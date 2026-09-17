#!/usr/bin/env node
/* =====================================================================
   daily-extract.mjs
   ---------------------------------------------------------------------
   Runs on a GitHub Actions schedule (see ../../.github/workflows/
   daily-extraction.yml). Unlike the browser app, this runs server-side
   so it calls each brand's API directly — no CORS proxy needed.

   What it does, once per run:
     1. App-only sign-in to Microsoft Graph (client credentials).
     2. Read every row in the SharePoint "SolarConnections" list.
     3. For each connection with DailyAutoExtract = true, pull "hourly"
        readings from its cursor date up to yesterday, respecting the
        brand's daily call quota (stops cleanly and resumes next run —
        same behaviour as fusionsolar_bot.py / solaredge_bot.py).
     4. Append new rows to SolarExtractionData.xlsx (Readings table).
     5. Persist the new cursor back onto the SharePoint list item.

   Required environment variables (set as GitHub Actions secrets):
     MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET   (app registration
       with Application permission Sites.Selected or Sites.ReadWrite.All,
       admin-consented)
     SP_SITE_HOSTNAME   e.g. contoso.sharepoint.com
     SP_SITE_PATH       e.g. /sites/SolarOps
   ===================================================================== */

import * as XLSX from "xlsx";
import crypto from "node:crypto";

const {
  MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET,
  SP_SITE_HOSTNAME, SP_SITE_PATH,
  SP_LIBRARY_NAME, // display name of the document library, e.g. "Process Improvement" — defaults to "Documents"
  SP_FOLDER_PATH,  // folder within that library, e.g. "API Data" — defaults to library root
} = process.env;

for (const v of ["MS_TENANT_ID", "MS_CLIENT_ID", "MS_CLIENT_SECRET", "SP_SITE_HOSTNAME", "SP_SITE_PATH"]) {
  if (!process.env[v]) { console.error(`Missing required env var ${v}`); process.exit(1); }
}

const GRAPH = "https://graph.microsoft.com/v1.0";
const WORKBOOK_NAME = "SolarExtractionData.xlsx";
const TABLE_NAME = "Readings";

const QUOTAS = {
  fusionsolar: { perDay: 25, delayMs: 2000 },
  solaredge: { perDay: 300, delayMs: 400 },
  sungrow: { perDay: 100, delayMs: 1000 },
  solis: { perDay: 100, delayMs: 1000 },
  growatt: { perDay: 100, delayMs: 1000 },
  sma: { perDay: 100, delayMs: 1000 },
};

async function main() {
  const token = await getAppOnlyToken();
  const siteId = await getSiteId(token);
  const driveId = await resolveDriveId(token, siteId);
  const listId = await getConnectionsListId(token, siteId);
  const items = await listItems(token, siteId, listId);

  for (const item of items) {
    const conn = itemToConnection(item);
    if (!conn.dailyAutoExtract) continue;
    console.log(`\n=== ${conn.companyName} (${conn.brand}) ===`);
    try {
      const rows = await extractForConnection(conn);
      if (rows.length) {
        await appendReadings(token, driveId, rows);
        console.log(`  -> appended ${rows.length} row(s) to ${WORKBOOK_NAME}`);
      } else {
        console.log("  -> nothing new to extract (quota reached, or already up to date)");
      }
      await saveCursor(token, siteId, listId, item.id, conn.cursor);
    } catch (err) {
      console.error(`  [!] ${conn.companyName}: ${err.message}`);
    }
  }
}

/* ---------------- per-brand extraction (mirrors js/brands.js) ---------------- */

async function extractForConnection(conn) {
  const quota = QUOTAS[conn.brand] || { perDay: 50, delayMs: 1000 };
  const yesterday = new Date(Date.now() - 86400000).toISOString().slice(0, 10);
  const startCursor = conn.cursor?.Hourly || yesterday;

  if (conn.brand === "fusionsolar") return extractFusionSolar(conn, startCursor, yesterday, quota);
  if (conn.brand === "solaredge") return extractSolarEdge(conn, startCursor, yesterday, quota);

  console.log(`  [i] ${conn.brand} historical endpoint is not confirmed yet (best-effort brand) — skipping automated pull. See js/brands.js docsNote.`);
  return [];
}

async function extractFusionSolar(conn, startCursor, endCursor, quota) {
  const c = conn.credentials;
  const regionMap = {
    intl: "https://intl.fusionsolar.huawei.com/thirdData",
    "ap-sg5": "https://uni005eu5.fusionsolar.huawei.com/thirdData",
    eu5: "https://eu5.fusionsolar.huawei.com/thirdData",
  };
  const base = c.baseUrl || regionMap[c.region] || regionMap.intl;

  const loginResp = await fetch(`${base}/login`, {
    method: "POST", headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ userName: c.username, systemCode: c.systemCode }),
  });
  const loginData = await loginResp.json();
  if (!loginData.success) throw new Error(loginData.message || "FusionSolar login failed");
  const token = loginData?.data?.xsrfToken || loginData?.xsrfToken || loginResp.headers.get("xsrf-token");
  if (!token) throw new Error("FusionSolar login OK but no XSRF-TOKEN returned");
  const headers = { "Content-Type": "application/json", "XSRF-TOKEN": token };

  const stationCodes = (conn.stations || []).map(s => s.id).join(",");
  if (!stationCodes) throw new Error("No cached stations for this connection — reconnect it in the web app to fetch stations first.");

  const rows = [];
  let calls = 0;
  const days = enumerateDays(startCursor, endCursor);
  for (const day of days) {
    if (calls >= quota.perDay) break;
    const ms = new Date(`${day}T12:00:00`).getTime();
    const resp = await fetch(`${base}/getKpiStationHour`, { method: "POST", headers, body: JSON.stringify({ stationCodes, collectTime: ms }) });
    const data = await resp.json();
    calls++;
    if (data.failCode === 407) break; // daily quota hit on Huawei's side
    for (const row of (data.data || [])) {
      rows.push({
        timestamp: row.collectTime, company: conn.companyName, brand: "FusionSolar",
        stationId: row.stationCode, stationName: nameFor(conn, row.stationCode), resolution: "Hourly",
        kwh: Number(row.dataItemMap?.inverter_power ?? row.dataItemMap?.product_power ?? 0),
      });
    }
    conn.cursor = conn.cursor || {}; conn.cursor.Hourly = day;
    await sleep(quota.delayMs);
  }
  return rows;
}

async function extractSolarEdge(conn, startCursor, endCursor, quota) {
  const c = conn.credentials;
  const base = c.baseUrl || "https://monitoringapi.solaredge.com";
  const resp = await fetch(
    `${base}/site/${c.siteId}/energyDetails?timeUnit=HOUR&meters=PRODUCTION`
    + `&startTime=${startCursor} 00:00:00&endTime=${endCursor} 23:59:59&api_key=${encodeURIComponent(c.apiKey)}`
  );
  const data = await resp.json();
  const meters = data?.energyDetails?.meters || [];
  const rows = [];
  for (const m of meters) {
    for (const v of (m.values || [])) {
      rows.push({
        timestamp: v.date, company: conn.companyName, brand: "SolarEdge",
        stationId: c.siteId, stationName: nameFor(conn, c.siteId), resolution: "Hourly",
        kwh: (v.value || 0) / 1000,
      });
    }
  }
  conn.cursor = conn.cursor || {}; conn.cursor.Hourly = endCursor;
  await sleep(quota.delayMs);
  return rows;
}

function nameFor(conn, stationId) {
  return (conn.stations || []).find(s => s.id === stationId)?.name || stationId;
}
function enumerateDays(start, end) {
  const days = []; const d = new Date(start); const e = new Date(end);
  for (; d <= e; d.setDate(d.getDate() + 1)) days.push(d.toISOString().slice(0, 10));
  return days;
}
function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

/* ---------------- Graph / SharePoint plumbing ---------------- */

async function getAppOnlyToken() {
  const resp = await fetch(`https://login.microsoftonline.com/${MS_TENANT_ID}/oauth2/v2.0/token`, {
    method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: MS_CLIENT_ID, client_secret: MS_CLIENT_SECRET,
      scope: "https://graph.microsoft.com/.default", grant_type: "client_credentials",
    }),
  });
  const data = await resp.json();
  if (!data.access_token) throw new Error(`Could not get app-only token: ${data.error_description || JSON.stringify(data)}`);
  return data.access_token;
}

async function gfetch(token, path, opts = {}) {
  const resp = await fetch(`${GRAPH}${path}`, { ...opts, headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json", ...(opts.headers || {}) } });
  if (!resp.ok) throw new Error(`Graph ${opts.method || "GET"} ${path} -> ${resp.status}: ${(await resp.text()).slice(0, 300)}`);
  if (resp.status === 204) return null;
  const ct = resp.headers.get("content-type") || "";
  return ct.includes("json") ? resp.json() : resp.arrayBuffer();
}

async function getSiteId(token) {
  const site = await gfetch(token, `/sites/${SP_SITE_HOSTNAME}:${SP_SITE_PATH}`);
  return site.id;
}

// Resolves the document library to write the workbook into, by display
// name (e.g. "Process Improvement") rather than assuming the default
// "Documents" library — mirrors js/msgraph.js#resolveDriveId in the web app.
async function resolveDriveId(token, siteId) {
  const wanted = (SP_LIBRARY_NAME || "").trim();
  if (!wanted || wanted.toLowerCase() === "documents") {
    const drive = await gfetch(token, `/sites/${siteId}/drive`);
    return drive.id;
  }
  const drives = await gfetch(token, `/sites/${siteId}/drives`);
  const match = (drives.value || []).find(d => d.name.toLowerCase() === wanted.toLowerCase());
  if (!match) throw new Error(`Document library "${wanted}" not found on this site — check SP_LIBRARY_NAME.`);
  return match.id;
}

function folderPrefix() {
  const f = (SP_FOLDER_PATH || "").trim().replace(/^\/+|\/+$/g, "");
  return f ? `${f}/` : "";
}

async function getConnectionsListId(token, siteId) {
  const lists = await gfetch(token, `/sites/${siteId}/lists?$filter=displayName eq 'SolarConnections'`);
  if (!lists.value?.length) throw new Error("SolarConnections list not found — create at least one connection in the web app first.");
  return lists.value[0].id;
}

async function listItems(token, siteId, listId) {
  const res = await gfetch(token, `/sites/${siteId}/lists/${listId}/items?expand=fields`);
  return res.value || [];
}

function itemToConnection(item) {
  const f = item.fields || {};
  return {
    id: item.id, companyName: f.Title, brand: f.Brand, region: f.Region,
    credentials: safeJson(f.CredentialsJson, {}), stations: safeJson(f.StationsJson, []),
    dailyAutoExtract: !!f.DailyAutoExtract, cursor: safeJson(f.CursorJson, {}),
  };
}
function safeJson(s, fallback) { try { return JSON.parse(s); } catch { return fallback; } }

async function saveCursor(token, siteId, listId, itemId, cursor) {
  await gfetch(token, `/sites/${siteId}/lists/${listId}/items/${itemId}/fields`, {
    method: "PATCH", body: JSON.stringify({ CursorJson: JSON.stringify(cursor || {}) }),
  });
}

async function ensureWorkbookItemId(token, driveId) {
  const path = `${folderPrefix()}${WORKBOOK_NAME}`;
  try {
    const item = await gfetch(token, `/drives/${driveId}/root:/${path}`);
    return item.id;
  } catch {
    // :/content PUT creates any missing intermediate folders automatically.
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([["Timestamp", "Company", "Brand", "StationId", "StationName", "Resolution", "kWh", "RunAt"]]);
    XLSX.utils.book_append_sheet(wb, ws, "Readings");
    const buf = XLSX.write(wb, { bookType: "xlsx", type: "buffer" });
    const resp = await fetch(`${GRAPH}/drives/${driveId}/root:/${path}:/content`, {
      method: "PUT", headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/octet-stream" }, body: buf,
    });
    const item = await resp.json();
    await gfetch(token, `/drives/${driveId}/items/${item.id}/workbook/worksheets('Readings')/tables/add`, {
      method: "POST", body: JSON.stringify({ address: "A1:H1", hasHeaders: true }),
    }).catch(() => {});
    await gfetch(token, `/drives/${driveId}/items/${item.id}/workbook/worksheets('Readings')/tables/Table1`, {
      method: "PATCH", body: JSON.stringify({ name: TABLE_NAME }),
    }).catch(() => {});
    return item.id;
  }
}

async function appendReadings(token, driveId, rows) {
  const itemId = await ensureWorkbookItemId(token, driveId);
  const runAt = new Date().toISOString();
  const values = rows.map(r => [r.timestamp, r.company, r.brand, r.stationId, r.stationName, r.resolution, r.kwh, runAt]);
  for (let i = 0; i < values.length; i += 500) {
    await gfetch(token, `/drives/${driveId}/items/${itemId}/workbook/tables/${TABLE_NAME}/rows/add`, {
      method: "POST", body: JSON.stringify({ values: values.slice(i, i + 500) }),
    });
  }
}

main().catch(err => { console.error(err); process.exit(1); });
