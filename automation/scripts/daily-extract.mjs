#!/usr/bin/env node
/* =====================================================================
   daily-extract.mjs
   ---------------------------------------------------------------------
   Runs on a GitHub Actions schedule (see ../../.github/workflows/
   daily-extraction.yml). Runs server-side, so it calls each brand's API
   directly — no CORS proxy needed for that part. For storage it talks
   to Google Sheets directly too, using its own copy of the service
   account credential (as GitHub secrets) — no Azure AD, no admin
   consent, nothing shared with the browser's copy of the credential.

   What it does, once per run:
     1. Sign a JWT with the service account key, exchange it for a
        Google access token (same pattern as the CORS proxy's /sheets
        route, just running in Node instead of a Workers isolate).
     2. Read every row in the sheet's "Connections" tab.
     3. For each connection with DailyAutoExtract = TRUE, pull "hourly"
        readings from its cursor date up to yesterday, respecting the
        brand's daily call quota (stops cleanly and resumes next run —
        same behaviour as fusionsolar_bot.py / solaredge_bot.py).
     4. Append new rows to the "Readings" tab.
     5. Write the new cursor back onto that connection's row.

   Required environment variables (set as GitHub Actions secrets):
     GOOGLE_SERVICE_ACCOUNT_EMAIL
     GOOGLE_SERVICE_ACCOUNT_KEY      (the PEM private key)
     GOOGLE_SPREADSHEET_ID
   ===================================================================== */

import crypto from "node:crypto";

const { GOOGLE_SERVICE_ACCOUNT_EMAIL, GOOGLE_SERVICE_ACCOUNT_KEY, GOOGLE_SPREADSHEET_ID } = process.env;

for (const v of ["GOOGLE_SERVICE_ACCOUNT_EMAIL", "GOOGLE_SERVICE_ACCOUNT_KEY", "GOOGLE_SPREADSHEET_ID"]) {
  if (!process.env[v]) { console.error(`Missing required env var ${v}`); process.exit(1); }
}

const CONNECTIONS_SHEET = "Connections";
const READINGS_SHEET = "Readings";

const QUOTAS = {
  fusionsolar: { perDay: 25, delayMs: 2000 },
  solaredge: { perDay: 300, delayMs: 400 },
  sungrow: { perDay: 100, delayMs: 1000 },
  solis: { perDay: 100, delayMs: 1000 },
  growatt: { perDay: 100, delayMs: 1000 },
  sma: { perDay: 100, delayMs: 1000 },
};

async function main() {
  const token = await getGoogleAccessToken();
  const connections = await listConnections(token);

  for (const conn of connections) {
    if (!conn.dailyAutoExtract) continue;
    console.log(`\n=== ${conn.companyName} (${conn.brand}) ===`);
    try {
      const rows = await extractForConnection(conn);
      if (rows.length) {
        await appendReadings(token, rows);
        console.log(`  -> appended ${rows.length} row(s) to the Readings tab`);
      } else {
        console.log("  -> nothing new to extract (quota reached, or already up to date)");
      }
      await saveCursor(token, conn);
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

/* ---------------- Google Sheets ---------------- */

async function sheetsFetch(token, path, opts = {}) {
  const resp = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${GOOGLE_SPREADSHEET_ID}${path}`, {
    ...opts,
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json", ...(opts.headers || {}) },
  });
  if (!resp.ok) throw new Error(`Sheets API ${opts.method || "GET"} ${path} -> ${resp.status}: ${(await resp.text()).slice(0, 300)}`);
  return resp.json();
}

async function listConnections(token) {
  const data = await sheetsFetch(token, `/values/${CONNECTIONS_SHEET}!A2:H1000`);
  const rows = data.values || [];
  return rows
    .map((r, i) => ({ rowIndex: i + 2, id: r[0], companyName: r[1], brand: r[2], region: r[3],
      credentials: safeJson(r[4], {}), stations: safeJson(r[5], []), dailyAutoExtract: r[6] === "TRUE", cursor: safeJson(r[7], {}) }))
    .filter(c => c.id && c.companyName);
}

async function saveCursor(token, conn) {
  await sheetsFetch(token, `/values/${CONNECTIONS_SHEET}!H${conn.rowIndex}?valueInputOption=RAW`, {
    method: "PUT", body: JSON.stringify({ values: [[JSON.stringify(conn.cursor || {})]] }),
  });
}

async function appendReadings(token, rows) {
  const runAt = new Date().toISOString();
  const values = rows.map(r => [r.timestamp, r.company, r.brand, r.stationId, r.stationName, r.resolution, r.kwh, runAt]);
  for (let i = 0; i < values.length; i += 500) {
    await sheetsFetch(token, `/values/${READINGS_SHEET}!A:H:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`, {
      method: "POST", body: JSON.stringify({ values: values.slice(i, i + 500) }),
    });
  }
}

function safeJson(s, fallback) { try { return JSON.parse(s); } catch { return fallback; } }

/* ---------------- Google service-account auth ---------------- */

async function getGoogleAccessToken() {
  const now = Math.floor(Date.now() / 1000);
  const header = { alg: "RS256", typ: "JWT" };
  const claim = {
    iss: GOOGLE_SERVICE_ACCOUNT_EMAIL,
    scope: "https://www.googleapis.com/auth/spreadsheets",
    aud: "https://oauth2.googleapis.com/token",
    iat: now, exp: now + 3600,
  };
  const signingInput = `${b64url(JSON.stringify(header))}.${b64url(JSON.stringify(claim))}`;
  const signature = crypto.sign("RSA-SHA256", Buffer.from(signingInput), GOOGLE_SERVICE_ACCOUNT_KEY.replace(/\\n/g, "\n"));
  const jwt = `${signingInput}.${b64urlFromBuffer(signature)}`;

  const resp = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: `grant_type=${encodeURIComponent("urn:ietf:params:oauth:grant-type:jwt-bearer")}&assertion=${jwt}`,
  });
  const data = await resp.json();
  if (!data.access_token) throw new Error(`Google auth failed: ${data.error_description || JSON.stringify(data)}`);
  return data.access_token;
}

function b64url(str) { return b64urlFromBuffer(Buffer.from(str)); }
function b64urlFromBuffer(buf) { return buf.toString("base64").replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, ""); }

main().catch(err => { console.error(err); process.exit(1); });
