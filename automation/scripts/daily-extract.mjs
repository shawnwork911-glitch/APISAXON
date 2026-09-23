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
     2. Read every row in the sheet's "Connections" tab (including each
        company's saved Template export settings — station→facility_id
        mapping and Template 1/2 choice, same ones set in the web app's
        Extraction tab).
     3. For each connection with DailyAutoExtract = TRUE, pull "hourly"
        readings from its cursor date up to yesterday, respecting the
        brand's daily call quota (stops cleanly and resumes next run —
        same behaviour as fusionsolar_bot.py / solaredge_bot.py).
     4. Append the raw rows to the "Readings" tab (unchanged — this is
        the audit trail regardless of template settings).
     5. If a connection has a Template 1/2 export configured, group its
        rows by facility_id (summing stations that share one, exactly
        like the web app's manual download) and append the formatted
        rows into a dedicated tab per facility — auto-created the first
        time that facility is seen, named e.g. "GEN3294_T1" — since this
        is an unattended script there's no download prompt, so the
        Sheet itself is where the Template-formatted output lives.
     6. Write the new cursor back onto that connection's row.

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
  const existingTabs = await listExistingTabTitles(token);
  const connections = await listConnections(token);

  for (const conn of connections) {
    if (!conn.dailyAutoExtract) continue;
    console.log(`\n=== ${conn.companyName} (${conn.brand}) ===`);
    try {
      const rows = await extractForConnection(conn);
      if (rows.length) {
        await appendReadings(token, rows);
        console.log(`  -> appended ${rows.length} row(s) to the Readings tab`);

        const settings = conn.templateSettings;
        if (settings && (settings.template === "1" || settings.template === "2")) {
          const groups = buildTemplateGroups(rows, conn.brand, settings);
          for (const g of groups) {
            const tabName = sanitizeTabName(`${g.facilityId}_T${settings.template}`);
            await ensureSheetTab(token, existingTabs, tabName, g.headers);
            await appendRowsToTab(token, tabName, g.rows);
            console.log(`  -> appended ${g.rows.length} row(s) to "${tabName}" (facility ${g.facilityId})`);
          }
        }
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
    "ap-sg5": "https://sg5.fusionsolar.huawei.com/thirdData",
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

/* ---------------- Template 1 / 2 conversion (mirrors js/templates.js) ----------------
   Duplicated here rather than imported, same reasoning as the JWT-signing helpers
   below being duplicated from the Cloudflare Worker: this script is a standalone
   deployable unit, and keeping it self-contained avoids a cross-module-system import
   (this file is ESM; js/templates.js is loaded as a plain <script> in the browser).
   Numerically matched against the same reference files as js/templates.js:
   GEN3294_Hourly_..._template_1.xlsx and GEN3294_STX_..._template_2.xlsx. */

function pad(n, len = 2) { return String(n).padStart(len, "0"); }
function wcToTrickEpoch(wc) { return Date.UTC(wc.y, wc.mo - 1, wc.d, wc.H, wc.Mi, wc.S); }
function trickEpochToWc(epoch) {
  const d = new Date(epoch);
  return { y: d.getUTCFullYear(), mo: d.getUTCMonth() + 1, d: d.getUTCDate(), H: d.getUTCHours(), Mi: d.getUTCMinutes(), S: d.getUTCSeconds() };
}
function fmtSpace(wc) { return `${wc.y}-${pad(wc.mo)}-${pad(wc.d)} ${pad(wc.H)}:${pad(wc.Mi)}:${pad(wc.S)}`; }
function fmtOffsetSuffix(h) {
  const sign = h < 0 ? "-" : "+";
  const abs = Math.abs(h);
  const hh = Math.floor(abs), mm = Math.round((abs - hh) * 60);
  return `${sign}${pad(hh)}:${pad(mm)}`;
}
function fmtISOWithOffset(wc, offsetHours) { return `${wc.y}-${pad(wc.mo)}-${pad(wc.d)}T${pad(wc.H)}:${pad(wc.Mi)}:${pad(wc.S)}${fmtOffsetSuffix(offsetHours)}`; }
function fmtISOZ(wc) { return `${wc.y}-${pad(wc.mo)}-${pad(wc.d)}T${pad(wc.H)}:${pad(wc.Mi)}:${pad(wc.S)}Z`; }

function wallClockFromRow(row, brandKey, utcOffsetHours) {
  if (brandKey === "fusionsolar") {
    const localTrick = Number(row.timestamp) + utcOffsetHours * 3600 * 1000;
    return trickEpochToWc(localTrick);
  }
  const s = String(row.timestamp).trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})[ T](\d{1,2}):(\d{1,2})(?::(\d{1,2}))?/);
  if (m) return { y: +m[1], mo: +m[2], d: +m[3], H: +m[4], Mi: +m[5], S: +(m[6] || 0) };
  const n = Number(row.timestamp);
  if (!isNaN(n)) return trickEpochToWc(n + utcOffsetHours * 3600 * 1000);
  return null;
}

function templateHeadersFor(template) {
  return template === "1"
    ? ["month", "day", "year", "datetime_start_local", "local_timezone", "datetime_start_utc", "value (KWh)", "value (MWh)", "facility_id"]
    : ["datetime_start_local", "local_timezone", "datetime_start_utc", "value", "unit_of_measurement", "meter_id", "eac_facility_id", "eac_registry_id"];
}

function facilityKeyFor(stationId, settings) {
  const assigned = (settings.stationFacility || {})[stationId];
  return (assigned && assigned.trim()) || stationId;
}

// Same grouping/summing rule as js/templates.js: stations sharing a facility_id
// have their hourly kWh summed into one row per hour; unassigned stations stay
// their own separate group (keyed by their own stationId).
function buildTemplateGroups(rows, brandKey, settings) {
  const headers = templateHeadersFor(settings.template);
  const buckets = new Map(); // facilityId -> hourEpoch -> summed kWh

  for (const row of rows) {
    const wc = wallClockFromRow(row, brandKey, settings.utcOffset ?? 8);
    if (!wc) continue;
    const hourWc = { ...wc, Mi: 0, S: 0 };
    const localTrick = wcToTrickEpoch(hourWc);
    const facilityId = facilityKeyFor(row.stationId, settings);
    if (!buckets.has(facilityId)) buckets.set(facilityId, new Map());
    const hourMap = buckets.get(facilityId);
    hourMap.set(localTrick, (hourMap.get(localTrick) || 0) + (row.kwh || 0));
  }

  const groups = [];
  for (const [facilityId, hourMap] of buckets.entries()) {
    const groupSettings = (settings.facilityGroups || {})[facilityId] || {};
    const entries = [...hourMap.entries()].sort((a, b) => a[0] - b[0]);
    const outRows = entries.map(([localTrick, kwhSum]) => {
      const wc = trickEpochToWc(localTrick);
      const kwhR = Math.round(kwhSum * 1e3) / 1e3;
      const mwhR = Math.round((kwhR / 1000) * 1e6) / 1e6;
      const offset = settings.utcOffset ?? 8;
      if (settings.template === "1") {
        const utcWc = trickEpochToWc(localTrick + offset * 3600 * 1000);
        return [wc.mo, wc.d, wc.y, fmtSpace(wc), settings.tzLabel || "SGT", fmtSpace(utcWc), kwhR, mwhR, facilityId];
      }
      const utcWc = trickEpochToWc(localTrick - offset * 3600 * 1000);
      const value = settings.unitOM === "kWh" ? kwhR : mwhR;
      return [fmtISOWithOffset(wc, offset), settings.tzLabel || "Asia/Singapore", fmtISOZ(utcWc), value,
              settings.unitOM || "MWh", groupSettings.meterId || facilityId, facilityId, groupSettings.eacRegistryId || "tigr"];
    });
    groups.push({ facilityId, headers, rows: outRows });
  }
  return groups;
}

function sanitizeTabName(name) {
  // Google Sheets tab-name rules: no : \ / ? * [ ], 1–100 chars, not blank.
  const cleaned = String(name).replace(/[:\\/?*\[\]]/g, "_").trim();
  return (cleaned || "facility").slice(0, 100);
}

/* ---------------- Google Sheets ---------------- */

async function sheetsFetch(token, path, opts = {}) {
  const resp = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${GOOGLE_SPREADSHEET_ID}${path}`, {
    ...opts,
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json", ...(opts.headers || {}) },
  });
  if (!resp.ok) throw new Error(`Sheets API ${opts.method || "GET"} ${path} -> ${resp.status}: ${(await resp.text()).slice(0, 300)}`);
  return resp.json();
}

async function listExistingTabTitles(token) {
  const data = await sheetsFetch(token, `?fields=sheets.properties.title`);
  return new Set((data.sheets || []).map(s => s.properties.title));
}

// Creates the tab (and writes its header row) the first time a facility/template
// combination is seen; a Set tracks what's already been created/confirmed this run
// so repeated facilities across connections don't re-check the Sheets API each time.
async function ensureSheetTab(token, existingTabs, tabName, headers) {
  if (existingTabs.has(tabName)) return;
  await sheetsFetch(token, `:batchUpdate`, {
    method: "POST", body: JSON.stringify({ requests: [{ addSheet: { properties: { title: tabName } } }] }),
  });
  await sheetsFetch(token, `/values/${tabName}!A1:${colLetter(headers.length)}1?valueInputOption=RAW`, {
    method: "PUT", body: JSON.stringify({ values: [headers] }),
  });
  existingTabs.add(tabName);
}

function colLetter(n) {
  let s = "";
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

async function appendRowsToTab(token, tabName, rows) {
  for (let i = 0; i < rows.length; i += 500) {
    await sheetsFetch(token, `/values/${tabName}!A:${colLetter(rows[0].length)}:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`, {
      method: "POST", body: JSON.stringify({ values: rows.slice(i, i + 500) }),
    });
  }
}

async function listConnections(token) {
  const data = await sheetsFetch(token, `/values/${CONNECTIONS_SHEET}!A2:I1000`);
  const rows = data.values || [];
  return rows
    .map((r, i) => ({ rowIndex: i + 2, id: r[0], companyName: r[1], brand: r[2], region: r[3],
      credentials: safeJson(r[4], {}), stations: safeJson(r[5], []), dailyAutoExtract: r[6] === "TRUE", cursor: safeJson(r[7], {}),
      templateSettings: safeJson(r[8], null) }))
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
