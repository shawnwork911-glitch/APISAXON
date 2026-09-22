/* =====================================================================
   SheetsProxy — Azure Function (HTTP trigger)
   ---------------------------------------------------------------------
   Azure Function equivalent of the Cloudflare Worker's /sheets route —
   see ../../cloudflare-worker/worker.js for the full design notes. Uses
   Node's built-in crypto module to sign the service-account JWT instead
   of Web Crypto, since this runs in Node rather than a Workers isolate.

   Required Function App settings (Configuration → Application settings):
     GOOGLE_SERVICE_ACCOUNT_EMAIL
     GOOGLE_SERVICE_ACCOUNT_KEY      (the PEM private key)
     GOOGLE_SPREADSHEET_ID
   ===================================================================== */

const crypto = require("crypto");

const CONNECTIONS_SHEET = "Connections";
const READINGS_SHEET = "Readings";

let cachedToken = null;

module.exports = async function (context, req) {
  if (req.method === "OPTIONS") { context.res = corsResponse(204, ""); return; }

  const { GOOGLE_SERVICE_ACCOUNT_EMAIL, GOOGLE_SERVICE_ACCOUNT_KEY, GOOGLE_SPREADSHEET_ID } = process.env;
  if (!GOOGLE_SERVICE_ACCOUNT_EMAIL || !GOOGLE_SERVICE_ACCOUNT_KEY || !GOOGLE_SPREADSHEET_ID) {
    context.res = corsResponse(500, JSON.stringify({ error: "Missing GOOGLE_SERVICE_ACCOUNT_EMAIL / GOOGLE_SERVICE_ACCOUNT_KEY / GOOGLE_SPREADSHEET_ID app settings. See docs/SETUP.md." }));
    return;
  }

  const { action, connection, id, rows } = req.body || {};
  try {
    const token = await getGoogleAccessToken(GOOGLE_SERVICE_ACCOUNT_EMAIL, GOOGLE_SERVICE_ACCOUNT_KEY);
    let result;
    if (action === "listConnections") result = await listConnections(GOOGLE_SPREADSHEET_ID, token);
    else if (action === "saveConnection") result = await saveConnection(GOOGLE_SPREADSHEET_ID, token, connection);
    else if (action === "deleteConnection") result = await deleteConnection(GOOGLE_SPREADSHEET_ID, token, id);
    else if (action === "appendReadings") result = await appendReadings(GOOGLE_SPREADSHEET_ID, token, rows);
    else { context.res = corsResponse(400, JSON.stringify({ error: `Unknown action '${action}'.` })); return; }
    context.res = corsResponse(200, JSON.stringify(result));
  } catch (err) {
    context.res = corsResponse(502, JSON.stringify({ error: err.message }));
  }
};

async function sheetsFetch(spreadsheetId, token, path, opts = {}) {
  const resp = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}${path}`, {
    ...opts,
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json", ...(opts.headers || {}) },
  });
  if (!resp.ok) throw new Error(`Sheets API ${opts.method || "GET"} ${path} -> ${resp.status}: ${(await resp.text()).slice(0, 300)}`);
  return resp.json();
}

async function listConnections(spreadsheetId, token) {
  const data = await sheetsFetch(spreadsheetId, token, `/values/${CONNECTIONS_SHEET}!A2:H1000`);
  const rows = data.values || [];
  return rows
    .map((r, i) => ({ rowIndex: i + 2, id: r[0], companyName: r[1], brand: r[2], region: r[3],
      credentials: safeJson(r[4], {}), stations: safeJson(r[5], []), dailyAutoExtract: r[6] === "TRUE", cursor: safeJson(r[7], {}) }))
    .filter(c => c.id && c.companyName);
}

async function saveConnection(spreadsheetId, token, conn) {
  const id = conn.id || crypto.randomUUID();
  const values = [[id, conn.companyName, conn.brand, conn.region || "", JSON.stringify(conn.credentials || {}),
    JSON.stringify(conn.stations || []), conn.dailyAutoExtract ? "TRUE" : "FALSE", JSON.stringify(conn.cursor || {})]];
  if (conn.id) {
    const rowIndex = await findRowIndex(spreadsheetId, token, CONNECTIONS_SHEET, conn.id);
    if (rowIndex) {
      await sheetsFetch(spreadsheetId, token, `/values/${CONNECTIONS_SHEET}!A${rowIndex}:H${rowIndex}?valueInputOption=RAW`, { method: "PUT", body: JSON.stringify({ values }) });
      return { id };
    }
  }
  await sheetsFetch(spreadsheetId, token, `/values/${CONNECTIONS_SHEET}!A:H:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`, { method: "POST", body: JSON.stringify({ values }) });
  return { id };
}

async function deleteConnection(spreadsheetId, token, id) {
  const rowIndex = await findRowIndex(spreadsheetId, token, CONNECTIONS_SHEET, id);
  if (!rowIndex) return { deleted: false };
  await sheetsFetch(spreadsheetId, token, `/values/${CONNECTIONS_SHEET}!A${rowIndex}:H${rowIndex}:clear`, { method: "POST", body: "{}" });
  return { deleted: true };
}

async function findRowIndex(spreadsheetId, token, sheetName, id) {
  const data = await sheetsFetch(spreadsheetId, token, `/values/${sheetName}!A2:A1000`);
  const rows = data.values || [];
  const i = rows.findIndex(r => r[0] === id);
  return i === -1 ? null : i + 2;
}

async function appendReadings(spreadsheetId, token, rows) {
  const runAt = new Date().toISOString();
  const values = (rows || []).map(r => [r.timestamp, r.company, r.brand, r.stationId, r.stationName, r.resolution, r.kwh, runAt]);
  for (let i = 0; i < values.length; i += 500) {
    await sheetsFetch(spreadsheetId, token, `/values/${READINGS_SHEET}!A:H:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`, { method: "POST", body: JSON.stringify({ values: values.slice(i, i + 500) }) });
  }
  return { appended: values.length };
}

function safeJson(s, fallback) { try { return JSON.parse(s); } catch { return fallback; } }

async function getGoogleAccessToken(email, pemKey) {
  if (cachedToken && cachedToken.expiresAt > Date.now() + 30000) return cachedToken.token;
  const now = Math.floor(Date.now() / 1000);
  const header = { alg: "RS256", typ: "JWT" };
  const claim = { iss: email, scope: "https://www.googleapis.com/auth/spreadsheets", aud: "https://oauth2.googleapis.com/token", iat: now, exp: now + 3600 };
  const signingInput = `${b64url(JSON.stringify(header))}.${b64url(JSON.stringify(claim))}`;
  const signature = crypto.sign("RSA-SHA256", Buffer.from(signingInput), pemKey.replace(/\\n/g, "\n"));
  const jwt = `${signingInput}.${b64urlFromBuffer(signature)}`;

  const resp = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: `grant_type=${encodeURIComponent("urn:ietf:params:oauth:grant-type:jwt-bearer")}&assertion=${jwt}`,
  });
  const data = await resp.json();
  if (!data.access_token) throw new Error(`Google auth failed: ${data.error_description || JSON.stringify(data)}`);
  cachedToken = { token: data.access_token, expiresAt: Date.now() + data.expires_in * 1000 };
  return cachedToken.token;
}

function b64url(str) { return b64urlFromBuffer(Buffer.from(str)); }
function b64urlFromBuffer(buf) { return buf.toString("base64").replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, ""); }

function corsResponse(status, body) {
  return { status, headers: { "Content-Type": "application/json", "Access-Control-Allow-Origin": "*", "Access-Control-Allow-Methods": "POST, OPTIONS", "Access-Control-Allow-Headers": "Content-Type" }, body };
}
