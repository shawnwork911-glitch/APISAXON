/* =====================================================================
   CorsProxy — Cloudflare Worker
   ---------------------------------------------------------------------
   Two jobs, both under this one Worker:

   1. POST /relay   — dumb CORS relay to allow-listed inverter-brand APIs
      (unchanged from before; browsers can't call these directly).

   2. POST /sheets  — talks to Google Sheets on the browser's behalf,
      using a Google service account whose private key lives ONLY in
      this Worker's secrets (set via `wrangler secret put`), never sent
      to or readable by the browser. This replaces the earlier
      SharePoint/Graph storage layer with no Azure AD app or admin
      consent needed — just your own Google Cloud project.

   ACCESS CONTROL
   Every request to both endpoints must carry a Google ID token (from
   "Sign in with Google" on the web page) in an Authorization: Bearer
   header. This Worker verifies that token against Google itself, then
   checks the verified email against the "Users" tab of the same
   spreadsheet — a request with no token, an invalid/expired token, or
   an email not on that list is rejected with 401/403 before anything
   else runs. This means the access control lives HERE, not just in the
   web page — calling this Worker's URL directly, bypassing the site
   entirely, is rejected the same way.

   Required Worker secrets/vars (see docs/SETUP.md):
     GOOGLE_SERVICE_ACCOUNT_EMAIL   (secret)
     GOOGLE_SERVICE_ACCOUNT_KEY     (secret — the PEM private key)
     GOOGLE_SPREADSHEET_ID          (var — not sensitive, the sheet's ID)
     GOOGLE_SIGNIN_CLIENT_ID        (var — not sensitive, the OAuth Client ID
                                      used for "Sign in with Google"; not a
                                      secret, this is the same value the
                                      browser uses openly)
   ===================================================================== */

const ALLOWED_HOSTS = new Set([
  "intl.fusionsolar.huawei.com",
  "monitoringapi.solaredge.com",
  "gateway.isolarcloud.com",
  "www.soliscloud.com",
  "openapi.growatt.com",
  "server.growatt.com",
  "async-auth.smaapis.de",
  "sandbox.smaapis.de",
]);
// FusionSolar issues a different regional API gateway per account (sg5, au5,
// eu5, la5, etc. — e.g. "sg5.fusionsolar.huawei.com"), so rather than list
// every region by hand, any subdomain of fusionsolar.huawei.com is allowed.
// (Not "uni..." domains — those are Huawei's human login portal, a
// different thing from the Northbound API gateway.)
const ALLOWED_SUFFIXES = [".fusionsolar.huawei.com"];

function isAllowedHost(hostname) {
  return ALLOWED_HOSTS.has(hostname) || ALLOWED_SUFFIXES.some(suf => hostname.endsWith(suf));
}

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type, Authorization",
};

export default {
  async fetch(request, env) {
    const { pathname } = new URL(request.url);
    if (request.method === "OPTIONS") return new Response(null, { status: 204, headers: CORS_HEADERS });

    // Every real request — /relay included — must present a verified,
    // allow-listed identity before anything else runs.
    const authResult = await authorize(request, env);
    if (authResult instanceof Response) return authResult; // 401/403 — rejected

    if (pathname === "/relay" && request.method === "POST") return handleRelay(request);
    if (pathname === "/sheets" && request.method === "POST") return handleSheets(request, env, authResult);

    return json(404, { error: "Not found. POST to /relay or /sheets." });
  },
};

/* ---------------------------- access control ---------------------------- */

const USERS_SHEET = "Users";

// Verifies the caller's Google ID token against Google itself (so a forged
// or expired token is rejected), confirms it was issued for THIS app (the
// aud check — otherwise a valid Google token from an unrelated app would
// pass), then looks up the verified email in the Users tab. Returns
// { email, role } on success, or a ready-to-send Response on failure.
async function authorize(request, env) {
  if (!env.GOOGLE_SIGNIN_CLIENT_ID) {
    return json(500, { error: "Worker is missing GOOGLE_SIGNIN_CLIENT_ID — set it in wrangler.toml [vars]. See docs/SETUP.md." });
  }
  const authHeader = request.headers.get("Authorization") || "";
  const idToken = authHeader.startsWith("Bearer ") ? authHeader.slice(7) : null;
  if (!idToken) return json(401, { error: "Not signed in.", code: "NO_TOKEN" });

  let info;
  try {
    const resp = await fetch(`https://oauth2.googleapis.com/tokeninfo?id_token=${encodeURIComponent(idToken)}`);
    if (!resp.ok) return json(401, { error: "Sign-in expired or invalid — please sign in again.", code: "BAD_TOKEN" });
    info = await resp.json();
  } catch {
    return json(401, { error: "Could not verify sign-in — please sign in again.", code: "BAD_TOKEN" });
  }
  if (info.aud !== env.GOOGLE_SIGNIN_CLIENT_ID) return json(401, { error: "Sign-in token was not issued for this app.", code: "BAD_TOKEN" });
  if (info.email_verified !== "true" && info.email_verified !== true) return json(403, { error: "Google account email is not verified." });

  const email = String(info.email || "").trim().toLowerCase();
  let role;
  try {
    role = await getGoogleAccessToken(env).then(token => getUserRole(env, token, email));
  } catch (err) {
    return json(502, { error: `Could not check the Users list: ${err.message}` });
  }
  if (!role) {
    return json(403, {
      error: `${email} is not on the approved users list yet. Ask your admin to add this email to the "Users" tab of the Google Sheet.`,
      code: "NOT_AUTHORIZED",
    });
  }
  return { email, role };
}

async function getUserRole(env, token, email) {
  const data = await sheetsFetch(env, token, `/values/${USERS_SHEET}!A2:B1000`);
  const rows = data.values || [];
  const match = rows.find(r => (r[0] || "").trim().toLowerCase() === email);
  return match ? (match[1] || "User").trim() : null;
}

/* ---------------------------- brand API relay ---------------------------- */

async function handleRelay(request) {
  let payload;
  try { payload = await request.json(); } catch { return json(400, { error: "Invalid JSON body." }); }
  const { url, method, headers, body, isForm } = payload || {};
  if (!url) return json(400, { error: "Missing 'url' in request body." });

  let target;
  try { target = new URL(url); } catch { return json(400, { error: "Invalid URL." }); }
  if (!isAllowedHost(target.hostname)) {
    return json(403, { error: `Host '${target.hostname}' is not on the allow-list. Add it to ALLOWED_HOSTS/ALLOWED_SUFFIXES in worker.js.` });
  }

  const fetchHeaders = new Headers(headers || {});
  if (isForm && !fetchHeaders.has("Content-Type")) fetchHeaders.set("Content-Type", "application/x-www-form-urlencoded");

  try {
    const upstream = await fetch(target.toString(), {
      method: method || "GET",
      headers: fetchHeaders,
      body: method && method.toUpperCase() !== "GET" ? body : undefined,
    });
    const text = await upstream.text();
    const respHeaders = {};
    upstream.headers.forEach((v, k) => (respHeaders[k] = v));
    return json(200, { status: upstream.status, body: text, headers: respHeaders });
  } catch (err) {
    return json(502, { error: `Upstream request failed: ${err.message}` });
  }
}

/* ---------------------------- Google Sheets backend ---------------------------- */

const CONNECTIONS_SHEET = "Connections";
const READINGS_SHEET = "Readings";
// Connections columns: RowId, Title, Brand, Region, CredentialsJson, StationsJson, DailyAutoExtract, CursorJson, TemplateSettingsJson, SubscriptionJson
// SubscriptionJson: { hourly: bool, monthly: bool, startDate: "YYYY-MM-DD", endDate: "YYYY-MM-DD" } — drives the
// scheduled GitHub Action (see automation/scripts/daily-extract.mjs): Hourly is pulled day-by-day up to
// min(endDate, yesterday); once Hourly has fully caught up to endDate AND endDate has actually passed, Monthly
// (if ticked) fires once for the whole range.
// Readings columns:    Timestamp, Company, Brand, StationId, StationName, Resolution, kWh, RunAt

async function handleSheets(request, env, auth) {
  if (!env.GOOGLE_SERVICE_ACCOUNT_EMAIL || !env.GOOGLE_SERVICE_ACCOUNT_KEY || !env.GOOGLE_SPREADSHEET_ID) {
    return json(500, { error: "Worker is missing GOOGLE_SERVICE_ACCOUNT_EMAIL / GOOGLE_SERVICE_ACCOUNT_KEY / GOOGLE_SPREADSHEET_ID — set them with `wrangler secret put` / in wrangler.toml [vars]. See docs/SETUP.md." });
  }
  let payload;
  try { payload = await request.json(); } catch { return json(400, { error: "Invalid JSON body." }); }
  const { action } = payload || {};

  try {
    const token = await getGoogleAccessToken(env);
    switch (action) {
      case "whoAmI": return json(200, { email: auth.email, role: auth.role });
      case "listConnections": return json(200, await listConnections(env, token));
      case "saveConnection": return json(200, await saveConnection(env, token, payload.connection));
      case "deleteConnection":
        if (auth.role !== "Admin") return json(403, { error: "Only Admins can remove a company connection." });
        return json(200, await deleteConnection(env, token, payload.id));
      case "appendReadings": return json(200, await appendReadings(env, token, payload.rows));
      case "listReadings": return json(200, await listReadings(env, token, payload.company));
      default: return json(400, { error: `Unknown action '${action}'.` });
    }
  } catch (err) {
    return json(502, { error: err.message });
  }
}

async function sheetsFetch(env, token, path, opts = {}) {
  const resp = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${env.GOOGLE_SPREADSHEET_ID}${path}`, {
    ...opts,
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json", ...(opts.headers || {}) },
  });
  if (!resp.ok) throw new Error(`Sheets API ${opts.method || "GET"} ${path} -> ${resp.status}: ${(await resp.text()).slice(0, 300)}`);
  return resp.json();
}

async function listConnections(env, token) {
  const data = await sheetsFetch(env, token, `/values/${CONNECTIONS_SHEET}!A2:J1000`);
  const rows = data.values || [];
  return rows
    .map((r, i) => ({ rowIndex: i + 2, id: r[0], companyName: r[1], brand: r[2], region: r[3],
      credentials: safeJson(r[4], {}), stations: safeJson(r[5], []), dailyAutoExtract: r[6] === "TRUE" || r[6] === true, cursor: safeJson(r[7], {}),
      templateSettings: safeJson(r[8], null), subscription: safeJson(r[9], null) }))
    .filter(c => c.id && c.companyName); // blank Title = soft-deleted row
}

async function saveConnection(env, token, conn) {
  const id = conn.id || crypto.randomUUID();
  const values = [[id, conn.companyName, conn.brand, conn.region || "", JSON.stringify(conn.credentials || {}),
    JSON.stringify(conn.stations || []), conn.dailyAutoExtract ? "TRUE" : "FALSE", JSON.stringify(conn.cursor || {}),
    JSON.stringify(conn.templateSettings || null), JSON.stringify(conn.subscription || null)]];

  if (conn.id) {
    const rowIndex = await findRowIndex(env, token, CONNECTIONS_SHEET, conn.id);
    if (rowIndex) {
      await sheetsFetch(env, token, `/values/${CONNECTIONS_SHEET}!A${rowIndex}:J${rowIndex}?valueInputOption=RAW`, {
        method: "PUT", body: JSON.stringify({ values }),
      });
      return { id };
    }
  }
  await sheetsFetch(env, token, `/values/${CONNECTIONS_SHEET}!A:J:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`, {
    method: "POST", body: JSON.stringify({ values }),
  });
  return { id };
}

async function deleteConnection(env, token, id) {
  const rowIndex = await findRowIndex(env, token, CONNECTIONS_SHEET, id);
  if (!rowIndex) return { deleted: false };
  // Clear rather than physically delete the row, so other rows' indices never shift underneath us.
  await sheetsFetch(env, token, `/values/${CONNECTIONS_SHEET}!A${rowIndex}:J${rowIndex}:clear`, { method: "POST", body: "{}" });
  return { deleted: true };
}

async function findRowIndex(env, token, sheetName, id) {
  const data = await sheetsFetch(env, token, `/values/${sheetName}!A2:A1000`);
  const rows = data.values || [];
  const i = rows.findIndex(r => r[0] === id);
  return i === -1 ? null : i + 2;
}

async function appendReadings(env, token, rows) {
  const runAt = new Date().toISOString();
  const values = (rows || []).map(r => [r.timestamp, r.company, r.brand, r.stationId, r.stationName, r.resolution, r.kwh, runAt]);
  for (let i = 0; i < values.length; i += 500) {
    await sheetsFetch(env, token, `/values/${READINGS_SHEET}!A:H:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`, {
      method: "POST", body: JSON.stringify({ values: values.slice(i, i + 500) }),
    });
  }
  return { appended: values.length };
}

// Reads back rows already synced to the Readings tab — used by the Compare
// view (hourly-vs-monthly). Filters by company server-side so the browser
// only receives what it needs, since Readings can grow large over time.
async function listReadings(env, token, companyName) {
  const data = await sheetsFetch(env, token, `/values/${READINGS_SHEET}!A2:H200000`);
  const rows = data.values || [];
  return rows
    .filter(r => !companyName || r[1] === companyName)
    .map(r => ({ timestamp: r[0], company: r[1], brand: r[2], stationId: r[3], stationName: r[4], resolution: r[5], kwh: Number(r[6]), runAt: r[7] }));
}

function safeJson(s, fallback) { try { return JSON.parse(s); } catch { return fallback; } }

/* ---------------------------- Google service-account auth ---------------------------- */
// Signs a JWT with the service account's private key (RS256, via Web Crypto)
// and exchanges it for a short-lived Google access token. Cached in module
// scope for the life of this Worker isolate to avoid re-signing every call.

let cachedToken = null; // { token, expiresAt }

async function getGoogleAccessToken(env) {
  if (cachedToken && cachedToken.expiresAt > Date.now() + 30000) return cachedToken.token;

  const now = Math.floor(Date.now() / 1000);
  const header = { alg: "RS256", typ: "JWT" };
  const claim = {
    iss: env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
    scope: "https://www.googleapis.com/auth/spreadsheets",
    aud: "https://oauth2.googleapis.com/token",
    iat: now, exp: now + 3600,
  };
  const signingInput = `${b64url(JSON.stringify(header))}.${b64url(JSON.stringify(claim))}`;
  const key = await importPrivateKey(env.GOOGLE_SERVICE_ACCOUNT_KEY);
  const sigBuf = await crypto.subtle.sign("RSASSA-PKCS1-v1_5", key, new TextEncoder().encode(signingInput));
  const jwt = `${signingInput}.${b64urlFromBuffer(sigBuf)}`;

  const resp = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: `grant_type=${encodeURIComponent("urn:ietf:params:oauth:grant-type:jwt-bearer")}&assertion=${jwt}`,
  });
  const data = await resp.json();
  if (!data.access_token) throw new Error(`Google auth failed: ${data.error_description || JSON.stringify(data)}`);
  cachedToken = { token: data.access_token, expiresAt: Date.now() + data.expires_in * 1000 };
  return cachedToken.token;
}

async function importPrivateKey(pem) {
  const cleaned = pem.replace(/\\n/g, "\n").replace(/-----BEGIN PRIVATE KEY-----/, "").replace(/-----END PRIVATE KEY-----/, "").replace(/\s/g, "");
  const der = Uint8Array.from(atob(cleaned), c => c.charCodeAt(0));
  return crypto.subtle.importKey("pkcs8", der, { name: "RSASSA-PKCS1-v1_5", hash: "SHA-256" }, false, ["sign"]);
}

function b64url(str) { return b64urlFromBuffer(new TextEncoder().encode(str)); }
function b64urlFromBuffer(buf) {
  const bytes = new Uint8Array(buf);
  let s = ""; for (const b of bytes) s += String.fromCharCode(b);
  return btoa(s).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

function json(status, obj) {
  return new Response(JSON.stringify(obj), { status, headers: { "Content-Type": "application/json", ...CORS_HEADERS } });
}
