/* =====================================================================
   Run context + Google service-account / Sheets helpers.

   Every outbound request goes through rc.fetch(), which counts it against
   the Worker's per-invocation subrequest limit (50 on the Workers Free
   plan, 1000 on Paid — set MAX_SUBREQUESTS to match yours). The scheduler
   stops cleanly before hitting that limit and continues on the next run,
   instead of being killed half-way through a job.
   ===================================================================== */

export class OutOfBudget extends Error {
  constructor() { super("Subrequest budget for this run used up — continues on the next run."); this.name = "OutOfBudget"; }
}

export function makeRunContext(env, { deadlineMs, now = () => Date.now(), sleep } = {}) {
  const max = Number(env.MAX_SUBREQUESTS) || 45;
  return {
    env,
    used: 0,
    max,
    deadline: deadlineMs ?? now() + 25000,
    now,
    sleep: sleep || (ms => new Promise(r => setTimeout(r, ms))),
    left() { return this.max - this.used; },
    timeLeft() { return this.deadline - this.now(); },
    async fetch(url, opts) {
      if (this.used >= this.max) throw new OutOfBudget();
      this.used++;
      return fetch(url, opts);
    },
  };
}

/* ---------------------------- Google access token ---------------------------- */

let cachedToken = null; // { token, expiresAt } — reused for the life of the isolate

export async function getGoogleAccessToken(rc) {
  if (cachedToken && cachedToken.expiresAt > Date.now() + 60000) return cachedToken.token;
  const { GOOGLE_SERVICE_ACCOUNT_EMAIL, GOOGLE_SERVICE_ACCOUNT_KEY } = rc.env;
  const now = Math.floor(Date.now() / 1000);
  const header = { alg: "RS256", typ: "JWT" };
  const claim = {
    iss: GOOGLE_SERVICE_ACCOUNT_EMAIL,
    scope: "https://www.googleapis.com/auth/spreadsheets",
    aud: "https://oauth2.googleapis.com/token",
    iat: now, exp: now + 3600,
  };
  const signingInput = `${b64url(JSON.stringify(header))}.${b64url(JSON.stringify(claim))}`;
  const key = await importPrivateKey(GOOGLE_SERVICE_ACCOUNT_KEY);
  const sig = await crypto.subtle.sign("RSASSA-PKCS1-v1_5", key, new TextEncoder().encode(signingInput));
  const jwt = `${signingInput}.${b64urlFromBuffer(sig)}`;

  const resp = await rc.fetch("https://oauth2.googleapis.com/token", {
    method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: `grant_type=${encodeURIComponent("urn:ietf:params:oauth:grant-type:jwt-bearer")}&assertion=${jwt}`,
  });
  const data = await resp.json();
  if (!data.access_token) throw new Error(`Google auth failed: ${data.error_description || JSON.stringify(data)}`);
  cachedToken = { token: data.access_token, expiresAt: Date.now() + data.expires_in * 1000 };
  return cachedToken.token;
}

async function importPrivateKey(pem) {
  const cleaned = String(pem || "").replace(/\\n/g, "\n")
    .replace(/-----BEGIN PRIVATE KEY-----/, "").replace(/-----END PRIVATE KEY-----/, "").replace(/\s/g, "");
  const der = Uint8Array.from(atob(cleaned), c => c.charCodeAt(0));
  return crypto.subtle.importKey("pkcs8", der, { name: "RSASSA-PKCS1-v1_5", hash: "SHA-256" }, false, ["sign"]);
}

function b64url(str) { return b64urlFromBuffer(new TextEncoder().encode(str)); }
function b64urlFromBuffer(buf) {
  const bytes = new Uint8Array(buf);
  let s = ""; for (const b of bytes) s += String.fromCharCode(b);
  return btoa(s).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

/* ---------------------------- Sheets ---------------------------- */

export async function sheets(rc, path, opts = {}) {
  const token = await getGoogleAccessToken(rc);
  const resp = await rc.fetch(`https://sheets.googleapis.com/v4/spreadsheets/${rc.env.GOOGLE_SPREADSHEET_ID}${path}`, {
    ...opts,
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json", ...(opts.headers || {}) },
  });
  if (!resp.ok) {
    const err = new Error(`Sheets API ${opts.method || "GET"} ${path.split("?")[0]} -> ${resp.status}: ${(await resp.text()).slice(0, 300)}`);
    err.status = resp.status;
    throw err;
  }
  return resp.json();
}

export function safeJson(s, fallback) { try { return s ? JSON.parse(s) : fallback; } catch { return fallback; } }
