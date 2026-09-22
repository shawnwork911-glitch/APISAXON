/* =====================================================================
   SolarLink Connect — Brand Registry
   ---------------------------------------------------------------------
   One entry per inverter-monitoring brand. Each entry describes:
     - the credential fields to render in "Add Company"
     - how to authenticate (login/token step, if any)
     - how to list stations/plants
     - how to pull hourly / daily / monthly readings
     - the brand's daily API call quota (used by the scheduler to pace
       requests and auto-resume across days, mirroring fusionsolar_bot.py
       and solaredge_bot.py)

   CONFIDENCE LEVEL
     verified   — taken directly from your Postman collections / bot
                  scripts and tested by you already (FusionSolar, SolarEdge)
     bestEffort — built from publicly published vendor docs or common
                  community integrations, NOT tested against a live
                  account. Treat endpoint paths/signing here as a
                  starting point — confirm against your own developer
                  portal access before relying on them.
   ===================================================================== */

const RATE_LIMIT_DEFAULTS = {
  fusionsolar: { perDay: 25, callDelayMs: 2000, windowMs: 5 * 60 * 1000, windowMax: 500 },
  solaredge:   { perDay: 300, callDelayMs: 400, windowMs: null, windowMax: null },
  sungrow:     { perDay: 100, callDelayMs: 1000, windowMs: null, windowMax: null },
  solis:       { perDay: 100, callDelayMs: 1000, windowMs: null, windowMax: null },
  growatt:     { perDay: 100, callDelayMs: 1000, windowMs: null, windowMax: null },
  sma:         { perDay: 100, callDelayMs: 1000, windowMs: null, windowMax: null },
};

/**
 * Every brand module implements:
 *   buildAuth(creds)              -> { headers, cookies, extra }  (called once per session)
 *   listStations(ctx)             -> [{ id, name, raw }]
 *   fetchSeries(ctx, {stationIds, resolution, startDate, endDate}) -> [{stationId, timestamp, kwh, raw}]
 * ctx = { call, creds, auth, proxyBaseUrl }
 * call(brandKey, { url, method, headers, body }) goes through the CORS proxy — see js/proxy-client.js
 */

const BRANDS = {
  /* ============================== FUSIONSOLAR (verified) ============================== */
  fusionsolar: {
    key: "fusionsolar",
    label: "FusionSolar",
    vendor: "Huawei",
    badge: "FS",
    color: "#e0812a",
    confidence: "verified",
    quota: RATE_LIMIT_DEFAULTS.fusionsolar,
    quotaNote: "Standard tier = 25 calls/day per account. Hourly/Daily/Monthly/Yearly KPI calls each take one station-batch per collectTime.",
    fields: [
      { key: "username", label: "API Username", type: "text", required: true,
        help: "The northbound API account username (often your portal login email)." },
      { key: "systemCode", label: "System Code", type: "password", required: true,
        help: "The API account's system code — this is the API password, issued when the northbound account was created (not always the same as your portal password)." },
      { key: "region", label: "Region Endpoint", type: "select", required: true,
        options: [
          { value: "intl", label: "Global — intl.fusionsolar.huawei.com", base: "https://intl.fusionsolar.huawei.com/thirdData" },
          { value: "ap-sg5", label: "Asia-Pacific — sg5", base: "https://sg5.fusionsolar.huawei.com/thirdData" },
          { value: "eu5", label: "Europe — eu5", base: "https://eu5.fusionsolar.huawei.com/thirdData" },
          { value: "custom", label: "Custom / override below", base: "" },
        ] },
      { key: "baseUrl", label: "Base URL", type: "text", required: false,
        placeholder: "https://intl.fusionsolar.huawei.com/thirdData",
        help: "Auto-filled from the Region Endpoint picked above. Edit it only if your account uses a different gateway than the listed regions (or pick \"Custom / override below\" and type it in directly)." },
    ],
    async buildAuth(creds, call) {
      const base = creds.baseUrl?.trim() || BRANDS.fusionsolar.fields[2].options.find(o => o.value === creds.region)?.base;
      const res = await call("fusionsolar", {
        url: `${base}/login`,
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: { userName: creds.username, systemCode: creds.systemCode },
      });
      if (!res.body?.success) {
        const raw = typeof res.body === "string" ? res.body.slice(0, 400) : JSON.stringify(res.body ?? "").slice(0, 400);
        throw new Error(res.body?.message || `Login failed (failCode ${res.body?.failCode ?? "?"}) — raw response: ${raw}`);
      }
      const token = res.body?.data?.xsrfToken || res.body?.xsrfToken || res.headers?.["xsrf-token"];
      if (!token) throw new Error("Login succeeded but no XSRF-TOKEN was returned.");
      return { base, headers: { "XSRF-TOKEN": token, "Content-Type": "application/json" } };
    },
    async listStations(ctx) {
      const res = await ctx.call("fusionsolar", {
        url: `${ctx.auth.base}/getStationList`, method: "POST",
        headers: ctx.auth.headers, body: {},
      });
      const list = res.body?.data || [];
      return list.map(s => ({ id: s.stationCode, name: s.stationName || s.stationCode, raw: s }));
    },
    // FusionSolar's Kpi endpoints take ONE collectTime timestamp per call and
    // return the whole bucket that timestamp falls in (day for Hour endpoint,
    // month for Day endpoint, year for Month endpoint) — so we iterate.
    resolutionEndpoint: {
      Hourly: "getKpiStationHour", Daily: "getKpiStationDay", Monthly: "getKpiStationMonth",
    },
    async fetchSeries(ctx, { stationIds, resolution, startDate, endDate }) {
      const endpoint = BRANDS.fusionsolar.resolutionEndpoint[resolution];
      const codes = stationIds.join(",");
      const cursors = iterateFusionCursors(resolution, startDate, endDate);
      const rows = [];
      for (const ms of cursors) {
        const res = await ctx.call("fusionsolar", {
          url: `${ctx.auth.base}/${endpoint}`, method: "POST",
          headers: ctx.auth.headers, body: { stationCodes: codes, collectTime: ms },
        });
        if (res.body?.failCode === 407) throw new RateLimitReached();
        for (const row of (res.body?.data || [])) {
          rows.push({
            stationId: row.stationCode, timestamp: row.collectTime,
            kwh: Number(row.dataItemMap?.inverter_power ?? row.dataItemMap?.product_power ?? row.dataItemMap?.PVYield ?? 0),
            raw: row,
          });
        }
      }
      return rows;
    },
  },

  /* ============================== SOLAREDGE (verified) ============================== */
  solaredge: {
    key: "solaredge",
    label: "SolarEdge",
    vendor: "SolarEdge",
    badge: "SE",
    color: "#2f6fd6",
    confidence: "verified",
    quota: RATE_LIMIT_DEFAULTS.solaredge,
    quotaNote: "300 calls/day per API key. Hourly windows max 30 days, Daily windows max 365 days, Monthly unlimited — the extractor auto-chunks long ranges.",
    fields: [
      { key: "siteId", label: "Site ID", type: "text", required: true },
      { key: "apiKey", label: "API Key", type: "password", required: true },
      { key: "baseUrl", label: "Base URL", type: "text", required: false,
        placeholder: "https://monitoringapi.solaredge.com" },
    ],
    async buildAuth(creds) {
      const base = creds.baseUrl?.trim() || "https://monitoringapi.solaredge.com";
      return { base, apiKey: creds.apiKey, siteId: creds.siteId };
    },
    async listStations(ctx) {
      // SolarEdge's model is one site per credential set, so "stations" here
      // is just the one site — validate it via /details.
      const res = await ctx.call("solaredge", {
        url: `${ctx.auth.base}/site/${ctx.auth.siteId}/details?api_key=${encodeURIComponent(ctx.auth.apiKey)}`,
        method: "GET",
      });
      if (!res.body?.details) throw new Error("Could not read site details — check Site ID / API Key.");
      return [{ id: ctx.auth.siteId, name: res.body.details.name || ctx.auth.siteId, raw: res.body.details }];
    },
    resolutionUnit: { Hourly: "HOUR", Daily: "DAY", Monthly: "MONTH" },
    maxWindowDays: { HOUR: 30, DAY: 365, MONTH: null },
    async fetchSeries(ctx, { resolution, startDate, endDate }) {
      const unit = BRANDS.solaredge.resolutionUnit[resolution];
      const chunks = chunkDateRange(startDate, endDate, BRANDS.solaredge.maxWindowDays[unit]);
      const rows = [];
      for (const [s, e] of chunks) {
        const res = await ctx.call("solaredge", {
          url: `${ctx.auth.base}/site/${ctx.auth.siteId}/energyDetails`
             + `?timeUnit=${unit}&meters=PRODUCTION&startTime=${s} 00:00:00&endTime=${e} 23:59:59`
             + `&api_key=${encodeURIComponent(ctx.auth.apiKey)}`,
          method: "GET",
        });
        const meters = res.body?.energyDetails?.meters || [];
        for (const m of meters) {
          for (const v of (m.values || [])) {
            rows.push({ stationId: ctx.auth.siteId, timestamp: v.date, kwh: (v.value || 0) / 1000, raw: v });
          }
        }
      }
      return rows;
    },
  },

  /* ============================== SUNGROW (best-effort, UNTESTED) ============================== */
  sungrow: {
    key: "sungrow",
    label: "Sungrow",
    vendor: "iSolarCloud",
    badge: "SG",
    color: "#e0922a",
    confidence: "bestEffort",
    quota: RATE_LIMIT_DEFAULTS.sungrow,
    quotaNote: "Sungrow does not publish a fixed public quota for the Basic Version Open API — 100/day is a conservative placeholder. Adjust once you see your Developer Portal quota.",
    docsNote: "Endpoint names are only visible after your iSolarCloud Developer Portal application is approved (Documentation → Open API v1). Update src in js/brands.js once confirmed.",
    fields: [
      { key: "appKey", label: "App Key", type: "text", required: true },
      { key: "accessKey", label: "Secret Key / X-access-key", type: "password", required: true },
      { key: "username", label: "iSolarCloud Username (email)", type: "text", required: true },
      { key: "password", label: "iSolarCloud Password", type: "password", required: true },
      { key: "psId", label: "Plant / Station ID (ps_id)", type: "text", required: false,
        help: "Leave blank to auto-fetch from Plant List Information Query once connected." },
      { key: "baseUrl", label: "Base URL / Gateway", type: "text", required: false,
        placeholder: "https://gateway.isolarcloud.com" },
    ],
    async buildAuth(creds, call) {
      const base = creds.baseUrl?.trim() || "https://gateway.isolarcloud.com";
      const res = await call("sungrow", {
        url: `${base}/v1/userService/login`, method: "POST",
        headers: { "Content-Type": "application/json", "x-access-key": creds.accessKey, "sys_code": "901" },
        body: { appkey: creds.appKey, user_account: creds.username, user_password: creds.password },
      });
      const token = res.body?.result_data?.token || res.body?.token;
      if (!token) throw new Error(res.body?.result_msg || "Login did not return a session token — endpoint/signature likely needs adjustment from your Developer Portal docs.");
      return { base, token, appKey: creds.appKey, accessKey: creds.accessKey, psId: creds.psId };
    },
    async listStations(ctx) {
      if (ctx.auth.psId) return [{ id: ctx.auth.psId, name: ctx.auth.psId, raw: {} }];
      // Placeholder call — confirm the real "Plant List Information Query" path/params
      // from Documentation → Open API (V1) in the Developer Portal.
      const res = await ctx.call("sungrow", {
        url: `${ctx.auth.base}/v1/powerStationService/getPowerStationList`, method: "POST",
        headers: { "Content-Type": "application/json", "x-access-key": ctx.auth.accessKey, token: ctx.auth.token },
        body: { appkey: ctx.auth.appKey },
      });
      const list = res.body?.result_data?.pageList || [];
      return list.map(p => ({ id: p.ps_id, name: p.ps_name || p.ps_id, raw: p }));
    },
    async fetchSeries() {
      throw new Error("Sungrow historical-data endpoint is not confirmed yet — see docsNote in js/brands.js. Add the real path once your Developer Portal access is approved.");
    },
  },

  /* ============================== SOLIS (best-effort, UNTESTED) ============================== */
  solis: {
    key: "solis",
    label: "Solis",
    vendor: "SolisCloud",
    badge: "SL",
    color: "#7c5cd6",
    confidence: "bestEffort",
    quota: RATE_LIMIT_DEFAULTS.solis,
    quotaNote: "SolisCloud's published quota varies by account tier — 100/day is a conservative placeholder.",
    docsNote: "SolisCloud API requires HMAC-SHA1 request signing (Content-MD5 + Date + method + resource, signed with your API Secret). Implemented in js/signing.js#solisSign — verify field names against your SolisCloud API v1 PDF.",
    fields: [
      { key: "apiId", label: "API ID (KeyId)", type: "text", required: true },
      { key: "apiSecret", label: "API Secret (KeySecret)", type: "password", required: true },
      { key: "baseUrl", label: "Base URL", type: "text", required: false,
        placeholder: "https://www.soliscloud.com:13333" },
    ],
    async buildAuth(creds) {
      return { base: creds.baseUrl?.trim() || "https://www.soliscloud.com:13333", apiId: creds.apiId, apiSecret: creds.apiSecret };
    },
    async listStations(ctx) {
      const path = "/v1/api/userStationList";
      const body = JSON.stringify({ pageNo: 1, pageSize: 100 });
      const signed = solisSign(ctx.auth.apiSecret, "POST", path, body);
      const res = await ctx.call("solis", {
        url: `${ctx.auth.base}${path}`, method: "POST",
        headers: { "Content-Type": "application/json;charset=UTF-8", "Content-MD5": signed.md5, Date: signed.date,
                   Authorization: `API ${ctx.auth.apiId}:${signed.sign}` },
        body,
      });
      const list = res.body?.data?.page?.records || [];
      return list.map(p => ({ id: p.id, name: p.stationName || p.id, raw: p }));
    },
    resolutionPath: { Hourly: "/v1/api/stationDayEnergyList", Daily: "/v1/api/stationMonthEnergyList", Monthly: "/v1/api/stationYearEnergyList" },
    async fetchSeries(ctx, { stationIds, resolution, startDate, endDate }) {
      const path = BRANDS.solis.resolutionPath[resolution];
      const rows = [];
      for (const id of stationIds) {
        const body = JSON.stringify({ id, money: "USD", time: startDate });
        const signed = solisSign(ctx.auth.apiSecret, "POST", path, body);
        const res = await ctx.call("solis", {
          url: `${ctx.auth.base}${path}`, method: "POST",
          headers: { "Content-Type": "application/json;charset=UTF-8", "Content-MD5": signed.md5, Date: signed.date,
                     Authorization: `API ${ctx.auth.apiId}:${signed.sign}` },
          body,
        });
        for (const v of (res.body?.data || [])) {
          rows.push({ stationId: id, timestamp: v.date || v.time, kwh: Number(v.energy || 0), raw: v });
        }
      }
      return rows;
    },
  },

  /* ============================== GROWATT (best-effort, UNTESTED) ============================== */
  growatt: {
    key: "growatt",
    label: "Growatt",
    vendor: "Growatt",
    badge: "GW",
    color: "#3fae5a",
    confidence: "bestEffort",
    quota: RATE_LIMIT_DEFAULTS.growatt,
    quotaNote: "No fixed published quota confirmed — 100/day is a conservative placeholder.",
    docsNote: "Growatt has two API surfaces in the wild: the official OpenAPI v1 (token-based, needs a Growatt-issued application token) and the legacy portal API (username/password session, undocumented, used by community integrations). Both are wired here behind the 'API Mode' selector — confirm which one your account actually has access to.",
    fields: [
      { key: "apiMode", label: "API Mode", type: "select", required: true,
        options: [{ value: "openapi", label: "Open API (token)" }, { value: "legacy", label: "Legacy portal (username/password)" }] },
      { key: "token", label: "API Token", type: "password", required: false, showIf: { apiMode: "openapi" } },
      { key: "username", label: "Portal Username", type: "text", required: false, showIf: { apiMode: "legacy" } },
      { key: "password", label: "Portal Password", type: "password", required: false, showIf: { apiMode: "legacy" } },
      { key: "baseUrl", label: "Base URL", type: "text", required: false,
        placeholder: "https://openapi.growatt.com  (or https://server.growatt.com for legacy)" },
    ],
    async buildAuth(creds, call) {
      if (creds.apiMode === "legacy") {
        const base = creds.baseUrl?.trim() || "https://server.growatt.com";
        const res = await call("growatt", {
          url: `${base}/login`, method: "POST",
          headers: { "Content-Type": "application/x-www-form-urlencoded" },
          body: `account=${encodeURIComponent(creds.username)}&password=${encodeURIComponent(md5Hex(creds.password))}&validateCode=`,
          isForm: true,
        });
        if (res.body?.result !== 1) throw new Error(res.body?.msg || "Legacy login failed.");
        return { base, mode: "legacy", cookies: res.cookies };
      }
      const base = creds.baseUrl?.trim() || "https://openapi.growatt.com";
      return { base, mode: "openapi", token: creds.token };
    },
    async listStations(ctx) {
      if (ctx.auth.mode === "openapi") {
        const res = await ctx.call("growatt", { url: `${ctx.auth.base}/v1/plant/list`, method: "GET", headers: { token: ctx.auth.token } });
        const list = res.body?.data?.plants || [];
        return list.map(p => ({ id: p.plant_id, name: p.name || p.plant_id, raw: p }));
      }
      const res = await ctx.call("growatt", { url: `${ctx.auth.base}/PlantListAPI.do`, method: "GET", headers: {}, cookies: ctx.auth.cookies });
      const list = res.body?.back?.data || [];
      return list.map(p => ({ id: p.plantId, name: p.plantName || p.plantId, raw: p }));
    },
    async fetchSeries() {
      throw new Error("Growatt historical-data endpoint is not confirmed for your account/mode yet — see docsNote in js/brands.js.");
    },
  },

  /* ============================== SMA (best-effort, UNTESTED) ============================== */
  sma: {
    key: "sma",
    label: "SMA",
    vendor: "Sunny Portal / ennexOS",
    badge: "SMA",
    color: "#d33b3b",
    confidence: "bestEffort",
    quota: RATE_LIMIT_DEFAULTS.sma,
    quotaNote: "No fixed published quota confirmed — 100/day is a conservative placeholder.",
    docsNote: "SMA uses OAuth2 with a one-time, per-plant-owner BC-Authorize consent step (async). Connecting a company here only stores client_id/client_secret; use 'Request Owner Consent' once before the first extraction, and re-run it if the owner revokes access.",
    fields: [
      { key: "clientId", label: "Client ID", type: "text", required: true },
      { key: "clientSecret", label: "Client Secret", type: "password", required: true },
      { key: "loginHint", label: "Plant Owner Email (login_hint)", type: "text", required: true },
      { key: "env", label: "Environment", type: "select", required: true,
        options: [
          { value: "prod", label: "Production", base: "https://async-auth.smaapis.de" },
          { value: "sandbox", label: "Sandbox", base: "https://sandbox.smaapis.de" },
        ] },
    ],
    async buildAuth(creds, call) {
      const base = BRANDS.sma.fields[3].options.find(o => o.value === creds.env)?.base;
      const res = await call("sma", {
        url: `${base}/oauth2/v2/token`, method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: `grant_type=client_credentials&client_id=${encodeURIComponent(creds.clientId)}&client_secret=${encodeURIComponent(creds.clientSecret)}`,
        isForm: true,
      });
      if (!res.body?.access_token) throw new Error(res.body?.error_description || "Could not obtain an access token.");
      return { base, accessToken: res.body.access_token, loginHint: creds.loginHint };
    },
    async requestConsent(ctx, call) {
      return call("sma", {
        url: `${ctx.auth.base}/oauth2/v2/bc-authorize`, method: "POST",
        headers: { Authorization: `Bearer ${ctx.auth.accessToken}`, "Content-Type": "application/x-www-form-urlencoded" },
        body: `login_hint=${encodeURIComponent(ctx.auth.loginHint)}`,
        isForm: true,
      });
    },
    async listStations() {
      throw new Error("SMA plant listing endpoint path is not confirmed yet — the Monitoring API's REST paths are only fully published in SMA's developer contract docs. See docsNote in js/brands.js.");
    },
    async fetchSeries() {
      throw new Error("SMA historical-data endpoint is not confirmed yet — see docsNote in js/brands.js.");
    },
  },
};

/* ---------------------------- helpers ---------------------------- */

class RateLimitReached extends Error {
  constructor() { super("Daily API call quota reached — extraction will auto-resume tomorrow."); this.name = "RateLimitReached"; }
}

function pad(n) { return String(n).padStart(2, "0"); }

// FusionSolar Kpi endpoints: one timestamp per bucket (a day for Hour, a month for Day, a year for Month)
function iterateFusionCursors(resolution, startDate, endDate) {
  const start = new Date(startDate), end = new Date(endDate);
  const cursors = [];
  if (resolution === "Hourly") {
    for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
      cursors.push(new Date(d).setHours(12, 0, 0, 0));
    }
  } else if (resolution === "Daily") {
    for (let d = new Date(start.getFullYear(), start.getMonth(), 1); d <= end; d.setMonth(d.getMonth() + 1)) {
      cursors.push(new Date(d).setDate(15));
    }
  } else {
    for (let y = start.getFullYear(); y <= end.getFullYear(); y++) {
      cursors.push(new Date(y, 6, 1).getTime());
    }
  }
  return cursors;
}

function chunkDateRange(startDate, endDate, maxDays) {
  const start = new Date(startDate), end = new Date(endDate);
  if (!maxDays) return [[fmt(start), fmt(end)]];
  const chunks = [];
  let cur = new Date(start);
  while (cur <= end) {
    const chunkEnd = new Date(Math.min(cur.getTime() + (maxDays - 1) * 86400000, end.getTime()));
    chunks.push([fmt(cur), fmt(chunkEnd)]);
    cur = new Date(chunkEnd.getTime() + 86400000);
  }
  return chunks;
  function fmt(d) { return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; }
}

if (typeof module !== "undefined") module.exports = { BRANDS, RateLimitReached };
