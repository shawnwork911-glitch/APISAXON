/* =====================================================================
   Google Sheets storage client (browser side)
   ---------------------------------------------------------------------
   Replaces the earlier SharePoint/Graph/MSAL storage layer. There's no
   "sign in" step here at all — a Google service account's private key
   must never reach the browser, so all reads/writes go through your
   deployed CORS proxy's /sheets endpoint, which holds that key server-
   side and signs requests on the browser's behalf. See
   api-proxy/cloudflare-worker/worker.js (or the Azure Function
   equivalent) for that side of it.
   ===================================================================== */

const SheetsClient = (() => {
  function loadConfig() {
    const raw = localStorage.getItem("slc.sheetsConfig");
    return raw ? JSON.parse(raw) : null;
  }
  function saveConfig(cfg) {
    localStorage.setItem("slc.sheetsConfig", JSON.stringify(cfg));
  }
  function isConfigured() {
    const cfg = loadConfig();
    return !!(cfg?.spreadsheetId && ProxyClient.getProxyBaseUrl());
  }
  function sheetUrl() {
    const cfg = loadConfig();
    return cfg?.spreadsheetId ? `https://docs.google.com/spreadsheets/d/${cfg.spreadsheetId}/edit` : null;
  }

  async function call(action, extra = {}) {
    const proxyBase = ProxyClient.getProxyBaseUrl();
    if (!proxyBase) throw new Error("No proxy URL configured yet — set it under Settings → CORS Proxy.");
    const cfg = loadConfig();
    if (!cfg?.spreadsheetId) throw new Error("No Google Sheet configured yet — set the Spreadsheet ID under Settings → Google Sheets.");
    const resp = await fetch(`${proxyBase}/sheets`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ action, ...extra }),
    });
    const data = await resp.json().catch(() => ({}));
    if (!resp.ok) throw new Error(data.error || `Sheets proxy error ${resp.status}`);
    return data;
  }

  async function listConnections() {
    return call("listConnections");
  }
  async function saveConnection(conn) {
    const res = await call("saveConnection", { connection: conn });
    return res.id;
  }
  async function deleteConnection(id) {
    return call("deleteConnection", { id });
  }
  async function appendReadings(rows) {
    return call("appendReadings", { rows });
  }
  async function listReadings(company) {
    return call("listReadings", { company });
  }

  return { loadConfig, saveConfig, isConfigured, sheetUrl, listConnections, saveConnection, deleteConnection, appendReadings, listReadings };
})();
