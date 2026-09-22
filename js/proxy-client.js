/* =====================================================================
   Proxy client
   ---------------------------------------------------------------------
   Inverter-brand APIs are called server-to-server in your Postman
   collections and Python bots; browsers block most of these cross-origin
   (no CORS headers from monitoringapi.solaredge.com / fusionsolar / etc).
   So every brand call in this app is relayed through a small serverless
   proxy you deploy yourself — see /api-proxy for an Azure Function and a
   Cloudflare Worker template. The proxy only forwards requests to an
   allow-listed set of hostnames (see ALLOWED_HOSTS in both templates) —
   it never sees or stores your credentials, it just relays them.
   ===================================================================== */

const ProxyClient = (() => {
  let proxyBaseUrl = localStorage.getItem("slc.proxyBaseUrl") || "";

  function setProxyBaseUrl(url) {
    proxyBaseUrl = url.replace(/\/$/, "");
    localStorage.setItem("slc.proxyBaseUrl", proxyBaseUrl);
  }
  function getProxyBaseUrl() { return proxyBaseUrl; }

  /**
   * call(brandKey, { url, method, headers, body, isForm })
   * Returns { status, body (parsed JSON if possible, else text), headers }
   */
  async function call(brandKey, req) {
    if (!proxyBaseUrl) {
      throw new Error("No proxy URL configured yet — set it under Settings → CORS Proxy before connecting a brand.");
    }
    const payload = {
      brand: brandKey,
      url: req.url,
      method: req.method || "GET",
      headers: req.headers || {},
      body: req.body != null ? (typeof req.body === "string" ? req.body : JSON.stringify(req.body)) : undefined,
      isForm: !!req.isForm,
    };
    const resp = await fetch(`${proxyBaseUrl}/relay`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    const text = await resp.text();
    let parsed;
    try { parsed = JSON.parse(text); } catch { parsed = text; }
    if (!resp.ok) {
      const msg = (parsed && parsed.error) || `Proxy error ${resp.status}`;
      throw new Error(msg);
    }
    return { status: parsed.status, body: safeParseJson(parsed.body), headers: parsed.headers || {}, cookies: parsed.cookies || [] };
  }

  function safeParseJson(v) {
    if (typeof v !== "string") return v;
    try { return JSON.parse(v); } catch { return v; }
  }

  return { call, setProxyBaseUrl, getProxyBaseUrl };
})();
