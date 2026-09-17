/* =====================================================================
   CorsProxy — Cloudflare Worker
   ---------------------------------------------------------------------
   Same contract as the Azure Function in ../azure-function: POST a
   { url, method, headers, body, isForm } JSON payload to this worker's
   /relay path and it forwards the request, returning
   { status, body, headers }. Use this if you'd rather not stand up an
   Azure Function — deploy with `wrangler deploy` (free tier is plenty
   for this volume of calls).
   ===================================================================== */

const ALLOWED_HOSTS = new Set([
  "intl.fusionsolar.huawei.com",
  "uni005eu5.fusionsolar.huawei.com",
  "eu5.fusionsolar.huawei.com",
  "monitoringapi.solaredge.com",
  "gateway.isolarcloud.com",
  "www.soliscloud.com",
  "openapi.growatt.com",
  "server.growatt.com",
  "async-auth.smaapis.de",
  "sandbox.smaapis.de",
]);

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type",
};

export default {
  async fetch(request) {
    const { pathname } = new URL(request.url);
    if (request.method === "OPTIONS") return new Response(null, { status: 204, headers: CORS_HEADERS });
    if (pathname !== "/relay" || request.method !== "POST") {
      return json(404, { error: "Not found. POST to /relay." });
    }

    let payload;
    try { payload = await request.json(); } catch { return json(400, { error: "Invalid JSON body." }); }
    const { url, method, headers, body, isForm } = payload || {};
    if (!url) return json(400, { error: "Missing 'url' in request body." });

    let target;
    try { target = new URL(url); } catch { return json(400, { error: "Invalid URL." }); }
    if (!ALLOWED_HOSTS.has(target.hostname)) {
      return json(403, { error: `Host '${target.hostname}' is not on the allow-list. Add it to ALLOWED_HOSTS in worker.js.` });
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
  },
};

function json(status, obj) {
  return new Response(JSON.stringify(obj), { status, headers: { "Content-Type": "application/json", ...CORS_HEADERS } });
}
