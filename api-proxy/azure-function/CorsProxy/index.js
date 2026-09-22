/* =====================================================================
   CorsProxy — Azure Function (HTTP trigger)
   ---------------------------------------------------------------------
   A dumb relay: the browser sends { brand, url, method, headers, body,
   isForm }, this function forwards it to the target host (only if that
   host is on the allow-list below) and returns the response so the
   browser's CORS restrictions don't apply. It never inspects, logs, or
   stores credentials — whatever the browser sends is what gets sent.

   Deploy: Azure Functions, Node.js 18+, HTTP trigger, anonymous or
   function-key auth (function-key recommended — see docs/SETUP.md).
   Route this function at POST /api/relay to match js/proxy-client.js.
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

module.exports = async function (context, req) {
  // CORS preflight
  if (req.method === "OPTIONS") {
    context.res = corsResponse(204, "");
    return;
  }

  const { url, method, headers, body, isForm } = req.body || {};
  if (!url) {
    context.res = corsResponse(400, JSON.stringify({ error: "Missing 'url' in request body." }));
    return;
  }

  let target;
  try { target = new URL(url); } catch {
    context.res = corsResponse(400, JSON.stringify({ error: "Invalid URL." }));
    return;
  }
  if (!ALLOWED_HOSTS.has(target.hostname)) {
    context.res = corsResponse(403, JSON.stringify({ error: `Host '${target.hostname}' is not on the allow-list. Add it to ALLOWED_HOSTS in CorsProxy/index.js.` }));
    return;
  }

  const fetchHeaders = { ...(headers || {}) };
  if (isForm && !fetchHeaders["Content-Type"] && !fetchHeaders["content-type"]) {
    fetchHeaders["Content-Type"] = "application/x-www-form-urlencoded";
  }

  try {
    const upstream = await fetch(target.toString(), {
      method: method || "GET",
      headers: fetchHeaders,
      body: method && method.toUpperCase() !== "GET" ? body : undefined,
    });
    const text = await upstream.text();
    const respHeaders = {};
    upstream.headers.forEach((v, k) => (respHeaders[k] = v));

    context.res = corsResponse(200, JSON.stringify({
      status: upstream.status,
      body: text,
      headers: respHeaders,
    }));
  } catch (err) {
    context.res = corsResponse(502, JSON.stringify({ error: `Upstream request failed: ${err.message}` }));
  }
};

function corsResponse(status, body) {
  return {
    status,
    headers: {
      "Content-Type": "application/json",
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "POST, OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type",
    },
    body,
  };
}
