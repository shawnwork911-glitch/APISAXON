/* =====================================================================
   Scheduler client — talks to the auto-continue scheduler Worker
   (api-proxy/cloudflare-scheduler). That Worker owns the Sheet's "Jobs"
   tab and continues each job every hour on its own, so a long or
   future-dated range finishes even when nobody has this page open.
   ===================================================================== */

const DEFAULT_SCHEDULER_URL = "https://solarlink-scheduler.saxon-solarlink.workers.dev";

const SchedulerClient = (() => {
  const KEY = "slc.schedulerUrl";
  function getBaseUrl() {
    try { return localStorage.getItem(KEY) ?? DEFAULT_SCHEDULER_URL; } catch { return DEFAULT_SCHEDULER_URL; }
  }
  function setBaseUrl(url) {
    try { localStorage.setItem(KEY, String(url || "").trim().replace(/\/$/, "")); } catch {}
  }
  function isConfigured() { return !!getBaseUrl(); }

  async function call(action, extra = {}) {
    const base = getBaseUrl();
    if (!base) throw new Error("No auto-continue scheduler URL set — add it under Settings.");
    const token = typeof AuthClient !== "undefined" ? AuthClient.getToken() : null;
    let resp;
    try {
      resp = await fetch(`${base}/jobs`, {
        method: "POST",
        headers: { "Content-Type": "application/json", ...(token ? { Authorization: `Bearer ${token}` } : {}) },
        body: JSON.stringify({ action, ...extra }),
      });
    } catch (e) {
      throw new Error(`Couldn't reach the auto-continue scheduler (${base}) — is it deployed? ${e.message}`);
    }
    const data = await resp.json().catch(() => ({}));
    if (!resp.ok) throw new Error(data.error || `Scheduler error ${resp.status}`);
    return data;
  }

  return {
    getBaseUrl, setBaseUrl, isConfigured,
    list: () => call("list").then(d => d.jobs || []),
    create: (job) => call("create", { job }).then(d => d.job),
    act: (action, jobId) => call(action, { jobId }).then(d => d.job),
  };
})();
