/* =====================================================================
   SolarLink auto-continue scheduler — Cloudflare Worker
   ---------------------------------------------------------------------
   A separate Worker from the CORS proxy (solarlink-cors-proxy), so the
   proxy you already have deployed doesn't need to change.

   scheduled()  — Cron Trigger (hourly, see wrangler.toml): continues
                  every active job in the Sheet's "Jobs" tab.
   POST /jobs   — the web app's Extraction tab: list / create / pause /
                  resume / cancel / runNow. Same access control as the
                  proxy: a Google ID token from "Sign in with Google",
                  checked against Google and the Sheet's Users tab.

   Secrets / vars (same values as the CORS proxy — see README.md):
     GOOGLE_SERVICE_ACCOUNT_EMAIL, GOOGLE_SERVICE_ACCOUNT_KEY   (secrets)
     GOOGLE_SPREADSHEET_ID, GOOGLE_SIGNIN_CLIENT_ID              (vars)
     MAX_SUBREQUESTS  (var, default 45 — Workers Free allows 50 per run;
                       raise to ~900 on Workers Paid)
     SETTLE_HOUR      (var, default 2 — local hour after which yesterday
                       counts as complete)
   ===================================================================== */

import { makeRunContext, sheets } from "./google.js";
import { runScheduler, readJobs, appendJob, writeJob, readConnections, SUPPORTED_BRANDS } from "./scheduler.js";
import { monthStart } from "./plan.js";

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type, Authorization",
};
const ISO_DATE = /^\d{4}-\d{2}-\d{2}$/;
const RESOLUTIONS = ["Hourly", "Daily", "Monthly"];

export default {
  async scheduled(event, env, ctx) {
    const rc = makeRunContext(env, { deadlineMs: Date.now() + 12 * 60 * 1000 }); // cron runs may last up to 15 min
    ctx.waitUntil(runScheduler(rc).then(s => console.log("scheduler run", JSON.stringify(s)))
      .catch(err => console.error("scheduler run failed:", err.stack || err.message)));
  },

  async fetch(request, env, ctx) {
    if (request.method === "OPTIONS") return new Response(null, { status: 204, headers: CORS_HEADERS });
    const { pathname } = new URL(request.url);
    if (pathname !== "/jobs" || request.method !== "POST") return json(404, { error: "Not found. POST to /jobs." });
    for (const v of ["GOOGLE_SERVICE_ACCOUNT_EMAIL", "GOOGLE_SERVICE_ACCOUNT_KEY", "GOOGLE_SPREADSHEET_ID", "GOOGLE_SIGNIN_CLIENT_ID"]) {
      if (!env[v]) return json(500, { error: `Scheduler Worker is missing ${v}. See api-proxy/cloudflare-scheduler/README.md.` });
    }

    // waitUntil work may run ~30s after the response, sharing this invocation's subrequest limit.
    const rc = makeRunContext(env, { deadlineMs: Date.now() + 27000 });
    const who = await authorize(request, rc);
    if (who instanceof Response) return who;

    let payload;
    try { payload = await request.json(); } catch { return json(400, { error: "Invalid JSON body." }); }

    try {
      switch (payload.action) {
        case "list": return json(200, { jobs: (await readJobs(rc)).map(publicJob) });
        case "create": return await createJob(rc, ctx, who, payload.job || {});
        case "pause": case "resume": case "cancel": case "runNow":
          return await changeJob(rc, ctx, who, payload.action, payload.jobId);
        default: return json(400, { error: `Unknown action '${payload.action}'.` });
      }
    } catch (err) {
      return json(502, { error: err.message });
    }
  },
};

/* ---------------------------- actions ---------------------------- */

async function createJob(rc, ctx, who, input) {
  const { connectionId, resolution, startDate, endDate } = input;
  const stationIds = Array.isArray(input.stationIds) ? input.stationIds.map(String) : [];
  const skip = Array.isArray(input.skip) ? input.skip.map(String).filter(s => ISO_DATE.test(s)).slice(0, 5000) : [];
  if (!RESOLUTIONS.includes(resolution)) return json(400, { error: "Resolution must be Hourly, Daily or Monthly." });
  if (!ISO_DATE.test(startDate || "") || !ISO_DATE.test(endDate || "")) return json(400, { error: "Start and end dates must be YYYY-MM-DD." });
  if (endDate < startDate) return json(400, { error: "End date is before start date." });
  if (!stationIds.length) return json(400, { error: "Pick at least one station." });

  const conns = await readConnections(rc);
  const conn = conns.find(c => c.id === connectionId);
  if (!conn) return json(404, { error: "Company connection not found." });
  if (!SUPPORTED_BRANDS.includes(conn.brand)) {
    return json(400, { error: `Auto-continue currently supports ${SUPPORTED_BRANDS.join(" and ")} only — ${conn.brand}'s historical-data endpoint isn't confirmed yet.` });
  }

  const jobs = await readJobs(rc);
  const start = resolution === "Monthly" ? monthStart(startDate) : startDate;
  const clash = jobs.find(j => j.connectionId === connectionId && j.resolution === resolution
    && ["active", "paused", "error"].includes(j.status) && j.startDate <= endDate && j.endDate >= start
    && j.stationIds.some(s => stationIds.includes(s)));
  if (clash) {
    return json(409, { error: `There's already an auto-continue ${resolution} job for ${conn.companyName} covering ${clash.startDate} → ${clash.endDate} (${clash.status}). Cancel it first, or pick a range that doesn't overlap.` });
  }

  const nowIso = new Date().toISOString();
  const job = await appendJob(rc, {
    jobId: crypto.randomUUID(), connectionId, company: conn.companyName, brand: conn.brand, resolution, stationIds,
    startDate: start, endDate, skip, cursor: "", status: "active", quotaDate: "", callsToday: 0,
    note: "Queued — first batch starting now.", lastRunAt: "", rowsSaved: 0, createdBy: who.email, createdAt: nowIso, updatedAt: nowIso, lastError: "",
  });
  ctx.waitUntil(runScheduler(rc, { onlyJobId: job.jobId }).catch(err => console.error("first run failed:", err.message)));
  return json(200, { job: publicJob(job) });
}

async function changeJob(rc, ctx, who, action, jobId) {
  const jobs = await readJobs(rc);
  const job = jobs.find(j => j.jobId === jobId);
  if (!job) return json(404, { error: "Job not found." });
  if (job.status === "done" || job.status === "cancelled") return json(400, { error: `This job is already ${job.status}.` });

  if (action === "pause") { job.status = "paused"; job.note = `Paused by ${who.email}.`; }
  if (action === "cancel") { job.status = "cancelled"; job.note = `Cancelled by ${who.email}. Readings already saved are kept.`; }
  if (action === "resume" || action === "runNow") {
    job.status = "active"; job.lastError = ""; job.note = "Running now…";
  }
  await writeJob(rc, job);
  if (job.status === "active") {
    ctx.waitUntil(runScheduler(rc, { onlyJobId: job.jobId }).catch(err => console.error("run failed:", err.message)));
  }
  return json(200, { job: publicJob(job) });
}

function publicJob(j) {
  const { rowIndex, skip, ...rest } = j;
  return { ...rest, skippedCount: (skip || []).length };
}

/* ---------------------------- access control (same rules as the CORS proxy) ---------------------------- */

async function authorize(request, rc) {
  const env = rc.env;
  const header = request.headers.get("Authorization") || "";
  const idToken = header.startsWith("Bearer ") ? header.slice(7) : null;
  if (!idToken) return json(401, { error: "Not signed in.", code: "NO_TOKEN" });

  let info;
  try {
    const resp = await rc.fetch(`https://oauth2.googleapis.com/tokeninfo?id_token=${encodeURIComponent(idToken)}`);
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
    const data = await sheets(rc, `/values/Users!A2:B1000`);
    const match = (data.values || []).find(r => (r[0] || "").trim().toLowerCase() === email);
    role = match ? (match[1] || "User").trim() : null;
  } catch (err) {
    return json(502, { error: `Could not check the Users list: ${err.message}` });
  }
  if (!role) return json(403, { error: `${email} is not on the approved users list.`, code: "NOT_AUTHORIZED" });
  return { email, role };
}

function json(status, obj) {
  return new Response(JSON.stringify(obj), { status, headers: { "Content-Type": "application/json", ...CORS_HEADERS } });
}
