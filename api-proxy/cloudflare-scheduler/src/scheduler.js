/* =====================================================================
   Auto-continue scheduler — the part that actually extracts.

   Runs every hour (Cron Trigger in wrangler.toml), and straight away
   when someone queues or clicks "Run now" on an auto-continue job.
   For each active job in the Sheet's "Jobs" tab it:

     1. works out what can be fetched now — from the day after the job's
        cursor up to the earlier of its end date and the latest COMPLETE
        local day (yesterday, once it's past SETTLE_HOUR local time);
     2. fetches it in API-call-sized spans, stopping at the brand's daily
        call quota (shared by every job on the same company connection);
     3. appends the rows to the Readings tab in exactly the format the
        browser extraction uses, then advances the job's cursor;
     4. marks the job done once its end date is saved — otherwise it just
        waits for the next run (tomorrow for quota, or the next day's data).

   Example: queued 3 Nov for 1 Oct → 31 Dec, FusionSolar Hourly
   (25 calls/day): 3 Nov fetches 1–25 Oct, 4 Nov fetches 26 Oct – 3 Nov
   (caught up), then one new day each day until 31 Dec is saved on 1 Jan.
   ===================================================================== */

import { sheets, safeJson, OutOfBudget } from "./google.js";
import { planJob, buildSpans, addDays, localNow, monthStart, monthEnd, pointKey, fmtDayLabel, pad } from "./plan.js";

export const QUOTAS = {
  // Keep in step with RATE_LIMIT_DEFAULTS in js/brands.js.
  fusionsolar: { perDay: 25, delayMs: 2000 },
  solaredge:   { perDay: 300, delayMs: 400 },
};
export const BRAND_LABEL = { fusionsolar: "FusionSolar", solaredge: "SolarEdge" };
export const SUPPORTED_BRANDS = Object.keys(QUOTAS);

// Network failures and 5xx responses: worth retrying at the next run rather than stopping the job.
class TransientError extends Error {}
function transientIfServerError(resp, what) {
  if (resp.status >= 500) throw new TransientError(`${what} returned ${resp.status}`);
}

export class QuotaHit extends Error {
  constructor(msg = "The brand's daily API limit was reached.") { super(msg); this.name = "QuotaHit"; }
}

/* ---------------------------- Jobs tab ---------------------------- */

export const JOBS_SHEET = "Jobs";
export const JOB_COLUMNS = ["JobId", "ConnectionId", "Company", "Brand", "Resolution", "StationIdsJson", "StartDate", "EndDate",
  "SkipJson", "Cursor", "Status", "QuotaDate", "CallsToday", "Note", "LastRunAt", "RowsSaved", "CreatedBy", "CreatedAt",
  "UpdatedAt", "LastError"];
const LAST_COL = String.fromCharCode(64 + JOB_COLUMNS.length); // "T"

function rowToJob(r, rowIndex) {
  return {
    rowIndex, jobId: r[0] || "", connectionId: r[1] || "", company: r[2] || "", brand: r[3] || "", resolution: r[4] || "",
    stationIds: safeJson(r[5], []), startDate: r[6] || "", endDate: r[7] || "", skip: safeJson(r[8], []),
    cursor: r[9] || "", status: r[10] || "active", quotaDate: r[11] || "", callsToday: Number(r[12]) || 0,
    note: r[13] || "", lastRunAt: r[14] || "", rowsSaved: Number(r[15]) || 0, createdBy: r[16] || "",
    createdAt: r[17] || "", updatedAt: r[18] || "", lastError: r[19] || "",
  };
}
function jobToRow(j) {
  return [j.jobId, j.connectionId, j.company, j.brand, j.resolution, JSON.stringify(j.stationIds || []), j.startDate, j.endDate,
    JSON.stringify(j.skip || []), j.cursor || "", j.status, j.quotaDate || "", j.callsToday || 0, j.note || "", j.lastRunAt || "",
    j.rowsSaved || 0, j.createdBy || "", j.createdAt || "", j.updatedAt || "", j.lastError || ""];
}

/** Reads the Jobs tab, creating it (with a header row) the first time. */
export async function readJobs(rc) {
  try {
    const data = await sheets(rc, `/values/${JOBS_SHEET}!A2:${LAST_COL}5000`);
    return (data.values || []).map((r, i) => rowToJob(r, i + 2)).filter(j => j.jobId);
  } catch (err) {
    if (err.status !== 400 || !/Unable to parse range/i.test(err.message)) throw err;
    await sheets(rc, `:batchUpdate`, { method: "POST", body: JSON.stringify({ requests: [{ addSheet: { properties: { title: JOBS_SHEET } } }] }) });
    await sheets(rc, `/values/${JOBS_SHEET}!A1:${LAST_COL}1?valueInputOption=RAW`, { method: "PUT", body: JSON.stringify({ values: [JOB_COLUMNS] }) });
    return [];
  }
}

export async function appendJob(rc, job) {
  const res = await sheets(rc, `/values/${JOBS_SHEET}!A:${LAST_COL}:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`, {
    method: "POST", body: JSON.stringify({ values: [jobToRow(job)] }),
  });
  const m = String(res.updates?.updatedRange || "").match(/!A(\d+)/);
  job.rowIndex = m ? Number(m[1]) : null;
  return job;
}

export async function writeJob(rc, job) {
  job.updatedAt = new Date(rc.now()).toISOString();
  await sheets(rc, `/values/${JOBS_SHEET}!A${job.rowIndex}:${LAST_COL}${job.rowIndex}?valueInputOption=RAW`, {
    method: "PUT", body: JSON.stringify({ values: [jobToRow(job)] }),
  });
}

async function writeJobsBatch(rc, jobs) {
  if (!jobs.length) return;
  const stamp = new Date(rc.now()).toISOString();
  jobs.forEach(j => { j.updatedAt = stamp; });
  await sheets(rc, `/values:batchUpdate`, {
    method: "POST",
    body: JSON.stringify({ valueInputOption: "RAW", data: jobs.map(j => ({ range: `${JOBS_SHEET}!A${j.rowIndex}:${LAST_COL}${j.rowIndex}`, values: [jobToRow(j)] })) }),
  });
}

/* ---------------------------- Connections / Readings / Audit ---------------------------- */

// Same column layout as the CORS proxy's Connections tab:
// RowId, Title, Brand, Region, CredentialsJson, StationsJson, DailyAutoExtract, CursorJson, TemplateSettingsJson, ...
// Read-only here — the scheduler never writes to Connections.
export async function readConnections(rc) {
  const data = await sheets(rc, `/values/Connections!A2:J1000`);
  return (data.values || [])
    .map(r => ({ id: r[0], companyName: r[1], brand: r[2], credentials: safeJson(r[4], {}), stations: safeJson(r[5], []), templateSettings: safeJson(r[8], null) }))
    .filter(c => c.id && c.companyName);
}

const READINGS_CHUNK = 10000;
function readingsAppendCost(n) { return Math.ceil(n / READINGS_CHUNK); }

async function appendReadings(rc, rows) {
  const runAt = new Date(rc.now()).toISOString();
  const values = rows.map(r => [r.timestamp, r.company, r.brand, r.stationId, r.stationName, r.resolution, r.kwh, runAt]);
  for (let i = 0; i < values.length; i += READINGS_CHUNK) {
    await sheets(rc, `/values/Readings!A:H:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`, {
      method: "POST", body: JSON.stringify({ values: values.slice(i, i + READINGS_CHUNK) }),
    });
  }
}

// Audit tab: Timestamp | User | Role | Action | Company | Details (same as the web app writes).
async function appendAudit(rc, entries) {
  if (!entries.length) return;
  await sheets(rc, `/values/Audit!A:F:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`, {
    method: "POST", body: JSON.stringify({ values: entries.map(e => [new Date(rc.now()).toISOString(), "Auto-continue scheduler", "System", e.action, e.company, e.details]) }),
  }).catch(() => {}); // auditing must never break extraction
}

/* ---------------------------- brand calls (mirror js/brands.js) ---------------------------- */

const FUSION_REGIONS = {
  intl: "https://intl.fusionsolar.huawei.com/thirdData",
  "ap-sg5": "https://sg5.fusionsolar.huawei.com/thirdData",
  eu5: "https://eu5.fusionsolar.huawei.com/thirdData",
};
const FUSION_ENDPOINT = { Hourly: "getKpiStationHour", Daily: "getKpiStationDay", Monthly: "getKpiStationMonth" };
const FUSION_BATCH = 100; // FusionSolar accepts up to 100 stationCodes per call

async function readJsonLoose(resp) {
  const text = await resp.text();
  try { return JSON.parse(text); } catch { return { _raw: text }; }
}

async function fusionLogin(rc, conn) {
  const c = conn.credentials || {};
  const base = (c.baseUrl || "").trim() || FUSION_REGIONS[c.region] || FUSION_REGIONS.intl;
  const resp = await rc.fetch(`${base}/login`, {
    method: "POST", headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ userName: c.username, systemCode: c.systemCode }),
  });
  transientIfServerError(resp, "FusionSolar login");
  const data = await readJsonLoose(resp);
  if (data.failCode === 407) throw new QuotaHit("FusionSolar refused the login: access frequency limit reached.");
  if (!data.success) throw new Error(`FusionSolar login failed: ${data.message || `failCode ${data.failCode ?? "?"}`}`);
  const token = data.data?.xsrfToken || data.xsrfToken || resp.headers.get("xsrf-token");
  if (!token) throw new Error("FusionSolar login succeeded but returned no XSRF-TOKEN.");
  return { base, headers: { "Content-Type": "application/json", "XSRF-TOKEN": token } };
}

// Local noon of a local date, as a real epoch — the collectTime FusionSolar expects.
function localNoonMs(iso, utcOffset) {
  const [y, m, d] = iso.split("-").map(Number);
  return Date.UTC(y, m - 1, d, 12) - utcOffset * 3600000;
}
// Wall clock of an epoch in the site's local time — same as TemplateExport.wallClockFromRow for FusionSolar.
function localStamp(epochMs, utcOffset) {
  const d = new Date(Number(epochMs) + utcOffset * 3600000);
  return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())} ${pad(d.getUTCHours())}:${pad(d.getUTCMinutes())}:${pad(d.getUTCSeconds())}`;
}

function inSpan(resolution, localDay, s, e) {
  return resolution === "Monthly"
    ? localDay.slice(0, 7) >= s.slice(0, 7) && localDay.slice(0, 7) <= e.slice(0, 7)
    : localDay >= s && localDay <= e;
}

async function fetchFusionSpan(rc, auth, conn, job, [s, e], utcOffset, counter, skip) {
  const endpoint = FUSION_ENDPOINT[job.resolution];
  const anchor = job.resolution === "Hourly" ? s : job.resolution === "Daily" ? `${s.slice(0, 7)}-15` : `${s.slice(0, 4)}-07-01`;
  const collectTime = localNoonMs(anchor, utcOffset);
  const rows = [];
  for (let i = 0; i < job.stationIds.length; i += FUSION_BATCH) {
    if (i > 0) await rc.sleep(QUOTAS.fusionsolar.delayMs);
    const resp = await rc.fetch(`${auth.base}/${endpoint}`, {
      method: "POST", headers: auth.headers,
      body: JSON.stringify({ stationCodes: job.stationIds.slice(i, i + FUSION_BATCH).join(","), collectTime }),
    });
    counter.calls++;
    transientIfServerError(resp, `FusionSolar ${endpoint}`);
    const data = await readJsonLoose(resp);
    if (data.failCode === 407) throw new QuotaHit();
    if (!data.success) throw new Error(`FusionSolar ${endpoint} failed: ${data.message || `failCode ${data.failCode ?? resp.status}`}`);
    for (const row of data.data || []) {
      const ts = localStamp(row.collectTime, utcOffset);
      const day = ts.slice(0, 10);
      if (!inSpan(job.resolution, day, s, e) || skip.has(pointKey(job.resolution, day))) continue;
      rows.push({ stationId: row.stationCode, timestamp: ts,
        kwh: Number(row.dataItemMap?.inverter_power ?? row.dataItemMap?.product_power ?? row.dataItemMap?.PVYield ?? 0) });
    }
  }
  return rows;
}

async function fetchSolarEdgeSpan(rc, conn, job, [s, e], counter, skip) {
  const c = conn.credentials || {};
  const base = (c.baseUrl || "").trim() || "https://monitoringapi.solaredge.com";
  const unit = { Hourly: "HOUR", Daily: "DAY", Monthly: "MONTH" }[job.resolution];
  const resp = await rc.fetch(`${base}/site/${encodeURIComponent(c.siteId)}/energyDetails?timeUnit=${unit}&meters=PRODUCTION`
    + `&startTime=${s}%2000:00:00&endTime=${e}%2023:59:59&api_key=${encodeURIComponent(c.apiKey)}`);
  counter.calls++;
  if (resp.status === 429) throw new QuotaHit();
  transientIfServerError(resp, "SolarEdge energyDetails");
  const data = await readJsonLoose(resp);
  if (!resp.ok) throw new Error(`SolarEdge energyDetails ${resp.status}: ${String(data.String || data.message || data._raw || "").slice(0, 200)}`);
  const rows = [];
  for (const m of data.energyDetails?.meters || []) {
    for (const v of m.values || []) {
      const day = String(v.date).slice(0, 10);
      if (!inSpan(job.resolution, day, s, e) || skip.has(pointKey(job.resolution, day))) continue;
      rows.push({ stationId: c.siteId, timestamp: v.date, kwh: (v.value || 0) / 1000 });
    }
  }
  return rows;
}

function pointsInSpan(resolution, s, e) {
  const pts = [];
  if (resolution === "Monthly") {
    for (let p = monthStart(s); p <= e; p = addDays(monthEnd(p), 1)) pts.push(p);
  } else {
    for (let d = s; d <= e; d = addDays(d, 1)) pts.push(d);
  }
  return pts;
}

/**
 * Fetches as much of plan.from → plan.availEnd as the quota, subrequest budget
 * and time allow. Never throws for brand/quota/budget problems — returns how
 * far it got so whatever was fetched is still saved.
 */
async function extractRange(rc, conn, job, plan, callsAllowed, utcOffset) {
  const spans = buildSpans(job.brand, job.resolution, plan.from, plan.availEnd);
  const skip = new Set(job.skip || []);
  const out = { rows: [], cursor: job.cursor, stop: null, quotaHit: false, error: null };
  const counter = { calls: 0 };
  const callsPerSpan = job.brand === "fusionsolar" ? Math.ceil(job.stationIds.length / FUSION_BATCH) : 1;
  const delay = QUOTAS[job.brand].delayMs;
  let auth = job.brand === "fusionsolar" ? null : {};

  for (const span of spans) {
    if (pointsInSpan(job.resolution, span[0], span[1]).every(p => skip.has(p))) { out.cursor = span[1]; continue; }
    if (counter.calls + callsPerSpan > callsAllowed) { out.stop = "quota"; break; }
    // Keep enough subrequests back to save the rows + the job (and audit).
    const reserve = readingsAppendCost(out.rows.length + 5000) + 3;
    if (rc.left() < callsPerSpan + (auth ? 0 : 1) + reserve) { out.stop = "budget"; break; }
    if (rc.timeLeft() < 8000 + callsPerSpan * delay) { out.stop = "time"; break; }
    try {
      if (!auth) auth = await fusionLogin(rc, conn);
      const rows = job.brand === "fusionsolar"
        ? await fetchFusionSpan(rc, auth, conn, job, span, utcOffset, counter, skip)
        : await fetchSolarEdgeSpan(rc, conn, job, span, counter, skip);
      const names = Object.fromEntries((conn.stations || []).map(st => [String(st.id), st.name]));
      for (const r of rows) {
        out.rows.push({ timestamp: r.timestamp, company: conn.companyName, brand: BRAND_LABEL[job.brand], stationId: r.stationId,
          stationName: names[String(r.stationId)] || r.stationId, resolution: job.resolution, kwh: r.kwh });
      }
      out.cursor = span[1];
    } catch (err) {
      if (err instanceof QuotaHit) { out.quotaHit = true; out.stop = "quota"; }
      else if (err instanceof OutOfBudget) { out.stop = "budget"; }
      else if (err instanceof TransientError || err instanceof TypeError) { out.transient = err.message; }
      else { out.error = err.message; }
      break;
    }
    await rc.sleep(delay);
  }
  out.calls = counter.calls;
  return out;
}

/* ---------------------------- the run ---------------------------- */

function waitingNote(plan, resolution) {
  const unit = resolution === "Monthly" ? "month" : "day";
  const savedThrough = addDays(plan.from, -1);
  return `Up to date through ${fmtDayLabel(savedThrough)} — next ${unit} becomes available ${fmtDayLabel(plan.availableOn)} (after ${pad(plan.settleHour)}:00 site time) and is fetched automatically.`;
}

function companyCallsToday(jobs, connectionId, today) {
  return jobs.filter(j => j.connectionId === connectionId && j.quotaDate === today).reduce((n, j) => n + (j.callsToday || 0), 0);
}

/**
 * opts.onlyJobId — process just that job (queue / "Run now").
 * Returns a short summary (also handy for the test harness).
 */
export async function runScheduler(rc, opts = {}) {
  const settleHour = Number(rc.env.SETTLE_HOUR ?? 2);
  const summary = { processed: [], skipped: [] };
  const jobs = await readJobs(rc);
  const candidates = jobs.filter(j => j.status === "active" && (!opts.onlyJobId || j.jobId === opts.onlyJobId));
  if (!candidates.length) return summary;
  const conns = await readConnections(rc);
  const dirty = new Set();
  const audit = [];
  const setNote = (j, note) => { if (j.note !== note) { j.note = note; dirty.add(j); } };

  for (const job of candidates) {
    const conn = conns.find(c => c.id === job.connectionId);
    if (!conn || conn.brand !== job.brand) {
      job.status = "error";
      job.lastError = conn ? `Company's brand changed (${conn.brand}).` : "Company connection not found — was it removed from the Dashboard?";
      setNote(job, "Stopped — see error."); dirty.add(job);
      continue;
    }
    const utcOffset = Number(conn.templateSettings?.utcOffset ?? 8);
    const nowMs = rc.now();
    const today = localNow(nowMs, utcOffset).today;
    let plan = planJob(job, nowMs, utcOffset, settleHour);

    if (plan.state === "done") {
      job.status = "done"; setNote(job, `Complete — ${job.rowsSaved} row(s) saved.`); dirty.add(job);
      audit.push({ action: "Auto-continue extraction complete", company: job.company, details: `${job.resolution} · ${job.startDate} → ${job.endDate}` });
      continue;
    }
    if (plan.state === "waiting") { setNote(job, waitingNote(plan, job.resolution)); summary.skipped.push(job.jobId); continue; }

    const quota = QUOTAS[job.brand];
    if (!quota) { job.status = "error"; job.lastError = `Auto-continue doesn't support ${job.brand} yet.`; dirty.add(job); continue; }
    const callsLeft = quota.perDay - companyCallsToday(jobs, job.connectionId, today);
    if (callsLeft <= 0) {
      setNote(job, `Daily API limit (${quota.perDay} calls) used for today — continues tomorrow (${fmtDayLabel(addDays(today, 1))}).`);
      summary.skipped.push(job.jobId);
      continue;
    }
    if (rc.left() < 6 || rc.timeLeft() < 10000) { setNote(job, "Queued — continues at the next hourly run."); break; }

    const res = await extractRange(rc, conn, job, plan, callsLeft, utcOffset);

    // Save rows first; only then move the cursor (so a failed save is simply refetched next time).
    let saved = true;
    if (res.rows.length) {
      try { await appendReadings(rc, res.rows); }
      catch (err) { saved = false; res.transient = `Fetched ${res.rows.length} row(s) but couldn't save them to Readings (${err.message}) — they'll be fetched again`; }
    }
    if (saved) { job.cursor = res.cursor; job.rowsSaved += res.rows.length; }

    if (job.quotaDate !== today) { job.quotaDate = today; job.callsToday = 0; }
    job.callsToday += res.calls;
    if (res.quotaHit) job.callsToday += Math.max(0, callsLeft - res.calls); // brand says we're out — treat today as used up
    job.lastRunAt = new Date(nowMs).toISOString();

    plan = planJob(job, rc.now(), utcOffset, settleHour);
    if (res.error) {
      job.status = "error"; job.lastError = res.error; job.note = "Stopped — see error. Press Retry once it's fixed.";
    } else if (res.transient) {
      job.lastError = ""; job.note = `Temporary problem: ${res.transient}. Retrying at the next hourly run.`;
    } else if (plan.state === "done") {
      job.status = "done"; job.lastError = ""; job.note = `Complete — ${job.rowsSaved} row(s) saved.`;
    } else if (plan.state === "waiting") {
      job.lastError = ""; job.note = waitingNote(plan, job.resolution);
    } else if (res.stop === "quota") {
      job.lastError = ""; job.note = `Daily API limit (${quota.perDay} calls) used for today — saved through ${fmtDayLabel(job.cursor || addDays(job.startDate, -1))}, continues tomorrow (${fmtDayLabel(addDays(today, 1))}).`;
    } else {
      job.lastError = ""; job.note = `Saved through ${fmtDayLabel(job.cursor || addDays(job.startDate, -1))} — continues at the next hourly run.`;
    }

    try { await writeJob(rc, job); dirty.delete(job); } catch { dirty.add(job); }
    audit.push({ action: res.error ? "Auto-continue extraction error" : "Auto-continue extraction",
      company: job.company, details: `${job.resolution} · ${job.startDate} → ${job.endDate} · ${res.rows.length} row(s) · ${res.calls} call(s) · saved through ${job.cursor || "nothing yet"}${res.error ? ` · ${res.error}` : ""}` });
    if (plan.state === "done") audit.push({ action: "Auto-continue extraction complete", company: job.company, details: `${job.resolution} · ${job.startDate} → ${job.endDate} · ${job.rowsSaved} row(s) total` });
    summary.processed.push({ jobId: job.jobId, rows: res.rows.length, calls: res.calls, cursor: job.cursor, stop: res.stop, error: res.error });
    if (res.stop === "budget" || res.stop === "time") break;
  }

  if (dirty.size && rc.left() >= 1) await writeJobsBatch(rc, [...dirty]).catch(() => {});
  if (audit.length && rc.left() >= 1) await appendAudit(rc, audit);
  return summary;
}
