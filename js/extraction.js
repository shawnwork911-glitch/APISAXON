/* =====================================================================
   Extraction engine
   ---------------------------------------------------------------------
   Mirrors the resumable, quota-aware design of fusionsolar_bot.py /
   solaredge_bot.py: a job tracks how much of a date range has been
   pulled, paces calls, stops cleanly at the daily quota, and can resume
   later (next click of "Run now", or the next scheduled GitHub Action
   run) picking up where it left off.
   ===================================================================== */

const ExtractionEngine = (() => {
  const jobs = new Map(); // jobId -> job state
  let idCounter = 1;

  function createJob({ connection, brand, resolution, startDate, endDate, stationIds, onProgress }) {
    const id = `job-${idCounter++}`;
    const job = {
      id, connection, brand, resolution, startDate, endDate, stationIds,
      status: "queued", // queued | running | paused | stopped | done | error
      callsMadeToday: 0,
      rowsCollected: [],
      error: null,
      onProgress: onProgress || (() => {}),
      cursorDate: startDate, // always exactly what was requested — see note below
    };
    jobs.set(id, job);
    return job;
  }
  // NOTE: cursorDate used to silently fall back to connection.cursor[resolution]
  // when one existed, meaning a wider/earlier startDate typed by the user could
  // get silently overridden and never actually fetched. Now the typed startDate
  // is always honored — the cursor is only ever used as a pre-filled suggestion
  // in the UI (see app.js), not a hidden override here.

  function getJob(id) { return jobs.get(id); }
  function pause(id) { const j = jobs.get(id); if (j && j.status === "running") j.status = "paused"; }
  function stop(id) { const j = jobs.get(id); if (j) j.status = "stopped"; }

  async function run(id, ctx) {
    const job = jobs.get(id);
    if (!job) return;
    job.status = "running";
    const brandDef = BRANDS[job.brand];
    const quota = brandDef.quota;

    try {
      if (job.brand === "solaredge") {
        // SolarEdge's fetchSeries already accepts the full date range and
        // chunks internally per brands.js's maxWindowDays — this needs
        // exactly one call covering job.cursorDate → job.endDate. (Bug fixed
        // here: this used to pass the end date as BOTH the start and end,
        // so only the very last day ever actually got requested.)
        if (job.callsMadeToday < quota.perDay) {
          const rows = await brandDef.fetchSeries(ctx, {
            stationIds: job.stationIds, resolution: job.resolution,
            startDate: job.cursorDate, endDate: job.endDate,
          });
          job.rowsCollected.push(...rows);
          job.callsMadeToday += 1;
          job.cursorDate = job.endDate;
          job.connection.cursor = job.connection.cursor || {};
          job.connection.cursor[job.resolution] = job.endDate;
          job.onProgress(job);
        } else {
          job.status = "paused";
          job.pausedReason = "quota";
          job.onProgress(job);
        }
      } else {
        // FusionSolar/etc: iterate date range day-by-day (or per-bucket) so we can
        // check the quota and persist a cursor between every single API call.
        const allCursorPoints = buildCursorPoints(job.brand, job.resolution, job.cursorDate, job.endDate);

        for (const point of allCursorPoints) {
          if (job.status !== "running") break; // paused or stopped
          if (job.callsMadeToday >= quota.perDay) {
            job.status = "paused";
            job.pausedReason = "quota";
            job.onProgress(job);
            break;
          }

          const rows = await brandDef.fetchSeries(ctx, {
            stationIds: job.stationIds,
            resolution: job.resolution,
            startDate: point, endDate: point,
          });
          job.rowsCollected.push(...rows);
          job.callsMadeToday += 1;
          job.cursorDate = point;
          job.connection.cursor = job.connection.cursor || {};
          job.connection.cursor[job.resolution] = point;
          job.onProgress(job);

          if (quota.callDelayMs) await sleep(quota.callDelayMs);
        }
      }

      if (job.status === "running") job.status = "done";
    } catch (err) {
      if (err instanceof RateLimitReached) {
        job.status = "paused";
        job.pausedReason = "quota";
      } else {
        job.status = "error";
        job.error = err.message;
      }
    }
    job.onProgress(job);
    return job;
  }

  // Day/month iteration for brands whose fetchSeries takes one bucket at a
  // time (FusionSolar) — SolarEdge is handled separately above, since its
  // fetchSeries takes the whole range in one call instead.
  //
  // Works on plain "YYYY-MM-DD" strings with UTC arithmetic only. (The old
  // version mixed new Date("YYYY-MM-DD") — parsed as UTC — with local-time
  // getters/setters, which shifted every day back by one in any browser west
  // of UTC, and its month stepping skipped months: 31 Jan + 1 month rolled to
  // 3 Mar, and a mid-month start dropped the end month entirely.) Monthly
  // points are always the 1st of each month, start and end months inclusive.
  function buildCursorPoints(brandKey, resolution, startDate, endDate) {
    const points = [];
    if (resolution === "Monthly") {
      let [y, m] = parseISO(startDate);
      const [ey, em] = parseISO(endDate);
      while (y < ey || (y === ey && m <= em)) {
        points.push(`${y}-${pad(m)}-01`);
        if (++m > 12) { m = 1; y++; }
      }
    } else {
      const [sy, sm, sd] = parseISO(startDate), [ey, em, ed] = parseISO(endDate);
      const endMs = Date.UTC(ey, em - 1, ed);
      for (let ms = Date.UTC(sy, sm - 1, sd); ms <= endMs; ms += 86400000) points.push(fmtUTC(ms));
    }
    return points;
  }

  // Where a job should START when continuing from a saved cursor. The cursor
  // records the last point that was successfully fetched, so resuming starts
  // at the point AFTER it — resuming AT the cursor re-fetched that whole day
  // (or month) again, which is exactly the duplicate boundary days found in
  // the Readings tab. Exception: if the cursor is today (or the current
  // month) or later, that bucket was still in progress when it was fetched,
  // so it's fetched again to pick up the rest of it.
  function nextStartAfter(resolution, cursorDate, today = localTodayISO()) {
    if (!cursorDate) return null;
    const [y, m, d] = parseISO(cursorDate);
    if (resolution === "Monthly") {
      const cursorMonth = `${y}-${pad(m)}`;
      if (cursorMonth >= today.slice(0, 7)) return `${cursorMonth}-01`;
      return m === 12 ? `${y + 1}-01-01` : `${y}-${pad(m + 1)}-01`;
    }
    if (cursorDate >= today) return cursorDate;
    return fmtUTC(Date.UTC(y, m - 1, d) + 86400000);
  }

  function parseISO(s) { return String(s).slice(0, 10).split("-").map(Number); }
  function pad(n) { return String(n).padStart(2, "0"); }
  function fmtUTC(ms) { const d = new Date(ms); return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`; }
  function localTodayISO() { const d = new Date(); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; }

  function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

  return { createJob, getJob, run, pause, stop, nextStartAfter, buildCursorPoints };
})();
