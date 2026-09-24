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
  function buildCursorPoints(brandKey, resolution, startDate, endDate) {
    const start = new Date(startDate), end = new Date(endDate);
    const points = [];
    const step = resolution === "Monthly" ? "month" : "day";
    for (let d = new Date(start); d <= end; ) {
      points.push(fmt(d));
      if (step === "month") d.setMonth(d.getMonth() + 1); else d.setDate(d.getDate() + 1);
    }
    return points;
    function fmt(d) { return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`; }
  }

  function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

  return { createJob, getJob, run, pause, stop };
})();
