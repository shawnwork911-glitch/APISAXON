/* =====================================================================
   Pure date logic for auto-continue jobs — no I/O, so it can be tested
   on its own (see ../test/simulate.mjs).

   All dates are plain "YYYY-MM-DD" strings in the SITE's local time
   (templateSettings.utcOffset, default +8 / Singapore), handled with UTC
   arithmetic only so the server's own timezone never matters.

   A job's `cursor` is the last LOCAL day already saved to Readings
   (for Monthly: the last day of the last month saved).
   ===================================================================== */

export function pad(n) { return String(n).padStart(2, "0"); }
export function parseISO(s) { return String(s).slice(0, 10).split("-").map(Number); }
export function isoFromMs(ms) {
  const d = new Date(ms);
  return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
}
export function addDays(iso, n) {
  const [y, m, d] = parseISO(iso);
  return isoFromMs(Date.UTC(y, m - 1, d) + n * 86400000);
}
export function monthStart(iso) { return `${String(iso).slice(0, 7)}-01`; }
export function monthEnd(iso) {
  const [y, m] = parseISO(iso);
  return `${y}-${pad(m)}-${pad(new Date(Date.UTC(y, m, 0)).getUTCDate())}`;
}
export function yearEnd(iso) { return `${String(iso).slice(0, 4)}-12-31`; }
export const minISO = (a, b) => (a < b ? a : b);
export const maxISO = (a, b) => (a > b ? a : b);

/** Local wall-clock "today" and hour for a UTC offset (hours). */
export function localNow(nowMs, utcOffsetHours) {
  const shifted = nowMs + utcOffsetHours * 3600000;
  return { today: isoFromMs(shifted), hour: new Date(shifted).getUTCHours() };
}

/**
 * The latest local day whose data is complete. A day counts as complete
 * once local time has passed `settleHour` on the following day — this
 * gives the vendor portal a little time to finalise yesterday's figures.
 */
export function latestCompleteDay(nowMs, utcOffsetHours, settleHour = 2) {
  const { today, hour } = localNow(nowMs, utcOffsetHours);
  return addDays(today, hour >= settleHour ? -1 : -2);
}

/**
 * Where the job stands right now.
 *   state "done"    — everything up to the end date is saved
 *   state "waiting" — caught up; the next day/month isn't complete yet
 *                     (availableOn = local date it becomes fetchable)
 *   state "ready"   — [from, availEnd] can be fetched now
 */
export function planJob(job, nowMs, utcOffsetHours = 8, settleHour = 2) {
  const monthly = job.resolution === "Monthly";
  const start = monthly ? monthStart(job.startDate) : job.startDate;
  const finalEnd = monthly ? monthEnd(job.endDate) : job.endDate;
  const from = job.cursor && job.cursor >= start ? addDays(job.cursor, 1) : start;
  if (from > finalEnd) return { state: "done", from, finalEnd };

  const complete = latestCompleteDay(nowMs, utcOffsetHours, settleHour);
  // Monthly: only whole months that have fully ended are fetched.
  const completeThrough = monthly
    ? (monthEnd(complete) === complete ? complete : addDays(monthStart(complete), -1))
    : complete;
  const availEnd = minISO(finalEnd, completeThrough);

  if (from > availEnd) {
    const lastNeeded = monthly ? monthEnd(from) : from;
    return { state: "waiting", from, finalEnd, availableOn: addDays(lastNeeded, 1), settleHour };
  }
  return { state: "ready", from, availEnd, finalEnd };
}

/**
 * Splits [from, to] into the API calls a brand needs, in order. Each span is
 * one call per station batch; the cursor only ever advances to a span's end
 * after that whole span has been fetched.
 */
export function buildSpans(brand, resolution, from, to) {
  const spans = [];
  if (brand === "fusionsolar") {
    // getKpiStationHour = one day per call, getKpiStationDay = one month,
    // getKpiStationMonth = one year (see js/brands.js).
    let s = from;
    while (s <= to) {
      const unitEnd = resolution === "Hourly" ? s : resolution === "Daily" ? monthEnd(s) : yearEnd(s);
      const e = minISO(unitEnd, to);
      spans.push([s, e]);
      s = addDays(e, 1);
    }
    return spans;
  }
  if (brand === "solaredge") {
    // energyDetails windows: HOUR max 30 days, DAY max 365, MONTH unlimited.
    const maxDays = { Hourly: 30, Daily: 365, Monthly: null }[resolution];
    if (!maxDays) return [[from, to]];
    let s = from;
    while (s <= to) {
      const e = minISO(addDays(s, maxDays - 1), to);
      spans.push([s, e]);
      s = addDays(e, 1);
    }
    return spans;
  }
  throw new Error(`Auto-continue isn't available for ${brand} yet — its historical-data endpoint isn't confirmed (see js/brands.js).`);
}

/** The point key the browser's extraction lock uses: a day, or the 1st of a month. */
export function pointKey(resolution, localDay) {
  return resolution === "Monthly" ? monthStart(localDay) : localDay;
}

export function fmtDayLabel(iso) {
  const [y, m, d] = parseISO(iso);
  const mon = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"][m - 1];
  return `${d} ${mon} ${y}`;
}
