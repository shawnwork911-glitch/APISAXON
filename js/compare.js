/* =====================================================================
   Compare — Hourly vs Monthly
   ---------------------------------------------------------------------
   Port of check_hourly_vs_monthly() from saxon_gen_bot.py: for each
   facility_id, sums Hourly-resolution kWh into monthly buckets and
   compares against the Monthly-resolution rows for that same
   facility+month, flagging any month where they disagree.

   Facility grouping matches TemplateExport exactly (same stationFacility
   mapping, same "unassigned station falls back to its own stationId"
   rule) — a facility with multiple stations has ALL of their Hourly rows
   summed, and ALL of their Monthly rows summed, before comparing.

   Hourly coverage — why it matters:
   A month that has a Monthly figure but NO Hourly rows at all simply
   hasn't been extracted hourly yet (e.g. the hourly backfill starts
   later than the monthly one). That is NOT an outage, so it's reported
   separately as "not extracted" instead of being flagged. A month whose
   Hourly rows cover only some of its days is still compared, but labelled
   "incomplete" so a gap there isn't mistaken for missing generation. Only
   a month with full hourly coverage that sums to ~0 against a real
   Monthly figure is called a FULL OUTAGE.

   Threshold: matches the Python script's shipped defaults (both
   threshold_kwh and threshold_pct set to 0), so the only real filter is
   its materiality floor — a month is flagged once its gap reaches at
   least 0.005 kWh.
   ===================================================================== */

const CompareEngine = (() => {
  const BRAND_LABEL_TO_KEY = { FusionSolar: "fusionsolar", SolarEdge: "solaredge" };

  function wallClock(row, utcOffset) {
    const brandKey = BRAND_LABEL_TO_KEY[row.brand] || String(row.brand || "").toLowerCase();
    let wc;
    try { wc = TemplateExport.wallClockFromRow(row, brandKey, utcOffset); } catch { return null; }
    if (!wc || ![wc.y, wc.mo, wc.d].every(Number.isFinite)) return null;
    return wc;
  }

  const pad = (n) => String(n).padStart(2, "0");
  const daysInMonth = (y, mo) => new Date(Date.UTC(y, mo, 0)).getUTCDate();
  function round2(n) { return Math.round(n * 100) / 100; }

  function localDateOf(iso) {
    if (!iso) return "";
    const d = new Date(iso);
    if (isNaN(d)) return String(iso).slice(0, 10);
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  }
  function todayISO() { return localDateOf(new Date().toISOString()); }

  // ["2026-09-03","2026-09-04","2026-09-05","2026-09-29","2026-09-30"]
  //   -> "2026-09-03 → 2026-09-05, 2026-09-29 → 2026-09-30 (not over yet)"
  function compressDays(dates, today) {
    if (!dates.length) return "";
    const nextDay = (s) => { const [y, m, d] = s.split("-").map(Number); const t = new Date(Date.UTC(y, m - 1, d + 1));
      return `${t.getUTCFullYear()}-${pad(t.getUTCMonth() + 1)}-${pad(t.getUTCDate())}`; };
    const ranges = [];
    let a = dates[0], b = dates[0];
    const push = () => ranges.push((a === b ? a : `${a} → ${b}`) + (b >= today ? " (not over yet)" : ""));
    for (const d of dates.slice(1)) { if (d === nextDay(b) && (d < today) === (b < today)) b = d; else { push(); a = b = d; } } // split past days from today/future
    push();
    return ranges.join(", ");
  }

  /**
   * readings: rows from SheetsClient.listReadings(companyName) — every
   * Resolution mixed together; this filters to Hourly/Monthly itself.
   * conn: the connection object (for its templateSettings — facility
   * mapping and UTC offset).
   * Returns {
   *   flagged: [...],          // months with a real gap
   *   notExtracted: [...],     // Monthly exists, Hourly never extracted — not checked
   *   totalChecked,            // facility-months actually compared
   *   hourlyMonths, monthlyMonths
   * }.
   */
  function compareHourlyVsMonthly(readings, conn) {
    const settings = conn.templateSettings || {};
    const utcOffset = settings.utcOffset ?? 8;

    const hourlySum = new Map();   // `${facilityId}|${month}` -> summed kWh
    const hourlyDays = new Map();  // `${facilityId}|${month}` -> Set of day-of-month with Hourly rows
    const monthlySum = new Map();
    const monthlyRunAt = new Map(); // `${facilityId}|${month}` -> newest RunAt among its Monthly rows
    const hourlyMonthsSeen = new Set();
    const monthlyMonthsSeen = new Set();

    for (const r of readings) {
      const res = String(r.resolution || "").trim();
      if (res !== "Hourly" && res !== "Monthly") continue;
      const wc = wallClock(r, utcOffset);
      if (!wc) continue;
      const month = `${wc.y}-${pad(wc.mo)}`;
      const facilityId = TemplateExport.facilityKeyFor(r.stationId, settings);
      const key = `${facilityId}|${month}`;
      const kwh = Number(r.kwh) || 0;
      if (res === "Hourly") {
        hourlySum.set(key, (hourlySum.get(key) || 0) + kwh);
        if (!hourlyDays.has(key)) hourlyDays.set(key, new Set());
        hourlyDays.get(key).add(wc.d);
        hourlyMonthsSeen.add(month);
      } else {
        monthlySum.set(key, (monthlySum.get(key) || 0) + kwh);
        monthlyMonthsSeen.add(month);
        const runAt = String(r.runAt || "");
        if (runAt > (monthlyRunAt.get(key) || "")) monthlyRunAt.set(key, runAt);
      }
    }

    const flagged = [];       // compared months that need attention (gap, incomplete hourly, or zero generation)
    const notExtracted = [];  // Monthly exists, Hourly never extracted for that month
    const monthlyMissing = []; // Hourly exists, no Monthly figure to compare against
    let totalChecked = 0;
    const today = todayISO();
    const monthRange = (month) => {
      const [y, mo] = month.split("-").map(Number);
      return compressDays(Array.from({ length: daysInMonth(y, mo) }, (_, i) => `${month}-${pad(i + 1)}`), today);
    };

    for (const [key, monthlyKwhRaw] of monthlySum.entries()) {
      const [facilityId, month] = key.split("|");
      const monthlyKwh = round2(monthlyKwhRaw);
      const days = hourlyDays.get(key);

      // Monthly figure only — Hourly was never extracted for this month.
      if (!days || !days.size) {
        notExtracted.push({ facilityId, month, monthlyKwh, missingRanges: monthRange(month),
          monthlyAsOf: localDateOf(monthlyRunAt.get(key)) });
        continue;
      }

      totalChecked++;
      const [y, mo] = month.split("-").map(Number);
      const totalDays = daysInMonth(y, mo);
      const coveredDays = days.size;
      const complete = coveredDays >= totalDays;
      const hourlyKwh = hourlySum.get(key) || 0;
      const diffKwh = round2(monthlyKwhRaw - hourlyKwh);
      const diffPct = monthlyKwhRaw ? round2((diffKwh / monthlyKwhRaw) * 100) : 0;
      const hasGap = Math.abs(diffKwh) >= 0.005;
      const zeroBoth = Math.abs(monthlyKwhRaw) < 0.005 && Math.abs(hourlyKwh) < 0.005;

      // A month is fine only if it's fully covered, the totals agree, and it
      // actually generated something. (Before, "totals agree" alone was enough,
      // so an incomplete month of 0 kWh in both — like a plant that's offline —
      // was silently reported as "no discrepancies".)
      if (complete && !hasGap && !zeroBoth) continue;

      const missing = [];
      for (let day = 1; day <= totalDays; day++) if (!days.has(day)) missing.push(`${month}-${pad(day)}`);
      const lastHourlyDate = `${month}-${pad(Math.max(...days))}`;
      const monthlyAsOf = localDateOf(monthlyRunAt.get(key));
      // A Monthly total fetched BEFORE the last hourly day is a snapshot of an
      // unfinished month — it can't include the later days' generation.
      const monthlyStale = hasGap && !!monthlyAsOf && monthlyAsOf <= lastHourlyDate;

      const issues = [];
      let kind;
      if (zeroBoth) {
        kind = "zero";
        issues.push("0 kWh in both Monthly and Hourly — check whether the plant was offline or not reporting");
      } else if (hasGap && complete && Math.abs(diffPct) >= 99.5 && Math.abs(hourlyKwh) < 0.005) {
        kind = "outage";
        issues.push("FULL OUTAGE (0 kWh in hourly data)");
      } else if (hasGap && complete) {
        kind = "mismatch";
        issues.push("hourly vs monthly mismatch");
      } else {
        kind = "incomplete";
      }
      if (!complete) issues.push(`Hourly incomplete — ${missing.length} of ${totalDays} days missing` + (hasGap ? "" : " (totals still agree)"));
      if (monthlyStale) issues.push(`Monthly figure fetched ${monthlyAsOf}, before the month's hourly data ended (${lastHourlyDate}) — re-run the Monthly extraction for ${month}`);

      flagged.push({
        facilityId, month, kind,
        monthlyKwh, hourlyKwh: round2(hourlyKwh),
        diffKwh, diffPct, hasGap,
        hourlyCoverage: `${coveredDays}/${totalDays} days`,
        missingDays: missing,
        missingRanges: compressDays(missing, today),
        monthlyAsOf: monthlyAsOf || "",
        monthlyStale,
        complete,
        issue: issues.join(" · "),
      });
    }

    // Hourly months with no Monthly figure at all — nothing to compare against.
    for (const [key, days] of hourlyDays.entries()) {
      if (monthlySum.has(key)) continue;
      const [facilityId, month] = key.split("|");
      const [y, mo] = month.split("-").map(Number);
      monthlyMissing.push({ facilityId, month, hourlyKwh: round2(hourlySum.get(key) || 0),
        hourlyCoverage: `${days.size}/${daysInMonth(y, mo)} days` });
    }

    const byFacMonth = (a, b) => a.facilityId.localeCompare(b.facilityId) || a.month.localeCompare(b.month);
    // Real kWh gaps first (largest first), then coverage/zero issues by facility and month.
    flagged.sort((a, b) => (b.hasGap && b.complete) - (a.hasGap && a.complete)
      || ((a.hasGap && a.complete) ? Math.abs(b.diffKwh) - Math.abs(a.diffKwh) : byFacMonth(a, b)));
    notExtracted.sort(byFacMonth);
    monthlyMissing.sort(byFacMonth);

    return {
      flagged,
      notExtracted,
      monthlyMissing,
      totalChecked,
      hourlyMonths: [...hourlyMonthsSeen].sort(),
      monthlyMonths: [...monthlyMonthsSeen].sort(),
    };
  }

  return { compareHourlyVsMonthly };
})();

if (typeof module !== "undefined") module.exports = CompareEngine;
