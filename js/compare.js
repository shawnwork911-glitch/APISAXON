/* =====================================================================
   Compare — Hourly vs Monthly
   ---------------------------------------------------------------------
   Direct port of check_hourly_vs_monthly() from saxon_gen_bot.py: for
   each facility_id, sums Hourly-resolution kWh into monthly buckets and
   compares against whatever Monthly-resolution rows already say for
   that same facility+month, flagging any month where they disagree.

   Facility grouping matches TemplateExport exactly (same stationFacility
   mapping, same "unassigned station falls back to its own stationId"
   rule) — a facility with multiple stations has ALL of their Hourly rows
   summed, and ALL of their Monthly rows summed, before comparing.

   Threshold: matches the Python script's shipped defaults (both
   threshold_kwh and threshold_pct set to 0), so the only real filter is
   its materiality floor — a month is flagged once its gap reaches at
   least 0.005 kWh. There is currently no separate configurable threshold.
   ===================================================================== */

const CompareEngine = (() => {
  const BRAND_LABEL_TO_KEY = { FusionSolar: "fusionsolar", SolarEdge: "solaredge" };

  function monthKeyFromRow(row, utcOffset) {
    const brandKey = BRAND_LABEL_TO_KEY[row.brand] || row.brand.toLowerCase();
    const wc = TemplateExport.wallClockFromRow(row, brandKey, utcOffset);
    if (!wc) return null;
    return `${wc.y}-${String(wc.mo).padStart(2, "0")}`;
  }

  function round2(n) { return Math.round(n * 100) / 100; }

  /**
   * readings: rows from SheetsClient.listReadings(companyName) — every
   * Resolution mixed together; this filters to Hourly/Monthly itself.
   * conn: the connection object (for its templateSettings — facility
   * mapping and UTC offset).
   * Returns { flagged: [...], totalChecked, hourlyMonths, monthlyMonths }.
   */
  function compareHourlyVsMonthly(readings, conn) {
    const settings = conn.templateSettings || {};
    const utcOffset = settings.utcOffset ?? 8;

    const hourlySum = new Map();  // `${facilityId}|${month}` -> summed kWh
    const monthlySum = new Map();
    const hourlyMonthsSeen = new Set();
    const monthlyMonthsSeen = new Set();

    for (const r of readings) {
      if (r.resolution !== "Hourly" && r.resolution !== "Monthly") continue;
      const month = monthKeyFromRow(r, utcOffset);
      if (!month) continue;
      const facilityId = TemplateExport.facilityKeyFor(r.stationId, settings);
      const key = `${facilityId}|${month}`;
      if (r.resolution === "Hourly") {
        hourlySum.set(key, (hourlySum.get(key) || 0) + (r.kwh || 0));
        hourlyMonthsSeen.add(month);
      } else {
        monthlySum.set(key, (monthlySum.get(key) || 0) + (r.kwh || 0));
        monthlyMonthsSeen.add(month);
      }
    }

    // Only checked where a Monthly figure actually exists to compare against
    // (matches the Python script dropping rows with no monthly_report).
    const flagged = [];
    for (const [key, monthlyKwh] of monthlySum.entries()) {
      const [facilityId, month] = key.split("|");
      const hourlyKwh = hourlySum.get(key) || 0;
      const diffKwh = round2(monthlyKwh - hourlyKwh);
      const diffPct = monthlyKwh ? round2((diffKwh / monthlyKwh) * 100) : 0;
      const hasGap = Math.abs(diffKwh) >= 0.005;
      if (!hasGap) continue;
      flagged.push({
        facilityId, month,
        monthlyKwh: round2(monthlyKwh), hourlyKwh: round2(hourlyKwh),
        diffKwh, diffPct,
        issue: Math.abs(diffPct) >= 99.5 ? "FULL OUTAGE (0 kWh in hourly data)" : "hourly vs monthly mismatch",
      });
    }
    flagged.sort((a, b) => b.diffKwh - a.diffKwh);

    return {
      flagged,
      totalChecked: monthlySum.size,
      hourlyMonths: [...hourlyMonthsSeen].sort(),
      monthlyMonths: [...monthlyMonthsSeen].sort(),
    };
  }

  return { compareHourlyVsMonthly };
})();

if (typeof module !== "undefined") module.exports = CompareEngine;
