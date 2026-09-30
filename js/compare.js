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
      }
    }

    const flagged = [];
    const notExtracted = [];
    let totalChecked = 0;

    for (const [key, monthlyKwhRaw] of monthlySum.entries()) {
      const [facilityId, month] = key.split("|");
      const monthlyKwh = round2(monthlyKwhRaw);
      const days = hourlyDays.get(key);

      // Monthly figure only — Hourly was never extracted for this month.
      if (!days || !days.size) {
        notExtracted.push({ facilityId, month, monthlyKwh });
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
      if (Math.abs(diffKwh) < 0.005) continue;

      let issue;
      if (!complete) issue = `Hourly incomplete — only ${coveredDays} of ${totalDays} days extracted`;
      else if (Math.abs(diffPct) >= 99.5 && Math.abs(hourlyKwh) < 0.005) issue = "FULL OUTAGE (0 kWh in hourly data)";
      else issue = "hourly vs monthly mismatch";

      flagged.push({
        facilityId, month,
        monthlyKwh, hourlyKwh: round2(hourlyKwh),
        diffKwh, diffPct,
        hourlyCoverage: `${coveredDays}/${totalDays} days`,
        complete,
        issue,
      });
    }

    // Real mismatches first (largest gap first), then incomplete-coverage months.
    flagged.sort((a, b) => (b.complete - a.complete) || (Math.abs(b.diffKwh) - Math.abs(a.diffKwh)));
    notExtracted.sort((a, b) => a.facilityId.localeCompare(b.facilityId) || a.month.localeCompare(b.month));

    return {
      flagged,
      notExtracted,
      totalChecked,
      hourlyMonths: [...hourlyMonthsSeen].sort(),
      monthlyMonths: [...monthlyMonthsSeen].sort(),
    };
  }

  return { compareHourlyVsMonthly };
})();

if (typeof module !== "undefined") module.exports = CompareEngine;
