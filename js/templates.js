/* =====================================================================
   Template export — Hourly Format (Template 1) / STX format (Template 2)
   ---------------------------------------------------------------------
   Ported from generation-data-converter.zip and cross-checked against
   fusionsolar_bot.py's save_hourly_template() (numerically matched
   against the real reference files GEN3294_Hourly_..._template_1.xlsx
   and GEN3294_STX_..._template_2.xlsx).

   TIMESTAMP MATH
     - Template 1's "datetime_start_utc" is NOT real UTC — it's
       local + offset (a deliberate, pre-existing convention).
     - Template 2's "datetime_start_utc" IS standard UTC — local − offset.
   All arithmetic operates on a normalized {y, mo, d, H, Mi, S} local
   wall-clock struct — wallClockFromRow() derives that correctly per
   brand, since FusionSolar's collectTime is a genuine epoch (needs the
   offset added to recover local time) while SolarEdge's date strings
   are already local (parsed as-is, no shift).

   FACILITY GROUPING (multi-station companies)
     Neither template has a per-row station/meter identifier that
     distinguishes multiple stations under one facility — Template 1 has
     no station column at all, and Template 2's meter_id is meant to
     represent the facility's own meter, not an individual inverter. So
     when several stations belong to the same reporting facility, their
     hourly readings are SUMMED into one row per hour per facility,
     exactly matching what these templates were designed to hold.
     stationFacility maps stationId -> facilityId (blank/unset stations
     fall back to using their own stationId as the facility key, so they
     still produce their own separate file rather than being silently
     dropped). facilityGroups holds any extra per-facility settings
     (Template 2's meter_id / eac_registry_id).
   ===================================================================== */

const TemplateExport = (() => {
  function pad(n, len = 2) { return String(n).padStart(len, "0"); }

  function wcToTrickEpoch(wc) { return Date.UTC(wc.y, wc.mo - 1, wc.d, wc.H, wc.Mi, wc.S); }
  function trickEpochToWc(epoch) {
    const d = new Date(epoch);
    return { y: d.getUTCFullYear(), mo: d.getUTCMonth() + 1, d: d.getUTCDate(), H: d.getUTCHours(), Mi: d.getUTCMinutes(), S: d.getUTCSeconds() };
  }

  function fmtSpace(wc) { return `${wc.y}-${pad(wc.mo)}-${pad(wc.d)} ${pad(wc.H)}:${pad(wc.Mi)}:${pad(wc.S)}`; }
  function fmtOffsetSuffix(h) {
    const sign = h < 0 ? "-" : "+";
    const abs = Math.abs(h);
    const hh = Math.floor(abs), mm = Math.round((abs - hh) * 60);
    return `${sign}${pad(hh)}:${pad(mm)}`;
  }
  function fmtISOWithOffset(wc, offsetHours) { return `${wc.y}-${pad(wc.mo)}-${pad(wc.d)}T${pad(wc.H)}:${pad(wc.Mi)}:${pad(wc.S)}${fmtOffsetSuffix(offsetHours)}`; }
  function fmtISOZ(wc) { return `${wc.y}-${pad(wc.mo)}-${pad(wc.d)}T${pad(wc.H)}:${pad(wc.Mi)}:${pad(wc.S)}Z`; }

  function wallClockFromRow(row, brandKey, utcOffsetHours) {
    if (brandKey === "fusionsolar") {
      const localTrick = Number(row.timestamp) + utcOffsetHours * 3600 * 1000;
      return trickEpochToWc(localTrick);
    }
    const s = String(row.timestamp).trim();
    let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})[ T](\d{1,2}):(\d{1,2})(?::(\d{1,2}))?/);
    if (m) return { y: +m[1], mo: +m[2], d: +m[3], H: +m[4], Mi: +m[5], S: +(m[6] || 0) };
    const n = Number(row.timestamp);
    if (!isNaN(n)) return trickEpochToWc(n + utcOffsetHours * 3600 * 1000);
    return null;
  }

  function headersFor(template) {
    return template === "1"
      ? ["month", "day", "year", "datetime_start_local", "local_timezone", "datetime_start_utc", "value (KWh)", "value (MWh)", "facility_id"]
      : ["datetime_start_local", "local_timezone", "datetime_start_utc", "value", "unit_of_measurement", "meter_id", "eac_facility_id", "eac_registry_id"];
  }

  function defaultSettings(template) {
    return template === "1"
      ? { template: "1", tzLabel: "SGT", utcOffset: 8, stationFacility: {}, facilityGroups: {} }
      : { template: "2", tzLabel: "Asia/Singapore", utcOffset: 8, unitOM: "MWh", stationFacility: {}, facilityGroups: {} };
  }

  function facilityKeyFor(stationId, settings) {
    const assigned = (settings.stationFacility || {})[stationId];
    return (assigned && assigned.trim()) || stationId; // unassigned stations stay their own separate group
  }

  /**
   * rows: [{ timestamp, stationId, kwh, raw }] — as produced by brands.js fetchSeries,
   * possibly spanning several stations in one extraction job.
   * brandKey: which brand these rows came from (affects timestamp parsing).
   * settings: from defaultSettings(), edited by the user — see facility grouping above.
   * Returns an array of { facilityId, headers, rows }, one per distinct facility,
   * each internally summed to one row per hour and sorted earliest → latest.
   */
  function build(rows, brandKey, settings) {
    const headers = headersFor(settings.template);
    // facilityId -> localTrickEpoch(hour) -> accumulated kwh
    const buckets = new Map();

    for (const row of rows) {
      const wc = wallClockFromRow(row, brandKey, settings.utcOffset);
      if (!wc) continue; // unparseable timestamp — skipped rather than guessed
      const hourWc = { ...wc, Mi: 0, S: 0 }; // group by the hour, in case of any sub-hour timestamps
      const localTrick = wcToTrickEpoch(hourWc);
      const facilityId = facilityKeyFor(row.stationId, settings);

      if (!buckets.has(facilityId)) buckets.set(facilityId, new Map());
      const hourMap = buckets.get(facilityId);
      hourMap.set(localTrick, (hourMap.get(localTrick) || 0) + (row.kwh || 0));
    }

    const groups = [];
    for (const [facilityId, hourMap] of buckets.entries()) {
      const groupSettings = (settings.facilityGroups || {})[facilityId] || {};
      const entries = [...hourMap.entries()].sort((a, b) => a[0] - b[0]);
      const outRows = entries.map(([localTrick, kwhSum]) => {
        const wc = trickEpochToWc(localTrick);
        const kwhR = Math.round(kwhSum * 1e3) / 1e3;
        const mwhR = Math.round((kwhR / 1000) * 1e6) / 1e6;
        if (settings.template === "1") {
          const utcWc = trickEpochToWc(localTrick + settings.utcOffset * 3600 * 1000); // local + offset (intentional)
          return [wc.mo, wc.d, wc.y, fmtSpace(wc), settings.tzLabel, fmtSpace(utcWc), kwhR, mwhR, facilityId];
        }
        const utcWc = trickEpochToWc(localTrick - settings.utcOffset * 3600 * 1000); // standard UTC
        const value = settings.unitOM === "MWh" ? mwhR : kwhR;
        return [fmtISOWithOffset(wc, settings.utcOffset), settings.tzLabel, fmtISOZ(utcWc), value,
                settings.unitOM, groupSettings.meterId || facilityId, facilityId, groupSettings.eacRegistryId || "tigr"];
      });
      groups.push({ facilityId, headers, rows: outRows });
    }
    groups.sort((a, b) => a.facilityId.localeCompare(b.facilityId));
    return groups;
  }

  return { headersFor, defaultSettings, facilityKeyFor, build, wallClockFromRow };
})();

if (typeof module !== "undefined") module.exports = TemplateExport;
