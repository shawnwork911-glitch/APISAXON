/* =====================================================================
   Template export — Hourly Format (Template 1) / STX format (Template 2)
   ---------------------------------------------------------------------
   Ported from generation-data-converter.zip and cross-checked against
   fusionsolar_bot.py's save_hourly_template(), which is the proven,
   already-in-production version of this logic. Both agree on:

     - Template 1's "datetime_start_utc" is NOT real UTC — it's
       local + offset (a deliberate, pre-existing convention, verified
       against real Hourly Format exports, not a bug).
     - Template 2's "datetime_start_utc" IS standard UTC — local − offset.

   All the wall-clock/offset arithmetic below operates on a normalized
   {y, mo, d, H, Mi, S} local wall-clock struct — callers are responsible
   for correctly deriving that struct from whichever raw timestamp
   representation a brand's API actually returns (see wallClockFromRow
   below), since brands differ:
     - FusionSolar's collectTime is a genuine Unix epoch (UTC) — the
       offset has to be added explicitly to recover local wall-clock time.
     - SolarEdge's date strings are already expressed in local site time
       — no shift needed, just parsed directly.
   ===================================================================== */

const TemplateExport = (() => {
  function pad(n, len = 2) { return String(n).padStart(len, "0"); }

  // A local wall-clock struct, encoded as an epoch value purely so we can
  // reuse Date's UTC getters as a formatting convenience — this epoch does
  // NOT represent a real point in time, it's "wall-clock numbers packed
  // into an epoch", exactly as the reference converter does internally.
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

  /**
   * Derives the correct LOCAL wall-clock struct for one extracted row,
   * given which brand it came from.
   *   row.timestamp — FusionSolar: epoch ms (true UTC). SolarEdge: a
   *                   "YYYY-MM-DD HH:MM:SS"-style string already in local time.
   *   utcOffsetHours — only used for FusionSolar's epoch shift.
   */
  function wallClockFromRow(row, brandKey, utcOffsetHours) {
    if (brandKey === "fusionsolar") {
      const localTrick = Number(row.timestamp) + utcOffsetHours * 3600 * 1000;
      return trickEpochToWc(localTrick);
    }
    // SolarEdge (and best-effort brands using a similar local-string convention).
    const s = String(row.timestamp).trim();
    let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})[ T](\d{1,2}):(\d{1,2})(?::(\d{1,2}))?/);
    if (m) return { y: +m[1], mo: +m[2], d: +m[3], H: +m[4], Mi: +m[5], S: +(m[6] || 0) };
    // Fallback: treat as epoch ms if it parses as a number.
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
      ? { template: "1", tzLabel: "SGT", utcOffset: 8, facilityId: "" }
      : { template: "2", tzLabel: "Asia/Singapore", utcOffset: 8, meterId: "", eacFacilityId: "", eacRegistryId: "tigr", unitOM: "MWh" };
  }

  /**
   * rows: [{ timestamp, stationId, kwh, raw }] — as produced by brands.js fetchSeries.
   * brandKey: which brand these rows came from (affects timestamp parsing — see above).
   * settings: from defaultSettings(), possibly edited by the user.
   * Returns { headers, rows: array of arrays }, sorted earliest → latest.
   */
  function build(rows, brandKey, settings) {
    const headers = headersFor(settings.template);
    const out = [];
    for (const row of rows) {
      const wc = wallClockFromRow(row, brandKey, settings.utcOffset);
      if (!wc) continue; // unparseable timestamp — skipped rather than guessed
      const localTrick = wcToTrickEpoch(wc);
      const kwhR = Math.round((row.kwh || 0) * 1e6) / 1e6;
      const mwhR = Math.round((kwhR / 1000) * 1e9) / 1e9;

      let arr;
      if (settings.template === "1") {
        const utcWc = trickEpochToWc(localTrick + settings.utcOffset * 3600 * 1000); // local + offset (intentional, non-standard)
        arr = [wc.mo, wc.d, wc.y, fmtSpace(wc), settings.tzLabel, fmtSpace(utcWc),
               Math.round(kwhR * 1000) / 1000, Math.round(mwhR * 1e6) / 1e6, settings.facilityId || ""];
      } else {
        const utcWc = trickEpochToWc(localTrick - settings.utcOffset * 3600 * 1000); // standard UTC
        const value = settings.unitOM === "MWh" ? mwhR : kwhR;
        arr = [fmtISOWithOffset(wc, settings.utcOffset), settings.tzLabel, fmtISOZ(utcWc), value,
               settings.unitOM, settings.meterId || "", settings.eacFacilityId || "", settings.eacRegistryId || ""];
      }
      out.push({ epoch: localTrick, arr });
    }
    out.sort((a, b) => a.epoch - b.epoch);
    return { headers, rows: out.map(o => o.arr) };
  }

  return { headersFor, defaultSettings, build };
})();

if (typeof module !== "undefined") module.exports = TemplateExport;
