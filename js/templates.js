/* =====================================================================
   Template export — Hourly Format (Template 1) / STX format (Template 2) / ACT format (Template 3)
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

  // Display name for a template id ("1", "2", "3"/ACT) — used in the UI and audit log.
  function templateLabel(template) {
    return TEMPLATE_INFO[template]?.label || `Template ${template}`;
  }

  // Name + one sample row per Hourly template, for the Export tab's hover preview.
  const TEMPLATE_INFO = {
    "1": { label: "Normal Hourly format", desc: "kWh and MWh, month/day/year columns, one facility_id",
           sample: [1, 1, 2026, "2026-01-01 08:00:00", "SGT", "2026-01-01 16:00:00", 135.4, 0.1354, "GEN4xx6"] },
    "2": { label: "STX format", desc: "ISO timestamps with offset, meter_id and registry per facility",
           sample: ["2026-01-01T08:00:00+08:00", "Asia/Singapore", "2026-01-01T00:00:00Z", 0.1354, "MWh", "", "GEN4xx6", "tigr"] },
    "3": { label: "ACT format", desc: "Excel date-times, MWh, registry always \"evident\"",
           sample: ["2026-01-01 08:00:00", "Asia/Singapore", "2026-01-01 00:00:00", "0.1354", "MWh", "", "GEN4xx6", "evident"] },
  };
  function templateInfo(template) { return TEMPLATE_INFO[template] || null; }

  // ACT format (template "3") — matches meter-data-template 2026.xlsx:
  // same 8 columns as Template 2, but datetime_start_local is a plain
  // "YYYY-MM-DD HH:MM:SS" local time (written as a real Excel date on
  // download), datetime_start_utc is standard UTC, value is always MWh,
  // and eac_registry_id is always "evident".
  const ACT = { tzLabel: "Asia/Singapore", unit: "MWh", registry: "evident" };

  function headersFor(template) {
    return template === "1"
      ? ["month", "day", "year", "datetime_start_local", "local_timezone", "datetime_start_utc", "value (kWh)", "value (MWh)", "facility_id"]
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
   * GROUPING (Export only) — which stations under one facility_id (AGG ID) are summed together.
   * grouping: { mode: "combined" | "separate" | "custom", separate: [stationId, …] }
   *   combined  every station sharing an AGG ID is summed together (default, old behaviour)
   *   separate  every station is kept on its own
   *   custom    stations listed in `separate` are kept on their own; the rest of
   *             their AGG ID's stations are still summed together
   * Returns { key, facilityId, stationId } — stationId is null for a combined group.
   * A station with no AGG ID assigned is already its own group either way.
   */
  function groupFor(stationId, settings, grouping) {
    const facilityId = facilityKeyFor(stationId, settings);
    if (facilityId === stationId || !isSeparate(stationId, grouping)) return { key: facilityId, facilityId, stationId: null };
    return { key: `${facilityId}\u0001${stationId}`, facilityId, stationId };
  }
  function isSeparate(stationId, grouping) {
    const mode = grouping?.mode || "combined";
    return mode === "separate" || (mode === "custom" && (grouping.separate || []).includes(stationId));
  }
  // Short text for summaries and the audit log. stationNames: { stationId: name }.
  function describeGrouping(grouping, stationNames = {}) {
    const mode = grouping?.mode || "combined";
    if (mode === "separate") return "every station separate";
    if (mode === "custom") {
      const sep = grouping.separate || [];
      return sep.length ? `kept separate: ${sep.map(id => stationNames[id] || id).join(", ")} · rest combined by AGG ID`
        : "all stations combined by AGG ID";
    }
    return "all stations combined by AGG ID";
  }

  /**
   * rows: [{ timestamp, stationId, kwh, raw }] — as produced by brands.js fetchSeries,
   * possibly spanning several stations in one extraction job.
   * brandKey: which brand these rows came from (affects timestamp parsing).
   * settings: from defaultSettings(), edited by the user — see facility grouping above.
   * grouping: see groupFor() — omitted means combined per AGG ID.
   * stationNames: optional { stationId: name } for labelling split groups.
   * Returns an array of { key, facilityId, stationId, stationName, members, headers, rows }
   * (members = names of the stations summed into it),
   * one per group, each internally summed to one row per hour and sorted
   * earliest → latest. The facility_id column always carries the AGG ID,
   * even when a facility is split per station. meter_id is left blank unless
   * a real meter ID was entered for the facility (never defaults to facility_id).
   */
  function build(rows, brandKey, settings, grouping, stationNames = {}) {
    const headers = headersFor(settings.template);
    // facilityId -> localTrickEpoch(hour) -> accumulated kwh
    const buckets = new Map();

    for (const row of rows) {
      const wc = wallClockFromRow(row, brandKey, settings.utcOffset);
      if (!wc) continue; // unparseable timestamp — skipped rather than guessed
      const hourWc = { ...wc, Mi: 0, S: 0 }; // group by the hour, in case of any sub-hour timestamps
      const localTrick = wcToTrickEpoch(hourWc);
      const g = groupFor(row.stationId, settings, grouping);

      if (!buckets.has(g.key)) buckets.set(g.key, { g, hours: new Map(), members: new Set() });
      buckets.get(g.key).members.add(stationNames[row.stationId] || row.stationName || row.stationId);
      const hourMap = buckets.get(g.key).hours;
      hourMap.set(localTrick, (hourMap.get(localTrick) || 0) + (row.kwh || 0));
    }

    const groups = [];
    for (const { g, hours: hourMap, members } of buckets.values()) {
      const facilityId = g.facilityId;
      const groupSettings = (settings.facilityGroups || {})[facilityId] || {};
      // meter_id stays empty unless a real one was entered. Older saved settings
      // auto-copied facility_id into meter_id, so treat that value as blank too.
      const rawMeter = String(groupSettings.meterId ?? "").trim();
      const meterId = rawMeter && rawMeter !== facilityId ? rawMeter : "";
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
        if (settings.template === "3") {
          return [fmtSpace(wc), ACT.tzLabel, fmtSpace(utcWc), mwhR, ACT.unit,
                  meterId, facilityId, ACT.registry];
        }
        // unit_of_measurement defaults to "MWh". Companies whose saved settings
        // started from the Normal Hourly defaults have no unitOM, which used to
        // leave this column empty (and the value in kWh).
        const unit = settings.unitOM === "kWh" ? "kWh" : "MWh";
        const value = unit === "MWh" ? mwhR : kwhR;
        // Same cause: those settings carry the Normal Hourly "SGT" label, but
        // STX's local_timezone is the IANA name.
        const tz = settings.tzLabel && settings.tzLabel !== "SGT" ? settings.tzLabel : "Asia/Singapore";
        return [fmtISOWithOffset(wc, settings.utcOffset), tz, fmtISOZ(utcWc), value,
                unit, meterId, facilityId, groupSettings.eacRegistryId || "tigr"];
      });
      groups.push({ key: g.key, facilityId, stationId: g.stationId, members: [...members].sort(),
        stationName: g.stationId ? (stationNames[g.stationId] || g.stationId) : "", headers, rows: outRows });
    }
    groups.sort((a, b) => a.facilityId.localeCompare(b.facilityId) || a.stationName.localeCompare(b.stationName));
    return groups;
  }

  return { headersFor, templateLabel, templateInfo, defaultSettings, facilityKeyFor, groupFor, isSeparate, describeGrouping, build, wallClockFromRow };
})();

if (typeof module !== "undefined") module.exports = TemplateExport;
