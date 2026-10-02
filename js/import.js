/* =====================================================================
   Excel import — generation data from a spreadsheet into Readings
   ---------------------------------------------------------------------
   Lets someone upload an .xlsx / .xls / .csv of generation data, say what
   each sheet holds (Hourly / Daily / Monthly / don't import), map each
   sheet's column titles to the Readings columns, preview the result, and
   append it to the Google Sheet's Readings tab. From then on it is
   ordinary Readings data: Compare and Export pick it up exactly like rows
   written by an Extraction run.

   COMPANY
     Typed in freely. A name matching an existing connection uses that
     connection's stations and facility_id settings. A new name is saved
     to the Connections tab as an "Excel import" company (pseudo-brand
     "excel", no API), with its stations taken from the file — so it shows
     up in the Dashboard, Compare and Export like any other company.

   OUTPUT SHAPE — identical to what the Extraction job writes:
     { timestamp: "YYYY-MM-DD HH:MM:SS" (local wall clock),
       company, brand (label), stationId, stationName, resolution, kwh }
   Timestamps are bucketed to the sheet's resolution:
     Hourly  → "YYYY-MM-DD HH:00:00"
     Daily   → "YYYY-MM-DD 00:00:00"
     Monthly → "YYYY-MM-01 00:00:00"
   Several rows landing in the same station + period are SUMMED, so the
   de-dupe rule (one row per Resolution + StationId + Timestamp, newest
   RunAt wins) never drops part of a total. A reading that appears on two
   sheets of the same resolution is kept from the first sheet only.

   STATIONS
     Matched against station IDs, then names, then facility_id (only when
     that facility has exactly one station). A sheet can also be assigned
     to one station. Import-only companies gain any new station values.

   EXISTING ROWS
     Before saving, the Sheet is read for the file's date range. By
     default readings that already exist are left alone; "Replace"
     appends them anyway, and the newest row then wins. A re-run after a
     partial failure is therefore safe.

   Column mappings are remembered per company + resolution in this
   browser (localStorage).
   ===================================================================== */

const ExcelImport = (() => {
  const STORE_KEY = "slc.importMappings";
  const CHUNK = 1000;          // rows per appendReadings call
  const PREVIEW_ROWS = 15;
  const MONTHS = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];

  const FIELDS = [
    { key: "timestamp", label: "Timestamp", required: true, help: "Date, or date and time, of each reading" },
    { key: "time", label: "Time", required: false, help: "Only if the time of day is in a separate column" },
    { key: "station", label: "Station", required: true, help: "Station ID, station name or facility_id — or one station for the whole sheet" },
    { key: "kwh", label: "kWh", required: true, help: "Generation for that period" },
  ];

  // Header guesses, in priority order (headers are compared lower-case, letters/digits only).
  const GUESS = {
    timestamp: ["timestamp", "datetime", "datetimestartlocal", "datetimelocal", "collecttime", "readingtime", "readingdate", "date", "period", "month", "intervalstart", "starttime", "time"],
    time: ["time", "hour", "hr", "interval", "timeofday", "starttime", "hourending"],
    station: ["station", "stationname", "stationid", "stationcode", "site", "sitename", "siteid", "plant", "plantname", "plantid", "facilityid", "facility", "genid", "meterid", "meter"],
    kwh: ["kwh", "valuekwh", "energykwh", "generationkwh", "yieldkwh", "pvyield", "pvyieldkwh", "yield", "generation", "energy", "production", "output", "value", "mwh", "valuemwh", "wh"],
  };
  const GUESS_CONTAINS = {
    timestamp: ["timestamp", "datetime", "date"],
    station: ["station", "site", "plant"],
    kwh: ["kwh", "mwh", "yield", "energy", "generation"],
  };

  let deps = {};
  let st = null; // the current import session

  const qs = (id) => document.getElementById(id);
  const pad = (n) => String(n).padStart(2, "0");
  const esc = (s) => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const norm = (s) => String(s ?? "").toLowerCase().replace(/[^a-z0-9]/g, "");
  const isBlank = (v) => v == null || String(v).trim() === "";
  const daysIn = (y, mo) => new Date(Date.UTC(y, mo, 0)).getUTCDate();
  const brandLabel = (conn) => (typeof BRANDS !== "undefined" && BRANDS[conn.brand]?.label) || conn.brand;
  const colLetter = (i) => (typeof XLSX !== "undefined" ? XLSX.utils.encode_col(i) : String(i + 1));

  /* ---------------------------- parsing ---------------------------- */

  const year = (y) => { y = +y; return y < 100 ? 2000 + y : y; };
  function ampm(H, ap) {
    if (!ap) return H;
    if (H === 12) H = 0;
    return ap.toUpperCase() === "PM" ? H + 12 : H;
  }
  function monthNum(name) {
    const i = MONTHS.indexOf(String(name).slice(0, 3).toLowerCase());
    return i < 0 ? null : i + 1;
  }
  function valid(wc) {
    return wc && [wc.y, wc.mo, wc.d, wc.H, wc.Mi, wc.S].every(Number.isFinite)
      && wc.y >= 1950 && wc.y <= 2100 && wc.mo >= 1 && wc.mo <= 12
      && wc.d >= 1 && wc.d <= daysIn(wc.y, wc.mo)
      && wc.H >= 0 && wc.H <= 24 && wc.Mi >= 0 && wc.Mi < 60 && wc.S >= 0 && wc.S < 60;
  }
  function fromUtcMs(ms) {
    const d = new Date(ms);
    return { y: d.getUTCFullYear(), mo: d.getUTCMonth() + 1, d: d.getUTCDate(), H: d.getUTCHours(),
      Mi: d.getUTCMinutes(), S: d.getUTCSeconds(), hasTime: true, isUtc: true };
  }
  function mk(y, mo, d, H, Mi, S, hasTime, ap) {
    const wc = { y: year(y), mo: +mo, d: +d, H: ampm(+(H || 0), ap), Mi: +(Mi || 0), S: +(S || 0), hasTime: !!hasTime, isUtc: false };
    return valid(wc) ? wc : null;
  }
  // "Z" or "+08:00" on the end → a real instant; return it as a UTC wall clock.
  function applyZone(wc, zone) {
    if (!wc || !zone) return wc;
    if (/^z$/i.test(zone)) return { ...wc, isUtc: true };
    const m = zone.match(/^([+-])(\d{2}):?(\d{2})$/);
    if (!m) return wc;
    const mins = (m[1] === "-" ? -1 : 1) * (+m[2] * 60 + +m[3]);
    return fromUtcMs(Date.UTC(wc.y, wc.mo - 1, wc.d, wc.H, wc.Mi, wc.S) - mins * 60000);
  }

  // One date/time cell → { y, mo, d, H, Mi, S, hasTime, isUtc } or null.
  // order: "DMY" | "MDY" — only used for ambiguous 03/04/2026-style dates.
  function parseDateCell(v, order) {
    if (v instanceof Date) {
      return isNaN(v) ? null : mk(v.getFullYear(), v.getMonth() + 1, v.getDate(), v.getHours(), v.getMinutes(), v.getSeconds(), true);
    }
    const s = String(v ?? "").trim();
    if (!s) return null;
    if (typeof v === "number" || /^\d+(\.\d+)?$/.test(s)) {
      const n = Number(s);
      if (n > 1e11) return fromUtcMs(n);           // epoch milliseconds (e.g. FusionSolar collectTime)
      if (n > 1e9) return fromUtcMs(n * 1000);     // epoch seconds
      if (Number.isInteger(n) && n >= 19500101 && n <= 21001231) {
        return mk(Math.floor(n / 10000), Math.floor(n / 100) % 100, n % 100, 0, 0, 0, false); // 20260131
      }
      if (n >= 18264 && n < 73051) {               // Excel serial date (1950–2099)
        const total = Math.round(n * 86400);
        const days = Math.floor(total / 86400), secs = total - days * 86400;
        const d = new Date(Date.UTC(1899, 11, 30) + days * 86400000);
        return mk(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate(),
          Math.floor(secs / 3600), Math.floor((secs % 3600) / 60), secs % 60, secs > 0);
      }
      return null;
    }
    const T = String.raw`(?:[ T,]+(\d{1,2}):(\d{2})(?::(\d{2})(?:\.\d+)?)?\s*([AaPp][Mm])?)?\s*(Z|[+-]\d{2}:?\d{2})?`;
    let m = s.match(new RegExp(String.raw`^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})${T}$`));
    if (m) return applyZone(mk(m[1], m[2], m[3], m[4], m[5], m[6], !!m[4], m[7]), m[8]);
    m = s.match(new RegExp(String.raw`^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2}|\d{4})${T}$`));
    if (m) {
      const [a, b] = [+m[1], +m[2]];
      const [d, mo] = order === "MDY" ? [b, a] : [a, b];
      return applyZone(mk(m[3], mo, d, m[4], m[5], m[6], !!m[4], m[7]), m[8]);
    }
    m = s.match(new RegExp(String.raw`^(\d{1,2})[ \-/]([A-Za-z]{3,9})[ \-/,]+(\d{2}|\d{4})${T}$`)); // 5 Jan 2026 [10:00]
    if (m && monthNum(m[2])) return applyZone(mk(m[3], monthNum(m[2]), m[1], m[4], m[5], m[6], !!m[4], m[7]), m[8]);
    m = s.match(new RegExp(String.raw`^([A-Za-z]{3,9})[ \-/]+(\d{1,2}),?[ \-/]+(\d{4})${T}$`));       // Jan 5, 2026
    if (m && monthNum(m[1])) return applyZone(mk(m[3], monthNum(m[1]), m[2], m[4], m[5], m[6], !!m[4], m[7]), m[8]);
    m = s.match(/^(\d{4})[-/.](\d{1,2})$/);                                                          // 2026-01
    if (m) return mk(m[1], m[2], 1, 0, 0, 0, false);
    m = s.match(/^([A-Za-z]{3,9})[\s\-/,']+(\d{2}|\d{4})$/);                                         // Jan 2026, Jan-26
    if (m && monthNum(m[1])) return mk(m[2], monthNum(m[1]), 1, 0, 0, 0, false);
    return null;
  }

  // A separate time-of-day cell → { H, Mi, S } or null.
  function parseTimeCell(v) {
    if (typeof v === "number") {
      if (v >= 0 && v < 1) { const s = Math.round(v * 86400); return { H: Math.floor(s / 3600), Mi: Math.floor((s % 3600) / 60), S: s % 60 }; }
      if (Number.isInteger(v) && v >= 0 && v <= 24) return { H: v, Mi: 0, S: 0 };
      if (v >= 18264) return parseTimeCell(v - Math.floor(v)); // a full date-time serial — keep its time part
      return null;
    }
    let s = String(v ?? "").trim();
    if (!s) return null;
    if (/^0?\.\d+$/.test(s)) return parseTimeCell(Number(s));
    s = s.split(/\s*(?:–|—|\bto\b|-(?=\s*\d))\s*/i)[0];           // "00:00 - 01:00" → "00:00"
    let m = s.match(/^(\d{1,2})(?::(\d{2}))?(?::(\d{2}))?\s*([AaPp][Mm])?$/);
    if (!m) { const n = s.match(/^(\d{2})(\d{2})$/); if (n) m = [s, n[1], n[2], "0", ""]; }  // "0100"
    if (!m) return null;
    const t = { H: ampm(+m[1], m[4]), Mi: +(m[2] || 0), S: +(m[3] || 0) };
    return t.H <= 24 && t.Mi < 60 && t.S < 60 ? t : null;
  }

  // Number, blank (null) or not-a-number (NaN). Accepts "1,234.5", "12.3 kWh".
  function parseNumber(v) {
    if (typeof v === "number") return Number.isFinite(v) ? v : NaN;
    let s = String(v ?? "").trim();
    if (!s || s === "-" || /^n\/?a$/i.test(s)) return null;
    s = s.replace(/[\s,]/g, "").replace(/[kKmM]?[wW][hH]$/, "");
    return /^[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?$/.test(s) ? Number(s) : NaN;
  }

  // Day-first vs month-first, from the dates in the column themselves.
  function detectDateOrder(values) {
    let dmy = 0, mdy = 0;
    for (const v of values) {
      const m = String(v ?? "").trim().match(/^(\d{1,2})[-/.](\d{1,2})[-/.]\d{2,4}/);
      if (!m) continue;
      if (+m[1] > 12) dmy++;
      if (+m[2] > 12) mdy++;
    }
    if (mdy > 0 && dmy === 0) return "MDY";
    return "DMY"; // default for SG/MY data, and when it can't be told apart
  }

  function stationMatcher(conn) {
    const stations = conn.stations || [];
    const k = (s) => String(s ?? "").trim().toLowerCase();
    const byId = new Map(), byName = new Map(), dupNames = new Set(), byFacility = new Map();
    for (const s of stations) {
      byId.set(k(s.id), s);
      const n = k(s.name);
      if (n) { if (byName.has(n) && byName.get(n).id !== s.id) dupNames.add(n); else byName.set(n, s); }
    }
    const sf = conn.templateSettings?.stationFacility || {};
    for (const s of stations) {
      const f = k(sf[s.id]);
      if (!f) continue;
      if (!byFacility.has(f)) byFacility.set(f, []);
      byFacility.get(f).push(s);
    }
    return (value) => {
      const key = k(value);
      if (!key) return { error: "blank" };
      if (byId.has(key)) return { station: byId.get(key) };
      if (dupNames.has(key)) return { error: "ambiguous" };
      if (byName.has(key)) return { station: byName.get(key) };
      const fac = byFacility.get(key);
      if (fac?.length === 1) return { station: fac[0] };
      if (fac?.length > 1) return { error: "multiFacility", count: fac.length };
      return { error: "unknown" };
    };
  }

  /* ---------------------------- company ---------------------------- */

  const MANUAL_BRAND = "excel";
  // Companies typed in on import (no inverter API) are stored as ordinary
  // connections with this pseudo-brand, so Compare, Export, the Dashboard
  // and the stations/facility_id modal all work for them unchanged.
  if (typeof BRANDS !== "undefined" && !BRANDS[MANUAL_BRAND]) {
    BRANDS[MANUAL_BRAND] = { key: MANUAL_BRAND, label: "Excel import", badge: "XL", color: "#2f8f7a",
      confidence: "verified", manual: true, fields: [] };
  }
  const isManual = (conn) => !!conn && conn.brand === MANUAL_BRAND;

  function findConn(name) {
    const k = String(name || "").trim().toLowerCase();
    return k ? (deps.getConnections?.() || []).find(c => String(c.companyName || "").trim().toLowerCase() === k) || null : null;
  }
  // The company as the parsers see it: an existing connection, or a blank one for a new name.
  function company() {
    return st.conn || { companyName: st.companyName, brand: MANUAL_BRAND, stations: [], templateSettings: { utcOffset: 8 } };
  }
  const newStationsAllowed = () => !st.conn || isManual(st.conn); // import-only companies just grow their station list
  const offsetOf = () => Number(company().templateSettings?.utcOffset ?? 8);

  /* ---------------------------- mapping memory ---------------------------- */

  const memKey = (sh) => `${st.companyName.trim().toLowerCase()}|${sh.role}`;
  function loadStore() { try { return JSON.parse(localStorage.getItem(STORE_KEY) || "{}"); } catch { return {}; } }
  function saveMapping(sh) {
    try {
      const store = loadStore();
      const byName = {};
      for (const f of FIELDS) {
        const v = sh.map[f.key];
        byName[f.key] = v === "" || v.startsWith("fixed:") ? v : (sh.headers[+v] ?? "");
      }
      store[memKey(sh)] = { map: byName, fixedName: sh.fixedName,
        opt: { unit: sh.opt.unit, dateOrder: sh.opt.dateOrder, tz: sh.opt.tz, hourEnding: sh.opt.hourEnding, unknown: sh.opt.unknown } };
      localStorage.setItem(STORE_KEY, JSON.stringify(store));
    } catch { /* storage full or blocked — just don't remember */ }
  }
  // A saved mapping only applies if every column it names exists on this sheet.
  function savedMappingFor(sh) {
    const saved = loadStore()[memKey(sh)];
    if (!saved) return null;
    const stations = company().stations || [];
    const map = {};
    for (const f of FIELDS) {
      const v = saved.map?.[f.key] ?? "";
      if (v === "" || v === "fixed:__new__") { map[f.key] = v; continue; }
      if (v.startsWith("fixed:")) {
        if (!stations.some(s => `fixed:${s.id}` === v)) return null;
        map[f.key] = v; continue;
      }
      const idx = sh.headers.findIndex(h => h === v);
      if (idx < 0) return null;
      map[f.key] = String(idx);
    }
    return { map, opt: saved.opt || {}, fixedName: saved.fixedName };
  }

  function guessMapping(sh) {
    const H = sh.headers.map(norm);
    const used = new Set();
    const pick = (field) => {
      for (const syn of GUESS[field]) {
        const i = H.findIndex((h, idx) => h === syn && !used.has(idx));
        if (i >= 0) return i;
      }
      for (const part of GUESS_CONTAINS[field] || []) {
        const i = H.findIndex((h, idx) => h.includes(part) && !used.has(idx));
        if (i >= 0) return i;
      }
      return -1;
    };
    const map = { timestamp: "", time: "", station: "", kwh: "" };
    for (const f of ["timestamp", "kwh", "station"]) {
      const i = pick(f);
      if (i >= 0) { map[f] = String(i); used.add(i); }
    }
    if (map.timestamp !== "" && /date/.test(H[+map.timestamp]) && !/time/.test(H[+map.timestamp])) {
      const i = H.findIndex((h, idx) => GUESS.time.includes(h) && !used.has(idx));
      if (i >= 0) map.time = String(i);
    }
    const stations = company().stations || [];
    if (map.station === "") map.station = stations.length === 1 ? `fixed:${stations[0].id}` : stations.length ? "" : "fixed:__new__";
    const kh = map.kwh === "" ? "" : H[+map.kwh];
    const unit = /mwh/.test(kh) ? "MWh" : /kwh/.test(kh) ? "kWh" : /wh$/.test(kh) ? "Wh" : "kWh";
    return { map, opt: { unit } };
  }

  /* ---------------------------- session ---------------------------- */

  function init(d) {
    deps = d || {};
    qs("btnImportClose").addEventListener("click", close);
    qs("btnImportCancel").addEventListener("click", close);
    qs("btnImportBack").addEventListener("click", back);
    qs("btnImportNext").addEventListener("click", next);
    qs("impFile").addEventListener("change", handleFile);
    qs("impCompany").addEventListener("input", () => { st.companyName = qs("impCompany").value; renderCompanyHint(); });
    qs("impExisting").addEventListener("change", () => { st.existingChoice = qs("impExisting").value; renderPreview(); });
    qs("importModalOverlay").addEventListener("click", (e) => { if (e.target.id === "importModalOverlay" && !st?.saving) close(); });
  }

  // opts: { companyId, resolution } — pre-filled from whichever tab opened it.
  function open(opts = {}) {
    if (st?.saving) { qs("importModalOverlay").classList.add("active"); return; } // a save is still running — show it
    const conns = deps.getConnections?.() || [];
    const pre = conns.find(c => c.id === opts.companyId);
    st = {
      step: 1, companyName: pre?.companyName || "", conn: null,
      defaultRole: ["Hourly", "Daily", "Monthly"].includes(opts.resolution) ? opts.resolution : "Hourly",
      fileName: "", wb: null, sheets: [], active: 0,
      existingKeys: null, existingError: "", existingChoice: "keep", checking: false, checkSeq: 0,
      saving: false, saved: false,
    };
    qs("impCompanyList").innerHTML = conns.map(c => `<option value="${esc(c.companyName)}">${esc(brandLabel(c))}</option>`).join("");
    qs("impCompany").value = st.companyName;
    qs("impFile").value = "";
    qs("impSheetsWrap").hidden = true;
    qs("impExisting").value = "keep";
    qs("impProgress").hidden = true;
    qs("impProgressFill").style.width = "0";
    renderCompanyHint();
    setStatus("");
    showStep(1);
    qs("importModalOverlay").classList.add("active");
    if (!st.companyName) qs("impCompany").focus();
  }

  function close() {
    if (st?.saving) return; // don't abandon a save half-way
    qs("importModalOverlay").classList.remove("active");
    st = null;
  }

  function setStatus(html, kind) {
    const el = qs("impStatus");
    el.innerHTML = html;
    el.style.color = kind === "err" ? "var(--err)" : kind === "ok" ? "var(--ok)" : "var(--muted)";
  }

  function showStep(n) {
    st.step = n;
    ["importStepFile", "importStepMap", "importStepPreview"].forEach((id, i) => { qs(id).hidden = i + 1 !== n; });
    qs("importStepper").querySelectorAll("li").forEach((li, i) => {
      li.classList.toggle("active", i + 1 === n);
      li.classList.toggle("done", i + 1 < n || st.saved);
    });
    qs("btnImportBack").hidden = n === 1 || st.saved;
    qs("btnImportCancel").textContent = st.saved ? "Close" : "Cancel";
    const nextBtn = qs("btnImportNext");
    nextBtn.hidden = st.saved;
    nextBtn.disabled = false;
    nextBtn.textContent = n === 1 ? "Next: map columns" : n === 2 ? "Next: preview" : "Save to Google Sheet";
  }

  function renderCompanyHint() {
    const el = qs("impCompanyHint");
    const name = st.companyName.trim();
    const conn = findConn(name);
    if (!name) { el.innerHTML = "Pick a company from the list, or type a new name."; el.style.color = ""; return; }
    if (conn && isManual(conn)) {
      el.innerHTML = `Existing Excel-import company · ${(conn.stations || []).length} station(s). New station values in the file are added to it.`;
    } else if (conn) {
      el.innerHTML = `Existing ${esc(brandLabel(conn))} company · ${(conn.stations || []).length} station(s). Its stations and facility_id settings are used to match rows.`;
    } else {
      el.innerHTML = `<span style="color:var(--accent);">New company.</span> It'll be added to the Dashboard as an Excel-import company (no API), with its stations taken from the file. Set facility_id there afterwards.`;
    }
  }

  /* ---------------------------- step 1: file & sheets ---------------------------- */

  function guessRole(name) {
    const n = String(name).toLowerCase();
    if (/hour|hrly|\bhr\b|interval/.test(n)) return "Hourly";
    if (/month|mthly|\bmth\b/.test(n)) return "Monthly";
    if (/daily|\bday/.test(n)) return "Daily";
    return "";
  }

  function handleFile() {
    const file = qs("impFile").files[0];
    if (!file) return;
    setStatus("Reading file…");
    const reader = new FileReader();
    reader.onerror = () => setStatus(`Couldn't read the file: ${esc(reader.error?.message || "unknown error")}`, "err");
    reader.onload = () => {
      try {
        const isText = /\.(csv|txt|tsv)$/i.test(file.name);
        // raw:true keeps CSV text as typed, so "03/04/2026" isn't guessed into a date by the parser.
        st.wb = XLSX.read(new Uint8Array(reader.result), { type: "array", raw: isText, cellDates: false });
        st.fileName = file.name;
        st.sheets = st.wb.SheetNames.map(name => loadSheet(name));
        const withData = st.sheets.filter(sh => sh.dataRows > 0);
        for (const sh of st.sheets) {
          if (!sh.dataRows) sh.role = "skip";
          else if (withData.length === 1) sh.role = guessRole(sh.name) || st.defaultRole;
          else sh.role = guessRole(sh.name); // several sheets: the person says which is which
        }
        renderSheets();
        setStatus("");
      } catch (e) {
        setStatus(`That file couldn't be opened as a spreadsheet (${esc(e.message)}).`, "err");
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function loadSheet(name) {
    const ws = st.wb.Sheets[name];
    const sh = {
      name,
      aoa: XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "", blankrows: true }),
      aoaText: XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: "", blankrows: true }),
      role: "", headerIdx: 0, headers: [], dataRows: 0,
      map: { timestamp: "", time: "", station: "", kwh: "" }, mapTouched: false, fromSaved: false,
      opt: { unit: "kWh", dateOrder: "auto", tz: "local", hourEnding: false, unknown: "skip" },
      fixedName: "", detectedOrder: "DMY", built: null,
    };
    // Header row = first of the top 20 rows with 2+ filled cells, mostly text.
    let idx = 0;
    for (let i = 0; i < Math.min(20, sh.aoa.length); i++) {
      const cells = (sh.aoa[i] || []).filter(c => !isBlank(c));
      const textCells = cells.filter(c => typeof c === "string" && !/^[\d.,\-/: ]+$/.test(c.trim()));
      if (cells.length >= 2 && textCells.length >= Math.ceil(cells.length / 2)) { idx = i; break; }
    }
    setHeaderRow(sh, idx);
    return sh;
  }

  function setHeaderRow(sh, idx) {
    sh.headerIdx = idx;
    const width = sh.aoa.reduce((w, r) => Math.max(w, (r || []).length), 0);
    const row = sh.aoa[idx] || [];
    sh.headers = Array.from({ length: width }, (_, i) => String(row[i] ?? "").trim() || `Column ${colLetter(i)}`);
    sh.dataRows = 0;
    for (let i = idx + 1; i < sh.aoa.length; i++) if ((sh.aoa[i] || []).some(c => !isBlank(c))) sh.dataRows++;
    sh.mapTouched = false;
  }

  const ROLE_LABEL = { Hourly: "Hourly data", Daily: "Daily data", Monthly: "Monthly data", skip: "Don't import" };

  function renderSheets() {
    qs("impSheetsWrap").hidden = !st.wb;
    if (!st.wb) return;
    const multi = st.sheets.filter(sh => sh.dataRows).length > 1;
    qs("impSheetsNote").innerHTML = multi
      ? `<strong style="color:var(--ink);">${esc(st.fileName)}</strong> has ${st.sheets.length} sheets. Choose what each one holds — for example the hourly sheet and the monthly sheet. Sheets set to <em>Don't import</em> are ignored.`
      : `<strong style="color:var(--ink);">${esc(st.fileName)}</strong>. Check the resolution of the data and the row the column titles are on.`;
    const opt = (v, cur) => `<option value="${v}"${v === cur ? " selected" : ""}>${ROLE_LABEL[v]}</option>`;
    qs("impSheetTable").innerHTML = `<thead><tr><th>Sheet</th><th>Contains</th><th>Titles on row</th><th>Data rows</th><th>Column titles found</th></tr></thead><tbody>` +
      st.sheets.map((sh, i) => `<tr>
          <td style="white-space:nowrap;"><strong>${esc(sh.name)}</strong></td>
          <td style="min-width:150px;">
            <select class="impRole" data-sheet="${i}"${sh.dataRows ? "" : " disabled"}>
              ${sh.role ? "" : `<option value="" selected>— choose —</option>`}
              ${opt("Hourly", sh.role)}${opt("Daily", sh.role)}${opt("Monthly", sh.role)}${opt("skip", sh.role)}
            </select></td>
          <td style="width:90px;"><input type="number" class="impHeaderRow" data-sheet="${i}" min="1" step="1" value="${sh.headerIdx + 1}"${sh.aoa.length ? "" : " disabled"}></td>
          <td>${sh.dataRows || `<span class="field-help">empty</span>`}</td>
          <td class="map-sample">${sh.dataRows ? sh.headers.slice(0, 8).map(h => `<code>${esc(h)}</code>`).join(" ") + (sh.headers.length > 8 ? " …" : "") : ""}</td>
        </tr>`).join("") + `</tbody>`;
    qs("impSheetTable").querySelectorAll(".impRole").forEach(sel => sel.addEventListener("change", () => {
      const sh = st.sheets[+sel.dataset.sheet];
      sh.role = sel.value; sh.mapTouched = false; sh.built = null;
      sel.querySelector('option[value=""]')?.remove();
    }));
    qs("impSheetTable").querySelectorAll(".impHeaderRow").forEach(inp => inp.addEventListener("change", () => {
      const sh = st.sheets[+inp.dataset.sheet];
      const n = Math.max(1, Math.floor(+inp.value || 1));
      setHeaderRow(sh, n - 1);
      if (!sh.dataRows) sh.role = "skip";
      renderSheets();
    }));
  }

  const chosenSheets = () => st.sheets.filter(sh => ["Hourly", "Daily", "Monthly"].includes(sh.role));

  /* ---------------------------- step 2: mapping ---------------------------- */

  function prepareSheet(sh) {
    if (!sh.mapTouched) {
      const saved = savedMappingFor(sh);
      const g = saved || guessMapping(sh);
      sh.map = { timestamp: "", time: "", station: "", kwh: "", ...g.map };
      Object.assign(sh.opt, g.opt);
      sh.fixedName = (saved && saved.fixedName) || sh.fixedName || st.companyName.trim();
      sh.fromSaved = !!saved;
    }
    const tsCol = sh.map.timestamp === "" ? -1 : +sh.map.timestamp;
    sh.detectedOrder = detectDateOrder(tsCol < 0 ? [] : sh.aoa.slice(sh.headerIdx + 1).map(r => r?.[tsCol]));
  }

  function missingFor(sh) {
    const miss = FIELDS.filter(f => f.required && sh.map[f.key] === "").map(f => f.label);
    if (sh.map.station === "fixed:__new__" && !String(sh.fixedName || "").trim()) miss.push("Station name");
    return miss;
  }

  function renderTabs() {
    const list = chosenSheets();
    qs("impSheetTabs").hidden = list.length < 2;
    qs("impSheetTabs").innerHTML = list.map((sh) => {
      const i = st.sheets.indexOf(sh);
      const ok = !missingFor(sh).length;
      return `<button type="button" class="import-tab${i === st.active ? " active" : ""}" data-sheet="${i}">
          <span class="import-tab-dot" style="background:${ok ? "var(--ok)" : "var(--warn)"};"></span>${esc(sh.role)} · ${esc(sh.name)}</button>`;
    }).join("");
    qs("impSheetTabs").querySelectorAll(".import-tab").forEach(b => b.addEventListener("click", () => {
      st.active = +b.dataset.sheet;
      renderMapping();
    }));
  }

  function columnOptions(sh, selected, noneLabel) {
    return `<option value="">${esc(noneLabel)}</option>` + sh.headers.map((h, i) =>
      `<option value="${i}"${String(i) === selected ? " selected" : ""}>${esc(colLetter(i))} · ${esc(h)}</option>`).join("");
  }

  function renderMapping() {
    const sh = st.sheets[st.active];
    prepareSheet(sh);
    renderTabs();
    const stations = company().stations || [];
    const rows = FIELDS.map(f => {
      let control;
      if (f.key === "station") {
        const fixed = sh.map.station.startsWith("fixed:");
        control = `<select class="impMap" data-field="station">
            ${columnOptions(sh, fixed ? "" : sh.map.station, "— choose a column —")}
            <optgroup label="Whole sheet is one station">
              ${stations.map(s => `<option value="fixed:${esc(s.id)}"${sh.map.station === `fixed:${s.id}` ? " selected" : ""}>${esc(s.name || s.id)}</option>`).join("")}
              ${newStationsAllowed() ? `<option value="fixed:__new__"${sh.map.station === "fixed:__new__" ? " selected" : ""}>${stations.length ? "Another station — type its name" : "One station — type its name"}</option>` : ""}
            </optgroup>
          </select>
          <input id="impFixedName" placeholder="Station name" value="${esc(sh.fixedName)}" style="margin-top:6px;"${sh.map.station === "fixed:__new__" ? "" : " hidden"}>`;
      } else {
        control = `<select class="impMap" data-field="${f.key}">${columnOptions(sh, sh.map[f.key], f.required ? "— choose a column —" : "— none —")}</select>`;
      }
      return `<tr>
          <td style="white-space:nowrap;"><strong>${esc(f.label)}</strong>${f.required ? ` <span style="color:var(--accent);">*</span>` : ""}
            <div class="field-help" style="white-space:normal;max-width:220px;">${esc(f.help)}</div></td>
          <td style="min-width:220px;">${control}</td>
          <td class="map-sample" id="impSample-${f.key}">${sampleHtml(sh, f.key)}</td>
        </tr>`;
    }).join("");
    qs("impMapTable").innerHTML = `<thead><tr><th>Readings column</th><th>Column on sheet “${esc(sh.name)}”</th><th>First values → how they'll be saved</th></tr></thead><tbody>${rows}</tbody>`;
    qs("impMapTable").querySelectorAll(".impMap").forEach(sel => sel.addEventListener("change", () => {
      sh.map[sel.dataset.field] = sel.value;
      sh.mapTouched = true;
      if (sel.dataset.field === "timestamp") {
        sh.detectedOrder = detectDateOrder(sel.value === "" ? [] : sh.aoa.slice(sh.headerIdx + 1).map(r => r?.[+sel.value]));
        renderOptions(sh);
      }
      if (sel.dataset.field === "kwh" && sel.value !== "") {
        const h = norm(sh.headers[+sel.value]);
        const unit = /mwh/.test(h) ? "MWh" : /kwh/.test(h) ? "kWh" : /wh$/.test(h) ? "Wh" : null;
        if (unit) { sh.opt.unit = unit; renderOptions(sh); }
      }
      if (sel.dataset.field === "station") qs("impFixedName").hidden = sel.value !== "fixed:__new__";
      refreshSamples(sh);
      renderTabs();
    }));
    qs("impFixedName").addEventListener("input", () => { sh.fixedName = qs("impFixedName").value; sh.mapTouched = true; refreshSamples(sh); renderTabs(); });
    qs("impMapNote").innerHTML = (sh.fromSaved
      ? `Filled in from the mapping last used for <strong>${esc(st.companyName.trim())}</strong> (${esc(sh.role)}). Check it still fits this sheet.`
      : `Columns were matched by their titles. Check each one against the sample values on the right.`)
      + (chosenSheets().length > 1 ? ` Each sheet is mapped separately — use the tabs above.` : "");
    renderOptions(sh);
  }

  function renderOptions(sh) {
    const off = offsetOf();
    const offLabel = `UTC${off >= 0 ? "+" : ""}${off}`;
    const opt = (v, label, cur) => `<option value="${esc(v)}"${v === cur ? " selected" : ""}>${esc(label)}</option>`;
    const orderName = { DMY: "day first", MDY: "month first" }[sh.detectedOrder];
    qs("impOptions").innerHTML = `
      <div class="field"><label>Unit of the kWh column</label>
        <select data-opt="unit">${opt("Wh", "Wh", sh.opt.unit)}${opt("kWh", "kWh", sh.opt.unit)}${opt("MWh", "MWh", sh.opt.unit)}</select></div>
      <div class="field"><label>Dates like 03/04/2026</label>
        <select data-opt="dateOrder">${opt("auto", `Detect (looks ${orderName})`, sh.opt.dateOrder)}${opt("DMY", "Day first — 3 April", sh.opt.dateOrder)}${opt("MDY", "Month first — 4 March", sh.opt.dateOrder)}</select></div>
      <div class="field"><label>Times on this sheet are</label>
        <select data-opt="tz">${opt("local", `Local time (${offLabel}${st.conn ? ", from company settings" : ""})`, sh.opt.tz)}${opt("utc", `UTC — shift ${off >= 0 ? "+" : ""}${off} h to local`, sh.opt.tz)}</select></div>
      <div class="field"${sh.role === "Hourly" ? "" : " hidden"}><label>Each hour is labelled by its</label>
        <select data-opt="hourEnding">${opt("start", "Start — 00:00 means 00:00–01:00", sh.opt.hourEnding ? "end" : "start")}${opt("end", "End — 01:00 means 00:00–01:00", sh.opt.hourEnding ? "end" : "start")}</select></div>
      <div class="field"${newStationsAllowed() ? " hidden" : ""}><label>Station values not in this company</label>
        <select data-opt="unknown">${opt("skip", "Skip those rows", sh.opt.unknown)}${opt("import", "Import as new station IDs", sh.opt.unknown)}</select></div>`;
    qs("impOptions").querySelectorAll("select[data-opt]").forEach(sel => sel.addEventListener("change", () => {
      const k = sel.dataset.opt;
      sh.opt[k] = k === "hourEnding" ? sel.value === "end" : sel.value;
      sh.mapTouched = true;
      refreshSamples(sh);
    }));
  }

  function refreshSamples(sh) {
    FIELDS.forEach(f => { const el = qs(`impSample-${f.key}`); if (el) el.innerHTML = sampleHtml(sh, f.key); });
  }

  // First few values of the mapped column, each with what it turns into.
  function sampleHtml(sh, field) {
    const v = sh.map[field];
    if (field === "station" && v === "fixed:__new__") {
      const n = String(sh.fixedName || "").trim();
      return n ? `Every row → <span class="good">${esc(n)}</span> <span style="opacity:.7;">(${findStation(n) ? "existing station" : "new station"})</span>` : `<span class="bad">Type the station name</span>`;
    }
    if (field === "station" && v.startsWith("fixed:")) {
      const s = (company().stations || []).find(x => `fixed:${x.id}` === v);
      return `Every row → <span class="good">${esc(s?.name || s?.id || "")}</span> <span style="opacity:.7;">(${esc(s?.id || "")})</span>`;
    }
    if (v === "") return field === "time" ? `Time is taken from the Timestamp column.` : `<span class="bad">Not mapped yet</span>`;
    const col = +v, out = [];
    const order = sh.opt.dateOrder === "auto" ? sh.detectedOrder : sh.opt.dateOrder;
    const match = field === "station" ? stationMatcher(company()) : null;
    for (let i = sh.headerIdx + 1; i < sh.aoa.length && out.length < 3; i++) {
      const raw = sh.aoa[i]?.[col], text = sh.aoaText[i]?.[col];
      if (isBlank(raw)) continue;
      let res;
      if (field === "timestamp") {
        const wc = parseDateCell(raw, order);
        res = wc ? `<span class="good">${wc.y}-${pad(wc.mo)}-${pad(wc.d)}${wc.hasTime ? ` ${pad(wc.H)}:${pad(wc.Mi)}` : ""}${wc.isUtc ? " UTC" : ""}</span>` : `<span class="bad">not a date</span>`;
      } else if (field === "time") {
        const t = parseTimeCell(raw);
        res = t ? `<span class="good">${pad(t.H)}:${pad(t.Mi)}</span>` : `<span class="bad">not a time</span>`;
      } else if (field === "station") {
        const m = match(raw);
        res = m.station ? `<span class="good">${esc(m.station.name || m.station.id)}</span>`
          : m.error === "multiFacility" ? `<span class="bad">facility with ${m.count} stations</span>`
          : newStationsAllowed() || sh.opt.unknown === "import" ? `<span style="color:var(--warn);">new station</span>` : `<span class="bad">not found</span>`;
      } else {
        const n = parseNumber(raw);
        const f = { Wh: 0.001, kWh: 1, MWh: 1000 }[sh.opt.unit];
        res = Number.isFinite(n) ? `<span class="good">${+(n * f).toFixed(6)} kWh</span>` : `<span class="bad">not a number</span>`;
      }
      out.push(`<div><span style="color:var(--ink);">${esc(text ?? raw)}</span> → ${res}</div>`);
    }
    return out.join("") || `<span class="bad">Column is empty</span>`;
  }

  function findStation(name) {
    const m = stationMatcher(company())(name);
    return m.station || null;
  }

  /* ---------------------------- build rows ---------------------------- */

  function buildRows(sh) {
    const conn = company();
    const resolution = sh.role, map = sh.map, opt = sh.opt;
    const off = offsetOf();
    const order = opt.dateOrder === "auto" ? sh.detectedOrder : opt.dateOrder;
    const factor = { Wh: 0.001, kWh: 1, MWh: 1000 }[opt.unit] || 1;
    const match = stationMatcher(conn);
    const allowNew = newStationsAllowed() || opt.unknown === "import";
    let fixed = null;
    if (map.station === "fixed:__new__") {
      const n = String(sh.fixedName || "").trim();
      fixed = findStation(n) || { id: n, name: n, isNew: true };
    } else if (map.station.startsWith("fixed:")) {
      fixed = (conn.stations || []).find(s => `fixed:${s.id}` === map.station);
    }
    const tsCol = +map.timestamp, timeCol = map.time === "" ? -1 : +map.time;
    const stCol = fixed ? -1 : +map.station, kwhCol = +map.kwh;

    const skipped = new Map();
    const skip = (reason, rowNo, value) => {
      const s = skipped.get(reason) || { count: 0, samples: [] };
      s.count++;
      if (s.samples.length < 3) s.samples.push(`row ${rowNo}${isBlank(value) ? "" : ` (“${String(value).slice(0, 40)}”)`}`);
      skipped.set(reason, s);
    };
    const agg = new Map();
    let read = 0, anyTime = false, negatives = 0;
    const newStations = new Set();
    if (fixed?.isNew) newStations.add(fixed.id);

    for (let i = sh.headerIdx + 1; i < sh.aoa.length; i++) {
      const row = sh.aoa[i] || [];
      if (row.every(isBlank)) continue;
      read++;
      const rowNo = i + 1;
      let wc = parseDateCell(row[tsCol], order);
      if (!wc) { skip("Date/time not recognised", rowNo, sh.aoaText[i]?.[tsCol]); continue; }
      if (timeCol >= 0) {
        const t = parseTimeCell(row[timeCol]);
        if (!t) { skip("Time not recognised", rowNo, sh.aoaText[i]?.[timeCol]); continue; }
        wc = { ...wc, H: t.H, Mi: t.Mi, S: t.S, hasTime: true };
      }
      const n = parseNumber(row[kwhCol]);
      if (n === null) { skip("Blank kWh value", rowNo); continue; }
      if (Number.isNaN(n)) { skip("kWh value is not a number", rowNo, sh.aoaText[i]?.[kwhCol]); continue; }

      let stationId, stationName;
      if (fixed) { stationId = fixed.id; stationName = fixed.name || fixed.id; }
      else {
        const m = match(row[stCol]);
        if (m.station) { stationId = m.station.id; stationName = m.station.name || m.station.id; }
        else if (m.error === "blank") { skip("Blank station", rowNo); continue; }
        else if (m.error === "multiFacility") { skip(`facility_id covers ${m.count} stations — can't tell which one`, rowNo, row[stCol]); continue; }
        else if (m.error === "ambiguous") { skip("Station name shared by two stations", rowNo, row[stCol]); continue; }
        else if (allowNew) { stationId = stationName = String(row[stCol]).trim(); newStations.add(stationId); }
        else { skip("Station not in this company", rowNo, row[stCol]); continue; }
      }
      if (wc.hasTime) anyTime = true;

      // To the local wall clock the Readings tab uses, then bucket to the resolution.
      let ms = Date.UTC(wc.y, wc.mo - 1, wc.d, wc.H, wc.Mi, wc.S);
      if (wc.isUtc || opt.tz === "utc") ms += off * 3600000;
      if (resolution === "Hourly" && opt.hourEnding) ms -= 3600000;
      const l = new Date(ms);
      const y = l.getUTCFullYear(), mo = l.getUTCMonth() + 1;
      const d = resolution === "Monthly" ? 1 : l.getUTCDate();
      const H = resolution === "Hourly" ? l.getUTCHours() : 0;
      const date = `${y}-${pad(mo)}-${pad(d)}`;
      const timestamp = `${date} ${pad(H)}:00:00`;

      const key = `${stationId}\u0001${timestamp}`;
      const e = agg.get(key) || { timestamp, date, stationId, stationName, kwh: 0, n: 0 };
      e.kwh += n * factor;
      e.n++;
      agg.set(key, e);
      if (n < 0) negatives++;
    }

    const rows = [...agg.values()]
      .map(e => ({ ...e, kwh: Math.round(e.kwh * 1e6) / 1e6 }))
      .sort((a, b) => a.stationName.localeCompare(b.stationName) || a.timestamp.localeCompare(b.timestamp));
    const combinedRows = rows.filter(r => r.n > 1);
    const errors = [];
    if (!rows.length) errors.push("No usable rows — check the column mapping and the skipped-row reasons below.");
    if (resolution === "Hourly" && rows.length && !anyTime) {
      errors.push("No time of day was found in any row, so every reading would land at 00:00. Map a Time column, or set this sheet to Daily or Monthly on step 1.");
    }
    return {
      read, rows, skipped, negatives, newStations: [...newStations],
      combined: { buckets: combinedRows.length, sourceRows: combinedRows.reduce((s, r) => s + r.n, 0) },
      fromDate: rows.length ? rows.reduce((m, r) => (r.date < m ? r.date : m), rows[0].date) : "",
      toDate: rows.length ? rows.reduce((m, r) => (r.date > m ? r.date : m), rows[0].date) : "",
      errors, dupOfSheet: new Map(),
    };
  }

  /* ---------------------------- step 3: preview ---------------------------- */

  const rowKey = (res, r) => `${res}\u0001${r.stationId}\u0001${r.timestamp}`;

  async function enterPreview() {
    const list = chosenSheets();
    // Build every sheet; a reading that an earlier sheet of the same resolution
    // already has is dropped here, so the Sheet never gets two versions of it.
    const seen = new Map(); // key -> sheet name
    for (const sh of list) {
      sh.built = buildRows(sh);
      sh.built.rows = sh.built.rows.filter(r => {
        const k = rowKey(sh.role, r);
        if (seen.has(k)) { sh.built.dupOfSheet.set(seen.get(k), (sh.built.dupOfSheet.get(seen.get(k)) || 0) + 1); return false; }
        seen.set(k, sh.name);
        return true;
      });
    }
    st.existingKeys = null;
    st.existingError = "";
    const ok = list.filter(sh => sh.built.rows.length);
    if (!ok.length || list.some(sh => sh.built.errors.length)) { st.checking = false; renderPreview(); return; }
    if (!st.conn) { st.existingKeys = new Set(); renderPreview(); return; } // brand-new company — nothing stored yet
    const from = ok.reduce((m, sh) => (sh.built.fromDate < m ? sh.built.fromDate : m), ok[0].built.fromDate);
    const to = ok.reduce((m, sh) => (sh.built.toDate > m ? sh.built.toDate : m), ok[0].built.toDate);
    const [ty, tm] = to.split("-").map(Number);
    const seq = ++st.checkSeq;
    st.checking = true;
    renderPreview();
    try {
      const existing = await SheetsClient.listReadings(st.conn.companyName, { fromDate: `${from.slice(0, 7)}-01`, toDate: `${ty}-${pad(tm)}-${pad(daysIn(ty, tm))}` });
      if (!st || seq !== st.checkSeq) return;
      st.existingKeys = new Set((existing || []).map(r => `${String(r.resolution).trim()}\u0001${r.stationId}\u0001${String(r.timestamp).trim()}`));
    } catch (e) {
      if (!st || seq !== st.checkSeq) return;
      st.existingError = e.message;
    }
    st.checking = false;
    renderPreview();
  }

  function rowsToSave(sh) {
    const b = sh.built;
    if (!b || b.errors.length) return [];
    if (!st.existingKeys || st.existingChoice === "replace") return b.rows;
    return b.rows.filter(r => !st.existingKeys.has(rowKey(sh.role, r)));
  }

  function sheetPreviewHtml(sh) {
    const b = sh.built;
    const exists = (r) => st.existingKeys?.has(rowKey(sh.role, r));
    const already = st.existingKeys ? b.rows.filter(exists).length : null;
    const skippedTotal = [...b.skipped.values()].reduce((s, x) => s + x.count, 0);
    const unit = sh.role === "Hourly" ? "hour" : sh.role === "Daily" ? "day" : "month";
    const stat = (num, label, color) => `<div class="import-stat"><b${color ? ` style="color:${color};"` : ""}>${num}</b><span>${label}</span></div>`;
    let html = `<div class="import-stats">
        ${stat(b.read, "rows read from the sheet")}
        ${stat(b.rows.length, `${sh.role.toLowerCase()} readings built`)}
        ${stat(st.checking ? "…" : already == null ? "?" : already, "already in the Sheet", already ? "var(--warn)" : "")}
        ${stat(skippedTotal, "rows skipped", skippedTotal ? "var(--err)" : "")}
      </div>`;
    const notes = [];
    notes.push(...b.errors.map(e => `<div class="callout warn" style="margin:0 0 8px;"><strong>Can't save this sheet yet</strong>${esc(e)}</div>`));
    if (b.rows.length) notes.push(`Covers <strong>${esc(b.fromDate)}</strong> to <strong>${esc(b.toDate)}</strong> (local time).`);
    if (b.combined.buckets) notes.push(`${b.combined.sourceRows} rows shared a station and ${unit}, so they were added together into ${b.combined.buckets} reading(s).`);
    for (const [other, n] of b.dupOfSheet) notes.push(`<span style="color:var(--warn);">${n} reading(s) are also on sheet “${esc(other)}” — that sheet's values are used.</span>`);
    if (b.newStations.length) {
      const names = esc(b.newStations.slice(0, 5).join(", ")) + (b.newStations.length > 5 ? ", …" : "");
      notes.push(newStationsAllowed()
        ? `${b.newStations.length} new station(s) will be added to the company: ${names}. Set their facility_id on the Dashboard so Compare and Export group them correctly.`
        : `<span style="color:var(--warn);">${b.newStations.length} station value(s) aren't in this company and will be saved as new station IDs</span> (${names}). Give them a facility_id on the Dashboard so Compare and Export group them correctly.`);
    }
    if (b.negatives) notes.push(`<span style="color:var(--warn);">${b.negatives} row(s) have a negative kWh value.</span>`);
    for (const [reason, s] of b.skipped) notes.push(`<span style="color:var(--err);">Skipped ${s.count}: ${esc(reason)}</span> <span style="opacity:.75;">— e.g. ${esc(s.samples.join(", "))}</span>`);
    html += notes.map(n => `<div style="margin-bottom:6px;">${n}</div>`).join("");

    const byStation = new Map();
    for (const r of b.rows) {
      const s = byStation.get(r.stationId) || { name: r.stationName, rows: 0, kwh: 0, first: r.date, last: r.date };
      s.rows++; s.kwh += r.kwh;
      if (r.date < s.first) s.first = r.date;
      if (r.date > s.last) s.last = r.date;
      byStation.set(r.stationId, s);
    }
    const fac = company().templateSettings?.stationFacility || {};
    if (byStation.size) {
      html += `<div class="table-scroll" style="max-height:160px;margin-top:8px;"><table><thead><tr><th>Station</th><th>facility_id</th><th>Readings</th><th>From</th><th>To</th><th>Total kWh</th></tr></thead><tbody>` +
        [...byStation.entries()].map(([id, s]) => `<tr><td>${esc(s.name)}${s.name !== id ? ` <span class="field-help">${esc(id)}</span>` : ""}</td>
          <td>${fac[id] ? esc(fac[id]) : `<span style="color:var(--warn);">not set</span>`}</td>
          <td>${s.rows}</td><td>${esc(s.first)}</td><td>${esc(s.last)}</td><td>${Math.round(s.kwh * 100) / 100}</td></tr>`).join("") + `</tbody></table></div>`;
    }
    const sample = b.rows.slice(0, PREVIEW_ROWS);
    if (sample.length) {
      html += `<div class="field-help" style="margin:10px 0 6px;">${b.rows.length > PREVIEW_ROWS ? `First ${PREVIEW_ROWS} of ${b.rows.length}` : `All ${b.rows.length}`} readings, as they'll be written to the Sheet.</div>
        <div class="table-scroll" style="max-height:220px;"><table><thead><tr><th>Timestamp</th><th>StationId</th><th>StationName</th><th>Resolution</th><th>kWh</th><th></th></tr></thead><tbody>` +
        sample.map(r => {
          const ex = exists(r);
          const tag = !st.existingKeys ? "" : !ex ? `<span class="pill ok">New</span>`
            : st.existingChoice === "replace" ? `<span class="pill warn">Replaces existing</span>` : `<span class="pill warn" style="opacity:.7;">Exists — skipped</span>`;
          return `<tr><td style="white-space:nowrap;">${esc(r.timestamp)}</td><td>${esc(r.stationId)}</td><td>${esc(r.stationName)}</td>
            <td>${esc(sh.role)}</td><td>${r.kwh}</td><td>${tag}</td></tr>`;
        }).join("") + `</tbody></table></div>`;
    }
    return html;
  }

  function renderPreview() {
    if (!st || st.step !== 3) return;
    const list = chosenSheets();
    qs("impPreviewSheets").innerHTML = list.map(sh => `<section class="import-sheet-preview">
        <h4>${esc(sh.role)} <span style="color:var(--muted);font-weight:400;">· sheet “${esc(sh.name)}”</span></h4>
        ${sh.built ? sheetPreviewHtml(sh) : ""}</section>`).join("");
    const anyExisting = st.existingKeys && list.some(sh => sh.built?.rows.some(r => st.existingKeys.has(rowKey(sh.role, r))));
    qs("impExistingWrap").hidden = !anyExisting;
    let note = "";
    if (st.checking) note = `Checking the Google Sheet for readings that already exist in this range…`;
    else if (st.existingError) note = `<span style="color:var(--warn);">Couldn't check the Sheet for existing readings (${esc(st.existingError)}).</span> Saving will still work; where a reading already exists, the newest one is used.`;
    qs("impPreviewNote").innerHTML = note;

    const blocked = list.filter(sh => sh.built?.errors.length);
    const counts = list.map(sh => ({ sh, n: rowsToSave(sh).length })).filter(x => x.n);
    const total = counts.reduce((s, x) => s + x.n, 0);
    let msg = "";
    if (blocked.length) {
      msg = `<span style="color:var(--err);">Fix ${blocked.map(sh => `“${esc(sh.name)}”`).join(", ")} first</span> — go Back to change the mapping, or set the sheet to <em>Don't import</em> on step 1.`;
    } else if (!st.checking && total) {
      msg = `${counts.map(x => `<strong style="color:var(--ink);">${x.n}</strong> ${esc(x.sh.role.toLowerCase())}`).join(" and ")} reading(s) will be added to the <code>Readings</code> tab for <strong>${esc(st.companyName.trim())}</strong>.`
        + (!st.conn ? ` <span style="color:var(--accent);">${esc(st.companyName.trim())} will be added to the Dashboard as an Excel-import company.</span>` : "");
    } else if (!st.checking && list.some(sh => sh.built?.rows.length)) {
      msg = `Every reading in this file is already in the Sheet. Choose <strong>Replace with the file's values</strong> above to overwrite them, or close.`;
    }
    qs("impSaveCount").innerHTML = msg;
    qs("btnImportNext").disabled = !!blocked.length || st.checking || !total || st.saving;
  }

  /* ---------------------------- save ---------------------------- */

  async function save() {
    const S = st; // this session, even if the modal is reopened while saving
    const list = chosenSheets().map(sh => ({ sh, rows: rowsToSave(sh) })).filter(x => x.rows.length);
    const total = list.reduce((s, x) => s + x.rows.length, 0);
    if (!total) return;
    S.saving = true;
    ["btnImportNext", "btnImportBack", "btnImportCancel"].forEach(id => { qs(id).disabled = true; });
    qs("impProgress").hidden = false;
    chosenSheets().forEach(saveMapping);

    let failure = null, saved = 0;
    const done = () => {
      S.saving = false;
      ["btnImportBack", "btnImportCancel"].forEach(id => { qs(id).disabled = false; });
    };

    // 1. A new import-only company (or new stations on one) goes into Connections first,
    //    so it's listed in Compare and Export as soon as its readings land.
    let conn = S.conn;
    const fileStations = new Map();
    list.forEach(x => x.rows.forEach(r => fileStations.set(r.stationId, r.stationName)));
    try {
      if (!conn) {
        setStatus(`Adding ${esc(S.companyName.trim())} to the Dashboard…`);
        conn = {
          companyName: S.companyName.trim(), brand: MANUAL_BRAND, credentials: {},
          stations: [...fileStations].map(([id, name]) => ({ id, name })),
          templateSettings: typeof TemplateExport !== "undefined" ? TemplateExport.defaultSettings("1") : { utcOffset: 8, stationFacility: {}, facilityGroups: {} },
          cursor: {}, source: "excel-import",
        };
        conn.id = await SheetsClient.saveConnection(conn);
        S.conn = conn;
        deps.audit?.("Add company", conn.companyName, `Excel import (no API) · ${conn.stations.length} station(s) from ${S.fileName}`);
      } else if (isManual(conn)) {
        const have = new Set((conn.stations || []).map(s => s.id));
        const add = [...fileStations].filter(([id]) => !have.has(id)).map(([id, name]) => ({ id, name }));
        if (add.length) {
          conn.stations = [...(conn.stations || []), ...add];
          await SheetsClient.saveConnection(conn);
        }
      }
    } catch (e) {
      done();
      setStatus(`Couldn't save the company to the Connections tab: ${esc(e.message)}. Nothing was imported — press Save to try again.`, "err");
      qs("btnImportNext").disabled = false;
      return;
    }

    // 2. The readings, in chunks.
    const label = brandLabel(conn);
    const all = list.flatMap(x => x.rows.map(r => ({ timestamp: r.timestamp, company: conn.companyName, brand: label,
      stationId: r.stationId, stationName: r.stationName, resolution: x.sh.role, kwh: r.kwh })));
    for (let i = 0; i < all.length; i += CHUNK) {
      setStatus(`Saving to Google Sheet… ${saved} / ${all.length}`);
      try {
        await SheetsClient.appendReadings(all.slice(i, i + CHUNK));
        saved += Math.min(CHUNK, all.length - i);
        qs("impProgressFill").style.width = `${Math.round((saved / all.length) * 100)}%`;
      } catch (e) { failure = e; break; }
    }
    done();

    const perSheet = list.map(x => `${x.sh.role} “${x.sh.name}” ${x.sh.built.fromDate} → ${x.sh.built.toDate} (${x.rows.length})`).join("; ");
    deps.audit?.(failure ? "Excel import failed" : "Excel import", conn.companyName,
      `${S.fileName} · ${perSheet} · ${saved} of ${all.length} row(s) saved`
      + (S.existingChoice === "replace" ? " · replace existing" : "") + (failure ? ` · FAILED: ${failure.message}` : ""));
    deps.onImported?.(conn);
    if (st !== S) return;

    if (failure) {
      setStatus(`Saved ${saved} of ${all.length} before an error: ${esc(failure.message)}. Press Save again to retry — readings that made it in are skipped automatically.`, "err");
      await enterPreview();
      return;
    }
    S.saved = true;
    showStep(3);
    const bySheet = list.map(x => `${x.rows.length} ${x.sh.role.toLowerCase()}`).join(" and ");
    setStatus(`✓ ${bySheet} reading(s) saved for ${esc(conn.companyName)}. They're now included in Compare and Export.`, "ok");
  }

  /* ---------------------------- navigation ---------------------------- */

  function next() {
    if (st.step === 1) {
      const before = st.conn ? st.conn.id : `new:${st.lastName || ""}`;
      st.companyName = qs("impCompany").value.trim();
      qs("impCompany").value = st.companyName;
      st.conn = findConn(st.companyName);
      st.lastName = st.companyName.toLowerCase();
      // Different company → its stations differ, so re-derive each sheet's mapping.
      if ((st.conn ? st.conn.id : `new:${st.lastName}`) !== before) st.sheets.forEach(sh => { sh.mapTouched = false; });
      if (!st.companyName) { setStatus("Type or pick a company first.", "err"); return; }
      if (!st.wb) { setStatus("Choose a file to upload.", "err"); return; }
      const undecided = st.sheets.filter(sh => sh.dataRows && !sh.role);
      if (undecided.length) { setStatus(`Choose what sheet ${undecided.map(sh => `“${esc(sh.name)}”`).join(", ")} contains — Hourly, Daily, Monthly or Don't import.`, "err"); return; }
      const list = chosenSheets();
      if (!list.length) { setStatus("Set at least one sheet to Hourly, Daily or Monthly data.", "err"); return; }
      const empty = list.filter(sh => !sh.dataRows);
      if (empty.length) { setStatus(`No data rows below the column titles on ${empty.map(sh => `“${esc(sh.name)}”`).join(", ")} — check “Titles on row”.`, "err"); return; }
      setStatus("");
      if (!list.includes(st.sheets[st.active])) st.active = st.sheets.indexOf(list[0]);
      list.forEach(prepareSheet);
      showStep(2);
      renderMapping();
    } else if (st.step === 2) {
      for (const sh of chosenSheets()) {
        const miss = missingFor(sh);
        const cols = ["timestamp", "time", "station", "kwh"].map(k => sh.map[k]).filter(v => v !== "" && !v.startsWith("fixed:"));
        const problem = miss.length ? `Map ${miss.join(", ")} first` : new Set(cols).size !== cols.length ? "The same column is mapped twice — each Readings column needs its own" : "";
        if (problem) {
          if (st.sheets[st.active] !== sh) { st.active = st.sheets.indexOf(sh); renderMapping(); }
          setStatus(`${problem}${chosenSheets().length > 1 ? ` (sheet “${esc(sh.name)}”)` : ""}.`, "err");
          return;
        }
      }
      setStatus("");
      chosenSheets().forEach(saveMapping);
      showStep(3);
      enterPreview();
    } else if (st.step === 3) {
      save();
    }
  }

  function back() {
    if (st.saving) return;
    setStatus("");
    qs("impProgress").hidden = true;
    qs("impProgressFill").style.width = "0";
    if (st.step === 3) { st.checkSeq++; st.checking = false; showStep(2); renderMapping(); }
    else if (st.step === 2) { showStep(1); renderSheets(); }
  }

  return { init, open, MANUAL_BRAND, _test: { parseDateCell, parseTimeCell, parseNumber, detectDateOrder, stationMatcher, guessRole } };
})();

if (typeof module !== "undefined") module.exports = ExcelImport;
