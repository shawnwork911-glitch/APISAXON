/* =====================================================================
   Excel import — generation data from a spreadsheet into Readings
   ---------------------------------------------------------------------
   Lets someone upload an .xlsx / .xls / .csv of generation data (from a
   monitoring portal, a meter export, a raw download from this app…),
   map its column titles to the Readings columns, preview the result,
   and append it to the Google Sheet's Readings tab. From then on it is
   ordinary Readings data: Compare and Export pick it up exactly like
   rows written by an Extraction run.

   OUTPUT SHAPE — identical to what the Extraction job writes:
     { timestamp: "YYYY-MM-DD HH:MM:SS" (local wall clock),
       company, brand (label), stationId, stationName, resolution, kwh }
   Timestamps are bucketed to the chosen resolution, the same way the
   APIs report them:
     Hourly  → "YYYY-MM-DD HH:00:00"
     Daily   → "YYYY-MM-DD 00:00:00"
     Monthly → "YYYY-MM-01 00:00:00"
   Several file rows landing in the same station + period (15-minute
   data, or hourly data imported as Monthly) are SUMMED into one row, so
   the de-dupe rule in app.js / the proxy (one row per Resolution +
   StationId + Timestamp, newest RunAt wins) never drops part of a total.

   STATIONS
     The Station column is matched against the company's station IDs,
     then station names, then facility_id (only when that facility has
     exactly one station — a multi-station facility can't be split).
     A whole file can also be assigned to one station.

   EXISTING ROWS
     Before saving, the Sheet is read for the file's date range. By
     default rows that already exist are left alone (only new ones are
     added); "Replace" appends them anyway, and because the newest row
     wins, the file's values then take over in Compare and Export.
     Either way a re-run after a partial failure is safe.

   Column mappings are remembered per company + resolution in this
   browser (localStorage), so a repeat upload of the same layout maps
   itself.
   ===================================================================== */

const ExcelImport = (() => {
  const STORE_KEY = "slc.importMappings";
  const CHUNK = 1000;          // rows per appendReadings call
  const PREVIEW_ROWS = 15;
  const MONTHS = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];

  const FIELDS = [
    { key: "timestamp", label: "Timestamp", required: true, help: "Date, or date and time, of each reading" },
    { key: "time", label: "Time", required: false, help: "Only if the time of day is in a separate column" },
    { key: "station", label: "Station", required: true, help: "Station ID, station name or facility_id — or one station for the whole file" },
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

  /* ---------------------------- mapping memory ---------------------------- */

  function loadStore() { try { return JSON.parse(localStorage.getItem(STORE_KEY) || "{}"); } catch { return {}; } }
  function saveMapping() {
    try {
      const store = loadStore();
      const byName = {};
      for (const f of FIELDS) {
        const v = st.map[f.key];
        byName[f.key] = v === "" || v.startsWith("fixed:") ? v : (st.headers[+v] ?? "");
      }
      store[`${st.conn.id}|${st.resolution}`] = { map: byName, opt: { unit: st.opt.unit, dateOrder: st.opt.dateOrder, tz: st.opt.tz, hourEnding: st.opt.hourEnding, unknown: st.opt.unknown } };
      localStorage.setItem(STORE_KEY, JSON.stringify(store));
    } catch { /* storage full or blocked — just don't remember */ }
  }
  // A saved mapping only applies if every column it names exists in this file.
  function savedMappingFor() {
    const saved = loadStore()[`${st.conn.id}|${st.resolution}`];
    if (!saved) return null;
    const map = {};
    for (const f of FIELDS) {
      const v = saved.map?.[f.key] ?? "";
      if (v === "") { map[f.key] = ""; continue; }
      if (v.startsWith("fixed:")) {
        if (!(st.conn.stations || []).some(s => `fixed:${s.id}` === v)) return null;
        map[f.key] = v; continue;
      }
      const idx = st.headers.findIndex(h => h === v);
      if (idx < 0) return null;
      map[f.key] = String(idx);
    }
    return { map, opt: saved.opt || {} };
  }

  function guessMapping() {
    const H = st.headers.map(norm);
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
    // A separate Time column only when the Timestamp column is a plain date.
    if (map.timestamp !== "" && /date/.test(H[+map.timestamp]) && !/time/.test(H[+map.timestamp])) {
      const i = H.findIndex((h, idx) => GUESS.time.includes(h) && !used.has(idx));
      if (i >= 0) map.time = String(i);
    }
    if (map.station === "" && (st.conn.stations || []).length === 1) map.station = `fixed:${st.conn.stations[0].id}`;
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
    qs("impSheet").addEventListener("change", () => { loadSheet(qs("impSheet").value); });
    qs("impHeaderRow").addEventListener("change", () => {
      const n = Math.max(1, Math.floor(+qs("impHeaderRow").value || 1));
      qs("impHeaderRow").value = n;
      setHeaderRow(n - 1);
    });
    qs("impCompany").addEventListener("change", () => { st.conn = currentConn(); st.mapTouched = false; renderFileInfo(); });
    qs("impResolution").addEventListener("change", () => { st.resolution = qs("impResolution").value; st.mapTouched = false; });
    qs("impExisting").addEventListener("change", () => { st.opt.existing = qs("impExisting").value; renderPreview(); });
    qs("importModalOverlay").addEventListener("click", (e) => { if (e.target.id === "importModalOverlay" && !st?.saving) close(); });
  }

  function currentConn() {
    return (deps.getConnections?.() || []).find(c => c.id === qs("impCompany").value) || null;
  }

  // opts: { companyId, resolution } — pre-selected from whichever tab opened it.
  function open(opts = {}) {
    if (st?.saving) { qs("importModalOverlay").classList.add("active"); return; } // a save is still running — show it
    const conns = deps.getConnections?.() || [];
    st = {
      step: 1, conn: null, resolution: ["Hourly", "Daily", "Monthly"].includes(opts.resolution) ? opts.resolution : "Hourly",
      fileName: "", wb: null, sheetName: "", aoa: [], aoaText: [], headerIdx: 0, headers: [],
      map: { timestamp: "", time: "", station: "", kwh: "" }, mapTouched: false,
      opt: { unit: "kWh", dateOrder: "auto", tz: "local", hourEnding: false, unknown: "skip", existing: "keep" },
      built: null, existingKeys: null, existingError: "", saving: false, saved: false,
    };
    const sel = qs("impCompany");
    sel.innerHTML = `<option value="">Select a company…</option>` + conns.map(c =>
      `<option value="${esc(c.id)}">${esc(c.companyName)} (${esc(brandLabel(c))})</option>`).join("");
    if (opts.companyId && conns.some(c => c.id === opts.companyId)) sel.value = opts.companyId;
    st.conn = currentConn();
    qs("impResolution").value = st.resolution;
    qs("impFile").value = "";
    qs("impSheetWrap").hidden = true;
    qs("impHeaderRowWrap").hidden = true;
    qs("impFileInfo").innerHTML = "";
    qs("impExisting").value = "keep";
    qs("impProgress").hidden = true;
    qs("impProgressFill").style.width = "0";
    setStatus("");
    showStep(1);
    qs("importModalOverlay").classList.add("active");
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
      li.classList.toggle("done", i + 1 < n);
    });
    qs("btnImportBack").hidden = n === 1 || st.saved;
    qs("btnImportCancel").textContent = st.saved ? "Close" : "Cancel";
    const nextBtn = qs("btnImportNext");
    nextBtn.hidden = st.saved;
    nextBtn.disabled = false;
    nextBtn.textContent = n === 1 ? "Next: map columns" : n === 2 ? "Next: preview" : "Save to Google Sheet";
  }

  /* ---------------------------- step 1: file ---------------------------- */

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
        const names = st.wb.SheetNames;
        qs("impSheet").innerHTML = names.map(n => `<option value="${esc(n)}">${esc(n)}</option>`).join("");
        qs("impSheetWrap").hidden = names.length < 2;
        // Default to the first sheet that actually has some rows.
        const first = names.find(n => (XLSX.utils.sheet_to_json(st.wb.Sheets[n], { header: 1, blankrows: false }) || []).length > 1) || names[0];
        qs("impSheet").value = first;
        loadSheet(first);
        setStatus("");
      } catch (e) {
        setStatus(`That file couldn't be opened as a spreadsheet (${esc(e.message)}).`, "err");
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function loadSheet(name) {
    const ws = st.wb.Sheets[name];
    st.sheetName = name;
    st.aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "", blankrows: true });
    st.aoaText = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: "", blankrows: true });
    // Header row = first of the top 20 rows with 2+ filled cells, mostly text.
    let idx = 0;
    for (let i = 0; i < Math.min(20, st.aoa.length); i++) {
      const cells = (st.aoa[i] || []).filter(c => !isBlank(c));
      const textCells = cells.filter(c => typeof c === "string" && !/^[\d.,\-/: ]+$/.test(c.trim()));
      if (cells.length >= 2 && textCells.length >= Math.ceil(cells.length / 2)) { idx = i; break; }
    }
    qs("impHeaderRowWrap").hidden = false;
    qs("impHeaderRow").value = idx + 1;
    setHeaderRow(idx);
  }

  function setHeaderRow(idx) {
    st.headerIdx = idx;
    const width = st.aoa.reduce((w, r) => Math.max(w, (r || []).length), 0);
    const row = st.aoa[idx] || [];
    st.headers = Array.from({ length: width }, (_, i) => String(row[i] ?? "").trim() || `Column ${colLetter(i)}`);
    st.mapTouched = false;
    renderFileInfo();
  }

  function dataRowCount() {
    let n = 0;
    for (let i = st.headerIdx + 1; i < st.aoa.length; i++) if ((st.aoa[i] || []).some(c => !isBlank(c))) n++;
    return n;
  }

  function renderFileInfo() {
    const box = qs("impFileInfo");
    if (!st?.wb) { box.innerHTML = ""; return; }
    const rows = dataRowCount();
    const noStations = st.conn && !(st.conn.stations || []).length;
    box.innerHTML = `<div class="field-help" style="margin-top:0;">
        <strong style="color:var(--ink);">${esc(st.fileName)}</strong>${st.wb.SheetNames.length > 1 ? `, sheet “${esc(st.sheetName)}”` : ""}:
        ${rows} data row(s), ${st.headers.length} column(s). Column titles found on row ${st.headerIdx + 1}:
        ${st.headers.map(h => `<code>${esc(h)}</code>`).join(" ")}
      </div>`
      + (noStations ? `<div class="callout warn" style="margin-bottom:0;"><strong>No stations cached for this company</strong>Station values can't be matched until the station list is fetched (Dashboard → Test &amp; Connect). You can still import them as new station IDs in the next step.</div>` : "");
  }

  /* ---------------------------- step 2: mapping ---------------------------- */

  function enterMapping() {
    if (!st.mapTouched) {
      const g = savedMappingFor() || guessMapping();
      st.map = { timestamp: "", time: "", station: "", kwh: "", ...g.map };
      Object.assign(st.opt, g.opt);
      st.fromSaved = !!savedMappingFor();
    }
    const tsCol = st.map.timestamp === "" ? -1 : +st.map.timestamp;
    st.detectedOrder = detectDateOrder(tsCol < 0 ? [] : st.aoa.slice(st.headerIdx + 1).map(r => r?.[tsCol]));
    renderMapping();
  }

  function columnOptions(selected, noneLabel) {
    return `<option value="">${esc(noneLabel)}</option>` + st.headers.map((h, i) =>
      `<option value="${i}"${String(i) === selected ? " selected" : ""}>${esc(colLetter(i))} · ${esc(h)}</option>`).join("");
  }

  function renderMapping() {
    const stations = st.conn.stations || [];
    const rows = FIELDS.map(f => {
      let control;
      if (f.key === "station") {
        control = `<select class="impMap" data-field="station">
            ${columnOptions(st.map.station.startsWith("fixed:") ? "" : st.map.station, "— choose a column —")}
            ${stations.length ? `<optgroup label="Whole file is one station">${stations.map(s =>
              `<option value="fixed:${esc(s.id)}"${st.map.station === `fixed:${s.id}` ? " selected" : ""}>${esc(s.name || s.id)}</option>`).join("")}</optgroup>` : ""}
          </select>`;
      } else {
        control = `<select class="impMap" data-field="${f.key}">${columnOptions(st.map[f.key], f.required ? "— choose a column —" : "— none —")}</select>`;
      }
      return `<tr>
          <td style="white-space:nowrap;"><strong>${esc(f.label)}</strong>${f.required ? ` <span style="color:var(--accent);">*</span>` : ""}
            <div class="field-help" style="white-space:normal;max-width:220px;">${esc(f.help)}</div></td>
          <td style="min-width:220px;">${control}</td>
          <td class="map-sample" id="impSample-${f.key}">${sampleHtml(f.key)}</td>
        </tr>`;
    }).join("");
    qs("impMapTable").innerHTML = `<thead><tr><th>Readings column</th><th>Column in your file</th><th>First values → how they'll be saved</th></tr></thead><tbody>${rows}</tbody>`;
    qs("impMapTable").querySelectorAll(".impMap").forEach(sel => sel.addEventListener("change", () => {
      st.map[sel.dataset.field] = sel.value;
      st.mapTouched = true;
      if (sel.dataset.field === "timestamp") {
        st.detectedOrder = detectDateOrder(sel.value === "" ? [] : st.aoa.slice(st.headerIdx + 1).map(r => r?.[+sel.value]));
        renderOptions();
      }
      if (sel.dataset.field === "kwh" && sel.value !== "") {
        const h = norm(st.headers[+sel.value]);
        const unit = /mwh/.test(h) ? "MWh" : /kwh/.test(h) ? "kWh" : /wh$/.test(h) ? "Wh" : null;
        if (unit) { st.opt.unit = unit; renderOptions(); }
      }
      refreshSamples();
    }));
    qs("impMapNote").innerHTML = st.fromSaved
      ? `Filled in from the mapping you last used for <strong>${esc(st.conn.companyName)}</strong> (${esc(st.resolution)}). Check it still fits this file.`
      : `Columns were matched by their titles. Check each one against the sample values on the right.`;
    renderOptions();
  }

  function renderOptions() {
    const off = Number(st.conn.templateSettings?.utcOffset ?? 8);
    const offLabel = `UTC${off >= 0 ? "+" : ""}${off}`;
    const opt = (v, label, cur) => `<option value="${esc(v)}"${v === cur ? " selected" : ""}>${esc(label)}</option>`;
    const orderName = { DMY: "day first", MDY: "month first" }[st.detectedOrder];
    qs("impOptions").innerHTML = `
      <div class="field"><label>Unit of the kWh column</label>
        <select data-opt="unit">${opt("Wh", "Wh", st.opt.unit)}${opt("kWh", "kWh", st.opt.unit)}${opt("MWh", "MWh", st.opt.unit)}</select></div>
      <div class="field"><label>Dates like 03/04/2026</label>
        <select data-opt="dateOrder">${opt("auto", `Detect (looks ${orderName})`, st.opt.dateOrder)}${opt("DMY", "Day first — 3 April", st.opt.dateOrder)}${opt("MDY", "Month first — 4 March", st.opt.dateOrder)}</select></div>
      <div class="field"><label>Times in the file are</label>
        <select data-opt="tz">${opt("local", `Local time (${offLabel}, from company settings)`, st.opt.tz)}${opt("utc", `UTC — shift ${off >= 0 ? "+" : ""}${off} h to local`, st.opt.tz)}</select></div>
      <div class="field"${st.resolution === "Hourly" ? "" : " hidden"}><label>Each hour is labelled by its</label>
        <select data-opt="hourEnding">${opt("start", "Start — 00:00 means 00:00–01:00", st.opt.hourEnding ? "end" : "start")}${opt("end", "End — 01:00 means 00:00–01:00", st.opt.hourEnding ? "end" : "start")}</select></div>
      <div class="field"><label>Station values not in this company</label>
        <select data-opt="unknown">${opt("skip", "Skip those rows", st.opt.unknown)}${opt("import", "Import as new station IDs", st.opt.unknown)}</select></div>`;
    qs("impOptions").querySelectorAll("select[data-opt]").forEach(sel => sel.addEventListener("change", () => {
      const k = sel.dataset.opt;
      st.opt[k] = k === "hourEnding" ? sel.value === "end" : sel.value;
      st.mapTouched = true;
      refreshSamples();
    }));
  }

  function refreshSamples() {
    FIELDS.forEach(f => { const el = qs(`impSample-${f.key}`); if (el) el.innerHTML = sampleHtml(f.key); });
  }

  // First few values of the mapped column, each with what it turns into.
  function sampleHtml(field) {
    const v = st.map[field];
    if (field === "station" && v.startsWith("fixed:")) {
      const s = (st.conn.stations || []).find(x => `fixed:${x.id}` === v);
      return `Every row → <span class="good">${esc(s?.name || s?.id || "")}</span> <span style="opacity:.7;">(${esc(s?.id || "")})</span>`;
    }
    if (v === "") return field === "time" ? `Time is taken from the Timestamp column.` : `<span class="bad">Not mapped yet</span>`;
    const col = +v, out = [];
    const order = st.opt.dateOrder === "auto" ? st.detectedOrder : st.opt.dateOrder;
    const match = field === "station" ? stationMatcher(st.conn) : null;
    for (let i = st.headerIdx + 1; i < st.aoa.length && out.length < 3; i++) {
      const raw = st.aoa[i]?.[col], text = st.aoaText[i]?.[col];
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
          : st.opt.unknown === "import" ? `<span style="color:var(--warn);">new station</span>` : `<span class="bad">not found</span>`;
      } else {
        const n = parseNumber(raw);
        const f = { Wh: 0.001, kWh: 1, MWh: 1000 }[st.opt.unit];
        res = Number.isFinite(n) ? `<span class="good">${+(n * f).toFixed(6)} kWh</span>` : `<span class="bad">not a number</span>`;
      }
      out.push(`<div><span style="color:var(--ink);">${esc(text ?? raw)}</span> → ${res}</div>`);
    }
    return out.join("") || `<span class="bad">Column is empty</span>`;
  }

  /* ---------------------------- build rows ---------------------------- */

  function buildRows() {
    const { conn, resolution, map, opt } = st;
    const off = Number(conn.templateSettings?.utcOffset ?? 8);
    const order = opt.dateOrder === "auto" ? st.detectedOrder : opt.dateOrder;
    const factor = { Wh: 0.001, kWh: 1, MWh: 1000 }[opt.unit] || 1;
    const match = stationMatcher(conn);
    const fixed = map.station.startsWith("fixed:") ? (conn.stations || []).find(s => `fixed:${s.id}` === map.station) : null;
    const tsCol = +map.timestamp, timeCol = map.time === "" ? -1 : +map.time;
    const stCol = fixed ? -1 : +map.station, kwhCol = +map.kwh;

    const skipped = new Map(); // reason -> { count, samples: [] }
    const skip = (reason, rowNo, value) => {
      const s = skipped.get(reason) || { count: 0, samples: [] };
      s.count++;
      if (s.samples.length < 3) s.samples.push(`row ${rowNo}${isBlank(value) ? "" : ` (“${String(value).slice(0, 40)}”)`}`);
      skipped.set(reason, s);
    };
    const agg = new Map();
    let read = 0, anyTime = false, negatives = 0;
    const newStations = new Set();

    for (let i = st.headerIdx + 1; i < st.aoa.length; i++) {
      const row = st.aoa[i] || [];
      if (row.every(isBlank)) continue;
      read++;
      const rowNo = i + 1;
      let wc = parseDateCell(row[tsCol], order);
      if (!wc) { skip("Date/time not recognised", rowNo, st.aoaText[i]?.[tsCol]); continue; }
      if (timeCol >= 0) {
        const t = parseTimeCell(row[timeCol]);
        if (!t) { skip("Time not recognised", rowNo, st.aoaText[i]?.[timeCol]); continue; }
        wc = { ...wc, H: t.H, Mi: t.Mi, S: t.S, hasTime: true };
      }
      const n = parseNumber(row[kwhCol]);
      if (n === null) { skip("Blank kWh value", rowNo); continue; }
      if (Number.isNaN(n)) { skip("kWh value is not a number", rowNo, st.aoaText[i]?.[kwhCol]); continue; }

      let stationId, stationName;
      if (fixed) { stationId = fixed.id; stationName = fixed.name || fixed.id; }
      else {
        const m = match(row[stCol]);
        if (m.station) { stationId = m.station.id; stationName = m.station.name || m.station.id; }
        else if (m.error === "blank") { skip("Blank station", rowNo); continue; }
        else if (m.error === "multiFacility") { skip(`facility_id covers ${m.count} stations — can't tell which one`, rowNo, row[stCol]); continue; }
        else if (m.error === "ambiguous") { skip("Station name shared by two stations", rowNo, row[stCol]); continue; }
        else if (opt.unknown === "import") { stationId = stationName = String(row[stCol]).trim(); newStations.add(stationId); }
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
      errors.push("No time of day was found in any row, so every reading would land at 00:00. Map a Time column, or import this file as Daily or Monthly.");
    }
    return {
      read, rows, skipped, negatives, newStations: [...newStations],
      combined: { buckets: combinedRows.length, sourceRows: combinedRows.reduce((s, r) => s + r.n, 0) },
      fromDate: rows.length ? rows.reduce((m, r) => (r.date < m ? r.date : m), rows[0].date) : "",
      toDate: rows.length ? rows.reduce((m, r) => (r.date > m ? r.date : m), rows[0].date) : "",
      errors,
    };
  }

  /* ---------------------------- step 3: preview ---------------------------- */

  async function enterPreview() {
    st.built = buildRows();
    st.existingKeys = null;
    st.existingError = "";
    renderPreview();
    if (st.built.errors.length) return;
    // Which of these rows are already in the Sheet?
    const lastDay = (iso) => { const [y, m] = iso.split("-").map(Number); return `${y}-${pad(m)}-${pad(daysIn(y, m))}`; };
    const seq = (st.checkSeq = (st.checkSeq || 0) + 1);
    st.checking = true;
    renderPreview();
    try {
      const existing = await SheetsClient.listReadings(st.conn.companyName, {
        fromDate: `${st.built.fromDate.slice(0, 7)}-01`, toDate: lastDay(st.built.toDate) });
      if (!st || seq !== st.checkSeq) return;
      st.existingKeys = new Set((existing || [])
        .filter(r => String(r.resolution).trim() === st.resolution)
        .map(r => `${r.stationId}\u0001${String(r.timestamp).trim()}`));
    } catch (e) {
      if (!st || seq !== st.checkSeq) return;
      st.existingError = e.message;
    }
    st.checking = false;
    renderPreview();
  }

  function rowsToSave() {
    const b = st.built;
    if (!b) return [];
    if (!st.existingKeys || st.opt.existing === "replace") return b.rows;
    return b.rows.filter(r => !st.existingKeys.has(`${r.stationId}\u0001${r.timestamp}`));
  }

  function renderPreview() {
    if (!st?.built) return;
    const b = st.built;
    const exists = (r) => st.existingKeys?.has(`${r.stationId}\u0001${r.timestamp}`);
    const already = st.existingKeys ? b.rows.filter(exists).length : null;
    const toSave = rowsToSave();
    const skippedTotal = [...b.skipped.values()].reduce((s, x) => s + x.count, 0);

    const stat = (num, label, color) => `<div class="import-stat"><b${color ? ` style="color:${color};"` : ""}>${num}</b><span>${label}</span></div>`;
    let html = `<div class="import-stats">
        ${stat(b.read, "rows read from the file")}
        ${stat(b.rows.length, `${st.resolution.toLowerCase()} readings built`)}
        ${stat(st.checking ? "…" : already == null ? "?" : already, "already in the Sheet", already ? "var(--warn)" : "")}
        ${stat(skippedTotal, "rows skipped", skippedTotal ? "var(--err)" : "")}
      </div>`;

    const notes = [];
    if (b.errors.length) notes.push(...b.errors.map(e => `<div class="callout warn" style="margin:0 0 8px;"><strong>Can't save yet</strong>${esc(e)}</div>`));
    if (b.rows.length) notes.push(`Covers <strong>${esc(b.fromDate)}</strong> to <strong>${esc(b.toDate)}</strong> (local time, ${esc(st.resolution)}).`);
    if (b.combined.buckets) notes.push(`${b.combined.sourceRows} file rows shared a station and ${st.resolution === "Hourly" ? "hour" : st.resolution === "Daily" ? "day" : "month"}, so they were added together into ${b.combined.buckets} reading(s).`);
    if (b.newStations.length) notes.push(`<span style="color:var(--warn);">${b.newStations.length} station value(s) aren't in this company and will be saved as new station IDs</span> (${esc(b.newStations.slice(0, 5).join(", "))}${b.newStations.length > 5 ? ", …" : ""}). Give them a facility_id on the Dashboard so Compare and Export group them correctly.`);
    if (b.negatives) notes.push(`<span style="color:var(--warn);">${b.negatives} row(s) have a negative kWh value.</span>`);
    for (const [reason, s] of b.skipped) notes.push(`<span style="color:var(--err);">Skipped ${s.count}: ${esc(reason)}</span> <span style="opacity:.75;">— e.g. ${esc(s.samples.join(", "))}</span>`);
    if (st.checking) notes.push(`Checking the Google Sheet for readings that already exist in this range…`);
    if (st.existingError) notes.push(`<span style="color:var(--warn);">Couldn't check the Sheet for existing readings (${esc(st.existingError)}).</span> Saving will still work; where a reading already exists, the newest one is used.`);
    html += notes.map(n => `<div style="margin-bottom:6px;">${n}</div>`).join("");
    qs("impSummary").innerHTML = html;

    qs("impExistingWrap").hidden = !already;
    qs("impSaveCount").innerHTML = b.errors.length ? "" : st.checking ? "" :
      `<strong style="color:var(--ink);">${toSave.length}</strong> reading(s) will be added to the <code>Readings</code> tab for <strong>${esc(st.conn.companyName)}</strong>.`;

    // Per-station summary
    const byStation = new Map();
    for (const r of b.rows) {
      const s = byStation.get(r.stationId) || { name: r.stationName, rows: 0, kwh: 0, first: r.date, last: r.date };
      s.rows++; s.kwh += r.kwh;
      if (r.date < s.first) s.first = r.date;
      if (r.date > s.last) s.last = r.date;
      byStation.set(r.stationId, s);
    }
    const fac = st.conn.templateSettings?.stationFacility || {};
    qs("impStationTable").innerHTML = byStation.size
      ? `<thead><tr><th>Station</th><th>facility_id</th><th>Readings</th><th>From</th><th>To</th><th>Total kWh</th></tr></thead><tbody>` +
        [...byStation.entries()].map(([id, s]) => `<tr><td>${esc(s.name)}${s.name !== id ? ` <span class="field-help">${esc(id)}</span>` : ""}</td>
          <td>${esc(fac[id] || "")}${fac[id] ? "" : `<span style="color:var(--warn);">not set</span>`}</td>
          <td>${s.rows}</td><td>${esc(s.first)}</td><td>${esc(s.last)}</td><td>${Math.round(s.kwh * 100) / 100}</td></tr>`).join("") + `</tbody>`
      : "";

    // First rows, exactly as they'll be written
    const sample = b.rows.slice(0, PREVIEW_ROWS);
    qs("impPreviewTable").innerHTML = sample.length
      ? `<thead><tr><th>Timestamp</th><th>StationId</th><th>StationName</th><th>Resolution</th><th>kWh</th><th></th></tr></thead><tbody>` +
        sample.map(r => {
          const ex = exists(r);
          const tag = !st.existingKeys ? "" : !ex ? `<span class="pill ok">New</span>`
            : st.opt.existing === "replace" ? `<span class="pill warn">Replaces existing</span>` : `<span class="pill warn" style="opacity:.7;">Exists — skipped</span>`;
          return `<tr><td style="white-space:nowrap;">${esc(r.timestamp)}</td><td>${esc(r.stationId)}</td><td>${esc(r.stationName)}</td>
            <td>${esc(st.resolution)}</td><td>${r.kwh}</td><td>${tag}</td></tr>`;
        }).join("") + `</tbody>`
      : "";
    qs("impPreviewCaption").textContent = b.rows.length > PREVIEW_ROWS ? `First ${PREVIEW_ROWS} of ${b.rows.length} readings, as they'll be written to the Sheet.` : b.rows.length ? `All ${b.rows.length} readings, as they'll be written to the Sheet.` : "";

    const nextBtn = qs("btnImportNext");
    nextBtn.disabled = !!b.errors.length || !!st.checking || !toSave.length || st.saving;
    if (!b.errors.length && !st.checking && !toSave.length && b.rows.length) {
      qs("impSaveCount").innerHTML = `Every reading in this file is already in the Sheet. Choose <strong>Replace with the file's values</strong> above to overwrite them, or close.`;
    }
  }

  /* ---------------------------- save ---------------------------- */

  async function save() {
    const S = st; // this session, even if the modal is reopened while saving
    const rows = rowsToSave();
    if (!rows.length) return;
    const conn = S.conn, resolution = S.resolution, label = brandLabel(conn);
    S.saving = true;
    qs("btnImportNext").disabled = true;
    qs("btnImportBack").disabled = true;
    qs("btnImportCancel").disabled = true;
    qs("impProgress").hidden = false;
    saveMapping();

    let saved = 0, failure = null;
    for (let i = 0; i < rows.length; i += CHUNK) {
      const chunk = rows.slice(i, i + CHUNK);
      setStatus(`Saving to Google Sheet… ${saved} / ${rows.length}`);
      try {
        await SheetsClient.appendReadings(chunk.map(r => ({
          timestamp: r.timestamp, company: conn.companyName, brand: label,
          stationId: r.stationId, stationName: r.stationName, resolution, kwh: r.kwh,
        })));
        saved += chunk.length;
        qs("impProgressFill").style.width = `${Math.round((saved / rows.length) * 100)}%`;
      } catch (e) { failure = e; break; }
    }

    S.saving = false;
    qs("btnImportBack").disabled = false;
    qs("btnImportCancel").disabled = false;
    const details = `${resolution} · ${S.fileName}${S.wb.SheetNames.length > 1 ? ` [${S.sheetName}]` : ""} · ${S.built.fromDate} → ${S.built.toDate} · ${saved} of ${rows.length} row(s) saved`
      + (S.opt.existing === "replace" ? " · replace existing" : "")
      + (failure ? ` · FAILED: ${failure.message}` : "");
    deps.audit?.(failure ? "Excel import failed" : "Excel import", conn.companyName, details);

    if (failure) {
      setStatus(`Saved ${saved} of ${rows.length} before an error: ${esc(failure.message)}. Press Save again to retry — readings that made it in are skipped automatically.`, "err");
      if (st !== S) return;
      if (saved) await enterPreview(); else renderPreview();
      return;
    }
    if (st !== S) { deps.onImported?.(conn, resolution); return; }
    S.saved = true;
    showStep(3);
    setStatus(`✓ ${saved} ${resolution} reading(s) saved to the Readings tab. They're now included in Compare and Export for ${esc(conn.companyName)}.`, "ok");
    deps.onImported?.(conn, resolution);
  }

  /* ---------------------------- navigation ---------------------------- */

  function next() {
    if (st.step === 1) {
      st.conn = currentConn();
      st.resolution = qs("impResolution").value;
      if (!st.conn) { setStatus("Pick a company first.", "err"); return; }
      if (!st.wb) { setStatus("Choose a file to upload.", "err"); return; }
      if (!dataRowCount()) { setStatus("No data rows below the column-title row — check the header row number.", "err"); return; }
      setStatus("");
      showStep(2);
      enterMapping();
    } else if (st.step === 2) {
      const missing = FIELDS.filter(f => f.required && st.map[f.key] === "").map(f => f.label);
      if (missing.length) { setStatus(`Map ${missing.join(", ")} first.`, "err"); return; }
      const cols = ["timestamp", "time", "station", "kwh"].map(k => st.map[k]).filter(v => v !== "" && !v.startsWith("fixed:"));
      if (new Set(cols).size !== cols.length) { setStatus("The same file column is mapped twice — each Readings column needs its own.", "err"); return; }
      setStatus("");
      saveMapping();
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
    if (st.step === 3) { st.checkSeq = (st.checkSeq || 0) + 1; st.checking = false; showStep(2); renderMapping(); }
    else if (st.step === 2) showStep(1);
  }

  return { init, open, _test: { parseDateCell, parseTimeCell, parseNumber, detectDateOrder, stationMatcher } };
})();

if (typeof module !== "undefined") module.exports = ExcelImport;
