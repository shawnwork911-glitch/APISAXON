/* =====================================================================
   SolarLink Connect — App shell
   ===================================================================== */

// Baked-in defaults so anyone opening this page (not just whoever first
// configured it) gets a working connection without typing anything in.
// Neither value is sensitive — the actual credential (the Google service
// account's private key) lives only in the deployed proxy, never here.
// Settings can still override these; whatever's saved in this browser's
// localStorage always wins over these defaults.
const DEFAULT_PROXY_BASE_URL = "https://solarlink-cors-proxy.saxon-solarlink.workers.dev";
const DEFAULT_SPREADSHEET_ID = "1Sq2AYWWDOBvlrrvcesw3bxrJHWE0xe-m5i_qkG46BCY";

const App = (() => {
  let connections = [];      // cached from Google Sheets (or local draft before first sync)
  let activeAuthByConn = {}; // connectionId -> { base, headers, ... } from buildAuth()
  let currentView = "dashboard";
  let pendingBrand = null;   // brand selected inside Add Company modal

  const els = {};

  function qs(id) { return document.getElementById(id); }

  async function boot() {
    cacheEls();
    wireStaticEvents();

    if (!ProxyClient.getProxyBaseUrl() && DEFAULT_PROXY_BASE_URL) {
      ProxyClient.setProxyBaseUrl(DEFAULT_PROXY_BASE_URL);
    }
    let cfg = SheetsClient.loadConfig();
    if (!cfg?.spreadsheetId && DEFAULT_SPREADSHEET_ID) {
      cfg = { ...(cfg || {}), spreadsheetId: DEFAULT_SPREADSHEET_ID };
      SheetsClient.saveConfig(cfg);
    }

    qs("cfgSpreadsheetId").value = cfg?.spreadsheetId || "";
    qs("cfgProxyUrl").value = ProxyClient.getProxyBaseUrl();
    updateConnectionBanner();
    if (SheetsClient.isConfigured()) await refreshConnections();
    renderBrandGrid();
    showView("dashboard");
  }

  function cacheEls() {
    ["dashboard", "extraction", "settings"].forEach(v => els[v] = qs(`view-${v}`));
  }

  function wireStaticEvents() {
    document.querySelectorAll(".nav-btn[data-view]").forEach(btn => {
      btn.addEventListener("click", () => showView(btn.dataset.view));
    });
    qs("btnAddCompany").addEventListener("click", openAddCompanyModal);
    qs("btnModalClose").addEventListener("click", closeAddCompanyModal);
    qs("btnStationsModalClose").addEventListener("click", closeStationsModal);
    qs("stationsModalOverlay").addEventListener("click", (e) => { if (e.target.id === "stationsModalOverlay") closeStationsModal(); });
    qs("btnModalCancel").addEventListener("click", closeAddCompanyModal);
    qs("btnSaveWithoutTest").addEventListener("click", () => saveConnectionFromModal(false));
    qs("btnTestConnect").addEventListener("click", () => saveConnectionFromModal(true));
    qs("btnSaveSettings").addEventListener("click", handleSaveSettings);
    qs("btnQueueExtraction").addEventListener("click", handleQueueExtraction);
    qs("extCompany").addEventListener("change", handleExtractionCompanyChange);
  }

  function showView(view) {
    currentView = view;
    Object.entries(els).forEach(([k, el]) => el.classList.toggle("active", k === view));
    document.querySelectorAll(".nav-btn[data-view]").forEach(b => b.classList.toggle("active", b.dataset.view === view));
    if (view === "extraction") populateExtractionCompanySelect();
  }

  /* ---------------------------- settings / connection status ---------------------------- */

  function updateConnectionBanner() {
    const banner = qs("connectionBanner");
    const configured = SheetsClient.isConfigured();
    banner.style.display = configured ? "none" : "flex";
    banner.classList.remove("banner-error");
    banner.textContent = "Set your Proxy base URL and Google Sheet ID under Settings to load and save company connections.";
    const sheetLink = qs("sheetLink");
    const url = SheetsClient.sheetUrl();
    if (sheetLink) sheetLink.href = url || "#";
    if (sheetLink) sheetLink.style.display = url ? "inline" : "none";
  }

  async function handleSaveSettings() {
    SheetsClient.saveConfig({ spreadsheetId: qs("cfgSpreadsheetId").value.trim() });
    ProxyClient.setProxyBaseUrl(qs("cfgProxyUrl").value.trim());
    updateConnectionBanner();
    qs("settingsSavedNote").textContent = "Saved.";
    qs("settingsSavedNote").style.color = "var(--ok)";
    if (SheetsClient.isConfigured()) {
      try {
        await refreshConnections();
        qs("settingsSavedNote").textContent = "Saved — connected to Google Sheet.";
      } catch (e) {
        qs("settingsSavedNote").textContent = `Saved, but: ${e.message}`;
        qs("settingsSavedNote").style.color = "var(--err)";
      }
    }
    setTimeout(() => (qs("settingsSavedNote").textContent = ""), 6000);
  }

  async function refreshConnections() {
    try {
      connections = await SheetsClient.listConnections();
    } catch (e) {
      console.warn("Could not load connections from Google Sheets:", e.message);
      connections = connections || [];
    }
    renderDashboard();
  }

  /* ---------------------------- dashboard ---------------------------- */

  function renderDashboard() {
    const root = qs("companyList");
    root.innerHTML = "";
    if (!connections.length) {
      root.innerHTML = `<div class="empty-state">No companies connected yet. Click <strong>+ Add Company</strong> to link your first inverter-brand account.</div>`;
      return;
    }
    for (const conn of connections) {
      const brand = BRANDS[conn.brand];
      const row = document.createElement("div");
      row.className = "company-row";
      row.innerHTML = `
        <div class="badge" style="background:${brand.color}">${brand.badge}</div>
        <div class="company-main" data-action="view-stations" data-id="${conn.id}">
          <div class="company-name">${escapeHtml(conn.companyName)}</div>
          <div class="company-sub">${brand.label} · ${(conn.stations || []).length} station(s)${conn.dailyAutoExtract ? " · Daily auto-extract on" : ""}</div>
        </div>
        <div class="company-status">${brand.confidence === "verified" ? "<span class=\"pill ok\">Ready</span>" : "<span class=\"pill warn\">Untested endpoints</span>"}</div>
        <div class="company-actions">
          <button class="btn btn-ghost" data-action="extract" data-id="${conn.id}">Extract</button>
          <button class="btn btn-ghost" data-action="delete" data-id="${conn.id}">Remove</button>
        </div>`;
      root.appendChild(row);
    }
    root.querySelectorAll("[data-action='view-stations']").forEach(el => el.addEventListener("click", () => openStationsModal(el.dataset.id)));
    root.querySelectorAll("[data-action='extract']").forEach(b => b.addEventListener("click", () => { showView("extraction"); qs("extCompany").value = b.dataset.id; handleExtractionCompanyChange(); }));
    root.querySelectorAll("[data-action='delete']").forEach(b => b.addEventListener("click", (e) => { e.stopPropagation(); handleDeleteConnection(b.dataset.id); }));
  }

  function openStationsModal(connId) {
    const conn = connections.find(c => c.id === connId);
    if (!conn) return;
    const brand = BRANDS[conn.brand];
    qs("stationsModalTitle").innerHTML = `<span class="badge" style="background:${brand.color};display:inline-flex;width:26px;height:26px;font-size:.62rem;vertical-align:middle;margin-right:8px;">${brand.badge}</span>${escapeHtml(conn.companyName)}`;
    const list = qs("stationsModalList");
    const stations = conn.stations || [];
    if (!stations.length) {
      list.innerHTML = `<div class="field-help">No stations cached yet — re-run "Test & Connect" on this company to fetch the station list.</div>`;
    } else {
      list.innerHTML = stations.map(s => `
        <div class="station-list-item">
          <div class="station-list-icon">⌂</div>
          <div>
            <div class="station-list-name">${escapeHtml(s.name)}</div>
            <div class="station-list-code">Station code: ${escapeHtml(s.id)}</div>
          </div>
        </div>`).join("");
    }
    qs("stationsModalOverlay").classList.add("active");
  }
  function closeStationsModal() { qs("stationsModalOverlay").classList.remove("active"); }

  async function handleDeleteConnection(id) {
    if (!confirm("Remove this company connection? This does not delete anything on the vendor side.")) return;
    await SheetsClient.deleteConnection(id);
    await refreshConnections();
  }

  /* ---------------------------- add company modal ---------------------------- */

  function renderBrandGrid() {
    const grid = qs("brandGrid");
    grid.innerHTML = "";
    Object.values(BRANDS).forEach(b => {
      const card = document.createElement("button");
      card.type = "button";
      card.className = "brand-card";
      card.dataset.brand = b.key;
      card.innerHTML = `<div class="brand-badge" style="background:${b.color}">${b.badge}</div><div>${b.label}</div>`;
      card.addEventListener("click", () => selectBrand(b.key));
      grid.appendChild(card);
    });
  }

  function openAddCompanyModal() {
    qs("modalOverlay").classList.add("active");
    qs("companyNameInput").value = "";
    pendingBrand = null;
    qs("credFields").innerHTML = "";
    qs("brandNotice").innerHTML = "";
    document.querySelectorAll(".brand-card").forEach(c => c.classList.remove("selected"));
  }
  function closeAddCompanyModal() { qs("modalOverlay").classList.remove("active"); }

  function selectBrand(key) {
    pendingBrand = key;
    document.querySelectorAll(".brand-card").forEach(c => c.classList.toggle("selected", c.dataset.brand === key));
    const brand = BRANDS[key];
    const wrap = qs("credFields");
    wrap.innerHTML = "";
    brand.fields.forEach(f => wrap.appendChild(renderField(f)));
    const notice = qs("brandNotice");
    if (brand.confidence === "bestEffort") {
      notice.innerHTML = `<div class="callout warn"><strong>Untested / best-effort</strong>${escapeHtml(brand.docsNote || "")}</div>`;
    } else {
      notice.innerHTML = `<div class="callout"><strong>Verified against your Postman collection</strong>${escapeHtml(brand.quotaNote || "")}</div>`;
    }
    wireRegionBaseUrlTie(brand);
  }

  // For brands whose "region" field maps to a known base URL (currently
  // FusionSolar), auto-fill the Base URL field from whichever region is
  // selected, so the correct endpoint doesn't have to be typed in by hand.
  // The field stays editable — typing a different value overrides the tie
  // until a region is picked again, matching the "override" label/help text.
  function wireRegionBaseUrlTie(brand) {
    const regionField = brand.fields.find(f => f.key === "region" && f.type === "select");
    const baseUrlField = brand.fields.find(f => f.key === "baseUrl");
    if (!regionField || !baseUrlField) return;
    const regionEl = qs("cred_region");
    const baseUrlEl = qs("cred_baseUrl");
    if (!regionEl || !baseUrlEl) return;

    const applyRegion = () => {
      const opt = regionField.options.find(o => o.value === regionEl.value);
      if (opt?.base) {
        baseUrlEl.value = opt.base;
        baseUrlEl.placeholder = opt.base;
      }
    };
    applyRegion(); // fill it for whatever region is selected by default
    regionEl.addEventListener("change", applyRegion);
  }

  function renderField(f) {
    const wrap = document.createElement("div");
    wrap.className = "field";
    const label = document.createElement("label");
    label.textContent = f.label;
    wrap.appendChild(label);
    let input;
    if (f.type === "select") {
      input = document.createElement("select");
      f.options.forEach(o => {
        const opt = document.createElement("option");
        opt.value = o.value; opt.textContent = o.label;
        input.appendChild(opt);
      });
    } else {
      input = document.createElement("input");
      input.type = f.type === "password" ? "password" : "text";
      if (f.placeholder) input.placeholder = f.placeholder;
    }
    input.id = `cred_${f.key}`;
    input.dataset.key = f.key;
    wrap.appendChild(input);
    if (f.help) {
      const help = document.createElement("div");
      help.className = "field-help";
      help.textContent = f.help;
      wrap.appendChild(help);
    }
    return wrap;
  }

  function readCredsFromModal() {
    const creds = {};
    qs("credFields").querySelectorAll("[data-key]").forEach(el => (creds[el.dataset.key] = el.value));
    return creds;
  }

  async function saveConnectionFromModal(testFirst) {
    if (!pendingBrand) { alert("Pick a brand first."); return; }
    const companyName = qs("companyNameInput").value.trim();
    if (!companyName) { alert("Give this connection a company/site name."); return; }
    const brand = BRANDS[pendingBrand];
    const creds = readCredsFromModal();
    for (const f of brand.fields) {
      if (f.required && !creds[f.key]) { alert(`"${f.label}" is required for ${brand.label}.`); return; }
    }

    let stations = [];
    if (testFirst) {
      qs("btnTestConnect").disabled = true;
      qs("btnTestConnect").textContent = "Testing…";
      try {
        const auth = await brand.buildAuth(creds, ProxyClient.call);
        const ctx = { call: ProxyClient.call, auth };
        stations = await brand.listStations(ctx);
        alert(`Connected — found ${stations.length} station(s).`);
      } catch (e) {
        alert(`Connection failed: ${e.message}`);
        qs("btnTestConnect").disabled = false;
        qs("btnTestConnect").textContent = "Test & Connect";
        return;
      }
      qs("btnTestConnect").disabled = false;
      qs("btnTestConnect").textContent = "Test & Connect";
    }

    const conn = { companyName, brand: pendingBrand, credentials: creds, stations, dailyAutoExtract: false, cursor: {} };
    try {
      await SheetsClient.saveConnection(conn);
    } catch (e) {
      alert(`Could not save connection (${e.message}). Check Settings → Proxy base URL and Google Sheet ID.`);
      return;
    }
    closeAddCompanyModal();
    await refreshConnections();
  }

  /* ---------------------------- extraction ---------------------------- */

  function populateExtractionCompanySelect() {
    const sel = qs("extCompany");
    sel.innerHTML = `<option value="">Select a company…</option>`;
    connections.forEach(c => {
      const opt = document.createElement("option");
      opt.value = c.id; opt.textContent = `${c.companyName} (${BRANDS[c.brand].label})`;
      sel.appendChild(opt);
    });
  }

  async function handleExtractionCompanyChange() {
    const id = qs("extCompany").value;
    const stationBox = qs("extStations");
    stationBox.innerHTML = "";
    if (!id) return;
    const conn = connections.find(c => c.id === id);
    if (!conn) return;
    (conn.stations || []).forEach(s => {
      const row = document.createElement("label");
      row.className = "station-row";
      row.innerHTML = `<input type="checkbox" value="${s.id}" checked> ${escapeHtml(s.name)}`;
      stationBox.appendChild(row);
    });
    if (!conn.stations?.length) {
      stationBox.innerHTML = `<div class="field-help">No stations cached yet — re-run "Test & Connect" on this company to fetch the station list.</div>`;
    }
  }

  async function handleQueueExtraction() {
    const id = qs("extCompany").value;
    const conn = connections.find(c => c.id === id);
    if (!conn) { alert("Pick a company first."); return; }
    const stationIds = [...qs("extStations").querySelectorAll("input:checked")].map(i => i.value);
    const resolution = qs("extResolution").value;
    const startDate = qs("extStart").value;
    const endDate = qs("extEnd").value;
    if (!startDate || !endDate) { alert("Pick a start and end date."); return; }

    const brand = BRANDS[conn.brand];
    let auth = activeAuthByConn[conn.id];
    if (!auth) {
      try {
        auth = await brand.buildAuth(conn.credentials, ProxyClient.call);
        activeAuthByConn[conn.id] = auth;
      } catch (e) { alert(`Could not authenticate: ${e.message}`); return; }
    }
    const ctx = { call: ProxyClient.call, auth };

    const jobRow = document.createElement("div");
    jobRow.className = "job-row";
    jobRow.innerHTML = `<div class="job-title">${escapeHtml(conn.companyName)} · ${resolution} · ${startDate} → ${endDate}</div>
      <div class="job-bar"><div class="job-bar-fill"></div></div>
      <div class="job-status">Starting…</div>`;
    qs("jobList").prepend(jobRow);

    const job = ExtractionEngine.createJob({
      connection: conn, brand: conn.brand, resolution, startDate, endDate, stationIds,
      onProgress: (j) => renderJobProgress(jobRow, j, startDate, endDate),
    });
    await ExtractionEngine.run(job.id, ctx);

    if (job.status === "done" && job.rowsCollected.length) {
      try {
        await SheetsClient.appendReadings(job.rowsCollected.map(r => ({
          timestamp: r.timestamp, company: conn.companyName, brand: brand.label,
          stationId: r.stationId, stationName: (conn.stations.find(s => s.id === r.stationId) || {}).name || r.stationId,
          resolution, kwh: r.kwh,
        })));
        jobRow.querySelector(".job-status").textContent += " · Synced to Google Sheet";
      } catch (e) {
        jobRow.querySelector(".job-status").textContent += ` · Sheet sync failed (${e.message})`;
      }
    }
    conn.cursor = job.connection.cursor;
    await SheetsClient.saveConnection(conn).catch(() => {});
    job._downloadRows = job.rowsCollected;
    jobRow._job = job;
  }

  function renderJobProgress(jobRow, job, startDate, endDate) {
    const total = Math.max(1, (new Date(endDate) - new Date(startDate)) / 86400000 + 1);
    const pct = Math.min(100, Math.round((job.callsMadeToday / total) * 100));
    jobRow.querySelector(".job-bar-fill").style.width = `${pct}%`;
    let statusText = `${job.rowsCollected.length} rows · ${job.status}`;
    if (job.status === "paused" && job.pausedReason === "quota") statusText += " (daily quota reached — resumes tomorrow / next run)";
    if (job.status === "error") statusText = `Error: ${job.error}`;
    jobRow.querySelector(".job-status").textContent = statusText;
    if ((job.status === "done" || job.status === "paused") && job.rowsCollected.length && !jobRow.querySelector(".job-download")) {
      const dl = document.createElement("button");
      dl.className = "btn btn-ghost job-download";
      dl.textContent = "Download Excel";
      dl.addEventListener("click", () => downloadRowsAsExcel(job));
      jobRow.appendChild(dl);
    }
  }

  function downloadRowsAsExcel(job) {
    const rows = job.rowsCollected.map(r => ({
      Timestamp: r.timestamp,
      Station: (job.connection.stations.find(s => s.id === r.stationId) || {}).name || r.stationId,
      Resolution: job.resolution,
      kWh: r.kwh,
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, job.resolution);
    XLSX.writeFile(wb, `${job.connection.companyName}_${job.resolution}_${job.startDate}_${job.endDate}.xlsx`);
  }

  function escapeHtml(s) { return String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c])); }

  return { boot };
})();

window.addEventListener("DOMContentLoaded", App.boot);
