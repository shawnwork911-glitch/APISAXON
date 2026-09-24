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
const DEFAULT_SIGNIN_CLIENT_ID = "336493659847-g7ip6cjvb4un4ouru4fod3605a50q2gu.apps.googleusercontent.com";

const App = (() => {
  let connections = [];      // cached from Google Sheets (or local draft before first sync)
  let activeAuthByConn = {}; // connectionId -> { base, headers, ... } from buildAuth()
  let currentView = "dashboard";
  let pendingBrand = null;   // brand selected inside Add Company modal
  let currentRole = null;    // "Admin" | "User" — set once whoAmI() succeeds

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
    if (!AuthClient.loadClientId() && DEFAULT_SIGNIN_CLIENT_ID) {
      AuthClient.saveClientId(DEFAULT_SIGNIN_CLIENT_ID);
    }
    qs("cfgSpreadsheetId").value = cfg?.spreadsheetId || "";
    qs("cfgProxyUrl").value = ProxyClient.getProxyBaseUrl();
    qs("cfgSignInClientId").value = AuthClient.loadClientId();

    startAuthGate();
  }

  /* ---------------------------- sign-in gate ---------------------------- */

  function startAuthGate(retriesLeft = 20) {
    const clientId = AuthClient.loadClientId();
    if (!clientId) {
      qs("authGateNotConfigured").hidden = false;
      return;
    }
    qs("authGateNotConfigured").hidden = true;

    if (typeof google === "undefined") {
      // The Google Sign-In script (accounts.google.com/gsi/client) hasn't
      // finished loading yet — wait briefly and retry rather than silently
      // showing an empty box. If it never loads (ad blocker, network
      // block), say so plainly instead of leaving the gate blank forever.
      if (retriesLeft > 0) { setTimeout(() => startAuthGate(retriesLeft - 1), 150); return; }
      const box = qs("authGateError");
      box.hidden = false;
      box.innerHTML = `<strong>Google Sign-In didn't load</strong>Check your internet connection or whether an ad/script blocker is blocking accounts.google.com, then reload the page.`;
      return;
    }
    const ok = AuthClient.init(onGoogleSignedIn);
    if (ok) AuthClient.renderButton(qs("authGateButton"));
  }

  async function onGoogleSignedIn() {
    qs("authGateError").hidden = true;
    qs("authGateButton").innerHTML = `<span class="field-help">Checking access…</span>`;
    try {
      const who = await SheetsClient.whoAmI();
      currentRole = who.role;
      qs("authGateOverlay").classList.remove("active");
      qs("authUserLabel").textContent = `${who.email} · ${who.role}`;
      await afterSignedIn();
    } catch (e) {
      AuthClient.signOut();
      qs("authGateButton").innerHTML = "";
      AuthClient.init(onGoogleSignedIn);
      AuthClient.renderButton(qs("authGateButton"));
      const box = qs("authGateError");
      box.hidden = false;
      box.innerHTML = `<strong>Sign-in didn't go through</strong>${escapeHtml(e.message)}`;
    }
  }

  async function afterSignedIn() {
    updateConnectionBanner();
    if (SheetsClient.isConfigured()) await refreshConnections();
    renderBrandGrid();
    showView("dashboard");
  }

  function handleSignOut() {
    AuthClient.signOut();
    location.reload(); // simplest reliable way back to a clean, gated state
  }

  function cacheEls() {
    ["dashboard", "extraction", "compare", "export", "settings"].forEach(v => els[v] = qs(`view-${v}`));
  }

  function wireStaticEvents() {
    DatePicker.attach(qs("extStart"));
    DatePicker.attach(qs("extEnd"));
    DatePicker.attach(qs("expStart"));
    DatePicker.attach(qs("expEnd"));
    qs("btnSignOut").addEventListener("click", handleSignOut);
    qs("authGateClientIdSave").addEventListener("click", () => {
      const id = qs("authGateClientIdInput").value.trim();
      if (!id) return;
      AuthClient.saveClientId(id);
      qs("cfgSignInClientId").value = id;
      startAuthGate();
    });
    document.querySelectorAll(".nav-btn[data-view]").forEach(btn => {
      btn.addEventListener("click", () => showView(btn.dataset.view));
    });
    qs("btnAddCompany").addEventListener("click", openAddCompanyModal);
    qs("btnModalClose").addEventListener("click", closeAddCompanyModal);
    qs("btnStationsModalClose").addEventListener("click", closeStationsModal);
    qs("btnStationsModalDone").addEventListener("click", closeStationsModal);
    qs("stationsModalOverlay").addEventListener("click", (e) => { if (e.target.id === "stationsModalOverlay") closeStationsModal(); });
    qs("stationsModalTable").addEventListener("input", (e) => {
      if (!e.target.classList.contains("modalStationFacilityInput")) return;
      stationsModalState.stationFacility[e.target.dataset.station] = e.target.value;
      renderStationsModalFacilityTable();
      const conn = connections.find(c => c.id === stationsModalConnId);
      if (conn) renderMissingFacilityWarning(conn);
    });
    qs("stationsModalTable").addEventListener("change", saveStationsModalSettings);
    qs("stationsModalFacilityTable").addEventListener("change", (e) => {
      const fid = e.target.dataset.facility;
      if (!fid) return;
      stationsModalState.facilityGroups[fid] = stationsModalState.facilityGroups[fid] || {};
      if (e.target.classList.contains("modalFacilityMeterInput")) stationsModalState.facilityGroups[fid].meterId = e.target.value;
      if (e.target.classList.contains("modalFacilityRegistrySelect")) stationsModalState.facilityGroups[fid].eacRegistryId = e.target.value;
      saveStationsModalSettings();
    });
    qs("btnModalCancel").addEventListener("click", closeAddCompanyModal);
    qs("btnSaveWithoutTest").addEventListener("click", () => saveConnectionFromModal(false));
    qs("btnTestConnect").addEventListener("click", () => saveConnectionFromModal(true));
    qs("btnSaveSettings").addEventListener("click", handleSaveSettings);
    qs("btnQueueExtraction").addEventListener("click", handleQueueExtraction);
    qs("extCompany").addEventListener("change", handleExtractionCompanyChange);
    qs("extResolution").addEventListener("change", suggestStartDateFromCursor);
    qs("btnRunCompare").addEventListener("click", handleRunCompare);
    qs("btnDownloadCompare").addEventListener("click", handleDownloadCompare);
    qs("btnRunExport").addEventListener("click", handleRunExport);
    qs("expResolution").addEventListener("change", syncExportFieldVisibility);
  }

  function showView(view) {
    currentView = view;
    Object.entries(els).forEach(([k, el]) => el.classList.toggle("active", k === view));
    document.querySelectorAll(".nav-btn[data-view]").forEach(b => b.classList.toggle("active", b.dataset.view === view));
    if (view === "extraction") { populateExtractionCompanySelect(); renderPersistedPendingJobs(); }
    if (view === "compare") populateCompareCompanySelect();
    if (view === "export") { populateExportCompanySelect(); syncExportFieldVisibility(); }
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
    const newClientId = qs("cfgSignInClientId").value.trim();
    const clientIdChanged = newClientId !== AuthClient.loadClientId();
    AuthClient.saveClientId(newClientId);
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
    if (clientIdChanged) {
      qs("settingsSavedNote").textContent = "Sign-in Client ID changed — reloading…";
      setTimeout(() => location.reload(), 1200);
      return;
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

  function missingFacilityStations(conn) {
    const map = (conn.templateSettings?.stationFacility) || {};
    return (conn.stations || []).filter(s => !(map[s.id] || "").trim());
  }

  function renderDashboard() {
    const root = qs("companyList");
    root.innerHTML = "";
    if (!connections.length) {
      root.innerHTML = `<div class="empty-state">No companies connected yet. Click <strong>+ Add Company</strong> to link your first inverter-brand account.</div>`;
      return;
    }
    for (const conn of connections) {
      const brand = BRANDS[conn.brand];
      const missing = missingFacilityStations(conn);
      const missingBadge = missing.length
        ? `<span class="pill warn" title="${escapeHtml(missing.map(s => s.name).join(", "))}">⚠ ${missing.length} missing facility_id</span>`
        : "";
      const row = document.createElement("div");
      row.className = "company-row";
      row.innerHTML = `
        <div class="badge" style="background:${brand.color}">${brand.badge}</div>
        <div class="company-main" data-action="view-stations" data-id="${conn.id}">
          <div class="company-name">${escapeHtml(conn.companyName)}</div>
          <div class="company-sub">${brand.label} · ${(conn.stations || []).length} station(s)</div>
        </div>
        <div class="company-status">${brand.confidence === "verified" ? "<span class=\"pill ok\">Ready</span>" : "<span class=\"pill warn\">Untested endpoints</span>"}${missingBadge}</div>
        <div class="company-actions">
          <button class="btn btn-ghost" data-action="extract" data-id="${conn.id}">Extract</button>
          ${currentRole === "Admin" ? `<button class="btn btn-ghost" data-action="delete" data-id="${conn.id}">Remove</button>` : ""}
        </div>`;
      root.appendChild(row);
    }
    root.querySelectorAll("[data-action='view-stations']").forEach(el => el.addEventListener("click", () => openStationsModal(el.dataset.id)));
    root.querySelectorAll("[data-action='extract']").forEach(b => b.addEventListener("click", () => { showView("extraction"); qs("extCompany").value = b.dataset.id; handleExtractionCompanyChange(); }));
    root.querySelectorAll("[data-action='delete']").forEach(b => b.addEventListener("click", (e) => { e.stopPropagation(); handleDeleteConnection(b.dataset.id); }));
  }

  // stationsModalState holds a working copy of the CURRENTLY OPEN company's
  // templateSettings while the modal is open (this is the only place these
  // settings are edited now — the Extraction tab's own copy was removed
  // since the Export tab supersedes it).
  let stationsModalState = null;
  let stationsModalConnId = null;

  function openStationsModal(connId) {
    const conn = connections.find(c => c.id === connId);
    if (!conn) return;
    stationsModalConnId = connId;
    const brand = BRANDS[conn.brand];
    qs("stationsModalTitle").innerHTML = `<span class="badge" style="background:${brand.color};display:inline-flex;width:26px;height:26px;font-size:.62rem;vertical-align:middle;margin-right:8px;">${brand.badge}</span>${escapeHtml(conn.companyName)}`;

    stationsModalState = conn.templateSettings ? JSON.parse(JSON.stringify(conn.templateSettings)) : TemplateExport.defaultSettings("1");
    stationsModalState.stationFacility = stationsModalState.stationFacility || {};
    stationsModalState.facilityGroups = stationsModalState.facilityGroups || {};

    renderStationsModalTable(conn);
    renderStationsModalFacilityTable();
    renderMissingFacilityWarning(conn);
    qs("stationsModalSavedNote").textContent = "";
    qs("stationsModalOverlay").classList.add("active");
    loadDataStoredSummary(conn);
  }

  async function loadDataStoredSummary(conn) {
    const table = qs("stationsModalDataStoredTable");
    table.innerHTML = `<tbody><tr><td class="field-help">Loading…</td></tr></tbody>`;
    try {
      const readings = await SheetsClient.listReadings(conn.companyName);
      renderDataStoredTable(table, readings);
    } catch (e) {
      table.innerHTML = `<tbody><tr><td class="field-help">Could not load: ${escapeHtml(e.message)}</td></tr></tbody>`;
    }
  }

  function renderDataStoredTable(table, readings) {
    if (!readings.length) {
      table.innerHTML = `<tbody><tr><td class="field-help">Nothing extracted yet for this company.</td></tr></tbody>`;
      return;
    }
    // Grouped by resolution — dates here are approximate (not timezone-corrected
    // per brand, unlike Template exports) since this is just a coverage overview.
    const byResolution = {};
    for (const r of readings) {
      const bucket = byResolution[r.resolution] || (byResolution[r.resolution] = { count: 0, min: null, max: null });
      bucket.count++;
      const n = Number(r.timestamp);
      const d = !isNaN(n) && n > 1e10 ? new Date(n) : new Date(r.timestamp);
      if (isNaN(d.getTime())) continue;
      if (!bucket.min || d < bucket.min) bucket.min = d;
      if (!bucket.max || d > bucket.max) bucket.max = d;
    }
    const fmt = d => d ? d.toISOString().slice(0, 10) : "?";
    const order = ["Hourly", "Daily", "Monthly"];
    const resolutions = Object.keys(byResolution).sort((a, b) => order.indexOf(a) - order.indexOf(b));
    table.innerHTML = `<thead><tr><th>Resolution</th><th>Rows</th><th>Earliest</th><th>Latest</th></tr></thead><tbody>` +
      resolutions.map(res => {
        const b = byResolution[res];
        return `<tr><td>${escapeHtml(res)}</td><td>${b.count}</td><td>${fmt(b.min)}</td><td>${fmt(b.max)}</td></tr>`;
      }).join("") + `</tbody>`;
  }

  function renderMissingFacilityWarning(conn) {
    const box = qs("stationsModalMissingWarning");
    const missing = (conn.stations || []).filter(s => !(stationsModalState.stationFacility[s.id] || "").trim());
    if (!missing.length) { box.hidden = true; box.textContent = ""; return; }
    box.hidden = false;
    box.innerHTML = `<strong>${missing.length} station(s) have no facility_id set</strong>`
      + `Each will export as its own separate facility until assigned: ${missing.map(s => escapeHtml(s.name)).join(", ")}.`;
  }

  function renderStationsModalTable(conn) {
    const wrap = qs("stationsModalTable");
    const stations = conn.stations || [];
    if (!stations.length) {
      wrap.innerHTML = `<tbody><tr><td class="field-help">No stations cached yet — re-run "Test & Connect" on this company to fetch the station list.</td></tr></tbody>`;
      return;
    }
    wrap.innerHTML = `<thead><tr><th>Station</th><th>facility_id</th></tr></thead><tbody>` +
      stations.map(s => `
        <tr>
          <td>${escapeHtml(s.name)}<div class="station-list-code">${escapeHtml(s.id)}</div></td>
          <td><input type="text" class="modalStationFacilityInput" data-station="${escapeHtml(s.id)}"
                     value="${escapeHtml(stationsModalState.stationFacility[s.id] || "")}"
                     placeholder="${escapeHtml(s.id)}"></td>
        </tr>`).join("") + `</tbody>`;
  }

  function renderStationsModalFacilityTable() {
    const wrap = qs("stationsModalFacilityTable");
    const distinct = [...new Set(Object.values(stationsModalState.stationFacility).filter(v => v && v.trim()))];
    if (!distinct.length) {
      wrap.innerHTML = `<tbody><tr><td class="field-help">No facility_id assigned yet — enter one above to see its group settings here.</td></tr></tbody>`;
      return;
    }
    // First time a facility_id is seen, auto-copy it into meter_id as an actual
    // value (not just a placeholder) — still editable/overridable afterwards.
    // Only fills it when meter_id has never been touched (undefined), so a
    // deliberately-cleared value doesn't get silently re-filled.
    distinct.forEach(fid => {
      stationsModalState.facilityGroups[fid] = stationsModalState.facilityGroups[fid] || {};
      if (stationsModalState.facilityGroups[fid].meterId === undefined) stationsModalState.facilityGroups[fid].meterId = fid;
    });
    wrap.innerHTML = `<thead><tr><th>facility_id</th><th>meter_id</th><th>eac_registry_id</th></tr></thead><tbody>` +
      distinct.map(fid => {
        const g = stationsModalState.facilityGroups[fid] || {};
        return `<tr>
          <td>${escapeHtml(fid)}</td>
          <td><input type="text" class="modalFacilityMeterInput" data-facility="${escapeHtml(fid)}" value="${escapeHtml(g.meterId ?? fid)}"></td>
          <td><select class="modalFacilityRegistrySelect" data-facility="${escapeHtml(fid)}">
                <option value="tigr" ${g.eacRegistryId !== "irec" ? "selected" : ""}>TIGR</option>
                <option value="irec" ${g.eacRegistryId === "irec" ? "selected" : ""}>I-REC</option>
              </select></td>
        </tr>`;
      }).join("") + `</tbody>`;
  }

  async function saveStationsModalSettings() {
    const conn = connections.find(c => c.id === stationsModalConnId);
    if (!conn || !stationsModalState) return;
    conn.templateSettings = stationsModalState;
    qs("stationsModalSavedNote").textContent = "Saving…";
    try {
      await SheetsClient.saveConnection(conn);
      qs("stationsModalSavedNote").textContent = "Saved.";
      // Keep the Extraction tab's own working copy in sync if it's currently
      // showing this same company, so switching tabs doesn't show stale data.
      if (qs("extCompany").value === conn.id) loadTemplateSettingsIntoForm(conn);
    } catch (e) {
      qs("stationsModalSavedNote").textContent = `Could not save: ${e.message}`;
      qs("stationsModalSavedNote").style.color = "var(--err)";
    }
  }

  function closeStationsModal() {
    qs("stationsModalOverlay").classList.remove("active");
    if (currentView === "dashboard") renderDashboard(); // refresh the missing-facility_id badge without a full reload
  }

  async function handleDeleteConnection(id) {
    if (!confirm("Remove this company connection? This does not delete anything on the vendor side.")) return;
    try {
      await SheetsClient.deleteConnection(id);
      await refreshConnections();
    } catch (e) {
      alert(`Could not remove: ${e.message}`);
    }
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

  /* ---------------------------- compare (hourly vs monthly) ---------------------------- */

  let lastCompareResult = null; // kept around for the Download Excel button

  function populateCompareCompanySelect() {
    const sel = qs("cmpCompany");
    sel.innerHTML = `<option value="">Select a company…</option>`;
    connections.forEach(c => {
      const opt = document.createElement("option");
      opt.value = c.id; opt.textContent = `${c.companyName} (${BRANDS[c.brand].label})`;
      sel.appendChild(opt);
    });
  }

  async function handleRunCompare() {
    const id = qs("cmpCompany").value;
    const conn = connections.find(c => c.id === id);
    if (!conn) { alert("Pick a company first."); return; }

    qs("btnRunCompare").disabled = true;
    qs("cmpStatus").textContent = "Reading Hourly and Monthly rows from your Google Sheet…";
    qs("cmpResultsCard").hidden = true;

    try {
      const readings = await SheetsClient.listReadings(conn.companyName);
      const result = CompareEngine.compareHourlyVsMonthly(readings, conn);
      lastCompareResult = { conn, result };
      renderCompareResults(conn, result);
      qs("cmpStatus").textContent = `Done — ${readings.length} row(s) read.`;
    } catch (e) {
      qs("cmpStatus").textContent = `Failed: ${e.message}`;
    }
    qs("btnRunCompare").disabled = false;
  }

  function renderCompareResults(conn, result) {
    qs("cmpResultsCard").hidden = false;
    const noHourly = !result.hourlyMonths.length;
    const noMonthly = !result.monthlyMonths.length;
    let summary = `Checked ${result.totalChecked} facility-month combination(s). `;
    if (noHourly || noMonthly) {
      summary += `<strong style="color:var(--warn);">Missing ${noHourly ? "Hourly" : "Monthly"} data entirely for this company — run that extraction first.</strong>`;
    } else if (!result.flagged.length) {
      summary += `<span style="color:var(--ok);">No discrepancies found — hourly and monthly figures match for every checked month.</span>`;
    } else {
      summary += `<strong style="color:var(--warn);">${result.flagged.length} month(s) flagged.</strong>`;
    }
    qs("cmpSummary").innerHTML = summary;

    const table = qs("cmpResultsTable");
    if (!result.flagged.length) {
      table.innerHTML = "";
      return;
    }
    table.innerHTML = `<thead><tr><th>facility_id</th><th>Month</th><th>Monthly report (kWh)</th><th>Hourly sum (kWh)</th><th>Diff (kWh)</th><th>Diff (%)</th><th>Issue</th></tr></thead><tbody>` +
      result.flagged.map(r => `
        <tr>
          <td>${escapeHtml(r.facilityId)}</td>
          <td>${escapeHtml(r.month)}</td>
          <td>${r.monthlyKwh}</td>
          <td>${r.hourlyKwh}</td>
          <td>${r.diffKwh}</td>
          <td>${r.diffPct}%</td>
          <td>${escapeHtml(r.issue)}</td>
        </tr>`).join("") + `</tbody>`;
  }

  function handleDownloadCompare() {
    if (!lastCompareResult) return;
    const { conn, result } = lastCompareResult;
    const headers = ["facility_id", "Month", "Monthly report (kWh)", "Hourly sum (kWh)", "Diff (kWh)", "Diff (%)", "Issue"];
    const rows = result.flagged.length
      ? result.flagged.map(r => [r.facilityId, r.month, r.monthlyKwh, r.hourlyKwh, r.diffKwh, r.diffPct, r.issue])
      : [["No discrepancies found", "", "", "", "", "", ""]];
    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "HourlyVsMonthly");
    XLSX.writeFile(wb, `${conn.companyName}_HourlyVsMonthly.xlsx`);
  }

  /* ---------------------------- export (Template 1/2 from stored Readings) ---------------------------- */

  function populateExportCompanySelect() {
    const sel = qs("expCompany");
    sel.innerHTML = `<option value="">Select a company…</option>`;
    connections.forEach(c => {
      const opt = document.createElement("option");
      opt.value = c.id; opt.textContent = `${c.companyName} (${BRANDS[c.brand].label})`;
      sel.appendChild(opt);
    });
  }

  function syncExportFieldVisibility() {
    const isHourly = qs("expResolution").value === "Hourly";
    qs("expTemplateWrap").hidden = !isHourly;
    qs("expMonthlyNote").hidden = isHourly;
  }

  async function handleRunExport() {
    const id = qs("expCompany").value;
    const conn = connections.find(c => c.id === id);
    if (!conn) { alert("Pick a company first."); return; }
    const resolution = qs("expResolution").value;
    const startDate = DatePicker.getISO(qs("expStart"));
    const endDate = DatePicker.getISO(qs("expEnd"));
    if (!startDate || !endDate) { alert("Pick a start and end date."); return; }

    qs("btnRunExport").disabled = true;
    qs("expStatus").textContent = `Reading ${resolution} rows from your Google Sheet…`;
    qs("expResultsCard").hidden = true;

    try {
      const readings = await SheetsClient.listReadings(conn.companyName);
      const brandKey = conn.brand;

      if (resolution === "Hourly") {
        // Facility mapping (station→facility_id, meter_id, eac_registry_id) still
        // comes from what's configured on the company — only the template
        // number (1 vs 2) is chosen fresh here, independent of that saved default.
        const baseSettings = conn.templateSettings ? JSON.parse(JSON.stringify(conn.templateSettings)) : TemplateExport.defaultSettings("1");
        baseSettings.template = qs("expTemplate").value;
        const utcOffset = baseSettings.utcOffset ?? 8;
        const filtered = readings.filter(r => {
          if (r.resolution !== "Hourly") return false;
          const wc = TemplateExport.wallClockFromRow(r, brandKey, utcOffset);
          if (!wc) return false;
          const localDate = `${wc.y}-${String(wc.mo).padStart(2, "0")}-${String(wc.d).padStart(2, "0")}`;
          return localDate >= startDate && localDate <= endDate;
        });
        const groups = TemplateExport.build(filtered, brandKey, baseSettings);
        renderExportResults(conn, { resolution, template: baseSettings.template }, groups, filtered.length, "Hourly");
        qs("expStatus").textContent = `Done — ${readings.length} row(s) read, ${filtered.length} in range.`;
      } else {
        const settingsForFacility = conn.templateSettings || {};
        const startMonth = `${startDate.slice(0, 7)}-01`;
        const filtered = readings.filter(r => {
          if (r.resolution !== "Monthly") return false;
          const wc = TemplateExport.wallClockFromRow(r, brandKey, settingsForFacility.utcOffset ?? 8);
          if (!wc) return false;
          const rowMonth = `${wc.y}-${String(wc.mo).padStart(2, "0")}-01`;
          return rowMonth >= startMonth && rowMonth <= endDate;
        });
        const groups = buildMonthlyExportGroups(filtered, brandKey, settingsForFacility);
        renderExportResults(conn, { resolution, template: null }, groups, filtered.length, "Monthly");
        qs("expStatus").textContent = `Done — ${readings.length} row(s) read, ${filtered.length} in range.`;
      }
    } catch (e) {
      qs("expStatus").textContent = `Failed: ${e.message}`;
    }
    qs("btnRunExport").disabled = false;
  }

  // The Monthly template is fixed and much simpler than Template 1/2: one
  // row per STATION per month (not summed across stations sharing a
  // facility_id — stationName stays its own column), tagged with its
  // facility_id under the header "GEN ID" to match the reference format.
  function buildMonthlyExportGroups(readings, brandKey, settings) {
    const headers = ["GEN ID", "stationName", "collectTime", "PVYield"];
    const byFacility = new Map(); // facilityId -> rows[]
    for (const r of readings) {
      const wc = TemplateExport.wallClockFromRow(r, brandKey, settings.utcOffset ?? 8);
      if (!wc) continue;
      const facilityId = TemplateExport.facilityKeyFor(r.stationId, settings);
      const collectTime = `${wc.y}-${String(wc.mo).padStart(2, "0")}-01 00:00:00`;
      const kwhR = Math.round((r.kwh || 0) * 100) / 100;
      if (!byFacility.has(facilityId)) byFacility.set(facilityId, []);
      byFacility.get(facilityId).push({ epoch: Date.UTC(wc.y, wc.mo - 1, 1), row: [facilityId, r.stationName, collectTime, kwhR] });
    }
    const groups = [];
    for (const [facilityId, entries] of byFacility.entries()) {
      entries.sort((a, b) => a.epoch - b.epoch);
      groups.push({ facilityId, headers, rows: entries.map(e => e.row) });
    }
    groups.sort((a, b) => a.facilityId.localeCompare(b.facilityId));
    return groups;
  }

  function renderExportResults(conn, exportKind, groups, rowCount, resolutionLabel) {
    qs("expResultsCard").hidden = false;
    qs("expSummary").innerHTML = groups.length
      ? `${rowCount} ${resolutionLabel.toLowerCase()} row(s) in range, grouped into <strong>${groups.length}</strong> facility file(s).`
      : `No ${resolutionLabel.toLowerCase()} rows found in that date range for this company.`;

    const list = qs("expFacilityList");
    list.innerHTML = "";
    groups.forEach(g => {
      const row = document.createElement("div");
      row.className = "export-facility-row";
      const formatLabel = exportKind.resolution === "Hourly" ? `Template ${exportKind.template}` : "Monthly template";
      row.innerHTML = `
        <div>
          <div style="font-weight:600;">${escapeHtml(g.facilityId)}</div>
          <div class="meta">${g.rows.length} row(s) · ${escapeHtml(formatLabel)}</div>
        </div>
        <button class="btn btn-primary" data-facility="${escapeHtml(g.facilityId)}">Download</button>`;
      row.querySelector("button").addEventListener("click", () => downloadExportFacility(conn, exportKind, g));
      list.appendChild(row);
    });
  }

  function downloadExportFacility(conn, exportKind, group) {
    const safeFacility = group.facilityId.replace(/[^a-z0-9_-]+/gi, "_");
    const ws = XLSX.utils.aoa_to_sheet([group.headers, ...group.rows]);
    const wb = XLSX.utils.book_new();
    if (exportKind.resolution === "Hourly") {
      const sheetName = exportKind.template === "1" ? "Data" : "MeterData";
      XLSX.utils.book_append_sheet(wb, ws, sheetName);
      XLSX.writeFile(wb, `${conn.companyName}_${safeFacility}_export_template_${exportKind.template}.xlsx`);
    } else {
      XLSX.utils.book_append_sheet(wb, ws, "Data");
      XLSX.writeFile(wb, `${conn.companyName}_${safeFacility}_export_monthly.xlsx`);
    }
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
    suggestStartDateFromCursor();
  }

  // Pre-fills Start Date with wherever this company+resolution last left off,
  // as a convenience default — purely a suggestion, never a hidden override.
  // Typing a different date (e.g. an earlier one, to backfill) is fully
  // respected; nothing silently resumes from the cursor instead anymore.
  function suggestStartDateFromCursor() {
    const id = qs("extCompany").value;
    const conn = connections.find(c => c.id === id);
    if (!conn) return;
    const resolution = qs("extResolution").value;
    const cursorDate = conn.cursor?.[resolution];
    if (cursorDate && !DatePicker.getISO(qs("extStart"))) {
      DatePicker.setFromISO(qs("extStart"), cursorDate);
    }
  }

  // A global queue so that clicking "Queue extraction" for a second company
  // while the first is still running doesn't let both jobs' API calls
  // interleave — only one job's actual calls run at a time. This matters
  // for two reasons: it avoids bursting the Google Sheets API's per-minute
  // read quota, and (more importantly) it stops two simultaneous jobs for
  // the SAME brand from collectively exceeding that brand's own daily call
  // quota — each job only tracks its own count, so without this, two
  // FusionSolar jobs running "at once" could together blow past 25/day
  // without either one knowing about the other.
  const extractionQueue = [];
  let queueRunning = false;

  function enqueueExtraction(task) {
    extractionQueue.push(task);
    runQueue();
  }
  async function runQueue() {
    if (queueRunning) return;
    queueRunning = true;
    while (extractionQueue.length) {
      const task = extractionQueue.shift();
      try { await task(); } catch (e) { console.error("Queued extraction failed:", e); }
    }
    queueRunning = false;
  }

  async function handleQueueExtraction() {
    const id = qs("extCompany").value;
    const conn = connections.find(c => c.id === id);
    if (!conn) { alert("Pick a company first."); return; }
    const stationIds = [...qs("extStations").querySelectorAll("input:checked")].map(i => i.value);
    const resolution = qs("extResolution").value;
    const startDate = DatePicker.getISO(qs("extStart"));
    const endDate = DatePicker.getISO(qs("extEnd"));
    if (!startDate || !endDate) { alert("Pick a start and end date."); return; }
    const today = new Date().toISOString().slice(0, 10);
    if (endDate > today) {
      alert("End date can't be in the future — data for days that haven't happened yet doesn't exist. Pick today or an earlier date.");
      return;
    }
    queueExtractionJob(conn, resolution, stationIds, startDate, endDate);
  }

  // Shared by both "Queue extraction" (fresh, from the form) and "Resume"
  // (reconstructed from a persisted pendingJob, no form re-entry needed).
  function queueExtractionJob(conn, resolution, stationIds, startDate, endDate) {
    const brand = BRANDS[conn.brand];

    // Created immediately so it shows up in the Jobs list right away, even
    // if it has to wait its turn behind another job that's already running.
    const jobRow = document.createElement("div");
    jobRow.className = "job-row";
    jobRow.innerHTML = `<div class="job-title">${escapeHtml(conn.companyName)} · ${resolution} · ${startDate} → ${endDate}</div>
      <div class="job-bar"><div class="job-bar-fill"></div></div>
      <div class="job-status">${queueRunning || extractionQueue.length ? "Waiting for other extraction(s) to finish…" : "Starting…"}</div>`;
    qs("jobList").prepend(jobRow);

    const job = ExtractionEngine.createJob({
      connection: conn, brand: conn.brand, resolution, startDate, endDate, stationIds,
      onProgress: (j) => renderJobProgress(jobRow, j, startDate, endDate),
    });
    liveJobKeys.add(`${conn.id}|${resolution}`); // so a persisted "Resume" card for the same job doesn't also render

    // Remember this as unfinished right away — saved to the Sheet, not just
    // in memory — so if the tab gets closed mid-run (or it just pauses on
    // quota), the NEXT time this company's shown, a "Resume" card appears
    // instead of the info being lost.
    conn.pendingJob = conn.pendingJob || {};
    conn.pendingJob[resolution] = { stationIds, startDate, endDate };
    SheetsClient.saveConnection(conn).catch(() => {});

    const cancelBtn = document.createElement("button");
    cancelBtn.className = "btn btn-ghost job-cancel";
    cancelBtn.textContent = "Cancel";
    cancelBtn.addEventListener("click", () => {
      // Cancels immediately whether it's already running or still waiting
      // in the queue — for a queued-but-not-started job this just marks it
      // stopped so runQueue() skips its real work when its turn comes.
      ExtractionEngine.stop(job.id);
      cancelBtn.disabled = true;
      cancelBtn.textContent = "Cancelling…";
    });
    jobRow.appendChild(cancelBtn);

    enqueueExtraction(async () => {
      if (job.status === "stopped") {
        jobRow.querySelector(".job-status").textContent = "Cancelled (was still waiting in the queue)";
        cancelBtn.remove();
        return;
      }

      let auth = activeAuthByConn[conn.id];
      if (!auth) {
        try {
          auth = await brand.buildAuth(conn.credentials, ProxyClient.call);
          activeAuthByConn[conn.id] = auth;
        } catch (e) { jobRow.querySelector(".job-status").textContent = `Error: could not authenticate — ${e.message}`; return; }
      }
      const ctx = { call: ProxyClient.call, auth };

      await ExtractionEngine.run(job.id, ctx);

      if ((job.status === "done" || job.status === "paused" || job.status === "stopped") && job.rowsCollected.length) {
        try {
          const utcOffset = conn.templateSettings?.utcOffset ?? 8;
          await SheetsClient.appendReadings(job.rowsCollected.map(r => ({
            timestamp: formatReadableTimestamp(r, conn.brand, utcOffset), company: conn.companyName, brand: brand.label,
            stationId: r.stationId, stationName: (conn.stations.find(s => s.id === r.stationId) || {}).name || r.stationId,
            resolution, kwh: r.kwh,
          })));
          jobRow.querySelector(".job-status").textContent += " · Synced to Google Sheet";
        } catch (e) {
          jobRow.querySelector(".job-status").textContent += ` · Sheet sync failed (${e.message})`;
        }
      }
      conn.cursor = job.connection.cursor;
      if (job.status === "done") {
        delete conn.pendingJob[resolution]; // fully caught up — nothing left to resume
      }
      await SheetsClient.saveConnection(conn).catch(() => {});
      job._downloadRows = job.rowsCollected;
      jobRow._job = job;
    });
  }

  // Companies with a saved pendingJob (paused/cancelled/tab-closed in a
  // previous session) get a "Resume" card here — same Jobs list, no form
  // re-entry needed. Skips anything already shown live this session.
  const liveJobKeys = new Set();

  function renderPersistedPendingJobs() {
    for (const conn of connections) {
      if (!conn.pendingJob) continue;
      for (const [resolution, pending] of Object.entries(conn.pendingJob)) {
        const key = `${conn.id}|${resolution}`;
        if (liveJobKeys.has(key)) continue;
        liveJobKeys.add(key);

        const effectiveStart = conn.cursor?.[resolution] || pending.startDate;
        const jobRow = document.createElement("div");
        jobRow.className = "job-row";
        jobRow.innerHTML = `<div class="job-title">${escapeHtml(conn.companyName)} · ${resolution} · ${pending.startDate} → ${pending.endDate}</div>
          <div class="job-bar"><div class="job-bar-fill"></div></div>
          <div class="job-status">Unfinished from a previous session — currently caught up to ${escapeHtml(effectiveStart)}.</div>`;
        const resumeBtn = document.createElement("button");
        resumeBtn.className = "btn btn-primary";
        resumeBtn.textContent = "Resume";
        resumeBtn.addEventListener("click", () => {
          jobRow.remove();
          queueExtractionJob(conn, resolution, pending.stationIds, effectiveStart, pending.endDate);
        });
        jobRow.appendChild(resumeBtn);
        qs("jobList").appendChild(jobRow);
      }
    }
  }

  function renderJobProgress(jobRow, job, startDate, endDate) {
    const total = Math.max(1, (new Date(endDate) - new Date(startDate)) / 86400000 + 1);
    const pct = Math.min(100, Math.round((job.callsMadeToday / total) * 100));
    jobRow.querySelector(".job-bar-fill").style.width = `${pct}%`;
    let statusText = `${job.rowsCollected.length} rows · ${job.status}`;
    if (job.status === "paused" && job.pausedReason === "quota") statusText += " (daily quota reached — resumes tomorrow / next run)";
    if (job.status === "stopped") statusText = `${job.rowsCollected.length} rows · Cancelled`;
    if (job.status === "error") statusText = `Error: ${job.error}`;
    jobRow.querySelector(".job-status").textContent = statusText;

    if (job.status !== "running" && job.status !== "queued") {
      const cancelBtn = jobRow.querySelector(".job-cancel");
      if (cancelBtn) cancelBtn.remove();
    }
    if ((job.status === "done" || job.status === "paused" || job.status === "stopped") && job.rowsCollected.length && !jobRow.querySelector(".job-download")) {
      const dl = document.createElement("button");
      dl.className = "btn btn-ghost job-download";
      dl.textContent = "Download Excel";
      dl.addEventListener("click", () => downloadRowsAsExcel(job));
      jobRow.appendChild(dl);
    }
  }

  // Always raw (Timestamp/Station/Resolution/kWh) — for Template 1/2, use
  // the Export tab, which reads accumulated Readings for any date range
  // rather than just this one job's in-memory rows.
  function downloadRowsAsExcel(job) {
    const utcOffset = job.connection.templateSettings?.utcOffset ?? 8;
    const rows = job.rowsCollected.map(r => ({
      Timestamp: formatReadableTimestamp(r, job.brand, utcOffset),
      Station: (job.connection.stations.find(s => s.id === r.stationId) || {}).name || r.stationId,
      Resolution: job.resolution,
      kWh: r.kwh,
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, job.resolution);
    XLSX.writeFile(wb, `${job.connection.companyName}_${job.resolution}_${job.startDate}_${job.endDate}_raw.xlsx`);
  }

  function escapeHtml(s) { return String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c])); }

  // FusionSolar's raw timestamp is an epoch number (e.g. 1782835200000) —
  // fine for calculation, unreadable if written straight into the Sheet.
  // SolarEdge's is already a readable local-time string from their API.
  // This makes both consistent: a plain "YYYY-MM-DD HH:MM:SS" local string,
  // reusing the exact same brand-aware wall-clock logic already verified
  // correct in Template exports.
  function formatReadableTimestamp(row, brandKey, utcOffset) {
    const wc = TemplateExport.wallClockFromRow(row, brandKey, utcOffset ?? 8);
    if (!wc) return String(row.timestamp);
    const pad = n => String(n).padStart(2, "0");
    return `${wc.y}-${pad(wc.mo)}-${pad(wc.d)} ${pad(wc.H)}:${pad(wc.Mi)}:${pad(wc.S)}`;
  }


  return { boot };
})();

window.addEventListener("DOMContentLoaded", App.boot);
