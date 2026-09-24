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
    DatePicker.attach(qs("subStart"));
    DatePicker.attach(qs("subEnd"));
    qs("btnSubModalClose").addEventListener("click", closeSubscriptionModal);
    qs("btnSubModalCancel").addEventListener("click", closeSubscriptionModal);
    qs("btnSubModalSave").addEventListener("click", handleSaveSubscription);
    qs("subModalOverlay").addEventListener("click", (e) => { if (e.target.id === "subModalOverlay") closeSubscriptionModal(); });
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
    qs("btnRunCompare").addEventListener("click", handleRunCompare);
    qs("btnDownloadCompare").addEventListener("click", handleDownloadCompare);
    qs("btnRunExport").addEventListener("click", handleRunExport);
  }

  function showView(view) {
    currentView = view;
    Object.entries(els).forEach(([k, el]) => el.classList.toggle("active", k === view));
    document.querySelectorAll(".nav-btn[data-view]").forEach(b => b.classList.toggle("active", b.dataset.view === view));
    if (view === "extraction") populateExtractionCompanySelect();
    if (view === "compare") populateCompareCompanySelect();
    if (view === "export") populateExportCompanySelect();
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
          <div class="company-sub">${brand.label} · ${(conn.stations || []).length} station(s)${subscriptionSummary(conn)}</div>
        </div>
        <div class="company-status">${brand.confidence === "verified" ? "<span class=\"pill ok\">Ready</span>" : "<span class=\"pill warn\">Untested endpoints</span>"}${missingBadge}</div>
        <div class="company-actions">
          <button class="btn ${conn.subscription?.hourly || conn.subscription?.monthly ? "btn-primary" : "btn-ghost"}" data-action="subscription" data-id="${conn.id}">
            Subscription${conn.subscription?.hourly || conn.subscription?.monthly ? " ✓" : ""}
          </button>
          <button class="btn btn-ghost" data-action="extract" data-id="${conn.id}">Extract</button>
          ${currentRole === "Admin" ? `<button class="btn btn-ghost" data-action="delete" data-id="${conn.id}">Remove</button>` : ""}
        </div>`;
      root.appendChild(row);
    }
    root.querySelectorAll("[data-action='view-stations']").forEach(el => el.addEventListener("click", () => openStationsModal(el.dataset.id)));
    root.querySelectorAll("[data-action='subscription']").forEach(b => b.addEventListener("click", (e) => { e.stopPropagation(); openSubscriptionModal(b.dataset.id); }));
    root.querySelectorAll("[data-action='extract']").forEach(b => b.addEventListener("click", () => { showView("extraction"); qs("extCompany").value = b.dataset.id; handleExtractionCompanyChange(); }));
    root.querySelectorAll("[data-action='delete']").forEach(b => b.addEventListener("click", (e) => { e.stopPropagation(); handleDeleteConnection(b.dataset.id); }));
  }

  function subscriptionSummary(conn) {
    const s = conn.subscription;
    if (!s || (!s.hourly && !s.monthly)) return "";
    const parts = [];
    if (s.hourly) parts.push("Hourly");
    if (s.monthly) parts.push("Monthly");
    return ` · Subscribed (${parts.join(" + ")}, ${s.startDate || "?"} → ${s.endDate || "?"})`;
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

  let subscriptionModalConnId = null;

  function openSubscriptionModal(id) {
    const conn = connections.find(c => c.id === id);
    if (!conn) return;
    subscriptionModalConnId = id;
    const s = conn.subscription || {};
    qs("subModalTitle").textContent = `${conn.companyName} — Subscription`;
    qs("subHourly").checked = !!s.hourly;
    qs("subMonthly").checked = !!s.monthly;
    if (s.startDate) DatePicker.setFromISO(qs("subStart"), s.startDate);
    else { qs("subStart").value = ""; qs("subStart").dataset.iso = ""; }
    if (s.endDate) DatePicker.setFromISO(qs("subEnd"), s.endDate);
    else { qs("subEnd").value = ""; qs("subEnd").dataset.iso = ""; }
    qs("subModalOverlay").classList.add("active");
  }
  function closeSubscriptionModal() { qs("subModalOverlay").classList.remove("active"); }

  async function handleSaveSubscription() {
    const conn = connections.find(c => c.id === subscriptionModalConnId);
    if (!conn) return;
    const hourly = qs("subHourly").checked;
    const monthly = qs("subMonthly").checked;
    const startDate = DatePicker.getISO(qs("subStart"));
    const endDate = DatePicker.getISO(qs("subEnd"));
    if ((hourly || monthly) && (!startDate || !endDate)) {
      alert("Pick a start and end date for the subscription.");
      return;
    }
    if (endDate && startDate && endDate < startDate) {
      alert("End date is before the start date.");
      return;
    }
    const prevSubscription = conn.subscription;
    const prevDaily = conn.dailyAutoExtract;
    conn.subscription = { hourly, monthly, startDate, endDate };
    conn.dailyAutoExtract = hourly || monthly; // kept for the "on/off" style display + backward compatibility
    try {
      await SheetsClient.saveConnection(conn);
      closeSubscriptionModal();
      renderDashboard();
    } catch (e) {
      conn.subscription = prevSubscription;
      conn.dailyAutoExtract = prevDaily;
      alert(`Could not save: ${e.message}`);
    }
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

  async function handleRunExport() {
    const id = qs("expCompany").value;
    const conn = connections.find(c => c.id === id);
    if (!conn) { alert("Pick a company first."); return; }
    const settings = conn.templateSettings;
    if (!settings || (settings.template !== "1" && settings.template !== "2")) {
      alert(`${conn.companyName} isn't set up for Template export yet — open it from the Dashboard and set a Template + facility_id assignments first (or configure it in the Extraction tab's Export format section).`);
      return;
    }
    const startDate = DatePicker.getISO(qs("expStart"));
    const endDate = DatePicker.getISO(qs("expEnd"));
    if (!startDate || !endDate) { alert("Pick a start and end date."); return; }

    qs("btnRunExport").disabled = true;
    qs("expStatus").textContent = "Reading Hourly rows from your Google Sheet…";
    qs("expResultsCard").hidden = true;

    try {
      const readings = await SheetsClient.listReadings(conn.companyName);
      const brandKey = conn.brand;
      const utcOffset = settings.utcOffset ?? 8;
      const filtered = readings.filter(r => {
        if (r.resolution !== "Hourly") return false;
        const wc = TemplateExport.wallClockFromRow(r, brandKey, utcOffset);
        if (!wc) return false;
        const localDate = `${wc.y}-${String(wc.mo).padStart(2, "0")}-${String(wc.d).padStart(2, "0")}`;
        return localDate >= startDate && localDate <= endDate;
      });
      const groups = TemplateExport.build(filtered, brandKey, settings);
      renderExportResults(conn, settings, groups, filtered.length);
      qs("expStatus").textContent = `Done — ${readings.length} row(s) read, ${filtered.length} in range.`;
    } catch (e) {
      qs("expStatus").textContent = `Failed: ${e.message}`;
    }
    qs("btnRunExport").disabled = false;
  }

  function renderExportResults(conn, settings, groups, rowCount) {
    qs("expResultsCard").hidden = false;
    qs("expSummary").innerHTML = groups.length
      ? `${rowCount} hourly row(s) in range, grouped into <strong>${groups.length}</strong> facility file(s).`
      : `No hourly rows found in that date range for this company.`;

    const list = qs("expFacilityList");
    list.innerHTML = "";
    groups.forEach(g => {
      const row = document.createElement("div");
      row.className = "export-facility-row";
      row.innerHTML = `
        <div>
          <div style="font-weight:600;">${escapeHtml(g.facilityId)}</div>
          <div class="meta">${g.rows.length} row(s) · Template ${escapeHtml(settings.template)}</div>
        </div>
        <button class="btn btn-primary" data-facility="${escapeHtml(g.facilityId)}">Download</button>`;
      row.querySelector("button").addEventListener("click", () => downloadExportFacility(conn, settings, g));
      list.appendChild(row);
    });
  }

  function downloadExportFacility(conn, settings, group) {
    const sheetName = settings.template === "1" ? "Data" : "MeterData";
    const filenameSuffix = settings.template === "1" ? "template_1" : "template_2";
    const ws = XLSX.utils.aoa_to_sheet([group.headers, ...group.rows]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    const safeFacility = group.facilityId.replace(/[^a-z0-9_-]+/gi, "_");
    XLSX.writeFile(wb, `${conn.companyName}_${safeFacility}_export_${filenameSuffix}.xlsx`);
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
    loadTemplateSettingsIntoForm(conn);
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

    const cancelBtn = document.createElement("button");
    cancelBtn.className = "btn btn-ghost job-cancel";
    cancelBtn.textContent = "Cancel";
    cancelBtn.addEventListener("click", () => {
      ExtractionEngine.stop(job.id);
      cancelBtn.disabled = true;
      cancelBtn.textContent = "Cancelling…";
    });
    jobRow.appendChild(cancelBtn);

    await ExtractionEngine.run(job.id, ctx);

    if ((job.status === "done" || job.status === "paused" || job.status === "stopped") && job.rowsCollected.length) {
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
    const rows = job.rowsCollected.map(r => ({
      Timestamp: r.timestamp,
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

  return { boot };
})();

window.addEventListener("DOMContentLoaded", App.boot);
