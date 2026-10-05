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
  let currentEmail = null;   // signed-in user's email — stamped on every audit entry
  let settingsUnlocked = false; // Settings tab is password-locked; re-locks whenever you leave it
  let stationsModalDirty = false;

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
      currentEmail = who.email;
      qs("authGateOverlay").classList.remove("active");
      qs("authUserLabel").textContent = `${who.email} · ${who.role}`;
      qs("navAudit").hidden = who.role !== "Admin";
      Audit.log("Sign in", "", `Role: ${who.role}`);
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
    Audit.log("Sign out", "", "");
    currentEmail = null;
    AuthClient.signOut();
    location.reload(); // simplest reliable way back to a clean, gated state
  }

  function cacheEls() {
    ["dashboard", "extraction", "compare", "export", "settings", "audit"].forEach(v => els[v] = qs(`view-${v}`));
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
    populateCompareMonthSelects();
    qs("btnSettingsUnlock").addEventListener("click", handleSettingsUnlock);
    qs("settingsPasswordInput").addEventListener("keydown", (e) => { if (e.key === "Enter") handleSettingsUnlock(); });
    qs("btnSettingsLockCancel").addEventListener("click", closeSettingsLock);
    qs("btnSettingsLockClose").addEventListener("click", closeSettingsLock);
    qs("btnSettingsRelock").addEventListener("click", () => { settingsUnlocked = false; Audit.log("Settings locked", "", ""); showView("dashboard"); });
    qs("btnAuditRefresh").addEventListener("click", loadAuditLog);
    qs("btnAuditDownload").addEventListener("click", downloadAuditLog);
    ["auditFilterUser", "auditFilterAction", "auditFilterText"].forEach(id => qs(id).addEventListener("input", renderAuditTable));
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
    qs("btnResumeAll").addEventListener("click", handleResumeAll);
    qs("expResolution").addEventListener("change", syncExportFieldVisibility);
    qs("expCompany").addEventListener("change", handleExportCompanyChange);

    // Excel import — same modal from both tabs, pre-set to whatever that tab has selected.
    ExcelImport.init({
      getConnections: () => connections,
      audit: (action, company, details) => Audit.log(action, company, details),
      onImported: handleExcelImported,
    });
    qs("btnImportFromCompare").addEventListener("click", () =>
      ExcelImport.open({ companyId: qs("cmpCompany").value, resolution: "Hourly" }));
    qs("btnImportFromExport").addEventListener("click", () =>
      ExcelImport.open({ companyId: qs("expCompany").value, resolution: qs("expResolution").value }));
  }

  // After an import, select that company on Compare and Export and reload
  // Compare's period list so the newly imported months show up straight away.
  // A new Excel-import company was also just added to Connections, so
  // reload the list before selecting it.
  async function handleExcelImported(conn) {
    await refreshConnections();
    const id = (connections.find(c => c.id === conn.id || c.companyName === conn.companyName) || conn).id;
    if (currentView === "compare") {
      populateCompareCompanySelect();
      qs("cmpCompany").value = id;
      refreshComparePeriods();
    } else if (currentView === "export") {
      populateExportCompanySelect();
      qs("expCompany").value = id;
    }
  }

  function showView(view) {
    if (view === "settings" && !settingsUnlocked) { openSettingsLock(); return; }
    if (view === "audit" && currentRole !== "Admin") { alert("Only Admins can view the audit log."); return; }
    if (currentView === "settings" && view !== "settings") settingsUnlocked = false; // leaving Settings re-locks it
    currentView = view;
    Object.entries(els).forEach(([k, el]) => el.classList.toggle("active", k === view));
    document.querySelectorAll(".nav-btn[data-view]").forEach(b => b.classList.toggle("active", b.dataset.view === view));
    if (view === "extraction") { populateExtractionCompanySelect(); renderPersistedPendingJobs(); }
    if (view === "compare") populateCompareCompanySelect();
    if (view === "export") { populateExportCompanySelect(); syncExportFieldVisibility(); }
    if (view === "audit") loadAuditLog();
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
    Audit.log("Settings saved", "", `Proxy: ${qs("cfgProxyUrl").value.trim() || "(blank)"} · Sheet ID: ${qs("cfgSpreadsheetId").value.trim() || "(blank)"}${clientIdChanged ? " · Sign-in Client ID changed" : ""}`);
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
        <div class="company-status">${brand.manual ? "<span class=\"pill ok\">Imported data</span>" : brand.confidence === "verified" ? "<span class=\"pill ok\">Ready</span>" : "<span class=\"pill warn\">Untested endpoints</span>"}${missingBadge}</div>
        <div class="company-actions">
          ${brand.manual
            ? `<button class="btn btn-ghost" data-action="import" data-id="${conn.id}">Import</button>`
            : `<button class="btn btn-ghost" data-action="extract" data-id="${conn.id}">Extract</button>`}
          ${currentRole === "Admin" ? `<button class="btn btn-ghost" data-action="delete" data-id="${conn.id}">Remove</button>` : ""}
        </div>`;
      root.appendChild(row);
    }
    root.querySelectorAll("[data-action='view-stations']").forEach(el => el.addEventListener("click", () => openStationsModal(el.dataset.id)));
    root.querySelectorAll("[data-action='extract']").forEach(b => b.addEventListener("click", () => { showView("extraction"); qs("extCompany").value = b.dataset.id; handleExtractionCompanyChange(); }));
    root.querySelectorAll("[data-action='import']").forEach(b => b.addEventListener("click", () => ExcelImport.open({ companyId: b.dataset.id })));
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
    stationsModalDirty = false;
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
    qs("stationsModalMonthlyCheck").innerHTML = "";
    try {
      const summary = await getReadingsSummary(conn);
      renderDataStoredSummary(table, summary);
      verifyUnfinishedMonthly(conn, summary);
    } catch (e) {
      table.innerHTML = `<tbody><tr><td class="field-help">Could not load: ${escapeHtml(e.message)}</td></tr></tbody>`;
    }
  }

  function renderDataStoredSummary(table, summary) {
    const order = ["Hourly", "Daily", "Monthly"];
    const keys = Object.keys(summary || {}).sort((a, b) => ((order.indexOf(a) + 1) || 99) - ((order.indexOf(b) + 1) || 99));
    if (!keys.length) {
      table.innerHTML = `<tbody><tr><td class="field-help">Nothing extracted yet for this company.</td></tr></tbody>`;
      return;
    }
    const cell = (k) => {
      const s = summary[k];
      if (k !== "Monthly") return [escapeHtml(s.earliest), escapeHtml(s.latest)];
      // Monthly rows are one per month (dated the 1st), so show "Sep 2026"
      // rather than a day. Latest = the latest month that's actually over;
      // rows for the current (or a later) month are checked separately by
      // verifyUnfinishedMonthly() and noted underneath.
      const months = monthsOfSummary(s);
      if (!months.length) return ["?", "?"];
      const now = new Date();
      const thisMonth = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;
      const finished = months.filter(m => m < thisMonth);
      const notOver = months.filter(m => m >= thisMonth);
      const latest = finished.length ? monthYearLabel(finished[finished.length - 1]) : "—";
      const note = notOver.length
        ? `<div class="field-help" style="margin-top:2px;">${escapeHtml(notOver.map(monthYearLabel).join(", "))} also stored, though that month isn't over — checked below</div>`
        : "";
      return [escapeHtml(monthYearLabel(months[0])), escapeHtml(latest) + note];
    };
    table.innerHTML = `<thead><tr><th>Resolution</th><th>Rows</th><th>Earliest</th><th>Latest</th></tr></thead><tbody>` +
      keys.map(k => { const [first, last] = cell(k);
        return `<tr><td>${escapeHtml(k)}</td><td>${summary[k].rows}</td><td>${first}</td><td>${last}</td></tr>`; }).join("") +
      `</tbody>`;
  }

  // Months a resolution has data for: the proxy's per-month list, or — if it
  // doesn't send one — every month between earliest and latest.
  function monthsOfSummary(s) {
    let months = (s?.months || []).map(m => String(m).slice(0, 7)).sort();
    if (!months.length && /^\d{4}-\d{2}/.test(s?.earliest || "") && /^\d{4}-\d{2}/.test(s?.latest || "")) {
      for (let [y, m] = String(s.earliest).slice(0, 7).split("-").map(Number); ; ) {
        const v = `${y}-${String(m).padStart(2, "0")}`;
        months.push(v);
        if (v >= String(s.latest).slice(0, 7) || months.length > 600) break;
        if (++m > 12) { m = 1; y++; }
      }
    }
    return months;
  }

  // A Monthly row for a month that isn't over can't be a full month's total,
  // so read those rows and test each one against the rest of the data:
  //   - a month that hasn't started, or a row fetched before its month began → wrong date
  //   - the same value as the previous month → that month stored again under the wrong date
  //   - the previous month's hourly total matching it instead of the previous month's own
  //     Monthly figure → dates shifted by a month
  //   - far more kWh than the days elapsed could produce (from last month's daily average)
  //   - 0 kWh → a placeholder, not real generation yet
  //   - otherwise a genuine month-to-date figure (cross-checked with hourly data if there is any)
  async function verifyUnfinishedMonthly(conn, summary) {
    const box = qs("stationsModalMonthlyCheck");
    const now = new Date();
    const thisMonth = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;
    const suspect = monthsOfSummary(summary?.Monthly).filter(m => m >= thisMonth);
    if (!suspect.length) { box.innerHTML = ""; return; }
    const [ty, tm] = thisMonth.split("-").map(Number);
    const prevMonth = tm === 1 ? `${ty - 1}-12` : `${ty}-${String(tm - 1).padStart(2, "0")}`;
    const last = suspect[suspect.length - 1];
    const label = monthYearLabel;
    box.innerHTML = `<div class="field-help" style="margin-top:10px;">Checking the ${escapeHtml(suspect.map(label).join(", "))} Monthly row(s)…</div>`;

    let rows;
    try {
      rows = dedupeReadings(await SheetsClient.listReadings(conn.companyName, { fromDate: `${prevMonth}-01`, toDate: monthEnd(`${last}-01`) }));
    } catch (e) {
      box.innerHTML = `<div class="field-help" style="margin-top:10px;color:var(--warn);">Couldn't check the ${escapeHtml(suspect.map(label).join(", "))} Monthly row(s): ${escapeHtml(e.message)}</div>`;
      return;
    }
    const monthOf = (r) => (extractDateOnly(r.timestamp) || "").slice(0, 7);
    const dayOf = (r) => Number((extractDateOnly(r.timestamp) || "").slice(8, 10));
    const res = (r) => String(r.resolution).trim();
    const monthly = (st, m) => rows.find(r => res(r) === "Monthly" && r.stationId === st && monthOf(r) === m);
    const hourly = (st, m, uptoDay = 31) => {
      const list = rows.filter(r => res(r) === "Hourly" && r.stationId === st && monthOf(r) === m && dayOf(r) <= uptoDay);
      return { kwh: list.reduce((t, r) => t + (Number(r.kwh) || 0), 0), days: new Set(list.map(dayOf)).size };
    };
    const daysIn = (m) => { const [y, mo] = m.split("-").map(Number); return new Date(Date.UTC(y, mo, 0)).getUTCDate(); };
    const close = (a, b, pct) => Math.abs(a - b) <= Math.max(0.01, Math.abs(b) * pct);
    const fmt = (n) => (Math.round(n * 100) / 100).toLocaleString("en-GB");

    const checks = [];
    for (const r of rows.filter(r => res(r) === "Monthly" && suspect.includes(monthOf(r)))) {
      const m = monthOf(r), kwh = Number(r.kwh) || 0;
      const fetched = r.runAt && !isNaN(new Date(r.runAt)) ? localISODate(new Date(r.runAt)) : "";
      const name = r.stationName || r.stationId;
      let verdict, kind;
      const prev = monthly(r.stationId, prevMonth);
      const prevKwh = prev ? Number(prev.kwh) || 0 : null;
      const prevHourly = hourly(r.stationId, prevMonth);
      if (m > thisMonth) {
        kind = "bad"; verdict = `${label(m)} hasn't started yet, so this row can't be real data — its date is wrong.`;
      } else if (fetched && fetched < `${m}-01`) {
        kind = "bad"; verdict = `Fetched on ${fetched}, before ${label(m)} began, so it can't be ${label(m)} data — its date is wrong.`;
      } else if (prevKwh != null && kwh > 0 && close(kwh, prevKwh, 0.005)) {
        kind = "bad"; verdict = `Same value as ${label(prevMonth)} (${fmt(prevKwh)} kWh) — looks like ${label(prevMonth)}'s total stored again under ${label(m)}.`;
      } else if (prevHourly.days >= daysIn(prevMonth) && kwh > 0 && close(kwh, prevHourly.kwh, 0.01)
                 && (prevKwh == null || !close(prevKwh, prevHourly.kwh, 0.01))) {
        kind = "bad"; verdict = `Matches ${label(prevMonth)}'s hourly total (${fmt(prevHourly.kwh)} kWh) rather than anything in ${label(m)} — Monthly dates look shifted by one month.`;
      } else {
        const fetchedDay = fetched && fetched.slice(0, 7) === m ? Number(fetched.slice(8, 10)) : null;
        const daysCovered = fetchedDay || now.getDate();
        const expected = prevKwh ? (prevKwh / daysIn(prevMonth)) * daysCovered : null;
        const h = hourly(r.stationId, m, daysCovered);
        if (kwh === 0) {
          kind = "info"; verdict = `0 kWh${fetched ? ` (fetched ${fetched})` : ""} — a placeholder for a month that has only just started, not real generation yet.`;
        } else if (expected != null && kwh > expected * 2 + 1 && kwh > prevKwh * 0.5) {
          kind = "bad"; verdict = `Too high for ${daysCovered} day(s): ${fmt(kwh)} kWh, against about ${fmt(expected)} kWh expected from ${label(prevMonth)}'s daily average. Probably a full month's total under the wrong date.`;
        } else if (h.days && !close(kwh, h.kwh, 0.05)) {
          kind = "warn"; verdict = `Month-to-date figure (1–${daysCovered} ${label(m)}), but it doesn't match the hourly data for those days (${fmt(h.kwh)} kWh over ${h.days} day(s)).`;
        } else {
          kind = "ok"; verdict = `Real month-to-date figure: 1–${daysCovered} ${label(m)}${fetched ? `, fetched ${fetched}` : ""}${h.days ? `, and it matches the hourly data` : ""}. Not a full month — extract ${label(m)} again after it ends to replace it.`;
        }
      }
      checks.push({ m, name, kwh, fetched, verdict, kind });
    }
    if (!checks.length) { box.innerHTML = ""; return; }
    const color = { bad: "var(--err)", warn: "var(--warn)", info: "var(--muted)", ok: "var(--ok)" };
    const icon = { bad: "✕", warn: "⚠", info: "○", ok: "✓" };
    const bad = checks.filter(c => c.kind === "bad").length;
    box.innerHTML = `
      <div class="callout${bad ? " warn" : ""}" style="margin:12px 0 8px;">
        <strong>Check of unfinished month(s): ${escapeHtml(suspect.map(label).join(", "))}</strong>
        ${bad ? `${bad} of ${checks.length} Monthly row(s) don't hold up — see below. Compare will show them against the hourly data until they're re-extracted.`
              : `The ${checks.length} Monthly row(s) for a month that isn't over were checked against the rest of the data.`}
      </div>
      <div class="table-scroll" style="max-height:200px;"><table>
        <thead><tr><th>Month</th><th>Station</th><th>kWh</th><th>Fetched</th><th>Result</th></tr></thead><tbody>
        ${checks.map(c => `<tr><td style="white-space:nowrap;">${escapeHtml(label(c.m))}</td><td>${escapeHtml(c.name)}</td><td>${fmt(c.kwh)}</td>
          <td style="white-space:nowrap;">${escapeHtml(c.fetched || "unknown")}</td>
          <td style="color:${color[c.kind]};">${icon[c.kind]} ${escapeHtml(c.verdict)}</td></tr>`).join("")}
        </tbody></table></div>`;
  }

  // "2026-09" or "2026-09-01" -> "Sep 2026" (built from the text itself, no timezone shifts).
  function monthYearLabel(ym) {
    const m = String(ym || "").match(/^(\d{4})-(\d{2})/);
    if (!m) return String(ym || "");
    return `${["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"][+m[2] - 1]} ${m[1]}`;
  }

  function renderDataStoredTable(table, readings) {
    if (!readings.length) {
      table.innerHTML = `<tbody><tr><td class="field-help">Nothing extracted yet for this company.</td></tr></tbody>`;
      return;
    }
    // Grouped by resolution. Dates are extracted as plain text from the
    // timestamp string — deliberately NOT run through `new Date(...)`, since
    // that would interpret an already-local time string as the *viewer's own
    // browser* timezone, then shift it when reformatting, silently rolling
    // the date back or forward depending on the viewer's own UTC offset.
    const byResolution = {};
    for (const r of readings) {
      const bucket = byResolution[r.resolution] || (byResolution[r.resolution] = { count: 0, min: null, max: null });
      bucket.count++;
      const dateStr = extractDateOnly(r.timestamp);
      if (!dateStr) continue;
      if (!bucket.min || dateStr < bucket.min) bucket.min = dateStr;
      if (!bucket.max || dateStr > bucket.max) bucket.max = dateStr;
    }
    const order = ["Hourly", "Daily", "Monthly"];
    const resolutions = Object.keys(byResolution).sort((a, b) => order.indexOf(a) - order.indexOf(b));
    table.innerHTML = `<thead><tr><th>Resolution</th><th>Rows</th><th>Earliest</th><th>Latest</th></tr></thead><tbody>` +
      resolutions.map(res => {
        const b = byResolution[res];
        return `<tr><td>${escapeHtml(res)}</td><td>${b.count}</td><td>${b.min || "?"}</td><td>${b.max || "?"}</td></tr>`;
      }).join("") + `</tbody>`;
  }

  // Current format: a plain "YYYY-MM-DD HH:MM:SS" local-time string (already
  // correctly localized when it was written) — just take the date prefix,
  // no Date object involved at all. Falls back to reading an older raw-epoch
  // row's UTC calendar date (getUTC*, never local getters) for rows written
  // before the readable-timestamp fix.
  function extractDateOnly(timestamp) {
    const s = String(timestamp).trim();
    const m = s.match(/^(\d{4}-\d{2}-\d{2})/);
    if (m) return m[1];
    const n = Number(s);
    if (!isNaN(n) && n > 1e10) {
      const d = new Date(n);
      return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}-${String(d.getUTCDate()).padStart(2, "0")}`;
    }
    return null;
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
      stationsModalDirty = true;
      // AGG IDs may have changed — refresh any grouping picker showing this company.
      refreshGroupingPickers(conn.id);
    } catch (e) {
      qs("stationsModalSavedNote").textContent = `Could not save: ${e.message}`;
      qs("stationsModalSavedNote").style.color = "var(--err)";
    }
  }

  function closeStationsModal() {
    qs("stationsModalOverlay").classList.remove("active");
    if (stationsModalDirty) {
      const conn = connections.find(c => c.id === stationsModalConnId);
      const map = conn?.templateSettings?.stationFacility || {};
      Audit.log("Edit company facility settings", conn?.companyName || "", Object.entries(map).map(([st, f]) => `${st}→${f || "(none)"}`).join(", "));
      stationsModalDirty = false;
    }
    if (currentView === "dashboard") renderDashboard(); // refresh the missing-facility_id badge without a full reload
  }

  async function handleDeleteConnection(id) {
    if (!confirm("Remove this company connection? This does not delete anything on the vendor side.")) return;
    const removed = connections.find(c => c.id === id);
    try {
      await SheetsClient.deleteConnection(id);
      Audit.log("Remove company", removed?.companyName || id, removed ? `Brand: ${BRANDS[removed.brand]?.label || removed.brand}` : "");
      await refreshConnections();
    } catch (e) {
      alert(`Could not remove: ${e.message}`);
    }
  }

  /* ---------------------------- add company modal ---------------------------- */

  function renderBrandGrid() {
    const grid = qs("brandGrid");
    grid.innerHTML = "";
    Object.values(BRANDS).filter(b => !b.manual).forEach(b => { // Excel-import companies are made from the import, not here
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

    const conn = { companyName, brand: pendingBrand, credentials: creds, stations, cursor: {} };
    try {
      await SheetsClient.saveConnection(conn);
    } catch (e) {
      alert(`Could not save connection (${e.message}). Check Settings → Proxy base URL and Google Sheet ID.`);
      return;
    }
    Audit.log("Add company", companyName, `Brand: ${brand.label} · ${testFirst ? `Tested, ${stations.length} station(s) found` : "Saved without test"}`);
    closeAddCompanyModal();
    await refreshConnections();
  }

  /* ---------------------------- extraction ---------------------------- */

  function populateExtractionCompanySelect() {
    const sel = qs("extCompany");
    sel.innerHTML = `<option value="">Select a company…</option>`;
    connections.filter(c => !BRANDS[c.brand]?.manual).forEach(c => { // Excel-import companies have no API to extract from
      const opt = document.createElement("option");
      opt.value = c.id; opt.textContent = `${c.companyName} (${BRANDS[c.brand].label})`;
      sel.appendChild(opt);
    });
    renderGroupingPicker("extGroupingWrap", null);
  }

  /* ---------------------------- AGG ID grouping picker ---------------------------- */
  // One choice per company, shared by Extraction (its Download Excel), Compare
  // and Export, so picking "separate" on one tab carries over to the others.
  // Kept for this session only — it doesn't change the company's saved settings.
  //   combined  stations sharing an AGG ID (facility_id) are summed (the old behaviour)
  //   separate  every station on its own
  //   custom    tick which AGG IDs to combine; unticked ones are split per station
  const groupingByConn = {};
  const GROUPING_WRAPS = ["extGroupingWrap", "cmpGroupingWrap", "expGroupingWrap"];

  function getGrouping(connId) {
    return groupingByConn[connId] || { mode: "combined", split: [] };
  }

  function stationNameMap(conn) {
    return Object.fromEntries((conn?.stations || []).map(s => [s.id, s.name]));
  }

  // AGG IDs that actually have more than one station — the only ones where
  // combining vs splitting makes a difference.
  function multiStationAggIds(conn) {
    const byAgg = new Map();
    for (const st of conn.stations || []) {
      const agg = TemplateExport.facilityKeyFor(st.id, conn.templateSettings || {});
      if (agg === st.id) continue; // no AGG ID assigned — always on its own
      if (!byAgg.has(agg)) byAgg.set(agg, []);
      byAgg.get(agg).push(st.name || st.id);
    }
    return [...byAgg.entries()].filter(([, names]) => names.length > 1).sort((a, b) => a[0].localeCompare(b[0]));
  }

  function renderGroupingPicker(wrapId, conn) {
    const wrap = qs(wrapId);
    if (!wrap) return;
    if (!conn) { wrap.innerHTML = ""; wrap.dataset.conn = ""; return; }
    wrap.dataset.conn = conn.id;
    const g = getGrouping(conn.id);
    const multi = multiStationAggIds(conn);
    const opt = (v, label) => `<option value="${v}"${g.mode === v ? " selected" : ""}>${label}</option>`;
    const customList = g.mode === "custom" && multi.length
      ? `<div class="grp-custom">${multi.map(([agg, names]) => `
          <label class="grp-row"><input type="checkbox" data-agg="${escapeHtml(agg)}"${(g.split || []).includes(agg) ? "" : " checked"}>
            <span><strong>${escapeHtml(agg)}</strong> <span class="grp-names">${names.length} stations: ${escapeHtml(names.join(", "))}</span></span></label>`).join("")}
         </div><div class="field-help">Ticked = combined into one. Unticked = each station kept separate.</div>`
      : "";
    const note = multi.length
      ? (g.mode === "custom" ? "" : `<div class="field-help">${multi.length} AGG ID(s) here have more than one station.</div>`)
      : `<div class="field-help">No AGG ID has more than one station in this company, so this choice makes no difference.</div>`;
    wrap.innerHTML = `<div class="field" style="margin:0;max-width:520px;">
        <label>AGG ID grouping</label>
        <select class="grp-mode">
          ${opt("combined", "Combine stations by AGG ID")}
          ${opt("separate", "Keep every station separate")}
          ${opt("custom", "Choose which AGG IDs to combine…")}
        </select>
      </div>${customList}${note}`;

    wrap.querySelector(".grp-mode").addEventListener("change", (e) => {
      const prev = getGrouping(conn.id);
      setGrouping(conn, { mode: e.target.value, split: prev.split || [] });
    });
    wrap.querySelectorAll("input[data-agg]").forEach(cb => cb.addEventListener("change", () => {
      const split = [...wrap.querySelectorAll("input[data-agg]")].filter(x => !x.checked).map(x => x.dataset.agg);
      setGrouping(conn, { mode: "custom", split });
    }));
  }

  function setGrouping(conn, grouping) {
    groupingByConn[conn.id] = grouping;
    refreshGroupingPickers(conn.id);
    // Results already on screen were built with the old choice.
    if (lastCompareResult?.conn.id === conn.id && !qs("cmpResultsCard").hidden)
      qs("cmpStatus").textContent = "AGG ID grouping changed — click Run comparison to update the results.";
    if (qs("expCompany").value === conn.id && !qs("expResultsCard").hidden)
      qs("expStatus").textContent = "AGG ID grouping changed — click Build export to update the files.";
  }

  function refreshGroupingPickers(connId) {
    const conn = connections.find(c => c.id === connId);
    GROUPING_WRAPS.forEach(id => { if (qs(id)?.dataset.conn === connId) renderGroupingPicker(id, conn); });
  }

  // Group label used in tables, file lists and the audit log.
  function groupLabel(g) {
    return g.stationName ? `${g.facilityId} · ${g.stationName}` : g.facilityId;
  }

  /* ---------------------------- compare (hourly vs monthly) ---------------------------- */

  let lastCompareResult = null; // kept around for the Download Excel button

  function populateCompareCompanySelect() {
    const sel = qs("cmpCompany");
    const keep = sel.value;
    sel.innerHTML = `<option value="">Select a company…</option>`;
    connections.forEach(c => {
      const opt = document.createElement("option");
      opt.value = c.id; opt.textContent = `${c.companyName} (${BRANDS[c.brand].label})`;
      sel.appendChild(opt);
    });
    if (keep && connections.some(c => c.id === keep)) sel.value = keep;
    if (!sel.dataset.periodHooked) {
      sel.dataset.periodHooked = "1";
      sel.addEventListener("change", refreshComparePeriods);
    }
    refreshComparePeriods();
  }

  // What's been extracted for a company, per resolution: { Hourly: { rows,
  // earliest, latest, months: ["2026-07", …] }, Monthly: {…} }. Uses the
  // proxy's light readingsSummary action; falls back to reading the rows if
  // the proxy hasn't been updated with it yet.
  async function getReadingsSummary(conn) {
    try {
      return await SheetsClient.readingsSummary(conn.companyName);
    } catch (e) {
      if (!/Unknown action/i.test(e.message)) throw e;
      const out = {};
      for (const r of dedupeReadings(await SheetsClient.listReadings(conn.companyName))) {
        const d = extractDateOnly(r.timestamp);
        if (!d) continue;
        const res = String(r.resolution).trim();
        const s = out[res] || (out[res] = { rows: 0, earliest: d, latest: d, months: [] });
        s.rows++;
        if (d < s.earliest) s.earliest = d;
        if (d > s.latest) s.latest = d;
        if (!s.months.includes(d.slice(0, 7))) s.months.push(d.slice(0, 7));
      }
      Object.values(out).forEach(s => s.months.sort());
      return out;
    }
  }

  // Compare period. The list is built from what's actually been extracted
  // for the selected company, from January of THIS year onwards:
  //   - a quarter is offered only if it contains Hourly or Monthly data;
  //   - each shows how many of its months have Hourly / Monthly data;
  //   - "Custom range…" From/To months run from the first to the last month
  //     with data (this year onwards).
  // Picking a quarter fills in From/To and locks them. Default: the latest
  // quarter that has data. Quarters: Q1 Jan–Mar, Q2 Apr–Jun, Q3 Jul–Sep, Q4 Oct–Dec.
  const QUARTER_MONTHS = ["Jan–Mar", "Apr–Jun", "Jul–Sep", "Oct–Dec"];
  let comparePeriodSeq = 0;

  function populateCompareMonthSelects() {
    qs("cmpPeriod").addEventListener("change", applyComparePeriod);
    setComparePeriodPlaceholder("Select a company first");
  }

  function setComparePeriodPlaceholder(text) {
    qs("cmpPeriod").innerHTML = `<option value="">${escapeHtml(text)}</option>`;
    qs("cmpPeriod").disabled = true;
    ["cmpFromMonth", "cmpToMonth"].forEach(id => { qs(id).innerHTML = ""; qs(id).disabled = true; });
    qs("btnRunCompare").disabled = true;
  }

  async function refreshComparePeriods() {
    const conn = connections.find(c => c.id === qs("cmpCompany").value);
    renderGroupingPicker("cmpGroupingWrap", conn);
    if (!conn) { setComparePeriodPlaceholder("Select a company first"); return; }
    const seq = ++comparePeriodSeq; // ignore answers for a company that's no longer selected
    const keep = qs("cmpPeriod").value;
    setComparePeriodPlaceholder("Checking extracted dates…");
    let summary;
    try {
      summary = await getReadingsSummary(conn);
    } catch (e) {
      if (seq === comparePeriodSeq) setComparePeriodPlaceholder(`Could not load dates (${e.message})`);
      return;
    }
    if (seq !== comparePeriodSeq) return;

    const yearStart = `${new Date().getFullYear()}-01`;
    const monthsOf = (res) => new Set((summary[res]?.months || []).filter(m => m >= yearStart));
    const hourly = monthsOf("Hourly"), monthly = monthsOf("Monthly");
    const withData = [...new Set([...hourly, ...monthly])].sort();
    if (!withData.length) {
      setComparePeriodPlaceholder(`No Hourly or Monthly data extracted for ${new Date().getFullYear()} yet`);
      return;
    }

    // Quarters that contain data, newest first.
    const now = new Date(), curKey = `Q${Math.ceil((now.getMonth() + 1) / 3)}-${now.getFullYear()}`;
    const quarters = new Map();
    for (const m of withData) {
      const [y, mo] = m.split("-").map(Number);
      const key = `Q${Math.ceil(mo / 3)}-${y}`;
      if (!quarters.has(key)) quarters.set(key, { y, q: Math.ceil(mo / 3) });
    }
    const opts = [...quarters.entries()].sort((a, b) => (b[1].y - a[1].y) || (b[1].q - a[1].q)).map(([key, { y, q }]) => {
      const qMonths = [0, 1, 2].map(i => `${y}-${String(q * 3 - 2 + i).padStart(2, "0")}`);
      const h = qMonths.filter(m => hourly.has(m)).length, mo = qMonths.filter(m => monthly.has(m)).length;
      return `<option value="${key}">Q${q} ${y} (${QUARTER_MONTHS[q - 1]}) · Hourly ${h}/3 · Monthly ${mo}/3${key === curKey ? " · in progress" : ""}</option>`;
    });
    qs("cmpPeriod").innerHTML = opts.join("") + `<option value="custom">Custom range…</option>`;
    qs("cmpPeriod").disabled = false;

    // Custom-range months: first → last month with data, continuous.
    const months = [];
    for (let [y, m] = withData[0].split("-").map(Number); ; ) {
      const v = `${y}-${String(m).padStart(2, "0")}`;
      months.push(v);
      if (v >= withData[withData.length - 1]) break;
      if (++m > 12) { m = 1; y++; }
    }
    const label = (v) => { const [y, m] = v.split("-").map(Number); return new Date(y, m - 1, 1).toLocaleString("en-GB", { month: "short", year: "numeric" }); };
    const tag = (v) => [hourly.has(v) && "H", monthly.has(v) && "M"].filter(Boolean).join("+") || "no data";
    const html = months.slice().reverse().map(v => `<option value="${v}">${label(v)} (${tag(v)})</option>`).join("");
    qs("cmpFromMonth").innerHTML = html;
    qs("cmpToMonth").innerHTML = html;

    qs("cmpPeriod").value = [...qs("cmpPeriod").options].some(o => o.value === keep) ? keep : qs("cmpPeriod").options[0].value;
    qs("btnRunCompare").disabled = false;
    applyComparePeriod();
  }

  function applyComparePeriod() {
    const v = qs("cmpPeriod").value;
    const custom = v === "custom";
    qs("cmpFromMonth").disabled = !custom;
    qs("cmpToMonth").disabled = !custom;
    const avail = [...qs("cmpToMonth").options].map(o => o.value); // newest first
    if (!avail.length) return;
    if (custom) {
      if (!qs("cmpFromMonth").value) qs("cmpFromMonth").value = avail[avail.length - 1];
      if (!qs("cmpToMonth").value) qs("cmpToMonth").value = avail[0];
      return;
    }
    const m = v.match(/^Q(\d)-(\d{4})$/);
    if (!m) return;
    const q = +m[1], y = +m[2];
    const qMonths = [0, 1, 2].map(i => `${y}-${String(q * 3 - 2 + i).padStart(2, "0")}`);
    // Clamp the quarter to months that exist in the list (e.g. an in-progress quarter).
    const inList = qMonths.filter(x => avail.includes(x));
    const first = inList[0] || qMonths[0], last = inList[inList.length - 1] || qMonths[2];
    if (avail.includes(first)) qs("cmpFromMonth").value = first;
    if (avail.includes(last)) qs("cmpToMonth").value = last;
  }

  // "Q3 2026" for a quarter, "2026-02 → 2026-05" for a custom range.
  function comparePeriodLabel(fromMonth, toMonth) {
    const v = qs("cmpPeriod").value;
    if (v !== "custom") { const [, q, y] = v.match(/^Q(\d)-(\d{4})$/); return `Q${q} ${y}`; }
    return fromMonth === toMonth ? fromMonth : `${fromMonth} → ${toMonth}`;
  }

  async function handleRunCompare() {
    const id = qs("cmpCompany").value;
    const conn = connections.find(c => c.id === id);
    if (!conn) { alert("Pick a company first."); return; }
    let fromMonth = qs("cmpFromMonth").value, toMonth = qs("cmpToMonth").value;
    const qm = qs("cmpPeriod").value.match(/^Q(\d)-(\d{4})$/);
    if (qm) { // a quarter always compares its full 3 months
      fromMonth = `${qm[2]}-${String(+qm[1] * 3 - 2).padStart(2, "0")}`;
      toMonth = `${qm[2]}-${String(+qm[1] * 3).padStart(2, "0")}`;
    }
    if (!fromMonth || !toMonth) { alert("Pick a period first."); return; }
    if (toMonth < fromMonth) { alert("To month is before From month — please fix the range."); return; }
    const fromDate = `${fromMonth}-01`, toDate = monthEnd(`${toMonth}-01`);
    const periodLabel = comparePeriodLabel(fromMonth, toMonth);

    qs("btnRunCompare").disabled = true;
    qs("cmpStatus").textContent = "Reading Hourly and Monthly rows from your Google Sheet…";
    qs("cmpResultsCard").hidden = true;

    try {
      // Range is also applied here, in case an older proxy ignores fromDate/toDate.
      const rawReadings = dedupeReadings(await SheetsClient.listReadings(conn.companyName, { fromDate, toDate }))
        .filter(r => { const d = String(r.timestamp).trim().slice(0, 10); return !/^\d{4}-\d{2}-\d{2}$/.test(d) || (d >= fromDate && d <= toDate); });
      const readings = normalizeReadingsForTemplates(rawReadings, conn.brand, conn.templateSettings?.utcOffset);
      const grouping = getGrouping(conn.id);
      let result = CompareEngine.compareHourlyVsMonthly(readings, conn, grouping);
      // If the engine parses the readable "YYYY-MM-DD HH:MM:SS" strings itself
      // (rather than via wallClockFromRow), the epoch-converted copy would read
      // as no data — so fall back to the rows exactly as stored.
      const hasHourly = rawReadings.some(r => String(r.resolution).trim() === "Hourly");
      const hasMonthly = rawReadings.some(r => String(r.resolution).trim() === "Monthly");
      if ((hasHourly && !result.hourlyMonths.length) || (hasMonthly && !result.monthlyMonths.length)) {
        const alt = CompareEngine.compareHourlyVsMonthly(rawReadings.map(r => ({ ...r, kwh: Number(r.kwh) })), conn, grouping);
        if (alt.hourlyMonths.length + alt.monthlyMonths.length > result.hourlyMonths.length + result.monthlyMonths.length) result = alt;
      }
      lastCompareResult = { conn, result, periodLabel, grouping };
      renderCompareResults(conn, result);
      qs("cmpStatus").textContent = `Done — ${readings.length} row(s) read for ${periodLabel}${periodLabel.startsWith("Q") ? ` (${fromMonth} → ${toMonth})` : ""}.`;
      Audit.log("Comparison", conn.companyName, `Hourly vs Monthly · ${periodLabel} · ${TemplateExport.describeGrouping(grouping)} · ${result.totalChecked} facility-month(s) checked · ${result.flagged.length} need attention · ${(result.notExtracted || []).length} hourly not extracted · ${(result.monthlyMissing || []).length} monthly missing`);
    } catch (e) {
      qs("cmpStatus").textContent = `Failed: ${e.message}`;
    }
    qs("btnRunCompare").disabled = false;
  }

  // One list of every row the Compare results show (and download): months
  // needing attention, months with no hourly data, months with no Monthly figure.
  // Every facility-month Compare knows about — OK ones included — in
  // facility, then month order. Used for both the table and the download.
  function compareRows(result) {
    const rows = [
      ...(result.ok || []),
      ...result.flagged,
      ...(result.notExtracted || []).map(r => ({
        facilityId: r.facilityId, stationId: r.stationId, stationName: r.stationName, month: r.month, kind: "notExtracted",
        monthlyKwh: r.monthlyKwh, hourlyKwh: "", diffKwh: "", diffPct: "",
        hourlyCoverage: "0 days", missingRanges: r.missingRanges || "whole month",
        monthlyAsOf: r.monthlyAsOf || "", monthlyStale: false, complete: false,
        issue: "Hourly not extracted for this month — not checked",
      })),
      ...(result.monthlyMissing || []).map(r => ({
        facilityId: r.facilityId, stationId: r.stationId, stationName: r.stationName, month: r.month, kind: "monthlyMissing",
        monthlyKwh: "", hourlyKwh: r.hourlyKwh, diffKwh: "", diffPct: "",
        hourlyCoverage: r.hourlyCoverage, missingRanges: r.missingRanges || "",
        monthlyAsOf: "", monthlyStale: false, complete: !r.missingRanges,
        issue: "No Monthly figure for this month — run a Monthly extraction to compare"
          + (r.missingRanges ? " · some hourly days missing" : ""),
      })),
    ];
    return rows.sort((a, b) => a.facilityId.localeCompare(b.facilityId)
      || (a.stationName || "").localeCompare(b.stationName || "") || a.month.localeCompare(b.month));
  }

  const COMPARE_STATUS = {
    ok:             { label: "✓ OK",               color: "var(--ok)" },
    mismatch:       { label: "⚠ Mismatch",         color: "var(--warn)" },
    outage:         { label: "⚠ Outage",           color: "var(--err)" },
    zero:           { label: "⚠ 0 kWh",            color: "var(--warn)" },
    incomplete:     { label: "◐ Incomplete hourly", color: "var(--warn)" },
    notExtracted:   { label: "○ No hourly",         color: "var(--muted)" },
    monthlyMissing: { label: "○ No monthly",        color: "var(--muted)" },
  };

  function renderCompareResults(conn, result) {
    qs("cmpResultsCard").hidden = false;
    const noHourly = !result.hourlyMonths.length;
    const noMonthly = !result.monthlyMonths.length;
    const notExtracted = result.notExtracted || [];
    const monthlyMissing = result.monthlyMissing || [];
    const okMonths = result.ok || [];
    const gaps = result.flagged.filter(f => f.hasGap && f.complete);
    const incomplete = result.flagged.filter(f => !f.complete);
    const zero = result.flagged.filter(f => f.kind === "zero");
    const stale = result.flagged.filter(f => f.monthlyStale);

    const lines = [];
    const groupingDesc = TemplateExport.describeGrouping(lastCompareResult?.grouping);
    let head = `Checked ${result.totalChecked} facility-month combination(s) <span style="color:var(--muted);">(${escapeHtml(groupingDesc)})</span>: `;
    if (noHourly || noMonthly) {
      head += `<strong style="color:var(--warn);">Missing ${noHourly ? "Hourly" : "Monthly"} data entirely for this company — run that extraction first.</strong>`;
    } else {
      head += `<span style="color:var(--ok);">${okMonths.length} OK</span>`
        + (result.flagged.length ? `, <strong style="color:var(--warn);">${result.flagged.length} need attention</strong>` : "")
        + (gaps.length ? ` (${gaps.length} with a kWh gap)` : "") + ".";
    }
    lines.push(head);
    if (incomplete.length) lines.push(`<span style="color:var(--warn);">${incomplete.length} month(s) with hourly data missing for some days</span> — see "Missing days".`);
    if (zero.length) lines.push(`<span style="color:var(--warn);">⚠ ${zero.length} month(s) show 0 kWh in both Monthly and Hourly (${escapeHtml(summarizeMonths(zero))})</span> — the plant may have been offline or not reporting.`);
    if (stale.length) lines.push(`<span style="color:var(--warn);">⚠ ${stale.length} Monthly figure(s) were fetched before that month's hourly data ended (${escapeHtml([...new Set(stale.map(f => f.month))].join(", "))})</span>, so they can't include the later days. Re-run the Monthly extraction for those month(s).`);
    if (notExtracted.length && !noHourly) lines.push(`Hourly not extracted (whole month missing): <strong>${escapeHtml(summarizeMonths(notExtracted))}</strong>. Run an Hourly extraction for those dates to include them.`);
    if (monthlyMissing.length && !noMonthly) lines.push(`No Monthly figure to compare against: <strong>${escapeHtml(summarizeMonths(monthlyMissing))}</strong>. Run a Monthly extraction for those month(s).`);
    if (result.unreadable?.count) lines.push(`<span style="color:var(--err);">⚠ ${result.unreadable.count} row(s) skipped — their timestamp couldn't be read</span> (e.g. ${escapeHtml(result.unreadable.samples.join("; "))}). Send this to your developer.`);
    // Keep whatever filters were chosen before a re-run, if they still apply.
    const prev = {
      facility: qs("cmpFilterFacility")?.value || "",
      month: qs("cmpFilterMonth")?.value || "",
      status: qs("cmpFilterStatus")?.value || "",
    };
    const all = compareRows(result);
    const facilities = [...new Set(all.map(r => r.facilityId))].sort();
    const months = [...new Set(all.map(r => r.month))].sort();
    const statuses = Object.keys(COMPARE_STATUS).filter(k => all.some(r => r.kind === k));
    const opt = (v, label, sel) => `<option value="${escapeHtml(v)}"${v === sel ? " selected" : ""}>${escapeHtml(label)}</option>`;
    const keep = (v, list) => (v === "" || list.includes(v) ? v : "");
    const fac = keep(prev.facility, facilities), mon = keep(prev.month, months);
    const sta = prev.status === "issues" ? "issues" : keep(prev.status, statuses);
    const countBy = (k) => all.filter(r => r.kind === k).length;

    qs("cmpSummary").innerHTML = lines.map((l, i) => `<div style="${i ? "margin-top:6px;" : ""}">${l}</div>`).join("")
      + `<div class="extraction-grid" style="grid-template-columns:1fr 1fr 1.2fr auto;align-items:end;margin-top:14px;gap:10px;">
           <div class="field" style="margin:0;"><label>facility_id</label>
             <select id="cmpFilterFacility">${opt("", `All facilities (${facilities.length})`, fac)}${facilities.map(f => opt(f, f, fac)).join("")}</select></div>
           <div class="field" style="margin:0;"><label>Month</label>
             <select id="cmpFilterMonth">${opt("", `All months (${months.length})`, mon)}${months.map(m => opt(m, m, mon)).join("")}</select></div>
           <div class="field" style="margin:0;"><label>Status</label>
             <select id="cmpFilterStatus">${opt("", "All statuses", sta)}${opt("issues", `All issues (${all.filter(r => r.kind !== "ok").length})`, sta)}
               ${statuses.map(k => opt(k, `${COMPARE_STATUS[k].label} (${countBy(k)})`, sta)).join("")}</select></div>
           <button class="btn" id="btnCmpClearFilters" style="height:38px;">Clear</button>
         </div>
         <div id="cmpFilterCount" class="field-help" style="margin-top:8px;"></div>`;
    ["cmpFilterFacility", "cmpFilterMonth", "cmpFilterStatus"].forEach(id => qs(id).addEventListener("change", () => renderCompareTable(result)));
    qs("btnCmpClearFilters").addEventListener("click", () => {
      ["cmpFilterFacility", "cmpFilterMonth", "cmpFilterStatus"].forEach(id => { qs(id).value = ""; });
      renderCompareTable(result);
    });
    renderCompareTable(result);
  }

  // Rows after the facility / month / status filters — shared by the table and the Excel download.
  function filteredCompareRows(result) {
    const fac = qs("cmpFilterFacility")?.value || "";
    const mon = qs("cmpFilterMonth")?.value || "";
    const sta = qs("cmpFilterStatus")?.value || "";
    return compareRows(result).filter(r => (!fac || r.facilityId === fac) && (!mon || r.month === mon)
      && (!sta || (sta === "issues" ? r.kind !== "ok" : r.kind === sta)));
  }

  function compareFilterDescription() {
    const parts = [
      qs("cmpFilterFacility")?.value && `facility ${qs("cmpFilterFacility").value}`,
      qs("cmpFilterMonth")?.value && `month ${qs("cmpFilterMonth").value}`,
      qs("cmpFilterStatus")?.value && `status ${qs("cmpFilterStatus").selectedOptions[0].textContent.replace(/\s*\(\d+\)$/, "")}`,
    ].filter(Boolean);
    return parts.join(", ");
  }

  function renderCompareTable(result) {
    const rows = filteredCompareRows(result);
    const total = compareRows(result).length;
    const desc = compareFilterDescription();
    if (qs("cmpFilterCount")) qs("cmpFilterCount").textContent = desc
      ? `Showing ${rows.length} of ${total} row(s) — filtered by ${desc}. Download Excel exports these rows.`
      : `Showing all ${total} row(s).`;
    const table = qs("cmpResultsTable");
    if (!rows.length) { table.innerHTML = `<tbody><tr><td class="field-help">No rows match these filters.</td></tr></tbody>`; return; }
    const pctCell = (v) => v === "" ? "" : `${v}%`;
    const combinedCell = `<span style="color:var(--muted);">combined</span>`;
    table.innerHTML = `<thead><tr><th>facility_id</th><th>Station</th><th>Month</th><th>Status</th><th>Monthly report (kWh)</th><th>Hourly sum (kWh)</th><th>Diff (kWh)</th><th>Diff (%)</th><th>Hourly coverage</th><th>Missing days (no hourly data)</th><th>Monthly fetched</th><th>Details</th></tr></thead><tbody>` +
      rows.map(r => {
        const st = COMPARE_STATUS[r.kind] || { label: r.kind, color: "var(--muted)" };
        return `
        <tr>
          <td>${escapeHtml(r.facilityId)}</td>
          <td>${r.stationName ? escapeHtml(r.stationName) : combinedCell}</td>
          <td style="white-space:nowrap;">${escapeHtml(r.month)}</td>
          <td style="white-space:nowrap;color:${st.color};font-weight:600;">${st.label}</td>
          <td>${r.monthlyKwh}</td>
          <td>${r.hourlyKwh}</td>
          <td>${r.diffKwh}</td>
          <td>${pctCell(r.diffPct)}</td>
          <td style="white-space:nowrap;">${escapeHtml(r.hourlyCoverage || "")}</td>
          <td>${r.missingRanges ? escapeHtml(r.missingRanges).replace(/, /g, "<br>") : `<span style="color:var(--muted);">none</span>`}</td>
          <td style="white-space:nowrap;${r.monthlyStale ? "color:var(--warn);font-weight:600;" : ""}">${escapeHtml(r.monthlyAsOf || "")}${r.monthlyStale ? " ⚠" : ""}</td>
          <td style="color:var(--muted);">${escapeHtml(r.issue)}</td>
        </tr>`; }).join("") + `</tbody>`;
  }

  // "GEN3228: 2026-01 → 2026-06; GEN1111: 2026-03" — compact list of months per facility.
  function summarizeMonths(list) {
    const byFac = new Map();
    list.forEach(x => { const k = groupLabel(x); if (!byFac.has(k)) byFac.set(k, []); byFac.get(k).push(x.month); });
    const next = (m) => { const [y, mo] = m.split("-").map(Number); return mo === 12 ? `${y + 1}-01` : `${y}-${String(mo + 1).padStart(2, "0")}`; };
    return [...byFac.entries()].map(([fac, months]) => {
      months.sort();
      const ranges = [];
      let a = months[0], b = months[0];
      for (const m of months.slice(1)) { if (m === next(b)) b = m; else { ranges.push(a === b ? a : `${a} → ${b}`); a = b = m; } }
      ranges.push(a === b ? a : `${a} → ${b}`);
      return `${fac}: ${ranges.join(", ")}`;
    }).join("; ");
  }

  function handleDownloadCompare() {
    if (!lastCompareResult) return;
    const { conn, result } = lastCompareResult;
    Audit.log("Download comparison", conn.companyName, `${TemplateExport.describeGrouping(lastCompareResult.grouping)} · ${filteredCompareRows(result).length} row(s)${compareFilterDescription() ? ` · filtered by ${compareFilterDescription()}` : ""}`);
    const headers = ["facility_id", "Station", "Month", "Status", "Monthly report (kWh)", "Hourly sum (kWh)", "Diff (kWh)", "Diff (%)", "Hourly coverage", "Missing days (no hourly data)", "Monthly fetched", "Details"];
    const rows = filteredCompareRows(result).map(r => [r.facilityId, r.stationName || "(combined)", r.month, (COMPARE_STATUS[r.kind] || {}).label || r.kind,
      r.monthlyKwh, r.hourlyKwh, r.diffKwh, r.diffPct,
      r.hourlyCoverage || "", r.missingRanges || "none", r.monthlyAsOf || "", r.issue]);
    if (!rows.length) rows.push(["No rows match these filters", ...headers.slice(1).map(() => "")]);
    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    ws["!cols"] = [10, 22, 9, 18, 14, 14, 11, 9, 13, 40, 13, 60].map(wch => ({ wch }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "HourlyVsMonthly");
    const period = (lastCompareResult.periodLabel || "").replace(/\s*→\s*/g, "_to_").replace(/\s+/g, "_");
    XLSX.writeFile(wb, `${conn.companyName}_HourlyVsMonthly${period ? `_${period}` : ""}.xlsx`);
  }

  /* ---------------------------- export (Template 1/2 from stored Readings) ---------------------------- */

  function populateExportCompanySelect() {
    const sel = qs("expCompany");
    const keep = sel.value;
    sel.innerHTML = `<option value="">Select a company…</option>`;
    connections.forEach(c => {
      const opt = document.createElement("option");
      opt.value = c.id; opt.textContent = `${c.companyName} (${BRANDS[c.brand].label})`;
      sel.appendChild(opt);
    });
    if (keep && connections.some(c => c.id === keep)) sel.value = keep;
    handleExportCompanyChange();
  }

  function handleExportCompanyChange() {
    renderGroupingPicker("expGroupingWrap", connections.find(c => c.id === qs("expCompany").value));
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
    if (endDate < startDate) { alert(`End date (${endDate}) is before start date (${startDate}) — please fix the range.`); return; }

    qs("btnRunExport").disabled = true;
    qs("expStatus").textContent = `Reading ${resolution} rows from your Google Sheet…`;
    qs("expResultsCard").hidden = true;

    try {
      const rawReadings = dedupeReadings(await SheetsClient.listReadings(conn.companyName,
        { fromDate: `${startDate.slice(0, 7)}-01`, toDate: endDate }));
      const brandKey = conn.brand;
      const readings = normalizeReadingsForTemplates(rawReadings, brandKey, conn.templateSettings?.utcOffset);
      const grouping = getGrouping(conn.id);
      const groupingDesc = TemplateExport.describeGrouping(grouping);

      if (resolution === "Hourly") {
        // Facility mapping (station→facility_id, meter_id, eac_registry_id) still
        // comes from what's configured on the company — only the template
        // number (1 vs 2) is chosen fresh here, independent of that saved default.
        const baseSettings = conn.templateSettings ? JSON.parse(JSON.stringify(conn.templateSettings)) : TemplateExport.defaultSettings("1");
        baseSettings.template = qs("expTemplate").value;
        const utcOffset = baseSettings.utcOffset ?? 8;
        const filtered = readings.filter(r => {
          if (String(r.resolution).trim() !== "Hourly") return false;
          const wc = TemplateExport.wallClockFromRow(r, brandKey, utcOffset);
          if (!wc) return false;
          const localDate = `${wc.y}-${String(wc.mo).padStart(2, "0")}-${String(wc.d).padStart(2, "0")}`;
          return localDate >= startDate && localDate <= endDate;
        });
        const groups = TemplateExport.build(filtered, brandKey, baseSettings, grouping, stationNameMap(conn));
        renderExportResults(conn, { resolution, template: baseSettings.template }, groups, filtered.length, "Hourly");
        Audit.log("Export built", conn.companyName, `Hourly · Template ${baseSettings.template} · ${startDate} → ${endDate} · ${groupingDesc} · ${filtered.length} row(s) · ${groups.length} file(s)`);
        qs("expStatus").textContent = `Done — ${readings.length} row(s) read, ${filtered.length} in range.` + (filtered.length ? "" : exportDiagnostic(readings, resolution, brandKey, conn.templateSettings?.utcOffset ?? 8, startDate, endDate));
      } else {
        const settingsForFacility = conn.templateSettings || {};
        const startMonth = `${startDate.slice(0, 7)}-01`;
        const filtered = readings.filter(r => {
          if (String(r.resolution).trim() !== "Monthly") return false;
          const wc = TemplateExport.wallClockFromRow(r, brandKey, settingsForFacility.utcOffset ?? 8);
          if (!wc) return false;
          const rowMonth = `${wc.y}-${String(wc.mo).padStart(2, "0")}-01`;
          return rowMonth >= startMonth && rowMonth <= endDate;
        });
        const groups = buildMonthlyExportGroups(filtered, brandKey, settingsForFacility, grouping, stationNameMap(conn));
        renderExportResults(conn, { resolution, template: null }, groups, filtered.length, "Monthly");
        Audit.log("Export built", conn.companyName, `Monthly · ${startDate} → ${endDate} · ${groupingDesc} · ${filtered.length} row(s) · ${groups.length} file(s)`);
        qs("expStatus").textContent = `Done — ${readings.length} row(s) read, ${filtered.length} in range.` + (filtered.length ? "" : exportDiagnostic(readings, resolution, brandKey, conn.templateSettings?.utcOffset ?? 8, startDate, endDate));
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
  // Combined AGG IDs: one file holding all their stations' rows. Split AGG IDs
  // (see the AGG ID grouping picker): one file per station.
  function buildMonthlyExportGroups(readings, brandKey, settings, grouping, stationNames = {}) {
    const headers = ["GEN ID", "stationName", "collectTime", "PVYield"];
    const byGroup = new Map(); // group key -> { g, entries }
    for (const r of readings) {
      const wc = TemplateExport.wallClockFromRow(r, brandKey, settings.utcOffset ?? 8);
      if (!wc) continue;
      const g = TemplateExport.groupFor(r.stationId, settings, grouping);
      const collectTime = `${wc.y}-${String(wc.mo).padStart(2, "0")}-01 00:00:00`;
      const kwhR = Math.round((r.kwh || 0) * 100) / 100;
      if (!byGroup.has(g.key)) byGroup.set(g.key, { g, entries: [] });
      byGroup.get(g.key).entries.push({ epoch: Date.UTC(wc.y, wc.mo - 1, 1), row: [g.facilityId, r.stationName, collectTime, kwhR] });
    }
    const groups = [];
    for (const { g, entries } of byGroup.values()) {
      entries.sort((a, b) => a.epoch - b.epoch);
      groups.push({ key: g.key, facilityId: g.facilityId, stationId: g.stationId,
        stationName: g.stationId ? (stationNames[g.stationId] || g.stationId) : "", headers, rows: entries.map(e => e.row) });
    }
    groups.sort((a, b) => a.facilityId.localeCompare(b.facilityId) || a.stationName.localeCompare(b.stationName));
    return groups;
  }

  function renderExportResults(conn, exportKind, groups, rowCount, resolutionLabel) {
    qs("expResultsCard").hidden = false;
    qs("expSummary").innerHTML = groups.length
      ? `${rowCount} ${resolutionLabel.toLowerCase()} row(s) in range, grouped into <strong>${groups.length}</strong> file(s) <span style="color:var(--muted);">(${escapeHtml(TemplateExport.describeGrouping(getGrouping(conn.id)))})</span>.`
      : `No ${resolutionLabel.toLowerCase()} rows found in that date range for this company.`;

    const list = qs("expFacilityList");
    list.innerHTML = "";
    groups.forEach(g => {
      const row = document.createElement("div");
      row.className = "export-facility-row";
      const formatLabel = exportKind.resolution === "Hourly" ? `Template ${exportKind.template}` : "Monthly template";
      row.innerHTML = `
        <div>
          <div style="font-weight:600;">${escapeHtml(g.facilityId)}${g.stationName ? ` <span style="font-weight:400;color:var(--muted);">· ${escapeHtml(g.stationName)}</span>` : ""}</div>
          <div class="meta">${g.rows.length} row(s) · ${escapeHtml(formatLabel)}${g.stationName ? " · this station only" : ""}</div>
        </div>
        <button class="btn btn-primary">Download</button>`;
      row.querySelector("button").addEventListener("click", () => downloadExportFacility(conn, exportKind, g));
      list.appendChild(row);
    });
  }

  function downloadExportFacility(conn, exportKind, group) {
    Audit.log("Export downloaded", conn.companyName, `${groupLabel(group)} · ${exportKind.resolution === "Hourly" ? `Template ${exportKind.template}` : "Monthly"} · ${group.rows.length} row(s)`);
    const safeFacility = (group.stationName ? `${group.facilityId}_${group.stationName}` : group.facilityId).replace(/[^a-z0-9_-]+/gi, "_");
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
    const conn = id ? connections.find(c => c.id === id) : null;
    renderGroupingPicker("extGroupingWrap", conn);
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
    // The cursor is the last day/month already fetched, so suggest the one after it.
    const suggested = ExtractionEngine.nextStartAfter(resolution, cursorDate);
    if (suggested && !DatePicker.getISO(qs("extStart"))) {
      DatePicker.setFromISO(qs("extStart"), suggested);
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
    if (endDate < startDate) { alert("End date is before start date — please fix the range."); return; }
    if (!stationIds.length) { alert("Tick at least one station."); return; }
    const today = localISODate(new Date()); // local date — toISOString() gave yesterday's date before 8am in UTC+8
    if (endDate > today) {
      alert("End date can't be in the future — data for days that haven't happened yet doesn't exist. Pick today or an earlier date.");
      return;
    }

    // ---- extraction lock: don't re-fetch what's already in the Readings tab ----
    const btn = qs("btnQueueExtraction");
    if (btn) btn.disabled = true;
    let plan;
    try {
      plan = await planExtraction(conn, resolution, stationIds, startDate, endDate);
    } catch (e) {
      if (btn) btn.disabled = false;
      alert(`Could not check which dates are already extracted (${e.message}). Try again in a moment.`);
      return;
    }
    if (btn) btn.disabled = false;

    const rangeText = `${resolution} · ${startDate} → ${endDate} · ${stationIds.length} station(s)`;
    if (!plan.done.length) { queueExtractionJob(conn, resolution, stationIds, startDate, endDate); return; }

    const choice = await askExtractionLock(conn, resolution, plan);
    if (choice === "missing") {
      Audit.log("Extraction — already-extracted dates skipped", conn.companyName,
        `${rangeText} · skipped ${plan.done.length}: ${plan.doneText} · fetching ${plan.missing.length}: ${plan.missingText}`);
      const first = plan.missing[0], last = plan.missing[plan.missing.length - 1];
      // Start/end narrowed to the missing span; anything already extracted inside it is skipped.
      queueExtractionJob(conn, resolution, stationIds,
        resolution === "Monthly" ? maxISO(first, startDate) : first,
        resolution === "Monthly" ? minISO(monthEnd(last), endDate) : last,
        { skipPoints: new Set(plan.done) });
    } else if (choice === "all") {
      Audit.log("Re-extraction (Admin override)", conn.companyName, `${rangeText} · ${plan.done.length} already-extracted ${plan.unit}(s) fetched again: ${plan.doneText}`);
      queueExtractionJob(conn, resolution, stationIds, startDate, endDate);
    } else {
      Audit.log("Extraction blocked — already extracted", conn.companyName, `${rangeText} · ${plan.doneText}`);
    }
  }

  /* ---------------------------- extraction lock ---------------------------- */
  // A day (Hourly/Daily) or month (Monthly) counts as ALREADY EXTRACTED when
  // every selected station has rows for it in the Readings tab AND those rows
  // were fetched after that day/month had ended. So today, the current month,
  // and anything fetched while it was still in progress are never locked —
  // they may still be missing data and can always be fetched again.
  function localISODate(d) {
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
  }
  function monthEnd(firstOfMonth) {
    const [y, m] = firstOfMonth.split("-").map(Number);
    return `${y}-${String(m).padStart(2, "0")}-${String(new Date(Date.UTC(y, m, 0)).getUTCDate()).padStart(2, "0")}`;
  }
  const maxISO = (a, b) => (a > b ? a : b);
  const minISO = (a, b) => (a < b ? a : b);

  async function planExtraction(conn, resolution, stationIds, startDate, endDate) {
    const points = ExtractionEngine.buildCursorPoints(conn.brand, resolution, startDate, endDate);
    const readings = dedupeReadings(await SheetsClient.listReadings(conn.companyName,
      { fromDate: `${startDate.slice(0, 7)}-01`, toDate: resolution === "Monthly" ? monthEnd(`${endDate.slice(0, 7)}-01`) : endDate }));
    const monthly = resolution === "Monthly";

    // point -> Map(stationId -> newest RunAt local date for that point)
    const seen = new Map();
    for (const r of readings) {
      if (String(r.resolution).trim() !== resolution) continue;
      const m = String(r.timestamp).trim().match(/^(\d{4})-(\d{2})-(\d{2})/);
      if (!m) continue;
      const point = monthly ? `${m[1]}-${m[2]}-01` : `${m[1]}-${m[2]}-${m[3]}`;
      if (!seen.has(point)) seen.set(point, new Map());
      const byStation = seen.get(point);
      const fetchedOn = r.runAt ? localISODate(new Date(r.runAt)) : "";
      if (fetchedOn > (byStation.get(r.stationId) || "")) byStation.set(r.stationId, fetchedOn);
    }

    const done = [], missing = [];
    for (const point of points) {
      const lastDay = monthly ? monthEnd(point) : point;
      const byStation = seen.get(point);
      const complete = !!byStation && stationIds.every(st => (byStation.get(st) || "") > lastDay);
      (complete ? done : missing).push(point);
    }
    const unit = monthly ? "month" : "day";
    const fmt = (list) => monthly ? summarizeMonthPoints(list) : summarizeDayPoints(list);
    return { points, done, missing, unit, doneText: fmt(done), missingText: fmt(missing) };
  }

  function summarizeDayPoints(days) {
    if (!days.length) return "none";
    const next = (s) => { const [y, m, d] = s.split("-").map(Number); return localISODate(new Date(y, m - 1, d + 1)); };
    const out = []; let a = days[0], b = days[0];
    for (const d of days.slice(1)) { if (d === next(b)) b = d; else { out.push(a === b ? a : `${a} → ${b}`); a = b = d; } }
    out.push(a === b ? a : `${a} → ${b}`);
    return out.join(", ");
  }
  function summarizeMonthPoints(months) {
    if (!months.length) return "none";
    return summarizeMonths(months.map(p => ({ facilityId: "", month: p.slice(0, 7) }))).replace(/^: /, "");
  }

  // Resolves "missing" | "all" | "cancel".
  function askExtractionLock(conn, resolution, plan) {
    const isAdmin = currentRole === "Admin";
    const allDone = !plan.missing.length;
    const unitPl = plan.unit + "(s)";
    qs("extLockTitle").textContent = allDone ? "🔒 Already extracted" : "Some dates already extracted";
    qs("extLockBody").innerHTML = `
      <p><strong>${escapeHtml(conn.companyName)}</strong> · ${escapeHtml(resolution)}</p>
      <p><strong>${plan.done.length} of ${plan.points.length} ${unitPl}</strong> in this range are already in the Readings tab for every selected station:</p>
      <div class="callout" style="margin:6px 0 12px;">${escapeHtml(plan.doneText)}</div>
      ${allDone
        ? `<p>Nothing new to fetch.${isAdmin ? " As an Admin you can re-extract it anyway (e.g. if the vendor corrected its figures) — this is recorded in the audit log." : " Ask an Admin if this data needs to be re-extracted."}</p>`
        : `<p><strong>${plan.missing.length} ${unitPl}</strong> still missing — only these will be fetched:</p>
           <div class="callout" style="margin:6px 0 12px;">${escapeHtml(plan.missingText)}</div>`}
      <p class="field-help" style="margin-top:4px;">Today, the current month, and anything fetched before it had ended are never locked.</p>`;
    qs("btnExtLockMissing").hidden = allDone;
    qs("btnExtLockAll").hidden = !isAdmin;
    qs("btnExtLockAll").textContent = allDone ? "Re-extract anyway (Admin)" : "Re-extract everything (Admin)";
    qs("btnExtLockCancel").textContent = allDone && !isAdmin ? "OK" : "Cancel";
    qs("extLockOverlay").classList.add("active");
    return new Promise(resolve => {
      const done = (v) => {
        qs("extLockOverlay").classList.remove("active");
        ["btnExtLockMissing", "btnExtLockAll", "btnExtLockCancel", "btnExtLockClose"].forEach(id => { qs(id).onclick = null; });
        resolve(v);
      };
      qs("btnExtLockMissing").onclick = () => done("missing");
      qs("btnExtLockAll").onclick = () => done("all");
      qs("btnExtLockCancel").onclick = () => done("cancel");
      qs("btnExtLockClose").onclick = () => done("cancel");
    });
  }

  // Shared by both "Queue extraction" (fresh, from the form) and "Resume"
  // (reconstructed from a persisted pendingJob, no form re-entry needed).
  function queueExtractionJob(conn, resolution, stationIds, startDate, endDate, opts = {}) {
    const brand = BRANDS[conn.brand];

    // Created immediately so it shows up in the Jobs list right away, even
    // if it has to wait its turn behind another job that's already running.
    const jobRow = document.createElement("div");
    jobRow.className = "job-row";
    jobRow.innerHTML = `<div class="job-title">${escapeHtml(conn.companyName)} · ${resolution} · ${startDate} → ${endDate}${opts.skipPoints?.size ? ` <span style="color:var(--muted);font-weight:400;">· ${opts.skipPoints.size} already-extracted ${resolution === "Monthly" ? "month" : "day"}(s) skipped</span>` : ""}</div>
      <div class="job-bar"><div class="job-bar-fill"></div></div>
      <div class="job-status">${queueRunning || extractionQueue.length ? "Waiting for other extraction(s) to finish…" : "Starting…"}</div>`;
    qs("jobList").prepend(jobRow);

    const job = ExtractionEngine.createJob({
      connection: conn, brand: conn.brand, resolution, startDate, endDate, stationIds,
      skipPoints: opts.skipPoints,
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

      // Snapshot so the cursor can be rolled back if the Sheet sync fails —
      // otherwise Resume would skip days whose rows never got saved.
      const cursorBefore = JSON.parse(JSON.stringify(conn.cursor || {}));
      let syncFailed = false;

      await ExtractionEngine.run(job.id, ctx);

      // "error" included: rows fetched before the error are valid and the
      // cursor has already moved past them, so they must be saved too.
      if (["done", "paused", "stopped", "error"].includes(job.status) && job.rowsCollected.length) {
        try {
          const utcOffset = conn.templateSettings?.utcOffset ?? 8;
          await SheetsClient.appendReadings(job.rowsCollected.map(r => ({
            timestamp: formatReadableTimestamp(r, conn.brand, utcOffset), company: conn.companyName, brand: brand.label,
            stationId: r.stationId, stationName: (conn.stations.find(s => s.id === r.stationId) || {}).name || r.stationId,
            resolution, kwh: r.kwh,
          })));
          jobRow.querySelector(".job-status").textContent += " · Synced to Google Sheet";
        } catch (e) {
          syncFailed = true;
          jobRow.querySelector(".job-status").textContent += ` · Sheet sync failed (${e.message}) — progress not saved, this range will be fetched again on the next run`;
        }
      }
      conn.cursor = syncFailed ? cursorBefore : job.connection.cursor;
      job.connection.cursor = conn.cursor;
      if (job.status === "done" && !syncFailed) {
        delete conn.pendingJob[resolution]; // fully caught up — nothing left to resume
      }
      await SheetsClient.saveConnection(conn).catch(() => {});
      Audit.log("Data extraction", conn.companyName,
        `${resolution} · ${startDate} → ${endDate} · ${stationIds.length} station(s) · Status: ${job.status} · ${job.rowsCollected.length} row(s)`
        + (syncFailed ? " · Sheet sync FAILED (cursor not advanced)" : ""));
      job._downloadRows = job.rowsCollected;
      jobRow._job = job;
    });
  }

  // Companies with a saved pendingJob (paused/cancelled/tab-closed in a
  // previous session) get a "Resume" card here — same Jobs list, no form
  // re-entry needed. Skips anything already shown live this session.
  const liveJobKeys = new Set();

  // Tracks every currently-rendered "Resume" card's action, so "Resume all"
  // can trigger them in one click instead of clicking each individually.
  const pendingResumeActions = [];

  function renderPersistedPendingJobs() {
    for (const conn of connections) {
      if (!conn.pendingJob) continue;
      for (const [resolution, pending] of Object.entries(conn.pendingJob)) {
        const key = `${conn.id}|${resolution}`;
        if (liveJobKeys.has(key)) continue;
        liveJobKeys.add(key);

        // Resume from the day/month AFTER the last one fetched (resuming AT the
        // cursor re-fetched it — the source of the duplicate boundary days).
        // A cursor from before this job's start (left over from an older run)
        // is ignored so the job's own start date is honoured.
        const cursor = conn.cursor?.[resolution];
        const effectiveStart = cursor && cursor >= pending.startDate
          ? ExtractionEngine.nextStartAfter(resolution, cursor)
          : pending.startDate;
        if (effectiveStart > pending.endDate) {
          delete conn.pendingJob[resolution]; // everything in range was already fetched
          SheetsClient.saveConnection(conn).catch(() => {});
          continue;
        }
        const jobRow = document.createElement("div");
        jobRow.className = "job-row";
        jobRow.innerHTML = `<div class="job-title">${escapeHtml(conn.companyName)} · ${resolution} · ${pending.startDate} → ${pending.endDate}</div>
          <div class="job-bar"><div class="job-bar-fill"></div></div>
          <div class="job-status">Unfinished from a previous session — ${cursor && cursor >= pending.startDate ? `fetched up to ${escapeHtml(cursor)}, ` : ""}resumes from ${escapeHtml(effectiveStart)}.</div>`;
        const resumeThis = () => {
          const idx = pendingResumeActions.indexOf(resumeThis);
          if (idx !== -1) pendingResumeActions.splice(idx, 1);
          qs("btnResumeAll").hidden = pendingResumeActions.length === 0;
          jobRow.remove();
          queueExtractionJob(conn, resolution, pending.stationIds, effectiveStart, pending.endDate);
        };
        const resumeBtn = document.createElement("button");
        resumeBtn.className = "btn btn-primary";
        resumeBtn.textContent = "Resume";
        resumeBtn.addEventListener("click", resumeThis);
        jobRow.appendChild(resumeBtn);
        qs("jobList").appendChild(jobRow);
        pendingResumeActions.push(resumeThis);
      }
    }
    qs("btnResumeAll").hidden = pendingResumeActions.length === 0;
  }

  function handleResumeAll() {
    // Snapshot first — each resumeThis() call removes its own row and enqueues
    // a real job, which shouldn't affect the ones still waiting to be triggered.
    const actions = pendingResumeActions.splice(0, pendingResumeActions.length);
    actions.forEach(fn => fn());
    qs("btnResumeAll").hidden = true;
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

  // Raw readings for this one job (for Template 1/2, use the Export tab, which
  // reads accumulated Readings for any date range rather than just this job's
  // in-memory rows). Follows the AGG ID grouping picker on the Extraction tab
  // for this company:
  //   - "All" sheet: every row with its AGG ID; combined AGG IDs are summed per
  //     timestamp into one row, split ones keep a row per station.
  //   - then one sheet per group (per AGG ID, or per station when split).
  function downloadRowsAsExcel(job) {
    const conn = connections.find(c => c.id === job.connection.id) || job.connection; // latest AGG IDs
    const settings = conn.templateSettings || {};
    const utcOffset = settings.utcOffset ?? 8;
    const grouping = getGrouping(conn.id);
    const names = stationNameMap(conn);

    const groups = new Map(); // group key -> { g, byTs: Map(ts -> { kwh, stations:Set }) }
    for (const r of job.rowsCollected) {
      const g = TemplateExport.groupFor(r.stationId, settings, grouping);
      const ts = formatReadableTimestamp(r, job.brand, utcOffset);
      if (!groups.has(g.key)) groups.set(g.key, { g, byTs: new Map() });
      const byTs = groups.get(g.key).byTs;
      if (!byTs.has(ts)) byTs.set(ts, { kwh: 0, stations: new Set() });
      const cell = byTs.get(ts);
      cell.kwh += Number(r.kwh) || 0;
      cell.stations.add(names[r.stationId] || r.stationId);
    }

    const sorted = [...groups.values()].sort((a, b) => a.g.facilityId.localeCompare(b.g.facilityId)
      || (names[a.g.stationId] || "").localeCompare(names[b.g.stationId] || ""));
    const toRows = ({ g, byTs }) => [...byTs.entries()].sort((a, b) => String(a[0]).localeCompare(String(b[0]))).map(([ts, c]) => ({
      Timestamp: ts,
      "AGG ID": g.facilityId,
      Station: g.stationId ? (names[g.stationId] || g.stationId)
        : c.stations.size > 1 ? `${c.stations.size} stations combined` : [...c.stations][0],
      Resolution: job.resolution,
      kWh: Math.round(c.kwh * 1e3) / 1e3,
    }));

    const wb = XLSX.utils.book_new();
    const used = new Set();
    const sheetName = (base) => { // Excel: max 31 chars, no []:*?/\, unique
      const clean = String(base).replace(/[\[\]:*?\/\\]/g, "_").slice(0, 31) || "Sheet";
      let name = clean, n = 2;
      while (used.has(name.toLowerCase())) { const suf = ` (${n++})`; name = clean.slice(0, 31 - suf.length) + suf; }
      used.add(name.toLowerCase());
      return name;
    };
    const allRows = sorted.flatMap(toRows);
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(allRows), sheetName("All"));
    for (const grp of sorted) {
      const label = grp.g.stationId ? `${grp.g.facilityId} ${names[grp.g.stationId] || grp.g.stationId}` : grp.g.facilityId;
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(toRows(grp)), sheetName(label));
    }
    Audit.log("Download raw extraction", conn.companyName, `${job.resolution} · ${job.startDate} → ${job.endDate} · ${TemplateExport.describeGrouping(grouping)} · ${job.rowsCollected.length} row(s) → ${allRows.length} row(s) in ${sorted.length} group(s)`);
    XLSX.writeFile(wb, `${conn.companyName}_${job.resolution}_${job.startDate}_${job.endDate}_raw.xlsx`);
  }

  /* ---------------------------- settings password lock ---------------------------- */
  // NOTE: this is a UI lock only. The check runs in the browser, so anyone who
  // opens DevTools can get past it — it stops casual/accidental changes, not a
  // determined user. Only the SHA-256 of the password is stored here.
  const SETTINGS_PASSWORD_SHA256 = "03ac674216f3e15c761ee1a5e255f067953623c8b388b4459e13f978d7c846f4";

  function openSettingsLock() {
    qs("settingsPasswordInput").value = "";
    qs("settingsLockError").textContent = "";
    qs("settingsLockOverlay").classList.add("active");
    setTimeout(() => qs("settingsPasswordInput").focus(), 30);
  }
  function closeSettingsLock() { qs("settingsLockOverlay").classList.remove("active"); }

  function handleSettingsUnlock() {
    const pw = qs("settingsPasswordInput").value;
    if (CryptoJS.SHA256(pw).toString() === SETTINGS_PASSWORD_SHA256) {
      settingsUnlocked = true;
      closeSettingsLock();
      Audit.log("Settings unlocked", "", "");
      showView("settings");
    } else {
      qs("settingsLockError").textContent = "Incorrect password.";
      qs("settingsPasswordInput").select();
      Audit.log("Settings unlock FAILED", "", "Wrong password entered");
    }
  }

  /* ---------------------------- audit log ---------------------------- */
  // Every entry goes to the Google Sheet's "Audit" tab via the proxy. If the
  // proxy can't take it (offline, or its appendAudit action isn't deployed
  // yet), the entry is kept in this browser and re-sent with the next one —
  // so nothing is lost, but it's only visible here until it syncs.
  const Audit = (() => {
    const PENDING_KEY = "slc.auditPending";
    const loadPending = () => { try { return JSON.parse(localStorage.getItem(PENDING_KEY) || "[]"); } catch { return []; } };
    const savePending = (arr) => { try { localStorage.setItem(PENDING_KEY, JSON.stringify(arr.slice(-2000))); } catch {} };
    let flushing = false;

    async function flush() {
      if (flushing) return;
      const pending = loadPending();
      if (!pending.length) return;
      flushing = true;
      try {
        await SheetsClient.appendAudit(pending);
        const now = loadPending();
        savePending(now.slice(pending.length)); // keep anything logged while we were sending
      } catch { flushing = false; return; /* stays pending; retried on the next log() */ }
      flushing = false;
      if (loadPending().length) return flush(); // entries logged mid-send go out right away
    }

    function log(action, company, details) {
      const entry = {
        timestamp: new Date().toISOString(),
        user: currentEmail || "(not signed in)",
        role: currentRole || "",
        action, company: company || "", details: details || "",
      };
      savePending([...loadPending(), entry]);
      flush();
    }
    return { log, flush, loadPending };
  })();

  let auditEntries = [];
  let auditSource = "";

  async function loadAuditLog() {
    qs("auditStatus").textContent = "Loading…";
    await Audit.flush();
    const pending = Audit.loadPending().map(e => ({ ...e, _pending: true }));
    try {
      const res = await SheetsClient.listAudit();
      const rows = Array.isArray(res) ? res : (res.rows || res.entries || []);
      auditEntries = [...rows, ...pending];
      auditSource = pending.length ? `${pending.length} entr${pending.length === 1 ? "y" : "ies"} not yet synced to the Sheet (shown with ⏳).` : "";
    } catch (e) {
      auditEntries = pending;
      auditSource = `Could not read the Audit tab from the Sheet (${e.message}). Showing only entries stored in this browser.`;
    }
    auditEntries.sort((a, b) => String(b.timestamp).localeCompare(String(a.timestamp)));
    const actions = [...new Set(auditEntries.map(e => e.action))].sort();
    const sel = qs("auditFilterAction"), keep = sel.value;
    sel.innerHTML = `<option value="">All actions</option>` + actions.map(a => `<option>${escapeHtml(a)}</option>`).join("");
    sel.value = actions.includes(keep) ? keep : "";
    const users = [...new Set(auditEntries.map(e => e.user))].sort();
    const usel = qs("auditFilterUser"), ukeep = usel.value;
    usel.innerHTML = `<option value="">All users</option>` + users.map(u => `<option>${escapeHtml(u)}</option>`).join("");
    usel.value = users.includes(ukeep) ? ukeep : "";
    renderAuditTable();
  }

  function filteredAuditEntries() {
    const u = qs("auditFilterUser").value, a = qs("auditFilterAction").value;
    const t = qs("auditFilterText").value.trim().toLowerCase();
    return auditEntries.filter(e => (!u || e.user === u) && (!a || e.action === a)
      && (!t || `${e.company} ${e.details}`.toLowerCase().includes(t)));
  }

  function fmtAuditTime(ts) {
    const d = new Date(ts);
    if (isNaN(d)) return String(ts);
    const p = n => String(n).padStart(2, "0");
    return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
  }

  function renderAuditTable() {
    const list = filteredAuditEntries();
    qs("auditStatus").textContent = `${list.length} of ${auditEntries.length} entr${auditEntries.length === 1 ? "y" : "ies"}. ${auditSource}`;
    const table = qs("auditTable");
    if (!list.length) { table.innerHTML = `<tbody><tr><td class="field-help">No audit entries match.</td></tr></tbody>`; return; }
    table.innerHTML = `<thead><tr><th>Time (your local)</th><th>User</th><th>Action</th><th>Company</th><th>Details</th></tr></thead><tbody>` +
      list.slice(0, 1000).map(e => `<tr>
        <td style="white-space:nowrap;">${e._pending ? "⏳ " : ""}${escapeHtml(fmtAuditTime(e.timestamp))}</td>
        <td>${escapeHtml(e.user)}${e.role ? `<div class="field-help" style="margin-top:0;">${escapeHtml(e.role)}</div>` : ""}</td>
        <td><span class="audit-action" data-kind="${escapeHtml(auditKind(e.action))}">${escapeHtml(e.action)}</span></td>
        <td>${escapeHtml(e.company)}</td>
        <td style="color:var(--muted);">${escapeHtml(e.details)}</td></tr>`).join("") + `</tbody>`;
  }

  function auditKind(action) {
    const a = String(action).toLowerCase();
    if (a.includes("failed") || a.includes("remove")) return "danger";
    if (a.includes("import")) return "import";
    if (a.includes("settings")) return "settings";
    if (a.includes("extraction")) return "extract";
    if (a.includes("export") || a.includes("download")) return "export";
    if (a.includes("compar")) return "compare";
    if (a.includes("company")) return "company";
    return "other";
  }

  function downloadAuditLog() {
    const list = filteredAuditEntries();
    Audit.log("Download audit log", "", `${list.length} entr${list.length === 1 ? "y" : "ies"}`);
    const ws = XLSX.utils.aoa_to_sheet([["Timestamp", "User", "Role", "Action", "Company", "Details"],
      ...list.map(e => [fmtAuditTime(e.timestamp), e.user, e.role, e.action, e.company, e.details])]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Audit");
    XLSX.writeFile(wb, `SolarLink_AuditLog_${new Date().toISOString().slice(0, 10)}.xlsx`);
  }

  function escapeHtml(s) { return String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c])); }

  // FusionSolar's raw timestamp is an epoch number (e.g. 1782835200000) —
  // fine for calculation, unreadable if written straight into the Sheet.
  // SolarEdge's is already a readable local-time string from their API.
  // This makes both consistent: a plain "YYYY-MM-DD HH:MM:SS" local string,
  // reusing the exact same brand-aware wall-clock logic already verified
  // correct in Template exports.
  // Rows are written to the Sheet with a readable "YYYY-MM-DD HH:MM:SS"
  // local-time timestamp (see formatReadableTimestamp). But
  // TemplateExport.wallClockFromRow expects each brand's RAW API timestamp —
  // for FusionSolar that's an epoch number — so for readable-string rows it
  // returns null and every row gets silently dropped from Export/Compare.
  // This converts readable rows back into a form wallClockFromRow accepts,
  // and verifies the round-trip so no row is ever shifted by the wrong offset.
  // Same rule as the proxy's listReadings: one row per (Resolution,
  // StationId, Timestamp), newest RunAt wins. Duplicates come from backfill
  // chunks resuming on the day the previous chunk ended, and from re-running
  // an extraction — left in, they're double-counted by Compare and Export.
  // Harmless no-op once the updated proxy is deployed (it already de-dupes).
  function dedupeReadings(readings) {
    const newest = new Map();
    for (const r of readings || []) {
      const key = `${r.resolution}\u0001${r.stationId}\u0001${String(r.timestamp).trim()}`;
      const prev = newest.get(key);
      if (!prev || String(r.runAt || "") >= String(prev.runAt || "")) newest.set(key, r);
    }
    return [...newest.values()];
  }

  function normalizeReadingsForTemplates(readings, brandKey, utcOffset) {
    const off = utcOffset ?? 8;
    // A wall-clock result only counts if every field is a real number AND it
    // lands on the same date/hour as the readable string. (A NaN-filled object
    // is truthy, so a plain truthiness check would wrongly accept it.)
    const matches = (wc, p) => !!wc && [wc.y, wc.mo, wc.d, wc.H].every(Number.isFinite)
      && (!p || (wc.y === p.y && wc.mo === p.mo && wc.d === p.d && wc.H === p.H));
    return readings.map(r => {
      const row = { ...r, readableTimestamp: r.readableTimestamp ?? r.timestamp, kwh: r.kwh === "" || r.kwh == null ? r.kwh : Number(r.kwh) };
      const m = String(r.timestamp).trim().match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})(?::(\d{2}))?/);
      if (!m) return matches(safeWallClock(row, brandKey, off)) ? row : row; // old epoch rows etc. — leave as-is
      const [y, mo, d, H, Mi, S] = [m[1], m[2], m[3], m[4], m[5], m[6] || "0"].map(Number);
      const parsed = { y, mo, d, H };
      if (matches(safeWallClock(row, brandKey, off), parsed)) return row; // e.g. SolarEdge: string is native
      const epoch = Date.UTC(y, mo - 1, d, H, Mi, S) - off * 3600000; // local wall clock -> UTC epoch
      for (const candidate of [epoch, String(epoch), Math.floor(epoch / 1000)]) {
        const test = { ...row, timestamp: candidate, collectTime: candidate };
        if (matches(safeWallClock(test, brandKey, off), parsed)) return test;
      }
      return row;
    });
  }

  function safeWallClock(row, brandKey, off) {
    try { return TemplateExport.wallClockFromRow(row, brandKey, off); } catch { return null; }
  }

  // Shown in the status line only when rows exist but none land in range,
  // so the cause is visible on screen instead of a bare "0 in range".
  function exportDiagnostic(readings, resolution, brandKey, off, startDate, endDate) {
    const same = readings.filter(r => String(r.resolution).trim() === resolution);
    if (!same.length) {
      const seen = [...new Set(readings.map(r => JSON.stringify(r.resolution)))].slice(0, 5).join(", ");
      return ` No rows with resolution "${resolution}" — values seen: ${seen || "none"}.`;
    }
    const s = same[0];
    const wc = safeWallClock(s, brandKey, off);
    return ` Range ${startDate} → ${endDate}. Sample row timestamp ${JSON.stringify(s.timestamp)} (${typeof s.timestamp})`
      + ` parsed as ${wc ? JSON.stringify({ y: wc.y, mo: wc.mo, d: wc.d, H: wc.H }) : "null"}.`;
  }

  function formatReadableTimestamp(row, brandKey, utcOffset) {
    const wc = TemplateExport.wallClockFromRow(row, brandKey, utcOffset ?? 8);
    if (!wc) return String(row.timestamp);
    const pad = n => String(n).padStart(2, "0");
    return `${wc.y}-${pad(wc.mo)}-${pad(wc.d)} ${pad(wc.H)}:${pad(wc.Mi)}:${pad(wc.S)}`;
  }


  return { boot };
})();

window.addEventListener("DOMContentLoaded", App.boot);
