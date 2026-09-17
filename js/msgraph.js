/* =====================================================================
   Microsoft sign-in + SharePoint storage
   ---------------------------------------------------------------------
   - Sign-in is delegated (MSAL.js, msal-browser via CDN) — the person
     using this page signs in with their own Microsoft account, and Graph
     calls run with their own SharePoint permissions. No secrets are
     baked into this static site.
   - Connections (brand + credential fields you enter in "Add Company")
     are stored as items in a SharePoint list called "SolarConnections",
     created automatically on first use.
   - Extracted readings are appended to an Excel table ("Readings" inside
     SolarExtractionData.xlsx) via the Graph Excel API — a table scales
     far better than a SharePoint list for dense time-series rows and is
     what "Download Excel" in the UI reads back from.

   SECURITY NOTE
   Brand credentials (API keys, iSolarCloud passwords, etc.) are stored
   as plain fields in the SolarConnections list so the extractor can use
   them. That means anyone with access to that SharePoint list can read
   them — protect the list the same way you'd protect the credentials
   themselves (restrict to the people who currently see them in the
   Postman/`.env` files). This app does not add its own encryption layer.
   ===================================================================== */

const MsGraph = (() => {
  let msalInstance = null;
  let account = null;
  let config = null; // { clientId, tenantId, siteUrl }
  let siteId = null;
  let driveId = null;
  const SCOPES = ["Sites.ReadWrite.All", "Files.ReadWrite.All", "User.Read"];

  function loadConfig() {
    const raw = localStorage.getItem("slc.msConfig");
    config = raw ? JSON.parse(raw) : null;
    return config;
  }
  function saveConfig(cfg) {
    config = cfg;
    localStorage.setItem("slc.msConfig", JSON.stringify(cfg));
  }

  async function init() {
    loadConfig();
    if (!config?.clientId || !config?.tenantId) return false;
    if (typeof msal === "undefined") {
      throw new Error("The MSAL library didn't load (script tag in index.html) — check your internet connection, ad blocker, or browser console for a blocked request, then reload.");
    }
    msalInstance = new msal.PublicClientApplication({
      auth: {
        clientId: config.clientId,
        authority: `https://login.microsoftonline.com/${config.tenantId}`,
        redirectUri: window.location.origin + window.location.pathname,
      },
      cache: { cacheLocation: "localStorage" },
    });
    await msalInstance.handleRedirectPromise();
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length) account = accounts[0];
    return !!account;
  }

  async function signIn() {
    if (!msalInstance) throw new Error("Set Client ID / Tenant ID under Settings first.");
    const res = await msalInstance.loginPopup({ scopes: SCOPES });
    account = res.account;
    return account;
  }
  function signOut() {
    if (msalInstance && account) msalInstance.logoutPopup({ account });
    account = null;
  }
  function isSignedIn() { return !!account; }
  function currentUser() { return account ? { name: account.name, username: account.username } : null; }

  async function getToken() {
    if (!account) throw new Error("Not signed in.");
    try {
      const res = await msalInstance.acquireTokenSilent({ scopes: SCOPES, account });
      return res.accessToken;
    } catch {
      const res = await msalInstance.acquireTokenPopup({ scopes: SCOPES, account });
      return res.accessToken;
    }
  }

  async function gfetch(path, opts = {}) {
    const token = await getToken();
    const resp = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
      ...opts,
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json", ...(opts.headers || {}) },
    });
    if (!resp.ok) {
      const errText = await resp.text();
      throw new Error(`Graph ${opts.method || "GET"} ${path} -> ${resp.status}: ${errText.slice(0, 300)}`);
    }
    if (resp.status === 204) return null;
    const ct = resp.headers.get("content-type") || "";
    return ct.includes("json") ? resp.json() : resp.arrayBuffer();
  }

  // config.siteUrl like "contoso.sharepoint.com:/sites/SolarOps"
  async function ensureSite() {
    if (siteId) return siteId;
    const site = await gfetch(`/sites/${config.siteUrl}`);
    siteId = site.id;
    driveId = await resolveDriveId();
    return siteId;
  }

  // Picks the document library to write the workbook into. config.library
  // is the library's *display name* as shown in SharePoint (e.g. "Process
  // Improvement") — this is NOT always "Documents", so we look it up by
  // name among the site's drives rather than assuming the default drive.
  async function resolveDriveId() {
    const wanted = (config.library || "").trim();
    if (!wanted || wanted.toLowerCase() === "documents") {
      const drive = await gfetch(`/sites/${siteId}/drive`);
      return drive.id;
    }
    const drives = await gfetch(`/sites/${siteId}/drives`);
    const match = (drives.value || []).find(d => d.name.toLowerCase() === wanted.toLowerCase());
    if (!match) {
      throw new Error(`Document library "${wanted}" not found on this site. Check the exact library name under Settings.`);
    }
    return match.id;
  }

  // Folder path within the chosen library, e.g. "API Data". Empty = library root.
  function folderPrefix() {
    const f = (config.folderPath || "").trim().replace(/^\/+|\/+$/g, "");
    return f ? `${f}/` : "";
  }

  /* ---------------- SolarConnections list ---------------- */

  async function ensureConnectionsList() {
    await ensureSite();
    const lists = await gfetch(`/sites/${siteId}/lists?$filter=displayName eq 'SolarConnections'`);
    if (lists.value?.length) return lists.value[0].id;
    const created = await gfetch(`/sites/${siteId}/lists`, {
      method: "POST",
      body: JSON.stringify({
        displayName: "SolarConnections",
        columns: [
          { name: "Brand", text: {} },
          { name: "Region", text: {} },
          { name: "CredentialsJson", text: { allowMultipleLines: true } },
          { name: "StationsJson", text: { allowMultipleLines: true } },
          { name: "DailyAutoExtract", boolean: {} },
          { name: "CursorJson", text: { allowMultipleLines: true } },
        ],
        list: { template: "genericList" },
      }),
    });
    return created.id;
  }

  async function listConnections() {
    const listId = await ensureConnectionsList();
    const res = await gfetch(`/sites/${siteId}/lists/${listId}/items?expand=fields`);
    return (res.value || []).map(itemToConnection);
  }

  function itemToConnection(item) {
    const f = item.fields || {};
    return {
      id: item.id,
      companyName: f.Title,
      brand: f.Brand,
      region: f.Region,
      credentials: safeJson(f.CredentialsJson, {}),
      stations: safeJson(f.StationsJson, []),
      dailyAutoExtract: !!f.DailyAutoExtract,
      cursor: safeJson(f.CursorJson, {}),
    };
  }
  function safeJson(s, fallback) { try { return JSON.parse(s); } catch { return fallback; } }

  async function saveConnection(conn) {
    const listId = await ensureConnectionsList();
    const fields = {
      Title: conn.companyName,
      Brand: conn.brand,
      Region: conn.region || "",
      CredentialsJson: JSON.stringify(conn.credentials || {}),
      StationsJson: JSON.stringify(conn.stations || []),
      DailyAutoExtract: !!conn.dailyAutoExtract,
      CursorJson: JSON.stringify(conn.cursor || {}),
    };
    if (conn.id) {
      await gfetch(`/sites/${siteId}/lists/${listId}/items/${conn.id}/fields`, { method: "PATCH", body: JSON.stringify(fields) });
      return conn.id;
    }
    const created = await gfetch(`/sites/${siteId}/lists/${listId}/items`, { method: "POST", body: JSON.stringify({ fields }) });
    return created.id;
  }

  async function deleteConnection(itemId) {
    const listId = await ensureConnectionsList();
    await gfetch(`/sites/${siteId}/lists/${listId}/items/${itemId}`, { method: "DELETE" });
  }

  /* ---------------- Readings workbook ---------------- */

  const WORKBOOK_NAME = "SolarExtractionData.xlsx";
  const TABLE_NAME = "Readings";

  async function ensureWorkbook() {
    await ensureSite();
    const path = `${folderPrefix()}${WORKBOOK_NAME}`;
    let item;
    try {
      item = await gfetch(`/drives/${driveId}/root:/${path}`);
    } catch {
      // Create a blank workbook (built client-side with SheetJS) then upload it.
      // If a folder path is set and doesn't exist yet, Graph's :/content PUT
      // creates any missing intermediate folders automatically.
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet([["Timestamp", "Company", "Brand", "StationId", "StationName", "Resolution", "kWh", "RunAt"]]);
      XLSX.utils.book_append_sheet(wb, ws, "Readings");
      const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
      const token = await getToken();
      const uploadResp = await fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${path}:/content`, {
        method: "PUT",
        headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/octet-stream" },
        body: wbout,
      });
      item = await uploadResp.json();
      // Turn the header range into a real Excel Table so rows/add works.
      await gfetch(`/drives/${driveId}/items/${item.id}/workbook/worksheets('Readings')/tables/add`, {
        method: "POST",
        body: JSON.stringify({ address: "A1:H1", hasHeaders: true }),
      }).catch(() => {}); // ignore if a default table already exists
      await gfetch(`/drives/${driveId}/items/${item.id}/workbook/worksheets('Readings')/tables/Table1`, {
        method: "PATCH", body: JSON.stringify({ name: TABLE_NAME }),
      }).catch(() => {});
    }
    return item.id;
  }

  async function appendReadings(rows) {
    // rows: [{ timestamp, company, brand, stationId, stationName, resolution, kwh }]
    const itemId = await ensureWorkbook();
    const runAt = new Date().toISOString();
    const values = rows.map(r => [r.timestamp, r.company, r.brand, r.stationId, r.stationName, r.resolution, r.kwh, runAt]);
    // Graph limits payload size — send in batches of 500 rows.
    for (let i = 0; i < values.length; i += 500) {
      const batch = values.slice(i, i + 500);
      await gfetch(`/drives/${driveId}/items/${itemId}/workbook/tables/${TABLE_NAME}/rows/add`, {
        method: "POST", body: JSON.stringify({ values: batch }),
      });
    }
  }

  async function workbookWebUrl() {
    await ensureWorkbook();
    const item = await gfetch(`/drives/${driveId}/root:/${folderPrefix()}${WORKBOOK_NAME}`);
    return item.webUrl;
  }

  return {
    saveConfig, loadConfig, init, signIn, signOut, isSignedIn, currentUser,
    listConnections, saveConnection, deleteConnection, appendReadings, workbookWebUrl,
  };
})();
