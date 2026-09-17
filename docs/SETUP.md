# SolarLink Connect — Setup Guide

This app has three moving parts that each need a one-time setup step:

1. **The web page itself** (`index.html` + `js/`) — hosted on GitHub Pages, this is
   what you and your team open in a browser to add companies and run manual/
   on-demand extractions.
2. **A CORS proxy** — a tiny serverless function that lets the browser talk to
   inverter-brand APIs at all (see "Why a proxy is required" below).
3. **A scheduled GitHub Action** — for true unattended daily/hourly extraction
   that doesn't depend on a browser tab being open.

Both (1) and (3) read/write the same data in your SharePoint site, so they
always agree on what's already been extracted.

---

## Why a proxy is required

Your Postman collections and Python bots call `monitoringapi.solaredge.com` and
`intl.fusionsolar.huawei.com` directly because they run outside a browser.
A browser enforces CORS: it will only let JavaScript read a cross-origin
response if that server explicitly allows it, and none of these six vendors'
monitoring APIs send those headers. So a static page hosted on GitHub Pages
cannot call them directly — every brand call in `index.html` is relayed
through a proxy you deploy, which forwards the request server-to-server (no
CORS rules apply between servers) and returns the result to the browser.

The proxy is intentionally "dumb" — it only forwards requests to an
allow-listed set of hostnames (edit `ALLOWED_HOSTS` in either template if you
need to add one) and never inspects or stores credentials.

### Option A — Azure Function (`api-proxy/azure-function`)
1. Install the [Azure Functions Core Tools](https://learn.microsoft.com/azure/azure-functions/functions-run-local).
2. `cd api-proxy/azure-function && npm install` (no runtime deps, but this
   creates the local folder Azure tooling expects).
3. `func azure functionapp publish <your-function-app-name>` (create the
   Function App in the Azure Portal first — Node.js 18+, Consumption plan is
   plenty).
4. In the Portal, copy the function's **URL** and **function key**, e.g.
   `https://your-func.azurewebsites.net/api/relay?code=...`. In the web app's
   Settings tab, set **Proxy base URL** to everything before `/relay`
   (the app appends `/relay` itself) — if you're using a function key, bake
   it into the URL as `.../api?code=XXXX` so it's appended correctly, or
   switch `authLevel` in `function.json` to `anonymous` for simplicity if
   this proxy isn't publicly discoverable.

### Option B — Cloudflare Worker (`api-proxy/cloudflare-worker`)
1. Install [Wrangler](https://developers.cloudflare.com/workers/wrangler/): `npm install -g wrangler`.
2. `cd api-proxy/cloudflare-worker && wrangler deploy`.
3. Set **Proxy base URL** in the web app's Settings to the printed
   `*.workers.dev` URL (again, everything before `/relay`).

---

## Microsoft 365 / SharePoint setup

### 1. Register an Azure AD app for the web page (delegated sign-in)
In [Azure Portal → App registrations → New registration](https://portal.azure.com):
- Name: `SolarLink Connect (web)`
- Supported account types: single tenant (your org only)
- Redirect URI: **Single-page application (SPA)**, set to the exact URL where
  you'll host `index.html` (e.g. `https://yourorg.github.io/solarlink-connect/`)
- After creation, under **API permissions**, add Microsoft Graph delegated
  permissions: `Sites.ReadWrite.All`, `Files.ReadWrite.All`, `User.Read`, then
  click **Grant admin consent**.
- Copy the **Application (client) ID** and **Directory (tenant) ID** — paste
  these into the web app's Settings tab.

No client secret is needed for this app — it's a public client (SPA), and
each person signs in with their own account and their own SharePoint
permissions.

### 2. Register a second Azure AD app for the GitHub Action (app-only)
The scheduled Action can't interactively sign in, so it needs its own
app-only credential:
- Name: `SolarLink Connect (automation)`
- Under **API permissions**, add Microsoft Graph **Application** permission
  `Sites.Selected` (recommended — see below) or `Sites.ReadWrite.All`, then
  **Grant admin consent**.
- Under **Certificates & secrets**, create a client secret and copy its
  value immediately (you won't see it again).
- If you used `Sites.Selected`, an admin also needs to grant this app
  `write` access to specifically your SharePoint site — see
  [Microsoft's Sites.Selected guide](https://learn.microsoft.com/en-us/sharepoint/dev/solution-guidance/security-apponly-azuread#restricted-app-only-permission-model).

Add these as **GitHub repo secrets** (Settings → Secrets and variables →
Actions):
| Secret | Value |
|---|---|
| `MS_TENANT_ID` | Directory (tenant) ID |
| `MS_CLIENT_ID` | Application (client) ID of the automation app |
| `MS_CLIENT_SECRET` | The client secret you copied |
| `SP_SITE_HOSTNAME` | e.g. `contoso.sharepoint.com` |
| `SP_SITE_PATH` | e.g. `/sites/SolarOps` |
| `SP_LIBRARY_NAME` | Document library display name, e.g. `Process Improvement`. Leave unset (or `Documents`) to use the site's default library. |
| `SP_FOLDER_PATH` | Folder within that library, e.g. `API Data`. Leave unset to use the library root. |

These four values need to match what you set in the web app's Settings tab
(SharePoint site / Document library / Folder path), so the browser and the
scheduled Action write to the exact same workbook.

### 3. SharePoint site
No manual prep needed — both the web app and the GitHub Action create what
they need the first time they run:
- A list called **SolarConnections** (one item per company/brand connection —
  this is where API credentials, cached station lists, the daily-auto-extract
  flag, and resume cursors live).
- A workbook called **SolarExtractionData.xlsx** with a table named
  **Readings** (one row per station/timestamp/kWh reading).

**Security note:** credentials in SolarConnections are stored as plain text
fields (list columns aren't encrypted at rest any differently than any other
SharePoint content) so the extractor can use them. Restrict that list to the
same people who currently have access to your `.env` files / Postman
environments.

---

## Hosting the web page on GitHub Pages

1. Push this whole folder to a GitHub repo.
2. Repo Settings → Pages → Deploy from a branch → root of `main` (or `/docs`
   if you prefer, adjusting paths).
3. GitHub gives you a URL like `https://yourorg.github.io/solarlink-connect/`
   — this must exactly match the SPA redirect URI you registered in Azure AD.

---

## Brands: verified vs. best-effort

| Brand | Status | Notes |
|---|---|---|
| FusionSolar | Verified | Built from your tested Postman collection + `fusionsolar_bot.py`. |
| SolarEdge | Verified | Built from your tested Postman collection + `solaredge_bot.py`. |
| Sungrow | Best-effort, untested | Login endpoint follows Sungrow's published guide; the historical-data query endpoint is only visible after your Developer Portal application is approved — update `js/brands.js` (`sungrow.fetchSeries`) once you have it. |
| Solis | Best-effort, untested | Uses SolisCloud's documented HMAC-SHA1 signing pattern (`js/signing.js#solisSign`) — verify field names/casing against your SolisCloud API PDF. |
| Growatt | Best-effort, untested | Two possible API surfaces wired in (Open API token vs. legacy portal login) — confirm which one your account actually has, and the historical-data endpoint is still a stub. |
| SMA | Best-effort, untested | OAuth2 + one-time plant-owner consent flow scaffolded; plant-listing and historical-data endpoints are stubs pending SMA's developer contract docs. |

As you get real docs/Postman collections for Sungrow, Solis, Growatt, or SMA,
share them and the corresponding sections of `js/brands.js` (browser) and
`automation/scripts/daily-extract.mjs` (scheduled job) can be filled in the
same way FusionSolar and SolarEdge were.
