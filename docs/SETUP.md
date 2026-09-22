# SolarLink Connect — Setup Guide

This app has three moving parts that each need a one-time setup step:

1. **The web page itself** (`index.html` + `js/`) — hosted on GitHub Pages, this is
   what you and your team open in a browser to add companies and run manual/
   on-demand extractions.
2. **A CORS proxy** — a tiny serverless function that lets the browser talk to
   inverter-brand APIs *and* Google Sheets (see "Why a proxy is required"
   below).
3. **A scheduled GitHub Action** — for true unattended daily/hourly extraction
   that doesn't depend on a browser tab being open.

Both (1) and (3) read/write the same Google Sheet, so they always agree on
what's already been extracted.

There's no Azure AD, no admin consent, and no "sign in" step anywhere in this
setup — everything runs off your own Google Cloud project and your own
GitHub repo.

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

The proxy also does one more job now: it holds the **Google service
account's private key**. That key must never reach the browser — anyone
could view-source a public GitHub Pages site and read it — so the browser
never talks to Google Sheets directly either. It sends save/load requests to
the proxy's `/sheets` route, and the proxy (which only you control, and
whose secrets are never sent to visitors) signs the actual Google API calls.

### Option A — Cloudflare Worker (`api-proxy/cloudflare-worker`) — recommended
1. Install [Wrangler](https://developers.cloudflare.com/workers/wrangler/): `npm install -g wrangler`.
2. `cd api-proxy/cloudflare-worker && wrangler login` (free Cloudflare account is fine).
3. Set the two sensitive values as encrypted secrets (never stored in a file):
   ```
   wrangler secret put GOOGLE_SERVICE_ACCOUNT_EMAIL
   wrangler secret put GOOGLE_SERVICE_ACCOUNT_KEY
   ```
   For the key, paste the **entire** `private_key` value from the service
   account's downloaded JSON file, including the `-----BEGIN PRIVATE
   KEY-----` / `-----END PRIVATE KEY-----` lines.
4. Edit `wrangler.toml` and set `GOOGLE_SPREADSHEET_ID` under `[vars]` (see
   "Google Sheet" below for where to find this ID) — this one isn't
   sensitive, so a plain var is fine.
5. `wrangler deploy`.
6. Set **Proxy base URL** in the web app's Settings to the printed
   `*.workers.dev` URL.

### Option B — Azure Function (`api-proxy/azure-function`)
Same idea, two functions (`CorsProxy` at `/api/relay`, `SheetsProxy` at
`/api/sheets`) instead of Worker routes. Deploy with the Azure Functions Core
Tools, and set `GOOGLE_SERVICE_ACCOUNT_EMAIL`, `GOOGLE_SERVICE_ACCOUNT_KEY`,
`GOOGLE_SPREADSHEET_ID` under the Function App's **Configuration →
Application settings** (these act like environment variables, not synced
anywhere public).

---

## Google Cloud service account

### 1. Create a project and enable the Sheets API
- [Google Cloud Console](https://console.cloud.google.com/) → create a new
  project (or reuse one) → **APIs & Services → Library** → search
  **Google Sheets API** → **Enable**.

### 2. Create the service account
- **APIs & Services → Credentials → Create credentials → Service account**.
- Give it a name (e.g. `solarlink-connect`), no roles needed at the project
  level — access is granted per-sheet instead (next step).
- Once created, open it → **Keys** tab → **Add key → Create new key → JSON**
  → downloads a `.json` file. This file contains the private key — treat it
  like a password. You'll paste two fields out of it (`client_email` and
  `private_key`) into your proxy's secrets and your GitHub secrets; you
  don't need to upload the file itself anywhere.

### 3. Create the Google Sheet
- Create a new Google Sheet (sheets.new).
- Rename the first tab to **`Connections`**, and add a second tab named
  **`Readings`** (Sheet → the `+` at the bottom).
- Add header rows (see `SolarLink_GoogleSheets_Template.xlsx` in this
  project for a ready-made starter you can just upload via **File → Import
  → Upload** on a new sheet, replacing the spreadsheet):
  - **Connections** row 1: `RowId, Title, Brand, Region, CredentialsJson, StationsJson, DailyAutoExtract, CursorJson`
  - **Readings** row 1: `Timestamp, Company, Brand, StationId, StationName, Resolution, kWh, RunAt`
- **Share the sheet** with your service account's email (the `client_email`
  field from the downloaded JSON, looks like
  `solarlink-connect@your-project.iam.gserviceaccount.com`) as **Editor**.
  This is the only "access grant" step in the whole setup — no admin
  approval needed, since you're sharing your own file the same way you'd
  share it with a colleague.
- Copy the **Spreadsheet ID** from its URL:
  `https://docs.google.com/spreadsheets/d/`**`THIS_PART`**`/edit` — this is
  what goes into `GOOGLE_SPREADSHEET_ID`.

**Security note:** brand credentials (API keys, iSolarCloud passwords, etc.)
are stored as plain text in the `Connections` tab so the extractor can use
them. Restrict who you share the sheet with the same way you'd protect the
credentials themselves.

---

## GitHub Actions secrets (for the scheduled daily/hourly extraction)

Repo → **Settings → Secrets and variables → Actions → New repository
secret**, add:

| Secret | Value |
|---|---|
| `GOOGLE_SERVICE_ACCOUNT_EMAIL` | The `client_email` field from the service account JSON |
| `GOOGLE_SERVICE_ACCOUNT_KEY` | The `private_key` field from the service account JSON (paste it exactly, including the BEGIN/END lines) |
| `GOOGLE_SPREADSHEET_ID` | Same ID as above |

You can reuse the **same** service account for both the proxy and the
GitHub Action — it doesn't need to be split into two credentials the way the
old SharePoint/Azure-AD setup did, since a service account key never has to
sit in a browser-reachable place and doesn't carry a "delegated vs.
app-only" distinction. Once the sheet is shared with it, both the proxy and
the Action can read/write it.

---

## Hosting the web page on GitHub Pages

1. Push this whole folder to a GitHub repo.
2. Repo Settings → Pages → Deploy from a branch → root of `main`.
3. GitHub gives you a URL like `https://yourorg.github.io/solarlink-connect/`.
4. Open it, go to **Settings**, fill in the **Proxy base URL** and **Google
   Spreadsheet ID**, save. No sign-in button anywhere — if both fields are
   set and the sheet is shared with the service account, the Dashboard
   should load (empty, until you add a company).

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
