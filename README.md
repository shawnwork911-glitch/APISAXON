# SolarLink Connect

A multi-brand inverter monitoring **API connection + data extraction** tool
for FusionSolar (Huawei), SolarEdge, Sungrow, Solis, Growatt, and SMA.

- **Hosted as a static site** (GitHub Pages) — `index.html` + `js/`.
- **Data lives in a Google Sheet** — connection credentials and extracted
  readings are stored there via a Google service account, held server-side
  by your proxy (never in the browser). No Azure AD, no admin consent.
- **Two ways to extract:** on-demand from the web page (pick a company,
  stations, resolution, date range → queue), or fully automated on a
  schedule via the included GitHub Action.

## Quick start

1. Read **[docs/SETUP.md](docs/SETUP.md)** first — there are three one-time
   setup steps (deploy a CORS proxy, create a Google service account +
   share a sheet with it, push this repo to GitHub Pages) before the app
   can do anything.
2. Open the hosted page, go to **Settings**, fill in the proxy URL and
   Google Spreadsheet ID, save.
3. **Dashboard → + Add Company** to connect your first brand.
4. **Extraction** tab to pull Hourly / Daily / Monthly data for a date range,
   or toggle **Daily auto-extract** on a company so the scheduled GitHub
   Action keeps it topped up automatically.

## Project layout

```
index.html                     the app shell
js/
  brands.js                    per-brand credential fields + API calls
  signing.js                   HMAC/MD5 request-signing helpers
  proxy-client.js              talks to the CORS proxy (brand API relay)
  sheets-client.js             talks to the proxy's Google Sheets relay
  extraction.js                quota-aware, resumable extraction job engine
  app.js                       UI wiring
api-proxy/
  azure-function/               CorsProxy (brand relay) + SheetsProxy (Sheets relay)
  cloudflare-worker/            same two jobs, one Worker (/relay and /sheets)
automation/
  scripts/daily-extract.mjs    the scheduled job's extraction logic
.github/workflows/
  daily-extraction.yml         runs daily-extract.mjs on a cron schedule
docs/SETUP.md                  full setup walkthrough
```

## What's verified vs. best-effort

FusionSolar and SolarEdge are built directly from your tested Postman
collections and Python bots. Sungrow, Solis, Growatt, and SMA are wired up
from each vendor's publicly published API docs but **have not been tested
against a live account** — see the table in `docs/SETUP.md` for exactly
what's confirmed vs. still a stub, and share real docs/Postman collections
for those brands to have them completed the same way.
