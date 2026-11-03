# Auto-continue scheduler (Cloudflare Worker)

Finishes long or future-dated extractions on its own. When **Auto-continue on the server** is ticked on the Extraction tab, the job goes here instead of running in the browser. From then on, the job keeps running whether or not anyone has the page open:

- **Every hour (at :07 UTC)** the Worker continues each active job. It fetches from the day after the last saved day up to the earlier of:
  - the job's end date
  - the latest complete day (yesterday, once it's past 02:00 site time)
- **Daily API limits** are respected (FusionSolar 25 calls/day, SolarEdge 300). When the limit is used up, the job waits and continues the next day. All jobs on the same company share that company's limit.
- **When a job catches up**, it waits. Each new day is fetched once it's complete (each new month, for Monthly).
- **When the end date is saved**, the job is marked **Complete**.

Rows go to the **Readings** tab in the same format as browser extraction, so Compare and Export work with them unchanged. Job progress is kept in a **Jobs** tab, which is created automatically on first use. Each run is also written to the **Audit** tab (user "Auto-continue scheduler").

Example: queued on 3 Nov for 1 Oct → 31 Dec, FusionSolar Hourly:

| When | What gets saved |
|---|---|
| 3 Nov | 1 Oct → ~25 Oct |
| 4 Nov | the rest, up to 3 Nov (caught up) |
| 5 Nov → 1 Jan | one new day per day |
| 1 Jan, 02:07 site time | 31 Dec saved → **Complete** |

`npm test` runs this scenario offline (fake Sheets, fake FusionSolar/SolarEdge, fake clock). It checks for:

- duplicate rows
- days fetched before they were complete
- days that went over the quota

## Deploy (once)

From this folder:

```bash
npx wrangler login                                   # same Cloudflare account as solarlink-cors-proxy
npx wrangler secret put GOOGLE_SERVICE_ACCOUNT_EMAIL # same value as the CORS proxy
npx wrangler secret put GOOGLE_SERVICE_ACCOUNT_KEY   # same private key as the CORS proxy
npx wrangler deploy
```

`wrangler deploy` prints the Worker's URL, for example `https://solarlink-scheduler.saxon-solarlink.workers.dev`. That URL is already the app's default. If yours is different, paste it into **Settings → Auto-continue scheduler URL**.

This is a separate Worker from `solarlink-cors-proxy`, so the deployed proxy is not touched.

## Limits and settings (`wrangler.toml` `[vars]`)

- **`MAX_SUBREQUESTS`** (default `45`): Workers Free allows 50 outbound requests per run, so each run stops just short of that and the next hourly run carries on. On Workers Paid, raise it to `900` so a run can do a whole day's quota in one go.
- **`SETTLE_HOUR`** (default `2`): the site-local hour after which yesterday counts as complete. The site timezone comes from the company's UTC offset in its template settings (default +8).
- **Brands:** only FusionSolar and SolarEdge are supported. The other brands' historical-data endpoints aren't confirmed yet (see `js/brands.js`).
- **Brand calls only count against the quota on the side that makes them.** The quota counter covers server-side jobs only. If someone also runs a manual browser extraction for the same FusionSolar account on the same day, the brand may refuse calls early (failCode 407). The job then just continues the next day.
- **Logs:** `npx wrangler tail solarlink-scheduler` shows each run live.
