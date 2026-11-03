/* Offline simulation of the scheduler — fake Google Sheets, fake FusionSolar
   (25 calls/day) and SolarEdge, fake clock. Run: node test/simulate.mjs
   Checks the 1 Oct → 31 Dec example queued on 3 Nov: every day saved exactly
   once, never a day before it's complete, quota respected, job done on 1 Jan. */

import { generateKeyPairSync } from "node:crypto";
import assert from "node:assert/strict";
import { makeRunContext } from "../src/google.js";
import { runScheduler, appendJob, readJobs, JOB_COLUMNS } from "../src/scheduler.js";
import { buildSpans, planJob } from "../src/plan.js";

const { privateKey } = generateKeyPairSync("rsa", { modulusLength: 2048 });
const env = {
  GOOGLE_SERVICE_ACCOUNT_EMAIL: "sa@test", GOOGLE_SERVICE_ACCOUNT_KEY: privateKey.export({ type: "pkcs8", format: "pem" }),
  GOOGLE_SPREADSHEET_ID: "SHEET", MAX_SUBREQUESTS: "45", SETTLE_HOUR: "2",
};

/* ---------- fake clock ---------- */
let NOW = Date.UTC(2026, 10, 3, 2, 0); // 3 Nov 2026 10:00 SGT
const now = () => NOW;
const sleep = async ms => { NOW += ms; };
const sgtDay = ms => new Date(ms + 8 * 3600e3).toISOString().slice(0, 10);

/* ---------- fake Sheets ---------- */
const tabs = {
  Connections: [["RowId"], ["c1", "Ding Fu", "fusionsolar", "ap-sg5", JSON.stringify({ username: "u", systemCode: "s", region: "ap-sg5" }),
    JSON.stringify([{ id: "NE=1", name: "Roof A" }, { id: "NE=2", name: "Roof B" }, { id: "NE=3", name: "Roof C" }]), "FALSE", "{}", JSON.stringify({ utcOffset: 8 })],
    ["c2", "Edge Co", "solaredge", "", JSON.stringify({ siteId: "999", apiKey: "k" }), JSON.stringify([{ id: "999", name: "Edge Site" }]), "FALSE", "{}", "null"]],
  Readings: [["Timestamp"]], Audit: [["Timestamp"]],
};
const rowsOf = (tab, from) => (tabs[tab] || []).slice(from - 1);

/* ---------- fake brands ---------- */
const fsCalls = new Map(); // sgt day -> calls
let fsTotal = 0, seCalls = 0, outageHits = 0;

globalThis.fetch = async (url, opts = {}) => {
  url = String(url);
  const ok = (obj, status = 200, headers = {}) => new Response(JSON.stringify(obj), { status, headers });
  if (url.startsWith("https://oauth2.googleapis.com/token")) return ok({ access_token: "tok", expires_in: 3600 });

  if (url.startsWith("https://sheets.googleapis.com/")) {
    const path = decodeURIComponent(url.split("/spreadsheets/SHEET")[1]);
    const body = opts.body ? JSON.parse(opts.body) : null;
    let m;
    if (path.startsWith(":batchUpdate")) { tabs[body.requests[0].addSheet.properties.title] = []; return ok({}); }
    if (path.startsWith("/values:batchUpdate")) {
      for (const d of body.data) { const r = +d.range.match(/!A(\d+)/)[1]; tabs.Jobs[r - 1] = d.values[0]; }
      return ok({});
    }
    if ((m = path.match(/^\/values\/(\w+)!A:[A-Z]:append/))) {
      const t = tabs[m[1]]; if (!t) return new Response("Unable to parse range", { status: 400 });
      const first = t.length + 1; t.push(...body.values);
      return ok({ updates: { updatedRange: `${m[1]}!A${first}:T${t.length}` } });
    }
    if ((m = path.match(/^\/values\/(\w+)!A(\d+):[A-Z](\d+)/)) && opts.method === "PUT") {
      tabs[m[1]] = tabs[m[1]] || []; tabs[m[1]][+m[2] - 1] = body.values[0]; return ok({});
    }
    if ((m = path.match(/^\/values\/(\w+)!A(\d+):[A-Z]\d+$/))) {
      if (!tabs[m[1]]) return new Response("Unable to parse range: Jobs", { status: 400 });
      return ok({ values: rowsOf(m[1], +m[2]) });
    }
    throw new Error("unhandled sheets " + path);
  }

  if (url.includes("fusionsolar")) {
    if (url.endsWith("/login")) return ok({ success: true, failCode: 0, data: null }, 200, { "xsrf-token": "x" });
    const day = sgtDay(NOW);
    // one-hour outage on 10 Nov — the job must retry next hour, not stop
    if (NOW >= Date.UTC(2026, 10, 9, 18) && NOW < Date.UTC(2026, 10, 9, 19)) { outageHits++; return new Response("Service Unavailable", { status: 503 }); }
    const used = fsCalls.get(day) || 0;
    if (used >= 25) return ok({ success: false, failCode: 407, message: "ACCESS_FREQUENCY_IS_TOO_HIGH" });
    fsCalls.set(day, used + 1); fsTotal++;
    const { stationCodes, collectTime } = JSON.parse(opts.body);
    const reqDay = sgtDay(collectTime);
    // a day's data must be complete (it's past that day in SGT) when requested
    assert.ok(reqDay < sgtDay(NOW), `requested ${reqDay} while it is still ${sgtDay(NOW)} SGT`);
    const dayStartUtc = Date.parse(reqDay + "T00:00:00Z") - 8 * 3600e3;
    const data = [];
    for (const code of stationCodes.split(","))
      for (let h = 0; h < 24; h++) data.push({ stationCode: code, collectTime: dayStartUtc + h * 3600e3, dataItemMap: { inverter_power: 1.5 } });
    return ok({ success: true, failCode: 0, data });
  }

  if (url.includes("monitoringapi.solaredge.com")) {
    seCalls++;
    const u = new URL(url);
    const s = u.searchParams.get("startTime").slice(0, 10), e = u.searchParams.get("endTime").slice(0, 10);
    const values = [];
    for (let d = Date.parse(s + "T00:00:00Z"); d <= Date.parse(e + "T00:00:00Z"); d += 86400e3)
      for (let h = 0; h < 24; h++) values.push({ date: `${new Date(d).toISOString().slice(0, 10)} ${String(h).padStart(2, "0")}:00:00`, value: 1000 });
    return ok({ energyDetails: { meters: [{ type: "Production", values }] } });
  }
  throw new Error("unhandled fetch " + url);
};

/* ---------- unit checks on the date logic ---------- */
assert.deepEqual(buildSpans("fusionsolar", "Daily", "2026-10-20", "2026-12-05"),
  [["2026-10-20", "2026-10-31"], ["2026-11-01", "2026-11-30"], ["2026-12-01", "2026-12-05"]]);
assert.equal(buildSpans("solaredge", "Hourly", "2026-10-01", "2026-12-31").length, 4);
// 01:00 SGT on 4 Nov: yesterday (3 Nov) not settled yet, so only through 2 Nov
assert.equal(planJob({ resolution: "Hourly", startDate: "2026-10-01", endDate: "2026-12-31", cursor: "2026-11-02" }, Date.UTC(2026, 10, 3, 17), 8).state, "waiting");
assert.equal(planJob({ resolution: "Hourly", startDate: "2026-10-01", endDate: "2026-12-31", cursor: "2026-11-02" }, Date.UTC(2026, 10, 3, 19), 8).availEnd, "2026-11-03");
// Monthly: November only once it has fully ended
assert.equal(planJob({ resolution: "Monthly", startDate: "2026-10-01", endDate: "2026-12-31", cursor: "2026-10-31" }, Date.UTC(2026, 10, 30, 17), 8).state, "waiting");
assert.equal(planJob({ resolution: "Monthly", startDate: "2026-10-01", endDate: "2026-12-31", cursor: "2026-10-31" }, Date.UTC(2026, 10, 30, 19), 8).availEnd, "2026-11-30");

/* ---------- the user's example ---------- */
const mk = async (over) => {
  const rc = makeRunContext(env, { now, sleep });
  await readJobs(rc); // creates Jobs tab first time
  return appendJob(rc, { jobId: over.jobId, connectionId: over.connectionId, company: over.company, brand: over.brand, resolution: "Hourly",
    stationIds: over.stationIds, startDate: "2026-10-01", endDate: "2026-12-31", skip: over.skip || [], cursor: "", status: "active",
    quotaDate: "", callsToday: 0, note: "", rowsSaved: 0, createdBy: "t", createdAt: "", updatedAt: "", lastError: "" });
};
await mk({ jobId: "J1", connectionId: "c1", company: "Ding Fu", brand: "fusionsolar", stationIds: ["NE=1", "NE=2", "NE=3"], skip: ["2026-10-05"] });
await mk({ jobId: "J2", connectionId: "c2", company: "Edge Co", brand: "solaredge", stationIds: ["999"] });
assert.equal(tabs.Jobs[0].length, JOB_COLUMNS.length);

// "queue" run (onlyJobId, ~27 s budget), then hourly cron runs until 3 Jan
{
  const rc = makeRunContext(env, { now, sleep, deadlineMs: NOW + 27000 });
  const s = await runScheduler(rc, { onlyJobId: "J1" });
  console.log("queue-time run:", JSON.stringify(s.processed), "subrequests", rc.used);
  assert.ok(rc.used <= 45);
}
const timeline = [];
let lastDay = sgtDay(NOW), doneAt = {};
NOW = Date.UTC(2026, 10, 3, 3, 7);
while (NOW < Date.UTC(2027, 0, 3)) {
  const rc = makeRunContext(env, { now, sleep, deadlineMs: NOW + 12 * 60e3 });
  const start = NOW;
  await runScheduler(rc);
  assert.ok(rc.used <= 45, `run used ${rc.used} subrequests`);
  const jobs = await readJobs(makeRunContext(env, { now, sleep }));
  for (const j of jobs) if (j.status === "done" && !doneAt[j.jobId]) doneAt[j.jobId] = sgtDay(NOW) + " " + new Date(NOW + 8 * 3600e3).toISOString().slice(11, 16) + " SGT";
  const d = sgtDay(NOW);
  if (d !== lastDay) { const j1 = jobs.find(j => j.jobId === "J1"); timeline.push(`${lastDay}: FS calls ${fsCalls.get(lastDay) || 0}, J1 saved through ${j1.cursor}`); lastDay = d; }
  NOW = start + 3600e3; // next hourly cron
}

console.log(timeline.slice(0, 6).join("\n"), "\n…\n" + timeline.slice(-3).join("\n"));
const jobs = await readJobs(makeRunContext(env, { now, sleep }));
console.log("final:", jobs.map(j => `${j.jobId} ${j.status} cursor=${j.cursor} rows=${j.rowsSaved} | ${j.note}`).join("\n       "));
console.log("done at:", doneAt);

// Readings: every day 1 Oct–31 Dec exactly 24 rows per station (except skipped 5 Oct for J1), no duplicates
const fsRows = tabs.Readings.slice(1).filter(r => r[1] === "Ding Fu");
const keys = new Set(fsRows.map(r => r[0] + "|" + r[3]));
assert.equal(keys.size, fsRows.length, "duplicate readings");
assert.equal(fsRows.length, (92 - 1) * 24 * 3);
assert.ok(!fsRows.some(r => r[0].startsWith("2026-10-05")));
assert.equal(fsRows[0][0], "2026-10-01 00:00:00");
const seRows = tabs.Readings.slice(1).filter(r => r[1] === "Edge Co");
assert.equal(new Set(seRows.map(r => r[0])).size, seRows.length);
assert.equal(seRows.length, 92 * 24);
for (const [d, n] of fsCalls) assert.ok(n <= 25, `${d} used ${n}`);
assert.ok(jobs.every(j => j.status === "done"));
assert.ok(outageHits > 0, 'outage not exercised');
console.log(`\nOK — outage hits ${outageHits}, FusionSolar calls ${fsTotal}, SolarEdge calls ${seCalls}, readings ${tabs.Readings.length - 1}`);
