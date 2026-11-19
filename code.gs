/**
 * Portfolio backend: a Google Sheet used as a database, exposed as an Apps Script web app.
 * Container-bound: create this script from the Sheet via Extensions > Apps Script.
 *
 * Security model
 *  - Anyone can READ published projects (doGet?action=public).
 *  - Only YOU can write or read private data. Every POST must carry a Google ID token.
 *    The token is verified against Google, and its email must equal OWNER_EMAIL.
 *  - Set two Script properties (Project Settings > Script properties):
 *      OWNER_EMAIL = your Google account email
 *      CLIENT_ID   = your OAuth Web client ID (see SETUP.md)
 */
const SCHEMA = {
  tasks:    ["id", "text", "tag", "done", "pos", "updated", "deleted"],
  projects: ["id", "title", "category", "icon", "description", "result", "tags", "status", "dataset",
             "live_url", "github_url", "kaggle_url", "question", "action", "published", "updated", "deleted"],
  kv:       ["key", "value", "updated"]
};
const KV_KEYS = ["journey", "ready", "formulas"];
const MAX_ROWS = 500, MAX_FIELD = 2000, TOMBSTONE_MS = 30 * 24 * 3600 * 1000;

/** Run once from the editor: creates the tabs and a dashboard of Sheets formulas. */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SCHEMA).forEach(function (name) {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    const cols = SCHEMA[name];
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight("bold");
    sh.getRange(2, 1, 998, cols.length).setNumberFormat("@");
    sh.setFrozenRows(1);
  });
  const d = ss.getSheetByName("dashboard") || ss.insertSheet("dashboard");
  d.clear();
  d.getRange("A1:B1").setValues([["Metric", "Value"]]).setFontWeight("bold");
  d.getRange("A2:A5").setValues([["Open tasks"], ["Done tasks"], ["Published projects"], ["Projects by category"]]);
  d.getRange("B2").setFormula('=SUMPRODUCT((tasks!A2:A500<>"")*(tasks!D2:D500="false")*(tasks!G2:G500<>"true"))');
  d.getRange("B3").setFormula('=SUMPRODUCT((tasks!A2:A500<>"")*(tasks!D2:D500="true")*(tasks!G2:G500<>"true"))');
  d.getRange("B4").setFormula('=SUMPRODUCT((projects!A2:A500<>"")*(projects!O2:O500="true")*(projects!Q2:Q500<>"true"))');
  d.getRange("B5").setFormula("=QUERY(projects!A1:Q500,\"select C, count(A) where A is not null and Q <> 'true' group by C label count(A) 'Projects'\",1)");
  d.setColumnWidth(1, 200);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** Public, read-only: health check and published projects. */
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || "ping";
  if (action === "public") {
    const keep = ["id", "title", "category", "icon", "description", "result", "tags", "status", "dataset",
                  "live_url", "github_url", "kaggle_url", "question", "action"];
    const rows = readTable_("projects")
      .filter(function (r) { return r.published === "true" && r.deleted !== "true"; })
      .map(function (r) { const o = {}; keep.forEach(function (k) { o[k] = r[k]; }); return o; });
    return json_({ ok: true, projects: rows, generated: new Date().toISOString() });
  }
  return json_({ ok: true, service: "portfolio-backend", time: new Date().toISOString() });
}

/** Owner-only: sync, list, upsert, remove. */
function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    verify_(req.token);
    switch (req.action) {
      case "sync":   return json_(sync_(req));
      case "list":   return json_({ ok: true, rows: readTable_(table_(req.table)) });
      case "upsert": return json_({ ok: true, rows: withLock_(function () { return mergeById_(table_(req.table), [req.row]); }) });
      case "remove": return json_({ ok: true, rows: withLock_(function () { return removeRow_(table_(req.table), req.id); }) });
      default: throw new Error("Unknown action.");
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function table_(name) {
  if (name !== "tasks" && name !== "projects") throw new Error("Unknown table.");
  return name;
}

function verify_(token) {
  const props = PropertiesService.getScriptProperties();
  const owner = String(props.getProperty("OWNER_EMAIL") || "").toLowerCase();
  const clientId = props.getProperty("CLIENT_ID");
  if (!owner || !clientId) throw new Error("Backend not configured: set OWNER_EMAIL and CLIENT_ID in Script properties.");
  if (!token) throw new Error("Sign in required.");
  const cache = CacheService.getScriptCache();
  const key = "t_" + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, token)).slice(0, 40);
  if (cache.get(key) === owner) return owner;
  const res = UrlFetchApp.fetch("https://oauth2.googleapis.com/tokeninfo?id_token=" + encodeURIComponent(token), { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error("Invalid or expired sign in. Sign in again.");
  const info = JSON.parse(res.getContentText());
  if (info.aud !== clientId) throw new Error("Sign in is not allowed for this app.");
  if (String(info.email_verified) !== "true") throw new Error("Your Google email is not verified.");
  if (Number(info.exp) * 1000 < Date.now()) throw new Error("Sign in expired. Sign in again.");
  if (String(info.email || "").toLowerCase() !== owner) throw new Error("This Google account is not allowed.");
  cache.put(key, owner, 300);
  return owner;
}

function readTable_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error("Missing tab '" + name + "'. Run setup() once.");
  const cols = SCHEMA[name], last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, cols.length).getValues()
    .filter(function (r) { return String(r[0]) !== ""; })
    .map(function (r) { const o = {}; cols.forEach(function (c, i) { o[c] = String(r[i]); }); return o; });
}

function writeTable_(name, rows) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  const cols = SCHEMA[name], last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, cols.length).clearContent();
  if (!rows.length) return;
  const rng = sh.getRange(2, 1, rows.length, cols.length);
  rng.setNumberFormat("@");   // plain text: blocks formula injection and keeps ids and booleans stable
  rng.setValues(rows.map(function (r) { return cols.map(function (c) { return r[c] === undefined ? "" : String(r[c]); }); }));
}

function clean_(table, obj) {
  const row = {};
  SCHEMA[table].forEach(function (c) {
    let v = obj && obj[c] !== undefined && obj[c] !== null ? String(obj[c]) : "";
    row[c] = v.length > MAX_FIELD ? v.slice(0, MAX_FIELD) : v;
  });
  return row;
}

/** Last-write-wins by id. Deleted rows stay as tombstones for 30 days so deletions sync. */
function mergeById_(table, incoming) {
  if (incoming.length > MAX_ROWS) throw new Error("Too many rows.");
  const map = {};
  readTable_(table).forEach(function (r) { map[r.id] = r; });
  incoming.forEach(function (raw) {
    const row = clean_(table, raw);
    if (!row.id) return;
    const cur = map[row.id];
    if (!cur || Number(row.updated || 0) >= Number(cur.updated || 0)) map[row.id] = row;
  });
  const now = Date.now();
  const rows = Object.keys(map).map(function (k) { return map[k]; }).filter(function (r) {
    return !(r.deleted === "true" && now - Number(r.updated || 0) > TOMBSTONE_MS);
  });
  writeTable_(table, rows);
  return rows;
}

function removeRow_(table, id) {
  const rows = readTable_(table).filter(function (r) { return r.id === id; });
  if (!rows.length) return readTable_(table);
  const r = rows[0];
  r.deleted = "true"; r.updated = String(Date.now());
  return mergeById_(table, [r]);
}

function mergeKv_(incoming) {
  const map = {};
  readTable_("kv").forEach(function (r) { map[r.key] = r; });
  Object.keys(incoming || {}).forEach(function (k) {
    if (KV_KEYS.indexOf(k) < 0) return;
    const inc = clean_("kv", { key: k, value: incoming[k].value, updated: incoming[k].updated });
    const cur = map[k];
    if (!cur || Number(inc.updated || 0) >= Number(cur.updated || 0)) map[k] = inc;
  });
  const rows = Object.keys(map).map(function (k) { return map[k]; });
  writeTable_("kv", rows);
  const out = {};
  rows.forEach(function (r) { out[r.key] = { value: r.value, updated: Number(r.updated || 0) }; });
  return out;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function sync_(req) {
  return withLock_(function () {
    return {
      ok: true,
      tasks: Array.isArray(req.tasks) ? mergeById_("tasks", req.tasks) : readTable_("tasks"),
      projects: Array.isArray(req.projects) ? mergeById_("projects", req.projects) : readTable_("projects"),
      kv: req.kv ? mergeKv_(req.kv) : mergeKv_({}),
      serverTime: Date.now()
    };
  });
}