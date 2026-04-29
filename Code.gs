/**
 * ⚡ Discipline — Habit Tracker · Google Apps Script backend (TABULAR v2)
 *
 * SETUP (one time):
 *   1. Create / open a Google Sheet.
 *   2. Extensions → Apps Script → paste this file → Save.
 *   3. Deploy ▸ New deployment ▸ type = "Web app"
 *        Execute as: Me     ·     Who has access: Anyone
 *      Copy the /exec URL → paste into SHEETS_URL in src/App.jsx.
 *   4. Visit the /exec URL once. The sheets below are auto-created.
 *
 * STORAGE MODEL
 *   ┌──────────┐   ┌─────────────────┐   ┌──────────┐   ┌──────────┐
 *   │ Schema   │   │ Entries         │   │ Meta     │   │ Data     │
 *   │ (one row │   │ (one row per    │   │ key|value│   │ (legacy, │
 *   │  per     │   │  day; columns   │   │ theme:…  │   │  read-   │
 *   │  field)  │   │  driven by      │   │          │   │  only on │
 *   │          │   │  schema)        │   │          │   │  boot)   │
 *   └──────────┘   └─────────────────┘   └──────────┘   └──────────┘
 *
 *   Schema columns: id, icon, label, type, op, threshold, holidayThreshold,
 *                   weight, enabled, unit, max, step, inverted, skipHoliday,
 *                   scoreMode, message, targetLabel, rules
 *     • `rules` is JSON (variable-shape)
 *
 *   Entries columns (auto-managed):
 *     date, isHoliday, score, met, total, criteria, <fieldId>, <fieldId>, …
 *     • `criteria` is JSON (per-field met/null/false map)
 *     • Booleans stored as TRUE/FALSE, time as "HH:MM" string, numbers as numbers
 *
 *   Meta sheet: free-form key/value rows. Currently used for `theme`.
 *
 * API (unchanged from v1, web client needs no changes)
 *   GET  /exec                    → { entries: [...], schema: [...], theme: {...} }
 *   POST /exec  {key, value}      → upserts; key ∈ {entries, schema, theme}
 */

const SHEET_SCHEMA  = "Schema";
const SHEET_ENTRIES = "Entries";
const SHEET_META    = "Meta";
const SHEET_LEGACY  = "Data";

// Schema columns in display order. `rules` is JSON-encoded.
const SCHEMA_COLS = [
  "id","icon","label","type","op","threshold","holidayThreshold",
  "weight","enabled","unit","max","step","inverted","skipHoliday",
  "scoreMode","message","targetLabel","rules"
];
// Entries columns that are NOT field IDs.
const ENTRY_META_COLS = ["date","isHoliday","score","met","total","criteria"];

/* ─────────────── helpers ─────────────── */
function ss_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error("Script must be bound to a Sheet (Extensions → Apps Script from inside a Sheet).");
  return ss;
}
function sheet_(name, headers) {
  const s = ss_();
  let sh = s.getSheetByName(name);
  if (!sh) {
    sh = s.insertSheet(name);
    if (headers && headers.length) sh.appendRow(headers);
  }
  return sh;
}
function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
function safeJson_(s, fallback) {
  if (s === null || s === undefined || s === "") return fallback;
  if (typeof s === "object") return s;
  try { return JSON.parse(s); } catch { return fallback; }
}
function toBool_(v) {
  if (typeof v === "boolean") return v;
  if (typeof v === "number") return v !== 0;
  const s = String(v).trim().toLowerCase();
  return s === "true" || s === "1" || s === "yes" || s === "y";
}

/* ─────────────── Schema ─────────────── */
function readSchema_() {
  // 1. Prefer Schema sheet
  const s = ss_().getSheetByName(SHEET_SCHEMA);
  if (s) {
    const values = s.getDataRange().getValues();
    if (values.length >= 2) {
      const head = values[0].map(String);
      const out = [];
      for (let i = 1; i < values.length; i++) {
        const row = values[i];
        if (!row[head.indexOf("id")]) continue;
        const obj = {};
        head.forEach((h, j) => {
          let v = row[j];
          if (v === "" || v === null) return;
          if (h === "rules") obj[h] = safeJson_(v, []);
          else if (h === "enabled" || h === "inverted" || h === "skipHoliday") obj[h] = toBool_(v);
          else if (h === "threshold" || h === "holidayThreshold" || h === "weight" || h === "max" || h === "step") obj[h] = Number(v);
          else obj[h] = v;
        });
        out.push(obj);
      }
      return out;
    }
  }
  // 2. Legacy fallback: Data sheet single cell.
  const legacy = ss_().getSheetByName(SHEET_LEGACY);
  if (legacy) {
    const rows = legacy.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === "schema") return safeJson_(rows[i][1], null);
    }
  }
  return null;
}
function writeSchema_(arr) {
  if (!Array.isArray(arr)) throw new Error("schema must be an array");
  const sh = sheet_(SHEET_SCHEMA, SCHEMA_COLS);
  // Replace contents wholesale (small table, simplest correctness).
  sh.clear();
  sh.appendRow(SCHEMA_COLS);
  const rows = arr.map(f => SCHEMA_COLS.map(c => {
    const v = f[c];
    if (v === undefined || v === null) return "";
    if (c === "rules") return JSON.stringify(v);
    return v;
  }));
  if (rows.length) sh.getRange(2, 1, rows.length, SCHEMA_COLS.length).setValues(rows);
}

/* ─────────────── Entries ─────────────── */
function entryColumns_() {
  // ENTRY_META_COLS first, then one column per schema field id (in schema order).
  const schema = readSchema_() || [];
  const ids = schema.map(f => f.id).filter(Boolean);
  return ENTRY_META_COLS.concat(ids);
}
function readEntries_() {
  // Cache the script timezone — used to safely format Date-typed cells back
  // to "HH:mm" / "yyyy-MM-dd" without UTC drift (Apps Script's JSON.stringify
  // would otherwise emit Date.toJSON() = UTC ISO, which is offset and, for
  // pre-1900 sentinel dates, also bitten by Local Mean Time).
  const tz = Session.getScriptTimeZone();
  // 1. Prefer Entries sheet
  const s = ss_().getSheetByName(SHEET_ENTRIES);
  if (s) {
    const values = s.getDataRange().getValues();
    if (values.length >= 2) {
      const head = values[0].map(String);
      const out = [];
      for (let i = 1; i < values.length; i++) {
        const row = values[i];
        if (!row[head.indexOf("date")]) continue;
        const obj = {};
        head.forEach((h, j) => {
          let v = row[j];
          if (v === "" || v === null) return;
          if (h === "criteria") { obj[h] = safeJson_(v, {}); return; }
          if (h === "isHoliday") { obj[h] = toBool_(v); return; }
          if (h === "score" || h === "met" || h === "total") { obj[h] = Number(v); return; }
          if (h === "date") {
            obj[h] = (v instanceof Date)
              ? Utilities.formatDate(v, tz, "yyyy-MM-dd")
              : String(v);
            return;
          }
          // Any other column may be a Date (Sheets auto-converts strings like
          // "06:07" into a time-only Date). Format in script TZ so the wall
          // clock value the user typed is preserved.
          if (v instanceof Date) {
            obj[h] = Utilities.formatDate(v, tz, "HH:mm");
            return;
          }
          obj[h] = v;
        });
        out.push(obj);
      }
      return out;
    }
  }
  // 2. Legacy fallback: Data sheet single cell
  const legacy = ss_().getSheetByName(SHEET_LEGACY);
  if (legacy) {
    const rows = legacy.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === "entries") return safeJson_(rows[i][1], []);
    }
  }
  return [];
}
function writeEntries_(arr) {
  if (!Array.isArray(arr)) throw new Error("entries must be an array");
  const cols = entryColumns_();
  const sh = sheet_(SHEET_ENTRIES, cols);
  sh.clear();
  sh.appendRow(cols);
  // Force every data column to plain-text format BEFORE writing values, so
  // Sheets doesn't auto-convert "06:30" to a time fraction or "2026-04-29"
  // to a Date object on round-trip. Only `score`/`met`/`total` should be
  // numeric, and `isHoliday` boolean — those still serialise fine as text.
  const dataRows = Math.max(arr.length, 1);
  sh.getRange(2, 1, dataRows, cols.length).setNumberFormat("@");
  // Stable sort by date asc for human readability.
  const sorted = arr.slice().sort((a,b) => String(a.date||"").localeCompare(String(b.date||"")));
  const rows = sorted.map(e => cols.map(c => {
    const v = e[c];
    if (v === undefined || v === null) return "";
    if (c === "criteria") return JSON.stringify(v);
    if (typeof v === "boolean") return v;
    return v;
  }));
  if (rows.length) sh.getRange(2, 1, rows.length, cols.length).setValues(rows);
}

/* ─────────────── Meta (theme & misc) ─────────────── */
function readMeta_(key) {
  const s = ss_().getSheetByName(SHEET_META);
  if (!s) return null;
  const rows = s.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === key) return safeJson_(rows[i][1], rows[i][1]);
  }
  return null;
}
function writeMeta_(key, value) {
  const sh = sheet_(SHEET_META, ["key","value"]);
  const rows = sh.getDataRange().getValues();
  const json = (typeof value === "object") ? JSON.stringify(value) : String(value);
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === key) { sh.getRange(i+1, 2).setValue(json); return; }
  }
  sh.appendRow([key, json]);
}

/* ─────────────── HTTP handlers ─────────────── */

/**
 * Shared-secret check. Set the secret once via the Apps Script editor:
 *   File ▸ Project Settings ▸ Script Properties ▸ Add “APP_SECRET” → your passcode
 * Or run this from the editor:  PropertiesService.getScriptProperties().setProperty('APP_SECRET','your-passcode');
 * If APP_SECRET is unset the API stays open (back-compat for first-time setup).
 */
function checkAuth_(provided) {
  const expected = PropertiesService.getScriptProperties().getProperty("APP_SECRET");
  if (!expected) return true;            // not configured → open (so you can do first-time setup)
  return String(provided || "") === String(expected);
}
function unauthorized_() {
  return jsonOut_({ error: "unauthorized", code: 401 });
}

function doGet(e) {
  try {
    const auth = e && e.parameter && e.parameter.k;
    if (!checkAuth_(auth)) return unauthorized_();
    return jsonOut_({
      entries: readEntries_(),
      schema:  readSchema_(),
      theme:   readMeta_("theme"),
    });
  } catch (err) {
    return jsonOut_({ error: String(err && err.message || err) });
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonOut_({ error: "Empty request body" });
    }
    const body = JSON.parse(e.postData.contents);
    if (!checkAuth_(body && body.k)) return unauthorized_();
    if (!body || typeof body.key !== "string") {
      return jsonOut_({ error: "Body must be { key, value, k }" });
    }
    switch (body.key) {
      case "schema":  writeSchema_(body.value);   break;
      case "entries": writeEntries_(body.value);  break;
      case "theme":   writeMeta_("theme", body.value); break;
      default: writeMeta_(body.key, body.value);  break;
    }
    return jsonOut_({ ok: true, key: body.key });
  } catch (err) {
    return jsonOut_({ error: String(err && err.message || err) });
  }
}

/* ─────────────── one-shot migration helper ───────────────
 * Run from the Apps Script editor (Run ▸ migrate_) once if you previously
 * stored data via v1's single-cell layout. It rewrites the data into the
 * new tabular sheets. Safe to re-run.
 */
function migrate_() {
  const legacy = ss_().getSheetByName(SHEET_LEGACY);
  if (!legacy) { Logger.log("No legacy Data sheet — nothing to migrate."); return; }
  const rows = legacy.getDataRange().getValues();
  let s = null, ents = null, theme = null;
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === "schema")  s     = safeJson_(rows[i][1], null);
    if (rows[i][0] === "entries") ents  = safeJson_(rows[i][1], []);
    if (rows[i][0] === "theme")   theme = safeJson_(rows[i][1], null);
  }
  if (s)     writeSchema_(s);
  if (ents)  writeEntries_(ents);
  if (theme) writeMeta_("theme", theme);
  Logger.log("Migration done. schema=%s entries=%s theme=%s",
    s ? s.length : 0, ents ? ents.length : 0, theme ? "yes" : "no");
}
