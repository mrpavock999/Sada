/**
 * ⚡ Discipline — Habit Tracker · Google Apps Script backend (STRUCTURED v3 — "DB revamp")
 *
 * SETUP (one time):
 *   1. Create / open a Google Sheet.
 *   2. Extensions → Apps Script → paste this file → Save.
 *   3. Deploy ▸ New deployment ▸ type = "Web app"
 *        Execute as: Me     ·     Who has access: Anyone
 *      Copy the /exec URL → paste into SHEETS_URL in src/App.jsx.
 *   4. Visit the /exec URL once. The sheets below are auto-created.
 *   5. If you are upgrading from the old v2 backend (single JSON cells in
 *      `Meta`/`FoodEntries`), run `migrateFoodDataToStructuredSheets_` ONCE
 *      from this editor (Run ▸ select function ▸ migrateFoodDataToStructuredSheets_)
 *      after deploying. It is safe to re-run; it never deletes the old sheets.
 *
 * WHY THIS REVAMP (v2 → v3)
 *   v2 stored food items / meal templates / meal types / daily meal logs as a
 *   single JSON-encoded cell per record (or even per whole collection). That
 *   is unreadable in the Sheet UI, risks hitting Sheets' ~50,000 char/cell
 *   limit as data grows, and every save did `sheet.clear()` then rewrote the
 *   WHOLE collection — so a mid-write failure (quota, network drop, bug)
 *   could leave a sheet truncated or blank with no way back.
 *
 *   v3 keeps Google Sheets as the ONLY datastore (no external DB) but:
 *     • Stores every collection as ONE ROW PER RECORD, ONE COLUMN PER FIELD
 *       — fully readable/sortable/filterable directly in Sheets.
 *     • NEVER calls `sheet.clear()` on a save. Saves upsert matching rows in
 *       place and append new ones; a record removed from the app's list is
 *       marked `archived = TRUE` instead of being deleted, so history (and
 *       any day/column referencing it) is never silently destroyed.
 *     • Validates every incoming payload BEFORE touching a sheet, so a bad
 *       request from a buggy frontend build can't corrupt stored data.
 *     • Snapshots food-item nutrients into each logged meal line at the time
 *       it was logged, so editing/deleting a food item later never rewrites
 *       history.
 *     • Makes a dated full-spreadsheet backup copy (via Drive) once per day,
 *       and appends every write to an audit `ChangeLog` sheet.
 *   The public GET/POST contract is UNCHANGED — the web app needs no changes.
 *
 * STORAGE MODEL (v3)
 *   Habit tracking (unchanged on-disk layout, now upsert-based):
 *     Schema   — one row per habit field (archived flag added; never cleared)
 *     Entries  — one row per logged day (columns driven by Schema incl.
 *                archived habit ids, so old values are never orphaned)
 *   Food tracking (NEW structured sheets — replaces v2's JSON-blob storage):
 *     FoodItems         — one row per food item
 *     MealTypes         — one row per meal category (breakfast/lunch/…)
 *     MealTemplates     — one row per template header
 *     MealTemplateItems — one row per (template, food item, qty) line
 *     FoodLog           — one row per logged meal instance (date, mealType)
 *     FoodLogItems      — one row per logged food line, WITH a nutrient
 *                         snapshot taken at log time (name/calories/protein/
 *                         carbs/fat/fibre) so later edits to FoodItems can
 *                         never retroactively change historical totals
 *   Small bounded config (kept as simple key→value rows; will never remotely
 *   approach the cell-size limit):
 *     Meta     — theme palette, food targets/habit-links (`foodSettings`)
 *   Safety / operations:
 *     ChangeLog — append-only audit trail: timestamp, key, record count
 *     (daily spreadsheet backup copies are created in the same Drive folder)
 *   Legacy / backup (left untouched, never written to by v3, never deleted):
 *     Data        — v1 single-cell fallback
 *     FoodEntries — v2 single-cell-per-day JSON food log
 *
 * API (unchanged — the web client needs no changes for this revamp)
 *   GET  /exec                → { entries, schema, theme, foodItems,
 *                                  foodSettings, mealTemplates, mealTypes,
 *                                  foodEntries }
 *   POST /exec {key, value}   → upserts; key ∈ {schema, entries, theme,
 *                                  foodItems, mealTypes, mealTemplates,
 *                                  foodEntries, foodSettings, ...}
 */

const SHEET_SCHEMA  = "Schema";
const SHEET_ENTRIES = "Entries";
const SHEET_META    = "Meta";
const SHEET_LEGACY  = "Data";
// v2 legacy single-JSON-cell-per-day food log — left in place as a read-only
// backup/migration source. v3 writes/reads food logs via SHEET_FOOD_LOG(_ITEMS).
const SHEET_FOOD_ENTRIES_LEGACY = "FoodEntries";

// v3 structured food sheets (one row per record — see header comment above).
const SHEET_FOOD_ITEMS          = "FoodItems";
const SHEET_MEAL_TYPES          = "MealTypes";
const SHEET_MEAL_TEMPLATES      = "MealTemplates";
const SHEET_MEAL_TEMPLATE_ITEMS = "MealTemplateItems";
const SHEET_FOOD_LOG            = "FoodLog";
const SHEET_FOOD_LOG_ITEMS      = "FoodLogItems";
const SHEET_CHANGELOG           = "ChangeLog";

// Schema columns in display order. `rules` is JSON-encoded. `archived` marks
// a habit the user deleted from the UI — kept (not removed) so its historical
// values in Entries are never orphaned/lost. `order` mirrors the app's list
// order so the Sheet reads top-to-bottom the same way the UI does.
const SCHEMA_COLS = [
  "id","icon","label","type","op","threshold","holidayThreshold",
  "weight","enabled","unit","max","step","inverted","skipHoliday",
  "scoreMode","message","targetLabel","trackingSince","rules","order","archived"
];
// Entries columns that are NOT field IDs.
const ENTRY_META_COLS = ["date","isHoliday","score","met","total","criteria"];

const FOOD_ITEM_COLS = [
  "id","icon","name","basis","qtyUnit","servingQty",
  "calories","protein","carbs","fat","fibre","enabled","order","archived",
];
const MEAL_TYPE_COLS = ["k","icon","label","order","archived"];
const MEAL_TEMPLATE_COLS = ["id","icon","name","mealType","order","archived"];
const MEAL_TEMPLATE_ITEM_COLS = ["templateId","itemId","qty","order"];
const FOOD_LOG_COLS = ["id","date","mealType","order","archived"];
const FOOD_LOG_ITEM_COLS = [
  "mealId","itemId","qty","order",
  "nameSnapshot","caloriesSnapshot","proteinSnapshot","carbsSnapshot","fatSnapshot","fibreSnapshot",
];

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

/* ─────────────── generic structured-sheet helpers (v3) ───────────────
 * These implement the "never clear(), always upsert-by-key, archive instead
 * of delete" pattern shared by every record-oriented sheet below.
 */
function readHeader_(sh) {
  const lastCol = sh.getLastColumn();
  if (!lastCol) return [];
  return sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
}
// Extends (never reorders/removes) a sheet's header with any columns in
// `requiredCols` it doesn't already have. Existing data/columns are untouched;
// older rows simply read back as blank ("") for the newly added column until
// they're next upserted — Apps Script pads short rows automatically on read.
function ensureHeader_(sh, currentHeader, requiredCols) {
  if (!currentHeader.length) {
    sh.getRange(1, 1, 1, requiredCols.length).setValues([requiredCols]);
    return requiredCols.slice();
  }
  const missing = requiredCols.filter(c => currentHeader.indexOf(c) === -1);
  if (!missing.length) return currentHeader;
  const newHeader = currentHeader.concat(missing);
  sh.getRange(1, 1, 1, newHeader.length).setValues([newHeader]);
  return newHeader;
}
// Upserts `records` (each MUST have a unique, non-empty `keyCol` value) into
// `sh` by that key: matching rows are updated in place, new ones appended.
// Any EXISTING row whose key is absent from `records` is marked archived
// (never deleted) — so removing something from an app list never destroys
// the row or anything that still references it. `jsonCols` lists fields that
// must be JSON-encoded (e.g. a habit's `rules`).
function upsertByKeyAndArchiveMissing_(sh, requiredCols, keyCol, records, jsonCols) {
  jsonCols = jsonCols || [];
  const header = ensureHeader_(sh, readHeader_(sh), requiredCols);
  const values = sh.getDataRange().getValues();
  const keyIdx = header.indexOf(keyCol);
  const archivedIdx = header.indexOf("archived");
  const existingRowByKey = {};
  for (let i = 1; i < values.length; i++) {
    const k = values[i][keyIdx];
    if (k !== "" && k !== null && k !== undefined) existingRowByKey[k] = i + 1;
  }
  const incomingKeys = new Set();
  records.forEach((rec, idx) => {
    const key = rec && rec[keyCol];
    if (key === undefined || key === null || key === "") {
      throw new Error(`Record missing required '${keyCol}'`);
    }
    incomingKeys.add(key);
    const rowVals = header.map(c => {
      if (c === "archived") return false;
      // Always derive `order` from the record's position in THIS save's
      // array — never trust an echoed-back value, so reordering in the app
      // (which just reorders the JS array, not a field on each object) is
      // reflected correctly every time, with no risk of a stale value
      // "freezing" a record's position after its first round-trip.
      if (c === "order") return idx;
      const v = rec[c];
      if (v === undefined || v === null) return "";
      if (jsonCols.indexOf(c) >= 0) return JSON.stringify(v);
      return v;
    });
    const r = existingRowByKey[key];
    if (r) sh.getRange(r, 1, 1, header.length).setValues([rowVals]);
    else sh.appendRow(rowVals);
  });
  if (archivedIdx >= 0) {
    for (let i = 1; i < values.length; i++) {
      const key = values[i][keyIdx];
      if (key !== "" && key !== null && key !== undefined &&
          !incomingKeys.has(key) && !toBool_(values[i][archivedIdx])) {
        sh.getRange(i + 1, archivedIdx + 1).setValue(true);
      }
    }
  }
}
// Reads back a record-oriented sheet into an array of plain objects, skipping
// archived rows unless `includeArchived` is true, and sorting by an `order`
// column (if present) so results match the app's intended list order.
function readRecordSheet_(sheetName, numericCols, jsonCols, includeArchived) {
  const s = ss_().getSheetByName(sheetName);
  if (!s) return null; // null = "sheet doesn't exist yet" (pre-migration), distinct from []
  const values = s.getDataRange().getValues();
  if (values.length < 2) return [];
  const head = values[0].map(String);
  const keyIdx = 0; // by convention, the first column of every record sheet is its unique key
  const archivedIdx = head.indexOf("archived");
  const orderIdx = head.indexOf("order");
  numericCols = numericCols || [];
  jsonCols = jsonCols || [];
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (row[keyIdx] === "" || row[keyIdx] === null || row[keyIdx] === undefined) continue;
    if (!includeArchived && archivedIdx >= 0 && toBool_(row[archivedIdx])) continue;
    const obj = {};
    head.forEach((h, j) => {
      const v = row[j];
      if (v === "" || v === null || v === undefined) return;
      if (h === "archived") { obj[h] = toBool_(v); return; }
      if (h === "enabled") { obj[h] = toBool_(v); return; }
      if (jsonCols.indexOf(h) >= 0) { obj[h] = safeJson_(v, null); return; }
      if (numericCols.indexOf(h) >= 0) { obj[h] = Number(v); return; }
      obj[h] = v;
    });
    obj.__order = orderIdx >= 0 ? (Number(row[orderIdx]) || 0) : out.length;
    out.push(obj);
  }
  out.sort((a, b) => a.__order - b.__order);
  out.forEach(o => { delete o.__order; });
  return out;
}

/* ─────────────── Schema (habits) ───────────────
 * `archived` habits are hidden from the app (readSchema_ excludes them by
 * default) but their row — and their historical value column in Entries —
 * is kept forever so past days never lose data just because a habit was
 * later removed from the UI.
 */
function readSchema_(includeArchived) {
  // 1. Prefer Schema sheet
  const s = ss_().getSheetByName(SHEET_SCHEMA);
  if (s) {
    const values = s.getDataRange().getValues();
    if (values.length >= 2) {
      const head = values[0].map(String);
      const archivedIdx = head.indexOf("archived");
      const orderIdx = head.indexOf("order");
      const out = [];
      for (let i = 1; i < values.length; i++) {
        const row = values[i];
        if (!row[head.indexOf("id")]) continue;
        if (!includeArchived && archivedIdx >= 0 && toBool_(row[archivedIdx])) continue;
        const obj = {};
        head.forEach((h, j) => {
          let v = row[j];
          if (v === "" || v === null) return;
          if (h === "rules") obj[h] = safeJson_(v, []);
          else if (h === "archived") obj[h] = toBool_(v);
          else if (h === "enabled" || h === "inverted" || h === "skipHoliday") obj[h] = toBool_(v);
          else if (h === "threshold" || h === "holidayThreshold" || h === "weight" || h === "max" || h === "step" || h === "order") obj[h] = Number(v);
          else obj[h] = v;
        });
        obj.__order = orderIdx >= 0 ? (Number(row[orderIdx]) || 0) : out.length;
        out.push(obj);
      }
      out.sort((a, b) => a.__order - b.__order);
      out.forEach(o => { delete o.__order; });
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
  upsertByKeyAndArchiveMissing_(sh, SCHEMA_COLS, "id", arr, ["rules"]);
}

/* ─────────────── Entries (habit log) ───────────────
 * One row per logged day. Saves UPSERT by date — a date missing from the
 * incoming payload is left completely untouched (never cleared), so a stale
 * or partial client-side copy can never wipe out history it doesn't know
 * about. `entryColumns_` always includes ARCHIVED habit ids too, so an old
 * habit's historical values stay readable even after it's deleted in the UI.
 */
function entryColumns_() {
  const schema = readSchema_(true) || []; // include archived — preserve their columns
  const ids = schema.map(f => f.id).filter(Boolean);
  return ENTRY_META_COLS.concat(ids);
}
function readEntries_() {
  // Cache the script timezone — used to safely format Date-typed cells back
  // to "HH:mm" / "yyyy-MM-dd" without UTC drift (Apps Script's JSON.stringify
  // would otherwise emit Date.toJSON() = UTC ISO, which is offset and, for
  // pre-1900 sentinel dates, also bitten by Local Mean Time).
  const tz = Session.getScriptTimeZone();
  // Pre-resolve toggle-typed field ids so old data with string "true"/"false"
  // values is coerced back to real booleans on read.
  const schemaForRead = readSchema_(true) || [];
  const toggleIds = {};
  schemaForRead.forEach(f => { if (f && f.type === "toggle") toggleIds[f.id] = true; });
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
          if (toggleIds[h]) { obj[h] = toBool_(v); return; }
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
  const schema = readSchema_(true) || [];
  const typeById = {};
  schema.forEach(f => { if (f && f.id) typeById[f.id] = f.type; });
  const sh = sheet_(SHEET_ENTRIES, cols);
  const header = ensureHeader_(sh, readHeader_(sh), cols);
  const values = sh.getDataRange().getValues();
  const dateIdx = header.indexOf("date");
  const existingRowByDate = {};
  for (let i = 1; i < values.length; i++) {
    const d = values[i][dateIdx];
    if (d !== "" && d !== null && d !== undefined) existingRowByDate[String(d)] = i + 1;
  }
  // Force plain-text format only on columns that hold human-readable
  // strings — `date` (YYYY-MM-DD) and any `time`-typed field (HH:mm).
  // Boolean columns must stay UNFORMATTED so Sheets stores them as native
  // booleans (TRUE/FALSE) instead of coercing them to lowercase strings
  // "true"/"false" — which, on read-back, both evaluate truthy.
  const lastRow = Math.max(sh.getLastRow(), 2);
  header.forEach((c, idx) => {
    const isString = c === "date" || typeById[c] === "time";
    if (isString) sh.getRange(2, idx + 1, lastRow - 1, 1).setNumberFormat("@");
  });
  // Upsert by date — dates NOT present in `arr` are left untouched, so a
  // stale/partial client payload can never wipe out history.
  arr.forEach(e => {
    const rowVals = header.map(c => {
      const v = e[c];
      if (v === undefined || v === null) return "";
      if (c === "criteria") return JSON.stringify(v);
      if (c === "isHoliday" || typeById[c] === "toggle") return toBool_(v);
      if (typeof v === "boolean") return v;
      return v;
    });
    const r = existingRowByDate[e.date];
    if (r) sh.getRange(r, 1, 1, header.length).setValues([rowVals]);
    else sh.appendRow(rowVals);
  });
}
// Cosmetic-only maintenance: re-sorts the Entries sheet chronologically.
// Safe to run any time from the Apps Script editor; never required for
// correctness (upserts above never depend on row order).
function resortEntriesByDate_() {
  const sh = ss_().getSheetByName(SHEET_ENTRIES);
  if (!sh) return;
  const values = sh.getDataRange().getValues();
  if (values.length < 3) return;
  const header = values[0];
  const body = values.slice(1).sort((a, b) => String(a[0] || "").localeCompare(String(b[0] || "")));
  sh.getRange(2, 1, body.length, header.length).setValues(body);
}

/* ─────────────── Food Items ─────────────── */
function readFoodItems_(includeArchived) {
  return readRecordSheet_(
    SHEET_FOOD_ITEMS,
    ["servingQty", "calories", "protein", "carbs", "fat", "fibre", "order"],
    [],
    includeArchived
  );
}
function writeFoodItems_(arr) {
  if (!Array.isArray(arr)) throw new Error("foodItems must be an array");
  const sh = sheet_(SHEET_FOOD_ITEMS, FOOD_ITEM_COLS);
  upsertByKeyAndArchiveMissing_(sh, FOOD_ITEM_COLS, "id", arr, []);
}

/* ─────────────── Meal Types ─────────────── */
function readMealTypes_(includeArchived) {
  return readRecordSheet_(SHEET_MEAL_TYPES, ["order"], [], includeArchived);
}
function writeMealTypes_(arr) {
  if (!Array.isArray(arr)) throw new Error("mealTypes must be an array");
  const sh = sheet_(SHEET_MEAL_TYPES, MEAL_TYPE_COLS);
  upsertByKeyAndArchiveMissing_(sh, MEAL_TYPE_COLS, "k", arr, []);
}

/* ─────────────── Meal Templates (header + line-item sheets) ───────────────
 * `MealTemplates` holds one row per template; `MealTemplateItems` holds one
 * row per (template, food item, qty) line. Saving only rewrites the item
 * lines belonging to templates present in this save — lines for any
 * template NOT included in the payload are left completely untouched.
 */
function readMealTemplates_(includeArchived) {
  const templates = readRecordSheet_(SHEET_MEAL_TEMPLATES, ["order"], [], includeArchived);
  if (templates === null) return null; // not migrated yet
  templates.forEach(t => { t.items = []; });
  const byId = {};
  templates.forEach(t => { byId[t.id] = t; });
  const itemsSheet = ss_().getSheetByName(SHEET_MEAL_TEMPLATE_ITEMS);
  if (itemsSheet) {
    const iv = itemsSheet.getDataRange().getValues();
    if (iv.length >= 2) {
      const ih = iv[0].map(String);
      const tIdx = ih.indexOf("templateId"), itemIdx = ih.indexOf("itemId");
      const qtyIdx = ih.indexOf("qty"), orderIdx = ih.indexOf("order");
      const lines = [];
      for (let i = 1; i < iv.length; i++) {
        const row = iv[i];
        const tid = row[tIdx];
        if (!tid || !byId[tid]) continue;
        lines.push({
          tid, itemId: row[itemIdx],
          qty: row[qtyIdx] === "" ? undefined : Number(row[qtyIdx]),
          order: Number(row[orderIdx]) || 0,
        });
      }
      lines.sort((a, b) => a.order - b.order);
      lines.forEach(l => { byId[l.tid].items.push({ itemId: l.itemId, qty: l.qty }); });
    }
  }
  return templates;
}
function writeMealTemplateItems_(arr) {
  const sh = sheet_(SHEET_MEAL_TEMPLATE_ITEMS, MEAL_TEMPLATE_ITEM_COLS);
  const values = sh.getDataRange().getValues();
  const header = values.length ? values[0].map(String) : MEAL_TEMPLATE_ITEM_COLS.slice();
  const tIdx = header.indexOf("templateId");
  const touchedIds = new Set(arr.map(t => t.id));
  const rowsToDelete = [];
  for (let i = 1; i < values.length; i++) {
    if (touchedIds.has(values[i][tIdx])) rowsToDelete.push(i + 1);
  }
  // Only delete+rewrite lines for templates in THIS save; untouched templates'
  // lines are never read, never deleted.
  rowsToDelete.sort((a, b) => b - a).forEach(r => sh.deleteRow(r));
  arr.forEach(t => {
    (t.items || []).forEach((it, idx) => {
      sh.appendRow([t.id, it.itemId, it.qty === undefined ? "" : it.qty, idx]);
    });
  });
}
function writeMealTemplates_(arr) {
  if (!Array.isArray(arr)) throw new Error("mealTemplates must be an array");
  const sh = sheet_(SHEET_MEAL_TEMPLATES, MEAL_TEMPLATE_COLS);
  upsertByKeyAndArchiveMissing_(sh, MEAL_TEMPLATE_COLS, "id", arr, []);
  writeMealTemplateItems_(arr);
}

/* ─────────────── Food Log (daily meals) ───────────────
 * `FoodLog` holds one row per logged MEAL instance; `FoodLogItems` holds one
 * row per logged FOOD line, with a nutrient snapshot captured at log time —
 * so later renaming/editing/archiving a food item never rewrites a day's
 * historical totals. Saves only touch meals for the DATES included in this
 * payload; other dates' meals are left completely untouched.
 */
function itemNutrients_(item, qty) {
  const ref = Number(item && item.servingQty) || 1;
  const q = Number(qty) || 0;
  const f = ref > 0 ? q / ref : 0;
  return {
    calories: (Number(item && item.calories) || 0) * f,
    protein:  (Number(item && item.protein)  || 0) * f,
    carbs:    (Number(item && item.carbs)    || 0) * f,
    fat:      (Number(item && item.fat)      || 0) * f,
    fibre:    (Number(item && item.fibre)    || 0) * f,
  };
}
function replaceMealItemLines_(itemsSheet, mealId, items, itemsById) {
  const values = itemsSheet.getDataRange().getValues();
  const header = values.length ? values[0].map(String) : FOOD_LOG_ITEM_COLS.slice();
  const mIdx = header.indexOf("mealId");
  const rowsToDelete = [];
  for (let i = 1; i < values.length; i++) {
    if (values[i][mIdx] === mealId) rowsToDelete.push(i + 1);
  }
  rowsToDelete.sort((a, b) => b - a).forEach(r => itemsSheet.deleteRow(r));
  items.forEach((it, idx) => {
    const item = itemsById[it.itemId] || {};
    const n = itemNutrients_(item, it.qty);
    itemsSheet.appendRow([
      mealId, it.itemId, it.qty === undefined ? "" : it.qty, idx,
      item.name || "", n.calories, n.protein, n.carbs, n.fat, n.fibre,
    ]);
  });
}
function readFoodEntries_() {
  const headerSheet = ss_().getSheetByName(SHEET_FOOD_LOG);
  if (!headerSheet) return null; // not migrated yet
  const hv = headerSheet.getDataRange().getValues();
  if (hv.length < 2) return [];
  const hh = hv[0].map(String);
  const idIdx = hh.indexOf("id"), dateIdx = hh.indexOf("date");
  const typeIdx = hh.indexOf("mealType"), orderIdx = hh.indexOf("order"), archivedIdx = hh.indexOf("archived");
  const mealsById = {};
  const byDate = {};
  for (let i = 1; i < hv.length; i++) {
    const row = hv[i];
    const id = row[idIdx];
    if (!id) continue;
    if (archivedIdx >= 0 && toBool_(row[archivedIdx])) continue;
    const date = row[dateIdx];
    const meal = { id, mealType: row[typeIdx], items: [], __order: Number(row[orderIdx]) || 0 };
    mealsById[id] = meal;
    (byDate[date] = byDate[date] || []).push(meal);
  }
  const itemsSheet = ss_().getSheetByName(SHEET_FOOD_LOG_ITEMS);
  if (itemsSheet) {
    const iv = itemsSheet.getDataRange().getValues();
    if (iv.length >= 2) {
      const ih = iv[0].map(String);
      const mIdx = ih.indexOf("mealId"), itemIdx = ih.indexOf("itemId");
      const qtyIdx = ih.indexOf("qty"), ordIdx = ih.indexOf("order");
      for (let i = 1; i < iv.length; i++) {
        const row = iv[i];
        const meal = mealsById[row[mIdx]];
        if (!meal) continue;
        meal.items.push({
          itemId: row[itemIdx],
          qty: row[qtyIdx] === "" ? undefined : Number(row[qtyIdx]),
          __order: Number(row[ordIdx]) || 0,
        });
      }
    }
  }
  Object.values(mealsById).forEach(m => {
    m.items.sort((a, b) => a.__order - b.__order).forEach(it => { delete it.__order; });
    delete m.__order;
  });
  return Object.keys(byDate).sort().map(date => ({
    date,
    meals: byDate[date].sort((a, b) => a.__order - b.__order).map(m => {
      const { __order, ...rest } = m;
      return rest;
    }),
  }));
}
function writeFoodEntries_(arr) {
  if (!Array.isArray(arr)) throw new Error("foodEntries must be an array");
  const headerSheet = sheet_(SHEET_FOOD_LOG, FOOD_LOG_COLS);
  const itemsSheet = sheet_(SHEET_FOOD_LOG_ITEMS, FOOD_LOG_ITEM_COLS);
  const itemsById = {};
  (readFoodItems_(true) || []).forEach(it => { itemsById[it.id] = it; });

  const header = ensureHeader_(headerSheet, readHeader_(headerSheet), FOOD_LOG_COLS);
  const hv = headerSheet.getDataRange().getValues();
  const dateIdx = header.indexOf("date"), idIdx = header.indexOf("id"), archivedIdx = header.indexOf("archived");
  const existingRowById = {};
  const mealIdsByDate = {};
  for (let i = 1; i < hv.length; i++) {
    const id = hv[i][idIdx], date = hv[i][dateIdx];
    if (id) existingRowById[id] = i + 1;
    if (date) (mealIdsByDate[date] = mealIdsByDate[date] || new Set()).add(id);
  }
  const datesTouched = new Set(arr.map(d => d.date));
  const incomingMealIds = new Set();
  arr.forEach(day => {
    (day.meals || []).forEach((m, idx) => {
      incomingMealIds.add(m.id);
      const rowVals = header.map(c => {
        if (c === "id") return m.id;
        if (c === "date") return day.date;
        if (c === "mealType") return m.mealType;
        if (c === "order") return idx;
        if (c === "archived") return false;
        return "";
      });
      const r = existingRowById[m.id];
      if (r) headerSheet.getRange(r, 1, 1, header.length).setValues([rowVals]);
      else headerSheet.appendRow(rowVals);
      replaceMealItemLines_(itemsSheet, m.id, m.items || [], itemsById);
    });
  });
  // Archive (never delete) meal headers that disappeared from a TOUCHED
  // date's payload. Meals on dates not present in `arr` are never touched.
  if (archivedIdx >= 0) {
    datesTouched.forEach(date => {
      (mealIdsByDate[date] || new Set()).forEach(id => {
        if (id && !incomingMealIds.has(id)) {
          const r = existingRowById[id];
          if (r) headerSheet.getRange(r, archivedIdx + 1).setValue(true);
        }
      });
    });
  }
}

/* ─────────────── Meta (small bounded config: theme, food targets/links) ───
 * Intentionally still simple key→value JSON rows — these are a handful of
 * colours / targets / links, nowhere near the ~50,000 char/cell limit, so
 * the structured-sheet treatment above would be pure overhead here.
 */
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

/* ─────────────── Safety: validation, audit log, backups ─────────────── */
// Rejects a malformed payload BEFORE any sheet is touched, so a bad request
// (buggy frontend build, corrupted localStorage, etc.) can never partially
// corrupt stored data.
function validateWrite_(key, value) {
  const isArr = Array.isArray(value);
  switch (key) {
    case "schema":
      if (!isArr) throw new Error("schema must be an array");
      value.forEach(f => { if (!f || !f.id) throw new Error("Each habit needs an 'id'"); });
      break;
    case "entries":
      if (!isArr) throw new Error("entries must be an array");
      value.forEach(e => { if (!e || !e.date) throw new Error("Each entry needs a 'date'"); });
      break;
    case "foodItems":
      if (!isArr) throw new Error("foodItems must be an array");
      value.forEach(f => { if (!f || !f.id) throw new Error("Each food item needs an 'id'"); });
      break;
    case "mealTypes":
      if (!isArr) throw new Error("mealTypes must be an array");
      value.forEach(f => { if (!f || !f.k) throw new Error("Each meal type needs a 'k'"); });
      break;
    case "mealTemplates":
      if (!isArr) throw new Error("mealTemplates must be an array");
      value.forEach(t => { if (!t || !t.id) throw new Error("Each meal template needs an 'id'"); });
      break;
    case "foodEntries":
      if (!isArr) throw new Error("foodEntries must be an array");
      value.forEach(d => {
        if (!d || !d.date || !Array.isArray(d.meals)) throw new Error("Each food log day needs a 'date' and a 'meals' array");
        d.meals.forEach(m => { if (!m || !m.id) throw new Error("Each logged meal needs an 'id'"); });
      });
      break;
    default:
      break; // Meta/theme/foodSettings: any JSON-serialisable value is accepted.
  }
}
// Append-only audit trail — never blocks a save if logging itself fails.
function logChange_(key, value) {
  try {
    const sh = sheet_(SHEET_CHANGELOG, ["timestamp", "key", "recordCount"]);
    const count = Array.isArray(value) ? value.length : 1;
    sh.appendRow([new Date(), key, count]);
  } catch (err) {
    Logger.log("ChangeLog append failed (non-fatal): " + (err && err.message));
  }
}
// Makes one full-spreadsheet Drive copy per calendar day as a restore point,
// independent of Sheets' own version history. Never blocks the request if
// Drive access fails (e.g. missing scope) — backups are best-effort safety,
// not a requirement for the app to function.
function createBackupIfDue_() {
  try {
    const props = PropertiesService.getScriptProperties();
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    if (props.getProperty("lastBackupDate") === today) return;
    const file = DriveApp.getFileById(ss_().getId());
    const parents = file.getParents();
    const folder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
    file.makeCopy(`${file.getName()} — backup ${today}`, folder);
    props.setProperty("lastBackupDate", today);
  } catch (err) {
    Logger.log("Backup skipped (non-fatal): " + (err && err.message));
  }
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
    createBackupIfDue_(); // cheap no-op after the first request of the day
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      return jsonOut_({
        entries: readEntries_(),
        schema:  readSchema_(),
        theme:   readMeta_("theme"),
        // Structured sheets are preferred; pre-migration installs fall back
        // to the old Meta/FoodEntries JSON blobs so nothing breaks before
        // `migrateFoodDataToStructuredSheets_` has been run.
        foodItems:     readFoodItems_()     ?? safeJson_(readMeta_("foodItems"), []),
        foodSettings:  readMeta_("foodSettings"),
        mealTemplates: readMealTemplates_() ?? safeJson_(readMeta_("mealTemplates"), []),
        mealTypes:     readMealTypes_()     ?? safeJson_(readMeta_("mealTypes"), null),
        foodEntries:   readFoodEntries_()   ?? [],
      });
    } finally {
      lock.releaseLock();
    }
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
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      validateWrite_(body.key, body.value); // throws BEFORE any sheet is touched
      switch (body.key) {
        case "schema":        writeSchema_(body.value);        break;
        case "entries":       writeEntries_(body.value);       break;
        case "theme":         writeMeta_("theme", body.value); break;
        case "foodItems":     writeFoodItems_(body.value);     break;
        case "mealTypes":     writeMealTypes_(body.value);     break;
        case "mealTemplates": writeMealTemplates_(body.value); break;
        case "foodEntries":   writeFoodEntries_(body.value);   break;
        default: writeMeta_(body.key, body.value);  break;
      }
      logChange_(body.key, body.value);
    } finally {
      lock.releaseLock();
    }
    return jsonOut_({ ok: true, key: body.key });
  } catch (err) {
    return jsonOut_({ error: String(err && err.message || err) });
  }
}

/* ─────────────── one-shot migration helpers ───────────────
 * Run from the Apps Script editor (Run ▸ select function ▸ Run). Both are
 * safe to re-run and never delete their source sheets/rows.
 */
function migrate_() {
  // v1 single-cell `Data` sheet → v2 tabular Schema/Entries/Meta (unchanged).
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
// v2 single-JSON-cell food data (Meta rows + legacy FoodEntries sheet) → v3
// structured FoodItems/MealTypes/MealTemplates(+Items)/FoodLog(+Items).
// The old Meta rows and the old FoodEntries sheet are left completely
// untouched, so this is non-destructive and safe to re-run.
function migrateFoodDataToStructuredSheets_() {
  const items = safeJson_(readMeta_("foodItems"), []);
  if (Array.isArray(items) && items.length) writeFoodItems_(items);

  const types = safeJson_(readMeta_("mealTypes"), []);
  if (Array.isArray(types) && types.length) writeMealTypes_(types);

  const templates = safeJson_(readMeta_("mealTemplates"), []);
  if (Array.isArray(templates) && templates.length) writeMealTemplates_(templates);

  const oldFoodEntriesSheet = ss_().getSheetByName(SHEET_FOOD_ENTRIES_LEGACY);
  let dayCount = 0;
  if (oldFoodEntriesSheet) {
    const values = oldFoodEntriesSheet.getDataRange().getValues();
    const days = [];
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (!row[0]) continue;
      days.push({ date: String(row[0]), meals: safeJson_(row[1], []) });
    }
    if (days.length) { writeFoodEntries_(days); dayCount = days.length; }
  }
  Logger.log(
    "Food data migration done. items=%s mealTypes=%s templates=%s foodLogDays=%s — old Meta/FoodEntries rows left untouched as a backup.",
    items.length || 0, types.length || 0, templates.length || 0, dayCount
  );
}
