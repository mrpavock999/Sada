// Apply the same normalisations to "Discipline Tracking V1 2.xlsx" that we did
// to the original file: invert junkFood + rename header, convert wake column
// to "HH:MM" text, and drop rows with empty wake.
const XLSX = require("xlsx");
const path = "Discipline Tracking V1 2.xlsx";

function toHHMM(v) {
  if (v === null || v === undefined || v === "") return "";
  if (v instanceof Date && !isNaN(v)) {
    return `${String(v.getHours()).padStart(2, "0")}:${String(v.getMinutes()).padStart(2, "0")}`;
  }
  if (typeof v === "string") {
    const m = v.match(/^(\d{1,2}):(\d{2})/);
    if (m) return `${String(Math.min(23, Number(m[1]))).padStart(2, "0")}:${m[2]}`;
    const num = Number(v);
    if (isFinite(num)) return toHHMM(num);
    return "";
  }
  if (typeof v !== "number" || !isFinite(v)) return "";
  const frac = v - Math.floor(v);
  const total = Math.round(frac * 24 * 60);
  const h = Math.floor(total / 60) % 24;
  const m = total % 60;
  return `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}`;
}

const isEmpty = (v) => v === null || v === undefined || String(v).trim() === "";

const wb = XLSX.readFile(path, { cellDates: false });
const ws = wb.Sheets[wb.SheetNames[0]];
const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
const headers = aoa[0];
console.log("Headers:", headers);

const wi = headers.findIndex((h) => /wake/i.test(String(h)));
const ji = headers.findIndex((h) => /junk/i.test(String(h)));
if (wi === -1) { console.error("Wake column not found"); process.exit(1); }

// 1. Invert junk food (only if header isn't already the inverted name)
let flipped = 0;
if (ji !== -1) {
  const alreadyInverted = /no\s*junk/i.test(String(headers[ji]));
  if (!alreadyInverted) {
    headers[ji] = "No Junk Food (1=Yes)";
    for (let i = 1; i < aoa.length; i++) {
      const v = aoa[i][ji];
      if (v === undefined || v === null || v === "") continue;
      let bool;
      if (typeof v === "boolean") bool = v;
      else if (typeof v === "number") bool = v !== 0;
      else {
        const s = String(v).trim().toLowerCase();
        bool = s === "1" || s === "true" || s === "yes" || s === "y";
      }
      aoa[i][ji] = bool ? 0 : 1;
      flipped++;
    }
    console.log(`Junk food: inverted ${flipped} rows, header → "No Junk Food (1=Yes)"`);
  } else {
    console.log("Junk food: header already inverted, skipping");
  }
}

// 2. Convert wake to HH:MM text
let waked = 0;
for (let i = 1; i < aoa.length; i++) {
  const before = aoa[i][wi];
  const after = toHHMM(before);
  aoa[i][wi] = after === "" ? null : after;
  if (after) waked++;
}
console.log(`Wake: normalised ${waked} cells to HH:MM text`);

// 3. Drop rows where wake is empty
const kept = [headers];
let dropped = 0;
for (let i = 1; i < aoa.length; i++) {
  if (isEmpty(aoa[i][wi])) { dropped++; continue; }
  kept.push(aoa[i]);
}
console.log(`Dropped ${dropped} rows with empty wake. Kept ${kept.length - 1} data rows.`);

const ws2 = XLSX.utils.aoa_to_sheet(kept);
if (ws["!cols"]) ws2["!cols"] = ws["!cols"];

// Force wake column cells to text
const range = XLSX.utils.decode_range(ws2["!ref"]);
for (let r = 1; r <= range.e.r; r++) {
  const addr = XLSX.utils.encode_cell({ r, c: wi });
  const cell = ws2[addr];
  if (!cell) continue;
  cell.t = "s";
  cell.v = String(cell.v);
  cell.w = cell.v;
  cell.z = "@";
}

wb.Sheets[wb.SheetNames[0]] = ws2;
XLSX.writeFile(wb, path);
console.log("Saved", path);
