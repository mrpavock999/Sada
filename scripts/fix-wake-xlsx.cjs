// Convert the Wake Time column from raw day-fractions / serials / Date cells
// into unambiguous "HH:MM" text strings, so neither Sheets nor the importer
// can drift the value via timezone-aware Date parsing.
const XLSX = require("xlsx");
const path = "Discipline Tracking V1.xlsx";

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

const wb = XLSX.readFile(path, { cellDates: false });
const ws = wb.Sheets[wb.SheetNames[0]];
const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
const headers = aoa[0];
const wi = headers.findIndex((h) => /wake/i.test(String(h)));
if (wi === -1) {
  console.error("Wake column not found");
  process.exit(1);
}
console.log("Wake column:", headers[wi], "@ index", wi);

let converted = 0;
let blank = 0;
const samples = [];
for (let i = 1; i < aoa.length; i++) {
  const before = aoa[i][wi];
  const after = toHHMM(before);
  if (after === "") {
    blank++;
    aoa[i][wi] = null;
    continue;
  }
  if (samples.length < 6) samples.push(`row ${i}: ${JSON.stringify(before)} → "${after}"`);
  aoa[i][wi] = after;
  converted++;
}
console.log("Converted:", converted, "blank:", blank);
console.log("Samples:\n  " + samples.join("\n  "));

const ws2 = XLSX.utils.aoa_to_sheet(aoa);
if (ws["!cols"]) ws2["!cols"] = ws["!cols"];

// Force the Wake Time column cells to be text-typed ('s') with @ format,
// so Excel/Sheets won't auto-parse them back into time fractions.
const range = XLSX.utils.decode_range(ws2["!ref"]);
for (let r = 1; r <= range.e.r; r++) {
  const addr = XLSX.utils.encode_cell({ r, c: wi });
  const cell = ws2[addr];
  if (!cell) continue;
  if (cell.v === null || cell.v === undefined || cell.v === "") {
    delete ws2[addr];
    continue;
  }
  cell.t = "s";
  cell.v = String(cell.v);
  cell.w = cell.v;
  cell.z = "@";
}

wb.Sheets[wb.SheetNames[0]] = ws2;
XLSX.writeFile(wb, path);
console.log("Saved", path);
