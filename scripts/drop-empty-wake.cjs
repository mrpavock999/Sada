// Drop rows where Wake Time is empty.
const XLSX = require("xlsx");
const path = "Discipline Tracking V1.xlsx";

const wb = XLSX.readFile(path, { cellDates: false });
const ws = wb.Sheets[wb.SheetNames[0]];
const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
const headers = aoa[0];
const wi = headers.findIndex((h) => /wake/i.test(String(h)));
if (wi === -1) {
  console.error("Wake column not found");
  process.exit(1);
}

const isEmpty = (v) => v === null || v === undefined || String(v).trim() === "";

const kept = [headers];
let dropped = 0;
for (let i = 1; i < aoa.length; i++) {
  if (isEmpty(aoa[i][wi])) { dropped++; continue; }
  kept.push(aoa[i]);
}
console.log("Dropped:", dropped, "Kept:", kept.length - 1);

const ws2 = XLSX.utils.aoa_to_sheet(kept);
if (ws["!cols"]) ws2["!cols"] = ws["!cols"];

// Preserve text typing for the Wake Time column.
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
