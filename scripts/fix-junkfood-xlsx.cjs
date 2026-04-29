const XLSX = require("xlsx");
const path = "Discipline Tracking V1.xlsx";
const wb = XLSX.readFile(path, { cellDates: false });
const ws = wb.Sheets[wb.SheetNames[0]];
const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });
const headers = aoa[0];
const jfIdx = headers.findIndex((h) => /junk/i.test(String(h)));
if (jfIdx === -1) {
  console.error("Junk Food column not found");
  process.exit(1);
}
console.log("Junk Food column:", headers[jfIdx], "@ index", jfIdx);
// Rename header to make app semantics explicit, and invert values.
headers[jfIdx] = "No Junk Food (1=Yes)";
let flipped = 0;
for (let i = 1; i < aoa.length; i++) {
  const v = aoa[i][jfIdx];
  if (v === undefined || v === null || v === "") continue;
  let bool;
  if (typeof v === "boolean") bool = v;
  else if (typeof v === "number") bool = v !== 0;
  else {
    const s = String(v).trim().toLowerCase();
    bool = s === "1" || s === "true" || s === "yes" || s === "y";
  }
  aoa[i][jfIdx] = bool ? 0 : 1; // invert: ate junk -> 0 (=did NOT keep clean), avoided junk -> 1
  flipped++;
}
console.log("Flipped", flipped, "rows");
const ws2 = XLSX.utils.aoa_to_sheet(aoa);
if (ws["!cols"]) ws2["!cols"] = ws["!cols"];
wb.Sheets[wb.SheetNames[0]] = ws2;
XLSX.writeFile(wb, path);
console.log("Saved", path);
