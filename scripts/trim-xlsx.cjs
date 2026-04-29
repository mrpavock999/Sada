const XLSX = require("xlsx");
const path = "Discipline Tracking V1.xlsx";
const wb = XLSX.readFile(path, { cellDates: false });
const ws = wb.Sheets[wb.SheetNames[0]];
const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });
const headers = aoa[0];
const cutoffSerial = 46139; // 27 Apr 2026
const cutoffDate = new Date("2026-04-27T23:59:59Z");
const dateCol = headers.findIndex((h) => /date/i.test(String(h)));
const kept = [
  headers,
  ...aoa.slice(1).filter((r) => {
    const v = r[dateCol];
    if (v === undefined || v === null || v === "") return false;
    if (typeof v === "number") return v <= cutoffSerial;
    const d = new Date(v);
    return !isNaN(d) && d <= cutoffDate;
  }),
];
console.log("Original rows:", aoa.length - 1, "→ Kept rows:", kept.length - 1);
const ws2 = XLSX.utils.aoa_to_sheet(kept);
if (ws["!cols"]) ws2["!cols"] = ws["!cols"];
wb.Sheets[wb.SheetNames[0]] = ws2;
XLSX.writeFile(wb, path);
console.log("Saved", path);
