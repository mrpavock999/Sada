const XLSX = require("xlsx");
const wb = XLSX.readFile("Discipline Tracking V1.xlsx", { cellDates: false });
const ws = wb.Sheets[wb.SheetNames[0]];
const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });
console.log("headers:", aoa[0]);
const wi = aoa[0].findIndex((h) => /wake/i.test(String(h)));
console.log("wake col index:", wi);
for (let i = 1; i <= 8 && i < aoa.length; i++) {
  console.log(i, JSON.stringify(aoa[i][0]), "wake=", JSON.stringify(aoa[i][wi]));
}
console.log("---raw cells (first 5 data rows)---");
for (let r = 1; r <= 5; r++) {
  const addr = XLSX.utils.encode_cell({ r, c: wi });
  console.log(addr, ws[addr]);
}
console.log("---unique types---");
const types = {};
for (let i = 1; i < aoa.length; i++) {
  const v = aoa[i][wi];
  const k = v === null || v === undefined || v === "" ? "empty" : typeof v;
  types[k] = (types[k] || 0) + 1;
}
console.log(types);
