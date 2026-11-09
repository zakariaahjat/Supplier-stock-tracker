const fs = require("fs");
const path = require("path");
const XLSX = require(path.join(__dirname, "..", "assets", "vendor", "xlsx.full.min.js"));

const purchaseRows = [
  ["Product Name", "Serial Number", "Supplier", "Purchase Date", "Invoice Number", "Notes"],
  ["UltraBook Pro 14", "UH14-2601-0001", "DIP HOSON", "2026-01-12", "DH-1042", ""],
  ["UltraBook Pro 14", "UH14-2601-0002", "DIP HOSON", "2026-01-12", "DH-1042", ""],
  ["UltraBook Pro 14", "UH14-2601-0003", "DIP HOSON", "2026-01-12", "DH-1042", ""],
  ["UltraBook Pro 14", "UH14-2601-0004", "DIP HOSON", "2026-01-12", "DH-1042", ""],
  ["UltraBook Pro 14", "UH14-2601-0005", "DIP HOSON", "2026-01-12", "DH-1042", ""],
  ["UltraBook Pro 14", "UH14-2601-0006", "DIP HOSON", "2026-01-12", "DH-1042", ""],
  ["UltraBook Air 13", "UA13-2602-0001", "DIP HOSON", "2026-02-03", "DH-1088", ""],
  ["UltraBook Air 13", "UA13-2602-0002", "DIP HOSON", "2026-02-03", "DH-1088", ""],
  ["UltraBook Air 13", "UA13-2602-0003", "DIP HOSON", "2026-02-03", "DH-1088", ""],
  ["UltraBook Air 13", "UA13-2602-0004", "DIP HOSON", "2026-02-03", "DH-1088", ""],
  ["UltraStation Mini", "USM-2603-0001", "DIP HOSON", "2026-03-18", "DH-1150", ""],
  ["UltraStation Mini", "USM-2603-0002", "DIP HOSON", "2026-03-18", "DH-1150", ""],
  ["UltraStation Mini", "USM-2603-0003", "DIP HOSON", "2026-03-18", "DH-1150", ""],
  ["UltraView 27", "UV27-2604-0001", "DIP HOSON", "2026-04-07", "DH-1194", ""],
  ["UltraView 27", "UV27-2604-0002", "DIP HOSON", "2026-04-07", "DH-1194", ""],
  ["UltraView 27", "UV27-2604-0003", "DIP HOSON", "2026-04-07", "DH-1194", ""],
  ["UltraView 27", "UV27-2604-0004", "DIP HOSON", "2026-04-07", "DH-1194", ""],
  ["UltraView 27", "UV27-2604-0005", "DIP HOSON", "2026-04-07", "DH-1194", ""],
  ["UltraGear K87", "K87-2605-0001", "DIP HOSON", "2026-05-21", "DH-1240", ""],
  ["UltraGear K87", "K87-2605-0002", "DIP HOSON", "2026-05-21", "DH-1240", ""],
  ["UltraGear K87", "K87-2605-0003", "DIP HOSON", "2026-05-21", "DH-1240", ""],
  ["UltraGear K87", "K87-2605-0004", "DIP HOSON", "2026-05-21", "DH-1240", ""],
  ["UltraGear K87", "K87-2605-0005", "DIP HOSON", "2026-05-21", "DH-1240", ""],
  ["UltraGear K87", "K87-2605-0006", "DIP HOSON", "2026-05-21", "DH-1240", ""],
  ["UltraGear K87", "K87-2605-0007", "DIP HOSON", "2026-05-21", "DH-1240", ""],
  ["UltraGear K87", "K87-2605-0008", "DIP HOSON", "2026-05-21", "DH-1240", ""],
  ["UltraBook Air 13", "", "DIP HOSON", "2026-02-03", "DH-1088", "Missing serial example"],
  ["UltraBook Air 13", "N/A", "DIP HOSON", "2026-02-03", "DH-1088", "Invalid serial example"],
  ["UltraBook Air 13", "UA13-2602-0002", "DIP HOSON", "2026-02-03", "DH-1088", "Duplicate serial example"],
  ["UltraBook Pro 14", "UH14-2601-0001", "TechSource", "2026-01-12", "TS-991", "Other supplier - excluded"]
];

const stockRows = [
  ["Product Name", "Serial Number", "Supplier", "Location", "Notes"],
  ["UltraBook Pro 14", "UH14-2601-0001", "DIP HOSON", "A1-01", ""],
  ["UltraBook Pro 14", "UH14-2601-0002", "TechSource", "A1-02", "Supplier conflict example"],
  ["UltraBook Pro 14", "UH14-2601-0004", "DIP HOSON", "A1-03", ""],
  ["UltraBook Pro 14", "UH14-2601-0006", "DIP HOSON", "A1-04", ""],
  ["UltraBook Pro 14", "OTH-PRO-001", "TechSource", "A1-05", "Other supplier stock"],
  ["UltraBook Air 13", "UA13-2602-0001", "DIP HOSON", "A2-01", ""],
  ["UltraBook Air 13", "UA13-2602-0003", "DIP HOSON", "A2-02", ""],
  ["UltraBook Air 13", "OTH-AIR-001", "TechSource", "A2-03", "Other supplier stock"],
  ["UltraBook Air 13", "UH14-2601-0005", "DIP HOSON", "A2-04", "Product mismatch example"],
  ["UltraStation Mini", "OTH-MINI-001", "TechSource", "B1-01", "Sold out for DIP HOSON"],
  ["UltraView 27", "UV27-2604-0001", "DIP HOSON", "C1-01", ""],
  ["UltraView 27", "UV27-2604-0002", "DIP HOSON", "C1-02", ""],
  ["UltraView 27", "UV27-2604-0003", "DIP HOSON", "C1-03", ""],
  ["UltraView 27", "UV27-2604-0004", "DIP HOSON", "C1-04", ""],
  ["UltraView 27", "UV27-2604-0005", "DIP HOSON", "C1-05", ""],
  ["UltraView 27", "N/A", "DIP HOSON", "C1-06", "Invalid serial example"],
  ["UltraGear K87", "K87-2605-0001", "DIP HOSON", "D1-01", ""],
  ["UltraGear K87", "K87-2605-0002", "DIP HOSON", "D1-02", ""],
  ["UltraGear K87", "K87-2605-0003", "DIP HOSON", "D1-03", ""],
  ["UltraGear K87", "K87-2605-0001", "DIP HOSON", "D1-04", "Duplicate serial example"],
  ["UltraGear K87", "", "DIP HOSON", "D1-05", "Missing serial example"],
  ["UltraDock USB-C Hub", "UD-2606-0001", "DIP HOSON", "E1-01", "Not in purchase file"],
  ["UltraDock USB-C Hub", "UD-2606-0002", "DIP HOSON", "E1-02", "Not in purchase file"],
  ["UltraDock USB-C Hub", "UD-2606-0003", "DIP HOSON", "E1-03", "Not in purchase file"]
];

function instructions(title, extra) {
  return [
    [title],
    [],
    ["How to use"],
    ["1", "Keep the header row and replace the data rows with your real data."],
    ["2", "Use one row per serialized unit. Never put multiple serials in one cell."],
    ["3", "Format the Serial Number column as Text before pasting long serials."],
    ["4", "Do not merge cells. Keep product names consistent between both workbooks."],
    ["5", extra],
    ["6", "Missing, invalid, duplicate, or conflicting serials are marked Unverified and excluded from every count."]
  ];
}

function writeSample(fileName, title, rows, sheetName, extra) {
  const workbook = XLSX.utils.book_new();
  const dataSheet = XLSX.utils.aoa_to_sheet(rows);
  dataSheet["!cols"] = rows[0].map((header) => ({ wch: header === "Notes" ? 34 : header === "Serial Number" ? 28 : 22 }));
  XLSX.utils.book_append_sheet(workbook, dataSheet, sheetName);
  const infoSheet = XLSX.utils.aoa_to_sheet(instructions(title, extra));
  infoSheet["!cols"] = [{ wch: 7 }, { wch: 96 }];
  XLSX.utils.book_append_sheet(workbook, infoSheet, "Instructions");
  const output = XLSX.write(workbook, { bookType: "xlsx", type: "buffer", compression: true });
  fs.writeFileSync(path.join(__dirname, "..", "samples", fileName), output);
}

writeSample(
  "DIP_HOSON_Purchases_Sample.xlsx",
  "UltraPC — DIP HOSON Purchases Sample",
  purchaseRows,
  "DIP HOSON Purchases",
  "If a Supplier column is present, only rows matching DIP HOSON are used. Other supplier rows are excluded."
);
writeSample(
  "Warehouse_Stock_Sample.xlsx",
  "UltraPC — Warehouse Stock Sample",
  stockRows,
  "Warehouse Stock",
  "The Supplier column is optional. If a matched serial explicitly names another supplier, it is flagged and not counted."
);
console.log("Sample workbooks created in samples/");
