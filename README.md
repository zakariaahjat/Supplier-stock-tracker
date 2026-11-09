# UltraPC — DIP HOSON Stock Tracker

A browser-based stock reconciliation dashboard for UltraPC. It imports an Excel
purchase file for **DIP HOSON** and a warehouse stock file, then uses **exact
serial-number matching** to show how many DIP HOSON units are still in stock.

Everything runs locally in the browser. No server, upload, account, or internet
connection is required after the files are on disk.

---

## What it does

- Imports DIP HOSON purchases and current warehouse stock from `.xlsx` / `.xls` / `.csv`.
- Compares serial numbers exactly, per product.
- Shows **Product Name**, **Purchased Quantity**, **Current Stock**,
  **Remaining DIP HOSON Stock**, and **Matching Serial Numbers**.
- Marks missing, invalid, duplicate, product-mismatched, and supplier-conflicting
  serials as **Unverified** instead of guessing.
- Exports a multi-sheet Excel report.

### Matching rules

1. Only **exact serial matches** from DIP HOSON purchases count as remaining stock.
2. **Purchased Quantity** is the number of unique valid DIP HOSON serial numbers,
   not the value in a quantity column.
3. **Current Stock** is every valid serialized warehouse unit for that product,
   regardless of supplier.
4. A serial only counts when the **product name and serial both match**.
5. Missing, placeholder, duplicate, mismatched, or conflicting serials are excluded
   from all counts and listed as **Unverified**.
6. If a warehouse row explicitly names a supplier other than DIP HOSON, it is never
   counted as DIP HOSON remaining stock. Other-supplier purchase rows are ignored.
7. Products in warehouse stock that do not appear in the DIP HOSON purchase file are
   excluded from the report.
8. Stock from different suppliers is never mixed, and quantities are never estimated.

---

## Run the app

### Option 1 — Local server (recommended)

Requires [Node.js](https://nodejs.org/).

```powershell
cd ultrapc-stock-tracker
node tools/server.js
```

Open <http://127.0.0.1:4173/>.

To use a different port:

```powershell
$env:PORT=8080; node tools/server.js
```

### Option 2 — Open the file directly

Double-click `index.html`, or open it in a modern browser. The app is fully static
and includes its own copy of SheetJS, so it works from `file://` as well.

> If a browser restricts local file access, use Option 1.

---

## Importing files

Each file needs two columns:

| Purpose | Required | Accepted header examples |
| --- | --- | --- |
| Product name | Yes | `Product Name`, `Product`, `Item`, `Description`, `Model`, `SKU` |
| Serial number | Yes | `Serial Number`, `Serial No`, `Serial`, `SN`, `Service Tag`, `IMEI` |
| Supplier | Optional | `Supplier`, `Vendor`, `Distributor`, `Brand` |
| Quantity | Optional | `Quantity`, `Qty`, `Units`, `Purchased Qty`, `Stock` |

- Headers are detected automatically and are case/spacing tolerant. French and
  Arabic header aliases are also recognised.
- Use **one serial number per row**. Multiple serials in one cell are rejected.
- Format serial columns as **Text** in Excel so long numbers are not converted to
  scientific notation.
- The first sheet with a detectable Product + Serial header is used. You can switch
  sheets in the upload card.

Not sure about the format? Use **Download template** in either upload card, or try
**Load demo data** to see the full workflow.

---

## Reading the dashboard

- **KPI cards** — purchased units, warehouse serials, remaining DIP HOSON units,
  match rate, and unverified serials.
- **Reconciliation table** — per-product results with expandable serial details.
- **Serial review** — every excluded serial with its reason and Excel row number.
- **Filters and sorting** — search, status filter, sort, and pagination.
- **Export results** — downloads
  `UltraPC_DIP_HOSON_Reconciliation_<date>.xlsx` with five sheets:
  - `Summary`
  - `Reconciliation`
  - `Matched Serials`
  - `Unverified Serials`
  - `Excluded Data`

Imported data and results are saved in the browser's local storage, so a refresh
does not lose your work. Use **Reset** to clear everything.

---

## Project structure

```
ultrapc-stock-tracker/
├─ index.html                     # Dashboard markup
├─ assets/
│  ├─ css/styles.css              # UltraPC blue/orange design and responsive rules
│  ├─ js/app.js                   # Import, parsing, matching, UI, export
│  └─ vendor/xlsx.full.min.js     # Local SheetJS build
├─ samples/
│  ├─ DIP_HOSON_Purchases_Sample.xlsx
│  └─ Warehouse_Stock_Sample.xlsx
├─ tools/
│  ├─ generate-samples.js         # Regenerates the sample workbooks
│  └─ server.js                   # Small static server (port 4173)
└─ README.md
```

Regenerate the sample workbooks with:

```powershell
node tools/generate-samples.js
```

---

## Troubleshooting

| Problem | Fix |
| --- | --- |
| “Could not find Product Name and Serial Number columns” | Rename the headers to `Product Name` and `Serial Number`, or use the template. |
| Serials show as `1.23457E+15` | Format the serial column as Text in Excel, then re-import. |
| Everything is Unverified | Check that both files use the same product names and that serials match exactly. |
| A product is missing from the report | It is not in the DIP HOSON purchase file, so it is excluded by design. |
| A known unit is not counted | Open **Serial review**; the serial may be duplicated, a placeholder, or assigned to another supplier/product. |

---

## Tech

- Static HTML, CSS, and vanilla JavaScript.
- [SheetJS](https://sheetjs.com/) (`xlsx.full.min.js`) for Excel parsing and export.
- Browser `localStorage` for persistence.

No data leaves the machine.
