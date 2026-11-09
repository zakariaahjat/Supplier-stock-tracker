/* UltraPC DIP HOSON Stock Tracker */
(() => {
  "use strict";

  const STORAGE_KEY = "ultrapc-dip-hoson-stock-v1";
  const STORAGE_VERSION = 1;
  const MAX_FILE_SIZE = 25 * 1024 * 1024;
  const PAGE_SIZE = 8;
  const REVIEW_LIMIT = 100;
  const DIP_HOSON_LABEL = "DIP HOSON";
  const UNSPECIFIED_KEY = "__unspecified_product__";
  const UNSPECIFIED_NAME = "Unspecified product";
  const ALLOWED_EXTENSIONS = new Set(["xlsx", "xls", "xlsm", "csv", "ods"]);
  const XLSX_LIB = window.XLSX;
  const numberFormat = new Intl.NumberFormat("en-US");
  const textCollator = new Intl.Collator(undefined, { numeric: true, sensitivity: "base" });

  const INVALID_SERIAL_VALUES = new Set([
    "NA", "NAN", "NONE", "NULL", "UNKNOWN", "UNK", "TBD", "TBA", "MISSING",
    "NOSERIAL", "NOSERIALNUMBER", "SERIAL", "SERIALNUMBER", "SN", "SERIE",
    "SERIES", "0", "00", "000", "0000", "X", "XX", "XXX", "TEMP", "TEST",
    "INVALID", "NOTAVAILABLE", "SANS", "SERIEABSENTE"
  ]);

  const HEADER_ALIASES = {
    product: [
      "productname", "product", "itemname", "item", "description", "designation",
      "produit", "article", "libelle", "model", "modele", "name", "nom", "sku",
      "reference", "productdescription", "producttitle", "اسم المنتج", "المنتج",
      "السلعة", "الوصف"
    ].map(headerKey),
    serial: [
      "serialnumber", "serialno", "serial", "serialnum", "serialnumber1", "sn",
      "snnumber", "serialtag", "servicetag", "imei", "numerodeserie", "nserie",
      "numserie", "numeroserie", "noserie", "serienumber", "serie", "الرقم التسلسلي",
      "رقم تسلسلي", "السيريال", "سيريال", "التسلسلي", "الرقم التسلسلى"
    ].map(headerKey),
    supplier: [
      "supplier", "suppliername", "vendor", "vendorname", "fournisseur",
      "distributeur", "distributor", "brand", "marque", "societe", "company",
      "المورد", "الموزع", "المزود"
    ].map(headerKey),
    quantity: [
      "quantity", "qty", "qte", "quantite", "quantityreceived", "units", "unit",
      "purchasedqty", "purchasedquantity", "stock", "currentstock", "onhand",
      "qtyinstock", "الكمية", "كمية", "العدد"
    ].map(headerKey)
  };

  const state = {
    purchases: null,
    stock: null,
    results: [],
    unverified: [],
    excludedStockProducts: [],
    summary: emptySummary(),
    filters: {
      search: "",
      status: "all",
      sort: "remaining-desc",
      page: 1
    },
    reviewFilters: {
      search: "",
      source: "all"
    },
    detailsKey: null,
    loading: { purchase: false, stock: false },
    pendingConfirm: null,
    persistenceDisabled: false,
    restored: false
  };

  const els = {};
  let toastCounter = 0;

  function emptySummary() {
    return {
      purchasedUnits: 0,
      warehouseSerials: 0,
      remainingUnits: 0,
      matchRate: 0,
      matchRatePct: 0,
      unverifiedCount: 0,
      excludedOtherSupplierRows: 0,
      stockOnlyProducts: 0,
      stockOnlySerials: 0,
      soldOutProducts: 0,
      matchedProducts: 0,
      unverifiedProducts: 0,
      purchaseProducts: 0,
      hasPurchases: false,
      hasStock: false
    };
  }

  function cacheElements() {
    const ids = [
      "todayChip", "menuBtn", "sidebar", "mobileOverlay", "topResetBtn", "topExportBtn", "demoBtn",
      "purchaseInput", "purchaseDropzone", "purchaseSummary", "purchaseFileName", "purchaseFileMeta",
      "purchaseMeta", "purchaseReplaceBtn", "purchaseRemoveBtn", "purchaseTemplateBtn", "purchaseLoading",
      "purchaseCard", "stockInput", "stockDropzone", "stockSummary", "stockFileName", "stockFileMeta",
      "stockMeta", "stockReplaceBtn", "stockRemoveBtn", "stockTemplateBtn", "stockLoading", "stockCard",
      "statPurchased", "statPurchasedSub", "statStock", "statStockSub", "statRemaining",
      "statRemainingSub", "statRate", "statRateBar", "statRateSub", "statUnverified",
      "statUnverifiedSub", "unverifiedStatCard", "resultCount", "productSearch", "statusFilter",
      "sortSelect", "tableExportBtn", "resultsBody", "resultsEmpty", "resultsTable", "paginationInfo",
      "pagination", "unverifiedCountPill", "reviewSearch", "reviewSourceFilter", "guardBanner",
      "reviewBody", "reviewEmpty", "reviewFooter", "reviewTable", "navUnverifiedBadge",
      "detailsDialog", "detailsContent", "confirmDialog", "confirmTitle", "confirmMessage",
      "confirmActionBtn", "confirmCancelBtn", "confirmIcon", "toastStack", "heroPurchased",
      "heroStock", "heroMatched"
    ];
    ids.forEach((id) => { els[id] = document.getElementById(id); });
  }

  function cleanText(value) {
    if (value === null || value === undefined) return "";
    if (value instanceof Date) {
      return Number.isNaN(value.getTime()) ? "" : value.toISOString().slice(0, 10);
    }
    return String(value)
      .replace(/\u00a0/g, " ")
      .replace(/[\u200b-\u200d\ufeff]/g, "")
      .replace(/\s+/g, " ")
      .trim();
  }

  function stripAccents(value) {
    return value.normalize("NFKD").replace(/[\u0300-\u036f]/g, "");
  }

  function headerKey(value) {
    return stripAccents(cleanText(value))
      .toLowerCase()
      .replace(/[^\p{L}\p{N}]+/gu, "");
  }

  function productKey(value) {
    const text = stripAccents(cleanText(value)).toUpperCase();
    return text || UNSPECIFIED_KEY;
  }

  function serialKey(value) {
    return stripAccents(cleanText(value))
      .replace(/^'+/, "")
      .replace(/[\u2010-\u2015]/g, "-")
      .toUpperCase();
  }

  function displayProduct(value) {
    return cleanText(value) || UNSPECIFIED_NAME;
  }

  function analyzeSupplier(value) {
    const text = cleanText(value);
    const key = headerKey(text);
    const compact = key.replace(/[^a-z0-9]/g, "");
    const isDip = compact.includes("diphoson") || compact.includes("diposon") || compact.includes("diphosson");
    return { text, key, isDip, isMissing: !text };
  }

  function validateSerial(value) {
    const text = cleanText(value);
    if (!text) return { valid: false, reason: "Serial number is missing." };
    if (/\d(?:\.\d+)?e[+-]?\d+/i.test(text)) {
      return { valid: false, reason: "Serial number appears in scientific notation; format the Excel column as text." };
    }
    if (/[,\n;|]/.test(text)) {
      return { valid: false, reason: "Multiple serial numbers found in one cell; use one serial per row." };
    }
    if (text.length < 3) return { valid: false, reason: "Serial number is too short to verify." };
    if (text.length > 120) return { valid: false, reason: "Serial number is too long to verify." };
    const compact = text.toUpperCase().replace(/[^A-Z0-9]/g, "");
    if (!compact) return { valid: false, reason: "Serial number contains no letters or digits." };
    if (INVALID_SERIAL_VALUES.has(compact)) {
      return { valid: false, reason: "Serial number is a placeholder, not a real identifier." };
    }
    if (/^[0\-\/._ ]+$/.test(text)) {
      return { valid: false, reason: "Serial number contains only zeros or punctuation." };
    }
    return { valid: true, reason: "" };
  }

  function findColumn(cells, aliases, used) {
    const normalized = cells.map(headerKey);
    for (const alias of aliases) {
      const index = normalized.findIndex((value, idx) => !used.has(idx) && value === alias);
      if (index !== -1) {
        used.add(index);
        return index;
      }
    }
    for (const alias of aliases) {
      if (alias.length < 4) continue;
      const index = normalized.findIndex((value, idx) => !used.has(idx) && value.includes(alias));
      if (index !== -1) {
        used.add(index);
        return index;
      }
    }
    return -1;
  }

  function detectHeaderRow(rows) {
    let best = null;
    const maxRows = Math.min(rows.length, 40);
    for (let rowIndex = 0; rowIndex < maxRows; rowIndex += 1) {
      const cells = (rows[rowIndex] || []).map(cleanText);
      const nonEmpty = cells.filter(Boolean).length;
      if (nonEmpty < 2) continue;
      const used = new Set();
      const mapping = {
        serial: findColumn(cells, HEADER_ALIASES.serial, used),
        product: findColumn(cells, HEADER_ALIASES.product, used),
        supplier: findColumn(cells, HEADER_ALIASES.supplier, used),
        quantity: findColumn(cells, HEADER_ALIASES.quantity, used)
      };
      let score = 0;
      if (mapping.serial !== -1) score += 10;
      if (mapping.product !== -1) score += 8;
      if (mapping.supplier !== -1) score += 2;
      if (mapping.quantity !== -1) score += 1;
      if (mapping.serial !== -1 && mapping.product !== -1) score += 12;
      score += Math.min(nonEmpty, 8) / 10;
      if (!best || score > best.score) {
        best = { rowIndex, score, mapping, headers: cells };
      }
    }
    if (!best || best.mapping.serial === -1 || best.mapping.product === -1) return null;
    return best;
  }

  function getSheetCandidates(workbook) {
    const candidates = [];
    for (const name of workbook.SheetNames) {
      const sheet = workbook.Sheets[name];
      if (!sheet) continue;
      const rows = XLSX_LIB.utils.sheet_to_json(sheet, {
        header: 1,
        raw: false,
        defval: "",
        blankrows: true
      });
      if (!rows.length) continue;
      const header = detectHeaderRow(rows);
      const dataRows = header ? Math.max(0, rows.length - header.rowIndex - 1) : 0;
      const score = header ? header.score + Math.min(dataRows, 500) / 1000 : -1;
      candidates.push({ name, rows, header, dataRows, score });
    }
    candidates.sort((a, b) => b.score - a.score);
    return candidates;
  }

  async function parseWorkbookFile(file, kind) {
    const extension = (file.name.split(".").pop() || "").toLowerCase();
    if (!ALLOWED_EXTENSIONS.has(extension)) {
      throw new Error("Unsupported file type. Use .xlsx, .xls, .xlsm, .csv, or .ods.");
    }
    if (file.size > MAX_FILE_SIZE) {
      throw new Error("File is larger than 25 MB. Split the workbook and try again.");
    }
    const buffer = await file.arrayBuffer();
    let workbook;
    try {
      workbook = XLSX_LIB.read(buffer, { type: "array", cellDates: false, cellText: true });
    } catch (error) {
      throw new Error("This workbook could not be read. Check that it is a valid Excel or CSV file.");
    }
    const candidates = getSheetCandidates(workbook);
    if (!candidates.length) throw new Error("The workbook is empty.");
    const candidate = candidates.find((item) => item.header);
    if (!candidate) {
      throw new Error("Could not find Product Name and Serial Number columns. Rename the headers or use the template.");
    }
    return buildDatasetFromSheet(kind, file.name, candidate);
  }

  function buildDatasetFromSheet(kind, fileName, sheetCandidate) {
    const { rows, header, name } = sheetCandidate;
    const processed = processRawRows(kind, rows, header);
    finalizeRecordStatuses(processed.records);
    const validRecords = processed.records.filter((record) => record.status === "valid");
    const unverifiedRecords = processed.records.filter((record) => record.status === "unverified");
    const uniqueValid = new Set(validRecords.map((record) => `${record.productKey}|${record.serialKey}`));
    const supplierConflictRows = validRecords.filter((record) => record.supplierConflict).length;
    return {
      kind,
      fileName,
      sheetName: name,
      importedAt: new Date().toISOString(),
      headerRow: header.rowIndex + 1,
      headers: header.headers,
      mapping: header.mapping,
      records: processed.records,
      dataRows: processed.dataRows,
      blankRows: processed.blankRows,
      ignoredOtherSupplierRows: processed.ignoredOtherSupplierRows,
      validCount: validRecords.length,
      unverifiedCount: unverifiedRecords.length,
      uniqueSerialCount: uniqueValid.size,
      supplierConflictRows,
      supplierColumnPresent: header.mapping.supplier !== -1
    };
  }

  function processRawRows(kind, rows, header) {
    const { mapping, rowIndex } = header;
    const records = [];
    let ignoredOtherSupplierRows = 0;
    let blankRows = 0;
    let dataRows = 0;

    for (let index = rowIndex + 1; index < rows.length; index += 1) {
      const row = rows[index] || [];
      const productRaw = cleanText(row[mapping.product]);
      const serialRaw = cleanText(row[mapping.serial]);
      const supplierRaw = mapping.supplier !== -1 ? cleanText(row[mapping.supplier]) : "";
      const quantityRaw = mapping.quantity !== -1 ? cleanText(row[mapping.quantity]) : "";
      if (!productRaw && !serialRaw && !supplierRaw && !quantityRaw) {
        blankRows += 1;
        continue;
      }
      dataRows += 1;
      const supplier = analyzeSupplier(supplierRaw);

      if (kind === "purchase" && mapping.supplier !== -1 && supplierRaw && !supplier.isDip) {
        ignoredOtherSupplierRows += 1;
        continue;
      }

      const reasons = [];
      let supplierConflict = false;
      if (kind === "purchase" && mapping.supplier !== -1 && !supplierRaw) {
        reasons.push("Supplier is missing; DIP HOSON ownership cannot be verified.");
      }
      if (!productRaw) reasons.push("Product name is missing.");
      const serialValidation = validateSerial(serialRaw);
      if (!serialValidation.valid) reasons.push(serialValidation.reason);
      if (kind === "stock" && mapping.supplier !== -1 && supplierRaw && !supplier.isDip) {
        supplierConflict = true;
      }

      records.push({
        id: `${kind}-${index + 1}-${records.length}`,
        source: kind,
        row: index + 1,
        productName: displayProduct(productRaw),
        productKey: productKey(productRaw),
        serial: serialRaw,
        serialKey: serialKey(serialRaw),
        serialDisplay: serialRaw || "(missing)",
        supplier: supplierRaw,
        supplierKey: supplier.key,
        supplierConflict,
        status: reasons.length ? "unverified" : "valid",
        reasons,
        reason: reasons.join(" ")
      });
    }

    return { records, ignoredOtherSupplierRows, blankRows, dataRows };
  }

  function finalizeRecordStatuses(records) {
    const validRecords = records.filter((record) => record.status === "valid");
    const serialProducts = new Map();
    validRecords.forEach((record) => {
      if (!serialProducts.has(record.serialKey)) serialProducts.set(record.serialKey, new Set());
      serialProducts.get(record.serialKey).add(record.productKey);
    });
    const conflictingSerials = new Set();
    serialProducts.forEach((productSet, key) => {
      if (productSet.size > 1) conflictingSerials.add(key);
    });
    records.forEach((record) => {
      if (record.status === "valid" && conflictingSerials.has(record.serialKey)) {
        record.status = "unverified";
        record.reasons.push("Serial number is linked to more than one product.");
      }
    });

    const seen = new Map();
    records.forEach((record) => {
      if (record.status !== "valid") return;
      const key = `${record.productKey}|${record.serialKey}`;
      if (seen.has(key)) {
        const first = seen.get(key);
        record.status = "unverified";
        record.reasons.push(`Duplicate serial number in this file (first seen on row ${first.row}).`);
      } else {
        seen.set(key, record);
      }
    });
    records.forEach((record) => { record.reason = record.reasons.join(" "); });
  }

  function makeUnverifiedEntry(source, record, fileName) {
    return {
      id: `${source}-${record.id}`,
      source,
      sourceLabel: source === "purchase" ? "DIP HOSON purchases" : "Warehouse stock",
      productKey: record.productKey,
      productName: record.productName,
      serial: record.serial || "(missing)",
      serialRaw: record.serial,
      reason: record.reason,
      row: record.row,
      file: fileName,
      supplier: record.supplier
    };
  }

  function computeReconciliation() {
    const purchases = state.purchases;
    const stock = state.stock;
    const hasPurchases = Boolean(purchases);
    const hasStock = Boolean(stock);
    const results = [];
    const unverified = [];

    const purchaseRecords = purchases?.records || [];
    const stockRecords = stock?.records || [];
    const validPurchaseRecords = purchaseRecords.filter((record) => record.status === "valid");
    const validStockRecords = stockRecords.filter((record) => record.status === "valid");

    purchaseRecords.forEach((record) => {
      if (record.status === "unverified") unverified.push(makeUnverifiedEntry("purchase", record, purchases.fileName));
    });
    stockRecords.forEach((record) => {
      if (record.status === "unverified") unverified.push(makeUnverifiedEntry("stock", record, stock.fileName));
    });

    const purchaseByProduct = new Map();
    purchaseRecords.forEach((record) => {
      if (!purchaseByProduct.has(record.productKey)) {
        purchaseByProduct.set(record.productKey, { productName: record.productName, valid: [], unverified: [] });
      }
      const entry = purchaseByProduct.get(record.productKey);
      if (record.status === "valid") entry.valid.push(record);
      else entry.unverified.push(record);
    });

    const stockByProduct = new Map();
    stockRecords.forEach((record) => {
      if (!stockByProduct.has(record.productKey)) {
        stockByProduct.set(record.productKey, { productName: record.productName, valid: [], unverified: [] });
      }
      const entry = stockByProduct.get(record.productKey);
      if (record.status === "valid") entry.valid.push(record);
      else entry.unverified.push(record);
    });

    const stockBySerial = new Map();
    validStockRecords.forEach((record) => {
      if (!stockBySerial.has(record.serialKey)) stockBySerial.set(record.serialKey, record);
    });

    purchaseByProduct.forEach((purchaseEntry, key) => {
      const stockEntry = stockByProduct.get(key);
      const purchaseValid = purchaseEntry.valid;
      const stockValid = stockEntry?.valid || [];
      const purchaseSerialMap = new Map();
      const stockSerialMap = new Map();
      purchaseValid.forEach((record) => purchaseSerialMap.set(record.serialKey, record));
      stockValid.forEach((record) => stockSerialMap.set(record.serialKey, record));

      const matchingSerials = [];
      const matchConflicts = [];

      purchaseSerialMap.forEach((purchaseRecord) => {
        const stockRecord = stockSerialMap.get(purchaseRecord.serialKey);
        if (stockRecord) {
          if (stockRecord.supplierConflict) {
            matchConflicts.push({
              id: `matching-supplier-${key}-${purchaseRecord.serialKey}-${stockRecord.row}`,
              source: "matching",
              sourceLabel: "Supplier conflict",
              productKey: key,
              productName: purchaseEntry.productName,
              serial: purchaseRecord.serial,
              reason: `Serial is in warehouse stock, but the stock row lists "${stockRecord.supplier || "another supplier"}" instead of DIP HOSON. Excluded to avoid supplier mixing.`,
              row: stockRecord.row,
              file: stock.fileName,
              supplier: stockRecord.supplier
            });
          } else {
            matchingSerials.push(purchaseRecord.serial);
          }
          return;
        }
        const otherStockRecord = stockBySerial.get(purchaseRecord.serialKey);
        if (otherStockRecord) {
          matchConflicts.push({
            id: `matching-product-${key}-${purchaseRecord.serialKey}-${otherStockRecord.row}`,
            source: "matching",
            sourceLabel: "Product mismatch",
            productKey: key,
            productName: purchaseEntry.productName,
            serial: purchaseRecord.serial,
            reason: `Serial is found in warehouse stock under "${otherStockRecord.productName}", not "${purchaseEntry.productName}". Not counted.`,
            row: otherStockRecord.row,
            file: stock.fileName,
            supplier: otherStockRecord.supplier
          });
        }
      });

      unverified.push(...matchConflicts);
      const productUnverified = unverified.filter((entry) => entry.productKey === key);
      const currentStock = new Set(stockValid.map((record) => record.serialKey)).size;
      const purchasedQty = purchaseSerialMap.size;
      const remaining = matchingSerials.length;
      const matchRate = purchasedQty ? remaining / purchasedQty : 0;
      let status = "sold-out";
      if (productUnverified.length) status = "unverified";
      else if (!hasStock) status = "awaiting-stock";
      else if (remaining > 0) status = "in-stock";

      results.push({
        productKey: key,
        productName: purchaseEntry.productName,
        purchasedQty,
        currentStock,
        remaining,
        matchingSerials: matchingSerials.sort((a, b) => textCollator.compare(a, b)),
        purchaseSerials: [...purchaseSerialMap.values()].map((record) => record.serial).sort((a, b) => textCollator.compare(a, b)),
        unverifiedCount: productUnverified.length,
        unverifiedEntries: productUnverified,
        matchRate,
        status
      });
    });

    const excludedStockProducts = [];
    stockByProduct.forEach((stockEntry, key) => {
      if (purchaseByProduct.has(key)) return;
      const serialCount = new Set(stockEntry.valid.map((record) => record.serialKey)).size;
      if (serialCount || stockEntry.unverified.length) {
        excludedStockProducts.push({
          productKey: key,
          productName: stockEntry.productName,
          serials: serialCount,
          unverified: stockEntry.unverified.length
        });
      }
    });
    excludedStockProducts.sort((a, b) => textCollator.compare(a.productName, b.productName));

    const purchasedUnits = results.reduce((sum, item) => sum + item.purchasedQty, 0);
    const remainingUnits = results.reduce((sum, item) => sum + item.remaining, 0);
    const warehouseSerials = new Set(validStockRecords.map((record) => record.serialKey)).size;
    const matchRate = purchasedUnits ? remainingUnits / purchasedUnits : 0;
    const stockOnlySerials = excludedStockProducts.reduce((sum, item) => sum + item.serials, 0);
    const unverifiedSorted = unverified.sort((a, b) => {
      const sourceOrder = { purchase: 0, stock: 1, matching: 2 };
      return (sourceOrder[a.source] ?? 9) - (sourceOrder[b.source] ?? 9)
        || textCollator.compare(a.productName, b.productName)
        || (a.row || 0) - (b.row || 0);
    });

    state.results = results.sort((a, b) => textCollator.compare(a.productName, b.productName));
    state.unverified = unverifiedSorted;
    state.excludedStockProducts = excludedStockProducts;
    state.summary = {
      purchasedUnits,
      warehouseSerials,
      remainingUnits,
      matchRate,
      matchRatePct: Math.round(matchRate * 1000) / 10,
      unverifiedCount: unverifiedSorted.length,
      excludedOtherSupplierRows: purchases?.ignoredOtherSupplierRows || 0,
      stockOnlyProducts: excludedStockProducts.length,
      stockOnlySerials,
      soldOutProducts: results.filter((item) => item.purchasedQty > 0 && item.remaining === 0).length,
      matchedProducts: results.filter((item) => item.remaining > 0).length,
      unverifiedProducts: results.filter((item) => item.unverifiedCount > 0).length,
      purchaseProducts: results.length,
      hasPurchases,
      hasStock
    };
  }

  function formatInt(value) {
    return numberFormat.format(Number(value) || 0);
  }

  function formatPct(value) {
    const number = Number(value) || 0;
    return `${Math.round(number * 1000) / 10}%`;
  }

  function escapeHtml(value) {
    return String(value ?? "").replace(/[&<>"']/g, (character) => ({
      "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
    }[character]));
  }

  function hashString(value) {
    let hash = 0;
    for (let index = 0; index < value.length; index += 1) {
      hash = ((hash << 5) - hash) + value.charCodeAt(index);
      hash |= 0;
    }
    return Math.abs(hash);
  }

  function productInitials(name) {
    const parts = String(name || "").replace(/[^\p{L}\p{N} ]/gu, " ").trim().split(/\s+/).filter(Boolean);
    if (!parts.length) return "?";
    return (parts[0][0] + (parts[1]?.[0] || "")).toUpperCase();
  }

  function renderAll() {
    renderSourceCard("purchase");
    renderSourceCard("stock");
    renderStats();
    renderHero();
    renderTable();
    renderReview();
    renderGuardBanner();
    const ready = Boolean(state.purchases && state.stock);
    els.topExportBtn.disabled = !ready;
    els.tableExportBtn.disabled = !ready;
  }

  function renderSourceCard(kind) {
    const dataset = kind === "purchase" ? state.purchases : state.stock;
    const isPurchase = kind === "purchase";
    const prefix = kind;
    const dropzone = els[`${prefix}Dropzone`];
    const summary = els[`${prefix}Summary`];
    const meta = els[`${prefix}Meta`];
    const loading = els[`${prefix}Loading`];
    const card = els[`${prefix}Card`];
    const input = els[`${prefix}Input`];
    const isLoading = Boolean(state.loading[kind]);

    card.classList.toggle("is-loading", isLoading);
    card.setAttribute("aria-busy", isLoading ? "true" : "false");
    loading.classList.toggle("hidden", !isLoading);
    input.disabled = isLoading;
    dropzone.classList.toggle("hidden", Boolean(dataset));
    summary.classList.toggle("hidden", !dataset);

    if (!dataset) {
      meta.innerHTML = "";
      return;
    }

    els[`${prefix}FileName`].textContent = dataset.fileName;
    els[`${prefix}FileMeta`].textContent = `${dataset.sheetName} · ${formatInt(dataset.dataRows)} data rows read`;
    els[`${prefix}FileName`].title = dataset.fileName;

    const chips = [];
    chips.push(`<span class="meta-chip success">${formatInt(dataset.validCount)} ${isPurchase ? "DIP HOSON" : "valid"} serials</span>`);
    if (dataset.unverifiedCount) {
      chips.push(`<span class="meta-chip warning">${formatInt(dataset.unverifiedCount)} unverified</span>`);
    }
    if (isPurchase && dataset.ignoredOtherSupplierRows) {
      chips.push(`<span class="meta-chip neutral">${formatInt(dataset.ignoredOtherSupplierRows)} other-supplier row${dataset.ignoredOtherSupplierRows === 1 ? "" : "s"} ignored</span>`);
    }
    if (!isPurchase && dataset.supplierConflictRows) {
      chips.push(`<span class="meta-chip neutral">${formatInt(dataset.supplierConflictRows)} other-supplier row${dataset.supplierConflictRows === 1 ? "" : "s"} checked</span>`);
    }
    chips.push(`<span class="meta-chip ghost">${dataset.supplierColumnPresent ? "Supplier column checked" : "No supplier column"}</span>`);
    chips.push(`<span class="meta-chip ghost">Header row ${formatInt(dataset.headerRow)}</span>`);
    meta.innerHTML = chips.join("");
  }

  function renderStats() {
    const summary = state.summary;
    const hasPurchases = Boolean(state.purchases);
    const hasStock = Boolean(state.stock);
    const hasData = hasPurchases || hasStock;

    els.statPurchased.textContent = hasPurchases ? formatInt(summary.purchasedUnits) : "0";
    els.statPurchasedSub.textContent = hasPurchases
      ? `${formatInt(summary.purchaseProducts)} product${summary.purchaseProducts === 1 ? "" : "s"} · unique serials`
      : "Unique purchase serials";

    els.statStock.textContent = hasStock ? formatInt(summary.warehouseSerials) : "—";
    els.statStockSub.textContent = hasStock
      ? `${formatInt(summary.stockOnlyProducts)} unlinked product${summary.stockOnlyProducts === 1 ? "" : "s"} excluded`
      : "Valid serials in stock file";

    els.statRemaining.textContent = hasStock && hasPurchases ? formatInt(summary.remainingUnits) : "—";
    els.statRemainingSub.textContent = hasStock && hasPurchases
      ? `${formatInt(summary.matchedProducts)} product${summary.matchedProducts === 1 ? "" : "s"} with matched stock`
      : "Exact serial matches only";

    els.statRate.textContent = hasStock && hasPurchases ? `${summary.matchRatePct}%` : "—";
    els.statRateBar.style.width = hasStock && hasPurchases ? `${Math.min(100, summary.matchRatePct)}%` : "0%";
    els.statRateSub.textContent = hasStock && hasPurchases
      ? `${formatInt(summary.remainingUnits)} of ${formatInt(summary.purchasedUnits)} purchased serials`
      : "Remaining ÷ purchased";

    els.statUnverified.textContent = formatInt(summary.unverifiedCount);
    els.statUnverifiedSub.textContent = summary.unverifiedCount
      ? `${formatInt(summary.unverifiedProducts)} product${summary.unverifiedProducts === 1 ? "" : "s"} affected · excluded`
      : hasData ? "No unverified serials" : "Excluded from all counts";
    els.unverifiedStatCard.classList.toggle("has-alert", summary.unverifiedCount > 0);
  }

  function renderHero() {
    const hasPurchases = Boolean(state.purchases);
    const hasStock = Boolean(state.stock);
    els.heroPurchased.textContent = hasPurchases ? formatInt(state.summary.purchasedUnits) : "—";
    els.heroStock.textContent = hasStock ? formatInt(state.summary.warehouseSerials) : "—";
    els.heroMatched.textContent = hasPurchases && hasStock ? formatInt(state.summary.remainingUnits) : "—";
  }

  function getFilteredResults() {
    const query = state.filters.search.trim().toLowerCase();
    let items = state.results.slice();
    if (query) {
      items = items.filter((item) => {
        if (item.productName.toLowerCase().includes(query)) return true;
        if (item.matchingSerials.some((serial) => serial.toLowerCase().includes(query))) return true;
        if (item.purchaseSerials.some((serial) => serial.toLowerCase().includes(query))) return true;
        return item.unverifiedEntries.some((entry) => String(entry.serial || "").toLowerCase().includes(query));
      });
    }
    if (state.filters.status !== "all") {
      items = items.filter((item) => item.status === state.filters.status);
    }
    return sortResults(items);
  }

  function sortResults(items) {
    const sorted = items.slice();
    switch (state.filters.sort) {
      case "product-asc":
        sorted.sort((a, b) => textCollator.compare(a.productName, b.productName));
        break;
      case "purchased-desc":
        sorted.sort((a, b) => b.purchasedQty - a.purchasedQty || textCollator.compare(a.productName, b.productName));
        break;
      case "current-desc":
        sorted.sort((a, b) => b.currentStock - a.currentStock || textCollator.compare(a.productName, b.productName));
        break;
      case "match-desc":
        sorted.sort((a, b) => b.matchRate - a.matchRate || textCollator.compare(a.productName, b.productName));
        break;
      default:
        sorted.sort((a, b) => b.remaining - a.remaining || b.purchasedQty - a.purchasedQty || textCollator.compare(a.productName, b.productName));
    }
    return sorted;
  }

  function renderStatusBadge(result, hasStock) {
    if (result.unverifiedCount > 0) {
      return `<span class="status status-unverified">Unverified · ${formatInt(result.unverifiedCount)}</span>`;
    }
    if (!hasStock) return `<span class="status status-pending">Awaiting stock</span>`;
    if (result.remaining > 0) return `<span class="status status-ok">Verified</span>`;
    return `<span class="status status-neutral">Verified</span>`;
  }

  function renderSerialsCell(result, hasStock) {
    if (!hasStock) return `<span class="muted">Pending stock import</span>`;
    if (!result.matchingSerials.length) return `<span class="muted">No matching serials</span>`;
    const visible = result.matchingSerials.slice(0, 2);
    const remaining = result.matchingSerials.length - visible.length;
    return `<div class="serial-chips">${visible.map((serial) => `<span class="serial-chip" title="${escapeHtml(serial)}">${escapeHtml(serial)}</span>`).join("")}${remaining > 0 ? `<button class="serial-more" type="button" data-details="${escapeHtml(result.productKey)}">+${remaining} more</button>` : ""}</div>`;
  }

  function renderTable() {
    const hasStock = Boolean(state.stock);
    const filtered = getFilteredResults();
    const total = filtered.length;
    const totalPages = Math.max(1, Math.ceil(total / PAGE_SIZE));
    state.filters.page = Math.min(Math.max(1, state.filters.page), totalPages);
    const start = (state.filters.page - 1) * PAGE_SIZE;
    const pageItems = filtered.slice(start, start + PAGE_SIZE);

    els.resultCount.textContent = `${formatInt(state.results.length)} product${state.results.length === 1 ? "" : "s"}`;
    els.resultsBody.innerHTML = pageItems.map((result) => {
      const hue = hashString(result.productKey) % 360;
      const initials = escapeHtml(productInitials(result.productName));
      const rate = Math.round(result.matchRate * 100);
      const statusNote = !hasStock ? "Waiting for warehouse file"
        : result.remaining > 0 ? "DIP HOSON units found"
        : "No DIP HOSON serials found";
      return `<tr data-product-key="${escapeHtml(result.productKey)}" tabindex="0" aria-label="View details for ${escapeHtml(result.productName)}">
        <td>
          <div class="product-cell">
            <span class="product-avatar" style="--avatar-hue:${hue}">${initials}</span>
            <span class="product-info">
              <button class="product-name-btn" type="button" data-details="${escapeHtml(result.productKey)}">${escapeHtml(result.productName)}</button>
              <span class="product-sub">${formatInt(result.purchasedQty)} unique purchase serial${result.purchasedQty === 1 ? "" : "s"}</span>
            </span>
          </div>
        </td>
        <td class="num"><strong>${formatInt(result.purchasedQty)}</strong></td>
        <td class="num">${hasStock ? `<strong>${formatInt(result.currentStock)}</strong>` : `<span class="muted">—</span>`}</td>
        <td class="num">
          <div class="remaining-cell"><strong>${hasStock ? formatInt(result.remaining) : "—"}</strong>${hasStock ? `<span>${rate}%</span>` : ""}</div>
        </td>
        <td>${renderSerialsCell(result, hasStock)}</td>
        <td>
          ${renderStatusBadge(result, hasStock)}
          <span class="status-note">${statusNote}</span>
        </td>
      </tr>`;
    }).join("");

    els.resultsTable.classList.toggle("hidden", total === 0);
    els.resultsEmpty.classList.toggle("hidden", total > 0);
    if (total === 0) {
      const strong = els.resultsEmpty.querySelector("strong");
      const paragraph = els.resultsEmpty.querySelector("p");
      if (!state.purchases) {
        strong.textContent = "No reconciliation yet";
        paragraph.textContent = "Import DIP HOSON purchases and warehouse stock to calculate verified remaining stock.";
      } else if (!state.stock) {
        strong.textContent = "Purchases loaded — add warehouse stock";
        paragraph.textContent = "Remaining DIP HOSON stock cannot be calculated until the current warehouse stock file is imported.";
      } else if (state.filters.search || state.filters.status !== "all") {
        strong.textContent = "No products match these filters";
        paragraph.textContent = "Try a different product name, serial number, or status filter.";
      } else {
        strong.textContent = "No DIP HOSON products found";
        paragraph.textContent = "The purchase file did not contain any usable product rows.";
      }
    }

    if (total === 0) {
      els.paginationInfo.textContent = "";
      els.pagination.innerHTML = "";
    } else {
      const end = Math.min(start + PAGE_SIZE, total);
      els.paginationInfo.textContent = `Showing ${formatInt(start + 1)}–${formatInt(end)} of ${formatInt(total)} product${total === 1 ? "" : "s"}`;
      renderPagination(totalPages);
    }
  }

  function renderPagination(totalPages) {
    if (totalPages <= 1) {
      els.pagination.innerHTML = "";
      return;
    }
    const current = state.filters.page;
    let first = Math.max(1, current - 2);
    let last = Math.min(totalPages, first + 4);
    first = Math.max(1, last - 4);
    const buttons = [];
    buttons.push(`<button class="page-btn" type="button" data-page="${current - 1}" ${current === 1 ? "disabled" : ""} aria-label="Previous page"><svg class="icon"><use href="#i-chevron-left"/></svg></button>`);
    for (let page = first; page <= last; page += 1) {
      buttons.push(`<button class="page-btn ${page === current ? "is-active" : ""}" type="button" data-page="${page}" aria-label="Page ${page}">${page}</button>`);
    }
    buttons.push(`<button class="page-btn" type="button" data-page="${current + 1}" ${current === totalPages ? "disabled" : ""} aria-label="Next page"><svg class="icon"><use href="#i-chevron-right"/></svg></button>`);
    els.pagination.innerHTML = buttons.join("");
  }

  function getFilteredUnverified() {
    const query = state.reviewFilters.search.trim().toLowerCase();
    let items = state.unverified.slice();
    if (state.reviewFilters.source !== "all") {
      items = items.filter((entry) => entry.source === state.reviewFilters.source);
    }
    if (query) {
      items = items.filter((entry) => [
        entry.productName,
        entry.serial,
        entry.reason,
        entry.sourceLabel,
        entry.file
      ].some((value) => String(value || "").toLowerCase().includes(query)));
    }
    return items;
  }

  function renderReview() {
    const filtered = getFilteredUnverified();
    const total = state.unverified.length;
    const visible = filtered.slice(0, REVIEW_LIMIT);
    els.unverifiedCountPill.textContent = `${formatInt(total)} unverified`;
    els.navUnverifiedBadge.textContent = formatInt(total);
    els.navUnverifiedBadge.classList.toggle("hidden", total === 0);

    els.reviewBody.innerHTML = visible.map((entry) => {
      const tagClass = entry.source === "purchase" ? "source-purchase"
        : entry.source === "stock" ? "source-stock" : "source-matching";
      const tagText = entry.source === "purchase" ? "DIP HOSON"
        : entry.source === "stock" ? "Warehouse" : entry.sourceLabel;
      return `<tr>
        <td><span class="source-tag ${tagClass}">${escapeHtml(tagText)}</span></td>
        <td>
          <div class="review-product">${escapeHtml(entry.productName)}</div>
          <div class="review-file" title="${escapeHtml(entry.file)}">${escapeHtml(entry.file)}</div>
        </td>
        <td><code class="serial-code" title="${escapeHtml(entry.serial)}">${escapeHtml(entry.serial)}</code></td>
        <td class="reason-cell">${escapeHtml(entry.reason)}</td>
        <td class="row-num">${entry.row ? formatInt(entry.row) : "—"}</td>
      </tr>`;
    }).join("");

    const hasRows = filtered.length > 0;
    els.reviewTable.classList.toggle("hidden", !hasRows);
    els.reviewEmpty.classList.toggle("hidden", hasRows);
    if (!hasRows) {
      const strong = els.reviewEmpty.querySelector("strong");
      const paragraph = els.reviewEmpty.querySelector("p");
      if (total === 0) {
        strong.textContent = "No unverified serials";
        paragraph.textContent = "Every counted unit is backed by a valid serial match.";
      } else {
        strong.textContent = "No unverified serials match these filters";
        paragraph.textContent = "Clear the search or choose another source to see the excluded rows.";
      }
    }

    if (filtered.length > REVIEW_LIMIT) {
      els.reviewFooter.textContent = `Showing the first ${REVIEW_LIMIT} of ${formatInt(filtered.length)} matching unverified rows. The Excel export includes every row.`;
    } else if (filtered.length !== total) {
      els.reviewFooter.textContent = `Showing ${formatInt(filtered.length)} of ${formatInt(total)} unverified rows.`;
    } else {
      els.reviewFooter.textContent = "";
    }
  }

  function renderGuardBanner() {
    const ignored = state.purchases?.ignoredOtherSupplierRows || 0;
    const stockOnlyProducts = state.summary.stockOnlyProducts;
    const stockOnlySerials = state.summary.stockOnlySerials;
    const hasExclusions = ignored > 0 || stockOnlyProducts > 0;
    els.guardBanner.classList.toggle("has-exclusions", hasExclusions);

    if (!state.purchases) {
      els.guardBanner.innerHTML = `<span class="guard-banner-icon"><svg class="icon"><use href="#i-shield"/></svg></span><div><strong>Supplier guard active.</strong> Import DIP HOSON purchases first. Warehouse stock alone is never added to the report.</div>`;
      return;
    }

    const notes = ["Only serials from DIP HOSON purchases are eligible for remaining stock."];
    if (ignored) notes.push(`${formatInt(ignored)} other-supplier purchase row${ignored === 1 ? " was" : "s were"} excluded.`);
    if (stockOnlyProducts) notes.push(`${formatInt(stockOnlyProducts)} warehouse product${stockOnlyProducts === 1 ? " is" : "s are"} not linked to DIP HOSON purchases and ${stockOnlyProducts === 1 ? "was" : "were"} excluded (${formatInt(stockOnlySerials)} serial${stockOnlySerials === 1 ? "" : "s"}).`);
    if (!ignored && !stockOnlyProducts) notes.push("No other-supplier purchase rows or unlinked warehouse products were included.");
    els.guardBanner.innerHTML = `<span class="guard-banner-icon"><svg class="icon"><use href="#i-shield"/></svg></span><div><strong>Supplier guard active.</strong> ${escapeHtml(notes.join(" "))}</div>`;
  }

  function openDetails(key) {
    const result = state.results.find((item) => item.productKey === key);
    if (!result) return;
    state.detailsKey = key;
    const hasStock = Boolean(state.stock);
    const rate = Math.round(result.matchRate * 100);
    const matchingList = result.matchingSerials.length
      ? `<div class="serial-list">${result.matchingSerials.map((serial) => `<span class="serial-chip">${escapeHtml(serial)}</span>`).join("")}</div>`
      : `<div class="detail-empty">${hasStock ? "No matching DIP HOSON serials were found in warehouse stock." : "Import warehouse stock to calculate matching serials."}</div>`;
    const unverifiedList = result.unverifiedEntries.length
      ? `<div class="unverified-list">${result.unverifiedEntries.map((entry) => {
        const tagClass = entry.source === "purchase" ? "source-purchase" : entry.source === "stock" ? "source-stock" : "source-matching";
        const tagText = entry.source === "purchase" ? "DIP HOSON" : entry.source === "stock" ? "Warehouse" : entry.sourceLabel;
        return `<div class="unverified-item">
          <span class="source-tag ${tagClass}">${escapeHtml(tagText)}</span>
          <div>
            <div class="unverified-meta"><code class="serial-code">${escapeHtml(entry.serial)}</code><span class="review-file">Row ${entry.row ? formatInt(entry.row) : "—"}</span></div>
            <p>${escapeHtml(entry.reason)}</p>
          </div>
        </div>`;
      }).join("")}</div>`
      : `<div class="detail-empty success">No unverified serials for this product.</div>`;

    els.detailsContent.innerHTML = `<div class="modal-head">
      <div><span class="modal-kicker">Product reconciliation</span><h2 id="detailsTitle">${escapeHtml(result.productName)}</h2><div class="detail-status">${renderStatusBadge(result, hasStock)}</div></div>
      <button class="icon-btn" type="button" data-close-dialog aria-label="Close details"><svg class="icon"><use href="#i-x"/></svg></button>
    </div>
    <div class="detail-stats">
      <div class="detail-stat"><span>Purchased Qty</span><strong>${formatInt(result.purchasedQty)}</strong></div>
      <div class="detail-stat"><span>Current Stock</span><strong>${hasStock ? formatInt(result.currentStock) : "—"}</strong></div>
      <div class="detail-stat"><span>Remaining DIP HOSON</span><strong class="orange">${hasStock ? formatInt(result.remaining) : "—"}</strong></div>
      <div class="detail-stat"><span>Match Rate</span><strong>${hasStock ? `${rate}%` : "—"}</strong></div>
    </div>
    <div class="detail-progress">
      <div class="detail-progress-head"><span>Exact serial match progress</span><span>${hasStock ? `${rate}%` : "Pending stock import"}</span></div>
      <div class="mini-track"><span style="width:${hasStock ? Math.min(100, rate) : 0}%"></span></div>
    </div>
    <div class="detail-section">
      <div class="detail-section-head"><h3>Matching serial numbers <span>(${formatInt(result.matchingSerials.length)})</span></h3>${result.matchingSerials.length ? `<button class="link-btn small" type="button" id="copySerialsBtn"><svg class="icon"><use href="#i-copy"/></svg>Copy all</button>` : ""}</div>
      ${matchingList}
    </div>
    <div class="detail-section">
      <div class="detail-section-head"><h3>Unverified serials <span>(${formatInt(result.unverifiedEntries.length)})</span></h3></div>
      ${unverifiedList}
    </div>
    <div class="detail-note"><strong>No quantity guessing:</strong> remaining stock is the count of exact serial matches between DIP HOSON purchases and current warehouse stock. Unverified rows are excluded.</div>`;

    if (typeof els.detailsDialog.showModal === "function") {
      els.detailsDialog.showModal();
      document.body.classList.add("modal-open");
    } else {
      els.detailsDialog.setAttribute("open", "");
    }
  }

  function closeDetails() {
    if (els.detailsDialog.open) els.detailsDialog.close();
    document.body.classList.remove("modal-open");
    state.detailsKey = null;
  }

  function showToast(message, type = "info", title = "") {
    const titles = { success: "Success", error: "Something needs attention", warning: "Heads up", info: "Info" };
    const icons = { success: "i-check", error: "i-alert", warning: "i-alert", info: "i-info" };
    const toast = document.createElement("div");
    toast.className = `toast ${type}`;
    const id = `toast-${++toastCounter}`;
    toast.id = id;
    toast.innerHTML = `<span class="toast-icon"><svg class="icon"><use href="#${icons[type] || icons.info}"/></svg></span>
      <div class="toast-copy"><strong>${escapeHtml(title || titles[type] || titles.info)}</strong><span>${escapeHtml(message)}</span></div>
      <button class="toast-close" type="button" aria-label="Dismiss notification"><svg class="icon"><use href="#i-x"/></svg></button>`;
    const stack = els.toastStack;
    while (stack.children.length >= 4) stack.firstElementChild.remove();
    stack.appendChild(toast);
    const dismiss = () => {
      if (!toast.isConnected) return;
      toast.classList.add("is-leaving");
      setTimeout(() => toast.remove(), 190);
    };
    toast.querySelector(".toast-close").addEventListener("click", dismiss);
    setTimeout(dismiss, type === "error" ? 7500 : 4800);
  }

  function setLoading(kind, loading) {
    state.loading[kind] = loading;
    renderSourceCard(kind);
  }

  async function handleFile(file, kind) {
    if (!file) return;
    const label = kind === "purchase" ? "DIP HOSON purchases" : "warehouse stock";
    setLoading(kind, true);
    try {
      const dataset = await parseWorkbookFile(file, kind);
      if (kind === "purchase") state.purchases = dataset;
      else state.stock = dataset;
      state.filters.page = 1;
      state.restored = false;
      refreshAll({ save: true });
      const details = [];
      details.push(`${formatInt(dataset.validCount)} valid serials`);
      if (dataset.unverifiedCount) details.push(`${formatInt(dataset.unverifiedCount)} unverified`);
      if (kind === "purchase" && dataset.ignoredOtherSupplierRows) details.push(`${formatInt(dataset.ignoredOtherSupplierRows)} other-supplier rows ignored`);
      showToast(`${label} imported from "${file.name}": ${details.join(", ")}.`, "success", "Workbook imported");
      if (state.purchases && state.stock) {
        showToast(`${formatInt(state.summary.remainingUnits)} DIP HOSON units remain after exact serial matching.`, "info", "Reconciliation updated");
      }
    } catch (error) {
      showToast(error.message || "The workbook could not be imported.", "error", "Import failed");
    } finally {
      setLoading(kind, false);
      const input = els[`${kind}Input`];
      if (input) input.value = "";
    }
  }

  function buildDemoDataset(kind) {
    const purchaseRows = [
      ["Product Name", "Serial Number", "Supplier", "Purchase Date", "Invoice No."],
      ["UltraBook Pro 14", "UH14-2601-0001", "DIP HOSON", "2026-01-12", "DH-1042"],
      ["UltraBook Pro 14", "UH14-2601-0002", "DIP HOSON", "2026-01-12", "DH-1042"],
      ["UltraBook Pro 14", "UH14-2601-0003", "DIP HOSON", "2026-01-12", "DH-1042"],
      ["UltraBook Pro 14", "UH14-2601-0004", "DIP HOSON", "2026-01-12", "DH-1042"],
      ["UltraBook Pro 14", "UH14-2601-0005", "DIP HOSON", "2026-01-12", "DH-1042"],
      ["UltraBook Pro 14", "UH14-2601-0006", "DIP HOSON", "2026-01-12", "DH-1042"],
      ["UltraBook Air 13", "UA13-2602-0001", "DIP HOSON", "2026-02-03", "DH-1088"],
      ["UltraBook Air 13", "UA13-2602-0002", "DIP HOSON", "2026-02-03", "DH-1088"],
      ["UltraBook Air 13", "UA13-2602-0003", "DIP HOSON", "2026-02-03", "DH-1088"],
      ["UltraBook Air 13", "UA13-2602-0004", "DIP HOSON", "2026-02-03", "DH-1088"],
      ["UltraStation Mini", "USM-2603-0001", "DIP HOSON", "2026-03-18", "DH-1150"],
      ["UltraStation Mini", "USM-2603-0002", "DIP HOSON", "2026-03-18", "DH-1150"],
      ["UltraStation Mini", "USM-2603-0003", "DIP HOSON", "2026-03-18", "DH-1150"],
      ["UltraView 27", "UV27-2604-0001", "DIP HOSON", "2026-04-07", "DH-1194"],
      ["UltraView 27", "UV27-2604-0002", "DIP HOSON", "2026-04-07", "DH-1194"],
      ["UltraView 27", "UV27-2604-0003", "DIP HOSON", "2026-04-07", "DH-1194"],
      ["UltraView 27", "UV27-2604-0004", "DIP HOSON", "2026-04-07", "DH-1194"],
      ["UltraView 27", "UV27-2604-0005", "DIP HOSON", "2026-04-07", "DH-1194"],
      ["UltraGear K87", "K87-2605-0001", "DIP HOSON", "2026-05-21", "DH-1240"],
      ["UltraGear K87", "K87-2605-0002", "DIP HOSON", "2026-05-21", "DH-1240"],
      ["UltraGear K87", "K87-2605-0003", "DIP HOSON", "2026-05-21", "DH-1240"],
      ["UltraGear K87", "K87-2605-0004", "DIP HOSON", "2026-05-21", "DH-1240"],
      ["UltraGear K87", "K87-2605-0005", "DIP HOSON", "2026-05-21", "DH-1240"],
      ["UltraGear K87", "K87-2605-0006", "DIP HOSON", "2026-05-21", "DH-1240"],
      ["UltraGear K87", "K87-2605-0007", "DIP HOSON", "2026-05-21", "DH-1240"],
      ["UltraGear K87", "K87-2605-0008", "DIP HOSON", "2026-05-21", "DH-1240"],
      ["UltraBook Air 13", "", "DIP HOSON", "2026-02-03", "DH-1088"],
      ["UltraBook Air 13", "N/A", "DIP HOSON", "2026-02-03", "DH-1088"],
      ["UltraBook Air 13", "UA13-2602-0002", "DIP HOSON", "2026-02-03", "DH-1088"],
      ["UltraBook Pro 14", "UH14-2601-0001", "TechSource", "2026-01-12", "TS-991"]
    ];

    const stockRows = [
      ["Product Name", "Serial Number", "Supplier", "Location"],
      ["UltraBook Pro 14", "UH14-2601-0001", "DIP HOSON", "A1-01"],
      ["UltraBook Pro 14", "UH14-2601-0002", "TechSource", "A1-02"],
      ["UltraBook Pro 14", "UH14-2601-0004", "DIP HOSON", "A1-03"],
      ["UltraBook Pro 14", "UH14-2601-0006", "DIP HOSON", "A1-04"],
      ["UltraBook Pro 14", "OTH-PRO-001", "TechSource", "A1-05"],
      ["UltraBook Air 13", "UA13-2602-0001", "DIP HOSON", "A2-01"],
      ["UltraBook Air 13", "UA13-2602-0003", "DIP HOSON", "A2-02"],
      ["UltraBook Air 13", "OTH-AIR-001", "TechSource", "A2-03"],
      ["UltraBook Air 13", "UH14-2601-0005", "DIP HOSON", "A2-04"],
      ["UltraStation Mini", "OTH-MINI-001", "TechSource", "B1-01"],
      ["UltraView 27", "UV27-2604-0001", "DIP HOSON", "C1-01"],
      ["UltraView 27", "UV27-2604-0002", "DIP HOSON", "C1-02"],
      ["UltraView 27", "UV27-2604-0003", "DIP HOSON", "C1-03"],
      ["UltraView 27", "UV27-2604-0004", "DIP HOSON", "C1-04"],
      ["UltraView 27", "UV27-2604-0005", "DIP HOSON", "C1-05"],
      ["UltraView 27", "N/A", "DIP HOSON", "C1-06"],
      ["UltraGear K87", "K87-2605-0001", "DIP HOSON", "D1-01"],
      ["UltraGear K87", "K87-2605-0002", "DIP HOSON", "D1-02"],
      ["UltraGear K87", "K87-2605-0003", "DIP HOSON", "D1-03"],
      ["UltraGear K87", "K87-2605-0001", "DIP HOSON", "D1-04"],
      ["UltraGear K87", "", "DIP HOSON", "D1-05"],
      ["UltraDock USB-C Hub", "UD-2606-0001", "DIP HOSON", "E1-01"],
      ["UltraDock USB-C Hub", "UD-2606-0002", "DIP HOSON", "E1-02"],
      ["UltraDock USB-C Hub", "UD-2606-0003", "DIP HOSON", "E1-03"]
    ];

    const rows = kind === "purchase" ? purchaseRows : stockRows;
    const header = {
      rowIndex: 0,
      headers: rows[0],
      mapping: { product: 0, serial: 1, supplier: 2, quantity: -1 }
    };
    return buildDatasetFromSheet(kind, kind === "purchase" ? "demo-dip-hoson-purchases.xlsx" : "demo-warehouse-stock.xlsx", {
      name: kind === "purchase" ? "DIP HOSON Purchases" : "Warehouse Stock",
      rows,
      header
    });
  }

  function applyDemoData() {
    state.purchases = buildDemoDataset("purchase");
    state.stock = buildDemoDataset("stock");
    state.filters = { search: "", status: "all", sort: "remaining-desc", page: 1 };
    state.reviewFilters = { search: "", source: "all" };
    state.restored = false;
    refreshAll({ save: true });
    showToast("Demo purchases and warehouse stock are loaded. Explore the serial review to see exclusions.", "success", "Demo data ready");
  }

  function loadDemoData() {
    if (state.purchases || state.stock) {
      openConfirm({
        title: "Replace current data with demo data?",
        message: "Both imported files and the current reconciliation will be replaced by the built-in demo workbook.",
        confirmLabel: "Load demo data",
        danger: false,
        onConfirm: applyDemoData
      });
    } else {
      applyDemoData();
    }
  }

  function refreshAll({ save = false } = {}) {
    computeReconciliation();
    renderAll();
    if (save) saveState();
  }

  function saveState() {
    if (state.persistenceDisabled) return;
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify({
        version: STORAGE_VERSION,
        savedAt: new Date().toISOString(),
        purchases: state.purchases,
        stock: state.stock
      }));
    } catch (error) {
      state.persistenceDisabled = true;
      showToast("The browser storage limit was reached. This session still works, but data will not be restored after refresh.", "warning", "Local save unavailable");
    }
  }

  function restoreState() {
    let raw;
    try {
      raw = localStorage.getItem(STORAGE_KEY);
    } catch (error) {
      return;
    }
    if (!raw) return;
    try {
      const saved = JSON.parse(raw);
      if (!saved || saved.version !== STORAGE_VERSION) return;
      const validPurchases = saved.purchases && Array.isArray(saved.purchases.records);
      const validStock = saved.stock && Array.isArray(saved.stock.records);
      if (!validPurchases && !validStock) return;
      if (validPurchases) state.purchases = saved.purchases;
      if (validStock) state.stock = saved.stock;
      state.restored = true;
    } catch (error) {
      try { localStorage.removeItem(STORAGE_KEY); } catch (removeError) { /* ignore */ }
    }
  }

  function removeDataset(kind) {
    const dataset = kind === "purchase" ? state.purchases : state.stock;
    if (!dataset) return;
    openConfirm({
      title: `Remove ${kind === "purchase" ? "DIP HOSON purchases" : "warehouse stock"}?`,
      message: `"${dataset.fileName}" will be removed. The reconciliation will be recalculated from the remaining file.`,
      confirmLabel: "Remove file",
      danger: true,
      onConfirm: () => {
        if (kind === "purchase") state.purchases = null;
        else state.stock = null;
        state.filters.page = 1;
        refreshAll({ save: true });
        showToast(`${kind === "purchase" ? "Purchase" : "Warehouse stock"} file removed.`, "info", "File removed");
      }
    });
  }

  function resetAll() {
    if (!state.purchases && !state.stock) {
      showToast("There is no imported data to clear.", "info");
      return;
    }
    openConfirm({
      title: "Clear all imported data?",
      message: "Both files, every reconciliation result, and the local session copy will be removed.",
      confirmLabel: "Clear data",
      danger: true,
      onConfirm: () => {
        state.purchases = null;
        state.stock = null;
        state.results = [];
        state.unverified = [];
        state.excludedStockProducts = [];
        state.summary = emptySummary();
        state.filters = { search: "", status: "all", sort: "remaining-desc", page: 1 };
        state.reviewFilters = { search: "", source: "all" };
        state.restored = false;
        els.productSearch.value = "";
        els.statusFilter.value = "all";
        els.sortSelect.value = "remaining-desc";
        els.reviewSearch.value = "";
        els.reviewSourceFilter.value = "all";
        try { localStorage.removeItem(STORAGE_KEY); } catch (error) { /* ignore */ }
        refreshAll();
        showToast("All imported data and local session storage were cleared.", "info", "Data cleared");
      }
    });
  }

  function openConfirm({ title, message, confirmLabel, danger, onConfirm }) {
    state.pendingConfirm = onConfirm;
    els.confirmTitle.textContent = title;
    els.confirmMessage.textContent = message;
    els.confirmActionBtn.textContent = confirmLabel;
    els.confirmActionBtn.className = `btn ${danger ? "btn-danger" : "btn-primary"}`;
    els.confirmIcon.className = `confirm-icon${danger ? "" : " is-primary"}`;
    els.confirmDialog.showModal();
    document.body.classList.add("modal-open");
  }

  function closeConfirm() {
    state.pendingConfirm = null;
    if (els.confirmDialog.open) els.confirmDialog.close();
    document.body.classList.remove("modal-open");
  }

  function excelSerialCell(serials) {
    const full = serials.join("\n");
    if (full.length <= 32000) return full;
    return `${full.slice(0, 31800)}\n… ${serials.length} serials total. See the "Matched Serials" sheet.`;
  }

  function exportResults() {
    if (!XLSX_LIB) {
      showToast("The Excel library did not load. Refresh the page and try again.", "error");
      return;
    }
    if (!state.purchases || !state.stock) {
      showToast("Import both the DIP HOSON purchases and warehouse stock files before exporting.", "warning", "Files still needed");
      return;
    }
    try {
      const workbook = XLSX_LIB.utils.book_new();
      const generated = new Date();
      const sortedResults = state.results.slice().sort((a, b) => textCollator.compare(a.productName, b.productName));

      const summaryRows = [
        ["UltraPC — DIP HOSON Stock Reconciliation"],
        ["Generated", generated.toLocaleString()],
        [],
        ["Metric", "Value"],
        ["DIP HOSON purchase file", state.purchases.fileName],
        ["Purchase sheet", state.purchases.sheetName],
        ["Purchased DIP HOSON units", state.summary.purchasedUnits],
        ["Warehouse stock file", state.stock.fileName],
        ["Warehouse sheet", state.stock.sheetName],
        ["Warehouse serials scanned", state.summary.warehouseSerials],
        ["Remaining DIP HOSON units", state.summary.remainingUnits],
        ["Match rate", `${state.summary.matchRatePct}%`],
        ["Unverified serials excluded", state.summary.unverifiedCount],
        ["Other-supplier purchase rows excluded", state.summary.excludedOtherSupplierRows],
        ["Stock-only products excluded", state.summary.stockOnlyProducts],
        ["Stock-only serials excluded", state.summary.stockOnlySerials],
        [],
        ["Rules"],
        ["1", "Only exact serial matches from DIP HOSON purchases are counted as remaining stock."],
        ["2", "Purchased Quantity is the count of unique valid DIP HOSON serial numbers, not a quantity column."],
        ["3", "Current Stock shows all valid serialized warehouse units for the product, regardless of supplier."],
        ["4", "Missing, invalid, duplicate, or conflicting serials are marked Unverified and excluded from all counts."],
        ["5", "A warehouse row that explicitly names another supplier is never counted as DIP HOSON remaining stock."],
        ["6", "Products in warehouse stock that are not in the DIP HOSON purchase file are excluded from the report."]
      ];
      const summarySheet = XLSX_LIB.utils.aoa_to_sheet(summaryRows);
      summarySheet["!cols"] = [{ wch: 42 }, { wch: 74 }];
      XLSX_LIB.utils.book_append_sheet(workbook, summarySheet, "Summary");

      const mainRows = [[
        "Product Name",
        "Purchased Quantity",
        "Current Stock",
        "Remaining DIP HOSON Stock",
        "Matching Serial Numbers",
        "Match Rate",
        "Serial Status",
        "Unverified Count"
      ]];
      sortedResults.forEach((result) => {
        mainRows.push([
          result.productName,
          result.purchasedQty,
          result.currentStock,
          result.remaining,
          excelSerialCell(result.matchingSerials),
          `${Math.round(result.matchRate * 1000) / 10}%`,
          result.unverifiedCount > 0 ? "Unverified" : "Verified",
          result.unverifiedCount
        ]);
      });
      const mainSheet = XLSX_LIB.utils.aoa_to_sheet(mainRows);
      mainSheet["!cols"] = [
        { wch: 34 }, { wch: 19 }, { wch: 15 }, { wch: 25 },
        { wch: 62 }, { wch: 12 }, { wch: 16 }, { wch: 18 }
      ];
      if (mainRows.length > 1) mainSheet["!autofilter"] = { ref: `A1:H${mainRows.length}` };
      XLSX_LIB.utils.book_append_sheet(workbook, mainSheet, "Reconciliation");

      const matchedRows = [["Product Name", "Serial Number", "Verification"]];
      sortedResults.forEach((result) => {
        result.matchingSerials.forEach((serial) => matchedRows.push([result.productName, serial, "Matched DIP HOSON"]));
      });
      const matchedSheet = XLSX_LIB.utils.aoa_to_sheet(matchedRows);
      matchedSheet["!cols"] = [{ wch: 36 }, { wch: 30 }, { wch: 20 }];
      if (matchedRows.length > 1) matchedSheet["!autofilter"] = { ref: `A1:C${matchedRows.length}` };
      XLSX_LIB.utils.book_append_sheet(workbook, matchedSheet, "Matched Serials");

      const unverifiedRows = [["Source", "Product Name", "Serial Number", "Reason", "Excel Row", "File"]];
      state.unverified.forEach((entry) => {
        unverifiedRows.push([
          entry.source === "purchase" ? "DIP HOSON purchases" : entry.source === "stock" ? "Warehouse stock" : entry.sourceLabel,
          entry.productName,
          entry.serial,
          entry.reason,
          entry.row || "",
          entry.file || ""
        ]);
      });
      const unverifiedSheet = XLSX_LIB.utils.aoa_to_sheet(unverifiedRows);
      unverifiedSheet["!cols"] = [{ wch: 22 }, { wch: 30 }, { wch: 24 }, { wch: 78 }, { wch: 11 }, { wch: 34 }];
      if (unverifiedRows.length > 1) unverifiedSheet["!autofilter"] = { ref: `A1:F${unverifiedRows.length}` };
      XLSX_LIB.utils.book_append_sheet(workbook, unverifiedSheet, "Unverified Serials");

      const excludedRows = [["Type", "Product Name", "Serialized Units", "Reason"]];
      if (state.summary.excludedOtherSupplierRows) {
        excludedRows.push([
          "Purchase rows",
          "—",
          state.summary.excludedOtherSupplierRows,
          "Supplier column named a supplier other than DIP HOSON; rows excluded from the report."
        ]);
      }
      state.excludedStockProducts.forEach((product) => {
        excludedRows.push([
          "Warehouse product",
          product.productName,
          product.serials,
          "Product is not present in the DIP HOSON purchase file; excluded from the DIP HOSON report."
        ]);
      });
      const excludedSheet = XLSX_LIB.utils.aoa_to_sheet(excludedRows);
      excludedSheet["!cols"] = [{ wch: 20 }, { wch: 34 }, { wch: 17 }, { wch: 84 }];
      if (excludedRows.length > 1) excludedSheet["!autofilter"] = { ref: `A1:D${excludedRows.length}` };
      XLSX_LIB.utils.book_append_sheet(workbook, excludedSheet, "Excluded Data");

      const datePart = generated.toISOString().slice(0, 10);
      XLSX_LIB.writeFile(workbook, `UltraPC_DIP_HOSON_Reconciliation_${datePart}.xlsx`, { compression: true });
      showToast("The reconciliation workbook includes results, matched serials, unverified rows, and exclusions.", "success", "Excel report exported");
    } catch (error) {
      showToast(error.message || "The Excel report could not be generated.", "error", "Export failed");
    }
  }

  function downloadTemplate(kind) {
    if (!XLSX_LIB) {
      showToast("The Excel library did not load. Refresh the page and try again.", "error");
      return;
    }
    try {
      const isPurchase = kind === "purchase";
      const headers = isPurchase
        ? ["Product Name", "Serial Number", "Supplier", "Purchase Date", "Invoice Number", "Notes"]
        : ["Product Name", "Serial Number", "Supplier", "Location", "Notes"];
      const instructions = [
        [isPurchase ? "UltraPC — DIP HOSON Purchases Template" : "UltraPC — Warehouse Stock Template"],
        [],
        ["How to use this template"],
        ["1", "Keep the header row and replace the sample-free rows with your data."],
        ["2", "Use one row per serialized unit. Do not put multiple serials in one cell."],
        ["3", "Format the Serial Number column as Text before pasting long serials."],
        ["4", "Do not merge cells. Keep the Product Name spelling consistent across both files."],
        ["5", isPurchase ? "If a Supplier column is present, only rows matching DIP HOSON are used." : "The Supplier column is optional. A conflicting supplier on a matched serial is flagged, never counted."],
        ["6", "Missing, invalid, duplicate, or conflicting serials will be marked Unverified and excluded."],
        [],
        ["Columns"],
        ...headers.map((header) => [header, header === "Serial Number" ? "Required. One unique serial per row." : header === "Product Name" ? "Required. Must match the other workbook." : "Optional."])
      ];
      const workbook = XLSX_LIB.utils.book_new();
      const dataSheet = XLSX_LIB.utils.aoa_to_sheet([headers]);
      dataSheet["!cols"] = headers.map((header) => ({ wch: header === "Notes" ? 38 : header === "Serial Number" ? 28 : 24 }));
      XLSX_LIB.utils.book_append_sheet(workbook, dataSheet, isPurchase ? "DIP HOSON Purchases" : "Warehouse Stock");
      const instructionsSheet = XLSX_LIB.utils.aoa_to_sheet(instructions);
      instructionsSheet["!cols"] = [{ wch: 7 }, { wch: 95 }];
      XLSX_LIB.utils.book_append_sheet(workbook, instructionsSheet, "Instructions");
      const fileName = isPurchase ? "UltraPC_DIP_HOSON_Purchases_Template.xlsx" : "UltraPC_Warehouse_Stock_Template.xlsx";
      XLSX_LIB.writeFile(workbook, fileName, { compression: true });
      showToast(`The ${isPurchase ? "purchase" : "warehouse stock"} template is ready to fill in.`, "success", "Template downloaded");
    } catch (error) {
      showToast(error.message || "The template could not be generated.", "error", "Template failed");
    }
  }

  async function copyText(text) {
    if (!text) return false;
    try {
      if (navigator.clipboard && window.isSecureContext) {
        await navigator.clipboard.writeText(text);
        return true;
      }
    } catch (error) { /* fall through to legacy copy */ }
    try {
      const textarea = document.createElement("textarea");
      textarea.value = text;
      textarea.style.position = "fixed";
      textarea.style.opacity = "0";
      textarea.style.pointerEvents = "none";
      document.body.appendChild(textarea);
      textarea.focus();
      textarea.select();
      const copied = document.execCommand("copy");
      textarea.remove();
      return copied;
    } catch (error) {
      return false;
    }
  }

  function setupDropzone(kind) {
    const dropzone = els[`${kind}Dropzone`];
    const input = els[`${kind}Input`];
    dropzone.addEventListener("click", () => {
      input.value = "";
      input.click();
    });
    dropzone.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        input.value = "";
        input.click();
      }
    });
    input.addEventListener("change", () => {
      const file = input.files?.[0];
      if (file) handleFile(file, kind);
    });
    ["dragenter", "dragover"].forEach((eventName) => {
      dropzone.addEventListener(eventName, (event) => {
        event.preventDefault();
        event.stopPropagation();
        dropzone.classList.add("is-dragover");
      });
    });
    ["dragleave", "dragend"].forEach((eventName) => {
      dropzone.addEventListener(eventName, (event) => {
        event.preventDefault();
        dropzone.classList.remove("is-dragover");
      });
    });
    dropzone.addEventListener("drop", (event) => {
      event.preventDefault();
      event.stopPropagation();
      dropzone.classList.remove("is-dragover");
      const file = event.dataTransfer?.files?.[0];
      if (file) handleFile(file, kind);
    });
  }

  function setupFileActions(kind) {
    const input = els[`${kind}Input`];
    const replaceButton = els[`${kind}ReplaceBtn`];
    const removeButton = els[`${kind}RemoveBtn`];
    const templateButton = els[`${kind}TemplateBtn`];
    replaceButton.addEventListener("click", () => {
      input.value = "";
      input.click();
    });
    removeButton.addEventListener("click", () => removeDataset(kind));
    templateButton.addEventListener("click", () => downloadTemplate(kind));
  }

  function closeSidebar() {
    els.sidebar.classList.remove("is-open");
    els.mobileOverlay.classList.remove("is-open");
    els.menuBtn.setAttribute("aria-expanded", "false");
  }

  function openSidebar() {
    els.sidebar.classList.add("is-open");
    els.mobileOverlay.classList.add("is-open");
    els.menuBtn.setAttribute("aria-expanded", "true");
  }

  function setupScrollSpy() {
    if (!("IntersectionObserver" in window)) return;
    const links = [...document.querySelectorAll("[data-nav]")];
    const sections = links
      .map((link) => document.querySelector(link.getAttribute("href")))
      .filter(Boolean);
    if (!sections.length) return;
    const observer = new IntersectionObserver((entries) => {
      const visible = entries
        .filter((entry) => entry.isIntersecting)
        .sort((a, b) => b.intersectionRatio - a.intersectionRatio)[0];
      if (!visible) return;
      links.forEach((link) => {
        link.classList.toggle("is-active", link.getAttribute("href") === `#${visible.target.id}`);
      });
    }, { rootMargin: "-22% 0px -62% 0px", threshold: [0, 0.2, 0.5] });
    sections.forEach((section) => observer.observe(section));
  }

  function bindEvents() {
    setupDropzone("purchase");
    setupDropzone("stock");
    setupFileActions("purchase");
    setupFileActions("stock");

    els.topExportBtn.addEventListener("click", exportResults);
    els.tableExportBtn.addEventListener("click", exportResults);
    els.topResetBtn.addEventListener("click", resetAll);
    els.demoBtn.addEventListener("click", loadDemoData);

    els.productSearch.addEventListener("input", () => {
      state.filters.search = els.productSearch.value;
      state.filters.page = 1;
      renderTable();
    });
    els.statusFilter.addEventListener("change", () => {
      state.filters.status = els.statusFilter.value;
      state.filters.page = 1;
      renderTable();
    });
    els.sortSelect.addEventListener("change", () => {
      state.filters.sort = els.sortSelect.value;
      state.filters.page = 1;
      renderTable();
    });
    els.reviewSearch.addEventListener("input", () => {
      state.reviewFilters.search = els.reviewSearch.value;
      renderReview();
    });
    els.reviewSourceFilter.addEventListener("change", () => {
      state.reviewFilters.source = els.reviewSourceFilter.value;
      renderReview();
    });

    els.resultsBody.addEventListener("click", (event) => {
      const detailsTrigger = event.target.closest("[data-details]");
      if (detailsTrigger) {
        openDetails(detailsTrigger.dataset.details);
        return;
      }
      const row = event.target.closest("tr[data-product-key]");
      if (row) openDetails(row.dataset.productKey);
    });
    els.resultsBody.addEventListener("keydown", (event) => {
      if (event.key !== "Enter" && event.key !== " ") return;
      const row = event.target.closest("tr[data-product-key]");
      if (!row || event.target.closest("button")) return;
      event.preventDefault();
      openDetails(row.dataset.productKey);
    });
    els.pagination.addEventListener("click", (event) => {
      const button = event.target.closest("[data-page]");
      if (!button || button.disabled) return;
      state.filters.page = Number(button.dataset.page);
      renderTable();
      document.getElementById("reconciliation").scrollIntoView({ behavior: "smooth", block: "start" });
    });

    els.menuBtn.addEventListener("click", () => {
      if (els.sidebar.classList.contains("is-open")) closeSidebar();
      else openSidebar();
    });
    els.mobileOverlay.addEventListener("click", closeSidebar);
    document.querySelectorAll("[data-nav]").forEach((link) => {
      link.addEventListener("click", () => {
        document.querySelectorAll("[data-nav]").forEach((item) => item.classList.remove("is-active"));
        link.classList.add("is-active");
        closeSidebar();
      });
    });
    document.addEventListener("keydown", (event) => {
      if (event.key === "Escape") closeSidebar();
    });

    els.detailsContent.addEventListener("click", async (event) => {
      if (event.target.closest("[data-close-dialog]")) {
        closeDetails();
        return;
      }
      if (event.target.closest("#copySerialsBtn")) {
        const result = state.results.find((item) => item.productKey === state.detailsKey);
        if (!result) return;
        const copied = await copyText(result.matchingSerials.join("\n"));
        showToast(copied ? `${formatInt(result.matchingSerials.length)} matching serials copied to the clipboard.` : "Clipboard access was blocked by the browser.", copied ? "success" : "warning", copied ? "Serials copied" : "Copy failed");
      }
    });
    els.detailsDialog.addEventListener("click", (event) => {
      if (event.target === els.detailsDialog) closeDetails();
    });
    els.detailsDialog.addEventListener("close", () => {
      document.body.classList.remove("modal-open");
      state.detailsKey = null;
    });

    els.confirmCancelBtn.addEventListener("click", closeConfirm);
    els.confirmActionBtn.addEventListener("click", () => {
      const action = state.pendingConfirm;
      state.pendingConfirm = null;
      if (els.confirmDialog.open) els.confirmDialog.close();
      document.body.classList.remove("modal-open");
      if (typeof action === "function") action();
    });
    els.confirmDialog.addEventListener("click", (event) => {
      if (event.target === els.confirmDialog) closeConfirm();
    });
    els.confirmDialog.addEventListener("close", () => {
      state.pendingConfirm = null;
      document.body.classList.remove("modal-open");
    });

    document.addEventListener("dragover", (event) => event.preventDefault());
    document.addEventListener("drop", (event) => {
      if (!event.target.closest(".dropzone")) event.preventDefault();
    });

    setupScrollSpy();
  }

  function setTodayChip() {
    try {
      els.todayChip.textContent = new Intl.DateTimeFormat("en-GB", {
        weekday: "short",
        day: "2-digit",
        month: "short",
        year: "numeric"
      }).format(new Date());
    } catch (error) {
      els.todayChip.textContent = new Date().toLocaleDateString();
    }
  }

  function init() {
    cacheElements();
    setTodayChip();
    bindEvents();
    restoreState();
    refreshAll();
    if (state.restored) {
      setTimeout(() => {
        showToast("Your previous imported files were restored from this browser.", "info", "Session restored");
      }, 450);
    }
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
