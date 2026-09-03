const HEADER_ALIASES = {
  code: ["code", "item code", "id"],
  description: ["description", "item description", "details"],
  quantity: ["quantity", "qty"],
  uom: ["unit", "uom", "unit of measure"],
  rate: ["rate", "unit rate", "price"],
  amount: ["amount", "total", "value"],
  category: ["trade", "category"],
  notes: ["notes", "remarks", "comment"],
};

const PREVIEW_COLUMNS = [
  "Code",
  "Description",
  "Quantity",
  "UoM",
  "Rate",
  "Amount",
  "Category",
  "Notes",
];

const WONDERBUILD_COLUMNS = [
  "Item Number",
  "Category",
  "Costing Item",
  "Location",
  "Quantity",
  "Wastage (%)",
  "Rounding",
  "UOM",
  "Cost (ex.)",
  "Cost Type",
  "Markup (%)",
  "Allowance Type (PC/PS)",
  "GST Free (Yes/No)",
  "Note",
  "SKU",
];

const BUILDXACT_COLUMNS = [
  "Category",
  "Description",
  "Quantity",
  "UOM",
  "UnitCost",
  "ItemType",
  "MarkupPercent",
  "ItemCode",
];

const DEFAULT_DESCRIPTION_MAX_LENGTH = 250;
const PREVIEW_ROW_LIMIT = 200;
const ANZSMM_TEMPLATE_URL =
  "./templates/ANZSMM%202022%20-%20TEMPLATE_V1.0.xlsx";
const ANZSMM_TEMPLATE_SHEET_START_INDEX = 1;

const state = {
  file: null,
  results: null,
  download: null,
};

const elements = {
  form: document.getElementById("converter-form"),
  fileInput: document.getElementById("file-input"),
  exportFormat: document.getElementById("export-format"),
  markupPercent: document.getElementById("markup-percent"),
  descriptionMaxLength: document.getElementById("description-max-length"),
  wonderbuildGstFree: document.getElementById("wonderbuild-gst-free"),
  exporterSettings: document.querySelectorAll("[data-exporter-settings]"),
  downloadButton: document.getElementById("download-button"),
  clearButton: document.getElementById("clear-button"),
  statusMessage: document.getElementById("status-message"),
  rowsRead: document.getElementById("rows-read"),
  rowsConverted: document.getElementById("rows-converted"),
  warningCount: document.getElementById("warning-count"),
  rejectedCount: document.getElementById("rejected-count"),
  warningsList: document.getElementById("warnings-list"),
  rejectionsList: document.getElementById("rejections-list"),
  previewHead: document.getElementById("preview-head"),
  previewBody: document.getElementById("preview-body"),
};

elements.fileInput.addEventListener("change", (event) => {
  state.file = event.target.files?.[0] ?? null;
  resetDownload();
  setStatus(
    state.file
      ? `Selected file: ${state.file.name}`
      : "Select a spreadsheet to begin.",
    "default"
  );
});

elements.exportFormat.addEventListener("change", () => {
  updateExporterSettingsVisibility();
});

elements.form.addEventListener("submit", async (event) => {
  event.preventDefault();

  if (!state.file) {
    setStatus("Please choose a .xlsx, .xls, or .csv file first.", "error");
    return;
  }

  if (typeof XLSX === "undefined") {
    setStatus("Spreadsheet library failed to load. Please refresh and try again.", "error");
    return;
  }

  try {
    setStatus("Reading workbook and converting rows...", "default");
    const formatKey = elements.exportFormat.value;
    const parsed =
      formatKey === "anzsmm"
        ? await parseAnzsmmInputFile(state.file)
        : await parseInputFile(state.file);
    const exportOptions = getExportOptions();
    const exported = await exportRows(
      parsed.convertedRows,
      formatKey,
      state.file.name,
      exportOptions
    );

    if (exported.warnings?.length) {
      parsed.warnings.push(...exported.warnings);
      parsed.summary.warningRows += exported.warnings.length;
    }

    state.results = parsed;
    state.download = exported;

    renderResults(parsed, exported.previewRows);
    setStatus(
      `Converted ${parsed.summary.convertedRows} rows for ${getExporterLabel(formatKey)}.`,
      "success"
    );
    elements.downloadButton.disabled = false;
  } catch (error) {
    console.error(error);
    resetDownload();
    renderResults(createEmptyResults(), []);
    setStatus(error.message || "The file could not be parsed.", "error");
  }
});

elements.downloadButton.addEventListener("click", () => {
  if (!state.download) {
    return;
  }

  saveConvertedFile(state.download).catch((error) => {
    console.error(error);
    setStatus(
      "Could not open a save location picker, so the browser used the normal download flow instead.",
      "default"
    );
    downloadBlob(state.download.blob, state.download.fileName);
  });
});

elements.clearButton.addEventListener("click", () => {
  resetApp();
});

updateExporterSettingsVisibility();

function setStatus(message, tone) {
  elements.statusMessage.textContent = message;
  elements.statusMessage.className = "status-message";

  if (tone === "error") {
    elements.statusMessage.classList.add("status-error");
  }

  if (tone === "success") {
    elements.statusMessage.classList.add("status-success");
  }
}

async function parseInputFile(file) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, {
    type: "array",
    cellDates: true,
    raw: true,
  });

  if (!workbook.SheetNames.length) {
    throw new Error("No worksheet was found in the uploaded file.");
  }

  const sheetContexts = buildSheetContexts(workbook);

  if (!sheetContexts.length) {
    throw new Error("No readable trade worksheets were found in the uploaded file.");
  }

  const transformed = transformWorkbookSheets(sheetContexts);

  return {
    sheetNames: sheetContexts.map((sheet) => sheet.sheetName),
    convertedRows: transformed.convertedRows,
    warnings: transformed.warningMessages,
    rejectedRows: transformed.rejectedRows,
    summary: {
      rowsRead: transformed.rowsRead,
      convertedRows: transformed.convertedRows.length,
      warningRows: transformed.warningRowCount,
      rejectedRows: transformed.rejectedRows.length,
    },
  };
}

async function parseAnzsmmInputFile(file) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, {
    type: "array",
    cellDates: true,
    raw: true,
  });

  if (!workbook.SheetNames.length) {
    throw new Error("No worksheet was found in the uploaded file.");
  }

  const sheetContexts = buildSheetContexts(workbook);
  if (!sheetContexts.length) {
    throw new Error("No readable CostX trade worksheets were found in the uploaded file.");
  }

  const transformed = transformAnzsmmSheets(sheetContexts);

  return {
    sheetNames: sheetContexts.map((sheet) => sheet.sheetName),
    convertedRows: transformed.convertedRows,
    warnings: transformed.warningMessages,
    rejectedRows: transformed.rejectedRows,
    summary: {
      rowsRead: transformed.rowsRead,
      convertedRows: transformed.convertedRows.length,
      warningRows: transformed.warningRowCount,
      rejectedRows: transformed.rejectedRows.length,
    },
  };
}

function transformAnzsmmSheets(sheetContexts) {
  const aggregate = {
    rowsRead: 0,
    convertedRows: [],
    warningMessages: [],
    warningRowCount: 0,
    rejectedRows: [],
  };

  sheetContexts.forEach((sheet) => {
    sheet.dataRows.forEach((row, index) => {
      const sourceRowNumber = sheet.headerRowIndex + 2 + index;

      if (isRowCompletelyBlank(row)) {
        return;
      }

      aggregate.rowsRead += 1;
      const rawRecord = buildRawRecord(row, sheet.headerMap);

      if (isTotalOnlyRow(rawRecord)) {
        aggregate.rejectedRows.push({
          sheetName: sheet.sheetName,
          sourceRowNumber,
          reason: "Skipped total-only row.",
          rawRecord,
        });
        return;
      }

      const warnings = [];
      if (!hasCellValue(rawRecord.code)) {
        warnings.push("Missing code");
      }
      if (!hasCellValue(rawRecord.description)) {
        warnings.push("Missing description");
      }

      if (warnings.length) {
        aggregate.warningRowCount += 1;
        aggregate.warningMessages.push(
          `${sheet.sheetName} row ${sourceRowNumber}: ${warnings.join("; ")}`
        );
      }

      aggregate.convertedRows.push({
        code: cleanWorkbookCellValue(rawRecord.code),
        description: cleanWorkbookCellValue(rawRecord.description),
        quantity: cleanWorkbookCellValue(rawRecord.quantity),
        uom: cleanWorkbookCellValue(rawRecord.uom),
        rate: cleanWorkbookCellValue(rawRecord.rate),
        amount: cleanWorkbookCellValue(rawRecord.amount),
        category: normaliseSheetCategory(sheet.sheetName),
        source_sheet_name: sheet.sheetName,
        source_row_number: sourceRowNumber,
        template_sheet_key: normaliseTradeKey(sheet.sheetName),
        warnings,
      });
    });
  });

  return aggregate;
}

function buildSheetContexts(workbook) {
  const sheetNames = workbook.SheetNames;
  const candidateNames =
    sheetNames.length > 1 ? sheetNames.slice(1) : sheetNames;

  return candidateNames
    .map((sheetName, sheetIndex) => {
      const worksheet = workbook.Sheets[sheetName];
      const matrix = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
        raw: false,
        defval: "",
        // Preserve worksheet row positions so reported source row numbers
        // stay aligned even when the CostX sheet contains separator rows.
        blankrows: true,
      });

      if (!matrix.length) {
        return null;
      }

      const { headerRowIndex, headerRow, dataRows } = detectHeaderRow(matrix);
      const headerMap = buildHeaderMap(headerRow);

      return {
        sheetName,
        sheetOrder: sheetIndex + 1,
        headerRowIndex,
        headerMap,
        dataRows,
      };
    })
    .filter(Boolean);
}

function detectHeaderRow(matrix) {
  let bestIndex = 0;
  let bestScore = -1;

  matrix.slice(0, 15).forEach((row, index) => {
    const score = row.reduce((count, cell) => {
      return isRecognisedHeader(cell) ? count + 1 : count;
    }, 0);

    if (score > bestScore) {
      bestScore = score;
      bestIndex = index;
    }
  });

  const headerRow = matrix[bestIndex] || [];
  const dataRows = matrix.slice(bestIndex + 1);

  return { headerRowIndex: bestIndex, headerRow, dataRows };
}

function isRecognisedHeader(value) {
  const normalised = normaliseHeaderValue(value);

  return Object.values(HEADER_ALIASES).some((aliases) =>
    aliases.some((alias) => normaliseHeaderValue(alias) === normalised)
  );
}

function buildHeaderMap(headerRow) {
  const map = {};

  headerRow.forEach((header, index) => {
    const canonicalKey = findCanonicalHeaderKey(header);
    if (canonicalKey && map[canonicalKey] === undefined) {
      map[canonicalKey] = index;
    }
  });

  return map;
}

function findCanonicalHeaderKey(header) {
  const normalised = normaliseHeaderValue(header);

  if (!normalised) {
    return null;
  }

  return (
    Object.keys(HEADER_ALIASES).find((key) =>
      HEADER_ALIASES[key].some((alias) => normaliseHeaderValue(alias) === normalised)
    ) || null
  );
}

function normaliseHeaderValue(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/[\s_-]+/g, " ")
    .replace(/[^\w\s]/g, "");
}

function transformWorkbookSheets(sheetContexts) {
  const aggregate = {
    rowsRead: 0,
    convertedRows: [],
    warningMessages: [],
    warningRowCount: 0,
    rejectedRows: [],
  };

  sheetContexts.forEach((sheet) => {
    const transformed = transformRows(
      sheet.dataRows,
      sheet.headerMap,
      sheet.headerRowIndex + 2,
      sheet.sheetName
    );

    aggregate.rowsRead += transformed.rowsRead;
    aggregate.convertedRows.push(...transformed.convertedRows);
    aggregate.warningMessages.push(...transformed.warningMessages);
    aggregate.warningRowCount += transformed.warningRowCount;
    aggregate.rejectedRows.push(...transformed.rejectedRows);
  });

  return aggregate;
}

function transformRows(dataRows, headerMap, sourceStartRowNumber, sheetName) {
  const convertedRows = [];
  const warningMessages = [];
  const rejectedRows = [];
  let rowsRead = 0;
  let warningRowCount = 0;

  dataRows.forEach((row, index) => {
    const sourceRowNumber = sourceStartRowNumber + index;

    if (isRowCompletelyBlank(row)) {
      return;
    }

    rowsRead += 1;

    const rawRecord = buildRawRecord(row, headerMap);

    if (isTotalOnlyRow(rawRecord)) {
      rejectedRows.push({
        sheetName,
        sourceRowNumber,
        reason: "Skipped total-only row.",
        rawRecord,
      });
      return;
    }

    if (isSectionHeaderRow(rawRecord)) {
      return;
    }

    const canonicalRow = createCanonicalRow(rawRecord, sourceRowNumber, sheetName);
    const validation = validateCanonicalRow(canonicalRow);

    canonicalRow.warnings = validation.warnings;

    if (validation.rejectReason) {
      rejectedRows.push({
        sheetName,
        sourceRowNumber,
        reason: validation.rejectReason,
        rawRecord,
      });
      return;
    }

    if (validation.warnings.length) {
      warningRowCount += 1;
      warningMessages.push(
        `${sheetName} row ${sourceRowNumber}: ${validation.warnings.join("; ")}`
      );
    }

    convertedRows.push(canonicalRow);
  });

  return {
    rowsRead,
    convertedRows,
    warningMessages,
    warningRowCount,
    rejectedRows,
  };
}

function buildRawRecord(row, headerMap) {
  return {
    code: getCellValue(row, headerMap.code),
    description: getCellValue(row, headerMap.description),
    quantity: getCellValue(row, headerMap.quantity),
    uom: getCellValue(row, headerMap.uom),
    rate: getCellValue(row, headerMap.rate),
    amount: getCellValue(row, headerMap.amount),
    category: getCellValue(row, headerMap.category),
    notes: getCellValue(row, headerMap.notes),
  };
}

function getCellValue(row, index) {
  if (index === undefined) {
    return "";
  }

  return row[index] ?? "";
}

function createCanonicalRow(rawRecord, sourceRowNumber, sheetName) {
  const description = cleanText(rawRecord.description);
  const category = cleanText(rawRecord.category) || normaliseSheetCategory(sheetName);

  return {
    code: cleanText(rawRecord.code),
    description,
    category,
    subcategory: "",
    quantity: parseNumericValue(rawRecord.quantity),
    uom: cleanText(rawRecord.uom),
    rate: parseNumericValue(rawRecord.rate),
    amount: parseNumericValue(rawRecord.amount),
    notes: cleanText(rawRecord.notes),
    location: "",
    source_sheet_name: sheetName,
    source_row_number: sourceRowNumber,
    warnings: [],
  };
}

function cleanText(value) {
  return String(value ?? "").trim().replace(/\s+/g, " ");
}

function cleanWorkbookCellValue(value) {
  if (typeof value === "string") {
    return value.trim();
  }

  return value ?? "";
}

function hasCellValue(value) {
  return cleanText(value) !== "";
}

function truncateText(value, maxLength) {
  return value.length > maxLength ? value.slice(0, maxLength) : value;
}

function parseNumericValue(value) {
  const text = cleanText(value);

  if (!text) {
    return null;
  }

  const normalised = text.replace(/,/g, "");
  const parsed = Number(normalised);

  return Number.isFinite(parsed) ? parsed : text;
}

function normaliseSheetCategory(sheetName) {
  return String(sheetName || "")
    .replace(/^\d+\s*-\s*/, "")
    .trim();
}

function normaliseTradeKey(sheetName) {
  return normaliseSheetCategory(sheetName)
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, " ")
    .trim();
}

function isRowCompletelyBlank(row) {
  return row.every((cell) => cleanText(cell) === "");
}

function isSummaryRow(rawRecord) {
  const combined = [
    rawRecord.code,
    rawRecord.description,
    rawRecord.category,
    rawRecord.notes,
  ]
    .map(cleanText)
    .join(" ")
    .toLowerCase();

  if (!combined) {
    return true;
  }

  const summaryMarkers = ["subtotal", "sub total", "total", "grand total"];
  const looksLikeSummary = summaryMarkers.some((marker) => combined.includes(marker));
  const missingCoreValues =
    cleanText(rawRecord.quantity) === "" &&
    cleanText(rawRecord.rate) === "" &&
    cleanText(rawRecord.amount) === "";

  return looksLikeSummary || (missingCoreValues && cleanText(rawRecord.description) === "");
}

function isTotalOnlyRow(rawRecord) {
  const labels = [
    cleanText(rawRecord.code),
    cleanText(rawRecord.description),
    cleanText(rawRecord.category),
    cleanText(rawRecord.notes),
  ].filter(Boolean);

  if (!labels.length) {
    return false;
  }

  if (labels.length > 1) {
    return false;
  }

  const onlyLabel = labels[0].toLowerCase();
  return ["total", "grand total", "subtotal", "sub total"].includes(onlyLabel);
}

function isSectionHeaderRow(rawRecord) {
  const code = cleanText(rawRecord.code);
  const description = cleanText(rawRecord.description);
  const quantity = cleanText(rawRecord.quantity);
  const uom = cleanText(rawRecord.uom);
  const rate = cleanText(rawRecord.rate);
  const amount = cleanText(rawRecord.amount);
  const notes = cleanText(rawRecord.notes);
  const category = cleanText(rawRecord.category);

  if (!description) {
    return false;
  }

  const hasOnlyLabelContent =
    quantity === "" &&
    uom === "" &&
    rate === "" &&
    amount === "" &&
    notes === "" &&
    category === "";

  return hasOnlyLabelContent;
}

function validateCanonicalRow(canonicalRow) {
  const warnings = [];

  if (canonicalRow.quantity === null || canonicalRow.quantity === "") {
    warnings.push("Missing quantity");
  } else if (typeof canonicalRow.quantity !== "number") {
    warnings.push("Invalid quantity");
  }

  if (canonicalRow.rate !== null && typeof canonicalRow.rate !== "number") {
    warnings.push("Invalid rate");
  }

  if (canonicalRow.amount !== null && typeof canonicalRow.amount !== "number") {
    warnings.push("Invalid amount");
  }

  if (!canonicalRow.code || !canonicalRow.description) {
    return {
      warnings,
      rejectReason: "Code and description are required for Wonderbuild import.",
    };
  }

  return { warnings, rejectReason: null };
}

async function exportRows(canonicalRows, formatKey, originalFileName, exportOptions) {
  const exporters = {
    generic: exportGenericPmImport,
    buildxact: exportBuildxact,
    wonderbuild: exportWonderbuild,
    anzsmm: exportAnzsmmWorkbook,
  };

  const exporter = exporters[formatKey] || exportGenericPmImport;
  return exporter(canonicalRows, originalFileName, exportOptions);
}

async function exportAnzsmmWorkbook(rows, originalFileName) {
  if (typeof ExcelJS === "undefined") {
    throw new Error(
      "The template workbook library failed to load. Please refresh and try again."
    );
  }

  const response = await fetch(ANZSMM_TEMPLATE_URL);
  if (!response.ok) {
    throw new Error(
      "The ANZSMM template could not be loaded. Please run the app from a local web server or GitHub Pages."
    );
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await response.arrayBuffer());

  const summarySheet = workbook.getWorksheet("Summary");
  if (summarySheet) {
    summarySheet.getCell("A1").value = stripExtension(originalFileName);
  }

  const rowsByTemplateSheet = groupRowsByTemplateSheet(rows);
  const templateSheets = workbook.worksheets.slice(
    ANZSMM_TEMPLATE_SHEET_START_INDEX
  );
  const templateSheetKeys = new Map(
    templateSheets.map((sheet) => [normaliseTradeKey(sheet.name), sheet])
  );
  const warnings = [];

  templateSheets.forEach((sheet) => {
    if (sheet.name === "Sheet1") {
      return;
    }

    const targetRows = rowsByTemplateSheet.get(normaliseTradeKey(sheet.name)) || [];
    populateAnzsmmTemplateSheet(sheet, targetRows);
    restoreTemplateHomeLink(sheet);
  });

  restoreTemplateSummaryLinks(summarySheet, workbook);

  rowsByTemplateSheet.forEach((_, key) => {
    if (!templateSheetKeys.has(key)) {
      const sourceName = rows.find((row) => row.template_sheet_key === key)
        ?.source_sheet_name;
      warnings.push(
        `${sourceName || key}: no matching ANZSMM template tab was found.`
      );
    }
  });

  // Excel recalculates the preserved summary and trade formulas when opened.
  workbook.calcProperties.fullCalcOnLoad = true;
  workbook.calcProperties.forceFullCalc = true;

  const arrayBuffer = await workbook.xlsx.writeBuffer();
  const previewRows = rows.slice(0, PREVIEW_ROW_LIMIT).map((row) => ({
    Code: row.code,
    Description: row.description,
    Quantity: row.quantity,
    UoM: row.uom,
    Rate: row.rate,
    Amount: row.amount,
    Category: row.category,
    Notes: "",
  }));

  return {
    blob: new Blob([arrayBuffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }),
    fileName: `${stripExtension(originalFileName)}_iqs.xlsx`,
    previewRows,
    previewColumns: PREVIEW_COLUMNS,
    warnings,
  };
}

function groupRowsByTemplateSheet(rows) {
  return rows.reduce((groups, row) => {
    const existing = groups.get(row.template_sheet_key) || [];
    existing.push(row);
    groups.set(row.template_sheet_key, existing);
    return groups;
  }, new Map());
}

function populateAnzsmmTemplateSheet(sheet, rows) {
  const dataStyle = captureRowStyle(sheet.getRow(3));
  const totalStyle = captureRowStyle(sheet.getRow(Math.min(sheet.rowCount, 4)));
  const dataRowHeight = sheet.getRow(3).height;
  const totalRowHeight = sheet.getRow(Math.min(sheet.rowCount, 4)).height;

  // The template's first two rows hold the trade title, HOME link, and headers.
  // Rebuilding from row 3 removes sample rows while preserving those controls.
  if (sheet.rowCount > 2) {
    sheet.spliceRows(3, sheet.rowCount - 2);
  }

  rows.forEach((sourceRow, index) => {
    const rowNumber = index + 3;
    const targetRow = sheet.getRow(rowNumber);
    applyRowStyle(targetRow, dataStyle, dataRowHeight);

    targetRow.getCell(1).value = sourceRow.code;
    targetRow.getCell(2).value = sourceRow.description;
    targetRow.getCell(3).value = sourceRow.quantity;
    targetRow.getCell(4).value = sourceRow.uom;
    targetRow.getCell(5).value = sourceRow.rate;

    if (isNumericCellValue(sourceRow.quantity) && isNumericCellValue(sourceRow.rate)) {
      targetRow.getCell(6).value = { formula: `C${rowNumber}*E${rowNumber}` };
    } else {
      // Do not retain the template's sample formula for a descriptive-only row.
      targetRow.getCell(6).value = null;
    }

    targetRow.commit();
  });

  if (rows.length) {
    const totalRowNumber = rows.length + 3;
    const totalRow = sheet.getRow(totalRowNumber);
    applyRowStyle(totalRow, totalStyle, totalRowHeight);
    totalRow.getCell(2).value = "Total";
    totalRow.getCell(6).value = {
      formula: `SUM(F2:F${Math.max(totalRowNumber - 1, 2)})`,
    };
    totalRow.commit();
  }
}

function restoreTemplateHomeLink(sheet) {
  sheet.getCell("E1").value = {
    text: "HOME",
    hyperlink: "#'Summary'!A1",
  };
}

function restoreTemplateSummaryLinks(summarySheet, workbook) {
  if (!summarySheet) {
    return;
  }

  summarySheet.eachRow((row) => {
    const totalFormula = row.getCell(3).formula;
    const targetSheetName = getFormulaSheetReference(totalFormula);
    const labelCell = row.getCell(2);

    if (!targetSheetName || !workbook.getWorksheet(targetSheetName) || !labelCell.value) {
      return;
    }

    labelCell.value = {
      text: String(labelCell.value),
      hyperlink: `#'${targetSheetName}'!A1`,
    };
  });
}

function getFormulaSheetReference(formula) {
  const match = String(formula || "").match(/'([^']+)'!/);
  return match?.[1] || null;
}

function captureRowStyle(row) {
  return Array.from({ length: 6 }, (_, index) => {
    const cell = row.getCell(index + 1);
    return {
      style: cloneExcelValue(cell.style),
      fill: cloneExcelValue(cell.fill),
      font: cloneExcelValue(cell.font),
      border: cloneExcelValue(cell.border),
      alignment: cloneExcelValue(cell.alignment),
      protection: cloneExcelValue(cell.protection),
      numFmt: cell.numFmt,
    };
  });
}

function applyRowStyle(row, cellStyles, height) {
  if (height) {
    row.height = height;
  }

  cellStyles.forEach((sourceStyle, index) => {
    const cell = row.getCell(index + 1);
    cell.style = cloneExcelValue(sourceStyle.style);
    cell.fill = cloneExcelValue(sourceStyle.fill);
    cell.font = cloneExcelValue(sourceStyle.font);
    cell.border = cloneExcelValue(sourceStyle.border);
    cell.alignment = cloneExcelValue(sourceStyle.alignment);
    cell.protection = cloneExcelValue(sourceStyle.protection);
    cell.numFmt = sourceStyle.numFmt;
  });
}

function cloneExcelValue(value) {
  return value ? JSON.parse(JSON.stringify(value)) : value;
}

function isNumericCellValue(value) {
  if (typeof value === "number") {
    return Number.isFinite(value);
  }

  return /^-?\d+(?:\.\d+)?$/.test(String(value ?? "").trim().replace(/,/g, ""));
}

function exportGenericPmImport(canonicalRows, originalFileName, exportOptions) {
  const rows = canonicalRows.map((row) => ({
    Code: sanitiseExportText(row.code),
    Description: sanitiseExportText(
      truncateText(row.description, exportOptions.descriptionMaxLength)
    ),
    Quantity: row.quantity ?? "",
    UoM: sanitiseExportText(row.uom),
    Rate: row.rate ?? "",
    Amount: row.amount ?? "",
    Category: sanitiseExportText(row.category),
    Notes: sanitiseExportText(row.notes),
  }));

  const worksheet = XLSX.utils.json_to_sheet(rows, { header: PREVIEW_COLUMNS });
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Generic Import");

  const arrayBuffer = XLSX.write(workbook, {
    bookType: "xlsx",
    type: "array",
  });

  return {
    blob: new Blob([arrayBuffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }),
    fileName: `${stripExtension(originalFileName)}_converted_generic.xlsx`,
    previewRows: rows.slice(0, PREVIEW_ROW_LIMIT),
    previewColumns: PREVIEW_COLUMNS,
  };
}

function exportBuildxact(canonicalRows, originalFileName, exportOptions) {
  const rows = canonicalRows.map((row) => ({
    Category: sanitiseExportText(row.category),
    Description: sanitiseExportText(
      truncateText(row.description, exportOptions.descriptionMaxLength)
    ),
    Quantity: row.quantity ?? "",
    UOM: sanitiseExportText(row.uom),
    UnitCost: row.rate ?? "",
    ItemType: "Material",
    MarkupPercent: exportOptions.markupPercent,
    ItemCode: sanitiseExportText(row.code),
  }));

  const worksheet = XLSX.utils.json_to_sheet(rows, { header: BUILDXACT_COLUMNS });
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Buildxact Import");

  const arrayBuffer = XLSX.write(workbook, {
    bookType: "xlsx",
    type: "array",
  });

  return {
    blob: new Blob([arrayBuffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }),
    fileName: `${stripExtension(originalFileName)}_BX.xlsx`,
    previewRows: rows.slice(0, PREVIEW_ROW_LIMIT),
    previewColumns: BUILDXACT_COLUMNS,
  };
}

function exportWonderbuild(canonicalRows, originalFileName, exportOptions) {
  const rows = canonicalRows.map((row) => ({
    "Item Number": sanitiseExportText(row.code),
    Category: sanitiseExportText(row.category),
    "Costing Item": sanitiseExportText(
      truncateText(row.description, exportOptions.descriptionMaxLength)
    ),
    Location: sanitiseExportText(row.location),
    Quantity: row.quantity ?? "",
    "Wastage (%)": "",
    Rounding: "",
    UOM: sanitiseExportText(row.uom),
    "Cost (ex.)": row.rate ?? "",
    "Cost Type": "",
    "Markup (%)": exportOptions.markupPercent,
    "Allowance Type (PC/PS)": "",
    "GST Free (Yes/No)": exportOptions.wonderbuildGstFree,
    Note: sanitiseExportText(row.notes),
    SKU: "",
  }));

  const worksheet = XLSX.utils.json_to_sheet(rows, { header: WONDERBUILD_COLUMNS });
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Wonderbuild Import");

  const arrayBuffer = XLSX.write(workbook, {
    bookType: "xlsx",
    type: "array",
  });

  return {
    blob: new Blob([arrayBuffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }),
    fileName: `${stripExtension(originalFileName)}_WB.xlsx`,
    previewRows: rows.slice(0, PREVIEW_ROW_LIMIT),
    previewColumns: WONDERBUILD_COLUMNS,
  };
}

function exportPlaceholderTemplate(canonicalRows, originalFileName, suffix, note, exportOptions) {
  const rows = canonicalRows.map((row) => ({
    Code: sanitiseExportText(row.code),
    Description: sanitiseExportText(
      truncateText(row.description, exportOptions.descriptionMaxLength)
    ),
    Quantity: row.quantity ?? "",
    UoM: sanitiseExportText(row.uom),
    Rate: row.rate ?? "",
    Amount: row.amount ?? "",
    Category: sanitiseExportText(row.category),
    Notes: sanitiseExportText(row.notes ? `${row.notes} | ${note}` : note),
  }));

  const worksheet = XLSX.utils.json_to_sheet(rows, { header: PREVIEW_COLUMNS });
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Placeholder Export");

  const arrayBuffer = XLSX.write(workbook, {
    bookType: "xlsx",
    type: "array",
  });

  return {
    blob: new Blob([arrayBuffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }),
    fileName: `${stripExtension(originalFileName)}_converted_${suffix}.xlsx`,
    previewRows: rows.slice(0, PREVIEW_ROW_LIMIT),
    previewColumns: PREVIEW_COLUMNS,
  };
}

function stripExtension(fileName) {
  return fileName.replace(/\.[^.]+$/, "");
}

function getExportOptions() {
  return {
    markupPercent: parseOptionalNumber(elements.markupPercent.value),
    descriptionMaxLength:
      parseOptionalPositiveInteger(elements.descriptionMaxLength.value) ||
      DEFAULT_DESCRIPTION_MAX_LENGTH,
    wonderbuildGstFree: elements.wonderbuildGstFree.value || "No",
  };
}

function parseOptionalNumber(value) {
  const text = cleanText(value);
  if (!text) {
    return "";
  }

  const parsed = Number(text);
  return Number.isFinite(parsed) ? parsed : "";
}

function parseOptionalPositiveInteger(value) {
  const text = cleanText(value);
  if (!text) {
    return null;
  }

  const parsed = Number(text);
  if (!Number.isInteger(parsed) || parsed < 1) {
    return null;
  }

  return parsed;
}

function sanitiseExportText(value) {
  return cleanText(value).replace(/,/g, "");
}

function renderResults(results, previewRows) {
  elements.rowsRead.textContent = String(results.summary.rowsRead);
  elements.rowsConverted.textContent = String(results.summary.convertedRows);
  elements.warningCount.textContent = String(results.summary.warningRows);
  elements.rejectedCount.textContent = String(results.summary.rejectedRows);

  renderMessageList(
    elements.warningsList,
    results.warnings,
    "No warnings generated."
  );
  renderMessageList(
    elements.rejectionsList,
    results.rejectedRows.map(
      (entry) => `${entry.sheetName} row ${entry.sourceRowNumber}: ${entry.reason}`
    ),
    "No rejected rows."
  );
  renderPreviewTable(previewRows, state.download?.previewColumns || PREVIEW_COLUMNS);
}

function renderMessageList(target, items, emptyMessage) {
  target.innerHTML = "";

  if (!items.length) {
    const listItem = document.createElement("li");
    listItem.textContent = emptyMessage;
    target.appendChild(listItem);
    return;
  }

  items.forEach((item) => {
    const listItem = document.createElement("li");
    listItem.textContent = item;
    target.appendChild(listItem);
  });
}

function renderPreviewTable(rows, columns) {
  elements.previewHead.innerHTML = "";
  elements.previewBody.innerHTML = "";

  const headRow = document.createElement("tr");
  columns.forEach((column) => {
    const cell = document.createElement("th");
    cell.textContent = column;
    headRow.appendChild(cell);
  });
  elements.previewHead.appendChild(headRow);

  if (!rows.length) {
    const row = document.createElement("tr");
    const cell = document.createElement("td");
    cell.colSpan = columns.length;
    cell.className = "empty-state";
    cell.textContent = "No converted rows available for preview.";
    row.appendChild(cell);
    elements.previewBody.appendChild(row);
    return;
  }

  rows.forEach((previewRow) => {
    const row = document.createElement("tr");

    columns.forEach((column) => {
      const cell = document.createElement("td");
      cell.textContent = previewRow[column] ?? "";
      row.appendChild(cell);
    });

    elements.previewBody.appendChild(row);
  });
}

function createEmptyResults() {
  return {
    summary: {
      rowsRead: 0,
      convertedRows: 0,
      warningRows: 0,
      rejectedRows: 0,
    },
    warnings: [],
    rejectedRows: [],
  };
}

function resetDownload() {
  state.download = null;
  elements.downloadButton.disabled = true;
}

function resetApp() {
  state.file = null;
  state.results = null;
  resetDownload();
  elements.form.reset();
  renderResults(createEmptyResults(), []);
  updateExporterSettingsVisibility();
  setStatus("Select a spreadsheet to begin.", "default");
}

function updateExporterSettingsVisibility() {
  const selectedExporter = elements.exportFormat.value;

  elements.exporterSettings.forEach((panel) => {
    const matches = panel.dataset.exporterSettings === selectedExporter;
    panel.hidden = !matches;
  });
}

function downloadBlob(blob, fileName) {
  const objectUrl = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = objectUrl;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(objectUrl);
}

async function saveConvertedFile(download) {
  if (!supportsFilePicker()) {
    downloadBlob(download.blob, download.fileName);
    setStatus(
      `Downloaded ${download.fileName} using the browser's default download location.`,
      "success"
    );
    return;
  }

  const extension = getFileExtension(download.fileName);
  const suggestedName = download.fileName;
  const handle = await window.showSaveFilePicker({
    suggestedName,
    types: [
      {
        description: extension === "csv" ? "CSV file" : "Excel workbook",
        accept:
          extension === "csv"
            ? { "text/csv": [".csv"] }
            : {
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"],
              },
      },
    ],
  });

  const writable = await handle.createWritable();
  await writable.write(download.blob);
  await writable.close();

  setStatus(`Saved ${download.fileName} to your chosen location.`, "success");
}

function supportsFilePicker() {
  return typeof window.showSaveFilePicker === "function";
}

function getFileExtension(fileName) {
  return fileName.split(".").pop()?.toLowerCase() || "";
}

function getExporterLabel(formatKey) {
  const labels = {
    generic: "Generic PM Import",
    buildxact: "Buildxact",
    wonderbuild: "Wonderbuild",
    anzsmm: "ANZSMM 2022 Workbook",
  };

  return labels[formatKey] || "selected export";
}
