const triggers = document.querySelectorAll("[data-target]");
const panels = document.querySelectorAll(".panel");

const mergeFilesInput = document.getElementById("merge-files");
const headerRowInput = document.getElementById("header-row");
const mergeRunButton = document.getElementById("merge-run");
const mergeStatus = document.getElementById("merge-status");
const mergeFileList = document.getElementById("merge-file-list");

const splitFilesInput = document.getElementById("split-files");
const splitHeaderRowInput = document.getElementById("split-header-row");
const splitRunButton = document.getElementById("split-run");
const splitStatus = document.getElementById("split-status");
const splitFileList = document.getElementById("split-file-list");
const splitRowSizeInput = document.getElementById("split-row-size");
const splitColumnNameInput = document.getElementById("split-column-name");
const rowSizeField = document.getElementById("row-size-field");
const columnNameField = document.getElementById("column-name-field");

function activatePanel(targetId) {
  const nextPanel = document.getElementById(targetId);

  if (!nextPanel) {
    return;
  }

  panels.forEach((panel) => {
    panel.classList.toggle("active", panel === nextPanel);
  });
}

function setStatus(element, message, tone = "") {
  element.textContent = message;
  element.className = "status-box";

  if (tone) {
    element.classList.add(tone);
  }
}

function formatFileSize(bytes) {
  if (bytes < 1024) {
    return `${bytes}B`;
  }

  if (bytes < 1024 * 1024) {
    return `${(bytes / 1024).toFixed(1)}KB`;
  }

  return `${(bytes / (1024 * 1024)).toFixed(1)}MB`;
}

function renderSelectedFiles(input, listElement, emptyMessage, readyMessage) {
  const files = Array.from(input.files || []);

  if (!files.length) {
    listElement.innerHTML = "";
    return emptyMessage;
  }

  listElement.innerHTML = files
    .map(
      (file) => `
        <div class="file-chip">
          <strong>${file.name}</strong>
          <span>${formatFileSize(file.size)}</span>
        </div>
      `
    )
    .join("");

  return `${files.length}개 파일이 선택되었습니다. ${readyMessage}`;
}

function getCheckedValue(name, fallback) {
  return document.querySelector(`input[name="${name}"]:checked`)?.value || fallback;
}

function normalizeSheetName(rawName, usedNames) {
  const baseName = String(rawName || "Sheet")
    .replace(/[\\/*?:[\]]/g, "_")
    .slice(0, 31) || "Sheet";
  let nextName = baseName;
  let suffix = 1;

  while (usedNames.has(nextName)) {
    const suffixText = `_${suffix}`;
    nextName = `${baseName.slice(0, 31 - suffixText.length)}${suffixText}`;
    suffix += 1;
  }

  usedNames.add(nextName);
  return nextName;
}

function safeFileSegment(rawValue) {
  return String(rawValue || "empty")
    .trim()
    .replace(/[<>:"/\\|?*\u0000-\u001f]/g, "_")
    .replace(/\s+/g, "_")
    .slice(0, 80) || "empty";
}

function readFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (event) => resolve(event.target.result);
    reader.onerror = () => reject(new Error(`${file.name} 파일을 읽지 못했습니다.`));
    reader.readAsArrayBuffer(file);
  });
}

function getCellDisplayFormat(cell) {
  if (!cell) {
    return "";
  }

  if (cell.z) {
    return cell.z;
  }

  return "";
}

function cloneCellValue(cell) {
  if (!cell) {
    return {
      v: "",
      t: "s",
      z: "",
    };
  }

  return {
    v: cell.v ?? "",
    t: cell.t || "s",
    z: getCellDisplayFormat(cell),
  };
}

function buildCellMatrix(sheet) {
  const ref = sheet["!ref"];

  if (!ref) {
    return [];
  }

  const range = XLSX.utils.decode_range(ref);
  const matrix = [];

  for (let rowIndex = range.s.r; rowIndex <= range.e.r; rowIndex += 1) {
    const row = [];

    for (let colIndex = range.s.c; colIndex <= range.e.c; colIndex += 1) {
      const address = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
      row.push(cloneCellValue(sheet[address]));
    }

    matrix.push(row);
  }

  return matrix;
}

function parseSheetRows(sheet, headerRowNumber) {
  const rows = buildCellMatrix(sheet);
  const headerIndex = headerRowNumber - 1;
  const headerRow = rows[headerIndex];

  if (!headerRow || !headerRow.length) {
    throw new Error("입력한 컬럼 행 번호를 해당 파일에서 찾을 수 없습니다.");
  }

  const normalizedHeaders = headerRow.map((cell, index) => {
    const value = String(cell?.v ?? "").trim();
    return value || `빈컬럼_${index + 1}`;
  });

  return {
    headers: normalizedHeaders,
    dataRows: rows.slice(headerIndex + 1),
  };
}

function buildSheetFromStructuredRows(headers, dataRows) {
  const sheet = {};
  const rowCount = dataRows.length + 1;
  const colCount = headers.length;

  headers.forEach((header, columnIndex) => {
    const address = XLSX.utils.encode_cell({ r: 0, c: columnIndex });
    sheet[address] = {
      v: header,
      t: "s",
    };
  });

  dataRows.forEach((row, rowIndex) => {
    headers.forEach((_, columnIndex) => {
      const cell = row[columnIndex] || { v: "", t: "s", z: "" };
      const address = XLSX.utils.encode_cell({ r: rowIndex + 1, c: columnIndex });

      sheet[address] = {
        v: cell.v ?? "",
        t: cell.t || "s",
      };

      if (cell.z) {
        sheet[address].z = cell.z;
      }
    });
  });

  sheet["!ref"] = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
    e: { r: Math.max(rowCount - 1, 0), c: Math.max(colCount - 1, 0) },
  });

  return sheet;
}

function createWorkbookFromRows(headers, dataRows, sheetName = "Sheet1") {
  const workbook = XLSX.utils.book_new();
  const sheet = buildSheetFromStructuredRows(headers, dataRows);
  XLSX.utils.book_append_sheet(workbook, sheet, normalizeSheetName(sheetName, new Set()));
  return workbook;
}

function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = url;
  link.download = fileName;
  link.click();
  URL.revokeObjectURL(url);
}

function buildMergedSingleSheet(parsedFiles) {
  const mergedHeaders = [];
  const headerSet = new Set();
  const mergedRows = [];

  parsedFiles.forEach(({ headers, dataRows }) => {
    headers.forEach((header) => {
      if (!headerSet.has(header)) {
        headerSet.add(header);
        mergedHeaders.push(header);
      }
    });

    dataRows.forEach((row) => {
      const rowMap = {};

      headers.forEach((header, index) => {
        rowMap[header] = row[index] || { v: "", t: "s", z: "" };
      });

      mergedRows.push(
        mergedHeaders.map((header) => rowMap[header] || { v: "", t: "s", z: "" })
      );
    });
  });

  return buildSheetFromStructuredRows(mergedHeaders, mergedRows);
}

function buildMergeWorkbook(parsedFiles, mode) {
  const workbook = XLSX.utils.book_new();

  if (mode === "single-sheet") {
    XLSX.utils.book_append_sheet(workbook, buildMergedSingleSheet(parsedFiles), "병합결과");
    return workbook;
  }

  const usedNames = new Set();

  parsedFiles.forEach(({ fileName, headers, dataRows }) => {
    const sheet = buildSheetFromStructuredRows(headers, dataRows);
    XLSX.utils.book_append_sheet(workbook, sheet, normalizeSheetName(fileName.replace(/\.[^.]+$/, ""), usedNames));
  });

  return workbook;
}

async function collectFirstSheetRows(files, headerRowNumber) {
  const parsedFiles = [];

  for (const file of files) {
    const buffer = await readFileAsArrayBuffer(file);
    const workbook = XLSX.read(buffer, {
      type: "array",
      cellNF: true,
      cellDates: true,
    });
    const firstSheetName = workbook.SheetNames[0];

    if (!firstSheetName) {
      throw new Error(`${file.name} 파일에 읽을 수 있는 시트가 없습니다.`);
    }

    parsedFiles.push({
      file,
      fileName: file.name,
      workbook,
      firstSheetName,
      ...parseSheetRows(workbook.Sheets[firstSheetName], headerRowNumber),
    });
  }

  return parsedFiles;
}

async function handleMerge() {
  const files = Array.from(mergeFilesInput.files || []);
  const headerRowNumber = Number(headerRowInput.value);

  if (!files.length) {
    setStatus(mergeStatus, "병합할 엑셀 파일을 하나 이상 선택해야 합니다.", "error");
    return;
  }

  if (!Number.isInteger(headerRowNumber) || headerRowNumber < 1) {
    setStatus(mergeStatus, "컬럼 행 번호는 1 이상의 정수여야 합니다.", "error");
    return;
  }

  setStatus(mergeStatus, "파일을 분석하고 병합 파일을 생성하는 중입니다.");

  try {
    const parsedFiles = await collectFirstSheetRows(files, headerRowNumber);
    const workbook = buildMergeWorkbook(parsedFiles, getCheckedValue("merge-mode", "single-sheet"));
    const outputName =
      getCheckedValue("merge-mode", "single-sheet") === "single-sheet"
        ? "merged_single_sheet.xlsx"
        : "merged_by_sheets.xlsx";

    XLSX.writeFile(workbook, outputName);
    setStatus(mergeStatus, `병합이 완료되었습니다. ${outputName} 파일이 다운로드됩니다.`, "success");
  } catch (error) {
    setStatus(mergeStatus, error.message || "병합 중 오류가 발생했습니다.", "error");
  }
}

function updateSplitOptionFields() {
  const mode = getCheckedValue("split-mode", "row");

  rowSizeField.classList.toggle("is-hidden", mode !== "row");
  columnNameField.classList.toggle("is-hidden", mode !== "column");
}

function addWorkbookToZip(zip, path, workbook) {
  const content = XLSX.write(workbook, { type: "array", bookType: "xlsx" });
  zip.file(path, content);
}

function addRowSplitFiles(zip, parsedFile, rowSize) {
  const baseName = parsedFile.fileName.replace(/\.[^.]+$/, "");

  for (let index = 0; index < parsedFile.dataRows.length; index += rowSize) {
    const chunk = parsedFile.dataRows.slice(index, index + rowSize);
    const workbook = createWorkbookFromRows(parsedFile.headers, chunk, parsedFile.firstSheetName);
    const partNumber = Math.floor(index / rowSize) + 1;
    addWorkbookToZip(zip, `${baseName}/${baseName}_rows_${partNumber}.xlsx`, workbook);
  }
}

function addColumnSplitFiles(zip, parsedFile, columnName) {
  const columnIndex = parsedFile.headers.indexOf(columnName);

  if (columnIndex === -1) {
    throw new Error(`${parsedFile.fileName} 파일에서 기준 컬럼명 "${columnName}"을 찾지 못했습니다.`);
  }

  const groupedRows = new Map();

  parsedFile.dataRows.forEach((row) => {
    const cell = row[columnIndex];
    const key = safeFileSegment(cell?.v ?? "빈값");

    if (!groupedRows.has(key)) {
      groupedRows.set(key, []);
    }

    groupedRows.get(key).push(row);
  });

  const baseName = parsedFile.fileName.replace(/\.[^.]+$/, "");

  groupedRows.forEach((rows, key) => {
    const workbook = createWorkbookFromRows(parsedFile.headers, rows, parsedFile.firstSheetName);
    addWorkbookToZip(zip, `${baseName}/${baseName}_${key}.xlsx`, workbook);
  });
}

function addSheetSplitFiles(zip, parsedFile, headerRowNumber) {
  const baseName = parsedFile.fileName.replace(/\.[^.]+$/, "");

  parsedFile.workbook.SheetNames.forEach((sheetName) => {
    const parsedSheet = parseSheetRows(parsedFile.workbook.Sheets[sheetName], headerRowNumber);
    const workbook = createWorkbookFromRows(parsedSheet.headers, parsedSheet.dataRows, sheetName);
    addWorkbookToZip(zip, `${baseName}/${baseName}_${safeFileSegment(sheetName)}.xlsx`, workbook);
  });
}

async function handleSplit() {
  const files = Array.from(splitFilesInput.files || []);
  const headerRowNumber = Number(splitHeaderRowInput.value);
  const mode = getCheckedValue("split-mode", "row");

  if (!files.length) {
    setStatus(splitStatus, "분할할 엑셀 파일을 하나 이상 선택해야 합니다.", "error");
    return;
  }

  if (!Number.isInteger(headerRowNumber) || headerRowNumber < 1) {
    setStatus(splitStatus, "컬럼 행 번호는 1 이상의 정수여야 합니다.", "error");
    return;
  }

  if (mode === "row") {
    const rowSize = Number(splitRowSizeInput.value);

    if (!Number.isInteger(rowSize) || rowSize < 1) {
      setStatus(splitStatus, "행 단위는 1 이상의 정수여야 합니다.", "error");
      return;
    }
  }

  if (mode === "column" && !splitColumnNameInput.value.trim()) {
    setStatus(splitStatus, "컬럼 기준 분할에서는 기준 컬럼명을 입력해야 합니다.", "error");
    return;
  }

  setStatus(splitStatus, "파일을 분석하고 분할 결과 ZIP을 생성하는 중입니다.");

  try {
    const parsedFiles = await collectFirstSheetRows(files, headerRowNumber);
    const zip = new JSZip();

    parsedFiles.forEach((parsedFile) => {
      if (mode === "row") {
        addRowSplitFiles(zip, parsedFile, Number(splitRowSizeInput.value));
        return;
      }

      if (mode === "column") {
        addColumnSplitFiles(zip, parsedFile, splitColumnNameInput.value.trim());
        return;
      }

      addSheetSplitFiles(zip, parsedFile, headerRowNumber);
    });

    const zipBlob = await zip.generateAsync({ type: "blob" });
    const outputName =
      mode === "row"
        ? "split_by_rows.zip"
        : mode === "column"
          ? "split_by_column.zip"
          : "split_by_sheet.zip";

    downloadBlob(zipBlob, outputName);
    setStatus(splitStatus, `분할이 완료되었습니다. ${outputName} 파일이 다운로드됩니다.`, "success");
  } catch (error) {
    setStatus(splitStatus, error.message || "분할 중 오류가 발생했습니다.", "error");
  }
}

triggers.forEach((trigger) => {
  trigger.addEventListener("click", (event) => {
    const { target } = trigger.dataset;

    if (!target) {
      return;
    }

    event.preventDefault();
    activatePanel(target);
  });
});

mergeFilesInput?.addEventListener("change", () => {
  const message = renderSelectedFiles(
    mergeFilesInput,
    mergeFileList,
    "파일을 선택한 뒤 병합 파일 생성을 실행하세요.",
    "병합 방식을 선택하고 실행하세요."
  );
  setStatus(mergeStatus, message);
});

splitFilesInput?.addEventListener("change", () => {
  const message = renderSelectedFiles(
    splitFilesInput,
    splitFileList,
    "파일을 선택한 뒤 분할 파일 생성을 실행하세요.",
    "분할 방식을 선택하고 실행하세요."
  );
  setStatus(splitStatus, message);
});

document.querySelectorAll('input[name="split-mode"]').forEach((radio) => {
  radio.addEventListener("change", updateSplitOptionFields);
});

mergeRunButton?.addEventListener("click", handleMerge);
splitRunButton?.addEventListener("click", handleSplit);

updateSplitOptionFields();
