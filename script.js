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
const peopleFileInput = document.getElementById("people-file");
const peopleHeaderRowInput = document.getElementById("people-header-row");
const peopleMonthInput = document.getElementById("people-month");
const peopleRunButton = document.getElementById("people-run");
const peopleStatus = document.getElementById("people-status");
const peopleFileList = document.getElementById("people-file-list");
const vacationFileInput = document.getElementById("vacation-file");
const vacationYearInput = document.getElementById("vacation-year");
const vacationHeaderRowInput = document.getElementById("vacation-header-row");
const vacationBasisInputs = document.querySelectorAll('input[name="vacation-basis"]');
const vacationRunButton = document.getElementById("vacation-run");
const vacationStatus = document.getElementById("vacation-status");
const vacationFileList = document.getElementById("vacation-file-list");

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
      const cell = row[columnIndex];
      const address = XLSX.utils.encode_cell({ r: rowIndex + 1, c: columnIndex });

      if (!cell) {
        return;
      }

      sheet[address] = {
        v: cell.v ?? "",
        t: cell.t || "s",
      };

      if (cell.f) {
        sheet[address].f = cell.f;
      }

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

function findHeaderIndex(headers, candidates) {
  return headers.findIndex((header) => candidates.includes(String(header).trim()));
}

function parseExcelDate(cell) {
  if (!cell || cell.v == null || cell.v === "") {
    return null;
  }

  if (cell.v instanceof Date) {
    return new Date(cell.v.getFullYear(), cell.v.getMonth(), cell.v.getDate());
  }

  if (typeof cell.v === "number") {
    const parsed = XLSX.SSF.parse_date_code(cell.v);

    if (!parsed) {
      return null;
    }

    return new Date(parsed.y, parsed.m - 1, parsed.d);
  }

  const raw = String(cell.v).trim();

  if (!raw) {
    return null;
  }

  const compactMatch = raw.match(/^(\d{4})(\d{2})(\d{2})$/);

  if (compactMatch) {
    return new Date(Number(compactMatch[1]), Number(compactMatch[2]) - 1, Number(compactMatch[3]));
  }

  const koreanMatch = raw.match(/^(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일$/);

  if (koreanMatch) {
    return new Date(Number(koreanMatch[1]), Number(koreanMatch[2]) - 1, Number(koreanMatch[3]));
  }

  const normalized = raw.replace(/[.]/g, "-").replace(/\//g, "-");
  const parsedDate = new Date(normalized);

  if (Number.isNaN(parsedDate.getTime())) {
    return null;
  }

  return new Date(parsedDate.getFullYear(), parsedDate.getMonth(), parsedDate.getDate());
}

function isBetweenMonth(date, monthStart, monthEnd) {
  return date && date >= monthStart && date <= monthEnd;
}

function isActiveAtMonthEnd(joinDate, leaveDate, monthEnd) {
  if (!joinDate || joinDate > monthEnd) {
    return false;
  }

  if (!leaveDate) {
    return true;
  }

  return leaveDate >= monthEnd;
}

function createSummarySheet(monthLabel, joiners, leavers, activeEmployees) {
  return XLSX.utils.aoa_to_sheet([
    ["기준 연월", monthLabel],
    ["당월 입사자 수", joiners.length],
    ["당월 퇴사자 수", leavers.length],
    ["말일 기준 재직 인원 수", activeEmployees.length],
  ]);
}

function getEmployeeKey(row, employeeIdIndex) {
  return String(row[employeeIdIndex]?.v ?? "").trim();
}

function pushUniqueEmployee(targetRows, seenKeys, employeeKey, row) {
  if (!employeeKey || seenKeys.has(employeeKey)) {
    return;
  }

  seenKeys.add(employeeKey);
  targetRows.push(row);
}

function createCell(value, type = "s", format = "") {
  const cell = {
    v: value,
    t: type,
  };

  if (format) {
    cell.z = format;
  }

  return cell;
}

function createFormulaCell(formula, resultType = "n", format = "") {
  const cell = {
    f: formula,
    v: resultType === "n" ? 0 : "",
    t: resultType,
  };

  if (format) {
    cell.z = format;
  }

  return cell;
}

function formatDateValue(date) {
  if (!date) {
    return "";
  }

  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function addMonthsKeepingRule(joinDate, monthsToAdd) {
  if (joinDate.getDate() === 1) {
    return new Date(joinDate.getFullYear(), joinDate.getMonth() + monthsToAdd, 1);
  }

  return new Date(joinDate.getFullYear(), joinDate.getMonth() + monthsToAdd, joinDate.getDate() - 1);
}

function getMonthlyLeaveAccrualDate(joinDate, sequence) {
  return addMonthsKeepingRule(joinDate, sequence);
}

function getCompletedYearsAtDate(joinDate, targetDate) {
  let years = targetDate.getFullYear() - joinDate.getFullYear();
  const anniversaryThisYear = new Date(
    targetDate.getFullYear(),
    joinDate.getMonth(),
    joinDate.getDate()
  );

  if (targetDate < anniversaryThisYear) {
    years -= 1;
  }

  return Math.max(years, 0);
}

function getAnnualLeaveDays(completedYears) {
  if (completedYears < 1) {
    return 0;
  }

  return 15 + Math.floor((completedYears - 1) / 2);
}

function countWorkedMonthsInPreviousYear(joinDate, leaveDate, year) {
  let months = 0;

  for (let monthIndex = 0; monthIndex < 12; monthIndex += 1) {
    const monthStart = new Date(year, monthIndex, 1);
    const monthEnd = new Date(year, monthIndex + 1, 0);

    if (joinDate <= monthEnd && (!leaveDate || leaveDate >= monthStart)) {
      months += 1;
    }
  }

  return months;
}

function getSelectedVacationBasis() {
  return document.querySelector('input[name="vacation-basis"]:checked')?.value || "fiscal";
}

function getFiscalAnnualAccrual(joinDate, leaveDate, selectedYear, monthNumber) {
  if (monthNumber !== 1) {
    return 0;
  }

  const janFirst = new Date(selectedYear, 0, 1);

  if (leaveDate && leaveDate < janFirst) {
    return 0;
  }

  const completedYears = getCompletedYearsAtDate(joinDate, janFirst);

  if (completedYears >= 1) {
    return getAnnualLeaveDays(completedYears);
  }

  if (joinDate.getFullYear() === selectedYear - 1) {
    const workedMonths = countWorkedMonthsInPreviousYear(joinDate, leaveDate, selectedYear - 1);
    return Number(((15 / 12) * workedMonths).toFixed(2));
  }

  return 0;
}

function getAnniversaryAnnualAccrual(joinDate, leaveDate, selectedYear, monthNumber) {
  const grantDate = new Date(selectedYear, joinDate.getMonth(), joinDate.getDate());

  if (grantDate.getMonth() + 1 !== monthNumber) {
    return 0;
  }

  if (leaveDate && leaveDate < grantDate) {
    return 0;
  }

  const completedYears = getCompletedYearsAtDate(joinDate, grantDate);
  return getAnnualLeaveDays(completedYears);
}

function getMonthlyAccrual(joinDate, leaveDate, selectedYear, monthNumber) {
  let accrualCount = 0;

  for (let sequence = 1; sequence <= 11; sequence += 1) {
    const accrualDate = getMonthlyLeaveAccrualDate(joinDate, sequence);

    if (accrualDate.getFullYear() !== selectedYear || accrualDate.getMonth() + 1 !== monthNumber) {
      continue;
    }

    if (leaveDate && leaveDate < accrualDate) {
      continue;
    }

    accrualCount += 1;
  }

  return accrualCount;
}

function buildVacationLedgerSheet(employees, selectedYear, monthNumber, basis, previousSheetName = "") {
  const daysInMonth = new Date(selectedYear, monthNumber, 0).getDate();
  const dayHeaders = Array.from({ length: daysInMonth }, (_, index) => `${index + 1}`);
  const isJanuary = monthNumber === 1;
  const summaryHeaders = ["전기잔여", "전월잔여", "연차발생", "월차발생", "당월사용", "잔여연차"];
  const headers = ["사번", "이름", "입사일", "퇴사일", ...summaryHeaders, ...dayHeaders];
  const priorCarryColumn = 4;
  const carryColumn = 5;
  const annualColumn = 6;
  const monthlyColumn = 7;
  const usedColumn = 8;
  const remainingColumn = 9;
  const dayStartColumn = 10;
  const dayEndColumn = dayStartColumn + daysInMonth - 1;
  const sheetName = `${selectedYear}년${monthNumber}월`;
  const rows = employees.map((employee, index) => {
    const excelRow = index + 2;
    const baseCells = [
      createCell(employee.employeeId),
      createCell(employee.employeeName),
      createCell(formatDateValue(employee.joinDate)),
      createCell(formatDateValue(employee.leaveDate)),
    ];
    const dayCells = Array.from({ length: daysInMonth }, () => null);
    const priorCarryCell = null;
    const carryRef =
      isJanuary
        ? createCell(0, "n", "0.00")
        : createFormulaCell(
            `'${previousSheetName}'!${XLSX.utils.encode_col(remainingColumn)}${excelRow}`,
            "n",
            "0.00"
          );
    const annualAccrual =
      basis === "fiscal"
        ? getFiscalAnnualAccrual(employee.joinDate, employee.leaveDate, selectedYear, monthNumber)
        : getAnniversaryAnnualAccrual(employee.joinDate, employee.leaveDate, selectedYear, monthNumber);
    const monthlyAccrual = getMonthlyAccrual(employee.joinDate, employee.leaveDate, selectedYear, monthNumber);
    const dayRange = `${XLSX.utils.encode_col(dayStartColumn)}${excelRow}:${XLSX.utils.encode_col(dayEndColumn)}${excelRow}`;
    const annualCell = createCell(annualAccrual, "n", "0.00");
    const monthlyCell = createCell(monthlyAccrual, "n", "0.00");
    const usedCell = createFormulaCell(`COUNTIF(${dayRange},"<>")`, "n", "0.00");
    const remainingFormula = `${XLSX.utils.encode_col(priorCarryColumn)}${excelRow}+${XLSX.utils.encode_col(carryColumn)}${excelRow}+${XLSX.utils.encode_col(annualColumn)}${excelRow}+${XLSX.utils.encode_col(monthlyColumn)}${excelRow}-${XLSX.utils.encode_col(usedColumn)}${excelRow}`;
    const remainingCell = createFormulaCell(remainingFormula, "n", "0.00");

    return [...baseCells, priorCarryCell, carryRef, annualCell, monthlyCell, usedCell, remainingCell, ...dayCells];
  });

  return {
    sheetName,
    sheet: buildSheetFromStructuredRows(headers, rows),
  };
}

async function handleVacationLedger() {
  const file = vacationFileInput.files?.[0];
  const selectedYear = Number(vacationYearInput.value);
  const headerRowNumber = Number(vacationHeaderRowInput.value);
  const basis = getSelectedVacationBasis();

  if (!file) {
    setStatus(vacationStatus, "연차관리대장을 만들 인사 엑셀 파일을 선택해야 합니다.", "error");
    return;
  }

  if (!Number.isInteger(selectedYear) || selectedYear < 1900) {
    setStatus(vacationStatus, "생성할 연도를 올바르게 입력해야 합니다.", "error");
    return;
  }

  if (!Number.isInteger(headerRowNumber) || headerRowNumber < 1) {
    setStatus(vacationStatus, "컬럼 행 번호는 1 이상의 정수여야 합니다.", "error");
    return;
  }

  setStatus(vacationStatus, "연차관리대장 파일을 생성하는 중입니다.");

  try {
    const parsedFiles = await collectFirstSheetRows([file], headerRowNumber);
    const parsedFile = parsedFiles[0];
    const headers = parsedFile.headers;
    const employeeIdIndex = findHeaderIndex(headers, ["사번"]);
    const nameIndex = findHeaderIndex(headers, ["이름", "성명"]);
    const joinIndex = findHeaderIndex(headers, ["입사일", "입사"]);
    const leaveIndex = findHeaderIndex(headers, ["퇴사일", "퇴사"]);

    if (employeeIdIndex === -1 || nameIndex === -1 || joinIndex === -1 || leaveIndex === -1) {
      throw new Error("필수 컬럼이 없습니다. 사번, 이름 또는 성명, 입사일 또는 입사, 퇴사일 또는 퇴사가 필요합니다.");
    }

    const employees = [];
    const seenEmployeeIds = new Set();

    parsedFile.dataRows.forEach((row) => {
      const employeeId = String(row[employeeIdIndex]?.v ?? "").trim();

      if (!employeeId || seenEmployeeIds.has(employeeId)) {
        return;
      }

      const joinDate = parseExcelDate(row[joinIndex]);

      if (!joinDate) {
        return;
      }

      seenEmployeeIds.add(employeeId);
      employees.push({
        employeeId,
        employeeName: String(row[nameIndex]?.v ?? "").trim(),
        joinDate,
        leaveDate: parseExcelDate(row[leaveIndex]),
      });
    });

    const workbook = XLSX.utils.book_new();
    let previousSheetName = "";

    for (let monthNumber = 1; monthNumber <= 12; monthNumber += 1) {
      const { sheetName, sheet } = buildVacationLedgerSheet(
        employees,
        selectedYear,
        monthNumber,
        basis,
        previousSheetName
      );
      XLSX.utils.book_append_sheet(workbook, sheet, sheetName);
      previousSheetName = sheetName;
    }

    const outputName = `vacation_ledger_${selectedYear}_${basis}.xlsx`;
    XLSX.writeFile(workbook, outputName);
    setStatus(vacationStatus, `연차관리대장 생성이 완료되었습니다. ${outputName} 파일이 다운로드됩니다.`, "success");
  } catch (error) {
    setStatus(vacationStatus, error.message || "연차관리대장 생성 중 오류가 발생했습니다.", "error");
  }
}

async function handlePeopleAnalysis() {
  const file = peopleFileInput.files?.[0];
  const headerRowNumber = Number(peopleHeaderRowInput.value);
  const monthValue = peopleMonthInput.value;

  if (!file) {
    setStatus(peopleStatus, "분석할 인사 엑셀 파일을 선택해야 합니다.", "error");
    return;
  }

  if (!Number.isInteger(headerRowNumber) || headerRowNumber < 1) {
    setStatus(peopleStatus, "컬럼 행 번호는 1 이상의 정수여야 합니다.", "error");
    return;
  }

  if (!monthValue) {
    setStatus(peopleStatus, "기준 연월을 선택해야 합니다.", "error");
    return;
  }

  setStatus(peopleStatus, "인사 데이터를 분석하고 결과 파일을 생성하는 중입니다.");

    try {
      const parsedFiles = await collectFirstSheetRows([file], headerRowNumber);
      const parsedFile = parsedFiles[0];
      const headers = parsedFile.headers;
      const employeeIdIndex = findHeaderIndex(headers, ["사번"]);
      const nameIndex = findHeaderIndex(headers, ["성명", "이름"]);
      const joinIndex = findHeaderIndex(headers, ["입사일", "입사"]);
      const leaveIndex = findHeaderIndex(headers, ["퇴사일", "퇴사"]);

      if (employeeIdIndex === -1) {
        throw new Error("필수 컬럼이 없습니다. 사번 컬럼은 반드시 있어야 합니다.");
      }

      if (nameIndex === -1) {
        throw new Error("필수 컬럼이 없습니다. 성명 또는 이름 컬럼 중 하나는 있어야 합니다.");
      }

      if (joinIndex === -1 || leaveIndex === -1) {
        throw new Error("필수 컬럼이 없습니다. 입사일(또는 입사), 퇴사일(또는 퇴사)이 필요합니다.");
      }

    const [year, month] = monthValue.split("-").map(Number);
      const monthStart = new Date(year, month - 1, 1);
      const monthEnd = new Date(year, month, 0);

      const joiners = [];
      const leavers = [];
      const activeEmployees = [];
      const joinerKeys = new Set();
      const leaverKeys = new Set();
      const activeKeys = new Set();

      parsedFile.dataRows.forEach((row) => {
        const employeeKey = getEmployeeKey(row, employeeIdIndex);
        const joinDate = parseExcelDate(row[joinIndex]);
        const leaveDate = parseExcelDate(row[leaveIndex]);

        if (isBetweenMonth(joinDate, monthStart, monthEnd)) {
          pushUniqueEmployee(joiners, joinerKeys, employeeKey, row);
        }

        if (isBetweenMonth(leaveDate, monthStart, monthEnd)) {
          pushUniqueEmployee(leavers, leaverKeys, employeeKey, row);
        }

        if (isActiveAtMonthEnd(joinDate, leaveDate, monthEnd)) {
          pushUniqueEmployee(activeEmployees, activeKeys, employeeKey, row);
        }
      });

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, createSummarySheet(monthValue, joiners, leavers, activeEmployees), "요약");
    XLSX.utils.book_append_sheet(workbook, buildSheetFromStructuredRows(headers, joiners), "입사자");
    XLSX.utils.book_append_sheet(workbook, buildSheetFromStructuredRows(headers, leavers), "퇴사자");
    XLSX.utils.book_append_sheet(workbook, buildSheetFromStructuredRows(headers, activeEmployees), "재직자");

    const outputName = `people_analysis_${monthValue}.xlsx`;
    XLSX.writeFile(workbook, outputName);
    setStatus(peopleStatus, `분석이 완료되었습니다. ${outputName} 파일이 다운로드됩니다.`, "success");
  } catch (error) {
    setStatus(peopleStatus, error.message || "인원 분석 중 오류가 발생했습니다.", "error");
  }
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

peopleFileInput?.addEventListener("change", () => {
  const message = renderSelectedFiles(
    peopleFileInput,
    peopleFileList,
    "파일과 기준 연월을 입력한 뒤 분석 파일 생성을 실행하세요.",
    "컬럼 행과 기준 연월을 입력하고 실행하세요."
  );
  setStatus(peopleStatus, message);
});

vacationFileInput?.addEventListener("change", () => {
  const message = renderSelectedFiles(
    vacationFileInput,
    vacationFileList,
    "파일과 생성 연도, 기준 방식을 입력한 뒤 연차관리대장 생성을 실행하세요.",
    "생성 연도와 기준 방식을 입력하고 실행하세요."
  );
  setStatus(vacationStatus, message);
});

document.querySelectorAll('input[name="split-mode"]').forEach((radio) => {
  radio.addEventListener("change", updateSplitOptionFields);
});

mergeRunButton?.addEventListener("click", handleMerge);
splitRunButton?.addEventListener("click", handleSplit);
peopleRunButton?.addEventListener("click", handlePeopleAnalysis);
vacationRunButton?.addEventListener("click", handleVacationLedger);

updateSplitOptionFields();
