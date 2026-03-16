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
const peoplePeriodModeInput = document.getElementById("people-period-mode");
const peoplePeriodYearInput = document.getElementById("people-period-year");
const peoplePeriodUnitInput = document.getElementById("people-period-unit");
const peoplePeriodUnitLabel = document.getElementById("people-period-unit-label");
const peoplePeriodUnitHelp = document.getElementById("people-period-unit-help");
const peoplePeriodColumnInput = document.getElementById("people-period-column");
const peopleSalaryColumnInput = document.getElementById("people-salary-column");
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
const salaryFilesInput = document.getElementById("salary-files");
const salaryHeaderRowInput = document.getElementById("salary-header-row");
const salaryPeriodColumnInput = document.getElementById("salary-period-column");
const salaryRunButton = document.getElementById("salary-run");
const salaryStatus = document.getElementById("salary-status");
const salaryFileList = document.getElementById("salary-file-list");
const salaryItemPicker = document.getElementById("salary-item-picker");
const salarySelectAllButton = document.getElementById("salary-select-all");
const workstatFileInput = document.getElementById("workstat-file");
const workstatHeaderRowInput = document.getElementById("workstat-header-row");
const workstatStartDateInput = document.getElementById("workstat-start-date");
const workstatRunButton = document.getElementById("workstat-run");
const workstatStatus = document.getElementById("workstat-status");
const workstatFileList = document.getElementById("workstat-file-list");
const workstatResult = document.getElementById("workstat-result");
const workstatForwardPreview = document.getElementById("workstat-forward-preview");
const workstatBackwardPreview = document.getElementById("workstat-backward-preview");
const severanceCompanyInput = document.getElementById("severance-company");
const severanceSiteInput = document.getElementById("severance-site");
const severanceNameInput = document.getElementById("severance-name");
const severanceBirthDateInput = document.getElementById("severance-birth-date");
const severanceJoinDateInput = document.getElementById("severance-join-date");
const severanceEndDateInput = document.getElementById("severance-end-date");
const severanceEmploymentTypeInput = document.getElementById("severance-employment-type");
const severanceMidStartInput = document.getElementById("severance-mid-start");
const severanceMidEndInput = document.getElementById("severance-mid-end");
const severanceWageInputs = [1, 2, 3, 4].map((index) => document.getElementById(`severance-wage-${index}`));
const severancePeriodLabels = [1, 2, 3, 4].map((index) => document.getElementById(`severance-period-label-${index}`));
const severanceBonusInput = document.getElementById("severance-bonus");
const severanceVacationPayInput = document.getElementById("severance-vacation-pay");
const severanceRunButton = document.getElementById("severance-run");
const severanceSavePdfButton = document.getElementById("severance-save-pdf");
const severanceStatus = document.getElementById("severance-status");
const severanceResult = document.getElementById("severance-result");
const severanceSummary = document.getElementById("severance-summary");
const severancePeriodTable = document.getElementById("severance-period-table");

let salaryHeaderOptions = [];

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

function normalizeHeaderText(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/\s+/g, "");
}

function headerIncludesCandidate(header, candidates) {
  const normalizedHeader = normalizeHeaderText(header);

  return candidates.some((candidate) => {
    const normalizedCandidate = normalizeHeaderText(candidate);
    return (
      normalizedHeader === normalizedCandidate ||
      normalizedHeader.includes(normalizedCandidate) ||
      normalizedCandidate.includes(normalizedHeader)
    );
  });
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

      if (cell.l) {
        sheet[address].l = cell.l;
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

function findHeaderIndexLoose(headers, candidates) {
  return headers.findIndex((header) => headerIncludesCandidate(header, candidates));
}

function parseExcelDate(cell) {
  if (!cell || cell.v == null || cell.v === "") {
    return null;
  }

  if (cell.v instanceof Date) {
    return new Date(cell.v.getFullYear(), cell.v.getMonth(), cell.v.getDate());
  }

  if (
    typeof cell.v === "number" &&
    /[ymd년월일/\-.]/i.test(String(cell.z || "")) &&
    cell.v > 20000 &&
    cell.v < 90000
  ) {
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

function getPeoplePeriodConfig() {
  const mode = peoplePeriodModeInput?.value || "monthly";
  const selectedYear = Number(peoplePeriodYearInput?.value);
  const selectedUnit = peoplePeriodUnitInput?.value;

  if (!Number.isInteger(selectedYear) || selectedYear < 1900) {
    throw new Error("분석할 연도를 올바르게 입력해야 합니다.");
  }

  const monthlyMonths = (startMonth, endMonth) =>
    Array.from({ length: endMonth - startMonth + 1 }, (_, index) => ({
      year: selectedYear,
      month: startMonth + index,
    }));

  if (mode === "monthly") {
    const month = Number(selectedUnit);

    if (!Number.isInteger(month) || month < 1 || month > 12) {
      throw new Error("분석할 월을 선택해야 합니다.");
    }

    return {
      mode,
      label: `${selectedYear}-${String(month).padStart(2, "0")}`,
      startDate: new Date(selectedYear, month - 1, 1),
      endDate: new Date(selectedYear, month, 0),
      months: monthlyMonths(month, month),
      monthCount: 1,
    };
  }

  if (mode === "quarterly") {
    const quarter = Number(selectedUnit);

    if (!Number.isInteger(quarter) || quarter < 1 || quarter > 4) {
      throw new Error("분석할 분기를 선택해야 합니다.");
    }

    const startMonth = (quarter - 1) * 3 + 1;
    const endMonth = startMonth + 2;

    return {
      mode,
      label: `${selectedYear}년 ${quarter}분기`,
      startDate: new Date(selectedYear, startMonth - 1, 1),
      endDate: new Date(selectedYear, endMonth, 0),
      months: monthlyMonths(startMonth, endMonth),
      monthCount: 3,
    };
  }

  if (mode === "half-yearly") {
    const half = Number(selectedUnit);

    if (half !== 1 && half !== 2) {
      throw new Error("분석할 반기를 선택해야 합니다.");
    }

    const startMonth = half === 1 ? 1 : 7;
    const endMonth = half === 1 ? 6 : 12;

    return {
      mode,
      label: `${selectedYear}년 ${half === 1 ? "상반기" : "하반기"}`,
      startDate: new Date(selectedYear, startMonth - 1, 1),
      endDate: new Date(selectedYear, endMonth, 0),
      months: monthlyMonths(startMonth, endMonth),
      monthCount: 6,
    };
  }

  return {
    mode: "yearly",
    label: `${selectedYear}년`,
    startDate: new Date(selectedYear, 0, 1),
    endDate: new Date(selectedYear, 12, 0),
    months: monthlyMonths(1, 12),
    monthCount: 12,
  };
}

function updatePeoplePeriodControls() {
  const mode = peoplePeriodModeInput?.value || "monthly";

  if (!peoplePeriodUnitInput || !peoplePeriodUnitLabel || !peoplePeriodUnitHelp || !peoplePeriodYearInput) {
    return;
  }

  if (!peoplePeriodYearInput.value) {
    peoplePeriodYearInput.value = String(new Date().getFullYear());
  }

  const optionConfigMap = {
    monthly: {
      label: "기준 월",
      help: "선택 월 기준으로 입사/퇴사와 말일 재직 인원을 분석합니다.",
      options: Array.from({ length: 12 }, (_, index) => ({
        value: String(index + 1),
        label: `${index + 1}월`,
      })),
    },
    quarterly: {
      label: "기준 분기",
      help: "선택 분기(3개월) 기준으로 입사/퇴사와 말일 재직 인원을 분석합니다.",
      options: Array.from({ length: 4 }, (_, index) => ({
        value: String(index + 1),
        label: `${index + 1}분기`,
      })),
    },
    "half-yearly": {
      label: "기준 반기",
      help: "선택 반기(6개월) 기준으로 입사/퇴사와 말일 재직 인원을 분석합니다.",
      options: [
        { value: "1", label: "상반기" },
        { value: "2", label: "하반기" },
      ],
    },
    yearly: {
      label: "기준 기간",
      help: "선택 연도 전체 기준으로 입사/퇴사와 말일 재직 인원을 분석합니다.",
      options: [{ value: "1", label: "연간" }],
    },
  };

  const config = optionConfigMap[mode];
  peoplePeriodUnitLabel.textContent = config.label;
  peoplePeriodUnitHelp.textContent = config.help;
  peoplePeriodUnitInput.innerHTML = config.options
    .map((option) => `<option value="${option.value}">${option.label}</option>`)
    .join("");
}

function calculatePeriodAverageHeadcount(employees, activeEmployeeCount, periodConfig, analysisMonths) {
  if (periodConfig.mode === "monthly") {
    return activeEmployeeCount;
  }

  const totalActiveEmployees = analysisMonths.reduce((sum, period) => {
    const monthEnd = new Date(period.year, period.month, 0);
    const monthActiveCount = employees.filter((employee) =>
      isActiveAtMonthEnd(employee.joinDate, employee.leaveDate, monthEnd)
    ).length;
    return sum + monthActiveCount;
  }, 0);

  return analysisMonths.length ? Number((totalActiveEmployees / analysisMonths.length).toFixed(2)) : 0;
}

function isPeriodInSelection(periodInfo, periodConfig) {
  if (!periodInfo) {
    return false;
  }

  return periodConfig.months.some((monthInfo) => monthInfo.year === periodInfo.year && monthInfo.month === periodInfo.month);
}

function getAvailablePeriodMonths(rows, periodColumnIndex, periodConfig) {
  const availableMonthMap = new Map();

  rows.forEach((row) => {
    const periodInfo = parsePeriodInfo(row[periodColumnIndex]);

    if (!isPeriodInSelection(periodInfo, periodConfig)) {
      return;
    }

    const key = `${periodInfo.year}-${String(periodInfo.month).padStart(2, "0")}`;

    if (!availableMonthMap.has(key)) {
      availableMonthMap.set(key, { year: periodInfo.year, month: periodInfo.month });
    }
  });

  return periodConfig.months.filter((monthInfo) => {
    const key = `${monthInfo.year}-${String(monthInfo.month).padStart(2, "0")}`;
    return availableMonthMap.has(key);
  });
}

function calculateSalaryAverageHeadcount(rows, employeeIdIndex, periodColumnIndex, salaryColumnIndex, periodConfig, analysisMonths) {
  if (periodConfig.mode === "monthly") {
    const paidEmployeeKeys = new Set();

    rows.forEach((row) => {
      const periodInfo = parsePeriodInfo(row[periodColumnIndex]);

      if (!isPeriodInSelection(periodInfo, periodConfig)) {
        return;
      }

      const salaryAmount = parseNumberValue(row[salaryColumnIndex]);

      if (salaryAmount === 0) {
        return;
      }

      const employeeKey = getEmployeeKey(row, employeeIdIndex);

      if (employeeKey) {
        paidEmployeeKeys.add(employeeKey);
      }
    });

    return paidEmployeeKeys.size;
  }

  const totalPaidEmployees = analysisMonths.reduce((sum, monthInfo) => {
    const paidEmployeeKeys = new Set();

    rows.forEach((row) => {
      const periodInfo = parsePeriodInfo(row[periodColumnIndex]);

      if (!periodInfo || periodInfo.year !== monthInfo.year || periodInfo.month !== monthInfo.month) {
        return;
      }

      const salaryAmount = parseNumberValue(row[salaryColumnIndex]);

      if (salaryAmount === 0) {
        return;
      }

      const employeeKey = getEmployeeKey(row, employeeIdIndex);

      if (employeeKey) {
        paidEmployeeKeys.add(employeeKey);
      }
    });

    return sum + paidEmployeeKeys.size;
  }, 0);

  return analysisMonths.length ? Number((totalPaidEmployees / analysisMonths.length).toFixed(2)) : 0;
}

function createSummarySheet(periodLabel, joiners, leavers, activeEmployees, salaryAverageHeadcount, periodAverageHeadcount) {
  return XLSX.utils.aoa_to_sheet([
    ["항목", "값", "비고"],
    ["분석 기간", periodLabel, "선택한 분석 구간입니다."],
    ["기간 입사자 수", joiners.length, "선택한 기간 안에 입사일이 포함된 인원 수입니다."],
    ["기간 퇴사자 수", leavers.length, "선택한 기간 안에 퇴사일이 포함된 인원 수입니다."],
    ["기간말 재직 인원 수", activeEmployees.length, "선택한 기간 종료일 기준 재직 중인 인원 수입니다."],
    ["급여 평균인원", salaryAverageHeadcount, "선택한 기간 안에서 실제 데이터가 존재하는 각 월별 급여 컬럼 값이 0이 아닌 인원 수를 실제 월수로 나눈 값입니다."],
    ["기간평균인원", periodAverageHeadcount, "선택한 기간 안에서 실제 데이터가 존재하는 각 월말 재직인원 수 합계를 실제 월수로 나눈 값입니다."],
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

function parseNumberValue(cell) {
  if (!cell || cell.v == null || cell.v === "") {
    return 0;
  }

  if (typeof cell.v === "number") {
    return cell.v;
  }

  const normalized = String(cell.v).replace(/,/g, "").trim();
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : 0;
}

function isNumericAmountCell(cell) {
  if (!cell || cell.v == null || cell.v === "") {
    return false;
  }

  if (parseExcelDate(cell)) {
    return false;
  }

  if (typeof cell.v === "number") {
    return true;
  }

  const normalized = String(cell.v).replace(/,/g, "").trim();

  if (!normalized) {
    return false;
  }

  return /^-?\d+(\.\d+)?$/.test(normalized);
}

function parsePeriodInfo(cell) {
  const dateValue = parseExcelDate(cell);

  if (dateValue) {
    return {
      year: dateValue.getFullYear(),
      month: dateValue.getMonth() + 1,
    };
  }

  const raw = String(cell?.v ?? "").trim();

  if (!raw) {
    return null;
  }

  const monthMatch = raw.match(/^(\d{4})[-./년\s]?(\d{1,2})/);

  if (monthMatch) {
    return {
      year: Number(monthMatch[1]),
      month: Number(monthMatch[2]),
    };
  }

  const compactMatch = raw.match(/^(\d{4})(\d{2})$/);

  if (compactMatch) {
    return {
      year: Number(compactMatch[1]),
      month: Number(compactMatch[2]),
    };
  }

  return null;
}

function formatMonthKey(period) {
  return `${period.year}-${String(period.month).padStart(2, "0")}`;
}

function formatQuarterKey(period) {
  const quarter = Math.floor((period.month - 1) / 3) + 1;
  return `${period.year}-Q${quarter}`;
}

function formatHalfYearKey(period) {
  const half = period.month <= 6 ? "H1" : "H2";
  return `${period.year}-${half}`;
}

function formatDisplayLabel(key, mode) {
  if (mode === "monthly") {
    return key;
  }

  if (mode === "quarterly") {
    const [year, quarter] = key.split("-");
    return `${year}년 ${quarter.replace("Q", "")}분기`;
  }

  if (mode === "half-yearly") {
    const [year, half] = key.split("-");
    return `${year}년 ${half === "H1" ? "상반기" : "하반기"}`;
  }

  return key.includes("_")
    ? key
    : `${key}년`;
}

function buildSalaryHeaderPicker(headers) {
  if (!headers.length) {
    salaryItemPicker.innerHTML = "<div class=\"policy-item\"><p>파일을 먼저 선택해 주세요.</p></div>";
    if (salarySelectAllButton) {
      salarySelectAllButton.textContent = "전체선택";
    }
    return;
  }

  salaryItemPicker.innerHTML = headers
    .map(
      (header, index) => `
        <label class="choice">
          <input type="checkbox" class="salary-item-option" data-header-index="${index}" />
          <span>${header}</span>
        </label>
      `
    )
    .join("");

  if (salarySelectAllButton) {
    salarySelectAllButton.textContent = "전체선택";
  }
}

function inferSalarySelectableHeaders(parsedFiles) {
  const excludedCandidates = [
    "사번",
    "이름",
    "성명",
    "성별",
    "주민번호",
    "주민등록번호",
    "생년월일",
    "부서",
    "부서명",
    "직급",
    "직급명",
    "직위",
    "직책",
    "직군",
    "직군명",
    "지급월",
    "지급일",
    "정산월",
    "정산일",
    "입사",
    "입사일",
    "퇴사",
    "퇴사일",
    "일자",
    "날짜",
  ];
  const headerSamples = new Map();

  parsedFiles.forEach((parsedFile) => {
    parsedFile.headers.forEach((header, index) => {
      if (headerIncludesCandidate(header, excludedCandidates)) {
        return;
      }

      if (!headerSamples.has(header)) {
        headerSamples.set(header, []);
      }

      parsedFile.dataRows.forEach((row) => {
        const cell = row[index];

        if (!cell || cell.v == null || cell.v === "") {
          return;
        }

        if (headerSamples.get(header).length < 20) {
          headerSamples.get(header).push(cell);
        }
      });
    });
  });

  return Array.from(headerSamples.keys());
}

async function refreshSalaryHeaderOptions() {
  const files = Array.from(salaryFilesInput.files || []);
  const headerRowNumber = Number(salaryHeaderRowInput.value);

  salaryHeaderOptions = [];
  salaryItemPicker.innerHTML = "";

  if (!files.length || !Number.isInteger(headerRowNumber) || headerRowNumber < 1) {
    return;
  }

  try {
    const parsedFiles = await collectFirstSheetRows(files, headerRowNumber);
    salaryHeaderOptions = inferSalarySelectableHeaders(parsedFiles);
    buildSalaryHeaderPicker(salaryHeaderOptions);
  } catch (error) {
    setStatus(salaryStatus, error.message || "급여 컬럼 목록을 읽는 중 오류가 발생했습니다.", "error");
  }
}

function getSelectedSalaryItems() {
  return Array.from(document.querySelectorAll(".salary-item-option:checked"))
    .map((checkbox) => salaryHeaderOptions[Number(checkbox.dataset.headerIndex)])
    .filter(Boolean);
}

function getSelectedSalaryPeriodMode() {
  return getCheckedValue("salary-period-mode", "monthly");
}

function getSelectedSalaryGroupMode() {
  return getCheckedValue("salary-group-mode", "overall");
}

function getSalaryGroupInfo(headers, row, groupMode, employeeIdIndex, nameIndex) {
  if (groupMode === "overall") {
    return { key: "전체", label: "전체" };
  }

  if (groupMode === "individual") {
    const employeeId = String(row[employeeIdIndex]?.v ?? "").trim();
    const employeeName = nameIndex === -1 ? "" : String(row[nameIndex]?.v ?? "").trim();
    return { key: employeeId, label: employeeName || employeeId };
  }

  const candidateMap = {
    department: ["부서", "부서명"],
    grade: ["직급", "직급명"],
    "job-family": ["직군", "직군명"],
  };
  const groupIndex = findHeaderIndexLoose(headers, candidateMap[groupMode] || []);

  if (groupIndex === -1) {
    throw new Error(`선택한 비교 형태에 필요한 컬럼이 없습니다: ${candidateMap[groupMode].join(", ")}`);
  }

  const key = String(row[groupIndex]?.v ?? "").trim() || "미분류";
  return { key, label: key };
}

function getSalaryPeriodKey(period, mode) {
  if (mode === "monthly") {
    return formatMonthKey(period);
  }

  if (mode === "quarterly") {
    return formatQuarterKey(period);
  }

  if (mode === "half-yearly") {
    return formatHalfYearKey(period);
  }

  return `${period.year}`;
}

function sortPeriodKeys(keys, mode) {
  const sorted = [...keys];

  sorted.sort((left, right) => {
    if (mode === "monthly") {
      return left.localeCompare(right);
    }

    if (mode === "quarterly") {
      const [leftYear, leftQuarter] = left.split("-Q").map(Number);
      const [rightYear, rightQuarter] = right.split("-Q").map(Number);
      return leftYear - rightYear || leftQuarter - rightQuarter;
    }

    if (mode === "half-yearly") {
      const [leftYear, leftHalf] = left.split("-");
      const [rightYear, rightHalf] = right.split("-");
      return Number(leftYear) - Number(rightYear) || leftHalf.localeCompare(rightHalf);
    }

    return Number(left) - Number(right);
  });

  return sorted;
}

function buildYearlyComparisonPairs(monthlyKeys) {
  const periods = monthlyKeys.map((key) => {
    const [year, month] = key.split("-").map(Number);
    return { year, month };
  });
  const latest = periods[periods.length - 1];
  const minYear = periods[0]?.year;
  const pairs = [];

  if (!latest || minYear == null) {
    return pairs;
  }

  for (let year = latest.year; year > minYear; year -= 1) {
    const previousYear = year - 1;

    if (year === latest.year && latest.month < 12) {
      pairs.push({
        currentKey: `${year}_YTD_${latest.month}`,
        previousKey: `${previousYear}_YTD_${latest.month}`,
        currentLabel: `${year}년 1~${latest.month}월`,
        previousLabel: `${previousYear}년 1~${latest.month}월`,
        filter: (period) => period.year === year && period.month <= latest.month,
        previousFilter: (period) => period.year === previousYear && period.month <= latest.month,
      });
      continue;
    }

    pairs.push({
      currentKey: `${year}`,
      previousKey: `${previousYear}`,
      currentLabel: `${year}년`,
      previousLabel: `${previousYear}년`,
      filter: (period) => period.year === year,
      previousFilter: (period) => period.year === previousYear,
    });
  }

  return pairs;
}

function buildComparisonWorkbook(dataset, selectedItems, mode) {
  const workbook = XLSX.utils.book_new();

  if (mode === "yearly") {
    const monthlyKeys = sortPeriodKeys(Array.from(dataset.monthlyPeriods.keys()), "monthly");
    const yearlyPairs = buildYearlyComparisonPairs(monthlyKeys);

    yearlyPairs.forEach((pair) => {
      const rows = [];
      dataset.groups.forEach((groupData, groupKey) => {
        const row = [groupKey, groupData.label];

        selectedItems.forEach((item) => {
          const currentValue = dataset.monthlyTimeline
            .filter((entry) => pair.filter(entry.period))
            .reduce((sum, entry) => sum + (entry.groups.get(groupKey)?.items[item] || 0), 0);
          const previousValue = dataset.monthlyTimeline
            .filter((entry) => pair.previousFilter(entry.period))
            .reduce((sum, entry) => sum + (entry.groups.get(groupKey)?.items[item] || 0), 0);

          row.push(previousValue, currentValue, currentValue - previousValue);
        });

        rows.push(row);
      });

      const headers = ["비교기준", "표시명"];
      selectedItems.forEach((item) => {
        headers.push(`${item}(${pair.previousLabel})`, `${item}(${pair.currentLabel})`, `${item} 증감액`);
      });

      XLSX.utils.book_append_sheet(
        workbook,
        XLSX.utils.aoa_to_sheet([headers, ...rows]),
        normalizeSheetName(`${pair.currentLabel}_${pair.previousLabel}`, new Set())
      );
    });

    return workbook;
  }

  const periodKeys = sortPeriodKeys(Array.from(dataset.periods.keys()), mode);

  for (let index = periodKeys.length - 1; index > 0; index -= 1) {
    const currentKey = periodKeys[index];
    const previousKey = periodKeys[index - 1];
    const currentGroups = dataset.periods.get(currentKey) || new Map();
    const previousGroups = dataset.periods.get(previousKey) || new Map();
    const allGroupKeys = Array.from(dataset.groups.keys());
    const rows = allGroupKeys.map((groupKey) => {
      const currentGroup = currentGroups.get(groupKey) || { items: {}, label: dataset.groups.get(groupKey)?.label || groupKey };
      const previousGroup = previousGroups.get(groupKey) || { items: {}, label: dataset.groups.get(groupKey)?.label || groupKey };
      const row = [groupKey, dataset.groups.get(groupKey)?.label || groupKey];

      selectedItems.forEach((item) => {
        const previousValue = previousGroup.items[item] || 0;
        const currentValue = currentGroup.items[item] || 0;
        row.push(previousValue, currentValue, currentValue - previousValue);
      });

      return row;
    });
    const currentLabel = formatDisplayLabel(currentKey, mode);
    const previousLabel = formatDisplayLabel(previousKey, mode);
    const headers = ["비교기준", "표시명"];

    selectedItems.forEach((item) => {
      headers.push(`${item}(${previousLabel})`, `${item}(${currentLabel})`, `${item} 증감액`);
    });

    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.aoa_to_sheet([headers, ...rows]),
      normalizeSheetName(`${currentLabel}_${previousLabel}`, new Set())
    );
  }

  return workbook;
}

async function handleSalaryAnalysis() {
  const files = Array.from(salaryFilesInput.files || []);
  const headerRowNumber = Number(salaryHeaderRowInput.value);
  const selectedItems = getSelectedSalaryItems();
  const periodColumnName = salaryPeriodColumnInput.value.trim();
  const periodMode = getSelectedSalaryPeriodMode();
  const groupMode = getSelectedSalaryGroupMode();

  if (!files.length) {
    setStatus(salaryStatus, "비교할 급여 엑셀 파일을 하나 이상 선택해야 합니다.", "error");
    return;
  }

  if (!Number.isInteger(headerRowNumber) || headerRowNumber < 1) {
    setStatus(salaryStatus, "컬럼 행 번호는 1 이상의 정수여야 합니다.", "error");
    return;
  }

  if (!periodColumnName) {
    setStatus(salaryStatus, "기간 비교 컬럼명을 입력해야 합니다.", "error");
    return;
  }

  if (!selectedItems.length) {
    setStatus(salaryStatus, "비교할 급여 항목을 하나 이상 선택해야 합니다.", "error");
    return;
  }

  setStatus(salaryStatus, "급여 비교 파일을 생성하는 중입니다.");

  try {
    const parsedFiles = await collectFirstSheetRows(files, headerRowNumber);
    const dataset = {
      periods: new Map(),
      groups: new Map(),
      monthlyPeriods: new Map(),
      monthlyTimeline: [],
    };

    parsedFiles.forEach((parsedFile) => {
      const headers = parsedFile.headers;
      const employeeIdIndex = findHeaderIndex(headers, ["사번"]);
      const nameIndex = findHeaderIndex(headers, ["이름", "성명"]);
      const periodIndex = findHeaderIndexLoose(headers, [periodColumnName]);

      if (employeeIdIndex === -1) {
        throw new Error("급여 비교분석에는 사번 컬럼이 반드시 필요합니다.");
      }

      if (periodIndex === -1) {
        throw new Error(`기간 비교 컬럼명을 찾지 못했습니다: ${periodColumnName}`);
      }

      const itemIndexes = selectedItems.map((item) => {
        const index = findHeaderIndex(headers, [item]);

        if (index === -1) {
          throw new Error(`${parsedFile.fileName} 파일에서 비교 항목을 찾지 못했습니다: ${item}`);
        }

        return { item, index };
      });

      parsedFile.dataRows.forEach((row) => {
        const periodInfo = parsePeriodInfo(row[periodIndex]);

        if (!periodInfo) {
          return;
        }

        const groupInfo = getSalaryGroupInfo(headers, row, groupMode, employeeIdIndex, nameIndex);
        const periodKey = getSalaryPeriodKey(periodInfo, periodMode);

        if (!dataset.periods.has(periodKey)) {
          dataset.periods.set(periodKey, new Map());
        }

        if (!dataset.groups.has(groupInfo.key)) {
          dataset.groups.set(groupInfo.key, { label: groupInfo.label });
        }

        const periodGroups = dataset.periods.get(periodKey);

        if (!periodGroups.has(groupInfo.key)) {
          periodGroups.set(groupInfo.key, { label: groupInfo.label, items: {} });
        }

        itemIndexes.forEach(({ item, index }) => {
          const amount = parseNumberValue(row[index]);
          const groupRecord = periodGroups.get(groupInfo.key);
          groupRecord.items[item] = (groupRecord.items[item] || 0) + amount;
        });

        const monthKey = formatMonthKey(periodInfo);

        if (!dataset.monthlyPeriods.has(monthKey)) {
          dataset.monthlyPeriods.set(monthKey, true);
        }
      });
    });

    const monthlyKeys = sortPeriodKeys(Array.from(dataset.monthlyPeriods.keys()), "monthly");
    dataset.monthlyTimeline = monthlyKeys.map((key) => {
      const [year, month] = key.split("-").map(Number);
      const groups = new Map();

      parsedFiles.forEach((parsedFile) => {
        const headers = parsedFile.headers;
        const employeeIdIndex = findHeaderIndex(headers, ["사번"]);
        const nameIndex = findHeaderIndex(headers, ["이름", "성명"]);
        const periodIndex = findHeaderIndexLoose(headers, [periodColumnName]);
        const itemIndexes = selectedItems.map((item) => ({ item, index: findHeaderIndex(headers, [item]) }));

        parsedFile.dataRows.forEach((row) => {
          const periodInfo = parsePeriodInfo(row[periodIndex]);

          if (!periodInfo || periodInfo.year !== year || periodInfo.month !== month) {
            return;
          }

          const groupInfo = getSalaryGroupInfo(headers, row, groupMode, employeeIdIndex, nameIndex);

          if (!groups.has(groupInfo.key)) {
            groups.set(groupInfo.key, { label: groupInfo.label, items: {} });
          }

          itemIndexes.forEach(({ item, index }) => {
            const amount = parseNumberValue(row[index]);
            const groupRecord = groups.get(groupInfo.key);
            groupRecord.items[item] = (groupRecord.items[item] || 0) + amount;
          });
        });
      });

      return {
        key,
        period: { year, month },
        groups,
      };
    });

    if (!dataset.periods.size) {
      throw new Error("비교 가능한 기간 데이터를 찾지 못했습니다.");
    }

    const workbook = buildComparisonWorkbook(dataset, selectedItems, periodMode);
    const outputName = `salary_comparison_${periodMode}_${groupMode}.xlsx`;
    XLSX.writeFile(workbook, outputName);
    setStatus(salaryStatus, `급여 비교분석이 완료되었습니다. ${outputName} 파일이 다운로드됩니다.`, "success");
  } catch (error) {
    setStatus(salaryStatus, error.message || "급여 비교분석 중 오류가 발생했습니다.", "error");
  }
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

function buildVacationAnnualLedgerSheet(employees, selectedYear) {
  const headers = [
    "사번",
    "이름",
    "입사일",
    "퇴사일",
    "전기잔여",
    "연차발생",
    "월차발생",
    ...Array.from({ length: 12 }, (_, index) => `${index + 1}월`),
    "합계",
    "잔여연차",
  ];
  const priorCarryColumn = 4;
  const annualAccrualColumn = 5;
  const monthlyAccrualColumn = 6;
  const januaryUsedColumn = 7;
  const totalColumn = januaryUsedColumn + 12;
  const remainingColumn = totalColumn + 1;
  const monthlyPriorCarryColumnIndex = 4;
  const monthlyCarryColumnIndex = 5;
  const monthlyAnnualColumnIndex = 6;
  const monthlyMonthlyColumnIndex = 7;
  const monthlyUsedColumnIndex = 8;

  const rows = employees.map((employee, index) => {
    const excelRow = index + 2;
    const baseCells = [
      createCell(employee.employeeId),
      createCell(employee.employeeName),
      createCell(formatDateValue(employee.joinDate)),
      createCell(formatDateValue(employee.leaveDate)),
    ];

    const januarySheetName = `${selectedYear}년1월`;
    const priorCarryCell = createFormulaCell(
      `IF('${januarySheetName}'!${XLSX.utils.encode_col(monthlyPriorCarryColumnIndex)}${excelRow}<>"",'${januarySheetName}'!${XLSX.utils.encode_col(monthlyPriorCarryColumnIndex)}${excelRow},'${januarySheetName}'!${XLSX.utils.encode_col(monthlyCarryColumnIndex)}${excelRow})`,
      "n",
      "0.00"
    );
    const annualAccrualCell = createFormulaCell(
      Array.from({ length: 12 }, (_, monthOffset) => {
        const monthNumber = monthOffset + 1;
        const sheetName = `${selectedYear}년${monthNumber}월`;
        return `'${sheetName}'!${XLSX.utils.encode_col(monthlyAnnualColumnIndex)}${excelRow}`;
      }).join("+"),
      "n",
      "0.00"
    );
    const monthlyAccrualCell = createFormulaCell(
      Array.from({ length: 12 }, (_, monthOffset) => {
        const monthNumber = monthOffset + 1;
        const sheetName = `${selectedYear}년${monthNumber}월`;
        return `'${sheetName}'!${XLSX.utils.encode_col(monthlyMonthlyColumnIndex)}${excelRow}`;
      }).join("+"),
      "n",
      "0.00"
    );
    const monthUsedCells = Array.from({ length: 12 }, (_, monthOffset) => {
      const monthNumber = monthOffset + 1;
      const sheetName = `${selectedYear}년${monthNumber}월`;
      return createFormulaCell(`'${sheetName}'!${XLSX.utils.encode_col(monthlyUsedColumnIndex)}${excelRow}`, "n", "0.00");
    });
    const totalCell = createFormulaCell(
      `SUM(${XLSX.utils.encode_col(januaryUsedColumn)}${excelRow}:${XLSX.utils.encode_col(totalColumn - 1)}${excelRow})`,
      "n",
      "0.00"
    );
    const remainingCell = createFormulaCell(
      `${XLSX.utils.encode_col(priorCarryColumn)}${excelRow}+${XLSX.utils.encode_col(annualAccrualColumn)}${excelRow}+${XLSX.utils.encode_col(monthlyAccrualColumn)}${excelRow}-${XLSX.utils.encode_col(totalColumn)}${excelRow}`,
      "n",
      "0.00"
    );

    return [...baseCells, priorCarryCell, annualAccrualCell, monthlyAccrualCell, ...monthUsedCells, totalCell, remainingCell];
  });

  const sheet = buildSheetFromStructuredRows(headers, rows);

  Array.from({ length: 12 }, (_, monthOffset) => {
    const monthNumber = monthOffset + 1;
    const columnIndex = januaryUsedColumn + monthOffset;
    const address = XLSX.utils.encode_cell({ r: 0, c: columnIndex });
    const sheetName = `${selectedYear}년${monthNumber}월`;

    if (!sheet[address]) {
      return;
    }

    sheet[address].l = {
      Target: `#'${sheetName}'!A1`,
      Tooltip: `${sheetName} 시트로 이동`,
    };
  });

  return {
    sheetName: "연간관리대장",
    sheet,
  };
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

  const sheet = buildSheetFromStructuredRows(headers, rows);
  const remainingHeaderAddress = XLSX.utils.encode_cell({ r: 0, c: remainingColumn });

  if (sheet[remainingHeaderAddress]) {
    sheet[remainingHeaderAddress].l = {
      Target: "#'연간관리대장'!A1",
      Tooltip: "연간관리대장 시트로 이동",
    };
  }

  return {
    sheetName,
    sheet,
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
    const { sheetName: annualSheetName, sheet: annualSheet } = buildVacationAnnualLedgerSheet(employees, selectedYear);
    XLSX.utils.book_append_sheet(workbook, annualSheet, annualSheetName);
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
  const periodColumnName = peoplePeriodColumnInput?.value.trim() || "";
  const salaryColumnName = peopleSalaryColumnInput?.value.trim() || "";

  if (!file) {
    setStatus(peopleStatus, "분석할 인사 엑셀 파일을 선택해야 합니다.", "error");
    return;
  }

  if (!Number.isInteger(headerRowNumber) || headerRowNumber < 1) {
    setStatus(peopleStatus, "컬럼 행 번호는 1 이상의 정수여야 합니다.", "error");
    return;
  }

  if (!periodColumnName) {
    setStatus(peopleStatus, "기간 판정에 사용할 기간 컬럼명을 입력해야 합니다.", "error");
    return;
  }

  if (!salaryColumnName) {
    setStatus(peopleStatus, "급여 평균인원 계산에 사용할 급여 컬럼명을 입력해야 합니다.", "error");
    return;
  }

  setStatus(peopleStatus, "인사 데이터를 분석하고 결과 파일을 생성하는 중입니다.");

  try {
    const periodConfig = getPeoplePeriodConfig();
    const parsedFiles = await collectFirstSheetRows([file], headerRowNumber);
    const parsedFile = parsedFiles[0];
    const headers = parsedFile.headers;
    const employeeIdIndex = findHeaderIndex(headers, ["사번"]);
    const nameIndex = findHeaderIndex(headers, ["성명", "이름"]);
    const joinIndex = findHeaderIndex(headers, ["입사일", "입사"]);
    const leaveIndex = findHeaderIndex(headers, ["퇴사일", "퇴사"]);
    const periodColumnIndex =
      findHeaderIndex(headers, [periodColumnName]) !== -1
        ? findHeaderIndex(headers, [periodColumnName])
        : findHeaderIndexLoose(headers, [periodColumnName]);
    const salaryColumnIndex =
      findHeaderIndex(headers, [salaryColumnName]) !== -1
        ? findHeaderIndex(headers, [salaryColumnName])
        : findHeaderIndexLoose(headers, [salaryColumnName]);

    if (employeeIdIndex === -1) {
      throw new Error("필수 컬럼이 없습니다. 사번 컬럼은 반드시 있어야 합니다.");
    }

    if (nameIndex === -1) {
      throw new Error("필수 컬럼이 없습니다. 성명 또는 이름 컬럼 중 하나는 있어야 합니다.");
    }

    if (joinIndex === -1 || leaveIndex === -1) {
      throw new Error("필수 컬럼이 없습니다. 입사일(또는 입사), 퇴사일(또는 퇴사)이 필요합니다.");
    }

    if (periodColumnIndex === -1) {
      throw new Error(`입력한 기간 컬럼명(${periodColumnName})을 파일에서 찾을 수 없습니다.`);
    }

    if (salaryColumnIndex === -1) {
      throw new Error(`입력한 급여 컬럼명(${salaryColumnName})을 파일에서 찾을 수 없습니다.`);
    }

    const employees = [];
    const seenEmployeeIds = new Set();
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

      if (employeeKey && !seenEmployeeIds.has(employeeKey) && joinDate) {
        seenEmployeeIds.add(employeeKey);
        employees.push({
          employeeId: employeeKey,
          employeeName: String(row[nameIndex]?.v ?? "").trim(),
          joinDate,
          leaveDate,
        });
      }

      if (joinDate && joinDate >= periodConfig.startDate && joinDate <= periodConfig.endDate) {
        pushUniqueEmployee(joiners, joinerKeys, employeeKey, row);
      }

      if (leaveDate && leaveDate >= periodConfig.startDate && leaveDate <= periodConfig.endDate) {
        pushUniqueEmployee(leavers, leaverKeys, employeeKey, row);
      }

      if (isActiveAtMonthEnd(joinDate, leaveDate, periodConfig.endDate)) {
        pushUniqueEmployee(activeEmployees, activeKeys, employeeKey, row);
      }
    });

    const analysisMonths = getAvailablePeriodMonths(parsedFile.dataRows, periodColumnIndex, periodConfig);

    if (!analysisMonths.length) {
      throw new Error("선택한 기간에 해당하는 데이터 월을 기간 컬럼에서 찾을 수 없습니다.");
    }

    const salaryAverageHeadcount = calculateSalaryAverageHeadcount(
      parsedFile.dataRows,
      employeeIdIndex,
      periodColumnIndex,
      salaryColumnIndex,
      periodConfig,
      analysisMonths
    );
    const periodAverageHeadcount = calculatePeriodAverageHeadcount(
      employees,
      activeEmployees.length,
      periodConfig,
      analysisMonths
    );

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(
      workbook,
      createSummarySheet(
        periodConfig.label,
        joiners,
        leavers,
        activeEmployees,
        salaryAverageHeadcount,
        periodAverageHeadcount
      ),
      "요약"
    );
    XLSX.utils.book_append_sheet(workbook, buildSheetFromStructuredRows(headers, joiners), "입사자");
    XLSX.utils.book_append_sheet(workbook, buildSheetFromStructuredRows(headers, leavers), "퇴사자");
    XLSX.utils.book_append_sheet(workbook, buildSheetFromStructuredRows(headers, activeEmployees), "재직자");

    const outputName = `people_analysis_${periodConfig.mode}_${periodConfig.label.replace(/[^\w가-힣-]/g, "_")}.xlsx`;
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

function parseWorkstatDate(cell) {
  if (!cell || cell.v == null || cell.v === "") {
    return null;
  }

  if (cell.v instanceof Date) {
    return new Date(cell.v.getFullYear(), cell.v.getMonth(), cell.v.getDate());
  }

  if (typeof cell.v === "number" && cell.v > 0 && cell.v < 90000) {
    const epoch = new Date(Date.UTC(1899, 11, 30));
    epoch.setUTCDate(epoch.getUTCDate() + Math.floor(cell.v));
    return new Date(epoch.getUTCFullYear(), epoch.getUTCMonth(), epoch.getUTCDate());
  }

  return parseExcelDate(cell);
}

function buildWorkstatRecords(parsedFile) {
  const headers = parsedFile.headers;
  const findIndexByCandidates = (candidates) =>
    headers.findIndex((header) => headerIncludesCandidate(header, candidates));

  const workDateIndex = findIndexByCandidates(["work_date", "workdate"]);
  const hoursIndex = findIndexByCandidates(["hours"]);
  const empIdIndex = findIndexByCandidates(["emp_id", "empid", "사번", "직원번호"]);
  const yearIndex = findIndexByCandidates(["년"]);
  const monthIndex = findIndexByCandidates(["월"]);
  const dayIndex = findIndexByCandidates(["일"]);
  const koreanHoursIndex = findIndexByCandidates(["근무시간"]);

  const records = [];

  if (workDateIndex !== -1 && hoursIndex !== -1) {
    parsedFile.dataRows.forEach((row) => {
      const workDate = parseWorkstatDate(row[workDateIndex]);

      if (!workDate) {
        return;
      }

      records.push({
        empId: String(row[empIdIndex]?.v ?? "ALL").trim() || "ALL",
        workDate,
        hours: parseNumberValue(row[hoursIndex]),
      });
    });

    return records;
  }

  if (yearIndex !== -1 && monthIndex !== -1 && dayIndex !== -1 && koreanHoursIndex !== -1) {
    parsedFile.dataRows.forEach((row) => {
      const year = parseNumberValue(row[yearIndex]);
      const month = parseNumberValue(row[monthIndex]);
      const day = parseNumberValue(row[dayIndex]);

      if (!year || !month || !day) {
        return;
      }

      const workDate = new Date(year, month - 1, day);

      if (Number.isNaN(workDate.getTime())) {
        return;
      }

      records.push({
        empId: String(row[empIdIndex]?.v ?? "ALL").trim() || "ALL",
        workDate,
        hours: parseNumberValue(row[koreanHoursIndex]),
      });
    });

    return records;
  }

  throw new Error("헤더 형식을 인식하지 못했습니다. `work_date` + `hours` 또는 `년`, `월`, `일`, `근무시간` 형식이 필요합니다.");
}

function addDays(date, days) {
  const next = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  next.setDate(next.getDate() + days);
  return next;
}

function roundToOneDecimal(value) {
  return Math.round(value * 10) / 10;
}

function compareWorkstatRows(left, right) {
  return (
    left.empId.localeCompare(right.empId, "ko") ||
    left.block - right.block
  );
}

function finalizeWorkstatRows(groupMap, dateFactory) {
  const sequenceMap = new Map();

  return Array.from(groupMap.values())
    .sort(compareWorkstatRows)
    .map((entry) => {
      const totalHours = roundToOneDecimal(entry.totalHours);
      const averageHours = roundToOneDecimal(totalHours / 4);
      const nextSequence = (sequenceMap.get(entry.empId) || 0) + 1;
      sequenceMap.set(entry.empId, nextSequence);
      const { fromDate, toDate } = dateFactory(entry.block);

      return {
        SEQ: nextSequence,
        FROM: formatDateValue(fromDate),
        TO: formatDateValue(toDate),
        "4주근무시간": totalHours,
        "4주근무시간/4": averageHours,
        퇴직산정주여부: averageHours >= 15 ? 1 : 0,
        emp_id: entry.empId,
      };
    });
}

function calculateWorkstatForward(records, startDate) {
  const filtered = records.filter((record) => record.workDate >= startDate);

  if (!filtered.length) {
    throw new Error("시작일 이후 데이터가 없습니다.");
  }

  const groups = new Map();

  filtered.forEach((record) => {
    const block = Math.floor((record.workDate - startDate) / (1000 * 60 * 60 * 24 * 28));
    const key = `${record.empId}__${block}`;

    if (!groups.has(key)) {
      groups.set(key, { empId: record.empId, block, totalHours: 0 });
    }

    groups.get(key).totalHours += record.hours;
  });

  return finalizeWorkstatRows(groups, (block) => {
    const fromDate = addDays(startDate, block * 28);
    return { fromDate, toDate: addDays(fromDate, 27) };
  });
}

function calculateWorkstatBackward(records) {
  const lastWorkedRecord = records
    .filter((record) => record.hours > 0)
    .sort((left, right) => right.workDate - left.workDate)[0];

  if (!lastWorkedRecord) {
    throw new Error("근무시간이 0보다 큰 행이 없습니다.");
  }

  const endDate = lastWorkedRecord.workDate;
  const filtered = records.filter((record) => record.workDate <= endDate);
  const groups = new Map();

  filtered.forEach((record) => {
    const block = Math.floor((endDate - record.workDate) / (1000 * 60 * 60 * 24 * 28));
    const key = `${record.empId}__${block}`;

    if (!groups.has(key)) {
      groups.set(key, { empId: record.empId, block, totalHours: 0 });
    }

    groups.get(key).totalHours += record.hours;
  });

  return finalizeWorkstatRows(groups, (block) => {
    const toDate = addDays(endDate, -block * 28);
    return { fromDate: addDays(toDate, -27), toDate };
  });
}

function buildWorkstatSheet(rows) {
  const headers = ["SEQ", "FROM", "TO", "4주근무시간", "4주근무시간/4", "퇴직산정주여부", "emp_id"];
  const matrix = [
    headers,
    ...rows.map((row) => headers.map((header) => row[header])),
  ];

  return XLSX.utils.aoa_to_sheet(matrix);
}

function renderWorkstatPreview(container, rows) {
  const headers = ["SEQ", "FROM", "TO", "4주근무시간", "4주근무시간/4", "퇴직산정주여부", "emp_id"];
  const previewRows = rows.slice(0, 20);

  if (!previewRows.length) {
    container.innerHTML = "<p>표시할 결과가 없습니다.</p>";
    return;
  }

  container.innerHTML = `
    <table class="result-table">
      <thead>
        <tr>${headers.map((header) => `<th>${header}</th>`).join("")}</tr>
      </thead>
      <tbody>
        ${previewRows
          .map(
            (row) => `
              <tr>${headers.map((header) => `<td>${row[header] ?? ""}</td>`).join("")}</tr>
            `
          )
          .join("")}
      </tbody>
    </table>
  `;
}

function hideWorkstatResult() {
  workstatResult?.classList.add("is-hidden");

  if (workstatForwardPreview) {
    workstatForwardPreview.innerHTML = "";
  }

  if (workstatBackwardPreview) {
    workstatBackwardPreview.innerHTML = "";
  }
}

async function handleWorkstatAnalysis() {
  const file = workstatFileInput.files?.[0];
  const headerRowNumber = Number(workstatHeaderRowInput.value);
  const startValue = workstatStartDateInput.value;

  if (!file) {
    setStatus(workstatStatus, "주휴 계산에 사용할 근무기록 엑셀 파일을 선택해야 합니다.", "error");
    return;
  }

  if (!Number.isInteger(headerRowNumber) || headerRowNumber < 1) {
    setStatus(workstatStatus, "컬럼 행 번호는 1 이상의 정수여야 합니다.", "error");
    return;
  }

  if (!startValue) {
    setStatus(workstatStatus, "시작일을 입력해야 합니다.", "error");
    return;
  }

  const [year, month, day] = startValue.split("-").map(Number);
  const startDate = new Date(year, month - 1, day);

  if (Number.isNaN(startDate.getTime())) {
    setStatus(workstatStatus, "시작일 형식이 올바르지 않습니다.", "error");
    return;
  }

  setStatus(workstatStatus, "근무기록을 분석하고 주휴 계산 파일을 생성하는 중입니다.");

  try {
    const parsedFiles = await collectFirstSheetRows([file], headerRowNumber);
    const records = buildWorkstatRecords(parsedFiles[0]);
    const forwardRows = calculateWorkstatForward(records, startDate);
    const backwardRows = calculateWorkstatBackward(records);
    const workbook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(workbook, buildWorkstatSheet(forwardRows), "정방향");
    XLSX.utils.book_append_sheet(workbook, buildWorkstatSheet(backwardRows), "역방향");

    renderWorkstatPreview(workstatForwardPreview, forwardRows);
    renderWorkstatPreview(workstatBackwardPreview, backwardRows);
    workstatResult?.classList.remove("is-hidden");

    const outputName = `weekly_holiday_${startValue}.xlsx`;
    XLSX.writeFile(workbook, outputName);
    setStatus(workstatStatus, `주휴 계산이 완료되었습니다. ${outputName} 파일이 다운로드됩니다.`, "success");
  } catch (error) {
    hideWorkstatResult();
    setStatus(workstatStatus, error.message || "주휴 계산 중 오류가 발생했습니다.", "error");
  }
}

function formatNumber(value, digits = 0) {
  if (!Number.isFinite(value)) {
    return "-";
  }

  return value.toLocaleString("ko-KR", {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits,
  });
}

function roundDownToTens(value) {
  return Math.floor((value || 0) / 10) * 10;
}

function roundUpToTens(value) {
  return Math.ceil((value || 0) / 10) * 10;
}

function parseDateInputValue(value) {
  if (!value) {
    return null;
  }

  const [year, month, day] = value.split("-").map(Number);
  const date = new Date(year, month - 1, day);
  return Number.isNaN(date.getTime()) ? null : date;
}

function endOfMonth(date) {
  return new Date(date.getFullYear(), date.getMonth() + 1, 0);
}

function startOfMonth(date) {
  return new Date(date.getFullYear(), date.getMonth(), 1);
}

function addMonthsClamped(date, months) {
  const year = date.getFullYear();
  const month = date.getMonth() + months;
  const targetYear = year + Math.floor(month / 12);
  const targetMonth = ((month % 12) + 12) % 12;
  const lastDay = new Date(targetYear, targetMonth + 1, 0).getDate();
  return new Date(targetYear, targetMonth, Math.min(date.getDate(), lastDay));
}

function diffDaysInclusive(fromDate, toDate) {
  return Math.floor((toDate - fromDate) / (1000 * 60 * 60 * 24)) + 1;
}

function buildSeverancePeriods(causeDate) {
  const overallStart = addMonthsClamped(causeDate, -3);
  const overallEnd = addDays(causeDate, -1);
  const periods = [];
  let cursor = overallStart;

  while (cursor <= overallEnd && periods.length < 4) {
    const periodEnd = endOfMonth(cursor) < overallEnd ? endOfMonth(cursor) : overallEnd;
    periods.push({
      from: new Date(cursor.getFullYear(), cursor.getMonth(), cursor.getDate()),
      to: periodEnd,
      days: diffDaysInclusive(cursor, periodEnd),
    });
    cursor = addDays(periodEnd, 1);
  }

  return {
    overallStart,
    overallEnd,
    totalDays: diffDaysInclusive(overallStart, overallEnd),
    periods,
  };
}

function renderSeverancePeriodLabels() {
  const endDate = parseDateInputValue(severanceEndDateInput?.value);

  severancePeriodLabels.forEach((label, index) => {
    if (label) {
      label.textContent = `${index + 1}구간 급여`;
    }
  });

  if (!endDate) {
    return;
  }

  const causeDate = addDays(endDate, 1);
  const { periods } = buildSeverancePeriods(causeDate);

  periods.forEach((period, index) => {
    if (severancePeriodLabels[index]) {
      severancePeriodLabels[index].textContent = `${formatDateValue(period.from)} ~ ${formatDateValue(period.to)} 급여`;
    }
  });
}

function hideSeveranceResult() {
  severanceResult?.classList.add("is-hidden");
  severanceSavePdfButton?.setAttribute("disabled", "disabled");

  if (severanceSummary) {
    severanceSummary.innerHTML = "";
  }

  if (severancePeriodTable) {
    severancePeriodTable.innerHTML = "";
  }
}

function buildSeverancePdfHtml() {
  const summaryHtml = severanceSummary?.innerHTML || "";
  const tableHtml = severancePeriodTable?.outerHTML || "";
  const title = severanceNameInput?.value?.trim()
    ? `${severanceNameInput.value.trim()} 퇴직금 계산내역`
    : "퇴직금 계산내역";

  return `<!DOCTYPE html>
<html lang="ko">
  <head>
    <meta charset="UTF-8" />
    <title>${title}</title>
    <style>
      body { font-family: "Malgun Gothic", "Apple SD Gothic Neo", sans-serif; margin: 24px; color: #2f261f; }
      h1 { margin: 0 0 16px; font-size: 24px; }
      .meta { margin-bottom: 20px; color: #6f5b4f; font-size: 12px; }
      .severance-summary { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 16px; margin-bottom: 20px; }
      .severance-card { border: 1px solid #d9c8bc; border-radius: 12px; padding: 16px; break-inside: avoid; }
      .severance-card h3 { margin: 0 0 12px; font-size: 16px; }
      .severance-metric { display: flex; justify-content: space-between; gap: 12px; padding: 7px 0; border-bottom: 1px solid #eee2d8; }
      .severance-metric:last-child { border-bottom: 0; }
      .result-table { width: 100%; border-collapse: collapse; }
      .result-table th, .result-table td { border: 1px solid #d9c8bc; padding: 10px; text-align: left; font-size: 12px; }
      .result-table th { background: #f7efe8; }
      @media print {
        body { margin: 12mm; }
        button { display: none; }
      }
    </style>
  </head>
  <body>
    <h1>${title}</h1>
    <div class="meta">출력일 ${formatDateValue(new Date())}</div>
    <div class="severance-summary">${summaryHtml}</div>
    ${tableHtml}
    <script>
      window.onload = function() {
        window.print();
      };
    </script>
  </body>
</html>`;
}

function handleSeverancePdfSave() {
  if (severanceResult?.classList.contains("is-hidden")) {
    setStatus(severanceStatus, "먼저 퇴직금 계산을 완료한 뒤 PDF 저장을 실행해 주세요.", "error");
    return;
  }

  const printWindow = window.open("", "_blank", "width=960,height=1200");
  if (!printWindow) {
    setStatus(severanceStatus, "팝업이 차단되어 PDF 저장 창을 열 수 없습니다. 팝업 허용 후 다시 시도해 주세요.", "error");
    return;
  }

  printWindow.document.open();
  printWindow.document.write(buildSeverancePdfHtml());
  printWindow.document.close();
}

function getNumericInputValue(input) {
  const normalizedValue = String(input?.value || "").replace(/,/g, "").trim();
  return Number(normalizedValue || 0) || 0;
}

function formatNumericInput(input, allowDecimal = false) {
  if (!input) {
    return;
  }

  const rawValue = String(input.value || "").replace(/,/g, "");
  const sanitizedValue = allowDecimal
    ? rawValue.replace(/[^\d.]/g, "").replace(/(\..*)\./g, "$1")
    : rawValue.replace(/\D/g, "");

  if (!sanitizedValue) {
    input.value = "";
    return;
  }

  if (allowDecimal) {
    const [integerPart, decimalPart] = sanitizedValue.split(".");
    const formattedInteger = (integerPart || "0").replace(/\B(?=(\d{3})+(?!\d))/g, ",");
    input.value = decimalPart !== undefined ? `${formattedInteger}.${decimalPart}` : formattedInteger;
    return;
  }

  input.value = sanitizedValue.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

function buildSeveranceSummaryCard(title, metrics) {
  return `
    <section class="severance-card">
      <h3>${title}</h3>
      ${metrics
        .map(
          ({ label, value }) => `
            <div class="severance-metric">
              <span>${label}</span>
              <strong>${value}</strong>
            </div>
          `
        )
        .join("")}
    </section>
  `;
}

function renderSeveranceResult(result) {
  severanceSummary.innerHTML = [
    buildSeveranceSummaryCard("기본 정보", [
      { label: "작성일", value: formatDateValue(new Date()) },
      { label: "산정사유발생일", value: formatDateValue(result.causeDate) },
      { label: "산정 대상 기간", value: `${formatDateValue(result.overallStart)} ~ ${formatDateValue(result.overallEnd)}` },
      { label: "산정 일수", value: `${formatNumber(result.eligibleDays)}일` },
      { label: "고용형태", value: severanceEmploymentTypeInput.value.trim() || "-" },
      { label: "성명", value: severanceNameInput.value.trim() || "-" },
    ]),
    buildSeveranceSummaryCard("임금 계산", [
      { label: "총임금액", value: `${formatNumber(result.totalWages)}원` },
      { label: "평균임금", value: `${formatNumber(result.averageWage)}원` },
      { label: "통상임금", value: `${formatNumber(result.ordinaryWage)}원` },
      { label: "적용기준", value: result.wageBasisLabel },
      { label: "기준금액", value: `${formatNumber(result.wageBasisAmount)}원` },
      { label: "퇴직금", value: `${formatNumber(result.retirementPay)}원` },
    ]),
    buildSeveranceSummaryCard("공제 반영", [
      { label: "퇴직소득총액", value: `${formatNumber(result.totalRetirementIncome)}원` },
      { label: "퇴직주민세", value: `${formatNumber(result.localTax)}원` },
      { label: "일반공제 합계", value: `${formatNumber(result.generalDeductionTotal)}원` },
      { label: "기타공제 합계", value: `${formatNumber(result.otherDeductionsTotal)}원` },
      { label: "일반공제 후 총액", value: `${formatNumber(result.afterGeneralDeduction)}원` },
      { label: "실지급대상", value: `${formatNumber(result.netPay)}원` },
    ]),
  ].join("");

  severancePeriodTable.innerHTML = `
    <thead>
      <tr>
        <th>구간</th>
        <th>기간</th>
        <th>일수</th>
        <th>급여(제수당)</th>
      </tr>
    </thead>
    <tbody>
      ${result.periods
        .map(
          (period, index) => `
            <tr>
              <td>${index + 1}구간</td>
              <td>${formatDateValue(period.from)} ~ ${formatDateValue(period.to)}</td>
              <td>${formatNumber(period.days)}일</td>
              <td>${formatNumber(period.wage, 2)}원</td>
            </tr>
          `
        )
        .join("")}
      <tr>
        <td>상여금</td>
        <td>3/12 반영</td>
        <td>-</td>
        <td>${formatNumber(result.bonusAdjusted)}원</td>
      </tr>
      <tr>
        <td>연차수당</td>
        <td>3/12 반영</td>
        <td>-</td>
        <td>${formatNumber(result.vacationAdjusted)}원</td>
      </tr>
      <tr>
        <td>기타수당</td>
        <td>${formatNumber(result.extraPayNumerator)} / ${formatNumber(result.extraPayDenominator)}</td>
        <td>-</td>
        <td>${formatNumber(result.extraAdjusted)}원</td>
      </tr>
    </tbody>
  `;

  severanceResult?.classList.remove("is-hidden");
}

function handleSeveranceCalculation() {
  const joinDate = parseDateInputValue(severanceJoinDateInput.value);
  const endDate = parseDateInputValue(severanceEndDateInput.value);
  const midStartDate = parseDateInputValue(severanceMidStartInput.value);
  const midEndDate = parseDateInputValue(severanceMidEndInput.value);

  if (!joinDate || !endDate) {
    setStatus(severanceStatus, "입사일과 종료일은 반드시 입력해야 합니다.", "error");
    hideSeveranceResult();
    return;
  }

  if (endDate < joinDate) {
    setStatus(severanceStatus, "종료일은 입사일보다 빠를 수 없습니다.", "error");
    hideSeveranceResult();
    return;
  }

  if ((midStartDate && !midEndDate) || (!midStartDate && midEndDate)) {
    setStatus(severanceStatus, "중간정산 기간은 시작일과 종료일을 모두 입력해야 합니다.", "error");
    hideSeveranceResult();
    return;
  }

  if (midStartDate && midEndDate && midEndDate < midStartDate) {
    setStatus(severanceStatus, "중간정산 종료일은 시작일보다 빠를 수 없습니다.", "error");
    hideSeveranceResult();
    return;
  }

  const causeDate = addDays(endDate, 1);
  const { overallStart, overallEnd, totalDays, periods } = buildSeverancePeriods(causeDate);
  const eligibleDays = totalDays;
  const serviceStartDate = midEndDate ? addDays(midEndDate, 1) : joinDate;
  const serviceEndDate = endDate;

  if (serviceStartDate > serviceEndDate) {
    setStatus(severanceStatus, "중간정산 종료일 이후의 계속근로기간이 없습니다.", "error");
    hideSeveranceResult();
    return;
  }

  const serviceDays = diffDaysInclusive(serviceStartDate, serviceEndDate);

  if (eligibleDays <= 0) {
    setStatus(severanceStatus, "평균임금 산정일수는 1일 이상이어야 합니다.", "error");
    hideSeveranceResult();
    return;
  }

  if (serviceDays < 365) {
    setStatus(
      severanceStatus,
      "법정 퇴직금은 계속근로기간이 1년 이상인 경우에만 발생합니다. 중간정산 이후 계속근로기간을 확인해 주세요.",
      "error"
    );
    hideSeveranceResult();
    return;
  }

  const wages = severanceWageInputs.map((input) => getNumericInputValue(input));
  const periodsWithWages = periods.map((period, index) => ({
    ...period,
    wage: wages[index] || 0,
  }));
  const periodWageTotal = periodsWithWages.reduce((sum, period) => sum + period.wage, 0);
  const bonusAdjusted = Math.round((getNumericInputValue(severanceBonusInput) * 3) / 12);
  const vacationAdjusted = Math.round((getNumericInputValue(severanceVacationPayInput) * 3) / 12);
  const totalWages = periodWageTotal + bonusAdjusted + vacationAdjusted;
  const averageWage = Math.round(totalWages / eligibleDays);
  const retirementPay = roundUpToTens((averageWage * 30 * serviceDays) / 365);
  const totalRetirementIncome = retirementPay;

  severanceSummary.innerHTML = [
    buildSeveranceSummaryCard("기본 정보", [
      { label: "작성일", value: formatDateValue(new Date()) },
      { label: "퇴직일", value: formatDateValue(serviceEndDate) },
      { label: "계속근로 시작일", value: formatDateValue(serviceStartDate) },
      { label: "계속근로기간", value: `${formatNumber(serviceDays)}일` },
      { label: "평균임금 산정 기간", value: `${formatDateValue(overallStart)} ~ ${formatDateValue(overallEnd)}` },
      { label: "평균임금 산정일수", value: `${formatNumber(eligibleDays)}일` },
      { label: "고용형태", value: severanceEmploymentTypeInput.value.trim() || "-" },
      { label: "성명", value: severanceNameInput.value.trim() || "-" },
    ]),
    buildSeveranceSummaryCard("법정 퇴직금 계산", [
      { label: "총임금액", value: `${formatNumber(totalWages)}원` },
      { label: "평균임금", value: `${formatNumber(averageWage)}원` },
      { label: "법정 산정식", value: "평균임금 30일분 x 계속근로기간 / 365" },
      { label: "퇴직금", value: `${formatNumber(retirementPay)}원` },
    ]),
    buildSeveranceSummaryCard("계산 결과", [
      { label: "퇴직소득총액", value: `${formatNumber(totalRetirementIncome)}원` },
    ]),
  ].join("");

  severancePeriodTable.innerHTML = `
    <thead>
      <tr>
        <th>구간</th>
        <th>기간</th>
        <th>일수</th>
        <th>급여(세전)</th>
      </tr>
    </thead>
    <tbody>
      ${periodsWithWages
        .map(
          (period, index) => `
            <tr>
              <td>${index + 1}구간</td>
              <td>${formatDateValue(period.from)} ~ ${formatDateValue(period.to)}</td>
              <td>${formatNumber(period.days)}일</td>
              <td>${formatNumber(period.wage, 2)}원</td>
            </tr>
          `
        )
        .join("")}
      <tr>
        <td>상여금</td>
        <td>3/12 반영</td>
        <td>-</td>
        <td>${formatNumber(bonusAdjusted)}원</td>
      </tr>
      <tr>
        <td>연차수당</td>
        <td>3/12 반영</td>
        <td>-</td>
        <td>${formatNumber(vacationAdjusted)}원</td>
      </tr>
    </tbody>
  `;

  severanceResult?.classList.remove("is-hidden");
  severanceSavePdfButton?.removeAttribute("disabled");
  setStatus(
    severanceStatus,
    "법정 퇴직금 기준으로 계산이 완료되었습니다. 평균임금 산정 기간과 계속근로기간을 함께 확인해 주세요.",
    "success"
  );
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
    "파일과 분석 구분, 기준 기간, 기간 컬럼명, 급여 컬럼명을 입력한 뒤 분석 파일 생성을 실행하세요.",
    "컬럼 행과 기준 기간, 기간 컬럼명, 급여 컬럼명을 입력하고 실행하세요."
  );
  setStatus(peopleStatus, message);
});

peoplePeriodModeInput?.addEventListener("change", updatePeoplePeriodControls);

vacationFileInput?.addEventListener("change", () => {
  const message = renderSelectedFiles(
    vacationFileInput,
    vacationFileList,
    "파일과 생성 연도, 기준 방식을 입력한 뒤 연차관리대장 생성을 실행하세요.",
    "생성 연도와 기준 방식을 입력하고 실행하세요."
  );
  setStatus(vacationStatus, message);
});

salaryFilesInput?.addEventListener("change", async () => {
  const message = renderSelectedFiles(
    salaryFilesInput,
    salaryFileList,
    "파일을 선택하고 비교 항목과 기준을 입력한 뒤 비교 파일 생성을 실행하세요.",
    "컬럼 행과 비교 기준을 입력한 뒤 비교 항목을 선택하세요."
  );
  setStatus(salaryStatus, message);
  await refreshSalaryHeaderOptions();
});

salaryHeaderRowInput?.addEventListener("change", async () => {
  await refreshSalaryHeaderOptions();
});

salarySelectAllButton?.addEventListener("click", () => {
  const checkboxes = Array.from(document.querySelectorAll(".salary-item-option"));
  const allChecked = checkboxes.length > 0 && checkboxes.every((checkbox) => checkbox.checked);

  checkboxes.forEach((checkbox) => {
    checkbox.checked = !allChecked;
  });

  salarySelectAllButton.textContent = allChecked ? "전체선택" : "전체해제";
});

salaryItemPicker?.addEventListener("change", () => {
  const checkboxes = Array.from(document.querySelectorAll(".salary-item-option"));
  const allChecked = checkboxes.length > 0 && checkboxes.every((checkbox) => checkbox.checked);

  if (salarySelectAllButton) {
    salarySelectAllButton.textContent = allChecked ? "전체해제" : "전체선택";
  }
});

workstatFileInput?.addEventListener("change", () => {
  hideWorkstatResult();
  const message = renderSelectedFiles(
    workstatFileInput,
    workstatFileList,
    "파일과 시작일을 입력한 뒤 주휴 계산 파일 생성을 실행하세요.",
    "컬럼 행과 시작일을 입력하고 실행하세요."
  );
  setStatus(workstatStatus, message);
});

workstatHeaderRowInput?.addEventListener("change", hideWorkstatResult);
workstatStartDateInput?.addEventListener("change", hideWorkstatResult);
severanceJoinDateInput?.addEventListener("change", () => {
  hideSeveranceResult();
  renderSeverancePeriodLabels();
});
severanceEndDateInput?.addEventListener("change", () => {
  hideSeveranceResult();
  renderSeverancePeriodLabels();
});
[
  severanceCompanyInput,
  severanceSiteInput,
  severanceNameInput,
  severanceBirthDateInput,
  severanceEmploymentTypeInput,
  severanceMidStartInput,
  severanceMidEndInput,
  severanceBonusInput,
  severanceVacationPayInput,
  ...severanceWageInputs,
].forEach((input) => {
  input?.addEventListener("input", hideSeveranceResult);
});

severanceWageInputs.forEach((input) => {
  input?.addEventListener("input", () => formatNumericInput(input, true));
});

[severanceBonusInput, severanceVacationPayInput].forEach((input) => {
  input?.addEventListener("input", () => formatNumericInput(input, false));
});

document.querySelectorAll('input[name="split-mode"]').forEach((radio) => {
  radio.addEventListener("change", updateSplitOptionFields);
});

mergeRunButton?.addEventListener("click", handleMerge);
splitRunButton?.addEventListener("click", handleSplit);
peopleRunButton?.addEventListener("click", handlePeopleAnalysis);
vacationRunButton?.addEventListener("click", handleVacationLedger);
salaryRunButton?.addEventListener("click", handleSalaryAnalysis);
workstatRunButton?.addEventListener("click", handleWorkstatAnalysis);
severanceRunButton?.addEventListener("click", handleSeveranceCalculation);
severanceSavePdfButton?.addEventListener("click", handleSeverancePdfSave);

updateSplitOptionFields();
updatePeoplePeriodControls();
renderSeverancePeriodLabels();
