"use strict";

const dashboardColumns = ["Site", "Sprint", "Start", "End", "Planned Hrs", "Actual Hrs", "Planned %", "Actual %", "Variance", "Delay days", "Health", "Blocked"];
const dashboardSummaryColumns = ["TOTAL Sites", "PLANNED HRS", "COMPLETED HRS", "OVERALL ACTUAL %", "DELAYED / AT RISK"];

function readDashboardWorksheet(workbook) {
  const normalize = value => String(value ?? "").trim().replace(/\s+/g, " ").toLowerCase();
  const sheetName = workbook.SheetNames.find(name => normalize(name) === "dashboard");
  if (!sheetName) throw new Error('Worksheet "Dashboard" is missing. Add it to populate the sprint delivery overview.');
  const sheet = workbook.Sheets[sheetName];
  const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: true });
  const headerIndex = matrix.findIndex(row => dashboardColumns.every(name => row.some(value => normalize(value) === normalize(name))));
  if (headerIndex < 0) throw new Error(`Worksheet "Dashboard" must contain these columns: ${dashboardColumns.join(", ")}.`);
  const columns = dashboardColumns.map(name => matrix[headerIndex].findIndex(value => normalize(value) === normalize(name)));
  const commentsColumn = matrix[headerIndex].findIndex(value => normalize(value) === "comments");
  const origin = XLSX.utils.decode_range(sheet["!ref"]).s;
  const firstRow = origin.r;
  const numberAt = (value, row, column, label, percentage = false) => {
    if (String(value ?? "").trim() === "") return null;
    let number = typeof value === "number" ? value : Number(String(value).trim().replace(/,/g, "").replace(/%$/, ""));
    const cell = sheet[XLSX.utils.encode_cell({ r: firstRow + row, c: origin.c + column })];
    // Excel stores percentage-formatted cells as fractions (0.2 means 20%).
    if (percentage && typeof value === "number" && (cell?.w?.includes("%") || cell?.z?.includes("%"))) number *= 100;
    if (!Number.isFinite(number)) throw new Error(`Dashboard row ${firstRow + row + 1}: ${label} must contain a number.`);
    return number;
  };
  const summaryHeaders = dashboardSummaryColumns;
  const summaryHeaderIndex = matrix.findIndex(row => summaryHeaders.every(name => row.some(value => normalize(value) === normalize(name))));
  const summaryRows = new Set();
  const labelNames = new Set(summaryHeaders.map(normalize));
  const labeledValues = summaryHeaders.map(label => {
    for (let row = 0; row < matrix.length; row++) {
      if (row === headerIndex) continue;
      const column = matrix[row].findIndex(value => normalize(value) === normalize(label));
      if (column < 0) continue;
      let valueRow = row, valueColumn = column + 1;
      // A summary header may label a value beside it, including merged-cell gaps.
      while (valueColumn < matrix[row].length && String(matrix[row][valueColumn] ?? "").trim() === "") valueColumn++;
      const rightValue = matrix[row][valueColumn];
      if (rightValue == null || labelNames.has(normalize(rightValue))) {
        valueRow = row + 1;
        valueColumn = column;
      }
      const value = matrix[valueRow]?.[valueColumn];
      if (valueRow === headerIndex || labelNames.has(normalize(value))) continue;
      // Keep summary sections out of the site table, without hiding site rows
      // when the summary is positioned to the right of that table.
      if (columns.includes(column)) summaryRows.add(row);
      if (columns.includes(valueColumn)) summaryRows.add(valueRow);
      return { value, row: valueRow, column: valueColumn, label };
    }
    return null;
  });
  const hasLabeledSummary = labeledValues.every(Boolean);
  let summary = null, summaryValueIndex = -1;
  if (hasLabeledSummary) {
    const keys = ["totalSites", "plannedHours", "completedHours", "overallActual", "delayedAtRisk"];
    summary = Object.fromEntries(keys.map((key, index) => {
      const entry = labeledValues[index];
      return [key, numberAt(entry.value, entry.row, entry.column, entry.label, key === "overallActual")];
    }));
  } else if (summaryHeaderIndex >= 0) {
    summaryRows.clear();
    const summaryColumns = summaryHeaders.map(name => matrix[summaryHeaderIndex].findIndex(value => normalize(value) === normalize(name)));
    // Read the first summary value row; never sum repeated workbook totals.
    const end = summaryHeaderIndex < headerIndex ? headerIndex : matrix.length;
    for (let row = summaryHeaderIndex + 1; row < end; row++) {
      if (summaryColumns.some(column => String(matrix[row][column] ?? "").trim() !== "")) {
        summaryValueIndex = row;
        break;
      }
    }
    if (summaryValueIndex >= 0) {
      const keys = ["totalSites", "plannedHours", "completedHours", "overallActual", "delayedAtRisk"];
      summary = Object.fromEntries(keys.map((key, index) => [key,
        numberAt(matrix[summaryValueIndex][summaryColumns[index]], summaryValueIndex, summaryColumns[index], summaryHeaders[index], key === "overallActual")]));
    }
  }
  const date1904 = Boolean(workbook.Workbook?.WBProps?.date1904);
  const sites = matrix.slice(headerIndex + 1).flatMap((row, index) => {
    const matrixRow = headerIndex + index + 1;
    if (summaryRows.has(matrixRow) || (summaryHeaderIndex !== headerIndex && (matrixRow === summaryHeaderIndex || matrixRow === summaryValueIndex))) return [];
    const values = columns.map(column => row[column] ?? "");
    if (values.every(value => String(value).trim() === "")) return [];
    const rowIndex = firstRow + headerIndex + index + 1;
    const numeric = (position, percentage = false) => numberAt(values[position], matrixRow, columns[position], dashboardColumns[position], percentage);
    const date = position => {
      const value = values[position];
      if (typeof value !== "number") return String(value).trim();
      const parsed = XLSX.SSF.parse_date_code(value, { date1904 });
      if (!parsed) throw new Error(`Dashboard row ${rowIndex + 1}: ${dashboardColumns[position]} contains an invalid date.`);
      const month = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"][parsed.m - 1];
      return `${String(parsed.d).padStart(2, "0")}-${month}-${parsed.y}`;
    };
    return [{ site: String(values[0]).trim(), sprint: String(values[1]).trim(), start: date(2), end: date(3),
      plannedHours: numeric(4), actualHours: numeric(5), plannedPercent: numeric(6, true), actualPercent: numeric(7, true),
      variance: numeric(8, true), delay: numeric(9), health: String(values[10]).trim(), blocked: numeric(11),
      comments: commentsColumn < 0 ? "" : String(row[commentsColumn] ?? "") }];
  });
  const summaryWarning = summary ? "" : summaryHeaderIndex < 0
    ? `Summary unavailable: add these Dashboard headers with values beside or below them: ${summaryHeaders.join(", ")}.`
    : "Summary unavailable: enter values beside or below the Dashboard summary headers.";
  return { sheetName, sites, summary, summaryWarning };
}

function renderDashboardWorksheet(data) {
  renderSummary({ ...dashboardData, summary: data.summary || dashboardData.summary });
  renderTable(data.sites);
  renderChart(data.sites);
}
