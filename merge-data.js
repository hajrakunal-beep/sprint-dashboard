"use strict";

const dailyDumpHeaders = ["ID", "Area Path", "Date", "Work Item Type", "Title 1", "Title 2", "Parent", "Assigned To", "State", "Original Estimate", "Remaining Work", "Completed Work"];
const normalizeDumpHeader = value => String(value ?? "").trim().replace(/\s+/g, " ").toLowerCase();
const dumpCellHasData = cell => Boolean(cell && (cell.f || (cell.v != null && String(cell.v).trim() !== "")));

function findDailyDumpHeaders(sheet) {
  if (!sheet?.["!ref"]) return null;
  const origin = XLSX.utils.decode_range(sheet["!ref"]).s;
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "", blankrows: true });
  for (let row = 0; row < rows.length; row++) {
    const labels = rows[row].map(normalizeDumpHeader);
    const columns = dailyDumpHeaders.map(header => labels.indexOf(normalizeDumpHeader(header)));
    if (columns.some(column => column < 0)) continue;
    if (dailyDumpHeaders.some(header => labels.filter(label => label === normalizeDumpHeader(header)).length !== 1)) {
      throw new Error("The daily dump header row contains duplicate column names.");
    }
    const tagColumns = labels.flatMap((label, column) => label === "tag" || label === "tags" ? [origin.c + column] : []);
    if (tagColumns.length > 1) throw new Error("The daily dump header row contains duplicate Tag/Tags columns.");
    return { row: origin.r + row, columns: columns.map(column => origin.c + column), tagColumn: tagColumns[0] };
  }
  return null;
}

function mergeDailyDumpWorkbooks(master, ado) {
  const masterName = master.SheetNames.find(name => normalizeDumpHeader(name) === "master daily dump");
  if (!masterName) throw new Error('The master workbook must contain a sheet named "Master Daily Dump".');
  const target = master.Sheets[masterName];
  const targetHeaders = findDailyDumpHeaders(target);
  if (!targetHeaders) throw new Error(`Master Daily Dump is missing required headers: ${dailyDumpHeaders.join(", ")}.`);
  const sources = ado.SheetNames.map(name => ({ name, sheet: ado.Sheets[name], headers: findDailyDumpHeaders(ado.Sheets[name]) })).filter(source => source.headers);
  const preferred = sources.find(source => normalizeDumpHeader(source.name) === "master daily dump");
  if (!preferred && sources.length > 1) throw new Error('Multiple ADO sheets contain the required headers. Name the intended sheet "Master Daily Dump".');
  const source = preferred || sources[0];
  if (!source) throw new Error(`The ADO workbook must contain a sheet with these headers: ${dailyDumpHeaders.join(", ")}.`);
  if (source.headers.tagColumn != null && targetHeaders.tagColumn == null) {
    throw new Error('The ADO dump contains tags. Add a "Tag" column to the Master Daily Dump header row, save and close the master, then select it again.');
  }
  const sourceColumns = [...source.headers.columns];
  const targetColumns = [...targetHeaders.columns];
  if (source.headers.tagColumn != null) {
    sourceColumns.push(source.headers.tagColumn);
    targetColumns.push(targetHeaders.tagColumn);
  }

  // Formatting below the data is not an occupied row. Internal gaps must
  // not cause later existing rows to be overwritten.
  let lastRow = targetHeaders.row;
  for (const [address, cell] of Object.entries(target)) {
    if (address.startsWith("!") || !dumpCellHasData(cell)) continue;
    lastRow = Math.max(lastRow, XLSX.utils.decode_cell(address).r);
  }
  const sourceRows = new Set();
  for (const [address, cell] of Object.entries(source.sheet)) {
    if (address.startsWith("!") || !dumpCellHasData(cell)) continue;
    const position = XLSX.utils.decode_cell(address);
    if (position.r > source.headers.row && sourceColumns.includes(position.c)) sourceRows.add(position.r);
  }
  const rows = [...sourceRows].sort((a, b) => a - b);
  if (!rows.length) throw new Error("The ADO worksheet contains no data rows to append.");
  if (lastRow + rows.length >= 1048576) throw new Error("The merged data exceeds the Excel worksheet row limit.");
  const source1904 = Boolean(ado.Workbook?.WBProps?.date1904);
  const target1904 = Boolean(master.Workbook?.WBProps?.date1904);
  rows.forEach((row, index) => {
    sourceColumns.forEach((column, headerIndex) => {
      const cell = source.sheet[XLSX.utils.encode_cell({ r: row, c: column })];
      if (!cell || cell.v == null) return;
      // Append values rather than ADO formulas pointing at unrelated cells.
      const copy = { t: cell.t, v: cell.v };
      if (cell.z) copy.z = cell.z;
      if (headerIndex === 2 && copy.t === "n" && source1904 !== target1904) copy.v += source1904 ? 1462 : -1462;
      target[XLSX.utils.encode_cell({ r: lastRow + 1 + index, c: targetColumns[headerIndex] })] = copy;
    });
  });
  const range = XLSX.utils.decode_range(target["!ref"]);
  range.e.r = Math.max(range.e.r, lastRow + rows.length);
  target["!ref"] = XLSX.utils.encode_range(range);
  if (target["!autofilter"]?.ref) {
    const filter = XLSX.utils.decode_range(target["!autofilter"].ref);
    filter.e.r = Math.max(filter.e.r, lastRow + rows.length);
    target["!autofilter"].ref = XLSX.utils.encode_range(filter);
  }
  return { workbook: master, appendedRows: rows.length, firstRow: lastRow + 2, columns: targetColumns };
}
