const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const XLSX = require('../vendor/xlsx.full.min.js');
const context = vm.createContext({ XLSX, TextEncoder, TextDecoder });
for (const name of ['merge-data.js', 'preserve-workbook.js']) vm.runInContext(fs.readFileSync(path.join(__dirname, '..', name), 'utf8'), context);
const headers = Array.from(vm.runInContext('dailyDumpHeaders', context));
const book = XLSX.utils.book_new();
const sheet = XLSX.utils.aoa_to_sheet([headers, [1, 'Existing', 46279], [], [2, 'Later', 46280]]);
sheet.C4.z = 'dd-mmm-yyyy';
XLSX.utils.book_append_sheet(book, sheet, 'Master Daily Dump');
XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([['Other sheet'], [42]]), 'Other Sheet');
const zip = XLSX.CFB.read(XLSX.write(book, { type: 'buffer', bookType: 'xlsx' }), { type: 'buffer' });
const text = name => Buffer.from(XLSX.CFB.find(zip, '/' + name).content).toString('utf8');
const put = (name, value) => XLSX.CFB.utils.cfb_add(zip, '/' + name, Buffer.from(value));
const validation = '<dataValidations count="1"><dataValidation type="list" allowBlank="1" sqref="I2:I4"><formula1>"Active,Closed"</formula1></dataValidation></dataValidations>';
const conditional = '<conditionalFormatting sqref="J2:J4"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFFF0000"/><color rgb="FF00FF00"/></colorScale></cfRule></conditionalFormatting>';
const originalTarget = text('xl/worksheets/sheet1.xml');
put('xl/worksheets/sheet1.xml', originalTarget.replace('</sheetData>', '<row r="5" ht="24" customHeight="1"><c r="B5" s="1"/><c r="M5" s="1"/></row><row r="9"><c r="C9" s="1"/></row></sheetData>').replace('</worksheet>', '<autoFilter ref="A1:L4"/>' + conditional + validation + '<tableParts count="1"><tablePart r:id="tableRel"/></tableParts></worksheet>'));
put('xl/worksheets/sheet2.xml', text('xl/worksheets/sheet2.xml').replace('</worksheet>', validation + conditional + '<extLst><ext uri="test"><custom xmlns="urn:test">Preserve me</custom></ext></extLst></worksheet>'));
put('xl/worksheets/_rels/sheet1.xml.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="tableRel" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>');
const table = '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="DailyDump" displayName="DailyDump" ref="A1:L4" totalsRowShown="0"><autoFilter ref="A1:L4"/><tableColumns count="12">' + headers.map((h, i) => `<tableColumn id="${i + 1}" name="${h}"/>`).join('') + '</tableColumns><tableStyleInfo name="TableStyleMedium2" showRowStripes="1"/></table>';
put('xl/tables/table1.xml', table);
put('[Content_Types].xml', text('[Content_Types].xml').replace('</Types>', '<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>'));
put('customXml/item1.xml', '<unknown>Untouched package payload</unknown>');
const originalBytes = XLSX.CFB.write(zip, { type: 'buffer', fileType: 'zip', compression: true });
const before = XLSX.CFB.read(originalBytes, { type: 'buffer' });
const ado = XLSX.utils.book_new();
const record = [123, 'A&B <Site> 😀', 46281, 'Task', 'Title text', 'Another', 1, 'Jane', 'Active', 8, 0, 8];
XLSX.utils.book_append_sheet(ado, XLSX.utils.aoa_to_sheet([[...headers].reverse(), [...record].reverse(), [], [...record].reverse()]), 'Export');
for (const bookType of ['xlsx', 'biff8', 'csv']) {
  const expected = bookType === 'csv' ? ['00123', 'A&B <Site> 😀', '2026-09-16', 'Task', 'Title, "quoted"\nsecond line', 'Another', '00001', 'José', 'Active', 1200, 0, 8] : record;
  const quote = value => '"' + String(value).replaceAll('"', '""') + '"';
  const csv = '\uFEFF' + [[...headers].reverse(), [...expected].reverse(), [], [...expected].reverse()].map(row => row.map(quote).join(',')).join('\r\n');
  const source = bookType === 'csv' ? new TextEncoder().encode(csv) : XLSX.write(ado, { type: 'array', bookType });
  const sourceName = bookType === 'csv' ? 'ADO.CSV' : 'dump.xlsx';
  const result = context.mergeDailyDumpPreservingWorkbook(originalBytes, source, sourceName);
  assert.equal(result.firstRow, 5);
  assert.equal(result.appendedRows, 2);
  const after = XLSX.CFB.read(result.bytes, { type: 'buffer' });
  assert.deepEqual(after.FullPaths, before.FullPaths);
  for (let i = 0; i < before.FullPaths.length; i++) {
    const name = before.FullPaths[i];
    if (name.endsWith('/xl/worksheets/sheet1.xml') || name.endsWith('/xl/tables/table1.xml')) continue;
    assert.deepEqual(Buffer.from(after.FileIndex[i].content || []), Buffer.from(before.FileIndex[i].content || []), `Changed unrelated workbook part: ${name}`);
  }
  const target = Buffer.from(XLSX.CFB.find(after, '/xl/worksheets/sheet1.xml').content).toString();
  assert.match(target, /dataValidation type="list" allowBlank="1" sqref="I2:I6"/);
  assert.match(target, /conditionalFormatting sqref="J2:J6"/);
  assert.match(target, /<row r="5" ht="24" customHeight="1">/);
  assert.match(target, /<c r="B5" s="1" t="inlineStr">/);
  assert.match(target, /<c r="C6" s="1"(?: t="inlineStr")?>/);
  assert.match(target, /<c r="M5" s="1"\/>/);
  assert.match(target, /<row r="9"><c r="C9" s="1"\/><\/row>/);
  assert.match(target, /<autoFilter ref="A1:L6"\/>/);
  const tableAfter = Buffer.from(XLSX.CFB.find(after, '/xl/tables/table1.xml').content).toString();
  assert.equal(tableAfter, table.replaceAll('A1:L4', 'A1:L6'));
  const readback = XLSX.read(result.bytes, { type: 'array', cellNF: true });
  for (const row of [4, 5]) {
    const values = headers.map((_, column) => readback.Sheets['Master Daily Dump'][XLSX.utils.encode_cell({ r: row, c: column })]?.v ?? '');
    assert.deepEqual(values, expected);
  }
  assert.equal(readback.Sheets['Master Daily Dump'].C6.z, 'dd-mmm-yyyy');
  const twice = context.mergeDailyDumpPreservingWorkbook(result.bytes, source, sourceName);
  assert.equal(twice.firstRow, 7);
}
const legacy = XLSX.write(book, { type: 'array', bookType: 'biff8' });
// Optional tags must survive both the column mapping and the ZIP patch writer.
const taggedMaster = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(taggedMaster, XLSX.utils.aoa_to_sheet([
  ['Tag', ...headers], ['Existing tag', 77, 'Existing site']
]), 'Master Daily Dump');
const taggedMasterBytes = XLSX.write(taggedMaster, { type: 'array', bookType: 'xlsx' });
for (const bookType of ['xlsx', 'biff8', 'csv']) {
  for (const tagHeader of [' Tag ', 'TAGS']) {
    const taggedAdo = XLSX.utils.book_new();
    const tags = ['Committed; Team A & B', '00123', ''];
    XLSX.utils.book_append_sheet(taggedAdo, XLSX.utils.aoa_to_sheet([
      [...headers, tagHeader], ...tags.map(tag => [...record, tag])
    ]), 'Export');
    const source = XLSX.write(taggedAdo, { type: 'array', bookType });
    const result = context.mergeDailyDumpPreservingWorkbook(taggedMasterBytes, source, `dump.${bookType === 'biff8' ? 'xls' : bookType}`);
    const saved = XLSX.read(result.bytes, { type: 'array' }).Sheets['Master Daily Dump'];
    assert.equal(saved.A2.v, 'Existing tag');
    assert.equal(result.appendedRows, 3);
    tags.forEach((tag, i) => {
      assert.equal(saved[`A${i + 3}`]?.v ?? '', tag);
      assert.equal(saved[`B${i + 3}`].v, bookType === 'csv' ? '123' : 123);
    });
    assert.throws(() => context.mergeDailyDumpPreservingWorkbook(originalBytes, source, `dump.${bookType === 'biff8' ? 'xls' : bookType}`), /Add a "Tag" column/);
  }
}
const duplicateTags = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(duplicateTags, XLSX.utils.aoa_to_sheet([[...headers, 'Tag', 'Tags'], [...record, 'A', 'B']]), 'Export');
assert.throws(() => context.mergeDailyDumpWorkbooks(taggedMaster, duplicateTags), /duplicate Tag\/Tags columns/);
// A missing optional source column leaves tags blank without rejecting the dump.
const noTagsResult = context.mergeDailyDumpPreservingWorkbook(taggedMasterBytes, XLSX.write(ado, { type: 'array', bookType: 'xlsx' }));
const noTagsSaved = XLSX.read(noTagsResult.bytes, { type: 'array' }).Sheets['Master Daily Dump'];
assert.equal(noTagsSaved.A2.v, 'Existing tag');
assert.equal(noTagsSaved.A3?.v, undefined);
assert.throws(() => context.mergeDailyDumpPreservingWorkbook(legacy, XLSX.write(ado, { type: 'array', bookType: 'xlsx' })), /master must be an .xlsx workbook/);
console.log('Passed: unrelated XLSX parts byte-identical; styles/dropdowns/rules/tables preserved; XLS/XLSX/CSV dumps, UTF-8 BOM, quoted multiline fields, leading-zero IDs, date text, numeric work and repeated merges.');
