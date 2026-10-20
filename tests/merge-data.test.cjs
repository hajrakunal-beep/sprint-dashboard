const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const XLSX = require('../vendor/xlsx.full.min.js');
const context = vm.createContext({ XLSX, TextEncoder, TextDecoder });
vm.runInContext(fs.readFileSync(path.join(__dirname, '../merge-data.js'), 'utf8'), context);
vm.runInContext(fs.readFileSync(path.join(__dirname, '../preserve-workbook.js'), 'utf8'), context);
const headers = Array.from(vm.runInContext('dailyDumpHeaders', context));
const record = [123, 'Alpha', 46279, 'Task', 'Story', 'Task title', 99, 'Jane', 'Active', 8, 0, 8];
function book(name, rows, origin = 'A1') {
  const result = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(result, XLSX.utils.aoa_to_sheet(rows, { origin }), name);
  return result;
}
function master() {
  const result = book('Master Daily Dump', [headers, [11, 'Existing'], [], [22, 'Later']]);
  result.Sheets['Master Daily Dump'].A20 = { t: 'z', s: {} };
  result.Sheets['Master Daily Dump']['!ref'] = 'A1:L20';
  result.Sheets['Master Daily Dump']['!autofilter'] = { ref: 'A1:L4' };
  XLSX.utils.book_append_sheet(result, XLSX.utils.aoa_to_sheet([['Total'], [3]]), 'Summary');
  result.Sheets.Summary.A2 = { t: 'n', v: 3, f: '1+2' };
  return result;
}
const reversed = [...headers].reverse();
for (const bookType of ['xlsx', 'biff8']) {
  const original = master();
  const input = book('ADO Export', [['Report'], reversed.map(h => ` ${h.toUpperCase()} `), [...record].reverse(), [], [...record].reverse()], 'C3');
  input.Sheets['ADO Export'].L5.z = 'dd-mmm-yyyy';
  const source = XLSX.read(XLSX.write(input, { type: 'buffer', bookType }), { type: 'buffer', cellNF: true });
  const oldRows = JSON.stringify(original.Sheets['Master Daily Dump'].A4);
  const result = context.mergeDailyDumpWorkbooks(original, source);
  assert.equal(result.appendedRows, 2);
  assert.equal(result.firstRow, 5);
  assert.equal(JSON.stringify(original.Sheets['Master Daily Dump'].A4), oldRows);
  assert.equal(original.Sheets.Summary.A2.f, '1+2');
  assert.equal(original.Sheets['Master Daily Dump']['!autofilter'].ref, 'A1:L6');
  const saved = XLSX.read(XLSX.write(result.workbook, { type: 'buffer', bookType: 'xlsx' }), { type: 'buffer', cellNF: true });
  for (const row of [4, 5]) {
    const values = headers.map((_, column) => saved.Sheets['Master Daily Dump'][XLSX.utils.encode_cell({ r: row, c: column })]?.v ?? '');
    assert.deepEqual(values, record);
  }
  assert.equal(saved.Sheets['Master Daily Dump'].C5.z, 'dd-mmm-yyyy');
  assert.equal(saved.Sheets.Summary.A2.f, '1+2');
}
const reorderedMaster = book('Master Daily Dump', [reversed], 'B2');
context.mergeDailyDumpWorkbooks(reorderedMaster, book('Dump', [headers, record]));
assert.equal(reorderedMaster.Sheets['Master Daily Dump'].M3.v, 123);
assert.equal(reorderedMaster.Sheets['Master Daily Dump'].B3.v, 8);
const dateSource = book('Dump', [headers, record]);
dateSource.Workbook = { WBProps: { date1904: true } };
const dateMaster = book('Master Daily Dump', [headers]);
context.mergeDailyDumpWorkbooks(dateMaster, dateSource);
assert.equal(dateMaster.Sheets['Master Daily Dump'].C2.v, record[2] + 1462);
assert.throws(() => context.mergeDailyDumpWorkbooks(book('Wrong', [headers]), dateSource), /must contain a sheet/);
assert.throws(() => context.mergeDailyDumpWorkbooks(master(), book('Dump', [headers.slice(1), record.slice(1)])), /must contain a sheet with these headers/);
assert.throws(() => context.mergeDailyDumpWorkbooks(master(), book('Dump', [headers])), /no data rows/);
assert.throws(() => context.mergeDailyDumpWorkbooks(master(), book('Dump', [[...headers, 'ID'], [...record, 123]])), /duplicate column/);
const ambiguous = book('One', [headers, record]);
XLSX.utils.book_append_sheet(ambiguous, XLSX.utils.aoa_to_sheet([headers, record]), 'Two');
assert.throws(() => context.mergeDailyDumpWorkbooks(master(), ambiguous), /Multiple ADO sheets/);

const elements = new Map();
for (const id of ['master-excel-file', 'ado-excel-file', 'merge-data-button', 'merge-data-status', 'master-file-status', 'ado-file-status']) {
  const classes = new Set();
  elements.set(id, { id, files: [], disabled: id === 'merge-data-button', value: '', attrs: {}, listeners: {},
    classList: { toggle(k, on) { on ? classes.add(k) : classes.delete(k); }, add(k) { classes.add(k); }, remove(k) { classes.delete(k); }, contains(k) { return classes.has(k); } },
    setAttribute(k, v) { this.attrs[k] = v; }, addEventListener(k, v) { this.listeners[k] = v; } });
}
let filename = 'master.xlsx';
let diskBytes = XLSX.write(master(), { type: 'array', bookType: 'xlsx' });
let permission = 'granted', failWrite = false, writes = 0, aborted = 0, pickerError;
let staleMasterReads = 0, masterReads = 0, failClose = false;
const staleError = () => Object.assign(new Error('An operation that depends on state cached in an interface object was made but the state had changed since it was read from disk.'), { name: 'NotReadableError' });
const handle = {
  getFile: async () => ({ name: filename, arrayBuffer: async () => {
    masterReads++;
    if (staleMasterReads > 0) { staleMasterReads--; throw staleError(); }
    return diskBytes;
  } }),
  requestPermission: async options => { assert.equal(options.mode, 'readwrite'); return permission; },
  createWritable: async () => {
    let pending;
    return {
      write: async bytes => { if (failWrite) throw new Error('File is locked'); pending = bytes; },
      close: async () => { if (failClose) throw staleError(); diskBytes = pending; writes++; },
      abort: async () => { aborted++; }
    };
  }
};
context.document = { getElementById: id => elements.get(id) };
context.window = { showOpenFilePicker: async options => {
  assert.equal(options.multiple, false);
  assert.equal(options.excludeAcceptAllOption, true);
  if (pickerError) throw pickerError;
  return [handle];
} };
context.XLSX = { ...XLSX, writeFile: () => { throw new Error('Downloads must not be used'); } };
vm.runInContext(fs.readFileSync(path.join(__dirname, '../settings.js'), 'utf8'), context);
const masterInput = elements.get('master-excel-file'), adoInput = elements.get('ado-excel-file'), button = elements.get('merge-data-button');
const file = (name, workbook) => ({ name, arrayBuffer: async () => XLSX.write(workbook, { type: 'array', bookType: 'xlsx' }) });
(async () => {
  await masterInput.listeners.click();
  assert.equal(button.disabled, true);
  adoInput.files = [{ name: 'bad.txt' }]; adoInput.listeners.change();
  assert.equal(button.disabled, true);
  assert.equal(adoInput.attrs['aria-invalid'], 'true');
  adoInput.files = [file('dump.XLSX', book('Export', [headers, record]))]; adoInput.listeners.change();
  assert.equal(button.disabled, false);
  const pending = button.listeners.click();
  assert.equal(button.disabled, true);
  assert.equal(masterInput.disabled, true);
  await pending;
  assert.equal(writes, 1);
  assert.equal(XLSX.read(diskBytes, { type: 'array' }).Sheets['Master Daily Dump'].A5.v, 123);
  assert.match(elements.get('merge-data-status').textContent, /Saved 1 new row starting at row 5.*directly to master.xlsx/);
  assert.equal(button.disabled, false);
  assert.equal(masterInput.disabled, false);
  await button.listeners.click();
  assert.equal(writes, 2);
  assert.equal(XLSX.read(diskBytes, { type: 'array' }).Sheets['Master Daily Dump'].A6.v, 123);
  const beforeError = diskBytes;
  permission = 'denied';
  await button.listeners.click();
  assert.equal(diskBytes, beforeError);
  assert.match(elements.get('merge-data-status').textContent, /Write access was not granted/);
  permission = 'granted'; failWrite = true;
  await button.listeners.click();
  assert.equal(aborted, 1);
  assert.equal(diskBytes, beforeError);
  assert.match(elements.get('merge-data-status').textContent, /File is locked/);
  assert.equal(button.disabled, false);
  failWrite = false;
  adoInput.files = [file('empty.xlsx', book('Empty', [headers]))]; adoInput.listeners.change();
  await button.listeners.click();
  assert.equal(writes, 2);
  assert.match(elements.get('merge-data-status').textContent, /Merge failed:.*no data rows/);
  assert.equal(button.disabled, false);
  filename = 'master.xls';
  diskBytes = XLSX.write(master(), { type: 'array', bookType: 'biff8' });
  const legacyBytes = diskBytes;
  await masterInput.listeners.click();
  assert.equal(diskBytes, legacyBytes);
  assert.match(elements.get('master-file-status').textContent, /Select an .xlsx master to preserve formatting/);
  filename = 'master.xlsx';
  diskBytes = XLSX.write(master(), { type: 'array', bookType: 'xlsx' });
  await masterInput.listeners.click();
  const csvText = [headers.join(','), '00123,CSV Site,2026-09-16,Task,Story,Title,00099,Jane,Active,8,0,8'].join('\r\n');
  adoInput.files = [{ name: 'dump.CSV', arrayBuffer: async () => new TextEncoder().encode(csvText).buffer }]; adoInput.listeners.change();
  assert.equal(button.disabled, false);
  await button.listeners.click();
  const csvSaved = XLSX.read(diskBytes, { type: 'array' }).Sheets['Master Daily Dump'];
  assert.equal(csvSaved.A5.v, '00123');
  assert.equal(csvSaved.C5.v, '2026-09-16');
  assert.equal(csvSaved.J5.v, 8);
  // A dump changed on disk after selection still uses the selected snapshot.
  let adoReads = 0, adoChanged = false;
  const snapshotBytes = XLSX.write(book('Dump', [headers, record]), { type: 'array', bookType: 'xlsx' });
  adoInput.files = [{ name: 'snapshot.xlsx', arrayBuffer: async () => {
    adoReads++;
    if (adoChanged) throw staleError();
    return snapshotBytes;
  } }]; adoInput.listeners.change();
  adoChanged = true;
  staleMasterReads = 1;
  const readsBeforeRetry = masterReads, writesBeforeRetry = writes;
  await button.listeners.click();
  assert.equal(adoReads, 1);
  assert.equal(masterReads, readsBeforeRetry + 2);
  assert.equal(writes, writesBeforeRetry + 1);
  assert.equal(XLSX.read(diskBytes, { type: 'array' }).Sheets['Master Daily Dump'].A6.v, 123);
  // Persistent read failures stop before writing and explain how to recover.
  const beforeStaleFailure = diskBytes;
  staleMasterReads = 2;
  await button.listeners.click();
  assert.equal(diskBytes, beforeStaleFailure);
  assert.equal(writes, writesBeforeRetry + 1);
  assert.match(elements.get('merge-data-status').textContent, /select the master again/);
  // A failed save is aborted, never retried.
  failClose = true;
  const abortsBeforeClose = aborted, readsBeforeClose = masterReads;
  await button.listeners.click();
  assert.equal(masterReads, readsBeforeClose + 1);
  assert.equal(aborted, abortsBeforeClose + 1);
  assert.equal(diskBytes, beforeStaleFailure);
  assert.match(elements.get('merge-data-status').textContent, /select both files again/);
  failClose = false;
  // Selection-time read errors remain handled until Merge is clicked.
  adoInput.files = [{ name: 'unreadable.xlsx', arrayBuffer: async () => { throw staleError(); } }];
  adoInput.listeners.change();
  await button.listeners.click();
  assert.equal(diskBytes, beforeStaleFailure);
  assert.match(elements.get('merge-data-status').textContent, /select the ADO file again/);
  adoInput.files = [file('reselected.xlsx', book('Dump', [headers, record]))]; adoInput.listeners.change();
  await button.listeners.click();
  assert.equal(writes, writesBeforeRetry + 2);
  adoInput.files = [file('dump.xlsx', book('Dump', [headers, record]))]; adoInput.listeners.change();
  pickerError = Object.assign(new Error('Cancelled'), { name: 'AbortError' });
  await masterInput.listeners.click();
  assert.equal(button.disabled, false); // Cancelling retains the previously chosen master.
  adoInput.files = []; adoInput.listeners.change();
  assert.equal(button.disabled, true);
  context.window = {};
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../settings.js'), 'utf8'), context);
  assert.equal(masterInput.disabled, true);
  assert.match(elements.get('merge-data-status').textContent, /browser with local file write access/);
  console.log('Passed: XLS/XLSX ADO merges, header mapping, direct writes, repeated merges, permission denial, write failure/abort, picker cancellation, legacy master rejection and unsupported browsers.');
})().catch(error => { console.error(error); process.exitCode = 1; });
