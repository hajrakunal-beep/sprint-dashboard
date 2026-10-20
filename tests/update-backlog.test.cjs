const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const XLSX = require('../vendor/xlsx.full.min.js');
const context = vm.createContext({ XLSX, TextEncoder, TextDecoder });
for (const name of ['preserve-workbook.js', 'update-backlog.js']) {
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', name), 'utf8'), context);
}
function fixture() {
  const book = XLSX.utils.book_new();
  const add = (name, rows) => XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(rows), name);
  add('Dashboard', [
    ['Site', 'Sprint', 'Start', 'End'],
    ['Alpha', 1, '01/10/2026', '10/10/2026'],
    ['Alpha', 2, '2026-10-11', '2026-10-20'],
    ['Beta', 1, '1-Oct-2026', '10-Oct-2026']
  ]);
  add('Consolidated Backlog', [
    ['Report'],
    ['Site', 'ID', 'Sprint No', 'User story / task', 'Estm. Hrs', 'Com. Hrs', 'Rem. Hrs', 'Notes'],
    ['Alpha', 101, 1, 'User Story', 99, 99, 99, 'keep'],
    ['Alpha', 102, 1, 'Bug', 99, 99, 99],
    ['Alpha', 103, 1, 'User Story', 99, 99, 99],
    ['Alpha', 101, 2, 'User Story', 99, 99, 99],
    ['Beta', 101, 1, 'User Story', 99, 99, 99],
    ['Alpha', 101, 1, 'Task', 99, 99, 99],
    ['Alpha', 101, 9, 'Bug', 99, 99, 99]
  ]);
  add('Master Daily Dump', [
    ['ID', 'Area Path', 'Date', 'Work Item Type', 'Parent', 'Original Estimate', 'Remaining Work', 'Completed Work'],
    [1, 'Project\\Alpha', '2026-10-01', 'Task', 101, 1, 1, 0],
    [1, 'Project\\Alpha', '2026-10-10T17:00:00', 'Task', 101, 4, 1, 3],
    [2, 'Project\\ALPHA', '10/10/2026', 'Task', '101', 6, 2, 4],
    [3, 'Project\\Alpha', '2026-10-11', 'Task', 101, 20, 8, 12],
    [4, 'Project\\Alpha', '2026-09-30', 'Task', 102, 999, 999, 999],
    [4, 'Project\\Alpha', '2026-10-01', 'Task', 102, 7, 4, 3],
    [5, 'Project\\Beta', '2026-10-10', 'Task', 101, 50, 20, 30],
    [6, 'Project\\Alpha', '2026-10-10', 'Bug', 102, 999, 999, 999],
    [7, 'Project\\Alpha', 'invalid', 'Task', 102, 999, 999, 999]
  ]);
  add('Other Sheet', [['Formula'], [2]]);
  book.Sheets['Other Sheet'].A2.f = '1+1';
  book.Sheets['Consolidated Backlog'].H3.f = '"keep"';
  book.Sheets['Consolidated Backlog'].E3.f = '90+9';
  book.Sheets['Consolidated Backlog'].E3.z = '0.00';
  return book;
}
function verify(book) {
  const sheet = book.Sheets['Consolidated Backlog'];
  const totals = row => ['E', 'F', 'G'].map(col => sheet[`${col}${row}`].v);
  assert.deepEqual(totals(3), [10, 7, 3]);
  assert.deepEqual(totals(4), [7, 3, 4]); // Bug has its own latest date (start boundary).
  assert.deepEqual(totals(5), [0, 0, 0]);
  assert.deepEqual(totals(6), [20, 12, 8]); // Sprint 2, same parent.
  assert.deepEqual(totals(7), [50, 30, 20]); // Another site, same parent.
  assert.deepEqual(totals(8), [99, 99, 99]); // Task rows untouched.
  assert.deepEqual(totals(9), [0, 0, 0]); // No matching Dashboard sprint.
  assert.equal(sheet.E3.f, undefined);
  assert.equal(sheet.E3.z, '0.00');
  assert.equal(sheet.H3.f, '"keep"');
}
const original = XLSX.write(fixture(), { type: 'buffer', bookType: 'xlsx' });
const result = context.updateBacklogPreservingWorkbook(original);
assert.equal(result.updatedRows, 6);
assert.equal(result.zeroRows, 2);
assert.match(result.noMatchDetails[0], /ID 103: No Task rows have Parent/);
assert.match(result.noMatchDetails[1], /No Dashboard row matches/);
verify(XLSX.read(result.bytes, { type: 'array', cellNF: true }));
const beforeZip = XLSX.CFB.read(original, { type: 'buffer' });
const afterZip = XLSX.CFB.read(result.bytes, { type: 'buffer' });
let changedParts = 0;
for (let index = 0; index < beforeZip.FullPaths.length; index++) {
  const entry = beforeZip.FileIndex[index];
  if (entry.type !== 2) continue;
  const after = XLSX.CFB.find(afterZip, beforeZip.FullPaths[index]);
  if (!Buffer.from(entry.content).equals(Buffer.from(after.content))) {
    changedParts++;
    assert.match(beforeZip.FullPaths[index], /xl\/worksheets\/sheet2.xml$/);
  }
}
assert.equal(changedParts, 1, 'Only Consolidated Backlog worksheet XML changes');

// Header detection works with different column order, case, and offset headers.
const reordered = fixture();
const sheet = reordered.Sheets.Dashboard;
const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
reordered.Sheets.Dashboard = XLSX.utils.aoa_to_sheet(rows.map(row => [...row].reverse()), { origin: 'C4' });
reordered.Sheets.Dashboard.F4.v = '  SITE  ';
assert.equal(context.calculateBacklogUpdates(reordered).updatedRows, 6);

// Excel dates, including fractional timestamps and the 1904 date system.
for (const date1904 of [false, true]) {
  const book = fixture();
  book.Workbook = { WBProps: { date1904 } };
  const base = Math.round((Date.UTC(2026, 9, 1) - Date.UTC(1899, 11, 30)) / 86400000) - (date1904 ? 1462 : 0);
  book.Sheets.Dashboard.C2 = { t: 'n', v: base };
  book.Sheets.Dashboard.D2 = { t: 'n', v: base + 9 };
  book.Sheets['Master Daily Dump'].C4 = { t: 'n', v: base + 9.75 };
  const update = context.calculateBacklogUpdates(book);
  assert.equal(update.updates.find(item => item.address === 'E3').value, 10);
}
// Reproduce the visible Parent 5533 tasks and the dashboard's known site alias.
const screenshotBook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(screenshotBook, XLSX.utils.aoa_to_sheet([
  ['Site', 'Sprint', 'Start', 'End'], ['Alingsas', 20, '01-Oct-2026', '14-Oct-2026']
]), 'Dashboard');
XLSX.utils.book_append_sheet(screenshotBook, XLSX.utils.aoa_to_sheet([
  ['Site', 'ID', 'Sprint No', 'User story / task', 'Estm. Hrs', 'Com. Hrs', 'Rem. Hrs'],
  ['MES Alingsas', 5533, 20, 'User Story', 0, 0, 0]
]), 'Consolidated Backlog');
XLSX.utils.book_append_sheet(screenshotBook, XLSX.utils.aoa_to_sheet([
  ['ID', 'Area Path', 'Date', 'Work Item Type', 'Parent', 'Original Estimate', 'Remaining Work', 'Completed Work'],
  ...[[5590,16,0,16], [5591,16,0,16], [5592,16,0,20], [5593,10,0,10], [5637,16,0,18], [5589,10,0,10], [5594,10,10,0]].map(([id, estimate, remaining, completed]) => [id, 'MES Alingsas', '14-Oct-2026', 'Task', 5533, estimate, remaining, completed])
]), 'Master Daily Dump');
let screenshotResult = context.calculateBacklogUpdates(screenshotBook);
assert.deepEqual(Array.from(screenshotResult.updates, update => update.value), [94, 90, 10]);
assert.equal(screenshotResult.zeroRows, 0);
// Re-importing active Task 5594 previously inflated estimate to 104 and remaining
// to 20. It must count once, while distinct IDs with equal hours still count.
XLSX.utils.sheet_add_aoa(screenshotBook.Sheets['Master Daily Dump'], [[5594, 'MES Alingsas', '14-Oct-2026', 'Task', 5533, 10, 10, 0]], { origin: 'A9' });
screenshotResult = context.calculateBacklogUpdates(screenshotBook);
assert.deepEqual(Array.from(screenshotResult.updates, update => update.value), [94, 90, 10]);
assert.equal(screenshotResult.duplicateTasks, 1);
const duplicateSaved = context.updateBacklogPreservingWorkbook(XLSX.write(screenshotBook, { type: 'buffer', bookType: 'xlsx' }));
const duplicateSheet = XLSX.read(duplicateSaved.bytes, { type: 'array' }).Sheets['Consolidated Backlog'];
assert.deepEqual(['E2', 'F2', 'G2'].map(address => duplicateSheet[address].v), [94, 90, 10]);
screenshotBook.Sheets['Master Daily Dump'].F9.v = 11;
assert.throws(() => context.calculateBacklogUpdates(screenshotBook), /Task 5594.*conflicting hours.*rows 8 and 9/);
screenshotBook.Sheets['Master Daily Dump'].F9.v = 10;
screenshotBook.Sheets.Dashboard.D2.v = '13-Oct-2026';
screenshotResult = context.calculateBacklogUpdates(screenshotBook);
assert.match(screenshotResult.noMatchDetails[0], /2026-10-14.*outside Dashboard range 2026-10-01 to 2026-10-13/);
screenshotBook.Sheets.Dashboard.D2.v = '14-Oct-2026';
screenshotBook.Sheets['Consolidated Backlog'].A2.v = 'Beta';
XLSX.utils.sheet_add_aoa(screenshotBook.Sheets.Dashboard, [['Beta', 20, '01-Oct-2026', '14-Oct-2026']], { origin: 'A3' });
assert.match(context.calculateBacklogUpdates(screenshotBook).noMatchDetails[0], /Area Path does not contain site "Beta"/);
screenshotBook.Sheets['Consolidated Backlog'].A2.v = 'Alingsas';
for (let row = 2; row <= 9; row++) screenshotBook.Sheets['Master Daily Dump'][`C${row}`].v = 'bad date';
assert.match(context.calculateBacklogUpdates(screenshotBook).noMatchDetails[0], /Date values could not be read/);

// All six Master Daily Dump aliases match their short backlog/Dashboard names,
// including Area Paths with subfolders and case/whitespace variations.
for (const [source, site] of [
  ['MES Alingsas', 'Alingsas'], ['MES FR Chassieu', 'Chassieu'],
  ['MES SE Vasteras', 'Vasteras'], ['MES_DE_Heidelberg', 'Heidelberg'],
  ['MES-Rakovski', 'Rakovski'], ['MES-Senatobia', 'Senatobia']
]) {
  for (const area of [source, `Project\\${source}\\Team`, ` ${source.toUpperCase()} `]) {
    const book = fixture();
    book.Sheets.Dashboard.A2.v = site;
    book.Sheets['Consolidated Backlog'].A3.v = site;
    for (const row of [2, 3, 4]) book.Sheets['Master Daily Dump'][`B${row}`].v = area;
    const result = context.calculateBacklogUpdates(book);
    assert.deepEqual(Array.from(result.updates.filter(item => ['E3', 'F3', 'G3'].includes(item.address)), item => item.value), [10, 7, 3], source);
  }
}

const missing = fixture(); delete missing.Sheets.Dashboard.B1;
assert.throws(() => context.calculateBacklogUpdates(missing), /Dashboard requires headers/);
const invalid = fixture(); invalid.Sheets.Dashboard.C2.v = '31/02/2026';
assert.throws(() => context.calculateBacklogUpdates(invalid), /Invalid Dashboard dates/);
const duplicate = fixture(); duplicate.Sheets.Dashboard.E1 = { t: 's', v: 'Site' }; duplicate.Sheets.Dashboard['!ref'] = 'A1:E4';
assert.throws(() => context.calculateBacklogUpdates(duplicate), /duplicate required headers/);
console.log('PASS: latest date per parent, site/sprint matching, inclusive dates, three totals, zero matches, skipped tasks, header mapping, Excel date systems, validation, and workbook preservation.');

// Excluded sites keep existing values and formulas, even with invalid sprint dates.
const excluded = fixture();
XLSX.utils.sheet_add_aoa(excluded.Sheets.Dashboard, [[' FROSINONE ', 1, 'invalid', 'invalid']], { origin: 'A5' });
XLSX.utils.sheet_add_aoa(excluded.Sheets['Consolidated Backlog'], [['Frosinone', 101, 1, 'User Story', 77, 55, 22], [' frosinone ', 102, 1, 'Bug', 44, 33, 11]], { origin: 'A10' });
excluded.Sheets['Consolidated Backlog'].E10.f = '70+7';
const excludedResult = context.updateBacklogPreservingWorkbook(XLSX.write(excluded, { type: 'buffer', bookType: 'xlsx' }));
const excludedSaved = XLSX.read(excludedResult.bytes, { type: 'array' }).Sheets['Consolidated Backlog'];
assert.equal(excludedResult.updatedRows, 6);
assert.equal(excludedResult.zeroRows, 2);
assert.equal(excludedSaved.E10.f, '70+7');
assert.deepEqual(['E10', 'F10', 'G10', 'E11', 'F11', 'G11'].map(address => excludedSaved[address].v), [77, 55, 22, 44, 33, 11]);

async function verifyControls() {
  const elements = new Map();
  const element = id => {
    if (!elements.has(id)) elements.set(id, { disabled: id === 'update-backlog-button', files: [], classList: { toggle() {}, add() {}, remove() {} }, listeners: {}, addEventListener(event, listener) { this.listeners[event] = listener; } });
    return elements.get(id);
  };
  let bytes = original, saved, writes = 0, aborts = 0, permission = 'granted', failWrite = false, failRead = false, cancel = false;
  const handle = {
    async getFile() { return { name: 'master.xlsx', arrayBuffer: async () => { if (failRead) throw new Error('Read failed'); return bytes; } }; },
    async requestPermission() { return permission; },
    async createWritable() { return {
      async write(value) { if (failWrite) throw new Error('Write failed'); saved = value; writes++; },
      async close() { bytes = saved; }, async abort() { aborts++; }
    }; }
  };
  const ui = vm.createContext({ XLSX, TextEncoder, TextDecoder, setTimeout: callback => callback(), window: { async showOpenFilePicker() { if (cancel) { const error = new Error('Cancelled'); error.name = 'AbortError'; throw error; } return [handle]; } }, document: {
    getElementById: element,
    createElement() { throw new Error('Downloads must not be created'); }
  } });
  for (const name of ['preserve-workbook.js', 'update-backlog.js']) vm.runInContext(fs.readFileSync(path.join(__dirname, '..', name), 'utf8'), ui);
  const input = element('backlog-excel-file'), button = element('update-backlog-button'), status = element('backlog-update-status');
  await button.listeners.click();
  assert.match(status.textContent, /Choose an .xlsx backlog workbook/);
  await input.listeners.click();
  assert.equal(button.disabled, false);
  await button.listeners.click();
  assert.equal(writes, 1);
  verify(XLSX.read(saved, { type: 'array', cellNF: true }));
  assert.match(status.textContent, /Updated 6 User Story\/Bug rows; 2 had no matches/);
  assert.match(status.textContent, /Saved directly to master.xlsx/);
  assert.equal(button.disabled, false); assert.equal(input.disabled, false);
  // Read external changes made after selection, and save to the same handle.
  const changed = fixture(); changed.Sheets['Master Daily Dump'].F3.v = 14;
  bytes = XLSX.write(changed, { type: 'buffer', bookType: 'xlsx' });
  await button.listeners.click();
  assert.equal(XLSX.read(saved, { type: 'array' }).Sheets['Consolidated Backlog'].E3.v, 20);
  assert.equal(writes, 2);
  cancel = true; await input.listeners.click(); cancel = false;
  permission = 'denied'; await button.listeners.click();
  assert.match(status.textContent, /Write access was not granted/);
  assert.equal(writes, 2); permission = 'granted';
  failWrite = true; await button.listeners.click(); failWrite = false;
  assert.match(status.textContent, /Write failed/); assert.equal(aborts, 1);
  failRead = true;
  await button.listeners.click();
  assert.match(status.textContent, /Update failed: Read failed/);
  assert.equal(writes, 2); assert.equal(input.disabled, false);
  const unsupported = vm.createContext({ document: { getElementById: element } });
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../update-backlog.js'), 'utf8'), unsupported);
  assert.equal(input.disabled, true); assert.equal(button.disabled, true);
  assert.match(status.textContent, /Direct file saving is unavailable/);
  console.log('PASS: direct saves, latest file reads, no downloads, permission denial, cancellation, write abort, read errors, and unsupported browsers.');
}
verifyControls().catch(error => { console.error(error); process.exitCode = 1; });
