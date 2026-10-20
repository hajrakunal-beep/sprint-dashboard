const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const XLSX = require('../vendor/xlsx.full.min.js');
const element = () => ({ addEventListener() {}, append() {}, replaceChildren() {} });
const context = vm.createContext({ XLSX, document: { getElementById: element, createElement: element, addEventListener() {} }, excelInput: element(), uploadStatus: {},
  dashboardDateFrom: { value: '2026-10-01' }, dashboardDateTo: { value: '2026-10-08' },
  masterDailyDumpSheet: () => 'Master Daily Dump',
  parseDashboardDate(value) {
    const date = new Date(`${value}T00:00:00Z`);
    return Number.isFinite(date.getTime()) && date.toISOString().slice(0, 10) === value ? date : null;
  }
});
const source = name => fs.readFileSync(path.join(__dirname, '..', name), 'utf8');
vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'site-filter.js'), 'utf8'), context);
vm.runInContext(source('site-status.js'), context);
const sprint = source('sprint-burndown.js');
vm.runInContext(sprint.slice(sprint.indexOf('function sprintSiteKey'), sprint.indexOf('let loadedSprintWorkbook')), context);
vm.runInContext(sprint.slice(sprint.indexOf('function sprintRowDate'), sprint.indexOf('function readSprintDashboardDates')), context);
context.dynamicSprintSites = new Set(['alingsas', 'vasteras', 'rakovski', 'heidelberg', 'senatobia', 'chassieu']);
vm.runInContext(sprint.slice(sprint.indexOf('function readSprintDashboardDates'), sprint.indexOf('function updateSprintDashboardDates')), context);
vm.runInContext(source('user-story-status.js'), context);
const headers = ['ID', 'Area Path', 'Date', 'Work Item Type', 'Assigned To', 'State', 'Original Estimate', 'Remaining Work', 'Completed Work'];
const book = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([
  ['Site', 'Start', 'End', 'Sprint'],
  ['MES Alingsas', '01/10/2026', '08/10/2026', 20],
  ['Vasteras', '2026-10-09', '2026-10-15', 21]
]), 'Dashboard');
XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([headers,
  [1, 'MES Alingsas', '30-Sep-2026', 'Task', 'Faiz', 'Active', 10, 9, 1],
  [1, 'MES Alingsas', '01-Oct-2026', 'Task', 'Faiz', 'Active', 10, 8, 2],
  [1, 'MES Alingsas', '08-Oct-2026', 'Task', 'Faiz', 'Closed', 10, 0, 10],
  [2, 'Alingsas', '08-Oct-2026', 'Task', 'Faiz', 'Closed', 5, 0, 5],
  [1, 'MES Alingsas', '09-Oct-2026', 'Task', 'Faiz', 'Active', 10, 9, 1],
  [3, 'Project\\MES SE Vasteras', '09-Oct-2026', 'Task', 'Mangesh', 'Active', 10, 5, 5],
  [3, 'MES SE Vasteras', '16-Oct-2026', 'Task', 'Mangesh', 'Closed', 10, 0, 10],
  [4, 'MES Alingsas', '', 'Task', 'Faiz', 'Active', 10, 10, 0],
  [5, 'MES-Senatobia', '08-Oct-2026', 'Task', 'Suman', 'Active', 10, 10, 0]
]), 'Master Daily Dump');
for (const bookType of ['xlsx', 'biff8']) {
  const parsed = XLSX.read(XLSX.write(book, { type: 'buffer', bookType }), { type: 'buffer' });
  const result = context.readSiteWorksheet(parsed);
  assert.equal(result.items.length, 3);
  assert.equal(result.items[0].completed, 10);
  assert.equal(result.items[0].date, '2026-10-08');
  assert.equal(result.items[1].id, '2');
  assert.equal(result.items[2].id, '3');
  assert.equal(result.items[2].date, '2026-10-09');
  assert.equal(result.warnings.length, 1);
  assert.match(result.warnings[0], /Senatobia.*Dashboard/);
  context.dashboardDateFrom.value = '2026-10-09';
  context.dashboardDateTo.value = '2026-10-15';
  const changed = context.readSiteWorksheet(parsed);
  assert.deepEqual(changed.items, result.items); // Settings dates do not change site ranges.
  context.dashboardDateFrom.value = '2026-10-01';
  context.dashboardDateTo.value = '2026-10-08';
}
context.dashboardDateFrom.value = '2026-10-15';
context.dashboardDateFrom.value = 'invalid';
assert.equal(context.readSiteWorksheet(book).items.length, 3);
// A replacement Dashboard range changes only its own site's snapshot.
XLSX.utils.sheet_add_aoa(book.Sheets.Dashboard, [['2026-10-01', '2026-10-01']], { origin: 'B2' });
let result = context.readSiteWorksheet(book);
assert.equal(result.items.length, 2);
assert.equal(result.items[0].completed, 2);
assert.equal(result.items[1].id, '3');
XLSX.utils.sheet_add_aoa(book.Sheets.Dashboard, [['2026-10-15', '2026-10-09']], { origin: 'B2' });
result = context.readSiteWorksheet(book);
assert.equal(result.items.length, 1);
assert.equal(result.items[0].id, '3');
assert.ok(result.warnings.some(warning => /Alingsas.*Dashboard/.test(warning)));
assert.throws(() => context.readSiteWorksheet({ SheetNames: [], Sheets: {} }), /Dashboard.*missing/);
console.log('Passed: Site wise status sprint ranges, inclusive boundaries, latest in-range snapshots, aliases, blank dates, missing site ranges, XLSX and XLS.');
