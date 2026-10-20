const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const XLSX = require('../vendor/xlsx.full.min.js');
class Element {
  constructor() { this.children = []; this.attrs = {}; }
  append(...children) { this.children.push(...children); }
  replaceChildren(...children) { this.children = children; }
  setAttribute(name, value) { this.attrs[name] = value; }
  addEventListener() {}
}
const elements = new Map();
const get = id => { if (!elements.has(id)) elements.set(id, new Element()); return elements.get(id); };
const context = vm.createContext({ XLSX,
  dashboardDateFrom: { value: '2026-10-01' }, dashboardDateTo: { value: '2026-10-08' },
  parseDashboardDate(value) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(value || '')) return null;
    const date = new Date(`${value}T00:00:00Z`);
    return Number.isFinite(date.getTime()) && date.toISOString().slice(0, 10) === value ? date : null;
  },
  document: { getElementById: get, createElement: () => new Element(), addEventListener() {} },
  svgElement(tag, attrs = {}, text) { const node = new Element(); node.attrs = attrs; node.textContent = text; return node; }
});
const source = name => fs.readFileSync(path.join(__dirname, '..', name), 'utf8');
const sprint = source('sprint-burndown.js');
context.dynamicSprintSites = new Set(['alingsas', 'vasteras', 'rakovski', 'heidelberg', 'senatobia', 'chassieu']);
context.sprintSiteKey = value => String(value).trim().toLowerCase().replace(/^mes /, '');
vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'site-filter.js'), 'utf8'), context);
vm.runInContext(sprint.slice(sprint.indexOf('function readSprintDashboardDates'), sprint.indexOf('function updateSprintDashboardDates')), context);
vm.runInContext(sprint.slice(sprint.indexOf('function sprintRowDate'), sprint.indexOf('function readSprintDashboardDates')), context);
vm.runInContext(source('user-story-status.js'), context);
vm.runInContext(source('burndown-chart.js'), context);
const book = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([
  ['Date', 'Area Path', 'Remaining Work', 'Completed Work', 'Sprint Start', 'Sprint End'],
  ['2026-09-30', 'MES Alingsas', 'invalid', 'invalid', '', ''],
  ['01-Oct-2026', 'MES Alingsas', 10, 0, '2026-09-01', '2026-09-15'],
  ['2026-10-01', 'MES Alingsas', 5, 2, '', ''],
  ['08/10/2026', 'MES Alingsas', 4, 11, '', ''],
  ['2026-10-09', 'MES Alingsas', 0, 15, '', ''],
  ['2026-10-15', 'MES Alingsas', 0, 15, '', ''],
  ['', 'MES Alingsas', 999, 999, '', ''],
  ['2026-09-30', 'Other', 999, 999, '', '']
]), 'Master Daily Dump');
XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([['Site', 'Sprint', 'Start', 'End'], ['Alingsas', 20, '2026-10-01', '2026-10-08']]), 'Dashboard');
for (const bookType of ['xlsx', 'biff8']) {
  const parsed = XLSX.read(XLSX.write(book, { type: 'buffer', bookType }), { type: 'buffer' });
  const data = context.readBurndownWorkbook(parsed);
  assert.equal(data.groups.length, 1);
  const group = data.groups[0];
  assert.equal(group.points.length, 2);
  assert.equal(group.points[0].remaining, 15);
  assert.equal(group.points[0].completed, 2);
  assert.equal(group.points[1].label, '2026-10-08');
  assert.equal(group.start.label, '2026-10-01');
  assert.equal(group.end.label, '2026-10-08');
}
context.updateSharedBurndown(book, 'report.xlsx');
assert.equal(get('burndown-site-charts').children.length, 1);
context.dashboardDateFrom.value = 'invalid';
context.dashboardDateTo.value = '2026-09-01';
context.refreshBurndownDateRange();
assert.match(get('burndown-site-charts').children[0].children[1].textContent, /2026-10-01.*2026-10-08/);
assert.equal(context.readBurndownWorkbook(book).groups[0].points[0].completed, 2);
// A second site has its own range and timeline, even with invalid Settings.
XLSX.utils.sheet_add_aoa(book.Sheets.Dashboard, [['Vasteras', 21, '2026-10-09', '2026-10-15']], { origin: 'A3' });
XLSX.utils.sheet_add_aoa(book.Sheets['Master Daily Dump'], [
  ['2026-10-08', 'Project\\MES Vasteras', 999, 999],
  ['2026-10-09', 'Project\\MES Vasteras', 10, 0],
  ['2026-10-15', 'Project\\MES Vasteras', 0, 10],
  ['2026-10-16', 'Project\\MES Vasteras', 999, 999]
], { origin: 'A10' });
let result = context.readBurndownWorkbook(book);
assert.equal(result.groups.length, 2);
assert.equal(result.groups[1].start.label, '2026-10-09');
assert.equal(result.groups[1].end.label, '2026-10-15');
assert.deepEqual(Array.from(result.groups[1].points, p => p.remaining), [10, 0]);
assert.equal(result.groups[1].days[0].label, '2026-10-09');
assert.equal(result.groups[1].days.at(-1).label, '2026-10-15');
assert.equal(result.groups[1].points.at(-1).ideal, 0);
XLSX.utils.sheet_add_aoa(book.Sheets.Dashboard, [['invalid']], { origin: 'C2' });
context.updateSharedBurndown(book, 'report.xlsx');
assert.equal(get('burndown-site-charts').children.length, 1);
assert.match(get('burndown-shared-status').textContent, /Alingsas.*Chart unavailable/);
XLSX.utils.sheet_add_aoa(book.Sheets.Dashboard, [['2026-11-01', '2026-11-15']], { origin: 'C3' });
context.updateSharedBurndown(book, 'report.xlsx');
assert.equal(get('burndown-site-charts').children.length, 0);
assert.match(get('burndown-shared-status').textContent, /No observations/);
assert.throws(() => context.readBurndownWorkbook({ SheetNames: [], Sheets: {} }), /Dashboard.*missing/);
context.resetSharedBurndown();
context.refreshBurndownDateRange();
assert.match(get('burndown-shared-status').textContent, /Upload an Excel/);
console.log('Passed: Burndown Dashboard per-site ranges and timelines, Settings independence, inclusive boundaries, totals, date formats, refresh, empty and invalid ranges, XLSX/XLS.');
