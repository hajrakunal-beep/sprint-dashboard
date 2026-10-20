const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const XLSX = require('../vendor/xlsx.full.min.js');
const context = vm.createContext({ XLSX, document: { getElementById: () => ({ addEventListener() {} }) } });
vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'site-filter.js'), 'utf8'), context);
vm.runInContext(fs.readFileSync(path.join(__dirname, '../burndown-chart.js'), 'utf8'), context);
const workbook = rows => {
  const result = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(result, XLSX.utils.aoa_to_sheet(rows), 'Master Daily Dump');
  return result;
};
const rows = [['Date', 'Remaining Work', 'Site'], ['2026-10-08', 20, 'Vestaras'], ['2026-10-01', 30, 'Vestaras'], ['2026-10-01', 10, 'Vestaras'], ['2026-10-15', 0, 'Vestaras'], ['2026-10-09', '', 'Vestaras'], ['2026-10-01', 10, 'Senotobia'], ['2026-10-15', 2, 'Senotobia']];
const missingMaster = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(missingMaster, XLSX.utils.aoa_to_sheet(rows), 'Other history');
assert.throws(() => context.readBurndownWorkbook(missingMaster), /Master Daily Dump.*missing/);
const isolatedMaster = workbook(rows);
XLSX.utils.book_append_sheet(isolatedMaster, XLSX.utils.aoa_to_sheet([rows[0], ['2026-10-01', 999, 'Unrelated site']]), 'Other history');
assert.equal(context.readBurndownWorkbook(isolatedMaster).groups.length, 2);
const invalidMaster = workbook([['Wrong headers']]);
XLSX.utils.book_append_sheet(invalidMaster, XLSX.utils.aoa_to_sheet(rows), 'Valid history elsewhere');
assert.throws(() => context.readBurndownWorkbook(invalidMaster), /Master Daily Dump.*Date, Site and Remaining Work/);
for (const bookType of ['xlsx', 'biff8']) {
  const parsed = XLSX.read(XLSX.write(workbook(rows), { bookType, type: 'buffer' }), { type: 'buffer' });
  const { groups } = context.readBurndownWorkbook(parsed);
  assert.equal(groups.length, 2);
  assert.equal(groups[0].site, 'Vestaras');
  assert.equal(groups[1].points[0].remaining, 10);
  const { points } = groups[0];
  assert.equal(points.length, 3);
  assert.equal(points[0].remaining, 40);
  assert.equal(points[1].ideal, 20);
  assert.equal(points[2].ideal, 0);
  assert.equal(points[0].label, '2026-10-01');
}
assert.equal(context.burndownDate(1).label, '1900-01-01');
assert.equal(context.burndownDate(1, true).label, '1904-01-02');
assert.throws(() => context.burndownDate('2026-02-30'), /date cells/);
assert.throws(() => context.readBurndownWorkbook(workbook([['ID', 'Remaining Work'], [1, 20]])), /Date, Site and Remaining Work/);
assert.throws(() => context.readBurndownWorkbook(workbook([['Date', 'Remaining Work', 'Site'], ['2026-10-01', -1, 'Site']])), /non-negative/);
assert.equal(context.readBurndownWorkbook(workbook([['Date', 'Remaining Work', 'Site'], ['2026-10-01', 20, 'Site']])).groups[0].points.length, 1);
const sprint = context.readBurndownWorkbook(workbook([
  ['Date', 'Site', 'Remaining Work', 'Sprint Start', 'Sprint End', 'Total Capacity', 'Remaining Capacity'],
  ['2026-10-02', 'Site A', 25, '2026-10-01', '2026-10-09', 40, 32],
  ['2026-10-05', 'Site A', 45, '', '', 40, 24]
])).groups[0];
assert.deepEqual(Array.from(sprint.days, day => day.label), ['2026-10-01', '2026-10-02', '2026-10-05', '2026-10-06', '2026-10-07', '2026-10-08', '2026-10-09']);
assert.equal(sprint.points[1].ideal, 40 * (1 - 2 / 6));
assert.equal(sprint.points[1].capacity, 24);
assert.equal(sprint.initial, 40);
assert.equal(sprint.explicitBounds, true);
assert.throws(() => context.readBurndownWorkbook(workbook([
  ['Date', 'Site', 'Remaining Work', 'Sprint End'], ['2026-10-08', 'Site', 5, '2026-10-07']
])), /within Sprint/);
assert.throws(() => context.readBurndownWorkbook(workbook([
  ['Date', 'Site', 'Remaining Work', 'Remaining Capacity'], ['2026-10-08', 'Site', 5, -1]
])), /Capacity/);
const weekend = context.readBurndownWorkbook(workbook([
  ['Date', 'Site', 'Remaining Work'], ['2026-10-02', 'Site', 10], ['2026-10-03', 'Site', 8], ['2026-10-05', 'Site', 5]
])).groups[0];
assert.equal(weekend.days.length, 3);
class Element {
  constructor(tag) { this.tag = tag; this.children = []; this.attrs = {}; }
  append(...children) { this.children.push(...children); }
  replaceChildren(...children) { this.children = children; }
  setAttribute(key, value) { this.attrs[key] = value; }
}
context.svgElement = (tag, attrs = {}, text) => {
  const node = new Element(tag);
  node.attrs = attrs;
  node.textContent = text;
  return node;
};
const chart = new Element('svg');
context.renderBurndownChart(sprint, chart);
assert.equal(chart.children.filter(node => node.tag === 'polygon').length, 1);
assert.ok(chart.children.some(node => node.attrs.stroke === '#38952c' && node.attrs['stroke-dasharray']));
const ideal = chart.children.find(node => node.tag === 'line' && node.attrs.stroke === '#707070');
assert.equal(ideal.attrs.x2, 615);
assert.equal(ideal.attrs.y2, 315);
assert.ok(!JSON.stringify(chart).includes('NaN'));
const area = chart.children.find(node => node.tag === 'polygon');
assert.ok(!area.attrs.points.endsWith('615,315'), 'Actual area must stop at the latest observation, not the future sprint end.');
const allTypes = [
  ['Date', 'Site', 'Work Item Type', 'Remaining Work', 'Completed Work'],
  ['2026-10-05', 'Site A', 'Task', 5, 30],
  ['2026-10-02', 'Site A', 'Bug', 8, 2],
  ['2026-10-02', 'Site A', 'User Story', 12, 3],
  ['2026-10-05', 'Site A', 'Bug', '', 10],
  ['2026-10-06', 'Site A', 'Task', '', 50],
  ['2026-10-02', 'Site B', 'Task', 100, 200]
];
for (const bookType of ['xlsx', 'biff8']) {
  const parsed = XLSX.read(XLSX.write(workbook(allTypes), { bookType, type: 'buffer' }), { type: 'buffer' });
  const totals = context.readBurndownWorkbook(parsed).groups;
  assert.deepEqual(Array.from(totals[0].points, point => [point.remaining, point.completed]), [[20, 5], [5, 40], [null, 50]]);
  assert.equal(totals[1].points[0].completed, 200);
  const comparison = new Element('svg');
  context.renderBurndownChart(totals[0], comparison);
  const completedLine = comparison.children.find(node => node.tag === 'polyline' && node.attrs.stroke === '#e87500');
  assert.ok(completedLine);
  assert.equal(completedLine.attrs.points.split(' ').length, 3);
  assert.ok(!JSON.stringify(comparison).includes('NaN'));
  assert.ok(comparison.children.some(node => node.tag === 'text' && node.textContent === 50), 'Y-axis includes completed hours above remaining hours.');
}
assert.equal(context.readBurndownWorkbook(workbook(rows)).groups[0].points[0].completed, null);
assert.throws(() => context.readBurndownWorkbook(workbook([
  ['Date', 'Site', 'Remaining Work', 'Completed Work'], ['2026-10-02', 'Site', '', -2]
])), /Completed Work/);
console.log('Passed: XLSX/XLS history, daily aggregation, chronological sorting, ideal line, blank observations, Excel date systems, and invalid history.');
console.log('Passed: working days, sprint bounds, capacity, scope increases, recorded weekends, and SVG area/trend rendering.');
console.log('Passed: remaining/completed totals across all work item types, site isolation, snapshot totals, missing hours, completed-only observations, and comparison rendering.');
