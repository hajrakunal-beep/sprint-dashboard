const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const XLSX = require('../vendor/xlsx.full.min.js');

class Element {
  constructor(tag = 'div') {
    this.tag = tag;
    this.children = [];
    this.attrs = {};
    this.listeners = {};
    this.style = {};
    this.textContent = '';
    this.hidden = true;
  }
  append(...children) { this.children.push(...children); }
  replaceChildren(...children) { this.children = children; }
  setAttribute(name, value) { this.attrs[name] = value; }
  addEventListener(name, handler) { this.listeners[name] = handler; }
  querySelector(tag) { return this.children.find(child => child.tag === tag); }
}
const elements = new Map();
const warnings = [];
function get(id) {
  if (!elements.has(id)) elements.set(id, new Element());
  return elements.get(id);
}
const context = vm.createContext({
  XLSX,
  window: { alert: message => warnings.push(message) },
  document: { getElementById: get, createElement: tag => new Element(tag), addEventListener() {} },
  excelInput: get('file'), uploadStatus: get('status'),
  svgElement(tag, attrs, text) { const element = new Element(tag); element.attrs = attrs; element.textContent = text; return element; },
  addCell(row, value) { const cell = new Element('td'); cell.textContent = value; row.append(cell); }
});
vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'site-filter.js'), 'utf8'), context);
vm.runInContext(fs.readFileSync(path.join(__dirname, '../site-status.js'), 'utf8'), context);
vm.runInContext(fs.readFileSync(path.join(__dirname, '../burndown-chart.js'), 'utf8'), context);

const headers = ['ID', 'Work Item Type', 'Assigned To', 'State', 'Original Estimate', 'Remaining Work', 'Completed Work', 'Site'];
const matrix = [headers,
  [5614, 'Task', 'Mangesh', 'New', 24, '', '', 'Vestaras'],
  [5696, 'Task', 'Mangesh', 'Closed', 20, '', 21, 'Vestaras'],
  [5741, 'Task', '', 'New', 6, 6, 0, 'Vestaras'],
  [5742, 'Task', '<script>alert(1)</script>', 'Active', 1.5, 0, 1.5, 'Vestaras']
];
function workbook(rows) {
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([['Cover']]), 'Cover');
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(rows), 'Master Daily Dump');
  return book;
}

(async () => {
  for (const bookType of ['xlsx', 'biff8']) {
    const buffer = XLSX.write(workbook(matrix), { type: 'buffer', bookType });
    const parsed = XLSX.read(buffer, { type: 'buffer' });
    const data = context.readSiteWorksheet(parsed);
    assert.equal(data.sheetName, 'Master Daily Dump');
    assert.equal(data.items.length, 4);
    assert.equal(data.items[0].completed, null);
    assert.equal(data.items[1].completed, 21);
    assert.equal(data.items[2].completed, 0);
    assert.equal(data.items[2].developer, 'Unassigned');
    assert.equal(data.items[1].state, 'Closed');
    context.renderSiteData(data, 'tasks.xlsx');
    context.updateSharedBurndown(parsed, 'tasks.xlsx', data);
    assert.equal(get('site-results-grid').hidden, false);
    assert.equal(get('burndown-site-charts').children.length, 1);
    assert.equal(get('burndown-site-charts').children[0].children[0].textContent, 'Vestaras');
    assert.match(get('burndown-shared-status').textContent, /Dated history is needed/);
    const card = get('site-results-grid').children[0];
    const chart = card.children[2].children[0];
    assert.equal(chart.style.minWidth, chart.style.width);
    assert.equal(chart.style.maxWidth, 'none');
    assert.equal(card.children[2].attrs.tabindex, '0');
    const groups = chart.children.filter(child => child.tag === 'g');
    assert.equal(groups.length, 4);
    const taskPositions = groups.map(group => group.children.find(child => child.attrs.class === 'work-item-id').attrs.x);
    assert.equal(taskPositions[1] - taskPositions[0], 160);
    assert.equal(taskPositions[3] - taskPositions[2], 160);
    assert.ok(groups[0].children.some(child => child.textContent === 'Status: New'));
    assert.ok(groups[0].children.some(child => child.textContent === 'Completion: N/A'));
    assert.ok(groups[1].children.some(child => child.textContent === 'Completion: 105%'));
    assert.ok(groups[2].children.some(child => child.textContent === 'Completion: 0%'));
    assert.ok(groups[3].children.some(child => child.textContent === 'Completion: 100%'));
    const overrun = groups[1].children.filter(child => child.tag === 'rect');
    assert.ok(overrun[1].attrs.height > overrun[0].attrs.height);
    assert.equal(overrun[0].attrs.y + overrun[0].attrs.height, overrun[1].attrs.y + overrun[1].attrs.height);
    assert.equal(overrun[0].attrs.width, 36);
    assert.equal(overrun[1].attrs.width, 36);
    assert.ok(overrun[0].attrs.x + overrun[0].attrs.width < overrun[1].attrs.x);
    assert.ok(chart.children.some(child => child.textContent === 'Task ID'));
    assert.ok(groups.every(group => !group.children.some(child => child.attrs.class === 'work-item-developer')));
    assert.match(card.children[0].children[0].textContent, /Developer\(s\): Mangesh, Unassigned/);
    assert.equal(card.children[0].children[1].textContent, 'Vestaras');
    const zeroLabels = groups[2].children.filter(child => child.textContent === '0 hrs');
    assert.equal(zeroLabels.length, 1);
    assert.match(groups[3].children[0].textContent, /Developer: <script>alert\(1\)<\/script>/);
    assert.ok(groups[0].children.some(child => child.textContent === '5614' && child.attrs.y > 310));
    const file = { name: bookType === 'xlsx' ? 'tasks.xlsx' : 'tasks.xls', arrayBuffer: async () => buffer };
    get('file').files = [file];
    await get('file').listeners.change();
    assert.equal(get('site-results-grid').hidden, true);
    assert.match(get('status').textContent, /Click Upload/);
    assert.equal(get('upload-excel-button').disabled, false);
    await get('upload-excel-button').listeners.click();
    assert.match(get('status').textContent, /Loaded 4 work items/);
    assert.equal(get('burndown-site-charts').children.length, 1);
    assert.equal(get('site-results-grid').hidden, false);
  }
  const tolerantHeaders = headers.map(header => ` ${header.toUpperCase().replaceAll(' ', '  ')} `);
  context.renderWorkItems([{ id: 'zero', workItemType: 'Task', state: 'Active', developer: 'Dev', estimate: 0, completed: 2 }], get('work-item-chart'));
  const zeroEstimateGroup = get('work-item-chart').children[0].children.find(child => child.tag === 'g');
  assert.ok(zeroEstimateGroup.children.some(child => child.textContent === 'Completion: N/A'));
  assert.ok(zeroEstimateGroup.children.some(child => child.textContent === 'User Story: N/A'));
  const mixedData = context.readSiteWorksheet(workbook([
    [...headers, 'Parent'],
    [5000, 'User Story', 'Story owner', 'Active', 1000, 500, 500, 'MES Alingsas', ''],
    [5791, 'Task', 'Task owner', 'Closed', 16, 0, 16, 'MES Alingsas', 5000],
    [5792, ' task ', 'Task owner', 'Active', 8, 4, 4, 'MES Alingsas', '5000'],
    [6000, 'Bug', 'Bug owner', 'Active', 500, 300, 200, 'MES Alingsas', ''],
    [7000, 'User Story', 'Other owner', 'New', 24, 24, 0, 'Story-only site', '']
  ]));
  assert.equal(mixedData.items.length, 5, 'Keep all imported rows available for other views.');
  assert.equal(mixedData.items[1].parent, '5000');
  context.renderSiteData(mixedData);
  const taskCards = get('site-results-grid').children;
  assert.equal(taskCards.length, 1);
  assert.equal(taskCards[0].children[0].children[0].textContent, 'Developer(s): Task owner');
  const taskChart = taskCards[0].children[2].children[0];
  const taskGroups = taskChart.children.filter(node => node.tag === 'g');
  assert.deepEqual(taskGroups.map(group => group.children.find(node => node.attrs.class === 'work-item-id').textContent), ['5791', '5792']);
  assert.ok(taskGroups.every(group => group.children.some(node => node.textContent === 'User Story: 5000')));
  assert.ok(taskGroups[0].children.some(node => node.textContent === 'Completion: 100%'));
  assert.ok(taskChart.children.some(node => node.tag === 'text' && node.textContent === 20), 'User Story and Bug hours must not inflate the Task chart scale.');
  const mixedHistory = context.readBurndownWorkbook(workbook([
    ['Date', 'Site', 'Work Item Type', 'Remaining Work'],
    ['2026-10-01', 'Site', 'User Story', 10],
    ['2026-10-01', 'Site', 'Task', 5]
  ]));
  assert.equal(mixedHistory.groups[0].points[0].remaining, 15, 'Burndown history still includes all work item types.');
  context.renderSiteData({ items: [mixedData.items[0]] });
  assert.match(get('site-results-grid').children[0].textContent, /No Task rows/);
  assert.equal(context.readSiteWorksheet(workbook([['Report title'], tolerantHeaders, matrix[1]])).items.length, 1);
  assert.throws(() => context.readSiteWorksheet(workbook([['Wrong headers']])), /required columns/);
  assert.throws(() => context.readSiteWorksheet(workbook([headers])), /no work items/);
  assert.throws(() => context.readSiteWorksheet(workbook([headers, [1, 'Task', 'Dev', 'New', 'bad', 0, 1]])), /non-negative/);
  assert.throws(() => context.readSiteWorksheet(workbook([headers, ['', 'Task', 'Dev', 'New', 1, 0, 1]])), /ID is missing/);
  get('file').files = [{ name: 'bad.xlsx', arrayBuffer: async () => new Uint8Array([0, 1, 2]).buffer }];
  await get('file').listeners.change();
  await get('upload-excel-button').listeners.click();
  assert.equal(get('site-results-grid').hidden, true);
  assert.match(get('status').textContent, /Could not read file/);
  get('file').files = [];
  await get('file').listeners.change();
  assert.equal(get('status').textContent, 'No file selected.');
  assert.equal(get('upload-excel-button').disabled, true);
  // A changed selection must invalidate a previous upload still being read.
  let finishRead;
  const pendingBytes = new Promise(resolve => { finishRead = resolve; });
  get('file').files = [{ name: 'old.xlsx', arrayBuffer: () => pendingBytes }];
  get('file').listeners.change();
  const pendingUpload = get('upload-excel-button').listeners.click();
  get('file').files = [{ name: 'new.xlsx' }];
  get('file').listeners.change();
  finishRead(XLSX.write(workbook(matrix), { type: 'buffer', bookType: 'xlsx' }));
  await pendingUpload;
  assert.equal(get('site-results-grid').hidden, true);
  assert.match(get('status').textContent, /Selected new.xlsx/);
  assert.equal(get('upload-excel-button').disabled, false);
  const groupedRows = [headers, matrix[1], [6614, 'Task', 'Suman', 'New', 20, '', '', 'Senotobia'], matrix[2], [6616, 'Task', 'Suman', 'Active', 4, 0, 4, ' Senotobia ']];
  const groupedData = context.readSiteWorksheet(workbook(groupedRows));
  context.renderSiteData(groupedData);
  const cards = get('site-results-grid').children;
  assert.equal(cards.length, 2);
  assert.equal(cards[0].children[0].children[1].textContent, 'Vestaras');
  assert.equal(cards[1].children[0].children[1].textContent, 'Senotobia');
  assert.equal(cards[1].children[0].children[0].textContent, 'Developer(s): Suman');
  const ids = card => card.children[2].children[0].children.filter(child => child.tag === 'g').map(group => group.children.find(child => child.attrs.class === 'work-item-id').textContent);
  assert.deepEqual(ids(cards[0]), ['5614', '5696']);
  assert.deepEqual(ids(cards[1]), ['6614', '6616']);
  const sharedBook = workbook([
    [...headers, 'Date'],
    [5614, 'Task', 'Mangesh', 'New', 40, 40, 0, 'Vestaras', '2026-10-01'],
    [5614, 'Task', 'Mangesh', 'Active', 40, 20, 20, 'Vestaras', '2026-10-08'],
    [6614, 'Task', 'Suman', 'New', 20, 20, 0, 'Senotobia', '2026-10-01'],
    [6614, 'Task', 'Suman', 'Active', 20, 4, 16, 'Senotobia', '2026-10-08']
  ]);
  context.updateSharedBurndown(sharedBook, 'shared.xlsx', groupedData);
  const chartCards = get('burndown-site-charts').children;
  assert.equal(chartCards.length, 2);
  assert.equal(chartCards[0].children[0].textContent, 'Vestaras');
  assert.equal(chartCards[1].children[0].textContent, 'Senotobia');
  const burndownSvg = chartCards[0].children.find(child => child.tag === 'svg');
  assert.ok(burndownSvg.children.some(child => child.tag === 'polygon' && child.attrs.fill === '#0078d4'));
  assert.ok(burndownSvg.children.some(child => child.tag === 'line' && child.attrs.stroke === '#707070'));
  context.resetSharedBurndown();
  assert.equal(get('burndown-site-charts').children.length, 0);
  const datedBook = workbook([
    [...headers, 'Date'],
    [...matrix[1], '2026-10-08'],
    [...matrix[2], '2026-10-01'],
    [...matrix[3], '2026-10-08'],
    [6614, 'Task', 'Suman', 'New', 20, 20, 0, 'Senotobia', '2026-10-01'],
    [6614, 'Task', 'Suman', 'Active', 20, 15, 5, 'Senotobia', '2026-10-05'],
    [...matrix[4], '']
  ]);
  const latest = context.readSiteWorksheet(datedBook);
  assert.deepEqual(Array.from(latest.items, item => item.id), ['5614', '5741', '6614']);
  assert.equal(latest.items[2].date, '2026-10-05');
  assert.equal(latest.items[2].completed, 5);
  const fullHistory = context.readBurndownWorkbook(datedBook);
  assert.equal(fullHistory.groups.find(group => group.site === 'Senotobia').points.length, 2);
  const decoy = XLSX.utils.aoa_to_sheet([headers, [999, 'Task', 'Other', 'New', 999, 999, 0, 'Other site']]);
  XLSX.utils.book_append_sheet(datedBook, decoy, 'Other tasks');
  assert.equal(context.readSiteWorksheet(datedBook).items.length, 3);
  const invalidMaster = workbook([['Wrong headers']]);
  XLSX.utils.book_append_sheet(invalidMaster, decoy, 'Valid tasks elsewhere');
  assert.throws(() => context.readSiteWorksheet(invalidMaster), /Master Daily Dump.*required columns/);
  const warningsBeforeMissingSheet = warnings.length;
  for (const bookType of ['xlsx', 'biff8']) {
    const missingMaster = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(missingMaster, decoy, 'Other tasks');
    assert.throws(() => context.readSiteWorksheet(missingMaster), /Master Daily Dump.*missing/);
    const buffer = XLSX.write(missingMaster, { type: 'buffer', bookType });
    get('file').files = [{ name: bookType === 'xlsx' ? 'missing.xlsx' : 'missing.xls', arrayBuffer: async () => buffer }];
    get('file').listeners.change();
    await get('upload-excel-button').listeners.click();
    assert.match(warnings[warnings.length - 1], /Master Daily Dump.*missing/);
    assert.equal(get('site-results-grid').hidden, true);
    assert.equal(get('burndown-site-charts').children.length, 0);
    assert.equal(get('upload-excel-button').disabled, false);
  }
  assert.equal(warnings.length - warningsBeforeMissingSheet, 2);
  context.renderSiteData(context.readSiteWorksheet(workbook(matrix)));
  assert.equal(get('site-results-grid').children.length, 1);
  const html = fs.readFileSync(path.join(__dirname, '../index.html'), 'utf8');
  assert.ok(!html.includes('site-data-table'));
  assert.ok(html.includes('id="upload-excel-button"'));
  console.log('Passed: explicit upload, results visibility, XLSX/XLS imports, equal-width side-by-side bars, Task ID labels, developer header, blank/zero/overrun hours, invalid uploads, and changed selection during upload.');
})().catch(error => { console.error(error); process.exitCode = 1; });
