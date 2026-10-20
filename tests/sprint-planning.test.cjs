const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const XLSX = require('../vendor/xlsx.full.min.js');

class Element {
  constructor() {
    this.children = []; this.attrs = {}; this.listeners = {}; this.style = {}; this.textContent = '';
    this.classList = { toggle() {}, contains: () => false };
  }
  append(...children) { this.children.push(...children); }
  get textContent() { return this._textContent + this.children.map(child => child.textContent).join(''); }
  set textContent(value) { this._textContent = String(value); this.children = []; }
  replaceChildren(...children) { this.children = children; }
  setAttribute(name, value) { this.attrs[name] = value; }
  addEventListener(name, handler) { this.listeners[name] = handler; }
}
const elements = new Map();
const get = id => { if (!elements.has(id)) elements.set(id, new Element()); return elements.get(id); };
const context = vm.createContext({ XLSX,
  document: { getElementById: get, createElement: () => new Element(), createElementNS: () => new Element(), body: get('body'), addEventListener() {} },
  window: { matchMedia: () => ({ matches: false }), alert() {} }
});
vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'site-filter.js'), 'utf8'), context);
for (const name of ['script.js', 'dashboard-sheet.js', 'site-status.js', 'sprint-burndown.js']) {
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', name), 'utf8'), context);
}
const duration = (index = 0) => get('burndown-cards').children[index].children[0].children[0].children[1].children[1].textContent;
const siteLabel = (index = 0) => get('burndown-cards').children[index].children[0].children[0].children[1].children[0].textContent;
const unavailable = '[Sprint dates unavailable]';
// Deliberately reordered headers, offset cells and a report title above the table.
const headers = ['Unused', ' End ', ' SITE ', 'Other', ' Sprint ', ' Start '];
function workbook(rows, columns = headers) {
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([['Sprint report'], columns, ...rows], { origin: 'C3' }), ' Dashboard ');
  return book;
}
async function upload(book, bookType = 'xlsx') {
  get('site-excel-file').files = [{ name: `master.${bookType}`, arrayBuffer: async () => XLSX.write(book, { type: 'array', bookType }) }];
  get('site-excel-file').listeners.change();
  assert.equal(duration(), unavailable);
  assert.equal(siteLabel(), 'Vasteras');
  await get('upload-excel-button').listeners.click();
}

(async () => {
  assert.equal(duration(), unavailable);
  for (const bookType of ['xlsx', 'xls']) {
    await upload(workbook([
      ['US 1', 46304, ' MES SE Vasteras ', 1, 20, 46293],
      ['Bug 2', '09-10-2026', 'Vestaras', 1, 20, '28-09-2026'],
      ['US 3', '23/10/2026', 'Rakovski', 1, 'Current', '12/10/2026'],
      ['US 4', '2026-10-30', 'Heidelberg', 1, 'current', '2026-10-19'],
      ['US 5', '09-Oct-2026', 'Senatobia', 1, 'current', '28-Sep-2026'],
      ['US 6', '2026-10-09', 'Alingsas', 1, 'Sprint 20', '2026-09-28'],
      ['US 7', '2027-01-08', 'Chassieu', 1, 'current', '2026-12-28'],
      ['US 8', '2026-10-09', 'Other', 1, 'current', '2026-09-28']
    ]), bookType);
    assert.equal(duration(), '[28-09-2026 to 09-10-2026]');
    assert.equal(duration(1), '[12-10-2026 to 23-10-2026]');
    assert.equal(duration(2), '[19-10-2026 to 30-10-2026]');
    assert.equal(duration(3), '[28-09-2026 to 09-10-2026]');
    assert.equal(duration(4), '[28-09-2026 to 09-10-2026]');
    assert.equal(siteLabel(), 'Vasteras\u00a0\u00a0Sprint 20');
    assert.equal(siteLabel(4), 'Alingsas\u00a0\u00a0Sprint 20');
    assert.equal(duration(5), '[28-12-2026 to 08-01-2027]');
    // Dashboard sprint dates load even when the workbook has no Master Daily Dump.
    assert.doesNotMatch(get('upload-status').textContent, /^Could not read file/);
    get('dashboard-date-from').value = '2026-01-01';
    get('dashboard-date-to').value = '2026-01-02';
    context.refreshSprintDateRange();
    assert.equal(duration(1), '[12-10-2026 to 23-10-2026]');
    assert.equal(siteLabel(), 'Vasteras\u00a0\u00a0Sprint 20');
  }
  await upload(workbook([
    ['', '2026-10-23', 'Vasteras', '', 'current', '2026-10-12'],
    ['', '2026-10-23', 'Rakovski', '', 'next', '2026-10-12']
  ]));
  assert.equal(duration(), '[12-10-2026 to 23-10-2026]');
  assert.equal(duration(1), '[12-10-2026 to 23-10-2026]');

  await upload(workbook([
    ['', '2026-10-09', 'Vasteras', '', 'current', '2026-09-28'],
    ['', '2026-10-23', 'Vasteras', '', 'current', '2026-10-12'],
    ['', '2026-03-06', 'Rakovski', '', 'current', '30-02-2026'],
    ['', '2026-10-12', 'Heidelberg', '', 'current', '2026-10-23'],
    ['', '', 'Senatobia', '', 'current', '2026-10-12']
  ]));
  for (let i = 0; i < 6; i++) assert.equal(duration(i), unavailable);
  assert.match(get('sprint-committed-status').textContent, /conflicting ranges/);

  const book1904 = workbook([['', 46304 - 1462, 'Vasteras', '', 'current', 46293 - 1462]]);
  book1904.Workbook = { WBProps: { date1904: true } };
  await upload(book1904);
  assert.equal(duration(), '[28-09-2026 to 09-10-2026]');

  await upload(workbook([], ['Site', 'Sprint', 'Start']));
  assert.equal(duration(), unavailable);
  assert.match(get('sprint-committed-status').textContent, /must contain these headers/);
  const missing = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(missing, XLSX.utils.aoa_to_sheet([['Site', 'Type', 'Sprint Start', 'Sprint End'], ['Vasteras', 'current', '2026-09-28', '2026-10-09']]), 'ELSP Sprint Planning');
  await upload(missing);
  assert.equal(duration(), unavailable);
  assert.match(get('sprint-committed-status').textContent, /Dashboard.*missing/);
  console.log('Passed: Dashboard upload, XLS/XLSX, site matching without a Current filter, repeated rows, date formats/1904, independent date loading, replacement, missing headers/sheets, invalid and conflicting ranges.');
})().catch(error => { console.error(error); process.exitCode = 1; });
