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
  replaceChildren(...children) { this.children = children; }
  setAttribute(name, value) { this.attrs[name] = value; }
  addEventListener(name, handler) { this.listeners[name] = handler; }
}
const elements = new Map();
const get = id => {
  if (!elements.has(id)) elements.set(id, new Element());
  return elements.get(id);
};
const context = vm.createContext({ XLSX,
  document: { getElementById: get, createElement: () => new Element(), createElementNS: () => new Element(),
    body: get('body'), addEventListener() {} },
  window: { matchMedia: () => ({ matches: false }), CSS: { supports: (_, value) => value === 'blue' }, alert() {} }
});
vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'site-filter.js'), 'utf8'), context);
for (const name of ['script.js', 'dashboard-sheet.js', 'site-status.js', 'burndown-chart.js']) {
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', name), 'utf8'), context);
}
const headers = ['Site', 'Sprint', 'Start', 'End', 'Planned Hrs', 'Actual Hrs', 'Planned %', 'Actual %', 'Variance', 'Delay', 'Health', 'Blocked'];
function workbook() {
  const book = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet([['Report'], headers,
    ['Alpha', 'Sprint 1', 46279, '25-Sep-2026', 100, 50, 0.2, '35%', '-7%', 2, ' green ', 1],
    [], ['Beta', 'Sprint 2', '', '', '1,200', 80, 70, 120, 50, 0, 'Amber', 0],
    ['<script>text</script>', 'Sprint 3', '', '', '', '', '', '', '', '', 'Blue', '']]);
  sheet.G3.z = '0.0%';
  XLSX.utils.book_append_sheet(book, sheet, 'Dashboard');
  return book;
}
(async () => {
  for (const bookType of ['xlsx', 'biff8']) {
    const bytes = XLSX.write(workbook(), { type: 'buffer', bookType });
    const data = context.readDashboardWorksheet(XLSX.read(bytes, { type: 'buffer' }));
    assert.equal(data.sites.length, 3);
    assert.equal(data.sites[0].start, '14-Sep-2026');
    assert.equal(data.sites[0].plannedPercent, 20);
    assert.equal(data.sites[0].actualPercent, 35);
    assert.equal(data.sites[0].variance, -7);
    assert.equal(data.sites[1].plannedHours, 1200);
    context.renderDashboardWorksheet(data);
    const rows = get('sprint-rows').children;
    assert.equal(rows[0].children[6].children[0].style.width, '20%');
    assert.equal(rows[0].children[7].children[0].style.width, '35%');
    assert.equal(rows[0].children[8].textContent, '-7.0%');
    assert.equal(rows[0].children[10].className, 'health health-green');
    assert.equal(rows[1].children[7].children[0].style.width, '100%');
    assert.equal(rows[1].children[7].children[1].textContent, '120.0%');
    assert.equal(rows[2].children[0].textContent, '<script>text</script>');
    assert.equal(rows[2].children[6].children[1].textContent, '—');
    assert.equal(rows[2].children[10].style.backgroundColor, 'blue');
    const comments = rows[0].children[12].children[0];
    assert.equal(comments.value, '');
    assert.equal(comments.rows, 1);
    assert.equal(comments.className, 'dashboard-comments');
    assert.equal(rows[0].children[6].children[0].className, 'cell-bar planned-progress-bar');
    assert.equal(comments.attrs['aria-label'], 'Comments for Alpha, Sprint 1');
    comments.value = 'Waiting for approval\n<script>plain text</script>';
    comments.listeners.input();
    assert.equal(comments.className, 'dashboard-comments has-comment');
    assert.equal(data.sites[0].comments, comments.value);
    context.renderDashboardWorksheet(data);
    assert.equal(get('sprint-rows').children[0].children[12].children[0].value, comments.value);
    assert.equal(get('sprint-rows').children[1].children[12].children[0].value, '');
    const editedComments = get('sprint-rows').children[0].children[12].children[0];
    assert.equal(editedComments.className, 'dashboard-comments has-comment');
    editedComments.value = '';
    editedComments.listeners.input();
    assert.equal(editedComments.className, 'dashboard-comments');
    get('site-excel-file').files = [{ name: 'dashboard.xlsx', arrayBuffer: async () => bytes }];
    get('site-excel-file').listeners.change();
    await get('upload-excel-button').listeners.click();
    assert.match(get('upload-status').textContent, /Loaded 3 sprint rows/);
    assert.equal(get('sprint-rows').children.length, 3);
    const chart = get('progress-chart');
    const bars = chart.children.filter(child => ['planned-bar', 'actual-bar'].includes(child.attrs.class));
    assert.equal(bars.length, 6);
    assert.equal(bars[0].attrs.class, 'planned-bar');
    assert.equal(bars[1].attrs.class, 'actual-bar');
    assert.equal(bars[0].attrs.height, 20 / 100 * 203);
    assert.equal(bars[1].attrs.height, 35 / 100 * 203);
    assert.equal(bars[0].children[0].textContent, 'Alpha — Planned: 20.0%');
    assert.equal(bars[1].children[0].textContent, 'Alpha — Actual: 35.0%');
    assert.ok(chart.children.some(child => child.textContent === 'Alpha'));
    assert.ok(chart.children.some(child => child.textContent === '20.0%'));
    assert.ok(chart.children.some(child => child.textContent === '35.0%'));
    const replacement = workbook();
    replacement.Sheets.Dashboard.A3.v = 'Updated site';
    replacement.Sheets.Dashboard.G3.v = 0.6;
    replacement.Sheets.Dashboard.H3.v = '45%';
    const updatedBytes = XLSX.write(replacement, { type: 'buffer', bookType });
    get('site-excel-file').files = [{ name: 'updated.xlsx', arrayBuffer: async () => updatedBytes }];
    get('site-excel-file').listeners.change();
    await get('upload-excel-button').listeners.click();
    const updatedBars = chart.children.filter(child => ['planned-bar', 'actual-bar'].includes(child.attrs.class));
    assert.equal(updatedBars.length, 6);
    assert.equal(updatedBars[0].attrs.height, 60 / 100 * 203);
    assert.equal(updatedBars[1].attrs.height, 45 / 100 * 203);
    assert.equal(updatedBars[0].children[0].textContent, 'Updated site — Planned: 60.0%');
    assert.equal(updatedBars[1].children[0].textContent, 'Updated site — Actual: 45.0%');
    assert.ok(!chart.children.some(child => child.textContent === 'Alpha'));
  }
  const invalid = workbook();
  const withComments = workbook();
  XLSX.utils.sheet_add_aoa(withComments.Sheets.Dashboard, [['Comments'], ['First line\nSecond line']], { origin: 'M2' });
  const commentedData = context.readDashboardWorksheet(withComments);
  assert.equal(commentedData.sites[0].comments, 'First line\nSecond line');
  assert.equal(commentedData.sites[1].comments, '');
  context.renderDashboardWorksheet(commentedData);
  assert.equal(get('sprint-rows').children[0].children[12].children[0].value, 'First line\nSecond line');
  invalid.Sheets.Dashboard.G3 = { t: 's', v: 'invalid' };
  assert.throws(() => context.readDashboardWorksheet(invalid), /Dashboard row 3: Planned %/);
  assert.throws(() => context.readDashboardWorksheet({ SheetNames: [] }), /Dashboard.*missing/);
  const missingHeader = workbook();
  delete missingHeader.Sheets.Dashboard.L2;
  assert.throws(() => context.readDashboardWorksheet(missingHeader), /must contain these columns/);
  const summaryHeaders = ['TOTAL SITES', 'PLANNED HRS', 'COMPLETED HRS', 'OVERALL ACTUAL %', 'DELAYED / AT RISK'];
  for (const bookType of ['xlsx', 'biff8']) {
    for (const placement of ['above', 'below', 'alongside']) {
      const book = XLSX.utils.book_new();
      const siteRow = ['One site', 'Sprint 1', '', '', 559, 100, 20, 15, -5, 0, 'Green', 0];
      const section = [summaryHeaders, [7, '559', 190, 0.34, 0]];
      const rows = placement === 'above' ? [...section, [], headers, siteRow]
        : placement === 'below' ? [headers, siteRow, [], ...section]
        : [[...headers, 'TOTAL SITES', 'COMPLETED HRS', 'OVERALL ACTUAL %', 'DELAYED / AT RISK'], [...siteRow, 7, 190, 0.34, 0]];
      const sheet = XLSX.utils.aoa_to_sheet(rows);
      const percentageAddress = placement === 'above' ? 'D2' : placement === 'below' ? 'D5' : 'O2';
      sheet[percentageAddress].z = '0.0%';
      XLSX.utils.book_append_sheet(book, sheet, 'Dashboard');
      const bytes = XLSX.write(book, { type: 'buffer', bookType });
      const data = context.readDashboardWorksheet(XLSX.read(bytes, { type: 'buffer' }));
      assert.equal(data.sites.length, 1, 'Summary sections must not create extra site rows');
      assert.equal(data.summary.totalSites, 7, 'Import supplied totals instead of counting site rows');
      assert.equal(data.summary.overallActual, 34);
      get('site-excel-file').files = [{ name: 'summary.xlsx', arrayBuffer: async () => bytes }];
      get('site-excel-file').listeners.change();
      assert.deepEqual(get('summary').children.map(card => card.children[1].textContent), ['—', '—', '—', '—', '—']);
      await get('upload-excel-button').listeners.click();
      assert.deepEqual(get('summary').children.map(card => String(card.children[1].textContent)), ['7', '559', '190', '34.0%', '0']);
      assert.match(get('upload-status').textContent, /Loaded Dashboard summary values/);
      assert.ok(get('summary').children.every(card => card.className === 'summary-card'));
    }
  }
  const invalidSummary = workbook();
  XLSX.utils.sheet_add_aoa(invalidSummary.Sheets.Dashboard, [summaryHeaders, [7, 559, 'invalid', '34%', 0]], { origin: 'A8' });
  assert.throws(() => context.readDashboardWorksheet(invalidSummary), /COMPLETED HRS must contain a number/);
  assert.match(context.readDashboardWorksheet(workbook()).summaryWarning, /Summary unavailable/);
  for (const bookType of ['xlsx', 'biff8']) {
    const book = XLSX.utils.book_new();
    const sheet = XLSX.utils.aoa_to_sheet([
      ['TOTAL SITES', 'PLANNED HRS', 'COMPLETED HRS', 'OVERALL  ACTUAL %', 'DELAYED / AT RISK'],
      [7, 559, 190, 0.34, 0], [], headers,
      ['One site', 'Sprint 1', '', '', 100, 50, 20, 15, -5, 0, 'Green', 0]
    ]);
    sheet.D2.z = '0.0%';
    XLSX.utils.book_append_sheet(book, sheet, 'Dashboard');
    const bytes = XLSX.write(book, { type: 'buffer', bookType });
    get('site-excel-file').files = [{ name: 'combined-summary.xlsx', arrayBuffer: async () => bytes }];
    get('site-excel-file').listeners.change();
    await get('upload-excel-button').listeners.click();
    assert.deepEqual(get('summary').children.map(card => card.children[0].textContent), summaryHeaders);
    assert.deepEqual(get('summary').children.map(card => String(card.children[1].textContent)), ['7', '559', '190', '34.0%', '0']);
    assert.equal(get('sprint-rows').children.length, 1);
    assert.match(get('upload-status').textContent, /Loaded Dashboard summary values/);
  }
  context.renderDashboardWorksheet({ sites: [] });
  assert.equal(get('sprint-rows').children[0].children[0].colSpan, 13);
  assert.deepEqual(get('summary').children.map(card => card.children[1].textContent), ['—', '—', '—', '—', '—']);
  for (const bookType of ['xlsx', 'biff8']) {
    for (const layout of ['beside', 'below', 'pairs']) {
      for (const placement of ['above', 'below']) {
        const book = XLSX.utils.book_new();
        const labels = summaryHeaders;
        const values = [8, '1,200', 480, 0.4, 2];
        const section = layout === 'beside' ? labels.map((label, index) => [label, '', values[index]])
          : layout === 'below' ? labels.flatMap((label, index) => [[label], [values[index]], []])
          : [labels.flatMap((label, index) => [label, values[index]])];
        const siteSection = [headers, ['Site A', 'Sprint 1', '', '', 100, 20, 50, 20, -30, 0, 'Green', 0]];
        const sectionOffset = placement === 'above' ? 0 : 3;
        const rows = placement === 'above' ? [...section, [], ...siteSection] : [...siteSection, [], ...section];
        const sheet = XLSX.utils.aoa_to_sheet(rows);
        const percentageCell = layout === 'beside' ? { r: sectionOffset + 3, c: 2 }
          : layout === 'below' ? { r: sectionOffset + 10, c: 0 } : { r: sectionOffset, c: 7 };
        sheet[XLSX.utils.encode_cell(percentageCell)].z = '0.0%';
        XLSX.utils.book_append_sheet(book, sheet, 'Dashboard');
        const bytes = XLSX.write(book, { type: 'buffer', bookType });
        get('site-excel-file').files = [{ name: 'summary-headers.xlsx', arrayBuffer: async () => bytes }];
        get('site-excel-file').listeners.change();
        await get('upload-excel-button').listeners.click();
        assert.deepEqual(get('summary').children.map(card => String(card.children[1].textContent)), ['8', '1200', '480', '40.0%', '2'], `${bookType}: ${layout}, ${placement}`);
        assert.equal(get('sprint-rows').children.length, 1, 'Header/value sections must not become site rows');
        assert.equal(get('sprint-rows').children[0].children[0].textContent, 'Site A');
      }
    }
  }
  console.log('Passed: XLSX/XLS Dashboard uploads, chart percentages and colors, replacement uploads, all rows, Excel dates, imported variance, blank cells, health colors, validation, and empty overview.');
})();
