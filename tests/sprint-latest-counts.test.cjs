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
const normalize = value => String(value ?? '').trim().replace(/\s+/g, ' ').toLowerCase();
const aliases = { 'mes se vasteras': 'Vasteras', 'mes-rakovski': 'Rakovski', 'mes_de_heidelberg': 'Heidelberg', 'mes-senatobia': 'Senatobia', 'mes alingsas': 'Alingsas', 'mes fr chassieu': 'Chassieu' };
const context = vm.createContext({ XLSX,
  document: { getElementById: get, createElement: () => new Element() },
  dashboardDateFrom: { value: '2026-10-01' }, dashboardDateTo: { value: '2026-10-15' },
  normalizeColumn: normalize,
  siteDisplayName: value => aliases[normalize(value)] || value,
  assigneeDisplayName: value => value,
  parseDashboardDate(value) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(value || '')) return null;
    const date = new Date(`${value}T00:00:00Z`);
    return Number.isFinite(date.getTime()) && date.toISOString().slice(0, 10) === value ? date : null;
  }
});
vm.runInContext(fs.readFileSync(path.join(__dirname, '../sprint-burndown.js'), 'utf8'), context);
const sites = ['Vasteras', 'Rakovski', 'Heidelberg', 'Senatobia', 'Alingsas', 'Chassieu'];
const areaNames = Object.keys(aliases);
const headers = ['ID', 'Area Path', 'Date', 'Work Item Type', 'Assigned To', 'State', 'Tags'];
function workbook() {
  const rows = [];
  sites.forEach((site, index) => {
    const latest = index === 0 ? '2026-10-15' : `2026-10-${String(14 - index).padStart(2, '0')}`;
    rows.push(
      [1, site, '2026-10-01', 'User Story', 'Dev', 'Closed', 'Committed; Tentative'],
      [2, site, '2026-09-30', 'Bug', 'Dev', 'Resolved', 'Committed'],
      [3, site, '2026-10-16', 'Bug', 'Dev', 'Done', 'Tentative'],
      [4, site, '', 'Bug', 'Dev', 'Active', 'Committed'],
      [5, site, 'invalid', 'Bug', 'Dev', 'Active', 'Tentative'],
      [1, areaNames[index], latest, 'User Story', 'Dev', ' CLOSED ', 'Priority; Committed; committed'],
      [6, site, latest, 'Bug', 'Dev', 'Resolved', ' tentative '],
      [7, site, latest, 'Bug', 'Dev', 'Active', 'Committed;Tentative'],
      [8, site, latest, 'Task', 'Dev', 'Closed', 'Committed;Tentative'],
      [9, site, latest, 'User Story', 'Dev', 'Done', 'Not Committed;Not Tentative'],
      [10, site, latest, 'Bug', 'Dev', 'Completed', '']
    );
  });
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([['Report'], headers, ...rows], { origin: 'C3' }), 'Master Daily Dump');
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([
    ['Site', 'Sprint', 'Start', 'End'],
    ...sites.map((site, index) => [areaNames[index], 20, '2026-10-01', '2026-10-15'])
  ]), 'Dashboard');
  return book;
}
for (const bookType of ['xlsx', 'biff8']) {
  const book = XLSX.read(XLSX.write(workbook(), { type: 'buffer', bookType }), { type: 'buffer' });
  const result = context.readSprintTagCounts(book);
  sites.forEach((site, index) => {
    const counts = result.resultsBySite.get(site.toLowerCase());
    assert.equal(counts.committed, 2, `${site}: latest snapshot committed only`);
    assert.equal(counts.tentative, 2, `${site}: latest snapshot tentative only`);
    assert.equal(counts.delivered, 3, `${site}: latest snapshot Closed/Resolved/Done only, regardless of tags`);
    assert.deepEqual([...counts.tentativeIds], ['Bug 6', 'Bug 7']);
    assert.equal(counts.skippedDates, 2);
    assert.equal(new Date(result.latestDatesBySite.get(site.toLowerCase())).toISOString().slice(0, 10), index === 0 ? '2026-10-15' : `2026-10-${String(14 - index).padStart(2, '0')}`);
  });
  context.updateSprintTagCounts(book);
  const metric = index => get('burndown-cards').children[index].children[1].children[0].children[1].textContent;
  assert.equal(metric(0), '2');
  const delivered = index => get('burndown-cards').children[index].children[1].children[2].children[1].textContent;
  sites.forEach((_, index) => {
    assert.equal(delivered(index), '3');
    assert.equal(get('burndown-cards').children[index].children[0].children[1].textContent, '150%');
  });
  assert.match(get('sprint-committed-status').textContent, /Vasteras on 2026-10-15/);
  context.dashboardDateFrom.value = 'invalid';
  context.dashboardDateTo.value = '2026-09-01';
  context.refreshSprintDateRange();
  sites.forEach((_, index) => assert.equal(metric(index), '2')); // Settings cannot change card counts.
  sites.forEach((_, index) => assert.equal(delivered(index), '3'));

  // Give each card its own range; both endpoints are inclusive.
  sites.forEach((_, index) => {
    const start = index % 2 === 0 ? '2026-10-01' : '2026-10-02';
    XLSX.utils.sheet_add_aoa(book.Sheets.Dashboard, [[start, start]], { origin: `C${index + 2}` });
  });
  context.updateSprintTagCounts(book);
  sites.forEach((_, index) => {
    assert.equal(metric(index), index % 2 === 0 ? '1' : '0');
    assert.equal(delivered(index), index % 2 === 0 ? '1' : '0');
  });
  const filtered = context.readSprintTagCounts(book);
  assert.deepEqual([...filtered.resultsBySite.get('vasteras').resolved], ['US 1']);
  assert.equal(filtered.resultsBySite.get('rakovski').inProgress.size, 0);
  assert.equal(filtered.developersBySite.has('rakovski'), false);

  // Invalid dates affect only their own card and clear earlier counts.
  XLSX.utils.sheet_add_aoa(book.Sheets.Dashboard, [['invalid']], { origin: 'C2' });
  context.updateSprintTagCounts(book);
  assert.equal(metric(0), '\u2014');
  assert.equal(metric(2), '1');
  assert.match(get('sprint-committed-status').textContent, /Vasteras: counts unavailable/);

}
// A later task-only snapshot must not reuse old story counts.
const taskOnly = workbook();
XLSX.utils.sheet_add_aoa(taskOnly.Sheets['Master Daily Dump'], [[10, 'Chassieu', '2026-10-15', 'Task', 'Dev', 'Active', 'Committed']], { origin: { r: XLSX.utils.decode_range(taskOnly.Sheets['Master Daily Dump']['!ref']).e.r + 1, c: 2 } });
const counts = context.readSprintTagCounts(taskOnly).resultsBySite.get('chassieu');
assert.equal(counts.committed, 0);
assert.equal(counts.tentative, 0);
assert.equal(counts.delivered, 0);
assert.equal(counts.tentativeIds.size, 0);
console.log('Passed: Dashboard per-site ranges, Settings independence, invalid site isolation, latest per-site sprint counts, all six sites, aliases, XLS/XLSX, offset headers, inclusive boundaries, refresh, empty ranges and task-only snapshots.');
