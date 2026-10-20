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
const committed = () => get('burndown-cards').children[0].children[1].children[0].children[1].textContent;
const tentative = () => get('burndown-cards').children[0].children[1].children[1].children[1].textContent;
const delivered = () => get('burndown-cards').children[0].children[1].children[2].children[1].textContent;
const developer = (index = 0) => get('burndown-cards').children[index].children[0].children[0].children[0].textContent;
const completion = (index = 0) => get('burndown-cards').children[index].children[0].children[1].textContent;
const inProgress = () => get('burndown-cards').children[0].children[3].children[0].children[1].textContent;
const resolved = () => get('burndown-cards').children[0].children[3].children[1].children[1].textContent;
const headers = ['Tag', 'State', 'Assigned To', 'Work Item Type', 'Date', 'Site', 'ID'];
function workbook(rows, columns = headers) {
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([['Report'], columns, ...rows], { origin: 'C3' }), 'Master Daily Dump');
  return book;
}
const rows = [
  ['Committed', 'Active', 'Jane', 'User Story', '2026-09-15', 'Vestaras', 1],
  [' committed ', 'Closed', 'Bob', ' BUG ', '2026-09-16', ' vEsTaRaS ', 2],
  ['Committed', 'New', 'Jane', 'User Story', '', 'Vasteras', 3],
  ['Committed', '', '', 'Bug', '', 'MES SE Vasteras', 4],
  ['Tentative', '', '', 'Bug', '', 'Vestaras', 5],
  ['Committed', '', '', 'Task', '', 'Vestaras', 6],
  ['Committed', '', '', 'User Story', '', 'Other', 7],
  ['Not Committed', '', '', 'Bug', '', 'Vestaras', 8],
  [],
  ['Committed', 'Active', 'Jane', 'User Story', '2026-09-17', 'Vestaras', 1],
  [' tentative ', 'Closed', 'Jane', ' user   story ', '', 'Vasteras', 9],
  ['TENTATIVE', 'New', 'Bob', 'Bug', '', 'MES SE Vasteras', 10],
  ['Tentative', '', '', 'Task', '', 'Vestaras', 11],
  ['Tentative', '', '', 'Bug', '', 'Other', 12],
  ['Not Tentative', '', '', 'Bug', '', 'Vestaras', 13],
  ['', 'Closed', 'Jane', 'User Story', '', 'Vestaras', 14],
  ['Other', ' resolved ', 'Bob', ' USER STORY ', '', 'Vasteras', 15],
  ['', 'DONE', '', 'User Story', '', 'MES SE Vasteras', 16],
  ['', 'Resolved', '', 'Bug', '', 'Vestaras', 17],
  ['', 'Done', '', 'Task', '', 'Vestaras', 18],
  ['', 'Closed', '', 'User Story', '', 'Other', 19],
  ['', 'Not Done', '', 'User Story', '', 'Vestaras', 20]
];
async function upload(book, bookType = 'xlsx', fillDates = true) {
  // Existing filter-rule fixtures use a valid date unless a date-filter test
  // explicitly supplies its own blank/invalid dates.
  const sheet = book.Sheets['Master Daily Dump'];
  if (sheet && fillDates) {
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: true });
    const header = rows.findIndex(row => row.includes('Date'));
    if (header >= 0) {
      const origin = XLSX.utils.decode_range(sheet['!ref']).s;
      const column = rows[header].indexOf('Date');
      rows.slice(header + 1).forEach((row, index) => {
        if (row.some(value => String(value).trim() !== '') && !row[column]) sheet[XLSX.utils.encode_cell({ r: origin.r + header + 1 + index, c: origin.c + column })] = { t: 's', v: '2026-09-16' };
      });
    }
  }
  get('site-excel-file').files = [{ name: bookType === 'xlsx' ? 'master.xlsx' : 'master.xls', arrayBuffer: async () => XLSX.write(book, { type: 'array', bookType }) }];
  get('site-excel-file').listeners.change();
  assert.equal(committed(), '—');
  assert.equal(tentative(), '—');
  assert.equal(delivered(), '—');
  assert.equal(completion(), 'N/A');
  assert.equal(inProgress(), 'NA');
  assert.equal(resolved(), 'NA');
  await get('upload-excel-button').listeners.click();
  if (committed() === '—') {
    assert.equal(tentative(), '—');
    assert.equal(delivered(), '—');
  }
}
(async () => {
  // These workbook fixtures cover September; do not depend on today's defaults.
  get('dashboard-date-from').value = '2026-09-14';
  get('dashboard-date-to').value = '2026-09-25';
  get('dashboard-date-from').listeners.change();
  assert.equal(committed(), '—');
  for (const bookType of ['xlsx', 'biff8']) {
    await upload(workbook(rows), bookType);
    assert.equal(committed(), '5'); // Count rows, including repeated IDs across days.
    assert.equal(tentative(), '3');
    assert.equal(delivered(), '6'); // Includes Closed/Resolved Bugs and stories, regardless of tag.
    assert.equal(completion(), '120%');
    assert.equal(completion(1), 'N/A');
    assert.equal(completion(5), 'N/A');
    assert.match(get('sprint-committed-status').textContent, /6 delivered User Story \/ Bug rows/);
    assert.match(get('sprint-committed-status').textContent, /3 tentative/);
    assert.match(get('sprint-committed-status').textContent, /Loaded 5/);
    assert.doesNotMatch(get('upload-status').textContent, /Could not read file/);
    const otherCard = get('burndown-cards').children[1];
    assert.equal(otherCard.children[1].children[0].children[1].textContent, '0');
  }
  await upload(workbook([rows[6]]));
  assert.equal(committed(), '0');
  assert.equal(tentative(), '0');
  assert.equal(delivered(), '0');
  assert.equal(completion(), 'N/A');
  await upload(workbook([rows[15]]));
  assert.equal(committed(), '0');
  assert.equal(tentative(), '0');
  assert.equal(delivered(), '1');
  assert.equal(completion(), 'N/A');
  await upload(workbook([rows[1], rows[18], ['', ' Done ', '', ' bug ', '', 'Vestaras', 21]]));
  assert.equal(delivered(), '3');
  await upload(workbook([rows[4]]));
  assert.equal(committed(), '0');
  assert.equal(tentative(), '1');
  assert.equal(delivered(), '0');
  await upload(workbook(rows.map(row => row.slice(1)), headers.slice(1)));
  assert.equal(committed(), '—');
  assert.match(get('sprint-committed-status').textContent, /must contain these headers.*Tag/);
  const missing = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(missing, XLSX.utils.aoa_to_sheet([['No data']]), 'Other');
  await upload(missing);
  assert.equal(committed(), '—');
  assert.match(get('sprint-committed-status').textContent, /worksheet "Master Daily Dump" is missing/);
  await upload(workbook([rows[0]]));
  assert.equal(committed(), '1');
  assert.equal(tentative(), '0');
  assert.equal(delivered(), '0');
  get('site-excel-file').files = [{ name: 'broken.xlsx', arrayBuffer: async () => { throw new Error('Broken workbook'); } }];
  get('site-excel-file').listeners.change();
  await get('upload-excel-button').listeners.click();
  assert.equal(committed(), '—');
  assert.equal(tentative(), '—');
  assert.equal(delivered(), '—');
  assert.match(get('upload-status').textContent, /Broken workbook/);
  const assignments = [
    ['Diogo Ribeiro EXT <diogo.ribeiro@it.abb.com>', 'MES_DE_Heidelberg', 'Diogo'],
    ['Faiz Mohammed Siddiqui EXT <faiz-mohammed.siddiqui@it.abb.com>', 'MES Alingsas', 'Faiz'],
    ['Joao Correia Rodrigues EXT <joao.correia-rodrigues@it.abb.com>', 'MES-Rakovski', 'Joao'],
    ['Mangesh Phatale EXT <mangesh.phatale@pt.abb.com>', 'Vestaras', 'Mangesh'],
    ['Nivruth Gummadavelly EXT <nivruth.gummadavelly@pt.abb.com>', 'MES FR Chassieu', 'Nivruth'],
    ['Ricardo Torres EXT <ricardo.torres@it.abb.com>', 'MES-Rakovski', 'Ricardo'],
    ['Shaalini Subu-Venkataraman EXT <shaalini.subu-venkataraman@it.abb.com>', 'MES FR Chassieu', 'Shaalini'],
    ['Suman Mannepalli EXT <suman.mannepalli@pt.abb.com>', 'MES-Senatobia', 'Suman'],
    ['Vinod Kumar Kuchi EXT <vinod-kumar.kuchi@de.abb.com>', 'MES_DE_Heidelberg', 'Vinod']
  ];
  const assignedRows = assignments.map(([name, site], i) => ['Committed', 'Active', name, 'User Story', '', site, i + 100]);
  assignedRows.push(['Tentative', 'New', ' MANGESH PHATALE EXT <MANGESH.PHATALE@PT.ABB.COM> ', 'Bug', '', 'MES SE Vasteras', 200]);
  assignedRows.push(['', '', '', 'Task', '', 'Vasteras', 201]);
  for (const bookType of ['xlsx', 'biff8']) {
    await upload(workbook(assignedRows), bookType);
    assert.equal(developer(), 'Mangesh');
    assert.equal(developer(1), 'Joao, Ricardo');
    assert.equal(developer(2), 'Diogo, Vinod');
    assert.equal(developer(3), 'Suman');
    assert.equal(developer(4), 'Faiz');
    assert.equal(developer(5), 'Nivruth, Shaalini');
  }
  await upload(workbook([['', 'Active', 'New Developer <new@example.com>', 'Task', '', 'Vestaras', 300]]));
  assert.equal(developer(), 'New Developer <new@example.com>');
  assert.equal(developer(1), 'Unassigned');
  const ratioRows = [
    ['Committed', 'Closed', 'Jane', 'User Story', '', 'Vestaras', 400],
    ['Committed', 'Active', 'Jane', 'User Story', '', 'Vestaras', 401],
    ['Committed', 'Active', 'Jane', 'Bug', '', 'Vestaras', 402]
  ];
  await upload(workbook(ratioRows));
  assert.equal(completion(), '33%');
  ratioRows[1][1] = 'Done';
  await upload(workbook(ratioRows));
  assert.equal(completion(), '67%');
  ratioRows[2][1] = 'Resolved';
  await upload(workbook(ratioRows));
  assert.equal(completion(), '100%');
  await upload(workbook([rows[0]]));
  assert.equal(completion(), '0%');
  const activeRows = [
    ['Committed', 'In Progress', 'Jane', 'User Story', '', 'Vestaras', 5588],
    ['Tentative', ' in   progress ', 'Jane', ' bug ', '', 'MES SE Vasteras', 5413],
    ['', 'In Progress', '', 'User Story', '', 'Vasteras', '005736'],
    ['', 'In Progress', '', 'User Story', '', 'Vestaras', 5588],
    ['', 'Active', '', 'User Story', '', 'Vestaras', 1],
    ['', ' ACTIVE ', '', 'Bug', '', 'Vestaras', 5],
    ['', 'Active', '', 'Task', '', 'Vestaras', 6],
    ['', 'Active', '', 'Bug', '', 'Other', 7],
    ['', 'In Progress', '', 'Task', '', 'Vestaras', 2],
    ['', 'In Progress', '', 'Bug', '', 'Other', 3],
    ['', 'Closed', '', 'Bug', '', 'Vestaras', 4],
    ['', 'In Progress', '', 'Bug', '', 'Vestaras', '']
  ];
  for (const bookType of ['xlsx', 'biff8']) {
    await upload(workbook(activeRows), bookType);
    assert.equal(inProgress(), 'US 5588 • Bug 5413 • US 005736 • US 1 • Bug 5');
  }
  await upload(workbook([rows[15]]));
  assert.equal(inProgress(), 'NA');
  await upload(workbook(activeRows.map(row => row.slice(1)), headers.slice(1)));
  assert.equal(inProgress(), 'NA');
  const resolvedRows = [
    ['Committed', 'Resolved', 'Jane', 'User Story', '', 'Vestaras', 100],
    ['Tentative', ' completed ', 'Jane', ' bug ', '', 'MES SE Vasteras', 101],
    ['', 'DONE', '', 'User Story', '', 'Vasteras', '00102'],
    ['', 'Resolved', '', 'User Story', '', 'Vestaras', 100],
    ['', 'Completed', '', 'User Story', '', 'Vestaras', 103],
    ['', 'Done', '', 'Bug', '', 'Vestaras', 104],
    ['', 'Resolved', '', 'Bug', '', 'Vestaras', 105],
    ['', 'Closed', '', 'User Story', '', 'Vestaras', 106],
    ['', ' CLOSED ', '', 'Bug', '', 'Vestaras', 110],
    ['', 'Resolved', '', 'Task', '', 'Vestaras', 107],
    ['', 'Completed', '', 'Bug', '', 'Other', 108],
    ['', 'Active', '', 'Bug', '', 'Vestaras', 109],
    ['', 'Done', '', 'Bug', '', 'Vestaras', '']
  ];
  for (const bookType of ['xlsx', 'biff8']) {
    await upload(workbook(resolvedRows), bookType);
    assert.equal(resolved(), 'US 100 • Bug 101 • US 00102 • US 103 • Bug 104 • Bug 105 • US 106 • Bug 110');
    assert.equal(inProgress(), 'Bug 109');
  }
  await upload(workbook([rows[0]]));
  assert.equal(resolved(), 'NA');
  await upload(workbook(resolvedRows.map(row => row.slice(1)), headers.slice(1)));
  assert.equal(resolved(), 'NA');
  const manyIds = Array.from({ length: 40 }, (_, index) => ['', index % 2 ? 'Closed' : 'Active', 'Jane', index % 3 ? 'Bug' : 'User Story', '', 'Vestaras', 1000 + index]);
  await upload(workbook(manyIds));
  const expectedIds = parity => manyIds.filter((_, index) => index % 2 === parity).map(row => `${row[3] === 'User Story' ? 'US' : 'Bug'} ${row[6]}`).join(' • ');
  assert.equal(inProgress(), expectedIds(0));
  assert.equal(resolved(), expectedIds(1));
  for (const index of [0, 1]) {
    const ids = get('burndown-cards').children[0].children[3].children[index].children[1];
    assert.equal(ids.children.filter(child => child.className === 'burndown-story-id').length, 20);
    assert.equal(ids.children.filter(child => child.className === 'burndown-id-separator').length, 19);
  }
  const tentativeDetails = () => get('burndown-cards').children[0].children[3].children[3].children[1].textContent;
  const filterRows = [
    ['Committed', 'Active', 'Start Person', 'User Story', 46279.75, 'Vestaras', 500],
    ['Committed', 'Closed', 'Middle Person', 'Bug', '2026-09-15T23:59:00', 'VastarAS', 501],
    ['Tentative', 'Done', 'End Person', 'User Story', '2026-09-25', 'MES SE Vasteras', 502],
    ['', 'Resolved', 'Named Date Person', 'Bug', '18-Sep-2026', 'Vestaras', 503],
    ['Committed', 'Active', 'Before Person', 'User Story', '2026-09-13', 'Vestaras', 504],
    ['Tentative', 'Done', 'After Person', 'Bug', '2026-09-26', 'Vestaras', 505],
    ['Committed', 'Active', 'Blank Date Person', 'Bug', '', 'Vestaras', 506],
    ['Committed', 'Done', 'Invalid Date Person', 'User Story', 'not a date', 'Vestaras', 507],
    ['Committed', 'Done', 'Invalid Day Person', 'User Story', '2026-02-30', 'Vestaras', 508],
    ['', 'Active', 'Other Site Person', 'Task', '2020-01-01', 'Rakovski', 509]
  ];
  // Use a supported alias; capitalization should not affect matching.
  filterRows[1][5] = ' vAsTeRaS ';
  const filteredBook = workbook(filterRows);
  filteredBook.Sheets['Master Daily Dump'].G6.z = 'dd-mmm-yyyy';
  await upload(filteredBook, 'xlsx', false);
  assert.equal(committed(), '2');
  assert.equal(tentative(), '1');
  assert.equal(delivered(), '3');
  assert.equal(completion(), '150%');
  assert.equal(inProgress(), 'US 500');
  assert.equal(resolved(), 'Bug 501 • US 502 • Bug 503');
  assert.equal(tentativeDetails(), 'US 502');
  assert.equal(developer(), 'Start Person, Middle Person, End Person, Named Date Person');
  assert.equal(developer(1), 'Unassigned');
  assert.match(get('sprint-committed-status').textContent, /Skipped 3/);
  const changeRange = (from, to) => {
    get('dashboard-date-from').value = from;
    get('dashboard-date-to').value = to;
    get('dashboard-date-from').listeners.change();
  };
  changeRange('2026-09-25', '2026-09-25');
  assert.equal(committed(), '0');
  assert.equal(tentative(), '1');
  assert.equal(delivered(), '1');
  assert.equal(resolved(), 'US 502');
  assert.equal(developer(), 'End Person');
  get('dashboard-date-to').value = '2026-09-26';
  get('dashboard-date-to').listeners.change();
  assert.equal(tentative(), '2');
  assert.equal(resolved(), 'US 502 • Bug 505');
  changeRange('2026-09-27', '2026-09-25');
  assert.equal(committed(), '—');
  assert.equal(inProgress(), 'NA');
  assert.match(get('sprint-committed-status').textContent, /valid From–To/);
  changeRange('2026-09-14', '2026-09-14');
  assert.equal(committed(), '1');
  assert.equal(inProgress(), 'US 500');
  assert.equal(developer(), 'Start Person');
  changeRange('2027-01-01', '2027-01-02');
  assert.equal(committed(), '0');
  assert.equal(delivered(), '0');
  assert.equal(developer(), 'Unassigned');
  assert.equal(tentativeDetails(), 'NA');
  changeRange('', '2026-09-25');
  assert.equal(committed(), '—');
  changeRange('2026-09-14', '2026-09-25');
  const book1904 = workbook([['Committed', 'Active', '1904 Person', 'Bug', 46279 - 1462, 'Vestaras', 600]]);
  book1904.Workbook = { WBProps: { date1904: true } };
  await upload(book1904, 'xlsx', false);
  assert.equal(committed(), '1');
  assert.equal(inProgress(), 'Bug 600');
  const tentativeRows = [
    ['Tentative', 'Active', 'Jane', 'User Story', '2026-09-14', 'Vestaras', 700],
    [' TENTATIVE ', 'Done', 'Bob', ' bug ', '2026-09-25', 'MES SE Vasteras', '00701'],
    ['Tentative', 'Closed', 'Jane', 'User Story', '2026-09-15', 'Vasteras', 700],
    ['Committed', 'Tentative', '', 'User Story', '2026-09-16', 'Vestaras', 702],
    ['Tentative', 'Tentative', '', 'Task', '2026-09-16', 'Vestaras', 703],
    ['Tentative', 'Tentative', '', 'Bug', '2026-09-16', 'Other', 704],
    ['Tentative', 'Tentative', '', 'Bug', '2026-09-26', 'Vestaras', 705],
    ['Tentative', 'Tentative', '', 'User Story', '2026-09-16', 'Vestaras', '']
  ];
  for (const bookType of ['xlsx', 'biff8']) {
    await upload(workbook(tentativeRows), bookType, false);
    assert.equal(tentativeDetails(), 'US 700 • Bug 00701');
    assert.equal(tentative(), '4'); // Tag rows count separately; duplicate/blank IDs do not repeat in the list.
  }
  changeRange('2026-09-25', '2026-09-25');
  assert.equal(tentativeDetails(), 'Bug 00701');
  changeRange('2026-09-17', '2026-09-18');
  assert.equal(tentativeDetails(), 'NA');
  const metricAt = (card, metric) => get('burndown-cards').children[card].children[1].children[metric].children[1].textContent;
  const detailAt = (card, detail) => get('burndown-cards').children[card].children[3].children[detail].children[1].textContent;
  const joao = assignments[2][0], ricardo = assignments[5][0];
  const rakovskiRows = [
    ['Committed', 'Active', joao, 'User Story', '2026-09-14', 'Rakovski', 900],
    ['Committed', 'Closed', ricardo, 'Bug', '2026-09-15', 'MES-Rakovski', 901],
    ['Tentative', 'Completed', ricardo, 'User Story', '2026-09-25', ' rAkOvSkI ', 902],
    ['Tentative', 'Done', joao, 'Bug', '2026-09-25', 'MES-Rakovski', 903],
    ['Committed', 'Active', 'Before Person', 'Bug', '2026-09-13', 'Rakovski', 904],
    ['Committed', 'Done', 'After Person', 'Bug', '2026-09-26', 'Rakovski', 905],
    ['Tentative', 'Active', 'Task Person', 'Task', '2026-09-16', 'Rakovski', 906],
    ['Committed', 'Active', 'Missing Date Person', 'Bug', '', 'Rakovski', 907],
    ['Committed', 'Active', 'Jane', 'User Story', '2026-09-16', 'Vestaras', 950]
  ];
  changeRange('2026-09-14', '2026-09-25');
  for (const bookType of ['xlsx', 'biff8']) {
    await upload(workbook(rakovskiRows), bookType, false);
    assert.equal(metricAt(1, 0), '2');
    assert.equal(metricAt(1, 1), '2');
    assert.equal(metricAt(1, 2), '2');
    assert.equal(completion(1), '100%');
    assert.equal(detailAt(1, 0), 'US 900');
    assert.equal(detailAt(1, 1), 'Bug 901 • US 902 • Bug 903');
    assert.equal(detailAt(1, 3), 'US 902 • Bug 903');
    assert.equal(developer(1), 'Joao, Ricardo, Task Person');
    assert.equal(committed(), '1');
    assert.equal(inProgress(), 'US 950');
    assert.match(get('sprint-committed-status').textContent, /Skipped 1 Rakovski/);
  }
  changeRange('2026-09-25', '2026-09-25');
  assert.equal(metricAt(1, 0), '0');
  assert.equal(metricAt(1, 1), '2');
  assert.equal(metricAt(1, 2), '1');
  assert.equal(completion(1), 'N/A');
  assert.equal(detailAt(1, 0), 'NA');
  assert.equal(detailAt(1, 1), 'US 902 • Bug 903');
  assert.equal(developer(1), 'Ricardo, Joao');
  assert.equal(committed(), '0');
  changeRange('2026-09-26', '2026-09-25');
  assert.equal(metricAt(1, 0), '—');
  assert.equal(detailAt(1, 1), 'NA');
  changeRange('2026-09-14', '2026-09-25');
  assert.equal(metricAt(1, 0), '2');
  const diogo = assignments[0][0], vinod = assignments[8][0];
  const heidelbergRows = rakovskiRows.filter(row => row[5].toLowerCase().includes('rakovski')).map((row, index) => {
    const copy = [...row];
    copy[2] = copy[2] === joao ? diogo : copy[2] === ricardo ? vinod : copy[2];
    copy[5] = index % 2 ? 'MES_DE_Heidelberg' : ' hEiDeLbErG ';
    copy[6] += 1000;
    return copy;
  });
  heidelbergRows.push(['', 'In Progress', diogo, 'User Story', '2026-09-16', 'Heidelberg', 1910]);
  for (const bookType of ['xlsx', 'biff8']) {
    await upload(workbook([...rakovskiRows, ...heidelbergRows]), bookType, false);
    assert.equal(metricAt(2, 0), '2');
    assert.equal(metricAt(2, 1), '2');
    assert.equal(metricAt(2, 2), '2');
    assert.equal(completion(2), '100%');
    assert.equal(detailAt(2, 0), 'US 1900 • US 1910');
    assert.equal(detailAt(2, 1), 'Bug 1901 • US 1902 • Bug 1903');
    assert.equal(detailAt(2, 3), 'US 1902 • Bug 1903');
    assert.equal(developer(2), 'Diogo, Vinod, Task Person');
    assert.match(get('sprint-committed-status').textContent, /Skipped 1 Heidelberg/);
    assert.equal(metricAt(1, 0), '2');
    assert.equal(detailAt(1, 0), 'US 900');
    assert.equal(inProgress(), 'US 950');
  }
  changeRange('2026-09-25', '2026-09-25');
  assert.equal(metricAt(2, 0), '0');
  assert.equal(metricAt(2, 1), '2');
  assert.equal(metricAt(2, 2), '1');
  assert.equal(completion(2), 'N/A');
  assert.equal(detailAt(2, 0), 'NA');
  assert.equal(detailAt(2, 1), 'US 1902 • Bug 1903');
  assert.equal(developer(2), 'Vinod, Diogo');
  changeRange('2026-09-26', '2026-09-25');
  assert.equal(metricAt(2, 0), '—');
  assert.equal(detailAt(2, 1), 'NA');
  changeRange('2026-09-14', '2026-09-25');
  assert.equal(metricAt(2, 0), '2');
  await upload(workbook(rakovskiRows), 'xlsx', false);
  assert.equal(metricAt(2, 0), '0');
  assert.equal(developer(2), 'Unassigned');
  assert.equal(detailAt(2, 1), 'NA');
  const suman = assignments[7][0];
  const senatobiaRows = rakovskiRows.filter(row => row[5].toLowerCase().includes('rakovski')).map((row, index) => {
    const copy = [...row];
    copy[2] = copy[2] === joao || copy[2] === ricardo ? suman : copy[2];
    copy[5] = index % 2 ? 'MES-Senatobia' : ' sEnAtObIa ';
    copy[6] += 2000;
    return copy;
  });
  senatobiaRows.push(['', 'In Progress', suman, 'Bug', '2026-09-16', 'Senatobia', 2910]);
  for (const bookType of ['xlsx', 'biff8']) {
    await upload(workbook([...rakovskiRows, ...heidelbergRows, ...senatobiaRows]), bookType, false);
    assert.equal(metricAt(3, 0), '2');
    assert.equal(metricAt(3, 1), '2');
    assert.equal(metricAt(3, 2), '2');
    assert.equal(completion(3), '100%');
    assert.equal(detailAt(3, 0), 'US 2900 • Bug 2910');
    assert.equal(detailAt(3, 1), 'Bug 2901 • US 2902 • Bug 2903');
    assert.equal(detailAt(3, 3), 'US 2902 • Bug 2903');
    assert.equal(developer(3), 'Suman, Task Person');
    assert.match(get('sprint-committed-status').textContent, /Skipped 1 Senatobia/);
    assert.equal(detailAt(2, 0), 'US 1900 • US 1910');
    assert.equal(detailAt(1, 0), 'US 900');
    assert.equal(inProgress(), 'US 950');
  }
  changeRange('2026-09-25', '2026-09-25');
  assert.equal(metricAt(3, 0), '0');
  assert.equal(metricAt(3, 1), '2');
  assert.equal(metricAt(3, 2), '1');
  assert.equal(completion(3), 'N/A');
  assert.equal(detailAt(3, 0), 'NA');
  assert.equal(detailAt(3, 1), 'US 2902 • Bug 2903');
  assert.equal(developer(3), 'Suman');
  changeRange('2026-09-26', '2026-09-25');
  assert.equal(metricAt(3, 0), '—');
  assert.equal(detailAt(3, 1), 'NA');
  changeRange('2026-09-14', '2026-09-25');
  assert.equal(metricAt(3, 0), '2');
  await upload(workbook(rakovskiRows), 'xlsx', false);
  assert.equal(metricAt(3, 0), '0');
  assert.equal(developer(3), 'Unassigned');
  assert.equal(detailAt(3, 1), 'NA');
  const faiz = assignments[1][0];
  const alingsasRows = rakovskiRows.filter(row => row[5].toLowerCase().includes('rakovski')).map((row, index) => {
    const copy = [...row];
    copy[2] = copy[2] === joao || copy[2] === ricardo ? faiz : copy[2];
    copy[5] = index % 2 ? 'MES Alingsas' : ' aLiNgSaS ';
    copy[6] += 3000;
    return copy;
  });
  alingsasRows.push(['', 'In Progress', faiz, 'Bug', '2026-09-16', 'Alingsas', 3910]);
  for (const bookType of ['xlsx', 'biff8']) {
    await upload(workbook([...rakovskiRows, ...heidelbergRows, ...senatobiaRows, ...alingsasRows]), bookType, false);
    assert.equal(metricAt(4, 0), '2');
    assert.equal(metricAt(4, 1), '2');
    assert.equal(metricAt(4, 2), '2');
    assert.equal(completion(4), '100%');
    assert.equal(detailAt(4, 0), 'US 3900 • Bug 3910');
    assert.equal(detailAt(4, 1), 'Bug 3901 • US 3902 • Bug 3903');
    assert.equal(detailAt(4, 3), 'US 3902 • Bug 3903');
    assert.equal(developer(4), 'Faiz, Task Person');
    assert.match(get('sprint-committed-status').textContent, /Skipped 1 Alingsas/);
    assert.equal(detailAt(3, 0), 'US 2900 • Bug 2910');
    assert.equal(detailAt(2, 0), 'US 1900 • US 1910');
    assert.equal(detailAt(1, 0), 'US 900');
    assert.equal(inProgress(), 'US 950');
  }
  changeRange('2026-09-25', '2026-09-25');
  assert.equal(metricAt(4, 0), '0');
  assert.equal(metricAt(4, 1), '2');
  assert.equal(metricAt(4, 2), '1');
  assert.equal(completion(4), 'N/A');
  assert.equal(detailAt(4, 0), 'NA');
  assert.equal(detailAt(4, 1), 'US 3902 • Bug 3903');
  assert.equal(developer(4), 'Faiz');
  changeRange('2026-09-26', '2026-09-25');
  assert.equal(metricAt(4, 0), '—');
  assert.equal(detailAt(4, 1), 'NA');
  changeRange('2026-09-14', '2026-09-25');
  assert.equal(metricAt(4, 0), '2');
  await upload(workbook(rakovskiRows), 'xlsx', false);
  assert.equal(metricAt(4, 0), '0');
  assert.equal(developer(4), 'Unassigned');
  assert.equal(detailAt(4, 1), 'NA');
  const nivruth = assignments[4][0], shaalini = assignments[6][0];
  const chassieuRows = rakovskiRows.filter(row => row[5].toLowerCase().includes('rakovski')).map((row, index) => {
    const copy = [...row];
    copy[2] = copy[2] === joao ? nivruth : copy[2] === ricardo ? shaalini : copy[2];
    copy[5] = index % 2 ? 'MES FR Chassieu' : ' cHaSsIeU ';
    copy[6] += 4000;
    return copy;
  });
  chassieuRows.push(['', 'In Progress', nivruth, 'Bug', '2026-09-16', 'Chassieu', 4910]);
  for (const bookType of ['xlsx', 'biff8']) {
    await upload(workbook([...rakovskiRows, ...heidelbergRows, ...senatobiaRows, ...alingsasRows, ...chassieuRows]), bookType, false);
    assert.equal(metricAt(5, 0), '2');
    assert.equal(metricAt(5, 1), '2');
    assert.equal(metricAt(5, 2), '2');
    assert.equal(completion(5), '100%');
    assert.equal(detailAt(5, 0), 'US 4900 • Bug 4910');
    assert.equal(detailAt(5, 1), 'Bug 4901 • US 4902 • Bug 4903');
    assert.equal(detailAt(5, 3), 'US 4902 • Bug 4903');
    assert.equal(developer(5), 'Nivruth, Shaalini, Task Person');
    assert.match(get('sprint-committed-status').textContent, /Skipped 1 Chassieu/);
    assert.equal(detailAt(4, 0), 'US 3900 • Bug 3910');
    assert.equal(detailAt(3, 0), 'US 2900 • Bug 2910');
    assert.equal(detailAt(2, 0), 'US 1900 • US 1910');
    assert.equal(detailAt(1, 0), 'US 900');
    assert.equal(inProgress(), 'US 950');
  }
  changeRange('2026-09-25', '2026-09-25');
  assert.equal(metricAt(5, 0), '0');
  assert.equal(metricAt(5, 1), '2');
  assert.equal(metricAt(5, 2), '1');
  assert.equal(completion(5), 'N/A');
  assert.equal(detailAt(5, 0), 'NA');
  assert.equal(detailAt(5, 1), 'US 4902 • Bug 4903');
  assert.equal(developer(5), 'Shaalini, Nivruth');
  changeRange('2026-09-26', '2026-09-25');
  assert.equal(metricAt(5, 0), '—');
  assert.equal(detailAt(5, 1), 'NA');
  changeRange('2026-09-14', '2026-09-25');
  assert.equal(metricAt(5, 0), '2');
  await upload(workbook(rakovskiRows), 'xlsx', false);
  assert.equal(metricAt(5, 0), '0');
  assert.equal(developer(5), 'Unassigned');
  assert.equal(detailAt(5, 1), 'NA');
  get('site-excel-file').files = []; get('site-excel-file').listeners.change();
  assert.equal(developer(), 'Unassigned');
  assert.equal(completion(), 'N/A');
  assert.equal(inProgress(), 'NA');
  assert.equal(resolved(), 'NA');
  assert.equal(tentativeDetails(), 'NA');
  assert.equal(metricAt(1, 0), '—');
  assert.equal(detailAt(1, 0), 'NA');
  assert.equal(detailAt(1, 1), 'NA');
  assert.equal(detailAt(1, 3), 'NA');
  assert.equal(metricAt(2, 0), '—');
  assert.equal(detailAt(2, 0), 'NA');
  assert.equal(detailAt(2, 1), 'NA');
  assert.equal(detailAt(2, 3), 'NA');
  assert.equal(metricAt(3, 0), '—');
  assert.equal(detailAt(3, 0), 'NA');
  assert.equal(detailAt(3, 1), 'NA');
  assert.equal(detailAt(3, 3), 'NA');
  assert.equal(metricAt(4, 0), '—');
  assert.equal(detailAt(4, 0), 'NA');
  assert.equal(detailAt(4, 1), 'NA');
  assert.equal(detailAt(4, 3), 'NA');
  assert.equal(metricAt(5, 0), '—');
  assert.equal(detailAt(5, 0), 'NA');
  assert.equal(detailAt(5, 1), 'NA');
  assert.equal(detailAt(5, 3), 'NA');
  changeRange('2026-09-15', '2026-09-20');
  assert.equal(committed(), '—'); // Cleared uploads must not reappear on date edits.
  // Semicolon-separated tags count once per category, regardless of order or repetition.
  for (const bookType of ['xlsx', 'xls']) {
    const multiTagRows = [];
    for (const site of ['Vasteras', 'Rakovski', 'Heidelberg', 'Senatobia', 'Alingsas', 'Chassieu']) {
      for (const [tag, type, date, id] of [
        ['Priority; Committed ;Team', 'User Story', '2026-09-16', 901],
        [' TENTATIVE; Priority ', 'Bug', '2026-09-16', 902],
        ['committed; tentative; committed; TENTATIVE;', 'Bug', '2026-09-16', 903],
        ['Not Committed;Not Tentative;Uncommitted;Tentatively', 'Bug', '2026-09-16', 904],
        ['', 'Bug', '2026-09-16', 905],
        ['Committed;Tentative', 'Task', '2026-09-16', 906],
        ['Committed;Tentative', 'Bug', '2026-09-25', 907]
      ]) multiTagRows.push([tag, 'Active', 'Developer', type, date, site, id]);
    }
    await upload(workbook(multiTagRows), bookType, false);
    for (let siteIndex = 0; siteIndex < 6; siteIndex++) {
      assert.equal(metricAt(siteIndex, 0), '2');
      assert.equal(metricAt(siteIndex, 1), '2');
      assert.equal(detailAt(siteIndex, 3), 'Bug 902 • Bug 903');
    }
  }
  console.log('Passed: Dashboard upload integration, XLS/XLSX, reordered/offset headers, site aliases, row counts, semicolon-separated tags, type/tag exclusions, zero counts, replacement/reset, missing headers/sheets and read failures.');
})().catch(error => { console.error(error); process.exitCode = 1; });
