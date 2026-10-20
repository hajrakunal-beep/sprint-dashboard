const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const XLSX = require('../vendor/xlsx.full.min.js');
class Element {
  constructor() { this.children = []; this.attrs = {}; this.style = {}; this.listeners = {}; }
  append(...children) { this.children.push(...children); }
  replaceChildren(...children) { this.children = children; }
  setAttribute(name, value) { this.attrs[name] = value; }
  addEventListener(name, handler) { this.listeners[name] = handler; }
  contains(target) { return target === this || this.children.some(child => child.contains(target)); }
  querySelector() { return this.children[0]; }
  focus() { this.focused = true; }
}
const elements = new Map();
const get = id => { if (!elements.has(id)) elements.set(id, new Element()); return elements.get(id); };
const documentListeners = {};
const context = vm.createContext({ XLSX, document: { getElementById: get, createElement: () => new Element(),
  addEventListener: (name, handler) => { documentListeners[name] = handler; } },
  dashboardDateFrom: { value: '2026-10-01' }, dashboardDateTo: { value: '2026-10-15' },
  parseDashboardDate: value => /^\d{4}-\d{2}-\d{2}$/.test(value) && Number.isFinite(new Date(value).getTime()) ? new Date(value) : null,
  normalizeColumn: value => String(value ?? '').trim().toLowerCase(),
  sprintSiteKey: value => {
    const aliases = { 'mes se vasteras': 'vasteras', 'vestaras': 'vasteras', 'mes-rakovski': 'rakovski',
      'mes_de_heidelberg': 'heidelberg', 'mes-senatobia': 'senatobia', 'mes fr chassieu': 'chassieu' };
    const key = String(value).trim().toLowerCase();
    return aliases[key] || key.replace(/^mes /, '');
  },
  masterDailyDumpSheet: book => { if (!book.Sheets['Master Daily Dump']) throw new Error('Master Daily Dump missing'); return 'Master Daily Dump'; },
  sprintRowDate: value => /^\d{4}-\d{2}-\d{2}$/.test(value) ? new Date(value) : null,
  formatSprintDuration: dates => `[${dates.sprintStartDate} to ${dates.sprintEndDate}]`
});
context.dynamicSprintSites = new Set(['alingsas', 'vasteras', 'rakovski', 'heidelberg', 'senatobia', 'chassieu']);
const sprintSource = fs.readFileSync(path.join(__dirname, '..', 'sprint-burndown.js'), 'utf8');
vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'site-filter.js'), 'utf8'), context);
vm.runInContext(sprintSource.slice(sprintSource.indexOf('function readSprintDashboardDates'), sprintSource.indexOf('function updateSprintDashboardDates')), context);
vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'user-story-status.js'), 'utf8'), context);
const headers = ['Area Path', 'ID', 'Work Item Type', 'Parent', 'Completed Work', 'Remaining Work', 'Original Estimate', 'Date'];
const book = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([headers,
  ['Project\\MES Alingsas', 100, 'User Story', '', '', '', '', '2026-10-01'],
  ['MES Alingsas', 101, 'Task', 100, 1, 9, 8, '2026-10-01'],
  ['MES Alingsas', 101, 'Task', 100, 8, 2, 8, '2026-10-02'],
  ['Alingsas', 102, 'Task', 100, 9, '', 10, '2026-10-02'],
  ['Alingsas', 103, 'Task', 999, 1, 0, 1, '2026-10-02'],
  ['Other', 200, 'User Story', '', '', '', '', '2026-10-02']
]), 'Master Daily Dump');
XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([
  ['Site', 'Start', 'End', ' Sprint '],
  ['Other', '2026-01-01', '2026-12-31', 'Current'],
  ['MES Alingsas', '01/10/2026', '15/10/2026', 'Sprint 20']
]), 'Dashboard');
const data = context.readUserStoryStatus(book);
assert.equal(data.dates.sprintStartDate, '2026-10-01');
assert.equal(data.dates.sprintEndDate, '2026-10-15');
const filteredBook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(filteredBook, book.Sheets['Dashboard'], 'Dashboard');
XLSX.utils.book_append_sheet(filteredBook, XLSX.utils.aoa_to_sheet([headers,
  ['Alingsas', 5533, 'User Story', '', '', '', '', '2026-10-01'],
  ['Alingsas', 5590, 'Task', 5533, 1, 9, 10, '2026-09-30'],
  ['Alingsas', 5590, 'Task', 5533, 8, 2, 10, '2026-10-15'],
  ['Alingsas', 5590, 'Task', 5533, 10, 0, 10, '2026-10-16'],
  ['Alingsas', 5591, 'Task', 5533, 3, 7, 10, '2026-10-01'],
  ['Alingsas', 5592, 'Task', 5533, 2, 8, 10, ''],
  ['Alingsas', 5593, 'Task', 5533, 'invalid', '', '', '2026-10-16'],
  ['Alingsas', 9999, 'User Story', '', '', '', '', '2026-09-30']
]), 'Master Daily Dump');
const filteredData = context.readUserStoryStatus(filteredBook);
assert.equal(filteredData.stories.length, 1);
assert.equal(filteredData.stories[0].tasks.length, 2);
assert.equal(filteredData.stories[0].tasks[0].completed, 8);
assert.ok(Math.abs(context.userStoryProgress(filteredData.stories[0].tasks) - 55) < 1e-10);
const noDashboard = { SheetNames: ['Master Daily Dump'], Sheets: { 'Master Daily Dump': book.Sheets['Master Daily Dump'] } };
assert.throws(() => context.readUserStoryStatus(noDashboard), /Dashboard.*missing/);
const conflictingBook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(conflictingBook, XLSX.utils.aoa_to_sheet([
  ['Site', 'Start', 'End', 'Sprint'],
  ['Alingsas', '2026-10-01', '2026-10-15', 'Current'],
  ['Alingsas', '2026-10-16', '2026-10-30', 'Current']
]), 'Dashboard');
assert.throws(() => context.readUserStorySprintDates(conflictingBook), /conflicting/);
context.dashboardDateFrom.value = '2026-10-15';
context.dashboardDateTo.value = '2026-10-15';
assert.equal(context.readUserStoryStatus(filteredBook).stories.length, 1);
context.dashboardDateFrom.value = 'invalid';
assert.equal(context.readUserStoryStatus(filteredBook).stories[0].tasks[0].completed, 8);
context.dashboardDateFrom.value = '2026-10-01';
context.dashboardDateTo.value = '2026-10-15';
assert.equal(data.stories.length, 1);
assert.equal(data.stories[0].id, '100');
assert.equal(data.stories[0].tasks.length, 2);
assert.equal(context.taskStoryProgress(data.stories[0].tasks[0]), 80);
assert.equal(context.taskStoryProgress(data.stories[0].tasks[1]), 90);
for (const [value, color] of [[0, 'blue'], [75, 'blue'], [75.1, 'orange'], [85, 'orange'], [85.1, 'purple'], [99.9, 'purple'], [100, 'green'], [null, 'neutral']]) {
  assert.equal(context.taskStoryColor(value), color);
}
assert.equal(context.taskStoryProgress({ completed: null, remaining: 2, estimate: 10 }), 0);
assert.equal(context.taskStoryProgress({ completed: 0, remaining: 0, estimate: 0 }), null);
assert.equal(context.userStoryProgress(data.stories[0].tasks), 85);
assert.equal(context.userStoryProgress([{ completed: 1, remaining: 0 }, { completed: 0, remaining: 9 }]), 10);
assert.equal(context.userStoryProgress([{ completed: 3, remaining: null, estimate: 4 }, { completed: 6, remaining: 0 }]), 90);
assert.equal(context.userStoryProgress([{ completed: 12, remaining: null, estimate: 10 }]), 100);
assert.equal(context.userStoryProgress([]), null);
assert.equal(context.userStoryProgress([{ completed: null, remaining: 1 }]), 0);
assert.equal(context.userStoryProgress([{ completed: 1, remaining: null, estimate: null }]), null);
assert.equal(context.userStoryProgress([{ completed: 0, remaining: 0 }]), null);
const tasks5533 = [
  { id: '5590', estimate: 16, remaining: null, completed: 16 },
  { id: '5591', estimate: 16, remaining: null, completed: 16 },
  { id: '5592', estimate: 16, remaining: null, completed: 20 },
  { id: '5593', estimate: 10, remaining: null, completed: 10 },
  { id: '5637', estimate: 16, remaining: null, completed: 18 },
  { id: '5589', estimate: 10, remaining: null, completed: 10 },
  { id: '5594', estimate: 10, remaining: 10, completed: null }
];
assert.equal(context.taskStoryProgress(tasks5533[6]), 0);
assert.equal(context.userStoryProgress(tasks5533), 90 / 94 * 100);
context.renderUserStoryStatus({ dates: data.dates, stories: [{ id: '5533', tasks: tasks5533 }] });
const bar5533 = get('user-story-cards').children[0].children[1].children[0].children[0];
assert.equal(bar5533.children[2].textContent, '95.7%');
assert.equal(bar5533.className, 'story-progress story-progress-purple');
context.renderUserStoryStatus({ ...data, stories: [...data.stories, { id: '110', tasks: [] }] });
assert.equal(get('user-story-cards').children.length, 1);
const card = get('user-story-cards').children[0];
assert.equal(card.children[0].children[0].textContent, '[2026-10-01 to 2026-10-15]');
assert.equal(card.children[0].children[1].textContent, 'Alingsas');
const trees = card.children[1];
assert.equal(trees.children.length, 2);
assert.equal(trees.children[1].children[0].children[1].textContent, 'US 110');
const tree = trees.children[0];
assert.equal(tree.children[0].children[1].textContent, 'US 100');
assert.equal(tree.children[0].children[2].textContent, '85%');
assert.equal(tree.children[0].children[0].style.width, '85%');
assert.equal(tree.children[0].className, 'story-progress story-progress-orange');
assert.equal(tree.children[0].attrs['aria-valuenow'], '85');
assert.equal(trees.children[1].children[0].children[2].textContent, '—');
assert.equal(get('user-story-status-message').textContent, '');
assert.equal(get('user-story-status-message').hidden, true);
const taskBar = tree.children[1].children[0].children[0];
assert.equal(taskBar.children[1].textContent, 'Task 101');
assert.equal(taskBar.children[2].textContent, '80%');
assert.equal(taskBar.children[0].style.width, '80%');
assert.equal(taskBar.attrs['aria-valuenow'], '80');
const allSitesBook = XLSX.utils.book_new();
const siteNames = ['Alingsas', 'Vasteras', 'Rakovski', 'Heidelberg', 'Senatobia', 'Chassieu'];
const areaNames = ['MES Alingsas', 'MES SE Vasteras', 'MES-Rakovski', 'MES_DE_Heidelberg', 'MES-Senatobia', 'MES FR Chassieu'];
XLSX.utils.book_append_sheet(allSitesBook, XLSX.utils.aoa_to_sheet([
  ['Site', 'Start', 'End', 'Sprint'],
  ...siteNames.map((site, index) => [site, `2026-10-${String(index + 1).padStart(2, '0')}`, `2026-10-${String(index + 8).padStart(2, '0')}`, 'Current'])
]), 'Dashboard');
XLSX.utils.book_append_sheet(allSitesBook, XLSX.utils.aoa_to_sheet([headers,
  ...areaNames.flatMap((site, index) => [
    // Reuse IDs to check that parents and daily snapshots stay within their site.
    [`Project\\${site}`, 100, 'User Story', '', '', '', '', `2026-10-${String(index + 1).padStart(2, '0')}`],
    [site, 101, 'Task', 100, index + 1, 9 - index, 10, `2026-10-${String(index + 8).padStart(2, '0')}`],
    [site, 999, 'User Story', '', '', '', '', `2026-10-${String(index + 9).padStart(2, '0')}`],
    [site, 101, 'Task', 100, 10, 0, 10, `2026-10-${String(index + 9).padStart(2, '0')}`]
  ])
]), 'Master Daily Dump');
assert.match(context.updateUserStoryStatus(allSitesBook), /Loaded 6 user stories across 6 sites/);
assert.equal(get('user-story-cards').children.length, 6);
siteNames.forEach((site, index) => {
  const siteData = context.readUserStoryStatus(allSitesBook, site);
  assert.equal(siteData.stories.length, 1);
  assert.equal(siteData.stories[0].tasks.length, 1);
  assert.equal(siteData.dates.sprintStartDate, `2026-10-${String(index + 1).padStart(2, '0')}`);
  assert.equal(siteData.stories[0].tasks[0].completed, index + 1);
  const siteCard = get('user-story-cards').children[index];
  assert.equal(siteCard.children[0].children[1].textContent, site);
  assert.equal(siteCard.children[1].children[0].children[0].children[2].textContent, `${(index + 1) * 10}%`);
});
// A missing range for one site must not prevent the other site cards from loading.
delete allSitesBook.Sheets['Dashboard'].A7;
assert.throws(() => context.readUserStorySprintDates(allSitesBook, 'Chassieu'), /Dashboard sprint dates/);
assert.match(context.updateUserStoryStatus(allSitesBook), /across 5 sites/);
assert.equal(get('user-story-cards').children.length, 6);
assert.match(get('user-story-cards').children[5].children[1].children[0].textContent, /Dashboard sprint dates/);
// Exercise checkbox changes against loaded data, including refresh and empty selection.
const options = get('user-story-site-options').children;
assert.equal(options.length, 7);
const allCheckbox = options[0].children[1];
const toggle = (index, checked) => {
  const checkbox = options[index].children[1];
  checkbox.checked = checked;
  checkbox.listeners.change();
};
siteNames.forEach((site, index) => {
  assert.equal(options[index + 1].children[0].textContent, site);
  assert.equal(options[index + 1].children[1].checked, true);
});
toggle(2, false);
assert.equal(get('user-story-cards').children.length, 5);
assert.ok(get('user-story-cards').children.every(card => card.children[0].children[1].textContent !== 'Vasteras'));
assert.equal(get('user-story-site-selection').textContent, '5 selected');
assert.equal(allCheckbox.checked, false);
assert.equal(allCheckbox.indeterminate, true);
context.updateUserStoryStatus(allSitesBook);
assert.equal(get('user-story-cards').children.length, 5);
toggle(2, true);
assert.equal(get('user-story-cards').children.length, 6);
assert.equal(get('user-story-cards').children[1].children[0].children[1].textContent, 'Vasteras');
assert.equal(get('user-story-site-selection').textContent, 'All');
assert.equal(allCheckbox.indeterminate, false);
toggle(0, false);
assert.equal(get('user-story-cards').children.length, 0);
assert.equal(get('user-story-status-message').hidden, false);
assert.match(get('user-story-status-message').textContent, /No sites selected/);
assert.equal(get('user-story-site-selection').textContent, 'None');
assert.ok(options.every(option => !option.children[1].checked));
toggle(4, true);
assert.equal(get('user-story-cards').children.length, 1);
assert.equal(get('user-story-cards').children[0].children[0].children[1].textContent, 'Heidelberg');
toggle(0, true);
assert.equal(get('user-story-cards').children.length, 6);
assert.ok(options.every(option => option.children[1].checked));
const filter = get('user-story-site-filter');
// The DOM mock needs to attach the options container to the details element.
// Escape should close the dropdown and return focus to its summary.
const summary = new Element();
filter.append(summary, get('user-story-site-options'));
filter.open = true;
documentListeners.click({ target: options[1].children[1] });
assert.equal(filter.open, true);
filter.listeners.keydown({ key: 'Escape' });
assert.equal(filter.open, false);
assert.equal(summary.focused, true);
filter.open = true;
documentListeners.click({ target: new Element() });
assert.equal(filter.open, false);
context.resetUserStoryStatus();
assert.equal(get('user-story-status-message').hidden, false);
assert.equal(get('user-story-cards').children.length, 0);
assert.match(context.updateUserStoryStatus({ SheetNames: [], Sheets: {} }), /unavailable/);
console.log('Passed: Dashboard site ranges, Settings independence, missing range isolation, story/task matching, latest snapshots, percentages, rendering, site checkbox filtering, refresh, dropdown dismissal and resets.');
