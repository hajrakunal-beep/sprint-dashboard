const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
class Element {
  constructor(tag = 'div') { this.tag = tag; this.children = []; this.listeners = {}; }
  append(...children) { this.children.push(...children); }
  replaceChildren(...children) { this.children = children; }
  setAttribute() {}
  addEventListener(name, handler) { this.listeners[name] = handler; }
  querySelector(tag) { return this.children.find(child => child.tag === tag); }
  contains(target) { return target === this || this.children.some(child => child.contains(target)); }
  focus() { this.focused = true; }
}
const elements = new Map();
const get = id => { if (!elements.has(id)) elements.set(id, new Element()); return elements.get(id); };
const listeners = {};
const context = vm.createContext({
  document: { getElementById: get, createElement: tag => new Element(tag), addEventListener: (name, fn) => { listeners[name] = fn; } },
  dashboardDateFrom: { value: '2026-10-01' }, dashboardDateTo: { value: '2026-10-15' },
  statusFilterSiteName: value => String(value).split('\\').pop().replace(/^MES /, ''),
});
vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'site-filter.js'), 'utf8'), context);
vm.runInContext(fs.readFileSync(path.join(__dirname, '../burndown-chart.js'), 'utf8'), context);
context.initializeBurndownSiteFilter();
// Test the filter independently of Excel parsing; date-filter tests cover the parser.
context.readBurndownWorkbook = data => data;
const options = get('burndown-chart-site-options');
const filter = get('burndown-chart-site-filter');
filter.append(new Element('summary'), options);
assert.equal(options.children.length, 7);
assert.ok(options.children.every(label => label.children[1].checked));
const data = { groups: ['MES Alingsas', 'Project\\MES Alingsas', 'Vasteras', 'Other'].map(site => ({ site, points: [] })) };
context.updateSharedBurndown(data, 'report.xlsx');
assert.equal(options.children.length, 8);
assert.equal(get('burndown-site-charts').children.length, 4);
const toggle = (index, checked) => {
  const input = options.children[index].children[1];
  input.checked = checked;
  input.listeners.change();
};
toggle(1, false);
assert.equal(get('burndown-site-charts').children.length, 2);
assert.equal(get('burndown-chart-site-selection').textContent, '6 selected');
assert.equal(options.children[0].children[1].indeterminate, true);
context.refreshBurndownDateRange();
assert.equal(get('burndown-site-charts').children.length, 2);
toggle(1, true);
assert.equal(get('burndown-site-charts').children.length, 4);
toggle(0, false);
assert.equal(get('burndown-site-charts').children.length, 0);
assert.match(get('burndown-shared-status').textContent, /No sites selected/);
assert.equal(get('burndown-chart-site-selection').textContent, 'None');
toggle(2, true);
assert.equal(get('burndown-site-charts').children[0].children[0].textContent, 'Vasteras');
toggle(0, true);
assert.equal(get('burndown-site-charts').children.length, 4);
filter.open = true;
listeners.click({ target: options.children[1].children[1] });
assert.equal(filter.open, true);
filter.listeners.keydown({ key: 'Escape' });
assert.equal(filter.open, false);
assert.equal(filter.querySelector('summary').focused, true);
filter.open = true;
listeners.click({ target: new Element() });
assert.equal(filter.open, false);
context.resetSharedBurndown();
toggle(1, false);
assert.equal(get('burndown-site-charts').children.length, 0);
assert.match(get('burndown-shared-status').textContent, /Upload an Excel/);
context.updateSharedBurndown(data, 'new-report.xlsx');
assert.equal(get('burndown-site-charts').children.length, 2);
console.log('Passed: Burndown site defaults, checkbox filtering, aliases, extra sites, refresh, select all, reset and dropdown dismissal.');
