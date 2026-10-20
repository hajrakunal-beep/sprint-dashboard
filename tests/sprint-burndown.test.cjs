const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

class Element {
  constructor() {
    this.children = []; this.attrs = {}; this.listeners = {}; this.textContent = '';
    this.active = false;
    this.classList = { toggle: (_, active) => { this.active = active; }, contains: () => false };
  }
  append(...children) { this.children.push(...children); }
  replaceChildren(...children) { this.children = children; }
  setAttribute(key, value) { this.attrs[key] = value; }
  removeAttribute(key) { delete this.attrs[key]; }
  addEventListener(key, handler) { this.listeners[key] = handler; }
  focus() { this.focused = true; }
}
const elements = new Map();
const get = id => {
  if (!elements.has(id)) elements.set(id, new Element());
  return elements.get(id);
};
const context = vm.createContext({ document: { getElementById: get, querySelector: get, createElement: () => new Element(), body: get('body'), addEventListener() {} }, window: { matchMedia: () => ({ matches: false }) } });
const read = name => fs.readFileSync(path.join(__dirname, '..', name), 'utf8');
vm.runInContext(read('script.js').replace(/renderDashboard\(dashboardData\);\s*$/, ''), context);
vm.runInContext(read('sprint-burndown.js'), context);
assert.equal(get('burndown-cards').children.length, 6);
const first = get('burndown-cards').children[0];
assert.equal(first.children[0].children[0].children[0].textContent, 'Phatale, Mangesh');
assert.equal(first.children[1].children[0].children[1].textContent, '3');
assert.equal(first.children[3].children[0].children[1].textContent, 'US 5588 • US 5736');
for (const view of ['sprint-burndown', 'site-status', 'burndown-chart', 'dashboard']) {
  get(`${view}-nav`).listeners.click({ preventDefault() {} });
  for (const id of ['sprint-burndown', 'site-status', 'burndown-chart', 'dashboard']) {
    assert.equal(get(id).hidden, id !== view);
    assert.equal(get(`${id}-nav`).active, id === view);
  }
  assert.equal(get(`#${view} h1`).focused, true);
}
context.renderSprintBurndown([{ developer: '<b>New developer</b>', site: 'New site', completionPercent: 75, committed: 4, delivered: 3, stories: { resolved: ['US 1'] } }]);
assert.equal(get('burndown-cards').children.length, 1);
assert.equal(get('burndown-cards').children[0].children[0].children[0].children[0].textContent, '<b>New developer</b>');
assert.equal(get('burndown-cards').children[0].children[0].children[1].textContent, '75%');
context.renderSprintBurndown([]);
assert.equal(get('burndown-cards').children[0].textContent, 'No sprint commitment data available.');
console.log('Passed: six reference cards, metrics and stories, three-view navigation, dynamic replacement, safe text, and empty data.');
