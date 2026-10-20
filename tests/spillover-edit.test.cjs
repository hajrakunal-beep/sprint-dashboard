const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

class Element {
  constructor() {
    this.children = [];
    this.textContent = '';
    this.attrs = {};
    this.listeners = {};
    this.classList = { toggle: (name, enabled) => {
      const classes = new Set(this.className.split(/\s+/));
      if (enabled) classes.add(name); else classes.delete(name);
      this.className = [...classes].join(' ');
    } };
  }
  append(...children) { this.children.push(...children); }
  replaceChildren(...children) { this.children = children; }
  setAttribute(name, value) { this.attrs[name] = value; }
  addEventListener(name, handler) { this.listeners[name] = handler; }
  focus() { this.focused = true; }
  blur() { this.listeners.blur?.(); }
}
const container = new Element();
const context = vm.createContext({
  document: { createElement: () => new Element(), getElementById: () => container },
  parseDashboardDate: () => null,
});
const source = fs.readFileSync(path.join(__dirname, '../sprint-burndown.js'), 'utf8');
vm.runInContext(source.replace(/resetSprintTagCounts\(\);\s*$/, ''), context);
const record = { site: 'Vasteras', spillover: 0 };
const render = () => context.renderSprintBurndown([record]);
const metric = () => container.children[0].children[1].children[3];
const value = () => metric().children[1];
const key = name => value().listeners.keydown({ key: name, preventDefault() {} });

render();
assert.equal(value().attrs.contenteditable, 'plaintext-only');
assert.equal(value().attrs['aria-label'], 'Spill over for Vasteras');
assert.equal(container.children[0].children[1].children[0].children[1].attrs.contenteditable, undefined);
for (const data of ['a', '-', '.', '1e2', '12abc', '\n']) {
  let prevented = false;
  value().listeners.beforeinput({ data, preventDefault() { prevented = true; } });
  assert.equal(prevented, true);
}
for (const data of ['123', null]) {
  value().listeners.beforeinput({ data, preventDefault() { assert.fail('Digits and deletions must be allowed'); } });
}
let pastePrevented = false;
value().listeners.paste({ clipboardData: { getData: () => '12abc' }, preventDefault() { pastePrevented = true; } });
assert.equal(pastePrevented, true);
value().listeners.paste({ clipboardData: { getData: () => '123' }, preventDefault() { assert.fail('Numeric paste must be allowed'); } });
value().textContent = '1a2.3';
value().listeners.input();
assert.equal(value().textContent, '123');
metric().listeners.click();
assert.equal(value().focused, true);
value().textContent = '12';
key('Enter');
assert.equal(record.spillover, 12);
assert.ok(!metric().className.includes('burndown-metric-zero'));
render();
assert.equal(value().textContent, '12');
value().textContent = '99';
key('Escape');
assert.equal(record.spillover, 12);
assert.equal(value().textContent, '12');
for (const invalid of ['', '-1', '1.5', 'abc', '9007199254740992']) {
  value().textContent = invalid;
  value().blur();
  assert.equal(record.spillover, 12);
  assert.equal(value().textContent, '12');
}
value().textContent = ' 0 ';
value().blur();
assert.equal(record.spillover, 0);
assert.ok(metric().className.includes('burndown-metric-zero'));
console.log('Passed: Spill over editing, click focus, save, cancel, validation, zero styling and retention on rerender.');

const storyRow = () => container.children[0].children[3].children[2];
const storyIds = () => storyRow().children[1];
const storyKey = key => storyIds().listeners.keydown({ key, preventDefault() {} });
assert.equal(storyIds().attrs.contenteditable, 'plaintext-only');
assert.equal(storyIds().textContent, 'NA');
assert.equal(container.children[0].children[3].children[0].children[1].attrs.contenteditable, undefined);
storyRow().listeners.click();
assert.equal(storyIds().focused, true);
storyIds().listeners.focus();
assert.equal(storyIds().textContent, '');
storyIds().textContent = 'US 123, Bug 456; US 789 • US 999\nUS 111';
storyKey('Enter');
assert.deepEqual(Array.from(record.stories.spillover), ['US 123', 'Bug 456', 'US 789', 'US 999', 'US 111']);
assert.equal(storyIds().children[1].textContent, ' • ');
assert.equal(record.spillover, 0);
render();
assert.equal(storyIds().children[0].textContent, 'US 123');
storyIds().textContent = 'US 222';
storyKey('Escape');
assert.equal(record.stories.spillover[0], 'US 123');
assert.equal(storyIds().children[0].textContent, 'US 123');
storyIds().textContent = '<b>US 333</b>';
storyIds().blur();
assert.equal(storyIds().children[0].textContent, '<b>US 333</b>');
for (const empty of ['', 'NA']) {
  storyIds().textContent = empty;
  storyIds().blur();
  assert.equal(record.stories.spillover.length, 0);
  assert.equal(storyIds().textContent, 'NA');
  assert.equal(storyIds().children.length, 0);
}
console.log('Passed: Spillover story editing, separators, save, cancel, safe text, clearing and retention on rerender.');
