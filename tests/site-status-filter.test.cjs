const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

class Element {
  constructor(tag = 'div') { this.tag = tag; this.children = []; this.listeners = {}; this.style = {}; }
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
  excelInput: get('file'), uploadStatus: {},
  sprintSiteKey(value) {
    return String(value).toLowerCase().replace(/^mes[ _-]*(se[ _-]*|fr[ _-]*|de[ _-]*)?/, '');
  },
  svgElement: tag => new Element(tag)
});
vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'site-filter.js'), 'utf8'), context);
vm.runInContext(fs.readFileSync(path.join(__dirname, '../site-status.js'), 'utf8'), context);
context.initializeStatusSiteFilter();
const options = get('site-status-site-options');
const filter = get('site-status-site-filter');
filter.append(new Element('summary'), options);
assert.equal(options.children.length, 7);
assert.ok(options.children.every(label => label.children[1].checked));
const item = site => ({ site, id: '1', workItemType: 'Task', developer: 'Dev', state: 'Active', estimate: 10, completed: 5 });
const data = { items: [item('Project\\MES Alingsas'), item('MES Alingsas'), item('MES SE Vasteras'), item('Other site')] };
context.renderSiteData(data);
assert.equal(options.children.length, 8); // Include additional workbook sites.
assert.equal(get('site-results-grid').children.length, 4);
const toggle = (index, checked) => {
  const input = options.children[index].children[1];
  input.checked = checked;
  input.listeners.change();
};
toggle(1, false);
assert.equal(get('site-results-grid').children.length, 2); // Both Alingsas paths hidden.
assert.equal(get('site-status-site-selection').textContent, '6 selected');
assert.equal(options.children[0].children[1].indeterminate, true);
context.renderSiteData(data);
assert.equal(get('site-results-grid').children.length, 2);
toggle(1, true);
assert.equal(get('site-results-grid').children.length, 4);
toggle(0, false);
assert.match(get('site-results-grid').children[0].textContent, /No sites/);
assert.equal(get('site-status-site-selection').textContent, 'None');
toggle(2, true);
assert.equal(get('site-results-grid').children.length, 1);
assert.equal(get('site-results-grid').children[0].children[0].children[1].textContent, 'Vasteras');
toggle(0, true);
assert.equal(get('site-results-grid').children.length, 4);
filter.open = true;
listeners.click({ target: options.children[1].children[1] });
assert.equal(filter.open, true);
filter.listeners.keydown({ key: 'Escape' });
assert.equal(filter.open, false);
assert.equal(filter.querySelector('summary').focused, true);
filter.open = true;
listeners.click({ target: new Element() });
assert.equal(filter.open, false);
console.log('Passed: Site wise status defaults, alias filtering, extra sites, refresh, select all, empty selection and dropdown dismissal.');
