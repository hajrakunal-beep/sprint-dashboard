const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { jsPDF } = require('../vendor/jspdf.umd.min.js');

const context = vm.createContext({ Intl, Date });
for (const name of ['report-footer.js', 'report-pdf.js']) {
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', name), 'utf8'), context);
}
const pdf = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4', compress: false });
context.addSprintReportBranding(pdf, 'Sprint health Status', 2);
const result = pdf.output();
// PDF painting operators consume the current path. W must act on the rectangle
// before it is painted, otherwise strict viewers receive an empty clipping path.
assert.doesNotMatch(result, /re\s+S\s+W\s+n/, 'Footer must not clip an already-consumed path');
assert.match(result, /re\s+W\s+n/, 'Footer must clip the rectangle before drawing its image');
assert.match(result, /\/I0 Do\s+Q\s+Q/, 'Footer clipping must be restored before later content');
console.log('PASS: real jsPDF footer has a valid clipping path and restores graphics state.');
