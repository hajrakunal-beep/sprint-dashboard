// Render the actual PDF with PDF.js: checking the captured HTML alone cannot
// catch PDF clipping errors that hide the body in browser PDF viewers.
const fs = require('node:fs');
const path = require('node:path');
const { pathToFileURL } = require('node:url');
const assert = require('node:assert/strict');

(async () => {
  const canvasModule = require(process.env.PDF_CANVAS_MODULE || '@napi-rs/canvas');
  for (const name of ['DOMMatrix', 'ImageData', 'Path2D']) globalThis[name] = canvasModule[name];
  const modulePath = require.resolve(process.env.PDFJS_MODULE || 'pdfjs-dist/legacy/build/pdf.mjs');
  const pdfjs = await import(pathToFileURL(modulePath).href);
  const input = process.argv[2] || 'tmp/pdfs/report.pdf';
  const output = process.argv[3] || 'tmp/pdfs/rendered';
  fs.mkdirSync(output, { recursive: true });
  const document = await pdfjs.getDocument({
    data: new Uint8Array(fs.readFileSync(input)),
    standardFontDataUrl: path.resolve(path.dirname(modulePath), '../../standard_fonts').replaceAll('\\', '/') + '/'
  }).promise;
  const blank = [];
  for (let number = 1; number <= document.numPages; number++) {
    const page = await document.getPage(number);
    const viewport = page.getViewport({ scale: 1.5 });
    const canvas = canvasModule.createCanvas(Math.ceil(viewport.width), Math.ceil(viewport.height));
    const context = canvas.getContext('2d');
    await page.render({ canvasContext: context, viewport }).promise;
    fs.writeFileSync(path.join(output, `page-${number}.png`), canvas.toBuffer('image/png'));
    if (number === 1) continue;
    // Exclude the heading and footer; these remained visible in the regression.
    const pixels = context.getImageData(Math.floor(canvas.width * .05), Math.floor(canvas.height * .14), Math.floor(canvas.width * .9), Math.floor(canvas.height * .72)).data;
    let ink = 0;
    for (let i = 0; i < pixels.length; i += 4) if (Math.min(pixels[i], pixels[i+1], pixels[i+2]) < 225) ink++;
    const fraction = ink / (pixels.length / 4);
    console.log(`Page ${number}: ${(fraction * 100).toFixed(2)}% visible body content`);
    if (fraction < .001) blank.push(number);
  }
  await document.destroy();
  assert.deepEqual(blank, [], 'PDF pages must contain visible content below the heading');
  console.log('PASS: every report page is visible in PDF.js.');
})().catch(error => { console.error(error); process.exitCode = 1; });
