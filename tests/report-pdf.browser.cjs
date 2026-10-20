const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const fs = require('node:fs');
const assert = require('node:assert/strict');
const path = require('node:path');
const XLSX = require('../../sprint-dashboard/vendor/xlsx.full.min.js');
(async () => {
  fs.mkdirSync('tmp/pdfs', { recursive: true });
  const browser = await chromium.launch({ channel: 'chrome', headless: true });
  try {
    const page = await browser.newPage({ viewport: { width: 1440, height: 1000 } });
    const errors = [];
    page.on('pageerror', error => errors.push(error.message));
    await page.goto('file:///' + path.resolve('sprint-dashboard/index.html').replaceAll('\\', '/'));
    assert.deepEqual(await page.evaluate(() => [
      formatSprintReportCoverDate(new Date('2026-10-16T18:29:59Z')),
      formatSprintReportCoverDate(new Date('2026-10-16T18:30:00Z')),
      formatSprintReportCoverDate(new Date('2027-01-01T12:00:00Z'))
    ]), ['16 October 2026', '17 October 2026', '01 January 2027']);
    let warnings = 0;
    page.on('dialog', async dialog => { assert.equal(dialog.message(), 'Upload Excel file to create report'); warnings++; await dialog.accept(); });
    await page.click('#create-report-button');
    assert.equal(warnings, 1);
    const sites = ['Vasteras', 'Rakovski', 'Heidelberg', 'Senatobia', 'Alingsas', 'Chassieu'];
    const book = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([
      ['TOTAL SITES', 'PLANNED HRS', 'COMPLETED HRS', 'OVERALL ACTUAL %', 'DELAYED / AT RISK'],
      [6, 600, 420, 70, 2], [],
      ['Site', 'Sprint', 'Start', 'End', 'Planned Hrs', 'Actual Hrs', 'Planned %', 'Actual %', 'Variance', 'Delay days', 'Health', 'Blocked', 'Comments'],
      ...sites.map((site, i) => [site, 20, '2026-10-01', '2026-10-15', 100, 70+i, 90, 70+i, -20+i, 2, i % 2 ? 'Green' : 'Amber', 0, 'Sprint delivery review'])
    ]), 'Dashboard');
    XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([
      ['ID', 'Area Path', 'Date', 'Work Item Type', 'Assigned To', 'State', 'Tags', 'Original Estimate', 'Remaining Work', 'Completed Work', 'Parent'],
      ...sites.flatMap((site, i) => [
        [300+i, site, '2026-10-01', 'Task', 'Developer '+site, 'Active', '', 10, 10, 0, 100+i],
        [100+i, site, '2026-10-15', 'User Story', 'Developer '+site, 'Closed', 'Committed', 10, 0, 10, ''],
        [200+i, site, '2026-10-15', 'Bug', 'Developer '+site, 'Active', 'Committed;Tentative', 10, 5, 5, ''],
        [300+i, site, '2026-10-15', 'Task', 'Developer '+site, 'Active', '', 10, 5, 5, 100+i]
      ])
    ]), 'Master Daily Dump');
    const file = { name: 'report-fixture.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: XLSX.write(book, { type: 'buffer', bookType: 'xlsx' }) };
    fs.writeFileSync('tmp/pdfs/report-fixture.xlsx', file.buffer);
    await page.setInputFiles('#site-excel-file', path.resolve('tmp/pdfs/report-fixture.xlsx'));
    await page.click('#create-report-button');
    assert.equal(warnings, 2, 'Selecting a file must not pass the upload check');
    await page.click('#upload-excel-button');
    await page.waitForFunction(() => document.getElementById('upload-status').textContent.includes('Loaded'));
    await page.click('#setting-nav');
    await page.fill('#dashboard-date-from', '2026-10-01');
    await page.dispatchEvent('#dashboard-date-from', 'change');
    await page.fill('#dashboard-date-to', '2026-10-15');
    await page.dispatchEvent('#dashboard-date-to', 'change');
    await page.click('#dashboard-nav');
    await page.locator('#sprint-rows textarea').first().fill('Edited comment included in the report.');
    await page.evaluate(() => {
      const original = window.html2canvas;
      window.html2canvas = async (...args) => {
        const canvas = await original(...args);
        (window.reportPreviews ||= []).push(canvas.toDataURL('image/png'));
        return canvas;
      };
    });
    const download = page.waitForEvent('download', { timeout: 60000 });
    await page.click('#create-report-button');
    const report = await download.catch(async error => { console.log(await page.locator('#report-status').textContent(), errors); throw error; });
    await report.saveAs('tmp/pdfs/report.pdf');
    await page.waitForFunction(() => !document.getElementById('create-report-button').disabled);
    assert.equal(await page.locator('.pdf-report-host').count(), 0);
    assert.match(report.suggestedFilename(), /2026-10-01_to_2026-10-15\.pdf$/);
    assert.equal((fs.readFileSync('tmp/pdfs/report.pdf').toString('latin1').match(/\/Type \/Page\b/g) || []).length, 7);
    const chartPagination = await page.evaluate(() => {
      const snapshot = buildSprintReportSnapshot();
      try {
        const pages = paginateReportBurndown(snapshot);
        const charts = pages.flatMap(p => [...p.querySelectorAll('.burndown-site-chart-card')]);
        const result = {
          pages: pages.length, charts: charts.length, svg: charts.filter(c => c.querySelector('svg')).length,
          headings: pages.map(p => p.dataset.reportTitle),
          overflow: pages.some(p => p.firstElementChild.getBoundingClientRect().height > 638)
        };
        pages.forEach(p => p.remove());
        snapshot.burndown.replaceChildren(snapshot.burndown.firstElementChild);
        const filtered = paginateReportBurndown(snapshot);
        result.filteredCount = filtered.flatMap(p => [...p.querySelectorAll('.burndown-site-chart-card')]).length;
        filtered.forEach(p => p.remove());
        snapshot.burndown.replaceChildren();
        snapshot.burndownMessage = 'No sites selected with observations.';
        const empty = paginateReportBurndown(snapshot);
        result.emptyMessage = empty.length === 1 && empty[0].textContent.includes(snapshot.burndownMessage);
        return result;
      } finally { snapshot.host.remove(); }
    });
    assert.deepEqual(chartPagination, { pages: 3, charts: 6, svg: 6, headings: Array(3).fill('Sprint burndown chart report'), overflow: false, filteredCount: 1, emptyMessage: true });
    const previews = await page.evaluate(() => window.reportPreviews);
    previews.forEach((data, i) => fs.writeFileSync(`tmp/pdfs/content-${i+1}.png`, Buffer.from(data.split(',')[1], 'base64')));
    const pagination = await page.evaluate(() => {
      const container = document.getElementById('user-story-cards');
      container.replaceChildren(...['Alingsas', 'Chassieu', 'Heidelberg', 'Rakovski', 'Senatobia', 'Vasteras'].map((site, siteIndex) => renderUserStorySiteCard({
        site, dates: { sprintStartDate: '2026-10-01', sprintEndDate: '2026-10-15' },
        stories: Array.from({length: siteIndex === 0 ? 1 : 7}, (_, storyIndex) => ({ id: `${siteIndex}-${storyIndex}`, tasks: Array.from({length: siteIndex === 0 ? 65 : 3}, (_, taskIndex) => ({ id: `${siteIndex}-${storyIndex}-${taskIndex}`, completed: 5, remaining: 5, estimate: 10 })) }))
      })));
      const snapshot = buildSprintReportSnapshot();
      try {
        const pages = paginateReportStories(snapshot);
        const labels = pages.flatMap(p => [...p.querySelectorAll('.user-story-task .story-progress-label')].map(x => x.textContent));
        const expected = [...container.querySelectorAll('.user-story-task .story-progress-label')].map(x => x.textContent);
        const overflow = pages.some(p => [...p.querySelectorAll('.pdf-story-column')].some(c => c.lastElementChild && c.lastElementChild.getBoundingClientRect().bottom - c.getBoundingClientRect().top > 638));
        return { count: pages.length, labels, expected, overflow, continuation: pages.some(p => p.textContent.includes('continued')) };
      } finally { snapshot.host.remove(); }
    });
    assert(pagination.count > 1);
    assert.deepEqual(pagination.labels, pagination.expected, 'Every task must appear once, in order');
    assert.equal(pagination.overflow, false, 'No card may overflow the content area');
    assert.equal(pagination.continuation, true);
    await page.setViewportSize({ width: 390, height: 844 });
    if (!await page.locator('body').evaluate(el => el.classList.contains('nav-collapsed'))) await page.click('#menu-toggle');
    const mobileDownload = page.waitForEvent('download', { timeout: 60000 });
    await page.click('#create-report-button');
    await (await mobileDownload).saveAs('tmp/pdfs/report-mobile.pdf');
    await page.waitForFunction(() => !document.getElementById('create-report-button').disabled);
    await page.evaluate(() => { window.html2canvas = () => Promise.reject(new Error('Simulated capture failure')); });
    await page.click('#create-report-button');
    await page.waitForFunction(() => document.getElementById('report-status').textContent.includes('Simulated capture failure'));
    assert.equal(await page.locator('.pdf-report-host').count(), 0);
    assert.equal(await page.locator('#create-report-button').isEnabled(), true);
    assert.deepEqual(errors, []);
    assert.equal((fs.readFileSync('tmp/pdfs/report-mobile.pdf').toString('latin1').match(/\/Type \/Page\b/g) || []).length, pagination.count + 6);
    console.log('PASS: dynamic India-time cover date; upload guards; desktop seven-page report; complete story/task pagination; six burndown charts on three pages; filtered and empty chart selections; no overflow; mobile download; failure recovery.');
  } finally { await browser.close(); }
})().catch(error => { console.error(error); process.exitCode = 1; });

