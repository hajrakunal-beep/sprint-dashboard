# Sprint Health Dashboard

A simple offline website using only HTML, CSS, and JavaScript. No installation,
Node.js, build process, chart library, or internet connection is required.

## Update Consolidated Backlog

In Settings, select an `.xlsx` workbook in **Update Consolidated Backlog** and
click **Update Backlog**. Allow the browser's write-access prompt. Updates are saved
directly to the selected workbook in its existing location; no download or new file
is created. Each update reads the latest saved workbook. Use desktop Chrome or Edge
with HTTPS or localhost. For SharePoint, select the OneDrive-synced workbook and
allow syncing to finish. Close the workbook in Excel before updating.

Required worksheet headers (case and surrounding whitespace are ignored):

- **Dashboard:** Site, Sprint, Start, End.
- **Consolidated Backlog:** Site, ID, Sprint No, User story / task, Estm. Hrs,
  Com. Hrs, Rem. Hrs.
- **Master Daily Dump:** ID, Area Path, Date, Work Item Type, Parent,
  Original Estimate, Remaining Work, Completed Work.

For each User Story or Bug backlog row, match Dashboard by Site and Sprint No
against Site and Sprint. Find Task dump rows whose Parent matches the backlog ID,
whose Area Path contains the site name (case-insensitive), and whose Date is within
Start and End inclusive. Select the latest matching calendar date separately for
each backlog row, and sum only the matching tasks on that date. Original Estimate
becomes Estm. Hrs, Completed Work becomes Com. Hrs, and Remaining Work becomes
Rem. Hrs. No matching tasks or Dashboard range results in three zeroes. Other
backlog row types are unchanged. Frosinone is excluded (ignoring case and surrounding
spaces): its backlog hours remain unchanged and its Dashboard dates are not checked.
These output cells are replaced with numeric
totals, including any existing formulas in those cells.

Site matching recognizes the dashboard's known MES aliases (for example,
MES Alingsas and Alingsas). The update status explains why a row had no matches:
missing Dashboard site/sprint, no matching Task Parent, Area Path mismatch,
unreadable task dates, or dates outside the Dashboard range.

Excel date cells and ISO, day/month/year, and day-month-name-year text dates are
supported; numeric text dates use day-first order. Invalid task dates are skipped.
Invalid or conflicting Dashboard ranges and missing/duplicate required headers
stop the update. Blank task hours count as zero; invalid hours on the selected
date stop the update. Repeated Task IDs on the selected date with identical hours
are counted once. Conflicting hours for the same Task ID stop the update and report
the source row numbers. Rows without a Task ID remain separate.
Only the Consolidated Backlog worksheet XML is patched, retaining existing cell
styles and other workbook parts. Pivot caches and formula results are not refreshed.

Run `node sprint-dashboard/tests/update-backlog.test.cjs` from the workspace root
to verify calculations, workbook preservation, and button behavior.

## Run locally in VS Code

1. Extract sprint-dashboard.zip.
2. In VS Code choose File > Open Folder and select the sprint-dashboard folder.
3. Open index.html in your browser by double-clicking it in File Explorer.
4. Edit the files in VS Code, save, and refresh the browser.

Optional: if you already use the VS Code Live Server extension, right-click
index.html and choose Open with Live Server for automatic refresh.

Keep index.html, style.css, and script.js together in the same folder.
Keep dashboard-sheet.js, site-filter.js, site-status.js and the vendor folder alongside them for Excel imports.
Keep merge-data.js, preserve-workbook.js and settings.js alongside them for Settings uploads and merging.

## Merge daily dumps in Settings

Choose the master workbook on the left and the ADO daily dump on the right
using an `.xlsx` master and an `.xlsx`, `.xls` or `.csv` ADO dump, then click **Merge Data**. The master must contain
a sheet named **Master Daily Dump**. Both sheets must contain these headers:

ID, Site, Date, Work Item Type, Title 1, Title 2, Parent, Assigned To, State,
Original Estimate, Remaining Work, Completed Work.

Tag values are also copied when the ADO dump includes a **Tag** or **Tags** column.
The master must have a **Tag** column (or **Tags**) on the same header row; its
position can differ from the ADO column. Tag text is preserved as supplied.
Files without a tag column can still be merged. If the ADO file has tags but the
master has no tag column, merging stops with instructions to add one.

CSV dumps should use UTF-8 encoding and the same headers. Quoted commas and
multiline fields are supported. IDs retain leading zeroes and date text is kept
as supplied; numeric work/estimate columns are imported as numbers.

Columns are matched by name, ignoring case and surrounding whitespace, so
column order can differ. The ADO sheet is detected by its headers; if multiple
sheets match, name the intended one **Master Daily Dump**.
Non-empty ADO rows are appended after the last populated master row. Existing
rows are retained, blank rows are skipped, and duplicate IDs are appended too.
ADO formula cells contribute their saved values; existing master formulas and
other worksheet data are retained. Excel date values and number formats are copied.

The selected master file is updated directly in its existing location, retaining
its `.xlsx` format. No new file or download is created. Allow the browser's
write-access prompt when merging. Each merge reads the latest master file, so the
next daily dump appends after previous merges. Clicking Merge again with the same
ADO dump appends those rows again. The ADO file is unchanged.
The ADO dump is read when selected. If you edit or replace the dump afterwards,
select it again to load the updated data. If merging reports a changed or unreadable
file, save and close both workbooks in Excel, wait for OneDrive/SharePoint or other
file sync to finish, then select both files again and retry. The master read is
retried once with a fresh file reference; saves are never automatically retried.
Direct saving requires File System Access support: use desktop Chrome or Edge
with HTTPS or a localhost server (such as VS Code Live Server). Unsupported
browsers show a message and disable merging. Files are processed in the browser.
The save patches the original XLSX ZIP package instead of exporting the whole
workbook through the Excel library. Other worksheets and unrelated package parts
(styles, dropdowns, conditional formatting, charts, pivot definitions, etc.) are
retained byte for byte. Existing Master Daily Dump cells are retained, appended
cells use existing master styles, and rules ending at the previous final data row
extend to the new rows. Target sheet filters and tables extend as well. Tables
with a totals row must have that row removed before merging. Existing pivot caches
and formula cached results are preserved; refresh/recalculate in Excel when needed.
Legacy `.xls` masters are blocked because this preservation method requires XLSX.
Save the original master as `.xlsx` in Excel first. Formatting lost by earlier
merges cannot be reconstructed; restore the original workbook or a backup.

Run merge checks with `node tests/merge-data.test.cjs` and dashboard checks with
`node tests/dashboard-sheet.test.cjs`.
Run package preservation checks with `node tests/preserve-workbook.test.cjs`.

## Dashboard Excel data

### Sprint Commitment Committed, Tentative and Delivered counts

All six sites use the rules below, independently per site.
Rakovski also accepts `MES-Rakovski`; Heidelberg also accepts `MES_DE_Heidelberg`.
Senatobia also accepts `MES-Senatobia`; Alingsas also accepts `MES Alingsas`.
Chassieu also accepts `MES FR Chassieu`. All six cards use the selected date range,
refresh on date changes, and clear previous data when a replacement file is selected.

Dashboard uploads also populate the first Sprint Commitment card's **Committed**
and **Tentative** values from **Master Daily Dump**, plus **Delivered** based on State.
Required headers are ID, Site, Date, Work Item
Type, Assigned To, State and Tag, in any order. Count each row whose Site is
Vestaras (also accepting the existing Vasteras and MES SE Vasteras names), Work
Item Type is User Story or Bug, and Tag contains Committed or Tentative as a
semicolon-separated tag into the corresponding count. Matching ignores case
and extra whitespace. For example, `Priority; Committed; Team` counts as Committed.
A row containing both Committed and Tentative counts once in each category;
repeating a tag in one cell does not increase the count. Tags such as
`Not Committed` do not match. Committed and Tentative use only each site's latest
available Date within that site's Dashboard date range. Earlier snapshots are excluded; sites
may use different latest dates. These remain row counts on that date, not unique-ID counts.
On each page load, Settings defaults to From = 15 days ago and To = today,
using the browser's local date. Manual date changes remain selected during
the session, including after Excel uploads. Reloading starts with fresh defaults.
All six Sprint Commitment cards match their site to the Dashboard worksheet and
use that row's Start and End dates, including both boundary dates. This range
controls counts, completion, assignees and ID lists. Settings dates do not
change these cards' data. Excel date cells (including 1904 workbooks),
YYYY-MM-DD dates/timestamps and DD-Mmm-YYYY text dates are supported in Master
Daily Dump. Blank/invalid row dates are skipped and reported. Missing, invalid,
reversed or conflicting Dashboard ranges make that site's counts unavailable;
other sites with valid ranges continue to load.
Delivered counts User Story and Bug rows for the same site whose State is Closed,
Resolved or Done, regardless of Tag, only on that site's latest available date
within that site's Dashboard range, matching Committed and Tentative. It excludes Tasks.
Each card's completion percentage is `Math.round(Delivered / Committed * 100)`.
It displays N/A when Committed is zero or either count is unavailable. Results
above 100% are retained when Delivered exceeds Committed.
These fields show zero for no matches and a dash before upload or when required
data is unavailable. Other card fields retain their current reference values.
All Sprint Commitment card names are populated from Assigned To rows for the
corresponding site. The nine configured assignee aliases display Diogo, Faiz,
Joao, Mangesh, Nivruth, Ricardo, Shaalini, Suman and Vinod. Multiple assignees
are shown once each, separated by commas. Unmapped names are kept as supplied;
cards without an assignee show Unassigned. Replacement uploads clear old names.
The Vasteras card's In progress detail lists IDs of User Story and Bug rows
whose State is In Progress or Active, irrespective of Tag. Prefixes are `US` and `Bug`.
Repeated IDs of the same type appear once; blank IDs are skipped. When no rows
match (or no valid workbook is loaded), the detail shows NA.
The Resolved detail lists unique User Story and Bug IDs with the same prefixes
when State is Resolved, Completed, Done or Closed, irrespective of Tag. It shows NA
when no matching IDs are available. This detail uses its own state rule; the
Delivered count continues to use Closed, Resolved or Done.
The Tentative detail lists unique User Story and Bug IDs whose tags include Tentative,
with `US` / `Bug` prefixes and dot separators, on the same latest date used for its count.
The top Tentative count uses the same semicolon-separated tag matching.
Run `node tests/sprint-committed.test.cjs` to check this upload flow.
Run `node tests/sprint-latest-counts.test.cjs` to check the latest-date count rule.

Select an `.xlsx` or `.xls` file and click **Upload**. The **Dashboard** worksheet
populates every non-empty row in **Sprint delivery overview** and updates the
planned versus actual chart. Required headers (case and extra spaces are ignored):
Site, Sprint, Start, End, Planned Hrs, Actual Hrs, Planned %, Actual %, Variance,
Delay, Health, Blocked. Rows above the headers and blank rows are skipped.

Percentages accept Excel percentage cells, text such as `20%`, or plain numbers
on a 0–100 scale (`20` means 20%). Blank percentages show a dash. Progress bars
are limited to 0–100%, while labels retain the supplied value. Variance comes
directly from Excel. Excel dates display as `14-Sep-2026`; text dates are kept.
Health names Green, Amber and Red retain the existing colors, regardless of case;
other valid CSS color names use their named color. Unrecognized health values
use gray. The reporting period remains independently configured.

The summary card reads **TOTAL SITES**, **PLANNED HRS**, **COMPLETED HRS**,
**OVERALL ACTUAL %**, and **DELAYED / AT RISK** as summary headers/labels in
Dashboard. Each label can have its value to the right (including gaps for merged
cells) or directly below it. For example, `TOTAL SITES | 7` and `OVERALL ACTUAL % | 34%`
can appear on separate rows. A horizontal header row with values underneath is
also supported. Put the summary section above or below the site table, or include
its headers alongside the site headers. When sharing the
site header row, PLANNED HRS uses that row's existing Planned Hrs column.
Overall Actual % displays as one summary value and supports the same percentage
formats as the chart. Matching ignores case and repeated spaces in the header.
Totals are
imported directly from Excel. Missing or blank summary values display a dash;
upload status reports missing summary headers. As Of continues to use today's date.

A workbook containing only Dashboard can populate the overview. Include
**Master Daily Dump** to also populate Site wise status and Burndown chart.
Upload status reports missing sheets or invalid columns.

## Site wise status

Site wise status matches each site's Area Path to Site in the Dashboard
worksheet and filters Master Daily Dump rows using that site's Start and End,
including both endpoints. It displays the latest available date per site
within that range. Blank or invalid row dates are excluded. Settings dates
do not affect this tab. Missing, invalid, reversed or conflicting Dashboard
ranges exclude the affected site's data and display a warning; valid sites
continue to load. Dashboard requires Site, Sprint, Start and End headers.

On **Dashboard**, select an `.xlsx` or `.xls` workbook in the compact upload strip
to the right of the date fields, then click **Upload**. Open **Site wise status**
to view the results. Selecting a file alone keeps results hidden. Both tabs read
the worksheet named **Master Daily Dump** for work item data. A missing sheet is
reported without preventing valid Dashboard data from loading. If no Dashboard data
can load either, a missing Master Daily Dump sheet triggers a warning popup.
This sheet must contain these headers (matching ignores
case and extra spaces): ID, Area Path, Date, Work Item Type, Assigned To, State, Original
Estimate, Remaining Work, and Completed Work. Rows above the headers are ignored.

Only rows with **Work Item Type = Task** are charted in Site wise status.
Tasks are grouped by Area Path, with one full-width card per row.
Only rows for each site's most recent date within its Dashboard range are shown.
Excel date cells, YYYY-MM-DD text, named month dates and day-month-year text
are supported. Undated rows are excluded.
The full workbook history remains available to Burndown chart.
Sites appear in their latest rows' workbook order. Matching ignores case and extra
spaces; blank sites appear as Unspecified site. Each card shows only that site's
Task IDs, its developer names from Assigned To in the header, and the Excel site
name at the top right. Each Task ID has two equal-width vertical bars:
grey for Original Estimate and green for Completed Work, side by side and
starting at zero. Hour values appear at the tops of the bars. Completed hours
above the estimate remain visible. Each Task ID also shows its State from Excel
and completion percentage (Completed Work / Original Estimate × 100).
Below the completion percentage, **User Story** displays the task's **Parent**
column value. A missing Parent column or blank value displays N/A. User Story
and other work item rows remain available to the Burndown chart, which continues
to aggregate all work item types.
Completion shows N/A when either hour value is blank or the estimate is zero;
values above 100% remain visible. Task IDs run along the horizontal axis and
hours along the vertical axis. Cards fit the available page width. Bars stay
36 pixels wide with 160 pixels of spacing per task. When the chart exceeds the
card width, a horizontal scrollbar appears below it. The chart can also be
focused and scrolled with the keyboard. Blank hours remain blank; zero hours show a zero
label. Missing assignees appear as Unassigned.

Hours must be non-negative numbers. Empty workbooks, missing headers, invalid
hours, and unreadable files show an error. Selecting another workbook replaces
the results. Files are read locally in the browser and are not sent to a server.
Imports use the bundled SheetJS CE 0.20.3 reader; its license is in
vendor/SHEETJS-LICENSE.txt. No internet connection is needed.

## Sprint Commitment

Click **Sprint Commitment** below Dashboard to view the six site cards from the
reference image. Each card shows developer names, site, completion percentage,
Committed / Tentative / Delivered / Spill over counts, and user stories grouped
as In progress, Resolved, Spillover, and Tentative. Cards adapt from three columns
to two or one on smaller screens.

Dashboard uploads populate all six cards from Master Daily Dump. Each site's
Dashboard Start and End dates filter its data. Completion is calculated from Delivered / Committed,
rounded to a whole percent. Spillover remains at its reference value until its
rule is configured.

The date range beside each site comes from **Dashboard** in the same
Dashboard upload. The sheet must have `Site`, `Sprint`, `Start`, and `End`
headers; columns can be reordered and a report title can appear above them.
Site and header matching ignore case and surrounding whitespace. The matching
site's `Start` and `End` supply its displayed range. The site name and `Sprint`
value appear together in bold. Numeric sprint values display as `Sprint 20`;
values already starting with `Sprint` are displayed as supplied. Rows do not
need a `Current` value.

Excel date cells (including the 1904 date system) and text dates such as
`28-09-2026`, `28/09/2026`, `2026-09-28`, or `28-Sep-2026` are supported.
Ranges display as `[28-09-2026 to 09-10-2026]` independently of Settings dates.
Repeated site rows must share the same range. Missing rows, invalid dates,
reversed ranges, or conflicting ranges display `[Sprint dates unavailable]`,
with details beneath the tab heading. Selecting or uploading another file
clears previous ranges.

The renderer accepts records with this shape:

```js
{
  developer: "Developer name",
  site: "Site name",
  sprintStartDate: "2026-09-28", sprintEndDate: "2026-10-09",
  completionPercent: 50,
  committed: 4, tentative: 2, delivered: 2, spillover: 0,
  stories: {
    inProgress: ["US 1234"], resolved: ["US 1235"],
    spillover: [], tentative: ["US 1236"]
  }
}
```

Empty story groups display NA. An empty records array displays an empty state.
Keep `sprint-burndown.js` alongside the other dashboard files.

## Burndown chart

Each site uses its matching Dashboard row's Start and End dates, including
both endpoints. Only observations within that site's range contribute to work
totals, capacity and chart lines. The X-axis and ideal trend use the same site
range. Settings dates do not affect the charts. Blank or invalid observation
dates are excluded. Missing, invalid, reversed or conflicting Dashboard ranges
exclude the affected site's chart and display a warning; valid sites continue
to load.

Click **Burndown chart** below Site wise status. It uses the workbook uploaded
on the Dashboard tab, with no separate upload. Site charts appear side by side,
two per row on wide screens and one on small screens. History is read only from
the Date, Area Path and Remaining Work columns in **Master Daily Dump**.
Date cells can be native Excel dates or text in YYYY-MM-DD format. Provide at
least two dates, covering the sprint start and end. Example:

| Date | Area Path | Remaining Work |
| --- | --- | --- |
| 2026-10-01 | Vestaras | 40 |
| 2026-10-08 | Vestaras | 20 |
| 2026-10-15 | Vestaras | 0 |

Rows for the same site and date are summed (one daily observation per task).
Each site has a blue remaining-work area, an orange completed-work line,
and a solid grey ideal remaining trend to zero. Add **Completed Work** to the
dated history worksheet to display completed hours. For each site and date:

- Remaining hours = sum of Remaining Work across all work item types.
- Completed hours = sum of Completed Work across all work item types.

These are daily snapshot totals. Completed hours are not summed across dates,
because workbook values may already represent cumulative work per item.
Blank hours are excluded from sums; a date with no values for a series is
left blank for that series. Without Completed Work, the remaining series
still displays and a note explains that completed history is needed.
The X-axis uses evenly spaced working days; weekends are omitted unless they
have recorded work or mark a sprint boundary. The Y-axis shows work hours
and scales to include both series. Actual work ends at the latest
observation; missing dates and blank values are not filled with zero.
Each site needs at least two dates to display a chart.

The Dashboard Start and End dates control each site's chart. **Total Capacity** sets the ideal
line's starting total; otherwise the
first observation's remaining work is used. These values must be consistent
for each site. **Remaining Capacity** adds a dashed green line using recorded
daily capacity. Capacity is a site-wide value, not a per-task sum; repeat the
same capacity on all task rows for a date. All work and capacity values use hours.
Missing capacity is left blank. Cards show latest remaining work, completed
work, and the latest observation date, with two charts per row on desktop and
one on mobile.

Upload one sprint's history at a time. A task snapshot containing only ID,
State, Original Estimate, Remaining Work, and Completed Work cannot provide a
historical burndown without dated observations; such files show site cards with
a missing-history message. A new Dashboard upload replaces the charts.
Files are processed locally.
Keep `burndown-chart.js` alongside the other dashboard files.

## User Story Status

Site wise status also provides the same header-row site dropdown. All sites
are selected by default; check or uncheck a site to show or hide its cards.
Additional sites in the uploaded workbook are included in its options, and
selections remain in effect when workbook data or Settings dates refresh.
The Burndown chart tab has the same filter beside its heading, with independent
site selections that hide or restore each site's chart.

The User Story Status tab displays Alingsas, Vasteras, Rakovski, Heidelberg,
Senatobia and Chassieu. The top-center Sites dropdown selects all sites by
default. Uncheck a site's checkbox to hide its card and check it to restore it;
All sites selects or clears every site. Selections remain in effect when data
or Settings dates refresh, and reset to All when the page is reloaded. Upload a workbook
on Dashboard with a Master Daily Dump sheet containing Area Path, ID,
Work Item Type, Parent, Completed Work, Remaining Work and Original Estimate.
Each site has one compact card, with sprint dates at the top left and the site
name at the top right. All story trees appear sequentially inside the card,
with connected Task bars matched by Parent = story ID. The desktop grid has
three cards per row; each card grows automatically to fit its story trees,
without an internal scrollbar.
Dashboard supplies each site's Start and End dates for both its card header
and its data range. The shared reader matches Site and requires Site, Sprint,
Start and End headers. Master Daily Dump rows within that site's range
(including both endpoints) are considered. Blank or invalid row dates are
excluded. The latest row per ID within the range is used. Settings dates do
not affect User Story Status. Missing, invalid, reversed or conflicting
Dashboard dates show an unavailable range and an explanation on the affected
site's card; other sites continue to load.
User story percentages use the summed Completed Work and total hours of all
child tasks, applying the same denominator and fallback to each task before
summing. This weights progress by hours rather than averaging percentages.
Blank Completed Work counts as zero. Stories with no tasks or missing total
hours display an unavailable percentage.

Task percentage = Completed Work / (Completed Work + Remaining Work) × 100.
When Remaining Work is blank, Original Estimate supplies the denominator.
Blank Completed Work counts as zero; a missing or zero denominator displays an
unavailable percentage.
Bar percentages are capped at 100%. Light blue covers 0–75%, orange covers
above 75–85%, purple covers above 85% and below 100%, and green marks 100%.
Hover over a task bar to inspect the three imported hour values.

## Files

The Dashboard overview includes an editable **Comments** column after Blocked.
Each row has a compact, single-line-height text box. Entered comments have a
grey background and remain editable. The Planned % table bars are light blue.
An optional Comments column in the uploaded
Dashboard worksheet supplies initial text; older workbooks start with blank
comments. Edits remain in the current session and are cleared by a new upload
or page refresh. They are not saved back to the Excel workbook.

| File | Purpose |
| --- | --- |
| index.html | Top navigation, Dashboard sidebar, headings, table structure, chart container |
| style.css | Theme colors, sizes, spacing, responsive layout, table styles |
| script.js | Hardcoded data and functions that render summary, table, and SVG chart |

## Change the values

At the top of script.js, edit dashboardData. For example:

```js
summary: {
  totalSites: 7,
  plannedHours: 559,
  completedHours: 190,
  overallActual: 34.0,
  delayedAtRisk: 0
}
```

Edit reportingPeriod and asOf to change the report dates. These are display
strings, not automatically calculated dates.

Each object in sites creates one table row and one pair of chart bars:

```js
{
  site: "Vasteras",
  sprint: "Sprint 43",
  start: "14-Sep-2026",
  end: "25-Sep-2026",
  plannedHours: 45,
  actualHours: 136,
  plannedPercent: 20,
  actualPercent: 17.8,
  delay: 1,
  health: "Green",
  blocked: 0
}
```

- Use numbers without quotes for hours, percentages, and counts.
- Percentages use the range 0–100. Enter 20 for 20%.
- Add/remove an object in sites to add/remove a site in both table and chart.
- Separate objects with commas.
- Health can be Green, Amber, Red, or Not Configured. It is manually supplied.
- Variance = actualPercent − plannedPercent, in percentage points. Displayed
  with a % symbol to match the reference. Values of −5 or below appear red;
  change the condition in renderTable if your business rule differs.
- Percentages are independent inputs, not actualHours / plannedHours.
- Summary values intentionally remain independent. The reference has 559
  planned / 190 completed hours, while its rows sum to 559 planned / 761 actual
  hours. Do not infer new formulas without confirming the reporting definitions.
- Bar lengths use a fixed 0–100 scale; keep input percentages in that range.

## Change the design

Edit the :root variables at the beginning of style.css:

```css
--navy: #07133e;
--teal: #00899b;
--orange: #ff7f0e;
--sidebar-width: 200px;
--topbar-height: 62px;
```

Change the Sprint Monitor brand and Delivery management caption in index.html.
The menu button collapses or expands the sidebar. On narrow screens it starts
collapsed; the wide table and chart can scroll horizontally.

## Connect dynamic data later

The rendering logic already accepts a data object. Replace the last line,
renderDashboard(dashboardData), with your API loading code when a backend exists:

```js
async function loadDashboard() {
  try {
    const response = await fetch("/api/sprint-dashboard");
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const data = await response.json();
    renderDashboard(data);
  } catch (error) {
    console.error("Dashboard could not load:", error);
    alert("Unable to load dashboard data. Please try again.");
  }
}
loadDashboard();
```

This endpoint is a future example, not included in the project. Its JSON must
have the same structure as dashboardData. Serve the page over HTTP when using
an API and configure backend CORS if the API uses a different origin. Database
credentials should stay in the backend, never in browser JavaScript.

## Troubleshooting

- No styling: confirm the name is style.css and it sits beside index.html.
- Empty table/chart: check the browser developer console (F12) for a syntax error
  after editing script.js, particularly missing commas or quotes.
- Old values: save all files and refresh the browser (Ctrl+F5 if necessary).

The source files contain comments marking the main editing points.

## PDF reports

Upload an Excel workbook on Dashboard, then select **Create Report**. Selecting a file alone does not pass the upload check. The button shows progress and starts the browser download automatically when the PDF is ready; browser download preferences still apply.

The report begins with the supplied ABB ELSP Bi-Weekly cover, preserving its original proportions. The cover date is generated at download time in Asia/Kolkata time using `DD Month YYYY` (for example, `17 October 2026`), independently of the selected reporting period. The printed date was removed from `assets/report-cover.png`; `report-cover.js` embeds that background for offline use, and the PDF overlays the current date as selectable text. The cover retains its own logos and artwork; numbered report pages follow in A4 landscape:
- **Sprint health Status - Reporting period (date range)**: the current dashboard summary, delivery table (including edited comments), and progress chart.
- **Sprint Commitment**: all six site cards, arranged in two rows of three.
- **User Story Status**: the cards currently shown on the User Story Status page, respecting its selected sites. This section begins on page 4 and uses as many pages as necessary. Three compact columns retain a fixed readable text size. Cards stay intact where possible; large cards split between stories or tasks, repeating site headers and identifying continued stories. Empty selections or unavailable data show the page's status message.
- **Sprint burndown chart report**: appended after User Story Status, using the Burndown Chart page's heading and selected sites. Two columns keep chart axes, dates, metrics and legends readable, typically placing two charts on each page. Cards stay intact, and more pages are added as needed. Source information, warnings and empty-selection messages are included. Every continuation page repeats the section heading and branded footer.

Every report page after the cover uses the supplied transparent 2026 Cognizant/ABB footer banner across the full page width, with no gray background. Its red symbol contains the actual page number, counting the cover as page 1. The cover has no footer banner. The banner's transparent margins are clipped in the PDF without stretching or altering the source image, and its alpha channel is preserved. The reporting period matches the dashboard, and the filename includes the selected dates. Content is captured at double resolution. The dashboard and commitment pages scale to fit their margins; User Story Status and Burndown Chart add pages instead of shrinking their content. Page content is rendered as images, while headings and page numbers remain selectable text.

`report-pdf.js` creates an isolated snapshot and cleans it up on success or failure. PDF generation runs locally using bundled jsPDF 4.2.1 and html2canvas 1.4.1; no workbook data is sent to a service. `report-footer.js` embeds the original `assets/report-footer.png` so exports also work when `index.html` is opened directly. Vendor license files accompany both libraries.

Browser regression check: from the workspace root, run `node sprint-dashboard/tests/report-pdf.browser.cjs` with Playwright installed and Chrome available. If using an external Playwright installation, set `PLAYWRIGHT_MODULE` to its module path. The check generates sample files under `tmp/pdfs/`, verifies missing/selected-file warnings, uploads an Excel fixture, downloads desktop/mobile reports, and checks error cleanup and button recovery.

PDF viewer checks: `node sprint-dashboard/tests/report-footer.test.cjs` verifies that the footer uses a valid, unpainted clipping path and restores its graphics state. `node sprint-dashboard/tests/report-pdf-render.cjs path/to/report.pdf` renders the actual PDF using PDF.js and checks for visible content below every report heading. It requires `pdfjs-dist` and `@napi-rs/canvas`; external installations can be selected with `PDFJS_MODULE` and `PDF_CANVAS_MODULE`.
