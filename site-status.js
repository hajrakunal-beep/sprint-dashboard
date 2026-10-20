"use strict";

const requiredSiteColumns = ["ID", "Area Path", "Date", "Work Item Type", "Assigned To", "State", "Original Estimate", "Remaining Work", "Completed Work"];
const normalizeColumn = value => String(value ?? "").trim().replace(/\s+/g, " ").toLowerCase();
// Display aliases only: keep imported site and assignee values intact.
const siteDisplayNames = new Map([
  ["MES Alingsas", "Alingsas"],
  ["MES FR Chassieu", "Chassieu"],
  ["MES SE Vasteras", "Vasteras"],
  ["MES_DE_Heidelberg", "Heidelberg"],
  ["MES-Rakovski", "Rakovski"],
  ["MES-Senatobia", "Senatobia"]
].map(([source, label]) => [normalizeColumn(source), label]));
const assigneeDisplayNames = new Map([
  ["Diogo Ribeiro EXT <diogo.ribeiro@it.abb.com>", "Diogo"],
  ["Faiz Mohammed Siddiqui EXT <faiz-mohammed.siddiqui@it.abb.com>", "Faiz"],
  ["Joao Correia Rodrigues EXT <joao.correia-rodrigues@it.abb.com>", "Joao"],
  ["Mangesh Phatale EXT <mangesh.phatale@pt.abb.com>", "Mangesh"],
  ["Nivruth Gummadavelly EXT <nivruth.gummadavelly@pt.abb.com>", "Nivruth"],
  ["Ricardo Torres EXT <ricardo.torres@it.abb.com>", "Ricardo"],
  ["Shaalini Subu-Venkataraman EXT <shaalini.subu-venkataraman@it.abb.com>", "Shaalini"],
  ["Suman Mannepalli EXT <suman.mannepalli@pt.abb.com>", "Suman"],
  ["Vinod Kumar Kuchi EXT <vinod-kumar.kuchi@de.abb.com>", "Vinod"]
].map(([source, label]) => [normalizeColumn(source), label]));
const siteDisplayName = value => siteDisplayNames.get(normalizeColumn(value)) ?? value;
const assigneeDisplayName = value => assigneeDisplayNames.get(normalizeColumn(value)) ?? value;
const hoursLabel = value => `${Number(value.toFixed(2))} hrs`;
const completionLabel = item => item.estimate > 0 && item.completed !== null
  ? `${Number((item.completed / item.estimate * 100).toFixed(1))}%`
  : "N/A";

const hiddenStatusSites = new Set();
let displayedSiteData = null;
let statusSiteFilterInitialized = false;

function statusFilterSiteName(value) {
  const leaf = String(value).trim().split(/[\\/]/).pop();
  const key = sprintSiteKey(leaf);
  return [...siteDisplayNames.values()].find(name => sprintSiteKey(name) === key) || siteDisplayName(leaf);
}

function initializeStatusSiteFilter() {
  if (statusSiteFilterInitialized) return;
  statusSiteFilterInitialized = true;
  updateStatusSiteFilter();
  initializeSiteFilterDismissal("site-status");
}

function updateStatusSiteFilter() {
  if (!statusSiteFilterInitialized) return;
  const names = [...new Set([
    ...dashboardSiteNames,
    ...(displayedSiteData?.items || []).filter(item => normalizeColumn(item.workItemType) === "task").map(item => statusFilterSiteName(item.site))
  ])];
  renderSiteFilter("site-status", names, hiddenStatusSites, () => {
    if (displayedSiteData) renderFilteredSiteData();
  });
}

function readSiteWorksheet(workbook) {
  const datesBySite = readSprintDashboardDates(workbook);
  const sheetName = masterDailyDumpSheet(workbook);
  const matrix = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "", raw: true });
  const headerIndex = matrix.findIndex(row => requiredSiteColumns.every(column => row.some(cell => normalizeColumn(cell) === normalizeColumn(column))));
  if (headerIndex < 0) throw new Error(`Worksheet "Master Daily Dump" must have all the required columns: ${requiredSiteColumns.join(", ")}.`);
  const headers = matrix[headerIndex].map((cell, index) => String(cell).trim() || `Column ${index + 1}`);
  const columnIndex = name => headers.findIndex(header => normalizeColumn(header) === normalizeColumn(name));
  const hourValue = (value, name, rowNumber) => {
    if (String(value).trim() === "") return null;
    const number = typeof value === "number" ? value : Number(String(value).trim().replace(/,/g, ""));
    if (!Number.isFinite(number) || number < 0) throw new Error(`Excel row ${rowNumber}: ${name} must contain a non-negative number of hours.`);
    return number;
  };
  let rows = matrix.slice(headerIndex + 1).map((cells, index) => ({ cells, rowNumber: headerIndex + index + 2 }))
    .filter(({ cells }) => cells.some(cell => String(cell).trim() !== ""));
  if (!rows.length) throw new Error(`Worksheet "${sheetName}" has no work items.`);
  const dateIndex = columnIndex("Date");
  const warnings = [];
  const unavailableSites = new Set();
  const latestBySite = new Map();
  rows = rows.filter(row => {
    const area = String(row.cells[columnIndex("Area Path")] ?? "").trim().split(/[\\/]/).pop();
    row.siteKey = sprintSiteKey(area);
    const range = datesBySite.get(row.siteKey);
    if (!range) {
      if (!unavailableSites.has(row.siteKey)) {
        warnings.push(`${statusFilterSiteName(area)}: Dashboard Start and End dates are missing, invalid, reversed, or conflicting. Site data is unavailable.`);
        unavailableSites.add(row.siteKey);
      }
      return false;
    }
    const date = userStoryDate(row.cells[dateIndex], Boolean(workbook.Workbook?.WBProps?.date1904));
    if (!date) return false;
    const label = date.toISOString().slice(0, 10);
    if (label < range.sprintStartDate || label > range.sprintEndDate) return false;
    row.date = { stamp: date.getTime(), label };
    latestBySite.set(row.siteKey, Math.max(latestBySite.get(row.siteKey) ?? -Infinity, row.date.stamp));
    return true;
  });
  rows = rows.filter(row => row.date.stamp === latestBySite.get(row.siteKey));
  const items = rows.map(({ cells, rowNumber, date }) => {
    const id = String(cells[columnIndex("ID")] ?? "").trim();
    if (!id) throw new Error(`Excel row ${rowNumber}: ID is missing.`);
    // Empty hours have no bar or value label.
    return {
      id,
      workItemType: String(cells[columnIndex("Work Item Type")] ?? "").trim(),
      parent: String(cells[columnIndex("Parent")] ?? "").trim(),
      date: date?.label ?? null,
      site: String(cells[columnIndex("Area Path")] ?? "").trim() || "Unspecified site",
      state: String(cells[columnIndex("State")] ?? "").trim() || "Not specified",
      developer: String(cells[columnIndex("Assigned To")] ?? "").trim() || "Unassigned",
      estimate: hourValue(cells[columnIndex("Original Estimate")] ?? "", "Original Estimate", rowNumber),
      completed: hourValue(cells[columnIndex("Completed Work")] ?? "", "Completed Work", rowNumber)
    };
  });
  return { sheetName, headers, items, warnings };
}

function renderWorkItems(items, container) {
  items = items.filter(item => normalizeColumn(item.workItemType) === "task");
  container.replaceChildren();
  const maximum = items.reduce((max, item) => Math.max(max, item.estimate ?? 0, item.completed ?? 0), 1);
  const tickStep = Math.max(1, Math.ceil(maximum / 5));
  const axisMax = tickStep * 5;
  const left = 65, right = 25, top = 50, baseline = 310, plotHeight = baseline - top;
  // Keep task spacing and bar widths fixed; wide charts scroll within the card.
  const groupWidth = 160;
  const width = Math.max(480, left + items.length * groupWidth + right);
  const plotWidth = width - left - right;
  const chart = svgElement("svg", { viewBox: `0 0 ${width} 445`, class: "work-vertical-chart", role: "img", "aria-label": "Original estimate and completed work in hours for Task rows only. Grey estimate bars and green completed work bars appear side by side. Each task shows its Excel State, completion percentage, and User Story ID from the Parent column." });
  chart.style.width = `${width}px`;
  chart.style.minWidth = `${width}px`;
  chart.style.maxWidth = "none";
  chart.append(svgElement("text", { x: left, y: 22 }, "Hours"));
  for (let tick = 0; tick <= 5; tick++) {
    const y = baseline - tick / 5 * plotHeight;
    chart.append(svgElement("line", { x1: left, y1: y, x2: width - right, y2: y, stroke: "#e5e9ee" }));
    chart.append(svgElement("text", { x: left - 12, y: y + 4, "text-anchor": "end" }, tick * tickStep));
  }
  items.forEach((item, index) => {
    const center = left + (index + 0.5) * groupWidth;
    const group = svgElement("g", { class: "work-item-bars" });
    group.append(svgElement("title", {}, `ID ${item.id}: original estimate ${item.estimate === null ? "blank" : hoursLabel(item.estimate)}, completed work ${item.completed === null ? "blank" : hoursLabel(item.completed)}. Developer: ${assigneeDisplayName(item.developer)}. Status: ${item.state}. Completion: ${completionLabel(item)}. User Story: ${item.parent || "N/A"}.`));
    [[item.estimate, "#b8bfc7", -22, "end", "estimate-value"], [item.completed, "#319455", 22, "start", "completed-value"]].forEach(([value, color, barOffset, anchor, className]) => {
      if (value === null) return;
      const height = value / axisMax * plotHeight;
      const y = baseline - height;
      const barCenter = center + barOffset;
      group.append(svgElement("rect", { x: barCenter - 18, y, width: 36, height, rx: 2, fill: color }));
      group.append(svgElement("text", { x: center + (barOffset < 0 ? -4 : 4), y: y - 8, "text-anchor": anchor, class: className }, hoursLabel(value)));
    });
    group.append(svgElement("text", { x: center, y: baseline + 25, "text-anchor": "middle", class: "work-item-id" }, item.id));
    group.append(svgElement("text", { x: center, y: baseline + 47, "text-anchor": "middle", class: "work-item-state" }, `Status: ${item.state}`));
    group.append(svgElement("text", { x: center, y: baseline + 69, "text-anchor": "middle", class: "work-item-completion" }, `Completion: ${completionLabel(item)}`));
    group.append(svgElement("text", { x: center, y: baseline + 91, "text-anchor": "middle", class: "work-item-parent" }, `User Story: ${item.parent || "N/A"}`));
    chart.append(group);
  });
  chart.append(svgElement("text", { x: left + plotWidth / 2, y: 435, "text-anchor": "middle" }, "Task ID"));
  container.append(chart);
}

function renderSiteData(data) {
  displayedSiteData = data;
  updateStatusSiteFilter();
  renderFilteredSiteData();
}

function renderFilteredSiteData() {
  const data = displayedSiteData;
  const container = document.getElementById("site-results-grid");
  container.replaceChildren();
  const sites = new Map();
  data.items.filter(item => normalizeColumn(item.workItemType) === "task").forEach(item => {
    if (hiddenStatusSites.has(statusFilterSiteName(item.site))) return;
    const key = normalizeColumn(item.site);
    if (!sites.has(key)) sites.set(key, { name: siteDisplayName(item.site), items: [] });
    sites.get(key).items.push(item);
  });
  const element = (tag, className, text) => {
    const node = document.createElement(tag);
    node.className = className;
    if (text !== undefined) node.textContent = text;
    return node;
  };
  if (!sites.size) container.append(element("p", "site-status-empty", hiddenStatusSites.size ? "No sites with Task rows selected. Select a site from the dropdown to view its status." : "No Task rows available within the site ranges from Dashboard."));
  (data.warnings || []).forEach(warning => container.append(element("p", "site-status-empty", warning)));
  [...sites.values()].forEach((site, index) => {
    const card = element("section", "panel site-results-card");
    const header = element("div", "site-results-header");
    const developers = element("p", "developer-names", `Developer(s): ${[...new Set(site.items.map(item => assigneeDisplayName(item.developer)))].join(", ")}`);
    const name = element("strong", "site-name", site.name);
    name.id = `site-name-${index}`;
    card.setAttribute("aria-labelledby", name.id);
    header.append(developers, name);
    const legend = element("div", "legend work-legend");
    [["estimate-key", "Original estimate (hrs)"], ["completed-key", "Completed work (hrs)"]].forEach(([className, label]) => {
      const entry = element("span", "");
      const swatch = element("i", className);
      swatch.setAttribute("aria-hidden", "true");
      entry.append(swatch, element("span", "", label));
      legend.append(entry);
    });
    const chart = element("div", "work-chart");
    chart.setAttribute("role", "region");
    chart.setAttribute("tabindex", "0");
    chart.setAttribute("aria-label", `${site.name}: work item hours by Task ID; scroll horizontally to view all tasks`);
    renderWorkItems(site.items, chart);
    card.append(header, legend, chart);
    container.append(card);
  });
  container.hidden = false;
  document.getElementById("site-status-empty").hidden = true;
}

let siteUploadVersion = 0;
let loadedStatusWorkbook = null;

function refreshStatusDateRange() {
  if (!loadedStatusWorkbook) return;
  if (typeof updateUserStoryStatus === "function") updateUserStoryStatus(loadedStatusWorkbook);
  try {
    renderSiteData(readSiteWorksheet(loadedStatusWorkbook));
  } catch (error) {
    displayedSiteData = null;
    document.getElementById("site-results-grid").replaceChildren();
    document.getElementById("site-results-grid").hidden = true;
    const message = document.getElementById("site-status-empty");
    message.hidden = false;
    message.textContent = error.message;
  }
}
const uploadExcelButton = document.getElementById("upload-excel-button");
document.getElementById("create-report-button").addEventListener("click", async () => {
  if (!loadedStatusWorkbook) {
    window.alert("Upload Excel file to create report");
    return;
  }
  await downloadSprintReport();
});
excelInput.addEventListener("change", () => {
  ++siteUploadVersion;
  loadedStatusWorkbook = null;
  displayedSiteData = null;
  document.getElementById("site-status-empty").textContent = "Upload an Excel file on the Dashboard tab to view site wise status.";
  if (typeof resetUserStoryStatus === "function") resetUserStoryStatus();
  const file = excelInput.files[0];
  if (typeof renderDashboardWorksheet === "function") renderDashboardWorksheet({ sites: [] });
  document.getElementById("site-results-grid").hidden = true;
  document.getElementById("site-status-empty").hidden = false;
  if (typeof resetSharedBurndown === "function") resetSharedBurndown();
  if (typeof resetSprintTagCounts === "function") resetSprintTagCounts();
  uploadExcelButton.disabled = true;
  uploadExcelButton.textContent = "Upload";
  if (!file) {
    uploadStatus.textContent = "No file selected.";
    return;
  }
  uploadExcelButton.disabled = false;
  uploadStatus.textContent = `Selected ${file.name}. Click Upload to display dashboard and work item data.`;
});

uploadExcelButton.addEventListener("click", async () => {
  const file = excelInput.files[0];
  if (!file || uploadExcelButton.disabled) return;
  const version = ++siteUploadVersion;
  loadedStatusWorkbook = null;
  displayedSiteData = null;
  if (typeof resetUserStoryStatus === "function") resetUserStoryStatus();
  if (typeof resetSprintTagCounts === "function") resetSprintTagCounts();
  document.getElementById("site-results-grid").hidden = true;
  uploadExcelButton.disabled = true;
  uploadExcelButton.textContent = "Uploading…";
  if (!/\.(xlsx|xls)$/i.test(file.name)) {
    excelInput.value = "";
    uploadStatus.textContent = "Please select an Excel file (.xlsx or .xls).";
    return;
  }
  uploadStatus.textContent = `Reading ${file.name}…`;
  try {
    if (typeof XLSX === "undefined") throw new Error("Excel reader could not load. Keep the vendor folder with the dashboard files and refresh the page.");
    const bytes = await file.arrayBuffer();
    if (version !== siteUploadVersion) return;
    const workbook = XLSX.read(bytes, { type: "array" });
    const messages = [];
    if (typeof updateUserStoryStatus === "function") messages.push(updateUserStoryStatus(workbook));
    if (typeof updateSharedBurndown === "function") updateSharedBurndown(workbook, file.name);
    let dashboardLoaded = false;
    const sprintTagsLoaded = typeof updateSprintTagCounts === "function" && updateSprintTagCounts(workbook);
    if (typeof readDashboardWorksheet === "function") {
      renderDashboardWorksheet({ sites: [] });
      try {
        const dashboard = readDashboardWorksheet(workbook);
        renderDashboardWorksheet(dashboard);
        dashboardLoaded = true;
        messages.push(`Loaded ${dashboard.sites.length} sprint rows from worksheet "${dashboard.sheetName}".`);
        if (dashboard.summary) messages.push("Loaded Dashboard summary values.");
        if (dashboard.summaryWarning) messages.push(dashboard.summaryWarning);
      } catch (error) {
        messages.push(error.message);
      }
    }
    try {
      const data = readSiteWorksheet(workbook);
      renderSiteData(data);
      messages.push(`Loaded ${data.items.length} work items from worksheet "${data.sheetName}".`);
      if (data.warnings?.length) messages.push(`Site wise status: ${data.warnings.join(" ")}`);
    } catch (error) {
      if (!dashboardLoaded && !sprintTagsLoaded) throw new Error([...messages, error.message].join(" "), { cause: error });
      messages.push(`Site wise status unavailable: ${error.message}`);
    }
    loadedStatusWorkbook = workbook;
    uploadStatus.textContent = messages.join(" ");
  } catch (error) {
    if (version !== siteUploadVersion) return;
    uploadStatus.textContent = `Could not read file. ${error.message || "Choose a valid Excel workbook."}`;
    if (error.code === "MISSING_MASTER_DAILY_DUMP" || error.cause?.code === "MISSING_MASTER_DAILY_DUMP") window.alert(error.message);
  } finally {
    if (version === siteUploadVersion) {
      uploadExcelButton.disabled = false;
      uploadExcelButton.textContent = "Upload";
    }
  }
});
