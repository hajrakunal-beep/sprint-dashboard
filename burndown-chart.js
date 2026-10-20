"use strict";

function masterDailyDumpSheet(workbook) {
  const sheetName = "Master Daily Dump";
  if (!workbook.SheetNames.includes(sheetName) || !workbook.Sheets[sheetName]) {
    const error = new Error('Required worksheet "Master Daily Dump" is missing. Please upload a workbook containing this sheet.');
    error.code = "MISSING_MASTER_DAILY_DUMP";
    throw error;
  }
  return sheetName;
}

function burndownDate(value, date1904 = false) {
  let year, month, day;
  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value, { date1904 });
    if (parsed) ({ y: year, m: month, d: day } = parsed);
  } else if (value instanceof Date) {
    year = value.getUTCFullYear(); month = value.getUTCMonth() + 1; day = value.getUTCDate();
  } else {
    const match = String(value ?? "").trim().match(/^(\d{4})-(\d{2})-(\d{2})(?:[T ].*)?$/);
    if (match) [, year, month, day] = match.map(Number);
  }
  const stamp = Date.UTC(Number(year), Number(month) - 1, Number(day));
  const date = new Date(stamp);
  if (!Number.isFinite(stamp) || date.getUTCFullYear() !== Number(year) || date.getUTCMonth() + 1 !== Number(month) || date.getUTCDate() !== Number(day)) {
    throw new Error("Use Excel date cells or YYYY-MM-DD dates in the Date column.");
  }
  return { stamp, label: date.toISOString().slice(0, 10) };
}

function readBurndownWorkbook(workbook) {
  const datesBySite = readSprintDashboardDates(workbook);
  const normalize = value => String(value ?? "").trim().replace(/\s+/g, " ").toLowerCase();
  const sheetName = masterDailyDumpSheet(workbook);
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "", raw: true });
  const headerIndex = rows.findIndex(row => ["date", "area path", "remaining work"].every(name => row.some(value => normalize(value) === name)));
  if (headerIndex < 0) throw new Error('Worksheet "Master Daily Dump" needs Date, Area Path and Remaining Work columns. A task snapshot without dates cannot show burndown over time.');
  const dateIndex = rows[headerIndex].findIndex(value => normalize(value) === "date");
  const remainingIndex = rows[headerIndex].findIndex(value => normalize(value) === "remaining work");
  const siteIndex = rows[headerIndex].findIndex(value => normalize(value) === "area path");
  const optionalIndex = name => rows[headerIndex].findIndex(value => normalize(value) === name);
  const completedIndex = optionalIndex("completed work");
  const capacityIndex = optionalIndex("remaining capacity"), totalIndex = optionalIndex("total capacity");
  const sites = new Map();
  const warnings = new Map();
  rows.slice(headerIndex + 1).forEach((row, index) => {
    if (!row.some(value => String(value).trim() !== "")) return;
    const site = String(row[siteIndex] ?? "").trim() || "Unspecified site";
    const siteKey = sprintSiteKey(site.split(/[\\/]/).pop());
    const range = datesBySite.get(siteKey);
    if (!range) {
      warnings.set(siteKey, `${burndownFilterSiteName(site)}: Dashboard Start and End dates are missing, invalid, reversed, or conflicting. Chart unavailable.`);
      return;
    }
    const parsedDate = userStoryDate(row[dateIndex], Boolean(workbook.Workbook?.WBProps?.date1904));
    if (!parsedDate) return;
    const date = { stamp: parsedDate.getTime(), label: parsedDate.toISOString().slice(0, 10) };
    if (date.label < range.sprintStartDate || date.label > range.sprintEndDate) return;
    const key = normalize(site);
    if (!sites.has(key)) sites.set(key, { site, range, daily: new Map() });
    const group = sites.get(key), daily = group.daily;
    const workHours = (column, label) => {
      if (column < 0 || String(row[column] ?? "").trim() === "") return null;
      const value = Number(String(row[column]).trim().replace(/,/g, ""));
      if (!Number.isFinite(value) || value < 0) throw new Error(`Row ${headerIndex + index + 2}: ${label} must be a non-negative number.`);
      return value;
    };
    const totalCapacity = workHours(totalIndex, "Capacity");
    if (totalCapacity !== null) {
      if (group.totalCapacity !== undefined && group.totalCapacity !== totalCapacity) throw new Error(`${site}: Total Capacity must be consistent across rows.`);
      group.totalCapacity = totalCapacity;
    }
    const remaining = workHours(remainingIndex, "Remaining Work");
    const completed = workHours(completedIndex, "Completed Work");
    if (remaining === null && completed === null) return;
    const entry = daily.get(date.stamp) || { ...date, remaining: null, completed: null };
    // Sum every work item type within this site/date snapshot. Do not sum
    // Completed Work across dates: exported values may already be cumulative.
    if (remaining !== null) entry.remaining = (entry.remaining ?? 0) + remaining;
    if (completed !== null) entry.completed = (entry.completed ?? 0) + completed;
    const capacity = workHours(capacityIndex, "Capacity");
    // Capacity is a site-wide daily value, repeated on each task row if needed.
    if (capacity !== null) {
      if (entry.capacity !== undefined && entry.capacity !== capacity) throw new Error(`${site}: Remaining Capacity must be consistent for each date.`);
      entry.capacity = capacity;
    }
    daily.set(date.stamp, entry);
  });
  const groups = [...sites.values()].filter(group => group.daily.size).map(({ site, range, daily, totalCapacity }) => {
    const points = [...daily.values()].sort((a, b) => a.stamp - b.stamp);
    const explicitBounds = true;
    const start = { stamp: parseDashboardDate(range.sprintStartDate).getTime(), label: range.sprintStartDate };
    const end = { stamp: parseDashboardDate(range.sprintEndDate).getTime(), label: range.sprintEndDate };
    const days = [];
    for (let stamp = start.stamp; stamp <= end.stamp; stamp += 86400000) {
      const weekday = new Date(stamp).getUTCDay();
      // Keep explicitly recorded weekend work; otherwise omit weekends.
      if ((weekday !== 0 && weekday !== 6) || daily.has(stamp) || stamp === start.stamp || stamp === end.stamp) {
        days.push({ stamp, label: new Date(stamp).toISOString().slice(0, 10) });
      }
    }
    const initial = totalCapacity ?? points[0]?.remaining ?? 0;
    points.forEach(point => { point.ideal = days.length > 1 ? initial * (1 - days.findIndex(day => day.stamp === point.stamp) / (days.length - 1)) : initial; });
    return { site, points, days, initial, start, end, explicitBounds };
  });
  return { sheetName, groups, warnings: [...warnings.values()] };
}

function renderBurndownChart(data, chart) {
  const { points, days, initial } = data;
  const width = 640, height = 410, left = 65, right = 25, top = 35, baseline = 315;
  const max = points.reduce((value, point) => Math.max(value, point.remaining ?? 0, point.completed ?? 0, point.capacity ?? 0), Math.max(initial, 1));
  const step = Math.max(1, Math.ceil(max / 5)), axisMax = step * 5;
  const x = point => left + days.findIndex(day => day.stamp === point.stamp) / Math.max(1, days.length - 1) * (width - left - right);
  const y = hours => baseline - hours / axisMax * (baseline - top);
  chart.setAttribute("viewBox", `0 0 ${width} ${height}`);
  chart.replaceChildren(svgElement("title", {}, `${data.site}: total remaining and completed work by date`));
  chart.append(svgElement("desc", {}, "Blue area shows total remaining hours and orange line shows total completed hours across all work item types for each recorded date. Grey line shows the ideal remaining trend. Green dashed line shows remaining capacity when available. Weekends without observations are omitted."));
  chart.append(svgElement("text", { x: left, y: 18 }, "Work (hours)"));
  for (let tick = 0; tick <= 5; tick++) {
    chart.append(svgElement("line", { x1: left, y1: y(tick * step), x2: width - right, y2: y(tick * step), stroke: "#d4dce4" }));
    chart.append(svgElement("text", { x: left - 10, y: y(tick * step) + 4, "text-anchor": "end" }, tick * step));
  }
  const drawWorkSeries = (key, color, area = false) => {
    let run = [];
    const flush = () => {
      if (!run.length) return;
      const coordinates = run.map(point => `${x(point)},${y(point[key])}`).join(" ");
      if (area) chart.append(svgElement("polygon", { points: `${x(run[0])},${baseline} ${coordinates} ${x(run[run.length - 1])},${baseline}`, fill: color, "fill-opacity": 0.9 }));
      chart.append(svgElement("polyline", { points: coordinates, fill: "none", stroke: color, "stroke-width": area ? 2 : 3 }));
      run.forEach(point => {
        const dot = svgElement("circle", { cx: x(point), cy: y(point[key]), r: 3, fill: color });
        dot.append(svgElement("title", {}, `${point.label}: ${key} ${point[key]} hrs; remaining ${point.remaining ?? "N/A"} hrs, completed ${point.completed ?? "N/A"} hrs`));
        chart.append(dot);
      });
      run = [];
    };
    points.forEach(point => {
      if (point[key] == null) flush();
      else run.push(point);
    });
    flush();
  };
  drawWorkSeries("remaining", "#0078d4", true);
  chart.append(svgElement("line", { x1: left, y1: y(initial), x2: width - right, y2: baseline, stroke: "#707070", "stroke-width": 3 }));
  let capacityRun = [];
  const drawCapacity = () => {
    if (capacityRun.length) chart.append(svgElement("polyline", { points: capacityRun.join(" "), fill: "none", stroke: "#38952c", "stroke-width": 3, "stroke-dasharray": "8 5" }));
    capacityRun = [];
  };
  days.forEach(day => {
    const point = points.find(point => point.stamp === day.stamp);
    if (point?.capacity === undefined) { drawCapacity(); return; }
    capacityRun.push(`${x(point)},${y(point.capacity)}`);
    chart.append(svgElement("circle", { cx: x(point), cy: y(point.capacity), r: 2.5, fill: "#38952c" }));
  });
  drawCapacity();
  drawWorkSeries("completed", "#e87500");
  const labelInterval = Math.max(1, Math.ceil(days.length / 10));
  days.forEach((day, index) => {
    if (index % labelInterval !== 0 && index !== days.length - 1) return;
    chart.append(svgElement("text", { x: x(day), y: baseline + 22, transform: `rotate(-35 ${x(day)} ${baseline + 22})`, "text-anchor": "end" }, day.label.slice(5)));
  });
  chart.append(svgElement("text", { x: width / 2, y: height - 8, "text-anchor": "middle" }, "Sprint working days"));
}

let loadedBurndownWorkbook = null;
let loadedBurndownFilename = "";
let displayedBurndownData = null;
let displayedBurndownIssue = "";
const hiddenBurndownSites = new Set();
let burndownSiteFilterInitialized = false;

function burndownFilterSiteName(value) {
  if (typeof statusFilterSiteName === "function") return statusFilterSiteName(value);
  return String(value).trim().split(/[\\/]/).pop();
}

function initializeBurndownSiteFilter() {
  if (burndownSiteFilterInitialized) return;
  burndownSiteFilterInitialized = true;
  updateBurndownSiteFilter();
  initializeSiteFilterDismissal("burndown-chart");
}

function updateBurndownSiteFilter() {
  if (!burndownSiteFilterInitialized) return;
  const names = [...new Set([
    ...dashboardSiteNames,
    ...(displayedBurndownData?.groups || []).map(group => burndownFilterSiteName(group.site))
  ])];
  renderSiteFilter("burndown-chart", names, hiddenBurndownSites, () => {
    if (displayedBurndownData) renderFilteredBurndownCards();
  });
}

function refreshBurndownDateRange() {
  if (loadedBurndownWorkbook) updateSharedBurndown(loadedBurndownWorkbook, loadedBurndownFilename);
}

function resetSharedBurndown() {
  displayedBurndownData = null;
  displayedBurndownIssue = "";
  loadedBurndownWorkbook = null;
  loadedBurndownFilename = "";
  document.getElementById("burndown-site-charts").replaceChildren();
  document.getElementById("burndown-shared-status").textContent = "Upload an Excel file on the Dashboard tab to view burndown charts.";
}

function updateSharedBurndown(workbook, filename, siteData = { items: [] }) {
  loadedBurndownWorkbook = workbook;
  loadedBurndownFilename = filename;
  let data, issue = "";
  try {
    data = readBurndownWorkbook(workbook);
  } catch (error) {
    issue = error.message;
    const sites = new Map();
    siteData.items.forEach(item => { if (!sites.has(normalizeColumn(item.site))) sites.set(normalizeColumn(item.site), item.site); });
    data = { groups: [...sites.values()].map(site => ({ site, points: [] })) };
  }
  displayedBurndownData = data;
  displayedBurndownIssue = issue;
  updateBurndownSiteFilter();
  renderFilteredBurndownCards();
}

function renderFilteredBurndownCards() {
  const data = displayedBurndownData;
  const issue = displayedBurndownIssue;
  const filename = loadedBurndownFilename;
  const container = document.getElementById("burndown-site-charts");
  container.replaceChildren();
  const node = (tag, className, text) => {
    const element = document.createElement(tag);
    element.className = className;
    if (text !== undefined) element.textContent = text;
    return element;
  };
  const groups = data.groups.filter(group => !hiddenBurndownSites.has(burndownFilterSiteName(group.site)));
  groups.forEach((group, index) => {
    const card = node("section", "panel burndown-site-chart-card");
    const heading = node("h2", "", group.site);
    heading.id = `burndown-chart-site-${index}`;
    card.setAttribute("aria-labelledby", heading.id);
    card.append(heading);
    if (group.points.length < 2) {
      card.append(node("p", "", issue || "At least two dated observations are needed to show progress for this site."));
    } else {
      card.append(node("p", "burndown-date-range", `${group.start.label} – ${group.end.label}`));
      const metrics = node("dl", "burndown-chart-metrics");
      const last = group.points[group.points.length - 1];
      [["Remaining work", last.remaining === null ? "N/A" : `${last.remaining} hrs`], ["Completed work", last.completed === null ? "N/A" : `${last.completed} hrs`], ["Latest observation", last.label]].forEach(([label, value]) => {
        const metric = node("div", "");
        metric.append(node("dt", "", label), node("dd", "", value));
        metrics.append(metric);
      });
      card.append(metrics);
      const legend = node("div", "legend burndown-line-legend");
      const keys = [["burndown-actual-key", "Remaining work"], ["burndown-completed-key", "Completed work"], ["burndown-ideal-key", "Ideal trend"]];
      if (group.points.some(point => point.capacity !== undefined)) keys.splice(1, 0, ["burndown-capacity-key", "Remaining capacity"]);
      keys.forEach(([className, label]) => {
        const entry = node("span", "");
        entry.append(node("i", className), node("span", "", label));
        legend.append(entry);
      });
      const chart = svgElement("svg", { role: "img", "aria-label": `${group.site}: total remaining and completed hours by sprint date` });
      renderBurndownChart(group, chart);
      card.append(chart, legend);
      if (!group.points.some(point => point.completed !== null)) card.append(node("p", "burndown-chart-note", "Supply dated Completed Work values to show the completed-work line."));
    }
    container.append(card);
  });
  document.getElementById("burndown-shared-status").textContent = issue
    ? `Source: ${filename}. ${issue}`
    : `Source: ${filename}. ${groups.length} sites. Date ranges: each site's Dashboard Start and End.${groups.length ? "" : data.groups.length ? " No sites selected with observations. Select a site from the dropdown to view its chart." : " No observations within the Dashboard site ranges."}${data.warnings?.length ? ` ${data.warnings.join(" ")}` : ""}`;
}

if (typeof initializeStatusSiteFilter === "function") initializeBurndownSiteFilter();
