"use strict";

function defaultDashboardPeriod(today = new Date()) {
  // Use the browser's local calendar date, then subtract calendar days in UTC.
  const end = new Date(Date.UTC(today.getFullYear(), today.getMonth(), today.getDate()));
  const start = new Date(end);
  start.setUTCDate(start.getUTCDate() - 15);
  return { from: start.toISOString().slice(0, 10), to: end.toISOString().slice(0, 10) };
}

// Reporting period defaults; summary values and sites come from Excel.
const initialDashboardPeriod = defaultDashboardPeriod();
const dashboardData = {
  reportingPeriod: formatDashboardPeriod(initialDashboardPeriod.from, initialDashboardPeriod.to),
  reportFrom: initialDashboardPeriod.from,
  reportTo: initialDashboardPeriod.to,
  asOf: new Intl.DateTimeFormat("en-GB", {
    day: "2-digit",
    month: "short",
    year: "numeric"
  }).format(new Date()).toUpperCase(),
  summary: { totalSites: null, plannedHours: null, completedHours: null, overallActual: null, delayedAtRisk: null },

  // Populated from the uploaded Dashboard worksheet.
  sites: []
};

const percent = value => value == null ? "—" : `${Number(value).toFixed(1)}%`;
const clampPercent = value => Math.max(0, Math.min(100, Number(value) || 0));

let selectedDashboardPeriod = null;
const dashboardDateFrom = document.getElementById("dashboard-date-from");
const dashboardDateTo = document.getElementById("dashboard-date-to");
const dashboardDateError = document.getElementById("dashboard-date-error");

function parseDashboardDate(value) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value || "")) return null;
  const date = new Date(`${value}T00:00:00Z`);
  return Number.isFinite(date.getTime()) && date.toISOString().slice(0, 10) === value ? date : null;
}

function formatDashboardPeriod(from, to) {
  const start = parseDashboardDate(from), end = parseDashboardDate(to);
  const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const includeYear = start.getUTCFullYear() !== end.getUTCFullYear();
  const label = date => `${String(date.getUTCDate()).padStart(2, "0")}-${months[date.getUTCMonth()]}${includeYear ? `-${date.getUTCFullYear()}` : ""}`;
  return `${label(start)} to ${label(end)}`;
}

function updateDashboardPeriod() {
  const from = dashboardDateFrom.value, to = dashboardDateTo.value;
  const validFrom = Boolean(parseDashboardDate(from));
  const validTo = Boolean(parseDashboardDate(to));
  const reversed = validFrom && validTo && from > to;
  dashboardDateFrom.setAttribute("aria-invalid", String(!validFrom || reversed));
  dashboardDateTo.setAttribute("aria-invalid", String(!validTo || reversed));
  dashboardDateError.hidden = validFrom && validTo && !reversed;
  dashboardDateError.textContent = dashboardDateError.hidden ? "" : reversed ? "To must be on or after From." : "Select both From and To dates.";
  if (!dashboardDateError.hidden) {
    if (typeof refreshSprintDateRange === "function") refreshSprintDateRange();
    return;
  }
  selectedDashboardPeriod = { from, to };
  document.getElementById("reporting-period").textContent = `– Reporting period (${formatDashboardPeriod(from, to)})`;
  if (typeof refreshSprintDateRange === "function") refreshSprintDateRange();
}

dashboardDateFrom.addEventListener("change", updateDashboardPeriod);
dashboardDateTo.addEventListener("change", updateDashboardPeriod);

function renderSummary(data) {
  if (!selectedDashboardPeriod && parseDashboardDate(data.reportFrom) && parseDashboardDate(data.reportTo) && data.reportFrom <= data.reportTo) {
    selectedDashboardPeriod = { from: data.reportFrom, to: data.reportTo };
    dashboardDateFrom.value = data.reportFrom;
    dashboardDateTo.value = data.reportTo;
  }
  const period = selectedDashboardPeriod ? formatDashboardPeriod(selectedDashboardPeriod.from, selectedDashboardPeriod.to) : data.reportingPeriod;
  document.getElementById("reporting-period").textContent = `– Reporting period (${period})`;
  document.getElementById("as-of").textContent = data.asOf;
  const s = data.summary;
  const cards = [["TOTAL SITES", s.totalSites], ["PLANNED HRS", s.plannedHours], ["COMPLETED HRS", s.completedHours], ["OVERALL ACTUAL %", percent(s.overallActual)], ["DELAYED / AT RISK", s.delayedAtRisk]];
  const container = document.getElementById("summary");
  container.replaceChildren();
  cards.forEach(([label, value]) => {
    const card = document.createElement("div");
    card.className = "summary-card";
    [["summary-label", label], ["summary-value", value]].forEach(([className, text]) => {
      const element = document.createElement("div");
      element.className = className;
      element.textContent = text ?? "—";
      card.append(element);
    });
    container.append(card);
  });
}

function addCell(row, value, className = "") {
  const cell = document.createElement("td");
  cell.textContent = value;
  cell.className = className;
  row.append(cell);
  return cell;
}

function addProgressCell(row, value, className = "") {
  const cell = addCell(row, "", "progress-cell");
  const bar = document.createElement("span");
  bar.className = `cell-bar${className ? ` ${className}` : ""}`;
  bar.style.width = `${clampPercent(value)}%`;
  bar.setAttribute("aria-hidden", "true");
  const label = document.createElement("span");
  label.className = "progress-value";
  label.textContent = percent(value);
  cell.append(bar, label);
}

function renderTable(sites) {
  const tbody = document.getElementById("sprint-rows");
  tbody.replaceChildren();
  const healthClasses = { Green: "health-green", Amber: "health-amber", Red: "health-red" };
  if (!sites.length) {
    const row = document.createElement("tr");
    const cell = addCell(row, "Upload an Excel workbook with a Dashboard sheet to view sprint delivery details.");
    cell.colSpan = 13;
    tbody.append(row);
    return;
  }
  sites.forEach(site => {
    const row = document.createElement("tr");
    [site.site, site.sprint, site.start, site.end, site.plannedHours, site.actualHours].forEach(value => addCell(row, value));
    addProgressCell(row, site.plannedPercent, "planned-progress-bar");
    addProgressCell(row, site.actualPercent);
    // Variance is calculated in percentage points, matching the reference display.
    const variance = Object.hasOwn(site, "variance") ? site.variance : site.actualPercent - site.plannedPercent;
    addCell(row, percent(variance), variance <= -5 ? "variance-negative" : "");
    addCell(row, site.delay);
    const healthName = String(site.health || "").trim().toLowerCase();
    const healthClass = Object.entries(healthClasses).find(([name]) => name.toLowerCase() === healthName)?.[1];
    const healthCell = addCell(row, site.health, `health ${healthClass || "health-unconfigured"}`);
    if (!healthClass && /^[a-z]+$/i.test(healthName) && window.CSS?.supports("color", healthName)) {
      healthCell.style.backgroundColor = healthName;
      healthCell.style.color = "#fff";
      healthCell.style.textShadow = "0 1px 3px #000, 0 0 2px #000";
    }
    addCell(row, site.blocked);
    const commentsCell = addCell(row, "", "comments-cell");
    const commentsInput = document.createElement("textarea");
    commentsInput.className = "dashboard-comments";
    commentsInput.rows = 1;
    commentsInput.wrap = "off";
    commentsInput.placeholder = "Add a comment…";
    commentsInput.setAttribute("aria-label", `Comments for ${site.site}, ${site.sprint}`);
    commentsInput.value = site.comments ?? "";
    const updateCommentsStyle = () => {
      commentsInput.className = `dashboard-comments${commentsInput.value.trim() ? " has-comment" : ""}`;
    };
    updateCommentsStyle();
    commentsInput.addEventListener("input", () => {
      site.comments = commentsInput.value;
      updateCommentsStyle();
    });
    commentsCell.append(commentsInput);
    tbody.append(row);
  });
}

// Native SVG chart: no Chart.js, CDN, framework, or internet connection needed.
function svgElement(tag, attributes = {}, text) {
  const element = document.createElementNS("http://www.w3.org/2000/svg", tag);
  Object.entries(attributes).forEach(([key, value]) => element.setAttribute(key, value));
  if (text !== undefined) element.textContent = text;
  return element;
}

function renderChart(sites) {
  const chart = document.getElementById("progress-chart");
  const width = Math.max(1120, sites.length * 145 + 80);
  const height = 275, left = 65, right = 15, top = 28, bottom = 44;
  const plotHeight = height - top - bottom;
  const baseline = height - bottom;
  chart.setAttribute("viewBox", `0 0 ${width} ${height}`);
  chart.style.minWidth = `${Math.max(800, sites.length * 110)}px`;
  chart.replaceChildren(
    svgElement("title", { id: "chart-title" }, "Planned versus actual progress by site"),
    svgElement("desc", { id: "chart-description" }, "Orange bars show planned progress and teal bars show actual progress, on a 0 to 100 percent scale. The same values are available in the table above.")
  );
  for (let tick = 0; tick <= 100; tick += 20) {
    const y = baseline - tick / 100 * plotHeight;
    chart.append(svgElement("line", { x1: left, y1: y, x2: width - right, y2: y, class: "grid-line" }));
    chart.append(svgElement("text", { x: left - 10, y: y + 4, "text-anchor": "end" }, percent(tick)));
  }
  if (!sites.length) {
    chart.append(svgElement("text", { x: width / 2, y: height / 2, "text-anchor": "middle" }, "No site data available"));
    return;
  }
  const groupWidth = (width - left - right) / sites.length;
  const barWidth = Math.min(48, groupWidth * 0.29);
  sites.forEach((site, index) => {
    const center = left + groupWidth * (index + 0.5);
    [[site.plannedPercent, "planned-bar", center - barWidth - 3, "Planned"], [site.actualPercent, "actual-bar", center + 3, "Actual"]].forEach(([value, className, x, label]) => {
      const barHeight = clampPercent(value) / 100 * plotHeight;
      const y = baseline - barHeight;
      const bar = svgElement("rect", { x, y, width: barWidth, height: barHeight, class: className });
      bar.append(svgElement("title", {}, `${site.site} — ${label}: ${percent(value)}`));
      chart.append(bar, svgElement("text", { x: x + barWidth / 2, y: y - 8, "text-anchor": "middle" }, percent(value)));
    });
    chart.append(svgElement("text", { x: center, y: baseline + 24, "text-anchor": "middle" }, site.site));
  });
}

function renderDashboard(data) {
  renderSummary(data);
  renderTable(data.sites);
  renderChart(data.sites);
}

const toggle = document.getElementById("menu-toggle");
function setNavigation(collapsed) {
  document.body.classList.toggle("nav-collapsed", collapsed);
  toggle.setAttribute("aria-expanded", String(!collapsed));
}
setNavigation(window.matchMedia("(max-width: 760px)").matches);
toggle.addEventListener("click", () => setNavigation(!document.body.classList.contains("nav-collapsed")));
document.addEventListener("keydown", event => {
  if (event.key === "Escape") setNavigation(true);
});
const dashboardViews = ["dashboard", "sprint-burndown", "site-status", "user-story-status", "burndown-chart", "setting"]
  .map(id => [id, document.getElementById(`${id}-nav`)]);

function showView(viewId) {
  dashboardViews.forEach(([id, link]) => {
    const active = id === viewId;
    document.getElementById(id).hidden = !active;
    link.classList.toggle("active", active);
    if (active) link.setAttribute("aria-current", "page");
    else link.removeAttribute("aria-current");
  });
  if (window.matchMedia("(max-width: 760px)").matches) setNavigation(true);
  const heading = document.querySelector(`#${viewId} h1`);
  heading.setAttribute("tabindex", "-1");
  heading.focus();
}

dashboardViews.forEach(([id, link]) => {
  link.addEventListener("click", event => {
    event.preventDefault();
    showView(id);
  });
});

const excelInput = document.getElementById("site-excel-file");
const uploadStatus = document.getElementById("upload-status");
renderDashboard(dashboardData);
