"use strict";

const userStorySites = dashboardSiteNames;
const selectedUserStorySites = new Set(userStorySites);
let userStoryStatusData = null;

function initializeUserStorySiteFilter() {
  const options = document.getElementById("user-story-site-options");
  const inputs = new Map();
  const syncSelection = () => {
    const count = selectedUserStorySites.size;
    document.getElementById("user-story-site-selection").textContent = count === userStorySites.length ? "All" : count === 0 ? "None" : `${count} selected`;
    inputs.forEach((input, site) => { input.checked = selectedUserStorySites.has(site); });
    all.checked = count === userStorySites.length;
    all.indeterminate = count > 0 && count < userStorySites.length;
    if (userStoryStatusData) renderFilteredUserStoryStatus();
  };
  const addOption = (text, onChange) => {
    const label = userStoryElement("label", "");
    const input = userStoryElement("input", "");
    input.type = "checkbox";
    input.checked = true;
    input.addEventListener("change", () => { onChange(input.checked); syncSelection(); });
    label.append(userStoryElement("span", "", text), input);
    options.append(label);
    return input;
  };
  const all = addOption("All sites", checked => {
    selectedUserStorySites.clear();
    if (checked) userStorySites.forEach(site => selectedUserStorySites.add(site));
  });
  userStorySites.forEach(site => inputs.set(site, addOption(site, checked => {
    if (checked) selectedUserStorySites.add(site);
    else selectedUserStorySites.delete(site);
  })));
  initializeSiteFilterDismissal("user-story");
}

function taskStoryProgress(task) {
  const completed = task.completed ?? 0;
  const total = task.remaining == null ? task.estimate : completed + task.remaining;
  return total > 0 ? Math.max(0, Math.min(100, completed / total * 100)) : null;
}

function userStoryProgress(tasks) {
  if (!tasks.length) return null;
  let completed = 0, total = 0;
  for (const task of tasks) {
    const taskCompleted = task.completed ?? 0;
    const taskTotal = task.remaining == null ? task.estimate : taskCompleted + task.remaining;
    if (taskTotal == null) return null;
    completed += taskCompleted;
    total += taskTotal;
  }
  return total > 0 ? Math.max(0, Math.min(100, completed / total * 100)) : null;
}

function taskStoryColor(value) {
  if (value == null) return "neutral";
  if (value <= 75) return "blue";
  if (value <= 85) return "orange";
  if (value < 100) return "purple";
  return "green";
}

function userStoryDate(value, date1904) {
  const dayFirst = /^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/.exec(String(value ?? "").trim());
  return sprintRowDate(dayFirst ? `${dayFirst[3]}-${dayFirst[2].padStart(2, "0")}-${dayFirst[1].padStart(2, "0")}` : value, date1904);
}

function readUserStorySprintDates(workbook, siteName = "Alingsas") {
  const dates = readSprintDashboardDates(workbook).get(sprintSiteKey(siteName));
  if (!dates) throw new Error(`${siteName}: Dashboard sprint dates are missing, invalid, reversed, or conflicting.`);
  return dates;
}

function readUserStoryStatus(workbook, siteName = "Alingsas") {
  const dates = readUserStorySprintDates(workbook, siteName);
  const start = parseDashboardDate(dates.sprintStartDate).getTime();
  const end = parseDashboardDate(dates.sprintEndDate).getTime();
  const sheetName = masterDailyDumpSheet(workbook);
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "", raw: true });
  const required = ["Area Path", "ID", "Date", "Work Item Type", "Parent", "Completed Work", "Remaining Work", "Original Estimate"];
  const header = rows.findIndex(row => required.every(name => row.some(cell => normalizeColumn(cell) === normalizeColumn(name))));
  if (header < 0) throw new Error(`Master Daily Dump must contain these headers: ${required.join(", ")}.`);
  const column = name => rows[header].findIndex(cell => normalizeColumn(cell) === normalizeColumn(name));
  const latest = new Map();
  const dateColumn = column("Date");
  const hour = (value, name, row) => {
    if (String(value ?? "").trim() === "") return null;
    const number = Number(String(value).trim().replace(/,/g, ""));
    if (!Number.isFinite(number) || number < 0) throw new Error(`Excel row ${row}: ${name} must contain non-negative hours.`);
    return number;
  };
  rows.slice(header + 1).forEach((row, index) => {
    const area = String(row[column("Area Path")] ?? "").trim().split(/[\\/]/).pop();
    if (sprintSiteKey(area) !== sprintSiteKey(siteName)) return;
    const type = normalizeColumn(row[column("Work Item Type")]);
    if (!["user story", "task"].includes(type)) return;
    const date = userStoryDate(row[dateColumn], Boolean(workbook.Workbook?.WBProps?.date1904));
    if (!date || date.getTime() < start || date.getTime() > end) return;
    const id = String(row[column("ID")] ?? "").trim();
    if (!id) throw new Error(`Excel row ${header + index + 2}: ID is missing.`);
    const item = { id, type, parent: String(row[column("Parent")] ?? "").trim(), stamp: date.getTime() };
    // Story progress is calculated from child task hours, not story row hours.
    if (type === "task") {
      for (const [key, name] of [["completed", "Completed Work"], ["remaining", "Remaining Work"], ["estimate", "Original Estimate"]]) {
        item[key] = hour(row[column(name)], name, header + index + 2);
      }
    }
    const previous = latest.get(id);
    if (!previous || item.stamp >= previous.stamp) latest.set(id, item);
  });
  const items = [...latest.values()];
  const stories = items.filter(item => item.type === "user story").map(story => ({
    ...story, tasks: items.filter(item => item.type === "task" && item.parent === story.id)
  }));
  return { site: siteName, stories, dates };
}

function userStoryElement(tag, className, text) {
  const element = document.createElement(tag);
  element.className = className;
  if (text !== undefined) element.textContent = text;
  return element;
}

function userStoryBar(label, value) {
  const bar = userStoryElement("div", `story-progress story-progress-${taskStoryColor(value)}`);
  const fill = userStoryElement("span", "story-progress-fill");
  fill.style.width = `${value ?? 0}%`;
  fill.setAttribute("aria-hidden", "true");
  const display = value == null ? "—" : `${Number(value.toFixed(1))}%`;
  bar.append(fill, userStoryElement("span", "story-progress-label", label), userStoryElement("span", "story-progress-percent", display));
  if (value != null) {
    bar.setAttribute("role", "progressbar");
    bar.setAttribute("aria-valuemin", "0");
    bar.setAttribute("aria-valuemax", "100");
    bar.setAttribute("aria-valuenow", String(value));
    bar.setAttribute("aria-label", label);
  } else bar.setAttribute("aria-label", `${label}: progress unavailable`);
  return bar;
}

function renderUserStoryStatus(data) {
  userStoryStatusData = data;
  renderFilteredUserStoryStatus();
}

function renderFilteredUserStoryStatus() {
  const container = document.getElementById("user-story-cards");
  container.replaceChildren();
  const message = document.getElementById("user-story-status-message");
  const data = userStoryStatusData;
  const sites = (data.sites || [data]).filter(site => selectedUserStorySites.has(site.site || "Alingsas"));
  message.hidden = sites.some(site => site.stories.length > 0);
  message.textContent = message.hidden ? "" : "No User Story rows found within the site's Dashboard date range.";
  if (!selectedUserStorySites.size) message.textContent = "No sites selected. Select a site from the dropdown to view its user stories.";
  sites.forEach(site => container.append(renderUserStorySiteCard(site)));
}

function renderUserStorySiteCard(data) {
  const siteName = data.site || "Alingsas";
  const card = userStoryElement("article", "user-story-card");
  card.setAttribute("aria-label", `${siteName} user stories and child tasks`);
  const header = userStoryElement("div", "user-story-card-header");
  header.append(userStoryElement("span", "user-story-dates", formatSprintDuration(data.dates || {})), userStoryElement("strong", "site-name", siteName));
  const trees = userStoryElement("div", "user-story-trees");
  trees.setAttribute("aria-label", `${siteName} user stories and tasks`);
  card.append(header, trees);
  if (!data.stories.length) trees.append(userStoryElement("p", "site-status-empty", data.error || "No user stories found within the site's Dashboard date range."));
  data.stories.forEach(story => {
    const tree = userStoryElement("section", "user-story-tree");
    tree.setAttribute("aria-label", `US ${story.id} and child tasks`);
    tree.append(userStoryBar(`US ${story.id}`, userStoryProgress(story.tasks)));
    const tasks = userStoryElement("ul", "user-story-tasks");
    story.tasks.forEach(task => {
      const row = userStoryElement("li", "user-story-task");
      const bar = userStoryBar(`Task ${task.id}`, taskStoryProgress(task));
      bar.title = `Original Estimate: ${task.estimate ?? "blank"} hrs; Completed Work: ${task.completed ?? "blank"} hrs; Remaining Work: ${task.remaining ?? "blank"} hrs`;
      row.append(bar);
      tasks.append(row);
    });
    if (story.tasks.length) tree.append(tasks);
    else tree.append(userStoryElement("p", "site-status-empty", "No child tasks found for this user story."));
    trees.append(tree);
  });
  return card;
}

function resetUserStoryStatus() {
  userStoryStatusData = null;
  document.getElementById("user-story-cards").replaceChildren();
  document.getElementById("user-story-status-message").hidden = false;
  document.getElementById("user-story-status-message").textContent = "Upload an Excel workbook from Dashboard to view user stories and tasks for all six sites.";
}

initializeUserStorySiteFilter();
if (typeof initializeStatusSiteFilter === "function") initializeStatusSiteFilter();

function updateUserStoryStatus(workbook) {
  try {
    const sites = userStorySites.map(site => {
      try { return readUserStoryStatus(workbook, site); }
      catch (error) { return { site, stories: [], dates: {}, error: error.message }; }
    });
    renderUserStoryStatus({ sites });
    const count = sites.reduce((total, site) => total + site.stories.length, 0);
    const errors = sites.filter(site => site.error).map(site => `${site.site}: ${site.error}`);
    return `Loaded ${count} user stories across ${sites.filter(site => !site.error).length} sites in User Story Status.${errors.length ? ` Some sites unavailable: ${errors.join(" ")}` : ""}`;
  } catch (error) {
    resetUserStoryStatus();
    const message = `User Story Status unavailable: ${error.message}`;
    document.getElementById("user-story-status-message").textContent = message;
    return message;
  }
}
