"use strict";

// All six cards' counts and details come from Master Daily Dump.
// Spillover is manually editable for each card during the current session.
// Per-site sprint dates come from the Dashboard worksheet.
const sprintBurndownData = [
  { developer: "Phatale, Mangesh", site: "Vasteras", completionPercent: 0, committed: null, tentative: null, delivered: null, spillover: 0,
    stories: { inProgress: ["US 5588", "US 5736"], resolved: [], spillover: [], tentative: ["US 5363"] } },
  { developer: "Correia Rodrigues, Joao Muthaih, Subramanian", site: "Rakovski", completionPercent: 50, committed: 4, tentative: 2, delivered: 2, spillover: 0,
    stories: { inProgress: ["Bug 5413", "US 5035", "US 5222"], resolved: ["BUG 5396", "BUG 5397"], spillover: [], tentative: ["US 5035", "US 4807"] } },
  { developer: "Ribeiro, Diogo Filipe", site: "Heidelberg", completionPercent: 0, committed: 8, tentative: 1, delivered: 0, spillover: 0,
    stories: { inProgress: ["BUG 4805", "BUG 4935"], resolved: [], spillover: [], tentative: ["US 5214"] } },
  { developer: "Mannepalli, Suman", site: "Senatobia", completionPercent: 0, committed: 4, tentative: 0, delivered: 0, spillover: 0,
    stories: { inProgress: ["US 5516", "US 4853"], resolved: [], spillover: [], tentative: [] } },
  { developer: "Siddiqui, Faiz Mohammed", site: "Alingsas", completionPercent: 0, committed: 1, tentative: 0, delivered: 0, spillover: 0,
    stories: { inProgress: ["US 5533"], resolved: [], spillover: [], tentative: [] } },
  { developer: "S, Shalin and Nivruth", site: "Chassieu", completionPercent: 0, committed: 8, tentative: 2, delivered: 1, spillover: 0,
    stories: { inProgress: ["US 5131", "US 5132", "US 5228", "US 5229", "US 5230", "US 4931", "US 5030"], resolved: ["US 4878"], spillover: [], tentative: ["US 5237", "US 4931"] } }
].map(record => ({ ...record, sprint: "", sprintStartDate: null, sprintEndDate: null }));

function formatSprintDuration(record) {
  const start = parseDashboardDate(record.sprintStartDate);
  const end = parseDashboardDate(record.sprintEndDate);
  if (!start || !end || start > end) return "[Sprint dates unavailable]";
  const format = value => value.split("-").reverse().join("-");
  return `[${format(record.sprintStartDate)} to ${format(record.sprintEndDate)}]`;
}

function burndownElement(tag, className, text) {
  const element = document.createElement(tag);
  element.className = className;
  if (text !== undefined) element.textContent = String(text);
  return element;
}

function makeSpilloverEditable(value, metric, record) {
  value.setAttribute("contenteditable", "plaintext-only");
  value.setAttribute("role", "textbox");
  value.setAttribute("inputmode", "numeric");
  value.setAttribute("aria-label", `Spill over for ${record.site || "this site"}`);
  value.setAttribute("title", "Click to edit Spill over. Enter to save; Escape to cancel.");
  value.setAttribute("spellcheck", "false");
  metric.addEventListener("click", () => value.focus());
  value.addEventListener("beforeinput", event => {
    if (event.data != null && /\D/.test(event.data)) event.preventDefault();
  });
  value.addEventListener("paste", event => {
    if (/\D/.test(event.clipboardData.getData("text/plain"))) event.preventDefault();
  });
  value.addEventListener("input", () => {
    const digits = value.textContent.replace(/\D/g, "");
    if (value.textContent !== digits) value.textContent = digits;
  });
  value.addEventListener("blur", () => {
    const text = value.textContent.trim();
    const count = Number(text);
    if (/^\d+$/.test(text) && Number.isSafeInteger(count)) record.spillover = count;
    value.textContent = String(record.spillover ?? 0);
    metric.classList.toggle("burndown-metric-zero", Number(record.spillover ?? 0) === 0);
  });
  value.addEventListener("keydown", event => {
    if (event.key === "Enter" || event.key === "Escape") {
      event.preventDefault();
      if (event.key === "Escape") value.textContent = String(record.spillover ?? 0);
      value.blur();
    }
  });
}

function renderSprintStoryIds(ids, stories) {
  ids.replaceChildren();
  ids.textContent = "";
  if (!stories.length) ids.textContent = "NA";
  stories.forEach((story, index) => {
    if (index) ids.append(burndownElement("span", "burndown-id-separator", " • "));
    ids.append(burndownElement("span", "burndown-story-id", story));
  });
}

function makeSpilloverStoriesEditable(ids, row, record) {
  ids.setAttribute("contenteditable", "plaintext-only");
  ids.setAttribute("role", "textbox");
  ids.setAttribute("aria-label", `Spillover user stories for ${record.site || "this site"}`);
  ids.setAttribute("title", "Click to edit Spillover IDs, separated by commas. Enter to save; Escape to cancel.");
  ids.setAttribute("spellcheck", "false");
  let cancelled = false;
  row.addEventListener("click", () => ids.focus());
  ids.addEventListener("focus", () => {
    if (!record.stories?.spillover?.length) ids.textContent = "";
  });
  ids.addEventListener("blur", () => {
    if (!cancelled) {
      const text = ids.textContent.trim();
      record.stories ??= {};
      record.stories.spillover = !text || /^na$/i.test(text) ? []
        : text.split(/[,;•\r\n]+/).map(id => id.trim()).filter(Boolean);
    }
    cancelled = false;
    renderSprintStoryIds(ids, record.stories?.spillover || []);
  });
  ids.addEventListener("keydown", event => {
    if (event.key === "Enter" || event.key === "Escape") {
      event.preventDefault();
      cancelled = event.key === "Escape";
      ids.blur();
    }
  });
}

function renderSprintBurndown(records) {
  const container = document.getElementById("burndown-cards");
  container.replaceChildren();
  if (!records.length) {
    container.append(burndownElement("p", "burndown-empty", "No sprint commitment data available."));
    return;
  }
  records.forEach((record, index) => {
    const card = burndownElement("article", "burndown-card");
    const header = burndownElement("div", "burndown-card-header");
    const identity = burndownElement("div", "burndown-identity");
    const name = burndownElement("h2", "burndown-developer", record.developer || "Unassigned");
    name.id = `burndown-developer-${index}`;
    card.setAttribute("aria-labelledby", name.id);
    const site = burndownElement("p", "burndown-site");
    site.append(
      burndownElement("strong", "burndown-site-name", `${record.site || "Site not specified"}${record.sprint ? `\u00a0\u00a0${/^sprint\b/i.test(record.sprint) ? record.sprint : `Sprint ${record.sprint}`}` : ""}`),
      burndownElement("span", "burndown-sprint-duration", formatSprintDuration(record))
    );
    identity.append(name, site);
    const completionPercent = record.committed > 0 && record.delivered != null
      ? Math.round(record.delivered / record.committed * 100)
      : null;
    const completion = burndownElement("span", "burndown-completion", completionPercent == null ? "N/A" : `${completionPercent}%`);
    if (completionPercent == null || completionPercent === 0) {
      completion.className += " burndown-completion-neutral";
    }
    completion.setAttribute("aria-label", `Completion: ${completion.textContent}`);
    header.append(identity, completion);
    const metrics = burndownElement("dl", "burndown-metrics");
    [["committed", "Committed"], ["tentative", "Tentative"], ["delivered", "Delivered"], ["spillover", "Spill over"]].forEach(([key, label]) => {
      const metric = burndownElement("div", `burndown-metric burndown-metric-${key}`);
      if (Number(record[key] ?? 0) === 0) metric.className += " burndown-metric-zero";
      const value = burndownElement("dd", "", record[key] ?? (["committed", "tentative", "delivered"].includes(key) ? "—" : 0));
      if (key === "spillover") makeSpilloverEditable(value, metric, record);
      metric.append(burndownElement("dt", "", label), value);
      metrics.append(metric);
    });
    const details = burndownElement("dl", "burndown-stories");
    [["inProgress", "In progress"], ["resolved", "Resolved"], ["spillover", "Spillover"], ["tentative", "Tentative"]].forEach(([key, label]) => {
      const row = burndownElement("div", "burndown-story-row");
      const stories = record.stories?.[key] || [];
      const ids = burndownElement("dd", "burndown-story-ids");
      renderSprintStoryIds(ids, stories);
      if (key === "spillover") makeSpilloverStoriesEditable(ids, row, record);
      row.append(burndownElement("dt", `burndown-story-label burndown-story-${key}`, label), ids);
      details.append(row);
    });
    card.append(header, metrics, burndownElement("h3", "burndown-details-title", "User story details"), details);
    container.append(card);
  });
}

function sprintSiteKey(value) {
  const displayed = siteDisplayName(String(value ?? "").trim());
  const key = normalizeColumn(displayed);
  return key === "vestaras" ? "vasteras" : key;
}

let loadedSprintWorkbook = null;
const dynamicSprintSites = new Set(["vasteras", "rakovski", "heidelberg", "senatobia", "alingsas", "chassieu"]);

function sprintRowDate(value, date1904) {
  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value, { date1904 });
    if (!parsed) return null;
    return parseDashboardDate(`${String(parsed.y).padStart(4, "0")}-${String(parsed.m).padStart(2, "0")}-${String(parsed.d).padStart(2, "0")}`);
  }
  if (value instanceof Date) return parseDashboardDate(value.toISOString().slice(0, 10));
  const text = String(value ?? "").trim();
  const iso = /^(\d{4}-\d{2}-\d{2})(?:[T ].*)?$/.exec(text);
  if (iso) return parseDashboardDate(iso[1]);
  const named = /^(\d{1,2})[- ]([a-z]{3})[- ](\d{4})$/i.exec(text);
  if (named) {
    const month = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"].indexOf(named[2].toLowerCase()) + 1;
    return parseDashboardDate(`${named[3]}-${String(month).padStart(2, "0")}-${named[1].padStart(2, "0")}`);
  }
  return null;
}

function readSprintDashboardDates(workbook) {
  const normalize = value => String(value ?? "").trim().replace(/\s+/g, " ").toLowerCase();
  const sheetName = workbook.SheetNames.find(name => normalize(name) === "dashboard");
  if (!sheetName) throw new Error('Sprint dates: worksheet "Dashboard" is missing.');
  const required = ["Site", "Sprint", "Start", "End"];
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "", raw: true });
  const headerIndex = rows.findIndex(row => required.every(header => row.some(cell => normalize(cell) === normalize(header))));
  if (headerIndex < 0) throw new Error(`Sprint dates: Dashboard must contain these headers: ${required.join(", ")}.`);
  const column = header => rows[headerIndex].findIndex(cell => normalize(cell) === normalize(header));
  const siteColumn = column("Site"), sprintColumn = column("Sprint"), startColumn = column("Start"), endColumn = column("End");
  const date1904 = Boolean(workbook.Workbook?.WBProps?.date1904);
  const parseDate = value => {
    // Text dates in this sheet use day-month-year; numeric cells are Excel dates.
    const numericText = /^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/.exec(String(value ?? "").trim());
    return numericText
      ? parseDashboardDate(`${numericText[3]}-${numericText[2].padStart(2, "0")}-${numericText[1].padStart(2, "0")}`)
      : sprintRowDate(value, date1904);
  };
  const datesBySite = new Map(), unavailableSites = new Set();
  for (const row of rows.slice(headerIndex + 1)) {
    const site = sprintSiteKey(String(row[siteColumn] ?? "").trim().split(/[\\/]/).pop());
    if (!dynamicSprintSites.has(site)) continue;
    const start = parseDate(row[startColumn]), end = parseDate(row[endColumn]);
    if (!start || !end || start > end) {
      unavailableSites.add(site);
      continue;
    }
    const dates = { sprint: String(row[sprintColumn] ?? "").trim(), sprintStartDate: start.toISOString().slice(0, 10), sprintEndDate: end.toISOString().slice(0, 10) };
    const previous = datesBySite.get(site);
    // Repeated site rows may share a range; conflicting ranges are ambiguous.
    if (previous && (previous.sprintStartDate !== dates.sprintStartDate || previous.sprintEndDate !== dates.sprintEndDate)) unavailableSites.add(site);
    datesBySite.set(site, dates);
  }
  unavailableSites.forEach(site => datesBySite.delete(site));
  return datesBySite;
}

function updateSprintDashboardDates(workbook) {
  sprintBurndownData.forEach(record => {
    record.sprint = "";
    record.sprintStartDate = null;
    record.sprintEndDate = null;
  });
  try {
    const datesBySite = readSprintDashboardDates(workbook);
    sprintBurndownData.forEach(record => Object.assign(record, datesBySite.get(sprintSiteKey(record.site)) || {}));
    const unavailable = sprintBurndownData.filter(record => !record.sprintStartDate).map(record => record.site);
    return {
      loaded: datesBySite.size > 0,
      message: `Dashboard: loaded sprint dates for ${datesBySite.size} sites.${unavailable.length ? ` Sprint dates unavailable for ${unavailable.join(", ")}: no matching site row, invalid dates, or conflicting ranges.` : ""}`
    };
  } catch (error) {
    return { loaded: false, message: error.message };
  }
}

function readSprintTagCounts(workbook) {
  const normalize = value => String(value ?? "").trim().replace(/\s+/g, " ").toLowerCase();
  const sheetName = workbook.SheetNames.find(name => normalize(name) === "master daily dump");
  if (!sheetName) throw new Error('Sprint Commitment: worksheet "Master Daily Dump" is missing.');
  const required = ["ID", "Area Path", "Date", "Work Item Type", "Assigned To", "State", "Tags"];
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "", raw: true });
  const headerIndex = rows.findIndex(row => required.every(header => row.some(cell => normalize(cell) === normalize(header))));
  if (headerIndex < 0) throw new Error(`Sprint Commitment: Master Daily Dump must contain these headers: ${required.join(", ")}.`);
  const column = header => rows[headerIndex].findIndex(cell => normalize(cell) === normalize(header));
  const datesBySite = readSprintDashboardDates(workbook);
  const rangesBySite = new Map([...datesBySite].map(([site, dates]) => [site, {
    from: parseDashboardDate(dates.sprintStartDate), to: parseDashboardDate(dates.sprintEndDate)
  }]));
  const dateColumn = column("Date");
  const origin = XLSX.utils.decode_range(workbook.Sheets[sheetName]["!ref"]).s;
  const date1904 = Boolean(workbook.Workbook?.WBProps?.date1904);
  const siteColumn = column("Area Path"), typeColumn = column("Work Item Type"), tagColumn = column("Tags"), stateColumn = column("State"), assigneeColumn = column("Assigned To"), idColumn = column("ID");
  const resultsBySite = new Map([...dynamicSprintSites].map(site => [site, {
    committed: rangesBySite.has(site) ? 0 : null, tentative: rangesBySite.has(site) ? 0 : null, delivered: rangesBySite.has(site) ? 0 : null,
    inProgress: new Set(), resolved: new Set(), tentativeIds: new Set(), skippedDates: 0
  }]));
  const developersBySite = new Map();
  // Choose one snapshot date per site, rather than adding daily snapshots together.
  const datedRows = rows.slice(headerIndex + 1).map((row, index) => {
    const cell = workbook.Sheets[sheetName][XLSX.utils.encode_cell({ r: origin.r + headerIndex + 1 + index, c: origin.c + dateColumn })];
    return { row, date: sprintRowDate(cell?.v, date1904) };
  });
  const latestDatesBySite = new Map();
  datedRows.forEach(({ row, date }) => {
    const site = sprintSiteKey(row[siteColumn]);
    const range = rangesBySite.get(site);
    if (!range || !date || date < range.from || date > range.to) return;
    latestDatesBySite.set(site, Math.max(latestDatesBySite.get(site) ?? -Infinity, date.getTime()));
  });
  for (const { row, date } of datedRows) {
    const siteKey = sprintSiteKey(row[siteColumn]);
    const result = resultsBySite.get(siteKey);
    const range = rangesBySite.get(siteKey);
    if (!result || !range) continue;
    if (!date) { result.skippedDates++; continue; }
    if (date < range.from || date > range.to) continue;
    const assignedTo = String(row[assigneeColumn] ?? "").trim();
    if (siteKey && assignedTo) {
      const name = assigneeDisplayName(assignedTo);
      if (!developersBySite.has(siteKey)) developersBySite.set(siteKey, new Map());
      // Repeated daily records should show the assignee only once per card.
      developersBySite.get(siteKey).set(normalize(name), name);
    }
    const type = normalize(row[typeColumn]);
    if (!result || !["user story", "bug"].includes(type)) continue;
    const tags = new Set(String(row[tagColumn] ?? "").split(";").map(normalize));
    const state = normalize(row[stateColumn]);
    const id = String(row[idColumn] ?? "").trim();
    if (["in progress", "active"].includes(state) && id) result.inProgress.add(`${type === "user story" ? "US" : "Bug"} ${id}`);
    if (["resolved", "completed", "done", "closed"].includes(state) && id) result.resolved.add(`${type === "user story" ? "US" : "Bug"} ${id}`);
    if (date.getTime() === latestDatesBySite.get(siteKey)) {
      if (tags.has("tentative") && id) result.tentativeIds.add(`${type === "user story" ? "US" : "Bug"} ${id}`);
      for (const tag of ["committed", "tentative"]) {
        if (tags.has(tag)) result[tag]++;
      }
      if (["closed", "resolved", "done"].includes(state)) result.delivered++;
    }
  }
  return { resultsBySite, developersBySite, latestDatesBySite };
}

function resetSprintTagCounts(clearWorkbook = true) {
  if (clearWorkbook) loadedSprintWorkbook = null;
  sprintBurndownData.forEach(record => {
    record.developer = "";
    if (clearWorkbook) {
      record.sprint = "";
      record.sprintStartDate = null;
      record.sprintEndDate = null;
    }
    if (!dynamicSprintSites.has(sprintSiteKey(record.site))) return;
    record.committed = null;
    record.tentative = null;
    record.delivered = null;
    record.stories.inProgress = [];
    record.stories.resolved = [];
    record.stories.tentative = [];
  });
  renderSprintBurndown(sprintBurndownData);
  document.getElementById("sprint-committed-status").textContent = "Upload an Excel file from Dashboard to load sprint data for all six sites.";
}

function updateSprintTagCounts(workbook) {
  loadedSprintWorkbook = workbook;
  const sprintDates = updateSprintDashboardDates(workbook);
  try {
    const { developersBySite, resultsBySite, latestDatesBySite } = readSprintTagCounts(workbook);
    const messages = [];
    sprintBurndownData.forEach(record => {
      record.developer = [...(developersBySite.get(sprintSiteKey(record.site))?.values() || [])].join(", ");
      const result = resultsBySite.get(sprintSiteKey(record.site));
      if (!result) return;
      const { committed, tentative, delivered, inProgress, resolved, tentativeIds, skippedDates } = result;
      Object.assign(record, { committed, tentative, delivered });
      record.stories.inProgress = [...inProgress];
      record.stories.resolved = [...resolved];
      record.stories.tentative = [...tentativeIds];
      if (committed == null) {
        messages.push(`${record.site}: counts unavailable because Dashboard sprint dates are unavailable.`);
        return;
      }
      const snapshot = latestDatesBySite.get(sprintSiteKey(record.site));
      messages.push(`Loaded ${committed} committed, ${tentative} tentative and ${delivered} delivered User Story / Bug rows for ${record.site}${snapshot == null ? " (no dated snapshot in range)" : ` on ${new Date(snapshot).toISOString().slice(0, 10)}`}.${skippedDates ? ` Skipped ${skippedDates} ${record.site} rows with blank or invalid dates.` : ""}`);
    });
    renderSprintBurndown(sprintBurndownData);
    document.getElementById("sprint-committed-status").textContent = `Master Daily Dump (per-site Dashboard Start�End dates): ${messages.join(" ")} ${sprintDates.message}`;
    return true;
  } catch (error) {
    resetSprintTagCounts(false);
    document.getElementById("sprint-committed-status").textContent = `${error.message} ${sprintDates.message}`;
    return sprintDates.loaded;
  }
}

function refreshSprintDateRange() {
  if (loadedSprintWorkbook) updateSprintTagCounts(loadedSprintWorkbook);
  if (typeof refreshStatusDateRange === "function") refreshStatusDateRange();
  if (typeof refreshBurndownDateRange === "function") refreshBurndownDateRange();
}

resetSprintTagCounts();
