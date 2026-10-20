"use strict";

function calculateBacklogUpdates(workbook) {
  const normalize = value => String(value ?? "").trim().replace(/\s+/g, " ").toLowerCase();
  // Use the same known site aliases as the dashboard's imported data.
  const siteAliases = new Map([
    ["mes alingsas", "alingsas"], ["mes fr chassieu", "chassieu"],
    ["mes se vasteras", "vasteras"], ["mes_de_heidelberg", "heidelberg"],
    ["mes-rakovski", "rakovski"], ["mes-senatobia", "senatobia"]
  ]);
  const siteKey = value => siteAliases.get(normalize(value)) || normalize(value);
  const locate = (name, headers) => {
    const key = workbook.SheetNames.find(item => normalize(item) === normalize(name));
    if (!key) throw new Error(`Missing worksheet: ${name}.`);
    const sheet = workbook.Sheets[key];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "", blankrows: true });
    const origin = XLSX.utils.decode_range(sheet["!ref"] || "A1").s;
    for (let index = 0; index < rows.length; index++) {
      const labels = rows[index].map(normalize);
      if (!headers.every(header => labels.includes(normalize(header)))) continue;
      if (headers.some(header => labels.filter(label => label === normalize(header)).length !== 1)) throw new Error(`${name} has duplicate required headers.`);
      const columns = headers.map(header => labels.indexOf(normalize(header)));
      return { key, sheet, columns: columns.map(column => column + origin.c), records: rows.slice(index + 1).map((row, offset) => ({ row: origin.r + index + offset + 1, values: columns.map(column => row[column] ?? "") })) };
    }
    throw new Error(`${name} requires headers: ${headers.join(", ")}.`);
  };
  const dashboard = locate("Dashboard", ["Site", "Sprint", "Start", "End"]);
  const backlog = locate("Consolidated Backlog", ["Site", "ID", "Sprint No", "User story / task", "Estm. Hrs", "Com. Hrs", "Rem. Hrs"]);
  const dump = locate("Master Daily Dump", ["ID", "Area Path", "Date", "Work Item Type", "Parent", "Original Estimate", "Remaining Work", "Completed Work"]);
  const day = value => {
    if (typeof value === "number") {
      const parsed = XLSX.SSF.parse_date_code(value, { date1904: Boolean(workbook.Workbook?.WBProps?.date1904) });
      return parsed ? Date.UTC(parsed.y, parsed.m - 1, parsed.d) : NaN;
    }
    const text = String(value ?? "").trim();
    let year, month, date;
    let match = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s].*)?$/.exec(text);
    if (match) [, year, month, date] = match;
    else if ((match = /^(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{4})(?:\s.*)?$/.exec(text))) [, date, month, year] = match;
    else if ((match = /^(\d{1,2})[\s-]([a-z]+)[\s-](\d{4})$/i.exec(text))) {
      date = match[1]; year = match[3]; month = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"].indexOf(match[2].slice(0, 3).toLowerCase()) + 1;
    } else return NaN;
    const result = new Date(Date.UTC(Number(year), Number(month) - 1, Number(date)));
    return result.getUTCFullYear() === Number(year) && result.getUTCMonth() + 1 === Number(month) && result.getUTCDate() === Number(date) ? result.getTime() : NaN;
  };
  const ranges = new Map();
  const rangeKey = (site, sprint) => JSON.stringify([siteKey(site), normalize(sprint)]);
  for (const { values: [site, sprint, start, end] } of dashboard.records) {
    if (normalize(site) === "frosinone") continue;
    if (!normalize(site) || !normalize(sprint)) continue;
    const from = day(start), to = day(end), key = rangeKey(site, sprint);
    if (!Number.isFinite(from) || !Number.isFinite(to) || from > to) throw new Error(`Invalid Dashboard dates for ${site}, sprint ${sprint}.`);
    const existing = ranges.get(key);
    if (existing && (existing.from !== from || existing.to !== to)) throw new Error(`Conflicting Dashboard date ranges for ${site}, sprint ${sprint}.`);
    ranges.set(key, { from, to });
  }
  const tasks = new Map();
  for (const { row, values: [taskId, area, date, type, parent, estimate, remaining, completed] } of dump.records) {
    if (normalize(type) !== "task" || !normalize(parent)) continue;
    const timestamp = day(date);
    const key = normalize(parent);
    if (!tasks.has(key)) tasks.set(key, []);
    tasks.get(key).push({ id: normalize(taskId), row: row + 1, area: siteKey(area), timestamp, hours: [estimate, completed, remaining] });
  }
  const updates = [], noMatchDetails = [];
  let updatedRows = 0, zeroRows = 0, duplicateTasks = 0;
  for (const { row, values: [site, id, sprint, type] } of backlog.records) {
    if (normalize(site) === "frosinone") continue;
    if (!["user story", "bug"].includes(normalize(type)) || !normalize(id)) continue;
    const range = ranges.get(rangeKey(site, sprint));
    const parentTasks = tasks.get(normalize(id)) || [];
    const siteTasks = siteKey(site) ? parentTasks.filter(task => task.area.includes(siteKey(site))) : [];
    const datedTasks = siteTasks.filter(task => Number.isFinite(task.timestamp));
    const matches = range ? datedTasks.filter(task => task.timestamp >= range.from && task.timestamp <= range.to) : [];
    const latest = matches.reduce((max, task) => Math.max(max, task.timestamp), -Infinity);
    const totals = [0, 0, 0];
    const countedTasks = new Map();
    for (const task of matches) {
      if (task.timestamp !== latest) continue;
      const hours = task.hours.map(value => {
        const number = value === "" ? 0 : Number(String(value).replace(/,/g, ""));
        if (!Number.isFinite(number)) throw new Error(`Invalid task hours for parent ${id}.`);
        return number;
      });
      // Daily exports can be appended more than once. Count each identified
      // task once on the selected date, without combining conflicting snapshots.
      const taskKey = task.id ? `id:${task.id}` : `row:${task.row}`;
      const previous = countedTasks.get(taskKey);
      if (previous) {
        if (hours.some((value, index) => value !== previous.hours[index])) throw new Error(`Task ${task.id} for Parent ${id} has conflicting hours on ${new Date(latest).toISOString().slice(0, 10)} in Master Daily Dump rows ${previous.row} and ${task.row}. Correct the conflicting rows and retry.`);
        duplicateTasks++;
        continue;
      }
      countedTasks.set(taskKey, { hours, row: task.row });
      hours.forEach((number, index) => { totals[index] += number; });
    }
    if (!matches.length) {
      zeroRows++;
      let reason;
      if (!range) reason = `No Dashboard row matches Site "${site}" and Sprint "${sprint}".`;
      else if (!parentTasks.length) reason = `No Task rows have Parent "${id}".`;
      else if (!siteTasks.length) reason = `Tasks for Parent "${id}" exist, but their Area Path does not contain site "${site}".`;
      else if (!datedTasks.length) reason = `Matching tasks exist, but their Date values could not be read.`;
      else {
        const dates = [...new Set(datedTasks.map(task => new Date(task.timestamp).toISOString().slice(0, 10)))].sort();
        reason = `Task dates ${dates[0]} to ${dates[dates.length - 1]} fall outside Dashboard range ${new Date(range.from).toISOString().slice(0, 10)} to ${new Date(range.to).toISOString().slice(0, 10)}.`;
      }
      noMatchDetails.push(`Backlog row ${row + 1}, ID ${id}: ${reason}`);
    }
    totals.forEach((value, index) => updates.push({ address: XLSX.utils.encode_cell({ r: row, c: backlog.columns[index + 4] }), value }));
    updatedRows++;
  }
  return { sheetName: backlog.key, updates, updatedRows, zeroRows, noMatchDetails, duplicateTasks };
}

// Change only the three output cells per story/bug in the original ZIP package.
function updateBacklogPreservingWorkbook(bytes) {
  const input = new Uint8Array(bytes);
  if (input[0] !== 0x50 || input[1] !== 0x4b) throw new Error("Select an .xlsx workbook to preserve its formatting.");
  const result = calculateBacklogUpdates(XLSX.read(bytes, { type: "array" }));
  const zip = XLSX.CFB.read(input, { type: "buffer" });
  const decoder = new TextDecoder(), encoder = new TextEncoder();
  const read = path => decoder.decode(new Uint8Array(XLSX.CFB.find(zip, `/${path}`).content));
  const relations = path => dumpXmlNodes(read(path))[0].children;
  const workbookPath = dumpZipPath("", relations("_rels/.rels").find(node => /\/officeDocument$/.test(node.attrs.Type))?.attrs.Target || "xl/workbook.xml");
  const workbook = dumpXmlNodes(read(workbookPath))[0];
  const sheet = workbook.children.find(node => node.name === "sheets").children.find(node => node.attrs.name === result.sheetName);
  const relId = Object.entries(sheet.attrs).find(([name]) => name.endsWith(":id"))[1];
  const relPath = `${workbookPath.slice(0, workbookPath.lastIndexOf("/") + 1)}_rels/${workbookPath.split("/").pop()}.rels`;
  const relation = relations(relPath).find(node => node.attrs.Id === relId);
  const sheetPath = dumpZipPath(workbookPath, relation.attrs.Target);
  const xml = read(sheetPath), root = dumpXmlNodes(xml)[0];
  const data = root.children.find(node => node.name === "sheetData");
  const edits = [], byRow = new Map();
  for (const update of result.updates) {
    const row = XLSX.utils.decode_cell(update.address).r + 1;
    if (!byRow.has(row)) byRow.set(row, []);
    byRow.get(row).push(update);
  }
  for (const row of data.children.filter(node => node.name === "row")) {
    const updates = byRow.get(Number(row.attrs.r));
    if (!updates) continue;
    const cells = new Map(row.children.filter(node => node.name === "c").map(cell => [cell.attrs.r, xml.slice(cell.start, cell.end)]));
    for (const { address, value } of updates) {
      const old = row.children.find(cell => cell.name === "c" && cell.attrs.r === address);
      const tag = old?.tag || (row.tag.includes(":") ? `${row.tag.split(":")[0]}:c` : "c");
      const prefix = tag.includes(":") ? `${tag.split(":")[0]}:` : "";
      const attrs = { ...old?.attrs, r: address, t: "n" };
      cells.set(address, `<${tag}${Object.entries(attrs).map(([name, text]) => ` ${name}="${dumpXmlEscape(String(text))}"`).join("")}><${prefix}v>${value}</${prefix}v></${tag}>`);
    }
    const content = [...cells].sort(([a], [b]) => XLSX.utils.decode_cell(a).c - XLSX.utils.decode_cell(b).c).map(([, text]) => text).join("");
    const extra = row.children.filter(node => node.name !== "c").map(node => xml.slice(node.start, node.end)).join("");
    edits.push({ start: row.start, end: row.end, text: xml.slice(row.start, row.openEnd).replace(/\/\s*>$/, ">") + content + extra + `</${row.tag}>` });
  }
  const entry = XLSX.CFB.find(zip, `/${sheetPath}`);
  entry.content = encoder.encode(dumpXmlApply(xml, edits)); entry.size = entry.content.length;
  return { ...result, bytes: new Uint8Array(XLSX.CFB.write(zip, { type: "buffer", fileType: "zip", compression: true })) };
}

if (typeof document !== "undefined") (() => {
  const input = document.getElementById("backlog-excel-file");
  const button = document.getElementById("update-backlog-button");
  const status = document.getElementById("backlog-update-status");
  let handle = null;
  let updating = false, selecting = false;
  const canSave = typeof window !== "undefined" && typeof window.showOpenFilePicker === "function";
  input.addEventListener("click", async () => {
    if (!canSave || updating || selecting) return;
    selecting = true; input.disabled = true; button.disabled = true;
    try {
      const [selected] = await window.showOpenFilePicker({ multiple: false, excludeAcceptAllOption: true, types: [{ description: "Excel workbook", accept: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"] } }] });
      const file = await selected.getFile();
      if (!/\.xlsx$/i.test(file.name)) throw new Error("Select an .xlsx workbook.");
      handle = selected;
      status.classList.remove("settings-file-error");
      status.textContent = `Selected: ${file.name}. Update Backlog will save directly to this file.`;
    } catch (error) {
      if (error.name !== "AbortError") {
        status.classList.add("settings-file-error");
        status.textContent = `Could not select backlog file: ${error.message}`;
      }
    } finally {
      selecting = false; input.disabled = false; button.disabled = false;
    }
  });
  button.addEventListener("click", async () => {
    if (updating || selecting || !canSave) return;
    if (!handle) {
      status.classList.add("settings-file-error");
      status.textContent = "Choose an .xlsx backlog workbook before clicking Update Backlog.";
      return;
    }
    updating = true;
    button.disabled = true; input.disabled = true;
    button.textContent = "Updating…";
    status.classList.remove("settings-file-error");
    status.textContent = "Updating backlog hours…";
    let writable;
    try {
      const permission = await handle.requestPermission({ mode: "readwrite" });
      if (permission !== "granted") throw new Error("Write access was not granted. Allow access to update the selected workbook.");
      // Let the browser paint progress before synchronous workbook processing.
      await new Promise(resolve => setTimeout(resolve, 0));
      if (typeof XLSX === "undefined" || typeof dumpXmlNodes !== "function") throw new Error("Required workbook scripts did not load. Refresh the page with Ctrl+Shift+R and try again.");
      // Always read the latest saved workbook, including changes since selection.
      const file = await handle.getFile();
      const result = updateBacklogPreservingWorkbook(await file.arrayBuffer());
      writable = await handle.createWritable();
      await writable.write(result.bytes);
      await writable.close();
      writable = null;
      const details = result.noMatchDetails.slice(0, 20);
      if (result.duplicateTasks) details.unshift(`Ignored ${result.duplicateTasks} repeated Task rows with the same ID, date, and hours.`);
      if (result.noMatchDetails.length > 20) details.push(`${result.noMatchDetails.length - 20} additional rows had no matches.`);
      status.textContent = [`Updated ${result.updatedRows} User Story/Bug rows; ${result.zeroRows} had no matches and were set to zero. Saved directly to ${file.name}.`, ...details].join("\n");
    } catch (error) {
      if (writable) { try { await writable.abort(); } catch (_) { /* Stream may already be closed. */ } }
      status.classList.add("settings-file-error");
      status.textContent = `Update failed: ${error.message || "Unable to read workbook."}`;
    } finally {
      updating = false; button.disabled = false; input.disabled = false; button.textContent = "Update Backlog";
    }
  });
  input.disabled = !canSave;
  button.disabled = !canSave;
  if (!canSave) status.textContent = "Direct file saving is unavailable. Open this dashboard in desktop Chrome or Edge using HTTPS or localhost.";
  if (typeof window !== "undefined") window.backlogUpdateReady = true;
})();
