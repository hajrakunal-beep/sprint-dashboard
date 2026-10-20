"use strict";

// Patch the original XLSX package. Do not serialize the master through XLSX.write:
// that would discard workbook features not represented in its in-memory model.
function dumpXmlDecode(value) {
  return value.replace(/&(#x[\da-f]+|#\d+|amp|lt|gt|quot|apos);/gi, (_, entity) => {
    if (entity[0] === "#") return String.fromCodePoint(entity[1].toLowerCase() === "x" ? parseInt(entity.slice(2), 16) : Number(entity.slice(1)));
    return { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'" }[entity];
  });
}

function dumpXmlEscape(value) {
  return String(value).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

// Keep source offsets so untouched XML stays exactly as authored by Excel,
// including namespaces, extensions and features the import library cannot read.
function dumpXmlNodes(xml) {
  const roots = [], stack = [];
  const tags = /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<(?:(?:"[^"]*")|(?:'[^']*')|[^'">])*>/g;
  for (const match of xml.matchAll(tags)) {
    const token = match[0];
    if (/^<[!?]/.test(token)) continue;
    if (token.startsWith("</")) {
      const node = stack.pop();
      if (!node || token.slice(2, -1).trim() !== node.tag) throw new Error("Invalid workbook XML.");
      node.closeStart = match.index;
      node.end = match.index + token.length;
      continue;
    }
    const tag = /^<([^\s/>]+)/.exec(token)?.[1];
    if (!tag) throw new Error("Invalid workbook XML tag.");
    const attrs = {};
    for (const attr of token.matchAll(/([^\s=<>]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g)) attrs[attr[1]] = dumpXmlDecode(attr[2] ?? attr[3]);
    const node = { tag, name: tag.split(":").pop(), attrs, start: match.index, openEnd: match.index + token.length, children: [] };
    (stack.length ? stack[stack.length - 1].children : roots).push(node);
    if (/\/\s*>$/.test(token)) { node.end = node.openEnd; node.closeStart = node.openEnd; node.empty = true; }
    else stack.push(node);
  }
  if (stack.length) throw new Error("Incomplete workbook XML.");
  return roots;
}

function dumpXmlSetAttribute(tag, name, value) {
  const escapedName = name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const pattern = new RegExp(`(\\s${escapedName}\\s*=\\s*)(?:"[^"]*"|'[^']*')`);
  if (pattern.test(tag)) return tag.replace(pattern, (_, lead) => `${lead}"${dumpXmlEscape(value)}"`);
  return tag.replace(/\/?\s*>$/, end => ` ${name}="${dumpXmlEscape(value)}"${end}`);
}

function dumpXmlApply(xml, edits) {
  for (const edit of edits.sort((a, b) => b.start - a.start)) xml = xml.slice(0, edit.start) + edit.text + xml.slice(edit.end);
  return xml;
}

function dumpZipPath(base, target) {
  const pieces = (target.startsWith("/") ? target.slice(1) : `${base.slice(0, base.lastIndexOf("/") + 1)}${target}`).split("/");
  const path = [];
  for (const piece of pieces) {
    if (piece === "..") path.pop();
    else if (piece && piece !== ".") path.push(piece);
  }
  return path.join("/");
}

function readAdoDumpWorkbook(bytes, filename = "") {
  if (!/\.csv$/i.test(filename)) return XLSX.read(bytes, { type: "array", cellNF: true, cellStyles: true });
  // Read CSV as UTF-8 text and keep identifiers and date text intact rather
  // than guessing types (which can remove leading zeroes or reinterpret dates).
  const csv = new TextDecoder("utf-8", { fatal: true }).decode(new Uint8Array(bytes)).replace(/^\uFEFF/, "");
  const workbook = XLSX.read(csv, { type: "string", raw: true });
  for (const name of workbook.SheetNames) {
    const sheet = workbook.Sheets[name], headers = findDailyDumpHeaders(sheet);
    if (!headers) continue;
    const numericColumns = headers.columns.slice(9);
    for (const [address, cell] of Object.entries(sheet)) {
      if (address.startsWith("!") || cell.t !== "s") continue;
      const position = XLSX.utils.decode_cell(address);
      if (position.r <= headers.row || !numericColumns.includes(position.c)) continue;
      const value = String(cell.v).trim();
      const numberText = /^[+-]?\d{1,3}(?:,\d{3})+(?:\.\d+)?$/.test(value) ? value.replace(/,/g, "") : value;
      if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(numberText)) continue;
      const number = Number(numberText);
      if (Number.isFinite(number)) { cell.t = "n"; cell.v = number; delete cell.w; }
    }
  }
  return workbook;
}

function mergeDailyDumpPreservingWorkbook(masterBytes, adoBytes, adoFilename = "") {
  const input = new Uint8Array(masterBytes);
  if (input[0] !== 0x50 || input[1] !== 0x4b) throw new Error("To preserve formatting, the master must be an .xlsx workbook. Save the original .xls master as .xlsx in Excel first.");
  const zip = XLSX.CFB.read(input, { type: "buffer" });
  const decoder = new TextDecoder("utf-8", { fatal: true }), encoder = new TextEncoder();
  const entry = path => XLSX.CFB.find(zip, `/${path}`);
  const read = path => {
    const file = entry(path);
    if (!file) throw new Error(`Missing workbook part: ${path}.`);
    return decoder.decode(new Uint8Array(file.content));
  };
  const replace = (path, xml) => {
    const file = entry(path);
    file.content = encoder.encode(xml);
    file.size = file.content.length;
  };
  const relationships = path => dumpXmlNodes(read(path))[0].children.filter(node => node.name === "Relationship");
  const workbookPath = dumpZipPath("", relationships("_rels/.rels").find(node => /\/officeDocument$/.test(node.attrs.Type))?.attrs.Target || "xl/workbook.xml");
  const workbookXml = dumpXmlNodes(read(workbookPath))[0];
  const sheetNode = workbookXml.children.find(node => node.name === "sheets")?.children.find(node => normalizeDumpHeader(node.attrs.name) === "master daily dump");
  if (!sheetNode) throw new Error('The master workbook must contain a sheet named "Master Daily Dump".');
  const relPath = `${workbookPath.slice(0, workbookPath.lastIndexOf("/") + 1)}_rels/${workbookPath.split("/").pop()}.rels`;
  const relationId = Object.entries(sheetNode.attrs).find(([name]) => name.endsWith(":id"))?.[1];
  const relation = relationships(relPath).find(node => node.attrs.Id === relationId);
  if (!relation || relation.attrs.TargetMode === "External") throw new Error("Cannot locate the Master Daily Dump worksheet.");
  const sheetPath = dumpZipPath(workbookPath, relation.attrs.Target);
  const xml = read(sheetPath), root = dumpXmlNodes(xml)[0];
  const data = root.children.find(node => node.name === "sheetData");
  if (!data) throw new Error("Master Daily Dump has no worksheet data section.");
  const options = { type: "array", cellNF: true, cellStyles: true };
  const master = XLSX.read(masterBytes, options);
  const target = master.Sheets[sheetNode.attrs.name];
  const header = findDailyDumpHeaders(target);
  const result = mergeDailyDumpWorkbooks(master, readAdoDumpWorkbook(adoBytes, adoFilename));
  const start = result.firstRow, end = start + result.appendedRows - 1;
  const prefix = data.tag.includes(":") ? `${data.tag.split(":")[0]}:` : "";
  const oldRows = data.children.filter(node => node.name === "row");
  const oldRowsByNumber = new Map(oldRows.map(row => [Number(row.attrs.r), row]));
  const templateStyles = new Map();
  let templateRow = null;
  for (const row of oldRows) {
    const number = Number(row.attrs.r);
    if (number <= header.row + 1 || number >= start) continue;
    templateRow = row;
    for (const cell of row.children.filter(node => node.name === "c")) {
      if (cell.attrs.s != null) templateStyles.set(XLSX.utils.decode_cell(cell.attrs.r).c, cell.attrs.s);
    }
  }
  const makeCell = (number, column, previous) => {
    const address = XLSX.utils.encode_cell({ r: number - 1, c: column });
    const cell = target[address];
    const style = previous?.attrs.s ?? templateStyles.get(column);
    const attrs = ` r="${address}"${style != null ? ` s="${dumpXmlEscape(style)}"` : ""}`;
    if (!cell || cell.v == null || cell.t === "z") return `<${prefix}c${attrs}/>`;
    if (cell.t === "n") {
      if (!Number.isFinite(cell.v)) throw new Error(`Invalid number at ${address}.`);
      return `<${prefix}c${attrs}><${prefix}v>${cell.v}</${prefix}v></${prefix}c>`;
    }
    if (cell.t === "b" || cell.t === "e") return `<${prefix}c${attrs} t="${cell.t}"><${prefix}v>${dumpXmlEscape(cell.t === "b" ? (cell.v ? 1 : 0) : XLSX.utils.format_cell(cell))}</${prefix}v></${prefix}c>`;
    const value = String(cell.v).replace(/_x[\da-f]{4}_/gi, match => `_x005F_${match.slice(1)}`).replace(/[\x00-\x08\x0b\x0c\x0e-\x1f]/g, char => `_x${char.charCodeAt(0).toString(16).padStart(4, "0")}_`);
    return `<${prefix}c${attrs} t="inlineStr"><${prefix}is><${prefix}t xml:space="preserve">${dumpXmlEscape(value)}</${prefix}t></${prefix}is></${prefix}c>`;
  };
  const rowXml = new Map();
  for (let number = start; number <= end; number++) {
    const existing = oldRowsByNumber.get(number);
    const cells = new Map();
    for (const cell of existing?.children.filter(node => node.name === "c") || []) cells.set(XLSX.utils.decode_cell(cell.attrs.r).c, xml.slice(cell.start, cell.end));
    for (const column of result.columns) {
      const previous = existing?.children.find(node => node.name === "c" && XLSX.utils.decode_cell(node.attrs.r).c === column);
      cells.set(column, makeCell(number, column, previous));
    }
    const content = [...cells].sort(([a], [b]) => a - b).map(([, text]) => text).join("");
    let open = existing ? xml.slice(existing.start, existing.openEnd).replace(/\/\s*>$/, ">") : `<${prefix}row r="${number}">`;
    if (!existing && templateRow) {
      for (const name of ["ht", "customHeight", "s", "customFormat"]) {
        if (templateRow.attrs[name] != null) open = dumpXmlSetAttribute(open, name, templateRow.attrs[name]);
      }
    }
    // A preformatted row's span may have been narrower than the data columns.
    if (/\sspans=/.test(open)) open = dumpXmlSetAttribute(open, "spans", `${Math.min(...result.columns) + 1}:${Math.max(...result.columns) + 1}`);
    rowXml.set(number, open + content + (existing?.children.filter(node => node.name !== "c").map(node => xml.slice(node.start, node.end)).join("") || "") + `</${prefix}row>`);
  }
  const dataEdits = [];
  for (const row of oldRows) {
    const replacement = rowXml.get(Number(row.attrs.r));
    if (replacement) { dataEdits.push({ start: row.start, end: row.end, text: replacement }); rowXml.delete(Number(row.attrs.r)); }
  }
  let followingIndex = 0;
  const insertionEdits = new Map();
  for (const [number, text] of rowXml) {
    while (followingIndex < oldRows.length && Number(oldRows[followingIndex].attrs.r) <= number) followingIndex++;
    const following = oldRows[followingIndex];
    const position = following ? following.start : data.closeStart;
    const edit = insertionEdits.get(position);
    if (edit) edit.text += text;
    else { const insertion = { start: position, end: position, text }; insertionEdits.set(position, insertion); dataEdits.push(insertion); }
  }
  if (data.empty) throw new Error("Master Daily Dump has no header rows.");
  const edits = [...dataEdits];
  for (const node of root.children) {
    if (node.name === "dimension" && node.attrs.ref) {
      edits.push({ start: node.start, end: node.openEnd, text: dumpXmlSetAttribute(xml.slice(node.start, node.openEnd), "ref", target["!ref"]) });
    } else if (node.name === "autoFilter" && node.attrs.ref) {
      const range = XLSX.utils.decode_range(node.attrs.ref); range.e.r = Math.max(range.e.r, end - 1);
      edits.push({ start: node.start, end: node.openEnd, text: dumpXmlSetAttribute(xml.slice(node.start, node.openEnd), "ref", XLSX.utils.encode_range(range)) });
    }
    // Extend rules covering the old final data row; leave other ranges unchanged.
    const rules = node.name === "conditionalFormatting" ? [node] : node.name === "dataValidations" ? node.children : [];
    for (const rule of rules) {
      if (!rule.attrs.sqref) continue;
      const sqref = rule.attrs.sqref.split(/\s+/).map(ref => {
        const range = XLSX.utils.decode_range(ref.replace(/\$/g, ""));
        if (range.e.r === start - 2 && range.s.r > header.row) { range.e.r = end - 1; return XLSX.utils.encode_range(range); }
        return ref;
      }).join(" ");
      if (sqref !== rule.attrs.sqref) edits.push({ start: rule.start, end: rule.openEnd, text: dumpXmlSetAttribute(xml.slice(rule.start, rule.openEnd), "sqref", sqref) });
    }
  }
  // Excel tables have a separate package part. Only extend tables belonging to
  // the target worksheet, preserving their columns, styles and all other XML.
  const tableParts = root.children.find(node => node.name === "tableParts");
  if (tableParts) {
    const sheetRels = `${sheetPath.slice(0, sheetPath.lastIndexOf("/") + 1)}_rels/${sheetPath.split("/").pop()}.rels`;
    const rels = relationships(sheetRels);
    for (const part of tableParts.children) {
      const id = Object.entries(part.attrs).find(([name]) => name.endsWith(":id"))?.[1];
      const rel = rels.find(node => node.attrs.Id === id);
      if (!rel) throw new Error("Cannot locate a Master Daily Dump table.");
      const path = dumpZipPath(sheetPath, rel.attrs.Target), tableXml = read(path), table = dumpXmlNodes(tableXml)[0];
      const range = XLSX.utils.decode_range(table.attrs.ref);
      if (range.e.r !== start - 2 || range.s.r !== header.row) continue;
      if (Number(table.attrs.totalsRowCount || 0)) throw new Error("Remove the totals row from the Master Daily Dump table before merging.");
      range.e.r = end - 1;
      const ref = XLSX.utils.encode_range(range), changes = [{ start: table.start, end: table.openEnd, text: dumpXmlSetAttribute(tableXml.slice(table.start, table.openEnd), "ref", ref) }];
      const filter = table.children.find(node => node.name === "autoFilter");
      if (filter) changes.push({ start: filter.start, end: filter.openEnd, text: dumpXmlSetAttribute(tableXml.slice(filter.start, filter.openEnd), "ref", ref) });
      replace(path, dumpXmlApply(tableXml, changes));
    }
  }
  replace(sheetPath, dumpXmlApply(xml, edits));
  const bytes = XLSX.CFB.write(zip, { type: "buffer", fileType: "zip", compression: true });
  return { bytes: new Uint8Array(bytes), appendedRows: result.appendedRows, firstRow: result.firstRow };
}
