const state = { file: null, roster: null, records: [], previewTimer: null, templateBuffer: null };
const REQUIRED_COLUMNS = ["工号", "常用姓名", "法定姓名"];
const DATE_COLUMNS = ["司龄起算日期", "入职日期"];
const elements = {
  fileInput: document.querySelector("#fileInput"), dropzone: document.querySelector("#dropzone"),
  dropTitle: document.querySelector("#dropTitle"), dropHint: document.querySelector("#dropHint"),
  removeFile: document.querySelector("#removeFile"), startDateInput: document.querySelector("#startDateInput"),
  endDateInput: document.querySelector("#endDateInput"), dateColumn: document.querySelector("#dateColumn"),
  prefixInput: document.querySelector("#prefixInput"), employeeType: document.querySelector("#employeeType"),
  hrNameInput: document.querySelector("#hrNameInput"),
  refreshButton: document.querySelector("#refreshButton"), statusText: document.querySelector("#statusText"),
  matchCount: document.querySelector("#matchCount"), emptyState: document.querySelector("#emptyState"),
  tableWrap: document.querySelector("#tableWrap"), resultBody: document.querySelector("#resultBody"),
  errorBox: document.querySelector("#errorBox"), outputSummary: document.querySelector("#outputSummary"),
  generateButton: document.querySelector("#generateButton"), generateLabel: document.querySelector("#generateLabel"),
};

function setInitialDate() {
  const today = new Date().toISOString().slice(0, 10);
  elements.startDateInput.value = localStorage.getItem("onboarding-start-date") || localStorage.getItem("onboarding-date") || today;
  elements.endDateInput.value = localStorage.getItem("onboarding-end-date") || elements.startDateInput.value;
  elements.prefixInput.value = localStorage.getItem("onboarding-prefix") ?? "E";
  elements.hrNameInput.value = localStorage.getItem("onboarding-hr-name") || "";
}
function showError(message) { elements.errorBox.textContent = message; elements.errorBox.hidden = false; }
function clearError() { elements.errorBox.hidden = true; elements.errorBox.textContent = ""; }
function formatName(record) { return record.legalName && record.preferredName !== record.legalName ? `${record.preferredName}（${record.legalName}）` : record.preferredName; }
function columnIndex(reference = "A") { return [...(reference.match(/[A-Z]+/)?.[0] || "A")].reduce((value, char) => value * 26 + char.charCodeAt(0) - 64, 0) - 1; }
function xmlDocument(text) {
  const document = new DOMParser().parseFromString(text, "application/xml");
  if (document.querySelector("parsererror")) throw new Error("文件中的 XML 内容无法解析。");
  return document;
}
function elementChildren(node, localName) { return [...node.children].filter(child => child.localName === localName); }
function descendants(node, localName) { return [...node.getElementsByTagNameNS("*", localName)]; }
async function readZipText(zip, path) {
  const entry = zip.file(path);
  if (!entry) throw new Error(`文件缺少必要内容：${path}`);
  return entry.async("string");
}

async function parseRoster(file) {
  let zip;
  try { zip = await JSZip.loadAsync(await file.arrayBuffer()); }
  catch { throw new Error("无法读取该文件，请上传有效的 .xlsx 花名册。"); }
  const sharedStrings = [];
  if (zip.file("xl/sharedStrings.xml")) {
    const shared = xmlDocument(await readZipText(zip, "xl/sharedStrings.xml"));
    descendants(shared, "si").forEach(item => sharedStrings.push(item.textContent || ""));
  }
  const workbook = xmlDocument(await readZipText(zip, "xl/workbook.xml"));
  const rels = xmlDocument(await readZipText(zip, "xl/_rels/workbook.xml.rels"));
  const targets = new Map(descendants(rels, "Relationship").map(rel => [rel.getAttribute("Id"), rel.getAttribute("Target")]));
  const candidates = [];
  for (const sheet of descendants(workbook, "sheet")) {
    if (sheet.getAttribute("state") === "hidden") continue;
    const relationId = sheet.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
    let target = (targets.get(relationId) || "").replace(/^\//, "");
    if (!target.startsWith("xl/")) target = `xl/${target}`;
    if (!zip.file(target)) continue;
    const sheetDocument = xmlDocument(await readZipText(zip, target));
    const rows = descendants(sheetDocument, "row").map(row => {
      const cells = new Map();
      elementChildren(row, "c").forEach(cell => {
        const position = columnIndex(cell.getAttribute("r"));
        const type = cell.getAttribute("t");
        let value = "";
        if (type === "inlineStr") value = cell.textContent || "";
        else {
          const raw = elementChildren(cell, "v")[0]?.textContent || "";
          value = type === "s" ? (sharedStrings[Number(raw)] || "") : raw;
        }
        cells.set(position, value);
      });
      const width = cells.size ? Math.max(...cells.keys()) + 1 : 0;
      return Array.from({ length: width }, (_, index) => cells.get(index) || "");
    }).filter(row => row.length);
    if (!rows.length) continue;
    const headers = rows[0].map(value => String(value).trim());
    const score = REQUIRED_COLUMNS.filter(value => headers.includes(value)).length + DATE_COLUMNS.filter(value => headers.includes(value)).length * 2;
    candidates.push({ score, name: sheet.getAttribute("name") || "工作表", headers, rows: rows.slice(1) });
  }
  if (!candidates.length) throw new Error("花名册中没有可读取的工作表。");
  candidates.sort((a, b) => b.score - a.score);
  const selected = candidates[0];
  const missing = REQUIRED_COLUMNS.filter(name => !selected.headers.includes(name));
  if (missing.length) throw new Error(`花名册缺少必要字段：${missing.join("、")}`);
  const dateColumns = DATE_COLUMNS.filter(name => selected.headers.includes(name));
  if (!dateColumns.length) throw new Error("花名册缺少“司龄起算日期”或“入职日期”字段。");
  const index = Object.fromEntries(selected.headers.map((name, position) => [name, position]));
  const normalizeDate = value => {
    const text = String(value ?? "").trim();
    const match = text.match(/^(\d{4})[./-](\d{1,2})[./-](\d{1,2})/);
    if (match) return `${match[1]}-${match[2].padStart(2, "0")}-${match[3].padStart(2, "0")}`;
    const serial = Number(text);
    if (Number.isFinite(serial) && serial >= 20000 && serial <= 80000) return new Date(Date.UTC(1899, 11, 30) + serial * 86400000).toISOString().slice(0, 10);
    return text;
  };
  const records = selected.rows.map(row => {
    const get = name => String(row[index[name]] ?? "").trim();
    return { employeeId: get("工号"), preferredName: get("常用姓名"), legalName: get("法定姓名"), employeeType: get("人员类型"),
      "司龄起算日期": normalizeDate(get("司龄起算日期")), "入职日期": normalizeDate(get("入职日期")) };
  }).filter(record => record.employeeId);
  return { sheetName: selected.name, dateColumns, records };
}

function naturalParts(value) { return value.toUpperCase().split(/(\d+)/).filter(Boolean).map(part => /^\d+$/.test(part) ? Number(part) : part); }
function compareNatural(left, right) {
  const a = naturalParts(left.employeeId), b = naturalParts(right.employeeId);
  for (let index = 0; index < Math.max(a.length, b.length); index++) {
    if (a[index] === b[index]) continue;
    if (a[index] === undefined) return -1;
    if (b[index] === undefined) return 1;
    return a[index] < b[index] ? -1 : 1;
  }
  return 0;
}
function filterRoster() {
  if (!state.roster) return [];
  const start = elements.startDateInput.value, end = elements.endDateInput.value;
  if (start && end && start > end) throw new Error("开始日期不能晚于结束日期。");
  const dateColumn = elements.dateColumn.value, prefix = elements.prefixInput.value.trim().toUpperCase(), employeeType = elements.employeeType.value;
  return state.roster.records.filter(record => {
    const date = record[dateColumn] || "";
    return (!start || date >= start) && (!end || date <= end) && (!prefix || record.employeeId.toUpperCase().startsWith(prefix))
      && (employeeType === "全部" || record.employeeType === employeeType);
  }).map(record => ({ ...record, selectedDate: record[dateColumn] })).sort(compareNatural);
}
function updateSelect(select, values, current) {
  const unique = [...new Set(values)];
  select.replaceChildren(...unique.map(value => { const option = document.createElement("option"); option.value = value; option.textContent = value; return option; }));
  select.value = unique.includes(current) ? current : unique[0];
}
function renderRecords(records) {
  state.records = records; elements.matchCount.textContent = records.length; elements.resultBody.replaceChildren();
  elements.emptyState.hidden = records.length > 0; elements.tableWrap.hidden = records.length === 0;
  elements.generateButton.disabled = records.length === 0 || !elements.startDateInput.value || !elements.endDateInput.value;
  elements.outputSummary.textContent = records.length ? `将生成 ${records.length} 页确认单` : "尚未准备生成";
  records.forEach((record, index) => {
    const compactDate = record.selectedDate.replaceAll("-", ""), row = document.createElement("tr");
    [String(index + 1).padStart(2, "0"), record.employeeId, formatName(record), record.selectedDate.replaceAll("-", "/"), record.employeeType || "—",
      `LD-${compactDate}-${record.employeeId}`].forEach(text => { const cell = document.createElement("td"); cell.textContent = text; row.append(cell); });
    elements.resultBody.append(row);
  });
}
async function setFile(file) {
  if (!file) return;
  if (!file.name.toLowerCase().endsWith(".xlsx")) return showError("请选择 .xlsx 格式的花名册。");
  clearError(); state.file = file; elements.dropzone.classList.add("loaded"); elements.dropTitle.textContent = file.name;
  elements.dropHint.textContent = `${(file.size / 1024).toFixed(1)} KB · 正在本地读取`; elements.refreshButton.disabled = true;
  try {
    state.roster = await parseRoster(file);
    updateSelect(elements.dateColumn, state.roster.dateColumns, elements.dateColumn.value);
    updateSelect(elements.employeeType, ["全部", ...state.roster.records.map(record => record.employeeType).filter(Boolean)], elements.employeeType.value);
    elements.dropHint.textContent = `${(file.size / 1024).toFixed(1)} KB · 已在浏览器本地读取`; preview();
  } catch (error) { state.roster = null; renderRecords([]); showError(error.message); elements.statusText.textContent = "花名册读取失败。"; }
  finally { elements.refreshButton.disabled = !state.roster; }
}
function clearFile(event) {
  event?.preventDefault(); event?.stopPropagation(); state.file = null; state.roster = null; state.records = []; elements.fileInput.value = "";
  elements.dropzone.classList.remove("loaded"); elements.dropTitle.textContent = "拖入花名册，或点击选择";
  elements.dropHint.textContent = "支持 .xlsx，数据仅在当前浏览器内处理"; elements.refreshButton.disabled = true;
  renderRecords([]); elements.statusText.textContent = "选择花名册后，将自动预览匹配人员。"; clearError();
}
function preview() {
  if (!state.roster) return; clearError();
  try { renderRecords(filterRoster()); elements.statusText.textContent = `已读取“${state.roster.sheetName}”工作表，共 ${state.roster.records.length} 条人员记录。`; }
  catch (error) { renderRecords([]); showError(error.message); }
}
function schedulePreview() {
  localStorage.setItem("onboarding-start-date", elements.startDateInput.value); localStorage.setItem("onboarding-end-date", elements.endDateInput.value);
  localStorage.setItem("onboarding-prefix", elements.prefixInput.value); clearTimeout(state.previewTimer); state.previewTimer = setTimeout(preview, 150);
}

function paragraphText(paragraph) { return [...paragraph.getElementsByTagNameNS("*", "t")].map(node => node.textContent || "").join(""); }
function replaceParagraph(paragraph, value) { const nodes = [...paragraph.getElementsByTagNameNS("*", "t")]; nodes[0].textContent = value; nodes.slice(1).forEach(node => node.textContent = ""); }
function replaceAfterLabel(paragraph, label, value) {
  const nodes = [...paragraph.getElementsByTagNameNS("*", "t")], labelIndex = nodes.findIndex(node => (node.textContent || "").includes(label));
  if (labelIndex < 0) throw new Error(`Word 模板缺少字段：${label}`);
  const target = Math.min(labelIndex + 1, nodes.length - 1); nodes[target].textContent = value; nodes.slice(target + 1).forEach(node => node.textContent = "");
}
function replaceDate(paragraph, value) {
  const nodes = [...paragraph.getElementsByTagNameNS("*", "t")]; let target = nodes.findIndex(node => /\d{4}[/.\-]\d{1,2}/.test(node.textContent || ""));
  if (target < 0) target = nodes.length - 1; nodes[target].textContent = value; nodes.slice(target + 1).forEach(node => node.textContent = "");
}
function addPageBreak(document, paragraph) {
  let properties = elementChildren(paragraph, "pPr")[0];
  if (!properties) { properties = document.createElementNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "w:pPr"); paragraph.insertBefore(properties, paragraph.firstChild); }
  properties.append(document.createElementNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "w:pageBreakBefore"));
}
function fillPage(document, nodes, record, pageIndex) {
  const paragraphs = nodes.filter(node => node.localName === "p");
  const contract = paragraphs.find(p => paragraphText(p).includes("劳动合同编号")), name = paragraphs.find(p => paragraphText(p).includes("员工姓名"));
  const employeeId = paragraphs.find(p => paragraphText(p).includes("工号")), date = paragraphs.find(p => paragraphText(p).includes("日期") && paragraphText(p).includes("Date"));
  const hr = paragraphs.find(p => paragraphText(p).includes("Confirmed") && paragraphText(p).includes("HR"));
  if (!contract || !name || !employeeId || !date) throw new Error("Word 模板缺少合同编号、姓名、工号或日期字段。");
  replaceParagraph(contract, `劳动合同编号：LD-${record.selectedDate.replaceAll("-", "")}-${record.employeeId}`);
  replaceAfterLabel(name, "员工姓名", formatName(record)); replaceAfterLabel(employeeId, "工号", record.employeeId); replaceDate(date, record.selectedDate.replaceAll("-", "/"));
  if (hr) {
    const nodes = [...hr.getElementsByTagNameNS("*", "t")];
    const target = nodes.find(node => node.textContent === "待填写");
    if (target) target.textContent = elements.hrNameInput.value.trim();
  }
  if (pageIndex) addPageBreak(document, contract);
  nodes.forEach(node => [...node.getElementsByTagNameNS("*", "bookmarkStart"), ...node.getElementsByTagNameNS("*", "bookmarkEnd")].forEach(bookmark => bookmark.remove()));
}
async function loadTemplate() {
  if (!state.templateBuffer) { const response = await fetch("./template.docx"); if (!response.ok) throw new Error("无法加载确认表模板。"); state.templateBuffer = await response.arrayBuffer(); }
  return state.templateBuffer.slice(0);
}
async function generateDocument(records) {
  const zip = await JSZip.loadAsync(await loadTemplate()), document = xmlDocument(await zip.file("word/document.xml").async("string"));
  const body = descendants(document, "body")[0], children = [...body.children], section = children.at(-1).cloneNode(true), templateNodes = children.slice(0, -1);
  body.replaceChildren();
  records.forEach((record, pageIndex) => { const nodes = templateNodes.map(node => node.cloneNode(true)); fillPage(document, nodes, record, pageIndex); nodes.forEach(node => body.append(node)); });
  body.append(section);
  const xml = new XMLSerializer().serializeToString(document).replace(/^<\?xml[^>]+>/, "");
  zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${xml}`);
  return zip.generateAsync({ type: "blob", compression: "DEFLATE", compressionOptions: { level: 6 }, mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
}
async function generate() {
  if (!state.records.length) return; clearError(); elements.generateButton.disabled = true; elements.generateButton.classList.add("busy"); elements.generateLabel.textContent = "正在生成 Word…";
  try {
    const blob = await generateDocument(state.records), start = elements.startDateInput.value.replaceAll("-", ""), end = elements.endDateInput.value.replaceAll("-", "");
    const period = start === end ? start : `${start}-${end}`, prefix = elements.prefixInput.value.trim().toUpperCase() || "全部";
    const link = document.createElement("a"); link.href = URL.createObjectURL(blob); link.download = `入职材料确认表_${period}_${prefix}_${state.records.length}人.docx`;
    document.body.append(link); link.click(); link.remove(); setTimeout(() => URL.revokeObjectURL(link.href), 1000); elements.outputSummary.textContent = `已生成 ${state.records.length} 页确认单`;
  } catch (error) { showError(error.message); }
  finally { elements.generateButton.disabled = state.records.length === 0; elements.generateButton.classList.remove("busy"); elements.generateLabel.textContent = "生成并下载确认单"; }
}

elements.fileInput.addEventListener("change", event => setFile(event.target.files[0])); elements.removeFile.addEventListener("click", clearFile);
elements.refreshButton.addEventListener("click", preview); elements.generateButton.addEventListener("click", generate);
elements.hrNameInput.addEventListener("input", () => localStorage.setItem("onboarding-hr-name", elements.hrNameInput.value));
[elements.startDateInput, elements.endDateInput, elements.dateColumn, elements.prefixInput, elements.employeeType].forEach(element => { element.addEventListener("change", schedulePreview); if (element.tagName === "INPUT") element.addEventListener("input", schedulePreview); });
["dragenter", "dragover"].forEach(type => elements.dropzone.addEventListener(type, event => { event.preventDefault(); elements.dropzone.classList.add("dragging"); }));
["dragleave", "drop"].forEach(type => elements.dropzone.addEventListener(type, event => { event.preventDefault(); elements.dropzone.classList.remove("dragging"); }));
elements.dropzone.addEventListener("drop", event => setFile(event.dataTransfer.files[0])); setInitialDate();
