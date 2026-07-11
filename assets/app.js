"use strict";

const state = {
  meta: null,
  profiles: null,
  registers: null,
  selectedYear: null,
  incidentCache: new Map(),
  currentRows: [],
  filteredRows: [],
  page: 1,
  pageSize: 25,
  activeView: "dashboard",
};

const el = id => document.getElementById(id);
const fmt = new Intl.NumberFormat("th-TH");
const fiscalMonths = [10, 11, 12, 1, 2, 3, 4, 5, 6, 7, 8, 9];
const monthLabels = {1:"ม.ค.",2:"ก.พ.",3:"มี.ค.",4:"เม.ย.",5:"พ.ค.",6:"มิ.ย.",7:"ก.ค.",8:"ส.ค.",9:"ก.ย.",10:"ต.ค.",11:"พ.ย.",12:"ธ.ค."};

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function normalize(value) {
  return String(value ?? "").toLocaleLowerCase("th-TH").replace(/\s+/g, " ").trim();
}

function getCol(name) {
  return state.meta.columns[name];
}

function valueAt(item, name) {
  return item.row[getCol(name)];
}

function riskCode(value) {
  const text = String(value ?? "").trim();
  const match = text.match(/^([A-Za-z]{2,5}\d{2,4})/);
  return match ? match[1].toUpperCase() : (text.split(":", 1)[0] || "ไม่ระบุ");
}

function riskType(value) {
  const code = riskCode(value);
  if (code.startsWith("C")) return "Clinical";
  if (code.startsWith("G")) return "Non-clinical";
  return "Other";
}

function isHighSeverity(value) {
  const key = String(value ?? "").trim().toUpperCase();
  if (["E","F","G","H","I"].includes(key)) return true;
  const n = Number(key);
  return Number.isFinite(n) && n >= 3;
}

function formatThaiDate(value) {
  if (!value) return "–";
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(value))) return String(value);
  const [y, m, d] = value.split("-").map(Number);
  return `${d} ${monthLabels[m]} ${y + 543}`;
}

function fiscalMonthFromDate(value) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(value ?? ""))) return "";
  return Number(String(value).slice(5, 7));
}

async function fetchJson(path) {
  const response = await fetch(path, {cache: "no-store"});
  if (!response.ok) throw new Error(`โหลด ${path} ไม่สำเร็จ (${response.status})`);
  return response.json();
}

async function loadYear(year) {
  const numericYear = Number(year);
  if (state.incidentCache.has(numericYear)) return state.incidentCache.get(numericYear);
  const data = await fetchJson(`data/incidents_${numericYear}.json`);
  state.incidentCache.set(numericYear, data);
  return data;
}

async function getSelectedItems() {
  if (state.selectedYear === "all") {
    const sets = await Promise.all(state.meta.years.map(loadYear));
    return sets.flatMap(set => set.rows.map(row => ({fy: set.fiscalYear, row})));
  }
  const set = await loadYear(state.selectedYear);
  return set.rows.map(row => ({fy: set.fiscalYear, row}));
}

function getSummary(year) {
  return state.meta.summaries.find(item => item.fiscalYear === Number(year));
}

function combineSummaries() {
  const list = state.meta.summaries;
  return {
    fiscalYear: "all",
    total: list.reduce((s, x) => s + x.total, 0),
    clinical: list.reduce((s, x) => s + x.clinical, 0),
    nonClinical: list.reduce((s, x) => s + x.nonClinical, 0),
    highSeverity: list.reduce((s, x) => s + x.highSeverity, 0),
    uniqueUnits: "–",
    monthly: [], severity: [], topRisks: [], topUnits: []
  };
}

function renderKpis(summary) {
  const cards = [
    ["อุบัติการณ์ทั้งหมด", summary.total, "รายการในชุดข้อมูล", "blue"],
    ["Clinical", summary.clinical, `${((summary.clinical / Math.max(summary.total, 1))*100).toFixed(1)}% ของทั้งหมด`, "teal"],
    ["Non-clinical", summary.nonClinical, `${((summary.nonClinical / Math.max(summary.total, 1))*100).toFixed(1)}% ของทั้งหมด`, "orange"],
    ["ระดับรุนแรงสูง", summary.highSeverity, "ระดับ E–I หรือ 3–5", "red"],
    ["หน่วยงานรายงาน", summary.uniqueUnits, summary.uniqueUnits === "–" ? "ดูแยกรายปี" : "หน่วยงานที่มีข้อมูล", "green"],
  ];
  el("kpiGrid").innerHTML = cards.map(([label, value, note, cls]) => `
    <article class="kpi ${cls}">
      <div class="label">${escapeHtml(label)}</div>
      <div class="value">${typeof value === "number" ? fmt.format(value) : escapeHtml(value)}</div>
      <div class="note">${escapeHtml(note)}</div>
    </article>`).join("");
}

function renderBarChart(targetId, data, maxItems = 12) {
  const target = el(targetId);
  const items = (data || []).slice(0, maxItems);
  if (!items.length) {
    target.innerHTML = `<div class="chart-empty">เลือกปีงบประมาณเพื่อดูข้อมูล</div>`;
    return;
  }
  const max = Math.max(...items.map(x => Number(x.count) || 0), 1);
  target.innerHTML = `<div class="bar-chart">${items.map(item => `
    <div class="bar-row" title="${escapeHtml(item.label)}: ${fmt.format(item.count)}">
      <div class="bar-label">${escapeHtml(item.label)}</div>
      <div class="bar-track"><div class="bar-fill" style="width:${Math.max(1, (item.count/max)*100)}%"></div></div>
      <div class="bar-value">${fmt.format(item.count)}</div>
    </div>`).join("")}</div>`;
}

function renderYearComparison() {
  const list = state.meta.summaries;
  const max = Math.max(...list.map(x => x.total), 1);
  el("yearComparisonChart").innerHTML = `<div class="year-bars">${list.map(item => `
    <div class="year-bar-item" title="ปี ${item.fiscalYear}: ${fmt.format(item.total)} รายการ">
      <div class="year-bar-value">${fmt.format(item.total)}</div>
      <div class="year-bar" style="height:${Math.max(4, (item.total/max)*205)}px"></div>
      <div class="year-bar-label">พ.ศ. ${item.fiscalYear}</div>
    </div>`).join("")}</div>`;
}

function renderDashboard() {
  const summary = state.selectedYear === "all" ? combineSummaries() : getSummary(state.selectedYear);
  const title = state.selectedYear === "all" ? "ภาพรวมทุกปีงบประมาณ" : `ภาพรวมปีงบประมาณ ${state.selectedYear}`;
  el("selectedYearTitle").innerHTML = `<div><h2>${title}</h2><p>ปีงบประมาณกำหนดจากไฟล์ต้นทาง ไม่คำนวณจากปีของวันที่</p></div>`;
  renderKpis(summary);
  renderYearComparison();
  renderBarChart("monthlyChart", summary.monthly || [], 12);
  renderBarChart("severityChart", summary.severity || [], 12);
  renderBarChart("riskChart", summary.topRisks || [], 10);
  renderBarChart("unitChart", summary.topUnits || [], 10);
}

function populateSelect(select, values, placeholder = "ทั้งหมด") {
  const current = select.value;
  select.innerHTML = `<option value="">${escapeHtml(placeholder)}</option>` + values.map(v => `<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join("");
  if (values.map(String).includes(current)) select.value = current;
}

function rebuildIncidentFilters() {
  const titles = state.currentRows.map(item => valueAt(item, "รหัส: เรื่องอุบัติการณ์"));
  const types = [...new Set(titles.map(riskType))].sort();
  const severities = [...new Set(state.currentRows.map(item => String(valueAt(item, "ความรุนแรง") ?? "ไม่ระบุ")))].sort((a,b) => a.localeCompare(b,"th",{numeric:true}));
  const units = [...new Set(state.currentRows.map(item => String(valueAt(item, "หน่วยงานที่บันทึกรายงาน") ?? "ไม่ระบุ")))].sort((a,b) => a.localeCompare(b,"th"));
  populateSelect(el("typeFilter"), types);
  populateSelect(el("severityFilter"), severities);
  populateSelect(el("unitFilter"), units);
  el("monthFilter").innerHTML = `<option value="">ทั้งหมด</option>` + fiscalMonths.map(m => `<option value="${m}">${monthLabels[m]}</option>`).join("");
}

function applyIncidentFilters() {
  const query = normalize(el("searchInput").value);
  const type = el("typeFilter").value;
  const severity = el("severityFilter").value;
  const unit = el("unitFilter").value;
  const month = el("monthFilter").value;
  const dateIdx = getCol("วันที่เกิดอุบัติการณ์");
  const titleIdx = getCol("รหัส: เรื่องอุบัติการณ์");
  const severityIdx = getCol("ความรุนแรง");
  const unitIdx = getCol("หน่วยงานที่บันทึกรายงาน");

  state.filteredRows = state.currentRows.filter(item => {
    const row = item.row;
    if (type && riskType(row[titleIdx]) !== type) return false;
    if (severity && String(row[severityIdx] ?? "ไม่ระบุ") !== severity) return false;
    if (unit && String(row[unitIdx] ?? "ไม่ระบุ") !== unit) return false;
    if (month && String(fiscalMonthFromDate(row[dateIdx])) !== month) return false;
    if (query) {
      const haystack = normalize(row.filter(v => v !== null && v !== "").join(" "));
      if (!haystack.includes(query)) return false;
    }
    return true;
  });
  state.page = 1;
  renderIncidentTable();
}

function renderIncidentTable() {
  const start = (state.page - 1) * state.pageSize;
  const pageItems = state.filteredRows.slice(start, start + state.pageSize);
  const fields = {
    date: "วันที่เกิดอุบัติการณ์",
    title: "รหัส: เรื่องอุบัติการณ์",
    sub: "เรื่องย่อย",
    severity: "ความรุนแรง",
    unit: "หน่วยงานที่บันทึกรายงาน",
    location: "สถานที่เกิดเหตุ",
    status: "สถานะ",
  };

  el("incidentTbody").innerHTML = pageItems.length ? pageItems.map((item, offset) => {
    const globalIndex = start + offset;
    const title = valueAt(item, fields.title);
    const type = riskType(title);
    const severity = valueAt(item, fields.severity);
    return `<tr>
      <td><span class="badge">${item.fy}</span></td>
      <td>${escapeHtml(formatThaiDate(valueAt(item, fields.date)))}</td>
      <td class="cell-title"><div class="ellipsis-2"><strong>${escapeHtml(title || "ไม่ระบุ")}</strong></div><span class="badge ${type.toLowerCase().replace(" ","-")}">${escapeHtml(type)}</span></td>
      <td class="cell-sub"><div class="ellipsis-2">${escapeHtml(valueAt(item, fields.sub) || "–")}</div></td>
      <td><span class="sev ${isHighSeverity(severity) ? "high" : ""}">${escapeHtml(severity ?? "–")}</span></td>
      <td class="cell-unit">${escapeHtml(valueAt(item, fields.unit) || "–")}</td>
      <td>${escapeHtml(valueAt(item, fields.location) || "–")}</td>
      <td class="cell-status"><div class="ellipsis-2">${escapeHtml(valueAt(item, fields.status) || "–")}</div></td>
      <td><button class="detail-btn" data-row-index="${globalIndex}">เปิดดู</button></td>
    </tr>`;
  }).join("") : `<tr><td colspan="9" class="chart-empty">ไม่พบข้อมูลตามเงื่อนไข</td></tr>`;

  el("resultCount").textContent = `${fmt.format(state.filteredRows.length)} รายการ`;
  el("dataSourceLabel").textContent = state.selectedYear === "all" ? "ข้อมูลปีงบประมาณ 2566–2569" : `ข้อมูลปีงบประมาณ ${state.selectedYear}`;
  document.querySelectorAll(".detail-btn").forEach(btn => btn.addEventListener("click", () => openDetail(Number(btn.dataset.rowIndex))));
  renderPagination();
}

function renderPagination() {
  const totalPages = Math.max(1, Math.ceil(state.filteredRows.length / state.pageSize));
  const current = Math.min(state.page, totalPages);
  state.page = current;
  const pages = [];
  const from = Math.max(1, current - 2);
  const to = Math.min(totalPages, current + 2);
  pages.push(`<button class="page-btn" data-page="${current-1}" ${current===1?"disabled":""}>‹</button>`);
  if (from > 1) pages.push(`<button class="page-btn" data-page="1">1</button>${from>2?"<span>…</span>":""}`);
  for (let p=from; p<=to; p++) pages.push(`<button class="page-btn ${p===current?"active":""}" data-page="${p}">${p}</button>`);
  if (to < totalPages) pages.push(`${to<totalPages-1?"<span>…</span>":""}<button class="page-btn" data-page="${totalPages}">${totalPages}</button>`);
  pages.push(`<button class="page-btn" data-page="${current+1}" ${current===totalPages?"disabled":""}>›</button>`);
  el("pagination").innerHTML = pages.join("");
  el("pagination").querySelectorAll("button:not(:disabled)").forEach(btn => btn.addEventListener("click", () => {
    state.page = Number(btn.dataset.page);
    renderIncidentTable();
    document.querySelector(".table-wrap").scrollTop = 0;
  }));
}

function openDetail(index) {
  const item = state.filteredRows[index];
  if (!item) return;
  const title = valueAt(item, "รหัส: เรื่องอุบัติการณ์") || "รายละเอียดอุบัติการณ์";
  el("dialogYearBadge").textContent = `ปีงบประมาณ ${item.fy}`;
  el("dialogTitle").textContent = title;
  const longFields = new Set(["สรุปประเด็นปัญหา","รายละเอียดการเกิด","การจัดการเบื้องต้น","ข้อเสนอแนะเพื่อการแก้ไข","สรุปการแก้ไขของกลุ่ม/หน่วยงานหลัก","การร่วมแก้ไขของกลุ่ม/หน่วยงานร่วม","การแก้ไขของกรรมการความเสี่ยง","มีการแก้ไขปัญหาอย่างไร","การทำ RCA/Contributing Factor อะไร อย่างไร","มีการปรับ/พัฒนาระบบ อะไร อย่างไร"]);
  el("dialogContent").innerHTML = state.meta.headers.map((header, i) => {
    let value = item.row[i];
    if (value === null || value === "") return "";
    if (header.startsWith("วันที่")) value = formatThaiDate(value);
    return `<dl class="detail-item ${longFields.has(header) ? "full" : ""}"><dt>${escapeHtml(header)}</dt><dd>${escapeHtml(value)}</dd></dl>`;
  }).join("");
  el("detailDialog").showModal();
}

function renderProfiles() {
  if (state.selectedYear === "all") {
    el("profileContent").innerHTML = `<div class="notice"><span class="notice-icon">i</span><div><strong>กรุณาเลือกปีงบประมาณ</strong><p>Risk Profile จัดทำแยกเป็นรายปี</p></div></div>`;
    return;
  }
  const rows = state.profiles[String(state.selectedYear)] || [];
  const sections = ["Clinical", "Non-clinical"];
  el("profileContent").innerHTML = sections.map(section => {
    const data = rows.filter(x => x.type === section);
    const title = section === "Clinical" ? "ความเสี่ยงด้านคลินิก (Clinical Risk)" : "ความเสี่ยงด้านทั่วไป (Non-clinical Risk)";
    return `<section class="profile-section panel">
      <h3>${title} — ปีงบประมาณ ${state.selectedYear}</h3>
      <div class="table-wrap"><table class="profile-table"><thead><tr><th>อันดับ</th><th>รายการความเสี่ยง</th><th>โอกาส</th><th>ผลกระทบ</th><th>คะแนน</th><th>ระดับ</th><th>มาตรการควบคุม/ป้องกัน/แก้ไข</th></tr></thead>
      <tbody>${data.map(x => `<tr><td>${x.rank}</td><td>${escapeHtml(x.risk)}</td><td>${escapeHtml(x.likelihood)}</td><td>${escapeHtml(x.impact)}</td><td><strong>${escapeHtml(x.score)}</strong></td><td><span class="level-pill level-${escapeHtml(x.level)}">${escapeHtml(x.level)}</span></td><td>${escapeHtml(x.control)}</td></tr>`).join("")}</tbody></table></div>
    </section>`;
  }).join("");
}

function renderRegisters() {
  if (state.selectedYear === "all") {
    el("registerContent").innerHTML = `<div class="notice"><span class="notice-icon">i</span><div><strong>กรุณาเลือกปีงบประมาณ</strong><p>Risk Register จัดทำแยกเป็นรายปี</p></div></div>`;
    return;
  }
  const query = normalize(el("registerSearch").value);
  const rows = (state.registers[String(state.selectedYear)] || []).filter(x => !query || normalize(Object.values(x).join(" ")).includes(query));
  el("registerContent").innerHTML = rows.length ? `<div class="register-grid">${rows.map(x => `
    <article class="register-card">
      <div class="meta"><span class="badge">${escapeHtml(String(x.riskId || "").replace("\n"," / "))}</span><span class="badge">${escapeHtml(x.quarter || "")}</span></div>
      <h3>${escapeHtml(x.title)}</h3>
      <p>${escapeHtml(x.description)}</p>
      <div class="register-score">
        <div class="score-box"><span>Likelihood</span><strong>${escapeHtml(x.likelihood)}</strong></div>
        <div class="score-box"><span>Impact</span><strong>${escapeHtml(x.impact)}</strong></div>
        <div class="score-box"><span>Risk Level</span><strong>${escapeHtml(String(x.level || "").replace("\n"," "))}</strong></div>
      </div>
      <div class="register-detail"><strong>มาตรการป้องกันและถ่ายโอน</strong><p>${escapeHtml(x.prevention)}</p></div>
      <div class="register-detail"><strong>การติดตามและควบคุม</strong><p>${escapeHtml(x.monitor)}</p></div>
      <div class="register-detail"><strong>แนวทางบรรเทาความเสียหาย</strong><p>${escapeHtml(x.mitigation)}</p></div>
    </article>`).join("")}</div>` : `<div class="chart-empty panel">ไม่พบข้อมูล</div>`;
}

function renderValidation() {
  el("validationList").innerHTML = state.meta.validation.map(item => `
    <div class="validation-item">
      <div><strong>ปีงบประมาณ ${item.year}</strong><br><small>${escapeHtml(item.sourceIncident)}</small></div>
      <div style="text-align:right"><strong>${fmt.format(item.incidentRows)} รายการ</strong><br><small>Profile ${item.profileRows} / Register ${item.registerRows}</small></div>
    </div>`).join("");
}

function csvCell(value) {
  const text = String(value ?? "");
  return `"${text.replaceAll('"','""')}"`;
}

function exportFilteredCsv() {
  const headers = ["ปีงบประมาณ", ...state.meta.headers];
  const lines = [headers.map(csvCell).join(",")];
  for (const item of state.filteredRows) lines.push([item.fy, ...item.row].map(csvCell).join(","));
  const blob = new Blob(["\ufeff" + lines.join("\r\n")], {type: "text/csv;charset=utf-8"});
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `risk_incidents_${state.selectedYear}_${new Date().toISOString().slice(0,10)}.csv`;
  a.click();
  URL.revokeObjectURL(url);
  showToast(`ส่งออก ${fmt.format(state.filteredRows.length)} รายการแล้ว`);
}

function showToast(message) {
  const toast = el("toast");
  toast.textContent = message;
  toast.classList.add("show");
  clearTimeout(showToast.timer);
  showToast.timer = setTimeout(() => toast.classList.remove("show"), 2600);
}

async function refreshYearDependentData() {
  el("loading").classList.remove("hidden");
  try {
    renderDashboard();
    state.currentRows = await getSelectedItems();
    rebuildIncidentFilters();
    applyIncidentFilters();
    renderProfiles();
    renderRegisters();
  } finally {
    el("loading").classList.add("hidden");
  }
}

function bindEvents() {
  document.querySelectorAll(".tab").forEach(tab => tab.addEventListener("click", () => {
    document.querySelectorAll(".tab").forEach(x => x.classList.remove("active"));
    document.querySelectorAll(".view").forEach(x => x.classList.remove("active"));
    tab.classList.add("active");
    state.activeView = tab.dataset.view;
    el(`view-${state.activeView}`).classList.add("active");
  }));

  el("yearSelect").addEventListener("change", async event => {
    state.selectedYear = event.target.value === "all" ? "all" : Number(event.target.value);
    await refreshYearDependentData();
  });

  let searchTimer;
  el("searchInput").addEventListener("input", () => { clearTimeout(searchTimer); searchTimer = setTimeout(applyIncidentFilters, 180); });
  ["typeFilter","severityFilter","unitFilter","monthFilter"].forEach(id => el(id).addEventListener("change", applyIncidentFilters));
  el("clearFiltersBtn").addEventListener("click", () => {
    ["searchInput","typeFilter","severityFilter","unitFilter","monthFilter"].forEach(id => el(id).value = "");
    applyIncidentFilters();
  });
  el("exportCsvBtn").addEventListener("click", exportFilteredCsv);
  el("registerSearch").addEventListener("input", renderRegisters);
  el("closeDialogBtn").addEventListener("click", () => el("detailDialog").close());
  el("detailDialog").addEventListener("click", event => {
    if (event.target === el("detailDialog")) el("detailDialog").close();
  });
}

async function init() {
  try {
    [state.meta, state.profiles, state.registers] = await Promise.all([
      fetchJson("data/meta.json"),
      fetchJson("data/profiles.json"),
      fetchJson("data/registers.json"),
    ]);
    state.selectedYear = state.meta.defaultYear;
    el("yearSelect").innerHTML = `<option value="all">ทุกปีงบประมาณ</option>` + state.meta.years.map(y => `<option value="${y}" ${y===state.selectedYear?"selected":""}>พ.ศ. ${y}</option>`).join("");
    bindEvents();
    renderValidation();
    await refreshYearDependentData();
  } catch (error) {
    console.error(error);
    el("loading").innerHTML = `<div class="panel" style="max-width:650px"><h2>ไม่สามารถเปิด Dashboard ได้</h2><p>${escapeHtml(error.message)}</p><p>หากเปิดไฟล์ index.html โดยตรง กรุณาเปิดผ่าน GitHub Pages หรือเว็บเซิร์ฟเวอร์ เนื่องจากเบราว์เซอร์ไม่อนุญาตให้โหลดไฟล์ JSON ผ่าน file://</p></div>`;
  }
}

document.addEventListener("DOMContentLoaded", init);
