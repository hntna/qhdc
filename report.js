/**
 * report.js - Logic hiển thị bảng tổng hợp Mảng, Dịch vụ và Chiến lược 5 năm
 */

// Trạng thái ứng dụng báo cáo
const reportState = {
  rawResponse: null,
  activeTab: 'tabTongHop',
  activeSubtab: 'subtabTongHop',
  unitMultiplier: 1.0, // 1.0 = Triệu USD (M$), 1000 = Nghìn USD (K$), 1000000 = USD
  unitLabel: 'M$',
  searchKeyword: '',
  filterMang: 'ALL'
};

document.addEventListener('DOMContentLoaded', () => {
  if (window.lucide) lucide.createIcons();
  initReportTabs();
  initReportEvents();
  loadReportDataFromServer();
});

// Chuyển đổi Tab & Subtab
function initReportTabs() {
  const tabBtns = document.querySelectorAll('.report-tab-btn, .report-subtab-btn');
  tabBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      const targetTabId = btn.getAttribute('data-tab') || btn.getAttribute('data-subtab');
      if (btn.classList.contains('report-subtab-btn')) {
        switchReportSubtab(targetTabId);
      } else {
        switchReportTab(targetTabId);
      }
    });
  });
}

function switchReportTab(tabId) {
  reportState.activeTab = tabId;
  document.querySelectorAll('.report-tab-btn').forEach(b => {
    b.classList.toggle('active', b.getAttribute('data-tab') === tabId);
  });
  document.querySelectorAll('.report-tab-content').forEach(c => {
    c.classList.toggle('active', c.id === tabId);
  });
  if (window.lucide) lucide.createIcons();
}

function switchReportSubtab(subtabId) {
  reportState.activeSubtab = subtabId;
  document.querySelectorAll('.report-subtab-btn').forEach(b => {
    b.classList.toggle('active', b.getAttribute('data-subtab') === subtabId);
  });
  document.querySelectorAll('.report-subtab-content').forEach(c => {
    c.classList.toggle('active', c.id === subtabId);
  });
  if (window.lucide) lucide.createIcons();
}

function initReportEvents() {
  // Tìm kiếm
  const searchInput = document.getElementById('searchReportInput');
  if (searchInput) {
    searchInput.addEventListener('input', (e) => {
      reportState.searchKeyword = (e.target.value || '').trim().toLowerCase();
      renderActiveTabContent();
    });
  }

  // Đổi đơn vị
  const unitSelect = document.getElementById('selectReportUnit');
  if (unitSelect) {
    unitSelect.addEventListener('change', (e) => {
      const val = e.target.value;
      if (val === 'K') {
        reportState.unitMultiplier = 1000.0;
        reportState.unitLabel = 'K$';
      } else if (val === 'USD') {
        reportState.unitMultiplier = 1000000.0;
        reportState.unitLabel = '$';
      } else {
        reportState.unitMultiplier = 1.0;
        reportState.unitLabel = 'M$';
      }
      document.querySelectorAll('.unit-label-text').forEach(el => el.textContent = reportState.unitLabel);
      renderAllTabs();
    });
  }

  // Nạp lại từ server
  const btnReload = document.getElementById('btnReloadReport');
  if (btnReload) {
    btnReload.addEventListener('click', () => {
      loadReportDataFromServer();
    });
  }

  // Chọn file Excel cục bộ
  const fileInput = document.getElementById('reportFileInput');
  if (fileInput) {
    fileInput.addEventListener('change', (e) => {
      const file = e.target.files[0];
      if (file) {
        parseLocalExcelFile(file);
      }
    });
  }
}

// Nạp dữ liệu từ endpoint backend
async function loadReportDataFromServer(customPath = null) {
  showLoading(true);
  try {
    let url = '/api/report-data';
    if (customPath) {
      url += `?file=${encodeURIComponent(customPath)}`;
    }
    const res = await fetch(url, { cache: 'no-store' });
    if (!res.ok) {
      throw new Error(`HTTP Error ${res.status}`);
    }
    const data = await res.json();
    if (!data.success) {
      throw new Error(data.error || 'Không tải được dữ liệu báo cáo');
    }

    reportState.rawResponse = data;
    updateFileBadge(data.file_path || data.file_name);
    renderAllTabs();
    showToast('Đã nạp thành công dữ liệu báo cáo!', 'success');
  } catch (err) {
    console.warn('Lỗi nạp từ server:', err);
    // Thử fallback sang đọc file nhúng sẵn hoặc Masterlist_Mau_Tach nếu chạy offline
    tryFallbackLocalLoad(err.message);
  } finally {
    showLoading(false);
  }
}

// Fallback nếu không có server Python
async function tryFallbackLocalLoad(errMsg) {
  try {
    const res = await fetch('Masterlist%202027-2028_Mau_Tach.xlsx');
    if (res.ok) {
      const buf = await res.arrayBuffer();
      parseExcelArrayBuffer(buf, 'Masterlist 2027-2028_Mau_Tach.xlsx');
      return;
    }
  } catch (e) {
    //
  }
  showToast(`Không thể nạp dữ liệu từ server: ${errMsg}. Bạn có thể chọn file Excel bằng nút "Nạp file Excel".`, 'warning');
}

function parseLocalExcelFile(file) {
  showLoading(true);
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      parseExcelArrayBuffer(e.target.result, file.name);
    } catch (err) {
      showToast('Lỗi khi đọc file Excel: ' + err.message, 'error');
      showLoading(false);
    }
  };
  reader.readAsArrayBuffer(file);
}

function parseExcelArrayBuffer(buffer, fileName) {
  if (typeof XLSX === 'undefined') {
    throw new Error('Chưa nạp thư viện XLSX (SheetJS)');
  }
  const wb = XLSX.read(buffer, { type: 'array', cellFormula: true, cellStyles: true });
  const sheetNames = wb.SheetNames;

  let nameM = sheetNames.find(s => s.toLowerCase().includes('mảng') || s.toLowerCase().includes('mang'));
  let nameDV = sheetNames.find(s => s.toLowerCase().includes('dịch vụ') || s.toLowerCase().includes('dich vu'));

  if (!nameM && !nameDV) {
    throw new Error('File không chứa sheet "TH theo Mảng" hoặc "TH theo Dịch vụ"');
  }

  // Client-side parser qua SheetJS
  const tablesMang = nameM ? parseSheetWithSheetJS(wb.Sheets[nameM], 'mang') : [];
  const tablesDV = nameDV ? parseSheetWithSheetJS(wb.Sheets[nameDV], 'dv') : [];
  const summary = computeClientSummary(tablesMang);
  const strategyComp = buildClientStrategyComp(tablesMang);

  reportState.rawResponse = {
    success: true,
    file_name: fileName,
    file_path: fileName,
    tables_mang: tablesMang,
    tables_dv: tablesDV,
    summary: summary,
    strategy_comparison: strategyComp
  };

  updateFileBadge(fileName);
  renderAllTabs();
  showLoading(false);
  showToast(`Đã đọc thành công file: ${fileName}`, 'success');
}

function parseSheetWithSheetJS(sheet, type) {
  const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1:Z150');
  const tables = [];
  let currentTable = null;

  for (let r = range.s.r; r <= range.e.r; r++) {
    const getVal = (c) => {
      const cell = sheet[XLSX.utils.encode_cell({ r, c })];
      return cell ? cell.v : null;
    };

    const b = String(getVal(1) || '').trim();
    const c = String(getVal(2) || '').trim();
    const d = getVal(3);
    const e = getVal(4);
    const f = getVal(5);
    const g = getVal(6);

    const isSec = ['I.', 'II.', 'III.', 'IV.', 'V.', 'VI.', 'VII.', 'VIII.', 'IX.', 'X.', 'Đầu tư', 'Hiện đại'].some(p => b.startsWith(p));
    if (isSec) {
      currentTable = { title: b, rows: [] };
      tables.push(currentTable);
      continue;
    }

    if (b === 'STT' || c === 'STT' || c === 'Danh mục' || c === 'Dịch vụ') continue;

    if (currentTable) {
      if (type === 'mang') {
        const name = c || b;
        if (name || d !== null || e !== null || f !== null) {
          const isTot = name.toLowerCase() === 'tổng' || b.toLowerCase() === 'tổng';
          const isSub = b === '-' || name.startsWith('-');
          const tot = typeof d === 'number' ? d : 0;
          const y27 = typeof e === 'number' ? e : 0;
          const y28 = typeof f === 'number' ? f : 0;
          const stratTot = typeof getVal(11) === 'number' ? getVal(11) : (typeof getVal(9) === 'number' ? getVal(9) : 0);
          const strat27 = typeof getVal(13) === 'number' ? getVal(13) : 0;
          const strat28 = typeof getVal(14) === 'number' ? getVal(14) : 0;

          if (name) {
            currentTable.rows.push({
              stt: b,
              name: isTot ? 'Tổng' : name.replace(/^-+\s*/, ''),
              is_total: isTot,
              is_sub: isSub,
              tong: tot || (y27 + y28),
              y2027: y27,
              y2028: y28,
              strat_tot: stratTot,
              strat_27: strat27,
              strat_28: strat28
            });
          }
        }
      } else {
        const name = String(d || c || b).trim();
        if (name || e !== null || f !== null || g !== null) {
          const isTot = name.toLowerCase() === 'tổng' || b.toLowerCase() === 'tổng';
          const isSub = b === '-' || name.startsWith('-');
          const tot = typeof e === 'number' ? e : 0;
          const y27 = typeof f === 'number' ? f : 0;
          const y28 = typeof g === 'number' ? g : 0;
          const stratTot = typeof getVal(10) === 'number' ? getVal(10) : 0;

          if (name) {
            currentTable.rows.push({
              stt: b,
              name: isTot ? 'Tổng' : name.replace(/^-+\s*/, ''),
              is_total: isTot,
              is_sub: isSub,
              tong: tot || (y27 + y28),
              y2027: y27,
              y2028: y28,
              strat_tot: stratTot
            });
          }
        }
      }
    }
  }
  return tables.filter(t => t.rows && t.rows.length > 0);
}

function computeClientSummary(tablesMang) {
  let tot = 0, y27 = 0, y28 = 0, strat = 0;
  const list = [];
  tablesMang.forEach(t => {
    const totRow = t.rows.find(r => r.is_total);
    const rowTot = totRow ? totRow.tong : t.rows.reduce((s, r) => s + (r.is_sub ? 0 : r.tong), 0);
    const row27 = totRow ? totRow.y2027 : t.rows.reduce((s, r) => s + (r.is_sub ? 0 : r.y2027), 0);
    const row28 = totRow ? totRow.y2028 : t.rows.reduce((s, r) => s + (r.is_sub ? 0 : r.y2028), 0);
    const rowStrat = totRow ? totRow.strat_tot : 0;

    tot += rowTot;
    y27 += row27;
    y28 += row28;
    strat += rowStrat;

    list.push({
      name: t.title,
      tong: rowTot,
      y2027: row27,
      y2028: row28,
      strat_tot: rowStrat
    });
  });

  list.forEach(item => {
    item.share = tot > 0 ? (item.tong / tot * 100) : 0;
  });

  return {
    total_investment: tot,
    total_2027: y27,
    total_2028: y28,
    total_strategy_5y: strat,
    share_2027: tot > 0 ? (y27 / tot * 100) : 0,
    share_2028: tot > 0 ? (y28 / tot * 100) : 0,
    by_mang: list
  };
}

function buildClientStrategyComp(tablesMang) {
  const comp = [];
  tablesMang.forEach(t => {
    const totRow = t.rows.find(r => r.is_total);
    if (totRow) {
      const qTot = totRow.tong;
      const sTot = totRow.strat_tot || 0;
      const s27_28 = (totRow.strat_27 || 0) + (totRow.strat_28 || 0);
      const delta = qTot - s27_28;
      comp.push({
        title: t.title,
        qhdc_2027: totRow.y2027,
        qhdc_2028: totRow.y2028,
        qhdc_total: qTot,
        strat_2027: totRow.strat_27 || 0,
        strat_2028: totRow.strat_28 || 0,
        strat_27_28: s27_28,
        strat_5y: sTot,
        delta: delta,
        rate: s27_28 > 0 ? (qTot / s27_28 * 100) : 0,
        status: delta > 0.05 ? `Vượt CL +${delta.toFixed(2)} M$` : (delta < -0.05 ? `Dưới CL ${delta.toFixed(2)} M$` : 'Vừa khớp'),
        status_color: delta > 0.05 ? 'orange' : (delta < -0.05 ? 'blue' : 'green')
      });
    }
  });
  return comp;
}

// Cập nhật nhãn file đang xem
function updateFileBadge(filePath) {
  const badge = document.getElementById('reportFileBadge');
  if (badge && filePath) {
    badge.textContent = filePath;
    badge.title = filePath;
  }
}

// Render toàn bộ các Tabs
function renderAllTabs() {
  if (!reportState.rawResponse) return;
  renderTabTongHop();
  renderTabMang();
  renderTabDichVu();
  renderTabStrategy();
}

function renderActiveTabContent() {
  if (reportState.activeTab === 'tabTongHop') renderTabTongHop();
  else if (reportState.activeTab === 'tabMang') renderTabMang();
  else if (reportState.activeTab === 'tabDichVu') renderTabDichVu();
  else if (reportState.activeTab === 'tabStrategy') renderTabStrategy();
}

// Format số
function fmtVal(num) {
  if (num === null || num === undefined || isNaN(num) || num === 0) return '0.00';
  const scaled = num * reportState.unitMultiplier;
  return scaled.toLocaleString('vi-VN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// TAB 1: TỔNG HỢP
function renderTabTongHop() {
  const summary = reportState.rawResponse?.summary;
  if (!summary) return;

  // KPI cards
  const elTot = document.getElementById('kpiTotalInv');
  if (elTot) elTot.textContent = fmtVal(summary.total_investment);

  const el27 = document.getElementById('kpi2027');
  if (el27) el27.textContent = fmtVal(summary.total_2027);

  const el28 = document.getElementById('kpi2028');
  if (el28) el28.textContent = fmtVal(summary.total_2028);

  const elStrat = document.getElementById('kpiStrat5Y');
  if (elStrat) elStrat.textContent = fmtVal(summary.total_strategy_5y);

  // Bảng tổng hợp các Mảng
  const tbody = document.getElementById('tableSummaryMangBody');
  if (!tbody) return;

  const rows = summary.by_mang || [];
  let html = '';

  rows.forEach((m, idx) => {
    html += `
      <tr>
        <td class="col-mang" style="font-weight: 700; color: #1e293b;">
          <span style="display: inline-block; width: 22px; color: #64748b;">${idx + 1}.</span>
          ${escapeHtml(m.name)}
        </td>
        <td class="col-num col-tot">${fmtVal(m.tong)}</td>
        <td class="col-num col-27">${fmtVal(m.y2027)}</td>
        <td class="col-num col-28">${fmtVal(m.y2028)}</td>
        <td class="col-num" style="color: #64748b; font-weight: 700; width: 100px;">
          ${m.share ? m.share.toFixed(1) + '%' : '-'}
        </td>
      </tr>
    `;
  });

  // Hàng tổng cộng
  html += `
    <tr class="row-total-table">
      <td class="col-mang" style="font-weight: 800; font-size: 0.95rem; color: #065f46;">
        TỔNG CỘNG TOÀN MẠNG
      </td>
      <td class="col-num col-tot" style="font-size: 1rem;">${fmtVal(summary.total_investment)}</td>
      <td class="col-num col-27">${fmtVal(summary.total_2027)}</td>
      <td class="col-num col-28">${fmtVal(summary.total_2028)}</td>
      <td class="col-num" style="color: #065f46; font-weight: 800;">100%</td>
    </tr>
  `;

  tbody.innerHTML = html;
}

// TAB 2: TỔNG HỢP THEO MẢNG (Hiển thị các bảng với đúng 4 cột: Mảng, Tổng, 2027, 2028)
function renderTabMang() {
  const container = document.getElementById('containerTablesMang');
  if (!container) return;

  const tables = reportState.rawResponse?.tables_mang || [];
  if (tables.length === 0) {
    container.innerHTML = `<div style="text-align: center; padding: 3rem; color: #94a3b8;">Chưa có dữ liệu sheet "TH theo Mảng"</div>`;
    return;
  }

  const kw = reportState.searchKeyword;
  let html = '';

  tables.forEach((t, tIdx) => {
    // Lọc theo từ khóa tìm kiếm nếu có
    const filteredRows = t.rows.filter(r => {
      if (!kw) return true;
      return r.name.toLowerCase().includes(kw) || t.title.toLowerCase().includes(kw);
    });

    if (filteredRows.length === 0) return;

    html += `
      <div class="table-card">
        <div class="table-card-header">
          <h3 class="table-card-title">
            <i data-lucide="layers" class="w-4 h-4 text-indigo-600"></i>
            <span>${escapeHtml(t.title)}</span>
          </h3>
          <div style="display: flex; gap: 8px; align-items: center;">
            <button type="button" class="btn btn-outline btn-xs" onclick="copyTableHtml('table_mang_${tIdx}', '${escapeHtml(t.title)}')">
              <i data-lucide="copy" class="w-3.5 h-3.5 inline mr-1"></i> Sao chép
            </button>
            <span class="table-badge-count">${filteredRows.length} dòng</span>
          </div>
        </div>
        <div style="overflow-x: auto;">
          <table class="report-table" id="table_mang_${tIdx}">
            <thead>
              <tr>
                <th class="col-mang">Mảng</th>
                <th class="col-num col-tot">Tổng (${reportState.unitLabel})</th>
                <th class="col-num col-27">2027 (${reportState.unitLabel})</th>
                <th class="col-num col-28">2028 (${reportState.unitLabel})</th>
              </tr>
            </thead>
            <tbody>
    `;

    filteredRows.forEach(r => {
      const isTot = r.is_total;
      const isSub = r.is_sub;
      const rowCls = isTot ? 'row-total-table' : (isSub ? 'row-sub-item' : '');

      html += `
        <tr class="${rowCls}">
          <td class="col-mang" style="${isTot ? 'font-weight: 800; color: #065f46;' : (isSub ? 'padding-left: 2rem; color: #475569;' : 'font-weight: 700;')}">
            ${isSub ? '<span style="color: #94a3b8; margin-right: 6px;">-</span>' : ''}
            ${escapeHtml(r.name)}
          </td>
          <td class="col-num col-tot" style="${isTot ? 'font-size: 0.95rem;' : ''}">${fmtVal(r.tong)}</td>
          <td class="col-num col-27">${fmtVal(r.y2027)}</td>
          <td class="col-num col-28">${fmtVal(r.y2028)}</td>
        </tr>
      `;
    });

    html += `
            </tbody>
          </table>
        </div>
      </div>
    `;
  });

  container.innerHTML = html;
  if (window.lucide) lucide.createIcons();
}

// TAB 3: TỔNG HỢP THEO DỊCH VỤ (Hiển thị các bảng với đúng 4 cột: Mảng, Tổng, 2027, 2028)
function renderTabDichVu() {
  const container = document.getElementById('containerTablesDichVu');
  if (!container) return;

  const tables = reportState.rawResponse?.tables_dv || [];
  if (tables.length === 0) {
    container.innerHTML = `<div style="text-align: center; padding: 3rem; color: #94a3b8;">Chưa có dữ liệu sheet "TH theo Dịch vụ"</div>`;
    return;
  }

  const kw = reportState.searchKeyword;
  let html = '';

  tables.forEach((t, tIdx) => {
    const filteredRows = t.rows.filter(r => {
      if (!kw) return true;
      return r.name.toLowerCase().includes(kw) || (r.service && r.service.toLowerCase().includes(kw)) || t.title.toLowerCase().includes(kw);
    });

    if (filteredRows.length === 0) return;

    html += `
      <div class="table-card">
        <div class="table-card-header">
          <h3 class="table-card-title">
            <i data-lucide="network" class="w-4 h-4 text-emerald-600"></i>
            <span>${escapeHtml(t.title)}</span>
          </h3>
          <div style="display: flex; gap: 8px; align-items: center;">
            <button type="button" class="btn btn-outline btn-xs" onclick="copyTableHtml('table_dv_${tIdx}', '${escapeHtml(t.title)}')">
              <i data-lucide="copy" class="w-3.5 h-3.5 inline mr-1"></i> Sao chép
            </button>
            <span class="table-badge-count">${filteredRows.length} dòng</span>
          </div>
        </div>
        <div style="overflow-x: auto;">
          <table class="report-table" id="table_dv_${tIdx}">
            <thead>
              <tr>
                <th class="col-mang">Mảng</th>
                <th class="col-num col-tot">Tổng (${reportState.unitLabel})</th>
                <th class="col-num col-27">2027 (${reportState.unitLabel})</th>
                <th class="col-num col-28">2028 (${reportState.unitLabel})</th>
              </tr>
            </thead>
            <tbody>
    `;

    filteredRows.forEach(r => {
      const isTot = r.is_total;
      const isSub = r.is_sub;
      const rowCls = isTot ? 'row-total-table' : (isSub ? 'row-sub-item' : '');

      html += `
        <tr class="${rowCls}">
          <td class="col-mang" style="${isTot ? 'font-weight: 800; color: #065f46;' : (isSub ? 'padding-left: 2rem; color: #475569;' : 'font-weight: 700;')}">
            ${isSub ? '<span style="color: #94a3b8; margin-right: 6px;">-</span>' : ''}
            ${escapeHtml(r.name)}
          </td>
          <td class="col-num col-tot" style="${isTot ? 'font-size: 0.95rem;' : ''}">${fmtVal(r.tong)}</td>
          <td class="col-num col-27">${fmtVal(r.y2027)}</td>
          <td class="col-num col-28">${fmtVal(r.y2028)}</td>
        </tr>
      `;
    });

    html += `
            </tbody>
          </table>
        </div>
      </div>
    `;
  });

  container.innerHTML = html;
  if (window.lucide) lucide.createIcons();
}

// TAB 4: SO SÁNH VỚI CHIẾN LƯỢC 5 NĂM
function renderTabStrategy() {
  const tbody = document.getElementById('tableStrategyBody');
  if (!tbody) return;

  const compList = reportState.rawResponse?.strategy_comparison || [];
  if (compList.length === 0) {
    tbody.innerHTML = `<tr><td colspan="10" style="text-align: center; padding: 2.5rem; color: #94a3b8;">Chưa có dữ liệu đối chiếu Chiến lược 5 năm</td></tr>`;
    return;
  }

  let html = '';
  let sumQ27 = 0, sumQ28 = 0, sumQTot = 0;
  let sumS27 = 0, sumS28 = 0, sumSTot = 0, sumS5Y = 0;

  compList.forEach((c, idx) => {
    sumQ27 += c.qhdc_2027;
    sumQ28 += c.qhdc_2028;
    sumQTot += c.qhdc_total;
    sumS27 += c.strat_2027;
    sumS28 += c.strat_2028;
    sumSTot += c.strat_27_28;
    sumS5Y += c.strat_5y;

    const badgeCls = `badge-${c.status_color || 'green'}`;

    html += `
      <tr>
        <td style="font-weight: 700; color: #1e293b;">
          <span style="display: inline-block; width: 22px; color: #64748b;">${idx + 1}.</span>
          ${escapeHtml(c.title)}
        </td>
        <td class="col-num col-27">${fmtVal(c.qhdc_2027)}</td>
        <td class="col-num col-28">${fmtVal(c.qhdc_2028)}</td>
        <td class="col-num col-tot">${fmtVal(c.qhdc_total)}</td>
        <td class="col-num" style="color: #6366f1;">${fmtVal(c.strat_2027)}</td>
        <td class="col-num" style="color: #059669;">${fmtVal(c.strat_2028)}</td>
        <td class="col-num" style="font-weight: 700; color: #d97706;">${fmtVal(c.strat_27_28)}</td>
        <td class="col-num" style="font-weight: 800; color: #b45309;">${fmtVal(c.strat_5y)}</td>
        <td class="col-num" style="font-weight: 700; color: ${c.delta > 0.05 ? '#e11d48' : '#059669'};">
          ${c.delta > 0 ? '+' : ''}${fmtVal(c.delta)}
        </td>
        <td style="text-align: center;">
          <span class="badge-status ${badgeCls}">${escapeHtml(c.status)}</span>
        </td>
      </tr>
    `;
  });

  // Hàng tổng
  const sumDelta = sumQTot - sumSTot;
  html += `
    <tr class="row-total-table" style="background: #fef3c7 !important;">
      <td style="font-weight: 800; font-size: 0.95rem; color: #92400e;">TỔNG CỘNG TOÀN BỘ</td>
      <td class="col-num col-27" style="font-weight: 800;">${fmtVal(sumQ27)}</td>
      <td class="col-num col-28" style="font-weight: 800;">${fmtVal(sumQ28)}</td>
      <td class="col-num col-tot" style="font-size: 1rem;">${fmtVal(sumQTot)}</td>
      <td class="col-num" style="font-weight: 700; color: #6366f1;">${fmtVal(sumS27)}</td>
      <td class="col-num" style="font-weight: 700; color: #059669;">${fmtVal(sumS28)}</td>
      <td class="col-num" style="font-weight: 800; color: #d97706;">${fmtVal(sumSTot)}</td>
      <td class="col-num" style="font-weight: 800; color: #b45309;">${fmtVal(sumS5Y)}</td>
      <td class="col-num" style="font-weight: 800; color: ${sumDelta > 0.05 ? '#e11d48' : '#059669'};">
        ${sumDelta > 0 ? '+' : ''}${fmtVal(sumDelta)}
      </td>
      <td style="text-align: center;">
        <span class="badge-status ${sumDelta > 0.05 ? 'badge-orange' : 'badge-green'}">
          ${sumDelta > 0.05 ? 'Vượt kế hoạch' : 'Trong ngân sách'}
        </span>
      </td>
    </tr>
  `;

  tbody.innerHTML = html;
}

// Sao chép bảng sang Clipboard (tab-separated)
function copyTableHtml(tableId, title) {
  const table = document.getElementById(tableId);
  if (!table) return;

  let text = '';
  const rows = table.querySelectorAll('tr');
  rows.forEach(tr => {
    const cells = tr.querySelectorAll('th, td');
    const rowText = Array.from(cells).map(c => c.textContent.trim().replace(/\s+/g, ' ')).join('\t');
    text += rowText + '\n';
  });

  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(text).then(() => {
      showToast(`Đã sao chép bảng "${title || tableId}" vào Clipboard! Bạn có thể dán trực tiếp vào Excel.`, 'success');
    }).catch(err => {
      showToast('Không thể sao chép: ' + err.message, 'error');
    });
  }
}

// Sao chép toàn bộ Tab/Subtab đang hiển thị
function copyCurrentTabTables() {
  let containerId = '';
  const current = reportState.activeSubtab || reportState.activeTab;
  if (current === 'subtabMang' || current === 'tabMang') containerId = 'containerTablesMang';
  else if (current === 'subtabDichVu' || current === 'tabDichVu') containerId = 'containerTablesDichVu';
  else if (current === 'subtabStratDetail' || current === 'tabStrategy') containerId = 'subtabStratDetail';
  else containerId = 'subtabTongHop';

  const container = document.getElementById(containerId);
  if (!container) return;

  const tables = container.querySelectorAll('table');
  if (tables.length === 0) return;

  let allText = '';
  tables.forEach((table, idx) => {
    const parentCard = table.closest('.table-card');
    const headerTitle = parentCard?.querySelector('.table-card-title')?.textContent?.trim() || `Bảng ${idx + 1}`;
    allText += `=== ${headerTitle} ===\n`;

    const rows = table.querySelectorAll('tr');
    rows.forEach(tr => {
      const cells = tr.querySelectorAll('th, td');
      const rowText = Array.from(cells).map(c => c.textContent.trim().replace(/\s+/g, ' ')).join('\t');
      allText += rowText + '\n';
    });
    allText += '\n';
  });

  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(allText).then(() => {
      showToast('Đã sao chép toàn bộ bảng của Tab hiện tại vào Clipboard! Bạn có thể dán trực tiếp vào Excel.', 'success');
    });
  }
}

// Xuất Excel nhanh bằng SheetJS
function exportReportExcel() {
  if (!reportState.rawResponse) {
    showToast('Chưa có dữ liệu để xuất Excel!', 'warning');
    return;
  }
  if (typeof XLSX === 'undefined') {
    showToast('Chưa có thư viện xuất Excel!', 'error');
    return;
  }

  const wb = XLSX.utils.book_new();

  // Sheet 1: TH theo Mảng
  const mangRows = [ ['Mảng / Hạng mục', 'Tổng (M$)', '2027 (M$)', '2028 (M$)'] ];
  (reportState.rawResponse.tables_mang || []).forEach(t => {
    mangRows.push([`=== ${t.title} ===`, '', '', '']);
    t.rows.forEach(r => {
      mangRows.push([ (r.is_sub ? '  - ' : '') + r.name, r.tong, r.y2027, r.y2028 ]);
    });
    mangRows.push(['', '', '', '']);
  });
  const wsM = XLSX.utils.aoa_to_sheet(mangRows);
  XLSX.utils.book_append_sheet(wb, wsM, 'TH theo Mảng (4 Cột)');

  // Sheet 2: TH theo Dịch vụ
  const dvRows = [ ['Mảng / Dịch vụ', 'Tổng (M$)', '2027 (M$)', '2028 (M$)'] ];
  (reportState.rawResponse.tables_dv || []).forEach(t => {
    dvRows.push([`=== ${t.title} ===`, '', '', '']);
    t.rows.forEach(r => {
      dvRows.push([ (r.is_sub ? '  - ' : '') + r.name, r.tong, r.y2027, r.y2028 ]);
    });
    dvRows.push(['', '', '', '']);
  });
  const wsDV = XLSX.utils.aoa_to_sheet(dvRows);
  XLSX.utils.book_append_sheet(wb, wsDV, 'TH theo Dịch vụ (4 Cột)');

  // Sheet 3: So sánh Chiến lược 5 năm
  const stratRows = [ ['Mảng', 'QHĐC 2027', 'QHĐC 2028', 'Tổng QHĐC', 'CL 2027', 'CL 2028', 'Tổng CL 27-28', 'Tổng CL 5 Năm', 'Chênh lệch', 'Trạng thái'] ];
  (reportState.rawResponse.strategy_comparison || []).forEach(c => {
    stratRows.push([ c.title, c.qhdc_2027, c.qhdc_2028, c.qhdc_total, c.strat_2027, c.strat_2028, c.strat_27_28, c.strat_5y, c.delta, c.status ]);
  });
  const wsStrat = XLSX.utils.aoa_to_sheet(stratRows);
  XLSX.utils.book_append_sheet(wb, wsStrat, 'So sánh Chiến lược 5 năm');

  const now = new Date();
  const dateStr = now.toISOString().slice(0, 10).replace(/-/g, '');
  XLSX.writeFile(wb, `BaoCao_TongHop_Mang_DichVu_QHDC_${dateStr}.xlsx`);
  showToast('Đã xuất file Excel báo cáo thành công!', 'success');
}

// Helpers
function escapeHtml(text) {
  if (!text) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function showLoading(show) {
  const el = document.getElementById('reportLoadingOverlay');
  if (el) el.style.display = show ? 'flex' : 'none';
}

function showToast(message, type = 'info') {
  const container = document.getElementById('toastContainer');
  if (!container) return;

  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;
  toast.innerHTML = `
    <div style="display: flex; align-items: center; gap: 8px;">
      <i data-lucide="${type === 'success' ? 'check-circle-2' : (type === 'error' ? 'alert-octagon' : 'info')}" class="w-4 h-4"></i>
      <span>${escapeHtml(message)}</span>
    </div>
  `;
  container.appendChild(toast);
  if (window.lucide) lucide.createIcons();

  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateY(-10px)';
    setTimeout(() => toast.remove(), 300);
  }, 3500);
}

// Window exports
window.copyTableHtml = copyTableHtml;
window.copyCurrentTabTables = copyCurrentTabTables;
window.exportReportExcel = exportReportExcel;
window.switchReportTab = switchReportTab;
window.switchReportSubtab = switchReportSubtab;
window.loadReportDataFromServer = loadReportDataFromServer;
window.renderAllReportTabs = renderAllTabs;
