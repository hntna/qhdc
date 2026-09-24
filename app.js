/**
 * Hệ thống Tổng hợp & So khớp Masterlist QHĐC 2027-2028
 * Client-side 100% (Offline, No Server)
 */

// Định nghĩa 7 mảng nghiệp vụ và tiêu đề chuẩn
const MANG_CONFIG = [
  { code: 'VT', name: 'Vô tuyến', tt: 'A', keywords: ['VÔ TUYẾN', 'VO TUYEN'] },
  { code: 'ML', name: 'Mạng lõi', tt: 'B', keywords: ['MẠNG LÕI', 'MANG LOI'] },
  { code: 'CDBR', name: 'CĐBR & Truyền hình', tt: 'C', keywords: ['CĐBR', 'CDBR', 'CỐ ĐỊNH BĂNG RỘNG', 'CO DINH BANG RONG'] },
  { code: 'CNTT', name: 'Công nghệ thông tin', tt: 'D', keywords: ['CNTT+VÍ', 'CNTT + VÍ', 'CNTT', 'CÔNG NGHỆ THÔNG TIN'] },
  { code: 'TD', name: 'Truyền dẫn', tt: 'E', keywords: ['TRUYỀN DẪN', 'TRUYEN DAN'] },
  { code: 'CD', name: 'Cơ điện', tt: 'F', keywords: ['CƠ ĐIỆN', 'CO DIEN'] },
  { code: 'HT', name: 'Triển khai hạ tầng', tt: 'G', keywords: ['TRIỂN KHAI HẠ TẦNG', 'TRIEN KHAI HA TANG', 'HẠ TẦNG', 'HA TANG'] }
];

// Tất cả từ khóa phân cách mảng cấp 1
const ALL_SECTION_HEADERS = [
  'VÔ TUYẾN', 'VO TUYEN',
  'MẠNG LÕI', 'MANG LOI',
  'CĐBR', 'CDBR',
  'CNTT+VÍ', 'CNTT + VÍ',
  'TRUYỀN DẪN', 'TRUYEN DAN',
  'CƠ ĐIỆN', 'CO DIEN',
  'TRIỂN KHAI HẠ TẦNG', 'TRIEN KHAI HA TANG'
];

// State quản lý toàn bộ ứng dụng
const state = {
  templateBuffer: null,
  templateWorkbook: null,
  templateFileName: 'Masterlist 2027-2028_Mau.xlsx',
  files: {
    VT: null,
    ML: null,
    CDBR: null,
    CNTT: null,
    TD: null,
    CD: null,
    HT: null
  },
  // Lưu trữ dữ liệu trích xuất theo từng mảng (để thay thế trực tiếp mảng tương ứng khi upload file mới)
  extractedByMang: {
    VT: [],
    ML: [],
    CDBR: [],
    CNTT: [],
    TD: [],
    CD: [],
    HT: []
  },
  mangHeaderInfo: {},   // Lưu thông tin tiêu đề và phạm vi subtotal của từng mảng
  validMaDVSet: new Set(),
  extractedData: [],    // Dữ liệu đã ghép theo thứ tự chuẩn 7 mảng
  validationIssues: [], // Danh sách các lỗi/cảnh báo
  groupsConfig: [],     // Cấu hình nhóm hạng mục
  activeGroupLevels: new Set([1, 2]), // Mặc định chỉ chọn Cấp 1 và Cấp 2
  collapsedGroupKeys: new Set(),            // Lưu các nhóm đang bị thu gọn
  previewColumns: {
    dvt: false,
    kl27: false,
    kl28: false,
    dg: false
  },
  exportBlob: null,
  exportFileName: 'Masterlist 2027-2028_Mau.xlsx',
  strategyProfiles: [],                     // Danh sách profiles Chiến lược 5 năm
  activeStrategyProfileId: 'profile_default', // ID profile đang được kích hoạt
  currentDraftId: null,                     // ID bản lưu đang mở
  currentDraftName: null,                   // Tên bản lưu đang mở
  serverDrafts: []                          // Danh sách bản lưu trên server
};

// ==================== KHỞI TẠO & SỰ KIỆN GIAO DIỆN ====================

document.addEventListener('DOMContentLoaded', () => {
  if (window.lucide) lucide.createIcons();
  initTabs();
  initUploadHandlers();
  initFilterHandlers();
  initGroupLevelFilterButtons();
  initActionButtons();
  initTemplateFile();
  makeTableResizable('validationTable');
  initStrategyComparison();
  initVTBConverter();
  initServerDrafts();
  initEditModalListeners();
  initPreviewInlineEditing();
  renderPreviewTable();
});

// Nạp file phôi mẫu: Luôn tự động lấy file Masterlist 2027-2028_Mau.xlsx, fallback sang bản nhúng nếu offline
async function initTemplateFile() {
  try {
    const res = await fetch('Masterlist%202027-2028_Mau.xlsx', { cache: 'no-store' });
    if (res.ok) {
      const buffer = await res.arrayBuffer();
      state.templateBuffer = buffer;
      state.templateWorkbook = XLSX.read(buffer, { type: 'array', cellFormula: true, cellStyles: true });
      extractValidMaDVFromTemplate(state.templateWorkbook);
      console.log('Đã tự động nạp thành công file mẫu Masterlist 2027-2028_Mau.xlsx');
      return;
    }
  } catch (err) {
    console.warn('Không tải được file mẫu qua fetch (chế độ offline file://), sử dụng bản nhúng sẵn:', err);
  }

  try {
    if (typeof getEmbeddedTemplateBuffer === 'function') {
      const buffer = getEmbeddedTemplateBuffer();
      state.templateBuffer = buffer;
      state.templateWorkbook = XLSX.read(buffer, { type: 'array', cellFormula: true, cellStyles: true });
      extractValidMaDVFromTemplate(state.templateWorkbook);
      console.log('Đã nạp file mẫu từ dữ liệu nhúng sẵn');
    }
  } catch (err) {
    console.error('Lỗi khi nạp file mẫu nhúng:', err);
  }
}

// Chức năng nạp file có dữ liệu sẵn có (có thể chứa 1 hoặc nhiều mảng nghiệp vụ)
async function loadMultiSectorFile(file) {
  try {
    showToast(`Đang phân tích file dữ liệu: ${file.name}...`, 'info');
    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array', cellFormula: true, cellStyles: true });
    const fileObj = {
      file: file,
      name: file.name,
      size: file.size,
      buffer: arrayBuffer,
      workbook: workbook
    };

    if (state.validMaDVSet.size === 0 && state.templateWorkbook) {
      extractValidMaDVFromTemplate(state.templateWorkbook);
    }

    const loadedMangs = [];
    let totalItemsLoaded = 0;

    for (const mang of MANG_CONFIG) {
      const items = extractItemsForMang(mang, fileObj);
      if (items && items.length > 0) {
        state.files[mang.code] = fileObj;
        state.extractedByMang[mang.code] = items;
        updateUploadBoxUI(mang.code, file.name, file.size);
        loadedMangs.push(`${mang.name} (${items.length} dòng)`);
        totalItemsLoaded += items.length;
      }
    }

    if (loadedMangs.length === 0) {
      showToast(`Không tìm thấy dữ liệu mảng nào trong file ${file.name}! Vui lòng kiểm tra lại cấu trúc sheet.`, 'warning');
      return;
    }

    // Xóa cache kết quả xuất để tạo mới theo dữ liệu vừa nạp
    state.exportBlob = null;
    rebuildExtractedData();
    showToast(`Đã nạp thành công ${loadedMangs.length} mảng: ${loadedMangs.join(', ')} (Tổng: ${totalItemsLoaded} dòng)!`, 'success');
  } catch (err) {
    console.error('Lỗi khi nạp file dữ liệu sẵn có:', err);
    showToast(`Lỗi khi đọc file: ${err.message}`, 'error');
  }
}
window.loadMultiSectorFile = loadMultiSectorFile;

// Toast thông báo
function showToast(message, type = 'info') {
  const container = document.getElementById('toastContainer');
  if (!container) return;
  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;
  let iconHtml = '<i data-lucide="info" class="w-4 h-4 text-indigo-600"></i>';
  if (type === 'success') iconHtml = '<i data-lucide="check-circle-2" class="w-4 h-4 text-emerald-600"></i>';
  if (type === 'error') iconHtml = '<i data-lucide="alert-circle" class="w-4 h-4 text-rose-600"></i>';
  if (type === 'warning') iconHtml = '<i data-lucide="alert-triangle" class="w-4 h-4 text-amber-600"></i>';
  toast.innerHTML = `<span style="display:inline-flex;align-items:center;">${iconHtml}</span><span>${message}</span>`;
  container.appendChild(toast);
  if (window.lucide) lucide.createIcons();
  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateX(100%)';
    toast.style.transition = 'all 0.3s ease';
    setTimeout(() => toast.remove(), 300);
  }, 4000);
}

function normalizeVietnameseSearchText(value) {
  return String(value || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd');
}

// Chuyển tab
function initTabs() {
  const tabBtns = document.querySelectorAll('.tab-btn');
  tabBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      tabBtns.forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
      setTimeout(() => { if (window.lucide) lucide.createIcons(); }, 10);
      btn.classList.add('active');
      const tabId = btn.getAttribute('data-tab');
      const target = document.getElementById(tabId);
      if (target) target.classList.add('active');
    });
  });
}

// ==================== XỬ LÝ UPLOAD FILE 7 MẢNG & FILE DỮ LIỆU CÓ SẴN ====================

function initUploadHandlers() {
  // 1. Xử lý nút và input "Nạp file có sẵn dữ liệu (1 hoặc nhiều mảng)"
  const inputMulti = document.getElementById('inputMultiFile');
  if (inputMulti) {
    inputMulti.addEventListener('change', async (e) => {
      if (e.target.files && e.target.files[0]) {
        await loadMultiSectorFile(e.target.files[0]);
        inputMulti.value = '';
      }
    });
  }

  // Hỗ trợ kéo thả file dữ liệu vào toàn bộ khu vực Section 1 (uploadCard)
  const uploadCard = document.getElementById('uploadCard');
  if (uploadCard) {
    uploadCard.addEventListener('dragover', (e) => {
      e.preventDefault();
      e.stopPropagation();
      uploadCard.classList.add('drag-over-card');
    });

    uploadCard.addEventListener('dragleave', (e) => {
      e.preventDefault();
      e.stopPropagation();
      uploadCard.classList.remove('drag-over-card');
    });

    uploadCard.addEventListener('drop', async (e) => {
      // Nếu thả vào một ô mảng cụ thể thì để ô đó tự xử lý
      if (e.target.closest('.upload-box')) return;
      e.preventDefault();
      e.stopPropagation();
      uploadCard.classList.remove('drag-over-card');

      if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
        if (e.dataTransfer.files.length === 1) {
          await loadMultiSectorFile(e.dataTransfer.files[0]);
        } else {
          // Nhiều file thả cùng lúc: tự động phân bổ theo mảng hoặc nạp tổng hợp
          for (let i = 0; i < e.dataTransfer.files.length; i++) {
            const f = e.dataTransfer.files[i];
            const detected = detectMangFromFileName(f.name);
            if (detected) {
              await assignFileToKey(detected, f);
            } else {
              await loadMultiSectorFile(f);
            }
          }
        }
      }
    });
  }

  // 2. Xử lý 7 ô mảng đầu vào
  const keys = ['VT', 'ML', 'CDBR', 'CNTT', 'TD', 'CD', 'HT'];

  keys.forEach(key => {
    const box = document.getElementById(`box_${key}`);
    const input = document.getElementById(`input_${key}`);

    if (input) {
      input.addEventListener('change', (e) => {
        if (e.target.files && e.target.files[0]) {
          assignFileToKey(key, e.target.files[0]);
        }
      });
    }

    if (box) {
      // Hỗ trợ kéo thả trực tiếp vào ô của mảng
      box.addEventListener('dragover', (e) => {
        e.preventDefault();
        e.stopPropagation();
        box.classList.add('drag-over');
      });

      box.addEventListener('dragleave', (e) => {
        e.preventDefault();
        e.stopPropagation();
        box.classList.remove('drag-over');
      });

      box.addEventListener('drop', (e) => {
        e.preventDefault();
        e.stopPropagation();
        box.classList.remove('drag-over');

        if (e.dataTransfer.files && e.dataTransfer.files[0]) {
          assignFileToKey(key, e.dataTransfer.files[0]);
        }
      });
    }
  });
}

// Tự động nhận diện mảng từ tên file nếu người dùng chọn nhiều file
function detectMangFromFileName(fileName) {
  const name = fileName.toUpperCase();
  if (name.includes('_TD_') || name.includes('TRUYEN DAN') || name.includes('TRUYỀN DẪN') || name.includes('_TD') || name.includes('TD_IP')) return 'TD';
  if (name.includes('_CNTT_') || name.includes('CONG NGHE THONG TIN') || name.includes('CNTT')) return 'CNTT';
  if (name.includes('_VT_') || name.includes('VO TUYEN') || name.includes('VÔ TUYẾN')) return 'VT';
  if (name.includes('_ML_') || name.includes('MANG LOI') || name.includes('MẠNG LÕI')) return 'ML';
  if (name.includes('_CDBR_') || name.includes('CO DINH') || name.includes('CỐ ĐỊNH') || name.includes('CDBR')) return 'CDBR';
  if (name.includes('_CD_') || name.includes('CO DIEN') || name.includes('CƠ ĐIỆN')) return 'CD';
  if (name.includes('_HT_') || name.includes('HA TANG') || name.includes('HẠ TẦNG')) return 'HT';
  return null;
}

// Gán file cho mảng: TỰ ĐỘNG TRÍCH XUẤT VÀ THAY THẾ DỮ LIỆU CỦA MẢNG ĐÓ
async function assignFileToKey(key, file) {
  try {
    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array', cellFormula: true, cellStyles: true });

    state.files[key] = {
      file: file,
      name: file.name,
      size: file.size,
      buffer: arrayBuffer,
      workbook: workbook
    };

    if (state.validMaDVSet.size === 0) {
      extractValidMaDVFromTemplate(workbook);
    }

    // Trích xuất cây hạng mục của đúng mảng này
    const mang = MANG_CONFIG.find(m => m.code === key);
    if (mang) {
      const items = extractItemsForMang(mang, state.files[key]);
      // THAY THẾ hoàn toàn dữ liệu cũ của mảng này:
      state.extractedByMang[key] = items;
      showToast(`Đã nhận diện và cập nhật ${items.length} dòng cho mảng [${mang.name}]!`, 'success');
    }

    updateUploadBoxUI(key, file.name, file.size);

    // Tự động tổng hợp lại toàn bộ dữ liệu 7 mảng
    rebuildExtractedData();
  } catch (err) {
    console.error('Lỗi nạp file:', err);
    showToast(`Lỗi khi đọc file ${file.name}: ${err.message}`, 'error');
  }
}

// Xóa file của 1 mảng
function removeFile(key) {
  state.files[key] = null;
  state.extractedByMang[key] = []; // Xóa trắng dữ liệu mảng tương ứng

  const box = document.getElementById(`box_${key}`);
  const infoBar = document.getElementById(`info_${key}`);
  if (box) box.classList.remove('has-file');
  if (infoBar) {
    infoBar.innerHTML = `
      <span class="file-name-text" id="name_${key}" style="color: var(--text-muted);">Chưa có file</span>
      <button class="btn btn-sm" onclick="event.stopPropagation(); document.getElementById('input_${key}').click()">Chọn file</button>
    `;
  }

  rebuildExtractedData();
  showToast(`Đã hủy file của mảng [${key}]`, 'info');
}

// Cập nhật giao diện ô upload sau khi nhận file
function updateUploadBoxUI(key, fileName, sizeBytes) {
  const box = document.getElementById(`box_${key}`);
  const infoBar = document.getElementById(`info_${key}`);
  if (box) box.classList.add('has-file');
  const sizeKb = (sizeBytes / 1024).toFixed(1);
  if (infoBar) {
    infoBar.innerHTML = `
      <div style="display: flex; flex-direction: column; overflow: hidden; max-width: 125px;">
        <span class="file-name-text" style="color: #047857; font-weight: 700;" title="${escapeHtml(fileName)}">${escapeHtml(fileName)}</span>
        <span style="font-size: 0.725rem; color: #64748b; font-weight: 600;">${sizeKb} KB</span>
      </div>
      <button class="btn-remove-file" title="Xóa file mảng này" onclick="event.stopPropagation(); removeFile('${key}')">✕</button>
    `;
  }
}

// Tự động ghép nối dữ liệu 7 mảng theo thứ tự chuẩn
function rebuildExtractedData() {
  state.extractedData = [];
  state.validationIssues = [];

  // Ghép các mảng theo đúng thứ tự chuẩn: VT -> ML -> CDBR -> CNTT -> TD -> CD -> HT
  for (const mang of MANG_CONFIG) {
    const items = state.extractedByMang[mang.code] || [];
    if (items.length === 0) continue;

    // 1. Tạo dòng Cấp 1 (Tiêu đề Mảng) cho mảng này
    const hInfo = state.mangHeaderInfo[mang.code] || {};
    const actualTT = (hInfo.origTT !== undefined && hInfo.origTT !== '') ? hInfo.origTT : (hInfo.tt || mang.tt);
    const actualND = (hInfo.nd !== undefined && hInfo.nd !== '') ? hInfo.nd : mang.name.toUpperCase();

    const cap1Item = {
      mangCode: mang.code,
      mangName: mang.name,
      fileName: (state.files[mang.code] ? state.files[mang.code].name : ''),
      isMangHeader: true,
      isGroup: true,
      origRow: hInfo.origRow || 10,
      tt: actualTT,
      origTT: actualTT,
      nd: actualND,
      levelNum: 1,
      level: 'CẤP 1',
      hasSubtotal: true,
      subtotalStartRow: hInfo.startRow || (items[0] ? items[0].origRow : 0),
      subtotalEndRow: hInfo.endRow || (items[items.length - 1] ? items[items.length - 1].origRow : 0),
      subtotalG: hInfo.formulaG || '-',
      subtotalH: hInfo.formulaH || '-',
      tt27Formula: hInfo.formulaG || null,
      tt28Formula: hInfo.formulaH || null,
      rangeG: hInfo.rangeG || null,
      rangeH: hInfo.rangeH || null,
      kl27: null,
      kl28: null,
      dg: null,
      dvt: '',
      groupKey: 'MANG_' + mang.code,
      parentKey: null
    };

    validateRow(cap1Item);
    state.extractedData.push(cap1Item);

    // 2. Thêm các dòng con của mảng
    items.forEach(item => {
      validateRow(item);
      state.extractedData.push(item);
    });
  }

  // Đánh lại số thứ tự index
  state.extractedData.forEach((item, idx) => {
    item.index = idx;
    if (item.origTT === undefined) {
      item.origTT = item.tt;
    }
  });

  autoDetectGroupsHierarchy();
  updateValidationKPIs();
  renderValidationTable();
  renderPreviewTable();
  renderHierarchyTable();
  renderStrategyComparisonTable();

  const hasData = state.extractedData.length > 0;
  const btnDownload = document.getElementById('btnOpenResultFile');
  if (btnDownload) {
    btnDownload.disabled = !hasData;
    btnDownload.title = hasData ? 'Tải file kết quả hoàn chỉnh về máy tính (Masterlist 2027-2028_KQ_ddmmyyyy_hhmm.xlsx)' : 'Vui lòng nạp ít nhất một file mảng để tải kết quả';
  }
  const btnSaveServer = document.getElementById('btnSaveToServer');
  if (btnSaveServer) {
    btnSaveServer.disabled = !hasData;
    btnSaveServer.title = hasData ? 'Lưu bản hiện tại lên Server để sau này nạp lại và tiếp tục chỉnh sửa' : 'Vui lòng nạp dữ liệu trước khi lưu lên Server';
  }
  const banner = document.getElementById('exportReadyBanner');
  if (banner) banner.style.display = hasData ? 'flex' : 'none';
  if (hasData) prepareQuickDownload();
}



// Trích xuất danh mục Mã DV hợp lệ từ sheet "TH theo DV"
function extractValidMaDVFromTemplate(workbook) {
  state.validMaDVSet.clear();
  const sheetName = workbook.SheetNames.find(s => s.toLowerCase().includes('th theo dv') || s.toLowerCase().includes('dv'));
  if (!sheetName) return;

  const ws = workbook.Sheets[sheetName];
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:M100');

  // Cột C (col index 2) là cột Mã DV
  for (let r = 2; r <= range.e.r; r++) {
    const cell = ws[XLSX.utils.encode_cell({ r: r, c: 2 })];
    if (cell && cell.v !== undefined && cell.v !== null) {
      const val = String(cell.v).trim();
      if (val && val !== 'Mã DV' && val !== 'Ko xóa' && val !== 'Nội dung') {
        state.validMaDVSet.add(val);
      }
    }
  }
  console.log('Loaded valid Mã DV set from template:', state.validMaDVSet);
}

// ==================== ENGINE TRÍCH XUẤT & SO KHỚP ====================

function initActionButtons() {
  const btnProcess = document.getElementById('btnProcess');
  if (btnProcess) btnProcess.addEventListener('click', runProcessingPipeline);

  const btnDownload = document.getElementById('btnOpenResultFile');
  if (btnDownload) {
    btnDownload.addEventListener('click', openResultFile);
  }

  const btnSaveServer = document.getElementById('btnSaveToServer');
  if (btnSaveServer) {
    btnSaveServer.addEventListener('click', openSaveDraftModal);
  }

  const btnConfirmSave = document.getElementById('btnConfirmSaveDraft');
  if (btnConfirmSave) {
    btnConfirmSave.addEventListener('click', confirmSaveDraft);
  }

  const btnConfirmSaveItm = document.getElementById('btnConfirmSaveItem');
  if (btnConfirmSaveItm) {
    btnConfirmSaveItm.addEventListener('click', confirmSaveItem);
  }

  const searchDraftsInput = document.getElementById('searchDraftsInput');
  if (searchDraftsInput) {
    searchDraftsInput.addEventListener('input', renderServerDraftsTable);
  }

  const btnReset = document.getElementById('btnReset');
  if (btnReset) {
    btnReset.addEventListener('click', () => {
      if (state.extractedData && state.extractedData.length > 0) {
        if (confirm('Bạn có chắc chắn muốn làm mới toàn bộ dữ liệu để bắt đầu bản khác?')) {
          location.reload();
        }
      } else {
        location.reload();
      }
    });
  }
  const btnApply = document.getElementById('btnApplyHierarchy');
  if (btnApply) {
    btnApply.addEventListener('click', () => {
      applyHierarchyConfiguration();
      showToast('Đã cập nhật công thức Subtotal và cấu hình phân cấp!', 'success');
    });
  }

  const btnReindex = document.getElementById('btnReindexTT');
  if (btnReindex) {
    btnReindex.addEventListener('click', reindexAllItems);
  }

  const btnFixOutline = document.getElementById('btnFixOutlineGroups');
  if (btnFixOutline) {
    btnFixOutline.addEventListener('click', fixOutlineGroups);
  }

  const btnRestoreTT = document.getElementById('btnRestoreOriginalTT');
  if (btnRestoreTT) {
    btnRestoreTT.addEventListener('click', restoreOriginalTT);
  }

  const btnSplitWalletDV = document.getElementById('btnSplitWalletDV');
  if (btnSplitWalletDV) {
    btnSplitWalletDV.addEventListener('click', () => applyWalletSplit('dv'));
  }

  const btnSplitWalletMang = document.getElementById('btnSplitWalletMang');
  if (btnSplitWalletMang) {
    btnSplitWalletMang.addEventListener('click', () => applyWalletSplit('mang'));
  }

  // Nút Mở file KQ ở Header
  const btnOpenFile = document.getElementById('btnOpenResultFile') || document.getElementById('btnOpenResultFolder');
  if (btnOpenFile) {
    btnOpenFile.addEventListener('click', openResultFile);
  }

  const btnCopyFolder = document.getElementById('btnCopyFolderPath');
  if (btnCopyFolder) {
    btnCopyFolder.addEventListener('click', () => {
      const input = document.getElementById('inputFolderPath');
      if (input) {
        input.select();
        navigator.clipboard.writeText(input.value);
        showToast('Đã sao chép đường dẫn: ' + input.value, 'success');
      }
    });
  }

  // Mở rộng tất cả & Thu gọn tất cả
  const btnExpandAll = document.getElementById('btnExpandAllGroups');
  if (btnExpandAll) {
    btnExpandAll.addEventListener('click', () => {
      state.collapsedGroupKeys.clear();
      renderHierarchyTable();
      showToast('Đã mở rộng toàn bộ các cấp!', 'info');
    });
  }

  const btnCollapseAll = document.getElementById('btnCollapseAllGroups');
  if (btnCollapseAll) {
    btnCollapseAll.addEventListener('click', () => {
      state.collapsedGroupKeys.clear();
      state.groupsConfig.forEach(g => {
        if (g.hasChildren) state.collapsedGroupKeys.add(g.key);
      });
      renderHierarchyTable();
      showToast('Đã thu gọn toàn bộ các cấp!', 'info');
    });
  }
}


// ==================== CHỨC NĂNG MỞ FILE KẾT QUẢ TRỰC TIẾP ====================

// Sinh tên file kết quả kèm dấu thời gian: Masterlist 2027-2028_KQ_ddmmyyyy_hhmm.xlsx
function getTimestampedExportFileName() {
  const now = new Date();
  const dd = String(now.getDate()).padStart(2, '0');
  const mm = String(now.getMonth() + 1).padStart(2, '0');
  const yyyy = now.getFullYear();
  const hh = String(now.getHours()).padStart(2, '0');
  const min = String(now.getMinutes()).padStart(2, '0');
  return `Masterlist 2027-2028_KQ_${dd}${mm}${yyyy}_${hh}${min}.xlsx`;
}

// Lấy buffer / blob của file kết quả mới nhất
async function getResultFileBlob() {
  // 1. Ưu tiên blob đã được build gần nhất trong phiên làm việc
  if (state.exportBlob) {
    return state.exportBlob;
  }
  // 2. Nếu đã có dữ liệu bóc tách từ các mảng, tự động build buffer mới nhất
  if (state.extractedData && state.extractedData.length > 0 && state.templateBuffer) {
    try {
      recalculateSubtotalFormulas();
      const buffer = await buildCleanMasterlist(state.templateBuffer, state.extractedByMang, state.files);
      state.exportBlob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      return state.exportBlob;
    } catch (e) {
      console.warn('Lỗi tự động tạo buffer kết quả:', e);
    }
  }
  // 3. Dự phòng từ buffer template trong bộ nhớ
  if (state.templateBuffer) {
    return new Blob([state.templateBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }
  if (typeof getEmbeddedTemplateBuffer === 'function') {
    const buf = getEmbeddedTemplateBuffer();
    if (buf) return new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }
  return null;
}

// Bấm "Download Excel": Tải file trực tiếp về máy tính với tên Masterlist 2027-2028_KQ_ddmmyyyy_hhmm.xlsx
async function openResultFile() {
  const downloadFileName = getTimestampedExportFileName();

  const blob = await getResultFileBlob();
  if (blob) {
    triggerDownloadBlob(blob, downloadFileName);
    showToast(`Đã tải về file: ${downloadFileName}`, 'success');
  } else {
    showToast('Chưa tìm thấy file kết quả để tải về!', 'error');
  }
}

function openResultFolder() {
  openResultFile();
}

window.openResultFile = openResultFile;
window.openResultFolder = openResultFile;

window.closeFolderModal = function() {
  const modal = document.getElementById('folderModal');
  if (modal) modal.style.display = 'none';
};

window.openFolderViaProtocol = function() {
  openResultFile();
};

// Tải hoặc mở file kết quả
function downloadOrExportResultFile() {
  openResultFile();
}

// Pipeline chính khi bấm nút "⚡ Tổng hợp & Kiểm tra"
async function runProcessingPipeline() {
  const btn = document.getElementById('btnProcess');
  btn.disabled = true;
  btn.innerHTML = `<span class="spinner"></span> Đang xử lý...`;

  try {
    if (!state.templateWorkbook) {
      for (const k of ['TD', 'CNTT', 'VT', 'ML', 'CDBR', 'CD', 'HT']) {
        if (state.files[k]) {
          extractValidMaDVFromTemplate(state.files[k].workbook);
          if (state.validMaDVSet.size > 0) break;
        }
      }
    }

    const hasAnyInput = MANG_CONFIG.some(m => state.files[m.code] !== null);
    if (!hasAnyInput) {
      showToast('Vui lòng nạp ít nhất một file mảng đầu vào!', 'error');
      btn.disabled = false;
      btn.innerHTML = `<span>⚡</span> Tổng hợp & Kiểm tra`;
      return;
    }

    // Quét và cập nhật dữ liệu cho từng mảng được nạp
    for (const mang of MANG_CONFIG) {
      const fileObj = state.files[mang.code];
      if (fileObj) {
        state.extractedByMang[mang.code] = extractItemsForMang(mang, fileObj);
      }
    }

    rebuildExtractedData();
    showToast(`Đã tổng hợp và so khớp thành công ${state.extractedData.length} dòng dữ liệu!`, 'success');
  } catch (err) {
    console.error('Lỗi pipeline xử lý:', err);
    showToast(`Có lỗi xảy ra: ${err.message}`, 'error');
  } finally {
    btn.disabled = false;
    btn.innerHTML = `<i data-lucide="zap" class="w-5 h-5 text-amber-300"></i><span>Tổng hợp & Kiểm tra</span>`;
    if (window.lucide) lucide.createIcons();
  }
}

// Trích xuất dải dòng phạm vi từ công thức Subtotal/Sum bất kỳ (hỗ trợ $, 109, SUM, nhiều dải ô, dấu cộng...)
function parseSubtotalRange(formula) {
  if (!formula || typeof formula !== 'string') return null;
  const clean = formula.trim().toUpperCase();

  // Bỏ qua các công thức tính thành tiền đơn giá * khối lượng (=F*D hoặc =F*E)
  if (clean.includes('*') && !clean.includes('SUBTOTAL') && !clean.includes('SUM')) return null;

  const hasSubtotal = clean.includes('SUBTOTAL');
  const hasSum = clean.includes('SUM');
  const hasPlus = clean.includes('+');

  if (!hasSubtotal && !hasSum && !hasPlus) return null;

  // Tìm các dải ô dạng G12:G50 hoặc $G$12:$G$50 hoặc G$12:G$50
  const rangeRegex = /(?:[A-Z\$]+)(\d+)\s*:\s*(?:[A-Z\$]+)(\d+)/g;
  let match;
  const ranges = [];
  let minRow = Infinity;
  let maxRow = -Infinity;

  while ((match = rangeRegex.exec(clean)) !== null) {
    const r1 = parseInt(match[1], 10);
    const r2 = parseInt(match[2], 10);
    if (!isNaN(r1) && !isNaN(r2)) {
      const s = Math.min(r1, r2);
      const e = Math.max(r1, r2);
      ranges.push({ start: s, end: e });
      minRow = Math.min(minRow, s);
      maxRow = Math.max(maxRow, e);
    }
  }

  // Nếu không có dải ô (dạng A:B), kiểm tra các ô đơn lẻ được cộng: G12 + G15 hoặc SUM(G12, G15)
  if (ranges.length === 0 && (hasSum || hasPlus || hasSubtotal)) {
    const cellRegex = /[A-Z\$]+(\d+)/g;
    const cells = [];
    while ((match = cellRegex.exec(clean)) !== null) {
      const r = parseInt(match[1], 10);
      if (!isNaN(r) && r > 0) {
        cells.push(r);
        minRow = Math.min(minRow, r);
        maxRow = Math.max(maxRow, r);
      }
    }
    if (cells.length > 0) {
      cells.sort((a, b) => a - b);
      ranges.push({ start: cells[0], end: cells[cells.length - 1], discrete: cells });
    }
  }

  if (ranges.length > 0 && minRow !== Infinity && maxRow !== -Infinity) {
    const parts = ranges.map(r => r.discrete ? r.discrete.join('+') : (r.start === r.end ? String(r.start) : `${r.start}:${r.end}`));
    const rangeStr = parts.join(', ');
    return {
      startRow: minRow,
      endRow: maxRow,
      rangeStr: rangeStr,
      ranges: ranges,
      formula: formula,
      clean: clean,
      isAggregate: true
    };
  }

  return null;
}

// Trích xuất cây hạng mục cho 1 mảng từ file đầu vào
function extractItemsForMang(mang, fileObj) {
  const wb = fileObj.workbook;
  const sheetName = wb.SheetNames.find(s => s.toUpperCase().includes('PL1.1') || s.toUpperCase().includes('ML2027')) || wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:Z3000');

  // Tìm dòng tiêu đề mảng ở Cột B (col index 1)
  let headerRow = -1;
  for (let r = 7; r <= range.e.r; r++) {
    const cellB = ws[XLSX.utils.encode_cell({ r: r, c: 1 })];
    if (!cellB || cellB.v === undefined) continue;
    const textB = String(cellB.v).trim().toUpperCase();

    const isMatch = mang.keywords.some(kw => textB === kw || (textB.includes(kw) && textB.length < kw.length + 5));
    if (isMatch) {
      headerRow = r;
      break;
    }
  }

  if (headerRow === -1) {
    console.warn(`Không tìm thấy tiêu đề cho mảng ${mang.name} trong file ${fileObj.name}`);
    return [];
  }

  // Tìm dòng kết thúc (trước tiêu đề mảng cấp cao kế tiếp)
  let endRow = range.e.r;

  const cellGHeader = ws[XLSX.utils.encode_cell({ r: headerRow, c: 6 })];
  const formulaGHeader = cellGHeader && (cellGHeader.f ? '=' + cellGHeader.f : (typeof cellGHeader.v === 'string' && cellGHeader.v.startsWith('=') ? cellGHeader.v : ''));
  const cellHHeader = ws[XLSX.utils.encode_cell({ r: headerRow, c: 7 })];
  const formulaHHeader = cellHHeader && (cellHHeader.f ? '=' + cellHHeader.f : (typeof cellHHeader.v === 'string' && cellHHeader.v.startsWith('=') ? cellHHeader.v : ''));

  const rangeGHeader = parseSubtotalRange(formulaGHeader);
  const rangeHHeader = parseSubtotalRange(formulaHHeader);
  const rangeHeader = rangeGHeader || rangeHHeader;
  if (rangeHeader && rangeHeader.endRow) {
    const parsedEndRow = rangeHeader.endRow - 1;
    if (parsedEndRow > headerRow && parsedEndRow <= range.e.r) {
      endRow = parsedEndRow;
      console.log(`Đã xác định phạm vi mảng ${mang.code} từ công thức Subtotal: dòng ${headerRow + 2} đến ${endRow + 1}`);
    }
  }

  if (endRow === range.e.r) {
    for (let r = headerRow + 1; r <= range.e.r; r++) {
      const cellA = ws[XLSX.utils.encode_cell({ r: r, c: 0 })];
      const cellB = ws[XLSX.utils.encode_cell({ r: r, c: 1 })];
      const valA = cellA && cellA.v !== undefined ? String(cellA.v).trim().toUpperCase() : '';
      const valB = cellB && cellB.v !== undefined ? String(cellB.v).trim().toUpperCase() : '';

      const isTopTT = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'P1', 'P2', 'P3'].includes(valA);
      const isMangKw = ALL_SECTION_HEADERS.some(kw => valB === kw || (valB.startsWith(kw) && valB.length < kw.length + 5));

      if (isTopTT && isMangKw) {
        endRow = r - 1;
        break;
      }
    }
  }

  const cellAHeader = ws[XLSX.utils.encode_cell({ r: headerRow, c: 0 })];
  const headerTT = cellAHeader ? (cellAHeader.w !== undefined && cellAHeader.w !== null ? String(cellAHeader.w).trim() : (cellAHeader.v !== undefined && cellAHeader.v !== null ? String(cellAHeader.v).trim() : '')) : '';
  const cellBHeader = ws[XLSX.utils.encode_cell({ r: headerRow, c: 1 })];
  const headerND = cellBHeader ? (cellBHeader.w !== undefined && cellBHeader.w !== null ? String(cellBHeader.w).trim() : (cellBHeader.v !== undefined && cellBHeader.v !== null ? String(cellBHeader.v).trim() : '')) : '';

  state.mangHeaderInfo[mang.code] = {
    origRow: headerRow + 1,
    tt: headerTT || mang.tt,
    origTT: headerTT || mang.tt,
    nd: headerND || mang.name.toUpperCase(),
    startRow: rangeHeader ? rangeHeader.startRow : headerRow + 2,
    endRow: endRow + 1,
    formulaG: formulaGHeader || (rangeHeader ? `=SUBTOTAL(9,G${headerRow + 2}:G${endRow + 1})` : '-'),
    formulaH: formulaHHeader || (rangeHeader ? `=SUBTOTAL(9,H${headerRow + 2}:H${endRow + 1})` : '-'),
    rangeG: rangeGHeader,
    rangeH: rangeHHeader
  };

  const items = [];
  for (let r = headerRow + 1; r <= endRow; r++) {
    const rowData = getRowValues(ws, r);
    if (!rowData.nd && !rowData.tt && !rowData.kl27 && !rowData.kl28 && !rowData.tt27Formula && !rowData.tt28Formula) continue; // Dòng trống

    // Parse công thức Subtotal / Sum ở cột G hoặc H
    const rangeG = parseSubtotalRange(rowData.tt27Formula);
    const rangeH = parseSubtotalRange(rowData.tt28Formula);
    const subMatch = rangeG || rangeH;
    const hasSubtotal = !!subMatch;
    const subtotalStartRow = subMatch ? subMatch.startRow : null;
    const subtotalEndRow = subMatch ? subMatch.endRow : null;

    const hasKL = (rowData.kl27 !== null && rowData.kl27 !== 0) || (rowData.kl28 !== null && rowData.kl28 !== 0);
    const hasDG = (rowData.dg !== null && rowData.dg !== 0);
    const isLeaf = (rowData.tt && rowData.tt.trim() === '-') || (rowData.dvt && rowData.dvt.trim() !== '') || hasDG || hasKL;

    const isGroup = hasSubtotal || (!isLeaf && rowData.tt.trim() !== '-' && (
      (rowData.tt && /^[A-Z0-9\.]+$/i.test(rowData.tt.trim().replace(/[\.\)\:\s]+$/, ''))) ||
      (rowData.nd && rowData.nd.trim() === rowData.nd.trim().toUpperCase() && rowData.nd.trim().length > 3)
    ));

    const item = {
      mangCode: mang.code,
      mangName: mang.name,
      fileName: fileObj.name,
      origRow: r + 1,
      tt: rowData.tt,
      origTT: rowData.tt,
      nd: rowData.nd,
      dvt: rowData.dvt,
      kl27: rowData.kl27,
      kl28: rowData.kl28,
      dg: rowData.dg,
      tt27: rowData.tt27,
      tt28: rowData.tt28,
      tt27Formula: rowData.tt27Formula,
      tt28Formula: rowData.tt28Formula,
      hasSubtotal: hasSubtotal,
      subtotalStartRow: subtotalStartRow,
      subtotalEndRow: subtotalEndRow,
      rangeG: rangeG,
      rangeH: rangeH,
      donViDT: rowData.donViDT,
      maMang: rowData.maMang, // Giữ nguyên giá trị gốc để check thiếu mã
      maDV: rowData.maDV,
      maLoai: rowData.maLoai,
      // Phân loại
      isGroup: isGroup,
      levelNum: isGroup ? 3 : 99,
      level: isGroup ? 'CẤP 3' : 'CHI TIẾT',
      subtotalG: rowData.tt27Formula ? rowData.tt27Formula : (hasSubtotal && subtotalStartRow && subtotalEndRow ? `=SUBTOTAL(9,G${subtotalStartRow}:G${subtotalEndRow})` : '-'),
      subtotalH: rowData.tt28Formula ? rowData.tt28Formula : (hasSubtotal && subtotalStartRow && subtotalEndRow ? `=SUBTOTAL(9,H${subtotalStartRow}:H${subtotalEndRow})` : '-'),
      parentIndex: -1
    };


    items.push(item);
  }

  return items;
}

// Đọc giá trị và công thức của 1 dòng
function getRowValues(ws, r) {
  const getCell = (c) => ws[XLSX.utils.encode_cell({ r: r, c: c })] || null;

  const cellA = getCell(0); // TT
  const cellB = getCell(1); // ND
  const cellC = getCell(2); // DVT
  const cellD = getCell(3); // KL 2027
  const cellE = getCell(4); // KL 2028
  const cellF = getCell(5); // Don gia USD
  const cellG = getCell(6); // Thanh tien 2027
  const cellH = getCell(7); // Thanh tien 2028
  const cellI = getCell(8); // Don vi DTMS
  const cellJ = getCell(9); // Ma mang
  const cellK = getCell(10); // Ma DV
  const cellL = getCell(11); // Ma loai

  const getCellText = (cell) => {
    if (!cell) return '';
    if (cell.w !== undefined && cell.w !== null) return String(cell.w).trim();
    if (cell.v !== undefined && cell.v !== null) return String(cell.v).trim();
    return '';
  };

  const parseNum = (cell) => {
    if (!cell || cell.v === undefined || cell.v === null || cell.v === '') return null;
    const n = parseFloat(cell.v);
    return isNaN(n) ? null : n;
  };

  const getFormula = (cell) => {
    if (!cell) return null;
    if (cell.f) return '=' + cell.f;
    if (typeof cell.v === 'string' && cell.v.startsWith('=')) return cell.v;
    return null;
  };

  return {
    tt: getCellText(cellA),
    nd: getCellText(cellB),
    dvt: getCellText(cellC),
    kl27: parseNum(cellD),
    kl28: parseNum(cellE),
    dg: parseNum(cellF),
    tt27: parseNum(cellG),
    tt28: parseNum(cellH),
    tt27Formula: getFormula(cellG),
    tt28Formula: getFormula(cellH),
    donViDT: getCellText(cellI),
    maMang: getCellText(cellJ),
    maDV: getCellText(cellK),
    maLoai: getCellText(cellL)
  };
}

// ==================== ENGINE KIỂM TRA & SO KHỚP LỖI ====================

function validateRow(item) {
  const hasKL = (item.kl27 !== null && item.kl27 !== 0) || (item.kl28 !== null && item.kl28 !== 0);

  // 1. Kiểm tra lỗi thành tiền: Cột G, H phải bằng đúng Khối lượng * Đơn giá
  const dg = item.dg || 0;

  if (item.kl27 !== null && item.kl27 !== 0) {
    const expected27 = item.kl27 * dg;
    if (item.tt27 !== null) {
      if (Math.abs(expected27 - item.tt27) > 1) {
        addIssue(item, 'MATH', 'G (2027)', item.tt27, expected27,
          `Thành tiền 2027 (${formatNumber(item.tt27)}) không bằng KL*ĐG (${formatNumber(expected27)}). Cần kiểm tra công thức.`);
      }
    }
  }

  if (item.kl28 !== null && item.kl28 !== 0) {
    const expected28 = item.kl28 * dg;
    if (item.tt28 !== null) {
      if (Math.abs(expected28 - item.tt28) > 1) {
        addIssue(item, 'MATH', 'H (2028)', item.tt28, expected28,
          `Thành tiền 2028 (${formatNumber(item.tt28)}) không bằng KL*ĐG (${formatNumber(expected28)}). Cần kiểm tra công thức.`);
      }
    }
  }

  // 2. Kiểm tra các dòng có khối lượng # 0 (chỉ kiểm tra các hạng mục chi tiết, bỏ qua dòng tổng/nhóm Subtotal)
  if (hasKL && !item.isGroup && !item.hasSubtotal) {
    // 2.1. Kiểm tra thiếu Mã loại (Cột L)
    if (!item.maLoai || !String(item.maLoai).trim()) {
      addIssue(item, 'MISSING_CODE', 'L (Mã loại)', 'Trống', 'VTTB / XL...',
        'Hạng mục có khối lượng nhưng thiếu Mã loại ở cột L (cần bổ sung VTTB, XL...).');
    }

    // 2.2. Kiểm tra thiếu Mã mảng (Cột J)
    if (!item.maMang || !String(item.maMang).trim()) {
      addIssue(item, 'MISSING_CODE', 'J (Mã mảng)', 'Trống', item.mangCode,
        `Hạng mục có khối lượng nhưng thiếu Mã mảng ở cột J (khuyến nghị: ${item.mangCode}).`);
    }

    // 2.3. Kiểm tra thiếu Mã DV (Cột K)
    if (!item.maDV || !String(item.maDV).trim()) {
      addIssue(item, 'MISSING_CODE', 'K (Mã DV)', 'Trống', 'Mã DV hợp lệ',
        'Hạng mục có khối lượng nhưng thiếu Mã DV ở cột K.');
    } else if (!item.skipMaDVValidation && state.validMaDVSet.size > 0 && !state.validMaDVSet.has(item.maDV)) {
      addIssue(item, 'MADV', 'K (Mã DV)', item.maDV, 'Có trong TH theo DV',
        `Mã DV "${item.maDV}" chưa có trong danh mục sheet "TH theo DV". Công thức SUMIFS sẽ bỏ sót dịch vụ này!`);
    }

    // 2.4. Thiếu đơn giá dự kiến
    if (item.dg === null || item.dg === 0) {
      addIssue(item, 'INFO', 'F (Đơn giá)', item.dg ?? 'Trống', '> 0',
        'Dòng có khối lượng nhưng đơn giá dự kiến đang để trống hoặc bằng 0.');
    }

    // 2.5. Thiếu ĐVT
    if (!item.dvt) {
      addIssue(item, 'INFO', 'C (ĐVT)', 'Trống', 'Có giá trị',
        'Hạng mục có khối lượng nhưng thiếu Đơn vị tính.');
    }

    // 2.6. Giá trị âm bất thường
    if ((item.kl27 !== null && item.kl27 < 0) || (item.kl28 !== null && item.kl28 < 0) || (item.dg !== null && item.dg < 0)) {
      addIssue(item, 'INFO', 'D/E/F', 'Giá trị âm', '>= 0',
        'Phát hiện số lượng hoặc đơn giá âm bất thường.');
    }
  }

  // 3. Kiểm tra lỗi công thức Excel (#REF!, #VALUE!...)
  for (const formula of [item.tt27Formula, item.tt28Formula]) {
    if (formula) {
      for (const errCode of ['#REF!', '#VALUE!', '#DIV/0!', '#N/A', '#NAME?']) {
        if (formula.includes(errCode)) {
          addIssue(item, 'FORMULA', 'Công thức', formula, 'Công thức hợp lệ',
            `Ô chứa lỗi công thức Excel: ${errCode}`);
        }
      }
    }
  }

  // 4. Kiểm tra so khớp dải Subtotal giữa Cột G (2027) và Cột H (2028)
  const dG = item.rangeG !== undefined ? item.rangeG : parseSubtotalRange(item.tt27Formula);
  const dH = item.rangeH !== undefined ? item.rangeH : parseSubtotalRange(item.tt28Formula);

  if (dG || dH) {
    if (dG && !dH) {
      addIssue(item, 'SUBTOTAL_MISMATCH', 'G/H (Subtotal)', item.tt27Formula || `Dải [${dG.rangeStr}]`, 'Thiếu ở H',
        `Cột G có công thức Subtotal (dải dòng [${dG.rangeStr}]) nhưng cột H không có công thức Subtotal tương ứng (đang để trống hoặc là giá trị/công thức khác).`);
    } else if (!dG && dH) {
      addIssue(item, 'SUBTOTAL_MISMATCH', 'G/H (Subtotal)', 'Thiếu ở G', item.tt28Formula || `Dải [${dH.rangeStr}]`,
        `Cột H có công thức Subtotal (dải dòng [${dH.rangeStr}]) nhưng cột G không có công thức Subtotal tương ứng (đang để trống hoặc là giá trị/công thức khác).`);
    } else if (dG && dH) {
      const isMismatch = (dG.rangeStr !== dH.rangeStr) || (dG.startRow !== dH.startRow) || (dG.endRow !== dH.endRow);
      if (isMismatch) {
        addIssue(item, 'SUBTOTAL_MISMATCH', 'G/H (Subtotal)', `G: [${dG.rangeStr}]`, `H: [${dH.rangeStr}]`,
          `Lệch dải Subtotal giữa Cột G và Cột H: Cột G tính dải dòng [${dG.rangeStr}] còn Cột H tính dải dòng [${dH.rangeStr}]. Cần kiểm tra và đồng bộ lại công thức.`);
      }
    }
  }
}

function addIssue(item, type, column, currentVal, expectedVal, message) {
  state.validationIssues.push({
    id: state.validationIssues.length + 1,
    type: type, // 'MATH', 'MISSING_CODE', 'MADV', 'INFO', 'FORMULA'
    mangCode: item.mangCode,
    mangName: item.mangName,
    fileName: item.fileName,
    origRow: item.origRow,
    nd: item.nd,
    kl27: item.kl27, // Khối lượng 2027 gốc
    kl28: item.kl28, // Khối lượng 2028 gốc
    dg: item.dg,     // Đơn giá hiện tại
    column: column,
    currentVal: currentVal !== null && currentVal !== undefined ? String(currentVal) : 'Trống',
    expectedVal: expectedVal !== null && expectedVal !== undefined ? String(expectedVal) : '',
    message: message
  });
}

// ==================== ENGINE NHẬN DẠNG & CẤU HÌNH NHÓM HẠNG MỤC (TREE VIEW) ====================

function autoDetectGroupsHierarchy() {
  state.groupsConfig = [];

  for (const mang of MANG_CONFIG) {
    const itemsOfMang = state.extractedData.filter(x => x.mangCode === mang.code);
    if (itemsOfMang.length === 0) continue;

    const hInfo = state.mangHeaderInfo[mang.code] || {};

    // 1. Nút Cấp 1: Tiêu đề Mảng
    const mangKey = 'MANG_' + mang.code;
    const actualTT = (hInfo.origTT !== undefined && hInfo.origTT !== '') ? hInfo.origTT : (hInfo.tt || mang.tt);
    const actualND = (hInfo.nd !== undefined && hInfo.nd !== '') ? hInfo.nd : mang.name.toUpperCase();
    const cap1Item = itemsOfMang.find(x => x.isMangHeader);
    const mangNode = {
      key: mangKey,
      isMangHeader: true,
      mangCode: mang.code,
      mangName: mang.name,
      itemIndex: cap1Item ? cap1Item.index : -1,
      origRow: hInfo.origRow || 10,
      subtotalStartRow: hInfo.startRow || (itemsOfMang[0] ? itemsOfMang[0].origRow : 0),
      subtotalEndRow: hInfo.endRow || (itemsOfMang[itemsOfMang.length - 1] ? itemsOfMang[itemsOfMang.length - 1].origRow : 0),
      tt: actualTT,
      origTT: actualTT,
      nd: actualND,
      levelNum: 1,
      level: 'CẤP 1',
      rank: 1,
      isGroup: true,
      hasSubtotal: true,
      parentKey: null,
      ancestors: [],
      hasChildren: false,
      subtotalG: hInfo.formulaG || '-',
      subtotalH: hInfo.formulaH || '-'
    };

    state.groupsConfig.push(mangNode);

    // Ngăn xếp (stack) để phân cấp lồng nhau chính xác theo công thức Subtotal và TT
    const groupStack = [mangNode];

    // 2. Xác định quan hệ cha-con và cấp bậc chính xác cho từng phần tử
    for (const it of itemsOfMang) {
      if (it.isMangHeader) continue;
      it.groupKey = (it.isGroup ? 'GRP_' : 'ITEM_') + it.index;

      if (it.isGroup) {
        // Pop các nhóm trên stack không thể là cha của it
        while (groupStack.length > 1) {
          const top = groupStack[groupStack.length - 1];

          // a. Nếu nhóm top có Subtotal/endRow và dòng hiện tại đã vượt quá phạm vi Subtotal của top -> top đã kết thúc
          if (top.subtotalEndRow && it.origRow > top.subtotalEndRow) {
            groupStack.pop();
            continue;
          }

          // b. Nếu top và it là 2 nhóm anh em liên tiếp (sequential siblings, vd 1 & 2, I.1 & I.2, a & b, I & II):
          // -> pop top để cùng nhận chung một nhóm cha!
          if (areSequentialSiblings(top.tt, it.tt)) {
            groupStack.pop();
            continue;
          }

          // c. Nếu dòng hiện tại nằm trong phạm vi Subtotal của top:
          // Top CHẮC CHẮN là cha trực tiếp của it (ví dụ 1034 có Subtotal G1035:G1048 thì 1035 là con của 1034)
          if (top.subtotalEndRow && it.origRow <= top.subtotalEndRow) {
            break;
          }

          break;
        }

        const parent = groupStack[groupStack.length - 1];
        let lvl = parent.levelNum + 1;

        it.levelNum = Math.min(6, Math.max(2, lvl));
        it.level = `CẤP ${it.levelNum}`;
        it.parentKey = parent.key;
        it.subtotalG = it.tt27Formula ? it.tt27Formula : (it.hasSubtotal && it.subtotalStartRow && it.subtotalEndRow ? `=SUBTOTAL(9,G${it.subtotalStartRow}:G${it.subtotalEndRow})` : '-');
        it.subtotalH = it.tt28Formula ? it.tt28Formula : (it.hasSubtotal && it.subtotalStartRow && it.subtotalEndRow ? `=SUBTOTAL(9,H${it.subtotalStartRow}:H${it.subtotalEndRow})` : '-');

        const effectiveEndRow = it.subtotalEndRow || parent.subtotalEndRow;

        groupStack.push({
          key: it.groupKey,
          origRow: it.origRow,
          levelNum: it.levelNum,
          subtotalStartRow: it.subtotalStartRow,
          subtotalEndRow: effectiveEndRow,
          hasSubtotal: it.hasSubtotal,
          tt: it.tt
        });
      } else {
        // Dòng chi tiết / vật tư
        while (groupStack.length > 1) {
          const top = groupStack[groupStack.length - 1];
          if (top.subtotalEndRow && it.origRow > top.subtotalEndRow) {
            groupStack.pop();
            continue;
          }
          break;
        }
        const parent = groupStack[groupStack.length - 1];
        it.parentKey = parent.key;
        it.levelNum = 99;
        it.level = 'CHI TIẾT';
        it.subtotalG = '-';
        it.subtotalH = '-';
      }
    }

    // 3. Xây dựng danh sách node phân cấp cho Tree View và đánh dấu hasChildren, ancestors
    for (const it of itemsOfMang) {
      if (it.isMangHeader) continue;

      const parentNode = (it.parentKey === mangKey) ? mangNode : state.groupsConfig.find(g => g.key === it.parentKey);
      if (parentNode) parentNode.hasChildren = true;

      const ancestors = parentNode ? (parentNode.ancestors ? [...parentNode.ancestors, parentNode.key] : [parentNode.key]) : [mangKey];

      const node = {
        key: it.groupKey,
        isMangHeader: false,
        mangCode: mang.code,
        mangName: mang.name,
        itemIndex: it.index,
        origRow: it.origRow,
        subtotalStartRow: it.subtotalStartRow || null,
        subtotalEndRow: it.subtotalEndRow || null,
        tt: it.tt,
        origTT: it.origTT || it.tt,
        nd: it.nd,
        dvt: it.dvt || '',
        kl27: it.kl27,
        kl28: it.kl28,
        dg: it.dg,
        levelNum: it.levelNum,
        level: it.level,
        isGroup: it.isGroup,
        hasSubtotal: it.hasSubtotal,
        parentKey: it.parentKey,
        ancestors: ancestors,
        hasChildren: false,
        subtotalG: it.subtotalG || '-',
        subtotalH: it.subtotalH || '-'
      };

      const outlineDepth = Math.min(7, Math.max(1, ancestors.length));
      node.outlineLevel = outlineDepth;
      it.outlineLevel = outlineDepth;

      state.groupsConfig.push(node);
    }
  }

  recalculateSubtotalFormulas();
}

// Kiểm tra 2 nhóm có cùng cấp và liên tiếp nhau (sequential siblings) hay không
// Ví dụ: 1 & 2, I.1 & I.2, a & b, I & II, E1 & E2...
function areSequentialSiblings(tt1, tt2) {
  if (!tt1 || !tt2) return false;
  const s1 = String(tt1).trim().replace(/[\.\)\:\s]+$/, '').replace(/^[\(\[]/, '');
  const s2 = String(tt2).trim().replace(/[\.\)\:\s]+$/, '').replace(/^[\(\[]/, '');
  if (!s1 || !s2) return false;

  // 1. Số nguyên thuần túy: 1, 2, 3... hoặc 01, 02...
  if (/^\d+$/.test(s1) && /^\d+$/.test(s2)) return true;

  // 2. Số phân cấp cùng tiền tố: 1.1 & 1.2, I.1 & I.2, 2.1.1 & 2.1.2...
  const dot1 = s1.lastIndexOf('.');
  const dot2 = s2.lastIndexOf('.');
  if (dot1 > 0 && dot2 > 0) {
    const pfx1 = s1.substring(0, dot1);
    const pfx2 = s2.substring(0, dot2);
    const suf1 = s1.substring(dot1 + 1);
    const suf2 = s2.substring(dot2 + 1);
    if (pfx1 === pfx2 && /^\d+$/.test(suf1) && /^\d+$/.test(suf2)) return true;
  }

  // 3. Số La Mã thuần túy: I, II, III, IV, V...
  const isRoman = s => /^(X{0,3})(IX|IV|V?I{0,3})$/i.test(s) || /^[IVXLCDM]+$/i.test(s);
  if (isRoman(s1) && isRoman(s2)) return true;

  // 4. Mã ký tự + số: E1 & E2, F1 & F2...
  if (/^[A-Z]\d+$/i.test(s1) && /^[A-Z]\d+$/i.test(s2)) return true;

  // 5. Chữ cái đơn: a & b, A & B...
  if (/^[a-zA-Z]$/.test(s1) && /^[a-zA-Z]$/.test(s2)) return true;

  return false;
}

// Xác định số cấp bậc (1..6) dựa vào TT hoặc nội dung
function determineLevelNumFromTT(tt, nd, parentLevel = 1) {
  if (tt) {
    let cleanTT = String(tt).trim().replace(/[\.\)\:\s]+$/, '').replace(/^[\(\[]/, '');
    if (!cleanTT) return null;

    // Cấp 1: Ký tự chữ cái đơn A, B, C...
    if (/^[A-Z]$/i.test(cleanTT)) return 1;

    // Cấp 2: Số La Mã: I, II, III, IV, V, VI, VII, VIII, IX, X... hoặc A1, B1...
    if (/^(X{0,3})(IX|IV|V?I{0,3})$/i.test(cleanTT) && cleanTT.length > 0) return 2;
    if (/^[IVXLCDM]+$/i.test(cleanTT) || /^[A-Z]\d+$/i.test(cleanTT)) return 2;

    // Dạng I.1, II.1, III.1...
    if (/^[IVXLCDM]+\.\d+$/i.test(cleanTT)) return 3;

    // Cấp 3: Số nguyên đơn 1, 2, 3...
    if (/^\d+$/.test(cleanTT)) return 3;

    // Cấp 4: 1.1, 1.2, 2.1...
    if (/^\d+\.\d+$/.test(cleanTT)) return 4;

    // Cấp 5: 1.1.1, 1.1.2...
    if (/^\d+\.\d+\.\d+$/.test(cleanTT)) return 5;

    // Cấp 6: 1.1.1.1, 1.1.1.2...
    if (/^\d+(\.\d+){3,}$/.test(cleanTT)) return 6;
  }
  return null;
}

// Đồng bộ công thức Subtotal cho các nhóm hạng mục theo đúng cây phân cấp hiện tại
function recalculateSubtotalFormulas() {
  const items = state.extractedData;
  if (!items || items.length === 0) return;

  // Xử lý theo từng mảng độc lập
  for (const mang of MANG_CONFIG) {
    const itemsOfMang = items.filter(x => x.mangCode === mang.code);
    if (itemsOfMang.length === 0) continue;

    const mangKey = 'MANG_' + mang.code;
    const mangNode = state.groupsConfig.find(g => g.key === mangKey);
    if (mangNode) {
      mangNode.hasChildren = false;
      mangNode.ancestors = [];
    }

    // 1. Tái tạo stack phân cấp (groupStack) dựa trên levelNum hiện tại của từng item
    const groupStack = [mangNode || {
      key: mangKey,
      levelNum: 1,
      isMangHeader: true,
      hasChildren: false,
      ancestors: []
    }];

    for (let mIdx = 0; mIdx < itemsOfMang.length; mIdx++) {
      const it = itemsOfMang[mIdx];
      if (it.isMangHeader) continue;

      if (!it.groupKey) {
        it.groupKey = (it.isGroup ? 'GRP_' : 'ITEM_') + it.index;
      }

      if (it.isGroup) {
        // Pop các nhóm có levelNum >= it.levelNum (vì chúng không thể là cha của it)
        while (groupStack.length > 1 && groupStack[groupStack.length - 1].levelNum >= it.levelNum) {
          groupStack.pop();
        }

        const parent = groupStack[groupStack.length - 1];
        it.parentKey = parent ? parent.key : mangKey;
        if (parent) parent.hasChildren = true;

        const ancestors = parent ? (parent.ancestors ? [...parent.ancestors, parent.key] : [parent.key]) : [mangKey];
        it.ancestors = ancestors;

        groupStack.push({
          key: it.groupKey,
          levelNum: it.levelNum,
          ancestors: ancestors,
          hasChildren: false
        });
      } else {
        // Dòng chi tiết
        const parent = groupStack[groupStack.length - 1];
        it.parentKey = parent ? parent.key : mangKey;
        if (parent) parent.hasChildren = true;
        it.ancestors = parent ? (parent.ancestors ? [...parent.ancestors, parent.key] : [parent.key]) : [mangKey];
      }
    }

    // 2. Tính toán chính xác phạm vi Subtotal (startRow, endRow) cho từng nhóm trong mảng
    for (let mIdx = 0; mIdx < itemsOfMang.length; mIdx++) {
      const cur = itemsOfMang[mIdx];
      if (cur.isMangHeader) continue;

      if (!cur.isGroup) {
        cur.hasSubtotal = false;
        cur.subtotalStartRow = null;
        cur.subtotalEndRow = null;
        cur.childStartIdxInMang = null;
        cur.childEndIdxInMang = null;
        cur.subtotalG = '-';
        cur.subtotalH = '-';
        cur.tt27Formula = null;
        cur.tt28Formula = null;

        const cfgNode = state.groupsConfig.find(g => g.itemIndex === cur.index || g.key === cur.groupKey);
        if (cfgNode) {
          cfgNode.subtotalG = '-';
          cfgNode.subtotalH = '-';
          cfgNode.hasSubtotal = false;
          cfgNode.subtotalStartRow = null;
          cfgNode.subtotalEndRow = null;
          cfgNode.parentKey = cur.parentKey;
          cfgNode.ancestors = cur.ancestors;
          cfgNode.isGroup = false;
          cfgNode.levelNum = 99;
          cfgNode.level = 'CHI TIẾT';
        }
        continue;
      }

      // Quét tất cả các dòng con/cháu thuộc quyền quản lý của cur
      let firstChildRow = null;
      let lastChildRow = null;
      let firstChildIdx = null;
      let lastChildIdx = null;

      for (let j = mIdx + 1; j < itemsOfMang.length; j++) {
        const next = itemsOfMang[j];
        if (next.isMangHeader) break;
        // Gặp nhóm cùng cấp hoặc cấp cao hơn -> kết thúc phạm vi của nhóm cur
        if (next.isGroup && next.levelNum <= cur.levelNum) break;

        const rowNum = next.origRow !== undefined && next.origRow !== null ? next.origRow : (next.index + 1);
        if (firstChildRow === null) {
          firstChildRow = rowNum;
          firstChildIdx = j;
        }
        lastChildRow = rowNum;
        lastChildIdx = j;
      }

      if (firstChildRow !== null && lastChildRow !== null) {
        cur.hasSubtotal = true;
        cur.subtotalStartRow = firstChildRow;
        cur.subtotalEndRow = lastChildRow;
        cur.childStartIdxInMang = firstChildIdx;
        cur.childEndIdxInMang = lastChildIdx;

        cur.subtotalG = `=SUBTOTAL(9,G${firstChildRow}:G${lastChildRow})`;
        cur.subtotalH = `=SUBTOTAL(9,H${firstChildRow}:H${lastChildRow})`;
        cur.tt27Formula = cur.subtotalG;
        cur.tt28Formula = cur.subtotalH;

        // Tính lại giá trị số thành tiền cho cur từ các dòng chi tiết con
        let sum27 = 0, sum28 = 0;
        let hasDetail = false;
        for (let j = firstChildIdx; j <= lastChildIdx; j++) {
          const ch = itemsOfMang[j];
          if (!ch.isGroup && !ch.isMangHeader) {
            hasDetail = true;
            if (ch.tt27 !== null && ch.tt27 !== undefined && !isNaN(ch.tt27)) {
              sum27 += ch.tt27;
            } else if (ch.kl27 !== null && ch.dg !== null) {
              sum27 += ch.kl27 * ch.dg;
            }
            if (ch.tt28 !== null && ch.tt28 !== undefined && !isNaN(ch.tt28)) {
              sum28 += ch.tt28;
            } else if (ch.kl28 !== null && ch.dg !== null) {
              sum28 += ch.kl28 * ch.dg;
            }
          }
        }
        if (hasDetail) {
          cur.tt27 = sum27;
          cur.tt28 = sum28;
        }
      } else {
        cur.hasSubtotal = false;
        cur.subtotalStartRow = null;
        cur.subtotalEndRow = null;
        cur.childStartIdxInMang = null;
        cur.childEndIdxInMang = null;
        cur.subtotalG = '-';
        cur.subtotalH = '-';
        cur.tt27Formula = null;
        cur.tt28Formula = null;
      }

      const cfgNode = state.groupsConfig.find(g => g.itemIndex === cur.index || g.key === cur.groupKey);
      if (cfgNode) {
        cfgNode.subtotalG = cur.subtotalG;
        cfgNode.subtotalH = cur.subtotalH;
        cfgNode.levelNum = cur.levelNum;
        cfgNode.level = cur.level;
        cfgNode.isGroup = cur.isGroup;
        cfgNode.hasSubtotal = cur.hasSubtotal;
        cfgNode.subtotalStartRow = cur.subtotalStartRow;
        cfgNode.subtotalEndRow = cur.subtotalEndRow;
        cfgNode.parentKey = cur.parentKey;
        cfgNode.ancestors = cur.ancestors;
        cfgNode.hasChildren = (firstChildRow !== null);
        const depth = Math.min(7, Math.max(1, (cur.ancestors ? cur.ancestors.length : 1)));
        cfgNode.outlineLevel = depth;
        cur.outlineLevel = depth;
      }
    }
  }

  // Cập nhật lại thuộc tính hasChildren cho toàn bộ groupsConfig
  state.groupsConfig.forEach(g => {
    if (g.isGroup) {
      g.hasChildren = state.groupsConfig.some(other => other.parentKey === g.key);
    }
  });
}

// Chuyển đổi số nguyên sang số La Mã (I, II, III, IV...)
function toRoman(num) {
  const lookup = [
    ['M', 1000], ['CM', 900], ['D', 500], ['CD', 400],
    ['C', 100], ['XC', 90], ['L', 50], ['XL', 40],
    ['X', 10], ['IX', 9], ['V', 5], ['IV', 4], ['I', 1]
  ];
  let roman = '';
  for (const [r, n] of lookup) {
    while (num >= n) {
      roman += r;
      num -= n;
    }
  }
  return roman || String(num);
}

// Cập nhật lại toàn bộ Chỉ mục (TT) theo đúng quy tắc người dùng:
// Cấp 1: A, B, C, D...
// Cấp 2: I, II, ...
// Cấp 3: 1, 2, ...
// Cấp 4: 1.1, 1.2 ...
// Cấp 5: 1.1.1, 1.1.2 ...
// Cấp 6: 1.1.1.1, 1.1.1.2 ...
// Dòng chi tiết / vật tư: -
function reindexAllItems() {
  if (!state.extractedData || state.extractedData.length === 0) {
    showToast('Chưa có dữ liệu để đánh lại chỉ mục.', 'warning');
    return;
  }

  for (const mang of MANG_CONFIG) {
    const mangKey = 'MANG_' + mang.code;
    const itemsOfMang = state.extractedData.filter(x => x.mangCode === mang.code);
    if (itemsOfMang.length === 0) continue;

    // Tiêu đề Mảng (Cấp 1)
    const mangNode = state.groupsConfig.find(g => g.key === mangKey);
    if (mangNode) {
      mangNode.tt = mang.tt;
    }
    const cap1Item = itemsOfMang.find(x => x.isMangHeader);
    if (cap1Item) {
      cap1Item.tt = mang.tt;
    }

    // Đánh số đệ quy theo cây quan hệ cha-con
    function assignTTRecursively(parentKey, parentTT, parentLevel) {
      const directChildren = itemsOfMang.filter(x => x.parentKey === parentKey && !x.isMangHeader);
      let gIdx = 0;
      for (const ch of directChildren) {
        if (ch.isGroup) {
          let newTT = '';
          if (parentLevel === 1) {
            newTT = toRoman(gIdx + 1);
          } else if (parentLevel === 2) {
            newTT = String(gIdx + 1);
          } else {
            newTT = `${parentTT}.${gIdx + 1}`;
          }
          ch.tt = newTT;
          gIdx++;

          const gNode = state.groupsConfig.find(g => g.itemIndex === ch.index);
          if (gNode) gNode.tt = newTT;

          assignTTRecursively(ch.groupKey, newTT, ch.levelNum);
        } else {
          ch.tt = '-';
          const gNode = state.groupsConfig.find(g => g.itemIndex === ch.index);
          if (gNode) gNode.tt = '-';
        }
      }
    }

    assignTTRecursively(mangKey, mang.tt, 1);
  }

  const btnRestore = document.getElementById('btnRestoreOriginalTT');
  if (btnRestore) btnRestore.style.display = 'inline-flex';

  renderHierarchyTable();
  renderPreviewTable();
  if (state.extractedData.length > 0) prepareQuickDownload();
  showToast('Đã cập nhật lại Chỉ mục: Cấp 1 (A, B...), Cấp 2 (I, II...), Cấp 3 (1, 2...), Cấp 4 (1.1, 1.2...), Cấp 5 (1.1.1, 1.1.2...)!', 'success');
}

// Khôi phục lại Chỉ mục (TT) gốc từ file đầu vào
function restoreOriginalTT() {
  if (!state.extractedData || state.extractedData.length === 0) {
    showToast('Chưa có dữ liệu để khôi phục.', 'warning');
    return;
  }

  state.extractedData.forEach(item => {
    if (item.origTT !== undefined) {
      item.tt = item.origTT;
    }
  });

  state.groupsConfig.forEach(g => {
    if (g.isMangHeader) {
      const hInfo = state.mangHeaderInfo[g.mangCode] || {};
      const orig = (hInfo.origTT !== undefined && hInfo.origTT !== '') ? hInfo.origTT : (g.origTT || g.tt);
      g.tt = orig;
    } else if (g.itemIndex >= 0 && state.extractedData[g.itemIndex]) {
      g.tt = state.extractedData[g.itemIndex].origTT !== undefined ? state.extractedData[g.itemIndex].origTT : (g.origTT || g.tt);
    }
  });

  const btnRestore = document.getElementById('btnRestoreOriginalTT');
  if (btnRestore) btnRestore.style.display = 'none';

  renderHierarchyTable();
  renderPreviewTable();
  if (state.extractedData.length > 0) prepareQuickDownload();
  showToast('Đã khôi phục lại toàn bộ Chỉ mục (TT) gốc từ file đầu vào!', 'info');
}

function rerunValidationForAllRows() {
  state.validationIssues = [];
  state.extractedData.forEach(item => validateRow(item));
  updateValidationKPIs();
  renderValidationTable();
}

function isWalletKeywordRow(item) {
  return normalizeVietnameseSearchText(item?.nd).includes('vi dien tu');
}

function applyWalletSplit(mode) {
  if (!state.extractedData || state.extractedData.length === 0) {
    showToast('Chưa có dữ liệu để tách Ví điện tử.', 'warning');
    return;
  }

  const walletConfig = mode === 'mang'
    ? { label: 'Tách mảng riêng', maMang: 'VI', maDV: 'TTKD', maLoai: 'VTTB' }
    : { label: 'Tách DV riêng', maMang: 'CNTT', maDV: 'VI', maLoai: 'VTTB' };

  let walletActive = false;
  let matchedMarkerCount = 0;
  let updatedCount = 0;

  state.extractedData.forEach(item => {
    if (item.isMangHeader) {
      walletActive = false;
      return;
    }

    if (isWalletKeywordRow(item)) {
      walletActive = true;
      matchedMarkerCount++;
    }

    if (!walletActive) return;

    item.maMang = walletConfig.maMang;
    item.maDV = walletConfig.maDV;
    item.maLoai = walletConfig.maLoai;
    item.walletSplitMode = mode;
    item.skipMaDVValidation = true;
    updatedCount++;

    const node = state.groupsConfig.find(g => g.itemIndex === item.index);
    if (node) {
      node.walletSplitMode = mode;
    }
  });

  if (updatedCount === 0) {
    showToast('Không tìm thấy dòng có chữ "Ví điện tử" để tách.', 'warning');
    return;
  }

  rerunValidationForAllRows();
  renderHierarchyTable();
  renderPreviewTable();
  renderStrategyComparisonTable();
  prepareQuickDownload();

  const markerText = matchedMarkerCount > 1 ? `${matchedMarkerCount} điểm bắt đầu` : '1 điểm bắt đầu';
  showToast(`Đã ${walletConfig.label} cho ${formatNumber(updatedCount)} dòng Ví điện tử (${markerText}).`, 'success');
}

// Toggle mở rộng / thu gọn 1 nhóm
window.toggleGroupCollapse = function(key) {
  if (state.collapsedGroupKeys.has(key)) {
    state.collapsedGroupKeys.delete(key);
  } else {
    state.collapsedGroupKeys.add(key);
  }
  renderHierarchyTable();
};

// Render bảng Cấu hình Nhóm hạng mục & Phân cấp (Tab 3)
function renderHierarchyTable() {
  const tbody = document.getElementById('hierarchyTableBody');
  const keyword = (document.getElementById('searchGroupInput').value || '').trim().toLowerCase();

  if (state.groupsConfig.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="11" style="text-align: center; padding: 2rem; color: var(--text-muted);">
          Chưa có dữ liệu phân cấp. Vui lòng nạp file và bấm "⚡ Tổng hợp & Kiểm tra".
        </td>
      </tr>
    `;
    document.getElementById('hierarchyFooterText').textContent = `Hiển thị 0 mục`;
    return;
  }

  let visibleNodes = state.groupsConfig.filter(node => {
    if (!state.activeGroupLevels.has(node.levelNum)) return false;

    if (node.ancestors && node.ancestors.length > 0) {
      const isAncestorCollapsed = node.ancestors.some(ancestorKey => state.collapsedGroupKeys.has(ancestorKey));
      if (isAncestorCollapsed) return false;
    }

    if (keyword) {
      const matchText = (node.nd || '') + ' ' + (node.tt || '') + ' ' + (node.mangName || '') + ' ' + (node.dvt || '');
      if (!matchText.toLowerCase().includes(keyword)) return false;
    }

    return true;
  });

  if (visibleNodes.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="11" style="text-align: center; padding: 2rem; color: var(--text-muted);">
          Không có hạng mục nào phù hợp với bộ lọc cấp bậc hoặc từ khóa.
        </td>
      </tr>
    `;
    document.getElementById('hierarchyFooterText').textContent = `Hiển thị 0 / ${state.groupsConfig.length} mục`;
    return;
  }

  const rowsHtml = visibleNodes.map(node => {
    const isCollapsed = state.collapsedGroupKeys.has(node.key);

    let toggleBtnHtml = `<span class="tree-toggle-btn empty-leaf">•</span>`;
    if (node.hasChildren) {
      const icon = isCollapsed ? '▶' : '▼';
      const title = isCollapsed ? 'Bấm để mở rộng' : 'Bấm để thu gọn';
      toggleBtnHtml = `<span class="tree-toggle-btn" title="${title}" onclick="toggleGroupCollapse('${node.key}')">${icon}</span>`;
    }

    if (node.isMangHeader) {
      return `
        <tr class="row-cap-1 sec-${node.mangCode}" data-key="${node.key}" data-type="mang">
          <td style="text-align: center; font-weight: 800; color: #475569;">${node.origRow ? node.origRow : '-'}</td>
          <td style="font-weight: 800; font-size: 0.9rem;">${escapeHtml(node.tt)}</td>
          <td><span class="badge-lvl badge-lvl-1">CẤP 1</span></td>
          <td style="text-align: center;">
            <span class="badge badge-green">Mảng chính</span>
          </td>
          <td>
            <div class="tree-indent-cell" style="padding-left: 0px;">
              ${toggleBtnHtml}
              <strong class="tree-node-title" onclick="toggleGroupCollapse('${node.key}')" style="font-size: 0.925rem; letter-spacing: 0.02em;">
                ${isCollapsed ? '📁' : '📂'} ${escapeHtml(node.nd)}
              </strong>
            </div>
          </td>
          <td style="text-align: center; color: var(--text-muted);">-</td>
          <td class="num-cell" style="color: var(--text-muted);">-</td>
          <td class="num-cell" style="color: var(--text-muted);">-</td>
          <td class="num-cell" style="color: var(--text-muted);">-</td>
          <td class="formula-cell" style="font-size: 0.75rem;">${escapeHtml(node.subtotalG || '=SUBTOTAL(...)')}</td>
          <td class="formula-cell" style="font-size: 0.75rem;">${escapeHtml(node.subtotalH || '=SUBTOTAL(...)')}</td>
        </tr>
      `;
    }

    const idx = node.itemIndex;
    const item = state.extractedData[idx];
    const isGroup = !!(item && item.isGroup);
    const indentPx = isGroup ? Math.min(180, (node.levelNum - 1) * 24) : Math.min(180, ((node.ancestors ? node.ancestors.length : 1) * 20));
    const rowClass = isGroup ? `row-cap-${node.levelNum}` : 'row-leaf';

    let treeContentHtml = '';
    if (node.hasChildren) {
      const icon = isCollapsed ? '▶' : '▼';
      const title = isCollapsed ? 'Bấm để mở rộng' : 'Bấm để thu gọn';
      treeContentHtml = `
        <span class="tree-toggle-btn" title="${title}" onclick="toggleGroupCollapse('${node.key}')">${icon}</span>
        <span class="tree-node-title" onclick="toggleGroupCollapse('${node.key}')" style="font-weight: ${node.levelNum <= 2 ? '800' : '600'}; color: ${node.levelNum <= 2 ? '#0f172a' : '#334155'};">
          ${isCollapsed ? '📁' : '📂'} ${escapeHtml(node.nd)}
        </span>
      `;
    } else if (isGroup) {
      treeContentHtml = `
        <span class="tree-toggle-btn empty-leaf">•</span>
        <span style="font-weight: 600; color: #0f172a;">
          📂 ${escapeHtml(node.nd)}
        </span>
      `;
    } else {
      treeContentHtml = `
        <span class="tree-leaf-icon">📄</span>
        <span style="color: #a1a1aa; font-size: 0.8rem;">
          ${escapeHtml(node.nd)}
        </span>
      `;
    }

    let levelCellHtml = '';
    if (isGroup) {
      levelCellHtml = `
        <select class="select-input sel-level sel-level-${node.levelNum}" onchange="changeNodeLevel('${node.key}', ${idx}, this.value)">
          <option value="2" ${node.levelNum === 2 ? 'selected' : ''}>🔵 Cấp 2</option>
          <option value="3" ${node.levelNum === 3 ? 'selected' : ''}>🟢 Cấp 3</option>
          <option value="4" ${node.levelNum === 4 ? 'selected' : ''}>🟠 Cấp 4</option>
          <option value="5" ${node.levelNum === 5 ? 'selected' : ''}>🟣 Cấp 5</option>
          <option value="6" ${node.levelNum === 6 ? 'selected' : ''}>🌸 Cấp 6</option>
        </select>
      `;
    } else {
      levelCellHtml = `
        <span class="badge-lvl badge-lvl-leaf">Chi tiết</span>
      `;
    }

    return `
      <tr class="${rowClass}" data-key="${node.key}" data-index="${idx}">
        <td style="text-align: center; color: var(--text-muted); font-size: 0.775rem; font-weight: 600;">${node.origRow ? node.origRow : (item && item.origRow ? item.origRow : idx + 1)}</td>
        <td style="font-weight: 700; font-family: monospace; font-size: 0.8rem;">${escapeHtml(node.tt)}</td>
        <td>${levelCellHtml}</td>
        <td style="text-align: center;">
          <input type="checkbox" class="chk-is-group" ${isGroup ? 'checked' : ''} onchange="toggleItemIsGroup(${idx}, this.checked)" style="cursor: pointer; width: 16px; height: 16px;">
        </td>
        <td>
          <div class="tree-indent-cell" style="padding-left: ${indentPx}px;">
            ${treeContentHtml}
          </div>
        </td>
        <td style="text-align: center; font-size: 0.775rem;">${escapeHtml(item ? (item.dvt || '-') : '-')}</td>
        <td class="num-cell" style="font-size: 0.775rem;">${item && item.kl27 !== null && item.kl27 !== undefined ? formatNumber(item.kl27) : '-'}</td>
        <td class="num-cell" style="font-size: 0.775rem;">${item && item.kl28 !== null && item.kl28 !== undefined ? formatNumber(item.kl28) : '-'}</td>
        <td class="num-cell" style="font-size: 0.775rem;">${item && item.dg !== null && item.dg !== undefined ? formatNumber(item.dg) : '-'}</td>
        <td class="formula-cell" style="font-size: 0.75rem;">${node.subtotalG || '-'}</td>
        <td class="formula-cell" style="font-size: 0.75rem;">${node.subtotalH || '-'}</td>
      </tr>
    `;
  }).join('');

  tbody.innerHTML = rowsHtml;
  if (window.lucide) lucide.createIcons();
  document.getElementById('hierarchyFooterText').textContent = `Hiển thị ${visibleNodes.length} / ${state.groupsConfig.length} mục`;
}

// Thay đổi cấp bậc trực tiếp từ dropdown trong Tab 3 (tự động tịnh tiến liên hoàn các cấp con)
window.changeNodeLevel = function(key, idx, newLevel) {
  const lvl = parseInt(newLevel, 10);
  const targetNode = state.groupsConfig.find(g => g.key === key);
  if (!targetNode) return;

  const oldLevel = targetNode.levelNum;
  const delta = lvl - oldLevel;

  // Cập nhật nhóm cha
  targetNode.levelNum = lvl;
  targetNode.level = `CẤP ${lvl}`;
  if (state.extractedData[idx]) {
    state.extractedData[idx].levelNum = lvl;
    state.extractedData[idx].level = `CẤP ${lvl}`;
  }

  // Tự động tịnh tiến (cascade) toàn bộ các nhóm con/cháu trực thuộc khi cấp cha tăng hoặc giảm
  let cascadedCount = 0;
  if (delta !== 0) {
    state.groupsConfig.forEach(g => {
      if (g.key !== key && g.ancestors && g.ancestors.includes(key)) {
        if (g.isGroup) {
          const newChildLvl = Math.min(6, Math.max(2, g.levelNum + delta));
          g.levelNum = newChildLvl;
          g.level = `CẤP ${newChildLvl}`;
          if (g.itemIndex >= 0 && state.extractedData[g.itemIndex]) {
            state.extractedData[g.itemIndex].levelNum = newChildLvl;
            state.extractedData[g.itemIndex].level = `CẤP ${newChildLvl}`;
          }
          cascadedCount++;
        }
      }
    });
  }

  recalculateSubtotalFormulas();
  renderHierarchyTable();
  renderPreviewTable();
  if (state.extractedData.length > 0) prepareQuickDownload();

  if (cascadedCount > 0) {
    showToast(`Đã đổi Cấp ${lvl} cho nhóm và tự động cập nhật ${cascadedCount} nhóm con trực thuộc!`, 'info');
  }
};

// Chuyển đổi trạng thái "Là Nhóm hạng mục?" trực tiếp từ checkbox trong Tab 3
window.toggleItemIsGroup = function(idx, isChecked) {
  if (state.extractedData[idx]) {
    state.extractedData[idx].isGroup = isChecked;
    if (isChecked) {
      if (!state.extractedData[idx].levelNum || state.extractedData[idx].levelNum >= 90) {
        state.extractedData[idx].levelNum = 4;
        state.extractedData[idx].level = 'CẤP 4';
      }
    } else {
      state.extractedData[idx].levelNum = 99;
      state.extractedData[idx].level = 'CHI TIẾT';
    }
  }
  autoDetectGroupsHierarchy();
  renderHierarchyTable();
  renderPreviewTable();
  if (state.extractedData.length > 0) prepareQuickDownload();
  showToast(isChecked ? 'Đã chuyển thành Nhóm hạng mục!' : 'Đã chuyển thành Hạng mục chi tiết!', 'success');
};

// Áp dụng cấu hình từ Tab 3
function applyHierarchyConfiguration() {
  const rows = document.querySelectorAll('#hierarchyTableBody tr[data-index]');
  rows.forEach(tr => {
    const itemIdx = parseInt(tr.getAttribute('data-index'), 10);
    if (isNaN(itemIdx) || !state.extractedData[itemIdx]) return;

    const chkGroup = tr.querySelector('.chk-is-group');
    const selLevel = tr.querySelector('.sel-level');

    const item = state.extractedData[itemIdx];
    item.isGroup = chkGroup ? chkGroup.checked : false;
    const newLvlNum = selLevel ? parseInt(selLevel.value, 10) : 3;
    item.levelNum = item.isGroup ? newLvlNum : 99;
    item.level = item.isGroup ? `CẤP ${newLvlNum}` : 'CHI TIẾT';

    const node = state.groupsConfig.find(g => g.itemIndex === itemIdx);
    if (node) {
      node.isGroup = item.isGroup;
      node.levelNum = item.levelNum;
      node.level = item.level;
    }
  });

  recalculateSubtotalFormulas();
  renderPreviewTable();
  renderHierarchyTable();
  if (state.extractedData.length > 0) prepareQuickDownload();
}

// Chuẩn hóa và gán lại toàn bộ Group dòng (Outline Level 1..5) trong Excel theo đúng cây phân cấp
function fixOutlineGroups() {
  if (!state.extractedData || state.extractedData.length === 0) {
    showToast('Chưa có dữ liệu để Fix Group dòng!', 'warning');
    return;
  }

  // 1. Nhận dạng lại cây phân cấp mới nhất
  autoDetectGroupsHierarchy();

  // 2. Tính toán outlineLevel chuẩn xác cho từng dòng dựa trên số cấp tổ tiên (ancestors)
  let countFixed = 0;
  state.extractedData.forEach(it => {
    const node = state.groupsConfig.find(g => g.itemIndex === it.index);
    if (node && node.ancestors) {
      it.outlineLevel = Math.min(7, Math.max(1, node.ancestors.length));
      node.outlineLevel = it.outlineLevel;
    } else {
      it.outlineLevel = it.isGroup ? 1 : 2;
    }
    countFixed++;
  });

  state.fixedOutlineEnabled = true;
  renderHierarchyTable();
  renderPreviewTable();
  if (state.extractedData.length > 0) prepareQuickDownload();
  showToast(`Đã chuẩn hóa thành công Group dòng (Outline Level 1..5) cho ${countFixed} dòng theo cây phân cấp!`, 'success');
}

// Khởi tạo các nút lọc Cấp bậc (Tab 3: Hiển thị theo cấp)
function initGroupLevelFilterButtons() {
  const container = document.getElementById('groupLevelButtons');
  if (!container) return;

  const buttons = container.querySelectorAll('.btn-level-pill');
  buttons.forEach(btn => {
    btn.addEventListener('click', () => {
      const levelAttr = btn.getAttribute('data-level');

      if (levelAttr === 'ALL') {
        if (state.activeGroupLevels.size >= 6) {
          state.activeGroupLevels.clear();
          state.activeGroupLevels.add(1);
          state.activeGroupLevels.add(2);
        } else {
          state.activeGroupLevels = new Set([1, 2, 3, 4, 5, 6, 99]);
        }
      } else {
        const lvl = parseInt(levelAttr, 10);
        if (state.activeGroupLevels.has(lvl)) {
          state.activeGroupLevels.delete(lvl);
        } else {
          state.activeGroupLevels.add(lvl);
        }
      }

      const isAllActive = (state.activeGroupLevels.size === 7);
      buttons.forEach(b => {
        const l = b.getAttribute('data-level');
        if (l === 'ALL') {
          b.classList.toggle('active', isAllActive);
        } else {
          b.classList.toggle('active', state.activeGroupLevels.has(parseInt(l, 10)));
        }
      });

      renderHierarchyTable();
    });
  });
}

// ==================== HIỂN THỊ DỮ LIỆU & GIAO DIỆN ====================

// Cập nhật thẻ KPI
function updateValidationKPIs() {
  const total = state.validationIssues.length;
  const mathCount = state.validationIssues.filter(x => x.type === 'MATH').length;
  const missingCodeCount = state.validationIssues.filter(x => x.type === 'MISSING_CODE').length;
  const madvCount = state.validationIssues.filter(x => x.type === 'MADV').length;
  const infoCount = state.validationIssues.filter(x => x.type === 'INFO').length;
  const formulaCount = state.validationIssues.filter(x => x.type === 'FORMULA').length;
  const subtotalMismatchCount = state.validationIssues.filter(x => x.type === 'SUBTOTAL_MISMATCH').length;


  document.getElementById('badgeValidationCount').textContent = total;
  document.getElementById('badgeRowCount').textContent = `${formatNumber(state.extractedData.length)} dòng`;

  const groupCount = state.extractedData.filter(x => x.isGroup).length;
  document.getElementById('badgeGroupCount').textContent = `${groupCount} nhóm`;

  const updateBadge = (id, count) => {
    const el = document.getElementById(id);
    if (!el) return;
    el.textContent = formatNumber(count);
    el.classList.toggle('badge-has-error', count > 0);
    el.classList.toggle('badge-zero', count === 0);
    const chip = el.closest('.chip');
    if (chip) {
      chip.classList.toggle('has-error', count > 0);
      chip.classList.toggle('has-zero', count === 0);
    }
  };

  updateBadge('countAllIssues', total);
  updateBadge('countMathIssues', mathCount);
  updateBadge('countCodeIssues', missingCodeCount);
  updateBadge('countMaDVIssues', madvCount);
  updateBadge('countInfoIssues', infoCount);
  updateBadge('countFormulaIssues', formulaCount);
  updateBadge('countSubtotalMismatch', subtotalMismatchCount);

  const madvAlert = document.getElementById('missingMaDVAlert');
  if (madvCount > 0) {
    const missingCodes = Array.from(new Set(state.validationIssues.filter(x => x.type === 'MADV').map(x => x.currentVal)));
    document.getElementById('missingMaDVText').innerHTML = `
      Phát hiện <strong>${missingCodes.length} Mã DV</strong> chưa có trong sheet <strong>"TH theo DV"</strong>:
      <span style="font-weight: 700; color: #b45309;">${escapeHtml(missingCodes.join(', '))}</span>.
      Vui lòng bổ sung vào cột C sheet "TH theo DV" để bảng tổng hợp theo DV không bị thiếu dữ liệu!
    `;
    madvAlert.style.display = 'flex';
  } else {
    madvAlert.style.display = 'none';
  }
}

// Helper tạo Badge Loại lỗi cho bảng Kiểm tra & So khớp
function getValidationIssueBadge(issue) {
  const type = issue.type || '';
  switch (type) {
    case 'MATH':
      return `<span class="badge-issue badge-issue-math">Lỗi Thành tiền</span>`;
    case 'MISSING_CODE':
      return `<span class="badge-issue badge-issue-code">Thiếu Mã loại/mảng/DV</span>`;
    case 'SUBTOTAL_MISMATCH':
      return `<span class="badge-issue badge-issue-subtotal">Lệch Subtotal G/H</span>`;
    case 'MADV':
      return `<span class="badge-issue badge-issue-madv">Mã DV ngoài TH DV</span>`;
    case 'INFO':
      return `<span class="badge-issue badge-issue-info">Thiếu Đơn giá / ĐVT</span>`;
    case 'FORMULA':
      return `<span class="badge-issue badge-issue-formula">Lỗi công thức Excel</span>`;
    default:
      return `<span class="badge-issue badge-issue-default">${escapeHtml(type || 'Cảnh báo')}</span>`;
  }
}

// Render bảng kiểm tra lỗi (Tab 1)
function renderValidationTable() {
  const tbody = document.getElementById('validationTableBody');
  const selectedType = document.querySelector('input[name="filterIssueType"]:checked')?.value || 'ALL';
  const keyword = (document.getElementById('searchValidationInput').value || '').trim().toLowerCase();

  let filtered = state.validationIssues;
  if (selectedType !== 'ALL') {
    filtered = filtered.filter(x => x.type === selectedType);
  }
  if (keyword) {
    filtered = filtered.filter(x =>
      (x.nd || '').toLowerCase().includes(keyword) ||
      (x.fileName || '').toLowerCase().includes(keyword) ||
      (x.mangName || '').toLowerCase().includes(keyword) ||
      (x.message || '').toLowerCase().includes(keyword) ||
      String(x.origRow).includes(keyword)
    );
  }

  if (filtered.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="10" style="text-align: center; padding: 2rem; color: #059669; font-weight: 600;">
          🎉 Không có cảnh báo hoặc lỗi nào phù hợp với bộ lọc!
        </td>
      </tr>
    `;
    document.getElementById('validationFooterText').textContent = `Hiển thị 0 / ${state.validationIssues.length} cảnh báo`;
    return;
  }

  const rowsHtml = filtered.map((issue, i) => {
    return `
      <tr>
        <td style="text-align: center;">${i + 1}</td>
        <td style="text-align: center;"><span class="badge badge-mang-${issue.mangCode}">${issue.mangCode}</span></td>
        <td style="text-align: center; font-weight: 700; color: #0284c7;">Dòng ${issue.origRow}</td>
        <td style="font-weight: 500; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="${escapeHtml(issue.nd)}">${escapeHtml(issue.nd)}</td>
        <td class="num-cell" style="font-family: monospace; font-weight: 600;">${issue.kl27 !== null && issue.kl27 !== undefined ? formatNumber(issue.kl27) : '-'}</td>
        <td class="num-cell" style="font-family: monospace; font-weight: 600;">${issue.kl28 !== null && issue.kl28 !== undefined ? formatNumber(issue.kl28) : '-'}</td>
        <td class="num-cell" style="font-family: monospace; color: #0f172a; font-weight: 700;">${issue.dg !== null && issue.dg !== undefined && !isNaN(issue.dg) ? formatNumber(issue.dg) : '-'}</td>
        <td style="text-align: center; white-space: nowrap;">${getValidationIssueBadge(issue)}</td>
        <td style="white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="${escapeHtml(issue.message)}">
          ${escapeHtml(issue.message)}
        </td>
        <td style="text-align: center;">
          <button type="button" class="btn btn-outline btn-xs" style="padding: 2px 7px; font-size: 0.725rem; font-weight: 700; color: #4338ca; border-color: #c7d2fe;" onclick="openEditItemFromValidation(${issue.origRow}, '${issue.mangCode}')" title="Sửa dòng lỗi này">
            <i data-lucide="edit-3" class="w-3 h-3 inline-block mr-0.5"></i> Sửa
          </button>
        </td>
      </tr>
    `;
  }).join('');

  tbody.innerHTML = rowsHtml;
  if (window.lucide) lucide.createIcons();

  const footerText = document.getElementById('validationFooterText');
  if (footerText) footerText.textContent = `Hiển thị ${filtered.length} / ${state.validationIssues.length} cảnh báo`;

  const visibleCountEl = document.getElementById('valVisibleCount');
  if (visibleCountEl) visibleCountEl.textContent = formatNumber(filtered.length);
  const totalCountEl = document.getElementById('valTotalCount');
  if (totalCountEl) totalCountEl.textContent = formatNumber(state.validationIssues.length);
}

// Helper lấy giá trị số Thành tiền 2027 của 1 dòng
function getItemVal27(item) {
  if (item.isMangHeader) {
    let s = 0, found = false;
    for (const it of state.extractedData) {
      if (!it.isGroup && it.mangCode === item.mangCode) {
        const v = getItemVal27(it);
        if (v !== null && v !== undefined && !isNaN(v)) { s += v; found = true; }
      }
    }
    return found ? s : 0;
  }
  if (item.isGroup) {
    if (item.subtotalStartRow && item.subtotalEndRow) {
      let s = 0, found = false;
      for (const it of state.extractedData) {
        if (!it.isGroup && it.mangCode === item.mangCode && it.origRow >= item.subtotalStartRow && it.origRow <= item.subtotalEndRow) {
          const v = getItemVal27(it);
          if (v !== null && v !== undefined && !isNaN(v)) { s += v; found = true; }
        }
      }
      if (found) return s;
    }
    if (item.groupKey) {
      let s = 0, found = false;
      for (const it of state.extractedData) {
        if (!it.isGroup && it.parentKey === item.groupKey) {
          const v = getItemVal27(it);
          if (v !== null && v !== undefined && !isNaN(v)) { s += v; found = true; }
        }
      }
      if (found) return s;
    }
    if (item.tt27 !== null && item.tt27 !== undefined && !isNaN(item.tt27)) {
      return item.tt27;
    }
    return null;
  }
  if (item.tt27 !== null && item.tt27 !== undefined && !isNaN(item.tt27)) {
    return item.tt27;
  }
  if (!item.isGroup && item.kl27 !== null && item.dg !== null) {
    return item.kl27 * item.dg;
  }
  return null;
}

// Helper lấy giá trị số Thành tiền 2028 của 1 dòng
function getItemVal28(item) {
  if (item.isMangHeader) {
    let s = 0, found = false;
    for (const it of state.extractedData) {
      if (!it.isGroup && it.mangCode === item.mangCode) {
        const v = getItemVal28(it);
        if (v !== null && v !== undefined && !isNaN(v)) { s += v; found = true; }
      }
    }
    return found ? s : 0;
  }
  if (item.isGroup) {
    if (item.subtotalStartRow && item.subtotalEndRow) {
      let s = 0, found = false;
      for (const it of state.extractedData) {
        if (!it.isGroup && it.mangCode === item.mangCode && it.origRow >= item.subtotalStartRow && it.origRow <= item.subtotalEndRow) {
          const v = getItemVal28(it);
          if (v !== null && v !== undefined && !isNaN(v)) { s += v; found = true; }
        }
      }
      if (found) return s;
    }
    if (item.groupKey) {
      let s = 0, found = false;
      for (const it of state.extractedData) {
        if (!it.isGroup && it.parentKey === item.groupKey) {
          const v = getItemVal28(it);
          if (v !== null && v !== undefined && !isNaN(v)) { s += v; found = true; }
        }
      }
      if (found) return s;
    }
    if (item.tt28 !== null && item.tt28 !== undefined && !isNaN(item.tt28)) {
      return item.tt28;
    }
    return null;
  }
  if (item.tt28 !== null && item.tt28 !== undefined && !isNaN(item.tt28)) {
    return item.tt28;
  }
  if (!item.isGroup && item.kl28 !== null && item.dg !== null) {
    return item.kl28 * item.dg;
  }
  return null;
}

// Kiểm tra xem item có con/hạng mục phụ thuộc nào cũng xuất hiện trong danh sách list không
function hasDescendantInList(item, list) {
  if (!item.isGroup) return false;
  return list.some(other => {
    if (other === item) return false;
    if (other.mangCode !== item.mangCode) return false;
    if (item.isMangHeader) {
      return true;
    }
    if (item.subtotalStartRow && item.subtotalEndRow) {
      if (other.origRow >= item.subtotalStartRow && other.origRow <= item.subtotalEndRow) {
        return true;
      }
    }
    if (other.parentKey && item.groupKey && other.parentKey === item.groupKey) {
      return true;
    }
    return false;
  });
}

const PREVIEW_COLUMN_DEFS = [
  { key: 'tt', label: 'TT', className: 'preview-col-tt', headerStyle: 'width: 75px; text-align: center;' },
  { key: 'nd', label: 'Nội dung / Hạng mục đầu tư, mua sắm (Cấp 2)', className: 'preview-col-nd' },
  { key: 'dvt', label: 'ĐVT', className: 'preview-col-dvt optional', headerStyle: 'width: 90px; text-align: center;', optional: true },
  { key: 'kl27', label: 'KL 2027', className: 'num-cell preview-col-kl optional', headerStyle: 'width: 120px;', optional: true },
  { key: 'kl28', label: 'KL 2028', className: 'num-cell preview-col-kl optional', headerStyle: 'width: 120px;', optional: true },
  { key: 'dg', label: 'Đơn giá', className: 'num-cell preview-col-dg optional', headerStyle: 'width: 140px;', optional: true },
  { key: 'tt27', label: 'Năm 2027', className: 'num-cell preview-col-money', headerStyle: 'width: 180px;' },
  { key: 'tt28', label: 'Năm 2028', className: 'num-cell preview-col-money', headerStyle: 'width: 180px;' },
  { key: 'total', label: 'Tổng', className: 'num-cell preview-col-money', headerStyle: 'width: 180px;' },
  { key: 'action', label: 'Sửa', className: 'preview-col-action', headerStyle: 'width: 60px; text-align: center;' }
];

function getVisiblePreviewColumns() {
  return PREVIEW_COLUMN_DEFS.filter(col => !col.optional || state.previewColumns[col.key]);
}

function renderPreviewTableHeader() {
  const thead = document.getElementById('previewTableHead');
  if (!thead) return;
  const headers = getVisiblePreviewColumns().map(col => {
    const style = col.headerStyle ? ` style="${col.headerStyle}"` : '';
    return `<th class="${col.className || ''}"${style}>${col.label}</th>`;
  }).join('');
  thead.innerHTML = `<tr>${headers}</tr>`;
}

function getPreviewVisibleColumnCount() {
  return getVisiblePreviewColumns().length;
}

function formatPreviewOptionalNumber(value) {
  return value !== null && value !== undefined && !isNaN(value) ? formatNumber(value) : '-';
}

// Render bảng xem trước Masterlist (Tab 2)
function renderPreviewTable() {
  const tbody = document.getElementById('previewTableBody');
  renderPreviewTableHeader();
  const visibleColumnCount = getPreviewVisibleColumnCount();
  if (state.extractedData.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="${visibleColumnCount}" style="text-align: center; padding: 2rem; color: var(--text-muted);">
          Chưa có dữ liệu. Vui lòng nạp file để xem trước kết quả.
        </td>
      </tr>
    `;
    document.getElementById('previewFooterText').textContent = `Hiển thị 0 / 0 dòng`;
    document.getElementById('sumTT2027').textContent = '0';
    document.getElementById('sumTT2028').textContent = '0';
    const sumTotEl = document.getElementById('sumTTTotal');
    if (sumTotEl) sumTotEl.textContent = '0';
    return;
  }

  const allowL1 = document.getElementById('chkLevel1')?.checked ?? true;
  const allowL2 = document.getElementById('chkLevel2')?.checked ?? true;
  const allowL3 = document.getElementById('chkLevel3')?.checked ?? true;
  const allowL4 = document.getElementById('chkLevel4')?.checked ?? true;
  const allowLeaf = document.getElementById('chkLevelLeaf')?.checked ?? true;

  const selectedMang = document.getElementById('selectMangFilter').value;
  const keyword = (document.getElementById('searchPreviewInput').value || '').trim().toLowerCase();

  let filtered = state.extractedData.filter(item => {
    if (selectedMang !== 'ALL' && item.mangCode !== selectedMang) return false;

    // Xác định xem dòng này là Chi tiết (Vật tư) hay là Nhóm
    const isLeaf = !item.isGroup || item.levelNum >= 90 || item.level === 'CHI TIẾT' || item.level === 'LEAF';

    if (isLeaf) {
      if (!allowLeaf) return false;
    } else {
      const lvl = item.levelNum;
      if (lvl === 1 || item.level === 'CẤP 1' || item.isMangHeader) {
        if (!allowL1) return false;
      } else if (lvl === 2 || item.level === 'CẤP 2') {
        if (!allowL2) return false;
      } else if (lvl === 3 || item.level === 'CẤP 3') {
        if (!allowL3) return false;
      } else if (lvl >= 4 || item.level === 'CẤP 4' || item.level === 'CẤP 5' || item.level === 'CẤP 6') {
        if (!allowL4) return false;
      } else {
        if (!allowL4) return false;
      }
    }

    if (keyword) {
      const matchText = (item.nd || '') + ' ' + (item.tt || '') + ' ' + (item.maDV || '') + ' ' + (item.maMang || '') + ' ' + (item.maLoai || '');
      if (!matchText.toLowerCase().includes(keyword)) return false;
    }

    return true;
  });

  if (filtered.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="${visibleColumnCount}" style="text-align: center; padding: 2rem; color: var(--text-muted);">
          Không có dữ liệu nào phù hợp với bộ lọc.
        </td>
      </tr>
    `;
    document.getElementById('previewFooterText').textContent = `Hiển thị 0 / ${state.extractedData.length} dòng`;
    document.getElementById('sumTT2027').textContent = '0';
    document.getElementById('sumTT2028').textContent = '0';
    const sumTotEl = document.getElementById('sumTTTotal');
    if (sumTotEl) sumTotEl.textContent = '0';
    return;
  }

  // Tính tổng Năm 2027 và Năm 2028 cho dòng đầu tiên (dòng Tổng)
  // Chỉ cộng các dòng không có con xuất hiện trong filtered để không bị trùng lặp giữa cha và con
  let sum27 = 0;
  let sum28 = 0;

  filtered.forEach(item => {
    if (!hasDescendantInList(item, filtered)) {
      const v27 = getItemVal27(item);
      const v28 = getItemVal28(item);
      if (v27 !== null && v27 !== undefined && !isNaN(v27)) sum27 += v27;
      if (v28 !== null && v28 !== undefined && !isNaN(v28)) sum28 += v28;
    }
  });

  // Dòng đầu tiên: Dòng Tổng với cell của cột Năm 2027, Năm 2028 và Tổng của các cell ở dưới
  const topTotalRowHtml = `
    <tr class="row-total-top">
      <td style="font-weight: 800; text-align: center;">-</td>
      <td class="cell-nd" style="font-weight: 800; color: #065f46;">
        <strong>Tổng</strong>
      </td>
      ${state.previewColumns.dvt ? '<td class="preview-col-dvt" style="text-align: center; color: #94a3b8;">-</td>' : ''}
      ${state.previewColumns.kl27 ? '<td class="num-cell preview-col-kl" style="color: #94a3b8;">-</td>' : ''}
      ${state.previewColumns.kl28 ? '<td class="num-cell preview-col-kl" style="color: #94a3b8;">-</td>' : ''}
      ${state.previewColumns.dg ? '<td class="num-cell preview-col-dg" style="color: #94a3b8;">-</td>' : ''}
      <td class="num-cell" style="font-weight: 800; color: #4f46e5;">${formatNumber(sum27)}</td>
      <td class="num-cell" style="font-weight: 800; color: #059669;">${formatNumber(sum28)}</td>
      <td class="num-cell" style="font-weight: 800; color: #047857;">${formatNumber(sum27 + sum28)}</td>
      <td style="text-align: center; color: #94a3b8;">-</td>
    </tr>
  `;

  const rowsHtml = filtered.map((item, idx) => {
    const isCap1 = item.levelNum === 1 || item.isMangHeader;
    const isGrp = item.isGroup;
    let rowClass = '';
    if (isCap1) rowClass = 'row-cap-1 sec-' + item.mangCode;
    else if (item.levelNum === 2) rowClass = 'row-cap-2';
    else if (item.levelNum === 3) rowClass = 'row-cap-3';
    else if (item.levelNum >= 4 && item.levelNum <= 6) rowClass = 'row-cap-4';

    let indentPx = 0;
    if (isCap1) indentPx = 0;
    else if (item.levelNum === 2) indentPx = 14;
    else if (item.levelNum === 3) indentPx = 28;
    else if (item.levelNum === 4) indentPx = 42;
    else if (item.levelNum === 5) indentPx = 56;
    else if (item.levelNum === 6) indentPx = 70;
    if (!isGrp) indentPx = 84;

    const val27 = getItemVal27(item);
    const val28 = getItemVal28(item);
    const has27 = (val27 !== null && val27 !== undefined);
    const has28 = (val28 !== null && val28 !== undefined);
    const rowTot = (val27 || 0) + (val28 || 0);

    const textTT27 = has27 ? formatNumber(val27) : '-';
    const textTT28 = has28 ? formatNumber(val28) : '-';
    const textTotal = (has27 || has28) ? formatNumber(rowTot) : '-';

    let displayTT27 = textTT27;
    let displayTT28 = textTT28;
    let displayTotal = textTotal;

    if (isGrp) {
      displayTT27 = `<strong style="color: ${isCap1 ? '#4338ca' : '#4f46e5'}; font-weight: ${isCap1 ? '800' : '700'};" title="${escapeHtml(item.subtotalG || '')}">${textTT27}</strong>`;
      displayTT28 = `<strong style="color: ${isCap1 ? '#047857' : '#059669'}; font-weight: ${isCap1 ? '800' : '700'};" title="${escapeHtml(item.subtotalH || '')}">${textTT28}</strong>`;
      displayTotal = `<strong style="color: #047857; font-weight: ${isCap1 ? '800' : '700'};">${textTotal}</strong>`;
    } else {
      if (has27 || has28) {
        displayTotal = `<strong style="color: #059669;">${textTotal}</strong>`;
      }
    }

    return `
      <tr class="${rowClass}">
        <td style="font-weight: ${isCap1 ? '800' : '600'}; font-family: monospace; text-align: center;">${escapeHtml(item.tt || '')}</td>
        <td class="cell-nd ${!isCap1 ? 'cell-editable' : ''}" data-item-idx="${item.index}" data-field="nd" title="${!isCap1 ? 'Bấm đúp để sửa nhanh tên hạng mục' : escapeHtml(item.nd || '')}">
          <span style="display: inline-block; width: ${indentPx}px;"></span>
          ${isCap1 ? `<strong style="font-size: 0.9rem; letter-spacing: 0.01em;">${escapeHtml(item.nd || '')}</strong>` : (isGrp ? `<strong>${escapeHtml(item.nd || '')}</strong>` : escapeHtml(item.nd || ''))}
        </td>
        ${state.previewColumns.dvt ? `<td class="preview-col-dvt ${!isCap1 ? 'cell-editable' : ''}" data-item-idx="${item.index}" data-field="dvt" title="${!isCap1 ? 'Bấm đúp để sửa nhanh ĐVT' : ''}" style="text-align: center; color: ${item.dvt ? '#334155' : '#94a3b8'};">${escapeHtml(item.dvt || '-')}</td>` : ''}
        ${state.previewColumns.kl27 ? `<td class="num-cell preview-col-kl ${!isCap1 ? 'cell-editable' : ''}" data-item-idx="${item.index}" data-field="kl27" title="${!isCap1 ? 'Bấm đúp để sửa nhanh KL 2027' : ''}">${formatPreviewOptionalNumber(item.kl27)}</td>` : ''}
        ${state.previewColumns.kl28 ? `<td class="num-cell preview-col-kl ${!isCap1 ? 'cell-editable' : ''}" data-item-idx="${item.index}" data-field="kl28" title="${!isCap1 ? 'Bấm đúp để sửa nhanh KL 2028' : ''}">${formatPreviewOptionalNumber(item.kl28)}</td>` : ''}
        ${state.previewColumns.dg ? `<td class="num-cell preview-col-dg ${!isCap1 ? 'cell-editable' : ''}" data-item-idx="${item.index}" data-field="dg" title="${!isCap1 ? 'Bấm đúp để sửa nhanh Đơn giá' : ''}">${formatPreviewOptionalNumber(item.dg)}</td>` : ''}
        <td class="num-cell">${displayTT27}</td>
        <td class="num-cell">${displayTT28}</td>
        <td class="num-cell">${displayTotal}</td>
        <td style="text-align: center;">
          ${!isCap1 ? `<button type="button" class="btn-row-edit" onclick="openEditItemModal(${item.index})" title="Chỉnh sửa chi tiết dòng này"><i data-lucide="edit-3" class="w-3.5 h-3.5"></i></button>` : ''}
        </td>
      </tr>
    `;
  }).join('');

  tbody.innerHTML = topTotalRowHtml + rowsHtml;
  if (window.lucide) lucide.createIcons();
  document.getElementById('previewFooterText').textContent = `Hiển thị ${filtered.length} / ${state.extractedData.length} dòng`;
  document.getElementById('sumTT2027').textContent = formatNumber(sum27);
  document.getElementById('sumTT2028').textContent = formatNumber(sum28);
  const sumTotEl = document.getElementById('sumTTTotal');
  if (sumTotEl) sumTotEl.textContent = formatNumber(sum27 + sum28);
}

// Sao chép toàn bộ bảng Preview vào Clipboard để dán vào Excel dạng bảng
async function copyPreviewTableToClipboard() {
  if (!state.extractedData || state.extractedData.length === 0) {
    showToast('Chưa có dữ liệu để sao chép!', 'warning');
    return;
  }

  const table = document.getElementById('previewTable');
  if (!table) return;

  const trList = table.querySelectorAll('tbody tr');
  if (trList.length === 0) {
    showToast('Không có dữ liệu trong bảng để sao chép!', 'warning');
    return;
  }

  // 1. Tạo dữ liệu dạng Tab-Separated (TSV) cho text/plain
  const tsvLines = [];
  tsvLines.push(['TT', 'Nội dung / Hạng mục đầu tư, mua sắm (Cấp 2)', 'Năm 2027', 'Năm 2028', 'Tổng'].join('\t'));

  trList.forEach(tr => {
    const tds = tr.querySelectorAll('td');
    if (tds.length === 5) {
      const line = [
        tds[0].innerText.replace(/\r?\n|\r/g, ' ').trim(),
        tds[1].innerText.replace(/\r?\n|\r/g, ' ').trim(),
        tds[2].innerText.replace(/\r?\n|\r/g, ' ').trim(),
        tds[3].innerText.replace(/\r?\n|\r/g, ' ').trim(),
        tds[4].innerText.replace(/\r?\n|\r/g, ' ').trim()
      ].join('\t');
      tsvLines.push(line);
    }
  });
  const tsvText = tsvLines.join('\r\n');

  // 2. Tạo HTML Table chuẩn Excel với kiểu định dạng cell rõ ràng
  let htmlTable = '<table border="1" style="border-collapse: collapse; font-family: Calibri, Arial, sans-serif; font-size: 11pt;">\r\n';
  htmlTable += '  <thead>\r\n';
  htmlTable += '    <tr style="background-color: #f1f5f9; font-weight: bold;">\r\n';
  htmlTable += '      <th style="border: 1px solid #94a3b8; padding: 6px; text-align: center;">TT</th>\r\n';
  htmlTable += '      <th style="border: 1px solid #94a3b8; padding: 6px; text-align: left;">Nội dung / Hạng mục đầu tư, mua sắm (Cấp 2)</th>\r\n';
  htmlTable += '      <th style="border: 1px solid #94a3b8; padding: 6px; text-align: right;">Năm 2027</th>\r\n';
  htmlTable += '      <th style="border: 1px solid #94a3b8; padding: 6px; text-align: right;">Năm 2028</th>\r\n';
  htmlTable += '      <th style="border: 1px solid #94a3b8; padding: 6px; text-align: right;">Tổng</th>\r\n';
  htmlTable += '    </tr>\r\n';
  htmlTable += '  </thead>\r\n';
  htmlTable += '  <tbody>\r\n';

  trList.forEach(tr => {
    const tds = tr.querySelectorAll('td');
    if (tds.length === 5) {
      const isTopTotal = tr.classList.contains('row-total-top');
      const isCap1 = tr.classList.contains('row-cap-1');
      const isGrp = tr.className && tr.className.includes('row-cap-');
      let trStyle = '';
      if (isTopTotal) {
        trStyle = 'background-color: #ecfdf5; font-weight: bold; color: #047857;';
      } else if (isCap1) {
        trStyle = 'background-color: #eef2ff; font-weight: bold; color: #3730a3;';
      } else if (isGrp) {
        trStyle = 'background-color: #f8fafc; font-weight: bold;';
      }

      htmlTable += `    <tr style="${trStyle}">\r\n`;
      htmlTable += `      <td style="border: 1px solid #cbd5e1; padding: 4px 8px; text-align: center;">${escapeHtml(tds[0].innerText.trim())}</td>\r\n`;
      htmlTable += `      <td style="border: 1px solid #cbd5e1; padding: 4px 8px; text-align: left;">${escapeHtml(tds[1].innerText.trim())}</td>\r\n`;
      htmlTable += `      <td style="border: 1px solid #cbd5e1; padding: 4px 8px; text-align: right;">${escapeHtml(tds[2].innerText.trim())}</td>\r\n`;
      htmlTable += `      <td style="border: 1px solid #cbd5e1; padding: 4px 8px; text-align: right;">${escapeHtml(tds[3].innerText.trim())}</td>\r\n`;
      htmlTable += `      <td style="border: 1px solid #cbd5e1; padding: 4px 8px; text-align: right;">${escapeHtml(tds[4].innerText.trim())}</td>\r\n`;
      htmlTable += '    </tr>\r\n';
    }
  });
  htmlTable += '  </tbody>\r\n</table>';

  function triggerCopySuccess() {
    ['copyBtnText', 'copyBtnIcon'].forEach(id => {
      const el = document.getElementById(id);
      if (el && id === 'copyBtnText') {
        const old = el.textContent;
        el.textContent = 'Đã sao chép!';
        setTimeout(() => { el.textContent = old; }, 2000);
      } else if (el && id === 'copyBtnIcon') {
        el.innerHTML = '<i data-lucide="check" class="w-4 h-4 text-emerald-600"></i>';
        if (window.lucide) lucide.createIcons();
        setTimeout(() => {
          el.innerHTML = '<i data-lucide="copy" class="w-4 h-4"></i>';
          if (window.lucide) lucide.createIcons();
        }, 2000);
      }
    });
    showToast('Đã sao chép bảng vào Clipboard! Bạn có thể dán (Ctrl+V) ngay vào Excel dạng bảng.', 'success');
  }

  try {
    if (navigator.clipboard && window.ClipboardItem) {
      const blobHtml = new Blob([htmlTable], { type: 'text/html' });
      const blobText = new Blob([tsvText], { type: 'text/plain' });
      await navigator.clipboard.write([
        new ClipboardItem({
          'text/html': blobHtml,
          'text/plain': blobText
        })
      ]);
      triggerCopySuccess();
    } else if (navigator.clipboard && navigator.clipboard.writeText) {
      await navigator.clipboard.writeText(tsvText);
      triggerCopySuccess();
    } else {
      fallbackCopy(tsvText);
      triggerCopySuccess();
    }
  } catch (err) {
    console.warn('ClipboardItem write failed, fallback to text/plain:', err);
    try {
      if (navigator.clipboard && navigator.clipboard.writeText) {
        await navigator.clipboard.writeText(tsvText);
      } else {
        fallbackCopy(tsvText);
      }
      triggerCopySuccess();
    } catch (e2) {
      fallbackCopy(tsvText);
      triggerCopySuccess();
    }
  }
}

function fallbackCopy(text) {
  const ta = document.createElement('textarea');
  ta.value = text;
  ta.style.position = 'fixed';
  ta.style.left = '-9999px';
  ta.style.top = '-9999px';
  document.body.appendChild(ta);
  ta.focus();
  ta.select();
  document.execCommand('copy');
  document.body.removeChild(ta);
}

// Sao chép toàn bộ bảng Kiểm tra & So khớp vào Clipboard để dán vào Excel dạng bảng
async function copyValidationTableToClipboard() {
  const tbody = document.getElementById('validationTableBody');
  const trList = tbody ? tbody.querySelectorAll('tr') : [];

  if (trList.length === 0 || (trList.length === 1 && trList[0].querySelector('td[colspan]'))) {
    showToast('Không có dữ liệu cảnh báo/lỗi để sao chép!', 'warning');
    return;
  }

  // 1. Tạo dữ liệu TSV (Tab-Separated Values) cho text/plain
  const headers = ['STT', 'Mảng', 'Vị trí', 'Hạng mục đầu tư / mua sắm', 'Khối lượng 27', 'Khối lượng 28', 'Đơn giá', 'Loại lỗi', 'Chi tiết'];
  const tsvLines = [headers.join('\t')];

  trList.forEach(tr => {
    const tds = tr.querySelectorAll('td');
    if (tds.length === 9) {
      const line = Array.from(tds).map(td => td.innerText.replace(/\r?\n|\r/g, ' ').trim()).join('\t');
      tsvLines.push(line);
    }
  });
  const tsvText = tsvLines.join('\r\n');

  // 2. Tạo HTML Table định dạng chuẩn Excel với viền, màu sắc và căn lề
  let htmlTable = '<table border="1" style="border-collapse: collapse; font-family: Calibri, Arial, sans-serif; font-size: 10pt;">\r\n';
  htmlTable += '  <thead>\r\n';
  htmlTable += '    <tr style="background-color: #f1f5f9; font-weight: bold; color: #1e293b;">\r\n';
  htmlTable += '      <th style="border: 1px solid #94a3b8; padding: 6px; text-align: center;">STT</th>\r\n';
  htmlTable += '      <th style="border: 1px solid #94a3b8; padding: 6px; text-align: center;">Mảng</th>\r\n';
  htmlTable += '      <th style="border: 1px solid #94a3b8; padding: 6px; text-align: center;">Vị trí</th>\r\n';
  htmlTable += '      <th style="border: 1px solid #94a3b8; padding: 6px; text-align: left;">Hạng mục đầu tư / mua sắm</th>\r\n';
  htmlTable += '      <th style="border: 1px solid #94a3b8; padding: 6px; text-align: right;">Khối lượng 27</th>\r\n';
  htmlTable += '      <th style="border: 1px solid #94a3b8; padding: 6px; text-align: right;">Khối lượng 28</th>\r\n';
  htmlTable += '      <th style="border: 1px solid #94a3b8; padding: 6px; text-align: right;">Đơn giá</th>\r\n';
  htmlTable += '      <th style="border: 1px solid #94a3b8; padding: 6px; text-align: center;">Loại lỗi</th>\r\n';
  htmlTable += '      <th style="border: 1px solid #94a3b8; padding: 6px; text-align: left;">Chi tiết</th>\r\n';
  htmlTable += '    </tr>\r\n';
  htmlTable += '  </thead>\r\n';
  htmlTable += '  <tbody>\r\n';

  trList.forEach(tr => {
    const tds = tr.querySelectorAll('td');
    if (tds.length === 9) {
      htmlTable += '    <tr>\r\n';
      htmlTable += `      <td style="border: 1px solid #cbd5e1; padding: 4px 6px; text-align: center;">${escapeHtml(tds[0].innerText.trim())}</td>\r\n`;
      htmlTable += `      <td style="border: 1px solid #cbd5e1; padding: 4px 6px; text-align: center; font-weight: bold;">${escapeHtml(tds[1].innerText.trim())}</td>\r\n`;
      htmlTable += `      <td style="border: 1px solid #cbd5e1; padding: 4px 6px; text-align: center; color: #0284c7; font-weight: bold;">${escapeHtml(tds[2].innerText.trim())}</td>\r\n`;
      htmlTable += `      <td style="border: 1px solid #cbd5e1; padding: 4px 8px; text-align: left;">${escapeHtml(tds[3].innerText.trim())}</td>\r\n`;
      htmlTable += `      <td style="border: 1px solid #cbd5e1; padding: 4px 6px; text-align: right; font-family: monospace;">${escapeHtml(tds[4].innerText.trim())}</td>\r\n`;
      htmlTable += `      <td style="border: 1px solid #cbd5e1; padding: 4px 6px; text-align: right; font-family: monospace;">${escapeHtml(tds[5].innerText.trim())}</td>\r\n`;
      htmlTable += `      <td style="border: 1px solid #cbd5e1; padding: 4px 6px; text-align: right; font-family: monospace; font-weight: bold;">${escapeHtml(tds[6].innerText.trim())}</td>\r\n`;
      htmlTable += `      <td style="border: 1px solid #cbd5e1; padding: 4px 6px; text-align: center; font-weight: bold;">${escapeHtml(tds[7].innerText.trim())}</td>\r\n`;
      htmlTable += `      <td style="border: 1px solid #cbd5e1; padding: 4px 8px; text-align: left;">${escapeHtml(tds[8].innerText.trim())}</td>\r\n`;
      htmlTable += '    </tr>\r\n';
    }
  });
  htmlTable += '  </tbody>\r\n</table>';

  function triggerValCopySuccess() {
    const textEl = document.getElementById('copyValBtnText');
    const iconEl = document.getElementById('copyValBtnIcon');
    if (textEl) {
      const old = textEl.textContent;
      textEl.textContent = 'Đã sao chép!';
      setTimeout(() => { textEl.textContent = old; }, 2000);
    }
    if (iconEl) {
      iconEl.innerHTML = '<i data-lucide="check" class="w-4 h-4 text-emerald-600"></i>';
      if (window.lucide) lucide.createIcons();
      setTimeout(() => {
        iconEl.innerHTML = '<i data-lucide="copy" class="w-4 h-4 text-orange-600"></i>';
        if (window.lucide) lucide.createIcons();
      }, 2000);
    }
    showToast('Đã sao chép bảng kiểm tra lỗi vào Clipboard! Bạn có thể dán (Ctrl+V) ngay vào Excel.', 'success');
  }

  try {
    if (navigator.clipboard && window.ClipboardItem) {
      const blobHtml = new Blob([htmlTable], { type: 'text/html' });
      const blobText = new Blob([tsvText], { type: 'text/plain' });
      await navigator.clipboard.write([
        new ClipboardItem({
          'text/html': blobHtml,
          'text/plain': blobText
        })
      ]);
      triggerValCopySuccess();
    } else if (navigator.clipboard && navigator.clipboard.writeText) {
      await navigator.clipboard.writeText(tsvText);
      triggerValCopySuccess();
    } else {
      fallbackCopy(tsvText);
      triggerValCopySuccess();
    }
  } catch (err) {
    console.warn('ClipboardItem write failed, fallback to text/plain:', err);
    try {
      if (navigator.clipboard && navigator.clipboard.writeText) {
        await navigator.clipboard.writeText(tsvText);
      } else {
        fallbackCopy(tsvText);
      }
      triggerValCopySuccess();
    } catch (e2) {
      console.error('Clipboard copy failed:', e2);
      showToast('Không thể sao chép vào clipboard. Vui lòng thử lại!', 'error');
    }
  }
}

// Cho phép kéo dãn chiều rộng các cột trong bảng bằng chuột
function makeTableResizable(tableId) {
  const table = document.getElementById(tableId);
  if (!table) return;

  const thList = table.querySelectorAll('thead th');
  if (!thList.length) return;

  table.style.tableLayout = 'fixed';

  thList.forEach(th => {
    if (th.querySelector('.col-resizer')) return;

    th.style.position = 'relative';

    const resizer = document.createElement('div');
    resizer.className = 'col-resizer';
    resizer.title = 'Kéo để thay đổi độ rộng cột';
    th.appendChild(resizer);

    let startX = 0;
    let startWidth = 0;
    let startTableWidth = 0;

    const onMouseDown = (e) => {
      e.preventDefault();
      e.stopPropagation();

      // Cố định độ rộng hiện tại của tất cả các cột nếu chưa có style width
      thList.forEach(cell => {
        if (!cell.style.width) {
          cell.style.width = cell.offsetWidth + 'px';
        }
      });

      startX = e.pageX;
      startWidth = th.offsetWidth;
      startTableWidth = table.offsetWidth;

      resizer.classList.add('is-resizing');
      document.body.classList.add('resizing-col');

      const onMouseMove = (moveEvent) => {
        const delta = moveEvent.pageX - startX;
        const newWidth = Math.max(35, startWidth + delta);
        const actualDelta = newWidth - startWidth;
        th.style.width = newWidth + 'px';
        th.style.minWidth = newWidth + 'px';
        table.style.width = Math.max(startTableWidth, startTableWidth + actualDelta) + 'px';
        table.style.minWidth = table.style.width;
      };

      const onMouseUp = () => {
        resizer.classList.remove('is-resizing');
        document.body.classList.remove('resizing-col');
        document.removeEventListener('mousemove', onMouseMove);
        document.removeEventListener('mouseup', onMouseUp);
      };

      document.addEventListener('mousemove', onMouseMove);
      document.addEventListener('mouseup', onMouseUp);
    };

    resizer.addEventListener('mousedown', onMouseDown);
  });
}

// Khởi tạo các sự kiện filter
function initFilterHandlers() {
  document.querySelectorAll('input[name="filterIssueType"]').forEach(radio => {
    radio.addEventListener('change', () => {
      document.querySelectorAll('input[name="filterIssueType"]').forEach(r => {
        r.closest('.chip').classList.toggle('active', r.checked);
      });
      renderValidationTable();
    });
  });

  const searchValInput = document.getElementById('searchValidationInput');
  if (searchValInput) {
    searchValInput.addEventListener('input', renderValidationTable);
  }

  const btnCopyVal = document.getElementById('btnCopyValidationTable');
  if (btnCopyVal) {
    btnCopyVal.addEventListener('click', copyValidationTableToClipboard);
  }

  ['chkLevel1', 'chkLevel2', 'chkLevel3', 'chkLevel4', 'chkLevelLeaf'].forEach(id => {
    const el = document.getElementById(id);
    if (el) {
      el.addEventListener('change', (e) => {
        e.target.closest('.chip').classList.toggle('active', e.target.checked);
        renderPreviewTable();
      });
    }
  });

  const selectMang = document.getElementById('selectMangFilter');
  if (selectMang) {
    selectMang.addEventListener('change', renderPreviewTable);
  }

  const searchPrevInput = document.getElementById('searchPreviewInput');
  if (searchPrevInput) {
    searchPrevInput.addEventListener('input', renderPreviewTable);
  }

  document.querySelectorAll('.preview-column-toggle').forEach(input => {
    const key = input.dataset.previewColumn;
    if (!key || !(key in state.previewColumns)) return;
    input.checked = Boolean(state.previewColumns[key]);
    const chip = input.closest('.chip');
    if (chip) chip.classList.toggle('active', input.checked);
    input.addEventListener('change', (event) => {
      state.previewColumns[key] = event.target.checked;
      const parentChip = event.target.closest('.chip');
      if (parentChip) parentChip.classList.toggle('active', event.target.checked);
      renderPreviewTable();
    });
  });

  const searchGrpInput = document.getElementById('searchGroupInput');
  if (searchGrpInput) {
    searchGrpInput.addEventListener('input', renderHierarchyTable);
  }

  const btnCopyPrev = document.getElementById('btnCopyPreviewTable');
  if (btnCopyPrev) {
    btnCopyPrev.addEventListener('click', copyPreviewTableToClipboard);
  }

  const btnCopyPrevBanner = document.getElementById('btnCopyPreviewTableBanner');
  if (btnCopyPrevBanner) {
    btnCopyPrevBanner.addEventListener('click', copyPreviewTableToClipboard);
  }
}

// ==================== ENGINE XUẤT EXCEL BẰNG EXCELJS ====================

function prepareQuickDownload() {
  const link = document.getElementById('linkDownloadDirect');
  if (!link) return;
  link.onclick = (e) => {
    e.preventDefault();
    exportToExcel();
  };
}

function normalizeXmlChunk(str) {
  return str.replace(/\s+/g, ' ').replace(/> </g, '><').trim();
}

class JSStyleMerger {
  constructor(tmplStylesXml) {
    this.tmplStylesXml = tmplStylesXml;
    
    // Parse numFmts
    this.numFmtsMatch = tmplStylesXml.match(/<numFmts(\s+count="(\d+)")?>(.*?)<\/numFmts>/s);
    this.numFmtsList = [];
    this.numFmtMap = {};
    this.maxNumFmtId = 163;
    if (this.numFmtsMatch) {
      const inner = this.numFmtsMatch[3];
      const re = /<numFmt\s+numFmtId="(\d+)"\s+formatCode="([^"]*)"\s*\/>/g;
      let m;
      while ((m = re.exec(inner)) !== null) {
        const id = parseInt(m[1], 10);
        const code = m[2];
        this.numFmtsList.push({ id, code, full: m[0] });
        this.numFmtMap[code] = id;
        if (id > this.maxNumFmtId) this.maxNumFmtId = id;
      }
    }

    // Parse fonts
    this.fontsMatch = tmplStylesXml.match(/<fonts(\s+[^>]*?)>(.*?)<\/fonts>/s);
    this.fontsList = [];
    if (this.fontsMatch) {
      const re = /<font(?:\s+[^>]*?)?>.*?<\/font>/gs;
      let m;
      while ((m = re.exec(this.fontsMatch[2])) !== null) {
        this.fontsList.push(normalizeXmlChunk(m[0]));
      }
    }

    // Parse fills
    this.fillsMatch = tmplStylesXml.match(/<fills(\s+[^>]*?)>(.*?)<\/fills>/s);
    this.fillsList = [];
    if (this.fillsMatch) {
      const re = /<fill(?:\s+[^>]*?)?>.*?<\/fill>/gs;
      let m;
      while ((m = re.exec(this.fillsMatch[2])) !== null) {
        this.fillsList.push(normalizeXmlChunk(m[0]));
      }
    }

    // Parse borders
    this.bordersMatch = tmplStylesXml.match(/<borders(\s+[^>]*?)>(.*?)<\/borders>/s);
    this.bordersList = [];
    if (this.bordersMatch) {
      const re = /<border(?:\s+[^>]*?)?>.*?<\/border>/gs;
      let m;
      while ((m = re.exec(this.bordersMatch[2])) !== null) {
        this.bordersList.push(normalizeXmlChunk(m[0]));
      }
    }

    // Parse cellXfs
    this.cellXfsMatch = tmplStylesXml.match(/<cellXfs(\s+[^>]*?)>(.*?)<\/cellXfs>/s);
    this.cellXfsList = [];
    if (this.cellXfsMatch) {
      const re = /<xf\s+[^>]*?(?:\/>|>.*?<\/xf>)/gs;
      let m;
      while ((m = re.exec(this.cellXfsMatch[2])) !== null) {
        this.cellXfsList.push(normalizeXmlChunk(m[0]));
      }
    }
  }

  addFont(fontXml) {
    const norm = normalizeXmlChunk(fontXml);
    const idx = this.fontsList.indexOf(norm);
    if (idx !== -1) return idx;
    this.fontsList.push(norm);
    return this.fontsList.length - 1;
  }

  addFill(fillXml) {
    const norm = normalizeXmlChunk(fillXml);
    const idx = this.fillsList.indexOf(norm);
    if (idx !== -1) return idx;
    this.fillsList.push(norm);
    return this.fillsList.length - 1;
  }

  addBorder(borderXml) {
    const norm = normalizeXmlChunk(borderXml);
    const idx = this.bordersList.indexOf(norm);
    if (idx !== -1) return idx;
    this.bordersList.push(norm);
    return this.bordersList.length - 1;
  }

  mapNumFmt(srcNumFmtId, srcNumFmtDict) {
    const id = parseInt(srcNumFmtId, 10);
    if (isNaN(id) || id < 164) return id || 0;
    const fmtCode = srcNumFmtDict[id];
    if (!fmtCode) return id;
    if (this.numFmtMap[fmtCode] !== undefined) {
      return this.numFmtMap[fmtCode];
    }
    this.maxNumFmtId++;
    const newId = this.maxNumFmtId;
    this.numFmtMap[fmtCode] = newId;
    this.numFmtsList.push({
      id: newId,
      code: fmtCode,
      full: `<numFmt numFmtId="${newId}" formatCode="${fmtCode}"/>`
    });
    return newId;
  }

  createMapperForSource(srcStylesXml) {
    const srcNumFmts = {};
    const nfMatch = srcStylesXml.match(/<numFmts(\s+[^>]*?)>(.*?)<\/numFmts>/s);
    if (nfMatch) {
      const re = /<numFmt\s+numFmtId="(\d+)"\s+formatCode="([^"]*)"\s*\/>/g;
      let m;
      while ((m = re.exec(nfMatch[2])) !== null) {
        srcNumFmts[parseInt(m[1], 10)] = m[2];
      }
    }

    const fontMap = {};
    const fMatch = srcStylesXml.match(/<fonts(\s+[^>]*?)>(.*?)<\/fonts>/s);
    if (fMatch) {
      const re = /<font(?:\s+[^>]*?)?>.*?<\/font>/gs;
      let m, idx = 0;
      while ((m = re.exec(fMatch[2])) !== null) {
        fontMap[idx] = this.addFont(m[0]);
        idx++;
      }
    }

    const fillMap = {};
    const fiMatch = srcStylesXml.match(/<fills(\s+[^>]*?)>(.*?)<\/fills>/s);
    if (fiMatch) {
      const re = /<fill(?:\s+[^>]*?)?>.*?<\/fill>/gs;
      let m, idx = 0;
      while ((m = re.exec(fiMatch[2])) !== null) {
        fillMap[idx] = this.addFill(m[0]);
        idx++;
      }
    }

    const borderMap = {};
    const bMatch = srcStylesXml.match(/<borders(\s+[^>]*?)>(.*?)<\/borders>/s);
    if (bMatch) {
      const re = /<border(?:\s+[^>]*?)?>.*?<\/border>/gs;
      let m, idx = 0;
      while ((m = re.exec(bMatch[2])) !== null) {
        borderMap[idx] = this.addBorder(m[0]);
        idx++;
      }
    }

    const styleMap = {};
    const xfMatch = srcStylesXml.match(/<cellXfs(\s+[^>]*?)>(.*?)<\/cellXfs>/s);
    if (xfMatch) {
      const re = /<xf\s+([^>]*?)(?:\/>|>.*?<\/xf>)/gs;
      let m, idx = 0;
      while ((m = re.exec(xfMatch[2])) !== null) {
        const fullTag = m[0];
        const attrs = m[1];

        const fontIdMatch = attrs.match(/fontId="(\d+)"/);
        const fillIdMatch = attrs.match(/fillId="(\d+)"/);
        const borderIdMatch = attrs.match(/borderId="(\d+)"/);
        const numFmtIdMatch = attrs.match(/numFmtId="(\d+)"/);

        const srcFontId = fontIdMatch ? parseInt(fontIdMatch[1], 10) : 0;
        const srcFillId = fillIdMatch ? parseInt(fillIdMatch[1], 10) : 0;
        const srcBorderId = borderIdMatch ? parseInt(borderIdMatch[1], 10) : 0;
        const srcNumFmtId = numFmtIdMatch ? parseInt(numFmtIdMatch[1], 10) : 0;

        const targetFontId = fontMap[srcFontId] !== undefined ? fontMap[srcFontId] : 0;
        const targetFillId = fillMap[srcFillId] !== undefined ? fillMap[srcFillId] : 0;
        const targetBorderId = borderMap[srcBorderId] !== undefined ? borderMap[srcBorderId] : 0;
        const targetNumFmtId = this.mapNumFmt(srcNumFmtId, srcNumFmts);

        let newAttrs = attrs
          .replace(/fontId="\d+"/, `fontId="${targetFontId}"`)
          .replace(/fillId="\d+"/, `fillId="${targetFillId}"`)
          .replace(/borderId="\d+"/, `borderId="${targetBorderId}"`)
          .replace(/numFmtId="\d+"/, `numFmtId="${targetNumFmtId}"`)
          .replace(/xfId="\d+"/, `xfId="0"`);
        
        if (!newAttrs.includes('xfId=')) {
          newAttrs += ' xfId="0"';
        }

        let newXf;
        if (fullTag.endsWith('/>')) {
          newXf = `<xf ${newAttrs.trim()}/>`;
        } else {
          const innerContent = fullTag.substring(fullTag.indexOf('>') + 1, fullTag.lastIndexOf('<'));
          newXf = `<xf ${newAttrs.trim()}>${innerContent}</xf>`;
        }

        const normXf = normalizeXmlChunk(newXf);
        const existingIdx = this.cellXfsList.indexOf(normXf);
        if (existingIdx !== -1) {
          styleMap[idx] = existingIdx;
        } else {
          const targetIdx = this.cellXfsList.length;
          this.cellXfsList.push(normXf);
          styleMap[idx] = targetIdx;
        }
        idx++;
      }
    }

    return styleMap;
  }

  buildMergedStylesXml() {
    let result = this.tmplStylesXml;

    const numFmtsXml = `<numFmts count="${this.numFmtsList.length}">${this.numFmtsList.map(x => x.full).join('')}</numFmts>`;
    if (this.numFmtsMatch) {
      result = result.replace(this.numFmtsMatch[0], numFmtsXml);
    } else {
      result = result.replace(/<styleSheet([^>]*?)>/, `<styleSheet$1>${numFmtsXml}`);
    }

    const fontsXml = `<fonts count="${this.fontsList.length}" x14ac:knownFonts="1">${this.fontsList.join('')}</fonts>`;
    result = result.replace(this.fontsMatch[0], fontsXml);

    const fillsXml = `<fills count="${this.fillsList.length}">${this.fillsList.join('')}</fills>`;
    result = result.replace(this.fillsMatch[0], fillsXml);

    const bordersXml = `<borders count="${this.bordersList.length}">${this.bordersList.join('')}</borders>`;
    result = result.replace(this.bordersMatch[0], bordersXml);

    const cellXfsXml = `<cellXfs count="${this.cellXfsList.length}">${this.cellXfsList.join('')}</cellXfs>`;
    result = result.replace(this.cellXfsMatch[0], cellXfsXml);

    return result;
  }
}

async function getSourceWorksheetXml(zip) {
  try {
    const wbXml = await zip.file('xl/workbook.xml')?.async('string');
    const relsXml = await zip.file('xl/_rels/workbook.xml.rels')?.async('string');
    if (wbXml && relsXml) {
      const sheetMatch = wbXml.match(/<sheet\s+[^>]*?name="([^"]*(?:PL1\.1|ML2027)[^"]*)"[^>]*?r:id="([^"]+)"/i) ||
                         wbXml.match(/<sheet\s+[^>]*?r:id="([^"]+)"[^>]*?name="([^"]*(?:PL1\.1|ML2027)[^"]*)"/i);
      if (sheetMatch) {
        const rId = sheetMatch[2] || sheetMatch[1];
        const relMatch = relsXml.match(new RegExp(`<Relationship\\s+[^>]*?Id="${rId}"[^>]*?Target="([^"]+)"`, 'i'));
        if (relMatch) {
          const target = relMatch[1].replace(/^\//, '');
          const fullPath = target.startsWith('xl/') ? target : `xl/${target}`;
          const sheetXml = await zip.file(fullPath)?.async('string');
          if (sheetXml) return sheetXml;
        }
      }
    }
  } catch (e) {}
  return await zip.file('xl/worksheets/sheet1.xml')?.async('string');
}

// Tạo buffer Excel sạch không lỗi XML: Cập nhật trực tiếp sheet1.xml (PL1.1) bên trong zip của phôi mẫu,
// đồng thời sao chép hoàn chỉnh 100% định dạng (font, in đậm, kẻ viền, tô màu, chiều cao dòng) từ file gốc,
// bảo tồn nguyên vẹn 100% các sheet khác (TH theo DV, PL1.2, PL1.3, Mã), styles, themes, và externalLinks.
async function buildCleanMasterlist(templateBuffer, extractedByMang, filesObj) {
  const JSZip = window.JSZip;
  if (!JSZip) {
    throw new Error('Thư viện JSZip chưa được nạp!');
  }
  const zip = await JSZip.loadAsync(templateBuffer);

  const tmplStylesXml = await zip.file('xl/styles.xml').async('string');
  const merger = new JSStyleMerger(tmplStylesXml);

  // Thu thập style và thông tin dòng gốc từ từng file mảng đầu vào
  const mangSourceData = {};
  if (filesObj) {
    for (const mang of MANG_CONFIG) {
      const code = mang.code;
      const fileObj = filesObj[code];
      if (fileObj && fileObj.buffer) {
        try {
          const srcZip = await JSZip.loadAsync(fileObj.buffer);
          const srcStylesXml = await srcZip.file('xl/styles.xml')?.async('string');
          let smap = {};
          if (srcStylesXml) {
            smap = merger.createMapperForSource(srcStylesXml);
          }

          const srcSheetXml = await getSourceWorksheetXml(srcZip);
          const sourceRows = {};
          if (srcSheetXml) {
            const rowRe = /<row\s+([^>]*?)r="(\d+)"([^>]*?)>(.*?)<\/row>/gs;
            let rm;
            while ((rm = rowRe.exec(srcSheetXml)) !== null) {
              const rNum = parseInt(rm[2], 10);
              const combinedAttrs = rm[1] + ' ' + rm[3];
              const htMatch = combinedAttrs.match(/ht="([^"]+)"/);
              const sMatch = combinedAttrs.match(/s="(\d+)"/);
              const olMatch = combinedAttrs.match(/outlineLevel="([^"]+)"/);

              const mappedRowS = sMatch && smap[parseInt(sMatch[1], 10)] !== undefined ? smap[parseInt(sMatch[1], 10)] : null;
              const cells = {};

              const cellRe = /<c\s+([^>]*?)r="([A-Z]+)\d+"([^>]*?)(?:\/>|>.*?<\/c>)/gs;
              let cm;
              while ((cm = cellRe.exec(rm[4])) !== null) {
                const col = cm[2];
                const cAttrs = cm[1] + ' ' + cm[3];
                const csMatch = cAttrs.match(/s="(\d+)"/);
                if (csMatch) {
                  const origS = parseInt(csMatch[1], 10);
                  cells[col] = smap[origS] !== undefined ? smap[origS] : null;
                }
              }

              sourceRows[rNum] = {
                ht: htMatch ? htMatch[1] : null,
                outlineLevel: olMatch ? olMatch[1] : null,
                mappedRowS: mappedRowS,
                cells: cells
              };
            }
          }

          mangSourceData[code] = { smap, sourceRows };
        } catch (errSrc) {
          console.warn(`Không thể đọc style từ file mảng ${code}:`, errSrc);
        }
      }
    }
  }

  let wbXml = await zip.file('xl/workbook.xml')?.async('string');
  const relsXml = await zip.file('xl/_rels/workbook.xml.rels')?.async('string');
  let masterlistSheetPath = 'xl/worksheets/sheet1.xml';
  if (wbXml && relsXml) {
    const sheetMatch = wbXml.match(/<sheet\s+[^>]*?name="([^"]*(?:PL1\.1|ML2027|Masterlist)[^"]*)"[^>]*?r:id="([^"]+)"/i) ||
                       wbXml.match(/<sheet\s+[^>]*?r:id="([^"]+)"[^>]*?name="([^"]*(?:PL1\.1|ML2027|Masterlist)[^"]*)"/i);
    if (sheetMatch) {
      const rId = sheetMatch[2] || sheetMatch[1];
      const relMatch = relsXml.match(new RegExp(`<Relationship\\s+[^>]*?Id="${rId}"[^>]*?Target="([^"]+)"`, 'i'));
      if (relMatch) {
        const target = relMatch[1].replace(/^\//, '');
        masterlistSheetPath = target.startsWith('xl/') ? target : `xl/${target}`;
      }
    }
  }

  let sheet1Xml = await zip.file(masterlistSheetPath).async('string');
  const matchData = sheet1Xml.match(/<sheetData>(.*?)<\/sheetData>/s);
  if (!matchData) throw new Error(`Không tìm thấy sheetData trong ${masterlistSheetPath} của file phôi`);

  const origSheetData = matchData[1];
  const rowRegex = /<row\s+[^>]*?r="([0-9]+)"[^>]*?>.*?<\/row>/gs;
  const rows1To9 = {};
  let m;
  while ((m = rowRegex.exec(origSheetData)) !== null) {
    const rInt = parseInt(m[1], 10);
    if (rInt <= 9) rows1To9[rInt] = m[0];
  }

  let currentRow = 10;
  const sectionMapping = {};
  const newRowsXml = [];

  const COLS = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q'];

  for (const mang of MANG_CONFIG) {
    const code = mang.code;
    const items = extractedByMang[code] || [];
    const secStart = currentRow;
    const secR = currentRow++;
    const itemRowMap = {};
    const childRows = [];
    const srcData = mangSourceData[code];
    const sourceRows = srcData ? srcData.sourceRows : {};

    for (let idx = 0; idx < items.length; idx++) {
      const it = items[idx];
      const rNum = currentRow++;
      itemRowMap[idx] = rNum;

      const srcRow = sourceRows[it.origRow];
      const srcCells = srcRow ? srcRow.cells : {};

      const sA = srcCells['A'] !== undefined && srcCells['A'] !== null ? srcCells['A'] : (it.isGroup ? 320 : 400);
      const sB = srcCells['B'] !== undefined && srcCells['B'] !== null ? srcCells['B'] : (it.isGroup ? 321 : 557);
      const sC = srcCells['C'] !== undefined && srcCells['C'] !== null ? srcCells['C'] : 397;
      const sD = srcCells['D'] !== undefined && srcCells['D'] !== null ? srcCells['D'] : 323;
      const sE = srcCells['E'] !== undefined && srcCells['E'] !== null ? srcCells['E'] : 323;
      const sF = srcCells['F'] !== undefined && srcCells['F'] !== null ? srcCells['F'] : 323;
      const sG = srcCells['G'] !== undefined && srcCells['G'] !== null ? srcCells['G'] : 323;
      const sH = srcCells['H'] !== undefined && srcCells['H'] !== null ? srcCells['H'] : 323;
      const sI = srcCells['I'] !== undefined && srcCells['I'] !== null ? srcCells['I'] : 399;
      const sJ = srcCells['J'] !== undefined && srcCells['J'] !== null ? srcCells['J'] : 399;
      const sK = srcCells['K'] !== undefined && srcCells['K'] !== null ? srcCells['K'] : 399;
      const sL = srcCells['L'] !== undefined && srcCells['L'] !== null ? srcCells['L'] : 399;

      const cellMap = {};
      if (it.tt) cellMap['A'] = `<c r="A${rNum}" s="${sA}" t="inlineStr"><is><t>${escapeXml(it.tt)}</t></is></c>`;
      else if ('A' in srcCells) cellMap['A'] = `<c r="A${rNum}" s="${sA}"/>`;

      if (it.nd) cellMap['B'] = `<c r="B${rNum}" s="${sB}" t="inlineStr"><is><t>${escapeXml(it.nd)}</t></is></c>`;
      else if ('B' in srcCells) cellMap['B'] = `<c r="B${rNum}" s="${sB}"/>`;

      if (it.dvt) cellMap['C'] = `<c r="C${rNum}" s="${sC}" t="inlineStr"><is><t>${escapeXml(it.dvt)}</t></is></c>`;
      else if ('C' in srcCells) cellMap['C'] = `<c r="C${rNum}" s="${sC}"/>`;

      if (it.kl27 !== null && it.kl27 !== undefined) cellMap['D'] = `<c r="D${rNum}" s="${sD}"><v>${it.kl27}</v></c>`;
      else if ('D' in srcCells) cellMap['D'] = `<c r="D${rNum}" s="${sD}"/>`;

      if (it.kl28 !== null && it.kl28 !== undefined) cellMap['E'] = `<c r="E${rNum}" s="${sE}"><v>${it.kl28}</v></c>`;
      else if ('E' in srcCells) cellMap['E'] = `<c r="E${rNum}" s="${sE}"/>`;

      if (it.dg !== null && it.dg !== undefined) cellMap['F'] = `<c r="F${rNum}" s="${sF}"><v>${it.dg}</v></c>`;
      else if ('F' in srcCells) cellMap['F'] = `<c r="F${rNum}" s="${sF}"/>`;

      if (!it.isGroup) {
        cellMap['G'] = `<c r="G${rNum}" s="${sG}"><f>F${rNum}*D${rNum}</f></c>`;
        cellMap['H'] = `<c r="H${rNum}" s="${sH}"><f>F${rNum}*E${rNum}</f></c>`;
      } else {
        if ('G' in srcCells) cellMap['G'] = `<c r="G${rNum}" s="${sG}"/>`;
        if ('H' in srcCells) cellMap['H'] = `<c r="H${rNum}" s="${sH}"/>`;
      }

      if (it.donViDT) cellMap['I'] = `<c r="I${rNum}" s="${sI}" t="inlineStr"><is><t>${escapeXml(it.donViDT)}</t></is></c>`;
      else if ('I' in srcCells) cellMap['I'] = `<c r="I${rNum}" s="${sI}"/>`;

      if (it.maMang) cellMap['J'] = `<c r="J${rNum}" s="${sJ}" t="inlineStr"><is><t>${escapeXml(it.maMang)}</t></is></c>`;
      else if ('J' in srcCells) cellMap['J'] = `<c r="J${rNum}" s="${sJ}"/>`;

      if (it.maDV) cellMap['K'] = `<c r="K${rNum}" s="${sK}" t="inlineStr"><is><t>${escapeXml(it.maDV)}</t></is></c>`;
      else if ('K' in srcCells) cellMap['K'] = `<c r="K${rNum}" s="${sK}"/>`;

      if (it.maLoai) cellMap['L'] = `<c r="L${rNum}" s="${sL}" t="inlineStr"><is><t>${escapeXml(it.maLoai)}</t></is></c>`;
      else if ('L' in srcCells) cellMap['L'] = `<c r="L${rNum}" s="${sL}"/>`;

      // Chuẩn hóa Group dòng (Outline Level 1..5) theo đúng cây phân cấp đã nhận dạng
      let ol = null;
      if (it.outlineLevel !== undefined && it.outlineLevel !== null) {
        ol = it.outlineLevel;
      } else {
        const node = state.groupsConfig ? state.groupsConfig.find(g => g.itemIndex === it.index) : null;
        if (node && node.ancestors) {
          ol = Math.min(7, Math.max(1, node.ancestors.length));
        } else if (srcRow && srcRow.outlineLevel) {
          ol = srcRow.outlineLevel;
        } else {
          ol = it.isGroup ? 1 : 2;
        }
      }

      const htAttr = (srcRow && srcRow.ht) ? ` ht="${srcRow.ht}" customHeight="1"` : '';
      const sRowAttr = (srcRow && srcRow.mappedRowS !== null) ? ` s="${srcRow.mappedRowS}" customFormat="1"` : (it.isGroup ? ' s="328"' : '');
      const olAttr = ol ? ` outlineLevel="${ol}"` : '';

      childRows.push({
        idx, rNum,
        rowAttrs: `${htAttr}${sRowAttr}${olAttr}`,
        cellMap, it, sG, sH
      });
    }

    const secEnd = currentRow - 1;
    sectionMapping[code] = { start: secStart, end: secEnd };

    const headerInfo = state.mangHeaderInfo ? state.mangHeaderInfo[code] : null;
    const secTT = (headerInfo && headerInfo.tt) ? headerInfo.tt : mang.tt;
    const secND = (headerInfo && headerInfo.nd) ? headerInfo.nd : mang.name.toUpperCase();
    const srcHeaderRow = (headerInfo && sourceRows) ? sourceRows[headerInfo.origRow] : null;
    const srcHeaderCells = srcHeaderRow ? srcHeaderRow.cells : {};

    const sec_sA = srcHeaderCells['A'] !== undefined && srcHeaderCells['A'] !== null ? srcHeaderCells['A'] : 320;
    const sec_sB = srcHeaderCells['B'] !== undefined && srcHeaderCells['B'] !== null ? srcHeaderCells['B'] : 321;
    const sec_sG = srcHeaderCells['G'] !== undefined && srcHeaderCells['G'] !== null ? srcHeaderCells['G'] : 323;
    const sec_sH = srcHeaderCells['H'] !== undefined && srcHeaderCells['H'] !== null ? srcHeaderCells['H'] : 323;
    const sec_sJ = srcHeaderCells['J'] !== undefined && srcHeaderCells['J'] !== null ? srcHeaderCells['J'] : 324;
    const secRowS = srcHeaderRow && srcHeaderRow.mappedRowS !== null ? ` s="${srcHeaderRow.mappedRowS}" customFormat="1"` : ' s="328" customFormat="1"';

    const secCellMap = {
      'A': `<c r="A${secR}" s="${sec_sA}" t="inlineStr"><is><t>${escapeXml(secTT)}</t></is></c>`,
      'B': `<c r="B${secR}" s="${sec_sB}" t="inlineStr"><is><t>${escapeXml(secND)}</t></is></c>`
    };
    if (secEnd > secStart) {
      secCellMap['G'] = `<c r="G${secR}" s="${sec_sG}"><f>SUBTOTAL(9,G${secStart + 1}:G${secEnd})</f></c>`;
      secCellMap['H'] = `<c r="H${secR}" s="${sec_sH}"><f>SUBTOTAL(9,H${secStart + 1}:H${secEnd})</f></c>`;
    }
    secCellMap['J'] = `<c r="J${secR}" s="${sec_sJ}" t="inlineStr"><is><t>${code}</t></is></c>`;
    newRowsXml.push(`<row r="${secR}" spans="1:17"${secRowS}>${COLS.map(c => secCellMap[c] || '').join('')}</row>`);

    for (const itemObj of childRows) {
      const it = itemObj.it;
      if (it.isGroup && it.hasSubtotal && it.subtotalStartRow && it.subtotalEndRow) {
        let cStart = null;
        let cEnd = null;
        for (let k = 0; k < items.length; k++) {
          const oR = items[k].origRow;
          if (oR && oR >= it.subtotalStartRow && oR <= it.subtotalEndRow) {
            if (cStart === null) cStart = itemRowMap[k];
            cEnd = itemRowMap[k];
          }
        }
        if ((cStart === null || cEnd === null) && it.childStartIdxInMang !== null && it.childStartIdxInMang !== undefined && it.childEndIdxInMang !== null && it.childEndIdxInMang !== undefined) {
          cStart = itemRowMap[it.childStartIdxInMang];
          cEnd = itemRowMap[it.childEndIdxInMang];
        }
        if (cStart !== null && cEnd !== null) {
          itemObj.cellMap['G'] = `<c r="G${itemObj.rNum}" s="${itemObj.sG}"><f>SUBTOTAL(9,G${cStart}:G${cEnd})</f></c>`;
          itemObj.cellMap['H'] = `<c r="H${itemObj.rNum}" s="${itemObj.sH}"><f>SUBTOTAL(9,H${cStart}:H${cEnd})</f></c>`;
        }
      }
      newRowsXml.push(`<row r="${itemObj.rNum}" spans="1:17"${itemObj.rowAttrs}>${COLS.map(c => itemObj.cellMap[c] || '').join('')}</row>`);
    }
  }

  const totalRows = currentRow - 1;
  const cdEnd = sectionMapping['CD'] ? sectionMapping['CD'].end : null;
  const tdEnd = sectionMapping['TD'] ? sectionMapping['TD'].end : null;
  const lastBeforeHt = cdEnd || tdEnd || totalRows;

  function updateRowSubtotals(rInt, startR, endR) {
    let rowStr = rows1To9[rInt] || '';
    rowStr = rowStr.replace(new RegExp(`(<c\\s+[^>]*?r="G${rInt}"[^>]*?>)(.*?)(</c>)`, 's'), `$1<f>SUBTOTAL(9,G${startR}:G${endR})</f>$3`);
    rowStr = rowStr.replace(new RegExp(`(<c\\s+[^>]*?r="H${rInt}"[^>]*?>)(.*?)(</c>)`, 's'), `$1<f>SUBTOTAL(9,H${startR}:H${endR})</f>$3`);
    rows1To9[rInt] = rowStr;
  }

  updateRowSubtotals(9, 10, lastBeforeHt);
  updateRowSubtotals(8, 9, totalRows);
  updateRowSubtotals(7, 8, totalRows);

  const allRows = [];
  for (let i = 1; i <= 9; i++) {
    if (rows1To9[i]) allRows.push(rows1To9[i]);
  }
  allRows.push(...newRowsXml);
  const newSheetData = `<sheetData>${allRows.join('')}</sheetData>`;

  sheet1Xml = sheet1Xml.slice(0, matchData.index) + newSheetData + sheet1Xml.slice(matchData.index + matchData[0].length);
  sheet1Xml = sheet1Xml.replace(/<dimension\s+ref="[^"]*"/, `<dimension ref="A1:Q${totalRows}"`);
  sheet1Xml = sheet1Xml.replace(/<autoFilter\s+ref="[^"]*"/, `<autoFilter ref="A7:Q${totalRows}"`);
  sheet1Xml = sheet1Xml.replace(/<dataValidations\s+count="[^"]*">.*?<\/dataValidations>/s, '');

  // Đảm bảo cấu hình Outline của Excel (summaryBelow="0" để dấu +/- nằm ở dòng tiêu đề nhóm phía trên)
  if (sheet1Xml.includes('<outlinePr')) {
    sheet1Xml = sheet1Xml.replace(/<outlinePr[^>]*\/>/, '<outlinePr summaryBelow="0" summaryRight="0"/>');
  } else if (sheet1Xml.includes('<sheetPr')) {
    sheet1Xml = sheet1Xml.replace(/<sheetPr(.*?)>/, '<sheetPr$1><outlinePr summaryBelow="0" summaryRight="0"/>');
  }

  zip.file(masterlistSheetPath, sheet1Xml);
  zip.file('xl/styles.xml', merger.buildMergedStylesXml());

  // Xóa calcChain để Excel tự tính toán lại công thức từ đầu khi mở file, tránh lỗi cache
  zip.remove('xl/calcChain.xml');

  let contentTypes = await zip.file('[Content_Types].xml').async('string');
  contentTypes = contentTypes.replace(/<Override\s+[^>]*?PartName="\/xl\/calcChain\.xml"[^>]*?\/>/g, '');
  zip.file('[Content_Types].xml', contentTypes);

  let wbRels = await zip.file('xl/_rels/workbook.xml.rels').async('string');
  wbRels = wbRels.replace(/<Relationship\s+[^>]*?Target="calcChain\.xml"[^>]*?\/>/g, '');
  zip.file('xl/_rels/workbook.xml.rels', wbRels);

  // Bắt buộc Excel tính toán lại 100% tất cả công thức ở TẤT CẢ các sheet khi mở file
  wbXml = await zip.file('xl/workbook.xml').async('string');
  if (wbXml.includes('<calcPr')) {
    wbXml = wbXml.replace(/<calcPr([^>]*?)\/?>/, (match, p1) => {
      let attrs = p1;
      if (!attrs.includes('fullCalcOnLoad')) attrs += ' fullCalcOnLoad="1"';
      if (!attrs.includes('forceFullCalculation')) attrs += ' forceFullCalculation="1"';
      if (!attrs.includes('calcMode')) attrs += ' calcMode="auto"';
      return `<calcPr${attrs}/>`;
    });
  } else {
    wbXml = wbXml.replace('</workbook>', '<calcPr fullCalcOnLoad="1" forceFullCalculation="1" calcMode="auto"/></workbook>');
  }
  zip.file('xl/workbook.xml', wbXml);

  return await zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
}


// Kích hoạt tải file Blob xuống trình duyệt
function triggerDownloadBlob(blob, fileName) {
  try {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }, 1000);
  } catch (e) {
    console.error('Lỗi tải blob:', e);
  }
}





// ==================== TIỆN ÍCH DÙNG CHUNG ====================

function formatNumber(val) {
  if (val === null || val === undefined || isNaN(val)) return '0';
  const num = typeof val === 'number' ? val : parseFloat(val);
  return num.toLocaleString('en-US', { maximumFractionDigits: 2 });
}

function escapeHtml(text) {
  if (!text) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function escapeXml(s) {
  if (s === null || s === undefined) return '';
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

// ==================== TÍNH NĂNG SO SÁNH VỚI CHIẾN LƯỢC 5 NĂM ====================

const DEFAULT_STRATEGY_DATA = {
  VT: { name: 'Vô tuyến', strat27: 37.04, strat28: 28.27 },
  ML: { name: 'Mạng lõi', strat27: 4.14, strat28: 5.15 },
  CDBR: { name: 'BRCĐ-TH', strat27: 0.95, strat28: 0.96 },
  CNTT: { name: 'CNTT', strat27: 12.92, strat28: 5.66 },
  TD: { name: 'Truyền dẫn', strat27: 21.02, strat28: 18.51 },
  CD: { name: 'Cơ điện', strat27: 13.07, strat28: 26.28 },
  HT: { name: 'Hạ tầng', strat27: 14.21, strat28: 18.04 },
  ATTT: { name: 'ATTT', strat27: 0.58, strat28: 0.68 },
  VI: { name: 'Ví', strat27: 7.62, strat28: 7.62 }
};

const STRATEGY_ROWS = [
  { key: 'KHONG_VI', name: 'Tổng đầu tư (không gồm ví) (m$)', isTotal: true, className: 'strat-row-total-not-wallet' },
  { key: 'TONG', name: 'Tổng đầu tư (m$)', isTotal: true, className: 'strat-row-total-all' },
  { key: 'VT', name: 'Vô tuyến', badge: 'VT', isTotal: false },
  { key: 'ML', name: 'Mạng lõi', badge: 'ML', isTotal: false },
  { key: 'CDBR', name: 'BRCĐ-TH', badge: 'CĐBR', isTotal: false },
  { key: 'CNTT', name: 'CNTT', badge: 'CNTT', isTotal: false },
  { key: 'TD', name: 'Truyền dẫn', badge: 'TD', isTotal: false },
  { key: 'CD', name: 'Cơ điện', badge: 'CĐ', isTotal: false },
  { key: 'HT', name: 'Hạ tầng', badge: 'HT', isTotal: false },
  { key: 'ATTT', name: 'ATTT', badge: 'ATTT', isTotal: false },
  { key: 'VI', name: 'Ví', badge: 'Ví', isTotal: false }
];

const STRATEGY_ITEM_KEYS = ['VT', 'ML', 'CDBR', 'CNTT', 'TD', 'CD', 'HT', 'ATTT', 'VI'];
const STRATEGY_KHONG_VI_KEYS = ['VT', 'ML', 'CDBR', 'CNTT', 'TD', 'CD', 'HT', 'ATTT'];
const STRATEGY_TOTAL_KEYS = ['KHONG_VI', 'TONG'];

function getNormalizedStrategyRowOrder(profile) {
  const rawOrder = Array.isArray(profile?.rowOrder) ? profile.rowOrder : STRATEGY_ITEM_KEYS;
  const orderedKeys = rawOrder.filter(key => STRATEGY_ITEM_KEYS.includes(key));
  const missingKeys = STRATEGY_ITEM_KEYS.filter(key => !orderedKeys.includes(key));
  return [...orderedKeys, ...missingKeys];
}

function getStrategyDisplayRows(profile) {
  const totals = STRATEGY_TOTAL_KEYS.map(key => STRATEGY_ROWS.find(row => row.key === key)).filter(Boolean);
  const details = getNormalizedStrategyRowOrder(profile)
    .map(key => STRATEGY_ROWS.find(row => row.key === key))
    .filter(Boolean);
  return [...totals, ...details];
}

function formatStratShare(part, total) {
  if (part === null || part === undefined || total === null || total === undefined) return '';
  const numerator = typeof part === 'number' ? part : parseFloat(part);
  const denominator = typeof total === 'number' ? total : parseFloat(total);
  if (isNaN(numerator) || isNaN(denominator) || Math.abs(denominator) < 0.0001) return '';
  const pct = (numerator / denominator) * 100;
  if (Math.abs(pct) < 0.0001) return '<span style="color: #94a3b8;">-</span>';
  return `${pct.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}%`;
}

function formatStratShareRaw(part, total) {
  if (part === null || part === undefined || total === null || total === undefined) return '';
  const numerator = typeof part === 'number' ? part : parseFloat(part);
  const denominator = typeof total === 'number' ? total : parseFloat(total);
  if (isNaN(numerator) || isNaN(denominator) || Math.abs(denominator) < 0.0001) return '';
  const pct = (numerator / denominator) * 100;
  if (Math.abs(pct) < 0.0001) return '-';
  return `${pct.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}%`;
}

// Khởi tạo các sự kiện cho tab Chiến lược 5 năm
function initStrategyComparison() {
  // Nút Sao chép bảng (Excel)
  const btnCopy = document.getElementById('btnCopyStrategyTable');
  if (btnCopy) {
    btnCopy.addEventListener('click', copyStrategyTableToClipboard);
  }

  const strategyBody = document.getElementById('strategyTableBody');
  if (strategyBody) {
    strategyBody.addEventListener('click', (event) => {
      const btn = event.target.closest('[data-strategy-move]');
      if (!btn) return;
      const key = btn.getAttribute('data-key');
      const direction = btn.getAttribute('data-strategy-move');
      moveStrategyRow(key, direction);
    });
  }

  // Chọn Profile
  const selProfile = document.getElementById('selectStrategyProfile');
  if (selProfile) {
    selProfile.addEventListener('change', (e) => {
      state.activeStrategyProfileId = e.target.value;
      renderStrategyComparisonTable();
    });
  }

  // Nút Thêm Profile
  const btnOpenModal = document.getElementById('btnOpenNewProfileModal');
  if (btnOpenModal) {
    btnOpenModal.addEventListener('click', () => openProfileModal('create'));
  }

  // Nút Chỉnh sửa số liệu Profile
  const btnEditModal = document.getElementById('btnEditProfileModal');
  if (btnEditModal) {
    btnEditModal.addEventListener('click', () => openProfileModal('edit'));
  }

  // Nút Xác nhận tạo/lưu Profile trong Modal
  const btnConfirm = document.getElementById('btnConfirmCreateProfile');
  if (btnConfirm) {
    btnConfirm.addEventListener('click', confirmSaveProfileModal);
  }

  // Các nút gán nhanh số liệu trong Modal
  const btnFillCurrent = document.getElementById('btnModalFillCurrent');
  if (btnFillCurrent) btnFillCurrent.addEventListener('click', () => fillModalData('current'));

  const btnFillDefault = document.getElementById('btnModalFillDefault');
  if (btnFillDefault) btnFillDefault.addEventListener('click', () => fillModalData('default'));

  const btnFillZero = document.getElementById('btnModalFillZero');
  if (btnFillZero) btnFillZero.addEventListener('click', () => fillModalData('zero'));

  // Nút Nhân bản Profile
  const btnClone = document.getElementById('btnCloneProfile');
  if (btnClone) {
    btnClone.addEventListener('click', cloneCurrentProfile);
  }

  // Nút Xóa Profile
  const btnDelete = document.getElementById('btnDeleteProfile');
  if (btnDelete) {
    btnDelete.addEventListener('click', deleteCurrentProfile);
  }


  // Tải danh sách profiles
  loadStrategyProfiles();
}

// Nạp danh sách profiles: ưu tiên localStorage, nếu chưa có thì nạp từ file tĩnh strategy_profiles.json
async function loadStrategyProfiles() {
  // 1. Kiểm tra trong localStorage (nếu người dùng đã từng lưu/sửa trên trình duyệt)
  try {
    const cached = localStorage.getItem('QHDC_STRATEGY_PROFILES');
    if (cached) {
      const parsed = JSON.parse(cached);
      if (parsed && Array.isArray(parsed.profiles) && parsed.profiles.length > 0) {
        state.strategyProfiles = parsed.profiles;
        state.activeStrategyProfileId = parsed.activeProfileId || parsed.profiles[0].id;
        renderStrategyProfileSelect();
        renderStrategyComparisonTable();
        return;
      }
    }
  } catch (e) {}

  // 2. Nạp từ file tĩnh strategy_profiles.json đi kèm dự án trên GitHub Pages
  try {
    const resp = await fetch('strategy_profiles.json');
    if (resp.ok) {
      const data = await resp.json();
      if (data && Array.isArray(data.profiles) && data.profiles.length > 0) {
        state.strategyProfiles = data.profiles;
        state.activeStrategyProfileId = data.activeProfileId || data.profiles[0].id;
        renderStrategyProfileSelect();
        renderStrategyComparisonTable();
        return;
      }
    }
  } catch (e) {}

  // 3. Phương án mặc định nếu chưa có dữ liệu
  if (!state.strategyProfiles || state.strategyProfiles.length === 0) {
    state.strategyProfiles = [
      {
        id: 'profile_default',
        name: 'Chiến lược 5 năm (Phương án Chuẩn)',
        description: 'Dữ liệu kế hoạch chiến lược 5 năm 2026-2030 theo phê duyệt',
        updatedAt: new Date().toISOString(),
        data: JSON.parse(JSON.stringify(DEFAULT_STRATEGY_DATA))
      }
    ];
    state.activeStrategyProfileId = 'profile_default';
  }

  renderStrategyProfileSelect();
  renderStrategyComparisonTable();
}

// Lưu profiles vào localStorage của trình duyệt (100% Client-side, không cần server)
async function saveStrategyProfiles() {
  const payload = {
    activeProfileId: state.activeStrategyProfileId,
    profiles: state.strategyProfiles
  };

  try {
    localStorage.setItem('QHDC_STRATEGY_PROFILES', JSON.stringify(payload));
    return true;
  } catch (e) {
    console.warn('Lỗi khi lưu vào localStorage:', e);
    return false;
  }
}

async function moveStrategyRow(key, direction) {
  const profile = getActiveStrategyProfile();
  if (!profile || !STRATEGY_ITEM_KEYS.includes(key)) return;

  const order = getNormalizedStrategyRowOrder(profile);
  const currentIndex = order.indexOf(key);
  if (currentIndex < 0) return;

  const targetIndex = direction === 'up' ? currentIndex - 1 : currentIndex + 1;
  if (targetIndex < 0 || targetIndex >= order.length) return;

  [order[currentIndex], order[targetIndex]] = [order[targetIndex], order[currentIndex]];
  profile.rowOrder = order;
  profile.updatedAt = new Date().toISOString();

  const saved = await saveStrategyProfiles();
  renderStrategyComparisonTable();
  showToast(saved ? 'Đã cập nhật thứ tự hiển thị dòng.' : 'Đã đổi thứ tự dòng, nhưng chưa lưu được vào bộ nhớ trình duyệt.', saved ? 'success' : 'warning');
}

// Lấy profile đang active
function getActiveStrategyProfile() {
  if (!state.strategyProfiles || state.strategyProfiles.length === 0) return null;
  const found = state.strategyProfiles.find(p => p.id === state.activeStrategyProfileId);
  return found || state.strategyProfiles[0];
}

// Hiển thị danh sách profile lên dropdown
function renderStrategyProfileSelect() {
  const sel = document.getElementById('selectStrategyProfile');
  if (!sel) return;
  sel.innerHTML = state.strategyProfiles.map(p => `
    <option value="${p.id}" ${p.id === state.activeStrategyProfileId ? 'selected' : ''}>
      ${escapeHtml(p.name)}
    </option>
  `).join('');
}

// Trích xuất số liệu QHĐC tự động thời gian thực (realtime) từ Masterlist
function getQHDCDataFromMasterlist() {
  const result = {
    VT: { qhdc27: null, qhdc28: null, qhdcTong: null, hasData: false },
    ML: { qhdc27: null, qhdc28: null, qhdcTong: null, hasData: false },
    CDBR: { qhdc27: null, qhdc28: null, qhdcTong: null, hasData: false },
    CNTT: { qhdc27: null, qhdc28: null, qhdcTong: null, hasData: false },
    TD: { qhdc27: null, qhdc28: null, qhdcTong: null, hasData: false },
    CD: { qhdc27: null, qhdc28: null, qhdcTong: null, hasData: false },
    HT: { qhdc27: null, qhdc28: null, qhdcTong: null, hasData: false },
    ATTT: { qhdc27: null, qhdc28: null, qhdcTong: null, hasData: false },
    VI: { qhdc27: null, qhdc28: null, qhdcTong: null, hasData: false }
  };

  if (!state.extractedData || state.extractedData.length === 0) {
    return result;
  }

  const totals = {
    VT: { tt27: 0, tt28: 0, count: 0 },
    ML: { tt27: 0, tt28: 0, count: 0 },
    CDBR: { tt27: 0, tt28: 0, count: 0 },
    CNTT: { tt27: 0, tt28: 0, count: 0 },
    TD: { tt27: 0, tt28: 0, count: 0 },
    CD: { tt27: 0, tt28: 0, count: 0 },
    HT: { tt27: 0, tt28: 0, count: 0 },
    ATTT: { tt27: 0, tt28: 0, count: 0 },
    VI: { tt27: 0, tt28: 0, count: 0 }
  };

  // Quét qua các dòng chi tiết trong state.extractedData
  for (const item of state.extractedData) {
    if (item.isMangHeader || item.isGroup) continue;

    let targetKey = item.mangCode;
    if (!targetKey) continue;

    // Tách riêng nhánh CNTT thành VI, ATTT, CNTT
    if (targetKey === 'CNTT') {
      const ma = (item.maMang || '').toUpperCase();
      const nd = (item.nd || '').toUpperCase();
      if (ma.includes('VI') || ma.includes('VÍ') || nd.includes('VÍ') || nd.includes('VI DIEN TU')) {
        targetKey = 'VI';
      } else if (ma.includes('ATTT') || nd.includes('ATTT') || nd.includes('AN TOÀN THÔNG TIN')) {
        targetKey = 'ATTT';
      } else {
        targetKey = 'CNTT';
      }
    }

    if (totals[targetKey]) {
      const val27 = (typeof item.tt27 === 'number' && !isNaN(item.tt27)) ? item.tt27 : ((item.kl27 && item.dg) ? item.kl27 * item.dg : 0);
      const val28 = (typeof item.tt28 === 'number' && !isNaN(item.tt28)) ? item.tt28 : ((item.kl28 && item.dg) ? item.kl28 * item.dg : 0);
      totals[targetKey].count++;
      totals[targetKey].tt27 += val27;
      totals[targetKey].tt28 += val28;
    }
  }

  // Đổi từ USD sang Triệu USD (M$)
  for (const k of Object.keys(totals)) {
    const fileLoaded = (k === 'VI' || k === 'ATTT') ? !!state.files['CNTT'] : !!state.files[k];
    if (fileLoaded && totals[k].count > 0) {
      const q27 = Math.round((totals[k].tt27 / 1000000) * 100) / 100;
      const q28 = Math.round((totals[k].tt28 / 1000000) * 100) / 100;
      result[k] = {
        qhdc27: q27,
        qhdc28: q28,
        qhdcTong: Math.round((q27 + q28) * 100) / 100,
        hasData: true
      };
    }
  }

  return result;
}

// Tính toán toàn bộ các giá trị của bảng so sánh (QHĐC realtime từ Masterlist, không sửa được; Chiến lược nhập theo Profile)
function computeStrategyTableData(profile) {
  if (!profile || !profile.data) return null;
  const data = profile.data;
  const qhdcLive = getQHDCDataFromMasterlist();

  // 1. Tính toán cho 9 mảng nghiệp vụ
  const rowData = {};
  for (const k of STRATEGY_ITEM_KEYS) {
    const item = data[k] || { name: k, strat27: 0, strat28: 0 };
    const q = qhdcLive[k];

    const s27 = parseFloat(item.strat27) || 0;
    const s28 = parseFloat(item.strat28) || 0;
    const sTong = s27 + s28;

    let qTong = null, q27 = null, q28 = null;
    let dTong = null, d27 = null, d28 = null;

    if (q && q.hasData) {
      q27 = q.qhdc27;
      q28 = q.qhdc28;
      qTong = q.qhdcTong;
      dTong = qTong - sTong;
      d27 = q27 - s27;
      d28 = q28 - s28;
    }

    rowData[k] = {
      hasQHDC: !!(q && q.hasData),
      qhdcTong: qTong, qhdc27: q27, qhdc28: q28,
      stratTong: sTong, strat27: s27, strat28: s28,
      diffTong: dTong, diff27: d27, diff28: d28
    };
  }

  // 2. Dòng 1: Tổng đầu tư (không gồm ví) (m$)
  let kv_s27 = 0, kv_s28 = 0;
  let kv_q27 = 0, kv_q28 = 0, hasAnyKV = false;

  for (const k of STRATEGY_KHONG_VI_KEYS) {
    kv_s27 += rowData[k].strat27;
    kv_s28 += rowData[k].strat28;
    if (rowData[k].hasQHDC) {
      hasAnyKV = true;
      kv_q27 += rowData[k].qhdc27;
      kv_q28 += rowData[k].qhdc28;
    }
  }

  const kv_sTong = kv_s27 + kv_s28;
  let kv_qTong = null, kv_dTong = null, kv_d27 = null, kv_d28 = null;

  if (hasAnyKV) {
    kv_qTong = kv_q27 + kv_q28;
    kv_dTong = kv_qTong - kv_sTong;
    kv_d27 = kv_q27 - kv_s27;
    kv_d28 = kv_q28 - kv_s28;
  }

  rowData['KHONG_VI'] = {
    hasQHDC: hasAnyKV,
    qhdcTong: kv_qTong, qhdc27: hasAnyKV ? kv_q27 : null, qhdc28: hasAnyKV ? kv_q28 : null,
    stratTong: kv_sTong, strat27: kv_s27, strat28: kv_s28,
    diffTong: kv_dTong, diff27: kv_d27, diff28: kv_d28
  };

  // 3. Dòng 2: Tổng đầu tư (m$) = Không ví + Ví
  const vi = rowData['VI'];
  const t_s27 = kv_s27 + vi.strat27;
  const t_s28 = kv_s28 + vi.strat28;
  const t_sTong = t_s27 + t_s28;

  let t_q27 = null, t_q28 = null, t_qTong = null;
  let t_dTong = null, t_d27 = null, t_d28 = null;
  const hasAnyTotal = hasAnyKV || vi.hasQHDC;

  if (hasAnyTotal) {
    t_q27 = (hasAnyKV ? kv_q27 : 0) + (vi.hasQHDC ? vi.qhdc27 : 0);
    t_q28 = (hasAnyKV ? kv_q28 : 0) + (vi.hasQHDC ? vi.qhdc28 : 0);
    t_qTong = t_q27 + t_q28;
    t_dTong = t_qTong - t_sTong;
    t_d27 = t_q27 - t_s27;
    t_d28 = t_q28 - t_s28;
  }

  rowData['TONG'] = {
    hasQHDC: hasAnyTotal,
    qhdcTong: t_qTong, qhdc27: t_q27, qhdc28: t_q28,
    stratTong: t_sTong, strat27: t_s27, strat28: t_s28,
    diffTong: t_dTong, diff27: t_d27, diff28: t_d28
  };

  return rowData;
}

// Định dạng số hiển thị trong bảng (Chuẩn theo hình ảnh: Viettel Red #EE0033, số căn giữa, Chênh lệch có dấu + và -, không có dữ liệu để trống)
// Định dạng số hiển thị trong bảng (Chuẩn SaaS hiện đại, số căn giữa, Chênh lệch có tag xanh/đỏ rõ ràng)
function formatStratCell(val, isDiff = false) {
  if (val === null || val === undefined) return '';
  const num = typeof val === 'number' ? val : parseFloat(val);
  if (isNaN(num)) return '';
  if (Math.abs(num) < 0.0001) {
    if (isDiff) return '<span class="strat-diff-zero">0.00</span>';
    return '<span style="color: #94a3b8;">-</span>';
  }
  const formatted = Math.abs(num).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  if (isDiff) {
    if (num > 0) {
      return `<span class="strat-diff-pos">+${formatted}</span>`;
    } else {
      return `<span class="strat-diff-neg">-${formatted}</span>`;
    }
  }
  return num < 0 ? `-${formatted}` : formatted;
}

function formatStratRaw(val, isDiff = false) {
  if (val === null || val === undefined) return '';
  const num = typeof val === 'number' ? val : parseFloat(val);
  if (isNaN(num)) return '';
  if (Math.abs(num) < 0.0001) {
    if (isDiff) return '0.00';
    return '-';
  }
  const formatted = Math.abs(num).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  if (isDiff) {
    return num > 0 ? `+${formatted}` : `-${formatted}`;
  }
  return num < 0 ? `-${formatted}` : formatted;
}

// Render toàn bộ bảng so sánh chiến lược
function renderStrategyComparisonTable() {
  const tbody = document.getElementById('strategyTableBody');
  if (!tbody) return;

  const profile = getActiveStrategyProfile();
  if (!profile) {
    tbody.innerHTML = `<tr><td colspan="12" style="text-align:center; padding: 2.5rem; color: #94a3b8;">Chưa có dữ liệu Profile</td></tr>`;
    return;
  }

  const computed = computeStrategyTableData(profile);
  let html = '';
  const rows = getStrategyDisplayRows(profile);
  const order = getNormalizedStrategyRowOrder(profile);
  const totalRow = computed['TONG'] || {};

  for (const r of rows) {
    const c = computed[r.key];
    const itemData = profile.data[r.key] || {};
    const trClass = r.className ? `class="${r.className}"` : '';
    const qhdcShare = formatStratShare(c.qhdcTong, totalRow.qhdcTong);
    const stratShare = formatStratShare(c.stratTong, totalRow.stratTong);

    if (r.isTotal) {
      // Dòng Tổng 1 & 2: các ô đều tính tự động
      html += `
        <tr ${trClass}>
          <td class="strat-cell-mang-total">${escapeHtml(r.name)}</td>
          <td id="strat_cell_${r.key}_qhdcTong" class="strat-cell-qhdc" style="font-weight: 800;">${formatStratCell(c.qhdcTong)}</td>
          <td id="strat_cell_${r.key}_qhdcShare" class="strat-cell-share strat-cell-qhdc" style="font-weight: 800;">${qhdcShare}</td>
          <td id="strat_cell_${r.key}_qhdc27" class="strat-cell-qhdc" style="font-weight: 800;">${formatStratCell(c.qhdc27)}</td>
          <td id="strat_cell_${r.key}_qhdc28" class="strat-cell-qhdc strat-border-group-right" style="font-weight: 800;">${formatStratCell(c.qhdc28)}</td>
          <td id="strat_cell_${r.key}_stratTong" style="font-weight: 800; color: #5b21b6;">${formatStratCell(c.stratTong)}</td>
          <td id="strat_cell_${r.key}_stratShare" class="strat-cell-share" style="font-weight: 800; color: #5b21b6;">${stratShare}</td>
          <td id="strat_cell_${r.key}_strat27" style="font-weight: 800; color: #5b21b6;">${formatStratCell(c.strat27)}</td>
          <td id="strat_cell_${r.key}_strat28" class="strat-border-group-right" style="font-weight: 800; color: #5b21b6;">${formatStratCell(c.strat28)}</td>
          <td id="strat_cell_${r.key}_diffTong" style="font-weight: 800;">${formatStratCell(c.diffTong, true)}</td>
          <td id="strat_cell_${r.key}_diff27" style="font-weight: 800;">${formatStratCell(c.diff27, true)}</td>
          <td id="strat_cell_${r.key}_diff28" style="font-weight: 800;">${formatStratCell(c.diff28, true)}</td>
        </tr>
      `;
    } else {
      // Các dòng mảng chi tiết (VT, ML, CDBR, CNTT, TD, CD, HT, ATTT, VI):
      // QHĐC: Lấy thời gian thực từ Masterlist, CỐ ĐỊNH KHÔNG CHO SỬA (chỉ hiển thị text)
      // Chiến lược: Cho phép nhập/sửa theo Profile kịch bản
      const valS27 = (itemData.strat27 !== undefined && itemData.strat27 !== null) ? itemData.strat27 : '';
      const valS28 = (itemData.strat28 !== undefined && itemData.strat28 !== null) ? itemData.strat28 : '';
      const rowIndex = order.indexOf(r.key);
      const canMoveUp = rowIndex > 0;
      const canMoveDown = rowIndex >= 0 && rowIndex < order.length - 1;

      html += `
        <tr ${trClass}>
          <td class="strat-cell-mang">
            <div class="strat-row-label-wrap">
              <span>${escapeHtml(r.name)}</span>
              <span class="strat-row-actions" aria-label="Điều chỉnh thứ tự dòng">
                <button type="button" class="strat-row-move-btn" data-key="${r.key}" data-strategy-move="up" ${canMoveUp ? '' : 'disabled'} title="Đưa dòng lên">
                  <i data-lucide="chevron-up" class="w-3.5 h-3.5"></i>
                </button>
                <button type="button" class="strat-row-move-btn" data-key="${r.key}" data-strategy-move="down" ${canMoveDown ? '' : 'disabled'} title="Đưa dòng xuống">
                  <i data-lucide="chevron-down" class="w-3.5 h-3.5"></i>
                </button>
              </span>
            </div>
          </td>
          <td id="strat_cell_${r.key}_qhdcTong" class="strat-cell-qhdc" style="font-weight: 700;">${formatStratCell(c.qhdcTong)}</td>
          <td id="strat_cell_${r.key}_qhdcShare" class="strat-cell-share strat-cell-qhdc">${qhdcShare}</td>
          <td id="strat_cell_${r.key}_qhdc27" class="strat-cell-qhdc">${formatStratCell(c.qhdc27)}</td>
          <td id="strat_cell_${r.key}_qhdc28" class="strat-cell-qhdc strat-border-group-right">${formatStratCell(c.qhdc28)}</td>
          <td id="strat_cell_${r.key}_stratTong" class="strat-cell-strat" style="font-weight: 700;">${formatStratCell(c.stratTong)}</td>
          <td id="strat_cell_${r.key}_stratShare" class="strat-cell-share strat-cell-strat">${stratShare}</td>
          <td id="strat_cell_${r.key}_strat27" class="strat-cell-strat">${formatStratCell(c.strat27)}</td>
          <td id="strat_cell_${r.key}_strat28" class="strat-cell-strat strat-border-group-right">${formatStratCell(c.strat28)}</td>
          <td id="strat_cell_${r.key}_diffTong">${formatStratCell(c.diffTong, true)}</td>
          <td id="strat_cell_${r.key}_diff27">${formatStratCell(c.diff27, true)}</td>
          <td id="strat_cell_${r.key}_diff28">${formatStratCell(c.diff28, true)}</td>
        </tr>
      `;
    }
  }

  tbody.innerHTML = html;
  if (window.lucide) lucide.createIcons();
}

// Cập nhật nội dung các ô tính toán trong DOM
function updateCalculatedCells() {
  const profile = getActiveStrategyProfile();
  if (!profile) return;
  const computed = computeStrategyTableData(profile);
  const totalRow = computed['TONG'] || {};

  for (const r of STRATEGY_ROWS) {
    const c = computed[r.key];
    const setCell = (id, html) => {
      const el = document.getElementById(id);
      if (el) el.innerHTML = html;
    };

    setCell(`strat_cell_${r.key}_qhdcTong`, formatStratCell(c.qhdcTong));
    setCell(`strat_cell_${r.key}_qhdcShare`, formatStratShare(c.qhdcTong, totalRow.qhdcTong));
    setCell(`strat_cell_${r.key}_qhdc27`, formatStratCell(c.qhdc27));
    setCell(`strat_cell_${r.key}_qhdc28`, formatStratCell(c.qhdc28));
    setCell(`strat_cell_${r.key}_stratTong`, formatStratCell(c.stratTong));
    setCell(`strat_cell_${r.key}_stratShare`, formatStratShare(c.stratTong, totalRow.stratTong));
    setCell(`strat_cell_${r.key}_diffTong`, formatStratCell(c.diffTong, true));
    setCell(`strat_cell_${r.key}_diff27`, formatStratCell(c.diff27, true));
    setCell(`strat_cell_${r.key}_diff28`, formatStratCell(c.diff28, true));

    if (r.isTotal) {
      setCell(`strat_cell_${r.key}_strat27`, formatStratCell(c.strat27));
      setCell(`strat_cell_${r.key}_strat28`, formatStratCell(c.strat28));
    }
  }
}

// Sao chép bảng so sánh sang Clipboard (Excel định dạng thanh lịch, hiện đại)
async function copyStrategyTableToClipboard() {
  const profile = getActiveStrategyProfile();
  if (!profile) return;

  const computed = computeStrategyTableData(profile);
  const rows = getStrategyDisplayRows(profile);
  const totalRow = computed['TONG'] || {};

  let html = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>SoSanhChienLuoc5Nam</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->
<style>
  table { border-collapse: collapse; font-family: Calibri, 'Segoe UI', Arial, sans-serif; font-size: 10pt; }
  th { border: 1px solid #cbd5e1; padding: 6px 10px; text-align: center; font-weight: bold; }
  td { border: 1px solid #cbd5e1; padding: 5px 8px; text-align: center; }
  .th-mang { background-color: #f1f5f9; color: #0f172a; text-align: left; padding-left: 10px; }
  .th-qhdc { background-color: #dbeafe; color: #1e40af; }
  .th-strat { background-color: #ede9fe; color: #5b21b6; }
  .th-diff { background-color: #f1f5f9; color: #334155; }
  .row-total { background-color: #f8fafc; font-weight: bold; }
</style>
</head>
<body>
<table>
<thead>
  <tr>
    <th rowspan="2" class="th-mang">Mảng nghiệp vụ</th>
    <th colspan="4" class="th-qhdc" style="border-right: 2px solid #94a3b8;">QHĐC 2027-2028 (M$)</th>
    <th colspan="4" class="th-strat" style="border-right: 2px solid #94a3b8;">CHIẾN LƯỢC 5 NĂM (M$)</th>
    <th colspan="3" class="th-diff">CHÊNH LỆCH SO VỚI CHIẾN LƯỢC (M$)</th>
  </tr>
  <tr>
    <th class="th-qhdc">Tổng</th>
    <th class="th-qhdc">Tỉ trọng</th>
    <th class="th-qhdc">2027</th>
    <th class="th-qhdc" style="border-right: 2px solid #94a3b8;">2028</th>
    <th class="th-strat">Tổng</th>
    <th class="th-strat">Tỉ trọng</th>
    <th class="th-strat">2027</th>
    <th class="th-strat" style="border-right: 2px solid #94a3b8;">2028</th>
    <th class="th-diff">Tổng</th>
    <th class="th-diff">2027</th>
    <th class="th-diff">2028</th>
  </tr>
</thead>
<tbody>`;

  let tsv = "Mảng\tQHĐC Tổng\tQHĐC Tỉ trọng\tQHĐC 2027\tQHĐC 2028\tChiến lược Tổng\tChiến lược Tỉ trọng\tChiến lược 2027\tChiến lược 2028\tChênh lệch Tổng\tChênh lệch 2027\tChênh lệch 2028\r\n";

  for (const r of rows) {
    const c = computed[r.key];
    const bold = r.isTotal ? 'font-weight: bold;' : '';
    const rowBg = r.isTotal ? 'background-color: #f8fafc;' : 'background-color: #ffffff;';
    const mangStyle = r.isTotal
      ? `background-color: #f1f5f9; color: #0f172a; font-weight: bold; text-align: left; padding: 6px 10px; border: 1px solid #cbd5e1;`
      : `background-color: #ffffff; color: #1e293b; font-weight: bold; text-align: left; padding-left: 12px; border: 1px solid #cbd5e1; padding: 5px 8px;`;

    const tdBase = `${rowBg} color: #0f172a; ${bold} border: 1px solid #cbd5e1; text-align: center; padding: 5px 8px; mso-number-format:'\\#\\,\\#\\#0\\.00';`;
    const tdDivider = `${rowBg} color: #0f172a; ${bold} border: 1px solid #cbd5e1; border-right: 2px solid #94a3b8; text-align: center; padding: 5px 8px; mso-number-format:'\\#\\,\\#\\#0\\.00';`;

    const diffTongColor = (c.diffTong && c.diffTong > 0) ? 'color: #047857;' : ((c.diffTong && c.diffTong < 0) ? 'color: #be123c;' : '');
    const diff27Color = (c.diff27 && c.diff27 > 0) ? 'color: #047857;' : ((c.diff27 && c.diff27 < 0) ? 'color: #be123c;' : '');
    const diff28Color = (c.diff28 && c.diff28 > 0) ? 'color: #047857;' : ((c.diff28 && c.diff28 < 0) ? 'color: #be123c;' : '');

    html += `
      <tr>
        <td style="${mangStyle}">${escapeHtml(r.name)}</td>
        <td style="${tdBase} color: #0369a1;">${formatStratRaw(c.qhdcTong)}</td>
        <td style="${tdBase} color: #0369a1;">${formatStratShareRaw(c.qhdcTong, totalRow.qhdcTong)}</td>
        <td style="${tdBase} color: #0369a1;">${formatStratRaw(c.qhdc27)}</td>
        <td style="${tdDivider} color: #0369a1;">${formatStratRaw(c.qhdc28)}</td>
        <td style="${tdBase} color: #5b21b6;">${formatStratRaw(c.stratTong)}</td>
        <td style="${tdBase} color: #5b21b6;">${formatStratShareRaw(c.stratTong, totalRow.stratTong)}</td>
        <td style="${tdBase} color: #5b21b6;">${formatStratRaw(c.strat27)}</td>
        <td style="${tdDivider} color: #5b21b6;">${formatStratRaw(c.strat28)}</td>
        <td style="${tdBase} ${diffTongColor}">${formatStratRaw(c.diffTong, true)}</td>
        <td style="${tdBase} ${diff27Color}">${formatStratRaw(c.diff27, true)}</td>
        <td style="${tdBase} ${diff28Color}">${formatStratRaw(c.diff28, true)}</td>
      </tr>
    `;

    tsv += `${r.name}\t${formatStratRaw(c.qhdcTong)}\t${formatStratShareRaw(c.qhdcTong, totalRow.qhdcTong)}\t${formatStratRaw(c.qhdc27)}\t${formatStratRaw(c.qhdc28)}\t${formatStratRaw(c.stratTong)}\t${formatStratShareRaw(c.stratTong, totalRow.stratTong)}\t${formatStratRaw(c.strat27)}\t${formatStratRaw(c.strat28)}\t${formatStratRaw(c.diffTong, true)}\t${formatStratRaw(c.diff27, true)}\t${formatStratRaw(c.diff28, true)}\r\n`;
  }

  html += `</tbody></table></body></html>`;

  function triggerStratCopySuccess() {
    const textEl = document.getElementById('copyStratBtnText');
    const iconEl = document.getElementById('copyStratBtnIcon');
    if (textEl) {
      const old = textEl.textContent;
      textEl.textContent = 'Đã sao chép!';
      setTimeout(() => { textEl.textContent = old; }, 2000);
    }
    if (iconEl) {
      iconEl.innerHTML = '<i data-lucide="check" class="w-4 h-4 text-emerald-600"></i>';
      if (window.lucide) lucide.createIcons();
      setTimeout(() => {
        iconEl.innerHTML = '<i data-lucide="copy" class="w-4 h-4 text-indigo-600"></i>';
        if (window.lucide) lucide.createIcons();
      }, 2000);
    }
    showToast('Đã sao chép bảng so sánh! Bạn có thể dán (Ctrl+V) trực tiếp vào Excel.', 'success');
  }

  try {
    if (navigator.clipboard && window.ClipboardItem) {
      const blobHtml = new Blob([html], { type: 'text/html' });
      const blobText = new Blob([tsv], { type: 'text/plain' });
      await navigator.clipboard.write([
        new ClipboardItem({
          'text/html': blobHtml,
          'text/plain': blobText
        })
      ]);
      triggerStratCopySuccess();
    } else if (navigator.clipboard) {
      await navigator.clipboard.writeText(tsv);
      triggerStratCopySuccess();
    }
  } catch (err) {
    console.error('Lỗi sao chép:', err);
    showToast('Lỗi khi sao chép: ' + err.message, 'error');
  }
}

// ==================== QUẢN LÝ PROFILE CHIẾN LƯỢC ====================

// Mở modal tạo mới hoặc chỉnh sửa Profile
function openProfileModal(mode = 'create') {
  const modal = document.getElementById('modalNewProfile');
  const inputName = document.getElementById('inputNewProfileName');
  const inputDesc = document.getElementById('inputNewProfileDesc');
  const titleEl = document.getElementById('modalProfileTitle');
  const iconEl = document.getElementById('modalProfileHeaderIcon');
  const saveTextEl = document.getElementById('btnModalSaveText');
  const modeEl = document.getElementById('modalProfileMode');
  const editIdEl = document.getElementById('modalProfileEditId');

  if (!modal) return;

  if (modeEl) modeEl.value = mode;

  let initialData = {};

  if (mode === 'create') {
    if (titleEl) titleEl.textContent = 'Thêm Profile Chiến lược Mới';
    if (saveTextEl) saveTextEl.textContent = 'Tạo Profile';
    if (iconEl) {
      iconEl.setAttribute('data-lucide', 'folder-plus');
      iconEl.className = 'w-5 h-5 text-rose-600';
    }
    if (inputName) inputName.value = `Chiến lược 5 năm (Phương án ${state.strategyProfiles.length + 1})`;
    if (inputDesc) inputDesc.value = 'Kịch bản phân bổ vốn đầu tư chiến lược 5 năm';
    if (editIdEl) editIdEl.value = '';

    // Khởi tạo số liệu từ profile đang chọn hoặc mặc định
    const active = getActiveStrategyProfile();
    initialData = active && active.data ? JSON.parse(JSON.stringify(active.data)) : JSON.parse(JSON.stringify(DEFAULT_STRATEGY_DATA));
  } else {
    // Mode edit
    const active = getActiveStrategyProfile();
    if (!active) {
      showToast('Không tìm thấy Profile để chỉnh sửa!', 'warning');
      return;
    }
    if (titleEl) titleEl.textContent = `Chỉnh sửa Profile: ${active.name}`;
    if (saveTextEl) saveTextEl.textContent = 'Cập nhật Profile';
    if (iconEl) {
      iconEl.setAttribute('data-lucide', 'edit-3');
      iconEl.className = 'w-5 h-5 text-indigo-600';
    }
    if (inputName) inputName.value = active.name;
    if (inputDesc) inputDesc.value = active.description || '';
    if (editIdEl) editIdEl.value = active.id;
    initialData = active.data ? JSON.parse(JSON.stringify(active.data)) : JSON.parse(JSON.stringify(DEFAULT_STRATEGY_DATA));
  }

  // Render bảng nhập liệu trong modal
  renderModalProfileTable(initialData);

  modal.style.display = 'flex';
  if (window.lucide) lucide.createIcons();
}

function closeProfileModal() {
  const modal = document.getElementById('modalNewProfile');
  if (modal) modal.style.display = 'none';
}

// Render các dòng nhập liệu trong Modal
function renderModalProfileTable(data) {
  const tbody = document.getElementById('modalProfileTableBody');
  if (!tbody) return;

  const itemRows = STRATEGY_ROWS.filter(r => !r.isTotal);
  let html = '';

  for (const r of itemRows) {
    const item = data[r.key] || { strat27: 0, strat28: 0 };
    const s27 = (item.strat27 !== null && item.strat27 !== undefined) ? item.strat27 : 0;
    const s28 = (item.strat28 !== null && item.strat28 !== undefined) ? item.strat28 : 0;
    const sTong = Math.round((parseFloat(s27) + parseFloat(s28)) * 100) / 100;

    html += `
      <tr>
        <td style="padding: 0.5rem 0.85rem; font-weight: 600; color: #1e293b; text-align: left;">
          ${escapeHtml(r.name)}
        </td>
        <td style="text-align: center; padding: 0.3rem 0.45rem;">
          <input type="number" step="0.01" class="modal-profile-input" data-key="${r.key}" data-field="strat27" value="${s27}" title="Chiến lược 2027 (M$)">
        </td>
        <td style="text-align: center; padding: 0.3rem 0.45rem;">
          <input type="number" step="0.01" class="modal-profile-input" data-key="${r.key}" data-field="strat28" value="${s28}" title="Chiến lược 2028 (M$)">
        </td>
        <td style="text-align: center; font-family: 'JetBrains Mono', monospace; font-weight: 700; color: #5b21b6; padding: 0.45rem 0.6rem;">
          <span id="modal_row_sum_${r.key}">${formatStratRaw(sTong)}</span>
        </td>
      </tr>
    `;
  }

  // 2 Dòng tổng tính toán tự động
  html += `
    <tr style="background: #f8fafc; font-weight: 800; border-top: 2px solid #cbd5e1;">
      <td style="padding: 0.5rem 0.75rem; color: #0f172a;">Tổng đầu tư (không gồm ví)</td>
      <td id="modal_total_kv_27" style="text-align: center; font-family: 'JetBrains Mono', monospace; color: #5b21b6; padding: 0.5rem 0.45rem;">0.00</td>
      <td id="modal_total_kv_28" style="text-align: center; font-family: 'JetBrains Mono', monospace; color: #5b21b6; padding: 0.5rem 0.45rem;">0.00</td>
      <td id="modal_total_kv_tong" style="text-align: center; font-family: 'JetBrains Mono', monospace; color: #5b21b6; padding: 0.5rem 0.6rem;">0.00</td>
    </tr>
    <tr style="background: #f1f5f9; font-weight: 800; border-top: 1px solid #cbd5e1; border-bottom: 2px solid #94a3b8;">
      <td style="padding: 0.5rem 0.75rem; color: #0f172a;">Tổng đầu tư (m$)</td>
      <td id="modal_total_all_27" style="text-align: center; font-family: 'JetBrains Mono', monospace; color: #5b21b6; padding: 0.5rem 0.45rem;">0.00</td>
      <td id="modal_total_all_28" style="text-align: center; font-family: 'JetBrains Mono', monospace; color: #5b21b6; padding: 0.5rem 0.45rem;">0.00</td>
      <td id="modal_total_all_tong" style="text-align: center; font-family: 'JetBrains Mono', monospace; color: #5b21b6; padding: 0.5rem 0.6rem;">0.00</td>
    </tr>
  `;

  tbody.innerHTML = html;

  // Gắn sự kiện tính toán tức thì khi gõ số
  tbody.querySelectorAll('.modal-profile-input').forEach(input => {
    input.addEventListener('input', updateModalProfileCalculations);
  });

  // Tính toán lần đầu
  updateModalProfileCalculations();
}

// Cập nhật tính toán tổng các dòng và tổng đầu tư trong modal
function updateModalProfileCalculations() {
  const tbody = document.getElementById('modalProfileTableBody');
  if (!tbody) return;

  const itemRows = STRATEGY_ROWS.filter(r => !r.isTotal);
  let kv27 = 0, kv28 = 0;
  let all27 = 0, all28 = 0;

  for (const r of itemRows) {
    const input27 = tbody.querySelector(`.modal-profile-input[data-key="${r.key}"][data-field="strat27"]`);
    const input28 = tbody.querySelector(`.modal-profile-input[data-key="${r.key}"][data-field="strat28"]`);
    const val27 = input27 ? (parseFloat(input27.value) || 0) : 0;
    const val28 = input28 ? (parseFloat(input28.value) || 0) : 0;
    const rowSum = Math.round((val27 + val28) * 100) / 100;

    const sumEl = document.getElementById(`modal_row_sum_${r.key}`);
    if (sumEl) sumEl.textContent = formatStratRaw(rowSum);

    if (r.key !== 'VI') {
      kv27 += val27;
      kv28 += val28;
    }
    all27 += val27;
    all28 += val28;
  }

  kv27 = Math.round(kv27 * 100) / 100;
  kv28 = Math.round(kv28 * 100) / 100;
  const kvTong = Math.round((kv27 + kv28) * 100) / 100;

  all27 = Math.round(all27 * 100) / 100;
  all28 = Math.round(all28 * 100) / 100;
  const allTong = Math.round((all27 + all28) * 100) / 100;

  const setEl = (id, val) => {
    const el = document.getElementById(id);
    if (el) el.textContent = formatStratRaw(val);
  };

  setEl('modal_total_kv_27', kv27);
  setEl('modal_total_kv_28', kv28);
  setEl('modal_total_kv_tong', kvTong);

  setEl('modal_total_all_27', all27);
  setEl('modal_total_all_28', all28);
  setEl('modal_total_all_tong', allTong);
}

// Điền nhanh số liệu vào Modal
function fillModalData(mode) {
  let source = {};
  if (mode === 'current') {
    const active = getActiveStrategyProfile();
    source = active && active.data ? active.data : DEFAULT_STRATEGY_DATA;
  } else if (mode === 'default') {
    source = DEFAULT_STRATEGY_DATA;
  } else {
    // zero
    source = {};
  }

  const tbody = document.getElementById('modalProfileTableBody');
  if (!tbody) return;

  const itemRows = STRATEGY_ROWS.filter(r => !r.isTotal);
  for (const r of itemRows) {
    const item = source[r.key] || { strat27: 0, strat28: 0 };
    const input27 = tbody.querySelector(`.modal-profile-input[data-key="${r.key}"][data-field="strat27"]`);
    const input28 = tbody.querySelector(`.modal-profile-input[data-key="${r.key}"][data-field="strat28"]`);
    if (input27) input27.value = (mode === 'zero') ? 0 : (item.strat27 || 0);
    if (input28) input28.value = (mode === 'zero') ? 0 : (item.strat28 || 0);
  }

  updateModalProfileCalculations();
}

// Xác nhận lưu Profile từ Modal (hỗ trợ cả Tạo mới và Chỉnh sửa)
async function confirmSaveProfileModal() {
  const inputName = document.getElementById('inputNewProfileName');
  const inputDesc = document.getElementById('inputNewProfileDesc');
  const modeEl = document.getElementById('modalProfileMode');
  const mode = modeEl ? (modeEl.value || 'create') : 'create';
  const editIdEl = document.getElementById('modalProfileEditId');
  const editId = editIdEl ? editIdEl.value : '';
  const tbody = document.getElementById('modalProfileTableBody');
  const btnSave = document.getElementById('btnConfirmSaveProfile');

  const name = inputName ? inputName.value.trim() : '';
  if (!name) {
    showToast('Vui lòng nhập tên Profile!', 'warning');
    if (inputName) inputName.focus();
    return;
  }

  const desc = inputDesc ? inputDesc.value.trim() : '';

  // Thu thập số liệu 9 mảng từ các ô input trong Modal
  const newData = {};
  const itemRows = STRATEGY_ROWS.filter(r => !r.isTotal);

  for (const r of itemRows) {
    const input27 = tbody ? tbody.querySelector(`.modal-profile-input[data-key="${r.key}"][data-field="strat27"]`) : null;
    const input28 = tbody ? tbody.querySelector(`.modal-profile-input[data-key="${r.key}"][data-field="strat28"]`) : null;
    const val27 = input27 ? (parseFloat(input27.value) || 0) : 0;
    const val28 = input28 ? (parseFloat(input28.value) || 0) : 0;

    newData[r.key] = {
      name: r.name,
      strat27: val27,
      strat28: val28
    };
  }

  if (btnSave) {
    btnSave.disabled = true;
    btnSave.innerHTML = '<span class="spinner"></span> Đang lưu...';
  }

  try {
    if (mode === 'create') {
      const activeForOrder = getActiveStrategyProfile();
      const newProfile = {
        id: 'profile_' + Date.now(),
        name: name,
        description: desc,
        updatedAt: new Date().toISOString(),
        data: newData,
        rowOrder: getNormalizedStrategyRowOrder(activeForOrder)
      };
      state.strategyProfiles.push(newProfile);
      state.activeStrategyProfileId = newProfile.id;
    } else {
      const existing = state.strategyProfiles.find(p => p.id === editId);
      if (existing) {
        existing.name = name;
        existing.description = desc;
        existing.updatedAt = new Date().toISOString();
        existing.data = newData;
        existing.rowOrder = getNormalizedStrategyRowOrder(existing);
      }
    }

    const saved = await saveStrategyProfiles();
    renderStrategyProfileSelect();
    renderStrategyComparisonTable();
    closeProfileModal();

    if (saved) {
      showToast(mode === 'create' ? `Đã tạo Profile [${name}] thành công!` : `Đã cập nhật Profile [${name}] thành công!`, 'success');
    } else {
      showToast(`⚠️ Không thể lưu Profile vào bộ nhớ trình duyệt!`, 'error');
    }
  } catch (err) {
    console.error('Lỗi khi lưu profile:', err);
    showToast(`Lỗi khi lưu profile: ${err.message}`, 'error');
  } finally {
    if (btnSave) {
      btnSave.disabled = false;
      btnSave.innerHTML = '<i data-lucide="check" class="w-4 h-4"></i> <span id="modalSaveProfileText">' + (mode === 'create' ? 'Lưu Profile' : 'Cập nhật Profile') + '</span>';
      if (window.lucide) lucide.createIcons();
    }
  }
}

async function cloneCurrentProfile() {
  const active = getActiveStrategyProfile();
  if (!active) return;

  const cloned = {
    id: 'profile_' + Date.now(),
    name: active.name + ' (Bản sao)',
    description: active.description,
    updatedAt: new Date().toISOString(),
    data: JSON.parse(JSON.stringify(active.data)),
    rowOrder: getNormalizedStrategyRowOrder(active)
  };

  state.strategyProfiles.push(cloned);
  state.activeStrategyProfileId = cloned.id;

  const saved = await saveStrategyProfiles();
  renderStrategyProfileSelect();
  renderStrategyComparisonTable();
  if (saved) {
    showToast(`Đã nhân bản Profile [${cloned.name}] thành công!`, 'success');
  } else {
    showToast(`⚠️ Không thể lưu Profile nhân bản!`, 'error');
  }
}

async function deleteCurrentProfile() {
  if (state.strategyProfiles.length <= 1) {
    showToast('Không thể xóa Profile duy nhất còn lại!', 'warning');
    return;
  }

  const active = getActiveStrategyProfile();
  if (!active) return;

  if (confirm(`Bạn có chắc chắn muốn xóa Profile [${active.name}]?`)) {
    state.strategyProfiles = state.strategyProfiles.filter(p => p.id !== active.id);
    state.activeStrategyProfileId = state.strategyProfiles[0].id;
    const saved = await saveStrategyProfiles();
    renderStrategyProfileSelect();
    renderStrategyComparisonTable();
    if (saved) {
      showToast(`Đã xóa Profile [${active.name}] thành công!`, 'info');
    } else {
      showToast(`⚠️ Không thể cập nhật trạng thái xóa!`, 'error');
    }
  }
}


window.closeNewProfileModal = closeProfileModal;
window.closeProfileModal = closeProfileModal;
window.openProfileModal = openProfileModal;
window.openNewProfileModal = openProfileModal;

// ==================== TÍCH HỢP CÔNG CỤ CHUYỂN ĐỔI MẪU CŨ VTB ====================

function initVTBConverter() {
  const btnOpen = document.getElementById('btnOpenVTBConverter');
  if (btnOpen) {
    btnOpen.addEventListener('click', openVTBConverterModal);
  }

  const inputOld = document.getElementById('inputVTBFileOld');
  if (inputOld) {
    inputOld.addEventListener('change', async (e) => {
      if (e.target.files && e.target.files[0]) {
        await handleVTBFileFromInput(e.target.files[0]);
      }
    });
  }

  const dropZone = document.getElementById('vtbDropZone');
  if (dropZone) {
    dropZone.addEventListener('dragover', (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropZone.classList.add('drag-over');
    });
    dropZone.addEventListener('dragleave', (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropZone.classList.remove('drag-over');
    });
    dropZone.addEventListener('drop', async (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropZone.classList.remove('drag-over');
      if (e.dataTransfer.files && e.dataTransfer.files[0]) {
        await handleVTBFileFromInput(e.dataTransfer.files[0]);
      }
    });
  }

  const btnDownload = document.getElementById('btnDownloadConvertedVTB');
  if (btnDownload) {
    btnDownload.addEventListener('click', downloadConvertedVTBFile);
  }

  const btnApply = document.getElementById('btnApplyConvertedToVT');
  if (btnApply) {
    btnApply.addEventListener('click', applyConvertedVTBToAppState);
  }

  const btnConvert = document.getElementById('btnConvertVTB');
  if (btnConvert) btnConvert.addEventListener('click', performVTBConversion);

  const selectAllButton = document.getElementById('btnVTBSelectAll');
  if (selectAllButton) selectAllButton.addEventListener('click', () => selectVTBMappingRows('all'));
  const selectUnassignedButton = document.getElementById('btnVTBSelectUnassigned');
  if (selectUnassignedButton) selectUnassignedButton.addEventListener('click', () => selectVTBMappingRows('unassigned'));
  const clearSelectionButton = document.getElementById('btnVTBClearSelection');
  if (clearSelectionButton) clearSelectionButton.addEventListener('click', () => selectVTBMappingRows('none'));
  const selectAllCheckbox = document.getElementById('vtbSelectAllCheckbox');
  if (selectAllCheckbox) {
    selectAllCheckbox.addEventListener('change', (event) => selectVTBMappingRows(event.target.checked ? 'all' : 'none'));
  }
  document.querySelectorAll('[data-vtb-technology]').forEach(button => {
    button.addEventListener('click', () => applyVTBTechnologyToSelection(button.dataset.vtbTechnology));
  });
  initVTBDragSelection();
}

function openVTBConverterModal() {
  const modal = document.getElementById('vtbConverterModal');
  if (modal) {
    modal.style.display = 'flex';
    if (window.lucide) lucide.createIcons();
  }
}

function closeVTBConverterModal() {
  const modal = document.getElementById('vtbConverterModal');
  if (modal) {
    modal.style.display = 'none';
  }
}

window.openVTBConverterModal = openVTBConverterModal;
window.closeVTBConverterModal = closeVTBConverterModal;

async function handleVTBFileFromInput(file) {
  try {
    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array', cellFormula: true, cellStyles: true });
    await handleVTBFileObject(file, workbook);
  } catch (err) {
    console.error('Lỗi khi đọc file VTB:', err);
    showToast(`Không thể đọc file: ${err.message}`, 'error');
  }
}

async function handleVTBFileObject(file, workbook) {
  if (!window.VTBConverter) {
    showToast('Chưa nạp module VTBConverter!', 'error');
    return;
  }

  const isOld = window.VTBConverter.isOldVTBFormat(workbook);
  if (!isOld) {
    showToast('File này không thuộc định dạng Mẫu cũ 37 cột của Vô tuyến!', 'warning');
  }

  // Cập nhật tên file trên giao diện
  const infoBar = document.getElementById('vtbSelectedFileInfo');
  const nameSpan = document.getElementById('vtbSelectedFileName');
  if (infoBar && nameSpan) {
    nameSpan.textContent = `${file.name} (${(file.size / 1024).toFixed(1)} KB)`;
    infoBar.style.display = 'inline-block';
  }

  try {
    const parsedData = window.VTBConverter.parseOldVTBWorkbook(workbook);
    const mappingRows = window.VTBConverter.buildVTBMappingRows(parsedData.items);
    const initialSplit = window.VTBConverter.splitOldVTBItems(parsedData.items, {});
    state.vtbConversion = {
      originalFile: file,
      originalFileName: file.name,
      oldWorkbook: workbook,
      parsedData,
      mappingRows,
      assignments: {},
      selectedIds: new Set(),
      unmappedAmounts: initialSplit.unmappedAmounts,
      convertedBuffer: null,
      convertedWorkbook: null,
      result: null
    };
    resetVTBConversionOutput();
    renderVTBAssignmentTable();
    showToast(`Đã đọc ${parsedData.items.length} vật tư. Hãy gán công nghệ trước khi chuyển đổi.`, 'success');
  } catch (err) {
    console.error('Lỗi khi phân tích file VTB:', err);
    showToast(`Lỗi khi phân tích file: ${err.message}`, 'error');
  }
}

function resetVTBConversionOutput() {
  const kpiSection = document.getElementById('vtbKPISection');
  if (kpiSection) kpiSection.style.display = 'none';
  const btnDownload = document.getElementById('btnDownloadConvertedVTB');
  const btnApply = document.getElementById('btnApplyConvertedToVT');
  if (btnDownload) btnDownload.disabled = true;
  if (btnApply) btnApply.disabled = true;
}

function renderVTBAssignmentTable() {
  const conversion = state.vtbConversion;
  const section = document.getElementById('vtbAssignmentSection');
  const tbody = document.getElementById('vtbAssignmentTableBody');
  if (!conversion || !section || !tbody) return;

  section.hidden = false;
  tbody.innerHTML = conversion.mappingRows.map(item => {
    const technology = conversion.assignments[item.id] || '';
    const suggestionLabel = item.suggestion ? `Gợi ý ${item.suggestion}` : 'Chọn';
    return `
      <tr data-vtb-mapping-id="${item.id}" class="${technology ? '' : 'is-unassigned'}">
        <td class="vtb-check-column"><input type="checkbox" class="vtb-row-checkbox" data-vtb-checkbox="${item.id}" aria-label="Chọn dòng ${item.row}"></td>
        <td style="text-align:center;font-weight:700;color:#64748b;">${item.row}</td>
        <td class="vtb-assignment-name" title="${escapeHtml(item.name)}">
          ${escapeHtml(item.name)}
          ${item.vendor ? `<span class="vtb-assignment-vendor">${escapeHtml(item.vendor)}</span>` : ''}
        </td>
        <td class="vtb-quantity-pair">${formatNumberVTB(item.coverage27)} / ${formatNumberVTB(item.coverage28)}</td>
        <td class="vtb-quantity-pair">${formatNumberVTB(item.capacity27)} / ${formatNumberVTB(item.capacity28)}</td>
        <td>
          <select class="vtb-tech-select ${technology ? '' : 'is-unassigned'}" data-vtb-select="${item.id}" aria-label="Công nghệ cho dòng ${item.row}">
            <option value="">${suggestionLabel}</option>
            ${['5G', '4G', '3G', '2G'].map(value => `<option value="${value}" ${technology === value ? 'selected' : ''}>${value}</option>`).join('')}
          </select>
        </td>
      </tr>
    `;
  }).join('');

  tbody.querySelectorAll('[data-vtb-select]').forEach(select => {
    select.addEventListener('change', event => {
      const id = event.target.dataset.vtbSelect;
      if (event.target.value) conversion.assignments[id] = event.target.value;
      else delete conversion.assignments[id];
      updateVTBAssignmentTableState();
    });
  });
  tbody.querySelectorAll('[data-vtb-checkbox]').forEach(checkbox => {
    checkbox.addEventListener('change', event => {
      setVTBMappingRowSelected(event.target.dataset.vtbCheckbox, event.target.checked);
    });
  });
  updateVTBAssignmentTableState();
}

function setVTBMappingRowSelected(id, selected) {
  const conversion = state.vtbConversion;
  if (!conversion) return;
  if (selected) conversion.selectedIds.add(String(id));
  else conversion.selectedIds.delete(String(id));
  updateVTBSelectionVisuals();
}

function updateVTBSelectionVisuals() {
  const conversion = state.vtbConversion;
  if (!conversion) return;
  document.querySelectorAll('[data-vtb-mapping-id]').forEach(row => {
    const selected = conversion.selectedIds.has(row.dataset.vtbMappingId);
    row.classList.toggle('is-selected', selected);
    const checkbox = row.querySelector('[data-vtb-checkbox]');
    if (checkbox) checkbox.checked = selected;
  });
  const count = conversion.selectedIds.size;
  const countLabel = document.getElementById('vtbSelectedCount');
  if (countLabel) countLabel.textContent = `${count} dòng đã chọn`;
  document.querySelectorAll('[data-vtb-technology]').forEach(button => {
    button.disabled = count === 0;
  });
  const allCheckbox = document.getElementById('vtbSelectAllCheckbox');
  if (allCheckbox) {
    const total = conversion.mappingRows.length;
    allCheckbox.checked = total > 0 && count === total;
    allCheckbox.indeterminate = count > 0 && count < total;
  }
}

function updateVTBAssignmentTableState() {
  const conversion = state.vtbConversion;
  if (!conversion) return;
  document.querySelectorAll('[data-vtb-mapping-id]').forEach(row => {
    const id = row.dataset.vtbMappingId;
    const technology = conversion.assignments[id] || '';
    row.classList.toggle('is-unassigned', !technology);
    const select = row.querySelector('[data-vtb-select]');
    if (select) {
      select.value = technology;
      select.classList.toggle('is-unassigned', !technology);
    }
  });

  const assignedCount = conversion.mappingRows.filter(row => conversion.assignments[row.id]).length;
  const total = conversion.mappingRows.length;
  const progress = document.getElementById('vtbAssignmentProgress');
  if (progress) {
    progress.textContent = `${assignedCount}/${total} đã gán`;
    progress.classList.toggle('is-complete', assignedCount === total);
  }
  const warning = document.getElementById('vtbSourceWarning');
  if (warning) {
    const amounts = conversion.unmappedAmounts || [];
    warning.hidden = amounts.length === 0;
    warning.textContent = amounts.length
      ? `Phát hiện ${amounts.length} ô có số lượng ở cột Giải nghẽn hoặc Củng cố. Cần xử lý dữ liệu nguồn trước khi chuyển đổi để không mất số liệu.`
      : '';
  }
  const ready = assignedCount === total && (conversion.unmappedAmounts || []).length === 0;
  const convertButton = document.getElementById('btnConvertVTB');
  if (convertButton) {
    convertButton.disabled = !ready;
    convertButton.title = ready ? 'Tách dòng và ghi vào Mẫu mới' : 'Gán đủ công nghệ và xử lý cảnh báo dữ liệu trước khi chuyển đổi';
  }
  updateVTBSelectionVisuals();
}

function selectVTBMappingRows(mode) {
  const conversion = state.vtbConversion;
  if (!conversion) return;
  if (mode === 'all') conversion.selectedIds = new Set(conversion.mappingRows.map(row => row.id));
  else if (mode === 'unassigned') conversion.selectedIds = new Set(conversion.mappingRows.filter(row => !conversion.assignments[row.id]).map(row => row.id));
  else conversion.selectedIds = new Set();
  updateVTBSelectionVisuals();
}

function applyVTBTechnologyToSelection(technology) {
  const conversion = state.vtbConversion;
  if (!conversion || conversion.selectedIds.size === 0) return;
  conversion.assignments = window.VTBConverter.applyVTBBulkTechnology(
    conversion.assignments,
    Array.from(conversion.selectedIds),
    technology
  );
  updateVTBAssignmentTableState();
}

function initVTBDragSelection() {
  const tbody = document.getElementById('vtbAssignmentTableBody');
  if (!tbody) return;
  let dragging = false;
  let selectMode = true;

  tbody.addEventListener('pointerdown', event => {
    if (event.button !== 0 || event.target.closest('select, input, button')) return;
    const row = event.target.closest('[data-vtb-mapping-id]');
    if (!row || !state.vtbConversion) return;
    dragging = true;
    const id = row.dataset.vtbMappingId;
    selectMode = !state.vtbConversion.selectedIds.has(id);
    setVTBMappingRowSelected(id, selectMode);
    event.preventDefault();
  });
  tbody.addEventListener('pointerover', event => {
    if (!dragging) return;
    const row = event.target.closest('[data-vtb-mapping-id]');
    if (row) setVTBMappingRowSelected(row.dataset.vtbMappingId, selectMode);
  });
  document.addEventListener('pointerup', () => { dragging = false; });
  document.addEventListener('pointercancel', () => { dragging = false; });
}

async function performVTBConversion() {
  const conversion = state.vtbConversion;
  if (!conversion) return;
  const button = document.getElementById('btnConvertVTB');
  if (button) button.disabled = true;
  try {
    const templateBuffer = await window.VTBConverter.loadVTBNewTemplateBuffer();
    const result = await window.VTBConverter.convertOldVTBToNewFormat(
      conversion.oldWorkbook,
      conversion.assignments,
      templateBuffer
    );
    conversion.result = result;
    conversion.convertedBuffer = result.convertedBuffer;
    conversion.convertedWorkbook = result.convertedWorkbook;
    renderVTBConversionResult(result, conversion.originalFileName);
    showToast(`Đã chuyển đổi ${result.mappedCount} dòng sang Masterlist mẫu mới.`, 'success');
  } catch (err) {
    console.error('Lỗi khi chuyển đổi file VTB:', err);
    showToast(`Lỗi khi chuyển đổi: ${err.message}`, 'error');
  } finally {
    updateVTBAssignmentTableState();
  }
}

function formatNumberVTB(num) {
  if (num === null || num === undefined || isNaN(num)) return '-';
  return Number(num).toLocaleString('vi-VN', { maximumFractionDigits: 2 });
}

function renderVTBConversionResult(res, fileName) {
  const kpiSection = document.getElementById('vtbKPISection');
  if (kpiSection) kpiSection.style.display = 'block';

  const kpiOldCount = document.getElementById('kpiOldItemCount');
  const kpiMappedCount = document.getElementById('kpiMappedCount');
  if (kpiOldCount) kpiOldCount.textContent = res.oldItemsCount;
  if (kpiMappedCount) kpiMappedCount.textContent = res.mappedCount;

  const kpiOld27 = document.getElementById('kpiOldTotal27');
  const kpiNew27 = document.getElementById('kpiNewTotal27');
  const kpiOld28 = document.getElementById('kpiOldTotal28');
  const kpiNew28 = document.getElementById('kpiNewTotal28');

  if (kpiOld27) kpiOld27.textContent = formatNumberVTB(res.totalOld27);
  if (kpiNew27) kpiNew27.textContent = formatNumberVTB(res.newTotal27);
  if (kpiOld28) kpiOld28.textContent = formatNumberVTB(res.totalOld28);
  if (kpiNew28) kpiNew28.textContent = formatNumberVTB(res.newTotal28);

  // Render bảng mapping
  const tbody = document.getElementById('vtbMappingTableBody');
  if (tbody) {
    tbody.innerHTML = '';
    res.mappedResults.forEach(item => {
      const tr = document.createElement('tr');
      let pillClass = 'vtb-pill-vp';
      if (item.targetDV === 'DLDĐ') pillClass = 'vtb-pill-dl';
      if (item.targetDV === 'ƯCTT') pillClass = 'vtb-pill-uctt';
      if (item.targetDV === 'VHKT') pillClass = 'vtb-pill-vhkt';
      if (item.targetDV === '5G') pillClass = 'vtb-pill-5g';

      tr.innerHTML = `
        <td style="text-align: center; font-weight: 700; color: #64748b;">${item.targetRow}</td>
        <td style="font-weight: 600; color: #1e293b;" title="${escapeHtml(item.targetName)}">${escapeHtml(item.targetName)}</td>
        <td style="text-align: center;"><span class="vtb-mapping-pill ${pillClass}">${item.targetDV}</span></td>
        <td style="text-align: right; font-weight: 700; color: ${item.kl27 ? '#047857' : '#94a3b8'};">${item.kl27 ? formatNumberVTB(item.kl27) : '-'}</td>
        <td style="text-align: right; font-weight: 700; color: ${item.kl28 ? '#047857' : '#94a3b8'};">${item.kl28 ? formatNumberVTB(item.kl28) : '-'}</td>
        <td style="text-align: right; color: #475569;">${item.dg ? formatNumberVTB(item.dg) : '-'}</td>
        <td style="color: #64748b;" title="${escapeHtml(item.oldName)}">
          <span style="font-weight: 600; color: #334155;">[${escapeHtml(item.oldPurpose)}]</span> ${escapeHtml(item.oldName)}
        </td>
      `;
      tbody.appendChild(tr);
    });
  }

  // Enable buttons
  const btnDownload = document.getElementById('btnDownloadConvertedVTB');
  const btnApply = document.getElementById('btnApplyConvertedToVT');
  if (btnDownload) btnDownload.disabled = false;
  if (btnApply) btnApply.disabled = false;

  if (window.lucide) lucide.createIcons();
}

function downloadConvertedVTBFile() {
  if (!state.vtbConversion || !state.vtbConversion.convertedBuffer) {
    showToast('Chưa có dữ liệu chuyển đổi để tải về!', 'error');
    return;
  }
  const blob = new Blob([state.vtbConversion.convertedBuffer], { 
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
  });
  const originalName = state.vtbConversion.originalFileName || 'Masterlist_VTB';
  const exportName = originalName.replace(/\.xlsx$/i, '') + '_Mau_moi.xlsx';
  triggerDownloadBlob(blob, exportName);
  showToast(`Đã tải về file Mẫu mới: ${exportName}`, 'success');
}

function applyConvertedVTBToAppState() {
  if (!state.vtbConversion || !state.vtbConversion.convertedBuffer) {
    showToast('Chưa có dữ liệu chuyển đổi để nạp vào App!', 'error');
    return;
  }

  const originalName = state.vtbConversion.originalFileName || 'Masterlist_VTB';
  const convertedFileName = originalName.replace(/\.xlsx$/i, '') + '_Mau_moi.xlsx';
  const convertedBuf = state.vtbConversion.convertedBuffer;
  const convertedWb = state.vtbConversion.convertedWorkbook;

  // Gán vào state.files['VT']
  state.files['VT'] = {
    file: new File([convertedBuf], convertedFileName, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
    name: convertedFileName,
    size: convertedBuf.byteLength,
    buffer: convertedBuf,
    workbook: convertedWb
  };

  const mang = MANG_CONFIG.find(m => m.code === 'VT');
  if (mang) {
    const items = extractItemsForMang(mang, state.files['VT']);
    state.extractedByMang['VT'] = items;
    showToast(`Đã nạp thành công ${items.length} dòng vào Mảng Vô Tuyến!`, 'success');
  }

  updateUploadBoxUI('VT', convertedFileName, convertedBuf.byteLength);
  rebuildExtractedData();
  closeVTBConverterModal();
}

// ==================== CHỈNH SỬA DỮ LIỆU & BẢN LƯU SERVER ====================

// Cập nhật dữ liệu của 1 hạng mục và đồng bộ toàn bộ hệ thống
function updateItemData(itemIndex, updatedFields) {
  const item = state.extractedData[itemIndex];
  if (!item) return;

  Object.assign(item, updatedFields);

  // Nếu không phải header mảng, tính lại thành tiền
  if (!item.isMangHeader) {
    const k27 = (item.kl27 !== null && item.kl27 !== undefined && item.kl27 !== '') ? Number(item.kl27) : null;
    const k28 = (item.kl28 !== null && item.kl28 !== undefined && item.kl28 !== '') ? Number(item.kl28) : null;
    const dg = (item.dg !== null && item.dg !== undefined && item.dg !== '') ? Number(item.dg) : null;

    item.kl27 = k27;
    item.kl28 = k28;
    item.dg = dg;

    if (k27 !== null && dg !== null && !isNaN(k27) && !isNaN(dg)) {
      item.tt27 = k27 * dg;
    } else {
      item.tt27 = null;
    }

    if (k28 !== null && dg !== null && !isNaN(k28) && !isNaN(dg)) {
      item.tt28 = k28 * dg;
    } else {
      item.tt28 = null;
    }

    item.tongTT = (item.tt27 || 0) + (item.tt28 || 0);
  }

  // Đồng bộ với mảng tương ứng trong state.extractedByMang
  if (item.mangCode && state.extractedByMang[item.mangCode]) {
    const list = state.extractedByMang[item.mangCode];
    const foundIdx = list.findIndex(x => x === item || (x.origRow === item.origRow && x.nd === item.nd));
    if (foundIdx >= 0) {
      Object.assign(list[foundIdx], item);
    }
  }

  // Hủy exportBlob cũ để build lại khi xuất file
  state.exportBlob = null;

  recalculateSubtotalFormulas();
  rebuildExtractedData();
}

// Mở modal chỉnh sửa chi tiết hạng mục
function openEditItemModal(itemIndex) {
  const item = state.extractedData[itemIndex];
  if (!item) return;

  const modal = document.getElementById('modalEditItem');
  if (!modal) return;

  document.getElementById('editItemIndex').value = itemIndex;

  const badgeMang = document.getElementById('editItemMangBadge');
  if (badgeMang) {
    badgeMang.className = `badge-mang badge-mang-${item.mangCode}`;
    badgeMang.textContent = item.mangCode;
  }
  const origRowEl = document.getElementById('editItemOrigRow');
  if (origRowEl) origRowEl.textContent = item.origRow || '-';
  const ttEl = document.getElementById('editItemTT');
  if (ttEl) ttEl.textContent = item.tt || '-';

  document.getElementById('editItemND').value = item.nd || '';
  document.getElementById('editItemDVT').value = item.dvt || '';
  document.getElementById('editItemKL27').value = (item.kl27 !== null && item.kl27 !== undefined) ? item.kl27 : '';
  document.getElementById('editItemKL28').value = (item.kl28 !== null && item.kl28 !== undefined) ? item.kl28 : '';
  document.getElementById('editItemDG').value = (item.dg !== null && item.dg !== undefined) ? item.dg : '';

  document.getElementById('editItemMaLoai').value = item.maLoai || '';
  document.getElementById('editItemMaMang').value = item.maMang || '';
  document.getElementById('editItemMaDV').value = item.maDV || '';

  const dl = document.getElementById('editItemMaDVSuggestions');
  if (dl && state.validMaDVSet.size > 0) {
    dl.innerHTML = Array.from(state.validMaDVSet).map(code => `<option value="${escapeHtml(code)}">`).join('');
  }

  const chkGroup = document.getElementById('editItemIsGroup');
  if (chkGroup) chkGroup.checked = !!item.isGroup;
  const selLevel = document.getElementById('editItemLevelNum');
  if (selLevel) selLevel.value = String(item.levelNum || (item.isGroup ? 2 : 99));

  updateEditItemMoneyPreview();
  modal.style.display = 'flex';
  if (window.lucide) lucide.createIcons();
}

// Mở modal sửa từ bảng Validation
function openEditItemFromValidation(origRow, mangCode) {
  const idx = state.extractedData.findIndex(x => x.mangCode === mangCode && x.origRow === origRow);
  if (idx >= 0) {
    openEditItemModal(idx);
  } else {
    showToast('Không tìm thấy dòng tương ứng trong bảng dữ liệu!', 'warning');
  }
}

// Tính realtime thành tiền trong modal sửa
function updateEditItemMoneyPreview() {
  const kl27Val = parseFloat(document.getElementById('editItemKL27')?.value);
  const kl28Val = parseFloat(document.getElementById('editItemKL28')?.value);
  const dgVal = parseFloat(document.getElementById('editItemDG')?.value);

  const tt27 = (!isNaN(kl27Val) && !isNaN(dgVal)) ? kl27Val * dgVal : 0;
  const tt28 = (!isNaN(kl28Val) && !isNaN(dgVal)) ? kl28Val * dgVal : 0;
  const total = tt27 + tt28;

  const p27 = document.getElementById('editItemPreviewTT27');
  if (p27) p27.textContent = tt27 > 0 ? formatNumber(tt27) + ' USD' : '-';
  const p28 = document.getElementById('editItemPreviewTT28');
  if (p28) p28.textContent = tt28 > 0 ? formatNumber(tt28) + ' USD' : '-';
  const pTot = document.getElementById('editItemPreviewTTTotal');
  if (pTot) pTot.textContent = total > 0 ? formatNumber(total) + ' USD' : '-';
}

function closeEditItemModal() {
  const modal = document.getElementById('modalEditItem');
  if (modal) modal.style.display = 'none';
}

function confirmSaveItem() {
  const idx = parseInt(document.getElementById('editItemIndex')?.value, 10);
  if (isNaN(idx) || !state.extractedData[idx]) return;

  const nd = document.getElementById('editItemND')?.value?.trim();
  if (!nd) {
    showToast('Tên hạng mục không được để trống!', 'warning');
    return;
  }

  const dvt = document.getElementById('editItemDVT')?.value?.trim() || '';
  const kl27Str = document.getElementById('editItemKL27')?.value;
  const kl28Str = document.getElementById('editItemKL28')?.value;
  const dgStr = document.getElementById('editItemDG')?.value;

  const kl27 = (kl27Str !== '' && !isNaN(parseFloat(kl27Str))) ? parseFloat(kl27Str) : null;
  const kl28 = (kl28Str !== '' && !isNaN(parseFloat(kl28Str))) ? parseFloat(kl28Str) : null;
  const dg = (dgStr !== '' && !isNaN(parseFloat(dgStr))) ? parseFloat(dgStr) : null;

  const maLoai = document.getElementById('editItemMaLoai')?.value?.trim() || '';
  const maMang = document.getElementById('editItemMaMang')?.value?.trim() || '';
  const maDV = document.getElementById('editItemMaDV')?.value?.trim() || '';

  const isGroup = document.getElementById('editItemIsGroup')?.checked || false;
  const levelNum = parseInt(document.getElementById('editItemLevelNum')?.value, 10) || (isGroup ? 2 : 99);

  updateItemData(idx, {
    nd: nd,
    dvt: dvt,
    kl27: kl27,
    kl28: kl28,
    dg: dg,
    maLoai: maLoai,
    maMang: maMang,
    maDV: maDV,
    isGroup: isGroup,
    levelNum: levelNum,
    level: levelNum <= 6 ? `CẤP ${levelNum}` : 'CHI TIẾT'
  });

  closeEditItemModal();
  showToast('Đã lưu thay đổi hạng mục thành công!', 'success');
}

function initEditModalListeners() {
  ['editItemKL27', 'editItemKL28', 'editItemDG'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.addEventListener('input', updateEditItemMoneyPreview);
  });
}

// Bấm đúp chuột để chỉnh sửa nhanh ngay trên bảng Xem trước (Preview)
function initPreviewInlineEditing() {
  const tbody = document.getElementById('previewTableBody');
  if (!tbody) return;

  tbody.addEventListener('dblclick', (e) => {
    const cell = e.target.closest('.cell-editable');
    if (!cell) return;
    if (cell.querySelector('input')) return;

    const itemIdx = parseInt(cell.getAttribute('data-item-idx'), 10);
    const field = cell.getAttribute('data-field');
    if (isNaN(itemIdx) || !field || !state.extractedData[itemIdx]) return;

    const item = state.extractedData[itemIdx];
    if (item.isMangHeader) return;

    const originalVal = (item[field] !== null && item[field] !== undefined) ? item[field] : '';
    const oldHtml = cell.innerHTML;

    const input = document.createElement('input');
    input.type = (field === 'kl27' || field === 'kl28' || field === 'dg') ? 'number' : 'text';
    if (input.type === 'number') input.step = 'any';
    input.className = 'cell-edit-input';
    input.value = originalVal;

    cell.innerHTML = '';
    cell.appendChild(input);
    input.focus();
    input.select();

    let committed = false;
    const commitChange = () => {
      if (committed) return;
      committed = true;
      const newValStr = input.value.trim();
      let newVal = newValStr;
      if (field === 'kl27' || field === 'kl28' || field === 'dg') {
        newVal = (newValStr !== '' && !isNaN(parseFloat(newValStr))) ? parseFloat(newValStr) : null;
      }
      if (newVal !== originalVal) {
        updateItemData(itemIdx, { [field]: newVal });
        showToast(`Đã cập nhật ${field.toUpperCase()} dòng ${item.origRow || itemIdx + 1}!`, 'success');
      } else {
        cell.innerHTML = oldHtml;
      }
    };

    input.addEventListener('blur', commitChange);
    input.addEventListener('keydown', (ke) => {
      if (ke.key === 'Enter') {
        ke.preventDefault();
        input.blur();
      } else if (ke.key === 'Escape') {
        ke.preventDefault();
        committed = true;
        cell.innerHTML = oldHtml;
      }
    });
  });
}

// ==================== BẢN LƯU SERVER (SERVER DRAFTS) ====================

async function initServerDrafts() {
  await updateServerDraftBadge();
}

async function updateServerDraftBadge() {
  const badge = document.getElementById('badgeServerDraftCount');
  if (!badge) return;
  try {
    const drafts = await fetchAllDraftsList();
    badge.textContent = drafts.length;
    badge.style.display = drafts.length > 0 ? 'inline-block' : 'none';
  } catch (e) {
    badge.textContent = '0';
  }
}

async function fetchAllDraftsList() {
  let serverList = [];
  try {
    const res = await fetch('/api/drafts', { cache: 'no-store' });
    if (res.ok) {
      const data = await res.json();
      if (data.success && Array.isArray(data.drafts)) {
        serverList = data.drafts;
      }
    }
  } catch (e) {
    // Offline
  }

  const localDrafts = JSON.parse(localStorage.getItem('qhdc_local_drafts') || '[]');
  const map = new Map();
  serverList.forEach(d => map.set(d.id, d));
  localDrafts.forEach(d => {
    if (!map.has(d.id)) map.set(d.id, d);
  });

  const allDrafts = Array.from(map.values());
  allDrafts.sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0));
  state.serverDrafts = allDrafts;
  return allDrafts;
}

function openSaveDraftModal() {
  if (!state.extractedData || state.extractedData.length === 0) {
    showToast('Chưa có dữ liệu để lưu! Vui lòng nạp ít nhất một mảng dữ liệu.', 'warning');
    return;
  }

  const modal = document.getElementById('modalSaveDraft');
  if (!modal) return;

  const activeMangs = MANG_CONFIG.filter(m => (state.extractedByMang[m.code] || []).length > 0);
  let sumTot = 0;
  state.extractedData.forEach(it => {
    if (!it.isGroup && !it.isMangHeader) {
      if (it.tt27) sumTot += Number(it.tt27);
      if (it.tt28) sumTot += Number(it.tt28);
    }
  });

  document.getElementById('saveDraftMangsCount').textContent = `${activeMangs.length} mảng (${activeMangs.map(m => m.code).join(', ')})`;
  document.getElementById('saveDraftRowCount').textContent = formatNumber(state.extractedData.length);
  document.getElementById('saveDraftTotalVal').textContent = formatNumber(sumTot);

  const now = new Date();
  const timeStr = `${String(now.getDate()).padStart(2, '0')}/${String(now.getMonth() + 1).padStart(2, '0')}/${now.getFullYear()} ${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;

  const nameInput = document.getElementById('inputSaveDraftName');
  if (nameInput) {
    nameInput.value = state.currentDraftName || `Bản lưu ngày ${timeStr}`;
  }

  const noteInput = document.getElementById('inputSaveDraftNote');
  if (noteInput) noteInput.value = '';

  const modeGroup = document.getElementById('saveDraftModeGroup');
  if (modeGroup) {
    if (state.currentDraftId) {
      modeGroup.style.display = 'block';
      const textOv = document.getElementById('textOverwriteDraftName');
      if (textOv) textOv.textContent = state.currentDraftName || state.currentDraftId;
      const rNew = document.querySelector('input[name="radioDraftSaveMode"][value="new"]');
      if (rNew) rNew.checked = true;
    } else {
      modeGroup.style.display = 'none';
    }
  }

  modal.style.display = 'flex';
  if (window.lucide) lucide.createIcons();
}

function closeSaveDraftModal() {
  const modal = document.getElementById('modalSaveDraft');
  if (modal) modal.style.display = 'none';
}

async function confirmSaveDraft() {
  const nameInput = document.getElementById('inputSaveDraftName');
  const noteInput = document.getElementById('inputSaveDraftNote');
  const name = (nameInput?.value || '').trim();
  if (!name) {
    showToast('Vui lòng nhập tên cho bản lưu!', 'warning');
    nameInput?.focus();
    return;
  }
  const note = (noteInput?.value || '').trim();
  const modeRadio = document.querySelector('input[name="radioDraftSaveMode"]:checked');
  const isOverwrite = (modeRadio && modeRadio.value === 'overwrite' && state.currentDraftId);

  const btnConfirm = document.getElementById('btnConfirmSaveDraft');
  const oldBtnHtml = btnConfirm.innerHTML;
  btnConfirm.disabled = true;
  btnConfirm.innerHTML = `<span class="spinner"></span> Đang lưu...`;

  try {
    let excelBase64 = null;
    try {
      if (state.templateBuffer) {
        recalculateSubtotalFormulas();
        const buffer = await buildCleanMasterlist(state.templateBuffer, state.extractedByMang, state.files);
        const bytes = new Uint8Array(buffer);
        let binary = '';
        const chunkSz = 0x8000;
        for (let i = 0; i < bytes.length; i += chunkSz) {
          binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunkSz));
        }
        excelBase64 = btoa(binary);
      }
    } catch (e) {
      console.warn('Lỗi tạo file Excel đính kèm bản lưu:', e);
    }

    const activeMangs = MANG_CONFIG.filter(m => (state.extractedByMang[m.code] || []).length > 0).map(m => m.code);
    let sum27 = 0, sum28 = 0;
    state.extractedData.forEach(it => {
      if (!it.isGroup && !it.isMangHeader) {
        if (it.tt27) sum27 += Number(it.tt27);
        if (it.tt28) sum28 += Number(it.tt28);
      }
    });

    const filesMeta = {};
    for (const m of MANG_CONFIG) {
      if (state.files[m.code]) {
        filesMeta[m.code] = {
          name: state.files[m.code].name,
          size: state.files[m.code].size
        };
      }
    }

    const payload = {
      id: isOverwrite ? state.currentDraftId : null,
      name: name,
      note: note,
      rowCount: state.extractedData.length,
      mangs: activeMangs,
      total2027: sum27,
      total2028: sum28,
      sessionData: {
        extractedByMang: state.extractedByMang,
        mangHeaderInfo: state.mangHeaderInfo,
        groupsConfig: state.groupsConfig,
        filesMeta: filesMeta
      },
      excelBase64: excelBase64
    };

    let savedDraft = null;
    let savedOnServer = false;

    try {
      const res = await fetch('/api/save-draft', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      if (res.ok) {
        const resData = await res.json();
        if (resData.success && resData.draft) {
          savedDraft = resData.draft;
          savedOnServer = true;
        }
      }
    } catch (netErr) {
      console.warn('Không kết nối được server, lưu vào localStorage dự phòng:', netErr);
    }

    if (!savedOnServer) {
      const now = new Date();
      const localId = isOverwrite ? state.currentDraftId : `local_draft_${Date.now()}`;
      savedDraft = {
        id: localId,
        name: name,
        note: note,
        timestamp: Date.now() / 1000,
        formattedTime: `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')} ${String(now.getDate()).padStart(2, '0')}/${String(now.getMonth() + 1).padStart(2, '0')}/${now.getFullYear()}`,
        rowCount: state.extractedData.length,
        mangs: activeMangs,
        total2027: sum27,
        total2028: sum28,
        sessionData: payload.sessionData,
        isLocal: true
      };
      const localDrafts = JSON.parse(localStorage.getItem('qhdc_local_drafts') || '[]');
      const existIdx = localDrafts.findIndex(d => d.id === localId);
      if (existIdx >= 0) localDrafts[existIdx] = savedDraft;
      else localDrafts.unshift(savedDraft);
      localStorage.setItem('qhdc_local_drafts', JSON.stringify(localDrafts));
    }

    state.currentDraftId = savedDraft.id;
    state.currentDraftName = savedDraft.name;

    closeSaveDraftModal();
    showToast(`Đã lưu bản lưu "${name}" thành công!`, 'success');
    updateServerDraftBadge();
  } catch (err) {
    console.error('Lỗi khi lưu bản lưu:', err);
    showToast(`Lỗi khi lưu bản lưu: ${err.message}`, 'error');
  } finally {
    btnConfirm.disabled = false;
    btnConfirm.innerHTML = oldBtnHtml;
    if (window.lucide) lucide.createIcons();
  }
}

function openServerDraftsModal() {
  const modal = document.getElementById('modalServerDrafts');
  if (!modal) return;
  modal.style.display = 'flex';
  fetchAndRenderServerDrafts();
  if (window.lucide) lucide.createIcons();
}

function closeServerDraftsModal() {
  const modal = document.getElementById('modalServerDrafts');
  if (modal) modal.style.display = 'none';
}

async function fetchAndRenderServerDrafts() {
  const tbody = document.getElementById('serverDraftsTableBody');
  if (!tbody) return;
  tbody.innerHTML = `
    <tr>
      <td colspan="6" style="text-align: center; padding: 2rem; color: #64748b;">
        <span class="spinner"></span> Đang tải danh sách bản lưu...
      </td>
    </tr>
  `;
  await fetchAllDraftsList();
  renderServerDraftsTable();
}

function renderServerDraftsTable() {
  const tbody = document.getElementById('serverDraftsTableBody');
  if (!tbody) return;
  const keyword = (document.getElementById('searchDraftsInput')?.value || '').trim().toLowerCase();
  let list = state.serverDrafts || [];
  if (keyword) {
    list = list.filter(d =>
      (d.name || '').toLowerCase().includes(keyword) ||
      (d.note || '').toLowerCase().includes(keyword) ||
      (d.formattedTime || '').toLowerCase().includes(keyword) ||
      (d.mangs || []).join(' ').toLowerCase().includes(keyword)
    );
  }

  if (list.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="6" style="text-align: center; padding: 2.5rem; color: #64748b;">
          ${keyword ? 'Không tìm thấy bản lưu nào phù hợp với từ khóa.' : 'Chưa có bản lưu nào trên server. Sau khi tải dữ liệu và chỉnh sửa, bạn hãy bấm "Lưu lên Server" để lưu trữ lại bản làm việc.'}
        </td>
      </tr>
    `;
    return;
  }

  tbody.innerHTML = list.map((d, i) => {
    const isCurrent = state.currentDraftId && state.currentDraftId === d.id;
    const currentPill = isCurrent ? `<span style="font-size: 0.7rem; background: #ecfdf5; color: #047857; border: 1px solid #a7f3d0; padding: 1px 6px; border-radius: 9999px; margin-left: 6px; font-weight: 700;">Đang mở</span>` : '';
    const noteHtml = d.note ? `<div style="font-size: 0.75rem; color: #64748b; margin-top: 3px; font-style: italic;">${escapeHtml(d.note)}</div>` : '';
    const mangBadges = (d.mangs || []).map(m => `<span class="badge-mang badge-mang-${m}" style="font-size: 0.65rem; padding: 1px 5px;">${m}</span>`).join(' ');
    const totVal = (d.total2027 || 0) + (d.total2028 || 0);

    return `
      <tr>
        <td style="text-align: center; font-weight: 700; color: #64748b;">${i + 1}</td>
        <td>
          <div style="font-weight: 700; color: #1e293b; display: flex; align-items: center;">
            <i data-lucide="file-spreadsheet" class="w-4 h-4 text-emerald-600 inline-block mr-1.5 shrink-0"></i>
            <span>${escapeHtml(d.name)}</span>
            ${currentPill}
          </div>
          ${noteHtml}
        </td>
        <td style="text-align: center; font-size: 0.775rem; color: #475569; font-weight: 600;">
          ${escapeHtml(d.formattedTime || '-')}
        </td>
        <td style="text-align: center;">
          <div style="font-weight: 700; font-size: 0.8rem; color: #0f172a;">${formatNumber(d.rowCount || 0)} dòng</div>
          <div style="margin-top: 3px; display: flex; gap: 3px; justify-content: center; flex-wrap: wrap;">${mangBadges}</div>
        </td>
        <td style="text-align: right; font-family: 'JetBrains Mono', monospace; font-weight: 700; color: #047857; font-size: 0.825rem;">
          ${totVal > 0 ? formatNumber(totVal) + ' USD' : '-'}
        </td>
        <td>
          <div class="draft-card-actions">
            <button type="button" class="btn btn-primary btn-xs" onclick="loadDraftIntoState('${d.id}')" title="Nạp bản này vào bảng làm việc để sửa tiếp">
              <i data-lucide="zap" class="w-3.5 h-3.5"></i>
              <span>Nạp bản này</span>
            </button>
            ${(d.hasExcel || !d.isLocal) ? `
              <button type="button" class="btn btn-outline btn-xs" onclick="downloadServerDraftExcel('${d.id}', '${escapeHtml(d.name)}')" title="Tải file Excel (.xlsx) của bản này về máy tính">
                <i data-lucide="download" class="w-3.5 h-3.5 text-emerald-600"></i>
                <span>Excel</span>
              </button>
            ` : ''}
            <button type="button" class="btn btn-outline btn-xs" style="color: #e11d48; border-color: #fecdd3;" onclick="deleteServerDraft('${d.id}', '${escapeHtml(d.name)}')" title="Xóa bản lưu này khỏi server">
              <i data-lucide="trash-2" class="w-3.5 h-3.5"></i>
            </button>
          </div>
        </td>
      </tr>
    `;
  }).join('');

  if (window.lucide) lucide.createIcons();
}

function downloadServerDraftExcel(draftId, draftName) {
  const url = `/api/draft-excel?id=${encodeURIComponent(draftId)}`;
  const a = document.createElement('a');
  a.href = url;
  a.download = `${draftName || 'Masterlist_Draft'}.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  showToast('Đang tải file Excel của bản lưu...', 'info');
}

async function deleteServerDraft(draftId, draftName) {
  if (!confirm(`Bạn có chắc chắn muốn xóa bản lưu "${draftName}" khỏi server?`)) return;

  try {
    try {
      await fetch(`/api/delete-draft?id=${encodeURIComponent(draftId)}`, { method: 'POST' });
    } catch (e) {}

    const localDrafts = JSON.parse(localStorage.getItem('qhdc_local_drafts') || '[]');
    const filtered = localDrafts.filter(d => d.id !== draftId);
    localStorage.setItem('qhdc_local_drafts', JSON.stringify(filtered));

    if (state.currentDraftId === draftId) {
      state.currentDraftId = null;
      state.currentDraftName = null;
    }

    showToast(`Đã xóa bản lưu "${draftName}"`, 'success');
    fetchAndRenderServerDrafts();
    updateServerDraftBadge();
  } catch (err) {
    showToast(`Lỗi khi xóa bản lưu: ${err.message}`, 'error');
  }
}

async function loadDraftIntoState(draftId) {
  try {
    showToast('Đang tải dữ liệu bản lưu từ server...', 'info');
    let draft = null;

    try {
      const res = await fetch(`/api/load-draft?id=${encodeURIComponent(draftId)}`);
      if (res.ok) {
        draft = await res.json();
      }
    } catch (e) {}

    if (!draft) {
      const localDrafts = JSON.parse(localStorage.getItem('qhdc_local_drafts') || '[]');
      draft = localDrafts.find(d => d.id === draftId);
    }

    if (!draft) {
      throw new Error('Không tìm thấy dữ liệu bản lưu!');
    }

    const sessionData = draft.sessionData || {};

    if (sessionData.extractedByMang) {
      state.extractedByMang = sessionData.extractedByMang;
    }
    if (sessionData.mangHeaderInfo) {
      state.mangHeaderInfo = sessionData.mangHeaderInfo;
    }
    if (sessionData.groupsConfig) {
      state.groupsConfig = sessionData.groupsConfig;
    }

    state.currentDraftId = draft.id;
    state.currentDraftName = draft.name;

    const filesMeta = sessionData.filesMeta || {};
    for (const mang of MANG_CONFIG) {
      const items = state.extractedByMang[mang.code] || [];
      if (items.length > 0) {
        const meta = filesMeta[mang.code] || {};
        const fileName = meta.name || `${draft.name} (${mang.code})`;
        const fileSize = meta.size || (items.length * 150);
        updateUploadBoxUI(mang.code, fileName, fileSize);
      } else {
        resetUploadBoxUI(mang.code);
      }
    }

    rebuildExtractedData();
    recalculateSubtotalFormulas();
    validateAllExtractedData();
    renderPreviewTable();
    renderValidationTable();
    renderHierarchyTable();
    syncStrategyComparisonRealtime();

    state.exportBlob = null;
    const btnDownload = document.getElementById('btnOpenResultFile');
    if (btnDownload) btnDownload.disabled = false;
    const btnSaveServer = document.getElementById('btnSaveToServer');
    if (btnSaveServer) btnSaveServer.disabled = false;

    closeServerDraftsModal();
    showToast(`Đã nạp bản lưu "${draft.name}" thành công! (${state.extractedData.length} dòng)`, 'success');
  } catch (err) {
    console.error('Lỗi khi nạp bản lưu:', err);
    showToast(`Không thể nạp bản lưu: ${err.message}`, 'error');
  }
}

window.openSaveDraftModal = openSaveDraftModal;
window.closeSaveDraftModal = closeSaveDraftModal;
window.confirmSaveDraft = confirmSaveDraft;
window.openServerDraftsModal = openServerDraftsModal;
window.closeServerDraftsModal = closeServerDraftsModal;
window.fetchAndRenderServerDrafts = fetchAndRenderServerDrafts;
window.renderServerDraftsTable = renderServerDraftsTable;
window.loadDraftIntoState = loadDraftIntoState;
window.downloadServerDraftExcel = downloadServerDraftExcel;
window.deleteServerDraft = deleteServerDraft;
window.openEditItemModal = openEditItemModal;
window.openEditItemFromValidation = openEditItemFromValidation;
window.closeEditItemModal = closeEditItemModal;
window.confirmSaveItem = confirmSaveItem;



