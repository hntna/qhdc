# Làm lại giao diện & điều hướng theo thương hiệu Viettel — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Chuyển giao diện app QHĐC 2027-2028 sang bộ nhận diện Viettel (tông đỏ) và tổ chức lại điều hướng thành sidebar 6 khu, giữ nguyên 100% logic JS.

**Architecture:** Reskin CSS-first + thêm một lớp "section" điều khiển bởi sidebar. Các khối `tab-content` hiện có (`tabValidation`, `tabPreview`, `tabHierarchy`, `tabStrategy`) được di chuyển vào 6 `section-panel`, giữ nguyên mọi `id`/`data-*`. Một file JS mới (`ui-nav.js`) đảm nhận chuyển section; `app.js`/`report.js`/`vtb_converter.js` không đổi hành vi. Không có framework test front-end; verify bằng (a) script Node kiểm tra bất biến cấu trúc HTML và (b) smoke test trình duyệt qua `python server.py`.

**Tech Stack:** HTML tĩnh + CSS thuần (`styles.css`, `report.css`) + JavaScript thuần (không build step). Server: `python server.py` (thư viện chuẩn). Icons: lucide. Node dùng để chạy script kiểm tra cấu trúc.

## Global Constraints

- **Không đổi hành vi JS:** giữ nguyên mọi `id`, `data-tab`, `data-subtab`, `name`, và cấu trúc mà JS truy vấn (`app.js`, `report.js`, `vtb_converter.js`). Chỉ được thêm class trình bày, di chuyển/bọc markup, và thêm listener điều hướng trong file mới `ui-nav.js`.
- **Danh sách id BẮT BUỘC còn sống** (mọi task phải giữ): `btnOpenVTBConverter`, `btnSaveToServer`, `btnOpenResultFile`, `btnReset`, `btnLoadMultiFile`, `inputMultiFile`, `btnOpenServerDrafts`, `badgeServerDraftCount`, `input_VT` `input_ML` `input_CDBR` `input_CNTT` `input_TD` `input_CD` `input_HT`, `tabValidation`, `tabPreview`, `tabHierarchy`, `tabStrategy`, `badgeValidationCount`, `badgeRowCount`, `badgeGroupCount`, `btnReindexTT`, `btnFixOutlineGroups`, `btnSplitWalletDV`, `btnSplitWalletMang`, `btnRestoreOriginalTT`, `selectReportUnit`, `btnReloadReport`.
- **Data-subtab bắt buộc còn sống:** `subtabTongHop`, `subtabMang`, `subtabDichVu`, `subtabStratDetail`.
- **Màu thương hiệu:** primary `#EE0033`, hover `#C4002A`, soft `#FFF1F3`, glow `rgba(238,0,51,0.18)`.
- **Giữ hệ 7 màu mảng** (VT/ML/CĐBR/CNTT/TĐ/CĐ/HT), chỉ hạ bão hòa; chữ trắng, tương phản WCAG AA.
- **Font:** giữ Inter + JetBrains Mono. **Offline:** không thêm phụ thuộc mạng ngoài bắt buộc.
- **Không thêm tính năng nghiệp vụ mới** trừ placeholder "Sắp có" cho Tách IP → Mảng.
- Commit sau mỗi task. Làm việc trên nhánh `feature/viettel-ui-redesign` (không commit thẳng `main`).

---

### Task 1: Guard cấu trúc HTML (regression test)

Viết script Node kiểm tra các bất biến cấu trúc PHẢI đúng cả trước và sau khi tái cấu trúc. Chạy được ngay trên HTML hiện tại (PASS), và giữ PASS sau mỗi task sau.

**Files:**
- Create: `tests/check_ui_structure.js`

**Interfaces:**
- Produces: lệnh `node tests/check_ui_structure.js` → exit 0 nếu mọi id/attr bắt buộc tồn tại, exit 1 + in danh sách thiếu nếu không.

- [ ] **Step 1: Viết script kiểm tra (đây là "test" cho công việc DOM)**

```js
// tests/check_ui_structure.js
// Kiem tra bat bien cau truc index.html: cac id/attr ma JS phu thuoc phai luon ton tai.
const fs = require('fs');
const path = require('path');

const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');

const REQUIRED_IDS = [
  'btnOpenVTBConverter','btnSaveToServer','btnOpenResultFile','btnReset',
  'btnLoadMultiFile','inputMultiFile','btnOpenServerDrafts','badgeServerDraftCount',
  'input_VT','input_ML','input_CDBR','input_CNTT','input_TD','input_CD','input_HT',
  'tabValidation','tabPreview','tabHierarchy','tabStrategy',
  'badgeValidationCount','badgeRowCount','badgeGroupCount',
  'btnReindexTT','btnFixOutlineGroups','btnSplitWalletDV','btnSplitWalletMang',
  'btnRestoreOriginalTT','selectReportUnit','btnReloadReport',
];
const REQUIRED_SUBTABS = ['subtabTongHop','subtabMang','subtabDichVu','subtabStratDetail'];

const missing = [];
for (const id of REQUIRED_IDS) {
  const re = new RegExp('id\\s*=\\s*["\']' + id + '["\']');
  if (!re.test(html)) missing.push('id=' + id);
}
for (const st of REQUIRED_SUBTABS) {
  if (!html.includes('data-subtab="' + st + '"')) missing.push('data-subtab=' + st);
}

if (missing.length) {
  console.error('FAIL - thieu ' + missing.length + ' phan tu bat buoc:');
  missing.forEach(m => console.error('  - ' + m));
  process.exit(1);
}
console.log('PASS - tat ca ' + (REQUIRED_IDS.length + REQUIRED_SUBTABS.length) + ' phan tu bat buoc con nguyen.');
process.exit(0);
```

- [ ] **Step 2: Chạy để xác nhận PASS trên HTML hiện tại**

Run: `node tests/check_ui_structure.js`
Expected: `PASS - tat ca 33 phan tu bat buoc con nguyen.`

- [ ] **Step 3: Commit**

```bash
git checkout -b feature/viettel-ui-redesign
git add tests/check_ui_structure.js
git commit -m "test: them guard kiem tra bat bien cau truc UI truoc khi reskin"
```

---

### Task 2: Design tokens Viettel trong `styles.css`

Thay bộ token màu trong `:root` sang tông đỏ Viettel; các component cũ tự đổi màu theo biến. Không đụng cấu trúc.

**Files:**
- Modify: `styles.css:6-65` (khối `:root`)

**Interfaces:**
- Produces: các biến `--primary`, `--primary-hover`, `--primary-soft`, `--primary-glow`, `--bg-app`, `--border-subtle`, `--text-*`, và cặp token 7 mảng đã hạ bão hòa. Task sau dùng lại các biến này.

- [ ] **Step 1: Thay khối `:root` (styles.css dòng 6-65)**

Thay toàn bộ nội dung `:root { ... }` hiện tại bằng:

```css
:root {
  /* Viettel Brand Light System */
  --bg-app: #F5F6F8;
  --bg-surface: #ffffff;
  --bg-card: #ffffff;
  --bg-card-hover: #FBFBFC;
  --bg-subtle: #F3F4F6;
  --bg-muted: #E5E7EB;

  --text-main: #1A1A1A;
  --text-secondary: #4B5563;
  --text-muted: #6B7280;
  --text-dim: #9CA3AF;

  --border-subtle: #E5E7EB;
  --border-strong: #D1D5DB;
  --border-focus: #EE0033;

  /* Primary: Viettel Red */
  --primary: #EE0033;
  --primary-hover: #C4002A;
  --primary-soft: #FFF1F3;
  --primary-glow: rgba(238, 0, 51, 0.18);

  /* Semantic (danger tach khoi primary) */
  --success: #047857; --success-bg: #ecfdf5; --success-border: #6ee7b7;
  --warning: #b45309; --warning-bg: #fffbeb; --warning-border: #fcd34d;
  --danger: #B91C1C;  --danger-bg: #FEF2F2;  --danger-border: #FCA5A5;
  --info: #0369a1;    --info-bg: #f0f9ff;    --info-border: #7dd3fc;

  /* 7 Mang (giu he mau, ha bao hoa nhe) */
  --c-vt: #ffffff;   --bg-vt: #3F3D9C;   --b-vt: #33327F;
  --c-ml: #ffffff;   --bg-ml: #7A3196;   --b-ml: #65297D;
  --c-cdbr: #ffffff; --bg-cdbr: #2586C0; --b-cdbr: #1F72A3;
  --c-cntt: #ffffff; --bg-cntt: #2F8F6E; --b-cntt: #26775B;
  --c-td: #ffffff;   --bg-td: #6FA31E;   --b-td: #5C871A;
  --c-cd: #ffffff;   --bg-cd: #C43050;   --b-cd: #A62742;
  --c-ht: #ffffff;   --bg-ht: #C0562A;   --b-ht: #A24824;

  /* Radius & Shadows */
  --radius-sm: 8px; --radius-md: 12px; --radius-lg: 16px; --radius-xl: 20px;
  --shadow-card: 0 1px 3px rgba(16,24,40,0.06), 0 1px 2px rgba(16,24,40,0.04);
  --shadow-card-hover: 0 8px 20px rgba(16,24,40,0.09), 0 2px 4px rgba(16,24,40,0.04);
  --shadow-dropdown: 0 14px 32px -4px rgba(16,24,40,0.15), 0 4px 12px -2px rgba(16,24,40,0.08);

  /* Layout */
  --sidebar-w: 248px;
  --sidebar-w-collapsed: 64px;
  --header-h: 60px;
}
```

- [ ] **Step 2: Cập nhật tiện ích màu indigo còn hardcode**

Trong `styles.css` (khối tiện ích dòng ~104-120), đổi các class indigo sang đỏ để không còn tím lạc lõng:

```css
.text-indigo-500 { color: #EE0033; }
.text-indigo-600 { color: #EE0033; }
.text-indigo-700 { color: #C4002A; }
.selection\:bg-indigo-500\/20::selection { background-color: rgba(238,0,51,0.18); }
.selection\:text-indigo-800::selection { color: #C4002A; }
```

- [ ] **Step 3: Verify — guard vẫn PASS + xem trình duyệt**

Run: `node tests/check_ui_structure.js`
Expected: `PASS - tat ca 33 phan tu bat buoc con nguyen.`

Run: `python server.py` rồi mở `http://localhost:8080/index.html`. Kỳ vọng: app hiển thị bình thường, các điểm nhấn tím trước đây (nút, badge, viền focus) chuyển sang đỏ Viettel; không vỡ layout; Console (F12) không có lỗi mới.

- [ ] **Step 4: Commit**

```bash
git add styles.css
git commit -m "feat(ui): ap dung bo token mau thuong hieu Viettel (do chu dao)"
```

---

### Task 3: Khung sidebar + lớp section trong `index.html`

Bọc nội dung hiện có vào 6 `section-panel`, thêm `<aside class="app-sidebar">`, gỡ `.tabs-nav` cũ. Di chuyển `#uploadCard` vào section "Nạp dữ liệu" và `#tabPreview` vào section "Trang chủ". Giữ nguyên mọi id.

**Files:**
- Modify: `index.html` (body: header dòng ~28-57; `main.main-layout` dòng ~60; `#uploadCard` dòng ~63-183; `.tabs-nav` dòng ~185-207; các `tab-content`)

**Interfaces:**
- Consumes: các id/khối từ HTML hiện tại (không đổi).
- Produces: khung DOM: `aside.app-sidebar` chứa 6 nút `.nav-item[data-section]` với `data-section` ∈ {`home`,`data`,`validation`,`tools`,`config`,`report`}; 6 `section.section-panel[data-panel]` cùng bộ khóa; class `.app-shell` bọc header + main. Task 4 (`ui-nav.js`) tiêu thụ `data-section`/`data-panel`.

- [ ] **Step 1: Thêm sidebar ngay sau `<body ...>` (trước `<header>`)**

```html
  <!-- Sidebar dieu huong Viettel -->
  <aside class="app-sidebar" id="appSidebar">
    <div class="sidebar-brand">
      <div class="sidebar-logo"><i data-lucide="bar-chart-3" class="w-5 h-5 text-white"></i></div>
      <div class="sidebar-brand-text">
        <span class="sidebar-brand-title">QHĐC 2027–2028</span>
        <span class="sidebar-brand-sub">Masterlist</span>
      </div>
      <button class="sidebar-toggle" id="btnSidebarToggle" title="Thu gọn / mở rộng menu">
        <i data-lucide="panel-left-close" class="w-4 h-4"></i>
      </button>
    </div>
    <nav class="sidebar-nav">
      <button class="nav-item active" data-section="home"><i data-lucide="home" class="w-4 h-4"></i><span>Trang chủ</span></button>
      <button class="nav-item" data-section="data"><i data-lucide="folder-input" class="w-4 h-4"></i><span>Nạp dữ liệu</span></button>
      <button class="nav-item" data-section="validation"><i data-lucide="shield-check" class="w-4 h-4"></i><span>Kiểm tra dữ liệu</span><span class="nav-badge badge-danger" id="navBadgeValidation">0</span></button>
      <button class="nav-item" data-section="tools"><i data-lucide="wrench" class="w-4 h-4"></i><span>Công cụ</span></button>
      <button class="nav-item" data-section="config"><i data-lucide="git-merge" class="w-4 h-4"></i><span>Cấu hình</span></button>
      <button class="nav-item" data-section="report"><i data-lucide="pie-chart" class="w-4 h-4"></i><span>Tổng hợp</span></button>
    </nav>
  </aside>
  <div class="app-shell" id="appShell">
```

- [ ] **Step 2: Đóng `.app-shell`**: thêm `</div>` ngay trước `</body>` (sau `#toastContainer`, trước `<script src="app.js"...>`). Giữ thẻ script ở nguyên vị trí.

- [ ] **Step 3: Gỡ khối `.tabs-nav`** (index.html dòng ~185-207): xóa toàn bộ `<div class="tabs-nav"> ... </div>`. (JS `initTabs` sẽ không tìm thấy `.tab-btn` → vô hại.)

- [ ] **Step 4: Bọc nội dung vào 6 section-panel.** Trong `<main class="main-layout">`, sắp lại theo thứ tự sau (di chuyển khối, KHÔNG đổi nội dung bên trong):

```html
    <!-- TRANG CHU -->
    <section class="section-panel active" data-panel="home">
      <div class="dashboard-kpis" id="dashboardKpis">
        <div class="kpi-card"><span class="kpi-label">Số dòng gộp</span><span class="kpi-value" id="kpiRows">0</span></div>
        <div class="kpi-card kpi-danger"><span class="kpi-label">Lỗi kiểm tra</span><span class="kpi-value" id="kpiErrors">0</span></div>
        <div class="kpi-card"><span class="kpi-label">Số nhóm</span><span class="kpi-value" id="kpiGroups">0</span></div>
        <div class="kpi-card"><span class="kpi-label">Mảng đã nạp</span><span class="kpi-value" id="kpiSectors">0/7</span></div>
      </div>
      <!-- di chuyen NGUYEN KHOI #tabPreview vao day -->
    </section>

    <!-- NAP DU LIEU -->
    <section class="section-panel" data-panel="data">
      <!-- di chuyen NGUYEN KHOI <section class="card" id="uploadCard"> vao day -->
    </section>

    <!-- KIEM TRA DU LIEU -->
    <section class="section-panel" data-panel="validation">
      <!-- di chuyen NGUYEN KHOI #tabValidation vao day -->
    </section>

    <!-- CONG CU (Task 6 se dien noi dung) -->
    <section class="section-panel" data-panel="tools">
      <div id="toolsGrid"></div>
    </section>

    <!-- CAU HINH -->
    <section class="section-panel" data-panel="config">
      <!-- di chuyen NGUYEN KHOI #tabHierarchy vao day -->
    </section>

    <!-- TONG HOP -->
    <section class="section-panel" data-panel="report">
      <!-- di chuyen NGUYEN KHOI #tabStrategy vao day -->
    </section>
```

- [ ] **Step 5: Ép các `tab-content` luôn hiển thị trong section của nó.** Với mỗi khối `#tabValidation`, `#tabPreview`, `#tabHierarchy`, `#tabStrategy`: thêm class `active` vào thẻ mở (ví dụ `<div id="tabPreview" class="tab-content active">`). Việc ẩn/hiện giờ do `.section-panel` đảm nhận.

- [ ] **Step 6: Verify guard**

Run: `node tests/check_ui_structure.js`
Expected: `PASS - tat ca 33 phan tu bat buoc con nguyen.`

(Chưa cần đẹp — Task 5 sẽ style sidebar. Lúc này layout có thể lệch, chấp nhận được.)

- [ ] **Step 7: Commit**

```bash
git add index.html
git commit -m "refactor(ui): tai cau truc DOM sang sidebar 6 khu + section-panel, giu nguyen id"
```

---

### Task 4: Điều hướng sidebar (`ui-nav.js`)

Thêm file JS mới chuyển section theo sidebar; kích hoạt render báo cáo khi vào khu Tổng hợp; đồng bộ badge; thu gọn sidebar. Không sửa `app.js`.

**Files:**
- Create: `ui-nav.js`
- Modify: `index.html:1308` (thêm `<script src="ui-nav.js" defer></script>` ngay sau app.js)

**Interfaces:**
- Consumes: `.nav-item[data-section]`, `.section-panel[data-panel]` (Task 3); các hàm render báo cáo toàn cục nếu có (`updateReportFromExtractedData`, `window.renderAllReportTabs`, `renderStrategyComparisonTable`) — gọi phòng thủ bằng `typeof`.
- Produces: hành vi chuyển khu; hàm `window.syncSidebarBadges()` để nơi khác gọi (tùy chọn).

- [ ] **Step 1: Tạo `ui-nav.js`**

```js
// ui-nav.js — Lop dieu huong sidebar (thuan trinh bay, khong dung logic nghiep vu)
(function () {
  function showSection(key) {
    document.querySelectorAll('.section-panel').forEach(p => {
      p.classList.toggle('active', p.getAttribute('data-panel') === key);
    });
    document.querySelectorAll('.nav-item').forEach(n => {
      n.classList.toggle('active', n.getAttribute('data-section') === key);
    });
    if (key === 'report') {
      try {
        if (typeof updateReportFromExtractedData === 'function' && window.state) {
          updateReportFromExtractedData(state.extractedData, state.extractedByMang);
        } else if (typeof window.renderAllReportTabs === 'function') {
          window.renderAllReportTabs();
        }
        if (typeof renderStrategyComparisonTable === 'function') renderStrategyComparisonTable();
      } catch (e) { console.warn('render report on nav:', e); }
    }
    if (window.lucide) setTimeout(() => lucide.createIcons(), 10);
  }

  function syncSidebarBadges() {
    const map = [['badgeValidationCount','navBadgeValidation'],['badgeValidationCount','kpiErrors'],
                 ['badgeRowCount','kpiRows'],['badgeGroupCount','kpiGroups']];
    map.forEach(([src, dst]) => {
      const s = document.getElementById(src), d = document.getElementById(dst);
      if (s && d) d.textContent = (s.textContent || '').replace(/[^0-9]/g, '') || '0';
    });
  }
  window.syncSidebarBadges = syncSidebarBadges;

  function init() {
    document.querySelectorAll('.nav-item').forEach(btn => {
      btn.addEventListener('click', () => showSection(btn.getAttribute('data-section')));
    });
    const toggle = document.getElementById('btnSidebarToggle');
    if (toggle) toggle.addEventListener('click', () => {
      document.getElementById('appSidebar').classList.toggle('collapsed');
      try { localStorage.setItem('qhdc_sidebar_collapsed',
        document.getElementById('appSidebar').classList.contains('collapsed') ? '1' : '0'); } catch (e) {}
    });
    try {
      if (localStorage.getItem('qhdc_sidebar_collapsed') === '1')
        document.getElementById('appSidebar').classList.add('collapsed');
    } catch (e) {}
    // Dong bo badge dinh ky (khong can sua app.js)
    setInterval(syncSidebarBadges, 1200);
    syncSidebarBadges();
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();
})();
```

- [ ] **Step 2: Nạp script** — thêm vào `index.html` ngay sau dòng 1308:

```html
  <script src="ui-nav.js" defer></script>
```

- [ ] **Step 3: Verify — smoke test trình duyệt**

Run: `python server.py` → mở `http://localhost:8080/index.html`.
Kỳ vọng: bấm lần lượt 6 mục sidebar → đúng khu hiển thị; vào "Tổng hợp" bảng báo cáo render; Console không lỗi. Nạp file mẫu `Masterlist 2027-2028_Mau.xlsx` qua khu "Nạp dữ liệu" → KPI Trang chủ cập nhật số dòng/lỗi/nhóm sau ~1-2s.

- [ ] **Step 4: Commit**

```bash
git add ui-nav.js index.html
git commit -m "feat(ui): them lop dieu huong sidebar + dong bo KPI/badge (ui-nav.js)"
```

---

### Task 5: CSS layout — sidebar, app-shell, header, KPI

Style khung mới: sidebar cố định trái, app-shell chừa lề trái, header thu gọn 2 dòng tiêu đề, thẻ KPI. Thêm vào cuối `styles.css`.

**Files:**
- Modify: `styles.css` (thêm khối mới ở cuối file); `index.html` header (dòng ~28-57: đổi tiêu đề 2 dòng, gỡ `style=""` inline trên nút)

**Interfaces:**
- Consumes: token `--sidebar-w`, `--header-h`, màu brand (Task 2); DOM sidebar/section (Task 3).

- [ ] **Step 1: Thêm CSS layout vào cuối `styles.css`**

```css
/* ===== Viettel Sidebar Layout ===== */
.app-sidebar {
  position: fixed; top: 0; left: 0; bottom: 0; width: var(--sidebar-w);
  background: var(--bg-surface); border-right: 1px solid var(--border-subtle);
  display: flex; flex-direction: column; z-index: 40; transition: width .18s ease;
}
.app-sidebar.collapsed { width: var(--sidebar-w-collapsed); }
.app-sidebar.collapsed .sidebar-brand-text,
.app-sidebar.collapsed .nav-item span:not(.nav-badge),
.app-sidebar.collapsed .nav-badge { display: none; }
.sidebar-brand { display: flex; align-items: center; gap: 10px; padding: 14px 16px; height: var(--header-h); border-bottom: 1px solid var(--border-subtle); }
.sidebar-logo { width: 34px; height: 34px; border-radius: 9px; background: var(--primary); display: grid; place-items: center; flex-shrink: 0; }
.sidebar-brand-text { display: flex; flex-direction: column; line-height: 1.15; overflow: hidden; }
.sidebar-brand-title { font-weight: 800; font-size: .92rem; color: var(--text-main); white-space: nowrap; }
.sidebar-brand-sub { font-size: .72rem; color: var(--text-muted); }
.sidebar-toggle { margin-left: auto; background: none; border: none; color: var(--text-muted); cursor: pointer; padding: 4px; border-radius: 6px; }
.sidebar-toggle:hover { background: var(--bg-subtle); color: var(--text-main); }
.sidebar-nav { padding: 10px 10px; display: flex; flex-direction: column; gap: 4px; overflow-y: auto; }
.nav-item { display: flex; align-items: center; gap: 11px; padding: 9px 12px; border-radius: 9px; border: none; background: none; cursor: pointer; color: var(--text-secondary); font-size: .87rem; font-weight: 600; text-align: left; position: relative; width: 100%; }
.nav-item span:not(.nav-badge) { white-space: nowrap; }
.nav-item:hover { background: var(--bg-subtle); color: var(--text-main); }
.nav-item.active { background: var(--primary-soft); color: var(--primary); }
.nav-item.active::before { content: ""; position: absolute; left: 0; top: 8px; bottom: 8px; width: 3px; border-radius: 0 3px 3px 0; background: var(--primary); }
.nav-badge { margin-left: auto; font-size: .68rem; font-weight: 700; padding: 1px 7px; border-radius: 999px; background: var(--bg-muted); color: var(--text-secondary); }
.nav-badge.badge-danger { background: var(--danger-bg); color: var(--danger); }

.app-shell { margin-left: var(--sidebar-w); transition: margin-left .18s ease; }
.app-sidebar.collapsed ~ .app-shell { margin-left: var(--sidebar-w-collapsed); }

/* Section panels */
.section-panel { display: none; }
.section-panel.active { display: block; }

/* KPI cards */
.dashboard-kpis { display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; margin-bottom: 20px; }
.kpi-card { background: var(--bg-card); border: 1px solid var(--border-subtle); border-radius: var(--radius-md); padding: 16px 18px; box-shadow: var(--shadow-card); display: flex; flex-direction: column; gap: 6px; }
.kpi-label { font-size: .78rem; color: var(--text-muted); font-weight: 600; }
.kpi-value { font-size: 1.6rem; font-weight: 800; color: var(--text-main); font-variant-numeric: tabular-nums; }
.kpi-card.kpi-danger .kpi-value { color: var(--danger); }

@media (max-width: 1100px) {
  .app-sidebar { width: var(--sidebar-w-collapsed); }
  .app-sidebar .sidebar-brand-text, .app-sidebar .nav-item span:not(.nav-badge), .app-sidebar .nav-badge { display: none; }
  .app-shell { margin-left: var(--sidebar-w-collapsed); }
  .dashboard-kpis { grid-template-columns: repeat(2, 1fr); }
}
```

- [ ] **Step 2: Header 2 dòng + chuẩn hóa nút.** Trong `index.html` header (dòng ~28-57): đổi khối `brand-title` một dòng thành 2 dòng, và **gỡ mọi `style="..."`** trên các nút. Thay khối brand + nút bằng:

```html
  <header class="app-header">
    <div class="header-container">
      <div class="brand-section">
        <div>
          <div class="brand-title">QHĐC 2027–2028</div>
          <div class="brand-sub">Tổng hợp &amp; So khớp Masterlist</div>
        </div>
      </div>
      <div class="header-actions">
        <button id="btnSaveToServer" class="btn btn-outline" disabled title="Lưu bản hiện tại lên Server để nạp lại chỉnh sửa tiếp">
          <i data-lucide="cloud-upload" class="w-4 h-4"></i><span>Lưu lên Server</span>
        </button>
        <button id="btnOpenResultFile" class="btn btn-primary" disabled title="Tải file kết quả hoàn chỉnh về máy">
          <i data-lucide="download" class="w-4 h-4"></i><span>File tổng hợp</span>
        </button>
        <button id="btnReset" class="btn btn-outline" title="Làm mới toàn bộ dữ liệu">
          <i data-lucide="rotate-ccw" class="w-4 h-4"></i><span>Làm mới</span>
        </button>
      </div>
    </div>
  </header>
```

(Lưu ý: `btnOpenVTBConverter` được chuyển sang trang Công cụ ở Task 6 — không còn ở header.)

- [ ] **Step 3: Thêm CSS header brand-sub + sticky** vào cuối `styles.css`:

```css
.app-header { position: sticky; top: 0; z-index: 30; background: var(--bg-surface); border-bottom: 1px solid var(--border-subtle); height: var(--header-h); }
.app-header .header-container { display: flex; align-items: center; justify-content: space-between; height: 100%; padding: 0 22px; }
.brand-title { font-weight: 800; font-size: 1rem; color: var(--text-main); line-height: 1.15; }
.brand-sub { font-size: .78rem; color: var(--text-muted); }
.header-actions { display: flex; align-items: center; gap: 10px; }
```

- [ ] **Step 4: Verify guard + trình duyệt**

Run: `node tests/check_ui_structure.js` → PASS.
Run: `python server.py` → mở app. Kỳ vọng: sidebar đỏ Viettel bên trái, nội dung chừa lề đúng, header 2 dòng gọn, nút đồng đều; bấm nút thu gọn sidebar chạy; thu nhỏ cửa sổ <1100px sidebar auto thu gọn.

- [ ] **Step 5: Commit**

```bash
git add styles.css index.html
git commit -m "feat(ui): layout sidebar + app-shell + header 2 dong + the KPI trang chu"
```

---

### Task 6: Trang Công cụ (gom chức năng rải rác)

Chuyển các nút công cụ vào `#toolsGrid` dạng thẻ hành động; thêm thẻ VTB và thẻ "Sắp có" cho Tách IP. Giữ nguyên id nút để handler cũ vẫn chạy.

**Files:**
- Modify: `index.html` (`#toolsGrid` trong panel `tools`; toolbar `#tabHierarchy` dòng ~427-450; header cũ chứa `btnOpenVTBConverter`)

**Interfaces:**
- Consumes: nút `btnReindexTT`, `btnFixOutlineGroups`, `btnSplitWalletDV`, `btnSplitWalletMang`, `btnOpenVTBConverter` (di chuyển, giữ id + class + onclick).

- [ ] **Step 1: Điền `#toolsGrid`** trong panel `data-panel="tools"`:

```html
      <div class="tools-grid" id="toolsGrid">
        <div class="tool-card">
          <div class="tool-card-head"><i data-lucide="refresh-cw" class="w-5 h-5"></i><h3>Chuyển đổi form cũ</h3></div>
          <p class="tool-card-desc">Chuyển file đầu vào Mẫu cũ (37 cột) sang Mẫu mới.</p>
          <button id="btnOpenVTBConverter" class="btn btn-primary btn-sm"><i data-lucide="arrow-right" class="w-4 h-4"></i><span>Mở công cụ</span></button>
        </div>
        <div class="tool-card">
          <div class="tool-card-head"><i data-lucide="binary" class="w-5 h-5"></i><h3>Đánh lại chỉ mục</h3></div>
          <p class="tool-card-desc">Tự động đánh lại toàn bộ số thứ tự/chỉ mục theo cây phân cấp.</p>
          <button id="btnReindexTT" class="btn btn-outline btn-sm"><i data-lucide="binary" class="w-4 h-4"></i><span>Cập nhật lại Chỉ mục</span></button>
        </div>
        <div class="tool-card">
          <div class="tool-card-head"><i data-lucide="list-tree" class="w-5 h-5"></i><h3>Fix Group dòng</h3></div>
          <p class="tool-card-desc">Chuẩn hóa Group dòng (Outline Level 1..5) theo cây phân cấp.</p>
          <button id="btnFixOutlineGroups" class="btn btn-outline btn-sm"><i data-lucide="list-tree" class="w-4 h-4"></i><span>Fix Group dòng</span></button>
        </div>
        <div class="tool-card">
          <div class="tool-card-head"><i data-lucide="split" class="w-5 h-5"></i><h3>Tách Ví → DV riêng</h3></div>
          <p class="tool-card-desc">Ví điện tử: Mã mảng CNTT, Mã DV VI, Mã loại VTTB.</p>
          <button id="btnSplitWalletDV" class="btn btn-outline btn-sm btn-wallet-split"><i data-lucide="split" class="w-4 h-4"></i><span>Tách Ví: DV riêng</span></button>
        </div>
        <div class="tool-card">
          <div class="tool-card-head"><i data-lucide="git-branch" class="w-5 h-5"></i><h3>Tách Ví → Mảng riêng</h3></div>
          <p class="tool-card-desc">Ví điện tử: Mã mảng VI, Mã DV TTKD, Mã loại VTTB.</p>
          <button id="btnSplitWalletMang" class="btn btn-outline btn-sm btn-wallet-split"><i data-lucide="git-branch" class="w-4 h-4"></i><span>Tách Ví: Mảng riêng</span></button>
        </div>
        <div class="tool-card tool-card-soon">
          <div class="tool-card-head"><i data-lucide="network" class="w-5 h-5"></i><h3>Tách IP → Mảng riêng</h3></div>
          <p class="tool-card-desc">Tách IP thành mảng riêng.</p>
          <button class="btn btn-outline btn-sm" disabled><span>Sắp có</span></button>
        </div>
      </div>
```

- [ ] **Step 2: Gỡ các nút gốc đã chuyển đi.** Trong toolbar `#tabHierarchy` (dòng ~427-450): xóa 4 nút `btnReindexTT`, `btnFixOutlineGroups`, `btnSplitWalletDV`, `btnSplitWalletMang` (giữ lại `btnRestoreOriginalTT` và ô tìm kiếm `searchGroupInput`). Trong header cũ đã bỏ `btnOpenVTBConverter` ở Task 5. Đảm bảo mỗi id trên **chỉ còn xuất hiện đúng 1 lần** trong file (grep kiểm tra ở Step 4).

- [ ] **Step 3: CSS thẻ công cụ** — thêm cuối `styles.css`:

```css
.tools-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(300px, 1fr)); gap: 16px; }
.tool-card { background: var(--bg-card); border: 1px solid var(--border-subtle); border-radius: var(--radius-md); padding: 18px; box-shadow: var(--shadow-card); display: flex; flex-direction: column; gap: 10px; }
.tool-card-head { display: flex; align-items: center; gap: 10px; color: var(--primary); }
.tool-card-head h3 { font-size: .95rem; font-weight: 700; color: var(--text-main); }
.tool-card-desc { font-size: .82rem; color: var(--text-muted); flex: 1; }
.tool-card .btn { align-self: flex-start; }
.tool-card-soon { opacity: .7; }
```

- [ ] **Step 4: Verify — mỗi id duy nhất + guard + smoke**

Run: `node tests/check_ui_structure.js` → PASS.
Run (đảm bảo không trùng id): `grep -oE 'id="(btnReindexTT|btnFixOutlineGroups|btnSplitWalletDV|btnSplitWalletMang|btnOpenVTBConverter)"' index.html | sort | uniq -c`
Expected: mỗi id đếm đúng `1`.
Run: `python server.py` → vào Công cụ, bấm từng nút: "Đánh lại chỉ mục"/"Fix Group"/"Tách Ví" hiện toast tương ứng (hoặc cảnh báo "chưa có dữ liệu" nếu chưa nạp), "Chuyển đổi form cũ" mở modal VTB; thẻ "Tách IP" disabled.

- [ ] **Step 5: Commit**

```bash
git add index.html styles.css
git commit -m "feat(ui): trang Cong cu gom cac chuc nang rai rac + placeholder Tach IP"
```

---

### Task 7: Reskin component — tab/sub-tab, bảng, nút, badge, upload, modal

Đồng bộ các component còn lại theo hệ token: sub-tab underline, bảng header sticky/zebra, bộ nút thống nhất, ô upload, modal.

**Files:**
- Modify: `styles.css` (các khối component tương ứng)

**Interfaces:**
- Consumes: token Task 2; DOM Task 3/6.

- [ ] **Step 1: Sub-tab underline** — cập nhật `.report-subtabs-nav` / `.report-subtab-btn` trong `styles.css`:

```css
.report-subtabs-nav { display: flex; gap: 4px; border-bottom: 1px solid var(--border-subtle); margin-bottom: 1.25rem; }
.report-subtab-btn { display: inline-flex; align-items: center; gap: 7px; padding: 10px 14px; border: none; background: none; cursor: pointer; color: var(--text-muted); font-weight: 600; font-size: .86rem; border-bottom: 2px solid transparent; margin-bottom: -1px; }
.report-subtab-btn:hover { color: var(--text-main); }
.report-subtab-btn.active { color: var(--primary); border-bottom-color: var(--primary); }
```

- [ ] **Step 2: Bảng** — cập nhật `.data-table` trong `styles.css`:

```css
.data-table { width: 100%; border-collapse: collapse; font-size: .82rem; }
.data-table thead th { position: sticky; top: 0; z-index: 2; background: var(--bg-subtle); color: var(--text-secondary); font-weight: 700; text-align: left; padding: 10px 12px; border-bottom: 1px solid var(--border-strong); white-space: nowrap; }
.data-table tbody td { padding: 8px 12px; border-bottom: 1px solid var(--border-subtle); color: var(--text-main); }
.data-table tbody tr:nth-child(even) { background: #FCFCFD; }
.data-table tbody tr:hover { background: var(--primary-soft); }
.data-table .num-cell { text-align: right; font-family: 'JetBrains Mono', monospace; font-variant-numeric: tabular-nums; }
```

(Giữ nguyên các rule đặc thù subtotal/strategy đã có; chỉ cập nhật rule chung ở trên.)

- [ ] **Step 3: Bộ nút thống nhất** — cập nhật `.btn` và biến thể trong `styles.css`:

```css
.btn { display: inline-flex; align-items: center; gap: 7px; padding: 8px 14px; border-radius: var(--radius-sm); font-size: .84rem; font-weight: 600; cursor: pointer; border: 1px solid transparent; transition: background .12s, border-color .12s, box-shadow .12s; line-height: 1; }
.btn:focus-visible { outline: none; box-shadow: 0 0 0 3px var(--primary-glow); }
.btn-sm { padding: 6px 11px; font-size: .8rem; }
.btn-primary { background: var(--primary); color: #fff; }
.btn-primary:hover:not(:disabled) { background: var(--primary-hover); }
.btn-outline { background: var(--bg-surface); border-color: var(--border-strong); color: var(--text-secondary); }
.btn-outline:hover:not(:disabled) { background: var(--bg-subtle); color: var(--text-main); }
.btn:disabled { opacity: .5; cursor: not-allowed; }
```

(Giữ `.btn-success`/`.btn-warning`/`.btn-light` hiện có nhưng cho kế thừa padding/radius mới; nếu chúng hardcode màu tím thì đổi sang token brand tương ứng.)

- [ ] **Step 4: Badge tab & modal** — cập nhật `styles.css`:

```css
.tab-badge, .table-badge-count { font-size: .7rem; font-weight: 700; padding: 2px 8px; border-radius: 999px; background: var(--bg-muted); color: var(--text-secondary); }
.tab-badge.badge-danger { background: var(--danger-bg); color: var(--danger); }
.tab-badge.badge-info { background: var(--info-bg); color: var(--info); }
.tab-badge.badge-warning { background: var(--warning-bg); color: var(--warning); }
.modal { border-radius: var(--radius-lg); box-shadow: var(--shadow-dropdown); }
.modal-overlay { background: rgba(16,24,40,0.45); }
```

- [ ] **Step 5: Verify — guard + smoke toàn diện**

Run: `node tests/check_ui_structure.js` → PASS.
Run: `python server.py`; nạp file mẫu; duyệt cả 6 khu; mở modal VTB & "Bản lưu Server"; đổi sub-tab trong Tổng hợp; cuộn bảng (header dính). Kỳ vọng: đồng bộ tông đỏ, không vỡ layout, Console không lỗi.

- [ ] **Step 6: Commit**

```bash
git add styles.css
git commit -m "style(ui): reskin sub-tab underline, bang, nut, badge, modal theo Viettel"
```

---

### Task 8: Đồng bộ `report.css` + dọn màu tím sót & smoke test cuối

Đồng bộ trang báo cáo và quét sạch màu indigo/tím còn hardcode; smoke test toàn luồng.

**Files:**
- Modify: `report.css`; `index.html` (banner `#tabStrategy` dòng ~491-535 nếu còn gradient tím inline); rà `styles.css`

**Interfaces:**
- Consumes: token brand.

- [ ] **Step 1: Quét màu tím còn sót**

Run: `grep -nE '#4338ca|#4f46e5|#6366f1|#3730a3|indigo' index.html styles.css report.css`
Với mỗi kết quả: đổi sang token brand phù hợp (`var(--primary)` / `var(--primary-hover)` / `var(--primary-soft)`), gồm cả gradient inline `strat-title-icon` (dòng ~493) và badge inline `style="background:#4338ca"` (dòng ~205 — nếu còn sau khi gỡ tabs-nav thì bỏ qua).

- [ ] **Step 2: Đồng bộ tokens trong `report.css`** — nếu `report.css` khai báo biến màu riêng, trỏ chúng về cùng giá trị brand (primary `#EE0033`, nền `#F5F6F8`, viền `#E5E7EB`); áp quy tắc bảng như Task 7 Step 2 cho bảng báo cáo.

- [ ] **Step 3: Verify — không còn tím + guard**

Run: `grep -nE '#4338ca|#4f46e5|#6366f1|#3730a3' index.html styles.css report.css`
Expected: không kết quả (exit 1 / rỗng).
Run: `node tests/check_ui_structure.js` → PASS.

- [ ] **Step 4: Smoke test toàn luồng (bắt buộc)**

Run: `python server.py` → `http://localhost:8080/index.html`. Thực hiện: nạp `Masterlist 2027-2028_Mau.xlsx` (khu Nạp dữ liệu) → Trang chủ hiện bảng + KPI → Kiểm tra dữ liệu lọc lỗi & sao chép → Công cụ chạy "Đánh lại chỉ mục" → Cấu hình xem cây nhóm → Tổng hợp đổi 4 sub-tab + Xuất Excel → nút "File tổng hợp" tải file. Kỳ vọng: mọi thao tác hoạt động như trước, giao diện đồng bộ Viettel, Console không lỗi.

- [ ] **Step 5: Commit**

```bash
git add report.css index.html styles.css
git commit -m "style(ui): dong bo report.css + don sach mau tim con sot"
```

---

### Task 9: Cập nhật README + hoàn tất nhánh

**Files:**
- Modify: `README.md` (mô tả giao diện mới)

- [ ] **Step 1: Cập nhật `README.md`** — đổi dòng mô tả `styles.css` (hiện: "Hệ thống giao diện Modern SaaS Light Mode") thành mô tả bộ nhận diện Viettel + điều hướng sidebar 6 khu; nếu README liệt kê 4 tab, cập nhật thành 6 khu sidebar.

- [ ] **Step 2: Verify build sạch**

Run: `node tests/check_ui_structure.js` → PASS.
Run: `node tests/vtb_converter.test.js` (nếu chạy được) → đảm bảo test converter cũ vẫn PASS (không bị ảnh hưởng).

- [ ] **Step 3: Commit & tổng kết nhánh**

```bash
git add README.md
git commit -m "docs: cap nhat README theo giao dien Viettel + sidebar 6 khu"
git log --oneline feature/viettel-ui-redesign ^main
```

Sau đó dùng skill `superpowers:finishing-a-development-branch` để chọn merge/PR.
