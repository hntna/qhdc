// ui-nav.js — Lop dieu huong sidebar (thuan trinh bay, khong dung logic nghiep vu)
(function () {
  function showSection(key) {
    document.querySelectorAll('.section-panel').forEach(p => {
      p.classList.toggle('active', p.getAttribute('data-panel') === key);
    });
    document.querySelectorAll('.nav-item').forEach(n => {
      n.classList.toggle('active', n.getAttribute('data-section') === key);
    });
    // Bao cao Tong hop hien nam trong khu "home" (Trang chu)
    if (key === 'home') {
      try {
        if (typeof updateReportFromExtractedData === 'function' && typeof state !== 'undefined' && state) {
          updateReportFromExtractedData(state.extractedData, state.extractedByMang);
        } else if (typeof window.renderAllReportTabs === 'function') {
          window.renderAllReportTabs();
        }
      } catch (e) { console.warn('render report on nav:', e); }
    }
    if (key === 'config') {
      try {
        const activeConfigTab = document.querySelector('.config-tab-btn.active');
        if (activeConfigTab && activeConfigTab.getAttribute('data-config-tab') === 'configTabStrategyProfile') {
          if (typeof renderStrategyComparisonTable === 'function') renderStrategyComparisonTable();
        }
      } catch (e) { console.warn('render config on nav:', e); }
    }
    if (key === 'admin-workspaces') {
      try {
        if (typeof window.renderAdminWorkspacesTable === 'function') {
          window.renderAdminWorkspacesTable();
        }
      } catch (e) { console.warn('render admin-workspaces on nav:', e); }
    }
    if (key === 'admin-users') {
      try {
        if (typeof window.renderAdminUsersTable === 'function') {
          window.renderAdminUsersTable();
        }
      } catch (e) { console.warn('render admin-users on nav:', e); }
    }
    if (window.lucide) setTimeout(() => lucide.createIcons(), 10);
  }

  // Chuyen tab trong Menu Cau hinh (Cau hinh nhom vs Cau hinh Profile Chien luoc)
  function switchConfigTab(tabId) {
    document.querySelectorAll('.config-tab-btn').forEach(btn => {
      btn.classList.toggle('active', btn.getAttribute('data-config-tab') === tabId);
    });
    document.querySelectorAll('.config-tab-content').forEach(content => {
      const isActive = content.id === tabId;
      content.style.display = isActive ? 'block' : 'none';
      content.classList.toggle('active', isActive);
    });
    if (tabId === 'configTabStrategyProfile') {
      if (typeof renderStrategyComparisonTable === 'function') {
        renderStrategyComparisonTable();
      }
    }
    if (window.lucide) setTimeout(() => lucide.createIcons(), 10);
  }
  window.switchConfigTab = switchConfigTab;

  window.showSection = showSection;
  window.switchSection = function (key) {
    if (key === 'tabConfig' || key === 'config') showSection('config');
    else showSection(key);
  };

  window.openConfigStrategyProfile = function () {
    showSection('config');
    switchConfigTab('configTabStrategyProfile');
  };

  // Chon menu con cua Trang chu (dong bo voi sub-tab bao cao)
  function selectHomeSubtab(subtabId) {
    if (!subtabId) return;
    showSection('home');
    document.querySelectorAll('.nav-subitem').forEach(n => {
      n.classList.toggle('active', n.getAttribute('data-subnav') === subtabId);
    });
    if (typeof switchReportSubtab === 'function') switchReportSubtab(subtabId);
    if (window.lucide) setTimeout(() => lucide.createIcons(), 10);
  }

  function syncSidebarBadges() {
    const map = [['badgeValidationCount','navBadgeValidation'],['badgeValidationCount','kpiErrors'],
                 ['badgeRowCount','kpiRows'],['badgeGroupCount','kpiGroups']];
    map.forEach(([src, dst]) => {
      const s = document.getElementById(src), d = document.getElementById(dst);
      if (s && d) {
        const raw = (s.textContent || '').replace(/[^0-9]/g, '');
        const val = parseInt(raw || '0', 10);
        d.textContent = val;
        if (dst === 'navBadgeValidation') {
          if (val > 0) {
            d.style.display = 'inline-flex';
            d.className = 'nav-badge badge-danger';
          } else {
            d.style.display = 'none';
          }
        }
      }
    });
    var sectorsEl = document.getElementById('kpiSectors');
    if (sectorsEl) {
      var loaded = document.querySelectorAll('.upload-box.has-file').length;
      sectorsEl.textContent = loaded + '/7';
    }
  }
  window.syncSidebarBadges = syncSidebarBadges;

  function init() {
    document.querySelectorAll('.nav-item').forEach(btn => {
      btn.addEventListener('click', () => showSection(btn.getAttribute('data-section')));
    });
    document.querySelectorAll('.nav-subitem').forEach(btn => {
      btn.addEventListener('click', () => selectHomeSubtab(btn.getAttribute('data-subnav')));
    });
    const toggle = document.getElementById('btnSidebarToggle');
    if (toggle) toggle.addEventListener('click', (e) => {
      e.stopPropagation();
      document.getElementById('appSidebar').classList.toggle('collapsed');
      if (window.lucide) setTimeout(() => lucide.createIcons(), 50);
      try { localStorage.setItem('qhdc_sidebar_collapsed',
        document.getElementById('appSidebar').classList.contains('collapsed') ? '1' : '0'); } catch (e) {}
    });
    const brand = document.querySelector('.sidebar-brand');
    if (brand) brand.addEventListener('click', () => {
      const sidebar = document.getElementById('appSidebar');
      if (sidebar && sidebar.classList.contains('collapsed')) {
        sidebar.classList.remove('collapsed');
        if (window.lucide) setTimeout(() => lucide.createIcons(), 50);
        try { localStorage.setItem('qhdc_sidebar_collapsed', '0'); } catch (e) {}
      }
    });
    try {
      if (localStorage.getItem('qhdc_sidebar_collapsed') === '1') {
        document.getElementById('appSidebar').classList.add('collapsed');
        if (window.lucide) setTimeout(() => lucide.createIcons(), 50);
      }
    } catch (e) {}
    // Dong bo badge dinh ky (khong can sua app.js)
    setInterval(syncSidebarBadges, 1200);
    syncSidebarBadges();
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();
})();
