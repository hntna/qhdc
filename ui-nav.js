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
