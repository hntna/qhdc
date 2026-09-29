// auth-ui.js - Điều khiển giao diện Đăng nhập, Quản lý User (Admin), Workspaces & Phân quyền
(function () {
  'use strict';

  // Khởi tạo ngay khi DOM sẵn sàng
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initAuthUI);
  } else {
    initAuthUI();
  }

  function initAuthUI() {
    setupStaticEventListeners();

    const service = window.FirebaseService;
    if (service && typeof service.onAuthStateChange === 'function') {
      service.onAuthStateChange(handleAuthStateChange);
    }

    // Tải danh sách workspace ban đầu
    loadWorkspacesDropdown();

    // Tự động đồng bộ số thông báo chuông
    setInterval(syncNotifyBellCount, 1000);
  }

  // ==================== GẮN SỰ KIỆN GIAO DIỆN CỐ ĐỊNH ====================

  function setupStaticEventListeners() {
    // 1. Toggle Sidebar từ Header (Ảnh 3)
    const btnHeaderSidebarToggle = document.getElementById('btnHeaderSidebarToggle');
    if (btnHeaderSidebarToggle) {
      btnHeaderSidebarToggle.addEventListener('click', () => {
        const sidebar = document.getElementById('appSidebar');
        if (sidebar) {
          sidebar.classList.toggle('collapsed');
          if (window.lucide) setTimeout(() => lucide.createIcons(), 50);
          try { localStorage.setItem('qhdc_sidebar_collapsed', sidebar.classList.contains('collapsed') ? '1' : '0'); } catch (e) {}
        }
      });
    }

    // 2. User Profile Dropdown Toggle (Click vào Avatar, Tên hoặc Mũi tên)
    const userTrigger = document.getElementById('userProfileHeaderTrigger');
    const userDropdown = document.getElementById('userProfileDropdown');
    if (userTrigger && userDropdown) {
      userTrigger.addEventListener('click', (e) => {
        e.stopPropagation();
        const isHidden = userDropdown.style.display === 'none' || !userDropdown.style.display;
        userDropdown.style.display = isHidden ? 'block' : 'none';
      });

      // Click ra ngoài để đóng dropdown
      document.addEventListener('click', (e) => {
        if (!userDropdown.contains(e.target) && !userTrigger.contains(e.target)) {
          userDropdown.style.display = 'none';
        }
      });
    }

    // 3. Các nút trong User Profile Dropdown
    const btnDropChangePass = document.getElementById('btnDropChangePassword');
    if (btnDropChangePass) {
      btnDropChangePass.addEventListener('click', () => {
        if (userDropdown) userDropdown.style.display = 'none';
        openChangePasswordModal();
      });
    }

    const btnDropAdminWs = document.getElementById('btnDropAdminWorkspaces');
    if (btnDropAdminWs) {
      btnDropAdminWs.addEventListener('click', () => {
        if (userDropdown) userDropdown.style.display = 'none';
        const navItemAdminWs = document.getElementById('navItemAdminWorkspaces');
        if (navItemAdminWs) navItemAdminWs.click();
      });
    }

    const btnDropAdminPanel = document.getElementById('btnDropAdminPanel');
    if (btnDropAdminPanel) {
      btnDropAdminPanel.addEventListener('click', () => {
        if (userDropdown) userDropdown.style.display = 'none';
        const navItemAdmin = document.getElementById('navItemAdminUsers');
        if (navItemAdmin) navItemAdmin.click();
      });
    }

    const formAdminCreateWs = document.getElementById('formAdminCreateWs');
    if (formAdminCreateWs) {
      formAdminCreateWs.addEventListener('submit', handleAdminCreateWorkspaceSubmit);
    }

    const btnDropLogout = document.getElementById('btnDropLogout');
    if (btnDropLogout) {
      btnDropLogout.addEventListener('click', () => {
        if (userDropdown) userDropdown.style.display = 'none';
        handleLogout();
      });
    }

    // 4. Nút Đăng nhập cho Guest
    const btnHeaderLogin = document.getElementById('btnHeaderLogin');
    if (btnHeaderLogin) {
      btnHeaderLogin.addEventListener('click', (e) => {
        e.stopPropagation();
        openLoginModal();
      });
    }

    const btnLegacyLogin = document.getElementById('btnAuthLogin');
    if (btnLegacyLogin) {
      btnLegacyLogin.addEventListener('click', openLoginModal);
    }

    // 5. Form đăng nhập & đổi mật khẩu
    const formLogin = document.getElementById('formLogin');
    if (formLogin) formLogin.addEventListener('submit', handleLoginSubmit);

    const formChangePass = document.getElementById('formChangePassword');
    if (formChangePass) formChangePass.addEventListener('submit', handleChangePasswordSubmit);

    // 6. Workspace selector & nút tạo workspace
    const selectWs = document.getElementById('selectWorkspace');
    if (selectWs) selectWs.addEventListener('change', handleWorkspaceChange);

    const formNewWs = document.getElementById('formNewWorkspace');
    if (formNewWs) formNewWs.addEventListener('submit', handleNewWorkspaceSubmit);

    // 7. Form tạo User mới trong Admin Panel
    const formAdminCreate = document.getElementById('formAdminCreateUser');
    if (formAdminCreate) formAdminCreate.addEventListener('submit', handleAdminCreateUserSubmit);

    // 8. Đồng bộ class active cho các chip chọn mảng phân quyền
    function syncChipState(chk) {
      if (!chk) return;
      const parentLabel = chk.closest('.chip-sector, .chip');
      if (parentLabel) {
        if (chk.checked) parentLabel.classList.add('active');
        else parentLabel.classList.remove('active');
      }
    }

    document.addEventListener('change', (e) => {
      if (e.target && (e.target.classList.contains('admin-sector-chk') || e.target.classList.contains('edit-sector-chk'))) {
        syncChipState(e.target);
      }
    });

    document.addEventListener('click', (e) => {
      const chip = e.target.closest('.chip-sector');
      if (chip) {
        const chk = chip.querySelector('input[type="checkbox"]');
        if (chk && e.target !== chk) {
          setTimeout(() => syncChipState(chk), 0);
        }
      }
    });
  }

  function getInitials(name) {
    if (!name) return 'U';
    const parts = name.trim().split(/\s+/);
    if (parts.length === 1) return parts[0].substring(0, 2).toUpperCase();
    return (parts[parts.length - 2][0] + parts[parts.length - 1][0]).toUpperCase();
  }

  function syncNotifyBellCount() {
    const valBadge = document.getElementById('badgeValidationCount');
    const notifyEl = document.getElementById('headerNotifyCount');
    if (valBadge && notifyEl) {
      const val = (valBadge.textContent || '0').replace(/[^0-9]/g, '') || '0';
      notifyEl.textContent = val;
    }
  }

  // ==================== XỬ LÝ AUTH STATE ====================

  function handleAuthStateChange(profile) {
    const service = window.FirebaseService;
    const userTrigger = document.getElementById('userProfileHeaderTrigger');
    const guestBlock = document.getElementById('guestHeaderBlock');
    const userDropdown = document.getElementById('userProfileDropdown');
    const avatarCircle = document.getElementById('userAvatarCircle');
    const headerUserName = document.getElementById('headerUserName');
    const headerUserDept = document.getElementById('headerUserDept');
    const dropUserFullName = document.getElementById('dropUserFullName');
    const dropUserRole = document.getElementById('dropUserRole');
    const dropUserSectors = document.getElementById('dropUserSectors');
    const btnDropAdmin = document.getElementById('btnDropAdminPanel');
    const navItemAdmin = document.getElementById('navItemAdminUsers');
    const btnSaveServer = document.getElementById('btnSaveToServer');

    if (profile) {
      // Đã đăng nhập (User / Admin)
      const isAdm = profile.role === 'admin';
      const sectorText = isAdm
        ? 'Toàn quyền (7 mảng)'
        : (profile.assignedSectors && profile.assignedSectors.length ? profile.assignedSectors.join(', ') : 'Chưa gán mảng');
      const displayName = profile.displayName || profile.username;

      if (userTrigger) userTrigger.style.display = 'flex';
      if (guestBlock) guestBlock.style.display = 'none';

      if (avatarCircle) {
        avatarCircle.textContent = getInitials(displayName);
        avatarCircle.style.background = isAdm ? '#e11d48' : '#2563eb';
      }
      if (headerUserName) headerUserName.textContent = displayName;
      if (headerUserDept) headerUserDept.textContent = isAdm ? 'Quản trị viên / TT, KTTC' : (profile.dept || 'TT, KTTC');

      if (dropUserFullName) dropUserFullName.textContent = displayName;
      if (dropUserRole) dropUserRole.textContent = isAdm ? 'Quản trị viên Hệ thống' : 'Chuyên viên Kế hoạch';
      if (dropUserSectors) dropUserSectors.textContent = `Mảng phụ trách: ${sectorText}`;
      const navItemAdminWs = document.getElementById('navItemAdminWorkspaces');
      const btnDropAdminWs = document.getElementById('btnDropAdminWorkspaces');

      if (btnDropAdminWs) btnDropAdminWs.style.display = 'flex';
      if (btnDropAdmin) btnDropAdmin.style.display = isAdm ? 'flex' : 'none';

      // Quản lý dự án: user thường được quản lý danh mục mình tạo, admin quản lý tất cả.
      if (navItemAdminWs) navItemAdminWs.style.display = 'flex';
      if (navItemAdmin) navItemAdmin.style.display = isAdm ? 'flex' : 'none';

      if (btnSaveServer) {
        btnSaveServer.disabled = false;
        btnSaveServer.style.opacity = '1';
      }

      renderAdminWorkspacesTable();

      // Nạp danh sách user nếu là admin
      if (isAdm) {
        renderAdminUsersTable();
      }
    } else {
      // Khách vãng lai (Guest)
      if (userTrigger) userTrigger.style.display = 'none';
      if (userDropdown) userDropdown.style.display = 'none';
      if (guestBlock) guestBlock.style.display = 'flex';

      const navItemAdminWs = document.getElementById('navItemAdminWorkspaces');
      const btnDropAdminWs = document.getElementById('btnDropAdminWorkspaces');

      if (navItemAdminWs) {
        navItemAdminWs.style.display = 'none';
        if (navItemAdminWs.classList.contains('active')) {
          const homeTab = document.querySelector('.nav-item[data-section="home"]');
          if (homeTab) homeTab.click();
        }
      }
      if (navItemAdmin) {
        navItemAdmin.style.display = 'none';
        if (navItemAdmin.classList.contains('active')) {
          const homeTab = document.querySelector('.nav-item[data-section="home"]');
          if (homeTab) homeTab.click();
        }
      }
      if (btnDropAdminWs) btnDropAdminWs.style.display = 'none';
      if (btnSaveServer) {
        btnSaveServer.disabled = true;
        btnSaveServer.style.opacity = '0.6';
      }
    }

    // Kiểm soát quyền xem Chiến lược 5 năm cho Khách (Guest)
    const isGuest = !profile;
    const cardStrat = document.querySelector('.kpi-card-strat');
    if (cardStrat) cardStrat.style.display = isGuest ? 'none' : 'flex';

    const subtabStratBtn = document.querySelector('.report-subtab-btn[data-subtab="subtabStratDetail"]');
    if (subtabStratBtn) subtabStratBtn.style.display = isGuest ? 'none' : 'inline-flex';

    const configStratTabBtn = document.querySelector('.config-tab-btn[data-config-tab="configTabStrategyProfile"]');
    if (configStratTabBtn) configStratTabBtn.style.display = isGuest ? 'none' : 'inline-flex';

    if (isGuest && typeof switchReportSubtab === 'function') {
      const activeSubtab = document.querySelector('.report-subtab-btn.active');
      if (activeSubtab && activeSubtab.getAttribute('data-subtab') === 'subtabStratDetail') {
        switchReportSubtab('subtabTongHop');
      }
    }

    if (typeof renderAllTabs === 'function') {
      renderAllTabs();
    }

    // Cập nhật quyền trên 7 ô nạp thẻ
    applySectorPermissions();

    // Tải lại danh sách dự án phù hợp với quyền mới
    loadWorkspacesDropdown();

    if (window.lucide) {
      window.lucide.createIcons();
    }
  }

  // ==================== PHÂN QUYỀN TRÊN 7 CARD MẢNG ====================

  function applySectorPermissions() {
    const service = window.FirebaseService;
    if (!service) return;

    const isGst = service.isGuest();
    const isAdm = service.isAdmin();
    const sectors = service.ALL_SECTORS;
    const wsId = document.getElementById('selectWorkspace')?.value;

    sectors.forEach(sec => {
      const box = document.getElementById(`box_${sec}`);
      const input = document.getElementById(`input_${sec}`);
      if (!box) return;

      const hasAccess = isAdm || (!isGst && service.canAccessSector(sec, wsId));

      let badgeLock = box.querySelector('.sector-lock-badge');
      if (!hasAccess) {
        box.classList.add('sector-locked');
        box.style.opacity = '0.65';
        box.style.filter = 'grayscale(0.3)';
        if (input) input.disabled = true;

        let lockMsg = '🔒 Chưa được phân quyền mảng này';
        if (isGst) {
          lockMsg = '🔒 Chỉ xem (Chưa đăng nhập)';
        } else if (wsId) {
          const user = service.getCurrentUser();
          const assignedWs = user?.assignedWorkspaces || [];
          if (!assignedWs.includes(wsId)) {
            lockMsg = '🔒 Dự án chỉ xem (Chưa gán quyền dự án)';
          }
        }

        if (!badgeLock) {
          badgeLock = document.createElement('div');
          badgeLock.className = 'sector-lock-badge';
          badgeLock.style.cssText = `
            font-size: 0.72rem;
            font-weight: 700;
            color: #b91c1c;
            background: #fee2e2;
            padding: 3px 8px;
            border-radius: 9999px;
            display: inline-flex;
            align-items: center;
            gap: 4px;
            margin-top: 6px;
          `;
          badgeLock.innerHTML = lockMsg;
          box.appendChild(badgeLock);
        } else {
          badgeLock.innerHTML = lockMsg;
        }

        box.onclick = function (e) {
          e.stopPropagation();
          e.preventDefault();
          showToast(lockMsg, 'warning');
          return false;
        };
      } else {
        box.classList.remove('sector-locked');
        box.style.opacity = '1';
        box.style.filter = 'none';
        if (input) input.disabled = false;
        if (badgeLock) badgeLock.remove();

        box.onclick = function () {
          if (input) input.click();
        };
      }
    });

    // Nút nạp nhiều file
    const btnLoadMulti = document.getElementById('btnLoadMultiFile');
    if (btnLoadMulti) {
      if (isGst) {
        btnLoadMulti.disabled = true;
        btnLoadMulti.title = 'Vui lòng đăng nhập để nạp file';
      } else {
        btnLoadMulti.disabled = false;
        btnLoadMulti.title = 'Chọn nhiều file hoặc kéo thả nhiều file vào đây';
      }
    }
  }

  window.applySectorPermissions = applySectorPermissions;

  // ==================== WORKSPACE (DANH MỤC / DỰ ÁN) ====================

  async function loadWorkspacesDropdown(selectedId = null) {
    const service = window.FirebaseService;
    if (!service) return;

    const select = document.getElementById('selectWorkspace');
    if (!select) return;

    try {
      const list = await service.listWorkspaces();
      select.innerHTML = '';

      if (list.length === 0) {
        select.innerHTML = '<option value="">(Không có danh mục nào)</option>';
        return;
      }

      list.forEach(ws => {
        const opt = document.createElement('option');
        opt.value = ws.id;
        const visibility = service.getWorkspaceVisibility ? service.getWorkspaceVisibility(ws) : (ws.isPublic ? 'public' : (ws.isShared ? 'shared' : 'private'));
        opt.textContent = ws.name;
        opt.dataset.visibility = visibility;
        opt.dataset.creator = ws.createdBy || '';
        opt.dataset.creatorUid = ws.creatorUid || '';
        select.appendChild(opt);
      });

      const savedWsId = (function () { try { return localStorage.getItem('qhdc_active_workspace_id'); } catch (e) { return null; } })();
      const targetId = selectedId || savedWsId;

      if (targetId && list.some(w => w.id === targetId)) {
        select.value = targetId;
      } else if (list.length > 0) {
        select.value = list[0].id;
      }

      try { localStorage.setItem('qhdc_active_workspace_id', select.value); } catch (e) {}

      updateWorkspaceBadgeText();
      applySectorPermissions();
      if (select.value && window.syncStrategyProfileWithActiveWorkspace) {
        window.syncStrategyProfileWithActiveWorkspace(select.value);
      }
      // Tải dữ liệu workspace được chọn
      if (select.value) {
        loadSelectedWorkspaceData(select.value);
      }
    } catch (err) {
      console.error('Lỗi nạp workspaces dropdown:', err);
    }
  }

  function updateWorkspaceBadgeText() {
    const select = document.getElementById('selectWorkspace');
    const badgeText = document.getElementById('currentWorkspaceNameText');
    if (select && badgeText) {
      const opt = select.options[select.selectedIndex];
      if (opt) badgeText.textContent = opt.textContent;
    }
  }

  async function handleWorkspaceChange(e) {
    const wsId = e.target.value;
    try { localStorage.setItem('qhdc_active_workspace_id', wsId); } catch (e) {}
    updateWorkspaceBadgeText();
    applySectorPermissions();
    if (wsId && window.syncStrategyProfileWithActiveWorkspace) {
      window.syncStrategyProfileWithActiveWorkspace(wsId);
    }
    if (wsId) {
      await loadSelectedWorkspaceData(wsId);
    }
    renderAdminWorkspacesTable();
  }

  async function loadSelectedWorkspaceData(workspaceId) {
    const service = window.FirebaseService;
    if (!service) return;

    try {
      const sectorsData = (await service.loadWorkspaceDataFromFirestore(workspaceId)) || {};

      const appState = window.state || (typeof state !== 'undefined' ? state : null);
      if (appState) {
        appState.exportBlob = null; // Luôn xóa cache kết quả xuất khi tải dữ liệu dự án mới
        appState.extractedByMang = appState.extractedByMang || {};
        appState.files = appState.files || {};
        let totalLoadedRows = 0;
        const loadedKeys = [];

        const allSectors = service.ALL_SECTORS || ['VT', 'ML', 'CDBR', 'CNTT', 'TD', 'CD', 'HT'];

        // Reset hoặc cập nhật từng mảng trong 7 mảng
        allSectors.forEach(secKey => {
          const item = sectorsData[secKey];
          if (item && item.rows && item.rows.length) {
            appState.extractedByMang[secKey] = item.rows;
            totalLoadedRows += item.rows.length;
            loadedKeys.push(secKey);

            // Cập nhật giao diện ô upload mảng
            if (typeof window.updateUploadBoxUI === 'function') {
              window.updateUploadBoxUI(secKey, `${item.rowCount || item.rows.length} dòng (${item.updatedByName || 'Server'})`, (item.rows.length * 150));
            } else {
              const nameEl = document.getElementById(`name_${secKey}`);
              if (nameEl) {
                nameEl.textContent = `${item.rowCount || item.rows.length} dòng (${item.updatedByName || 'Server'})`;
                nameEl.style.color = '#059669';
              }
              const box = document.getElementById(`box_${secKey}`);
              if (box) box.classList.add('has-file');
            }
          } else {
            // Không có dữ liệu trong workspace này -> reset trắng mảng này
            appState.extractedByMang[secKey] = [];
            appState.files[secKey] = null;
            if (typeof window.resetUploadBoxUI === 'function') {
              window.resetUploadBoxUI(secKey);
            } else {
              const nameEl = document.getElementById(`name_${secKey}`);
              if (nameEl) {
                nameEl.textContent = 'Chưa có file';
                nameEl.style.color = 'var(--text-muted)';
              }
              const box = document.getElementById(`box_${secKey}`);
              if (box) box.classList.remove('has-file');
            }
          }
        });

        // Tái tạo extractedData tổng hợp
        appState.extractedData = [];
        allSectors.forEach(k => {
          if (appState.extractedByMang[k] && appState.extractedByMang[k].length) {
            appState.extractedData.push(...appState.extractedByMang[k]);
          }
        });

        // Luôn luôn rebuild và vẽ lại tất cả các bảng (kể cả khi rỗng 0 dòng)
        if (typeof window.rebuildExtractedData === 'function') {
          window.rebuildExtractedData();
        } else if (typeof window.processLoadedData === 'function') {
          window.processLoadedData();
        } else if (typeof window.rebuildAll === 'function') {
          window.rebuildAll();
        }

        if (typeof updateReportFromExtractedData === 'function') {
          updateReportFromExtractedData(appState.extractedData, appState.extractedByMang);
        }
        if (typeof renderAllTabs === 'function') {
          renderAllTabs();
        }

        if (totalLoadedRows > 0) {
          showToast(`Đã nạp ${totalLoadedRows} dòng từ ${loadedKeys.length} mảng (${loadedKeys.join(', ')}) của dự án`, 'success');
        } else {
          showToast(`Dự án hiện chưa có dữ liệu mảng nào.`, 'info');
        }
      }
    } catch (err) {
      console.error('Lỗi tải dữ liệu workspace:', err);
    }
  }

  function openNewWorkspaceModal() {
    const service = window.FirebaseService;
    if (service && service.isGuest()) {
      showToast('Vui lòng đăng nhập để tạo danh mục mới!', 'warning');
      openLoginModal();
      return;
    }

    const modal = document.getElementById('modalNewWorkspace');
    const inputName = document.getElementById('inputNewWsName');
    if (inputName) inputName.value = '';
    const sharedRadio = document.getElementById('newWsVisShared');
    if (sharedRadio) sharedRadio.checked = true;

    // Nạp danh sách Strategy Profiles
    const selProfile = document.getElementById('inputNewWsStrategyProfile');
    if (selProfile) {
      const profiles = (window.state && window.state.strategyProfiles) ? window.state.strategyProfiles : [];
      selProfile.innerHTML = profiles.map(p => `
        <option value="${p.id}" ${p.id === 'profile_default' ? 'selected' : ''}>
          ${escapeHtml(p.name)}
        </option>
      `).join('');
    }

    if (modal) modal.style.display = 'flex';
  }

  window.closeNewWorkspaceModal = function () {
    const modal = document.getElementById('modalNewWorkspace');
    if (modal) modal.style.display = 'none';
  };

  async function handleNewWorkspaceSubmit(e) {
    e.preventDefault();
    const service = window.FirebaseService;
    const inputName = document.getElementById('inputNewWsName');
    const name = inputName ? inputName.value.trim() : '';
    const visibility = document.querySelector('input[name="newWsVisibility"]:checked')?.value || 'shared';
    const strategyProfileId = document.getElementById('inputNewWsStrategyProfile')?.value || null;

    if (!name) {
      showToast('Vui lòng nhập tên danh mục!', 'warning');
      return;
    }

    try {
      const newWs = await service.createWorkspace(name, visibility, strategyProfileId);
      showToast(`Đã tạo danh mục "${name}" thành công!`, 'success');
      window.closeNewWorkspaceModal();
      await loadWorkspacesDropdown(newWs.id);
    } catch (err) {
      showToast(`Lỗi tạo danh mục: ${err.message}`, 'error');
    }
  }

  // ==================== LƯU DỮ LIỆU MẢNG THEO DANH MỤC (< 1MB GUARD) ====================

  async function handleSaveToFirestore() {
    const service = window.FirebaseService;
    if (!service) return;

    if (service.isGuest()) {
      showToast('Vui lòng đăng nhập để lưu dữ liệu lên server!', 'warning');
      openLoginModal();
      return;
    }

    const select = document.getElementById('selectWorkspace');
    const workspaceId = select ? select.value : null;
    if (!workspaceId) {
      showToast('Vui lòng chọn một Danh mục/Dự án để lưu!', 'warning');
      return;
    }

    const appState = window.state || (typeof state !== 'undefined' ? state : null);
    if (!appState || !appState.extractedByMang) {
      showToast('Chưa có dữ liệu nào được nạp!', 'warning');
      return;
    }

    const availableMangs = Object.keys(appState.extractedByMang).filter(
      k => (appState.extractedByMang[k] || []).length > 0
    );

    if (availableMangs.length === 0) {
      showToast('Chưa có dữ liệu mảng nào! Vui lòng nạp ít nhất một file.', 'warning');
      return;
    }

    // Lọc ra các mảng mà user hiện tại có quyền lưu
    const permittedMangs = availableMangs.filter(k => service.canAccessSector(k, workspaceId));

    if (permittedMangs.length === 0) {
      const curUser = service.getCurrentUser();
      const wsName = select ? select.options[select.selectedIndex]?.textContent : 'Dự án hiện tại';
      const assignedWs = curUser?.assignedWorkspaces || [];
      if (!assignedWs.includes(workspaceId)) {
        showToast(`Tài khoản "${curUser?.username || 'Bạn'}" chưa được phân quyền thao tác trên dự án "${wsName}"!`, 'error');
      } else {
        const assigned = curUser?.assignedSectors || [];
        const assignedNames = assigned.map(k => service.SECTOR_NAMES?.[k] || k).join(', ');
        showToast(`Tài khoản "${curUser?.username || 'Bạn'}" chỉ có quyền quản lý mảng [${assignedNames || 'Không có mảng nào'}] trong dự án "${wsName}". Dữ liệu hiện nạp (${availableMangs.join(', ')}) không có mảng nào bạn được quyền lưu!`, 'error');
      }
      return;
    }

    const btn = document.getElementById('btnSaveToServer');
    const originalText = btn.innerHTML;
    btn.disabled = true;
    btn.innerHTML = `<span class="spinner"></span> Đang lưu...`;

    try {
      let savedCount = 0;
      for (const secKey of permittedMangs) {
        const rows = appState.extractedByMang[secKey];
        // Lưu từng mảng, tự động chặn nếu >= 1MB
        await service.saveSectorDataToFirestore(workspaceId, secKey, rows, null);
        savedCount++;
      }

      showToast(`Đã lưu thành công ${savedCount} mảng (${permittedMangs.join(', ')}) vào danh mục!`, 'success');
      appState.exportBlob = null; // Xóa cache để lượt xuất sau tạo mới theo dữ liệu vừa lưu
      await loadWorkspacesDropdown(workspaceId);
    } catch (err) {
      console.error('Lỗi khi lưu Firestore:', err);
      alert(`⚠️ LỖI LƯU DỮ LIỆU:\n\n${err.message}`);
    } finally {
      btn.disabled = false;
      btn.innerHTML = originalText;
    }
  }

  window.saveCurrentWorkspaceDataToFirebase = handleSaveToFirestore;

  // ==================== MODAL ĐĂNG NHẬP & ĐỔI MẬT KHẨU ====================

  function openLoginModal() {
    const modal = document.getElementById('modalLogin');
    const errBox = document.getElementById('loginErrorMessage');
    if (errBox) errBox.style.display = 'none';
    const inputU = document.getElementById('loginUsername');
    const inputP = document.getElementById('loginPassword');
    if (inputU) inputU.value = '';
    if (inputP) inputP.value = '';
    if (modal) modal.style.display = 'flex';
    if (inputU) inputU.focus();
  }
  window.openLoginModal = openLoginModal;

  window.closeLoginModal = function () {
    const modal = document.getElementById('modalLogin');
    if (modal) modal.style.display = 'none';
  };

  async function handleLoginSubmit(e) {
    if (e && typeof e.preventDefault === 'function') e.preventDefault();
    const service = window.FirebaseService;
    const inputU = document.getElementById('loginUsername');
    const inputP = document.getElementById('loginPassword');
    const btnSubmit = document.getElementById('btnLoginSubmit');
    const errBox = document.getElementById('loginErrorMessage');

    if (btnSubmit && btnSubmit.disabled) return;

    const username = inputU ? inputU.value.trim() : '';
    const password = inputP ? inputP.value : '';

    if (!username || !password) {
      if (errBox) {
        errBox.textContent = 'Vui lòng nhập Tên đăng nhập và Mật khẩu!';
        errBox.style.display = 'block';
      }
      return;
    }

    btnSubmit.disabled = true;
    btnSubmit.innerHTML = `<span class="spinner"></span> Đang xác thực...`;
    if (errBox) errBox.style.display = 'none';

    try {
      const res = await service.loginUser(username, password);
      showToast(res.isNewAdmin ? 'Khởi tạo tài khoản Quản trị viên (Admin) thành công!' : `Xin chào, ${res.profile.displayName || res.profile.username}!`, 'success');
      window.closeLoginModal();
      if (inputP) inputP.value = '';
    } catch (err) {
      if (errBox) {
        errBox.textContent = err.message || 'Đăng nhập thất bại!';
        errBox.style.display = 'block';
      }
    } finally {
      btnSubmit.disabled = false;
      btnSubmit.innerHTML = `<span>Đăng nhập</span>`;
    }
  }
  window.handleLoginSubmit = handleLoginSubmit;

  async function handleLogout() {
    if (!confirm('Bạn có chắc chắn muốn đăng xuất?')) return;
    const service = window.FirebaseService;
    try {
      await service.logoutUser();
      showToast('Đã đăng xuất tài khoản.', 'info');
    } catch (err) {
      showToast(`Lỗi: ${err.message}`, 'error');
    }
  }

  function openChangePasswordModal() {
    const modal = document.getElementById('modalChangePassword');
    const errBox = document.getElementById('changePassErrorMessage');
    if (errBox) errBox.style.display = 'none';
    const inputP = document.getElementById('newPasswordInput');
    const inputCP = document.getElementById('confirmNewPasswordInput');
    if (inputP) inputP.value = '';
    if (inputCP) inputCP.value = '';
    if (modal) modal.style.display = 'flex';
  }

  window.closeChangePasswordModal = function () {
    const modal = document.getElementById('modalChangePassword');
    if (modal) modal.style.display = 'none';
  };

  async function handleChangePasswordSubmit(e) {
    e.preventDefault();
    const service = window.FirebaseService;
    const inputP = document.getElementById('newPasswordInput');
    const inputCP = document.getElementById('confirmNewPasswordInput');
    const errBox = document.getElementById('changePassErrorMessage');
    const btnSubmit = document.getElementById('btnChangePassSubmit');

    const pass = inputP ? inputP.value : '';
    const confirmPass = inputCP ? inputCP.value : '';

    if (!pass || pass.length < 6) {
      if (errBox) {
        errBox.textContent = 'Mật khẩu mới phải có tối thiểu 6 ký tự!';
        errBox.style.display = 'block';
      }
      return;
    }

    if (pass !== confirmPass) {
      if (errBox) {
        errBox.textContent = 'Mật khẩu xác nhận không trùng khớp!';
        errBox.style.display = 'block';
      }
      return;
    }

    btnSubmit.disabled = true;
    btnSubmit.innerHTML = `<span class="spinner"></span> Đang đổi...`;

    try {
      await service.changeUserPassword(pass);
      showToast('Đổi mật khẩu thành công!', 'success');
      window.closeChangePasswordModal();
    } catch (err) {
      if (errBox) {
        errBox.textContent = err.message || 'Lỗi khi đổi mật khẩu.';
        errBox.style.display = 'block';
      }
    } finally {
      btnSubmit.disabled = false;
      btnSubmit.innerHTML = `<span>Cập nhật mật khẩu</span>`;
    }
  }

  // ==================== ADMIN PANEL (QUẢN LÝ TÀI KHOẢN) ====================

  async function populateAdminWorkspaceCheckboxes() {
    const service = window.FirebaseService;
    if (!service || !service.isAdmin()) return;
    const container = document.getElementById('adminNewUserWorkspacesContainer');
    if (!container) return;
    try {
      const allWs = await service.adminGetAllWorkspaces();
      const sharedWs = allWs.filter(w => {
        const vis = service.getWorkspaceVisibility ? service.getWorkspaceVisibility(w) : (w.visibility || 'shared');
        return vis === 'shared' || vis === 'public';
      });
      if (!sharedWs.length) {
        container.innerHTML = `<span style="font-size:0.8rem; color:#94a3b8; font-style:italic;">Chưa có dự án chung nào trên hệ thống.</span>`;
        return;
      }
      container.innerHTML = sharedWs.map(ws => `
        <label class="chip-sector">
          <input type="checkbox" class="admin-ws-chk" value="${escapeHtml(ws.id)}">
          <span class="chip-check-icon">✓</span>
          <span>${escapeHtml(ws.name)}</span>
          <span style="font-size:0.7rem; color:#64748b; margin-left:2px;">(${escapeHtml(ws.createdBy || 'Hệ thống')})</span>
        </label>
      `).join('');
    } catch (err) {
      console.error('Lỗi tải danh sách dự án cho form tạo user:', err);
    }
  }

  async function populateEditUserWorkspaceCheckboxes(selectedWsIds = []) {
    const service = window.FirebaseService;
    const container = document.getElementById('editUserWorkspacesContainer');
    if (!container || !service) return;
    try {
      const allWs = await service.adminGetAllWorkspaces();
      const sharedWs = allWs.filter(w => {
        const vis = service.getWorkspaceVisibility ? service.getWorkspaceVisibility(w) : (w.visibility || 'shared');
        return vis === 'shared' || vis === 'public';
      });
      if (!sharedWs.length) {
        container.innerHTML = `<span style="font-size:0.8rem; color:#94a3b8; font-style:italic;">Chưa có dự án chung nào trên hệ thống.</span>`;
        return;
      }
      container.innerHTML = sharedWs.map(ws => {
        const isChecked = selectedWsIds.includes(ws.id);
        return `
          <label class="chip-sector ${isChecked ? 'active' : ''}">
            <input type="checkbox" class="edit-ws-chk" value="${escapeHtml(ws.id)}" ${isChecked ? 'checked' : ''} onchange="this.closest('.chip-sector').classList.toggle('active', this.checked)">
            <span class="chip-check-icon">✓</span>
            <span>${escapeHtml(ws.name)}</span>
            <span style="font-size:0.7rem; color:#64748b; margin-left:2px;">(${escapeHtml(ws.createdBy || 'Hệ thống')})</span>
          </label>
        `;
      }).join('');
    } catch (err) {
      console.error('Lỗi tải danh sách dự án cho modal edit user:', err);
    }
  }

  async function renderAdminUsersTable() {
    const service = window.FirebaseService;
    if (!service || !service.isAdmin()) return;

    const tbody = document.getElementById('adminUsersTableBody');
    if (!tbody) return;

    // Luôn làm mới danh sách checkbox dự án ở form tạo tài khoản
    populateAdminWorkspaceCheckboxes();

    try {
      const [users, allWorkspaces] = await Promise.all([
        service.adminGetUsersList(),
        service.adminGetAllWorkspaces()
      ]);

      if (users.length === 0) {
        tbody.innerHTML = `<tr><td colspan="7" style="text-align:center; padding: 2rem; color: #64748b;">Chưa có tài khoản nào khác.</td></tr>`;
        return;
      }

      tbody.innerHTML = users.map((u, i) => {
        const isAdm = u.role === 'admin';
        const roleBadge = isAdm
          ? `<span class="badge" style="background:#e11d48; color:#fff; font-weight:700;">Admin</span>`
          : `<span class="badge" style="background:#0284c7; color:#fff; font-weight:700;">User</span>`;

        const sectorBadges = isAdm
          ? `<span style="font-weight:700; color:#059669; font-size:0.8rem;">Toàn bộ 7 mảng</span>`
          : (u.assignedSectors || []).map(s => `<span class="badge-mang badge-mang-${s}" style="font-size:0.68rem; padding: 2px 7px;">${s}</span>`).join(' ') || '<span style="color:#94a3b8; font-style:italic;">Chưa gán</span>';

        const wsBadges = isAdm
          ? `<span style="font-weight:700; color:#059669; font-size:0.8rem;">Toàn bộ dự án</span>`
          : (u.assignedWorkspaces && u.assignedWorkspaces.length > 0)
            ? u.assignedWorkspaces.map(wId => {
                const wsObj = allWorkspaces.find(w => w.id === wId);
                const name = wsObj ? wsObj.name : wId;
                return `<span class="badge" style="background:#ede9fe; color:#5b21b6; font-size:0.7rem; padding: 2px 8px; border: 1px solid #ddd6fe; border-radius: 9999px;">${escapeHtml(name)}</span>`;
              }).join(' ')
            : '<span style="color:#94a3b8; font-style:italic;">Chưa gán</span>';

        const uSecsJson = JSON.stringify(u.assignedSectors || []).replace(/"/g, '&quot;');
        const uWsJson = JSON.stringify(u.assignedWorkspaces || []).replace(/"/g, '&quot;');

        return `
          <tr>
            <td style="text-align:center; font-weight:700; color:#64748b;">${i + 1}</td>
            <td>
              <div style="font-weight:700; color:#0f172a;">${escapeHtml(u.username)}</div>
            </td>
            <td style="font-weight:600; color:#334155;">${escapeHtml(u.displayName || u.username)}</td>
            <td style="text-align:center;">${roleBadge}</td>
            <td><div style="display:flex; gap:4px; flex-wrap:wrap;">${sectorBadges}</div></td>
            <td><div style="display:flex; gap:4px; flex-wrap:wrap;">${wsBadges}</div></td>
            <td style="text-align:center;">
              ${!isAdm ? `
                <button type="button" class="btn btn-outline btn-xs" onclick="window.openEditUserSectorsModal('${u.uid}', '${escapeHtml(u.username)}', ${uSecsJson}, ${uWsJson})" title="Sửa phân quyền">
                  <i data-lucide="edit-3" class="w-3.5 h-3.5"></i> Sửa quyền
                </button>
                <button type="button" class="btn btn-outline btn-xs" style="color:#e11d48; border-color:#fecdd3;" onclick="window.adminDeleteUserHandler('${u.uid}', '${escapeHtml(u.username)}')" title="Xóa tài khoản">
                  <i data-lucide="trash-2" class="w-3.5 h-3.5"></i>
                </button>
              ` : '<span style="font-size:0.75rem; color:#94a3b8; font-weight:600;">Hệ thống</span>'}
            </td>
          </tr>
        `;
      }).join('');

      if (window.lucide) window.lucide.createIcons();
    } catch (err) {
      console.error('Lỗi nạp bảng users admin:', err);
    }
  }

  async function handleAdminCreateUserSubmit(e) {
    e.preventDefault();
    const service = window.FirebaseService;
    if (!service || !service.isAdmin()) return;

    const inputU = document.getElementById('adminNewUsername');
    const inputP = document.getElementById('adminNewPassword');
    const inputN = document.getElementById('adminNewDisplayName');
    const btnSubmit = document.getElementById('btnAdminCreateUserSubmit');

    const username = inputU ? inputU.value.trim() : '';
    const password = inputP ? inputP.value : '';
    const displayName = inputN ? inputN.value.trim() : '';

    const checkedSectors = [];
    document.querySelectorAll('.admin-sector-chk:checked').forEach(chk => {
      checkedSectors.push(chk.value);
    });

    const checkedWorkspaces = [];
    document.querySelectorAll('.admin-ws-chk:checked').forEach(chk => {
      checkedWorkspaces.push(chk.value);
    });

    if (!username || !password) {
      showToast('Vui lòng nhập Tên đăng nhập và Mật khẩu!', 'warning');
      return;
    }

    btnSubmit.disabled = true;
    btnSubmit.innerHTML = `<span class="spinner"></span> Đang tạo...`;

    try {
      await service.adminCreateUser(username, password, displayName, checkedSectors, checkedWorkspaces);
      showToast(`Đã tạo tài khoản "${username}" thành công!`, 'success');
      if (inputU) inputU.value = '';
      if (inputP) inputP.value = '';
      if (inputN) inputN.value = '';
      document.querySelectorAll('.admin-sector-chk, .admin-ws-chk').forEach(c => {
        c.checked = false;
        const parent = c.closest('.chip-sector, .chip');
        if (parent) parent.classList.remove('active');
      });
      renderAdminUsersTable();
    } catch (err) {
      showToast(`Lỗi tạo tài khoản: ${err.message}`, 'error');
    } finally {
      btnSubmit.disabled = false;
      btnSubmit.innerHTML = `<i data-lucide="user-plus" class="w-4 h-4"></i><span>Tạo tài khoản</span>`;
      if (window.lucide) window.lucide.createIcons();
    }
  }

  window.adminDeleteUserHandler = async function (uid, username) {
    if (!confirm(`Bạn có chắc chắn muốn xóa tài khoản "${username}"?`)) return;
    const service = window.FirebaseService;
    try {
      await service.adminDeleteUser(uid);
      showToast(`Đã xóa tài khoản "${username}".`, 'success');
      renderAdminUsersTable();
    } catch (err) {
      showToast(`Lỗi: ${err.message}`, 'error');
    }
  };

  // Modal sửa quyền mảng và dự án của user
  window.openEditUserSectorsModal = function (uid, username, currentSectors = [], currentWorkspaces = []) {
    const modal = document.getElementById('modalEditUserSectors');
    if (!modal) return;
    document.getElementById('editUserSectorsUid').value = uid;
    document.getElementById('editUserSectorsUsername').textContent = username;

    document.querySelectorAll('.edit-sector-chk').forEach(chk => {
      const isChecked = currentSectors.includes(chk.value);
      chk.checked = isChecked;
      const parent = chk.closest('.chip-sector, .chip');
      if (parent) {
        if (isChecked) parent.classList.add('active');
        else parent.classList.remove('active');
      }
    });

    populateEditUserWorkspaceCheckboxes(currentWorkspaces);

    modal.style.display = 'flex';
  };

  window.closeEditUserSectorsModal = function () {
    const modal = document.getElementById('modalEditUserSectors');
    if (modal) modal.style.display = 'none';
  };

  window.saveEditUserSectors = async function () {
    const service = window.FirebaseService;
    const uid = document.getElementById('editUserSectorsUid').value;
    const checkedSectors = [];
    document.querySelectorAll('.edit-sector-chk:checked').forEach(c => checkedSectors.push(c.value));

    const checkedWorkspaces = [];
    document.querySelectorAll('.edit-ws-chk:checked').forEach(c => checkedWorkspaces.push(c.value));

    try {
      await service.adminUpdateUserPermissions(uid, checkedSectors, checkedWorkspaces);
      showToast('Đã cập nhật phân quyền mảng & dự án thành công!', 'success');
      window.closeEditUserSectorsModal();
      renderAdminUsersTable();
      applySectorPermissions();
    } catch (err) {
      showToast(`Lỗi: ${err.message}`, 'error');
    }
  };

  // ==================== QUẢN LÝ DỰ ÁN / DANH MỤC (ADMIN PANEL) ====================

  async function renderAdminWorkspacesTable() {
    const service = window.FirebaseService;
    if (!service || service.isGuest()) return;

    const tbody = document.getElementById('adminWorkspacesTableBody');
    if (!tbody) return;

    // Cập nhật dropdown strategy profile ở form tạo mới
    const adminNewWsSel = document.getElementById('adminNewWsStrategyProfile');
    if (adminNewWsSel) {
      const profiles = (window.state && window.state.strategyProfiles) ? window.state.strategyProfiles : [];
      adminNewWsSel.innerHTML = profiles.map(p => `
        <option value="${p.id}" ${p.id === 'profile_default' ? 'selected' : ''}>
          ${escapeHtml(p.name)}
        </option>
      `).join('');
    }

    try {
      const workspaces = await service.adminGetAllWorkspaces();
      if (!workspaces.length) {
        tbody.innerHTML = `<tr><td colspan="7" style="text-align:center; padding: 2rem; color: #64748b;">Chưa có danh mục nào.</td></tr>`;
        return;
      }

      const activeWsId = document.getElementById('selectWorkspace')?.value;
      const profilesList = (window.state && window.state.strategyProfiles) ? window.state.strategyProfiles : [];

      tbody.innerHTML = workspaces.map((ws, idx) => {
        const isDefault = ws.id === 'ws_toan_quoc_2027_2028';
        const isActive = ws.id === activeWsId;
        const canManage = service.canManageWorkspace ? service.canManageWorkspace(ws) : service.isAdmin();
        const visibility = service.getWorkspaceVisibility ? service.getWorkspaceVisibility(ws) : (ws.isPublic ? 'public' : (ws.isShared ? 'shared' : 'private'));
        const visibilityMeta = {
          public: {
            label: 'Public',
            desc: 'Guest xem; user xem/sửa',
            badgeClass: 'badge-visibility-public'
          },
          shared: {
            label: 'Dự án chung',
            desc: 'User xem/sửa; guest không xem',
            badgeClass: 'badge-visibility-shared'
          },
          private: {
            label: 'Mình tôi',
            desc: 'Chỉ chủ sở hữu và admin',
            badgeClass: 'badge-visibility-private'
          }
        }[visibility] || {
          label: visibility || 'Chung',
          desc: '',
          badgeClass: 'badge-visibility-shared'
        };
        const visibilityBadge = `
          <div style="display:inline-flex; flex-direction:column; align-items:center; gap:2px;">
            <span class="badge-visibility ${visibilityMeta.badgeClass}">${visibilityMeta.label}</span>
            <span style="font-size:0.68rem; color:#475569; font-weight:500;">${visibilityMeta.desc}</span>
          </div>
        `;

        // Profile Chiến lược 5 năm gắn với dự án này
        const stratProfileId = ws.strategyProfileId;
        let matchedProfile = profilesList.find(p => p.id === stratProfileId);
        if (!matchedProfile) {
          const n = (ws.name || '').toLowerCase();
          if (n.includes('lào') || n.includes('lao')) {
            matchedProfile = profilesList.find(p => p.id === 'profile_1789883317033' || (p.name || '').toLowerCase().includes('lào'));
          } else if (n.includes('mozambique') || n.includes('movitel')) {
            matchedProfile = profilesList.find(p => p.id === 'profile_1789854757103' || (p.name || '').toLowerCase().includes('mozambique'));
          }
        }
        const profileName = matchedProfile ? matchedProfile.name : (stratProfileId ? stratProfileId : 'Chuẩn (Mặc định)');
        const profileBadge = `
          <div style="display:inline-flex; align-items:center; gap:4px;">
            <span class="badge" style="background:#4338ca; color:#ffffff !important; font-size:0.75rem; font-weight:700; padding:2px 8px; border-radius:6px;" title="${escapeHtml(profileName)}">
              ${escapeHtml(profileName)}
            </span>
          </div>
        `;

        // Dữ liệu các mảng đã có
        const loadedKeys = Object.keys(ws.sectorsLoaded || {});
        let sectorsContent = '<span style="color:#94a3b8; font-style:italic;">Chưa có dữ liệu</span>';
        if (loadedKeys.length > 0) {
          sectorsContent = loadedKeys.map(k => `
            <span class="badge-mang badge-mang-${k}" style="font-size:0.68rem; padding: 2px 6px;" title="${k}: ${ws.sectorsLoaded[k]} dòng">
              ${k}: <strong>${ws.sectorsLoaded[k]}</strong>
            </span>
          `).join(' ');
        }

        const dateStr = ws.createdAt ? new Date(ws.createdAt).toLocaleDateString('vi-VN') : '';

        return `
          <tr style="${isActive ? 'background: #eff6ff; border-left: 4px solid #2563eb;' : ''}">
            <td style="text-align:center; font-weight:700; color:#64748b;">${idx + 1}</td>
            <td>
              <div style="display:flex; align-items:center; gap:8px;">
                <span style="font-weight:700; font-size:0.92rem; color:#0f172a;">${escapeHtml(ws.name)}</span>
                ${isActive ? '<span class="badge" style="background:#2563eb; color:#fff !important; font-size:0.68rem; padding:2px 7px; border-radius:9999px;">● Đang làm việc (Sidebar)</span>' : ''}
              </div>
              <div style="font-size:0.75rem; color:#64748b; margin-top:2px;">ID: ${escapeHtml(ws.id)}</div>
            </td>
            <td style="text-align:center;">${visibilityBadge}</td>
            <td style="text-align:center;">${profileBadge}</td>
            <td>
              <div style="display:flex; gap:4px; flex-wrap:wrap; align-items:center;">
                ${sectorsContent}
              </div>
            </td>
            <td style="text-align:center;">
              <div style="font-weight:600; color:#334155; font-size:0.82rem;">${escapeHtml(ws.createdBy || 'Hệ thống')}</div>
              <div style="font-size:0.72rem; color:#94a3b8;">${dateStr}</div>
            </td>
            <td style="text-align:center;">
              <div style="display:inline-flex; gap:6px;">
                ${canManage ? `
                <button type="button" class="btn btn-outline btn-xs" onclick="window.openEditWorkspaceModal('${ws.id}', '${escapeHtml(ws.name).replace(/'/g, "\\'")}', '${visibility}', '${escapeHtml(ws.strategyProfileId || '')}')" title="Sửa thông tin danh mục">
                  <i data-lucide="edit-3" class="w-3.5 h-3.5"></i> Sửa
                </button>
                ` : ''}
                ${(!isDefault && canManage) ? `
                  <button type="button" class="btn btn-outline btn-xs" style="color:#e11d48; border-color:#fecdd3;" onclick="window.adminDeleteWorkspaceHandler('${ws.id}', '${escapeHtml(ws.name).replace(/'/g, "\\'")}')" title="Xóa danh mục">
                    <i data-lucide="trash-2" class="w-3.5 h-3.5"></i> Xóa
                  </button>
                ` : ''}
                ${!canManage ? '<span style="font-size:0.75rem; color:#94a3b8; font-weight:600;">Chỉ xem</span>' : ''}
              </div>
            </td>
          </tr>
        `;
      }).join('');

      if (window.lucide) window.lucide.createIcons();
    } catch (err) {
      console.error('Lỗi renderAdminWorkspacesTable:', err);
    }
  }

  window.renderAdminWorkspacesTable = renderAdminWorkspacesTable;

  window.adminSelectWorkspaceHandler = async function (wsId) {
    const select = document.getElementById('selectWorkspace');
    if (select) {
      select.value = wsId;
      select.dispatchEvent(new Event('change'));
      showToast('Đã chuyển sang danh mục được chọn!', 'info');
      renderAdminWorkspacesTable();
    }
  };

  window.adminDeleteWorkspaceHandler = async function (wsId, name) {
    if (!confirm(`Bạn có chắc chắn muốn xóa danh mục "${name}"?\n(Dữ liệu các mảng trong danh mục này sẽ bị xóa khỏi hệ thống)`)) return;
    const service = window.FirebaseService;
    try {
      await service.deleteWorkspace(wsId);
      showToast(`Đã xóa danh mục "${name}".`, 'success');
      await loadWorkspacesDropdown();
      renderAdminWorkspacesTable();
    } catch (err) {
      showToast(`Lỗi: ${err.message}`, 'error');
    }
  };

  async function handleAdminCreateWorkspaceSubmit(e) {
    e.preventDefault();
    const service = window.FirebaseService;
    if (!service || service.isGuest()) return;

    const inputName = document.getElementById('adminNewWsNameInput');
    const btnSubmit = document.getElementById('btnAdminCreateWsSubmit');

    const name = inputName ? inputName.value.trim() : '';
    const visibility = document.querySelector('input[name="adminNewWsVisibility"]:checked')?.value || 'shared';
    const strategyProfileId = document.getElementById('adminNewWsStrategyProfile')?.value || null;

    if (!name) {
      showToast('Vui lòng nhập tên danh mục!', 'warning');
      return;
    }

    btnSubmit.disabled = true;
    btnSubmit.innerHTML = `<span class="spinner"></span> Đang tạo...`;

    try {
      const newWs = await service.createWorkspace(name, visibility, strategyProfileId);
      showToast(`Đã tạo danh mục "${name}" thành công!`, 'success');
      if (inputName) inputName.value = '';
      const sharedRadio = document.getElementById('adminNewWsVisShared');
      if (sharedRadio) sharedRadio.checked = true;
      await loadWorkspacesDropdown(newWs.id);
      renderAdminWorkspacesTable();
    } catch (err) {
      showToast(`Lỗi tạo danh mục: ${err.message}`, 'error');
    } finally {
      btnSubmit.disabled = false;
      btnSubmit.innerHTML = `<i data-lucide="plus-circle" class="w-4 h-4"></i><span>Tạo Danh Mục</span>`;
      if (window.lucide) window.lucide.createIcons();
    }
  }

  // Modal Sửa Danh mục
  window.openEditWorkspaceModal = function (wsId, name, visibility, currentStratProfileId) {
    const modal = document.getElementById('modalEditWorkspace');
    if (!modal) return;
    document.getElementById('editWsId').value = wsId;
    document.getElementById('editWsName').value = name;
    const radio = document.querySelector(`input[name="editWsVisibility"][value="${visibility || 'shared'}"]`);
    if (radio) radio.checked = true;

    // Nạp dropdown Strategy Profile
    const selProfile = document.getElementById('editWsStrategyProfile');
    if (selProfile) {
      const profiles = (window.state && window.state.strategyProfiles) ? window.state.strategyProfiles : [];
      let activeProfileId = currentStratProfileId;
      if (!activeProfileId && window.FirebaseService && window.FirebaseService.getWorkspaceStrategyProfile) {
        activeProfileId = window.FirebaseService.getWorkspaceStrategyProfile(wsId);
      }
      selProfile.innerHTML = profiles.map(p => `
        <option value="${p.id}" ${p.id === activeProfileId ? 'selected' : ''}>
          ${escapeHtml(p.name)}
        </option>
      `).join('');
    }

    modal.style.display = 'flex';
  };

  window.closeEditWorkspaceModal = function () {
    const modal = document.getElementById('modalEditWorkspace');
    if (modal) modal.style.display = 'none';
  };

  window.saveEditWorkspaceHandler = async function () {
    const service = window.FirebaseService;
    const wsId = document.getElementById('editWsId').value;
    const name = document.getElementById('editWsName').value.trim();
    const visibility = document.querySelector('input[name="editWsVisibility"]:checked')?.value || 'shared';
    const strategyProfileId = document.getElementById('editWsStrategyProfile')?.value || null;

    if (!name) {
      showToast('Tên danh mục không được để trống!', 'warning');
      return;
    }

    try {
      await service.updateWorkspace(wsId, { name, visibility, strategyProfileId });
      showToast('Đã cập nhật danh mục thành công!', 'success');
      window.closeEditWorkspaceModal();

      const currentActiveId = document.getElementById('selectWorkspace')?.value;
      if (wsId === currentActiveId && window.syncStrategyProfileWithActiveWorkspace) {
        window.syncStrategyProfileWithActiveWorkspace(wsId);
      }

      await loadWorkspacesDropdown(currentActiveId);
      renderAdminWorkspacesTable();
    } catch (err) {
      showToast(`Lỗi: ${err.message}`, 'error');
    }
  };

  // ==================== TIỆN ÍCH HELPER ====================

  function showToast(msg, type = 'info') {
    const bgColors = {
      success: '#059669',
      error: '#dc2626',
      warning: '#d97706',
      info: '#2563eb'
    };
    const toast = document.createElement('div');
    toast.className = 'app-toast';
    toast.style.cssText = `
      position: fixed;
      bottom: 24px;
      right: 24px;
      background: ${bgColors[type] || '#1e293b'};
      color: #ffffff;
      padding: 12px 20px;
      border-radius: 10px;
      font-size: 0.88rem;
      font-weight: 600;
      box-shadow: 0 10px 25px rgba(0,0,0,0.2);
      z-index: 99999;
      display: flex;
      align-items: center;
      gap: 10px;
      animation: modalScaleUp 0.2s ease;
    `;
    toast.textContent = msg;
    document.body.appendChild(toast);
    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transition = 'opacity 0.3s ease';
      setTimeout(() => toast.remove(), 300);
    }, 3500);
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

})();
