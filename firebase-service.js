// firebase-service.js - Quản lý Tài khoản, Phân quyền & Lưu trữ Dữ liệu QHDC3
// Tương thích 100% cả môi trường Trực tuyến (Firebase) và Cục bộ / Offline / file:// protocol

(function () {
  'use strict';

  // 8 mảng chuẩn của hệ thống
  const ALL_SECTORS = ['VT', 'ML', 'CDBR', 'CNTT', 'TD', 'IP', 'CD', 'HT'];
  const SECTOR_NAMES = {
    VT: 'Vô tuyến',
    ML: 'Mạng lõi',
    CDBR: 'CĐBR & Truyền hình',
    CNTT: 'Công nghệ thông tin',
    TD: 'Truyền dẫn quang',
    IP: 'Truyền dẫn IP',
    CD: 'Cơ điện',
    HT: 'Triển khai hạ tầng'
  };

  const STORAGE_KEYS = {
    CURRENT_USER: 'qhdc_current_user',
    USERS: 'qhdc_users_db',
    WORKSPACES: 'qhdc_workspaces_db',
    SECTOR_DATA_PREFIX: 'qhdc_sector_data_'
  };

  let currentUserProfile = null;
  const authListeners = [];
  const workspaceListeners = [];

  // ==================== KHỞI TẠO DỮ LIỆU MẶC ĐỊNH ====================

  function getLocalData(key, defaultVal) {
    try {
      const raw = localStorage.getItem(key);
      return raw ? JSON.parse(raw) : defaultVal;
    } catch (e) {
      console.warn('Lỗi đọc localStorage:', e);
      return defaultVal;
    }
  }

  function setLocalData(key, val) {
    try {
      localStorage.setItem(key, JSON.stringify(val));
    } catch (e) {
      console.warn('Lỗi ghi localStorage:', e);
      if (e.name === 'QuotaExceededError' || e.code === 22 || e.code === 1014) {
        throw new Error('Dung lượng lưu trữ của trình duyệt (localStorage) đã đầy. Vui lòng xóa bớt dữ liệu các dự án cũ!');
      }
      throw e;
    }
  }

  // Khởi tạo workspaces mặc định nếu chưa có
  function initDefaultWorkspaces() {
    let wsList = getLocalData(STORAGE_KEYS.WORKSPACES, null);
    if (!wsList || !wsList.length) {
      wsList = [
        {
          id: 'ws_toan_quoc_2027_2028',
          name: 'QHĐC Toàn quốc 2027–2028',
          visibility: 'public',
          isPublic: true,
          isShared: true,
          strategyProfileId: 'profile_default',
          createdBy: 'admin',
          createdAt: new Date().toISOString()
        }
      ];
      setLocalData(STORAGE_KEYS.WORKSPACES, wsList);
    } else {
      // Auto-link profile cho các workspace cũ nếu chưa gán
      let changed = false;
      wsList.forEach(w => {
        if (!w.strategyProfileId) {
          const n = (w.name || '').toLowerCase();
          if (n.includes('lào') || n.includes('lao')) {
            w.strategyProfileId = 'profile_1789883317033';
            changed = true;
          } else if (n.includes('mozambique') || n.includes('movitel')) {
            w.strategyProfileId = 'profile_1789854757103';
            changed = true;
          } else if (w.id === 'ws_toan_quoc_2027_2028') {
            w.strategyProfileId = 'profile_default';
            changed = true;
          }
        }
      });
      if (changed) {
        setLocalData(STORAGE_KEYS.WORKSPACES, wsList);
      }
    }
    return wsList;
  }

  // Khôi phục phiên đăng nhập trước đó nếu có
  function initSession() {
    initDefaultWorkspaces();
    const savedUser = getLocalData(STORAGE_KEYS.CURRENT_USER, null);
    if (savedUser) {
      currentUserProfile = savedUser;
    }
  }

  initSession();

  // ==================== AUTH & PHÂN QUYỀN ====================

  function onAuthStateChange(callback) {
    authListeners.push(callback);
    // Kích hoạt ngay callback với trạng thái hiện tại
    try {
      callback(currentUserProfile);
    } catch (e) {
      console.error(e);
    }
  }

  function notifyAuthListeners() {
    authListeners.forEach(cb => {
      try {
        cb(currentUserProfile);
      } catch (e) {
        console.error(e);
      }
    });
  }

  /**
   * Đăng nhập tài khoản bằng Username & Password
   */
  async function loginUser(username, password) {
    const cleanUser = (username || '').trim().toLowerCase();
    const cleanPass = (password || '').trim();

    if (!cleanUser || !cleanPass) {
      throw new Error('Vui lòng nhập tên đăng nhập và mật khẩu.');
    }

    let users = getLocalData(STORAGE_KEYS.USERS, []);

    // Nếu chưa có tài khoản nào và người dùng đăng nhập bằng "admin"
    if (cleanUser === 'admin') {
      let adminAccount = users.find(u => u.username.toLowerCase() === 'admin');
      if (!adminAccount) {
        // Tự động khởi tạo tài khoản admin đầu tiên
        adminAccount = {
          uid: 'uid_admin',
          username: 'admin',
          password: cleanPass,
          displayName: 'Quản trị viên Hệ thống',
          dept: 'TT, KTTC',
          role: 'admin',
          assignedSectors: [...ALL_SECTORS],
          createdAt: new Date().toISOString()
        };
        users.push(adminAccount);
        setLocalData(STORAGE_KEYS.USERS, users);

        currentUserProfile = { ...adminAccount };
        delete currentUserProfile.password;
        setLocalData(STORAGE_KEYS.CURRENT_USER, currentUserProfile);
        notifyAuthListeners();
        return { success: true, profile: currentUserProfile, isNewAdmin: true };
      }
    }

    const found = users.find(u => u.username.toLowerCase() === cleanUser);
    if (!found) {
      throw new Error(`Tài khoản "${username}" không tồn tại. Vui lòng liên hệ Admin tạo tài khoản.`);
    }

    if (found.password !== cleanPass) {
      throw new Error('Mật khẩu không chính xác.');
    }

    currentUserProfile = { ...found };
    delete currentUserProfile.password;
    setLocalData(STORAGE_KEYS.CURRENT_USER, currentUserProfile);
    notifyAuthListeners();

    return { success: true, profile: currentUserProfile };
  }

  /**
   * Đăng xuất tài khoản
   */
  async function logoutUser() {
    currentUserProfile = null;
    try {
      localStorage.removeItem(STORAGE_KEYS.CURRENT_USER);
    } catch (e) {}
    notifyAuthListeners();
    return true;
  }

  /**
   * Đổi mật khẩu tài khoản hiện tại
   */
  async function changeUserPassword(newPassword) {
    if (!currentUserProfile) throw new Error('Chưa đăng nhập.');
    if (!newPassword || newPassword.length < 6) {
      throw new Error('Mật khẩu mới phải có tối thiểu 6 ký tự.');
    }

    const users = getLocalData(STORAGE_KEYS.USERS, []);
    const idx = users.findIndex(u => u.uid === currentUserProfile.uid);
    if (idx !== -1) {
      users[idx].password = newPassword;
      setLocalData(STORAGE_KEYS.USERS, users);
    }
    return true;
  }

  function getCurrentUser() {
    return currentUserProfile;
  }

  function isGuest() {
    return !currentUserProfile;
  }

  function isAdmin() {
    return currentUserProfile && currentUserProfile.role === 'admin';
  }

  /**
   * Kiểm tra quyền thao tác trên một mảng trong một dự án (workspace)
   * - Nếu là người tạo/chủ sở hữu dự án đó (hoặc Admin): Toàn quyền mọi mảng trong dự án.
   * - Nếu là dự án khác/dự án chung: Chỉ có quyền trên các mảng được giao (assignedSectors).
   */
  function canAccessSector(sectorKey, workspaceOrId = null) {
    if (!currentUserProfile) return false;
    if (currentUserProfile.role === 'admin') return true;

    // 1. Xác định workspace cần kiểm tra
    let ws = null;
    let wsId = null;
    const list = getLocalData(STORAGE_KEYS.WORKSPACES, []);

    if (workspaceOrId && typeof workspaceOrId === 'object') {
      ws = workspaceOrId;
      wsId = ws.id;
    } else {
      wsId = workspaceOrId || (typeof document !== 'undefined' ? document.getElementById('selectWorkspace')?.value : null);
      if (wsId) {
        ws = list.find(w => w.id === wsId);
      }
    }

    // 2. Nếu là dự án riêng do chính họ tạo (Owner): Toàn quyền với mọi mảng trong dự án này!
    if (ws && canManageWorkspace(ws)) {
      return true;
    }

    // 3. Nếu là dự án ở chế độ "private" và không phải owner -> Không có quyền
    if (ws && getWorkspaceVisibility(ws) === 'private') {
      return false;
    }

    // 4. Nếu là dự án chung (shared / public):
    // Cần thỏa mãn CẢ HAI điều kiện:
    const assignedWs = currentUserProfile.assignedWorkspaces || [];
    const targetWsId = wsId || ws?.id;
    if (!targetWsId || !assignedWs.includes(targetWsId)) {
      return false;
    }

    const assigned = currentUserProfile.assignedSectors || [];
    return assigned.includes(sectorKey);
  }

  // ==================== QUẢN LÝ USER (ADMIN PANEL) ====================

  async function adminGetUsersList() {
    if (!isAdmin()) throw new Error('Chỉ Admin mới có quyền xem danh sách user.');
    const users = getLocalData(STORAGE_KEYS.USERS, []);
    return users.map(u => {
      const copy = { ...u };
      delete copy.password;
      return copy;
    });
  }

  async function adminCreateUser(username, password, displayName, assignedSectors = [], assignedWorkspaces = []) {
    if (!isAdmin()) throw new Error('Chỉ Admin mới có quyền tạo tài khoản.');
    const cleanUser = (username || '').trim().toLowerCase();
    const cleanPass = (password || '').trim();

    if (!cleanUser || !cleanPass) throw new Error('Vui lòng điền đủ Tên đăng nhập và Mật khẩu.');
    if (cleanPass.length < 6) throw new Error('Mật khẩu phải từ 6 ký tự trở lên.');

    const users = getLocalData(STORAGE_KEYS.USERS, []);
    if (users.some(u => u.username.toLowerCase() === cleanUser)) {
      throw new Error(`Tên đăng nhập "${username}" đã tồn tại!`);
    }

    const newUser = {
      uid: 'uid_' + Date.now() + '_' + Math.random().toString(36).substring(2, 7),
      username: cleanUser,
      password: cleanPass,
      displayName: (displayName || '').trim() || cleanUser,
      dept: 'TT, KTTC',
      role: 'user',
      assignedSectors: assignedSectors || [],
      assignedWorkspaces: assignedWorkspaces || [],
      createdAt: new Date().toISOString()
    };

    users.push(newUser);
    setLocalData(STORAGE_KEYS.USERS, users);
    return newUser;
  }

  async function adminUpdateUserPermissions(uid, newSectors, newWorkspaces = null) {
    if (!isAdmin()) throw new Error('Chỉ Admin mới có quyền sửa phân quyền.');
    const users = getLocalData(STORAGE_KEYS.USERS, []);
    const user = users.find(u => u.uid === uid);
    if (!user) throw new Error('Không tìm thấy tài khoản.');

    user.assignedSectors = newSectors;
    if (Array.isArray(newWorkspaces)) {
      user.assignedWorkspaces = newWorkspaces;
    }
    setLocalData(STORAGE_KEYS.USERS, users);

    if (currentUserProfile && currentUserProfile.uid === uid) {
      currentUserProfile.assignedSectors = newSectors;
      if (Array.isArray(newWorkspaces)) {
        currentUserProfile.assignedWorkspaces = newWorkspaces;
      }
      setLocalData(STORAGE_KEYS.CURRENT_USER, currentUserProfile);
      notifyAuthListeners();
    }
    return true;
  }

  async function adminDeleteUser(uid) {
    if (!isAdmin()) throw new Error('Chỉ Admin mới có quyền xóa tài khoản.');
    let users = getLocalData(STORAGE_KEYS.USERS, []);
    const target = users.find(u => u.uid === uid);
    if (target && target.username.toLowerCase() === 'admin') {
      throw new Error('Không thể xóa tài khoản Quản trị viên tối cao.');
    }
    users = users.filter(u => u.uid !== uid);
    setLocalData(STORAGE_KEYS.USERS, users);
    return true;
  }

  // ==================== WORKSPACE (DANH MỤC / DỰ ÁN) ====================

  /**
   * Xác định chế độ phân quyền của danh mục:
   * - 'public': Khách xem được, all user xem và chỉnh sửa được
   * - 'shared': Dự án chung - all user xem và chỉnh sửa được, Khách không xem được
   * - 'private': Mình tôi - chỉ người tạo mới xem và chỉnh sửa được
   */
  function getWorkspaceVisibility(ws) {
    if (!ws) return 'shared';
    if (ws.visibility && ['public', 'shared', 'private'].includes(ws.visibility)) {
      return ws.visibility;
    }
    if (ws.isPublic) return 'public';
    if (ws.isShared) return 'shared';
    return 'private';
  }

  /**
   * Quyền quản trị danh mục (Sửa tên, Đổi chế độ phân quyền, Xóa):
   * User tạo danh mục nào thì được set quyền cho danh mục đó, ngoại trừ Admin được set tất cả.
   */
  function canManageWorkspace(ws) {
    if (!ws || isGuest()) return false;
    if (isAdmin()) return true;
    const curUser = (currentUserProfile?.username || '').trim().toLowerCase();
    const curUid = currentUserProfile?.uid;
    const wsCreator = (ws.createdBy || '').trim().toLowerCase();
    return (wsCreator && wsCreator === curUser) ||
           (ws.creatorUid && ws.creatorUid === curUid);
  }

  /**
   * Quyền nạp/chỉnh sửa mảng trong danh mục:
   * - Public: Tất cả chuyên viên đã đăng nhập (theo phân quyền mảng của họ)
   * - Shared: Tất cả chuyên viên đã đăng nhập (theo phân quyền mảng của họ)
   * - Private: Chỉ người tạo danh mục (hoặc Admin)
   */
  function canEditWorkspaceSectors(ws) {
    if (!ws || isGuest()) return false;
    if (isAdmin()) return true;
    const vis = getWorkspaceVisibility(ws);
    if (vis === 'public' || vis === 'shared') return true;
    const curUser = (currentUserProfile?.username || '').trim().toLowerCase();
    const curUid = currentUserProfile?.uid;
    const wsCreator = (ws.createdBy || '').trim().toLowerCase();
    return (wsCreator && wsCreator === curUser) ||
           (ws.creatorUid && ws.creatorUid === curUid);
  }

  async function listWorkspaces() {
    const list = getLocalData(STORAGE_KEYS.WORKSPACES, []);
    return list.filter(w => {
      const vis = getWorkspaceVisibility(w);
      if (isAdmin()) return true;
      if (isGuest()) {
        return vis === 'public';
      }
      // User thường đã đăng nhập:
      // - Public và Dự án chung: Xem được
      // - Mình tôi: Chỉ người tạo xem được
      if (vis === 'public' || vis === 'shared') return true;
      return (w.createdBy && w.createdBy === currentUserProfile.username) ||
             (w.creatorUid && w.creatorUid === currentUserProfile.uid);
    });
  }

  async function createWorkspace(name, visibility = 'shared', strategyProfileId = null) {
    if (isGuest()) throw new Error('Vui lòng đăng nhập để tạo danh mục.');
    const cleanName = (name || '').trim();
    if (!cleanName) throw new Error('Tên danh mục không được để trống.');

    const list = getLocalData(STORAGE_KEYS.WORKSPACES, []);
    const wsId = 'ws_' + Date.now();
    const vis = ['public', 'shared', 'private'].includes(visibility) ? visibility : 'shared';

    const newWs = {
      id: wsId,
      name: cleanName,
      visibility: vis,
      isPublic: vis === 'public',
      isShared: vis === 'shared' || vis === 'public',
      strategyProfileId: strategyProfileId || null,
      createdBy: currentUserProfile.username,
      creatorUid: currentUserProfile.uid,
      createdAt: new Date().toISOString()
    };
    list.push(newWs);
    setLocalData(STORAGE_KEYS.WORKSPACES, list);
    return newWs;
  }

  async function updateWorkspace(workspaceId, { name, visibility, strategyProfileId }) {
    if (isGuest()) throw new Error('Vui lòng đăng nhập.');
    const list = getLocalData(STORAGE_KEYS.WORKSPACES, []);
    const ws = list.find(w => w.id === workspaceId);
    if (!ws) throw new Error('Không tìm thấy danh mục.');

    // Kiểm tra quyền: Người tạo hoặc Admin
    if (!canManageWorkspace(ws)) {
      throw new Error('Bạn không có quyền chỉnh sửa danh mục này (chỉ người tạo hoặc Admin mới có quyền).');
    }

    if (name && name.trim()) ws.name = name.trim();
    if (visibility && ['public', 'shared', 'private'].includes(visibility)) {
      ws.visibility = visibility;
      ws.isPublic = visibility === 'public';
      ws.isShared = visibility === 'shared' || visibility === 'public';
    }
    if (strategyProfileId !== undefined) {
      ws.strategyProfileId = strategyProfileId || null;
    }
    ws.updatedAt = new Date().toISOString();

    setLocalData(STORAGE_KEYS.WORKSPACES, list);
    return ws;
  }

  async function setWorkspaceStrategyProfile(workspaceId, strategyProfileId) {
    if (isGuest()) throw new Error('Vui lòng đăng nhập.');
    const list = getLocalData(STORAGE_KEYS.WORKSPACES, []);
    const ws = list.find(w => w.id === workspaceId);
    if (!ws) throw new Error('Không tìm thấy danh mục/dự án.');

    if (!isAdmin() && !canManageWorkspace(ws)) {
      throw new Error('Chỉ Quản trị viên (Admin) hoặc người tạo dự án mới có quyền gán Profile Chiến lược.');
    }

    ws.strategyProfileId = strategyProfileId || null;
    ws.updatedAt = new Date().toISOString();
    setLocalData(STORAGE_KEYS.WORKSPACES, list);
    return ws;
  }

  function getWorkspaceStrategyProfile(workspaceId) {
    const list = getLocalData(STORAGE_KEYS.WORKSPACES, []);
    const ws = list.find(w => w.id === workspaceId);
    return ws ? ws.strategyProfileId : null;
  }

  async function toggleWorkspacePublic(workspaceId, isPublic) {
    const list = getLocalData(STORAGE_KEYS.WORKSPACES, []);
    const ws = list.find(w => w.id === workspaceId);
    if (!ws) throw new Error('Không tìm thấy danh mục.');
    if (!canManageWorkspace(ws)) throw new Error('Chỉ người tạo hoặc Admin mới có quyền thay đổi trạng thái danh mục.');

    const nextVis = isPublic ? 'public' : 'shared';
    return updateWorkspace(workspaceId, { visibility: nextVis });
  }

  async function deleteWorkspace(workspaceId) {
    if (isGuest()) throw new Error('Vui lòng đăng nhập.');
    if (workspaceId === 'ws_toan_quoc_2027_2028') {
      throw new Error('Không thể xóa danh mục mặc định toàn quốc.');
    }
    const list = getLocalData(STORAGE_KEYS.WORKSPACES, []);
    const ws = list.find(w => w.id === workspaceId);
    if (!ws) throw new Error('Không tìm thấy danh mục.');

    if (!canManageWorkspace(ws)) {
      throw new Error('Bạn không có quyền xóa danh mục này (chỉ người tạo hoặc Admin mới có quyền).');
    }

    const filtered = list.filter(w => w.id !== workspaceId);
    setLocalData(STORAGE_KEYS.WORKSPACES, filtered);
    return true;
  }

  // ==================== LƯU TRỮ DỮ LIỆU MẢNG (< 1MB GUARD) ====================

  /**
   * Lưu dữ liệu mảng, kiểm tra nghiêm ngặt kích thước < 1MB
   */
  async function saveSectorDataToFirestore(workspaceId, sectorKey, rows, rawBase64 = null) {
    if (isGuest()) throw new Error('Vui lòng đăng nhập để lưu dữ liệu.');
    if (!canAccessSector(sectorKey, workspaceId)) {
      throw new Error(`Bạn không có quyền thao tác trên mảng ${SECTOR_NAMES[sectorKey] || sectorKey} trong dự án này.`);
    }

    const list = getLocalData(STORAGE_KEYS.WORKSPACES, []);
    const ws = list.find(w => w.id === workspaceId);
    if (ws && !canEditWorkspaceSectors(ws)) {
      throw new Error(`Danh mục "${ws.name}" ở chế độ "Mình tôi", chỉ người tạo (${ws.createdBy}) hoặc Admin mới có quyền nạp/chỉnh sửa.`);
    }

    const payload = {
      workspaceId,
      sectorKey,
      rowCount: (rows || []).length,
      rows: rows || [],
      rawBase64: rawBase64 || null,
      updatedByUid: currentUserProfile.uid,
      updatedByName: currentUserProfile.displayName || currentUserProfile.username,
      updatedAt: new Date().toISOString()
    };

    const jsonStr = JSON.stringify(payload);
    const sizeBytes = new Blob([jsonStr]).size;
    const MAX_ALLOWED_BYTES = 950 * 1024; // 950KB để đảm bảo an toàn hạn mức Firestore 1MB

    if (sizeBytes > MAX_ALLOWED_BYTES) {
      if (payload.rawBase64) {
        payload.rawBase64 = null;
        const strippedSize = new Blob([JSON.stringify(payload)]).size;
        if (strippedSize > MAX_ALLOWED_BYTES) {
          throw new Error(
            `Dữ liệu mảng ${sectorKey} quá lớn (${(strippedSize / 1024 / 1024).toFixed(2)} MB), vượt quá hạn mức 1MB của Firestore. Vui lòng kiểm tra lại file!`
          );
        }
      } else {
        throw new Error(
          `Dữ liệu mảng ${sectorKey} quá lớn (${(sizeBytes / 1024 / 1024).toFixed(2)} MB), vượt quá hạn mức 1MB của Firestore. Vui lòng kiểm tra lại file!`
        );
      }
    }

    const storageKey = `${STORAGE_KEYS.SECTOR_DATA_PREFIX}${workspaceId}_${sectorKey}`;
    setLocalData(storageKey, payload);

    return { success: true, sizeBytes };
  }

  /**
   * Tải toàn bộ dữ liệu 7 mảng của 1 workspace
   */
  /**
   * Xóa dữ liệu của một mảng trong workspace (kiểm tra phân quyền)
   */
  async function deleteSectorDataFromFirestore(workspaceId, sectorKey) {
    if (isGuest()) throw new Error('Vui lòng đăng nhập.');
    if (!canAccessSector(sectorKey, workspaceId)) {
      throw new Error(`Bạn không có quyền xóa dữ liệu mảng ${SECTOR_NAMES[sectorKey] || sectorKey} trong dự án này.`);
    }

    const list = getLocalData(STORAGE_KEYS.WORKSPACES, []);
    const ws = list.find(w => w.id === workspaceId);
    if (ws && !canEditWorkspaceSectors(ws)) {
      throw new Error(`Danh mục "${ws.name}" ở chế độ "Mình tôi", chỉ người tạo (${ws.createdBy}) hoặc Admin mới có quyền thao tác.`);
    }

    const storageKey = `${STORAGE_KEYS.SECTOR_DATA_PREFIX}${workspaceId}_${sectorKey}`;
    try {
      localStorage.removeItem(storageKey);
    } catch (e) {
      console.warn('Lỗi removeItem:', e);
    }

    return true;
  }

  async function loadWorkspaceDataFromFirestore(workspaceId) {
    const sectorsData = {};
    ALL_SECTORS.forEach(secKey => {
      const storageKey = `${STORAGE_KEYS.SECTOR_DATA_PREFIX}${workspaceId}_${secKey}`;
      const item = getLocalData(storageKey, null);
      if (item) {
        sectorsData[secKey] = item;
      }
    });
    return sectorsData;
  }

  async function adminGetAllWorkspaces() {
    if (isGuest()) throw new Error('Vui lòng đăng nhập để quản lý danh mục.');
    const source = getLocalData(STORAGE_KEYS.WORKSPACES, []);
    const list = isAdmin()
      ? source
      : source.filter(w => {
          const vis = getWorkspaceVisibility(w);
          if (vis === 'public' || vis === 'shared') return true;
          return (w.createdBy && w.createdBy === currentUserProfile.username) ||
                 (w.creatorUid && w.creatorUid === currentUserProfile.uid);
        });
    return list.map(ws => {
      const sectors = {};
      let totalRows = 0;
      ALL_SECTORS.forEach(secKey => {
        const item = getLocalData(`${STORAGE_KEYS.SECTOR_DATA_PREFIX}${ws.id}_${secKey}`, null);
        if (item && item.rowCount > 0) {
          sectors[secKey] = item.rowCount;
          totalRows += item.rowCount;
        }
      });
      return {
        ...ws,
        sectorsLoaded: sectors,
        totalRows
      };
    });
  }

  async function adminUpdateWorkspace(workspaceId, { name, visibility, isPublic, isShared, strategyProfileId }) {
    if (!isAdmin()) throw new Error('Chỉ Admin mới có quyền cập nhật danh mục.');
    const list = getLocalData(STORAGE_KEYS.WORKSPACES, []);
    const ws = list.find(w => w.id === workspaceId);
    if (!ws) throw new Error('Không tìm thấy danh mục.');

    if (name && name.trim()) ws.name = name.trim();
    if (visibility && ['public', 'shared', 'private'].includes(visibility)) {
      ws.visibility = visibility;
      ws.isPublic = visibility === 'public';
      ws.isShared = visibility === 'shared' || visibility === 'public';
    } else if (typeof isPublic === 'boolean' || typeof isShared === 'boolean') {
      let nextVisibility = getWorkspaceVisibility(ws);
      if (isPublic === true) {
        nextVisibility = 'public';
      } else if (isShared === true) {
        nextVisibility = 'shared';
      } else if (isShared === false) {
        nextVisibility = 'private';
      }
      ws.visibility = nextVisibility;
      ws.isPublic = nextVisibility === 'public';
      ws.isShared = nextVisibility === 'shared' || nextVisibility === 'public';
    }
    if (strategyProfileId !== undefined) {
      ws.strategyProfileId = strategyProfileId || null;
    }
    ws.updatedAt = new Date().toISOString();

    setLocalData(STORAGE_KEYS.WORKSPACES, list);
    return ws;
  }

  // Xuất ra window.FirebaseService
  window.FirebaseService = {
    loginUser,
    logoutUser,
    changeUserPassword,
    getCurrentUser,
    isGuest,
    isAdmin,
    canAccessSector,
    adminCreateUser,
    adminGetUsersList,
    adminUpdateUserPermissions,
    adminDeleteUser,
    listWorkspaces,
    createWorkspace,
    updateWorkspace,
    setWorkspaceStrategyProfile,
    getWorkspaceStrategyProfile,
    toggleWorkspacePublic,
    deleteWorkspace,
    canManageWorkspace,
    canEditWorkspaceSectors,
    getWorkspaceVisibility,
    adminGetAllWorkspaces,
    adminUpdateWorkspace,
    saveSectorDataToFirestore,
    deleteSectorDataFromFirestore,
    loadWorkspaceDataFromFirestore,
    onAuthStateChange,
    ALL_SECTORS,
    SECTOR_NAMES
  };

})();
