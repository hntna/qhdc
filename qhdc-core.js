// qhdc-core.js - Quản lý tài khoản, phân quyền danh mục và lưu trữ dữ liệu QHDC3
// Dữ liệu nghiệp vụ lưu tập trung trên Firebase/Firestore.

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
    CURRENT_USER: 'qhdc_current_user'
  };

  let currentUserProfile = null;
  const authListeners = [];
  let cachedWorkspaces = [];
  let firestoreDb = null;
  let firestoreEnabled = false;
  let firebaseInitError = null;

  const FIRESTORE_COLLECTIONS = {
    USERS: 'qhdc_users',
    WORKSPACES: 'qhdc_workspaces'
  };

  const DEFAULT_WORKSPACE = {
    id: 'ws_toan_quoc_2027_2028',
    name: 'QHĐC Toàn quốc 2027–2028',
    visibility: 'public',
    isPublic: true,
    isShared: true,
    createdBy: 'admin',
    creatorUid: 'uid_admin'
  };

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
    }
  }

  function initFirestore() {
    const config = window.QHDC_FIREBASE_CONFIG;
    if (!config || !config.apiKey || !config.projectId) return false;
    if (!window.firebase || typeof window.firebase.initializeApp !== 'function') {
      firebaseInitError = 'Chưa tải Firebase SDK.';
      return false;
    }

    try {
      const app = window.firebase.apps && window.firebase.apps.length
        ? window.firebase.app()
        : window.firebase.initializeApp(config);
      firestoreDb = app.firestore ? app.firestore() : window.firebase.firestore();
      firestoreEnabled = !!firestoreDb;
      return firestoreEnabled;
    } catch (err) {
      firebaseInitError = err;
      console.warn('Không khởi tạo được Firestore:', err);
      firestoreDb = null;
      firestoreEnabled = false;
      return false;
    }
  }

  function isFirestoreEnabled() {
    return firestoreEnabled && !!firestoreDb;
  }

  function requireFirestore() {
    if (isFirestoreEnabled()) return;
    const detail = firebaseInitError
      ? ` (${firebaseInitError.message || firebaseInitError})`
      : '';
    throw new Error(`Chưa kết nối Firebase/Firestore${detail}. Kiểm tra qhdc-firebase-config.js, Firebase SDK và quyền Firestore.`);
  }

  function toPlainData(value) {
    return JSON.parse(JSON.stringify(value || null));
  }

  function makeDocId(prefix) {
    return `${prefix}_${Date.now()}_${Math.random().toString(36).substring(2, 8)}`;
  }

  function estimateBytes(value) {
    return new Blob([JSON.stringify(value || null)]).size;
  }

  function splitRowsIntoChunks(rows, maxBytes = 700 * 1024) {
    const chunks = [];
    let current = [];
    let currentBytes = estimateBytes({ rows: current });

    (rows || []).forEach(row => {
      const nextBytes = estimateBytes({ rows: current.concat([row]) });
      if (current.length && nextBytes > maxBytes) {
        chunks.push(current);
        current = [row];
        currentBytes = estimateBytes({ rows: current });
      } else {
        current.push(row);
        currentBytes = nextBytes;
      }
    });

    if (current.length || !chunks.length) chunks.push(current);
    return chunks;
  }

  async function commitInBatches(writeFns) {
    const MAX_BATCH_WRITES = 450;
    for (let i = 0; i < writeFns.length; i += MAX_BATCH_WRITES) {
      const batch = firestoreDb.batch();
      writeFns.slice(i, i + MAX_BATCH_WRITES).forEach(fn => fn(batch));
      await batch.commit();
    }
  }

  function workspaceRef(workspaceId) {
    return firestoreDb.collection(FIRESTORE_COLLECTIONS.WORKSPACES).doc(workspaceId);
  }

  async function getWorkspaceById(workspaceId) {
    if (cachedWorkspaces && cachedWorkspaces.length) {
      const found = cachedWorkspaces.find(w => w.id === workspaceId);
      if (found) return found;
    }
    requireFirestore();
    const doc = await workspaceRef(workspaceId).get();
    const ws = doc.exists ? normalizeWorkspace({ id: doc.id, ...doc.data() }) : null;
    if (ws) {
      const idx = cachedWorkspaces.findIndex(w => w.id === ws.id);
      if (idx >= 0) cachedWorkspaces[idx] = ws;
      else cachedWorkspaces.push(ws);
    }
    return ws;
  }

  function normalizeWorkspace(ws) {
    const vis = getWorkspaceVisibility(ws);
    return {
      ...ws,
      visibility: vis,
      isPublic: vis === 'public',
      isShared: vis === 'shared' || vis === 'public',
      strategyProfileId: ws.strategyProfileId || null
    };
  }

  async function ensureDefaultWorkspaceRemote() {
    if (!isFirestoreEnabled()) return;
    const ref = workspaceRef(DEFAULT_WORKSPACE.id);
    const doc = await ref.get();
    if (!doc.exists) {
      await ref.set({
        ...DEFAULT_WORKSPACE,
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString()
      }, { merge: true });
    }
  }

  async function getRemoteUsersByUsername(username) {
    const snap = await firestoreDb
      .collection(FIRESTORE_COLLECTIONS.USERS)
      .where('username', '==', username)
      .limit(1)
      .get();
    return snap.empty ? null : { uid: snap.docs[0].id, ...snap.docs[0].data() };
  }

  // Khôi phục phiên đăng nhập trước đó nếu có
  function initSession() {
    initFirestore();
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
    requireFirestore();

    if (isFirestoreEnabled()) {
      let found = await getRemoteUsersByUsername(cleanUser);
      if (cleanUser === 'admin' && !found) {
        found = {
          uid: 'uid_admin',
          username: 'admin',
          password: cleanPass,
          displayName: 'Quản trị viên Hệ thống',
          dept: 'TT, KTTC',
          role: 'admin',
          assignedSectors: [...ALL_SECTORS],
          createdAt: new Date().toISOString()
        };
        await firestoreDb.collection(FIRESTORE_COLLECTIONS.USERS).doc(found.uid).set(found, { merge: true });
        await ensureDefaultWorkspaceRemote();
      }

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
      return { success: true, profile: currentUserProfile, isNewAdmin: cleanUser === 'admin' && found.createdAt };
    }

    throw new Error('Chưa kết nối Firebase/Firestore.');
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
    requireFirestore();
    await firestoreDb.collection(FIRESTORE_COLLECTIONS.USERS).doc(currentUserProfile.uid).set({
      password: newPassword,
      updatedAt: new Date().toISOString()
    }, { merge: true });
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
   * Kiểm tra quyền thao tác trên một mảng
   * - Nếu là người tạo/chủ sở hữu dự án đó (hoặc Admin): Toàn quyền mọi mảng trong dự án.
   * - Nếu là dự án chung (shared / public): Phải được phân quyền dự án (assignedWorkspaces)
   *   VÀ phân quyền mảng chuyên môn (assignedSectors).
   */
  function canAccessSector(sectorKey, workspaceOrId = null) {
    if (!currentUserProfile) return false;
    if (currentUserProfile.role === 'admin') return true;

    // 1. Xác định workspace cần kiểm tra
    let ws = null;
    let wsCreator = null;
    let wsCreatorUid = null;
    let wsId = null;

    if (workspaceOrId && typeof workspaceOrId === 'object') {
      ws = workspaceOrId;
      wsId = ws.id;
      wsCreator = ws.createdBy;
      wsCreatorUid = ws.creatorUid;
    } else {
      wsId = workspaceOrId || (typeof document !== 'undefined' ? document.getElementById('selectWorkspace')?.value : null);
      if (wsId) {
        ws = cachedWorkspaces.find(w => w.id === wsId);
        if (ws) {
          wsCreator = ws.createdBy;
          wsCreatorUid = ws.creatorUid;
        } else if (typeof document !== 'undefined') {
          const select = document.getElementById('selectWorkspace');
          const opt = select?.querySelector(`option[value="${wsId}"]`);
          if (opt) {
            wsCreator = opt.dataset.creator;
            wsCreatorUid = opt.dataset.creatorUid;
            ws = { id: wsId, createdBy: wsCreator, creatorUid: wsCreatorUid, visibility: opt.dataset.visibility };
          }
        }
      }
    }

    // 2. Nếu là người tạo danh mục (hoặc Admin): Toàn quyền với mọi mảng trong danh mục đó!
    const curUser = (currentUserProfile?.username || '').trim().toLowerCase();
    const curUid = currentUserProfile?.uid;
    const isOwner = (wsCreator && wsCreator.trim().toLowerCase() === curUser) ||
                    (wsCreatorUid && wsCreatorUid === curUid) ||
                    (ws && canManageWorkspace(ws));
    if (isOwner) {
      return true;
    }

    // 3. Nếu là dự án ở chế độ "private" (Mình tôi) và không phải owner -> Không có quyền
    if (ws && getWorkspaceVisibility(ws) === 'private') {
      return false;
    }

    // 4. Nếu là dự án chung (shared / public):
    // Cần thỏa mãn CẢ HAI điều kiện:
    // a) Phải được gán quyền thao tác trên Dự án này:
    const assignedWs = currentUserProfile.assignedWorkspaces || [];
    const targetWsId = wsId || ws?.id;
    if (!targetWsId || !assignedWs.includes(targetWsId)) {
      return false;
    }

    // b) Phải được gán quyền trên Mảng chuyên môn này:
    const assignedSec = currentUserProfile.assignedSectors || [];
    return assignedSec.includes(sectorKey);
  }

  // ==================== QUẢN LÝ USER (ADMIN PANEL) ====================

  async function adminGetUsersList() {
    if (!isAdmin()) throw new Error('Chỉ Admin mới có quyền xem danh sách user.');
    requireFirestore();
    if (isFirestoreEnabled()) {
      const snap = await firestoreDb.collection(FIRESTORE_COLLECTIONS.USERS).get();
      return snap.docs.map(doc => {
        const copy = { uid: doc.id, ...doc.data() };
        delete copy.password;
        return copy;
      });
    }
    return [];
  }

  async function adminCreateUser(username, password, displayName, assignedSectors = [], assignedWorkspaces = []) {
    if (!isAdmin()) throw new Error('Chỉ Admin mới có quyền tạo tài khoản.');
    requireFirestore();
    const cleanUser = (username || '').trim().toLowerCase();
    const cleanPass = (password || '').trim();

    if (!cleanUser || !cleanPass) throw new Error('Vui lòng điền đủ Tên đăng nhập và Mật khẩu.');
    if (cleanPass.length < 6) throw new Error('Mật khẩu phải từ 6 ký tự trở lên.');

    const newUser = {
      uid: makeDocId('uid'),
      username: cleanUser,
      password: cleanPass,
      displayName: (displayName || '').trim() || cleanUser,
      dept: 'TT, KTTC',
      role: 'user',
      assignedSectors: assignedSectors || [],
      assignedWorkspaces: assignedWorkspaces || [],
      createdAt: new Date().toISOString()
    };

    if (isFirestoreEnabled()) {
      const existing = await getRemoteUsersByUsername(cleanUser);
      if (existing) throw new Error(`Tên đăng nhập "${username}" đã tồn tại!`);
      await firestoreDb.collection(FIRESTORE_COLLECTIONS.USERS).doc(newUser.uid).set(newUser);
      return newUser;
    }

    throw new Error('Chưa kết nối Firebase/Firestore.');
  }

  async function adminUpdateUserPermissions(uid, newSectors, newWorkspaces = null) {
    if (!isAdmin()) throw new Error('Chỉ Admin mới có quyền sửa phân quyền.');
    requireFirestore();
    if (isFirestoreEnabled()) {
      const updateData = {
        assignedSectors: newSectors,
        updatedAt: new Date().toISOString()
      };
      if (Array.isArray(newWorkspaces)) {
        updateData.assignedWorkspaces = newWorkspaces;
      }
      await firestoreDb.collection(FIRESTORE_COLLECTIONS.USERS).doc(uid).set(updateData, { merge: true });

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
    throw new Error('Chưa kết nối Firebase/Firestore.');
  }

  async function adminDeleteUser(uid) {
    if (!isAdmin()) throw new Error('Chỉ Admin mới có quyền xóa tài khoản.');
    requireFirestore();
    if (isFirestoreEnabled()) {
      const ref = firestoreDb.collection(FIRESTORE_COLLECTIONS.USERS).doc(uid);
      const doc = await ref.get();
      const target = doc.exists ? doc.data() : null;
      if (target && target.username && target.username.toLowerCase() === 'admin') {
        throw new Error('Không thể xóa tài khoản Quản trị viên tối cao.');
      }
      await ref.delete();
      return true;
    }
    throw new Error('Chưa kết nối Firebase/Firestore.');
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
    const curUser = (currentUserProfile?.username || '').trim().toLowerCase();
    const curUid = currentUserProfile?.uid;
    const wsCreator = (ws.createdBy || '').trim().toLowerCase();
    if ((wsCreator && wsCreator === curUser) || (ws.creatorUid && ws.creatorUid === curUid)) {
      return true;
    }
    const vis = getWorkspaceVisibility(ws);
    if (vis === 'public' || vis === 'shared') {
      const assignedWs = currentUserProfile?.assignedWorkspaces || [];
      return assignedWs.includes(ws.id);
    }
    return false;
  }

  async function listWorkspaces() {
    requireFirestore();
    if (isFirestoreEnabled()) {
      await ensureDefaultWorkspaceRemote();
      const snap = await firestoreDb.collection(FIRESTORE_COLLECTIONS.WORKSPACES).get();
      const list = snap.docs.map(doc => normalizeWorkspace({ id: doc.id, ...doc.data() }));
      cachedWorkspaces = list;
      return list.filter(w => {
        const vis = getWorkspaceVisibility(w);
        if (isAdmin()) return true;
        if (isGuest()) return vis === 'public';
        if (vis === 'public' || vis === 'shared') return true;
        const curUser = (currentUserProfile?.username || '').trim().toLowerCase();
        const curUid = currentUserProfile?.uid;
        const wsCreator = (w.createdBy || '').trim().toLowerCase();
        return (wsCreator && wsCreator === curUser) ||
               (w.creatorUid && w.creatorUid === curUid);
      });
    }

    return [];
  }

  async function createWorkspace(name, visibility = 'shared', strategyProfileId = null) {
    if (isGuest()) throw new Error('Vui lòng đăng nhập để tạo danh mục.');
    requireFirestore();
    const cleanName = (name || '').trim();
    if (!cleanName) throw new Error('Tên danh mục không được để trống.');

    const wsId = makeDocId('ws');
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
    if (isFirestoreEnabled()) {
      await workspaceRef(wsId).set(newWs);
      cachedWorkspaces.push(newWs);
      return newWs;
    }
    throw new Error('Chưa kết nối Firebase/Firestore.');
  }

  async function updateWorkspace(workspaceId, { name, visibility, strategyProfileId }) {
    if (isGuest()) throw new Error('Vui lòng đăng nhập.');
    requireFirestore();
    const ws = await getWorkspaceById(workspaceId);
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
      try {
        const localMap = JSON.parse(localStorage.getItem('qhdc_ws_strategy_profiles') || '{}');
        localMap[workspaceId] = ws.strategyProfileId;
        localStorage.setItem('qhdc_ws_strategy_profiles', JSON.stringify(localMap));
      } catch (e) {}
    }
    ws.updatedAt = new Date().toISOString();

    if (isFirestoreEnabled()) {
      await workspaceRef(workspaceId).set(normalizeWorkspace(ws), { merge: true });
      return normalizeWorkspace(ws);
    }

    throw new Error('Chưa kết nối Firebase/Firestore.');
  }

  async function setWorkspaceStrategyProfile(workspaceId, strategyProfileId) {
    if (isGuest()) throw new Error('Vui lòng đăng nhập.');
    let ws = null;
    try {
      ws = await getWorkspaceById(workspaceId);
    } catch (e) {
      console.warn('getWorkspaceById warning:', e);
    }

    if (!ws) {
      if (workspaceId === DEFAULT_WORKSPACE.id) {
        ws = { ...DEFAULT_WORKSPACE };
      } else {
        ws = { id: workspaceId, name: 'Dự án', createdBy: currentUserProfile?.username || 'admin' };
      }
    }

    if (!isAdmin() && !canManageWorkspace(ws)) {
      throw new Error('Chỉ Quản trị viên (Admin) hoặc người tạo dự án mới có quyền gán Profile Chiến lược.');
    }

    ws.strategyProfileId = strategyProfileId || null;
    ws.updatedAt = new Date().toISOString();

    if (cachedWorkspaces) {
      const idx = cachedWorkspaces.findIndex(w => w.id === workspaceId);
      if (idx >= 0) cachedWorkspaces[idx] = ws;
      else cachedWorkspaces.push(ws);
    }

    try {
      const localMap = JSON.parse(localStorage.getItem('qhdc_ws_strategy_profiles') || '{}');
      localMap[workspaceId] = ws.strategyProfileId;
      localStorage.setItem('qhdc_ws_strategy_profiles', JSON.stringify(localMap));
    } catch (e) {}

    if (isFirestoreEnabled()) {
      try {
        await workspaceRef(workspaceId).set(normalizeWorkspace(ws), { merge: true });
      } catch (err) {
        console.warn('Lỗi ghi Firestore setWorkspaceStrategyProfile:', err);
      }
    }

    return normalizeWorkspace(ws);
  }

  function getWorkspaceStrategyProfile(workspaceId) {
    if (!workspaceId) return null;
    if (cachedWorkspaces && cachedWorkspaces.length) {
      const ws = cachedWorkspaces.find(w => w.id === workspaceId);
      if (ws && ws.strategyProfileId) return ws.strategyProfileId;
    }
    try {
      const localMap = JSON.parse(localStorage.getItem('qhdc_ws_strategy_profiles') || '{}');
      if (localMap[workspaceId]) return localMap[workspaceId];
    } catch (e) {}
    return null;
  }

  async function toggleWorkspacePublic(workspaceId, isPublic) {
    requireFirestore();
    const ws = await getWorkspaceById(workspaceId);
    if (!ws) throw new Error('Không tìm thấy danh mục.');
    if (!canManageWorkspace(ws)) throw new Error('Chỉ người tạo hoặc Admin mới có quyền thay đổi trạng thái danh mục.');

    const nextVis = isPublic ? 'public' : 'shared';
    return updateWorkspace(workspaceId, { visibility: nextVis });
  }

  async function deleteWorkspace(workspaceId) {
    if (isGuest()) throw new Error('Vui lòng đăng nhập.');
    requireFirestore();
    if (workspaceId === 'ws_toan_quoc_2027_2028') {
      throw new Error('Không thể xóa danh mục mặc định toàn quốc.');
    }
    const ws = await getWorkspaceById(workspaceId);
    if (!ws) throw new Error('Không tìm thấy danh mục.');

    if (!canManageWorkspace(ws)) {
      throw new Error('Bạn không có quyền xóa danh mục này (chỉ người tạo hoặc Admin mới có quyền).');
    }

    if (isFirestoreEnabled()) {
      const ref = workspaceRef(workspaceId);
      const sectorsSnap = await ref.collection('sectors').get();
      const writeFns = [];
      for (const sectorDoc of sectorsSnap.docs) {
        const chunksSnap = await sectorDoc.ref.collection('chunks').get();
        chunksSnap.docs.forEach(chunkDoc => writeFns.push(batch => batch.delete(chunkDoc.ref)));
        writeFns.push(batch => batch.delete(sectorDoc.ref));
      }
      writeFns.push(batch => batch.delete(ref));
      await commitInBatches(writeFns);
      return true;
    }

    throw new Error('Chưa kết nối Firebase/Firestore.');
  }

  // ==================== LƯU TRỮ DỮ LIỆU MẢNG (< 1MB GUARD) ====================

  /**
   * Lưu dữ liệu mảng, kiểm tra nghiêm ngặt kích thước < 1MB
   */
  async function saveSectorDataToFirestore(workspaceId, sectorKey, rows, rawBase64 = null) {
    if (isGuest()) throw new Error('Vui lòng đăng nhập để lưu dữ liệu.');
    requireFirestore();
    if (!canAccessSector(sectorKey, workspaceId)) {
      throw new Error(`Bạn không có quyền thao tác trên mảng ${SECTOR_NAMES[sectorKey] || sectorKey}.`);
    }

    const ws = await getWorkspaceById(workspaceId);
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

    if (isFirestoreEnabled()) {
      const cleanRows = toPlainData(rows || []);
      const chunks = splitRowsIntoChunks(cleanRows);
      const sectorRef = workspaceRef(workspaceId).collection('sectors').doc(sectorKey);
      const oldChunksSnap = await sectorRef.collection('chunks').get();
      const writeFns = [];

      oldChunksSnap.docs.forEach(doc => writeFns.push(batch => batch.delete(doc.ref)));
      writeFns.push(batch => batch.set(sectorRef, {
        workspaceId,
        sectorKey,
        rowCount: cleanRows.length,
        chunkCount: chunks.length,
        sizeBytes: estimateBytes({ rows: cleanRows }),
        updatedByUid: currentUserProfile.uid,
        updatedByName: currentUserProfile.displayName || currentUserProfile.username,
        updatedAt: new Date().toISOString()
      }, { merge: true }));

      chunks.forEach((chunkRows, index) => {
        writeFns.push(batch => batch.set(sectorRef.collection('chunks').doc(String(index).padStart(4, '0')), {
          index,
          rows: chunkRows
        }));
      });

      await commitInBatches(writeFns);
      await workspaceRef(workspaceId).set({
        updatedAt: new Date().toISOString(),
        updatedByUid: currentUserProfile.uid,
        updatedByName: currentUserProfile.displayName || currentUserProfile.username
      }, { merge: true });
      return { success: true, sizeBytes: estimateBytes({ rows: cleanRows }), chunkCount: chunks.length };
    }

    throw new Error('Chưa kết nối Firebase/Firestore.');
  }

  /**
   * Tải toàn bộ dữ liệu 7 mảng của 1 workspace
   */
  async function loadWorkspaceDataFromFirestore(workspaceId) {
    requireFirestore();
    const sectorsData = {};
    if (isFirestoreEnabled()) {
      for (const secKey of ALL_SECTORS) {
        const sectorRef = workspaceRef(workspaceId).collection('sectors').doc(secKey);
        const doc = await sectorRef.get();
        if (!doc.exists) continue;
        const meta = doc.data();
        let rows = meta.rows || [];
        if (meta.chunkCount) {
          const chunksSnap = await sectorRef.collection('chunks').orderBy('index').get();
          rows = [];
          chunksSnap.docs.forEach(chunkDoc => {
            const chunk = chunkDoc.data();
            rows.push(...(chunk.rows || []));
          });
        }
        sectorsData[secKey] = {
          ...meta,
          rows,
          rowCount: meta.rowCount || rows.length
        };
      }
      return sectorsData;
    }

    return sectorsData;
  }

  /**
   * Xóa dữ liệu của một mảng trong danh mục (kiểm tra phân quyền)
   */
  async function deleteSectorDataFromFirestore(workspaceId, sectorKey) {
    if (isGuest()) throw new Error('Vui lòng đăng nhập.');
    requireFirestore();

    if (!canAccessSector(sectorKey, workspaceId)) {
      throw new Error(`Bạn không có quyền xóa dữ liệu mảng ${SECTOR_NAMES[sectorKey] || sectorKey} trong danh mục này.`);
    }

    const ws = await getWorkspaceById(workspaceId);
    if (ws && !canEditWorkspaceSectors(ws)) {
      throw new Error(`Danh mục "${ws.name}" ở chế độ "Mình tôi", chỉ người tạo (${ws.createdBy}) hoặc Admin mới có quyền thao tác.`);
    }

    if (isFirestoreEnabled()) {
      const sectorRef = workspaceRef(workspaceId).collection('sectors').doc(sectorKey);
      const oldChunksSnap = await sectorRef.collection('chunks').get();
      const writeFns = [];
      oldChunksSnap.docs.forEach(doc => writeFns.push(batch => batch.delete(doc.ref)));
      writeFns.push(batch => batch.delete(sectorRef));
      await commitInBatches(writeFns);

      await workspaceRef(workspaceId).set({
        updatedAt: new Date().toISOString(),
        updatedByUid: currentUserProfile.uid,
        updatedByName: currentUserProfile.displayName || currentUserProfile.username
      }, { merge: true });

      return true;
    }

    throw new Error('Chưa kết nối Firebase/Firestore.');
  }

  async function adminGetAllWorkspaces() {
    if (isGuest()) throw new Error('Vui lòng đăng nhập để quản lý danh mục.');
    requireFirestore();
    if (isFirestoreEnabled()) {
      const sourceSnap = await firestoreDb.collection(FIRESTORE_COLLECTIONS.WORKSPACES).get();
      const source = sourceSnap.docs.map(doc => normalizeWorkspace({ id: doc.id, ...doc.data() }));
      const list = isAdmin()
        ? source
        : source.filter(w => {
            const vis = getWorkspaceVisibility(w);
            if (vis === 'public' || vis === 'shared') return true;
            return (w.createdBy && w.createdBy === currentUserProfile.username) ||
                   (w.creatorUid && w.creatorUid === currentUserProfile.uid);
          });

      const result = [];
      for (const ws of list) {
        const sectors = {};
        let totalRows = 0;
        const sectorsSnap = await workspaceRef(ws.id).collection('sectors').get();
        sectorsSnap.docs.forEach(doc => {
          const data = doc.data();
          if (data && data.rowCount > 0) {
            sectors[doc.id] = data.rowCount;
            totalRows += data.rowCount;
          }
        });
        result.push({ ...ws, sectorsLoaded: sectors, totalRows });
      }
      return result;
    }

    return [];
  }

  async function adminUpdateWorkspace(workspaceId, { name, visibility, isPublic, isShared }) {
    if (!isAdmin()) throw new Error('Chỉ Admin mới có quyền cập nhật danh mục.');
    requireFirestore();
    const ws = await getWorkspaceById(workspaceId);
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
    ws.updatedAt = new Date().toISOString();

    if (isFirestoreEnabled()) {
      await workspaceRef(workspaceId).set(normalizeWorkspace(ws), { merge: true });
      return normalizeWorkspace(ws);
    }

    throw new Error('Chưa kết nối Firebase/Firestore.');
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
    isFirestoreEnabled,
    getBackendMode: () => isFirestoreEnabled() ? 'firestore' : 'not-connected',
    getFirebaseInitError: () => firebaseInitError,
    ALL_SECTORS,
    SECTOR_NAMES
  };

})();
