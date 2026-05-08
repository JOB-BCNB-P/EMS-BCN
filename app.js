// ======================== STATE ========================
let APP = {
  currentUser: null, currentRole: null, currentPage: 'dashboard', sidebarOpen: false,
  allData: [],
  config: { system_title: 'ระบบบริหารจัดการงานวิชาการ (EMS-BCNB)', college_name: 'วิทยาลัยพยาบาลบรมราชชนนี กรุงเทพ' },
  permissions: { admin: { dashboard: 1, students: 1, subjects: 1, schedule: 1, grades: 1, engResults: 1, teachers: 1, teacherDirectory: 1, services: 1, tracking: 1, gradeTracking: 1, fileTracking: 1, leave: 1, settings: 1, loginLog: 1 }, academic: { dashboard: 1, students: 1, subjects: 1, schedule: 1, grades: 1, engResults: 1, teachers: 1, teacherDirectory: 1, services: 1, tracking: 1, gradeTracking: 1, fileTracking: 1, leave: 1, settings: 1 }, teacher: { dashboard: 1, students: 1, subjects: 1, grades: 1, engResults: 1, tracking: 1, gradeTracking: 1, fileTracking: 1, leave: 1 }, classTeacher: { dashboard: 1, students: 1, subjects: 1, grades: 1, engResults: 1, tracking: 1, gradeTracking: 1, fileTracking: 1, leave: 1 }, student: { dashboard: 1, students: 1, grades: 1, engResults: 1, leave: 1 }, executive: { dashboard: 1, students: 1, subjects: 1, schedule: 1, grades: 1, engResults: 1, teachers: 1, teacherDirectory: 1, tracking: 1, gradeTracking: 1, fileTracking: 1, leave: 1 } },
  filters: { semester: '', academicYear: '', search: '', yearLevel: '' },
  pagination: { page: 1, perPage: 10 }
};

function getDataByType(t) { return APP.allData.filter(d => d.type === t) }

// ======================== HOMEROOM NUMBERS (per year level) ========================
// Stored in localStorage so admin/academic can configure without modifying the spreadsheet
const DEFAULT_HOMEROOMS = { '1': '', '2': '', '3': '', '4': '' };
function getHomeroomNumbers() {
  try {
    const raw = localStorage.getItem('homeroomNumbers');
    if (!raw) return { ...DEFAULT_HOMEROOMS };
    const parsed = JSON.parse(raw);
    return { ...DEFAULT_HOMEROOMS, ...parsed };
  } catch { return { ...DEFAULT_HOMEROOMS }; }
}
function setHomeroomNumber(yr, value) {
  const all = getHomeroomNumbers();
  all[String(yr)] = (value || '').trim();
  localStorage.setItem('homeroomNumbers', JSON.stringify(all));
}
function promptEditHomeroom(yr) {
  const role = APP.currentUser && APP.currentUser.role;
  if (role !== 'admin' && role !== 'academic') {
    showToast && showToast('เฉพาะผู้ดูแลระบบ/งานวิชาการเท่านั้นที่แก้ไขได้', 'error');
    return;
  }
  const current = getHomeroomNumbers()[String(yr)] || '';
  const v = prompt(`กรอกหมายเลขห้องเรียนประจำของชั้นปี ${yr}\n(เช่น 101, 202, ห้อง A เป็นต้น)`, current);
  if (v === null) return;
  setHomeroomNumber(yr, v);
  if (typeof renderCurrentPage === 'function') renderCurrentPage();
}

// ======================== LOGIN ACTIVITY LOG ========================
// Save login/logout events to Google Sheet "login_log" tab
const LOGIN_LOG_ROLE_LABEL = { admin: 'ผู้ดูแลระบบ', academic: 'เจ้าหน้าที่งานวิชาการ', teacher: 'อาจารย์', classTeacher: 'อาจารย์ประจำชั้น', student: 'นักศึกษา', executive: 'ผู้บริหาร' };
async function logLoginEvent(eventType, userInfo) {
  // eventType: 'login' | 'logout' | 'login_failed'
  try {
    const now = new Date();
    const pad = n => String(n).padStart(2, '0');
    const localTimestamp = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())} ${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;
    const obj = {
      type: 'login_log',
      event_type: eventType,
      user_name: (userInfo && userInfo.name) || '-',
      role: (userInfo && userInfo.role) || '-',
      role_label: (userInfo && userInfo.role && LOGIN_LOG_ROLE_LABEL[userInfo.role]) || '-',
      identifier: (userInfo && userInfo.identifier) || '',
      user_agent: (navigator && navigator.userAgent) ? String(navigator.userAgent).slice(0, 250) : '',
      timestamp: localTimestamp,
      created_at: now.toISOString()
    };
    if (GSheetDB && GSheetDB.hasWriteAccess && GSheetDB.hasWriteAccess()) {
      // Fire and forget - don't block UI
      GSheetDB.create(obj).catch(err => console.warn('logLoginEvent error:', err));
    } else {
      console.warn('logLoginEvent: no write access (Apps Script URL not configured)');
    }
  } catch (err) {
    console.warn('logLoginEvent failed:', err);
  }
}

// Look up student name from student_id
function getStudentName(studentId) {
  if (!studentId) return '';
  const stu = getDataByType('student').find(s => s.student_id === studentId);
  return stu ? stu.name : '';
}

// Build student <option> list for forms
function studentOptionsHTML(selectedId) {
  const students = getDataByType('student');
  return `<option value="">-- เลือกนักศึกษา --</option>` +
    students.map(s => `<option value="${s.student_id || ''}" ${(s.student_id || '') === (selectedId || '') ? 'selected' : ''}>${s.student_id || ''} — ${s.name || ''}</option>`).join('');
}
function studentDatalistHTML(listId) {
  const students = getDataByType('student');
  return `<datalist id="${listId}">${students.map(s => `<option value="${s.student_id || ''}">${s.student_id || ''} — ${s.name || ''}</option>`).join('')}</datalist>`;
}

// ======================== GOOGLE SHEET INIT ========================
function saveGSheetConfig() {
  const input = document.getElementById('gsheetUrlInput');
  const scriptInput = document.getElementById('gsheetScriptUrl');
  const errEl = document.getElementById('configError');
  errEl.classList.add('hidden');
  const sheetId = GSheetDB.extractSheetId(input.value);
  if (!sheetId) {
    errEl.textContent = 'กรุณาวาง URL ของ Google Sheet หรือ Spreadsheet ID ที่ถูกต้อง';
    errEl.classList.remove('hidden');
    return;
  }
  const scriptUrl = (scriptInput ? scriptInput.value : '').trim() || '';
  GSheetDB.storeConfig({ spreadsheetId: sheetId, scriptUrl: scriptUrl });
  initGSheet(sheetId, scriptUrl);
}

function resetGSheetConfig() {
  const allUsers = getDataByType('user');
  const hasAdmin = allUsers.some(u => normalizeRole(u.role) === 'admin');

  if (hasAdmin) {
    const p = prompt("กรุณากรอกรหัสผ่านผู้ดูแลระบบ (Admin) 6 หลัก เพื่อเปลี่ยน Google Sheet:");
    if (!p) return;
    const adminUser = allUsers.find(u => normalizeRole(u.role) === 'admin' && cleanPassword(u.password) === p.trim());
    if (!adminUser) {
      alert("รหัสผ่านผู้ดูแลระบบไม่ถูกต้อง");
      return;
    }
  }

  GSheetDB.clearConfig();
  GSheetDB.destroy();
  showScreen('configScreen');
}

function showScreen(id) {
  ['loadingScreen', 'configScreen', 'loginScreen', 'mainApp'].forEach(s => {
    const el = document.getElementById(s);
    if (s === id) {
      el.classList.remove('hidden');
      el.classList.add('flex');
    } else {
      el.classList.add('hidden');
      el.classList.remove('flex');
    }
  });
}

function loadPermissions() {
  const perms = getDataByType('permission');
  perms.forEach(p => {
    if (p.role && p.module && APP.permissions[p.role]) {
      APP.permissions[p.role][p.module] = parseInt(p.value, 10) || 0;
    }
  });
}

function initGSheet(sheetId, scriptUrl) {
  showScreen('loadingScreen');
  GSheetDB.init({ spreadsheetId: sheetId, scriptUrl: scriptUrl || '' }, (data) => {
    APP.allData = data;
    loadPermissions();
    if (APP.currentUser) {
      buildSidebar();
      renderCurrentPage();
    }
    updateNotifBadge();
  }).then(r => {
    if (r.isOk) {
      showScreen('loginScreen');
    } else {
      alert('เชื่อมต่อ Google Sheet ไม่สำเร็จ: ' + (r.error || 'ตรวจสอบว่า Sheet เป็น Public และไม่ได้ลบ Tab ที่จำเป็นทิ้ง'));
      document.body.innerHTML = `<div class="h-screen flex flex-col items-center justify-center text-center p-6 bg-gray-50"><h2 class="text-2xl font-bold text-red-500 mb-2">เชื่อมต่อฐานข้อมูลไม่สำเร็จ 🚨</h2><p class="text-gray-600">${r.error || 'โปรดตรวจสอบลิงก์ Google Sheet ว่าถูกต้อง หรือสิทธิ์การเข้าถึงเป็น Anyone with the link'}</p></div>`;
    }
  });
}

async function refreshData() {
  showToast('กำลังรีเฟรชข้อมูล...');
  await GSheetDB.refresh();
  showToast('รีเฟรชข้อมูลสำเร็จ');
}

async function debugConnection() {
  const result = await GSheetDB.debugTab('user');
  const totalData = APP.allData.length;
  const userCount = getDataByType('user').length;
  let msg = `=== Debug Info ===\n`;
  msg += `ข้อมูลทั้งหมดในระบบ: ${totalData} แถว\n`;
  msg += `ข้อมูล type=user: ${userCount} แถว\n\n`;
  msg += `=== Raw "user" tab ===\n`;
  msg += JSON.stringify(result, null, 2);
  if (userCount > 0) {
    msg += `\n\n=== User data ===\n`;
    msg += JSON.stringify(getDataByType('user'), null, 2);
  }
  alert(msg);
}

// ======================== READ-ONLY NOTICE ========================
function readOnlyNotice() {
  showToast('ระบบเป็นแบบอ่านอย่างเดียว — แก้ไขข้อมูลใน Google Sheet', 'error');
}

// Default config (hardcoded fallback)
const DEFAULT_SPREADSHEET_ID = '1SjucS8W7syfiS7I9PyQQonwSYbHT5TwSHeo9B1cGQ9U';
const DEFAULT_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbzyTFKNmxCzEpqz-BNOTe3N3DREkHKhwMnsWjlU0AmM0YqaVrWrqFARHAI88MEKPulrMQ/exec';

function getActiveConfig() {
  const stored = GSheetDB.getStoredConfig();
  return {
    spreadsheetId: (stored && stored.spreadsheetId) ? stored.spreadsheetId : DEFAULT_SPREADSHEET_ID,
    scriptUrl: (stored && stored.scriptUrl) ? stored.scriptUrl : DEFAULT_SCRIPT_URL
  };
}

// Boot
(() => {
  const cfg = getActiveConfig();
  initGSheet(cfg.spreadsheetId, cfg.scriptUrl);
})();

// ======================== LOGIN ========================
// Map Thai role names to English equivalents for matching
function normalizeRole(role) {
  const r = String(role || '').trim().toLowerCase();
  if (r === 'admin' || r === 'ผู้ดูแลระบบ') return 'admin';
  if (r === 'academic' || r === 'เจ้าหน้าที่งานวิชาการ') return 'academic';
  if (r === 'teacher' || r === 'อาจารย์') return 'teacher';
  if (r === 'classteacher' || r === 'อาจารย์ประจำชั้น') return 'classTeacher';
  if (r === 'student' || r === 'นักศึกษา') return 'student';
  if (r === 'executive' || r === 'ผู้บริหาร') return 'executive';
  return r;
}

// Clean password value from Google Sheet (handles .0, scientific notation, etc.)
function cleanPassword(val) {
  let s = String(val || '').trim();
  s = s.replace(/\.0$/, '');
  if (s.includes('E') || s.includes('e')) { try { s = String(Math.round(Number(s))); } catch (_) { } }
  return s;
}
function updateLoginFields() {
  const role = document.getElementById('loginRole').value;
  const f = document.getElementById('loginFields');
  if (role === 'admin') {
    f.innerHTML = `<div class="mb-4"><label class="block text-sm font-medium text-gray-700 mb-2">รหัสผ่าน 6 หลัก</label><input type="password" id="adminPass" maxlength="6" pattern="[0-9]{6}" class="w-full border border-gray-300 rounded-xl px-4 py-3 focus:ring-2 focus:ring-primary focus:border-primary outline-none" placeholder="กรอกรหัสผ่าน 6 หลัก" onkeypress="if(event.key==='Enter')handleLogin()"></div>`;
  } else if (role === 'student') {
    f.innerHTML = `<div class="mb-4"><label class="block text-sm font-medium text-gray-700 mb-2">เลขบัตรประชาชน 13 หลัก</label><input type="text" id="studentNID" maxlength="13" class="w-full border border-gray-300 rounded-xl px-4 py-3 focus:ring-2 focus:ring-primary focus:border-primary outline-none" placeholder="กรอกเลขบัตรประชาชน 13 หลัก" onkeypress="if(event.key==='Enter')handleLogin()"></div>`;
  } else if (role === 'executive' || role === 'classTeacher') {
    f.innerHTML = `<div class="mb-4"><label class="block text-sm font-medium text-gray-700 mb-2">Username</label><input type="text" id="loginUsername" class="w-full border border-gray-300 rounded-xl px-4 py-3 focus:ring-2 focus:ring-primary focus:border-primary outline-none" placeholder="กรอก Username"></div><div class="mb-4"><label class="block text-sm font-medium text-gray-700 mb-2">รหัสผ่าน</label><input type="password" id="loginUserPass" class="w-full border border-gray-300 rounded-xl px-4 py-3 focus:ring-2 focus:ring-primary focus:border-primary outline-none" placeholder="รหัสผ่าน" onkeypress="if(event.key==='Enter')handleLogin()"></div>`;
  } else {
    f.innerHTML = `<div class="mb-4"><label class="block text-sm font-medium text-gray-700 mb-2">E-mail</label><input type="email" id="teacherEmail" class="w-full border border-gray-300 rounded-xl px-4 py-3 focus:ring-2 focus:ring-primary focus:border-primary outline-none" placeholder="E-mail"></div><div class="mb-4"><label class="block text-sm font-medium text-gray-700 mb-2">รหัสผ่าน</label><input type="password" id="teacherPass" class="w-full border border-gray-300 rounded-xl px-4 py-3 focus:ring-2 focus:ring-primary focus:border-primary outline-none" placeholder="รหัสผ่าน" onkeypress="if(event.key==='Enter')handleLogin()"></div>`;
  }
}
updateLoginFields();

function handleLogin() {
  const role = document.getElementById('loginRole').value;
  const err = document.getElementById('loginError');
  err.classList.add('hidden');
  if (role === 'admin') {
    const p = document.getElementById('adminPass').value;
    if (!/^\d{6}$/.test(p)) { err.textContent = 'กรุณากรอกรหัสผ่าน 6 หลัก (ตัวเลขเท่านั้น)'; err.classList.remove('hidden'); return }
    const adminUser = getDataByType('user').find(u => {
      if (normalizeRole(u.role) !== 'admin') return false;
      return cleanPassword(u.password) === p;
    });
    if (!adminUser) {
      const allUsers = getDataByType('user');
      if (allUsers.length === 0) {
        err.innerHTML = 'ไม่พบข้อมูลผู้ใช้ — ตรวจสอบว่า Google Sheet มี tab ชื่อ "user" และ Share เป็น Public แล้ว<br><button onclick="debugConnection()" class="mt-2 text-xs underline text-primary">ตรวจสอบการเชื่อมต่อ</button>';
      } else {
        const adminUsers = allUsers.filter(u => normalizeRole(u.role) === 'admin');
        if (adminUsers.length === 0) {
          err.textContent = 'ไม่พบผู้ใช้ที่มี role=admin ในระบบ (พบ user ' + allUsers.length + ' คน)';
        } else {
          err.innerHTML = 'รหัสผ่านไม่ถูกต้อง<br><button onclick="debugConnection()" class="mt-2 text-xs underline text-primary">ตรวจสอบการเชื่อมต่อ</button> <button onclick="GSheetDB.clearConfig();location.reload()" class="mt-2 text-xs underline text-red-500 ml-3">รีเซ็ต Google Sheet</button>';
        }
      }
      err.classList.remove('hidden'); return
    }
    APP.currentUser = { name: adminUser.name || 'ผู้ดูแลระบบ', role: 'admin' };
  } else if (role === 'student') {
    const nid = document.getElementById('studentNID').value.trim();
    if (!/^\d{13}$/.test(nid)) { err.textContent = 'กรุณากรอกเลขบัตรประชาชน 13 หลัก'; err.classList.remove('hidden'); return }
    // Match with norm(), also handle sheet stripping leading zeros (compare as numbers)
    const stu = getDataByType('student').find(s => {
      const stored = norm(s.national_id);
      return stored === nid || stored === String(Number(nid)) || nid === String(Number(stored)).padStart(13, '0');
    });
    if (!stu) { err.textContent = 'ไม่พบข้อมูลนักศึกษา กรุณาตรวจสอบเลขบัตรประชาชน'; err.classList.remove('hidden'); return }
    APP.currentUser = { name: stu.name, role: 'student', data: stu };
  } else if (role === 'teacher') {
    const em = document.getElementById('teacherEmail').value;
    const pw = document.getElementById('teacherPass').value;
    if (!em) { err.textContent = 'กรุณากรอก E-mail'; err.classList.remove('hidden'); return }
    if (!pw) { err.textContent = 'กรุณากรอกรหัสผ่าน'; err.classList.remove('hidden'); return }
    const user = getDataByType('user').find(u => normalizeRole(u.role) === 'teacher' && String(u.email || '').trim().toLowerCase() === em.trim().toLowerCase() && cleanPassword(u.password) === pw);
    if (!user) { err.textContent = 'E-mail หรือรหัสผ่านไม่ถูกต้อง'; err.classList.remove('hidden'); return }
    const t = getDataByType('teacher').find(x => x.email === em);
    APP.currentUser = t ? { name: t.name, role: 'teacher', data: t } : { name: user.name || em, role: 'teacher', email: em };
  } else if (role === 'academic') {
    const em = document.getElementById('teacherEmail').value;
    const pw = document.getElementById('teacherPass').value;
    if (!em) { err.textContent = 'กรุณากรอก E-mail'; err.classList.remove('hidden'); return }
    if (!pw) { err.textContent = 'กรุณากรอกรหัสผ่าน'; err.classList.remove('hidden'); return }
    const user = getDataByType('user').find(u => normalizeRole(u.role) === 'academic' && String(u.email || '').trim().toLowerCase() === em.trim().toLowerCase() && cleanPassword(u.password) === pw);
    if (!user) { err.textContent = 'E-mail หรือรหัสผ่านไม่ถูกต้อง'; err.classList.remove('hidden'); return }
    APP.currentUser = { name: user.name || em, role: 'academic', email: em };
  } else if (role === 'classTeacher') {
    const uname = document.getElementById('loginUsername').value.trim();
    const pw = document.getElementById('loginUserPass').value;
    if (!uname) { err.textContent = 'กรุณากรอก Username'; err.classList.remove('hidden'); return }
    if (!pw) { err.textContent = 'กรุณากรอกรหัสผ่าน'; err.classList.remove('hidden'); return }
    const user = getDataByType('user').find(u => normalizeRole(u.role) === 'classTeacher' && (String(u.username || '').trim().toLowerCase() === uname.toLowerCase() || String(u.name || '').trim().toLowerCase() === uname.toLowerCase()) && cleanPassword(u.password) === pw);
    if (!user) { err.textContent = 'Username หรือรหัสผ่านไม่ถูกต้อง'; err.classList.remove('hidden'); return }
    const t = getDataByType('teacher').find(x => (x.name || '').trim().toLowerCase() === (user.name || '').trim().toLowerCase() || (x.email || '') === (user.email || ''));
    APP.currentUser = t ? { name: t.name, role: 'classTeacher', data: t, responsible_year: t.responsible_year || user.responsible_year || '1' } : { name: user.name || uname, role: 'classTeacher', responsible_year: user.responsible_year || '1' };
  } else if (role === 'executive') {
    const uname = document.getElementById('loginUsername').value.trim();
    const pw = document.getElementById('loginUserPass').value;
    if (!uname) { err.textContent = 'กรุณากรอก Username'; err.classList.remove('hidden'); return }
    if (!pw) { err.textContent = 'กรุณากรอกรหัสผ่าน'; err.classList.remove('hidden'); return }
    const user = getDataByType('user').find(u => normalizeRole(u.role) === 'executive' && (String(u.username || '').trim().toLowerCase() === uname.toLowerCase() || String(u.name || '').trim().toLowerCase() === uname.toLowerCase()) && cleanPassword(u.password) === pw);
    if (!user) { err.textContent = 'Username หรือรหัสผ่านไม่ถูกต้อง'; err.classList.remove('hidden'); return }
    APP.currentUser = { name: user.name || uname, role: 'executive' };
  }
  APP.currentRole = APP.currentUser.role;
  // Log login event to Google Sheet
  try {
    let identifier = '';
    if (APP.currentRole === 'student' && APP.currentUser.data) identifier = APP.currentUser.data.student_id || '';
    else if (APP.currentUser.email) identifier = APP.currentUser.email;
    logLoginEvent('login', { name: APP.currentUser.name, role: APP.currentRole, identifier });
  } catch (e) { /* ignore */ }
  showScreen('mainApp');
  document.getElementById('currentUserName').textContent = APP.currentUser.name;
  document.getElementById('currentUserRole').textContent = { admin: 'ผู้ดูแลระบบ', academic: 'เจ้าหน้าที่งานวิชาการ', teacher: 'อาจารย์', classTeacher: 'อาจารย์ประจำชั้น', student: 'นักศึกษา', executive: 'ผู้บริหาร' }[APP.currentRole];
  buildSidebar();
  navigateTo('dashboard');
  lucide.createIcons();
}

function handleLogout() {
  // Log logout event before clearing state
  try {
    if (APP.currentUser) {
      let identifier = '';
      if (APP.currentRole === 'student' && APP.currentUser.data) identifier = APP.currentUser.data.student_id || '';
      else if (APP.currentUser.email) identifier = APP.currentUser.email;
      logLoginEvent('logout', { name: APP.currentUser.name, role: APP.currentRole, identifier });
    }
  } catch (e) { /* ignore */ }
  APP.currentUser = null; APP.currentRole = null; APP.currentPage = 'dashboard';
  showScreen('loginScreen');
}

// ======================== SIDEBAR ========================
function buildSidebar() {
  const r = APP.currentRole;
  const p = APP.permissions[r] || {};
  let items = [];
  if (p.dashboard) items.push({ id: 'dashboard', icon: 'layout-dashboard', label: 'หน้าหลัก' });

  // Registration dropdown — permission-driven for all roles
  let regSub = [];
  if (r === 'student') {
    if (p.students) regSub.push({ id: 'studentInfo', label: 'ข้อมูลนักศึกษา' });
  } else {
    if (p.students) regSub.push({ id: 'students', label: 'ข้อมูลนักศึกษา' });
  }
  if (p.subjects) regSub.push({ id: 'subjects', label: 'รายวิชาที่เปิดสอน' });
  if (p.schedule) regSub.push({ id: 'schedule', label: 'ตารางเรียน/ตารางสอบ' });
  if (p.grades) regSub.push({ id: 'grades', label: 'ผลการเรียน' });
  if (p.engResults) regSub.push({ id: 'engResults', label: 'ผลสอบภาษาอังกฤษ' });
  if (p.teachers) regSub.push({ id: 'teachers', label: 'ข้อมูลอาจารย์' });
  if (regSub.length) items.push({ id: 'registration', icon: 'book-open', label: 'ระบบทะเบียน', sub: regSub });

  if (p.teacherDirectory) items.push({ id: 'teacherDirectory', icon: 'award', label: 'ทำเนียบอาจารย์' });

  // Tracking dropdown
  let trackSub = [];
  if (p.tracking) trackSub.push({ id: 'tracking', label: 'ส่งรายละเอียดรายวิชา' });
  if (p.tracking) trackSub.push({ id: 'resultTracking', label: 'ส่งผลการดำเนินงานรายวิชา' });
  if (p.gradeTracking) trackSub.push({ id: 'gradeTracking', label: 'ส่งเกรดรายวิชา' });
  if (p.fileTracking) trackSub.push({ id: 'fileTracking', label: 'ส่งแฟ้มรายวิชา' });
  if (trackSub.length) items.push({ id: 'trackingGroup', icon: 'clipboard-list', label: 'ติดตามการส่ง', sub: trackSub });

  if (p.leave) items.push({ id: 'leave', icon: 'calendar-off', label: 'ระบบการลาของนักศึกษา' });
  if (p.services) items.push({ id: 'services', icon: 'grid', label: 'บริการอื่นๆ' });
  if ((r === 'admin' || r === 'academic') && p.settings) items.push({ id: 'settings', icon: 'settings', label: 'ตั้งค่าระบบ' });
  if (r === 'admin' && p.loginLog) items.push({ id: 'loginLog', icon: 'log-in', label: 'บันทึกการเข้าใช้ระบบ' });

  const nav = document.getElementById('sidebarNav');
  nav.innerHTML = items.map(it => {
    if (it.sub) {
      return `<div class="dropdown-item">
        <button onclick="toggleDropdown(this)" class="w-full flex items-center justify-between px-3 py-2.5 rounded-xl text-sm text-gray-700 hover:bg-surface transition">
          <span class="flex items-center gap-3"><i data-lucide="${it.icon}" class="w-5 h-5 flex-shrink-0"></i>${it.label}</span>
          <i data-lucide="chevron-down" class="w-4 h-4 transition-transform"></i>
        </button>
        <div class="dropdown-menu ml-8 mt-1 space-y-1">
          ${it.sub.map(s => `<button onclick="navigateTo('${s.id}')" class="nav-item w-full text-left px-3 py-2 rounded-lg text-sm text-gray-600 hover:bg-surface hover:text-primary transition" data-page="${s.id}">${s.label}</button>`).join('')}
        </div>
      </div>`;
    }
    return `<button onclick="navigateTo('${it.id}')" class="nav-item w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-sm text-gray-700 hover:bg-surface hover:text-primary transition" data-page="${it.id}"><i data-lucide="${it.icon}" class="w-5 h-5 flex-shrink-0"></i>${it.label}</button>`;
  }).join('');
  lucide.createIcons();
}

function toggleDropdown(btn) {
  const p = btn.parentElement;
  p.classList.toggle('dropdown-open');
  const chev = btn.querySelector('[data-lucide="chevron-down"]');
  if (chev) chev.style.transform = p.classList.contains('dropdown-open') ? 'rotate(180deg)' : '';
}

function toggleSidebar() {
  const sb = document.getElementById('sidebar');
  const ov = document.getElementById('sidebarOverlay');
  APP.sidebarOpen = !APP.sidebarOpen;
  if (APP.sidebarOpen) { sb.classList.remove('-translate-x-full'); ov.classList.remove('hidden') }
  else { sb.classList.add('-translate-x-full'); ov.classList.add('hidden') }
}

// ======================== NAVIGATION ========================
function navigateTo(page) {
  APP.currentPage = page;
  APP.pagination.page = 1;
  APP.filters.search = '';
  APP.filters.semester = '';
  APP.filters.academicYear = '';
  APP.filters.yearLevel = '';
  APP.filters._gradeStudent = '';
  APP.filters._engStudent = '';
  APP.filters._gradeSearch = '';
  APP.filters._engSearch = '';
  APP.filters._trackingYear = '';
  APP.filters._gradeTrackingYear = '';
  APP.filters._fileTrackingYear = '';
  APP.filters._studentYearLevel = '';
  APP.filters._engAdvisor = '';
  APP.filters._gradeAdvisor = '';
  APP.filters._engYear = '';
  APP.filters._engYearLevel = '';
  APP.filters._gradeYearLevel = '';
  APP._directoryTab = 'all';
  APP.filters._directoryYear = '';
  APP.filters._pageYear = '';
  document.querySelectorAll('.nav-item').forEach(n => {
    n.classList.toggle('bg-primaryLight', n.dataset.page === page);
    n.classList.toggle('text-primary', n.dataset.page === page);
    n.classList.toggle('font-semibold', n.dataset.page === page);
  });
  renderCurrentPage();
  if (APP.sidebarOpen) toggleSidebar();
}

function renderCurrentPage() {
  const mc = document.getElementById('mainContent');
  const p = APP.currentPage;
  const r = APP.currentRole;
  mc.innerHTML = '<div class="fade-in">' + getPageContent(p, r) + '</div>';
  lucide.createIcons();
  initPageScripts(p);
}

// ======================== HELPERS ========================
// Normalize value from Google Sheet (strip .0, trim whitespace)
function norm(v) { return String(v || '').replace(/\.0$/, '').trim() }

// Convert YYYY-MM-DD (Gregorian) to DD/MM/YYYY (Buddhist Era)
function toBuddhistDate(dateStr) {
  if (!dateStr) return '';
  const s = String(dateStr).trim();
  // Match YYYY-MM-DD format
  const m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) {
    const y = parseInt(m[1], 10);
    const mo = String(parseInt(m[2], 10)).padStart(2, '0');
    const d = String(parseInt(m[3], 10)).padStart(2, '0');
    return `${d}/${mo}/${y + 543}`;
  }
  // Google Sheets gviz date format: Date(yyyy,m,d) — month is 0-indexed
  const gs = s.match(/^Date\((\d+),(\d+),(\d+)(?:,\d+,\d+,\d+)?\)$/);
  if (gs) {
    const y = parseInt(gs[1], 10);
    const mo = String(parseInt(gs[2], 10) + 1).padStart(2, '0');
    const d = String(parseInt(gs[3], 10)).padStart(2, '0');
    return `${d}/${mo}/${y + 543}`;
  }
  // ISO datetime e.g. 2026-05-07T00:00:00.000Z
  const iso = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})T/);
  if (iso) {
    const y = parseInt(iso[1], 10);
    const mo = String(parseInt(iso[2], 10)).padStart(2, '0');
    const d = String(parseInt(iso[3], 10)).padStart(2, '0');
    return `${d}/${mo}/${y + 543}`;
  }
  // DD/MM/YYYY where YYYY is CE — convert to BE
  const dmyCE = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (dmyCE) {
    const y = parseInt(dmyCE[3], 10);
    if (y < 2500) {
      const dd = String(parseInt(dmyCE[1], 10)).padStart(2, '0');
      const mo = String(parseInt(dmyCE[2], 10)).padStart(2, '0');
      return `${dd}/${mo}/${y + 543}`;
    }
  }
  // If already in DD/MM/YYYY format with BE year, leave as-is
  return s;
}

// Format multiple dates (comma-separated) to Buddhist
function toBuddhistDateList(dateStr) {
  if (!dateStr) return '';
  const s = String(dateStr).trim();
  // Single Google Sheets date literal — don't split (commas are inside the parens)
  if (/^Date\(\d+,\d+,\d+(?:,\d+,\d+,\d+)?\)$/.test(s)) {
    return toBuddhistDate(s);
  }
  // Otherwise split multi-day list on commas/semicolons
  return s.split(/[,;]/).map(d => toBuddhistDate(d.trim())).filter(Boolean).join(', ');
}
function parseDate(v) {
  if (!v) return null;
  const s = String(v).trim();
  if (!s) return null;
  // Google Sheets JSON date: Date(yyyy,m,d) — month is 0-indexed
  const gsMatch = s.match(/^Date\((\d+),(\d+),(\d+)\)$/);
  if (gsMatch) return new Date(Number(gsMatch[1]), Number(gsMatch[2]), Number(gsMatch[3]));
  // Excel serial number (e.g. 44927)
  if (/^\d{5}$/.test(s)) {
    const d = new Date((Number(s) - 25569) * 86400 * 1000);
    return isNaN(d) ? null : d;
  }
  // dd/mm/yyyy or dd-mm-yyyy (Buddhist or CE)
  const dmyMatch = s.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})$/);
  if (dmyMatch) {
    let [, d, m, y] = dmyMatch;
    let yr = Number(y);
    if (yr > 2400) yr = yr - 543; // พ.ศ. → ค.ศ.
    return new Date(yr, Number(m) - 1, Number(d));
  }
  // yyyy-mm-dd
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return new Date(s);
  const d = new Date(s);
  return isNaN(d) ? null : d;
}
function formatDate(v) {
  const d = parseDate(v);
  if (!d || isNaN(d)) return v || '';
  const dd = String(d.getDate()).padStart(2, '0');
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const yyyy = d.getFullYear() + 543; // CE → พ.ศ.
  return `${dd}/${mm}/${yyyy}`;
}
// Normalize date input (พ.ศ. or ค.ศ.) to dd/mm/yyyy CE for storage
function normalizeDateInput(v) {
  if (!v) return '';
  const d = parseDate(v);
  if (!d || isNaN(d)) return v;
  const dd = String(d.getDate()).padStart(2, '0');
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  return `${dd}/${mm}/${d.getFullYear()}`; // store as CE
}
function semLabel(v) {
  const s = norm(v);
  if (s === '3' || s.includes('ร้อน')) return 'ฤดูร้อน';
  if (s === '1') return '1';
  if (s === '2') return '2';
  return s;
}
function normSem(v) {
  const s = norm(v);
  if (s.includes('ร้อน') || s === '3') return '3';
  return s;
}

// Loading overlay for save/edit/delete operations
function showLoading(msg = 'กำลังบันทึก...') {
  let overlay = document.getElementById('loadingOverlay');
  if (!overlay) {
    overlay = document.createElement('div');
    overlay.id = 'loadingOverlay';
    overlay.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,0.35);display:flex;align-items:center;justify-content:center;z-index:9999;';
    document.body.appendChild(overlay);
  }
  overlay.innerHTML = `<div class="bg-white rounded-2xl px-8 py-6 shadow-2xl flex flex-col items-center gap-3 fade-in">
    <div class="w-10 h-10 border-4 border-primary border-t-transparent rounded-full animate-spin"></div>
    <p class="text-sm text-gray-700 font-medium" id="loadingMsg">${msg}</p>
  </div>`;
  overlay.classList.remove('hidden');
}

function hideLoading() {
  const overlay = document.getElementById('loadingOverlay');
  if (overlay) overlay.classList.add('hidden');
}

function showToast(msg, type = 'success') {
  const c = document.getElementById('toastContainer');
  const colors = type === 'success' ? 'bg-green-500' : type === 'loading' ? 'bg-blue-500' : 'bg-red-500';
  const icon = type === 'success' ? 'check-circle' : type === 'loading' ? 'loader' : 'alert-circle';
  const d = document.createElement('div');
  d.className = `toast ${colors} text-white px-5 py-3.5 rounded-xl shadow-lg text-sm flex items-center gap-2`;
  d.innerHTML = `<i data-lucide="${icon}" class="w-5 h-5 ${type === 'loading' ? 'animate-spin' : ''}"></i>${msg}`;
  if (type === 'loading') d.id = 'loadingToast';
  c.appendChild(d);
  lucide.createIcons();
  if (type !== 'loading') setTimeout(() => { d.style.transition = 'opacity .3s, transform .3s'; d.style.opacity = '0'; d.style.transform = 'translateX(30px)'; setTimeout(() => d.remove(), 300) }, 2500);
  return d;
}

function hideLoadingToast() {
  const t = document.getElementById('loadingToast');
  if (t) { t.style.transition = 'opacity .2s'; t.style.opacity = '0'; setTimeout(() => t.remove(), 200) }
}

// Wrap save action with loading state on button
async function withLoading(btnOrForm, asyncFn) {
  const btn = btnOrForm ? (btnOrForm.tagName === 'FORM' ? btnOrForm.querySelector('[type="submit"]') : btnOrForm) : null;
  const origText = btn ? btn.innerHTML : '';
  if (btn) { btn.disabled = true; btn.innerHTML = '<i data-lucide="loader" class="w-4 h-4 animate-spin inline mr-1"></i>กำลังบันทึก...'; lucide.createIcons(); }
  try {
    await asyncFn();
  } finally {
    if (btn) { btn.disabled = false; btn.innerHTML = origText; lucide.createIcons(); }
  }
}

function showModal(title, content, onConfirm, maxWidth) {
  const mc = document.getElementById('modalContainer');
  mc.classList.remove('hidden');
  window._modalConfirm = onConfirm || null;
  const mw = maxWidth || 'max-w-lg';
  mc.innerHTML = `<div class="modal-overlay fixed inset-0 flex items-center justify-center p-4 z-50" onclick="if(event.target===this)closeModal()">
    <div class="modal-content bg-white rounded-2xl shadow-2xl w-full ${mw} max-h-[85vh] overflow-auto">
      <div class="p-5 border-b flex items-center justify-between"><h3 class="font-bold text-lg">${title}</h3><button onclick="closeModal()" class="p-1 rounded hover:bg-gray-100"><i data-lucide="x" class="w-5 h-5"></i></button></div>
      <div class="p-5">${content}</div>
      ${onConfirm ? `<div class="p-4 border-t flex justify-end gap-2"><button onclick="closeModal()" class="px-4 py-2 rounded-xl border hover:bg-gray-50">ยกเลิก</button><button onclick="if(window._modalConfirm)window._modalConfirm()" class="px-4 py-2 rounded-xl bg-primary text-white hover:bg-primaryDark">ยืนยัน</button></div>` : ''}
    </div>
  </div>`;
  // Animate in
  requestAnimationFrame(() => {
    const overlay = mc.querySelector('.modal-overlay');
    const box = mc.querySelector('.modal-content');
    if (overlay) { overlay.style.opacity = '0'; overlay.style.transition = 'opacity .25s ease'; requestAnimationFrame(() => overlay.style.opacity = '1') }
    if (box) { box.style.transform = 'scale(0.95) translateY(10px)'; box.style.opacity = '0'; box.style.transition = 'transform .25s ease, opacity .25s ease'; requestAnimationFrame(() => { box.style.transform = 'scale(1) translateY(0)'; box.style.opacity = '1' }) }
  });
  lucide.createIcons();
}

function closeModal() {
  const mc = document.getElementById('modalContainer');
  const overlay = mc.querySelector('.modal-overlay');
  const box = mc.querySelector('.modal-content');
  if (box) { box.style.transition = 'transform .2s ease, opacity .2s ease'; box.style.transform = 'scale(0.95) translateY(10px)'; box.style.opacity = '0' }
  if (overlay) { overlay.style.transition = 'opacity .2s ease'; overlay.style.opacity = '0' }
  setTimeout(() => { mc.classList.add('hidden'); mc.innerHTML = '' }, 220);
}

function showNotifications() { document.getElementById('notifPanel').style.transform = 'translateX(0)'; renderNotifications() }
function closeNotifications() { document.getElementById('notifPanel').style.transform = 'translateX(100%)' }
function renderNotifications() {
  const ann = getDataByType('announcement').slice(-20).reverse();
  document.getElementById('notifList').innerHTML = ann.length ? ann.map(a => `<div class="p-3 bg-surface rounded-xl"><p class="font-medium text-sm">${a.announcement_title || ''}</p><p class="text-xs text-gray-500 mt-1">${a.announcement_date || ''}</p><p class="text-xs text-gray-600 mt-1">${a.announcement_content || ''}</p></div>`).join('') : '<p class="text-gray-400 text-center text-sm">ไม่มีการแจ้งเตือน</p>';
}
function updateNotifBadge() {
  const b = document.getElementById('notifBadge');
  const c = getDataByType('announcement').length;
  if (c > 0) { b.textContent = c > 99 ? '99+' : c; b.classList.remove('hidden') } else b.classList.add('hidden');
}

function paginationHTML(total, perPage, page, onChange) {
  const pages = Math.ceil(total / perPage) || 1;
  if (pages <= 1) return '';
  let h = '<div class="flex items-center justify-center gap-1 mt-4">';
  h += `<button onclick="${onChange}(1)" class="px-3 py-1 rounded-lg border text-sm hover:bg-gray-50 ${page === 1 ? 'opacity-40' : ''}" title="หน้าแรก">«</button>`;
  h += `<button onclick="${onChange}(${Math.max(1, page - 1)})" class="px-3 py-1 rounded-lg border text-sm hover:bg-gray-50 ${page === 1 ? 'opacity-40' : ''}" title="ก่อนหน้า">‹</button>`;
  for (let i = 1; i <= Math.min(pages, 7); i++) {
    const pg = pages <= 7 ? i : i <= 3 ? i : i <= 5 ? page - 3 + i : pages - 7 + i;
    h += `<button onclick="${onChange}(${Math.min(pages, Math.max(1, pg))})" class="px-3 py-1 rounded-lg text-sm ${pg === page ? 'bg-primary text-white' : 'border hover:bg-gray-50'}">${Math.min(pages, Math.max(1, pg))}</button>`;
  }
  h += `<button onclick="${onChange}(${Math.min(pages, page + 1)})" class="px-3 py-1 rounded-lg border text-sm hover:bg-gray-50 ${page === pages ? 'opacity-40' : ''}" title="ถัดไป">›</button>`;
  h += `<button onclick="${onChange}(${pages})" class="px-3 py-1 rounded-lg border text-sm hover:bg-gray-50 ${page === pages ? 'opacity-40' : ''}" title="หน้าสุดท้าย">»</button>`;
  h += '</div>';
  return h;
}

function filterBar(opts = {}) {
  const searchVal = APP.filters.search || '';
  let h = '<div class="flex flex-wrap gap-3 mb-4">';
  h += `<div class="flex-1 min-w-[200px] relative"><i data-lucide="search" class="absolute left-3 top-3 w-4 h-4 text-gray-400"></i><input type="text" placeholder="ค้นหา..." value="${searchVal}" class="w-full pl-10 pr-4 py-2.5 border border-gray-200 rounded-xl text-sm focus:ring-2 focus:ring-primary focus:border-primary outline-none" oninput="APP.filters.search=this.value;APP.pagination.page=1;clearTimeout(window._searchTimer);window._searchTimer=setTimeout(()=>renderCurrentPage(),300)"></div>`;
  if (opts.semester !== false) {
    const sem = APP.filters.semester || '';
    h += `<select onchange="APP.filters.semester=this.value;APP.pagination.page=1;renderCurrentPage()" class="border border-gray-200 rounded-xl px-3 py-2.5 text-sm"><option value="">ทุกภาคการศึกษา</option><option value="1" ${sem === '1' ? 'selected' : ''}>ภาคการศึกษาที่ 1</option><option value="2" ${sem === '2' ? 'selected' : ''}>ภาคการศึกษาที่ 2</option><option value="3" ${sem === '3' ? 'selected' : ''}>ภาคฤดูร้อน</option></select>`;
  }
  if (opts.year !== false) {
    const yr = APP.filters.academicYear || '';
    h += `<select onchange="APP.filters.academicYear=this.value;APP.pagination.page=1;renderCurrentPage()" class="border border-gray-200 rounded-xl px-3 py-2.5 text-sm"><option value="">ทุกปีการศึกษา</option><option value="2567" ${yr === '2567' ? 'selected' : ''}>2567</option><option value="2568" ${yr === '2568' ? 'selected' : ''}>2568</option><option value="2569" ${yr === '2569' ? 'selected' : ''}>2569</option></select>`;
  }
  if (opts.yearLevel) {
    const yl = APP.filters.yearLevel || '';
    h += `<select onchange="APP.filters.yearLevel=this.value;APP.pagination.page=1;renderCurrentPage()" class="border border-gray-200 rounded-xl px-3 py-2.5 text-sm"><option value="">ทุกชั้นปี</option><option value="1" ${yl === '1' ? 'selected' : ''}>ชั้นปี 1</option><option value="2" ${yl === '2' ? 'selected' : ''}>ชั้นปี 2</option><option value="3" ${yl === '3' ? 'selected' : ''}>ชั้นปี 3</option><option value="4" ${yl === '4' ? 'selected' : ''}>ชั้นปี 4</option></select>`;
  }
  h += '</div>';
  return h;
}

// Reusable academic year picker bar
function yearPickerBar(allData, label) {
  const allYears = [...new Set(allData.map(d => d.academic_year).filter(Boolean))].sort().reverse();
  if (!allYears.length) allYears.push('2568');
  const sel = APP.filters._pageYear || '';
  return `<div class="bg-white rounded-2xl p-4 border border-blue-100 mb-4">
    <div class="flex items-center gap-3">
      <label class="text-sm font-medium text-gray-700">${label || 'ปีการศึกษา'}:</label>
      <select onchange="APP.filters._pageYear=this.value;APP.pagination.page=1;renderCurrentPage()" class="border border-gray-200 rounded-xl px-3 py-2 text-sm">
        <option value="">-- เลือกปีการศึกษา --</option>
        ${allYears.map(y => '<option value="' + y + '"' + (sel === y ? ' selected' : '') + '>' + y + '</option>').join('')}
      </select>
      ${sel ? '<span class="text-xs text-gray-500">แสดงข้อมูลปีการศึกษา ' + sel + '</span>' : ''}
    </div>
  </div>`;
}

function noYearSelectedMsg(pageName) {
  return '<div class="bg-blue-50 rounded-2xl p-8 text-center text-blue-600 mt-4"><i data-lucide="info" class="w-6 h-6 inline mr-2"></i>กรุณาเลือกปีการศึกษาเพื่อดูข้อมูล' + pageName + '</div>';
}

function applyFilters(data) {
  let d = data;

  if (APP.filters.search) { const s = APP.filters.search.toLowerCase(); d = d.filter(x => Object.values(x).some(v => String(v).toLowerCase().includes(s))) }
  if (APP.filters.semester) d = d.filter(x => normSem(x.semester) === APP.filters.semester);
  if (APP.filters.academicYear) d = d.filter(x => norm(x.academic_year) === APP.filters.academicYear);
  if (APP.filters.yearLevel) d = d.filter(x => norm(x.year_level) === APP.filters.yearLevel);
  return d;
}

function paginate(data) {
  const s = (APP.pagination.page - 1) * APP.pagination.perPage;
  return data.slice(s, s + APP.pagination.perPage);
}

function csvUploadBtn(type, fields) {
  return `<button onclick="triggerCSVUpload('${type}','${fields}')" class="flex items-center gap-2 px-4 py-2 border border-primary text-primary rounded-xl hover:bg-primaryLight text-sm"><i data-lucide="upload" class="w-4 h-4"></i>Upload CSV</button>
  <input type="file" id="csvInput_${type}" accept=".csv" class="hidden" onchange="handleCSVUpload(event,'${type}','${fields}')">`;
}

function triggerCSVUpload(type) { document.getElementById('csvInput_' + type).click() }

async function handleCSVUpload(e, type, fieldsStr) {
  const file = e.target.files[0]; if (!file) return;
  const text = await file.text();
  const lines = text.split('\n').filter(l => l.trim());
  if (lines.length < 2) { showToast('ไฟล์ CSV ไม่มีข้อมูล', 'error'); return }
  const headers = lines[0].split(',').map(h => h.trim().replace(/"/g, ''));
  let count = 0;
  for (let i = 1; i < lines.length; i++) {
    if (count + getDataByType(type).length >= 999) { showToast('ข้อมูลเต็ม (สูงสุด 999 รายการ)', 'error'); break }
    const vals = lines[i].split(',').map(v => v.trim().replace(/"/g, ''));
    const obj = { type, created_at: new Date().toISOString() };
    headers.forEach((h, idx) => { obj[h] = vals[idx] || '' });
    const r = await GSheetDB.create(obj);

    if (r.isOk) count++;
  }
  showToast(`นำเข้าข้อมูลสำเร็จ ${count} รายการ`);
  e.target.value = '';
}

// ======================== PAGE CONTENT ========================
function getPageContent(page, role) {
  switch (page) {
    case 'dashboard': return dashboardPage();
    case 'students': return studentsPage();
    case 'studentInfo': return studentInfoPage();
    case 'subjects': return subjectsPage();
    case 'schedule': return schedulePage();
    case 'grades': return gradesPage();
    case 'engResults': return engResultsPage();
    case 'teachers': return teachersPage();
    case 'teacherDirectory': return teacherDirectoryPage();
    case 'services': return servicesPage();
    case 'tracking': return trackingPage();
    case 'resultTracking': return resultTrackingPage();
    case 'gradeTracking': return gradeTrackingPage();
    case 'fileTracking': return fileTrackingPage();
    case 'leave': return leavePage();
    case 'settings': return settingsPage();
    case 'loginLog': return loginLogPage();
    default: return '<p>ไม่พบหน้าที่ต้องการ</p>';
  }
}

// ======================== DASHBOARD ========================
function dashboardPage() {
  const students = getDataByType('student');
  const teachers = getDataByType('teacher');
  const allEngResults = getDataByType('eng_result');
  const engPassRecords = allEngResults.filter(e => e.eng_status === 'ผ่าน');
  // Count unique students who passed (each student counted only once)
  const engPassStudentIds = [...new Set(engPassRecords.map(e => e.student_id))];
  const announcements = getDataByType('announcement').slice(-5).reverse();
  const r = APP.currentRole;

  let stats = '';
  if (r === 'admin' || r === 'academic') {
    // Teacher breakdown by department (count only active)
    const activeTeachers = teachers.filter(t => (t.teacher_status || 'ปฏิบัติงานอยู่') === 'ปฏิบัติงานอยู่');
    const deptMap = {};
    activeTeachers.forEach(t => {
      const dept = t.department || 'ไม่ระบุสาขา';
      deptMap[dept] = (deptMap[dept] || 0) + 1;
    });
    const deptCards = Object.entries(deptMap).map(([dept, count]) =>
      `<div class="bg-white rounded-xl p-3 border border-blue-100 flex items-center gap-3">
        <div class="w-10 h-10 bg-emerald-100 rounded-lg flex items-center justify-center"><i data-lucide="briefcase" class="w-5 h-5 text-emerald-600"></i></div>
        <div><p class="text-xs text-gray-500">${dept}</p><p class="text-lg font-bold text-gray-800">${count} <span class="text-xs font-normal text-gray-500">คน</span></p></div>
      </div>`
    ).join('');

    stats = `
    <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4 mb-6">
      ${statCard('users', 'จำนวนนักศึกษาทั้งหมด', students.length, 'คน', 'bg-blue-500')}
      ${statCard('briefcase', 'จำนวนอาจารย์ (ปฏิบัติงาน)', activeTeachers.length, 'คน', 'bg-emerald-500')}
      ${statCard('check-circle', 'นักศึกษาสอบผ่านภาษาอังกฤษ', engPassStudentIds.length, 'คน', 'bg-amber-500')}
    </div>
    <h3 class="font-bold mb-3 text-gray-800">จำนวนอาจารย์แยกสาขาวิชา</h3>
    <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3 mb-6">${deptCards || '<p class="text-gray-400 text-sm col-span-3">ไม่มีข้อมูลสาขาวิชา</p>'}</div>
    <h3 class="font-bold mb-3 text-gray-800">นักศึกษารายชั้นปี</h3>
    <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 mb-6">
      ${(() => {
        const _hr = getHomeroomNumbers(); const _canEdit = APP.currentUser && (APP.currentUser.role === 'admin' || APP.currentUser.role === 'academic'); return [1, 2, 3, 4].map(yr => {
          const yrStudents = students.filter(s => norm(s.year_level) === String(yr));
          const yrEngPassUnique = [...new Set(engPassRecords.filter(e => yrStudents.some(s => s.student_id === e.student_id)).map(e => e.student_id))];
          const homeroom = _hr[String(yr)] || '';
          return `<div class="bg-white rounded-2xl p-4 border border-blue-100">
          <p class="text-sm text-gray-500">ชั้นปี ${yr}</p>
          <div class="flex gap-3 mt-2">
            <div><p class="text-2xl font-bold text-primary">${yrStudents.length}</p><p class="text-xs text-gray-500">นักศึกษา</p></div>
            <div><p class="text-2xl font-bold text-green-500">${yrEngPassUnique.length}</p><p class="text-xs text-gray-500">ผ่าน ENG</p></div>
          </div>
          <div class="mt-3 pt-3 border-t border-gray-100 flex items-center justify-between gap-2">
            <div class="flex items-center gap-2 min-w-0">
              <i data-lucide="door-open" class="w-4 h-4 text-blue-500 flex-shrink-0"></i>
              <div class="min-w-0">
                <p class="text-xs text-gray-500 leading-tight">ห้องเรียนประจำ</p>
                <p class="text-sm font-bold text-gray-800 truncate">${homeroom ? homeroom : '<span class="text-gray-400 font-normal">ยังไม่กำหนด</span>'}</p>
              </div>
            </div>
            ${_canEdit ? `<button onclick="promptEditHomeroom('${yr}')" title="แก้ไขหมายเลขห้องเรียนประจำ" class="text-xs text-blue-600 hover:text-blue-800 hover:bg-blue-50 rounded-lg p-1.5 flex-shrink-0"><i data-lucide="pencil" class="w-3.5 h-3.5"></i></button>` : ''}
          </div>
        </div>`;
        }).join('');
      })()}
    </div>`;
  } else if (r === 'teacher') {
    const myStudents = students.filter(s => s.advisor === APP.currentUser.name);
    const myEngPassUnique = [...new Set(allEngResults.filter(e => e.eng_status === 'ผ่าน' && myStudents.some(s => s.student_id === e.student_id)).map(e => e.student_id))];
    const teacherYrBreakdown = ['1', '2', '3', '4'].map(yr => {
      const yrStu = myStudents.filter(s => norm(s.year_level) === yr);
      const yrPass = [...new Set(allEngResults.filter(e => e.eng_status === 'ผ่าน' && yrStu.some(s => s.student_id === e.student_id)).map(e => e.student_id))];
      return { yr, count: yrStu.length, pass: yrPass.length };
    }).filter(x => x.count > 0);
    const teacherTooltip = teacherYrBreakdown.map(x => `<div class="flex justify-between gap-4"><span>ชั้นปี ${x.yr}:</span><span>${x.count} คน (ผ่าน ENG ${x.pass})</span></div>`).join('');
    stats = `
    <div class="bg-white rounded-2xl p-5 border border-blue-100 mb-4"><p class="text-sm text-gray-500">ข้อมูลตนเอง</p><p class="font-bold text-lg">${APP.currentUser.name}</p></div>
    <div class="grid grid-cols-1 sm:grid-cols-2 gap-4 mb-6">
      <div class="group relative card-stat bg-white rounded-2xl p-5 border border-blue-100 cursor-pointer">
        <div class="flex items-center gap-4"><div class="w-12 h-12 bg-blue-500 rounded-xl flex items-center justify-center"><i data-lucide="users" class="w-6 h-6 text-white"></i></div><div><p class="text-sm text-gray-500">นักศึกษาในที่ปรึกษา</p><p class="text-2xl font-bold text-gray-800">${myStudents.length} <span class="text-sm font-normal text-gray-500">คน</span></p></div></div>
        ${teacherYrBreakdown.length ? `<div class="hidden group-hover:block absolute left-0 top-full mt-1 z-20 bg-gray-800 text-white text-xs rounded-xl p-3 shadow-lg w-64">${teacherTooltip}</div>` : ''}
      </div>
      ${statCard('check-circle', 'นักศึกษาสอบผ่านภาษาอังกฤษ', myEngPassUnique.length, 'คน', 'bg-amber-500')}
    </div>`;
  } else if (r === 'classTeacher') {
    const yr = APP.currentUser.responsible_year || '1';
    const myStudents = students.filter(s => norm(s.year_level) === norm(yr));
    const roomA = myStudents.filter(s => norm(s.room).toUpperCase() === 'A');
    const roomB = myStudents.filter(s => norm(s.room).toUpperCase() === 'B');
    const engPassA = [...new Set(allEngResults.filter(e => e.eng_status === 'ผ่าน' && roomA.some(s => s.student_id === e.student_id)).map(e => e.student_id))];
    const engFailA = roomA.filter(s => !engPassA.includes(s.student_id));
    const engPassB = [...new Set(allEngResults.filter(e => e.eng_status === 'ผ่าน' && roomB.some(s => s.student_id === e.student_id)).map(e => e.student_id))];
    const engFailB = roomB.filter(s => !engPassB.includes(s.student_id));
    stats = `
    <div class="bg-white rounded-2xl p-5 border border-blue-100 mb-4"><p class="text-sm text-gray-500">อาจารย์ประจำชั้นปีที่ ${yr}</p><p class="font-bold text-lg">${APP.currentUser.name}</p></div>
    <div class="grid grid-cols-1 sm:grid-cols-2 gap-4 mb-6">
      <div class="card-stat bg-white rounded-2xl p-5 border border-blue-100">
        <div class="flex items-center gap-4 mb-3"><div class="w-12 h-12 bg-blue-500 rounded-xl flex items-center justify-center"><i data-lucide="users" class="w-6 h-6 text-white"></i></div><div><p class="text-sm text-gray-500">จำนวนนักศึกษา</p><p class="text-2xl font-bold text-gray-800">${myStudents.length} <span class="text-sm font-normal text-gray-500">คน</span></p></div></div>
        <div class="flex gap-3 text-sm border-t pt-3">
          <span class="px-3 py-1 bg-blue-50 text-blue-700 rounded-lg">ห้อง A: <b>${roomA.length}</b> คน</span>
          <span class="px-3 py-1 bg-purple-50 text-purple-700 rounded-lg">ห้อง B: <b>${roomB.length}</b> คน</span>
        </div>
      </div>
      <div class="card-stat bg-white rounded-2xl p-5 border border-blue-100">
        <div class="flex items-center gap-4 mb-3"><div class="w-12 h-12 bg-amber-500 rounded-xl flex items-center justify-center"><i data-lucide="check-circle" class="w-6 h-6 text-white"></i></div><div><p class="text-sm text-gray-500">ผลสอบภาษาอังกฤษ</p><p class="text-2xl font-bold text-gray-800">${engPassA.length + engPassB.length} <span class="text-sm font-normal text-gray-500">คนผ่าน</span></p></div></div>
        <div class="border-t pt-3 space-y-2 text-sm">
          <div class="flex items-center gap-2"><span class="font-medium text-gray-600 w-14">ห้อง A:</span><span class="px-2 py-0.5 bg-green-50 text-green-700 rounded-lg">ผ่าน <b>${engPassA.length}</b></span><span class="px-2 py-0.5 bg-red-50 text-red-700 rounded-lg">ไม่ผ่าน <b>${engFailA.length}</b></span></div>
          <div class="flex items-center gap-2"><span class="font-medium text-gray-600 w-14">ห้อง B:</span><span class="px-2 py-0.5 bg-green-50 text-green-700 rounded-lg">ผ่าน <b>${engPassB.length}</b></span><span class="px-2 py-0.5 bg-red-50 text-red-700 rounded-lg">ไม่ผ่าน <b>${engFailB.length}</b></span></div>
        </div>
      </div>
    </div>`;
  } else if (r === 'student' && APP.currentUser.data) {
    const myLeaves = getDataByType('leave').filter(l => l.name === APP.currentUser.data.name);
    const subjectMap = {};
    myLeaves.forEach(l => {
      const subj = l.subject_name || 'ไม่ระบุ';
      if (!subjectMap[subj]) subjectMap[subj] = { hours: 0, percent: 0, count: 0 };
      subjectMap[subj].hours += Number(l.leave_hours) || 0;
      subjectMap[subj].count++;
      const pct = Number(l.leave_percent) || 0;
      if (pct > subjectMap[subj].percent) subjectMap[subj].percent = pct;
    });
    const leaveRows = Object.entries(subjectMap).map(([subj, info]) => {
      const pct = info.percent;
      const colorClass = pct >= 20 ? 'bg-red-100 text-red-700 font-bold' : pct >= 15 ? 'bg-yellow-100 text-yellow-700 font-semibold' : 'bg-green-100 text-green-700';
      return `<tr class="border-t hover:bg-gray-50">
        <td class="px-4 py-2 text-sm">${subj}</td>
        <td class="px-4 py-2 text-sm text-center">${info.hours}</td>
        <td class="px-4 py-2 text-sm text-center"><span class="px-2 py-1 rounded-full text-xs ${colorClass}">${pct}%</span></td>
        <td class="px-4 py-2 text-sm text-center">${info.count} ครั้ง</td>
      </tr>`;
    }).join('');
    stats = `
    <div class="bg-white rounded-2xl p-5 border border-blue-100 mb-4"><p class="text-sm text-gray-500">ข้อมูลนักศึกษา</p><p class="font-bold text-lg">${APP.currentUser.name}</p></div>
    ${Object.keys(subjectMap).length ? `
    <div class="bg-white rounded-2xl p-5 border border-blue-100 mb-6">
      <h3 class="font-bold mb-3 flex items-center gap-2"><i data-lucide="bar-chart-3" class="w-5 h-5 text-primary"></i>เปอร์เซ็นต์การลาแต่ละรายวิชา</h3>
      <div class="overflow-x-auto"><table class="w-full text-sm">
        <thead><tr class="bg-surface text-left"><th class="px-4 py-2 font-semibold">รายวิชา</th><th class="px-4 py-2 font-semibold text-center">ชม.ลารวม</th><th class="px-4 py-2 font-semibold text-center">%ลา</th><th class="px-4 py-2 font-semibold text-center">จำนวนครั้ง</th></tr></thead>
        <tbody>${leaveRows}</tbody>
      </table></div>
      <div class="mt-3 flex flex-wrap gap-3 text-xs text-gray-500">
        <span class="flex items-center gap-1"><span class="w-3 h-3 rounded-full bg-red-500"></span> ≥ 20% เกินเกณฑ์</span>
        <span class="flex items-center gap-1"><span class="w-3 h-3 rounded-full bg-yellow-500"></span> 15-19% ใกล้เกินเกณฑ์</span>
        <span class="flex items-center gap-1"><span class="w-3 h-3 rounded-full bg-green-500"></span> < 15% ปกติ</span>
      </div>
    </div>` : '<div class="bg-blue-50 rounded-2xl p-4 mb-6 text-sm text-blue-700"><i data-lucide="info" class="w-4 h-4 inline"></i> ยังไม่มีข้อมูลการลา</div>'}`;
  }

  return `<h2 class="text-xl font-bold text-gray-800 mb-4"><i data-lucide="layout-dashboard" class="w-6 h-6 inline mr-2"></i>หน้าหลัก</h2>
  ${stats}
  <div class="grid grid-cols-1 lg:grid-cols-2 gap-4">
    <div class="bg-white rounded-2xl p-5 border border-blue-100">
      <h3 class="font-bold mb-3 flex items-center gap-2"><i data-lucide="calendar" class="w-5 h-5 text-primary"></i>ปฏิทินการศึกษา</h3>
      <div id="dashCalendar"></div>
    </div>
    <div class="bg-white rounded-2xl p-5 border border-blue-100">
      <h3 class="font-bold mb-3 flex items-center gap-2"><i data-lucide="megaphone" class="w-5 h-5 text-primary"></i>ประกาศ</h3>
      ${announcements.length ? announcements.map(a => `<div class="p-3 bg-surface rounded-xl mb-2"><p class="font-medium text-sm">${a.announcement_title || ''}</p><p class="text-xs text-gray-500">${a.announcement_date || ''}</p><p class="text-xs text-gray-600 mt-1">${(a.announcement_content || '').substring(0, 100)}</p></div>`).join('') : '<p class="text-gray-400 text-sm text-center py-8">ยังไม่มีประกาศ</p>'}
    </div>
  </div>
  <div class="bg-gradient-to-r from-blue-50 to-sky-50 rounded-2xl p-5 border border-blue-200 mt-6">
    <h3 class="font-bold mb-3 flex items-center gap-2 text-primary"><i data-lucide="phone" class="w-5 h-5"></i>ติดต่อสอบถาม</h3>
    <p class="text-sm text-gray-700 mb-3"><i data-lucide="phone-call" class="w-4 h-4 inline mr-1"></i>โทร. <a href="tel:023542320" class="font-semibold text-primary hover:underline">02 354 2320</a></p>
    <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-2">
      <div class="flex items-center gap-2 bg-white rounded-xl px-3 py-2 border border-blue-100"><span class="w-2 h-2 bg-primary rounded-full"></span><span class="text-sm">งานวิชาการ (admin) ต่อ <b>310</b></span></div>
      <div class="flex items-center gap-2 bg-white rounded-xl px-3 py-2 border border-blue-100"><span class="w-2 h-2 bg-emerald-500 rounded-full"></span><span class="text-sm">งานบริหารหลักสูตร ต่อ <b>311</b></span></div>
      <div class="flex items-center gap-2 bg-white rounded-xl px-3 py-2 border border-blue-100"><span class="w-2 h-2 bg-amber-500 rounded-full"></span><span class="text-sm">งานข้อสอบและวัดประเมินผล ต่อ <b>320</b></span></div>
      <div class="flex items-center gap-2 bg-white rounded-xl px-3 py-2 border border-blue-100"><span class="w-2 h-2 bg-purple-500 rounded-full"></span><span class="text-sm">งานดิจิตอล ต่อ <b>322</b></span></div>
      <div class="flex items-center gap-2 bg-white rounded-xl px-3 py-2 border border-blue-100"><span class="w-2 h-2 bg-pink-500 rounded-full"></span><span class="text-sm">งานเลขา ต่อ <b>330</b></span></div>
      <div class="flex items-center gap-2 bg-white rounded-xl px-3 py-2 border border-blue-100"><span class="w-2 h-2 bg-red-500 rounded-full"></span><span class="text-sm">งานทะเบียน ต่อ <b>340</b></span></div>
    </div>
  </div>`;
}

function statCard(icon, label, value, unit, color) {
  return `<div class="card-stat bg-white rounded-2xl p-5 border border-blue-100"><div class="flex items-center gap-4"><div class="w-12 h-12 ${color} rounded-xl flex items-center justify-center"><i data-lucide="${icon}" class="w-6 h-6 text-white"></i></div><div><p class="text-sm text-gray-500">${label}</p><p class="text-2xl font-bold text-gray-800">${value} <span class="text-sm font-normal text-gray-500">${unit}</span></p></div></div></div>`;
}

// ======================== STUDENTS ========================
function studentsPage() {
  const isAdmin = APP.currentRole === 'admin' || APP.currentRole === 'academic';
  const isClassTeacher = APP.currentRole === 'classTeacher';
  const canEdit = isAdmin || APP.currentRole === 'teacher' || isClassTeacher;
  const allStudents = getDataByType('student');
  const selectedYearLevel = APP.filters._studentYearLevel || '';

  // ClassTeacher: room-based selector
  if (isClassTeacher) {
    const yr = APP.currentUser.responsible_year || '1';
    const myYrStudents = allStudents.filter(s => norm(s.year_level) === norm(yr));
    const selectedRoom = APP.filters._studentRoom || '';
    const roomACount = myYrStudents.filter(s => norm(s.room).toUpperCase() === 'A').length;
    const roomBCount = myYrStudents.filter(s => norm(s.room).toUpperCase() === 'B').length;

    let headerHtml = `<div class="flex flex-wrap items-center justify-between gap-3 mb-4">
      <h2 class="text-xl font-bold text-gray-800"><i data-lucide="users" class="w-6 h-6 inline mr-2"></i>ข้อมูลนักศึกษา ชั้นปีที่ ${yr}</h2>
    </div>
    <div class="grid grid-cols-2 sm:grid-cols-3 gap-3 mb-4">
      <div class="bg-white rounded-2xl p-4 border border-blue-100 text-center"><p class="text-sm text-gray-500">ทั้งหมด</p><p class="text-2xl font-bold text-primary">${myYrStudents.length} <span class="text-sm font-normal text-gray-500">คน</span></p></div>
      <div class="bg-white rounded-2xl p-4 border border-blue-100 text-center"><p class="text-sm text-gray-500">ห้อง A</p><p class="text-2xl font-bold text-blue-600">${roomACount} <span class="text-sm font-normal text-gray-500">คน</span></p></div>
      <div class="bg-white rounded-2xl p-4 border border-blue-100 text-center"><p class="text-sm text-gray-500">ห้อง B</p><p class="text-2xl font-bold text-purple-600">${roomBCount} <span class="text-sm font-normal text-gray-500">คน</span></p></div>
    </div>
    <div class="bg-white rounded-2xl p-4 border border-blue-100 mb-4">
      <div class="flex items-center gap-3">
        <label class="text-sm font-medium text-gray-700">เลือกห้อง:</label>
        <div class="flex gap-2">
          <button onclick="APP.filters._studentRoom='';APP.pagination.page=1;renderCurrentPage()" class="px-4 py-2 rounded-xl text-sm font-medium ${!selectedRoom ? 'bg-primary text-white' : 'bg-gray-100 text-gray-600 hover:bg-gray-200'}">ทั้งหมด</button>
          <button onclick="APP.filters._studentRoom='A';APP.pagination.page=1;renderCurrentPage()" class="px-4 py-2 rounded-xl text-sm font-medium ${selectedRoom === 'A' ? 'bg-primary text-white' : 'bg-gray-100 text-gray-600 hover:bg-gray-200'}">ห้อง A</button>
          <button onclick="APP.filters._studentRoom='B';APP.pagination.page=1;renderCurrentPage()" class="px-4 py-2 rounded-xl text-sm font-medium ${selectedRoom === 'B' ? 'bg-primary text-white' : 'bg-gray-100 text-gray-600 hover:bg-gray-200'}">ห้อง B</button>
        </div>
      </div>
    </div>`;

    let data = myYrStudents;
    if (selectedRoom) data = data.filter(s => norm(s.room).toUpperCase() === selectedRoom.toUpperCase());
    data = applyFilters(data);
    const total = data.length;
    const paged = paginate(data);

    return headerHtml + `
    ${filterBar({ semester: false, year: false, yearLevel: false })}
    <div class="bg-white rounded-2xl border border-blue-100 overflow-hidden">
      <div class="overflow-x-auto"><table class="w-full text-sm">
        <thead><tr class="bg-surface text-left"><th class="px-4 py-3 font-semibold">รหัสนักศึกษา</th><th class="px-4 py-3 font-semibold">ชื่อ-สกุล</th><th class="px-4 py-3 font-semibold">ห้อง</th><th class="px-4 py-3 font-semibold">รุ่นที่</th><th class="px-4 py-3 font-semibold">สถานภาพ</th><th class="px-4 py-3"></th></tr></thead>
        <tbody>${paged.length ? paged.map(s => `<tr class="border-t hover:bg-gray-50">
          <td class="px-4 py-3">${s.student_id || ''}</td><td class="px-4 py-3 font-medium">${s.name || ''}</td><td class="px-4 py-3">${s.room || '-'}</td><td class="px-4 py-3">${s.batch || ''}</td>
          <td class="px-4 py-3"><span class="px-2 py-1 rounded-full text-xs ${s.status === 'กำลังศึกษา' ? 'bg-green-100 text-green-700' : s.status === 'สำเร็จการศึกษา' ? 'bg-blue-100 text-blue-700' : 'bg-yellow-100 text-yellow-700'}">${s.status || ''}</span></td>
          <td class="px-4 py-3"><div class="flex gap-1"><button onclick="showStudentDetail('${s.__backendId}')" class="text-gray-400 hover:text-primary" title="ดูข้อมูล"><i data-lucide="eye" class="w-4 h-4"></i></button></div></td></tr>`).join('') : '<tr><td colspan="6" class="px-4 py-8 text-center text-gray-400">ไม่มีข้อมูล</td></tr>'}</tbody>
      </table></div>
    </div>
    ${paginationHTML(total, APP.pagination.perPage, APP.pagination.page, 'changePage')}`;
  }

  // Teacher: button-based year selector
  if (APP.currentRole === 'teacher') {
    const myAllStudents = allStudents.filter(s => s.advisor === APP.currentUser.name);
    const selectedYearLevel = APP.filters._studentYearLevel || '';

    let headerHtml = `<div class="flex flex-wrap items-center justify-between gap-3 mb-4">
      <h2 class="text-xl font-bold text-gray-800"><i data-lucide="users" class="w-6 h-6 inline mr-2"></i>ข้อมูลนักศึกษา</h2>
    </div>
    <div class="bg-white rounded-2xl p-4 border border-blue-100 mb-4">
      <div class="flex items-center gap-3 flex-wrap">
        <label class="text-sm font-medium text-gray-700">ชั้นปี:</label>
        <div class="flex gap-2 flex-wrap">
          <button onclick="APP.filters._studentYearLevel='';APP.pagination.page=1;renderCurrentPage()" class="px-4 py-2 rounded-xl text-sm font-medium ${!selectedYearLevel ? 'bg-primary text-white' : 'bg-gray-100 text-gray-600 hover:bg-gray-200'}">แสดงทั้งหมด</button>
          ${['1', '2', '3', '4'].map(yr => `<button onclick="APP.filters._studentYearLevel='${yr}';APP.pagination.page=1;renderCurrentPage()" class="px-4 py-2 rounded-xl text-sm font-medium ${selectedYearLevel === yr ? 'bg-primary text-white' : 'bg-gray-100 text-gray-600 hover:bg-gray-200'}">ชั้นปี ${yr}</button>`).join('')}
        </div>
      </div>
    </div>`;

    let data = myAllStudents;
    if (selectedYearLevel) data = data.filter(s => norm(s.year_level) === selectedYearLevel);
    data = applyFilters(data);
    const total = data.length;
    const paged = paginate(data);

    return headerHtml + `
    ${filterBar({ semester: false, year: false, yearLevel: false })}
    <div class="bg-white rounded-2xl border border-blue-100 overflow-hidden">
      <div class="overflow-x-auto"><table class="w-full text-sm">
        <thead><tr class="bg-surface text-left"><th class="px-4 py-3 font-semibold">รหัสนักศึกษา</th><th class="px-4 py-3 font-semibold">ชื่อ-สกุล</th><th class="px-4 py-3 font-semibold">ชั้นปี</th><th class="px-4 py-3 font-semibold">รุ่นที่</th><th class="px-4 py-3 font-semibold">สถานภาพ</th><th class="px-4 py-3"></th></tr></thead>
        <tbody>${paged.length ? paged.map(s => `<tr class="border-t hover:bg-gray-50">
          <td class="px-4 py-3">${s.student_id || ''}</td><td class="px-4 py-3 font-medium">${s.name || ''}</td><td class="px-4 py-3">${s.year_level || ''}</td><td class="px-4 py-3">${s.batch || ''}</td>
          <td class="px-4 py-3"><span class="px-2 py-1 rounded-full text-xs ${s.status === 'กำลังศึกษา' ? 'bg-green-100 text-green-700' : s.status === 'สำเร็จการศึกษา' ? 'bg-blue-100 text-blue-700' : 'bg-yellow-100 text-yellow-700'}">${s.status || ''}</span></td>
          <td class="px-4 py-3"><div class="flex gap-1"><button onclick="showStudentDetail('${s.__backendId}')" class="text-gray-400 hover:text-primary" title="ดูข้อมูล"><i data-lucide="eye" class="w-4 h-4"></i></button></div></td></tr>`).join('') : '<tr><td colspan="6" class="px-4 py-8 text-center text-gray-400">ไม่มีข้อมูล</td></tr>'}</tbody>
      </table></div>
    </div>
    ${paginationHTML(total, APP.pagination.perPage, APP.pagination.page, 'changePage')}`;
  }

  let headerHtml = `<div class="flex flex-wrap items-center justify-between gap-3 mb-4">
    <h2 class="text-xl font-bold text-gray-800"><i data-lucide="users" class="w-6 h-6 inline mr-2"></i>ข้อมูลนักศึกษา</h2>
    ${isAdmin ? `<div class="flex gap-2"><button onclick="showAddStudentModal()" class="flex items-center gap-2 px-4 py-2 bg-primary text-white rounded-xl hover:bg-primaryDark text-sm"><i data-lucide="plus" class="w-4 h-4"></i>เพิ่มนักศึกษา</button>${csvUploadBtn('student', 'name,student_id,batch,status,phone,email,parent_name,parent_phone,advisor,year_level,room,national_id')}</div>` : ''}
  </div>
  <div class="bg-white rounded-2xl p-4 border border-blue-100 mb-4">
    <div class="flex items-center gap-3">
      <label class="text-sm font-medium text-gray-700">ชั้นปี:</label>
      <select onchange="APP.filters._studentYearLevel=this.value;APP.pagination.page=1;renderCurrentPage()" class="border border-gray-200 rounded-xl px-3 py-2 text-sm">
        <option value="">-- เลือกชั้นปี --</option>
        <option value="1" ${selectedYearLevel === '1' ? 'selected' : ''}>ชั้นปี 1</option>
        <option value="2" ${selectedYearLevel === '2' ? 'selected' : ''}>ชั้นปี 2</option>
        <option value="3" ${selectedYearLevel === '3' ? 'selected' : ''}>ชั้นปี 3</option>
        <option value="4" ${selectedYearLevel === '4' ? 'selected' : ''}>ชั้นปี 4</option>
      </select>
      ${selectedYearLevel ? '<span class="text-xs text-gray-500">แสดงข้อมูลชั้นปี ' + selectedYearLevel + '</span>' : ''}
    </div>
  </div>`;

  if (!selectedYearLevel) return headerHtml + noYearSelectedMsg('นักศึกษา (กรุณาเลือกชั้นปี)');

  let data = allStudents.filter(s => norm(s.year_level) === selectedYearLevel);
  data = applyFilters(data);
  const total = data.length;
  const paged = paginate(data);

  return headerHtml + `
  ${filterBar({ semester: false, year: false, yearLevel: false })}
  <div class="bg-white rounded-2xl border border-blue-100 overflow-hidden">
    <div class="overflow-x-auto"><table class="w-full text-sm">
      <thead><tr class="bg-surface text-left"><th class="px-4 py-3 font-semibold">รหัสนักศึกษา</th><th class="px-4 py-3 font-semibold">ชื่อ-สกุล</th><th class="px-4 py-3 font-semibold">ชั้นปี</th><th class="px-4 py-3 font-semibold">รุ่นที่</th><th class="px-4 py-3 font-semibold">สถานภาพ</th><th class="px-4 py-3"></th></tr></thead>
      <tbody>${paged.length ? paged.map(s => `<tr class="border-t hover:bg-gray-50">
        <td class="px-4 py-3">${s.student_id || ''}</td><td class="px-4 py-3 font-medium">${s.name || ''}</td><td class="px-4 py-3">${s.year_level || ''}</td><td class="px-4 py-3">${s.batch || ''}</td>
        <td class="px-4 py-3"><span class="px-2 py-1 rounded-full text-xs ${s.status === 'กำลังศึกษา' ? 'bg-green-100 text-green-700' : s.status === 'สำเร็จการศึกษา' ? 'bg-blue-100 text-blue-700' : 'bg-yellow-100 text-yellow-700'}">${s.status || ''}</span></td>
        <td class="px-4 py-3"><div class="flex gap-1"><button onclick="showStudentDetail('${s.__backendId}')" class="text-gray-400 hover:text-primary" title="ดูข้อมูล"><i data-lucide="eye" class="w-4 h-4"></i></button>${isAdmin ? `<button onclick="showEditStudentModal('${s.__backendId}')" class="text-blue-400 hover:text-blue-600" title="แก้ไข"><i data-lucide="pencil" class="w-4 h-4"></i></button><button onclick="deleteRecord('${s.__backendId}')" class="text-red-400 hover:text-red-600" title="ลบ"><i data-lucide="trash-2" class="w-4 h-4"></i></button>` : ''}</div></td></tr>`).join('') : '<tr><td colspan="6" class="px-4 py-8 text-center text-gray-400">ไม่มีข้อมูล</td></tr>'}</tbody>
    </table></div>
  </div>
  ${paginationHTML(total, APP.pagination.perPage, APP.pagination.page, 'changePage')}`;
}

function studentInfoPage() {
  // Student self-view
  const stu = APP.currentUser.data;
  if (!stu) return '<div class="bg-white rounded-2xl p-8 text-center border"><p class="text-gray-400">ไม่พบข้อมูลนักศึกษา กรุณาติดต่อผู้ดูแลระบบ</p></div>';
  return `<h2 class="text-xl font-bold text-gray-800 mb-4">ข้อมูลนักศึกษา</h2>
  <div class="bg-white rounded-2xl p-6 border border-blue-100">
    <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
      ${infoRow('ชื่อ-สกุล', stu.name)}${infoRow('รหัสนักศึกษา', stu.student_id)}${infoRow('รุ่นที่', stu.batch)}${infoRow('สถานภาพ', stu.status)}
      ${infoRow('ชั้นปี', stu.year_level)}${infoRow('ห้อง', stu.room)}${infoRow('โทร', stu.phone)}
      ${infoRow('E-mail', stu.email)}${infoRow('ผู้ปกครอง', stu.parent_name)}${infoRow('โทรผู้ปกครอง', stu.parent_phone)}
      ${infoRow('อาจารย์ที่ปรึกษา', stu.advisor)}
    </div>
    </div>
  </div>`;
}

function infoRow(l, v) { return `<div><p class="text-xs text-gray-500">${l}</p><p class="font-medium">${v || '-'}</p></div>` }

// Helper: show "รหัส ชื่อวิชา" or just "ชื่อวิชา" if no code
function subjectLabel(code, name) { return code ? `${code} ${name || ''}` : name || '' }

function showAddStudentModal() {
  showModal('เพิ่มนักศึกษา', `
    <form id="addStudentForm" class="space-y-3">
      <div class="grid grid-cols-1 md:grid-cols-2 gap-3">
        <div><label class="block text-xs text-gray-600 mb-1">ชื่อ-สกุล *</label><input name="name" required class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">รหัสนักศึกษา *</label><input name="student_id" required class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">รุ่นที่</label><input name="batch" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="เช่น 36"></div>
        <div><label class="block text-xs text-gray-600 mb-1">เลขบัตรประชาชน</label><input name="national_id" maxlength="13" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">สถานภาพ</label><select name="status" class="w-full border rounded-xl px-3 py-2 text-sm"><option>กำลังศึกษา</option><option>พักการศึกษา</option><option>สำเร็จการศึกษา</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">ชั้นปี</label><select name="year_level" class="w-full border rounded-xl px-3 py-2 text-sm"><option>1</option><option>2</option><option>3</option><option>4</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">ห้อง</label><input name="room" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">โทรศัพท์</label><input name="phone" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">E-mail</label><input name="email" type="email" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">ชื่อผู้ปกครอง</label><input name="parent_name" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">โทรผู้ปกครอง</label><input name="parent_phone" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">อาจารย์ที่ปรึกษา</label><input name="advisor" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      </div>
      <button type="submit" class="w-full mt-3 bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึก</button>
    </form>
  `);
  document.getElementById('addStudentForm').onsubmit = async (e) => {
    e.preventDefault();
    await withLoading(e.target, async () => {
      const fd = new FormData(e.target);
      const obj = { type: 'student', created_at: new Date().toISOString() };
      fd.forEach((v, k) => obj[k] = v);
      if (APP.allData.filter(d => d.type === 'student').length >= 999) { showToast('ข้อมูลเต็ม', 'error'); return }
      const r = await GSheetDB.create(obj);
      if (r.isOk) { showToast('เพิ่มนักศึกษาสำเร็จ'); closeModal() } else showToast('เกิดข้อผิดพลาด', 'error');
    });
  };
}

function showStudentDetail(id) {
  const s = APP.allData.find(d => d.__backendId === id);
  if (!s) return;
  showModal('ข้อมูลนักศึกษา', `<div class="grid grid-cols-2 gap-3">
    ${infoRow('ชื่อ-สกุล', s.name)}${infoRow('รหัสนักศึกษา', s.student_id)}${infoRow('รุ่นที่', s.batch)}${infoRow('สถานภาพ', s.status)}
    ${infoRow('ชั้นปี', s.year_level)}${infoRow('ห้อง', s.room)}${infoRow('โทร', s.phone)}${infoRow('E-mail', s.email)}
    ${infoRow('ผู้ปกครอง', s.parent_name)}${infoRow('โทรผู้ปกครอง', s.parent_phone)}${infoRow('อาจารย์ที่ปรึกษา', s.advisor)}
  </div>`);
}

// ======================== SUBJECTS ========================
function subjectsPage() {
  const isAdmin = APP.currentRole === 'admin' || APP.currentRole === 'academic';
  const canEdit = isAdmin || APP.currentRole === 'teacher' || APP.currentRole === 'classTeacher';
  const allSubjects = getDataByType('subject');
  const selectedYear = APP.filters._pageYear || '';

  let headerHtml = `<div class="flex flex-wrap items-center justify-between gap-3 mb-4">
    <h2 class="text-xl font-bold text-gray-800"><i data-lucide="book-open" class="w-6 h-6 inline mr-2"></i>รายวิชาที่เปิดสอน</h2>
    ${isAdmin ? `<div class="flex gap-2"><button onclick="showAddSubjectModal()" class="flex items-center gap-2 px-4 py-2 bg-primary text-white rounded-xl hover:bg-primaryDark text-sm"><i data-lucide="plus" class="w-4 h-4"></i>เพิ่มรายวิชา</button>${csvUploadBtn('subject', 'subject_name,coordinator,year_level,room,credits,semester,academic_year')}</div>` : ''}
  </div>`;
  headerHtml += yearPickerBar(allSubjects, 'ปีการศึกษา');

  if (!selectedYear) return headerHtml + noYearSelectedMsg('รายวิชา');

  let data = applyFilters(allSubjects.filter(s => (s.academic_year || '') === selectedYear));
  if (APP.currentRole === 'classTeacher') data = data.filter(s => norm(s.year_level) === norm(APP.currentUser.responsible_year || '1'));
  if (APP.currentRole === 'student' && APP.currentUser.data) data = data.filter(s => norm(s.year_level) === norm(APP.currentUser.data.year_level));
  const total = data.length; const paged = paginate(data);

  return headerHtml + `
  ${filterBar({ semester: true, year: false, yearLevel: true })}
  <div class="bg-white rounded-2xl border border-blue-100 overflow-hidden">
    <div class="overflow-x-auto"><table class="w-full text-sm">
      <thead><tr class="bg-surface text-left"><th class="px-4 py-3 font-semibold">รหัสวิชา</th><th class="px-4 py-3 font-semibold">ชื่อรายวิชา</th><th class="px-4 py-3 font-semibold">ผู้ประสานงาน</th><th class="px-4 py-3 font-semibold">ชั้นปี</th><th class="px-4 py-3 font-semibold">หน่วยกิต</th><th class="px-4 py-3 font-semibold">ภาค/ปี</th>${isAdmin ? '<th class="px-4 py-3"></th>' : ''}</tr></thead>
      <tbody>${paged.length ? paged.map(s => `<tr class="border-t hover:bg-gray-50">
        <td class="px-4 py-3 font-mono text-primary">${s.subject_code || ''}</td><td class="px-4 py-3 font-medium">${s.subject_name || ''}</td><td class="px-4 py-3">${s.coordinator || ''}</td>
        <td class="px-4 py-3">${s.year_level || ''}</td><td class="px-4 py-3">${s.credits || ''}</td>
        <td class="px-4 py-3">${semLabel(s.semester)}/${s.academic_year || ''}</td>
        ${isAdmin ? `<td class="px-4 py-3"><div class="flex gap-1"><button onclick="showEditSubjectModal('${s.__backendId}')" class="text-blue-400 hover:text-blue-600" title="แก้ไข"><i data-lucide="pencil" class="w-4 h-4"></i></button><button onclick="deleteRecord('${s.__backendId}')" class="text-red-400 hover:text-red-600" title="ลบ"><i data-lucide="trash-2" class="w-4 h-4"></i></button></div></td>` : ''}</tr>`).join('') : '<tr><td colspan="7" class="px-4 py-8 text-center text-gray-400">ไม่มีข้อมูล</td></tr>'}</tbody>
    </table></div>
  </div>
  ${paginationHTML(total, APP.pagination.perPage, APP.pagination.page, 'changePage')}`;
}

function showAddSubjectModal() {
  showModal('เพิ่มรายวิชา', `
    <form id="addSubjectForm" class="space-y-3">
      <div><label class="block text-xs text-gray-600 mb-1">ชื่อรายวิชา *</label><input name="subject_name" required class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <div><label class="block text-xs text-gray-600 mb-1">อาจารย์ผู้ประสานงาน (คั่นด้วย ,)</label><input name="coordinator" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="อ.ก, อ.ข"></div>
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-xs text-gray-600 mb-1">ชั้นปี</label><select name="year_level" class="w-full border rounded-xl px-3 py-2 text-sm"><option>1</option><option>2</option><option>3</option><option>4</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">ห้อง</label><input name="room" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">หน่วยกิต</label><input name="credits" type="number" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">ภาคการศึกษา</label><select name="semester" class="w-full border rounded-xl px-3 py-2 text-sm"><option value="1">1</option><option value="2">2</option><option value="3">ฤดูร้อน</option></select></div>
      </div>
      <div><label class="block text-xs text-gray-600 mb-1">ปีการศึกษา</label><input name="academic_year" class="w-full border rounded-xl px-3 py-2 text-sm" value="2568"></div>
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึก</button>
    </form>
  `);
  document.getElementById('addSubjectForm').onsubmit = async (e) => {
    e.preventDefault();
    await withLoading(e.target, async () => {
      const fd = new FormData(e.target);
      const obj = { type: 'subject', created_at: new Date().toISOString() };
      fd.forEach((v, k) => obj[k] = k === 'credits' ? Number(v) : v);
      if (APP.allData.filter(d => d.type === 'subject').length >= 999) { showToast('ข้อมูลเต็ม', 'error'); return }
      const r = await GSheetDB.create(obj);
      if (r.isOk) { showToast('เพิ่มรายวิชาสำเร็จ'); closeModal() } else showToast('เกิดข้อผิดพลาด', 'error');
    });
  };
}

// ======================== SCHEDULE ========================
function schedulePage() {
  return `<div class="flex flex-wrap items-center justify-between gap-3 mb-4">
    <h2 class="text-xl font-bold text-gray-800"><i data-lucide="calendar" class="w-6 h-6 inline mr-2"></i>ตารางเรียน/ตารางสอบ</h2>
    <div class="flex gap-2">
      <button onclick="showAddScheduleModal()" class="flex items-center gap-2 px-4 py-2 bg-primary text-white rounded-xl hover:bg-primaryDark text-sm"><i data-lucide="plus" class="w-4 h-4"></i>เพิ่มตาราง</button>
      ${csvUploadBtn('schedule', 'subject_name,schedule_date,schedule_time,schedule_type,room,year_level')}
    </div>
  </div>
  <div class="bg-white rounded-2xl border border-blue-100 p-5">
    <div id="scheduleCalendar"></div>
  </div>
  <div class="mt-4 bg-white rounded-2xl border border-blue-100 overflow-hidden">
    <div class="overflow-x-auto"><table class="w-full text-sm">
      <thead><tr class="bg-surface text-left"><th class="px-4 py-3 font-semibold">วันที่</th><th class="px-4 py-3 font-semibold">เวลา</th><th class="px-4 py-3 font-semibold">รายวิชา</th><th class="px-4 py-3 font-semibold">ประเภท</th><th class="px-4 py-3 font-semibold">ห้อง</th><th class="px-4 py-3"></th></tr></thead>
      <tbody>${getDataByType('schedule').sort((a, b) => (a.schedule_date || '').localeCompare(b.schedule_date || '')).slice(0, 20).map(s => `<tr class="border-t hover:bg-gray-50">
        <td class="px-4 py-3">${s.schedule_date || ''}</td><td class="px-4 py-3">${s.schedule_time || ''}</td>
        <td class="px-4 py-3">${s.subject_name || ''}</td>
        <td class="px-4 py-3"><span class="px-2 py-1 rounded-full text-xs ${s.schedule_type === 'สอบ' ? 'bg-red-100 text-red-700' : 'bg-blue-100 text-blue-700'}">${s.schedule_type || 'เรียน'}</span></td>
        <td class="px-4 py-3">${s.room || ''}</td>
        <td class="px-4 py-3"><div class="flex gap-1"><button onclick="showEditScheduleModal('${s.__backendId}')" class="text-blue-400 hover:text-blue-600" title="แก้ไข"><i data-lucide="pencil" class="w-4 h-4"></i></button><button onclick="deleteRecord('${s.__backendId}')" class="text-red-400 hover:text-red-600" title="ลบ"><i data-lucide="trash-2" class="w-4 h-4"></i></button></div></td>
      </tr>`).join('') || '<tr><td colspan="6" class="px-4 py-8 text-center text-gray-400">ไม่มีข้อมูล</td></tr>'}</tbody>
    </table></div>
  </div>`;
}

function showAddScheduleModal() {
  showModal('เพิ่มตารางเรียน/สอบ', `
    <form id="addScheduleForm" class="space-y-3">
      <div><label class="block text-xs text-gray-600 mb-1">รายวิชา</label><input name="subject_name" required class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-xs text-gray-600 mb-1">วันที่</label><input name="schedule_date" type="date" required class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">เวลา</label><input name="schedule_time" type="time" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">ประเภท</label><select name="schedule_type" class="w-full border rounded-xl px-3 py-2 text-sm"><option>เรียน</option><option>สอบ</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">ห้อง</label><input name="room" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">ชั้นปี</label><select name="year_level" class="w-full border rounded-xl px-3 py-2 text-sm"><option>1</option><option>2</option><option>3</option><option>4</option></select></div>
      </div>
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึก</button>
    </form>
  `);
  document.getElementById('addScheduleForm').onsubmit = async (e) => {
    e.preventDefault();
    await withLoading(e.target, async () => {
      const fd = new FormData(e.target);
      const obj = { type: 'schedule', created_at: new Date().toISOString() }; fd.forEach((v, k) => obj[k] = v);
      const r = await GSheetDB.create(obj);
      if (r.isOk) { showToast('เพิ่มตารางสำเร็จ'); closeModal() } else showToast('เกิดข้อผิดพลาด', 'error');
    });
  };
}

// ======================== GRADES ========================
function gradesPage() {
  const isAdmin = APP.currentRole === 'admin' || APP.currentRole === 'academic';
  const isExecutive = APP.currentRole === 'executive';
  const canEdit = isAdmin || APP.currentRole === 'teacher' || APP.currentRole === 'classTeacher';
  const isStudent = APP.currentRole === 'student';
  const canFilterByAdvisor = isAdmin || isExecutive;
  let allGrades = getDataByType('grade');

  // Build student list for non-student roles
  let studentSelector = '';
  let selectedStudentName = APP.filters._gradeStudent || '';

  if (!isStudent) {
    let studentList = getDataByType('student');
    if (APP.currentRole === 'classTeacher') {
      const yr = APP.currentUser.responsible_year || '1';
      studentList = studentList.filter(s => norm(s.year_level) === norm(yr));
    }
    if (APP.currentRole === 'teacher') {
      studentList = studentList.filter(s => s.advisor === APP.currentUser.name);
    }

    // Year level + Advisor filter for admin/academic/executive
    let advisorSelector = '';
    if (canFilterByAdvisor) {
      // Year level filter
      const selectedGradeYear = APP.filters._gradeYearLevel || '';
      if (selectedGradeYear) {
        studentList = studentList.filter(s => norm(s.year_level) === selectedGradeYear);
      }
      const yearLevelSelector = `<div class="bg-white rounded-2xl p-4 border border-blue-100 mb-4">
        <label class="block text-sm font-medium text-gray-700 mb-2"><i data-lucide="layers" class="w-4 h-4 inline mr-1"></i>กรองตามชั้นปี</label>
        <div class="flex flex-wrap gap-2">
          <button onclick="APP.filters._gradeYearLevel='';APP.filters._gradeStudent='';APP.filters._gradeSearch='';APP.pagination.page=1;renderCurrentPage()" class="px-4 py-2 rounded-xl text-sm font-medium ${selectedGradeYear === '' ? 'bg-primary text-white' : 'bg-gray-100 text-gray-600 hover:bg-gray-200'}">ทุกชั้นปี</button>
          ${['1', '2', '3', '4'].map(yr => `<button onclick="APP.filters._gradeYearLevel='${yr}';APP.filters._gradeStudent='';APP.filters._gradeSearch='';APP.pagination.page=1;renderCurrentPage()" class="px-4 py-2 rounded-xl text-sm font-medium ${selectedGradeYear === yr ? 'bg-primary text-white' : 'bg-gray-100 text-gray-600 hover:bg-gray-200'}">ชั้นปี ${yr}</button>`).join('')}
        </div>
        ${selectedGradeYear ? `<p class="text-xs text-gray-500 mt-2"><i data-lucide="info" class="w-3 h-3 inline mr-1"></i>แสดงเฉพาะนักศึกษาชั้นปีที่ ${selectedGradeYear} (${studentList.length} คน)</p>` : ''}
      </div>`;

      const allAdvisors = [...new Set(studentList.map(s => s.advisor).filter(Boolean))].sort();
      const selectedAdvisor = APP.filters._gradeAdvisor || '';
      if (selectedAdvisor) {
        studentList = studentList.filter(s => (s.advisor || '') === selectedAdvisor);
      }
      advisorSelector = `${yearLevelSelector}<div class="bg-white rounded-2xl p-4 border border-blue-100 mb-4">
        <label class="block text-sm font-medium text-gray-700 mb-2"><i data-lucide="user-check" class="w-4 h-4 inline mr-1"></i>กรองตามอาจารย์ที่ปรึกษา</label>
        <select onchange="APP.filters._gradeAdvisor=this.value;APP.filters._gradeStudent='';APP.filters._gradeSearch='';APP.pagination.page=1;renderCurrentPage()" class="w-full border border-gray-200 rounded-xl px-4 py-2.5 text-sm">
          <option value="">-- แสดงนักศึกษาทั้งหมด --</option>
          ${allAdvisors.map(a => `<option value="${a}" ${selectedAdvisor === a ? 'selected' : ''}>${a}</option>`).join('')}
        </select>
        ${selectedAdvisor ? `<p class="text-xs text-gray-500 mt-2"><i data-lucide="info" class="w-3 h-3 inline mr-1"></i>แสดงเฉพาะนักศึกษาในความดูแลของ ${selectedAdvisor} (${studentList.length} คน)</p>` : ''}
      </div>`;
    }

    // Store student list for search
    window._gradeStudentList = studentList;
    const searchVal = APP.filters._gradeSearch || '';
    let filteredList = studentList;
    if (searchVal) {
      const q = searchVal.toLowerCase();
      filteredList = studentList.filter(s => (s.name || '').toLowerCase().includes(q) || (s.student_id || '').toLowerCase().includes(q));
    }
    studentSelector = `${advisorSelector}<div class="bg-white rounded-2xl p-4 border border-blue-100 mb-4">
      <label class="block text-sm font-medium text-gray-700 mb-2"><i data-lucide="user-search" class="w-4 h-4 inline mr-1"></i>เลือกนักศึกษา</label>
      <div class="flex gap-2 mb-2">
        <div class="flex-1 relative"><i data-lucide="search" class="absolute left-3 top-2.5 w-4 h-4 text-gray-400"></i><input type="text" placeholder="พิมพ์ค้นหาชื่อหรือรหัส..." value="${searchVal}" class="w-full pl-10 pr-4 py-2 border border-gray-200 rounded-xl text-sm" oninput="clearTimeout(window._gradeSearchTimer);window._gradeSearchTimer=setTimeout(()=>{APP.filters._gradeSearch=this.value;APP.filters._gradeStudent='';APP.pagination.page=1;renderCurrentPage()},300)"></div>
      </div>
      <select onchange="APP.filters._gradeStudent=this.value;APP.pagination.page=1;renderCurrentPage()" class="w-full border border-gray-200 rounded-xl px-4 py-2.5 text-sm">
        <option value="">-- กรุณาเลือกนักศึกษา (${filteredList.length} คน) --</option>
        ${filteredList.map(s => `<option value="${s.student_id || s.name}" ${selectedStudentName === (s.student_id || s.name) ? 'selected' : ''}>${s.student_id || ''} — ${s.name || ''}</option>`).join('')}
      </select>
    </div>`;
  }

  // Filter data — match by student_id
  let data;
  if (isStudent && APP.currentUser.data) {
    const sid = norm(APP.currentUser.data.student_id);
    data = allGrades.filter(g => norm(g.student_id) === sid);
  } else if (selectedStudentName) {
    const sid = norm(selectedStudentName);
    data = allGrades.filter(g => norm(g.student_id) === sid);
  } else {
    data = [];
  }
  data = applyFilters(data);
  const total = data.length; const paged = paginate(data);

  // GPA calc
  let gpaSection = '';
  if (data.length && (isStudent || selectedStudentName)) {
    const gradeMap = { 'A': 4, 'B+': 3.5, 'B': 3, 'C+': 2.5, 'C': 2, 'D+': 1.5, 'D': 1, 'F': 0 };
    let totalCredits = 0, totalPoints = 0;
    data.forEach(g => { const gv = gradeMap[g.grade]; const cr = Number(g.credits) || 3; if (gv !== undefined) { totalPoints += gv * cr; totalCredits += cr } });
    const gpax = totalCredits ? ((totalPoints / totalCredits).toFixed(2)) : 'N/A';
    gpaSection = `<div class="bg-gradient-to-r from-primary to-accent text-white rounded-2xl p-5 mb-4 flex items-center justify-between">
      <div><p class="text-sm opacity-90">เกรดเฉลี่ยสะสม (GPAX)</p><p class="text-3xl font-bold">${gpax}</p></div>
      ${isStudent ? `<button onclick="showTranscript()" class="px-4 py-2 bg-white bg-opacity-20 rounded-xl hover:bg-opacity-30 text-sm">ใบแสดงผลการเรียน</button>` : `<button onclick="showTranscriptForStudent('${selectedStudentName.replace(/'/g, "\\'")}')" class="px-4 py-2 bg-white bg-opacity-20 rounded-xl hover:bg-opacity-30 text-sm">ใบแสดงผลการเรียน</button>`}
    </div>`;
  }

  // Show prompt if no student selected (non-student roles)
  let noSelectionMsg = '';
  if (!isStudent && !selectedStudentName) {
    noSelectionMsg = `<div class="bg-white rounded-2xl border border-blue-100 p-8 text-center text-gray-400"><i data-lucide="user-search" class="w-10 h-10 mx-auto mb-3 text-gray-300"></i><p>กรุณาเลือกนักศึกษาเพื่อดูผลการเรียน</p></div>`;
  }

  return `<div class="flex flex-wrap items-center justify-between gap-3 mb-4">
    <h2 class="text-xl font-bold text-gray-800"><i data-lucide="file-text" class="w-6 h-6 inline mr-2"></i>ผลการเรียน</h2>
    ${isAdmin ? `<div class="flex gap-2"><button onclick="showAddGradeModal()" class="flex items-center gap-2 px-4 py-2 bg-primary text-white rounded-xl hover:bg-primaryDark text-sm"><i data-lucide="plus" class="w-4 h-4"></i>เพิ่มผลการเรียน</button>${csvUploadBtn('grade', 'student_id,subject_code,subject_name,grade,credits,semester,academic_year')}</div>` : ''}
  </div>
  ${studentSelector}
  ${noSelectionMsg || `${filterBar()}
  ${gpaSection}
  <div class="bg-white rounded-2xl border border-blue-100 overflow-hidden">
    <div class="overflow-x-auto"><table class="w-full text-sm">
      <thead><tr class="bg-surface text-left"><th class="px-4 py-3 font-semibold">รหัสวิชา</th><th class="px-4 py-3 font-semibold">รายวิชา</th><th class="px-4 py-3 font-semibold">เกรด</th><th class="px-4 py-3 font-semibold">หน่วยกิต</th><th class="px-4 py-3 font-semibold">ภาค/ปี</th>${isAdmin ? '<th class="px-4 py-3"></th>' : ''}</tr></thead>
      <tbody>${paged.length ? paged.map(g => `<tr class="border-t hover:bg-gray-50">
        <td class="px-4 py-3 font-mono text-primary">${g.subject_code || ''}</td><td class="px-4 py-3">${g.subject_name || ''}</td>
        <td class="px-4 py-3"><span class="px-2 py-1 rounded-full text-xs font-bold ${g.grade === 'F' ? 'bg-red-100 text-red-700' : 'bg-green-100 text-green-700'}">${g.grade || ''}</span></td>
        <td class="px-4 py-3">${g.credits || ''}</td><td class="px-4 py-3">${semLabel(g.semester)}/${g.academic_year || ''}</td>
        ${isAdmin ? `<td class="px-4 py-3"><div class="flex gap-1"><button onclick="showEditGradeModal('${g.__backendId}')" class="text-blue-400 hover:text-blue-600" title="แก้ไข"><i data-lucide="pencil" class="w-4 h-4"></i></button><button onclick="deleteRecord('${g.__backendId}')" class="text-red-400 hover:text-red-600" title="ลบ"><i data-lucide="trash-2" class="w-4 h-4"></i></button></div></td>` : ''}</tr>`).join('') : '<tr><td colspan="5" class="px-4 py-8 text-center text-gray-400">ไม่มีข้อมูล</td></tr>'}</tbody>
    </table></div>
  </div>
  ${paginationHTML(total, APP.pagination.perPage, APP.pagination.page, 'changePage')}`}`;
}

function showAddGradeModal() {
  showModal('เพิ่มผลการเรียน', `
    <form id="addGradeForm" class="space-y-3">
      <div><label class="block text-xs text-gray-600 mb-1">นักศึกษา * (พิมพ์รหัสหรือเลือก)</label><input list="addGradeStudentList" name="student_id" required class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="พิมพ์รหัสนักศึกษา...">${studentDatalistHTML('addGradeStudentList')}</div>
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-xs text-gray-600 mb-1">รหัสวิชา</label><input name="subject_code" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">รายวิชา</label><input name="subject_name" required class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">เกรด</label><select name="grade" class="w-full border rounded-xl px-3 py-2 text-sm"><option>A</option><option>B+</option><option>B</option><option>C+</option><option>C</option><option>D+</option><option>D</option><option>F</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">หน่วยกิต</label><input name="credits" type="number" value="3" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">ภาคการศึกษา</label><select name="semester" class="w-full border rounded-xl px-3 py-2 text-sm"><option value="1">1</option><option value="2">2</option></select></div>
      </div>
      <div><label class="block text-xs text-gray-600 mb-1">ปีการศึกษา</label><input name="academic_year" value="2568" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึก</button>
    </form>
  `);
  document.getElementById('addGradeForm').onsubmit = async (e) => {
    e.preventDefault();
    await withLoading(e.target, async () => {
      const fd = new FormData(e.target);
      const obj = { type: 'grade', created_at: new Date().toISOString() }; fd.forEach((v, k) => obj[k] = k === 'credits' ? Number(v) : v);
      const r = await GSheetDB.create(obj);
      if (r.isOk) { showToast('เพิ่มผลการเรียนสำเร็จ'); closeModal() } else showToast('เกิดข้อผิดพลาด', 'error');
    });
  };
}

function showTranscript() {
  const stu = APP.currentUser.data; if (!stu) return;
  _renderTranscript(stu);
}

function showTranscriptForStudent(studentKey) {
  if (!studentKey) return;
  const stu = getDataByType('student').find(s => s.student_id === studentKey || s.name === studentKey);
  if (!stu) { showToast('ไม่พบข้อมูลนักศึกษา', 'error'); return; }
  _renderTranscript(stu);
}

function _renderTranscript(stu) {
  let grades = getDataByType('grade').filter(g => g.student_id === stu.student_id);
  // Respect current filters (academic year / semester) selected on grades page
  const filterYear = APP.filters.academicYear || '';
  const filterSem = APP.filters.semester || '';
  if (filterYear) grades = grades.filter(g => norm(g.academic_year) === filterYear);
  if (filterSem) grades = grades.filter(g => normSem(g.semester) === filterSem);

  const gradeMap = { 'A': 4, 'B+': 3.5, 'B': 3, 'C+': 2.5, 'C': 2, 'D+': 1.5, 'D': 1, 'F': 0 };

  const semesters = {};
  grades.forEach(g => {
    const key = `${g.semester || '1'}/${g.academic_year || ''}`;
    if (!semesters[key]) semesters[key] = { semester: g.semester || '1', year: g.academic_year || '', grades: [] };
    semesters[key].grades.push(g);
  });
  const semKeys = Object.keys(semesters).sort();

  let totalCreditsAll = 0, totalPointsAll = 0;
  let tableRows = '';
  semKeys.forEach(key => {
    const sem = semesters[key];
    tableRows += `<tr class="bg-blue-50"><td colspan="4" class="px-3 py-2 font-bold text-primary text-center">ภาคการศึกษาที่ ${sem.semester} ปีการศึกษา ${sem.year}</td></tr>`;
    let semCredits = 0, semPoints = 0;
    sem.grades.forEach(g => {
      const cr = Number(g.credits) || 0;
      const gv = gradeMap[g.grade];
      tableRows += `<tr class="border-t border-gray-200"><td class="px-3 py-1.5 font-mono text-xs">${g.subject_code || ''}</td><td class="px-3 py-1.5 text-xs">${g.subject_name || ''}</td><td class="px-3 py-1.5 text-center text-xs">${g.credits || ''}</td><td class="px-3 py-1.5 text-center text-xs font-bold">${g.grade || ''}</td></tr>`;
      if (gv !== undefined) { semPoints += gv * cr; semCredits += cr; }
    });
    const semGpa = semCredits ? (semPoints / semCredits).toFixed(2) : 'N/A';
    totalCreditsAll += semCredits; totalPointsAll += semPoints;
    tableRows += `<tr class="border-t bg-gray-50"><td colspan="2" class="px-3 py-1.5 text-xs text-right font-semibold">จำนวนหน่วยกิตรวม: ${semCredits}</td><td colspan="2" class="px-3 py-1.5 text-xs text-right font-semibold">คะแนนเฉลี่ย: ${semGpa}</td></tr>`;
  });

  const gpax = totalCreditsAll ? (totalPointsAll / totalCreditsAll).toFixed(2) : 'N/A';
  const studentProgram = stu.program || 'หลักสูตรพยาบาลศาสตรบัณฑิต';
  const studentLevel = stu.level || 'ปริญญาตรี';
  const studentBatch = stu.batch || '';
  const studentYearLevel = stu.year_level || '';
  const stuNameSafe = (stu.name || '').replace(/'/g, "\\'");

  const filterBadge = (filterYear || filterSem)
    ? `<div class="bg-blue-50 border border-blue-200 rounded-xl p-2 mb-3 text-xs text-blue-800 text-center"><i data-lucide="filter" class="w-3 h-3 inline mr-1"></i>แสดงเฉพาะ${filterSem ? ' ภาคการศึกษาที่ ' + filterSem : ''}${filterYear ? ' ปีการศึกษา ' + filterYear : ''}</div>`
    : '';

  showModal('ใบรายงานผลการเรียน', `
    ${filterBadge}
    <div id="transcriptContent" class="bg-white p-4 relative overflow-hidden" style="max-width:700px;margin:auto;">
      <div aria-hidden="true" style="position:absolute;inset:0;pointer-events:none;display:flex;align-items:center;justify-content:center;z-index:1;">
        <div style="transform:rotate(-30deg);font-size:54px;font-weight:800;color:rgba(220,38,38,0.18);white-space:nowrap;letter-spacing:6px;">ฉบับสำเนาไม่ใช่ตัวจริง</div>
      </div>
      <div class="relative" style="z-index:2;">
      <div class="text-center mb-3">
        <img src="https://cdn.jsdelivr.net/gh/JOB-BCNB-P/LOGO/Logo%20Thai.png" alt="Logo" style="width:60px;height:auto;margin:0 auto 6px auto;display:block;" onerror="this.style.display='none'">
        <p class="font-bold text-sm">${APP.config.college_name}</p>
        <p class="text-xs text-gray-600">ใบรายงานผลการเรียนนักศึกษารายภาคการศึกษา</p>
        <p class="text-xs text-gray-600">${studentProgram} ระดับ ${studentLevel}${studentBatch ? ' รุ่นที่ ' + studentBatch : ''}${studentYearLevel ? ' ชั้นปีที่ ' + studentYearLevel : ''}</p>
      </div>
      <div class="flex justify-between text-xs mb-3 px-1">
        <div><span class="text-gray-500">รหัสนักศึกษา:</span> <strong>${stu.student_id || ''}</strong></div>
        <div><span class="text-gray-500">ชื่อ-สกุล:</span> <strong>${stu.name || ''}</strong></div>
      </div>
      <table class="w-full text-sm border border-gray-300" style="border-collapse:collapse">
        <thead><tr class="bg-surface"><th class="px-3 py-2 text-left text-xs border border-gray-300" style="width:22%">รหัสวิชา</th><th class="px-3 py-2 text-left text-xs border border-gray-300" style="width:48%">รายวิชา</th><th class="px-3 py-2 text-center text-xs border border-gray-300" style="width:15%">หน่วยกิต</th><th class="px-3 py-2 text-center text-xs border border-gray-300" style="width:15%">ระดับคะแนน</th></tr></thead>
        <tbody>${tableRows}</tbody>
      </table>
      <div class="mt-3 border border-gray-300 rounded p-2 text-xs">
        <div class="flex justify-between"><span>รวมหน่วยกิตตลอดปีการศึกษา: <strong>${totalCreditsAll}</strong></span><span>คะแนนเฉลี่ยตลอดปีการศึกษา: <strong>${gpax}</strong></span></div>
        <div class="flex justify-between mt-1"><span>รวมหน่วยกิตสะสมตลอดหลักสูตร: <strong>${totalCreditsAll}</strong></span><span>คะแนนเฉลี่ยสะสมตลอดหลักสูตร: <strong>${gpax}</strong></span></div>
      </div>
      <div class="mt-3 text-xs text-gray-500">
        <p class="font-semibold mb-1">หมายเหตุ</p>
        <p>A : ดีเยี่ยม &nbsp; B+ : ดีมาก &nbsp; B : ดี &nbsp; C+ : ค่อนข้างดี &nbsp; C : พอใช้ &nbsp; D+ : อ่อน &nbsp; D : อ่อนมาก &nbsp; F : ตก</p>
      </div>
      </div>
    </div>
    <div class="flex justify-center mt-4">
      <button onclick="downloadTranscriptPDF('${stuNameSafe}')" class="flex items-center gap-2 px-6 py-2.5 bg-primary text-white rounded-xl hover:bg-primaryDark text-sm"><i data-lucide="download" class="w-4 h-4"></i>ดาวน์โหลด PDF</button>
    </div>
  `, null, 'max-w-3xl');
  setTimeout(() => lucide.createIcons(), 100);
}

async function downloadTranscriptPDF(studentKey) {
  const stu = getDataByType('student').find(s => s.student_id === studentKey || s.name === studentKey) || (APP.currentUser.data && (APP.currentUser.data.name === studentKey || APP.currentUser.data.student_id === studentKey) ? APP.currentUser.data : null);
  if (!stu) return;
  let grades = getDataByType('grade').filter(g => g.student_id === stu.student_id);
  // Respect current filters (academic year / semester) selected on grades page
  const filterYear = APP.filters.academicYear || '';
  const filterSem = APP.filters.semester || '';
  if (filterYear) grades = grades.filter(g => norm(g.academic_year) === filterYear);
  if (filterSem) grades = grades.filter(g => normSem(g.semester) === filterSem);
  const gradeMap = { 'A': 4, 'B+': 3.5, 'B': 3, 'C+': 2.5, 'C': 2, 'D+': 1.5, 'D': 1, 'F': 0 };

  const semesters = {};
  grades.forEach(g => {
    const key = `${g.semester || '1'}/${g.academic_year || ''}`;
    if (!semesters[key]) semesters[key] = { semester: g.semester || '1', year: g.academic_year || '', grades: [] };
    semesters[key].grades.push(g);
  });
  const semKeys = Object.keys(semesters).sort();
  let totalCreditsAll = 0, totalPointsAll = 0;
  let tableHTML = '';
  semKeys.forEach(key => {
    const sem = semesters[key];
    tableHTML += `<tr style="background:#e8f4fd"><td colspan="4" style="padding:4px 8px;font-weight:bold;text-align:center;font-size:12px;border:1px solid #999">ภาคการศึกษาที่ ${sem.semester} ปีการศึกษา ${sem.year}</td></tr>`;
    let semCredits = 0, semPoints = 0;
    sem.grades.forEach(g => {
      const cr = Number(g.credits) || 0; const gv = gradeMap[g.grade];
      tableHTML += `<tr><td style="padding:3px 8px;font-size:11px;border:1px solid #999;font-family:monospace">${g.subject_code || ''}</td><td style="padding:3px 8px;font-size:11px;border:1px solid #999">${g.subject_name || ''}</td><td style="padding:3px 8px;font-size:11px;text-align:center;border:1px solid #999">${g.credits || ''}</td><td style="padding:3px 8px;font-size:11px;text-align:center;font-weight:bold;border:1px solid #999">${g.grade || ''}</td></tr>`;
      if (gv !== undefined) { semPoints += gv * cr; semCredits += cr; }
    });
    const semGpa = semCredits ? (semPoints / semCredits).toFixed(2) : 'N/A';
    totalCreditsAll += semCredits; totalPointsAll += semPoints;
    tableHTML += `<tr style="background:#f9fafb"><td colspan="2" style="padding:3px 8px;font-size:11px;text-align:right;font-weight:600;border:1px solid #999">จำนวนหน่วยกิตรวม: ${semCredits}</td><td colspan="2" style="padding:3px 8px;font-size:11px;text-align:right;font-weight:600;border:1px solid #999">คะแนนเฉลี่ย: ${semGpa}</td></tr>`;
  });
  const gpax = totalCreditsAll ? (totalPointsAll / totalCreditsAll).toFixed(2) : 'N/A';

  let logoBase64 = '';
  try { const resp = await fetch('https://cdn.jsdelivr.net/gh/JOB-BCNB-P/LOGO/Logo%20Thai.png'); if (resp.ok) { const blob = await resp.blob(); logoBase64 = await new Promise(r => { const fr = new FileReader(); fr.onload = () => r(fr.result); fr.readAsDataURL(blob); }); } } catch (e) { }

  const studentProgram = stu.program || 'หลักสูตรพยาบาลศาสตรบัณฑิต';
  const studentLevel = stu.level || 'ปริญญาตรี';

  // Build watermark rows that repeat across the entire page (works on every PDF page)
  const watermarkText = 'ฉบับสำเนาไม่ใช่ตัวจริง';
  let watermarkRows = '';
  for (let i = 0; i < 8; i++) {
    watermarkRows += `<div class="wm-row">${Array(3).fill(`<span class="wm-text">${watermarkText}</span>`).join('')}</div>`;
  }

  const htmlContent = `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>ใบรายงานผลการเรียน - ${stu.name || ''}</title>
    <style>
      @import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;600;700&display=swap');
      *{font-family:'Sarabun',sans-serif;margin:0;padding:0;box-sizing:border-box}
      html,body{position:relative}
      body{padding:20px 40px;font-size:12px;color:#333;min-height:100vh}
      @media print{body{padding:15px 30px}@page{size:A4;margin:15mm 20mm}}
      table{width:100%;border-collapse:collapse}
      /* Watermark layer - fixed across all PDF pages */
      .watermark-layer{position:fixed;top:0;left:0;width:100%;height:100%;pointer-events:none;z-index:9999;overflow:hidden;display:flex;flex-direction:column;justify-content:space-around;transform:rotate(-30deg);transform-origin:center center;}
      .wm-row{display:flex;justify-content:space-around;width:140%;margin-left:-20%;}
      .wm-text{font-size:42px;font-weight:800;color:rgba(220,38,38,0.16);white-space:nowrap;letter-spacing:4px;}
      .content{position:relative;z-index:1;}
      -webkit-print-color-adjust:exact;
      .watermark-layer{-webkit-print-color-adjust:exact;print-color-adjust:exact;}
    </style></head><body>
    <div class="watermark-layer" aria-hidden="true">${watermarkRows}</div>
    <div class="content">
    <div style="text-align:center;margin-bottom:12px">${logoBase64 ? `<img src="${logoBase64}" style="width:55px;height:auto;margin-bottom:4px">` : ''}<div style="font-weight:700;font-size:14px">${APP.config.college_name}</div><div style="font-size:12px;color:#555">ใบรายงานผลการเรียนนักศึกษารายภาคการศึกษา</div><div style="font-size:11px;color:#555">${studentProgram} ระดับ ${studentLevel}</div>${(filterYear || filterSem) ? `<div style="font-size:11px;color:#1e3a8a;font-weight:600;margin-top:2px">เฉพาะ${filterSem ? ' ภาคการศึกษาที่ ' + filterSem : ''}${filterYear ? ' ปีการศึกษา ' + filterYear : ''}</div>` : ''}<div style="font-size:11px;color:#dc2626;font-weight:700;margin-top:4px;letter-spacing:2px">— ฉบับสำเนาไม่ใช่ตัวจริง —</div></div>
    <div style="display:flex;justify-content:space-between;margin-bottom:10px;font-size:11px"><div>รหัสนักศึกษา: <strong>${stu.student_id || ''}</strong></div><div>ชื่อ-สกุล: <strong>${stu.name || ''}</strong></div></div>
    <table><thead><tr style="background:#e8f4fd"><th style="padding:5px 8px;text-align:left;font-size:11px;border:1px solid #999;width:22%">รหัสวิชา</th><th style="padding:5px 8px;text-align:left;font-size:11px;border:1px solid #999;width:48%">รายวิชา</th><th style="padding:5px 8px;text-align:center;font-size:11px;border:1px solid #999;width:15%">หน่วยกิต</th><th style="padding:5px 8px;text-align:center;font-size:11px;border:1px solid #999;width:15%">ระดับคะแนน</th></tr></thead><tbody>${tableHTML}</tbody></table>
    <div style="margin-top:10px;border:1px solid #999;padding:6px 10px;font-size:11px"><div style="display:flex;justify-content:space-between"><span>รวมหน่วยกิตตลอดปีการศึกษา: <strong>${totalCreditsAll}</strong></span><span>คะแนนเฉลี่ยตลอดปีการศึกษา: <strong>${gpax}</strong></span></div><div style="display:flex;justify-content:space-between;margin-top:3px"><span>รวมหน่วยกิตสะสมตลอดหลักสูตร: <strong>${totalCreditsAll}</strong></span><span>คะแนนเฉลี่ยสะสมตลอดหลักสูตร: <strong>${gpax}</strong></span></div></div>
    <div style="margin-top:10px;font-size:10px;color:#666"><div style="font-weight:600;margin-bottom:3px">หมายเหตุ</div><div>A : ดีเยี่ยม &nbsp; B+ : ดีมาก &nbsp; B : ดี &nbsp; C+ : ค่อนข้างดี &nbsp; C : พอใช้ &nbsp; D+ : อ่อน &nbsp; D : อ่อนมาก &nbsp; F : ตก</div></div>
    </div>
    <script>window.onload=function(){window.print()}<\/script></body></html>`;
  const w = window.open('', '_blank');
  if (w) { w.document.write(htmlContent); w.document.close(); } else { showToast('กรุณาอนุญาต Popup เพื่อดาวน์โหลด PDF', 'error'); }
}

// ======================== ENG RESULTS ========================
function engResultsPage() {
  const isAdmin = APP.currentRole === 'admin' || APP.currentRole === 'academic';
  const isExecutive = APP.currentRole === 'executive';
  const canEdit = isAdmin || APP.currentRole === 'teacher' || APP.currentRole === 'classTeacher';
  const isStudent = APP.currentRole === 'student';
  const canFilterByAdvisor = isAdmin || isExecutive;
  let allEng = getDataByType('eng_result');

  // Academic year picker
  const allEngYears = [...new Set(allEng.map(e => e.academic_year).filter(Boolean))].sort().reverse();
  if (!allEngYears.length) allEngYears.push('2568');
  const selectedEngYear = APP.filters._engYear || '';
  let yearPickerHtml = `<div class="bg-white rounded-2xl p-4 border border-blue-100 mb-4">
    <div class="flex items-center gap-3">
      <label class="text-sm font-medium text-gray-700">ปีการศึกษา:</label>
      <select onchange="APP.filters._engYear=this.value;APP.pagination.page=1;renderCurrentPage()" class="border border-gray-200 rounded-xl px-3 py-2 text-sm">
        <option value="">-- ทุกปีการศึกษา --</option>
        ${allEngYears.map(y => '<option value="' + y + '"' + (selectedEngYear === y ? ' selected' : '') + '>' + y + '</option>').join('')}
      </select>
      ${selectedEngYear ? '<span class="text-xs text-gray-500">แสดงข้อมูลปีการศึกษา ' + selectedEngYear + '</span>' : ''}
    </div>
  </div>`;

  // Filter eng results by academic year
  if (selectedEngYear) {
    allEng = allEng.filter(e => (e.academic_year || '') === selectedEngYear);
  }

  // Build student selector for non-student roles
  let studentSelector = '';
  let selectedStudentName = APP.filters._engStudent || '';

  if (!isStudent) {
    let studentList = getDataByType('student');
    if (APP.currentRole === 'classTeacher') {
      const yr = APP.currentUser.responsible_year || '1';
      studentList = studentList.filter(s => norm(s.year_level) === norm(yr));
    }
    if (APP.currentRole === 'teacher') {
      studentList = studentList.filter(s => s.advisor === APP.currentUser.name);
    }

    // Year level + Advisor filter for admin/academic/executive
    let advisorSelector = '';
    if (canFilterByAdvisor) {
      // Year level filter
      const selectedEngYrLevel = APP.filters._engYearLevel || '';
      if (selectedEngYrLevel) {
        studentList = studentList.filter(s => norm(s.year_level) === selectedEngYrLevel);
      }
      const yearLevelSelector = `<div class="bg-white rounded-2xl p-4 border border-blue-100 mb-4">
        <label class="block text-sm font-medium text-gray-700 mb-2"><i data-lucide="layers" class="w-4 h-4 inline mr-1"></i>กรองตามชั้นปี</label>
        <div class="flex flex-wrap gap-2">
          <button onclick="APP.filters._engYearLevel='';APP.filters._engStudent='';APP.filters._engSearch='';APP.pagination.page=1;renderCurrentPage()" class="px-4 py-2 rounded-xl text-sm font-medium ${selectedEngYrLevel === '' ? 'bg-primary text-white' : 'bg-gray-100 text-gray-600 hover:bg-gray-200'}">ทุกชั้นปี</button>
          ${['1', '2', '3', '4'].map(yr => `<button onclick="APP.filters._engYearLevel='${yr}';APP.filters._engStudent='';APP.filters._engSearch='';APP.pagination.page=1;renderCurrentPage()" class="px-4 py-2 rounded-xl text-sm font-medium ${selectedEngYrLevel === yr ? 'bg-primary text-white' : 'bg-gray-100 text-gray-600 hover:bg-gray-200'}">ชั้นปี ${yr}</button>`).join('')}
        </div>
        ${selectedEngYrLevel ? `<p class="text-xs text-gray-500 mt-2"><i data-lucide="info" class="w-3 h-3 inline mr-1"></i>แสดงเฉพาะนักศึกษาชั้นปีที่ ${selectedEngYrLevel} (${studentList.length} คน)</p>` : ''}
      </div>`;

      const allAdvisors = [...new Set(studentList.map(s => s.advisor).filter(Boolean))].sort();
      const selectedAdvisor = APP.filters._engAdvisor || '';
      if (selectedAdvisor) {
        studentList = studentList.filter(s => (s.advisor || '') === selectedAdvisor);
      }
      advisorSelector = `${yearLevelSelector}<div class="bg-white rounded-2xl p-4 border border-blue-100 mb-4">
        <label class="block text-sm font-medium text-gray-700 mb-2"><i data-lucide="user-check" class="w-4 h-4 inline mr-1"></i>กรองตามอาจารย์ที่ปรึกษา</label>
        <select onchange="APP.filters._engAdvisor=this.value;APP.filters._engStudent='';APP.filters._engSearch='';APP.pagination.page=1;renderCurrentPage()" class="w-full border border-gray-200 rounded-xl px-4 py-2.5 text-sm">
          <option value="">-- แสดงนักศึกษาทั้งหมด --</option>
          ${allAdvisors.map(a => `<option value="${a}" ${selectedAdvisor === a ? 'selected' : ''}>${a}</option>`).join('')}
        </select>
        ${selectedAdvisor ? `<p class="text-xs text-gray-500 mt-2"><i data-lucide="info" class="w-3 h-3 inline mr-1"></i>แสดงเฉพาะนักศึกษาในความดูแลของ ${selectedAdvisor} (${studentList.length} คน)</p>` : ''}
      </div>`;
    }

    const searchVal = APP.filters._engSearch || '';
    let filteredList = studentList;
    if (searchVal) {
      const q = searchVal.toLowerCase();
      filteredList = studentList.filter(s => (s.name || '').toLowerCase().includes(q) || (s.student_id || '').toLowerCase().includes(q));
    }
    studentSelector = `${advisorSelector}<div class="bg-white rounded-2xl p-4 border border-blue-100 mb-4">
      <label class="block text-sm font-medium text-gray-700 mb-2"><i data-lucide="user-search" class="w-4 h-4 inline mr-1"></i>เลือกนักศึกษา</label>
      <div class="flex gap-2 mb-2">
        <div class="flex-1 relative"><i data-lucide="search" class="absolute left-3 top-2.5 w-4 h-4 text-gray-400"></i><input type="text" placeholder="พิมพ์ค้นหาชื่อหรือรหัส..." value="${searchVal}" class="w-full pl-10 pr-4 py-2 border border-gray-200 rounded-xl text-sm" oninput="clearTimeout(window._engSearchTimer);window._engSearchTimer=setTimeout(()=>{APP.filters._engSearch=this.value;APP.filters._engStudent='';APP.pagination.page=1;renderCurrentPage()},300)"></div>
      </div>
      <select onchange="APP.filters._engStudent=this.value;APP.pagination.page=1;renderCurrentPage()" class="w-full border border-gray-200 rounded-xl px-4 py-2.5 text-sm">
        <option value="">-- กรุณาเลือกนักศึกษา (${filteredList.length} คน) --</option>
        ${filteredList.map(s => `<option value="${s.student_id || s.name}" ${selectedStudentName === (s.student_id || s.name) ? 'selected' : ''}>${s.student_id || ''} — ${s.name || ''}</option>`).join('')}
      </select>
    </div>`;
  }

  // Filter data — match by student_id
  let data;
  if (isStudent && APP.currentUser.data) {
    const sid = norm(APP.currentUser.data.student_id);
    data = allEng.filter(e => norm(e.student_id) === sid);
  } else if (selectedStudentName) {
    const sid = norm(selectedStudentName);
    data = allEng.filter(e => norm(e.student_id) === sid);
  } else {
    data = [];
  }
  data = applyFilters(data);
  const total = data.length; const paged = paginate(data);

  // Show prompt if no student selected
  let noSelectionMsg = '';
  if (!isStudent && !selectedStudentName) {
    noSelectionMsg = `<div class="bg-white rounded-2xl border border-blue-100 p-8 text-center text-gray-400"><i data-lucide="user-search" class="w-10 h-10 mx-auto mb-3 text-gray-300"></i><p>กรุณาเลือกนักศึกษาเพื่อดูผลสอบภาษาอังกฤษ</p></div>`;
  }

  // Build summary stats (pass/fail counts only)
  let summaryTableHtml = '';
  if (!isStudent) {
    // Get students list for summary
    let summaryStudents = getDataByType('student');
    if (APP.currentRole === 'classTeacher') {
      const yr = APP.currentUser.responsible_year || '1';
      summaryStudents = summaryStudents.filter(s => norm(s.year_level) === norm(yr));
    }
    if (APP.currentRole === 'teacher') {
      summaryStudents = summaryStudents.filter(s => s.advisor === APP.currentUser.name);
    }
    // Apply advisor filter if set
    const selectedAdvisor = APP.filters._engAdvisor || '';
    if (canFilterByAdvisor && selectedAdvisor) {
      summaryStudents = summaryStudents.filter(s => (s.advisor || '') === selectedAdvisor);
    }

    // Count unique students who passed / not passed
    const passedIds = new Set(allEng.filter(e => e.eng_status === 'ผ่าน').map(e => norm(e.student_id)));
    const passedCount = summaryStudents.filter(s => passedIds.has(norm(s.student_id))).length;
    const notPassedCount = summaryStudents.length - passedCount;

    // Breakdown for tooltip — classTeacher: by room, others: by year level
    let tooltipHtml = '';
    if (APP.currentRole === 'classTeacher') {
      const roomAstu = summaryStudents.filter(s => norm(s.room).toUpperCase() === 'A');
      const roomBstu = summaryStudents.filter(s => norm(s.room).toUpperCase() === 'B');
      const roomApass = roomAstu.filter(s => passedIds.has(s.student_id)).length;
      const roomBpass = roomBstu.filter(s => passedIds.has(s.student_id)).length;
      const roomAfail = roomAstu.length - roomApass;
      const roomBfail = roomBstu.length - roomBpass;
      const tooltipRows = `
        <div class="flex items-center justify-between gap-4 py-1 border-b border-gray-100">
          <span class="text-xs text-gray-600">ห้อง A</span>
          <div class="flex gap-3"><span class="text-xs font-semibold text-green-600">ผ่าน ${roomApass}</span><span class="text-xs font-semibold text-red-500">ไม่ผ่าน ${roomAfail}</span></div>
        </div>
        <div class="flex items-center justify-between gap-4 py-1">
          <span class="text-xs text-gray-600">ห้อง B</span>
          <div class="flex gap-3"><span class="text-xs font-semibold text-green-600">ผ่าน ${roomBpass}</span><span class="text-xs font-semibold text-red-500">ไม่ผ่าน ${roomBfail}</span></div>
        </div>`;
      tooltipHtml = `<div class="absolute z-20 bottom-full left-1/2 -translate-x-1/2 mb-2 w-48 bg-white border border-blue-100 rounded-xl shadow-xl p-3 hidden group-hover:block pointer-events-none">
        <p class="text-xs font-bold text-gray-700 mb-2">สรุปรายห้อง</p>
        ${tooltipRows}
        <div class="absolute bottom-[-6px] left-1/2 -translate-x-1/2 w-3 h-3 bg-white border-r border-b border-blue-100 rotate-45"></div>
      </div>`;
    } else {
      const yearBreakdown = [1, 2, 3, 4].map(yr => {
        const yrStudents = summaryStudents.filter(s => norm(s.year_level) === String(yr));
        if (!yrStudents.length) return null;
        const yrPassed = yrStudents.filter(s => passedIds.has(s.student_id)).length;
        const yrFailed = yrStudents.length - yrPassed;
        return { yr, total: yrStudents.length, passed: yrPassed, failed: yrFailed };
      }).filter(Boolean);
      if (yearBreakdown.length) {
        const tooltipRows = yearBreakdown.map(y =>
          `<div class="flex items-center justify-between gap-4 py-1 border-b border-gray-100 last:border-0">
            <span class="text-xs text-gray-600">ชั้นปี ${y.yr}</span>
            <div class="flex gap-3"><span class="text-xs font-semibold text-green-600">ผ่าน ${y.passed}</span><span class="text-xs font-semibold text-red-500">ไม่ผ่าน ${y.failed}</span></div>
          </div>`).join('');
        tooltipHtml = `<div class="absolute z-20 bottom-full left-1/2 -translate-x-1/2 mb-2 w-52 bg-white border border-blue-100 rounded-xl shadow-xl p-3 hidden group-hover:block pointer-events-none">
          <p class="text-xs font-bold text-gray-700 mb-2">สรุปรายชั้นปี</p>
          ${tooltipRows}
          <div class="absolute bottom-[-6px] left-1/2 -translate-x-1/2 w-3 h-3 bg-white border-r border-b border-blue-100 rotate-45"></div>
        </div>`;
      }
    }

    summaryTableHtml = `
    <div class="bg-white rounded-2xl border border-blue-100 p-5 mb-4">
      <h3 class="font-bold text-gray-800 mb-3 flex items-center gap-2"><i data-lucide="bar-chart-3" class="w-5 h-5 text-primary"></i>สรุปผลสอบภาษาอังกฤษ</h3>
      <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
        <div class="relative group cursor-default">
          ${statCard('check-circle', 'สอบผ่าน', passedCount, 'คน', 'bg-green-500')}
          ${tooltipHtml}
        </div>
        <div class="relative group cursor-default">
          ${statCard('x-circle', 'ยังไม่ผ่าน', notPassedCount, 'คน', 'bg-red-500')}
          ${tooltipHtml}
        </div>
      </div>
    </div>`;
  }

  return `<div class="flex flex-wrap items-center justify-between gap-3 mb-4">
    <h2 class="text-xl font-bold text-gray-800"><i data-lucide="languages" class="w-6 h-6 inline mr-2"></i>ผลสอบภาษาอังกฤษ</h2>
    ${isAdmin ? `<div class="flex gap-2"><button onclick="showAddEngModal()" class="flex items-center gap-2 px-4 py-2 bg-primary text-white rounded-xl hover:bg-primaryDark text-sm"><i data-lucide="plus" class="w-4 h-4"></i>เพิ่มผลสอบ</button>${csvUploadBtn('eng_result', 'student_id,eng_score,eng_type,eng_attempt,eng_date,eng_status,academic_year')}</div>` : ''}
  </div>
  ${summaryTableHtml}
  ${yearPickerHtml}
  ${studentSelector}
  ${noSelectionMsg || `<div class="bg-white rounded-2xl border border-blue-100 p-4 mb-4">
    <h3 class="font-bold text-gray-800 mb-3 flex items-center gap-2"><i data-lucide="file-text" class="w-5 h-5 text-primary"></i>รายละเอียดผลสอบรายบุคคล</h3>
    <div class="grid grid-cols-1 sm:grid-cols-2 gap-4 mb-4">
      ${statCard('check-circle', 'ผ่าน', data.filter(e => e.eng_status === 'ผ่าน').length, 'ครั้ง', 'bg-green-500')}
      ${statCard('x-circle', 'ไม่ผ่าน', data.filter(e => e.eng_status === 'ไม่ผ่าน').length, 'ครั้ง', 'bg-red-500')}
    </div>
  </div>
  <div class="bg-white rounded-2xl border border-blue-100 overflow-hidden">
    <div class="overflow-x-auto"><table class="w-full text-sm">
      <thead><tr class="bg-surface text-left">
        <th class="px-4 py-3 font-semibold">รูปแบบ</th>
        <th class="px-4 py-3 font-semibold">Listening</th>
        <th class="px-4 py-3 font-semibold">Grammar</th>
        <th class="px-4 py-3 font-semibold">Reading</th>
        <th class="px-4 py-3 font-semibold">คะแนนรวม</th>
        <th class="px-4 py-3 font-semibold">ระดับ</th>
        <th class="px-4 py-3 font-semibold">สอบครั้งที่</th>
        <th class="px-4 py-3 font-semibold">วันที่สอบ</th>
        <th class="px-4 py-3 font-semibold">ปีการศึกษา</th>
        <th class="px-4 py-3 font-semibold">สถานะ</th>
        ${isAdmin ? '<th class="px-4 py-3"></th>' : ''}
      </tr></thead>
      <tbody>${paged.length ? paged.map(e => {
    const isSbch = e.eng_type === 'สบช.';
    const isAbsent = e.eng_status === 'ไม่เข้าสอบ';
    const level = isAbsent ? '' : (e.eng_level || (isSbch ? getEngLevel(Number(e.eng_score) || 0) : ''));
    return `<tr class="border-t hover:bg-gray-50">
          <td class="px-4 py-3 font-medium">${e.eng_type || ''}</td>
          <td class="px-4 py-3 text-center">${isSbch ? (e.eng_listening || '-') : '-'}</td>
          <td class="px-4 py-3 text-center">${isSbch ? (e.eng_grammar || '-') : '-'}</td>
          <td class="px-4 py-3 text-center">${isSbch ? (e.eng_reading || '-') : '-'}</td>
          <td class="px-4 py-3 text-center font-semibold">${e.eng_score || ''}</td>
          <td class="px-4 py-3"><span class="text-xs ${isSbch ? 'font-medium text-blue-700' : 'text-gray-400'}">${level || '-'}</span></td>
          <td class="px-4 py-3 text-center">${e.eng_attempt || ''}</td>
          <td class="px-4 py-3">${formatDate(e.eng_date)}</td>
          <td class="px-4 py-3">${e.academic_year || ''}</td>
          <td class="px-4 py-3"><span class="px-2 py-1 rounded-full text-xs ${e.eng_status === 'ผ่าน' ? 'bg-green-100 text-green-700' : e.eng_status === 'ไม่เข้าสอบ' ? 'bg-orange-100 text-orange-700' : 'bg-red-100 text-red-700'}">${e.eng_status || ''}</span></td>
          ${isAdmin ? `<td class="px-4 py-3"><div class="flex gap-1"><button onclick="showEditEngModal('${e.__backendId}')" class="text-blue-400 hover:text-blue-600" title="แก้ไข"><i data-lucide="pencil" class="w-4 h-4"></i></button><button onclick="deleteRecord('${e.__backendId}')" class="text-red-400 hover:text-red-600" title="ลบ"><i data-lucide="trash-2" class="w-4 h-4"></i></button></div></td>` : ''}
        </tr>`;
  }).join('') : '<tr><td colspan="11" class="px-4 py-8 text-center text-gray-400">ไม่มีข้อมูล</td></tr>'}</tbody>
    </table></div>
  </div>
  ${paginationHTML(total, APP.pagination.perPage, APP.pagination.page, 'changePage')}`}`;
}

// ---- Eng helpers ----
function getEngLevel(score) {
  const s = Number(score) || 0;
  if (s >= 91) return 'Proficiency';
  if (s >= 81) return 'Advanced';
  if (s >= 61) return 'Upper Intermediate';
  if (s >= 41) return 'Intermediate';
  if (s >= 21) return 'Elementary';
  return 'Beginner';
}
function updateEngTypeForm(prefix) {
  const type = document.getElementById(prefix + 'EngType').value;
  const sbch = document.getElementById(prefix + 'EngSbch');
  const other = document.getElementById(prefix + 'EngOther');
  if (sbch) sbch.style.display = type === 'สบช.' ? 'block' : 'none';
  if (other) other.style.display = type === 'สบช.' ? 'none' : 'block';
}
function calcEngSbchTotal(prefix) {
  const l = Number(document.getElementById(prefix + 'EngL').value) || 0;
  const g = Number(document.getElementById(prefix + 'EngG').value) || 0;
  const r = Number(document.getElementById(prefix + 'EngR').value) || 0;
  const total = l + g + r;
  const totalEl = document.getElementById(prefix + 'EngTotal');
  if (totalEl && total > 0) totalEl.value = total;
  updateEngSbchLevelStatus(prefix);
}
function calcEngSbchFromTotal(prefix) {
  updateEngSbchLevelStatus(prefix);
}
function updateEngSbchLevelStatus(prefix) {
  const totalEl = document.getElementById(prefix + 'EngTotal');
  const levelEl = document.getElementById(prefix + 'EngLevel');
  const statusEl = document.getElementById(prefix + 'EngStatus');
  const total = Number(totalEl ? totalEl.value : 0) || 0;
  if (levelEl) levelEl.value = total > 0 ? getEngLevel(total) : '';
  if (statusEl) {
    if (total > 0) {
      statusEl.textContent = total >= 41 ? 'ผ่าน' : 'ไม่ผ่าน';
      statusEl.className = 'inline-block px-3 py-1 rounded-full text-sm font-semibold ' + (total >= 41 ? 'bg-green-100 text-green-700' : 'bg-red-100 text-red-700');
    } else {
      statusEl.textContent = '-';
      statusEl.className = 'text-sm text-gray-400';
    }
  }
}
const ENG_TYPES = ['สบช.', 'TOEIC', 'CU-TEP', 'IELTS', 'TOEIC-ITP', 'TOEFL', 'TU-GET'];
function toggleEngAbsent(prefix) {
  const checked = document.getElementById(prefix + 'EngAbsent').checked;
  const sbch = document.getElementById(prefix + 'EngSbch');
  const other = document.getElementById(prefix + 'EngOther');
  const typeEl = document.getElementById(prefix + 'EngType');
  if (checked) {
    if (sbch) sbch.style.display = 'none';
    if (other) other.style.display = 'none';
    if (typeEl) typeEl.disabled = true;
  } else {
    if (typeEl) { typeEl.disabled = false; updateEngTypeForm(prefix); }
  }
}
function engTypeOptions(selected) {
  return `<option value="">-- เลือกรูปแบบ --</option>` + ENG_TYPES.map(t => `<option value="${t}" ${selected === t ? 'selected' : ''}>${t}</option>`).join('');
}
function engYearOptions(selected) {
  const existing = [...new Set(getDataByType('eng_result').map(e => e.academic_year).filter(Boolean))].sort().reverse();
  const currentBE = String(new Date().getFullYear() + 543);
  const years = [...new Set([currentBE, ...existing])].sort().reverse();
  return years.map(y => `<option value="${y}" ${y === (selected || currentBE) ? 'selected' : ''}>${y}</option>`).join('');
}

function showAddEngModal() {
  showModal('เพิ่มผลสอบภาษาอังกฤษ', `
    <form id="addEngForm" class="space-y-3">
      <div><label class="block text-xs text-gray-600 mb-1">นักศึกษา * (พิมพ์รหัสหรือเลือก)</label><input list="addEngStudentList" name="student_id" required class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="พิมพ์รหัสนักศึกษา...">${studentDatalistHTML('addEngStudentList')}</div>
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-xs text-gray-600 mb-1">รูปแบบการสอบ *</label>
          <select id="addEngType" onchange="updateEngTypeForm('add')" class="w-full border rounded-xl px-3 py-2 text-sm">${engTypeOptions('')}</select>
        </div>
        <div><label class="block text-xs text-gray-600 mb-1">สอบครั้งที่</label><input id="addEngAttempt" type="number" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="เช่น 1, 2"></div>
        <div><label class="block text-xs text-gray-600 mb-1">วันที่สอบ <span class="text-gray-400">(วว/ดด/ปปปป พ.ศ. หรือ ค.ศ.)</span></label><input id="addEngDate" type="text" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="เช่น 05/04/2568 หรือ 05/04/2025"></div>
        <div><label class="block text-xs text-gray-600 mb-1">ปีการศึกษา</label><input id="addEngYear" type="text" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="เช่น 2568"></div>
      </div>
      <!-- สบช. fields -->
      <div id="addEngSbch" style="display:none">
        <p class="text-xs font-semibold text-blue-700 mb-2">คะแนนรายทักษะ (สบช.) <span class="font-normal text-gray-400">— ไม่บังคับ หากไม่มีกรอกแค่คะแนนรวม</span></p>
        <div class="grid grid-cols-3 gap-2 mb-2">
          <div><label class="block text-xs text-gray-600 mb-1">Listening</label><input id="addEngL" type="number" min="0" max="100" oninput="calcEngSbchTotal('add')" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="0-100"></div>
          <div><label class="block text-xs text-gray-600 mb-1">Grammar</label><input id="addEngG" type="number" min="0" max="100" oninput="calcEngSbchTotal('add')" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="0-100"></div>
          <div><label class="block text-xs text-gray-600 mb-1">Reading</label><input id="addEngR" type="number" min="0" max="100" oninput="calcEngSbchTotal('add')" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="0-100"></div>
        </div>
        <div class="grid grid-cols-2 gap-2">
          <div><label class="block text-xs text-gray-600 mb-1">คะแนนรวม</label><input id="addEngTotal" type="number" min="0" max="300" oninput="calcEngSbchFromTotal('add')" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="กรอกคะแนนรวม"></div>
          <div><label class="block text-xs text-gray-600 mb-1">ระดับผลการสอบ (อัตโนมัติ)</label><input id="addEngLevel" readonly class="w-full border rounded-xl px-3 py-2 text-sm bg-gray-50"></div>
        </div>
        <div class="mt-2"><label class="block text-xs text-gray-600 mb-1">สถานะ (อัตโนมัติ)</label><div class="border rounded-xl px-3 py-2 bg-gray-50 min-h-[36px] flex items-center"><span id="addEngStatus" class="text-sm text-gray-400">-</span></div></div>
      </div>
      <!-- Other type fields -->
      <div id="addEngOther" style="display:none">
        <div class="grid grid-cols-2 gap-3">
          <div><label class="block text-xs text-gray-600 mb-1">คะแนนรวม</label><input id="addEngOtherScore" type="number" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
          <div><label class="block text-xs text-gray-600 mb-1">สถานะ *</label><select id="addEngOtherStatus" class="w-full border rounded-xl px-3 py-2 text-sm"><option value="ผ่าน">ผ่าน</option><option value="ไม่ผ่าน">ไม่ผ่าน</option></select></div>
        </div>
      </div>
      <!-- ไม่เข้าสอบ -->
      <label class="flex items-center gap-2 cursor-pointer select-none">
        <input type="checkbox" id="addEngAbsent" onchange="toggleEngAbsent('add')" class="w-4 h-4 rounded border-gray-300 text-red-500 focus:ring-red-400">
        <span class="text-sm text-red-600 font-medium">ไม่เข้าสอบ</span>
      </label>
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึก</button>
    </form>
  `);
  document.getElementById('addEngForm').onsubmit = async (ev) => {
    ev.preventDefault();
    await withLoading(ev.target, async () => {
      const absent = document.getElementById('addEngAbsent').checked;
      const engType = document.getElementById('addEngType').value;
      if (!absent && !engType) { showToast('กรุณาเลือกรูปแบบการสอบ', 'error'); return; }
      const studentId = ev.target.querySelector('[name="student_id"]').value;
      const attempt = document.getElementById('addEngAttempt').value;
      const date = normalizeDateInput(document.getElementById('addEngDate').value);
      const year = document.getElementById('addEngYear').value;
      const obj = { type: 'eng_result', created_at: new Date().toISOString(), student_id: studentId, eng_type: engType, eng_attempt: Number(attempt) || '', eng_date: date, academic_year: year };
      if (absent) {
        obj.eng_status = 'ไม่เข้าสอบ';
        obj.eng_score = '';
      } else if (engType === 'สบช.') {
        const l = Number(document.getElementById('addEngL').value) || 0;
        const g = Number(document.getElementById('addEngG').value) || 0;
        const r = Number(document.getElementById('addEngR').value) || 0;
        const total = Number(document.getElementById('addEngTotal').value) || 0;
        if (total === 0) { showToast('กรุณากรอกคะแนนรวม', 'error'); return; }
        if (l) obj.eng_listening = l;
        if (g) obj.eng_grammar = g;
        if (r) obj.eng_reading = r;
        obj.eng_score = total; obj.eng_level = getEngLevel(total);
        obj.eng_status = total >= 41 ? 'ผ่าน' : 'ไม่ผ่าน';
      } else {
        obj.eng_score = Number(document.getElementById('addEngOtherScore').value) || '';
        obj.eng_status = document.getElementById('addEngOtherStatus').value;
      }
      const res = await GSheetDB.create(obj);
      if (res.isOk) { showToast('เพิ่มผลสอบสำเร็จ'); closeModal(); } else showToast('เกิดข้อผิดพลาด', 'error');
    });
  };
}

// ======================== EVAL TEACHER (REMOVED) ========================
// Teacher evaluation feature has been removed from the system.
// The following Google Sheet tabs are no longer used: eval_form, evaluation
/* removed_eval_block_start
function renderEvalTeacherFiltered(data, byTeacher, teacherName) {
  const filtered = data.filter(e => (e.teacher_name || e.name) === teacherName);
  const selTeacher = byTeacher[teacherName];
  let html = '';
  if (selTeacher) {
    html += `<div class="bg-white rounded-2xl p-4 border border-blue-100 mb-4"><p class="font-bold">${teacherName}</p><p class="text-sm text-gray-500">${selTeacher.subject || ''}</p><div class="flex items-center gap-2 mt-2"><span class="text-2xl font-bold text-primary">${(selTeacher.total / selTeacher.count).toFixed(1)}</span><span class="text-gray-400">/5 (${selTeacher.count} ผลประเมิน)</span></div></div>`;
  }
  html += `<div class="bg-white rounded-2xl border border-blue-100 overflow-hidden"><div class="overflow-x-auto"><table class="w-full text-sm">
    <thead><tr class="bg-surface text-left"><th class="px-4 py-3 font-semibold">รายวิชา</th><th class="px-4 py-3 font-semibold">อาจารย์</th><th class="px-4 py-3 font-semibold">นักศึกษา</th><th class="px-4 py-3 font-semibold">คะแนน</th><th class="px-4 py-3 font-semibold">ภาค/ปี</th></tr></thead>
    <tbody>${filtered.length ? filtered.slice(0, 30).map(e => '<tr class="border-t hover:bg-gray-50"><td class="px-4 py-3">' + (e.subject_name || '') + '</td><td class="px-4 py-3">' + (e.teacher_name || e.name || '') + '</td><td class="px-4 py-3">' + (e.student_name || '') + '</td><td class="px-4 py-3 font-bold">' + (e.eval_score || '') + '/5</td><td class="px-4 py-3">' + (e.semester || '') + '/' + (e.academic_year || '') + '</td></tr>').join('') : '<tr><td colspan="5" class="px-4 py-8 text-center text-gray-400">ยังไม่มีผลประเมิน</td></tr>'}</tbody>
  </table></div></div>`;
  return html;
}

function evalTeacherPage() {
  const isAdmin = APP.currentRole === 'admin' || APP.currentRole === 'academic';
  const isStudent = APP.currentRole === 'student';
  const evalForms = getDataByType('eval_form');
  const evaluations = getDataByType('evaluation');

  if (isStudent) {
    // Student: see forms created by admin, fill in scores
    const myEvals = evaluations.filter(e => e.student_name === (APP.currentUser.data?.name || ''));
    const availableForms = evalForms.filter(f => {
      // Filter forms that student hasn't submitted yet
      return f.status === 'เปิด' && !myEvals.some(e => e.eval_form_id === f.__backendId);
    });
    const submittedForms = evalForms.filter(f => {
      return myEvals.some(e => e.eval_form_id === f.__backendId);
    });

    return `<h2 class="text-xl font-bold text-gray-800 mb-4"><i data-lucide="star" class="w-6 h-6 inline mr-2"></i>ประเมินอาจารย์ผู้สอน</h2>
    
    ${availableForms.length ? `<h3 class="font-bold mb-3 text-green-700 flex items-center gap-2"><i data-lucide="clipboard-list" class="w-5 h-5"></i>แบบประเมินที่ยังไม่ได้ทำ (${availableForms.length})</h3>
    <div class="grid grid-cols-1 md:grid-cols-2 gap-4 mb-6">${availableForms.map(f => `<div class="bg-white rounded-2xl p-5 border border-green-200 hover:shadow-md transition cursor-pointer" onclick="showStudentEvalForm('${f.__backendId}')">
      <div class="flex items-center justify-between mb-2">
        <span class="px-2 py-1 rounded-full text-xs bg-green-100 text-green-700">รอประเมิน</span>
        <span class="text-xs text-gray-400">${semLabel(f.semester)}/${f.academic_year || ''}</span>
      </div>
      <p class="font-bold text-gray-800">${f.subject_code ? f.subject_code + ' ' : ''}${f.subject_name || ''}</p>
      <p class="text-sm text-gray-500 mt-1">อาจารย์: ${f.teacher_name || ''}</p>
      <p class="text-xs text-gray-400 mt-2">${f.eval_items ? f.eval_items.split(',').length || 0 : 0} หัวข้อประเมิน</p>
    </div>`).join('')}</div>` : '<div class="bg-green-50 rounded-2xl p-6 text-center mb-6"><p class="text-green-600">ไม่มีแบบประเมินที่ต้องทำ</p></div>'}

    <h3 class="font-bold mb-3 text-gray-600">ประวัติที่ประเมินแล้ว (${submittedForms.length})</h3>
    <div class="bg-white rounded-2xl border border-blue-100 overflow-hidden">
      <div class="overflow-x-auto"><table class="w-full text-sm">
        <thead><tr class="bg-surface"><th class="px-4 py-3 text-left">รายวิชา</th><th class="px-4 py-3 text-left">อาจารย์</th><th class="px-4 py-3">คะแนนเฉลี่ย</th><th class="px-4 py-3">ภาค/ปี</th></tr></thead>
        <tbody>${myEvals.length ? myEvals.map(e => `<tr class="border-t"><td class="px-4 py-3">${e.subject_name || ''}</td><td class="px-4 py-3">${e.teacher_name || e.name || ''}</td><td class="px-4 py-3 text-center font-bold text-primary">${e.eval_score || ''}/5</td><td class="px-4 py-3 text-center">${semLabel(e.semester)}/${e.academic_year || ''}</td></tr>`).join('') : '<tr><td colspan="4" class="py-6 text-center text-gray-400">ยังไม่มีประวัติ</td></tr>'}</tbody>
      </table></div>
    </div>`;
  }

  // ===== Admin view: manage eval forms + see results =====
  if (isAdmin) {
    const data = applyFilters(evaluations);
    // Summary by teacher
    const byTeacher = {};
    data.forEach(e => {
      const tname = e.teacher_name || e.name || '';
      if (!tname) return;
      if (!byTeacher[tname]) byTeacher[tname] = { total: 0, count: 0, subject: e.subject_name };
      byTeacher[tname].total += Number(e.eval_score) || 0;
      byTeacher[tname].count++;
    });

    const selEvalTeacher = APP.filters._evalTeacher || '';
    const teacherFilter = `<select onchange="APP.filters._evalTeacher=this.value;renderCurrentPage()" class="border border-gray-200 rounded-xl px-3 py-2.5 text-sm"><option value="">-- เลือกอาจารย์ --</option>${[...new Set(data.map(e => e.teacher_name || e.name).filter(Boolean))].map(t => `<option ${t === selEvalTeacher ? 'selected' : ''}>${t}</option>`).join('')}</select>`;

    return `<h2 class="text-xl font-bold text-gray-800 mb-4"><i data-lucide="star" class="w-6 h-6 inline mr-2"></i>ระบบประเมินอาจารย์ผู้สอน</h2>
    
    <div class="bg-white rounded-2xl p-5 border border-blue-100 mb-6">
      <div class="flex items-center justify-between mb-4">
        <h3 class="font-bold">จัดการแบบประเมิน</h3>
        <button onclick="showCreateEvalFormModal()" class="flex items-center gap-2 px-4 py-2 bg-primary text-white rounded-xl hover:bg-primaryDark text-sm"><i data-lucide="plus" class="w-4 h-4"></i>สร้างแบบประเมิน</button>
      </div>
      <div class="overflow-x-auto"><table class="w-full text-sm">
        <thead><tr class="bg-surface text-left"><th class="px-4 py-3 font-semibold">รายวิชา</th><th class="px-4 py-3 font-semibold">อาจารย์</th><th class="px-4 py-3 font-semibold">หัวข้อประเมิน</th><th class="px-4 py-3 font-semibold">ภาค/ปี</th><th class="px-4 py-3 font-semibold">สถานะ</th><th class="px-4 py-3 font-semibold">ผู้ตอบ</th><th class="px-4 py-3"></th></tr></thead>
        <tbody>${evalForms.length ? evalForms.map(f => {
      const respCount = evaluations.filter(e => e.eval_form_id === f.__backendId).length;
      return `<tr class="border-t hover:bg-gray-50">
          <td class="px-4 py-3 font-medium">${f.subject_code ? f.subject_code + ' ' : ''}${f.subject_name || ''}</td>
          <td class="px-4 py-3">${f.teacher_name || ''}</td>
          <td class="px-4 py-3 text-xs">${(f.eval_items || '').split(',').filter(Boolean).join(', ')}</td>
          <td class="px-4 py-3">${semLabel(f.semester)}/${f.academic_year || ''}</td>
          <td class="px-4 py-3"><span class="px-2 py-1 rounded-full text-xs ${f.status === 'เปิด' ? 'bg-green-100 text-green-700' : 'bg-gray-100 text-gray-500'}">${f.status || 'เปิด'}</span></td>
          <td class="px-4 py-3 font-bold text-primary">${respCount}</td>
          <td class="px-4 py-3"><div class="flex gap-1"><button onclick="showEditEvalFormModal('${f.__backendId}')" class="text-blue-400 hover:text-blue-600" title="แก้ไข"><i data-lucide="pencil" class="w-4 h-4"></i></button><button onclick="deleteRecord('${f.__backendId}')" class="text-red-400 hover:text-red-600" title="ลบ"><i data-lucide="trash-2" class="w-4 h-4"></i></button></div></td>
        </tr>`}).join('') : '<tr><td colspan="7" class="px-4 py-8 text-center text-gray-400">ยังไม่มีแบบประเมิน</td></tr>'}</tbody>
      </table></div>
    </div>

    <h3 class="font-bold mb-3">ผลประเมินรวม</h3>
    <div class="flex flex-wrap gap-3 mb-4">${teacherFilter}</div>
    ${APP.filters._evalTeacher ? renderEvalTeacherFiltered(data, byTeacher, APP.filters._evalTeacher) : '<div class="bg-blue-50 rounded-2xl p-6 text-center text-blue-600"><i data-lucide="info" class="w-5 h-5 inline mr-1"></i>กรุณาเลือกอาจารย์เพื่อดูผลประเมิน</div>'}`;
  }

  // ===== Teacher view: see own results =====
  let data = evaluations.filter(e => (e.teacher_name || e.name) === APP.currentUser.name);
  data = applyFilters(data);
  let totalScore = 0, totalCount = 0;
  data.forEach(e => { totalScore += Number(e.eval_score) || 0; totalCount++ });
  const avgScore = totalCount ? (totalScore / totalCount).toFixed(1) : 'N/A';

  return `<h2 class="text-xl font-bold text-gray-800 mb-4"><i data-lucide="star" class="w-6 h-6 inline mr-2"></i>ผลประเมินของฉัน</h2>
  <div class="bg-gradient-to-r from-primary to-accent text-white rounded-2xl p-5 mb-4 flex items-center justify-between">
    <div><p class="text-sm opacity-90">คะแนนเฉลี่ยรวม</p><p class="text-3xl font-bold">${avgScore}/5</p></div>
    <div><p class="text-sm opacity-90">จำนวนผู้ประเมิน</p><p class="text-3xl font-bold">${totalCount} <span class="text-sm font-normal opacity-80">คน</span></p></div>
  </div>
  <div class="bg-white rounded-2xl border border-blue-100 overflow-hidden">
    <div class="overflow-x-auto"><table class="w-full text-sm">
      <thead><tr class="bg-surface text-left"><th class="px-4 py-3 font-semibold">รายวิชา</th><th class="px-4 py-3 font-semibold">นักศึกษา</th><th class="px-4 py-3 font-semibold">คะแนน</th><th class="px-4 py-3 font-semibold">ภาค/ปี</th></tr></thead>
      <tbody>${data.length ? data.map(e => `<tr class="border-t hover:bg-gray-50"><td class="px-4 py-3">${e.subject_name || ''}</td><td class="px-4 py-3">${e.student_name || ''}</td><td class="px-4 py-3 font-bold">${e.eval_score || ''}/5</td><td class="px-4 py-3">${semLabel(e.semester)}/${e.academic_year || ''}</td></tr>`).join('') : '<tr><td colspan="4" class="px-4 py-8 text-center text-gray-400">ยังไม่มีผลประเมิน</td></tr>'}</tbody>
    </table></div>
  </div>`;
}

// Admin: Create eval form
function showCreateEvalFormModal() {
  const subjects = getDataByType('subject');
  showModal('สร้างแบบประเมินอาจารย์', `
    <form id="createEvalFormForm" class="space-y-3">
      <div><label class="block text-xs text-gray-600 mb-1">รายวิชา *</label>
        <select name="subject_name" required class="w-full border rounded-xl px-3 py-2 text-sm" onchange="onEvalFormSubjectChange(this)">
          <option value="">เลือกรายวิชา</option>
          ${subjects.map(s => `<option value="${s.subject_name}" data-code="${s.subject_code || ''}" data-coord="${s.coordinator || ''}">${s.subject_code ? s.subject_code + ' ' : ''}${s.subject_name}</option>`).join('')}
        </select>
      </div>
      <div><label class="block text-xs text-gray-600 mb-1">รหัสวิชา</label><input name="subject_code" id="evalFormSubCode" readonly class="w-full border rounded-xl px-3 py-2 text-sm bg-gray-50"></div>
      <div><label class="block text-xs text-gray-600 mb-1">อาจารย์ผู้สอน *</label><input name="teacher_name" id="evalFormTeacher" required class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-xs text-gray-600 mb-1">ภาคการศึกษา</label><select name="semester" class="w-full border rounded-xl px-3 py-2 text-sm"><option value="1">1</option><option value="2">2</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">ปีการศึกษา</label><input name="academic_year" value="2568" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      </div>
      <div>
        <label class="block text-xs text-gray-600 mb-1">หัวข้อประเมิน * (คั่นด้วยเครื่องหมาย ,)</label>
        <textarea name="eval_items" required rows="3" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="เนื้อหาการสอน,เทคนิคการสอน,สื่อการสอน,การวัดผล,ความตรงต่อเวลา"></textarea>
        <p class="text-xs text-gray-400 mt-1">ตัวอย่าง: เนื้อหาการสอน,เทคนิคการสอน,สื่อการสอน,การวัดผล,ความตรงต่อเวลา</p>
      </div>
      <div><label class="block text-xs text-gray-600 mb-1">สถานะ</label>
        <select name="status" class="w-full border rounded-xl px-3 py-2 text-sm"><option value="เปิด">เปิดรับประเมิน</option><option value="ปิด">ปิดรับประเมิน</option></select>
      </div>
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">สร้างแบบประเมิน</button>
    </form>
  `);
  document.getElementById('createEvalFormForm').onsubmit = async (e) => {
    e.preventDefault();
    await withLoading(e.target, async () => {
      const fd = new FormData(e.target);
      const obj = { type: 'eval_form', created_at: new Date().toISOString() };
      fd.forEach((v, k) => obj[k] = v);
      const r = await GSheetDB.create(obj);
      if (r.isOk) { showToast('สร้างแบบประเมินสำเร็จ'); closeModal() } else showToast('เกิดข้อผิดพลาด', 'error');
    });
  };
}

function onEvalFormSubjectChange(sel) {
  const opt = sel.options[sel.selectedIndex];
  document.getElementById('evalFormSubCode').value = opt.dataset.code || '';
  document.getElementById('evalFormTeacher').value = opt.dataset.coord || '';
}

// Student: fill eval form
function showStudentEvalForm(formId) {
  const form = APP.allData.find(d => d.__backendId === formId);
  if (!form) return;
  const items = (form.eval_items || '').split(',').map(s => s.trim()).filter(Boolean);
  if (!items.length) { showToast('แบบประเมินนี้ยังไม่มีหัวข้อ', 'error'); return }

  let formHTML = `<div class="mb-4 p-3 bg-blue-50 rounded-xl">
    <p class="font-bold">${form.subject_code ? form.subject_code + ' ' : ''}${form.subject_name || ''}</p>
    <p class="text-sm text-gray-600">อาจารย์: ${form.teacher_name || ''} | ภาค ${semLabel(form.semester)}/${form.academic_year || ''}</p>
  </div>
  <form id="studentEvalForm" class="space-y-4">
    <input type="hidden" name="eval_form_id" value="${formId}">
    <input type="hidden" name="subject_name" value="${form.subject_name || ''}">
    <input type="hidden" name="subject_code" value="${form.subject_code || ''}">
    <input type="hidden" name="teacher_name" value="${form.teacher_name || ''}">
    <input type="hidden" name="semester" value="${form.semester || ''}">
    <input type="hidden" name="academic_year" value="${form.academic_year || ''}">
    <p class="text-sm text-gray-500">ให้คะแนนแต่ละหัวข้อ (1-5)</p>`;

  items.forEach((item, idx) => {
    formHTML += `<div class="border rounded-xl p-3">
      <p class="font-medium text-sm mb-2">${idx + 1}. ${item}</p>
      <div class="flex gap-2">${[1, 2, 3, 4, 5].map(n => `<button type="button" onclick="setItemScore(${idx},${n})" class="eval-item-${idx} w-9 h-9 rounded-full border-2 border-gray-300 flex items-center justify-center text-sm hover:border-yellow-400 hover:bg-yellow-50 transition">${n}</button>`).join('')}</div>
      <input type="hidden" name="score_${idx}" id="scoreInput_${idx}" value="0">
    </div>`;
  });

  formHTML += `<button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">ส่งผลประเมิน</button></form>`;

  showModal('ประเมินอาจารย์ผู้สอน', formHTML);

  document.getElementById('studentEvalForm').onsubmit = async (ev) => {
    ev.preventDefault();
    const fd = new FormData(ev.target);
    // Calculate average score
    let total = 0, count = 0;
    items.forEach((item, idx) => {
      const s = Number(fd.get('score_' + idx)) || 0;
      if (s > 0) { total += s; count++ }
    });
    if (count < items.length) { showToast('กรุณาให้คะแนนทุกหัวข้อ', 'error'); return }
    await withLoading(ev.target, async () => {
      const avg = (total / count).toFixed(1);
      const scoreDetail = items.map((item, idx) => `${item}:${fd.get('score_' + idx)}`).join('|');
      const obj = {
        type: 'evaluation',
        eval_form_id: fd.get('eval_form_id'),
        subject_name: fd.get('subject_name'),
        subject_code: fd.get('subject_code'),
        teacher_name: fd.get('teacher_name'),
        name: fd.get('teacher_name'),
        student_name: APP.currentUser.data?.name || '',
        eval_score: avg,
        eval_detail: scoreDetail,
        semester: fd.get('semester'),
        academic_year: fd.get('academic_year'),
        created_at: new Date().toISOString()
      };
      const r = await GSheetDB.create(obj);
      if (r.isOk) { showToast('ส่งผลประเมินสำเร็จ'); closeModal(); renderCurrentPage() } else showToast('เกิดข้อผิดพลาด', 'error');
    });
  };
}

function setItemScore(idx, score) {
  document.getElementById('scoreInput_' + idx).value = score;
  document.querySelectorAll('.eval-item-' + idx).forEach((btn, i) => {
    btn.classList.toggle('bg-yellow-400', i < score);
    btn.classList.toggle('text-white', i < score);
    btn.classList.toggle('border-yellow-400', i < score);
  });
}

function showAddEvalModal() {
  showCreateEvalFormModal();
}
removed_eval_block_end */

// ======================== TEACHERS ========================
function teachersPage() {
  const isAdmin = APP.currentRole === 'admin';
  const isExecutive = APP.currentRole === 'executive';
  const isAcademic = APP.currentRole === 'academic';
  let allTeachers = applyFilters(getDataByType('teacher'));

  // Department filter
  const allDepts = [...new Set(allTeachers.map(t => t.department).filter(Boolean))].sort();
  const selectedDept = APP.filters._teacherDept || '';
  let data = selectedDept ? allTeachers.filter(t => (t.department || '') === selectedDept) : [];

  const total = data.length; const paged = paginate(data);

  const deptFilter = `<div class="bg-white rounded-2xl p-4 border border-blue-100 mb-4">
    <div class="flex items-center gap-3">
      <label class="text-sm font-medium text-gray-700">สาขาวิชา:</label>
      <select onchange="APP.filters._teacherDept=this.value;APP.pagination.page=1;renderCurrentPage()" class="border border-gray-200 rounded-xl px-3 py-2 text-sm">
        <option value="">-- เลือกสาขาวิชา --</option>
        ${allDepts.map(d => `<option value="${d}" ${selectedDept === d ? 'selected' : ''}>${d}</option>`).join('')}
      </select>
      ${selectedDept ? `<span class="text-xs text-gray-500">แสดงสาขา: ${selectedDept}</span>` : ''}
    </div>
  </div>`;

  // Executive & Academic sees basic info only (ชื่อ-สกุล, ตำแหน่ง, สาขาวิชา, โทร, E-mail)
  if (isExecutive || isAcademic) {
    if (!selectedDept) {
      return `<div class="flex flex-wrap items-center justify-between gap-3 mb-4">
        <h2 class="text-xl font-bold text-gray-800"><i data-lucide="briefcase" class="w-6 h-6 inline mr-2"></i>ข้อมูลอาจารย์</h2>
      </div>
      ${deptFilter}
      ${noYearSelectedMsg('อาจารย์ (กรุณาเลือกสาขาวิชา)')}`;
    }
    return `<div class="flex flex-wrap items-center justify-between gap-3 mb-4">
      <h2 class="text-xl font-bold text-gray-800"><i data-lucide="briefcase" class="w-6 h-6 inline mr-2"></i>ข้อมูลอาจารย์</h2>
    </div>
    ${deptFilter}
    ${filterBar({ semester: false, year: false })}
    <div class="bg-white rounded-2xl border border-blue-100 overflow-hidden">
      <div class="overflow-x-auto"><table class="w-full text-sm">
        <thead><tr class="bg-surface text-left"><th class="px-4 py-3 font-semibold">ชื่อ-สกุล</th><th class="px-4 py-3 font-semibold">ตำแหน่ง</th><th class="px-4 py-3 font-semibold">สาขาวิชา</th><th class="px-4 py-3 font-semibold">สถานะ</th><th class="px-4 py-3 font-semibold">โทร</th><th class="px-4 py-3 font-semibold">E-mail</th></tr></thead>
        <tbody>${paged.length ? paged.map(t => {
      const st = t.teacher_status || 'ปฏิบัติงานอยู่'; const stColor = st === 'ปฏิบัติงานอยู่' ? 'bg-green-100 text-green-700' : 'bg-yellow-100 text-yellow-700'; return `<tr class="border-t hover:bg-gray-50">
          <td class="px-4 py-3 font-medium">${t.name || ''}</td><td class="px-4 py-3">${t.position || ''}</td>
          <td class="px-4 py-3">${t.department || ''}</td><td class="px-4 py-3"><span class="px-2 py-1 rounded-full text-xs ${stColor}">${st}</span></td><td class="px-4 py-3">${t.phone || ''}</td><td class="px-4 py-3">${t.email || ''}</td></tr>`
    }).join('') : '<tr><td colspan="6" class="px-4 py-8 text-center text-gray-400">ไม่มีข้อมูล</td></tr>'}</tbody>
      </table></div>
    </div>
    ${paginationHTML(total, APP.pagination.perPage, APP.pagination.page, 'changePage')}`;
  }

  if (!selectedDept) {
    return `<div class="flex flex-wrap items-center justify-between gap-3 mb-4">
      <h2 class="text-xl font-bold text-gray-800"><i data-lucide="briefcase" class="w-6 h-6 inline mr-2"></i>ข้อมูลอาจารย์</h2>
      ${isAdmin ? `<div class="flex gap-2"><button onclick="showAddTeacherModal()" class="flex items-center gap-2 px-4 py-2 bg-primary text-white rounded-xl hover:bg-primaryDark text-sm"><i data-lucide="plus" class="w-4 h-4"></i>เพิ่มอาจารย์</button></div>` : ''}
    </div>
    ${deptFilter}
    ${noYearSelectedMsg('อาจารย์ (กรุณาเลือกสาขาวิชา)')}`;
  }

  return `<div class="flex flex-wrap items-center justify-between gap-3 mb-4">
    <h2 class="text-xl font-bold text-gray-800"><i data-lucide="briefcase" class="w-6 h-6 inline mr-2"></i>ข้อมูลอาจารย์</h2>
    ${isAdmin ? `<div class="flex gap-2"><button onclick="showAddTeacherModal()" class="flex items-center gap-2 px-4 py-2 bg-primary text-white rounded-xl hover:bg-primaryDark text-sm"><i data-lucide="plus" class="w-4 h-4"></i>เพิ่มอาจารย์</button></div>` : ''}
  </div>
  ${deptFilter}
  ${filterBar({ semester: false, year: false })}
  <div class="bg-white rounded-2xl border border-blue-100 overflow-hidden">
    <div class="overflow-x-auto"><table class="w-full text-sm">
      <thead><tr class="bg-surface text-left"><th class="px-4 py-3 font-semibold">ชื่อ-สกุล</th><th class="px-4 py-3 font-semibold">ตำแหน่ง</th><th class="px-4 py-3 font-semibold">สาขาวิชา</th><th class="px-4 py-3 font-semibold">สถานะ</th><th class="px-4 py-3"></th></tr></thead>
      <tbody>${paged.length ? paged.map(t => {
    const st = t.teacher_status || 'ปฏิบัติงานอยู่'; const stColor = st === 'ปฏิบัติงานอยู่' ? 'bg-green-100 text-green-700' : 'bg-yellow-100 text-yellow-700'; return `<tr class="border-t hover:bg-gray-50">
        <td class="px-4 py-3 font-medium">${t.name || ''}</td><td class="px-4 py-3">${t.position || ''}</td>
        <td class="px-4 py-3">${t.department || ''}</td>
        <td class="px-4 py-3"><span class="px-2 py-1 rounded-full text-xs ${stColor}">${st}</span></td>
        <td class="px-4 py-3"><div class="flex gap-1"><button onclick="showTeacherDetail('${t.__backendId}')" class="text-gray-400 hover:text-primary" title="ดูข้อมูล"><i data-lucide="eye" class="w-4 h-4"></i></button>${isAdmin ? `<button onclick="showEditTeacherModal('${t.__backendId}')" class="text-blue-400 hover:text-blue-600" title="แก้ไข"><i data-lucide="pencil" class="w-4 h-4"></i></button><button onclick="deleteRecord('${t.__backendId}')" class="text-red-400 hover:text-red-600" title="ลบ"><i data-lucide="trash-2" class="w-4 h-4"></i></button>` : ''}</div></td></tr>`
  }).join('') : '<tr><td colspan="5" class="px-4 py-8 text-center text-gray-400">ไม่มีข้อมูล</td></tr>'}</tbody>
    </table></div>
  </div>
  ${paginationHTML(total, APP.pagination.perPage, APP.pagination.page, 'changePage')}`;
}

function showTeacherDetail(id) {
  const t = APP.allData.find(d => d.__backendId === id); if (!t) return;
  const stBadge = (t.teacher_status || 'ปฏิบัติงานอยู่') === 'ปฏิบัติงานอยู่' ? '<span class="px-2 py-1 rounded-full text-xs bg-green-100 text-green-700">ปฏิบัติงานอยู่</span>' : '<span class="px-2 py-1 rounded-full text-xs bg-yellow-100 text-yellow-700">ลาศึกษาต่อ</span>';
  showModal('ข้อมูลอาจารย์', `<div class="grid grid-cols-2 gap-3">
    ${infoRow('ชื่อ-สกุล', t.name)}${infoRow('ตำแหน่ง', t.position)}${infoRow('สาขาวิชา', t.department)}
    <div><p class="text-xs text-gray-500">สถานะ</p><p class="font-medium mt-1">${stBadge}</p></div>
    ${infoRow('โทร', t.phone)}${infoRow('E-mail', t.email)}${infoRow('ชั้นปีที่รับผิดชอบ', t.responsible_year)}
    ${infoRow('เลขบัญชีธนาคาร', t.bank_account)}
  </div>
  <div class="mt-3">${infoRow('ที่อยู่', t.address)}</div>`);
}

function showAddTeacherModal() {
  showModal('เพิ่มอาจารย์', `
    <form id="addTeacherForm" class="space-y-3">
      <div><label class="block text-xs text-gray-600 mb-1">ชื่อ-สกุล</label><input name="name" required class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-xs text-gray-600 mb-1">ตำแหน่ง</label><input name="position" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">สาขาวิชา</label><input name="department" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">โทรศัพท์</label><input name="phone" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">E-mail</label><input name="email" type="email" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">ชั้นปีที่รับผิดชอบ (ถ้ามี)</label><select name="responsible_year" class="w-full border rounded-xl px-3 py-2 text-sm"><option value="">ไม่มี</option><option>1</option><option>2</option><option>3</option><option>4</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">สถานะ</label><select name="teacher_status" class="w-full border rounded-xl px-3 py-2 text-sm"><option>ปฏิบัติงานอยู่</option><option>ลาศึกษาต่อ</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">เลขบัญชีธนาคาร</label><input name="bank_account" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      </div>
      <div><label class="block text-xs text-gray-600 mb-1">ที่อยู่</label><textarea name="address" rows="2" class="w-full border rounded-xl px-3 py-2 text-sm"></textarea></div>
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึก</button>
    </form>
  `);
  document.getElementById('addTeacherForm').onsubmit = async (e) => {
    e.preventDefault();
    await withLoading(e.target, async () => {
      const fd = new FormData(e.target);
      const obj = { type: 'teacher', created_at: new Date().toISOString() }; fd.forEach((v, k) => obj[k] = v);
      const r = await GSheetDB.create(obj);
      if (r.isOk) { showToast('เพิ่มอาจารย์สำเร็จ'); closeModal() } else showToast('เกิดข้อผิดพลาด', 'error');
    });
  };
}


// ======================== TEACHER DIRECTORY (ทำเนียบอาจารย์) ========================

// Helper: render multi-value items as bullet list
function multiValList(val) {
  if (!val) return '-';
  const items = val.split('||').map(v => v.trim()).filter(Boolean);
  if (!items.length) return '-';
  if (items.length === 1) return items[0];
  return items.map(v => `<span class="block">• ${v}</span>`).join('');
}

// Helper: dynamic multi-input field
function multiInputField(name, label, placeholder, values) {
  const items = values ? values.split('||').map(v => v.trim()).filter(Boolean) : [''];
  if (!items.length) items.push('');
  return `
    <div>
      <label class="block text-xs text-gray-600 mb-1">${label} <span class="text-gray-400">(กดปุ่ม + เพื่อเพิ่มรายการ)</span></label>
      <div id="multi_${name}" class="space-y-2">
        ${items.map((v, i) => `<div class="flex gap-2"><input name="${name}__multi" value="${v}" class="flex-1 border rounded-xl px-3 py-2 text-sm" placeholder="${placeholder}">${i > 0 ? `<button type="button" onclick="this.parentElement.remove()" class="text-red-400 hover:text-red-600 px-2" title="ลบ">✕</button>` : '<div class="w-8"></div>'}</div>`).join('')}
      </div>
      <button type="button" onclick="addMultiInput('${name}','${placeholder}')" class="mt-1 text-xs text-primary hover:underline flex items-center gap-1"><span>＋</span> เพิ่มรายการ</button>
    </div>`;
}

function addMultiInput(name, placeholder) {
  const container = document.getElementById('multi_' + name);
  if (!container) return;
  const div = document.createElement('div');
  div.className = 'flex gap-2';
  div.innerHTML = `<input name="${name}__multi" value="" class="flex-1 border rounded-xl px-3 py-2 text-sm" placeholder="${placeholder}"><button type="button" onclick="this.parentElement.remove()" class="text-red-400 hover:text-red-600 px-2" title="ลบ">✕</button>`;
  container.appendChild(div);
}

// Collect multi inputs into || separated string
function collectMultiInputs(form, name) {
  const inputs = form.querySelectorAll(`[name="${name}__multi"]`);
  return Array.from(inputs).map(inp => inp.value.trim()).filter(Boolean).join(' || ');
}

// Helper: mask national ID — show first 9 digits, last 4 as xxxx
function maskNationalId(nid) {
  if (!nid) return '-';
  const s = String(nid).trim();
  if (s.length < 5) return s;
  return s.substring(0, s.length - 4) + 'xxxx';
}

// Export to PDF via print
function exportTeacherDirectoryPDF() {
  let data = getDataByType('teacher_directory');
  const activeTab = APP._directoryTab || 'all';
  const selectedYear = APP.filters._directoryYear || '';
  if (selectedYear) data = data.filter(d => (d.academic_year || '') === selectedYear);
  if (activeTab === 'curriculum') data = data.filter(d => (d.teacher_category || '') === 'อาจารย์ประจำหลักสูตร');
  else if (activeTab === 'regular') data = data.filter(d => (d.teacher_category || '') === 'อาจารย์ประจำ');

  if (!data.length) { showToast('ไม่มีข้อมูลสำหรับส่งออก', 'error'); return; }

  const tabLabel = activeTab === 'all' ? 'ทั้งหมด' : activeTab === 'curriculum' ? 'อาจารย์ประจำหลักสูตร' : 'อาจารย์ประจำ';
  const title = `ทำเนียบอาจารย์ — ${tabLabel}` + (selectedYear ? ` ปีการศึกษา ${selectedYear}` : '');

  const headers = ['#', 'ชื่อ-สกุล', 'เลขบัตรประชาชน', 'เลขใบประกอบวิชาชีพ', 'ตำแหน่งทางวิชาการ', 'วุฒิการศึกษา', 'ประสบการณ์สอน (ปี)', 'ประสบการณ์ปฏิบัติการ (ปี)', 'ผลงานวิชาการ (ย้อนหลัง 5 ปี)', 'ประเภทอาจารย์'];

  let tableRows = '';
  data.forEach((row, i) => {
    tableRows += `<tr>
      <td style="text-align:center">${i + 1}</td>
      <td>${(row.name || '').replace(/\|\|/g, ', ')}</td>
      <td style="text-align:center">${maskNationalId(row.national_id)}</td>
      <td style="text-align:center">${row.license_no || ''}</td>
      <td>${row.academic_position || ''}</td>
      <td>${(row.education || '').replace(/\|\|/g, '<br>')}</td>
      <td style="text-align:center">${row.nursing_teaching_years || ''}</td>
      <td style="text-align:center">${row.nursing_practice_years || ''}</td>
      <td>${(row.academic_work || '').replace(/\|\|/g, '<br>')}</td>
      <td>${row.teacher_category || ''}</td>
    </tr>`;
  });

  const printHTML = `<!DOCTYPE html>
<html lang="th">
<head>
  <meta charset="UTF-8">
  <title>${title}</title>
  <link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;600;700&display=swap" rel="stylesheet">
  <style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body { font-family: 'Sarabun', sans-serif; font-size: 11px; color: #333; }
    @page { size: A4 landscape; margin: 12mm 10mm; }
    h1 { text-align: center; font-size: 18px; font-weight: 700; margin-bottom: 4px; color: #1e6fba; }
    .subtitle { text-align: center; font-size: 12px; color: #666; margin-bottom: 12px; }
    table { width: 100%; border-collapse: collapse; }
    th, td { border: 1px solid #ccc; padding: 4px 6px; vertical-align: top; word-break: break-word; }
    th { background: #1e6fba; color: #fff; font-weight: 600; text-align: center; font-size: 11px; }
    tr:nth-child(even) { background: #f7fafd; }
    tr:hover { background: #eef5fb; }
    @media print {
      body { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    }
  </style>
</head>
<body>
  <h1>${title}</h1>
  <div class="subtitle">วิทยาลัยพยาบาลบรมราชชนนี กรุงเทพ &bull; พิมพ์เมื่อ ${new Date().toLocaleDateString('th-TH', { year: 'numeric', month: 'long', day: 'numeric' })}</div>
  <table>
    <thead><tr>${headers.map(h => `<th>${h}</th>`).join('')}</tr></thead>
    <tbody>${tableRows}</tbody>
  </table>
</body>
</html>`;

  const printWin = window.open('', '_blank');
  printWin.document.write(printHTML);
  printWin.document.close();
  printWin.onload = function () {
    setTimeout(() => { printWin.print(); }, 500);
  };
  showToast('เปิดหน้าต่างพิมพ์ PDF แล้ว — กรุณาเลือก "Save as PDF"');
}

function renderDirectoryDataSection(paged, total, catAll, catCurriculum, catRegular, activeTab, isAdmin) {
  let html = '<div class="grid grid-cols-3 gap-3 mb-4">';
  html += '<div class="card-stat bg-white rounded-2xl p-4 border border-blue-100 text-center cursor-pointer ' + (activeTab === 'all' ? 'ring-2 ring-primary' : '') + '" onclick="APP._directoryTab=\'all\';APP.pagination.page=1;renderCurrentPage()"><p class="text-2xl font-bold text-primary">' + catAll + '</p><p class="text-xs text-gray-500">ทั้งหมด</p></div>';
  html += '<div class="card-stat bg-white rounded-2xl p-4 border border-purple-100 text-center cursor-pointer ' + (activeTab === 'curriculum' ? 'ring-2 ring-purple-500' : '') + '" onclick="APP._directoryTab=\'curriculum\';APP.pagination.page=1;renderCurrentPage()"><p class="text-2xl font-bold text-purple-600">' + catCurriculum + '</p><p class="text-xs text-gray-500">อ.ประจำหลักสูตร</p></div>';
  html += '<div class="card-stat bg-white rounded-2xl p-4 border border-blue-100 text-center cursor-pointer ' + (activeTab === 'regular' ? 'ring-2 ring-blue-500' : '') + '" onclick="APP._directoryTab=\'regular\';APP.pagination.page=1;renderCurrentPage()"><p class="text-2xl font-bold text-blue-600">' + catRegular + '</p><p class="text-xs text-gray-500">อ.ประจำ</p></div>';
  html += '</div>';
  html += filterBar({ semester: false, year: false });
  html += '<div class="bg-white rounded-2xl border border-blue-100 overflow-hidden"><div class="overflow-x-auto"><table class="w-full text-sm">';
  html += '<thead><tr class="bg-surface text-left"><th class="px-4 py-3 font-semibold">ชื่อ-สกุล</th><th class="px-4 py-3 font-semibold">ตำแหน่งทางวิชาการ</th><th class="px-4 py-3 font-semibold">ประเภท</th><th class="px-4 py-3"></th></tr></thead>';
  html += '<tbody>';
  if (paged.length) {
    paged.forEach(t => {
      const catColor = (t.teacher_category || '') === 'อาจารย์ประจำหลักสูตร' ? 'bg-purple-100 text-purple-700' : 'bg-blue-100 text-blue-700';
      html += '<tr class="border-t hover:bg-gray-50">';
      html += '<td class="px-4 py-3 font-medium">' + (t.name || '') + '</td>';
      html += '<td class="px-4 py-3">' + (t.academic_position || '-') + '</td>';
      html += '<td class="px-4 py-3"><span class="px-2 py-1 rounded-full text-xs ' + catColor + '">' + (t.teacher_category || '-') + '</span></td>';
      html += '<td class="px-4 py-3"><div class="flex gap-1">';
      html += '<button onclick="showTeacherDirectoryDetail(\'' + t.__backendId + '\')" class="text-gray-400 hover:text-primary" title="ดูข้อมูล"><i data-lucide="eye" class="w-4 h-4"></i></button>';
      if (isAdmin) {
        html += '<button onclick="showEditTeacherDirectoryModal(\'' + t.__backendId + '\')" class="text-blue-400 hover:text-blue-600" title="แก้ไข"><i data-lucide="pencil" class="w-4 h-4"></i></button>';
        html += '<button onclick="deleteRecord(\'' + t.__backendId + '\')" class="text-red-400 hover:text-red-600" title="ลบ"><i data-lucide="trash-2" class="w-4 h-4"></i></button>';
      }
      html += '</div></td></tr>';
    });
  } else {
    html += '<tr><td colspan="4" class="px-4 py-8 text-center text-gray-400">ไม่มีข้อมูล</td></tr>';
  }
  html += '</tbody></table></div></div>';
  html += paginationHTML(total, APP.pagination.perPage, APP.pagination.page, 'changePage');
  return html;
}

function teacherDirectoryPage() {
  const isAdmin = APP.currentRole === 'admin' || APP.currentRole === 'academic';
  let allData = applyFilters(getDataByType('teacher_directory'));

  // Academic year filter
  const allYears = [...new Set(allData.map(d => d.academic_year).filter(Boolean))].sort().reverse();
  const selectedYear = APP.filters._directoryYear || '';
  if (selectedYear) allData = allData.filter(d => (d.academic_year || '') === selectedYear);
  let data = allData;

  // Counts by category
  const catAll = data.length;
  const catCurriculum = data.filter(d => (d.teacher_category || '') === 'อาจารย์ประจำหลักสูตร').length;
  const catRegular = data.filter(d => (d.teacher_category || '') === 'อาจารย์ประจำ').length;

  // Tab filter
  const activeTab = APP._directoryTab || 'all';
  if (activeTab === 'curriculum') data = data.filter(d => (d.teacher_category || '') === 'อาจารย์ประจำหลักสูตร');
  else if (activeTab === 'regular') data = data.filter(d => (d.teacher_category || '') === 'อาจารย์ประจำ');

  const total = data.length; const paged = paginate(data);

  function tabBtn(id, label, count) {
    const active = activeTab === id;
    return `<button onclick="APP._directoryTab='${id}';APP.pagination.page=1;renderCurrentPage()" class="px-4 py-2 rounded-xl text-sm font-medium transition ${active ? 'bg-primary text-white shadow' : 'bg-white text-gray-600 hover:bg-surface border border-gray-200'}">${label} <span class="ml-1 px-1.5 py-0.5 rounded-full text-xs ${active ? 'bg-white/20' : 'bg-gray-100'}">${count}</span></button>`;
  }

  return `<div class="flex flex-wrap items-center justify-between gap-3 mb-4">
    <h2 class="text-xl font-bold text-gray-800"><i data-lucide="award" class="w-6 h-6 inline mr-2"></i>ทำเนียบอาจารย์</h2>
    <div class="flex gap-2">
      <button onclick="exportTeacherDirectoryPDF()" class="flex items-center gap-2 px-4 py-2 bg-red-500 text-white rounded-xl hover:bg-red-600 text-sm"><i data-lucide="file-text" class="w-4 h-4"></i>ส่งออก PDF</button>
      ${isAdmin ? `<button onclick="showAddTeacherDirectoryModal()" class="flex items-center gap-2 px-4 py-2 bg-primary text-white rounded-xl hover:bg-primaryDark text-sm"><i data-lucide="plus" class="w-4 h-4"></i>เพิ่มอาจารย์</button>${csvUploadBtn('teacher_directory', 'name,national_id,license_no,academic_position,education,nursing_teaching_years,nursing_teaching_exp,nursing_practice_years,nursing_practice_exp,academic_work,academic_year,teacher_category')}` : ''}
    </div>
  </div>

  <div class="bg-white rounded-2xl p-4 border border-blue-100 mb-4">
    <div class="flex items-center gap-3">
      <label class="text-sm font-medium text-gray-700">ปีการศึกษา:</label>
      <select onchange="APP.filters._directoryYear=this.value;APP.pagination.page=1;renderCurrentPage()" class="border border-gray-200 rounded-xl px-3 py-2 text-sm">
        <option value="">-- เลือกปีการศึกษา --</option>
        ${allYears.map(y => `<option value="${y}" ${selectedYear === y ? 'selected' : ''}>${y}</option>`).join('')}
      </select>
      ${selectedYear ? `<span class="text-xs text-gray-500">แสดงข้อมูลปีการศึกษา ${selectedYear}</span>` : ''}
    </div>
  </div>

  ${selectedYear ? renderDirectoryDataSection(paged, total, catAll, catCurriculum, catRegular, activeTab, isAdmin) : noYearSelectedMsg('ทำเนียบอาจารย์')}`;
}

function showTeacherDirectoryDetail(id) {
  const t = APP.allData.find(d => d.__backendId === id); if (!t) return;
  const catColor = (t.teacher_category || '') === 'อาจารย์ประจำหลักสูตร' ? 'bg-purple-100 text-purple-700' : 'bg-blue-100 text-blue-700';

  function detailList(val) {
    if (!val) return '<span class="text-gray-400">-</span>';
    const items = val.split('||').map(v => v.trim()).filter(Boolean);
    if (!items.length) return '<span class="text-gray-400">-</span>';
    if (items.length === 1) return `<span class="text-sm">${items[0]}</span>`;
    return '<ul class="list-disc list-inside text-sm space-y-1">' + items.map(v => `<li>${v}</li>`).join('') + '</ul>';
  }

  showModal('ข้อมูลทำเนียบอาจารย์', `
    <div class="space-y-3">
      <div class="flex items-center gap-3 mb-2"><div class="w-12 h-12 bg-primary rounded-full flex items-center justify-center"><i data-lucide="user" class="w-6 h-6 text-white"></i></div><div><p class="font-bold text-lg">${t.name || '-'}</p><span class="px-2 py-1 rounded-full text-xs ${catColor}">${t.teacher_category || '-'}</span></div></div>
      <div class="grid grid-cols-2 gap-3">
        ${infoRow('เลขบัตรประชาชน', maskNationalId(t.national_id))}
        ${infoRow('เลขใบประกอบวิชาชีพ', t.license_no)}
        ${infoRow('ตำแหน่งทางวิชาการ', t.academic_position)}
        ${infoRow('ปีการศึกษา', t.academic_year)}
      </div>
      <div class="bg-surface rounded-xl p-3"><p class="text-xs text-gray-500 mb-1 font-semibold">วุฒิการศึกษา</p>${detailList(t.education)}</div>
      <div class="bg-surface rounded-xl p-3"><p class="text-xs text-gray-500 mb-1 font-semibold">ประสบการณ์สอนทางการพยาบาล ${t.nursing_teaching_years ? '<span class="text-primary font-bold">(' + t.nursing_teaching_years + ' ปี)</span>' : ''}</p>${detailList(t.nursing_teaching_exp)}</div>
      <div class="bg-surface rounded-xl p-3"><p class="text-xs text-gray-500 mb-1 font-semibold">ประสบการณ์ปฏิบัติการพยาบาล ${t.nursing_practice_years ? '<span class="text-primary font-bold">(' + t.nursing_practice_years + ' ปี)</span>' : ''}</p>${detailList(t.nursing_practice_exp)}</div>
      <div class="bg-surface rounded-xl p-3"><p class="text-xs text-gray-500 mb-1 font-semibold">ผลงานวิชาการ (ย้อนหลัง 5 ปี)</p>${detailList(t.academic_work)}</div>
    </div>
  `);
}

function showAddTeacherDirectoryModal() {
  showModal('เพิ่มอาจารย์ (ทำเนียบ)', `
    <form id="addTeacherDirForm" class="space-y-3" style="max-height:70vh;overflow-y:auto;padding-right:4px">
      <div><label class="block text-xs text-gray-600 mb-1">ชื่อ-สกุล *</label><input name="name" required class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-xs text-gray-600 mb-1">เลขบัตรประชาชน</label><input name="national_id" maxlength="13" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="13 หลัก"></div>
        <div><label class="block text-xs text-gray-600 mb-1">เลขใบประกอบวิชาชีพ</label><input name="license_no" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      </div>
      <div><label class="block text-xs text-gray-600 mb-1">ตำแหน่งทางวิชาการ</label><input name="academic_position" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="เช่น ผศ.ดร., รศ."></div>
      ${multiInputField('education', 'วุฒิการศึกษา', 'เช่น พย.บ., พย.ม., ปร.ด.', '')}
      <div>
        <label class="block text-xs text-gray-600 mb-1">ประสบการณ์สอนทางการพยาบาล <span class="text-gray-400">(กดปุ่ม + เพื่อเพิ่มรายการ)</span></label>
        <div class="flex items-center gap-2 mb-1"><span class="text-xs text-gray-500 w-16 text-right flex-shrink-0">จำนวนปี</span><input name="nursing_teaching_years" type="number" min="0" class="w-20 border rounded-xl px-3 py-2 text-sm" placeholder="ปี"></div>
        ${multiInputField('nursing_teaching_exp', 'รายละเอียด', 'เช่น สอนวิชาการพยาบาลผู้ใหญ่', '')}
      </div>
      <div>
        <label class="block text-xs text-gray-600 mb-1">ประสบการณ์ปฏิบัติการพยาบาล <span class="text-gray-400">(กดปุ่ม + เพื่อเพิ่มรายการ)</span></label>
        <div class="flex items-center gap-2 mb-1"><span class="text-xs text-gray-500 w-16 text-right flex-shrink-0">จำนวนปี</span><input name="nursing_practice_years" type="number" min="0" class="w-20 border rounded-xl px-3 py-2 text-sm" placeholder="ปี"></div>
        ${multiInputField('nursing_practice_exp', 'รายละเอียด', 'เช่น พยาบาลวิชาชีพ รพ.รามาธิบดี', '')}
      </div>
      ${multiInputField('academic_work', 'ผลงานวิชาการ (ย้อนหลัง 5 ปี)', 'เช่น บทความวิจัยเรื่อง...', '')}
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-xs text-gray-600 mb-1">ปีการศึกษา *</label><input name="academic_year" required class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="เช่น 2568" value="2568"></div>
        <div><label class="block text-xs text-gray-600 mb-1">ประเภทอาจารย์ *</label>
          <select name="teacher_category" required class="w-full border rounded-xl px-3 py-2 text-sm">
            <option value="">-- เลือกประเภท --</option>
            <option>อาจารย์ประจำหลักสูตร</option>
            <option>อาจารย์ประจำ</option>
          </select>
        </div>
      </div>
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึก</button>
    </form>
  `);
  document.getElementById('addTeacherDirForm').onsubmit = async (e) => {
    e.preventDefault();
    await withLoading(e.target, async () => {
      const form = e.target;
      const obj = { type: 'teacher_directory', created_at: new Date().toISOString() };
      obj.name = form.querySelector('[name="name"]').value;
      obj.national_id = form.querySelector('[name="national_id"]').value;
      obj.license_no = form.querySelector('[name="license_no"]').value;
      obj.academic_position = form.querySelector('[name="academic_position"]').value;
      obj.education = collectMultiInputs(form, 'education');
      obj.nursing_teaching_years = form.querySelector('[name="nursing_teaching_years"]').value;
      obj.nursing_teaching_exp = collectMultiInputs(form, 'nursing_teaching_exp');
      obj.nursing_practice_years = form.querySelector('[name="nursing_practice_years"]').value;
      obj.nursing_practice_exp = collectMultiInputs(form, 'nursing_practice_exp');
      obj.academic_work = collectMultiInputs(form, 'academic_work');
      obj.academic_year = form.querySelector('[name="academic_year"]').value;
      obj.teacher_category = form.querySelector('[name="teacher_category"]').value;
      const r = await GSheetDB.create(obj);
      if (r.isOk) { showToast('เพิ่มอาจารย์สำเร็จ'); closeModal() } else showToast('เกิดข้อผิดพลาด', 'error');
    });
  };
}

function showEditTeacherDirectoryModal(id) {
  const t = APP.allData.find(d => d.__backendId === id); if (!t) return;
  showModal('แก้ไขข้อมูลทำเนียบอาจารย์', `
    <form id="editTeacherDirForm" class="space-y-3" style="max-height:70vh;overflow-y:auto;padding-right:4px">
      <div><label class="block text-xs text-gray-600 mb-1">ชื่อ-สกุล</label><input name="name" value="${t.name || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-xs text-gray-600 mb-1">เลขบัตรประชาชน</label><input name="national_id" value="${t.national_id || ''}" maxlength="13" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">เลขใบประกอบวิชาชีพ</label><input name="license_no" value="${t.license_no || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      </div>
      <div><label class="block text-xs text-gray-600 mb-1">ตำแหน่งทางวิชาการ</label><input name="academic_position" value="${t.academic_position || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      ${multiInputField('education', 'วุฒิการศึกษา', 'เช่น พย.บ., พย.ม., ปร.ด.', t.education || '')}
      <div>
        <label class="block text-xs text-gray-600 mb-1">ประสบการณ์สอนทางการพยาบาล <span class="text-gray-400">(กดปุ่ม + เพื่อเพิ่มรายการ)</span></label>
        <div class="flex items-center gap-2 mb-1"><span class="text-xs text-gray-500 w-16 text-right flex-shrink-0">จำนวนปี</span><input name="nursing_teaching_years" type="number" min="0" value="${t.nursing_teaching_years || ''}" class="w-20 border rounded-xl px-3 py-2 text-sm" placeholder="ปี"></div>
        ${multiInputField('nursing_teaching_exp', 'รายละเอียด', 'เช่น สอนวิชาการพยาบาลผู้ใหญ่', t.nursing_teaching_exp || '')}
      </div>
      <div>
        <label class="block text-xs text-gray-600 mb-1">ประสบการณ์ปฏิบัติการพยาบาล <span class="text-gray-400">(กดปุ่ม + เพื่อเพิ่มรายการ)</span></label>
        <div class="flex items-center gap-2 mb-1"><span class="text-xs text-gray-500 w-16 text-right flex-shrink-0">จำนวนปี</span><input name="nursing_practice_years" type="number" min="0" value="${t.nursing_practice_years || ''}" class="w-20 border rounded-xl px-3 py-2 text-sm" placeholder="ปี"></div>
        ${multiInputField('nursing_practice_exp', 'รายละเอียด', 'เช่น พยาบาลวิชาชีพ รพ.รามาธิบดี', t.nursing_practice_exp || '')}
      </div>
      ${multiInputField('academic_work', 'ผลงานวิชาการ (ย้อนหลัง 5 ปี)', 'เช่น บทความวิจัยเรื่อง...', t.academic_work || '')}
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-xs text-gray-600 mb-1">ปีการศึกษา</label><input name="academic_year" value="${t.academic_year || ''}" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="เช่น 2568"></div>
        <div><label class="block text-xs text-gray-600 mb-1">ประเภทอาจารย์</label>
          <select name="teacher_category" class="w-full border rounded-xl px-3 py-2 text-sm">
            <option value="">-- เลือกประเภท --</option>
            <option ${(t.teacher_category || '') === 'อาจารย์ประจำหลักสูตร' ? 'selected' : ''}>อาจารย์ประจำหลักสูตร</option>
            <option ${(t.teacher_category || '') === 'อาจารย์ประจำ' ? 'selected' : ''}>อาจารย์ประจำ</option>
          </select>
        </div>
      </div>
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึกการแก้ไข</button>
    </form>
  `);
  document.getElementById('editTeacherDirForm').onsubmit = async (e) => {
    e.preventDefault();
    await withLoading(e.target, async () => {
      const form = e.target;
      const rec = APP.allData.find(d => d.__backendId === id); if (!rec) return;
      rec.name = form.querySelector('[name="name"]').value;
      rec.national_id = form.querySelector('[name="national_id"]').value;
      rec.license_no = form.querySelector('[name="license_no"]').value;
      rec.academic_position = form.querySelector('[name="academic_position"]').value;
      rec.education = collectMultiInputs(form, 'education');
      rec.nursing_teaching_years = form.querySelector('[name="nursing_teaching_years"]').value;
      rec.nursing_teaching_exp = collectMultiInputs(form, 'nursing_teaching_exp');
      rec.nursing_practice_years = form.querySelector('[name="nursing_practice_years"]').value;
      rec.nursing_practice_exp = collectMultiInputs(form, 'nursing_practice_exp');
      rec.academic_work = collectMultiInputs(form, 'academic_work');
      rec.academic_year = form.querySelector('[name="academic_year"]').value;
      rec.teacher_category = form.querySelector('[name="teacher_category"]').value;
      const r = await GSheetDB.update(rec);
      if (r.isOk) { showToast('แก้ไขข้อมูลสำเร็จ'); closeModal(); renderCurrentPage(); } else showToast('เกิดข้อผิดพลาด: ' + (r.error || ''), 'error');
    });
  };
}

// ======================== SERVICES ========================
function servicesPage() {
  const announcements = getDataByType('announcement').slice(-10).reverse();
  const docRequests = getDataByType('doc_request').slice(-20).reverse();

  return `<h2 class="text-xl font-bold text-gray-800 mb-4"><i data-lucide="grid" class="w-6 h-6 inline mr-2"></i>บริการอื่นๆ</h2>
  <div class="grid grid-cols-1 lg:grid-cols-2 gap-4">
    <div class="bg-white rounded-2xl p-5 border border-blue-100">
      <div class="flex items-center justify-between mb-4"><h3 class="font-bold">ข่าวสาร/แจ้งเตือน</h3><button onclick="showAddAnnouncementModal()" class="text-primary hover:underline text-sm">+ เพิ่มประกาศ</button></div>
      ${announcements.length ? announcements.map(a => `<div class="p-3 bg-surface rounded-xl mb-2 flex justify-between items-start">
        <div><p class="font-medium text-sm">${a.announcement_title || ''}</p><p class="text-xs text-gray-500">${a.announcement_date || ''} · ${a.event_type || 'ทั่วไป'}</p><p class="text-xs text-gray-600 mt-1">${(a.announcement_content || '').substring(0, 80)}</p></div>
        <div class="flex gap-1 ml-2"><button onclick="showEditAnnouncementModal('${a.__backendId}')" class="text-blue-400 hover:text-blue-600" title="แก้ไข"><i data-lucide="pencil" class="w-4 h-4"></i></button><button onclick="deleteRecord('${a.__backendId}')" class="text-red-400 hover:text-red-600" title="ลบ"><i data-lucide="trash-2" class="w-4 h-4"></i></button></div>
      </div>`).join('') : '<p class="text-gray-400 text-center py-6 text-sm">ไม่มีประกาศ</p>'}
    </div>
    <div class="bg-white rounded-2xl p-5 border border-blue-100">
      <h3 class="font-bold mb-4">คำร้องขอเอกสาร</h3>
      ${docRequests.length ? docRequests.map(d => `<div class="p-3 bg-surface rounded-xl mb-2 flex justify-between items-center">
        <div><p class="font-medium text-sm">${d.name || ''} - ${d.doc_request_type || ''}</p><p class="text-xs text-gray-500">${d.created_at ? new Date(d.created_at).toLocaleDateString('th-TH') : ''}</p></div>
        <div class="flex items-center gap-2">
          <select onchange="updateDocStatus('${d.__backendId}',this.value)" class="text-xs border rounded-lg px-2 py-1">
            <option ${d.doc_status === 'รอดำเนินการ' ? 'selected' : ''}>รอดำเนินการ</option>
            <option ${d.doc_status === 'กำลังดำเนินการ' ? 'selected' : ''}>กำลังดำเนินการ</option>
            <option ${d.doc_status === 'เสร็จสิ้น' ? 'selected' : ''}>เสร็จสิ้น</option>
          </select>
        </div>
      </div>`).join('') : '<p class="text-gray-400 text-center py-6 text-sm">ไม่มีคำร้อง</p>'}
    </div>
  </div>`;
}

function showAddAnnouncementModal() {
  showModal('เพิ่มประกาศ/แจ้งเตือน', `
    <form id="addAnnForm" class="space-y-3">
      <div><label class="block text-xs text-gray-600 mb-1">เรื่อง</label><input name="announcement_title" required class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <div><label class="block text-xs text-gray-600 mb-1">เนื้อหา</label><textarea name="announcement_content" rows="3" class="w-full border rounded-xl px-3 py-2 text-sm"></textarea></div>
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-xs text-gray-600 mb-1">วันที่</label><input name="announcement_date" type="date" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">ประเภท</label><select name="event_type" class="w-full border rounded-xl px-3 py-2 text-sm"><option>ทั่วไป</option><option>สอบ</option><option>วันหยุด</option><option>กิจกรรม</option></select></div>
      </div>
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึก</button>
    </form>
  `);
  document.getElementById('addAnnForm').onsubmit = async (e) => {
    e.preventDefault();
    await withLoading(e.target, async () => {
      const fd = new FormData(e.target);
      const obj = { type: 'announcement', created_at: new Date().toISOString() }; fd.forEach((v, k) => obj[k] = v);
      const r = await GSheetDB.create(obj);
      if (r.isOk) { showToast('เพิ่มประกาศสำเร็จ'); closeModal() } else showToast('เกิดข้อผิดพลาด', 'error');
    });
  };
}

async function updateDocStatus(id, status) {
  const rec = APP.allData.find(d => d.__backendId === id); if (!rec) return;
  rec.doc_status = status;
  const r = await GSheetDB.update(rec);
  if (r.isOk) showToast('อัปเดตสถานะสำเร็จ'); else showToast('เกิดข้อผิดพลาด', 'error');
}

// ======================== TRACKING ========================
function trackingPage() {
  const isAdmin = APP.currentRole === 'admin' || APP.currentRole === 'academic';
  const isExecutive = APP.currentRole === 'executive';
  const canApprove = isAdmin || isExecutive;
  const canEdit = isAdmin || APP.currentRole === 'teacher' || APP.currentRole === 'classTeacher';
  let data = getDataByType('tracking').filter(t => t.subject_name && t.subject_name.trim());
  if (APP.currentRole === 'teacher') data = data.filter(t => t.coordinator === APP.currentUser.name);

  // Year filter for stats
  const selectedYear = APP.filters._trackingYear || '';
  const allSubjects = getDataByType('subject');
  const subjectsFiltered = selectedYear ? allSubjects.filter(s => norm(s.academic_year) === selectedYear) : allSubjects;
  const dataForStats = selectedYear ? data.filter(t => norm(t.academic_year) === selectedYear) : data;

  // Only show not-submitted and stats when year is selected
  let statsSection = '';
  let notSubmittedSection = '';
  if (selectedYear) {
    // Find subjects not yet in tracking
    const trackedKeys = dataForStats.map(t => `${norm(t.subject_name)}|${norm(t.semester)}|${norm(t.academic_year)}`);
    const notSubmitted = subjectsFiltered.filter(s => !trackedKeys.includes(`${norm(s.subject_name)}|${norm(s.semester)}|${norm(s.academic_year)}`));

    // Status counts
    const completed = dataForStats.filter(t => t.deputy_sign === 'เสร็จสิ้น').length;
    const inProgress = dataForStats.filter(t => (t.class_teacher_check === 'เสร็จสิ้น' || t.academic_propose === 'เสร็จสิ้น') && t.deputy_sign !== 'เสร็จสิ้น').length;
    const pending = dataForStats.length - completed - inProgress;

    statsSection = `<div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3 mt-3">
      ${statCard('alert-circle', 'ยังไม่ส่ง', notSubmitted.length, 'วิชา', 'bg-red-500')}
      ${statCard('clock', 'รอดำเนินการ', pending, 'วิชา', 'bg-yellow-500')}
      ${statCard('loader', 'กำลังดำเนินการ', inProgress, 'วิชา', 'bg-blue-500')}
      ${statCard('check-circle', 'เสร็จสิ้น', completed, 'วิชา', 'bg-green-500')}
    </div>`;
    if (notSubmitted.length) {
      notSubmittedSection = `<div class="bg-red-50 rounded-2xl p-4 border border-red-200 mb-4">
        <h3 class="font-bold text-red-700 mb-2 text-sm flex items-center gap-2"><i data-lucide="alert-triangle" class="w-4 h-4"></i>รายวิชาที่ยังไม่ส่งรายละเอียด (${notSubmitted.length} วิชา)</h3>
        <div class="flex flex-wrap gap-2">${notSubmitted.map(s => `<span class="px-3 py-1 bg-white border border-red-200 rounded-lg text-xs text-red-700">${s.subject_code ? s.subject_code + ' ' : ''}${s.subject_name || ''} <span class="text-gray-400">(ภาค ${s.semester || ''})</span></span>`).join('')}</div>
      </div>`;
    }
  }

  // Year options
  const allYears = [...new Set([...allSubjects.map(s => s.academic_year), ...data.map(t => t.academic_year)].filter(Boolean))].sort();

  // Apply general filters to table data (also filter by selected year)
  if (selectedYear) data = data.filter(t => norm(t.academic_year) === selectedYear);
  data = applyFilters(data);
  const total = data.length; const paged = paginate(data);

  const isClassTeacher = APP.currentRole === 'classTeacher';

  return `<div class="flex flex-wrap items-center justify-between gap-3 mb-4">
    <h2 class="text-xl font-bold text-gray-800"><i data-lucide="file-check" class="w-6 h-6 inline mr-2"></i>ติดตามการส่งรายละเอียดรายวิชา</h2>
    ${canEdit ? `<button onclick="showAddTrackingModal()" class="flex items-center gap-2 px-4 py-2 bg-primary text-white rounded-xl hover:bg-primaryDark text-sm"><i data-lucide="plus" class="w-4 h-4"></i>เพิ่มข้อมูล</button>` : ''}
  </div>
  <div class="bg-white rounded-2xl p-4 border border-blue-100 mb-4">
    <div class="flex items-center gap-3 mb-3">
      <label class="text-sm font-medium text-gray-700">ปีการศึกษา:</label>
      <select onchange="APP.filters._trackingYear=this.value;renderCurrentPage()" class="border border-gray-200 rounded-xl px-3 py-2 text-sm">
        <option value="">-- เลือกปีการศึกษา --</option>
        ${allYears.map(y => `<option value="${y}" ${selectedYear === y ? 'selected' : ''}>${y}</option>`).join('')}
      </select>
    </div>
    ${statsSection}
  </div>
  ${selectedYear ? `${notSubmittedSection}
  ${filterBar({ yearLevel: true })}
  <div class="bg-white rounded-2xl border border-blue-100 overflow-hidden">
    <div class="overflow-x-auto"><table class="w-full text-sm">
      <thead><tr class="bg-surface text-left"><th class="px-4 py-3 font-semibold">รายวิชา</th><th class="px-4 py-3 font-semibold">ทฤษฎี/ปฏิบัติ</th><th class="px-4 py-3 font-semibold">ชั้นปี</th><th class="px-4 py-3 font-semibold">ภาค/ปี</th><th class="px-4 py-3 font-semibold">ผู้ประสานงาน</th><th class="px-4 py-3 font-semibold">อ.ประจำชั้นตรวจ</th><th class="px-4 py-3 font-semibold">วิชาการเสนอ</th><th class="px-4 py-3 font-semibold">รอง ผอ.ลงนาม</th><th class="px-4 py-3 font-semibold">วันอนุมัติ</th><th class="px-4 py-3 font-semibold">ไฟล์</th><th class="px-4 py-3 font-semibold">หมายเหตุ</th>${isAdmin ? '<th class="px-4 py-3"></th>' : ''}</tr></thead>
      <tbody>${paged.length ? paged.map(t => {
    const sb = (s) => s === 'เสร็จสิ้น' ? '<span class="px-2 py-1 rounded-full text-xs bg-green-100 text-green-700">✓ เสร็จสิ้น</span>' : s === 'ส่งกลับแก้ไข' ? '<span class="px-2 py-1 rounded-full text-xs bg-red-100 text-red-700">↩ ส่งกลับแก้ไข</span>' : s === 'กำลังดำเนินการ' ? '<span class="px-2 py-1 rounded-full text-xs bg-yellow-100 text-yellow-700">⏳ ดำเนินการ</span>' : '<span class="px-2 py-1 rounded-full text-xs bg-gray-100 text-gray-500">รอ</span>';
    const isLate = t.is_late === 'ใช่' || t.is_late === 'late';
    const id = t.__backendId;

    const ctCell = (isClassTeacher && (t.class_teacher_check === 'รอ' || t.class_teacher_check === 'ส่งกลับแก้ไข' || !t.class_teacher_check))
      ? `<div class="flex flex-col gap-1"><button onclick="updateTrackingField('${id}','class_teacher_check','เสร็จสิ้น')" class="px-2 py-1 bg-green-500 hover:bg-green-600 text-white rounded-lg text-xs">✓ ตรวจสอบแล้ว</button><button onclick="updateTrackingField('${id}','class_teacher_check','ส่งกลับแก้ไข')" class="px-2 py-1 bg-orange-500 hover:bg-orange-600 text-white rounded-lg text-xs">↩ ส่งกลับแก้ไข</button></div>`
      : (isAdmin ? `<select onchange="updateTrackingField('${id}','class_teacher_check',this.value)" class="text-xs border rounded px-1 py-0.5"><option ${t.class_teacher_check === 'รอ' ? 'selected' : ''}>รอ</option><option ${t.class_teacher_check === 'เสร็จสิ้น' ? 'selected' : ''}>เสร็จสิ้น</option><option ${t.class_teacher_check === 'ส่งกลับแก้ไข' ? 'selected' : ''}>ส่งกลับแก้ไข</option></select>` : sb(t.class_teacher_check));

    const acCell = (isAdmin && t.class_teacher_check === 'เสร็จสิ้น' && (t.academic_propose === 'รอ' || t.academic_propose === 'ส่งกลับแก้ไข' || !t.academic_propose))
      ? `<div class="flex flex-col gap-1"><button onclick="updateTrackingField('${id}','academic_propose','เสร็จสิ้น')" class="px-2 py-1 bg-green-500 hover:bg-green-600 text-white rounded-lg text-xs">✓ ส่งเสนอรองฯ แล้ว</button><button onclick="updateTrackingField('${id}','academic_propose','ส่งกลับแก้ไข')" class="px-2 py-1 bg-orange-500 hover:bg-orange-600 text-white rounded-lg text-xs">↩ ส่งกลับแก้ไข</button></div>`
      : (isAdmin ? `<select onchange="updateTrackingField('${id}','academic_propose',this.value)" class="text-xs border rounded px-1 py-0.5"><option ${t.academic_propose === 'รอ' ? 'selected' : ''}>รอ</option><option ${t.academic_propose === 'เสร็จสิ้น' ? 'selected' : ''}>เสร็จสิ้น</option><option ${t.academic_propose === 'ส่งกลับแก้ไข' ? 'selected' : ''}>ส่งกลับแก้ไข</option></select>` : sb(t.academic_propose));

    const dpCell = (isExecutive && t.academic_propose === 'เสร็จสิ้น' && (t.deputy_sign === 'รอ' || t.deputy_sign === 'ส่งกลับแก้ไข' || !t.deputy_sign))
      ? `<div class="flex flex-col gap-1"><button onclick="updateTrackingField('${id}','deputy_sign','เสร็จสิ้น')" class="px-2 py-1 bg-green-500 hover:bg-green-600 text-white rounded-lg text-xs">✓ ลงนามแล้ว</button><button onclick="updateTrackingField('${id}','deputy_sign','ส่งกลับแก้ไข')" class="px-2 py-1 bg-orange-500 hover:bg-orange-600 text-white rounded-lg text-xs">↩ ส่งกลับแก้ไข</button></div>`
      : (isAdmin ? `<select onchange="updateTrackingField('${id}','deputy_sign',this.value)" class="text-xs border rounded px-1 py-0.5"><option ${t.deputy_sign === 'รอ' ? 'selected' : ''}>รอ</option><option ${t.deputy_sign === 'เสร็จสิ้น' ? 'selected' : ''}>เสร็จสิ้น</option><option ${t.deputy_sign === 'ส่งกลับแก้ไข' ? 'selected' : ''}>ส่งกลับแก้ไข</option></select>` : sb(t.deputy_sign));

    let fCell = '<span class="text-xs text-gray-400">-</span>';
    if (t.deputy_sign === 'เสร็จสิ้น') {
      if (t.file_link) {
        fCell = `<a href="${t.file_link}" target="_blank" class="px-2 py-1 bg-blue-100 text-blue-700 rounded-lg text-xs hover:bg-blue-200 inline-flex items-center gap-1"><i data-lucide="file-text" class="w-3 h-3"></i>ดาวน์โหลด</a>`;
        if (isAdmin) fCell += ` <button onclick="promptTrackingFileLink('${id}')" class="text-blue-400 hover:text-blue-600 ml-1" title="แก้ไขลิงก์"><i data-lucide="pencil" class="w-3 h-3"></i></button>`;
      } else if (isAdmin) {
        fCell = `<button onclick="promptTrackingFileLink('${id}')" class="px-2 py-1 bg-emerald-500 hover:bg-emerald-600 text-white rounded-lg text-xs inline-flex items-center gap-1"><i data-lucide="upload" class="w-3 h-3"></i>เพิ่มไฟล์ PDF</button>`;
      }
    }

    return `<tr class="border-t hover:bg-gray-50 ${isLate ? 'bg-red-50' : ''}">
        <td class="px-4 py-3 font-medium">${t.subject_name || ''}${isLate ? ' <span class="text-xs text-red-500 font-normal">(ส่งช้า)</span>' : ''}</td>
        <td class="px-4 py-3">${t.theory_practice || ''}</td>
        <td class="px-4 py-3">${t.year_level || ''}</td>
        <td class="px-4 py-3">${semLabel(t.semester)}/${t.academic_year || ''}</td>
        <td class="px-4 py-3 text-xs">${t.coordinator || ''}</td>
        <td class="px-4 py-3">${ctCell}</td>
        <td class="px-4 py-3">${acCell}</td>
        <td class="px-4 py-3">${dpCell}</td>
        <td class="px-4 py-3">${t.approved_date ? toBuddhistDate(t.approved_date) : '-'}</td>
        <td class="px-4 py-3">${fCell}</td>
        <td class="px-4 py-3 text-xs text-gray-500">${t.remarks || ''}</td>
        ${isAdmin ? `<td class="px-4 py-3"><div class="flex gap-1"><button onclick="showEditTrackingModal('${id}')" class="text-blue-400 hover:text-blue-600" title="แก้ไข"><i data-lucide="pencil" class="w-4 h-4"></i></button><button onclick="deleteRecord('${id}')" class="text-red-400 hover:text-red-600" title="ลบ"><i data-lucide="trash-2" class="w-4 h-4"></i></button></div></td>` : ''}
      </tr>`
  }).join('') : '<tr><td colspan="12" class="px-4 py-8 text-center text-gray-400">ไม่มีข้อมูล</td></tr>'}</tbody>
    </table></div>
  </div>
  ${paginationHTML(total, APP.pagination.perPage, APP.pagination.page, 'changePage')}` : noYearSelectedMsg('ติดตามการส่งรายละเอียดรายวิชา')}`;
}

function showAddTrackingModal() {
  const subjects = getDataByType('subject');
  const subjectOptions = [...new Set(subjects.map(s => s.subject_name).filter(Boolean))].sort()
    .map(name => `<option value="${name.replace(/"/g, '&quot;')}">${name}</option>`).join('');
  const teachers = getDataByType('teacher');
  const teacherList = [...new Set(teachers.map(t => (t.name || '').trim()).filter(Boolean))].sort();
  const myName = (APP.currentUser && APP.currentUser.name || '').trim();
  const coordCheckboxes = teacherList.map(name => {
    const isMe = name === myName;
    return `<label class="flex items-center gap-2 p-1.5 rounded-lg hover:bg-gray-50 cursor-pointer ${isMe ? 'bg-blue-50' : ''}"><input type="checkbox" class="coord-check w-4 h-4 rounded border-gray-300 text-primary focus:ring-primary" value="${name.replace(/"/g, '&quot;')}" ${isMe ? 'checked' : ''}><span class="text-sm">${name}${isMe ? ' <span class="text-xs text-blue-500">(คุณ)</span>' : ''}</span></label>`;
  }).join('');
  showModal('เพิ่มรายละเอียดรายวิชา', `
    <form id="addTrackingForm" class="space-y-3">
      <div><label class="block text-xs text-gray-600 mb-1">ชื่อรายวิชา *</label><select name="subject_name" required class="w-full border rounded-xl px-3 py-2 text-sm"><option value="">-- เลือกรายวิชา --</option>${subjectOptions}</select></div>
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-xs text-gray-600 mb-1">ทฤษฎี/ปฏิบัติ</label><select name="theory_practice" class="w-full border rounded-xl px-3 py-2 text-sm"><option>ทฤษฎี</option><option>ปฏิบัติ</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">ชั้นปี</label><select name="year_level" class="w-full border rounded-xl px-3 py-2 text-sm"><option>1</option><option>2</option><option>3</option><option>4</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">ห้อง</label><input name="room" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">ภาคการศึกษา</label><select name="semester" class="w-full border rounded-xl px-3 py-2 text-sm"><option value="1">1</option><option value="2">2</option><option value="3">ฤดูร้อน</option></select></div>
      </div>
      <div>
        <label class="block text-xs text-gray-600 mb-1">ผู้ประสานงานรายวิชา (เลือกจากรายชื่อ หรือพิมพ์เพิ่มเติม)</label>
        <div class="max-h-40 overflow-y-auto border rounded-xl p-2 bg-gray-50 space-y-0.5">${coordCheckboxes || '<p class="text-xs text-gray-400 p-2">ยังไม่มีข้อมูลอาจารย์ในระบบ</p>'}</div>
        <input type="text" id="trackingCoordExtra" class="w-full border rounded-xl px-3 py-2 text-sm mt-2" placeholder="พิมพ์ชื่อเพิ่มเติม (คั่นด้วยเครื่องหมาย ,)">
        <input type="hidden" name="coordinator" id="trackingCoordHidden">
      </div>
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึก</button>
    </form>
  `);
  document.getElementById('addTrackingForm').onsubmit = async (e) => {
    e.preventDefault();
    // Sync coordinator checkboxes to hidden field
    const checked = [...e.target.querySelectorAll('.coord-check:checked')].map(cb => cb.value);
    const extraText = (document.getElementById('trackingCoordExtra') || {}).value || '';
    const extraNames = extraText.split(',').map(s => s.trim()).filter(Boolean);
    const allNames = [...checked, ...extraNames].filter(Boolean);
    const coordHidden = document.getElementById('trackingCoordHidden');
    if (coordHidden) coordHidden.value = allNames.join(', ');
    await withLoading(e.target, async () => {
      const fd = new FormData(e.target);
      const obj = { type: 'tracking', class_teacher_check: 'รอ', academic_propose: 'รอ', deputy_sign: 'รอ', approved_date: '', created_at: new Date().toISOString() };
      fd.forEach((v, k) => obj[k] = v);
      const r = await GSheetDB.create(obj);
      if (r.isOk) { showToast('เพิ่มสำเร็จ'); closeModal() } else showToast('เกิดข้อผิดพลาด', 'error');
    });
  };
}

async function updateTrackingField(id, field, value) {
  const el = window.event ? (window.event.target.closest('button') || window.event.target.closest('select')) : null;
  await withLoading(el, async () => {
    const rec = APP.allData.find(d => d.__backendId === id); if (!rec) return;
    rec[field] = value;
    if (field === 'deputy_sign' && value === 'เสร็จสิ้น') rec.approved_date = new Date().toISOString().split('T')[0];

    const r = await GSheetDB.update(rec);
    if (r.isOk) { showToast('อัปเดตสำเร็จ'); renderCurrentPage(); } else showToast('เกิดข้อผิดพลาด', 'error');
  });
}

function promptTrackingFileLink(id) {
  const rec = APP.allData.find(d => d.__backendId === id); if (!rec) return;
  showModal('เพิ่ม/แก้ไขลิงก์ไฟล์ PDF', `
    <div class="space-y-3">
      <div class="bg-blue-50 rounded-xl p-3 text-sm space-y-1">
        <p><span class="text-gray-500">รายวิชา:</span> <strong>${rec.subject_name || '-'}</strong></p>
        <p><span class="text-gray-500">ภาค/ปี:</span> <strong>${semLabel(rec.semester)}/${rec.academic_year || ''}</strong></p>
      </div>
      <form id="trackingFileLinkForm" class="space-y-3">
        <div>
          <label class="block text-xs text-gray-600 mb-1">ลิงก์ไฟล์ PDF (Google Drive หรือ URL อื่น)</label>
          <input name="file_link" id="trackingFileLinkInput" value="${rec.file_link || ''}" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="https://drive.google.com/file/d/...">
        </div>
        <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark flex items-center justify-center gap-2"><i data-lucide="save" class="w-4 h-4"></i>บันทึกลิงก์</button>
      </form>
    </div>
  `);
  setTimeout(() => { lucide.createIcons(); const inp = document.getElementById('trackingFileLinkInput'); if (inp) inp.focus(); }, 50);
  const f = document.getElementById('trackingFileLinkForm');
  if (f) f.onsubmit = async (ev) => {
    ev.preventDefault();
    const link = (new FormData(f).get('file_link') || '').toString().trim();
    closeModal();
    await withLoading(null, async () => {
      rec.file_link = link;
      const r = await GSheetDB.update(rec);
      if (r.isOk) showToast('บันทึกลิงก์สำเร็จ'); else showToast('เกิดข้อผิดพลาด', 'error');
    });
  };
}


// ======================== RESULT TRACKING ========================
function resultTrackingPage() {
  const isAdmin = APP.currentRole === 'admin' || APP.currentRole === 'academic';
  const isExecutive = APP.currentRole === 'executive';
  const isClassTeacher = APP.currentRole === 'classTeacher';
  const canEdit = isAdmin || APP.currentRole === 'teacher' || APP.currentRole === 'classTeacher';
  let data = getDataByType('result_tracking').filter(t => t.subject_name && t.subject_name.trim());
  if (APP.currentRole === 'teacher') data = data.filter(t => t.coordinator === APP.currentUser.name);

  const selectedYear = APP.filters._resultTrackingYear || '';
  const allSubjects = getDataByType('subject');
  const subjectsFiltered = selectedYear ? allSubjects.filter(s => norm(s.academic_year) === selectedYear) : allSubjects;
  const dataForStats = selectedYear ? data.filter(t => norm(t.academic_year) === selectedYear) : data;

  let statsSection = '';
  let notSubmittedSection = '';
  if (selectedYear) {
    const trackedKeys = dataForStats.map(t => `${norm(t.subject_name)}|${norm(t.semester)}|${norm(t.academic_year)}`);
    const notSubmitted = subjectsFiltered.filter(s => !trackedKeys.includes(`${norm(s.subject_name)}|${norm(s.semester)}|${norm(s.academic_year)}`));

    const completed = dataForStats.filter(t => t.deputy_sign === 'เสร็จสิ้น').length;
    const inProgress = dataForStats.filter(t => (t.class_teacher_check === 'เสร็จสิ้น' || t.academic_propose === 'เสร็จสิ้น') && t.deputy_sign !== 'เสร็จสิ้น').length;
    const pending = dataForStats.length - completed - inProgress;

    statsSection = `<div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3 mt-3">
      ${statCard('alert-circle', 'ยังไม่ส่ง', notSubmitted.length, 'วิชา', 'bg-red-500')}
      ${statCard('clock', 'รอดำเนินการ', pending, 'วิชา', 'bg-yellow-500')}
      ${statCard('loader', 'กำลังดำเนินการ', inProgress, 'วิชา', 'bg-blue-500')}
      ${statCard('check-circle', 'เสร็จสิ้น', completed, 'วิชา', 'bg-green-500')}
    </div>`;
    if (notSubmitted.length) {
      notSubmittedSection = `<div class="bg-red-50 rounded-2xl p-4 border border-red-200 mb-4">
        <h3 class="font-bold text-red-700 mb-2 text-sm flex items-center gap-2"><i data-lucide="alert-triangle" class="w-4 h-4"></i>รายวิชาที่ยังไม่ส่งผลการดำเนินงาน (${notSubmitted.length} วิชา)</h3>
        <div class="flex flex-wrap gap-2">${notSubmitted.map(s => `<span class="px-3 py-1 bg-white border border-red-200 rounded-lg text-xs text-red-700">${s.subject_code ? s.subject_code + ' ' : ''}${s.subject_name || ''} <span class="text-gray-400">(ภาค ${s.semester || ''})</span></span>`).join('')}</div>
      </div>`;
    }
  }

  const allYears = [...new Set([...allSubjects.map(s => s.academic_year), ...data.map(t => t.academic_year)].filter(Boolean))].sort();

  if (selectedYear) data = data.filter(t => norm(t.academic_year) === selectedYear);
  data = applyFilters(data);
  const total = data.length; const paged = paginate(data);

  return `<div class="flex flex-wrap items-center justify-between gap-3 mb-4">
    <h2 class="text-xl font-bold text-gray-800"><i data-lucide="bar-chart-2" class="w-6 h-6 inline mr-2"></i>ติดตามการส่งผลการดำเนินงานรายวิชา</h2>
    ${canEdit ? `<button onclick="showAddResultTrackingModal()" class="flex items-center gap-2 px-4 py-2 bg-primary text-white rounded-xl hover:bg-primaryDark text-sm"><i data-lucide="plus" class="w-4 h-4"></i>เพิ่มข้อมูล</button>` : ''}
  </div>
  <div class="bg-white rounded-2xl p-4 border border-blue-100 mb-4">
    <div class="flex items-center gap-3 mb-3">
      <label class="text-sm font-medium text-gray-700">ปีการศึกษา:</label>
      <select onchange="APP.filters._resultTrackingYear=this.value;renderCurrentPage()" class="border border-gray-200 rounded-xl px-3 py-2 text-sm">
        <option value="">-- เลือกปีการศึกษา --</option>
        ${allYears.map(y => `<option value="${y}" ${selectedYear === y ? 'selected' : ''}>${y}</option>`).join('')}
      </select>
    </div>
    ${statsSection}
  </div>
  ${selectedYear ? `${notSubmittedSection}
  ${filterBar({ yearLevel: true })}
  <div class="bg-white rounded-2xl border border-blue-100 overflow-hidden">
    <div class="overflow-x-auto"><table class="w-full text-sm">
      <thead><tr class="bg-surface text-left"><th class="px-4 py-3 font-semibold">รายวิชา</th><th class="px-4 py-3 font-semibold">ทฤษฎี/ปฏิบัติ</th><th class="px-4 py-3 font-semibold">ชั้นปี</th><th class="px-4 py-3 font-semibold">ภาค/ปี</th><th class="px-4 py-3 font-semibold">ผู้ประสานงาน</th><th class="px-4 py-3 font-semibold">อ.ประจำชั้นตรวจ</th><th class="px-4 py-3 font-semibold">วิชาการเสนอ</th><th class="px-4 py-3 font-semibold">รอง ผอ.ลงนาม</th><th class="px-4 py-3 font-semibold">วันอนุมัติ</th><th class="px-4 py-3 font-semibold">ไฟล์</th><th class="px-4 py-3 font-semibold">หมายเหตุ</th>${isAdmin ? '<th class="px-4 py-3"></th>' : ''}</tr></thead>
      <tbody>${paged.length ? paged.map(t => {
    const sb = (s) => s === 'เสร็จสิ้น' ? '<span class="px-2 py-1 rounded-full text-xs bg-green-100 text-green-700">✓ เสร็จสิ้น</span>' : s === 'ส่งกลับแก้ไข' ? '<span class="px-2 py-1 rounded-full text-xs bg-red-100 text-red-700">↩ ส่งกลับแก้ไข</span>' : s === 'กำลังดำเนินการ' ? '<span class="px-2 py-1 rounded-full text-xs bg-yellow-100 text-yellow-700">⏳ ดำเนินการ</span>' : '<span class="px-2 py-1 rounded-full text-xs bg-gray-100 text-gray-500">รอ</span>';
    const isLate = t.is_late === 'ใช่' || t.is_late === 'late';
    const id = t.__backendId;

    const ctCell = (isClassTeacher && (t.class_teacher_check === 'รอ' || t.class_teacher_check === 'ส่งกลับแก้ไข' || !t.class_teacher_check))
      ? `<div class="flex flex-col gap-1"><button onclick="updateTrackingField('${id}','class_teacher_check','เสร็จสิ้น')" class="px-2 py-1 bg-green-500 hover:bg-green-600 text-white rounded-lg text-xs">✓ ตรวจสอบแล้ว</button><button onclick="updateTrackingField('${id}','class_teacher_check','ส่งกลับแก้ไข')" class="px-2 py-1 bg-orange-500 hover:bg-orange-600 text-white rounded-lg text-xs">↩ ส่งกลับแก้ไข</button></div>`
      : (isAdmin ? `<select onchange="updateTrackingField('${id}','class_teacher_check',this.value)" class="text-xs border rounded px-1 py-0.5"><option ${t.class_teacher_check === 'รอ' ? 'selected' : ''}>รอ</option><option ${t.class_teacher_check === 'เสร็จสิ้น' ? 'selected' : ''}>เสร็จสิ้น</option><option ${t.class_teacher_check === 'ส่งกลับแก้ไข' ? 'selected' : ''}>ส่งกลับแก้ไข</option></select>` : sb(t.class_teacher_check));

    const acCell = (isAdmin && t.class_teacher_check === 'เสร็จสิ้น' && (t.academic_propose === 'รอ' || t.academic_propose === 'ส่งกลับแก้ไข' || !t.academic_propose))
      ? `<div class="flex flex-col gap-1"><button onclick="updateTrackingField('${id}','academic_propose','เสร็จสิ้น')" class="px-2 py-1 bg-green-500 hover:bg-green-600 text-white rounded-lg text-xs">✓ ส่งเสนอรองฯ แล้ว</button><button onclick="updateTrackingField('${id}','academic_propose','ส่งกลับแก้ไข')" class="px-2 py-1 bg-orange-500 hover:bg-orange-600 text-white rounded-lg text-xs">↩ ส่งกลับแก้ไข</button></div>`
      : (isAdmin ? `<select onchange="updateTrackingField('${id}','academic_propose',this.value)" class="text-xs border rounded px-1 py-0.5"><option ${t.academic_propose === 'รอ' ? 'selected' : ''}>รอ</option><option ${t.academic_propose === 'เสร็จสิ้น' ? 'selected' : ''}>เสร็จสิ้น</option><option ${t.academic_propose === 'ส่งกลับแก้ไข' ? 'selected' : ''}>ส่งกลับแก้ไข</option></select>` : sb(t.academic_propose));

    const dpCell = (isExecutive && t.academic_propose === 'เสร็จสิ้น' && (t.deputy_sign === 'รอ' || t.deputy_sign === 'ส่งกลับแก้ไข' || !t.deputy_sign))
      ? `<div class="flex flex-col gap-1"><button onclick="updateTrackingField('${id}','deputy_sign','เสร็จสิ้น')" class="px-2 py-1 bg-green-500 hover:bg-green-600 text-white rounded-lg text-xs">✓ ลงนามแล้ว</button><button onclick="updateTrackingField('${id}','deputy_sign','ส่งกลับแก้ไข')" class="px-2 py-1 bg-orange-500 hover:bg-orange-600 text-white rounded-lg text-xs">↩ ส่งกลับแก้ไข</button></div>`
      : (isAdmin ? `<select onchange="updateTrackingField('${id}','deputy_sign',this.value)" class="text-xs border rounded px-1 py-0.5"><option ${t.deputy_sign === 'รอ' ? 'selected' : ''}>รอ</option><option ${t.deputy_sign === 'เสร็จสิ้น' ? 'selected' : ''}>เสร็จสิ้น</option><option ${t.deputy_sign === 'ส่งกลับแก้ไข' ? 'selected' : ''}>ส่งกลับแก้ไข</option></select>` : sb(t.deputy_sign));

    let fCell = '<span class="text-xs text-gray-400">-</span>';
    if (t.deputy_sign === 'เสร็จสิ้น') {
      if (t.file_link) {
        fCell = `<a href="${t.file_link}" target="_blank" class="px-2 py-1 bg-blue-100 text-blue-700 rounded-lg text-xs hover:bg-blue-200 inline-flex items-center gap-1"><i data-lucide="file-text" class="w-3 h-3"></i>ดาวน์โหลด</a>`;
        if (isAdmin) fCell += ` <button onclick="promptTrackingFileLink('${id}')" class="text-blue-400 hover:text-blue-600 ml-1" title="แก้ไขลิงก์"><i data-lucide="pencil" class="w-3 h-3"></i></button>`;
      } else if (isAdmin) {
        fCell = `<button onclick="promptTrackingFileLink('${id}')" class="px-2 py-1 bg-emerald-500 hover:bg-emerald-600 text-white rounded-lg text-xs inline-flex items-center gap-1"><i data-lucide="upload" class="w-3 h-3"></i>เพิ่มไฟล์ PDF</button>`;
      }
    }

    return `<tr class="border-t hover:bg-gray-50 ${isLate ? 'bg-red-50' : ''}">
        <td class="px-4 py-3 font-medium">${t.subject_name || ''}${isLate ? ' <span class="text-xs text-red-500 font-normal">(ส่งช้า)</span>' : ''}</td>
        <td class="px-4 py-3">${t.theory_practice || ''}</td>
        <td class="px-4 py-3">${t.year_level || ''}</td>
        <td class="px-4 py-3">${semLabel(t.semester)}/${t.academic_year || ''}</td>
        <td class="px-4 py-3 text-xs">${t.coordinator || ''}</td>
        <td class="px-4 py-3">${ctCell}</td>
        <td class="px-4 py-3">${acCell}</td>
        <td class="px-4 py-3">${dpCell}</td>
        <td class="px-4 py-3">${t.approved_date ? toBuddhistDate(t.approved_date) : '-'}</td>
        <td class="px-4 py-3">${fCell}</td>
        <td class="px-4 py-3 text-xs text-gray-500">${t.remarks || ''}</td>
        ${isAdmin ? `<td class="px-4 py-3"><div class="flex gap-1"><button onclick="showEditTrackingModal('${id}')" class="text-blue-400 hover:text-blue-600" title="แก้ไข"><i data-lucide="pencil" class="w-4 h-4"></i></button><button onclick="deleteRecord('${id}')" class="text-red-400 hover:text-red-600" title="ลบ"><i data-lucide="trash-2" class="w-4 h-4"></i></button></div></td>` : ''}
      </tr>`
  }).join('') : '<tr><td colspan="12" class="px-4 py-8 text-center text-gray-400">ไม่มีข้อมูล</td></tr>'}</tbody>
    </table></div>
  </div>
  ${paginationHTML(total, APP.pagination.perPage, APP.pagination.page, 'changePage')}` : noYearSelectedMsg('ติดตามการส่งผลการดำเนินงานรายวิชา')}`;
}

function showAddResultTrackingModal() {
  const subjects = getDataByType('subject');
  const subjectOptions = [...new Set(subjects.map(s => s.subject_name).filter(Boolean))].sort()
    .map(name => `<option value="${name.replace(/"/g, '&quot;')}">${name}</option>`).join('');
  const teachers = getDataByType('teacher');
  const teacherList = [...new Set(teachers.map(t => (t.name || '').trim()).filter(Boolean))].sort();
  const myName = (APP.currentUser && APP.currentUser.name || '').trim();
  const coordCheckboxes = teacherList.map(name => {
    const isMe = name === myName;
    return `<label class="flex items-center gap-2 p-1.5 rounded-lg hover:bg-gray-50 cursor-pointer ${isMe ? 'bg-blue-50' : ''}"><input type="checkbox" class="coord-check w-4 h-4 rounded border-gray-300 text-primary focus:ring-primary" value="${name.replace(/"/g, '&quot;')}" ${isMe ? 'checked' : ''}><span class="text-sm">${name}${isMe ? ' <span class="text-xs text-blue-500">(คุณ)</span>' : ''}</span></label>`;
  }).join('');
  showModal('เพิ่มข้อมูลติดตามการส่งผลการดำเนินงานรายวิชา', `
    <form id="addResultTrackingForm" class="space-y-3">
      <div><label class="block text-xs text-gray-600 mb-1">ชื่อรายวิชา *</label><select name="subject_name" required class="w-full border rounded-xl px-3 py-2 text-sm"><option value="">-- เลือกรายวิชา --</option>${subjectOptions}</select></div>
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-xs text-gray-600 mb-1">ทฤษฎี/ปฏิบัติ</label><select name="theory_practice" class="w-full border rounded-xl px-3 py-2 text-sm"><option>ทฤษฎี</option><option>ปฏิบัติ</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">ชั้นปี</label><select name="year_level" class="w-full border rounded-xl px-3 py-2 text-sm"><option>1</option><option>2</option><option>3</option><option>4</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">ภาคการศึกษา</label><select name="semester" class="w-full border rounded-xl px-3 py-2 text-sm"><option value="1">1</option><option value="2">2</option><option value="3">ฤดูร้อน</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">ปีการศึกษา</label><input name="academic_year" value="2568" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      </div>
      <div>
        <label class="block text-xs text-gray-600 mb-1">ผู้ประสานงานรายวิชา (เลือกจากรายชื่อ หรือพิมพ์เพิ่มเติม)</label>
        <div class="max-h-40 overflow-y-auto border rounded-xl p-2 bg-gray-50 space-y-0.5">${coordCheckboxes || '<p class="text-xs text-gray-400 p-2">ยังไม่มีข้อมูลอาจารย์ในระบบ</p>'}</div>
        <input type="text" id="resultTrackingCoordExtra" class="w-full border rounded-xl px-3 py-2 text-sm mt-2" placeholder="พิมพ์ชื่อเพิ่มเติม (คั่นด้วยเครื่องหมาย ,)">
        <input type="hidden" name="coordinator" id="resultTrackingCoordHidden">
      </div>
      <div><label class="block text-xs text-gray-600 mb-1">หมายเหตุ</label><input name="remarks" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึก</button>
    </form>
  `);
  document.getElementById('addResultTrackingForm').onsubmit = async (e) => {
    e.preventDefault();
    const checked = [...e.target.querySelectorAll('.coord-check:checked')].map(cb => cb.value);
    const extraText = (document.getElementById('resultTrackingCoordExtra') || {}).value || '';
    const extraNames = extraText.split(',').map(s => s.trim()).filter(Boolean);
    const allNames = [...checked, ...extraNames].filter(Boolean);
    const coordHidden = document.getElementById('resultTrackingCoordHidden');
    if (coordHidden) coordHidden.value = allNames.join(', ');
    await withLoading(e.target, async () => {
      const fd = new FormData(e.target);
      const obj = { type: 'result_tracking', class_teacher_check: 'รอ', academic_propose: 'รอ', deputy_sign: 'รอ', approved_date: '', created_at: new Date().toISOString() };
      fd.forEach((v, k) => obj[k] = v);
      const r = await GSheetDB.create(obj);
      if (r.isOk) { showToast('เพิ่มสำเร็จ'); closeModal() } else showToast('เกิดข้อผิดพลาด', 'error');
    });
  };
}
// ======================== GRADE TRACKING ========================
function gradeTrackingPage() {
  const isAdmin = APP.currentRole === 'admin' || APP.currentRole === 'academic';
  const isExecutive = APP.currentRole === 'executive';
  const isClassTeacher = APP.currentRole === 'classTeacher';
  const canEdit = isAdmin || APP.currentRole === 'teacher' || APP.currentRole === 'classTeacher';
  let data = getDataByType('grade_tracking').filter(t => t.subject_name && t.subject_name.trim());
  if (APP.currentRole === 'teacher') data = data.filter(t => t.coordinator === APP.currentUser.name);

  const selectedYear = APP.filters._gradeTrackingYear || '';
  const allSubjects = getDataByType('subject');
  const subjectsFiltered = selectedYear ? allSubjects.filter(s => norm(s.academic_year) === selectedYear) : allSubjects;
  const dataForStats = selectedYear ? data.filter(t => norm(t.academic_year) === selectedYear) : data;

  let statsSection = '';
  let notSubmittedSection = '';
  if (selectedYear) {
    const trackedKeys = dataForStats.map(t => `${norm(t.subject_name)}|${norm(t.semester)}|${norm(t.academic_year)}`);
    const notSubmitted = subjectsFiltered.filter(s => !trackedKeys.includes(`${norm(s.subject_name)}|${norm(s.semester)}|${norm(s.academic_year)}`));

    const completed = dataForStats.filter(t => t.deputy_sign === 'เสร็จสิ้น').length;
    const inProgress = dataForStats.filter(t => (t.coordinator_check === 'เสร็จสิ้น' || t.academic_check === 'เสร็จสิ้น') && t.deputy_sign !== 'เสร็จสิ้น').length;
    const pending = dataForStats.length - completed - inProgress;

    statsSection = `<div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3 mt-3">
      ${statCard('alert-circle', 'ยังไม่ส่งเกรด', notSubmitted.length, 'วิชา', 'bg-red-500')}
      ${statCard('clock', 'รอดำเนินการ', pending, 'วิชา', 'bg-yellow-500')}
      ${statCard('loader', 'กำลังดำเนินการ', inProgress, 'วิชา', 'bg-blue-500')}
      ${statCard('check-circle', 'เสร็จสิ้น', completed, 'วิชา', 'bg-green-500')}
    </div>`;
    if (notSubmitted.length) {
      notSubmittedSection = `<div class="bg-red-50 rounded-2xl p-4 border border-red-200 mb-4">
        <h3 class="font-bold text-red-700 mb-2 text-sm flex items-center gap-2"><i data-lucide="alert-triangle" class="w-4 h-4"></i>รายวิชาที่ยังไม่ส่งเกรด (${notSubmitted.length} วิชา)</h3>
        <div class="flex flex-wrap gap-2">${notSubmitted.map(s => `<span class="px-3 py-1 bg-white border border-red-200 rounded-lg text-xs text-red-700">${s.subject_code ? s.subject_code + ' ' : ''}${s.subject_name || ''} <span class="text-gray-400">(ภาค ${s.semester || ''})</span></span>`).join('')}</div>
      </div>`;
    }
  }

  const allYears = [...new Set([...allSubjects.map(s => s.academic_year), ...data.map(t => t.academic_year)].filter(Boolean))].sort();

  if (selectedYear) data = data.filter(t => norm(t.academic_year) === selectedYear);
  data = applyFilters(data);
  const total = data.length; const paged = paginate(data);

  return `<div class="flex flex-wrap items-center justify-between gap-3 mb-4">
    <h2 class="text-xl font-bold text-gray-800"><i data-lucide="clipboard-check" class="w-6 h-6 inline mr-2"></i>ติดตามการส่งเกรดรายวิชา</h2>
    ${canEdit ? `<button onclick="showAddGradeTrackingModal()" class="flex items-center gap-2 px-4 py-2 bg-primary text-white rounded-xl hover:bg-primaryDark text-sm"><i data-lucide="plus" class="w-4 h-4"></i>เพิ่มข้อมูล</button>` : ''}
  </div>
  <div class="bg-white rounded-2xl p-4 border border-blue-100 mb-4">
    <div class="flex items-center gap-3 mb-3">
      <label class="text-sm font-medium text-gray-700">ปีการศึกษา:</label>
      <select onchange="APP.filters._gradeTrackingYear=this.value;renderCurrentPage()" class="border border-gray-200 rounded-xl px-3 py-2 text-sm">
        <option value="">-- เลือกปีการศึกษา --</option>
        ${allYears.map(y => `<option value="${y}" ${selectedYear === y ? 'selected' : ''}>${y}</option>`).join('')}
      </select>
    </div>
    ${statsSection}
  </div>
  ${selectedYear ? `${notSubmittedSection}
  ${filterBar({ yearLevel: true })}
  <div class="bg-white rounded-2xl border border-blue-100 overflow-hidden">
    <div class="overflow-x-auto"><table class="w-full text-sm">
      <thead><tr class="bg-surface text-left"><th class="px-4 py-3 font-semibold">รายวิชา</th><th class="px-4 py-3 font-semibold">ทฤษฎี/ปฏิบัติ</th><th class="px-4 py-3 font-semibold">ชั้นปี</th><th class="px-4 py-3 font-semibold">ภาค/ปี</th><th class="px-4 py-3 font-semibold">ผู้ประสานงาน</th><th class="px-4 py-3 font-semibold">อ.ประจำชั้นตรวจ</th><th class="px-4 py-3 font-semibold">วิชาการเสนอ</th><th class="px-4 py-3 font-semibold">รอง ผอ.ลงนาม</th><th class="px-4 py-3 font-semibold">วันอนุมัติ</th><th class="px-4 py-3 font-semibold">ไฟล์</th><th class="px-4 py-3 font-semibold">หมายเหตุ</th>${isAdmin ? '<th class="px-4 py-3"></th>' : ''}</tr></thead>
      <tbody>${paged.length ? paged.map(t => {
    const sb = (s) => s === 'เสร็จสิ้น' ? '<span class="px-2 py-1 rounded-full text-xs bg-green-100 text-green-700">✓ เสร็จสิ้น</span>' : s === 'ส่งกลับแก้ไข' ? '<span class="px-2 py-1 rounded-full text-xs bg-red-100 text-red-700">↩ ส่งกลับแก้ไข</span>' : s === 'กำลังดำเนินการ' ? '<span class="px-2 py-1 rounded-full text-xs bg-yellow-100 text-yellow-700">⏳ ดำเนินการ</span>' : '<span class="px-2 py-1 rounded-full text-xs bg-gray-100 text-gray-500">รอ</span>';
    const isLate = t.is_late === 'ใช่' || t.is_late === 'late';
    const id = t.__backendId;

    const ctCell = (isClassTeacher && (t.coordinator_check === 'รอ' || t.coordinator_check === 'ส่งกลับแก้ไข' || !t.coordinator_check))
      ? `<div class="flex flex-col gap-1"><button onclick="updateTrackingField('${id}','coordinator_check','เสร็จสิ้น')" class="px-2 py-1 bg-green-500 hover:bg-green-600 text-white rounded-lg text-xs">✓ ตรวจสอบแล้ว</button><button onclick="updateTrackingField('${id}','coordinator_check','ส่งกลับแก้ไข')" class="px-2 py-1 bg-orange-500 hover:bg-orange-600 text-white rounded-lg text-xs">↩ ส่งกลับแก้ไข</button></div>`
      : (isAdmin ? `<select onchange="updateTrackingField('${id}','coordinator_check',this.value)" class="text-xs border rounded px-1 py-0.5"><option ${t.coordinator_check === 'รอ' ? 'selected' : ''}>รอ</option><option ${t.coordinator_check === 'เสร็จสิ้น' ? 'selected' : ''}>เสร็จสิ้น</option><option ${t.coordinator_check === 'ส่งกลับแก้ไข' ? 'selected' : ''}>ส่งกลับแก้ไข</option></select>` : sb(t.coordinator_check));

    const acCell = (isAdmin && t.coordinator_check === 'เสร็จสิ้น' && (t.academic_check === 'รอ' || t.academic_check === 'ส่งกลับแก้ไข' || !t.academic_check))
      ? `<div class="flex flex-col gap-1"><button onclick="updateTrackingField('${id}','academic_check','เสร็จสิ้น')" class="px-2 py-1 bg-green-500 hover:bg-green-600 text-white rounded-lg text-xs">✓ ส่งเสนอรองฯ แล้ว</button><button onclick="updateTrackingField('${id}','academic_check','ส่งกลับแก้ไข')" class="px-2 py-1 bg-orange-500 hover:bg-orange-600 text-white rounded-lg text-xs">↩ ส่งกลับแก้ไข</button></div>`
      : (isAdmin ? `<select onchange="updateTrackingField('${id}','academic_check',this.value)" class="text-xs border rounded px-1 py-0.5"><option ${t.academic_check === 'รอ' ? 'selected' : ''}>รอ</option><option ${t.academic_check === 'เสร็จสิ้น' ? 'selected' : ''}>เสร็จสิ้น</option><option ${t.academic_check === 'ส่งกลับแก้ไข' ? 'selected' : ''}>ส่งกลับแก้ไข</option></select>` : sb(t.academic_check));

    const dpCell = (isExecutive && t.academic_check === 'เสร็จสิ้น' && (t.deputy_sign === 'รอ' || t.deputy_sign === 'ส่งกลับแก้ไข' || !t.deputy_sign))
      ? `<div class="flex flex-col gap-1"><button onclick="updateTrackingField('${id}','deputy_sign','เสร็จสิ้น')" class="px-2 py-1 bg-green-500 hover:bg-green-600 text-white rounded-lg text-xs">✓ ลงนามแล้ว</button><button onclick="updateTrackingField('${id}','deputy_sign','ส่งกลับแก้ไข')" class="px-2 py-1 bg-orange-500 hover:bg-orange-600 text-white rounded-lg text-xs">↩ ส่งกลับแก้ไข</button></div>`
      : (isAdmin ? `<select onchange="updateTrackingField('${id}','deputy_sign',this.value)" class="text-xs border rounded px-1 py-0.5"><option ${t.deputy_sign === 'รอ' ? 'selected' : ''}>รอ</option><option ${t.deputy_sign === 'เสร็จสิ้น' ? 'selected' : ''}>เสร็จสิ้น</option><option ${t.deputy_sign === 'ส่งกลับแก้ไข' ? 'selected' : ''}>ส่งกลับแก้ไข</option></select>` : sb(t.deputy_sign));

    let fCell = '<span class="text-xs text-gray-400">-</span>';
    if (t.deputy_sign === 'เสร็จสิ้น') {
      if (t.file_link) {
        fCell = `<a href="${t.file_link}" target="_blank" class="px-2 py-1 bg-blue-100 text-blue-700 rounded-lg text-xs hover:bg-blue-200 inline-flex items-center gap-1"><i data-lucide="file-text" class="w-3 h-3"></i>ดาวน์โหลด</a>`;
        if (isAdmin) fCell += ` <button onclick="promptTrackingFileLink('${id}')" class="text-blue-400 hover:text-blue-600 ml-1" title="แก้ไขลิงก์"><i data-lucide="pencil" class="w-3 h-3"></i></button>`;
      } else if (isAdmin) {
        fCell = `<button onclick="promptTrackingFileLink('${id}')" class="px-2 py-1 bg-emerald-500 hover:bg-emerald-600 text-white rounded-lg text-xs inline-flex items-center gap-1"><i data-lucide="upload" class="w-3 h-3"></i>เพิ่มไฟล์ PDF</button>`;
      }
    }

    return `<tr class="border-t hover:bg-gray-50 ${isLate ? 'bg-red-50' : ''}">
        <td class="px-4 py-3 font-medium">${t.subject_name || ''}${isLate ? ' <span class="text-xs text-red-500 font-normal">(ส่งช้า)</span>' : ''}</td>
        <td class="px-4 py-3">${t.theory_practice || ''}</td>
        <td class="px-4 py-3">${t.year_level || ''}</td>
        <td class="px-4 py-3">${semLabel(t.semester)}/${t.academic_year || ''}</td>
        <td class="px-4 py-3 text-xs">${t.coordinator || ''}</td>
        <td class="px-4 py-3">${ctCell}</td>
        <td class="px-4 py-3">${acCell}</td>
        <td class="px-4 py-3">${dpCell}</td>
        <td class="px-4 py-3">${t.approved_date ? toBuddhistDate(t.approved_date) : '-'}</td>
        <td class="px-4 py-3">${fCell}</td>
        <td class="px-4 py-3 text-xs text-gray-500">${t.remarks || ''}</td>
        ${isAdmin ? `<td class="px-4 py-3"><div class="flex gap-1"><button onclick="showEditTrackingModal('${id}')" class="text-blue-400 hover:text-blue-600" title="แก้ไข"><i data-lucide="pencil" class="w-4 h-4"></i></button><button onclick="deleteRecord('${id}')" class="text-red-400 hover:text-red-600" title="ลบ"><i data-lucide="trash-2" class="w-4 h-4"></i></button></div></td>` : ''}
      </tr>`
  }).join('') : '<tr><td colspan="12" class="px-4 py-8 text-center text-gray-400">ไม่มีข้อมูล</td></tr>'}</tbody>
    </table></div>
  </div>
  ${paginationHTML(total, APP.pagination.perPage, APP.pagination.page, 'changePage')}` : noYearSelectedMsg('ติดตามการส่งเกรดรายวิชา')}`;
}
function showAddGradeTrackingModal() {
  const subjects = getDataByType('subject');
  const subjectOptions = [...new Set(subjects.map(s => s.subject_name).filter(Boolean))].sort()
    .map(name => `<option value="${name.replace(/"/g, '&quot;')}">${name}</option>`).join('');
  const teachers = getDataByType('teacher');
  const teacherList = [...new Set(teachers.map(t => (t.name || '').trim()).filter(Boolean))].sort();
  const myName = (APP.currentUser && APP.currentUser.name || '').trim();
  const coordCheckboxes = teacherList.map(name => {
    const isMe = name === myName;
    return `<label class="flex items-center gap-2 p-1.5 rounded-lg hover:bg-gray-50 cursor-pointer ${isMe ? 'bg-blue-50' : ''}"><input type="checkbox" class="coord-check w-4 h-4 rounded border-gray-300 text-primary focus:ring-primary" value="${name.replace(/"/g, '&quot;')}" ${isMe ? 'checked' : ''}><span class="text-sm">${name}${isMe ? ' <span class="text-xs text-blue-500">(คุณ)</span>' : ''}</span></label>`;
  }).join('');
  showModal('เพิ่มข้อมูลติดตามการส่งเกรดรายวิชา', `
    <form id="addGradeTrackingForm" class="space-y-3">
      <div><label class="block text-xs text-gray-600 mb-1">ชื่อรายวิชา *</label><select name="subject_name" required class="w-full border rounded-xl px-3 py-2 text-sm"><option value="">-- เลือกรายวิชา --</option>${subjectOptions}</select></div>
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-xs text-gray-600 mb-1">ทฤษฎี/ปฏิบัติ</label><select name="theory_practice" class="w-full border rounded-xl px-3 py-2 text-sm"><option>ทฤษฎี</option><option>ปฏิบัติ</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">ชั้นปี</label><select name="year_level" class="w-full border rounded-xl px-3 py-2 text-sm"><option>1</option><option>2</option><option>3</option><option>4</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">ภาคการศึกษา</label><select name="semester" class="w-full border rounded-xl px-3 py-2 text-sm"><option value="1">1</option><option value="2">2</option><option value="3">ฤดูร้อน</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">ปีการศึกษา</label><input name="academic_year" value="2568" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      </div>
      <div>
        <label class="block text-xs text-gray-600 mb-1">ผู้ประสานงานรายวิชา (เลือกจากรายชื่อ หรือพิมพ์เพิ่มเติม)</label>
        <div class="max-h-40 overflow-y-auto border rounded-xl p-2 bg-gray-50 space-y-0.5">${coordCheckboxes || '<p class="text-xs text-gray-400 p-2">ยังไม่มีข้อมูลอาจารย์ในระบบ</p>'}</div>
        <input type="text" id="gradeTrackingCoordExtra" class="w-full border rounded-xl px-3 py-2 text-sm mt-2" placeholder="พิมพ์ชื่อเพิ่มเติม (คั่นด้วยเครื่องหมาย ,)">
        <input type="hidden" name="coordinator" id="gradeTrackingCoordHidden">
      </div>
      <div><label class="block text-xs text-gray-600 mb-1">หมายเหตุ</label><input name="remarks" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึก</button>
    </form>
  `);
  document.getElementById('addGradeTrackingForm').onsubmit = async (e) => {
    e.preventDefault();
    const checked = [...e.target.querySelectorAll('.coord-check:checked')].map(cb => cb.value);
    const extraText = (document.getElementById('gradeTrackingCoordExtra') || {}).value || '';
    const extraNames = extraText.split(',').map(s => s.trim()).filter(Boolean);
    const allNames = [...checked, ...extraNames].filter(Boolean);
    const coordHidden = document.getElementById('gradeTrackingCoordHidden');
    if (coordHidden) coordHidden.value = allNames.join(', ');
    await withLoading(e.target, async () => {
      const fd = new FormData(e.target);
      const obj = { type: 'grade_tracking', coordinator_check: 'รอ', academic_check: 'รอ', deputy_sign: 'รอ', approved_date: '', created_at: new Date().toISOString() };
      fd.forEach((v, k) => obj[k] = v);
      const r = await GSheetDB.create(obj);
      if (r.isOk) { showToast('เพิ่มสำเร็จ'); closeModal() } else showToast('เกิดข้อผิดพลาด', 'error');
    });
  };
}

// ======================== FILE TRACKING ========================
function fileTrackingPage() {
  const isAdmin = APP.currentRole === 'admin' || APP.currentRole === 'academic';
  const isExecutive = APP.currentRole === 'executive';
  const isClassTeacher = APP.currentRole === 'classTeacher';
  const canEdit = isAdmin || APP.currentRole === 'teacher' || APP.currentRole === 'classTeacher';
  let data = getDataByType('file_tracking').filter(t => t.subject_name && t.subject_name.trim());
  if (APP.currentRole === 'teacher') data = data.filter(t => t.coordinator === APP.currentUser.name);

  const selectedYear = APP.filters._fileTrackingYear || '';
  const allSubjects = getDataByType('subject');
  const subjectsFiltered = selectedYear ? allSubjects.filter(s => norm(s.academic_year) === selectedYear) : allSubjects;
  const dataForStats = selectedYear ? data.filter(t => norm(t.academic_year) === selectedYear) : data;

  let statsSection = '';
  let notSubmittedSection = '';
  if (selectedYear) {
    const trackedKeys = dataForStats.map(t => `${norm(t.subject_name)}|${norm(t.semester)}|${norm(t.academic_year)}`);
    const notSubmitted = subjectsFiltered.filter(s => !trackedKeys.includes(`${norm(s.subject_name)}|${norm(s.semester)}|${norm(s.academic_year)}`));

    const completed = dataForStats.filter(t => t.deputy_sign === 'เสร็จสิ้น').length;
    const inProgress = dataForStats.filter(t => (t.coordinator_check === 'เสร็จสิ้น' || t.academic_check === 'เสร็จสิ้น') && t.deputy_sign !== 'เสร็จสิ้น').length;
    const pending = dataForStats.length - completed - inProgress;

    statsSection = `<div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3 mt-3">
      ${statCard('alert-circle', 'ยังไม่ส่งแฟ้ม', notSubmitted.length, 'วิชา', 'bg-red-500')}
      ${statCard('clock', 'รอดำเนินการ', pending, 'วิชา', 'bg-yellow-500')}
      ${statCard('loader', 'กำลังดำเนินการ', inProgress, 'วิชา', 'bg-blue-500')}
      ${statCard('check-circle', 'เสร็จสิ้น', completed, 'วิชา', 'bg-green-500')}
    </div>`;
    if (notSubmitted.length) {
      notSubmittedSection = `<div class="bg-red-50 rounded-2xl p-4 border border-red-200 mb-4">
        <h3 class="font-bold text-red-700 mb-2 text-sm flex items-center gap-2"><i data-lucide="alert-triangle" class="w-4 h-4"></i>รายวิชาที่ยังไม่ส่งแฟ้ม (${notSubmitted.length} วิชา)</h3>
        <div class="flex flex-wrap gap-2">${notSubmitted.map(s => `<span class="px-3 py-1 bg-white border border-red-200 rounded-lg text-xs text-red-700">${s.subject_code ? s.subject_code + ' ' : ''}${s.subject_name || ''} <span class="text-gray-400">(ภาค ${s.semester || ''})</span></span>`).join('')}</div>
      </div>`;
    }
  }

  const allYears = [...new Set([...allSubjects.map(s => s.academic_year), ...data.map(t => t.academic_year)].filter(Boolean))].sort();

  if (selectedYear) data = data.filter(t => norm(t.academic_year) === selectedYear);
  data = applyFilters(data);
  const total = data.length; const paged = paginate(data);

  return `<div class="flex flex-wrap items-center justify-between gap-3 mb-4">
    <h2 class="text-xl font-bold text-gray-800"><i data-lucide="folder-check" class="w-6 h-6 inline mr-2"></i>ติดตามการส่งแฟ้มรายวิชา</h2>
    ${canEdit ? `<button onclick="showAddFileTrackingModal()" class="flex items-center gap-2 px-4 py-2 bg-primary text-white rounded-xl hover:bg-primaryDark text-sm"><i data-lucide="plus" class="w-4 h-4"></i>เพิ่มข้อมูล</button>` : ''}
  </div>
  <div class="bg-white rounded-2xl p-4 border border-blue-100 mb-4">
    <div class="flex items-center gap-3 mb-3">
      <label class="text-sm font-medium text-gray-700">ปีการศึกษา:</label>
      <select onchange="APP.filters._fileTrackingYear=this.value;renderCurrentPage()" class="border border-gray-200 rounded-xl px-3 py-2 text-sm">
        <option value="">-- เลือกปีการศึกษา --</option>
        ${allYears.map(y => `<option value="${y}" ${selectedYear === y ? 'selected' : ''}>${y}</option>`).join('')}
      </select>
    </div>
    ${statsSection}
  </div>
  ${selectedYear ? `${notSubmittedSection}
  ${filterBar({ yearLevel: true })}
  <div class="bg-white rounded-2xl border border-blue-100 overflow-hidden">
    <div class="overflow-x-auto"><table class="w-full text-sm">
      <thead><tr class="bg-surface text-left"><th class="px-4 py-3 font-semibold">รายวิชา</th><th class="px-4 py-3 font-semibold">ทฤษฎี/ปฏิบัติ</th><th class="px-4 py-3 font-semibold">ชั้นปี</th><th class="px-4 py-3 font-semibold">ภาค/ปี</th><th class="px-4 py-3 font-semibold">ผู้ประสานงาน</th><th class="px-4 py-3 font-semibold">อ.ประจำชั้นตรวจ</th><th class="px-4 py-3 font-semibold">วิชาการเสนอ</th><th class="px-4 py-3 font-semibold">รอง ผอ.ลงนาม</th><th class="px-4 py-3 font-semibold">วันอนุมัติ</th><th class="px-4 py-3 font-semibold">ไฟล์</th><th class="px-4 py-3 font-semibold">หมายเหตุ</th>${isAdmin ? '<th class="px-4 py-3"></th>' : ''}</tr></thead>
      <tbody>${paged.length ? paged.map(t => {
    const sb = (s) => s === 'เสร็จสิ้น' ? '<span class="px-2 py-1 rounded-full text-xs bg-green-100 text-green-700">✓ เสร็จสิ้น</span>' : s === 'ส่งกลับแก้ไข' ? '<span class="px-2 py-1 rounded-full text-xs bg-red-100 text-red-700">↩ ส่งกลับแก้ไข</span>' : s === 'กำลังดำเนินการ' ? '<span class="px-2 py-1 rounded-full text-xs bg-yellow-100 text-yellow-700">⏳ ดำเนินการ</span>' : '<span class="px-2 py-1 rounded-full text-xs bg-gray-100 text-gray-500">รอ</span>';
    const isLate = t.is_late === 'ใช่' || t.is_late === 'late';
    const id = t.__backendId;

    const ctCell = (isClassTeacher && (t.coordinator_check === 'รอ' || t.coordinator_check === 'ส่งกลับแก้ไข' || !t.coordinator_check))
      ? `<div class="flex flex-col gap-1"><button onclick="updateTrackingField('${id}','coordinator_check','เสร็จสิ้น')" class="px-2 py-1 bg-green-500 hover:bg-green-600 text-white rounded-lg text-xs">✓ ตรวจสอบแล้ว</button><button onclick="updateTrackingField('${id}','coordinator_check','ส่งกลับแก้ไข')" class="px-2 py-1 bg-orange-500 hover:bg-orange-600 text-white rounded-lg text-xs">↩ ส่งกลับแก้ไข</button></div>`
      : (isAdmin ? `<select onchange="updateTrackingField('${id}','coordinator_check',this.value)" class="text-xs border rounded px-1 py-0.5"><option ${t.coordinator_check === 'รอ' ? 'selected' : ''}>รอ</option><option ${t.coordinator_check === 'เสร็จสิ้น' ? 'selected' : ''}>เสร็จสิ้น</option><option ${t.coordinator_check === 'ส่งกลับแก้ไข' ? 'selected' : ''}>ส่งกลับแก้ไข</option></select>` : sb(t.coordinator_check));

    const acCell = (isAdmin && t.coordinator_check === 'เสร็จสิ้น' && (t.academic_check === 'รอ' || t.academic_check === 'ส่งกลับแก้ไข' || !t.academic_check))
      ? `<div class="flex flex-col gap-1"><button onclick="updateTrackingField('${id}','academic_check','เสร็จสิ้น')" class="px-2 py-1 bg-green-500 hover:bg-green-600 text-white rounded-lg text-xs">✓ ส่งเสนอรองฯ แล้ว</button><button onclick="updateTrackingField('${id}','academic_check','ส่งกลับแก้ไข')" class="px-2 py-1 bg-orange-500 hover:bg-orange-600 text-white rounded-lg text-xs">↩ ส่งกลับแก้ไข</button></div>`
      : (isAdmin ? `<select onchange="updateTrackingField('${id}','academic_check',this.value)" class="text-xs border rounded px-1 py-0.5"><option ${t.academic_check === 'รอ' ? 'selected' : ''}>รอ</option><option ${t.academic_check === 'เสร็จสิ้น' ? 'selected' : ''}>เสร็จสิ้น</option><option ${t.academic_check === 'ส่งกลับแก้ไข' ? 'selected' : ''}>ส่งกลับแก้ไข</option></select>` : sb(t.academic_check));

    const dpCell = (isExecutive && t.academic_check === 'เสร็จสิ้น' && (t.deputy_sign === 'รอ' || t.deputy_sign === 'ส่งกลับแก้ไข' || !t.deputy_sign))
      ? `<div class="flex flex-col gap-1"><button onclick="updateTrackingField('${id}','deputy_sign','เสร็จสิ้น')" class="px-2 py-1 bg-green-500 hover:bg-green-600 text-white rounded-lg text-xs">✓ ลงนามแล้ว</button><button onclick="updateTrackingField('${id}','deputy_sign','ส่งกลับแก้ไข')" class="px-2 py-1 bg-orange-500 hover:bg-orange-600 text-white rounded-lg text-xs">↩ ส่งกลับแก้ไข</button></div>`
      : (isAdmin ? `<select onchange="updateTrackingField('${id}','deputy_sign',this.value)" class="text-xs border rounded px-1 py-0.5"><option ${t.deputy_sign === 'รอ' ? 'selected' : ''}>รอ</option><option ${t.deputy_sign === 'เสร็จสิ้น' ? 'selected' : ''}>เสร็จสิ้น</option><option ${t.deputy_sign === 'ส่งกลับแก้ไข' ? 'selected' : ''}>ส่งกลับแก้ไข</option></select>` : sb(t.deputy_sign));

    let fCell = '<span class="text-xs text-gray-400">-</span>';
    if (t.deputy_sign === 'เสร็จสิ้น') {
      if (t.file_link) {
        fCell = `<a href="${t.file_link}" target="_blank" class="px-2 py-1 bg-blue-100 text-blue-700 rounded-lg text-xs hover:bg-blue-200 inline-flex items-center gap-1"><i data-lucide="file-text" class="w-3 h-3"></i>ดาวน์โหลด</a>`;
        if (isAdmin) fCell += ` <button onclick="promptTrackingFileLink('${id}')" class="text-blue-400 hover:text-blue-600 ml-1" title="แก้ไขลิงก์"><i data-lucide="pencil" class="w-3 h-3"></i></button>`;
      } else if (isAdmin) {
        fCell = `<button onclick="promptTrackingFileLink('${id}')" class="px-2 py-1 bg-emerald-500 hover:bg-emerald-600 text-white rounded-lg text-xs inline-flex items-center gap-1"><i data-lucide="upload" class="w-3 h-3"></i>เพิ่มไฟล์ PDF</button>`;
      }
    }

    return `<tr class="border-t hover:bg-gray-50 ${isLate ? 'bg-red-50' : ''}">
        <td class="px-4 py-3 font-medium">${t.subject_name || ''}${isLate ? ' <span class="text-xs text-red-500 font-normal">(ส่งช้า)</span>' : ''}</td>
        <td class="px-4 py-3">${t.theory_practice || ''}</td>
        <td class="px-4 py-3">${t.year_level || ''}</td>
        <td class="px-4 py-3">${semLabel(t.semester)}/${t.academic_year || ''}</td>
        <td class="px-4 py-3 text-xs">${t.coordinator || ''}</td>
        <td class="px-4 py-3">${ctCell}</td>
        <td class="px-4 py-3">${acCell}</td>
        <td class="px-4 py-3">${dpCell}</td>
        <td class="px-4 py-3">${t.approved_date ? toBuddhistDate(t.approved_date) : '-'}</td>
        <td class="px-4 py-3">${fCell}</td>
        <td class="px-4 py-3 text-xs text-gray-500">${t.remarks || ''}</td>
        ${isAdmin ? `<td class="px-4 py-3"><div class="flex gap-1"><button onclick="showEditTrackingModal('${id}')" class="text-blue-400 hover:text-blue-600" title="แก้ไข"><i data-lucide="pencil" class="w-4 h-4"></i></button><button onclick="deleteRecord('${id}')" class="text-red-400 hover:text-red-600" title="ลบ"><i data-lucide="trash-2" class="w-4 h-4"></i></button></div></td>` : ''}
      </tr>`
  }).join('') : '<tr><td colspan="12" class="px-4 py-8 text-center text-gray-400">ไม่มีข้อมูล</td></tr>'}</tbody>
    </table></div>
  </div>
  ${paginationHTML(total, APP.pagination.perPage, APP.pagination.page, 'changePage')}` : noYearSelectedMsg('ติดตามการส่งแฟ้มรายวิชา')}`;
}
function showAddFileTrackingModal() {
  const subjects = getDataByType('subject');
  const subjectOptions = [...new Set(subjects.map(s => s.subject_name).filter(Boolean))].sort()
    .map(name => `<option value="${name.replace(/"/g, '&quot;')}">${name}</option>`).join('');
  const teachers = getDataByType('teacher');
  const teacherList = [...new Set(teachers.map(t => (t.name || '').trim()).filter(Boolean))].sort();
  const myName = (APP.currentUser && APP.currentUser.name || '').trim();
  const coordCheckboxes = teacherList.map(name => {
    const isMe = name === myName;
    return `<label class="flex items-center gap-2 p-1.5 rounded-lg hover:bg-gray-50 cursor-pointer ${isMe ? 'bg-blue-50' : ''}"><input type="checkbox" class="coord-check w-4 h-4 rounded border-gray-300 text-primary focus:ring-primary" value="${name.replace(/"/g, '&quot;')}" ${isMe ? 'checked' : ''}><span class="text-sm">${name}${isMe ? ' <span class="text-xs text-blue-500">(คุณ)</span>' : ''}</span></label>`;
  }).join('');
  showModal('เพิ่มข้อมูลติดตามการส่งแฟ้มรายวิชา', `
    <form id="addFileTrackingForm" class="space-y-3">
      <div><label class="block text-xs text-gray-600 mb-1">ชื่อรายวิชา *</label><select name="subject_name" required class="w-full border rounded-xl px-3 py-2 text-sm"><option value="">-- เลือกรายวิชา --</option>${subjectOptions}</select></div>
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-xs text-gray-600 mb-1">ทฤษฎี/ปฏิบัติ</label><select name="theory_practice" class="w-full border rounded-xl px-3 py-2 text-sm"><option>ทฤษฎี</option><option>ปฏิบัติ</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">ชั้นปี</label><select name="year_level" class="w-full border rounded-xl px-3 py-2 text-sm"><option>1</option><option>2</option><option>3</option><option>4</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">ภาคการศึกษา</label><select name="semester" class="w-full border rounded-xl px-3 py-2 text-sm"><option value="1">1</option><option value="2">2</option><option value="3">ฤดูร้อน</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">ปีการศึกษา</label><input name="academic_year" value="2568" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      </div>
      <div>
        <label class="block text-xs text-gray-600 mb-1">ผู้ประสานงานรายวิชา (เลือกจากรายชื่อ หรือพิมพ์เพิ่มเติม)</label>
        <div class="max-h-40 overflow-y-auto border rounded-xl p-2 bg-gray-50 space-y-0.5">${coordCheckboxes || '<p class="text-xs text-gray-400 p-2">ยังไม่มีข้อมูลอาจารย์ในระบบ</p>'}</div>
        <input type="text" id="fileTrackingCoordExtra" class="w-full border rounded-xl px-3 py-2 text-sm mt-2" placeholder="พิมพ์ชื่อเพิ่มเติม (คั่นด้วยเครื่องหมาย ,)">
        <input type="hidden" name="coordinator" id="fileTrackingCoordHidden">
      </div>
      <div><label class="block text-xs text-gray-600 mb-1">หมายเหตุ</label><input name="remarks" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึก</button>
    </form>
  `);
  document.getElementById('addFileTrackingForm').onsubmit = async (e) => {
    e.preventDefault();
    const checked = [...e.target.querySelectorAll('.coord-check:checked')].map(cb => cb.value);
    const extraText = (document.getElementById('fileTrackingCoordExtra') || {}).value || '';
    const extraNames = extraText.split(',').map(s => s.trim()).filter(Boolean);
    const allNames = [...checked, ...extraNames].filter(Boolean);
    const coordHidden = document.getElementById('fileTrackingCoordHidden');
    if (coordHidden) coordHidden.value = allNames.join(', ');
    await withLoading(e.target, async () => {
      const fd = new FormData(e.target);
      const obj = { type: 'file_tracking', coordinator_check: 'รอ', academic_check: 'รอ', deputy_sign: 'รอ', approved_date: '', created_at: new Date().toISOString() };
      fd.forEach((v, k) => obj[k] = v);
      const r = await GSheetDB.create(obj);
      if (r.isOk) { showToast('เพิ่มสำเร็จ'); closeModal() } else showToast('เกิดข้อผิดพลาด', 'error');
    });
  };
}

// ======================== LEAVE ========================
// Helper: Check if all coordinator approvals for same student + same date are done
function allCoordinatorsApproved(leaveRec) {
  const allLeaves = getDataByType('leave');
  const sameBatch = allLeaves.filter(l =>
    l.name === leaveRec.name &&
    l.leave_date === leaveRec.leave_date &&
    l.leave_status !== 'ปฏิเสธ'
  );
  if (sameBatch.length <= 1) return leaveRec.coordinator_approval === 'อนุมัติ';
  return sameBatch.every(l => l.coordinator_approval === 'อนุมัติ');
}

// Helper: Build leave percent summary table HTML
function leavePercentSummaryHTML(leaveRecords, groupBy) {
  // groupBy: 'subject' (for student/teacher) or 'student_subject' (for classTeacher)
  const map = {};
  leaveRecords.forEach(l => {
    const key = groupBy === 'student_subject' ? (l.name + '|||' + l.subject_name) : l.subject_name;
    if (!map[key]) map[key] = { name: l.name || '', subject: l.subject_name || '', hours: 0, percent: 0, count: 0 };
    map[key].hours += Number(l.leave_hours) || 0;
    map[key].count++;
    const pct = Number(l.leave_percent) || 0;
    if (pct > map[key].percent) map[key].percent = pct;
  });
  const entries = Object.values(map);
  if (!entries.length) return '';
  const showName = groupBy === 'student_subject';
  const rows = entries.map(info => {
    const pct = info.percent;
    const colorClass = pct >= 20 ? 'bg-red-100 text-red-700 font-bold' : pct >= 15 ? 'bg-yellow-100 text-yellow-700 font-semibold' : 'bg-green-100 text-green-700';
    return `<tr class="border-t hover:bg-gray-50">
      ${showName ? `<td class="px-4 py-2 text-sm">${info.name}</td>` : ''}
      <td class="px-4 py-2 text-sm">${info.subject}</td>
      <td class="px-4 py-2 text-sm text-center">${info.hours}</td>
      <td class="px-4 py-2 text-sm text-center"><span class="px-2 py-1 rounded-full text-xs ${colorClass}">${pct}%</span></td>
      <td class="px-4 py-2 text-sm text-center">${info.count}</td>
    </tr>`;
  }).join('');
  return `<div class="bg-white rounded-2xl p-5 border border-blue-100 mb-4">
    <h3 class="font-bold mb-3 flex items-center gap-2"><i data-lucide="bar-chart-3" class="w-5 h-5 text-primary"></i>สรุปเปอร์เซ็นต์การลาแต่ละรายวิชา</h3>
    <div class="overflow-x-auto"><table class="w-full text-sm">
      <thead><tr class="bg-surface text-left">
        ${showName ? '<th class="px-4 py-2 font-semibold">ชื่อนักศึกษา</th>' : ''}
        <th class="px-4 py-2 font-semibold">รายวิชา</th>
        <th class="px-4 py-2 font-semibold text-center">ชม.ลารวม</th>
        <th class="px-4 py-2 font-semibold text-center">%ลา</th>
        <th class="px-4 py-2 font-semibold text-center">จำนวนครั้ง</th>
      </tr></thead>
      <tbody>${rows}</tbody>
    </table></div>
    <div class="mt-3 flex flex-wrap gap-3 text-xs text-gray-500">
      <span class="flex items-center gap-1"><span class="w-3 h-3 rounded-full bg-red-500"></span> ≥ 20% เกินเกณฑ์</span>
      <span class="flex items-center gap-1"><span class="w-3 h-3 rounded-full bg-yellow-500"></span> 15-19% ใกล้เกินเกณฑ์</span>
      <span class="flex items-center gap-1"><span class="w-3 h-3 rounded-full bg-green-500"></span> < 15% ปกติ</span>
    </div>
  </div>`;
}

// Filter handlers for leave page (admin/executive view)
function setLeaveYearLevel(v) {
  APP.filters._leaveYearLevel = v;
  // Reset student filter when year level changes (since student list depends on year)
  APP.filters._leaveStudent = '';
  APP.pagination.page = 1;
  renderCurrentPage();
}
function setLeaveStudent(v) {
  APP.filters._leaveStudent = v;
  APP.pagination.page = 1;
  renderCurrentPage();
}
function clearLeaveFilters() {
  APP.filters._leaveYearLevel = '';
  APP.filters._leaveStudent = '';
  APP.pagination.page = 1;
  renderCurrentPage();
}

// SVG pie chart for a specific student's leave hours grouped by subject
function renderLeavePieChart(leaveRecords, studentName) {
  const byS = {};
  leaveRecords.forEach(l => {
    if (!l.subject_name) return;
    const k = l.subject_name;
    byS[k] = (byS[k] || 0) + (Number(l.leave_hours) || 0);
  });
  const entries = Object.entries(byS).filter(([_, v]) => v > 0);
  if (entries.length === 0) {
    return `<div class="bg-blue-50 rounded-2xl p-4 mb-4 text-sm text-blue-700"><i data-lucide="info" class="w-4 h-4 inline"></i> ${studentName} ยังไม่มีข้อมูลการลา</div>`;
  }
  const total = entries.reduce((s, [_, v]) => s + v, 0);
  const colors = ['#1e6fba', '#0ea5e9', '#10b981', '#f59e0b', '#8b5cf6', '#ef4444', '#ec4899', '#14b8a6', '#6366f1', '#f97316'];
  const cx = 110, cy = 110, r = 95;
  let cumAngle = -Math.PI / 2;
  let pathHTML = '';
  let legendHTML = '';
  entries.forEach(([name, hours], idx) => {
    const angle = (hours / total) * 2 * Math.PI;
    const color = colors[idx % colors.length];
    if (entries.length === 1) {
      pathHTML += `<circle cx="${cx}" cy="${cy}" r="${r}" fill="${color}"/>`;
    } else {
      const x1 = cx + r * Math.cos(cumAngle);
      const y1 = cy + r * Math.sin(cumAngle);
      cumAngle += angle;
      const x2 = cx + r * Math.cos(cumAngle);
      const y2 = cy + r * Math.sin(cumAngle);
      const largeArc = angle > Math.PI ? 1 : 0;
      pathHTML += `<path d="M ${cx} ${cy} L ${x1.toFixed(2)} ${y1.toFixed(2)} A ${r} ${r} 0 ${largeArc} 1 ${x2.toFixed(2)} ${y2.toFixed(2)} Z" fill="${color}" stroke="white" stroke-width="2"/>`;
    }
    const pct = ((hours / total) * 100).toFixed(1);
    legendHTML += `<div class="flex items-center gap-2 text-sm py-1"><span class="w-3 h-3 rounded flex-shrink-0" style="background:${color}"></span><span class="flex-1">${name}</span><span class="text-gray-600 font-medium">${hours} ชม. (${pct}%)</span></div>`;
  });
  return `<div class="bg-white rounded-2xl p-5 border border-blue-100 mb-4">
    <h3 class="font-bold mb-4 flex items-center gap-2"><i data-lucide="pie-chart" class="w-5 h-5 text-primary"></i>กราฟวงกลมการลาแยกตามรายวิชา — ${studentName}</h3>
    <div class="flex flex-wrap gap-6 items-center">
      <svg width="220" height="220" viewBox="0 0 220 220" class="flex-shrink-0">${pathHTML}</svg>
      <div class="flex-1 min-w-[260px]">
        ${legendHTML}
        <div class="border-t pt-2 mt-2 font-semibold text-gray-800 flex justify-between"><span>รวมทั้งหมด</span><span>${total} ชม.</span></div>
      </div>
    </div>
  </div>`;
}

function leavePage() {
  const isStudent = APP.currentRole === 'student';
  const isAdmin = APP.currentRole === 'admin' || APP.currentRole === 'academic';
  const isTeacher = APP.currentRole === 'teacher';
  const isClassTeacher = APP.currentRole === 'classTeacher';
  const isExecutive = APP.currentRole === 'executive';
  const canApprove = isTeacher || isClassTeacher || isExecutive;
  const canEdit = isAdmin || isTeacher || isClassTeacher;

  let data = getDataByType('leave');

  if (isStudent && APP.currentUser.data) data = data.filter(l => l.name === APP.currentUser.data.name);
  if (isTeacher) {
    const myName = (APP.currentUser.name || '').trim();
    data = data.filter(l => {
      const sub = getDataByType('subject').find(s => s.subject_name === l.subject_name);
      if (!sub || !sub.coordinator) return false;
      // Support multiple coordinators separated by comma / และ / and
      const coords = String(sub.coordinator).split(/[,;|]|\sและ\s|\sand\s/).map(c => c.trim()).filter(Boolean);
      return coords.some(c => c === myName || c.includes(myName) || myName.includes(c));
    });
  }
  if (isClassTeacher) {
    const yr = APP.currentUser.responsible_year || '1';
    const stuNames = getDataByType('student')
      .filter(s => norm(s.year_level) === norm(yr))
      .map(s => (s.name || '').trim());
    const stuNameSet = new Set(stuNames);
    data = data.filter(l => {
      const lname = (l.name || '').trim();
      if (stuNameSet.has(lname)) return true;
      // Fallback: also check if class_teacher field on the leave matches current user's name
      if (l.class_teacher && (l.class_teacher || '').trim() === (APP.currentUser.name || '').trim()) return true;
      return false;
    });
  }
  data = applyFilters(data);

  // Admin/Executive: extra year-level + student filters (joined to student records)
  const showAdminFilters = isAdmin || isExecutive;
  let leaveYearLevel = '';
  let leaveStudentName = '';
  if (showAdminFilters) {
    leaveYearLevel = APP.filters._leaveYearLevel || '';
    leaveStudentName = APP.filters._leaveStudent || '';
    if (leaveYearLevel) {
      const stuNameSet = new Set(
        getDataByType('student')
          .filter(s => norm(s.year_level) === norm(leaveYearLevel))
          .map(s => (s.name || '').trim())
      );
      data = data.filter(l => stuNameSet.has((l.name || '').trim()));
    }
    if (leaveStudentName) {
      data = data.filter(l => (l.name || '').trim() === leaveStudentName);
    }
  }

  const total = data.length; const paged = paginate(data);

  // Pending-approval count for current role
  let pendingCount = 0;
  if (isTeacher) pendingCount = data.filter(l => (l.coordinator_approval || 'รอ') === 'รอ' && l.leave_status !== 'ปฏิเสธ').length;
  if (isClassTeacher) pendingCount = data.filter(l => allCoordinatorsApproved(l) && (l.class_teacher_approval || 'รอ') === 'รอ' && l.leave_status !== 'ปฏิเสธ').length;
  if (isExecutive) pendingCount = data.filter(l => (l.coordinator_approval === 'อนุมัติ') && (l.class_teacher_approval === 'อนุมัติ') && (l.deputy_approval || 'รอ') === 'รอ' && l.leave_status !== 'ปฏิเสธ').length;

  let form = '';
  if (isStudent) {
    const stuYearLevel = APP.currentUser.data?.year_level || '';
    const allSubjects = getDataByType('subject');
    const subjects = stuYearLevel ? allSubjects.filter(s => norm(s.year_level) === norm(stuYearLevel)) : allSubjects;
    // Auto-detect class teacher from teacher records by responsible_year matching student's year_level
    const allTeachers = getDataByType('teacher');
    const classTeacherRec = stuYearLevel
      ? allTeachers.find(t => norm(t.responsible_year) === norm(stuYearLevel))
      : null;
    const classTeacherName = classTeacherRec ? classTeacherRec.name : '';
    form = `<div class="bg-white rounded-2xl p-5 border border-blue-100 mb-4">
      <h3 class="font-bold mb-3">กรอกข้อมูลการลา</h3>
      <form id="leaveForm" class="space-y-3">
        <div class="grid grid-cols-1 md:grid-cols-2 gap-3">
          <div><label class="block text-xs text-gray-600 mb-1">ชื่อ-สกุล</label><input name="name" value="${APP.currentUser.data?.name || ''}" readonly class="w-full border rounded-xl px-3 py-2 text-sm bg-gray-50"></div>
          <div><label class="block text-xs text-gray-600 mb-1">อาจารย์ประจำชั้น (ชั้นปี ${stuYearLevel || '-'})</label><input name="class_teacher" value="${classTeacherName}" readonly class="w-full border rounded-xl px-3 py-2 text-sm bg-gray-50" placeholder="${classTeacherName ? '' : 'ยังไม่มีอาจารย์ประจำชั้นในระบบ'}"></div>
          <div><label class="block text-xs text-gray-600 mb-1">ภาคการศึกษา</label><select name="semester" class="w-full border rounded-xl px-3 py-2 text-sm"><option value="1">1</option><option value="2">2</option></select></div>
          <div><label class="block text-xs text-gray-600 mb-1">ปีการศึกษา</label><input name="academic_year" value="2568" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
          <div class="md:col-span-2"><label class="block text-xs text-gray-600 mb-1">ประเภทการลา</label><select name="leave_type" required class="w-full border rounded-xl px-3 py-2 text-sm" onchange="onLeaveTypeChange(this.value)"><option value="">เลือก</option><option value="ลาป่วย">ลาป่วย</option><option value="ลากิจ">ลากิจ</option><option value="ลาพบแพทย์">ลาพบแพทย์</option></select></div>
        </div>
        <div>
          <label class="block text-xs text-gray-600 mb-1">วันที่ลา (เลือกได้หลายวัน)</label>
          <div id="leaveDateList" class="space-y-2">
            <div class="flex items-center gap-2 leave-date-row">
              <input type="date" class="leave-date-input flex-1 border rounded-xl px-3 py-2 text-sm" required onchange="validateLeaveDate()">
              <span class="be-display text-xs text-gray-500 min-w-[90px]"></span>
              <button type="button" onclick="removeLeaveDateRow(this)" class="px-2 py-1 text-red-500 hover:bg-red-50 rounded-lg text-xs"><i data-lucide="x" class="w-4 h-4"></i></button>
            </div>
          </div>
          <button type="button" onclick="addLeaveDateRow()" class="mt-2 text-xs text-blue-600 hover:underline flex items-center gap-1"><i data-lucide="plus" class="w-3 h-3"></i> เพิ่มวันที่ลา</button>
          <input type="hidden" name="leave_date" id="leaveDateHidden">
          <p class="text-xs text-gray-400 mt-1">วันที่จะแสดงในรูปแบบ พ.ศ. (เช่น 07/05/2569)</p>
        </div>
        <div>
          <label class="block text-xs text-gray-600 mb-2 font-semibold">เลือกรายวิชาที่ต้องการลา (เลือกได้หลายวิชา กรอกจำนวนชั่วโมงที่ละวิชา)</label>
          <div id="leaveSubjectList" class="space-y-1 max-h-64 overflow-y-auto border rounded-xl p-3 bg-gray-50">
            ${subjects.map(s => `<div class="flex items-start gap-3 p-2 rounded-lg hover:bg-white transition leave-subject-row">
              <input type="checkbox" class="leave-subject-check w-4 h-4 mt-1 rounded border-gray-300 text-primary focus:ring-primary" value="${(s.subject_name || '').replace(/"/g, '&quot;')}" data-coordinator="${(s.coordinator || '').replace(/"/g, '&quot;')}" onchange="toggleLeaveSubjectHours(this)">
              <div class="flex-1">
                <div class="text-sm">${s.subject_code ? s.subject_code + ' ' : ''}${s.subject_name || ''}</div>
                <div class="text-xs text-gray-500 mt-0.5"><i data-lucide="user" class="w-3 h-3 inline"></i> อ.ผู้ประสานรายวิชา: <span class="font-medium text-gray-700">${s.coordinator || '-'}</span></div>
              </div>
              <div class="leave-subject-hours hidden flex items-center gap-1">
                <input type="number" min="1" class="w-20 border rounded-lg px-2 py-1 text-sm text-center leave-hours-input" placeholder="ชม.">
                <span class="text-xs text-gray-400">ชม.</span>
              </div>
            </div>`).join('')}
          </div>
          <p class="text-xs text-gray-400 mt-1">เลือกรายวิชาแล้วกรอกจำนวนชั่วโมงที่ลาในแต่ละวิชา</p>
        </div>
        <div>
          <label class="block text-xs text-gray-600 mb-1">เหตุผลการลา <span class="text-red-500">*</span></label>
          <textarea name="leave_reason" required rows="3" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="กรุณากรอกเหตุผลการลา"></textarea>
        </div>
        <div id="leaveExtra" class="space-y-3"></div>
        <div id="leaveValidation" class="text-red-500 text-xs hidden"></div>
        <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">ส่งใบลา</button>
      </form>
    </div>`;
  }

  const pendingBanner = canApprove && pendingCount > 0
    ? `<div class="bg-amber-50 border border-amber-200 rounded-2xl p-3 mb-4 flex items-center gap-3">
        <div class="w-10 h-10 bg-amber-500 rounded-full flex items-center justify-center text-white font-bold">${pendingCount}</div>
        <div>
          <p class="text-sm font-semibold text-amber-900">มีใบลา ${pendingCount} รายการรอการอนุมัติของคุณ</p>
          <p class="text-xs text-amber-700">${isTeacher ? 'คุณคืออาจารย์ผู้ประสานรายวิชา — โปรดกรอก % การลาและกดอนุมัติ' : isClassTeacher ? 'อาจารย์ผู้ประสานได้อนุมัติแล้ว — รอคุณอนุมัติเป็นลำดับถัดไป' : 'ผ่านการอนุมัติจาก ปสน. และ ปจช. แล้ว — รอผู้บริหารอนุมัติเป็นขั้นสุดท้าย'}</p>
        </div>
      </div>`
    : '';

  // Leave percent summary tables
  let summaryTable = '';
  if (isStudent && APP.currentUser.data) {
    summaryTable = leavePercentSummaryHTML(data, 'subject');
  } else if (isTeacher) {
    summaryTable = leavePercentSummaryHTML(data, 'student_subject');
  } else if (isClassTeacher) {
    summaryTable = leavePercentSummaryHTML(data, 'student_subject');
  }

  // Admin/Executive: year-level + student filter card + pie chart
  let adminFilterCard = '';
  let pieChartCard = '';
  if (showAdminFilters) {
    // Build student dropdown options based on selected year level
    const allStudents = getDataByType('student');
    const studentsForDropdown = leaveYearLevel
      ? allStudents.filter(s => norm(s.year_level) === norm(leaveYearLevel))
      : allStudents;
    const studentOptions = studentsForDropdown
      .map(s => (s.name || '').trim())
      .filter(Boolean)
      .sort()
      .map(name => `<option value="${name.replace(/"/g, '&quot;')}" ${leaveStudentName === name ? 'selected' : ''}>${name}</option>`)
      .join('');
    adminFilterCard = `<div class="bg-white rounded-2xl p-4 border border-blue-100 mb-4">
      <div class="flex flex-wrap items-end gap-3">
        <div>
          <label class="block text-xs text-gray-600 mb-1">ชั้นปี</label>
          <select onchange="setLeaveYearLevel(this.value)" class="border border-gray-200 rounded-xl px-3 py-2 text-sm">
            <option value="">ทุกชั้นปี</option>
            <option value="1" ${leaveYearLevel === '1' ? 'selected' : ''}>ชั้นปี 1</option>
            <option value="2" ${leaveYearLevel === '2' ? 'selected' : ''}>ชั้นปี 2</option>
            <option value="3" ${leaveYearLevel === '3' ? 'selected' : ''}>ชั้นปี 3</option>
            <option value="4" ${leaveYearLevel === '4' ? 'selected' : ''}>ชั้นปี 4</option>
          </select>
        </div>
        <div class="flex-1 min-w-[200px]">
          <label class="block text-xs text-gray-600 mb-1">นักศึกษา (เลือกเพื่อดูกราฟวงกลม)</label>
          <select onchange="setLeaveStudent(this.value)" class="w-full border border-gray-200 rounded-xl px-3 py-2 text-sm">
            <option value="">${leaveYearLevel ? 'ทุกคนในชั้นปี ' + leaveYearLevel : 'ทุกคน'}</option>
            ${studentOptions}
          </select>
        </div>
        ${(leaveYearLevel || leaveStudentName) ? `<button onclick="clearLeaveFilters()" class="px-3 py-2 text-xs text-gray-600 hover:bg-gray-100 rounded-xl border border-gray-200"><i data-lucide="x" class="w-3 h-3 inline"></i> ล้างตัวกรอง</button>` : ''}
      </div>
    </div>`;
    if (leaveStudentName) {
      pieChartCard = renderLeavePieChart(data, leaveStudentName);
    }
  }

  return `<h2 class="text-xl font-bold text-gray-800 mb-4"><i data-lucide="calendar-off" class="w-6 h-6 inline mr-2"></i>ระบบการลาของนักศึกษา</h2>
  ${isAdmin ? `<div class="flex gap-2 mb-4"><button onclick="showAddLeaveModal()" class="flex items-center gap-2 px-4 py-2 bg-primary text-white rounded-xl hover:bg-primaryDark text-sm"><i data-lucide="plus" class="w-4 h-4"></i>เพิ่มข้อมูลการลา</button>${csvUploadBtn('leave', 'name,subject_name,leave_hours,leave_percent,semester,academic_year,leave_date,leave_type')}</div>` : ''}
  ${pendingBanner}
  ${form}
  ${adminFilterCard}
  ${pieChartCard}
  ${summaryTable}
  ${filterBar()}
  <div class="bg-white rounded-2xl border border-blue-100 overflow-hidden">
    <div class="overflow-x-auto"><table class="w-full text-sm">
      <thead><tr class="bg-surface text-left"><th class="px-4 py-3 font-semibold">ชื่อ-สกุล</th><th class="px-4 py-3 font-semibold">รายวิชา</th><th class="px-4 py-3 font-semibold">ประเภท</th><th class="px-4 py-3 font-semibold">เหตุผลการลา</th><th class="px-4 py-3 font-semibold">ชม.</th><th class="px-4 py-3 font-semibold">%ลา</th><th class="px-4 py-3 font-semibold">วันที่</th><th class="px-4 py-3 font-semibold">ภาค/ปี</th><th class="px-4 py-3 font-semibold">สถานะ</th>${(canApprove || isStudent || isAdmin) ? '<th class="px-4 py-3 font-semibold">' + (canApprove ? 'การอนุมัติ' : 'ขั้นการอนุมัติ') + '</th>' : ''}${isAdmin ? '<th class="px-4 py-3"></th>' : ''}</tr></thead>
      <tbody>${paged.length ? paged.map(l => {
    const getStatusBadge = (status) => {
      if (status === 'รออนุมัติ') return '<span class="px-2 py-1 rounded-full text-xs bg-yellow-100 text-yellow-700">⏳ รออนุมัติ</span>';
      if (status === 'อนุมัติแล้ว') return '<span class="px-2 py-1 rounded-full text-xs bg-green-100 text-green-700">✓ อนุมัติแล้ว</span>';
      if (status === 'ปฏิเสธ') return '<span class="px-2 py-1 rounded-full text-xs bg-red-100 text-red-700">✕ ปฏิเสธ</span>';
      return '<span class="px-2 py-1 rounded-full text-xs bg-gray-100 text-gray-700">รอส่ง</span>';
    };
    const coordApproved = l.coordinator_approval === 'อนุมัติ';
    const classApproved = l.class_teacher_approval === 'อนุมัติ';
    const deputyApproved = l.deputy_approval === 'อนุมัติ';
    const isRejected = l.leave_status === 'ปฏิเสธ' || l.coordinator_approval === 'ปฏิเสธ' || l.class_teacher_approval === 'ปฏิเสธ' || l.deputy_approval === 'ปฏิเสธ';

    // Workflow step indicator
    const stepIcon = (state) => state === 'อนุมัติ'
      ? '<span class="w-5 h-5 rounded-full bg-green-500 text-white text-[10px] flex items-center justify-center" title="อนุมัติแล้ว">✓</span>'
      : state === 'ปฏิเสธ'
        ? '<span class="w-5 h-5 rounded-full bg-red-500 text-white text-[10px] flex items-center justify-center" title="ปฏิเสธ">✕</span>'
        : '<span class="w-5 h-5 rounded-full bg-gray-200 text-gray-500 text-[10px] flex items-center justify-center" title="รอ">⏳</span>';
    const workflowSteps = `<div class="flex items-center gap-1 mb-1.5" title="ลำดับการอนุมัติ">
      <div class="flex items-center gap-1" title="อ.ผู้ประสาน">${stepIcon(l.coordinator_approval)}<span class="text-[10px] text-gray-500">ปสน.</span></div>
      <span class="text-gray-300 text-[10px]">→</span>
      <div class="flex items-center gap-1" title="อ.ประจำชั้น">${stepIcon(l.class_teacher_approval)}<span class="text-[10px] text-gray-500">ปจช.</span></div>
      <span class="text-gray-300 text-[10px]">→</span>
      <div class="flex items-center gap-1" title="ผู้บริหาร">${stepIcon(l.deputy_approval)}<span class="text-[10px] text-gray-500">ผบห.</span></div>
    </div>`;

    // Determine current approval stage label (used for student/admin view)
    let currentStage = '';
    if (l.leave_status === 'อนุมัติแล้ว') currentStage = '<span class="px-2 py-1 rounded-full text-xs bg-green-100 text-green-700">✓ อนุมัติครบแล้ว</span>';
    else if (isRejected) currentStage = '<span class="px-2 py-1 rounded-full text-xs bg-red-100 text-red-700">✕ ใบลาถูกปฏิเสธ</span>';
    else if (!coordApproved) currentStage = '<span class="px-2 py-1 rounded-full text-xs bg-amber-50 text-amber-700 border border-amber-200">⏳ รออาจารย์ผู้ประสานรายวิชา</span>';
    else if (!classApproved) currentStage = '<span class="px-2 py-1 rounded-full text-xs bg-blue-50 text-blue-700 border border-blue-200">⏳ รออาจารย์ประจำชั้น</span>';
    else if (!deputyApproved) currentStage = '<span class="px-2 py-1 rounded-full text-xs bg-purple-50 text-purple-700 border border-purple-200">⏳ รอผู้บริหาร</span>';

    const approvalField = isTeacher ? 'coordinator_approval' : isClassTeacher ? 'class_teacher_approval' : isExecutive ? 'deputy_approval' : '';
    const myApprovalStatus = approvalField ? (l[approvalField] || 'รอ') : '';

    // Sequential rule: classTeacher waits for ALL coordinators (all subjects same student+date); deputy waits for classTeacher
    const allCoordsOk = allCoordinatorsApproved(l);
    const cannotActYet = (isClassTeacher && !allCoordsOk) || (isExecutive && (!coordApproved || !classApproved));
    const waitingForLabel = isClassTeacher ? (allCoordsOk ? 'รออาจารย์ผู้ประสานอนุมัติก่อน' : 'รอ coordinator วิชาอื่นอนุมัติครบก่อน') : 'รออาจารย์ประจำชั้นอนุมัติก่อน';

    const renderApprovalCell = () => {
      // Student/admin view: show only the workflow + current stage
      if (!canApprove) {
        return `${workflowSteps}${currentStage || '<span class="px-2 py-1 rounded-full text-xs bg-gray-100 text-gray-700">รอส่ง</span>'}`;
      }
      if (isRejected && myApprovalStatus !== 'ปฏิเสธ' && myApprovalStatus !== 'อนุมัติ') {
        return `${workflowSteps}<span class="px-2 py-1 rounded-full text-xs bg-red-100 text-red-700">ใบลาถูกปฏิเสธแล้ว</span>`;
      }
      if (myApprovalStatus === 'อนุมัติ') return `${workflowSteps}<span class="px-2 py-1 rounded-full text-xs bg-green-100 text-green-700"><i data-lucide="check" class="w-3 h-3 inline"></i> คุณอนุมัติแล้ว</span>`;
      if (myApprovalStatus === 'ปฏิเสธ') return `${workflowSteps}<span class="px-2 py-1 rounded-full text-xs bg-red-100 text-red-700"><i data-lucide="x" class="w-3 h-3 inline"></i> คุณปฏิเสธแล้ว</span>`;
      if (cannotActYet) {
        return `${workflowSteps}<span class="px-2 py-1 rounded-full text-xs bg-amber-50 text-amber-700 border border-amber-200">⏳ ${waitingForLabel}</span>`;
      }
      // Coordinator (teacher) — percent modal; class teacher — note modal; executive — note modal
      const onclickApprove = isTeacher
        ? `showLeaveApprovalModal('${l.__backendId}','${l.leave_percent || ''}')`
        : isClassTeacher
          ? `showClassTeacherApprovalModal('${l.__backendId}')`
          : isExecutive
            ? `showExecutiveApprovalModal('${l.__backendId}')`
            : `approveLeave('${l.__backendId}','${approvalField}')`;
      return `${workflowSteps}<div class="flex gap-1">
          <button onclick="${onclickApprove}" class="px-2 py-1 bg-green-500 hover:bg-green-600 text-white rounded-lg text-xs flex items-center gap-1"><i data-lucide="check" class="w-3 h-3"></i>อนุมัติ</button>
          <button onclick="rejectLeave('${l.__backendId}','${approvalField}')" class="px-2 py-1 bg-red-500 hover:bg-red-600 text-white rounded-lg text-xs flex items-center gap-1"><i data-lucide="x" class="w-3 h-3"></i>ปฏิเสธ</button>
        </div>`;
    };
    const approvalButtons = renderApprovalCell();

    const reasonText = l.leave_reason ? String(l.leave_reason) : '';
    const reasonShort = reasonText.length > 60 ? reasonText.substring(0, 60) + '...' : reasonText;
    const reasonTitle = reasonText.split('"').join('&quot;');
    const reasonCell = reasonText
      ? '<span class="text-xs text-gray-700" title="' + reasonTitle + '">' + reasonShort + '</span>'
      : '<span class="text-xs text-gray-400">-</span>';
    return `<tr class="border-t hover:bg-gray-50">
        <td class="px-4 py-3">${l.name || ''}</td><td class="px-4 py-3">${l.subject_name || ''}</td>
        <td class="px-4 py-3"><span class="px-2 py-1 rounded-full text-xs ${l.leave_type === 'ลาป่วย' ? 'bg-red-100 text-red-700' : l.leave_type === 'ลาพบแพทย์' ? 'bg-purple-100 text-purple-700' : 'bg-orange-100 text-orange-700'}">${l.leave_type || ''}</span></td>
        <td class="px-4 py-3 max-w-[200px]">${reasonCell}</td>
        <td class="px-4 py-3">${l.leave_hours || ''}</td><td class="px-4 py-3">${l.leave_percent || '-'}%</td>
        <td class="px-4 py-3 text-xs">${toBuddhistDateList(l.leave_date) || '-'}</td><td class="px-4 py-3">${semLabel(l.semester)}/${l.academic_year || ''}</td>
        <td class="px-4 py-3">${getStatusBadge(l.leave_status)}</td>
        ${(canApprove || isStudent || isAdmin) ? `<td class="px-4 py-3">${approvalButtons}</td>` : ''}
        ${isAdmin ? `<td class="px-4 py-3"><div class="flex gap-1"><button onclick="showEditLeaveModal('${l.__backendId}')" class="text-blue-400 hover:text-blue-600" title="แก้ไข"><i data-lucide="pencil" class="w-4 h-4"></i></button><button onclick="deleteRecord('${l.__backendId}')" class="text-red-400 hover:text-red-600" title="ลบ"><i data-lucide="trash-2" class="w-4 h-4"></i></button></div></td>` : ''}</tr>`
  }).join('') : '<tr><td colspan="11" class="px-4 py-8 text-center text-gray-400">ไม่มีข้อมูล</td></tr>'}</tbody>
    </table></div>
  </div>
  ${paginationHTML(total, APP.pagination.perPage, APP.pagination.page, 'changePage')}`;
}

function showAddLeaveModal() {
  showModal('เพิ่มข้อมูลการลา', `
    <form id="addLeaveForm" class="space-y-3">
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-xs text-gray-600 mb-1">ชื่อ-สกุล</label><input name="name" required class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">รายวิชา</label><input name="subject_name" required class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">จำนวนชั่วโมง</label><input name="leave_hours" type="number" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">% การลา</label><input name="leave_percent" type="number" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">วันที่ลา</label><input name="leave_date" type="date" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">ประเภท</label><select name="leave_type" class="w-full border rounded-xl px-3 py-2 text-sm"><option>ลาป่วย</option><option>ลากิจ</option><option>ลาพบแพทย์</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">ภาคการศึกษา</label><select name="semester" class="w-full border rounded-xl px-3 py-2 text-sm"><option value="1">1</option><option value="2">2</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">ปีการศึกษา</label><input name="academic_year" value="2568" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      </div>
      <div><label class="block text-xs text-gray-600 mb-1">เหตุผลการลา</label><textarea name="leave_reason" rows="2" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="กรอกเหตุผลการลา"></textarea></div>
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึก</button>
    </form>
  `);
  document.getElementById('addLeaveForm').onsubmit = async (e) => {
    e.preventDefault();
    await withLoading(e.target, async () => {
      const fd = new FormData(e.target);
      const obj = { type: 'leave', created_at: new Date().toISOString() }; fd.forEach((v, k) => obj[k] = (k === 'leave_hours' || k === 'leave_percent') ? Number(v) : v);
      obj.leave_status = 'รออนุมัติ';
      obj.coordinator_approval = 'รอ';
      obj.class_teacher_approval = 'รอ';
      obj.deputy_approval = 'รอ';
      const r = await GSheetDB.create(obj);
      if (r.isOk) { showToast('เพิ่มข้อมูลการลาสำเร็จ'); closeModal() } else showToast('เกิดข้อผิดพลาด', 'error');
    });
  };
}

// ======================== LOGIN LOG (admin only) ========================
function loginLogPage() {
  if (APP.currentRole !== 'admin') {
    return '<div class="bg-white rounded-2xl border border-blue-100 p-8 text-center text-gray-500">เฉพาะผู้ดูแลระบบเท่านั้นที่ดูหน้านี้ได้</div>';
  }
  const logs = getDataByType('login_log');
  // Sort newest first by created_at / timestamp
  const sortedLogs = [...logs].sort((a, b) => {
    const ta = a.created_at || a.timestamp || '';
    const tb = b.created_at || b.timestamp || '';
    return tb.localeCompare(ta);
  });

  // Filters
  const fRole = APP.filters._loginLogRole || '';
  const fEvent = APP.filters._loginLogEvent || '';
  const fSearch = (APP.filters._loginLogSearch || '').toLowerCase();
  const fDateFrom = APP.filters._loginLogFrom || '';
  const fDateTo = APP.filters._loginLogTo || '';

  let filtered = sortedLogs;
  if (fRole) filtered = filtered.filter(l => norm(l.role) === fRole);
  if (fEvent) filtered = filtered.filter(l => norm(l.event_type) === fEvent);
  if (fSearch) filtered = filtered.filter(l => (l.user_name || '').toLowerCase().includes(fSearch) || (l.identifier || '').toLowerCase().includes(fSearch));
  if (fDateFrom) filtered = filtered.filter(l => (l.timestamp || '').slice(0, 10) >= fDateFrom);
  if (fDateTo) filtered = filtered.filter(l => (l.timestamp || '').slice(0, 10) <= fDateTo);

  const total = filtered.length;
  const paged = paginate(filtered);

  // Summary stats (today, last 7 days, total unique users today)
  const today = new Date().toISOString().slice(0, 10);
  const past7Date = new Date(); past7Date.setDate(past7Date.getDate() - 7);
  const past7 = past7Date.toISOString().slice(0, 10);
  const todayLogins = sortedLogs.filter(l => l.event_type === 'login' && (l.timestamp || '').slice(0, 10) === today);
  const last7Logins = sortedLogs.filter(l => l.event_type === 'login' && (l.timestamp || '').slice(0, 10) >= past7);
  const uniqueTodayUsers = [...new Set(todayLogins.map(l => l.user_name))].length;

  const eventBadge = (ev) => {
    if (ev === 'login') return '<span class="px-2 py-1 rounded-full text-xs font-semibold bg-green-100 text-green-700"><i data-lucide="log-in" class="w-3 h-3 inline"></i> เข้าสู่ระบบ</span>';
    if (ev === 'logout') return '<span class="px-2 py-1 rounded-full text-xs font-semibold bg-gray-100 text-gray-700"><i data-lucide="log-out" class="w-3 h-3 inline"></i> ออกจากระบบ</span>';
    if (ev === 'login_failed') return '<span class="px-2 py-1 rounded-full text-xs font-semibold bg-red-100 text-red-700"><i data-lucide="x-circle" class="w-3 h-3 inline"></i> ล็อกอินไม่สำเร็จ</span>';
    return `<span class="px-2 py-1 rounded-full text-xs font-semibold bg-gray-100 text-gray-600">${ev || '-'}</span>`;
  };

  return `<div class="flex flex-wrap items-center justify-between gap-3 mb-4">
    <h2 class="text-xl font-bold text-gray-800"><i data-lucide="log-in" class="w-6 h-6 inline mr-2"></i>บันทึกการเข้าใช้งานระบบ</h2>
    <div class="flex gap-2">
      <button onclick="exportLoginLogCSV()" class="flex items-center gap-2 px-4 py-2 bg-emerald-500 text-white rounded-xl hover:bg-emerald-600 text-sm"><i data-lucide="download" class="w-4 h-4"></i>Export CSV</button>
      <button onclick="refreshData()" class="flex items-center gap-2 px-4 py-2 bg-primary text-white rounded-xl hover:bg-primaryDark text-sm"><i data-lucide="refresh-cw" class="w-4 h-4"></i>รีเฟรช</button>
    </div>
  </div>

  <div class="grid grid-cols-1 sm:grid-cols-3 gap-4 mb-4">
    <div class="bg-white rounded-2xl p-4 border border-blue-100 flex items-center gap-3">
      <div class="w-10 h-10 bg-blue-100 rounded-lg flex items-center justify-center"><i data-lucide="calendar-check" class="w-5 h-5 text-blue-600"></i></div>
      <div><p class="text-xs text-gray-500">เข้าใช้งานวันนี้</p><p class="text-xl font-bold text-gray-800">${todayLogins.length} <span class="text-xs font-normal text-gray-500">ครั้ง</span></p></div>
    </div>
    <div class="bg-white rounded-2xl p-4 border border-blue-100 flex items-center gap-3">
      <div class="w-10 h-10 bg-emerald-100 rounded-lg flex items-center justify-center"><i data-lucide="users" class="w-5 h-5 text-emerald-600"></i></div>
      <div><p class="text-xs text-gray-500">ผู้ใช้ที่เข้าระบบวันนี้</p><p class="text-xl font-bold text-gray-800">${uniqueTodayUsers} <span class="text-xs font-normal text-gray-500">คน</span></p></div>
    </div>
    <div class="bg-white rounded-2xl p-4 border border-blue-100 flex items-center gap-3">
      <div class="w-10 h-10 bg-amber-100 rounded-lg flex items-center justify-center"><i data-lucide="trending-up" class="w-5 h-5 text-amber-600"></i></div>
      <div><p class="text-xs text-gray-500">เข้าใช้งาน 7 วันล่าสุด</p><p class="text-xl font-bold text-gray-800">${last7Logins.length} <span class="text-xs font-normal text-gray-500">ครั้ง</span></p></div>
    </div>
  </div>

  <div class="bg-white rounded-2xl p-4 border border-blue-100 mb-4">
    <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-5 gap-3">
      <div>
        <label class="block text-xs text-gray-600 mb-1">ค้นหา (ชื่อ/รหัส/อีเมล)</label>
        <input type="text" value="${fSearch}" placeholder="พิมพ์เพื่อค้นหา..." class="w-full border border-gray-200 rounded-xl px-3 py-2 text-sm" oninput="clearTimeout(window._loginLogSearchTimer);window._loginLogSearchTimer=setTimeout(()=>{APP.filters._loginLogSearch=this.value;APP.pagination.page=1;renderCurrentPage()},300)">
      </div>
      <div>
        <label class="block text-xs text-gray-600 mb-1">บทบาท</label>
        <select onchange="APP.filters._loginLogRole=this.value;APP.pagination.page=1;renderCurrentPage()" class="w-full border border-gray-200 rounded-xl px-3 py-2 text-sm">
          <option value="">ทั้งหมด</option>
          <option value="admin" ${fRole === 'admin' ? 'selected' : ''}>ผู้ดูแลระบบ</option>
          <option value="academic" ${fRole === 'academic' ? 'selected' : ''}>งานวิชาการ</option>
          <option value="teacher" ${fRole === 'teacher' ? 'selected' : ''}>อาจารย์</option>
          <option value="classTeacher" ${fRole === 'classTeacher' ? 'selected' : ''}>อ.ประจำชั้น</option>
          <option value="student" ${fRole === 'student' ? 'selected' : ''}>นักศึกษา</option>
          <option value="executive" ${fRole === 'executive' ? 'selected' : ''}>ผู้บริหาร</option>
        </select>
      </div>
      <div>
        <label class="block text-xs text-gray-600 mb-1">เหตุการณ์</label>
        <select onchange="APP.filters._loginLogEvent=this.value;APP.pagination.page=1;renderCurrentPage()" class="w-full border border-gray-200 rounded-xl px-3 py-2 text-sm">
          <option value="">ทั้งหมด</option>
          <option value="login" ${fEvent === 'login' ? 'selected' : ''}>เข้าสู่ระบบ</option>
          <option value="logout" ${fEvent === 'logout' ? 'selected' : ''}>ออกจากระบบ</option>
          <option value="login_failed" ${fEvent === 'login_failed' ? 'selected' : ''}>ล็อกอินไม่สำเร็จ</option>
        </select>
      </div>
      <div>
        <label class="block text-xs text-gray-600 mb-1">ตั้งแต่วันที่</label>
        <input type="date" value="${fDateFrom}" onchange="APP.filters._loginLogFrom=this.value;APP.pagination.page=1;renderCurrentPage()" class="w-full border border-gray-200 rounded-xl px-3 py-2 text-sm">
      </div>
      <div>
        <label class="block text-xs text-gray-600 mb-1">ถึงวันที่</label>
        <input type="date" value="${fDateTo}" onchange="APP.filters._loginLogTo=this.value;APP.pagination.page=1;renderCurrentPage()" class="w-full border border-gray-200 rounded-xl px-3 py-2 text-sm">
      </div>
    </div>
    ${(fRole || fEvent || fSearch || fDateFrom || fDateTo) ? `<div class="mt-3 flex justify-end"><button onclick="APP.filters._loginLogRole='';APP.filters._loginLogEvent='';APP.filters._loginLogSearch='';APP.filters._loginLogFrom='';APP.filters._loginLogTo='';APP.pagination.page=1;renderCurrentPage()" class="text-xs text-blue-600 hover:underline"><i data-lucide="x" class="w-3 h-3 inline"></i> ล้างตัวกรอง</button></div>` : ''}
  </div>

  <div class="bg-white rounded-2xl border border-blue-100 overflow-hidden">
    <div class="overflow-x-auto"><table class="w-full text-sm">
      <thead><tr class="bg-surface text-left">
        <th class="px-4 py-3 font-semibold">วันที่/เวลา</th>
        <th class="px-4 py-3 font-semibold">ชื่อผู้ใช้</th>
        <th class="px-4 py-3 font-semibold">บทบาท</th>
        <th class="px-4 py-3 font-semibold">รหัส/อีเมล</th>
        <th class="px-4 py-3 font-semibold">เหตุการณ์</th>
      </tr></thead>
      <tbody>${paged.length ? paged.map(l => `<tr class="border-t hover:bg-gray-50">
        <td class="px-4 py-3 font-mono text-xs text-gray-600">${l.timestamp || (l.created_at || '').replace('T', ' ').slice(0, 19) || '-'}</td>
        <td class="px-4 py-3 font-medium">${l.user_name || '-'}</td>
        <td class="px-4 py-3 text-xs">${l.role_label || LOGIN_LOG_ROLE_LABEL[l.role] || l.role || '-'}</td>
        <td class="px-4 py-3 text-xs text-gray-500">${l.identifier || '-'}</td>
        <td class="px-4 py-3">${eventBadge(l.event_type)}</td>
      </tr>`).join('') : '<tr><td colspan="5" class="px-4 py-8 text-center text-gray-400">ไม่มีข้อมูลการเข้าใช้งาน</td></tr>'}</tbody>
    </table></div>
  </div>
  ${paginationHTML(total, APP.pagination.perPage, APP.pagination.page, 'changePage')}

  <div class="mt-4 bg-blue-50 border border-blue-200 rounded-xl p-3 text-xs text-blue-800">
    <i data-lucide="info" class="w-4 h-4 inline mr-1"></i>
    หมายเหตุ: ระบบบันทึกการเข้า/ออกระบบลงในแท็บ <code class="bg-white px-1 rounded">login_log</code> ของ Google Sheet โดยอัตโนมัติ หากยังไม่เคยมีข้อมูล โปรดสร้างแท็บใหม่พร้อมหัวคอลัมน์: <code class="bg-white px-1 rounded">event_type, user_name, role, role_label, identifier, user_agent, timestamp, created_at</code>
  </div>`;
}

function exportLoginLogCSV() {
  const logs = getDataByType('login_log');
  if (!logs.length) { showToast('ไม่มีข้อมูลให้ส่งออก', 'error'); return; }
  const headers = ['timestamp', 'user_name', 'role_label', 'event_type', 'identifier', 'user_agent'];
  const rows = [headers.join(',')].concat(
    logs.map(l => headers.map(h => {
      const v = String(l[h] || '').replace(/"/g, '""');
      return `"${v}"`;
    }).join(','))
  );
  const csv = '﻿' + rows.join('\n');
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `login_log_${new Date().toISOString().slice(0, 10)}.csv`;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  URL.revokeObjectURL(url);
  showToast('ส่งออก CSV สำเร็จ');
}

// ======================== SETTINGS ========================
function settingsPage() {
  const roles = ['admin', 'academic', 'executive', 'teacher', 'classTeacher', 'student'];
  const modules = ['dashboard', 'students', 'subjects', 'schedule', 'grades', 'engResults', 'teachers', 'teacherDirectory', 'services', 'tracking', 'gradeTracking', 'fileTracking', 'leave'];
  const moduleLabels = { dashboard: 'หน้าหลัก', students: 'ข้อมูลนักศึกษา', subjects: 'รายวิชา', schedule: 'ตารางเรียน/สอบ', grades: 'ผลการเรียน', engResults: 'ผลสอบ ENG', teachers: 'ข้อมูลอาจารย์', teacherDirectory: 'ทำเนียบอาจารย์', services: 'บริการอื่นๆ', tracking: 'ติดตามการส่งรายละเอียดรายวิชา', gradeTracking: 'ติดตามการส่งเกรดรายวิชา', fileTracking: 'ติดตามส่งแฟ้มรายวิชา', leave: 'ระบบการลาของนักศึกษา' };
  const roleLabels = { admin: 'ผู้ดูแลระบบ', academic: 'เจ้าหน้าที่งานวิชาการ', executive: 'ผู้บริหาร', teacher: 'อาจารย์', classTeacher: 'อ.ประจำชั้น', student: 'นักศึกษา' };

  const users = applyFilters(getDataByType('user'));
  const total = users.length; const paged = paginate(users);
  const usersTable = paged.map(u => `<tr class="border-t hover:bg-gray-50">
    <td class="px-4 py-3">${u.name || ''}</td>
    <td class="px-4 py-3">${u.email || u.national_id || ''}</td>
    <td class="px-4 py-3"><span class="px-2 py-1 rounded-full text-xs bg-surface">${roleLabels[u.role] || u.role}</span></td>
    <td class="px-4 py-3"><div class="flex gap-1"><button onclick="showEditUserModal('${u.__backendId}')" class="text-blue-400 hover:text-blue-600" title="แก้ไข"><i data-lucide="pencil" class="w-4 h-4"></i></button><button onclick="deleteRecord('${u.__backendId}')" class="text-red-400 hover:text-red-600" title="ลบ"><i data-lucide="trash-2" class="w-4 h-4"></i></button></div></td>
  </tr>`).join('');

  return `<h2 class="text-xl font-bold text-gray-800 mb-6"><i data-lucide="settings" class="w-6 h-6 inline mr-2"></i>ตั้งค่าระบบ</h2>
  
  <div class="bg-white rounded-2xl p-5 border border-blue-100 mb-6">
    <div class="flex items-center justify-between mb-4">
      <h3 class="font-bold">จัดการผู้ใช้งาน (${total} คน)</h3>
      <button onclick="showAddUserModal()" class="flex items-center gap-2 px-4 py-2 bg-primary text-white rounded-xl hover:bg-primaryDark text-sm"><i data-lucide="plus" class="w-4 h-4"></i>เพิ่มผู้ใช้</button>
    </div>
    ${filterBar({ semester: false, year: false })}
    <div class="overflow-x-auto">
      <table class="w-full text-sm table-fixed">
        <thead><tr class="bg-surface text-left"><th class="px-4 py-3 font-semibold" style="width: 30%;">ชื่อ-สกุล</th><th class="px-4 py-3 font-semibold" style="width: 35%;">ชื่อผู้ใช้/Email/รหัสประชาชน</th><th class="px-4 py-3 font-semibold" style="width: 25%;">บทบาท</th><th class="px-4 py-3" style="width: 10%;"></th></tr></thead>
        <tbody>${usersTable || '<tr><td colspan="4" class="px-4 py-8 text-center text-gray-400">ยังไม่มีผู้ใช้</td></tr>'}</tbody>
      </table>
    </div>
    ${paginationHTML(total, APP.pagination.perPage, APP.pagination.page, 'changePage')}
  </div>
  
  <div class="bg-white rounded-2xl p-5 border border-blue-100">
    <h3 class="font-bold mb-4">สิทธิ์การเข้าถึงระบบ</h3>
    <div class="overflow-x-auto"><table class="w-full text-sm table-fixed" style="min-width:600px">
      <thead><tr class="bg-surface"><th class="px-3 py-2 text-left font-semibold" style="width: 20%;">โมดูล</th>${roles.map(r => `<th class="px-3 py-2 text-center font-semibold" style="width: 20%;">${roleLabels[r]}</th>`).join('')}</tr></thead>
      <tbody>${modules.map(m => `<tr class="border-t hover:bg-gray-50"><td class="px-3 py-2 font-medium">${moduleLabels[m]}</td>${roles.map(r => `<td class="px-3 py-2 text-center"><label class="inline-flex"><input type="checkbox" ${APP.permissions[r]?.[m] ? 'checked' : ''} onchange="togglePermission('${r}','${m}',this.checked)" class="w-4 h-4 rounded border-gray-300 text-primary focus:ring-primary"></label></td>`).join('')}</tr>`).join('')}</tbody>
    </table></div>
  </div>
  ${APP.currentRole === 'admin' ? `
  <div class="bg-white rounded-2xl p-5 border border-orange-200 mt-6">
    <h3 class="font-bold mb-1 flex items-center gap-2"><i data-lucide="database" class="w-5 h-5 text-orange-500"></i>เชื่อมต่อฐานข้อมูล</h3>
    <p class="text-xs text-gray-400 mb-4">เฉพาะผู้ดูแลระบบ (Admin) เท่านั้นที่เปลี่ยนได้</p>
    <div class="space-y-3">
      <div>
        <label class="block text-xs text-gray-600 mb-1">Google Sheet URL หรือ Spreadsheet ID</label>
        <input id="adminSheetUrl" value="${getActiveConfig().spreadsheetId || ''}" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="https://docs.google.com/spreadsheets/d/xxxxx/edit">
      </div>
      <div class="bg-green-50 border border-green-200 rounded-xl p-3">
        <label class="block text-xs text-green-700 font-semibold mb-1">Apps Script URL (สำหรับแก้ไขข้อมูลผ่านเว็บ)</label>
        <input id="adminScriptUrl" value="${getActiveConfig().scriptUrl || ''}" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="https://script.google.com/macros/s/xxxxx/exec">
        <p class="text-xs text-gray-400 mt-1">ถ้าไม่กรอก = อ่านอย่างเดียว | กรอก = เพิ่ม/แก้ไข/ลบข้อมูลได้</p>
      </div>
      <div class="flex gap-2">
        <button onclick="saveAdminGSheetConfig()" class="flex items-center gap-2 px-4 py-2 bg-orange-500 text-white rounded-xl hover:bg-orange-600 text-sm"><i data-lucide="save" class="w-4 h-4"></i>บันทึกและเชื่อมต่อใหม่</button>
        <button onclick="resetToDefaultConfig()" class="flex items-center gap-2 px-4 py-2 bg-gray-200 text-gray-700 rounded-xl hover:bg-gray-300 text-sm"><i data-lucide="rotate-ccw" class="w-4 h-4"></i>คืนค่าเริ่มต้น</button>
      </div>
    </div>
  </div>` : ''}`;
}

function saveAdminGSheetConfig() {
  const urlInput = document.getElementById('adminSheetUrl');
  const scriptInput = document.getElementById('adminScriptUrl');
  if (!urlInput) return;
  const sheetId = GSheetDB.extractSheetId(urlInput.value) || urlInput.value.trim();
  if (!sheetId) { showToast('กรุณากรอก Google Sheet URL หรือ Spreadsheet ID', 'error'); return }
  const scriptUrl = scriptInput ? scriptInput.value.trim() : '';
  if (!confirm('ต้องการเปลี่ยนการเชื่อมต่อ Google Sheet ใช่ไหม?\nระบบจะโหลดข้อมูลใหม่ทั้งหมด')) return;
  GSheetDB.storeConfig({ spreadsheetId: sheetId, scriptUrl: scriptUrl });
  location.reload();
}

function resetToDefaultConfig() {
  if (!confirm('ต้องการคืนค่าเริ่มต้นของการเชื่อมต่อ Google Sheet ใช่ไหม?')) return;
  GSheetDB.clearConfig();
  location.reload();
}

function showAddUserModal() {
  showModal('เพิ่มผู้ใช้งาน', `
    <form id="addUserForm" class="space-y-3">
      <div><label class="block text-xs text-gray-600 mb-1">บทบาท *</label>
        <select name="role" required onchange="onUserRoleChange(this.value)" class="w-full border rounded-xl px-3 py-2 text-sm">
          <option value="">เลือกบทบาท</option>
          <option value="admin">ผู้ดูแลระบบ</option>
          <option value="academic">เจ้าหน้าที่งานวิชาการ</option>
          <option value="executive">ผู้บริหาร</option>
          <option value="teacher">อาจารย์</option>
          <option value="classTeacher">อาจารย์ประจำชั้น</option>
          <option value="student">นักศึกษา</option>
        </select>
      </div>
      <div><label class="block text-xs text-gray-600 mb-1">ชื่อ-สกุล *</label>
        <input name="name" required class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <div id="userCredFields" class="space-y-3"></div>
      <div><label class="block text-xs text-gray-600 mb-1">ชั้นปีที่รับผิดชอบ (อ.ประจำชั้นเท่านั้น)</label>
        <select name="responsible_year" class="w-full border rounded-xl px-3 py-2 text-sm">
          <option value="">ไม่มี</option>
          <option>1</option>
          <option>2</option>
          <option>3</option>
          <option>4</option>
        </select>
      </div>
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึก</button>
    </form>
  `);
  document.getElementById('addUserForm').onsubmit = async (e) => {
    e.preventDefault();
    const fd = new FormData(e.target);
    const role = fd.get('role');
    if (!role) { showToast('กรุณาเลือกบทบาท', 'error'); return }
    if (APP.allData.filter(d => d.type === 'user').length >= 999) { showToast('ข้อมูลเต็ม', 'error'); return }
    const obj = { type: 'user', name: fd.get('name'), role, created_at: new Date().toISOString() };
    if (role === 'admin') { obj.password = fd.get('password') || '123456' }
    else if (role === 'student') { obj.national_id = fd.get('national_id') }
    else { obj.email = fd.get('email'); obj.password = fd.get('password') || '123456' }
    const resp_yr = fd.get('responsible_year');
    if (resp_yr) obj.responsible_year = resp_yr;

    const r = await GSheetDB.create(obj);
    if (r.isOk) { showToast('เพิ่มผู้ใช้สำเร็จ'); closeModal() } else showToast('เกิดข้อผิดพลาด', 'error');
  };
}

function onUserRoleChange(role) {
  const fieldsDiv = document.getElementById('userCredFields');
  if (!fieldsDiv) return;
  fieldsDiv.innerHTML = '';
  if (role === 'admin') {
    fieldsDiv.innerHTML = `<div><label class="block text-xs text-gray-600 mb-1">รหัสผ่าน (6 หลัก) *</label>
      <input name="password" maxlength="6" pattern="[0-9]{6}" required class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="6 หลัก ตัวเลข"></div>`;
  } else if (role === 'student') {
    fieldsDiv.innerHTML = `<div><label class="block text-xs text-gray-600 mb-1">เลขบัตรประชาชน 13 หลัก *</label>
      <input name="national_id" maxlength="13" pattern="[0-9]{13}" required class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="13 หลัก"></div>`;
  } else if (role === 'teacher' || role === 'classTeacher' || role === 'executive' || role === 'academic') {
    fieldsDiv.innerHTML = `<div><label class="block text-xs text-gray-600 mb-1">E-mail *</label>
      <input name="email" type="email" required class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <div><label class="block text-xs text-gray-600 mb-1">รหัสผ่าน (ถ้าไม่ระบุจะเป็น 123456)</label>
      <input name="password" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="กรุณาตั้งรหัสผ่าน"></div>`;
  }
}

async function togglePermission(role, module, checked) {
  if (!APP.permissions[role]) APP.permissions[role] = {};
  APP.permissions[role][module] = checked ? 1 : 0;

  // Rebuild sidebar immediately so tabs reflect the change
  buildSidebar();

  const existing = getDataByType('permission').find(p => p.role === role && String(p.module) === String(module));
  let result;

  if (existing) {
    existing.value = String(checked ? 1 : 0);
    result = await GSheetDB.update(existing);
  } else {
    result = await GSheetDB.create({
      type: 'permission',
      role: role,
      module: module,
      value: checked ? 1 : 0
    });
  }

  if (result && result.isOk) {
    showToast('อัปเดตสิทธิ์ในฐานข้อมูลสำเร็จ');
  } else {
    showToast('เกิดข้อผิดพลาดในการบันทึกสิทธิ์', 'error');
  }
}

// ======================== LEAVE VALIDATION (Student) ========================
function onLeaveTypeChange(type) {
  const extra = document.getElementById('leaveExtra');
  if (!extra) return;
  extra.innerHTML = '';
  if (type === 'ลาป่วย') {
    extra.innerHTML = `<div class="bg-yellow-50 p-3 rounded-xl text-xs text-yellow-700"><i data-lucide="alert-triangle" class="w-4 h-4 inline"></i> ลาป่วย 3 วันขึ้นไป ต้องแนบใบรับรองแพทย์ (.jpg, .pdf, .png)</div>
    <div id="sickCertUpload" class="hidden"><label class="block text-xs text-gray-600 mb-1">แนบใบรับรองแพทย์ *</label><input type="file" accept=".jpg,.pdf,.png" class="w-full text-sm" name="medical_cert"></div>`;
  } else if (type === 'ลากิจ') {
    extra.innerHTML = `<div class="bg-blue-50 p-3 rounded-xl text-xs text-blue-700"><i data-lucide="info" class="w-4 h-4 inline"></i> ลากิจต้องส่งล่วงหน้า 1-2 วัน หากส่งช้าโปรดระบุเหตุผลในช่องเหตุผลการลาด้านบน</div>`;
  } else if (type === 'ลาพบแพทย์') {
    extra.innerHTML = `<div class="bg-purple-50 p-3 rounded-xl text-xs text-purple-700"><i data-lucide="info" class="w-4 h-4 inline"></i> ลาพบแพทย์ต้องส่งล่วงหน้า 3 วันทำการ และแนบใบนัดแพทย์</div>
    <div><label class="block text-xs text-gray-600 mb-1">แนบใบนัดแพทย์ * (.jpg, .pdf, .png)</label><input type="file" accept=".jpg,.pdf,.png" class="w-full text-sm" name="appointment_doc"></div>`;
  }
  lucide.createIcons();
  // Re-check dates to show/hide late reason field immediately
  if (typeof validateLeaveDate === 'function') validateLeaveDate();
}

function validateLeaveDate() {
  // Update Buddhist date display next to each date input + sync hidden field
  syncLeaveDates();
  const form = document.getElementById('leaveForm'); if (!form) return;
  const typeSelect = form.querySelector('[name="leave_type"]');
  if (!typeSelect) return;
  const type = typeSelect.value;
  // Use the earliest selected date for late-leave check
  const inputs = form.querySelectorAll('.leave-date-input');
  let earliestDate = null;
  inputs.forEach(inp => {
    if (inp.value) {
      const d = new Date(inp.value);
      if (!earliestDate || d < earliestDate) earliestDate = d;
    }
  });
  if (!earliestDate) return;
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const diffDays = Math.ceil((earliestDate - today) / (1000 * 60 * 60 * 24));

  // Note: leave_reason is now always required for all leave types — handled in submit validation
}

// ----- Multi-date leave row helpers -----
function addLeaveDateRow() {
  const list = document.getElementById('leaveDateList'); if (!list) return;
  const row = document.createElement('div');
  row.className = 'flex items-center gap-2 leave-date-row';
  row.innerHTML = `
    <input type="date" class="leave-date-input flex-1 border rounded-xl px-3 py-2 text-sm" required onchange="validateLeaveDate()">
    <span class="be-display text-xs text-gray-500 min-w-[90px]"></span>
    <button type="button" onclick="removeLeaveDateRow(this)" class="px-2 py-1 text-red-500 hover:bg-red-50 rounded-lg text-xs"><i data-lucide="x" class="w-4 h-4"></i></button>`;
  list.appendChild(row);
  if (window.lucide) lucide.createIcons();
}

function removeLeaveDateRow(btn) {
  const list = document.getElementById('leaveDateList'); if (!list) return;
  const rows = list.querySelectorAll('.leave-date-row');
  if (rows.length <= 1) {
    // Just clear the value if last row
    const inp = btn.parentElement.querySelector('.leave-date-input');
    if (inp) { inp.value = ''; }
    syncLeaveDates();
    return;
  }
  btn.parentElement.remove();
  syncLeaveDates();
}

function syncLeaveDates() {
  const list = document.getElementById('leaveDateList'); if (!list) return;
  const hidden = document.getElementById('leaveDateHidden');
  const dates = [];
  list.querySelectorAll('.leave-date-row').forEach(row => {
    const inp = row.querySelector('.leave-date-input');
    const beSpan = row.querySelector('.be-display');
    if (inp && inp.value) {
      dates.push(inp.value);
      if (beSpan) beSpan.textContent = '(พ.ศ. ' + toBuddhistDate(inp.value) + ')';
    } else if (beSpan) {
      beSpan.textContent = '';
    }
  });
  if (hidden) hidden.value = dates.join(',');
}

function updateLeaveCoordinator(subjectName) {
  const sub = getDataByType('subject').find(s => s.subject_name === subjectName);
  const el = document.getElementById('leaveCoordinator');
  if (el) el.value = sub?.coordinator || '';
}

function toggleLeaveSubjectHours(checkbox) {
  const row = checkbox.closest('.leave-subject-row');
  if (!row) return;
  const hoursDiv = row.querySelector('.leave-subject-hours');
  const hoursInput = row.querySelector('.leave-hours-input');
  if (checkbox.checked) {
    if (hoursDiv) hoursDiv.classList.remove('hidden');
    if (hoursInput) hoursInput.required = true;
  } else {
    if (hoursDiv) hoursDiv.classList.add('hidden');
    if (hoursInput) { hoursInput.required = false; hoursInput.value = ''; }
  }
}

// updateEvalTeacherOptions / setEvalScore removed (eval feature removed)

// ======================== INIT PAGE SCRIPTS ========================
function initPageScripts(page) {
  if (page === 'dashboard') { renderCalendar('dashCalendar') }
  if (page === 'schedule') { renderCalendar('scheduleCalendar') }

  // Student leave form (multi-subject)
  const leaveForm = document.getElementById('leaveForm');
  if (leaveForm) {
    leaveForm.onsubmit = async (e) => {
      e.preventDefault();
      if (typeof syncLeaveDates === 'function') syncLeaveDates();

      const fd = new FormData(leaveForm);
      const leaveDate = fd.get('leave_date');
      const type = fd.get('leave_type');
      const today = new Date(); today.setHours(0, 0, 0, 0);
      const valEl = document.getElementById('leaveValidation');

      if (!leaveDate) {
        if (valEl) { valEl.textContent = 'กรุณาเลือกวันที่ลาอย่างน้อย 1 วัน'; valEl.classList.remove('hidden') }
        return;
      }
      if (!type) {
        if (valEl) { valEl.textContent = 'กรุณาเลือกประเภทการลา'; valEl.classList.remove('hidden') }
        return;
      }

      // Collect checked subjects with hours
      const checkedSubjects = [];
      document.querySelectorAll('.leave-subject-check:checked').forEach(cb => {
        const row = cb.closest('.leave-subject-row');
        const hoursInput = row ? row.querySelector('.leave-hours-input') : null;
        const hours = hoursInput ? Number(hoursInput.value) : 0;
        checkedSubjects.push({
          subject_name: cb.value,
          coordinator: cb.dataset.coordinator || '',
          leave_hours: hours
        });
      });

      if (checkedSubjects.length === 0) {
        if (valEl) { valEl.textContent = 'กรุณาเลือกรายวิชาอย่างน้อย 1 วิชา'; valEl.classList.remove('hidden') }
        return;
      }
      const missingHours = checkedSubjects.find(s => !s.leave_hours || s.leave_hours <= 0);
      if (missingHours) {
        if (valEl) { valEl.textContent = 'กรุณากรอกจำนวนชั่วโมงให้ครบทุกรายวิชาที่เลือก'; valEl.classList.remove('hidden') }
        return;
      }

      // Date validation
      const dateList = String(leaveDate).split(',').map(d => d.trim()).filter(Boolean);
      const dateObjs = dateList.map(d => new Date(d)).sort((a, b) => a - b);
      const earliest = dateObjs[0];
      const diff = Math.ceil((earliest - today) / (1000 * 60 * 60 * 24));

      // Reason is required for all leave types
      const reasonText = (fd.get('leave_reason') || '').toString().trim();
      if (!reasonText) {
        if (valEl) { valEl.textContent = 'กรุณากรอกเหตุผลการลา'; valEl.classList.remove('hidden') }
        return;
      }
      if (type === 'ลาป่วย' && diff <= -3) {
        const cert = fd.get('medical_cert');
        if (!cert || !cert.name) {
          if (valEl) { valEl.textContent = 'ลาป่วย 3 วันขึ้นไป ต้องแนบใบรับรองแพทย์'; valEl.classList.remove('hidden') }
          return;
        }
      }
      if (type === 'ลาพบแพทย์') {
        if (diff < 3) {
          if (valEl) { valEl.textContent = 'ลาพบแพทย์ต้องส่งล่วงหน้าอย่างน้อย 3 วันทำการ'; valEl.classList.remove('hidden') }
          return;
        }
        const appt = fd.get('appointment_doc');
        if (!appt || !appt.name) {
          if (valEl) { valEl.textContent = 'กรุณาแนบใบนัดแพทย์'; valEl.classList.remove('hidden') }
          return;
        }
      }

      if (valEl) valEl.classList.add('hidden');
      await withLoading(leaveForm, async () => {
        let successCount = 0;
        for (const subj of checkedSubjects) {
          const obj = {
            type: 'leave', created_at: new Date().toISOString(),
            name: fd.get('name'),
            subject_name: subj.subject_name,
            coordinator: subj.coordinator,
            leave_hours: subj.leave_hours,
            semester: fd.get('semester'),
            academic_year: fd.get('academic_year'),
            leave_date: leaveDate,
            leave_type: type,
            leave_reason: fd.get('leave_reason') || '',
            class_teacher: fd.get('class_teacher') || '',
            leave_status: 'รออนุมัติ',
            coordinator_approval: 'รอ',
            class_teacher_approval: 'รอ',
            deputy_approval: 'รอ'
          };
          if (APP.allData.filter(d => d.type === 'leave').length >= 999) { showToast('ข้อมูลเต็ม', 'error'); break; }
          const r = await GSheetDB.create(obj);
          if (r.isOk) successCount++;
        }
        if (successCount > 0) {
          showToast(`ส่งใบลาสำเร็จ ${successCount} รายวิชา`);
          leaveForm.reset();
          // Uncheck all and hide hours
          document.querySelectorAll('.leave-subject-check').forEach(cb => { cb.checked = false; toggleLeaveSubjectHours(cb); });
        } else {
          showToast('เกิดข้อผิดพลาด', 'error');
        }
      });
    };
  }

  // Eval form removed (teacher evaluation feature has been removed from system)
}

// ======================== CALENDAR ========================
function renderCalendar(containerId) {
  const el = document.getElementById(containerId); if (!el) return;
  const now = new Date();
  const year = now.getFullYear(); const month = now.getMonth();
  const firstDay = new Date(year, month, 1).getDay();
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const monthNames = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];

  const events = [...getDataByType('announcement'), ...getDataByType('schedule')];

  let h = `<div class="text-center font-bold mb-3">${monthNames[month]} ${year + 543}</div>`;
  h += '<div class="grid grid-cols-7 gap-px text-center text-xs">';
  ['อา', 'จ', 'อ', 'พ', 'พฤ', 'ศ', 'ส'].forEach(d => h += `<div class="py-1 font-semibold text-gray-500">${d}</div>`);
  for (let i = 0; i < firstDay; i++)h += '<div></div>';
  for (let d = 1; d <= daysInMonth; d++) {
    const dateStr = `${year}-${String(month + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
    const dayEvents = events.filter(e => (e.schedule_date || e.announcement_date || '').startsWith(dateStr));
    const isToday = d === now.getDate();
    h += `<div class="cal-day p-1 min-h-[40px] rounded-lg ${isToday ? 'bg-primary text-white' : ''}">
      <div class="text-xs ${isToday ? 'font-bold' : ''}">${d}</div>
      ${dayEvents.slice(0, 2).map(e => `<div class="cal-event ${e.schedule_type === 'สอบ' || e.event_type === 'สอบ' ? 'bg-red-200 text-red-800' : e.event_type === 'วันหยุด' ? 'bg-green-200 text-green-800' : 'bg-blue-200 text-blue-800'}">${(e.subject_name || e.announcement_title || '').substring(0, 6)}</div>`).join('')}
    </div>`;
  }
  h += '</div>';
  el.innerHTML = h;
}

// ======================== LEAVE APPROVAL ========================
async function approveLeave(id, approvalField, extra) {
  const rec = APP.allData.find(d => d.__backendId === id); if (!rec) return;
  rec[approvalField] = 'อนุมัติ';
  if (extra && typeof extra === 'object') {
    Object.keys(extra).forEach(k => { if (extra[k] !== undefined && extra[k] !== null && extra[k] !== '') rec[k] = extra[k]; });
  }
  // If all approvals done, mark as approved
  if (rec.coordinator_approval === 'อนุมัติ' && rec.class_teacher_approval === 'อนุมัติ' && rec.deputy_approval === 'อนุมัติ') {
    rec.leave_status = 'อนุมัติแล้ว';
  }

  showToast('กำลังบันทึก...', 'loading');
  const r = await GSheetDB.update(rec);
  hideLoadingToast && hideLoadingToast();
  if (r.isOk) { showToast('อนุมัติการลาสำเร็จ'); renderCurrentPage() } else showToast('เกิดข้อผิดพลาด: ' + (r.error || ''), 'error');
}

async function rejectLeave(id, approvalField) {
  const rec = APP.allData.find(d => d.__backendId === id); if (!rec) return;
  if (!confirm('ยืนยันการปฏิเสธใบลานี้?')) return;
  rec[approvalField] = 'ปฏิเสธ';
  rec.leave_status = 'ปฏิเสธ';

  showToast('กำลังบันทึก...', 'loading');
  const r = await GSheetDB.update(rec);
  hideLoadingToast && hideLoadingToast();
  if (r.isOk) { showToast('ปฏิเสธการลาสำเร็จ'); renderCurrentPage() } else showToast('เกิดข้อผิดพลาด: ' + (r.error || ''), 'error');
}

// Coordinator approves leave + fills in leave_percent
function showLeaveApprovalModal(id, currentPercent) {
  const rec = APP.allData.find(d => d.__backendId === id); if (!rec) return;
  showModal('อนุมัติใบลา', `
    <div class="space-y-3">
      <div class="bg-blue-50 rounded-xl p-3 text-sm space-y-1">
        <p><span class="text-gray-500">นักศึกษา:</span> <strong>${rec.name || '-'}</strong></p>
        <p><span class="text-gray-500">รายวิชา:</span> <strong>${rec.subject_name || '-'}</strong></p>
        <p><span class="text-gray-500">ประเภท:</span> <strong>${rec.leave_type || '-'}</strong> | <span class="text-gray-500">วันที่:</span> <strong>${toBuddhistDateList(rec.leave_date) || '-'}</strong> | <span class="text-gray-500">ชม.:</span> <strong>${rec.leave_hours || '-'}</strong></p>
        ${rec.leave_reason ? `<p><span class="text-gray-500">เหตุผล:</span> ${rec.leave_reason}</p>` : ''}
      </div>
      <form id="leaveApprovalForm" class="space-y-3">
        <div>
          <label class="block text-xs text-gray-600 mb-1">% การลาในรายวิชานี้ <span class="text-red-500">*</span></label>
          <div class="relative">
            <input name="leave_percent" id="leavePercentInput" type="number" step="0.01" min="0" max="100" required value="${currentPercent || ''}" class="w-full border rounded-xl px-3 py-2 pr-10 text-sm" placeholder="เช่น 5.5">
            <span class="absolute right-3 top-2.5 text-gray-400 text-sm">%</span>
          </div>
          <p class="text-xs text-gray-500 mt-1">กรอกเปอร์เซ็นต์การลาของนักศึกษาในรายวิชาตนเอง</p>
        </div>
        <div>
          <label class="block text-xs text-gray-600 mb-1">หมายเหตุ (ถ้ามี)</label>
          <textarea name="approval_note" rows="2" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="หมายเหตุจากอาจารย์ผู้ประสาน"></textarea>
        </div>
        <button type="submit" class="w-full bg-green-500 hover:bg-green-600 text-white py-2.5 rounded-xl flex items-center justify-center gap-2"><i data-lucide="check" class="w-4 h-4"></i>ยืนยันอนุมัติ</button>
      </form>
    </div>
  `);
  setTimeout(() => { lucide.createIcons(); const inp = document.getElementById('leavePercentInput'); if (inp) inp.focus(); }, 50);
  const f = document.getElementById('leaveApprovalForm');
  if (f) f.onsubmit = async (ev) => {
    ev.preventDefault();
    const fd = new FormData(f);
    const percent = fd.get('leave_percent');
    if (!percent && percent !== '0') { showToast('กรุณากรอก % การลา', 'error'); return; }
    const extra = { leave_percent: Number(percent) };
    const note = fd.get('approval_note');
    if (note) extra.coordinator_note = note;
    closeModal();
    await approveLeave(id, 'coordinator_approval', extra);
  };
}

// Class teacher (อาจารย์ประจำชั้น) approves leave + optional note
function showClassTeacherApprovalModal(id) {
  const rec = APP.allData.find(d => d.__backendId === id); if (!rec) return;
  showModal('อนุมัติใบลา (อาจารย์ประจำชั้น)', `
    <div class="space-y-3">
      <div class="bg-blue-50 rounded-xl p-3 text-sm space-y-1">
        <p><span class="text-gray-500">นักศึกษา:</span> <strong>${rec.name || '-'}</strong></p>
        <p><span class="text-gray-500">รายวิชา:</span> <strong>${rec.subject_name || '-'}</strong></p>
        <p><span class="text-gray-500">ประเภท:</span> <strong>${rec.leave_type || '-'}</strong> | <span class="text-gray-500">วันที่:</span> <strong>${toBuddhistDateList(rec.leave_date) || '-'}</strong> | <span class="text-gray-500">ชม.:</span> <strong>${rec.leave_hours || '-'}</strong></p>
        ${rec.leave_percent ? `<p><span class="text-gray-500">% การลา:</span> <strong>${rec.leave_percent}%</strong></p>` : ''}
        ${rec.leave_reason ? `<p><span class="text-gray-500">เหตุผล:</span> ${rec.leave_reason}</p>` : ''}
        ${rec.coordinator_note ? `<p><span class="text-gray-500">บันทึก ปสน.:</span> ${rec.coordinator_note}</p>` : ''}
      </div>
      <form id="classTeacherApprovalForm" class="space-y-3">
        <div>
          <label class="block text-xs text-gray-600 mb-1">หมายเหตุ (ถ้ามี — ไม่บังคับ)</label>
          <textarea name="class_teacher_note" id="classTeacherNoteInput" rows="3" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="หมายเหตุจากอาจารย์ประจำชั้น (ถ้าไม่มีก็กดยืนยันได้เลย)"></textarea>
        </div>
        <button type="submit" class="w-full bg-green-500 hover:bg-green-600 text-white py-2.5 rounded-xl flex items-center justify-center gap-2"><i data-lucide="check" class="w-4 h-4"></i>ยืนยันอนุมัติ</button>
      </form>
    </div>
  `);
  setTimeout(() => { lucide.createIcons(); const inp = document.getElementById('classTeacherNoteInput'); if (inp) inp.focus(); }, 50);
  const f = document.getElementById('classTeacherApprovalForm');
  if (f) f.onsubmit = async (ev) => {
    ev.preventDefault();
    const fd = new FormData(f);
    const note = (fd.get('class_teacher_note') || '').toString().trim();
    const extra = {};
    if (note) extra.class_teacher_note = note;
    closeModal();
    await approveLeave(id, 'class_teacher_approval', extra);
  };
}

// Executive (deputy) approves leave + optional note
function showExecutiveApprovalModal(id) {
  const rec = APP.allData.find(d => d.__backendId === id); if (!rec) return;
  showModal('อนุมัติใบลา (ผู้บริหาร)', `
    <div class="space-y-3">
      <div class="bg-blue-50 rounded-xl p-3 text-sm space-y-1">
        <p><span class="text-gray-500">นักศึกษา:</span> <strong>${rec.name || '-'}</strong></p>
        <p><span class="text-gray-500">รายวิชา:</span> <strong>${rec.subject_name || '-'}</strong></p>
        <p><span class="text-gray-500">ประเภท:</span> <strong>${rec.leave_type || '-'}</strong> | <span class="text-gray-500">วันที่:</span> <strong>${toBuddhistDateList(rec.leave_date) || '-'}</strong> | <span class="text-gray-500">ชม.:</span> <strong>${rec.leave_hours || '-'}</strong></p>
        ${rec.leave_percent ? `<p><span class="text-gray-500">% การลา:</span> <strong>${rec.leave_percent}%</strong></p>` : ''}
        ${rec.leave_reason ? `<p><span class="text-gray-500">เหตุผล:</span> ${rec.leave_reason}</p>` : ''}
        ${rec.coordinator_note ? `<p><span class="text-gray-500">บันทึก ปสน.:</span> ${rec.coordinator_note}</p>` : ''}
        ${rec.class_teacher_note ? `<p><span class="text-gray-500">บันทึก ปจช.:</span> ${rec.class_teacher_note}</p>` : ''}
      </div>
      <form id="execApprovalForm" class="space-y-3">
        <div>
          <label class="block text-xs text-gray-600 mb-1">หมายเหตุ (ถ้ามี — ไม่บังคับ)</label>
          <textarea name="deputy_note" id="deputyNoteInput" rows="3" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="หมายเหตุจากผู้บริหาร (ถ้าไม่มีก็กดยืนยันได้เลย)"></textarea>
        </div>
        <button type="submit" class="w-full bg-green-500 hover:bg-green-600 text-white py-2.5 rounded-xl flex items-center justify-center gap-2"><i data-lucide="check" class="w-4 h-4"></i>ยืนยันอนุมัติ</button>
      </form>
    </div>
  `);
  setTimeout(() => { lucide.createIcons(); const inp = document.getElementById('deputyNoteInput'); if (inp) inp.focus(); }, 50);
  const f = document.getElementById('execApprovalForm');
  if (f) f.onsubmit = async (ev) => {
    ev.preventDefault();
    const fd = new FormData(f);
    const note = (fd.get('deputy_note') || '').toString().trim();
    const extra = {};
    if (note) extra.deputy_note = note;
    closeModal();
    await approveLeave(id, 'deputy_approval', extra);
  };
}





// ======================== GLOBAL ACTIONS ========================
function changePage(p) { APP.pagination.page = p; renderCurrentPage() }

async function deleteRecord(id) {
  const rec = APP.allData.find(d => d.__backendId === id); if (!rec) return;
  showModal('ยืนยันการลบ', '<p class="text-center text-gray-600">คุณต้องการลบรายการนี้หรือไม่?</p>', async () => {
    showToast('กำลังลบข้อมูล...', 'loading');
    const r = await GSheetDB.delete(rec);
    hideLoadingToast();
    if (r.isOk) { showToast('ลบสำเร็จ'); closeModal(); renderCurrentPage() } else { showToast('เกิดข้อผิดพลาด', 'error'); closeModal() }
  });
}

// ======================== GENERIC EDIT HELPER ========================
async function editRecord(id, formId) {
  const rec = APP.allData.find(d => d.__backendId === id); if (!rec) return;
  const form = document.getElementById(formId); if (!form) return;
  const btn = form.querySelector('[type="submit"]');
  const origText = btn ? btn.innerHTML : '';
  if (btn) { btn.disabled = true; btn.innerHTML = '<span class="flex items-center justify-center gap-2"><i data-lucide="loader" class="w-4 h-4 animate-spin"></i>กำลังบันทึก...</span>'; lucide.createIcons() }
  const fd = new FormData(form);
  fd.forEach((v, k) => { if (k !== '__backendId') rec[k] = v });
  const r = await GSheetDB.update(rec);
  if (btn) { btn.disabled = false; btn.innerHTML = origText; lucide.createIcons() }
  if (r.isOk) { showToast('แก้ไขข้อมูลสำเร็จ'); closeModal(); renderCurrentPage() } else showToast('เกิดข้อผิดพลาด: ' + (r.error || ''), 'error');
}

// ======================== EDIT MODALS ========================
function showEditStudentModal(id) {
  const s = APP.allData.find(d => d.__backendId === id); if (!s) return;
  showModal('แก้ไขข้อมูลนักศึกษา', `
    <form id="editStudentForm" class="space-y-3">
      <div class="grid grid-cols-1 md:grid-cols-2 gap-3">
        <div><label class="block text-xs text-gray-600 mb-1">ชื่อ-สกุล</label><input name="name" value="${s.name || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">รหัสนักศึกษา</label><input name="student_id" value="${s.student_id || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">รุ่นที่</label><input name="batch" value="${s.batch || ''}" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="เช่น 36"></div>
        <div><label class="block text-xs text-gray-600 mb-1">เลขบัตรประชาชน</label><input name="national_id" value="${s.national_id || ''}" maxlength="13" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">สถานภาพ</label><select name="status" class="w-full border rounded-xl px-3 py-2 text-sm"><option ${s.status === 'กำลังศึกษา' ? 'selected' : ''}>กำลังศึกษา</option><option ${s.status === 'พักการศึกษา' ? 'selected' : ''}>พักการศึกษา</option><option ${s.status === 'สำเร็จการศึกษา' ? 'selected' : ''}>สำเร็จการศึกษา</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">ชั้นปี</label><select name="year_level" class="w-full border rounded-xl px-3 py-2 text-sm"><option ${norm(s.year_level) === '1' ? 'selected' : ''}>1</option><option ${norm(s.year_level) === '2' ? 'selected' : ''}>2</option><option ${norm(s.year_level) === '3' ? 'selected' : ''}>3</option><option ${norm(s.year_level) === '4' ? 'selected' : ''}>4</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">ห้อง</label><input name="room" value="${s.room || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">โทรศัพท์</label><input name="phone" value="${s.phone || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">E-mail</label><input name="email" value="${s.email || ''}" type="email" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">ชื่อผู้ปกครอง</label><input name="parent_name" value="${s.parent_name || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">โทรผู้ปกครอง</label><input name="parent_phone" value="${s.parent_phone || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">อาจารย์ที่ปรึกษา</label><input name="advisor" value="${s.advisor || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      </div>
      <button type="submit" class="w-full mt-3 bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึกการแก้ไข</button>
    </form>
  `);
  document.getElementById('editStudentForm').onsubmit = (e) => { e.preventDefault(); editRecord(id, 'editStudentForm') };
}

function showEditSubjectModal(id) {
  const s = APP.allData.find(d => d.__backendId === id); if (!s) return;
  showModal('แก้ไขรายวิชา', `
    <form id="editSubjectForm" class="space-y-3">
      <div><label class="block text-xs text-gray-600 mb-1">รหัสวิชา</label><input name="subject_code" value="${s.subject_code || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <div><label class="block text-xs text-gray-600 mb-1">ชื่อรายวิชา</label><input name="subject_name" value="${s.subject_name || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <div><label class="block text-xs text-gray-600 mb-1">ผู้ประสานงาน</label><input name="coordinator" value="${s.coordinator || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-xs text-gray-600 mb-1">ชั้นปี</label><select name="year_level" class="w-full border rounded-xl px-3 py-2 text-sm"><option ${norm(s.year_level) === '1' ? 'selected' : ''}>1</option><option ${norm(s.year_level) === '2' ? 'selected' : ''}>2</option><option ${norm(s.year_level) === '3' ? 'selected' : ''}>3</option><option ${norm(s.year_level) === '4' ? 'selected' : ''}>4</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">ห้อง</label><input name="room" value="${s.room || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">หน่วยกิต</label><input name="credits" type="number" value="${s.credits || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">ภาคการศึกษา</label><select name="semester" class="w-full border rounded-xl px-3 py-2 text-sm"><option value="1" ${norm(s.semester) === '1' ? 'selected' : ''}>1</option><option value="2" ${norm(s.semester) === '2' ? 'selected' : ''}>2</option><option value="3" ${norm(s.semester) === '3' ? 'selected' : ''}>ฤดูร้อน</option></select></div>
      </div>
      <div><label class="block text-xs text-gray-600 mb-1">ปีการศึกษา</label><input name="academic_year" value="${s.academic_year || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึกการแก้ไข</button>
    </form>
  `);
  document.getElementById('editSubjectForm').onsubmit = (e) => { e.preventDefault(); editRecord(id, 'editSubjectForm') };
}

function showEditScheduleModal(id) {
  const s = APP.allData.find(d => d.__backendId === id); if (!s) return;
  showModal('แก้ไขตารางเรียน/สอบ', `
    <form id="editScheduleForm" class="space-y-3">
      <div><label class="block text-xs text-gray-600 mb-1">รายวิชา</label><input name="subject_name" value="${s.subject_name || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-xs text-gray-600 mb-1">วันที่</label><input name="schedule_date" type="date" value="${s.schedule_date || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">เวลา</label><input name="schedule_time" type="time" value="${s.schedule_time || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">ประเภท</label><select name="schedule_type" class="w-full border rounded-xl px-3 py-2 text-sm"><option ${s.schedule_type === 'เรียน' ? 'selected' : ''}>เรียน</option><option ${s.schedule_type === 'สอบ' ? 'selected' : ''}>สอบ</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">ห้อง</label><input name="room" value="${s.room || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">ชั้นปี</label><select name="year_level" class="w-full border rounded-xl px-3 py-2 text-sm"><option ${norm(s.year_level) === '1' ? 'selected' : ''}>1</option><option ${norm(s.year_level) === '2' ? 'selected' : ''}>2</option><option ${norm(s.year_level) === '3' ? 'selected' : ''}>3</option><option ${norm(s.year_level) === '4' ? 'selected' : ''}>4</option></select></div>
      </div>
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึกการแก้ไข</button>
    </form>
  `);
  document.getElementById('editScheduleForm').onsubmit = (e) => { e.preventDefault(); editRecord(id, 'editScheduleForm') };
}

function showEditGradeModal(id) {
  const g = APP.allData.find(d => d.__backendId === id); if (!g) return;
  showModal('แก้ไขผลการเรียน', `
    <form id="editGradeForm" class="space-y-3">
      <div><label class="block text-xs text-gray-600 mb-1">นักศึกษา</label><select name="student_id" class="w-full border rounded-xl px-3 py-2 text-sm">${studentOptionsHTML(g.student_id)}</select></div>
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-xs text-gray-600 mb-1">รหัสวิชา</label><input name="subject_code" value="${g.subject_code || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">รายวิชา</label><input name="subject_name" value="${g.subject_name || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">เกรด</label><select name="grade" class="w-full border rounded-xl px-3 py-2 text-sm"><option ${g.grade === 'A' ? 'selected' : ''}>A</option><option ${g.grade === 'B+' ? 'selected' : ''}>B+</option><option ${g.grade === 'B' ? 'selected' : ''}>B</option><option ${g.grade === 'C+' ? 'selected' : ''}>C+</option><option ${g.grade === 'C' ? 'selected' : ''}>C</option><option ${g.grade === 'D+' ? 'selected' : ''}>D+</option><option ${g.grade === 'D' ? 'selected' : ''}>D</option><option ${g.grade === 'F' ? 'selected' : ''}>F</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">หน่วยกิต</label><input name="credits" type="number" value="${g.credits || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">ภาคการศึกษา</label><select name="semester" class="w-full border rounded-xl px-3 py-2 text-sm"><option value="1" ${norm(g.semester) === '1' ? 'selected' : ''}>1</option><option value="2" ${norm(g.semester) === '2' ? 'selected' : ''}>2</option></select></div>
      </div>
      <div><label class="block text-xs text-gray-600 mb-1">ปีการศึกษา</label><input name="academic_year" value="${g.academic_year || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึกการแก้ไข</button>
    </form>
  `);
  document.getElementById('editGradeForm').onsubmit = (e) => { e.preventDefault(); editRecord(id, 'editGradeForm') };
}

function showEditEngModal(id) {
  const e = APP.allData.find(d => d.__backendId === id); if (!e) return;
  const isSbch = e.eng_type === 'สบช.';
  const initTotal = Number(e.eng_score) || 0;
  const initLevel = e.eng_level || (isSbch ? getEngLevel(initTotal) : '');
  const initStatus = e.eng_status || (isSbch ? (initTotal >= 41 ? 'ผ่าน' : 'ไม่ผ่าน') : '');
  showModal('แก้ไขผลสอบภาษาอังกฤษ', `
    <form id="editEngForm" class="space-y-3">
      <div><label class="block text-xs text-gray-600 mb-1">นักศึกษา</label><select name="student_id" class="w-full border rounded-xl px-3 py-2 text-sm">${studentOptionsHTML(e.student_id)}</select></div>
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-xs text-gray-600 mb-1">รูปแบบการสอบ *</label>
          <select id="editEngType" onchange="updateEngTypeForm('edit')" class="w-full border rounded-xl px-3 py-2 text-sm">${engTypeOptions(e.eng_type || '')}</select>
        </div>
        <div><label class="block text-xs text-gray-600 mb-1">สอบครั้งที่</label><input id="editEngAttempt" type="number" value="${e.eng_attempt || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">วันที่สอบ <span class="text-gray-400">(วว/ดด/ปปปป พ.ศ. หรือ ค.ศ.)</span></label><input id="editEngDate" type="text" value="${formatDate(e.eng_date)}" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="เช่น 05/04/2568 หรือ 05/04/2025"></div>
        <div><label class="block text-xs text-gray-600 mb-1">ปีการศึกษา</label><input id="editEngYear" type="text" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="เช่น 2568" value="${e.academic_year || ''}"></div>
      </div>
      <!-- สบช. fields -->
      <div id="editEngSbch" style="display:${isSbch ? 'block' : 'none'}">
        <p class="text-xs font-semibold text-blue-700 mb-2">คะแนนรายทักษะ (สบช.)</p>
        <div class="grid grid-cols-3 gap-2 mb-2">
          <div><label class="block text-xs text-gray-600 mb-1">Listening</label><input id="editEngL" type="number" min="0" max="100" value="${e.eng_listening || ''}" oninput="calcEngSbchTotal('edit')" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="0-100"></div>
          <div><label class="block text-xs text-gray-600 mb-1">Grammar</label><input id="editEngG" type="number" min="0" max="100" value="${e.eng_grammar || ''}" oninput="calcEngSbchTotal('edit')" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="0-100"></div>
          <div><label class="block text-xs text-gray-600 mb-1">Reading</label><input id="editEngR" type="number" min="0" max="100" value="${e.eng_reading || ''}" oninput="calcEngSbchTotal('edit')" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="0-100"></div>
        </div>
        <div class="grid grid-cols-2 gap-2">
          <div><label class="block text-xs text-gray-600 mb-1">คะแนนรวม (อัตโนมัติ)</label><input id="editEngTotal" type="number" value="${initTotal || ''}" readonly class="w-full border rounded-xl px-3 py-2 text-sm bg-gray-50"></div>
          <div><label class="block text-xs text-gray-600 mb-1">ระดับผลการสอบ (อัตโนมัติ)</label><input id="editEngLevel" value="${initLevel}" readonly class="w-full border rounded-xl px-3 py-2 text-sm bg-gray-50"></div>
        </div>
        <div class="mt-2"><label class="block text-xs text-gray-600 mb-1">สถานะ (อัตโนมัติ)</label><div class="border rounded-xl px-3 py-2 bg-gray-50 min-h-[36px] flex items-center"><span id="editEngStatus" class="inline-block px-3 py-1 rounded-full text-sm font-semibold ${initStatus === 'ผ่าน' ? 'bg-green-100 text-green-700' : 'bg-red-100 text-red-700'}">${initStatus || '-'}</span></div></div>
      </div>
      <!-- Other type fields -->
      <div id="editEngOther" style="display:${isSbch ? 'none' : 'block'}">
        <div class="grid grid-cols-2 gap-3">
          <div><label class="block text-xs text-gray-600 mb-1">คะแนนรวม</label><input id="editEngOtherScore" type="number" value="${!isSbch ? (e.eng_score || '') : ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
          <div><label class="block text-xs text-gray-600 mb-1">สถานะ *</label><select id="editEngOtherStatus" class="w-full border rounded-xl px-3 py-2 text-sm"><option value="ผ่าน" ${e.eng_status === 'ผ่าน' ? 'selected' : ''}>ผ่าน</option><option value="ไม่ผ่าน" ${e.eng_status === 'ไม่ผ่าน' ? 'selected' : ''}>ไม่ผ่าน</option></select></div>
        </div>
      </div>
      <!-- ไม่เข้าสอบ -->
      <label class="flex items-center gap-2 cursor-pointer select-none">
        <input type="checkbox" id="editEngAbsent" onchange="toggleEngAbsent('edit')" class="w-4 h-4 rounded border-gray-300 text-red-500 focus:ring-red-400" ${e.eng_status === 'ไม่เข้าสอบ' ? 'checked' : ''}>
        <span class="text-sm text-red-600 font-medium">ไม่เข้าสอบ</span>
      </label>
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึกการแก้ไข</button>
    </form>
  `);
  // If currently absent, hide score fields
  if (e.eng_status === 'ไม่เข้าสอบ') toggleEngAbsent('edit');
  document.getElementById('editEngForm').onsubmit = async (ev) => {
    ev.preventDefault();
    await withLoading(ev.target, async () => {
      const absent = document.getElementById('editEngAbsent').checked;
      const engType = document.getElementById('editEngType').value;
      if (!absent && !engType) { showToast('กรุณาเลือกรูปแบบการสอบ', 'error'); return; }
      const studentId = ev.target.querySelector('[name="student_id"]').value;
      const attempt = document.getElementById('editEngAttempt').value;
      const date = normalizeDateInput(document.getElementById('editEngDate').value);
      const year = document.getElementById('editEngYear').value;
      const obj = { ...e, student_id: studentId, eng_type: engType, eng_attempt: Number(attempt) || '', eng_date: date, academic_year: year };
      if (absent) {
        obj.eng_status = 'ไม่เข้าสอบ';
        obj.eng_score = ''; obj.eng_listening = ''; obj.eng_grammar = ''; obj.eng_reading = ''; obj.eng_level = '';
      } else if (engType === 'สบช.') {
        const l = Number(document.getElementById('editEngL').value) || 0;
        const g = Number(document.getElementById('editEngG').value) || 0;
        const r = Number(document.getElementById('editEngR').value) || 0;
        const total = l + g + r;
        obj.eng_listening = l; obj.eng_grammar = g; obj.eng_reading = r;
        obj.eng_score = total; obj.eng_level = getEngLevel(total);
        obj.eng_status = total >= 41 ? 'ผ่าน' : 'ไม่ผ่าน';
      } else {
        obj.eng_score = Number(document.getElementById('editEngOtherScore').value) || '';
        obj.eng_status = document.getElementById('editEngOtherStatus').value;
        obj.eng_listening = ''; obj.eng_grammar = ''; obj.eng_reading = ''; obj.eng_level = '';
      }
      const res = await GSheetDB.update(obj);
      if (res.isOk) { showToast('แก้ไขสำเร็จ'); closeModal(); } else showToast('เกิดข้อผิดพลาด', 'error');
    });
  };
}

// showEditEvalFormModal removed (teacher evaluation feature has been removed)

function showEditTeacherModal(id) {
  const t = APP.allData.find(d => d.__backendId === id); if (!t) return;
  showModal('แก้ไขข้อมูลอาจารย์', `
    <form id="editTeacherForm" class="space-y-3">
      <div><label class="block text-xs text-gray-600 mb-1">ชื่อ-สกุล</label><input name="name" value="${t.name || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-xs text-gray-600 mb-1">ตำแหน่ง</label><input name="position" value="${t.position || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">สาขาวิชา</label><input name="department" value="${t.department || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">โทรศัพท์</label><input name="phone" value="${t.phone || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">E-mail</label><input name="email" value="${t.email || ''}" type="email" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">ชั้นปีที่รับผิดชอบ</label><select name="responsible_year" class="w-full border rounded-xl px-3 py-2 text-sm"><option value="">ไม่มี</option><option ${norm(t.responsible_year) === '1' ? 'selected' : ''}>1</option><option ${norm(t.responsible_year) === '2' ? 'selected' : ''}>2</option><option ${norm(t.responsible_year) === '3' ? 'selected' : ''}>3</option><option ${norm(t.responsible_year) === '4' ? 'selected' : ''}>4</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">สถานะ</label><select name="teacher_status" class="w-full border rounded-xl px-3 py-2 text-sm"><option ${(t.teacher_status || 'ปฏิบัติงานอยู่') === 'ปฏิบัติงานอยู่' ? 'selected' : ''}>ปฏิบัติงานอยู่</option><option ${t.teacher_status === 'ลาศึกษาต่อ' ? 'selected' : ''}>ลาศึกษาต่อ</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">เลขบัญชีธนาคาร</label><input name="bank_account" value="${t.bank_account || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      </div>
      <div><label class="block text-xs text-gray-600 mb-1">ที่อยู่</label><textarea name="address" rows="2" class="w-full border rounded-xl px-3 py-2 text-sm">${t.address || ''}</textarea></div>
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึกการแก้ไข</button>
    </form>
  `);
  document.getElementById('editTeacherForm').onsubmit = (e) => { e.preventDefault(); editRecord(id, 'editTeacherForm') };
}

function showEditAnnouncementModal(id) {
  const a = APP.allData.find(d => d.__backendId === id); if (!a) return;
  showModal('แก้ไขประกาศ', `
    <form id="editAnnForm" class="space-y-3">
      <div><label class="block text-xs text-gray-600 mb-1">เรื่อง</label><input name="announcement_title" value="${a.announcement_title || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <div><label class="block text-xs text-gray-600 mb-1">เนื้อหา</label><textarea name="announcement_content" rows="3" class="w-full border rounded-xl px-3 py-2 text-sm">${a.announcement_content || ''}</textarea></div>
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-xs text-gray-600 mb-1">วันที่</label><input name="announcement_date" type="date" value="${a.announcement_date || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">ประเภท</label><select name="event_type" class="w-full border rounded-xl px-3 py-2 text-sm"><option ${a.event_type === 'ทั่วไป' ? 'selected' : ''}>ทั่วไป</option><option ${a.event_type === 'สอบ' ? 'selected' : ''}>สอบ</option><option ${a.event_type === 'วันหยุด' ? 'selected' : ''}>วันหยุด</option><option ${a.event_type === 'กิจกรรม' ? 'selected' : ''}>กิจกรรม</option></select></div>
      </div>
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึกการแก้ไข</button>
    </form>
  `);
  document.getElementById('editAnnForm').onsubmit = (e) => { e.preventDefault(); editRecord(id, 'editAnnForm') };
}

function showEditTrackingModal(id) {
  const t = APP.allData.find(d => d.__backendId === id); if (!t) return;
  showModal('แก้ไขข้อมูลติดตามรายวิชา', `
    <form id="editTrackingForm" class="space-y-3">
      <div><label class="block text-xs text-gray-600 mb-1">ชื่อรายวิชา</label><input name="subject_name" value="${t.subject_name || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-xs text-gray-600 mb-1">ทฤษฎี/ปฏิบัติ</label><select name="theory_practice" class="w-full border rounded-xl px-3 py-2 text-sm"><option ${t.theory_practice === 'ทฤษฎี' ? 'selected' : ''}>ทฤษฎี</option><option ${t.theory_practice === 'ปฏิบัติ' ? 'selected' : ''}>ปฏิบัติ</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">ชั้นปี</label><select name="year_level" class="w-full border rounded-xl px-3 py-2 text-sm"><option ${norm(t.year_level) === '1' ? 'selected' : ''}>1</option><option ${norm(t.year_level) === '2' ? 'selected' : ''}>2</option><option ${norm(t.year_level) === '3' ? 'selected' : ''}>3</option><option ${norm(t.year_level) === '4' ? 'selected' : ''}>4</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">ห้อง</label><input name="room" value="${t.room || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">ภาคการศึกษา</label><select name="semester" class="w-full border rounded-xl px-3 py-2 text-sm"><option value="1" ${norm(t.semester) === '1' ? 'selected' : ''}>1</option><option value="2" ${norm(t.semester) === '2' ? 'selected' : ''}>2</option></select></div>
      </div>
      <div><label class="block text-xs text-gray-600 mb-1">ผู้ประสานงาน</label><input name="coordinator" value="${t.coordinator || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <div><label class="block text-xs text-gray-600 mb-1">วันอนุมัติ</label><input name="approved_date" type="date" value="${t.approved_date || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึกการแก้ไข</button>
    </form>
  `);
  document.getElementById('editTrackingForm').onsubmit = (e) => { e.preventDefault(); editRecord(id, 'editTrackingForm') };
}

function showEditLeaveModal(id) {
  const l = APP.allData.find(d => d.__backendId === id); if (!l) return;
  showModal('แก้ไขข้อมูลการลา', `
    <form id="editLeaveForm" class="space-y-3">
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-xs text-gray-600 mb-1">ชื่อ-สกุล</label><input name="name" value="${l.name || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">รายวิชา</label><input name="subject_name" value="${l.subject_name || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">จำนวนชั่วโมง</label><input name="leave_hours" type="number" value="${l.leave_hours || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">% การลา</label><input name="leave_percent" type="number" value="${l.leave_percent || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">วันที่ลา</label><input name="leave_date" type="date" value="${l.leave_date || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">ประเภท</label><select name="leave_type" class="w-full border rounded-xl px-3 py-2 text-sm"><option ${l.leave_type === 'ลาป่วย' ? 'selected' : ''}>ลาป่วย</option><option ${l.leave_type === 'ลากิจ' ? 'selected' : ''}>ลากิจ</option><option ${l.leave_type === 'ลาพบแพทย์' ? 'selected' : ''}>ลาพบแพทย์</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">สถานะ</label><select name="leave_status" class="w-full border rounded-xl px-3 py-2 text-sm"><option ${l.leave_status === 'รออนุมัติ' ? 'selected' : ''}>รออนุมัติ</option><option ${l.leave_status === 'อนุมัติแล้ว' ? 'selected' : ''}>อนุมัติแล้ว</option><option ${l.leave_status === 'ปฏิเสธ' ? 'selected' : ''}>ปฏิเสธ</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">ภาคการศึกษา</label><select name="semester" class="w-full border rounded-xl px-3 py-2 text-sm"><option value="1" ${norm(l.semester) === '1' ? 'selected' : ''}>1</option><option value="2" ${norm(l.semester) === '2' ? 'selected' : ''}>2</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">ปีการศึกษา</label><input name="academic_year" value="${l.academic_year || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      </div>
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึกการแก้ไข</button>
    </form>
  `);
  document.getElementById('editLeaveForm').onsubmit = (e) => { e.preventDefault(); editRecord(id, 'editLeaveForm') };
}

function showEditUserModal(id) {
  const u = APP.allData.find(d => d.__backendId === id); if (!u) return;
  const roleLabels = { admin: 'ผู้ดูแลระบบ', academic: 'เจ้าหน้าที่งานวิชาการ', executive: 'ผู้บริหาร', teacher: 'อาจารย์', classTeacher: 'อาจารย์ประจำชั้น', student: 'นักศึกษา' };
  showModal('แก้ไขผู้ใช้งาน', `
    <form id="editUserForm" class="space-y-3">
      <div><label class="block text-xs text-gray-600 mb-1">ชื่อ-สกุล</label><input name="name" value="${u.name || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <div><label class="block text-xs text-gray-600 mb-1">บทบาท</label><select name="role" class="w-full border rounded-xl px-3 py-2 text-sm"><option value="admin" ${u.role === 'admin' ? 'selected' : ''}>ผู้ดูแลระบบ</option><option value="academic" ${u.role === 'academic' ? 'selected' : ''}>เจ้าหน้าที่งานวิชาการ</option><option value="executive" ${u.role === 'executive' ? 'selected' : ''}>ผู้บริหาร</option><option value="teacher" ${u.role === 'teacher' ? 'selected' : ''}>อาจารย์</option><option value="classTeacher" ${u.role === 'classTeacher' ? 'selected' : ''}>อาจารย์ประจำชั้น</option><option value="student" ${u.role === 'student' ? 'selected' : ''}>นักศึกษา</option></select></div>
      <div><label class="block text-xs text-gray-600 mb-1">E-mail</label><input name="email" value="${u.email || ''}" type="email" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <div><label class="block text-xs text-gray-600 mb-1">รหัสผ่าน</label><input name="password" value="${u.password || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <div><label class="block text-xs text-gray-600 mb-1">เลขบัตรประชาชน</label><input name="national_id" value="${u.national_id || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <div><label class="block text-xs text-gray-600 mb-1">ชั้นปีที่รับผิดชอบ</label><select name="responsible_year" class="w-full border rounded-xl px-3 py-2 text-sm"><option value="">ไม่มี</option><option ${norm(u.responsible_year) === '1' ? 'selected' : ''}>1</option><option ${norm(u.responsible_year) === '2' ? 'selected' : ''}>2</option><option ${norm(u.responsible_year) === '3' ? 'selected' : ''}>3</option><option ${norm(u.responsible_year) === '4' ? 'selected' : ''}>4</option></select></div>
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึกการแก้ไข</button>
    </form>
  `);
  document.getElementById('editUserForm').onsubmit = (e) => { e.preventDefault(); editRecord(id, 'editUserForm') };
}

// Init icons
lucide.createIcons();
