// ======================== STATE ========================
let APP = {
  currentUser: null, currentRole: null, currentPage: 'dashboard', sidebarOpen: false,
  allData: [],
  config: { system_title: 'ระบบบริหารจัดการงานวิชาการ (EMS-BCNB)', college_name: 'วิทยาลัยพยาบาลบรมราชชนนี กรุงเทพ' },
  permissions: { admin: { dashboard: 1, students: 1, subjects: 1, schedule: 1, grades: 1, engResults: 1, evalTeacher: 1, teachers: 1, teacherDirectory: 1, services: 1, tracking: 1, gradeTracking: 1, fileTracking: 1, leave: 1, settings: 1 }, academic: { dashboard: 1, students: 1, subjects: 1, schedule: 1, grades: 1, engResults: 1, evalTeacher: 1, teachers: 1, teacherDirectory: 1, services: 1, tracking: 1, gradeTracking: 1, fileTracking: 1, leave: 1, settings: 1 }, teacher: { dashboard: 1, students: 1, subjects: 1, grades: 1, engResults: 1, evalTeacher: 1, tracking: 1, gradeTracking: 1, fileTracking: 1, leave: 1 }, classTeacher: { dashboard: 1, students: 1, subjects: 1, grades: 1, engResults: 1, tracking: 1, gradeTracking: 1, fileTracking: 1, leave: 1 }, student: { dashboard: 1, students: 1, grades: 1, evalTeacher: 1, leave: 1 }, executive: { dashboard: 1, students: 1, subjects: 1, schedule: 1, grades: 1, engResults: 1, evalTeacher: 1, teachers: 1, teacherDirectory: 1, tracking: 1, gradeTracking: 1, fileTracking: 1, leave: 1 } },
  filters: { semester: '', academicYear: '', search: '', yearLevel: '' },
  pagination: { page: 1, perPage: 10 }
};

function getDataByType(t) { return APP.allData.filter(d => d.type === t) }

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
          const debugPwds = adminUsers.map(u => '"' + String(u.password || '') + '"→"' + cleanPassword(u.password) + '"').join(', ');
          err.innerHTML = 'รหัสผ่านไม่ถูกต้อง (พบ admin ' + adminUsers.length + ' คน, ค่าที่กรอก="' + p + '", ค่าในระบบ: ' + debugPwds + ')<br><button onclick="debugConnection()" class="mt-2 text-xs underline text-primary">ตรวจสอบการเชื่อมต่อ</button> <button onclick="GSheetDB.clearConfig();location.reload()" class="mt-2 text-xs underline text-red-500 ml-3">รีเซ็ต Google Sheet</button>';
        }
      }
      err.classList.remove('hidden'); return
    }
    APP.currentUser = { name: adminUser.name || 'ผู้ดูแลระบบ', role: 'admin' };
  } else if (role === 'student') {
    const nid = document.getElementById('studentNID').value;
    if (!/^\d{13}$/.test(nid)) { err.textContent = 'กรุณากรอกเลขบัตรประชาชน 13 หลัก'; err.classList.remove('hidden'); return }
    const stu = getDataByType('student').find(s => s.national_id === nid);
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
    const em = document.getElementById('teacherEmail').value;
    const pw = document.getElementById('teacherPass').value;
    if (!em) { err.textContent = 'กรุณากรอก E-mail'; err.classList.remove('hidden'); return }
    if (!pw) { err.textContent = 'กรุณากรอกรหัสผ่าน'; err.classList.remove('hidden'); return }
    const user = getDataByType('user').find(u => normalizeRole(u.role) === 'classTeacher' && String(u.email || '').trim().toLowerCase() === em.trim().toLowerCase() && cleanPassword(u.password) === pw);
    if (!user) { err.textContent = 'E-mail หรือรหัสผ่านไม่ถูกต้อง'; err.classList.remove('hidden'); return }
    const t = getDataByType('teacher').find(x => x.email === em);
    APP.currentUser = t ? { name: t.name, role: 'classTeacher', data: t, responsible_year: t.responsible_year || user.responsible_year || '1' } : { name: user.name || em, role: 'classTeacher', email: em, responsible_year: user.responsible_year || '1' };
  } else if (role === 'executive') {
    const em = document.getElementById('teacherEmail').value;
    const pw = document.getElementById('teacherPass').value;
    if (!em) { err.textContent = 'กรุณากรอก E-mail'; err.classList.remove('hidden'); return }
    if (!pw) { err.textContent = 'กรุณากรอกรหัสผ่าน'; err.classList.remove('hidden'); return }
    const user = getDataByType('user').find(u => normalizeRole(u.role) === 'executive' && String(u.email || '').trim().toLowerCase() === em.trim().toLowerCase() && cleanPassword(u.password) === pw);
    if (!user) { err.textContent = 'E-mail หรือรหัสผ่านไม่ถูกต้อง'; err.classList.remove('hidden'); return }
    APP.currentUser = { name: user.name || em, role: 'executive', email: em };
  }
  APP.currentRole = APP.currentUser.role;
  showScreen('mainApp');
  document.getElementById('currentUserName').textContent = APP.currentUser.name;
  document.getElementById('currentUserRole').textContent = { admin: 'ผู้ดูแลระบบ', academic: 'เจ้าหน้าที่งานวิชาการ', teacher: 'อาจารย์', classTeacher: 'อาจารย์ประจำชั้น', student: 'นักศึกษา', executive: 'ผู้บริหาร' }[APP.currentRole];
  buildSidebar();
  navigateTo('dashboard');
  lucide.createIcons();
}

function handleLogout() {
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
  if (p.evalTeacher) regSub.push({ id: 'evalTeacher', label: 'ประเมินอาจารย์ผู้สอน' });
  if (p.teachers) regSub.push({ id: 'teachers', label: 'ข้อมูลอาจารย์' });
  if (regSub.length) items.push({ id: 'registration', icon: 'book-open', label: 'ระบบทะเบียน', sub: regSub });

  if (p.teacherDirectory) items.push({ id: 'teacherDirectory', icon: 'award', label: 'ทำเนียบอาจารย์' });

  // Tracking dropdown
  let trackSub = [];
  if (p.tracking) trackSub.push({ id: 'tracking', label: 'ส่งรายละเอียดรายวิชา' });
  if (p.gradeTracking) trackSub.push({ id: 'gradeTracking', label: 'ส่งเกรดรายวิชา' });
  if (p.fileTracking) trackSub.push({ id: 'fileTracking', label: 'ส่งแฟ้มรายวิชา' });
  if (trackSub.length) items.push({ id: 'trackingGroup', icon: 'clipboard-list', label: 'ติดตามการส่ง', sub: trackSub });

  if (p.leave) items.push({ id: 'leave', icon: 'calendar-off', label: 'ระบบการลาของนักศึกษา' });
  if (p.services) items.push({ id: 'services', icon: 'grid', label: 'บริการอื่นๆ' });
  if ((r === 'admin' || r === 'academic') && p.settings) items.push({ id: 'settings', icon: 'settings', label: 'ตั้งค่าระบบ' });

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
  APP.filters._evalTeacher = '';
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
  const btn = btnOrForm.tagName === 'FORM' ? btnOrForm.querySelector('[type="submit"]') : btnOrForm;
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
  if (APP.filters.semester) d = d.filter(x => norm(x.semester) === APP.filters.semester);
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
    case 'evalTeacher': return evalTeacherPage();
    case 'teachers': return teachersPage();
    case 'teacherDirectory': return teacherDirectoryPage();
    case 'services': return servicesPage();
    case 'tracking': return trackingPage();
    case 'gradeTracking': return gradeTrackingPage();
    case 'fileTracking': return fileTrackingPage();
    case 'leave': return leavePage();
    case 'settings': return settingsPage();
    default: return '<p>ไม่พบหน้าที่ต้องการ</p>';
  }
}

// ======================== DASHBOARD ========================
function dashboardPage() {
  const students = getDataByType('student');
  const teachers = getDataByType('teacher');
  const engPass = getDataByType('eng_result').filter(e => e.eng_status === 'ผ่าน');
  const announcements = getDataByType('announcement').slice(-5).reverse();
  const r = APP.currentRole;

  let stats = '';
  if (r === 'admin' || r === 'academic') {
    // Teacher breakdown by department
    const deptMap = {};
    teachers.forEach(t => {
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
      ${statCard('briefcase', 'จำนวนอาจารย์ทั้งหมด', teachers.length, 'คน', 'bg-emerald-500')}
      ${statCard('check-circle', 'สอบผ่านภาษาอังกฤษ', engPass.length, 'คน', 'bg-amber-500')}
    </div>
    <h3 class="font-bold mb-3 text-gray-800">จำนวนอาจารย์แยกสาขาวิชา</h3>
    <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3 mb-6">${deptCards || '<p class="text-gray-400 text-sm col-span-3">ไม่มีข้อมูลสาขาวิชา</p>'}</div>
    <h3 class="font-bold mb-3 text-gray-800">นักศึกษารายชั้นปี</h3>
    <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 mb-6">
      ${[1, 2, 3, 4].map(yr => {
      const yrStudents = students.filter(s => norm(s.year_level) === String(yr));
      const yrEngPass = engPass.filter(e => yrStudents.some(s => s.student_id === e.student_id));
      return `<div class="bg-white rounded-2xl p-4 border border-blue-100">
          <p class="text-sm text-gray-500">ชั้นปี ${yr}</p>
          <div class="flex gap-3 mt-2">
            <div><p class="text-2xl font-bold text-primary">${yrStudents.length}</p><p class="text-xs text-gray-500">นักศึกษา</p></div>
            <div><p class="text-2xl font-bold text-green-500">${yrEngPass.length}</p><p class="text-xs text-gray-500">ผ่าน ENG</p></div>
          </div>
        </div>`;
    }).join('')}
    </div>`;
  } else if (r === 'teacher') {
    const myStudents = students.filter(s => s.advisor === APP.currentUser.name);
    const myEngPass = getDataByType('eng_result').filter(e => e.eng_status === 'ผ่าน' && myStudents.some(s => s.student_id === e.student_id));
    stats = `
    <div class="bg-white rounded-2xl p-5 border border-blue-100 mb-4"><p class="text-sm text-gray-500">ข้อมูลตนเอง</p><p class="font-bold text-lg">${APP.currentUser.name}</p></div>
    <div class="grid grid-cols-1 sm:grid-cols-2 gap-4 mb-6">
      ${statCard('users', 'นักศึกษาในที่ปรึกษา', myStudents.length, 'คน', 'bg-blue-500')}
      ${statCard('check-circle', 'สอบผ่านภาษาอังกฤษ', myEngPass.length, 'คน', 'bg-amber-500')}
    </div>`;
  } else if (r === 'classTeacher') {
    const yr = APP.currentUser.responsible_year || '1';
    const myStudents = students.filter(s => norm(s.year_level) === norm(yr));
    const myEngPass = getDataByType('eng_result').filter(e => e.eng_status === 'ผ่าน' && myStudents.some(s => s.student_id === e.student_id));
    stats = `
    <div class="bg-white rounded-2xl p-5 border border-blue-100 mb-4"><p class="text-sm text-gray-500">อาจารย์ประจำชั้นปีที่ ${yr}</p><p class="font-bold text-lg">${APP.currentUser.name}</p></div>
    <div class="grid grid-cols-1 sm:grid-cols-2 gap-4 mb-6">
      ${statCard('users', 'จำนวนนักศึกษา', myStudents.length, 'คน', 'bg-blue-500')}
      ${statCard('check-circle', 'สอบผ่านภาษาอังกฤษ', myEngPass.length, 'คน', 'bg-amber-500')}
    </div>`;
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

  let headerHtml = `<div class="flex flex-wrap items-center justify-between gap-3 mb-4">
    <h2 class="text-xl font-bold text-gray-800"><i data-lucide="users" class="w-6 h-6 inline mr-2"></i>ข้อมูลนักศึกษา</h2>
    ${isAdmin ? `<div class="flex gap-2">${APP.currentRole === 'admin' ? `<button onclick="showPromoteYearModal()" class="flex items-center gap-2 px-4 py-2 bg-emerald-500 text-white rounded-xl hover:bg-emerald-600 text-sm"><i data-lucide="arrow-up-circle" class="w-4 h-4"></i>เลื่อนชั้นปี</button>` : ''}<button onclick="showAddStudentModal()" class="flex items-center gap-2 px-4 py-2 bg-primary text-white rounded-xl hover:bg-primaryDark text-sm"><i data-lucide="plus" class="w-4 h-4"></i>เพิ่มนักศึกษา</button>${csvUploadBtn('student', 'name,student_id,batch,status,phone,email,parent_name,parent_phone,advisor,year_level,room,national_id')}</div>` : ''}
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
  if (isClassTeacher) data = data.filter(s => norm(s.year_level) === norm(APP.currentUser.responsible_year || '1'));
  if (APP.currentRole === 'teacher') data = data.filter(s => s.advisor === APP.currentUser.name);
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
    const fd = new FormData(e.target);
    const obj = { type: 'student', created_at: new Date().toISOString() };
    fd.forEach((v, k) => obj[k] = v);
    if (APP.allData.filter(d => d.type === 'student').length >= 999) { showToast('ข้อมูลเต็ม', 'error'); return }

    const r = await GSheetDB.create(obj);
    if (r.isOk) { showToast('เพิ่มนักศึกษาสำเร็จ'); closeModal() } else showToast('เกิดข้อผิดพลาด', 'error');
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
        <td class="px-4 py-3">${s.semester || ''}/${s.academic_year || ''}</td>
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
    e.preventDefault(); const fd = new FormData(e.target);
    const obj = { type: 'subject', created_at: new Date().toISOString() };
    fd.forEach((v, k) => obj[k] = k === 'credits' ? Number(v) : v);
    if (APP.allData.filter(d => d.type === 'subject').length >= 999) { showToast('ข้อมูลเต็ม', 'error'); return }

    const r = await GSheetDB.create(obj);
    if (r.isOk) { showToast('เพิ่มรายวิชาสำเร็จ'); closeModal() } else showToast('เกิดข้อผิดพลาด', 'error');
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
    e.preventDefault(); const fd = new FormData(e.target);
    const obj = { type: 'schedule', created_at: new Date().toISOString() }; fd.forEach((v, k) => obj[k] = v);

    const r = await GSheetDB.create(obj);
    if (r.isOk) { showToast('เพิ่มตารางสำเร็จ'); closeModal() } else showToast('เกิดข้อผิดพลาด', 'error');
  };
}

// ======================== GRADES ========================
function gradesPage() {
  const isAdmin = APP.currentRole === 'admin' || APP.currentRole === 'academic';
  const canEdit = isAdmin || APP.currentRole === 'teacher' || APP.currentRole === 'classTeacher';
  const isStudent = APP.currentRole === 'student';
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
    // Store student list for search
    window._gradeStudentList = studentList;
    const searchVal = APP.filters._gradeSearch || '';
    let filteredList = studentList;
    if (searchVal) {
      const q = searchVal.toLowerCase();
      filteredList = studentList.filter(s => (s.name || '').toLowerCase().includes(q) || (s.student_id || '').toLowerCase().includes(q));
    }
    studentSelector = `<div class="bg-white rounded-2xl p-4 border border-blue-100 mb-4">
      <label class="block text-sm font-medium text-gray-700 mb-2"><i data-lucide="user-search" class="w-4 h-4 inline mr-1"></i>เลือกนักศึกษา</label>
      <div class="flex gap-2 mb-2">
        <div class="flex-1 relative"><i data-lucide="search" class="absolute left-3 top-2.5 w-4 h-4 text-gray-400"></i><input type="text" placeholder="พิมพ์ค้นหาชื่อหรือรหัส..." value="${searchVal}" class="w-full pl-10 pr-4 py-2 border border-gray-200 rounded-xl text-sm" onkeyup="APP.filters._gradeSearch=this.value;APP.filters._gradeStudent='';APP.pagination.page=1;renderCurrentPage()"></div>
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
    data = allGrades.filter(g => g.student_id === APP.currentUser.data.student_id);
  } else if (selectedStudentName) {
    data = allGrades.filter(g => g.student_id === selectedStudentName);
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
        <td class="px-4 py-3">${g.credits || ''}</td><td class="px-4 py-3">${g.semester || ''}/${g.academic_year || ''}</td>
        ${isAdmin ? `<td class="px-4 py-3"><div class="flex gap-1"><button onclick="showEditGradeModal('${g.__backendId}')" class="text-blue-400 hover:text-blue-600" title="แก้ไข"><i data-lucide="pencil" class="w-4 h-4"></i></button><button onclick="deleteRecord('${g.__backendId}')" class="text-red-400 hover:text-red-600" title="ลบ"><i data-lucide="trash-2" class="w-4 h-4"></i></button></div></td>` : ''}</tr>`).join('') : '<tr><td colspan="5" class="px-4 py-8 text-center text-gray-400">ไม่มีข้อมูล</td></tr>'}</tbody>
    </table></div>
  </div>
  ${paginationHTML(total, APP.pagination.perPage, APP.pagination.page, 'changePage')}`}`;
}

function showAddGradeModal() {
  showModal('เพิ่มผลการเรียน', `
    <form id="addGradeForm" class="space-y-3">
      <div><label class="block text-xs text-gray-600 mb-1">นักศึกษา *</label><select name="student_id" required class="w-full border rounded-xl px-3 py-2 text-sm">${studentOptionsHTML()}</select></div>
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
    e.preventDefault(); const fd = new FormData(e.target);
    const obj = { type: 'grade', created_at: new Date().toISOString() }; fd.forEach((v, k) => obj[k] = k === 'credits' ? Number(v) : v);

    const r = await GSheetDB.create(obj);
    if (r.isOk) { showToast('เพิ่มผลการเรียนสำเร็จ'); closeModal() } else showToast('เกิดข้อผิดพลาด', 'error');
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
  const grades = getDataByType('grade').filter(g => g.student_id === stu.student_id);
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

  showModal('ใบรายงานผลการเรียน', `
    <div id="transcriptContent" class="bg-white p-4" style="max-width:700px;margin:auto;">
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
    <div class="flex justify-center mt-4">
      <button onclick="downloadTranscriptPDF('${stuNameSafe}')" class="flex items-center gap-2 px-6 py-2.5 bg-primary text-white rounded-xl hover:bg-primaryDark text-sm"><i data-lucide="download" class="w-4 h-4"></i>ดาวน์โหลด PDF</button>
    </div>
  `, null, 'max-w-3xl');
  setTimeout(() => lucide.createIcons(), 100);
}

async function downloadTranscriptPDF(studentKey) {
  const stu = getDataByType('student').find(s => s.student_id === studentKey || s.name === studentKey) || (APP.currentUser.data && (APP.currentUser.data.name === studentKey || APP.currentUser.data.student_id === studentKey) ? APP.currentUser.data : null);
  if (!stu) return;
  const grades = getDataByType('grade').filter(g => g.student_id === stu.student_id);
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

  const htmlContent = `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>ใบรายงานผลการเรียน - ${stu.name || ''}</title>
    <style>@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;600;700&display=swap');*{font-family:'Sarabun',sans-serif;margin:0;padding:0;box-sizing:border-box}body{padding:20px 40px;font-size:12px;color:#333}@media print{body{padding:15px 30px}@page{size:A4;margin:15mm 20mm}}table{width:100%;border-collapse:collapse}</style></head><body>
    <div style="text-align:center;margin-bottom:12px">${logoBase64 ? `<img src="${logoBase64}" style="width:55px;height:auto;margin-bottom:4px">` : ''}<div style="font-weight:700;font-size:14px">${APP.config.college_name}</div><div style="font-size:12px;color:#555">ใบรายงานผลการเรียนนักศึกษารายภาคการศึกษา</div><div style="font-size:11px;color:#555">${studentProgram} ระดับ ${studentLevel}</div></div>
    <div style="display:flex;justify-content:space-between;margin-bottom:10px;font-size:11px"><div>รหัสนักศึกษา: <strong>${stu.student_id || ''}</strong></div><div>ชื่อ-สกุล: <strong>${stu.name || ''}</strong></div></div>
    <table><thead><tr style="background:#e8f4fd"><th style="padding:5px 8px;text-align:left;font-size:11px;border:1px solid #999;width:22%">รหัสวิชา</th><th style="padding:5px 8px;text-align:left;font-size:11px;border:1px solid #999;width:48%">รายวิชา</th><th style="padding:5px 8px;text-align:center;font-size:11px;border:1px solid #999;width:15%">หน่วยกิต</th><th style="padding:5px 8px;text-align:center;font-size:11px;border:1px solid #999;width:15%">ระดับคะแนน</th></tr></thead><tbody>${tableHTML}</tbody></table>
    <div style="margin-top:10px;border:1px solid #999;padding:6px 10px;font-size:11px"><div style="display:flex;justify-content:space-between"><span>รวมหน่วยกิตตลอดปีการศึกษา: <strong>${totalCreditsAll}</strong></span><span>คะแนนเฉลี่ยตลอดปีการศึกษา: <strong>${gpax}</strong></span></div><div style="display:flex;justify-content:space-between;margin-top:3px"><span>รวมหน่วยกิตสะสมตลอดหลักสูตร: <strong>${totalCreditsAll}</strong></span><span>คะแนนเฉลี่ยสะสมตลอดหลักสูตร: <strong>${gpax}</strong></span></div></div>
    <div style="margin-top:10px;font-size:10px;color:#666"><div style="font-weight:600;margin-bottom:3px">หมายเหตุ</div><div>A : ดีเยี่ยม &nbsp; B+ : ดีมาก &nbsp; B : ดี &nbsp; C+ : ค่อนข้างดี &nbsp; C : พอใช้ &nbsp; D+ : อ่อน &nbsp; D : อ่อนมาก &nbsp; F : ตก</div></div>
    <script>window.onload=function(){window.print()}<\/script></body></html>`;
  const w = window.open('', '_blank');
  if (w) { w.document.write(htmlContent); w.document.close(); } else { showToast('กรุณาอนุญาต Popup เพื่อดาวน์โหลด PDF', 'error'); }
}

// ======================== ENG RESULTS ========================
function engResultsPage() {
  const isAdmin = APP.currentRole === 'admin' || APP.currentRole === 'academic';
  const canEdit = isAdmin || APP.currentRole === 'teacher' || APP.currentRole === 'classTeacher';
  const isStudent = APP.currentRole === 'student';
  let allEng = getDataByType('eng_result');

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
    const searchVal = APP.filters._engSearch || '';
    let filteredList = studentList;
    if (searchVal) {
      const q = searchVal.toLowerCase();
      filteredList = studentList.filter(s => (s.name || '').toLowerCase().includes(q) || (s.student_id || '').toLowerCase().includes(q));
    }
    studentSelector = `<div class="bg-white rounded-2xl p-4 border border-blue-100 mb-4">
      <label class="block text-sm font-medium text-gray-700 mb-2"><i data-lucide="user-search" class="w-4 h-4 inline mr-1"></i>เลือกนักศึกษา</label>
      <div class="flex gap-2 mb-2">
        <div class="flex-1 relative"><i data-lucide="search" class="absolute left-3 top-2.5 w-4 h-4 text-gray-400"></i><input type="text" placeholder="พิมพ์ค้นหาชื่อหรือรหัส..." value="${searchVal}" class="w-full pl-10 pr-4 py-2 border border-gray-200 rounded-xl text-sm" onkeyup="APP.filters._engSearch=this.value;APP.filters._engStudent='';APP.pagination.page=1;renderCurrentPage()"></div>
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
    data = allEng.filter(e => e.student_id === APP.currentUser.data.student_id);
  } else if (selectedStudentName) {
    data = allEng.filter(e => e.student_id === selectedStudentName);
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

  return `<div class="flex flex-wrap items-center justify-between gap-3 mb-4">
    <h2 class="text-xl font-bold text-gray-800"><i data-lucide="languages" class="w-6 h-6 inline mr-2"></i>ผลสอบภาษาอังกฤษ</h2>
    ${isAdmin ? `<div class="flex gap-2"><button onclick="showAddEngModal()" class="flex items-center gap-2 px-4 py-2 bg-primary text-white rounded-xl hover:bg-primaryDark text-sm"><i data-lucide="plus" class="w-4 h-4"></i>เพิ่มผลสอบ</button>${csvUploadBtn('eng_result', 'student_id,eng_score,eng_type,eng_status')}</div>` : ''}
  </div>
  ${studentSelector}
  ${noSelectionMsg || `<div class="grid grid-cols-1 sm:grid-cols-2 gap-4 mb-4">
    ${statCard('check-circle', 'ผ่าน', data.filter(e => e.eng_status === 'ผ่าน').length, 'คน', 'bg-green-500')}
    ${statCard('x-circle', 'ไม่ผ่าน', data.filter(e => e.eng_status === 'ไม่ผ่าน').length, 'คน', 'bg-red-500')}
  </div>
  <div class="bg-white rounded-2xl border border-blue-100 overflow-hidden">
    <div class="overflow-x-auto"><table class="w-full text-sm">
      <thead><tr class="bg-surface text-left"><th class="px-4 py-3 font-semibold">คะแนน</th><th class="px-4 py-3 font-semibold">รูปแบบ</th><th class="px-4 py-3 font-semibold">สถานะ</th>${isAdmin ? '<th class="px-4 py-3"></th>' : ''}</tr></thead>
      <tbody>${paged.length ? paged.map(e => `<tr class="border-t hover:bg-gray-50">
        <td class="px-4 py-3">${e.eng_score || ''}</td><td class="px-4 py-3">${e.eng_type || ''}</td>
        <td class="px-4 py-3"><span class="px-2 py-1 rounded-full text-xs ${e.eng_status === 'ผ่าน' ? 'bg-green-100 text-green-700' : 'bg-red-100 text-red-700'}">${e.eng_status || ''}</span></td>
        ${isAdmin ? `<td class="px-4 py-3"><div class="flex gap-1"><button onclick="showEditEngModal('${e.__backendId}')" class="text-blue-400 hover:text-blue-600" title="แก้ไข"><i data-lucide="pencil" class="w-4 h-4"></i></button><button onclick="deleteRecord('${e.__backendId}')" class="text-red-400 hover:text-red-600" title="ลบ"><i data-lucide="trash-2" class="w-4 h-4"></i></button></div></td>` : ''}</tr>`).join('') : '<tr><td colspan="4" class="px-4 py-8 text-center text-gray-400">ไม่มีข้อมูล</td></tr>'}</tbody>
    </table></div>
  </div>
  ${paginationHTML(total, APP.pagination.perPage, APP.pagination.page, 'changePage')}`}`;
}

function showAddEngModal() {
  showModal('เพิ่มผลสอบภาษาอังกฤษ', `
    <form id="addEngForm" class="space-y-3">
      <div><label class="block text-xs text-gray-600 mb-1">นักศึกษา *</label><select name="student_id" required class="w-full border rounded-xl px-3 py-2 text-sm">${studentOptionsHTML()}</select></div>
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-xs text-gray-600 mb-1">คะแนน</label><input name="eng_score" type="number" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">รูปแบบการสอบ</label><input name="eng_type" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="TOEIC, IELTS..."></div>
      </div>
      <div><label class="block text-xs text-gray-600 mb-1">สถานะ</label><select name="eng_status" class="w-full border rounded-xl px-3 py-2 text-sm"><option>ผ่าน</option><option>ไม่ผ่าน</option></select></div>
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึก</button>
    </form>
  `);
  document.getElementById('addEngForm').onsubmit = async (e) => {
    e.preventDefault(); const fd = new FormData(e.target);
    const obj = { type: 'eng_result', created_at: new Date().toISOString() }; fd.forEach((v, k) => obj[k] = k === 'eng_score' ? Number(v) : v);

    const r = await GSheetDB.create(obj);
    if (r.isOk) { showToast('เพิ่มผลสอบสำเร็จ'); closeModal() } else showToast('เกิดข้อผิดพลาด', 'error');
  };
}

// ======================== EVAL TEACHER ========================
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
        <span class="text-xs text-gray-400">${f.semester || ''}/${f.academic_year || ''}</span>
      </div>
      <p class="font-bold text-gray-800">${f.subject_code ? f.subject_code + ' ' : ''}${f.subject_name || ''}</p>
      <p class="text-sm text-gray-500 mt-1">อาจารย์: ${f.teacher_name || ''}</p>
      <p class="text-xs text-gray-400 mt-2">${f.eval_items ? f.eval_items.split(',').length || 0 : 0} หัวข้อประเมิน</p>
    </div>`).join('')}</div>` : '<div class="bg-green-50 rounded-2xl p-6 text-center mb-6"><p class="text-green-600">ไม่มีแบบประเมินที่ต้องทำ</p></div>'}

    <h3 class="font-bold mb-3 text-gray-600">ประวัติที่ประเมินแล้ว (${submittedForms.length})</h3>
    <div class="bg-white rounded-2xl border border-blue-100 overflow-hidden">
      <div class="overflow-x-auto"><table class="w-full text-sm">
        <thead><tr class="bg-surface"><th class="px-4 py-3 text-left">รายวิชา</th><th class="px-4 py-3 text-left">อาจารย์</th><th class="px-4 py-3">คะแนนเฉลี่ย</th><th class="px-4 py-3">ภาค/ปี</th></tr></thead>
        <tbody>${myEvals.length ? myEvals.map(e => `<tr class="border-t"><td class="px-4 py-3">${e.subject_name || ''}</td><td class="px-4 py-3">${e.teacher_name || e.name || ''}</td><td class="px-4 py-3 text-center font-bold text-primary">${e.eval_score || ''}/5</td><td class="px-4 py-3 text-center">${e.semester || ''}/${e.academic_year || ''}</td></tr>`).join('') : '<tr><td colspan="4" class="py-6 text-center text-gray-400">ยังไม่มีประวัติ</td></tr>'}</tbody>
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
          <td class="px-4 py-3">${f.semester || ''}/${f.academic_year || ''}</td>
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
      <tbody>${data.length ? data.map(e => `<tr class="border-t hover:bg-gray-50"><td class="px-4 py-3">${e.subject_name || ''}</td><td class="px-4 py-3">${e.student_name || ''}</td><td class="px-4 py-3 font-bold">${e.eval_score || ''}/5</td><td class="px-4 py-3">${e.semester || ''}/${e.academic_year || ''}</td></tr>`).join('') : '<tr><td colspan="4" class="px-4 py-8 text-center text-gray-400">ยังไม่มีผลประเมิน</td></tr>'}</tbody>
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
    e.preventDefault(); const fd = new FormData(e.target);
    const obj = { type: 'eval_form', created_at: new Date().toISOString() };
    fd.forEach((v, k) => obj[k] = v);
    const r = await GSheetDB.create(obj);
    if (r.isOk) { showToast('สร้างแบบประเมินสำเร็จ'); closeModal() } else showToast('เกิดข้อผิดพลาด', 'error');
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
    <p class="text-sm text-gray-600">อาจารย์: ${form.teacher_name || ''} | ภาค ${form.semester || ''}/${form.academic_year || ''}</p>
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
    const avg = (total / count).toFixed(1);

    // Build score detail string
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

// ======================== TEACHERS ========================
function teachersPage() {
  const isAdmin = APP.currentRole === 'admin';
  const isExecutive = APP.currentRole === 'executive';
  const isAcademic = APP.currentRole === 'academic';
  let data = applyFilters(getDataByType('teacher'));

  // Department filter
  const allDepts = [...new Set(data.map(t => t.department).filter(Boolean))].sort();
  const selectedDept = APP.filters._teacherDept || '';
  if (selectedDept) data = data.filter(t => (t.department || '') === selectedDept);

  const total = data.length; const paged = paginate(data);

  const deptFilter = `<div class="bg-white rounded-2xl p-4 border border-blue-100 mb-4">
    <div class="flex items-center gap-3">
      <label class="text-sm font-medium text-gray-700">สาขาวิชา:</label>
      <select onchange="APP.filters._teacherDept=this.value;APP.pagination.page=1;renderCurrentPage()" class="border border-gray-200 rounded-xl px-3 py-2 text-sm">
        <option value="">-- ทั้งหมด --</option>
        ${allDepts.map(d => `<option value="${d}" ${selectedDept === d ? 'selected' : ''}>${d}</option>`).join('')}
      </select>
      ${selectedDept ? `<span class="text-xs text-gray-500">แสดงสาขา: ${selectedDept}</span>` : ''}
    </div>
  </div>`;

  // Executive & Academic sees basic info only (ชื่อ-สกุล, ตำแหน่ง, สาขาวิชา, โทร, E-mail)
  if (isExecutive || isAcademic) {
    return `<div class="flex flex-wrap items-center justify-between gap-3 mb-4">
      <h2 class="text-xl font-bold text-gray-800"><i data-lucide="briefcase" class="w-6 h-6 inline mr-2"></i>ข้อมูลอาจารย์</h2>
    </div>
    ${deptFilter}
    ${filterBar({ semester: false, year: false })}
    <div class="bg-white rounded-2xl border border-blue-100 overflow-hidden">
      <div class="overflow-x-auto"><table class="w-full text-sm">
        <thead><tr class="bg-surface text-left"><th class="px-4 py-3 font-semibold">ชื่อ-สกุล</th><th class="px-4 py-3 font-semibold">ตำแหน่ง</th><th class="px-4 py-3 font-semibold">สาขาวิชา</th><th class="px-4 py-3 font-semibold">โทร</th><th class="px-4 py-3 font-semibold">E-mail</th></tr></thead>
        <tbody>${paged.length ? paged.map(t => `<tr class="border-t hover:bg-gray-50">
          <td class="px-4 py-3 font-medium">${t.name || ''}</td><td class="px-4 py-3">${t.position || ''}</td>
          <td class="px-4 py-3">${t.department || ''}</td><td class="px-4 py-3">${t.phone || ''}</td><td class="px-4 py-3">${t.email || ''}</td></tr>`).join('') : '<tr><td colspan="5" class="px-4 py-8 text-center text-gray-400">ไม่มีข้อมูล</td></tr>'}</tbody>
      </table></div>
    </div>
    ${paginationHTML(total, APP.pagination.perPage, APP.pagination.page, 'changePage')}`;
  }

  return `<div class="flex flex-wrap items-center justify-between gap-3 mb-4">
    <h2 class="text-xl font-bold text-gray-800"><i data-lucide="briefcase" class="w-6 h-6 inline mr-2"></i>ข้อมูลอาจารย์</h2>
    ${isAdmin ? `<div class="flex gap-2"><button onclick="showAddTeacherModal()" class="flex items-center gap-2 px-4 py-2 bg-primary text-white rounded-xl hover:bg-primaryDark text-sm"><i data-lucide="plus" class="w-4 h-4"></i>เพิ่มอาจารย์</button></div>` : ''}
  </div>
  ${deptFilter}
  ${filterBar({ semester: false, year: false })}
  <div class="bg-white rounded-2xl border border-blue-100 overflow-hidden">
    <div class="overflow-x-auto"><table class="w-full text-sm">
      <thead><tr class="bg-surface text-left"><th class="px-4 py-3 font-semibold">ชื่อ-สกุล</th><th class="px-4 py-3 font-semibold">ตำแหน่ง</th><th class="px-4 py-3 font-semibold">สาขาวิชา</th><th class="px-4 py-3"></th></tr></thead>
      <tbody>${paged.length ? paged.map(t => `<tr class="border-t hover:bg-gray-50">
        <td class="px-4 py-3 font-medium">${t.name || ''}</td><td class="px-4 py-3">${t.position || ''}</td>
        <td class="px-4 py-3">${t.department || ''}</td>
        <td class="px-4 py-3"><div class="flex gap-1"><button onclick="showTeacherDetail('${t.__backendId}')" class="text-gray-400 hover:text-primary" title="ดูข้อมูล"><i data-lucide="eye" class="w-4 h-4"></i></button>${isAdmin ? `<button onclick="showEditTeacherModal('${t.__backendId}')" class="text-blue-400 hover:text-blue-600" title="แก้ไข"><i data-lucide="pencil" class="w-4 h-4"></i></button><button onclick="deleteRecord('${t.__backendId}')" class="text-red-400 hover:text-red-600" title="ลบ"><i data-lucide="trash-2" class="w-4 h-4"></i></button>` : ''}</div></td></tr>`).join('') : '<tr><td colspan="4" class="px-4 py-8 text-center text-gray-400">ไม่มีข้อมูล</td></tr>'}</tbody>
    </table></div>
  </div>
  ${paginationHTML(total, APP.pagination.perPage, APP.pagination.page, 'changePage')}`;
}

function showTeacherDetail(id) {
  const t = APP.allData.find(d => d.__backendId === id); if (!t) return;
  showModal('ข้อมูลอาจารย์', `<div class="grid grid-cols-2 gap-3">
    ${infoRow('ชื่อ-สกุล', t.name)}${infoRow('ตำแหน่ง', t.position)}${infoRow('สาขาวิชา', t.department)}
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
        <div><label class="block text-xs text-gray-600 mb-1">เลขบัญชีธนาคาร</label><input name="bank_account" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      </div>
      <div><label class="block text-xs text-gray-600 mb-1">ที่อยู่</label><textarea name="address" rows="2" class="w-full border rounded-xl px-3 py-2 text-sm"></textarea></div>
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึก</button>
    </form>
  `);
  document.getElementById('addTeacherForm').onsubmit = async (e) => {
    e.preventDefault(); const fd = new FormData(e.target);
    const obj = { type: 'teacher', created_at: new Date().toISOString() }; fd.forEach((v, k) => obj[k] = v);

    const r = await GSheetDB.create(obj);
    if (r.isOk) { showToast('เพิ่มอาจารย์สำเร็จ'); closeModal() } else showToast('เกิดข้อผิดพลาด', 'error');
  };
}

// ======================== PROMOTE YEAR (เลื่อนชั้นปี) ========================
function showPromoteYearModal() {
  const students = getDataByType('student').filter(s => s.status === 'กำลังศึกษา');
  const y1 = students.filter(s => norm(s.year_level) === '1').length;
  const y2 = students.filter(s => norm(s.year_level) === '2').length;
  const y3 = students.filter(s => norm(s.year_level) === '3').length;
  const y4 = students.filter(s => norm(s.year_level) === '4').length;

  showModal('เลื่อนชั้นปีนักศึกษา', `
    <div class="space-y-4">
      <div class="bg-amber-50 border border-amber-200 rounded-xl p-4 text-sm text-amber-800">
        <p class="font-semibold mb-1">⚠️ คำเตือน</p>
        <p>เลือกเลื่อนทีละชั้นปี หรือเลื่อนทั้งหมดพร้อมกัน (เฉพาะสถานภาพ "กำลังศึกษา")</p>
      </div>

      <div class="space-y-2">
        <div class="flex items-center justify-between bg-blue-50 rounded-xl p-3">
          <div><p class="font-medium text-sm">ชั้นปี 1 → ชั้นปี 2</p><p class="text-xs text-gray-500">${y1} คน</p></div>
          <button onclick="executePromoteYear('1')" class="px-4 py-2 bg-primary text-white rounded-xl hover:bg-primaryDark text-sm" ${y1 === 0 ? 'disabled class="px-4 py-2 bg-gray-300 text-gray-500 rounded-xl text-sm cursor-not-allowed"' : ''}>เลื่อน ปี 1→2</button>
        </div>
        <div class="flex items-center justify-between bg-blue-50 rounded-xl p-3">
          <div><p class="font-medium text-sm">ชั้นปี 2 → ชั้นปี 3</p><p class="text-xs text-gray-500">${y2} คน</p></div>
          <button onclick="executePromoteYear('2')" class="px-4 py-2 bg-primary text-white rounded-xl hover:bg-primaryDark text-sm" ${y2 === 0 ? 'disabled class="px-4 py-2 bg-gray-300 text-gray-500 rounded-xl text-sm cursor-not-allowed"' : ''}>เลื่อน ปี 2→3</button>
        </div>
        <div class="flex items-center justify-between bg-blue-50 rounded-xl p-3">
          <div><p class="font-medium text-sm">ชั้นปี 3 → ชั้นปี 4</p><p class="text-xs text-gray-500">${y3} คน</p></div>
          <button onclick="executePromoteYear('3')" class="px-4 py-2 bg-primary text-white rounded-xl hover:bg-primaryDark text-sm" ${y3 === 0 ? 'disabled class="px-4 py-2 bg-gray-300 text-gray-500 rounded-xl text-sm cursor-not-allowed"' : ''}>เลื่อน ปี 3→4</button>
        </div>
        <div class="flex items-center justify-between bg-amber-50 rounded-xl p-3">
          <div><p class="font-medium text-sm">ชั้นปี 4 → สำเร็จการศึกษา</p><p class="text-xs text-gray-500">${y4} คน</p></div>
          <button onclick="executePromoteYear('4')" class="px-4 py-2 bg-amber-500 text-white rounded-xl hover:bg-amber-600 text-sm" ${y4 === 0 ? 'disabled class="px-4 py-2 bg-gray-300 text-gray-500 rounded-xl text-sm cursor-not-allowed"' : ''}>จบการศึกษา</button>
        </div>
      </div>

      <div class="border-t pt-3">
        <button onclick="executePromoteYear('all')" class="w-full bg-emerald-500 text-white py-2.5 rounded-xl hover:bg-emerald-600 font-semibold">เลื่อนทั้งหมดพร้อมกัน (ปี 1→2→3→4→จบ)</button>
      </div>
      <button onclick="closeModal()" class="w-full bg-gray-200 text-gray-700 py-2.5 rounded-xl hover:bg-gray-300">ยกเลิก</button>
    </div>
  `);
}

async function executePromoteYear(targetYear) {
  const students = getDataByType('student').filter(s => s.status === 'กำลังศึกษา');

  let batch = [];
  if (targetYear === 'all') {
    batch = students;
  } else {
    batch = students.filter(s => norm(s.year_level) === targetYear);
  }

  if (!batch.length) { showToast('ไม่มีนักศึกษาที่ต้องเลื่อนชั้นปี', 'error'); return; }

  const label = targetYear === 'all' ? 'ทั้งหมด' : targetYear === '4' ? 'ชั้นปี 4 → สำเร็จการศึกษา' : `ชั้นปี ${targetYear} → ${parseInt(targetYear) + 1}`;
  if (!confirm(`ยืนยันเลื่อนชั้นปี: ${label} (${batch.length} คน)?`)) return;

  let successCount = 0; let errorCount = 0;
  showToast('กำลังเลื่อนชั้นปี... กรุณารอสักครู่');

  if (targetYear === 'all') {
    // Process from year 4 down to year 1
    for (const yr of ['4', '3', '2', '1']) {
      const yrBatch = batch.filter(s => norm(s.year_level) === yr);
      for (const s of yrBatch) {
        if (yr === '4') { s.status = 'สำเร็จการศึกษา'; } else { s.year_level = String(parseInt(yr) + 1); }
        const r = await GSheetDB.update(s);
        if (r.isOk) successCount++; else errorCount++;
      }
    }
  } else {
    for (const s of batch) {
      if (targetYear === '4') { s.status = 'สำเร็จการศึกษา'; } else { s.year_level = String(parseInt(targetYear) + 1); }
      const r = await GSheetDB.update(s);
      if (r.isOk) successCount++; else errorCount++;
    }
  }

  closeModal();
  if (errorCount === 0) {
    showToast(`เลื่อนชั้นปีสำเร็จ ${successCount} คน`);
  } else {
    showToast(`เลื่อนชั้นปีสำเร็จ ${successCount} คน / ผิดพลาด ${errorCount} คน`, 'error');
  }
  renderCurrentPage();
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

// Export to Excel
function exportTeacherDirectoryExcel() {
  let data = getDataByType('teacher_directory');
  const activeTab = APP._directoryTab || 'all';
  const selectedYear = APP.filters._directoryYear || '';
  if (selectedYear) data = data.filter(d => (d.academic_year || '') === selectedYear);
  if (activeTab === 'curriculum') data = data.filter(d => (d.teacher_category || '') === 'อาจารย์ประจำหลักสูตร');
  else if (activeTab === 'regular') data = data.filter(d => (d.teacher_category || '') === 'อาจารย์ประจำ');

  if (!data.length) { showToast('ไม่มีข้อมูลสำหรับส่งออก', 'error'); return; }

  // Build CSV with BOM for Excel Thai support
  const headers = ['ชื่อ-สกุล', 'เลขบัตรประชาชน', 'เลขใบประกอบวิชาชีพ', 'ตำแหน่งทางวิชาการ', 'วุฒิการศึกษา', 'ประสบการณ์สอน (ปี)', 'ประสบการณ์สอนทางการพยาบาล', 'ประสบการณ์ปฏิบัติการ (ปี)', 'ประสบการณ์ปฏิบัติการพยาบาล', 'ผลงานวิชาการ (ย้อนหลัง 5 ปี)', 'ปีการศึกษา', 'ประเภทอาจารย์'];
  const fields = ['name', 'national_id', 'license_no', 'academic_position', 'education', 'nursing_teaching_years', 'nursing_teaching_exp', 'nursing_practice_years', 'nursing_practice_exp', 'academic_work', 'academic_year', 'teacher_category'];

  function csvEscape(val) {
    const s = (val || '').replace(/\|\|/g, ', ');
    if (s.includes(',') || s.includes('"') || s.includes('\n')) return '"' + s.replace(/"/g, '""') + '"';
    return s;
  }

  let csv = '\uFEFF'; // BOM
  csv += headers.map(csvEscape).join(',') + '\n';
  data.forEach(row => {
    csv += fields.map(f => {
      if (f === 'national_id') return csvEscape(maskNationalId(row[f]));
      return csvEscape(row[f]);
    }).join(',') + '\n';
  });

  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  const tabLabel = activeTab === 'all' ? 'ทั้งหมด' : activeTab === 'curriculum' ? 'ประจำหลักสูตร' : 'ประจำ';
  a.href = url;
  a.download = `ทำเนียบอาจารย์_${tabLabel}_${new Date().toISOString().slice(0, 10)}.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  showToast('ส่งออกไฟล์สำเร็จ');
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
      <button onclick="exportTeacherDirectoryExcel()" class="flex items-center gap-2 px-4 py-2 bg-emerald-500 text-white rounded-xl hover:bg-emerald-600 text-sm"><i data-lucide="download" class="w-4 h-4"></i>ส่งออก Excel</button>
      ${isAdmin ? `<button onclick="showAddTeacherDirectoryModal()" class="flex items-center gap-2 px-4 py-2 bg-primary text-white rounded-xl hover:bg-primaryDark text-sm"><i data-lucide="plus" class="w-4 h-4"></i>เพิ่มอาจารย์</button>${csvUploadBtn('teacher_directory', 'name,national_id,license_no,academic_position,education,nursing_teaching_years,nursing_teaching_exp,nursing_practice_years,nursing_practice_exp,academic_work,academic_year,teacher_category')}` : ''}
    </div>
  </div>

  <div class="bg-white rounded-2xl p-4 border border-blue-100 mb-4">
    <div class="flex items-center gap-3">
      <label class="text-sm font-medium text-gray-700">ปีการศึกษา:</label>
      <select onchange="APP.filters._directoryYear=this.value;APP.pagination.page=1;renderCurrentPage()" class="border border-gray-200 rounded-xl px-3 py-2 text-sm">
        <option value="">ทั้งหมด</option>
        ${allYears.map(y => `<option value="${y}" ${selectedYear === y ? 'selected' : ''}>${y}</option>`).join('')}
      </select>
      ${selectedYear ? `<span class="text-xs text-gray-500">แสดงข้อมูลปีการศึกษา ${selectedYear}</span>` : ''}
    </div>
  </div>

  ${selectedYear ? renderDirectoryDataSection(paged, total, catAll, catCurriculum, catRegular, activeTab, isAdmin) : '<div class="bg-blue-50 rounded-2xl p-8 text-center text-blue-600 mt-4"><i data-lucide="info" class="w-6 h-6 inline mr-2"></i>กรุณาเลือกปีการศึกษาเพื่อดูข้อมูลทำเนียบอาจารย์</div>'}`;
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
    e.preventDefault(); const form = e.target;
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
    e.preventDefault(); const form = e.target;
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
    e.preventDefault(); const fd = new FormData(e.target);
    const obj = { type: 'announcement', created_at: new Date().toISOString() }; fd.forEach((v, k) => obj[k] = v);

    const r = await GSheetDB.create(obj);
    if (r.isOk) { showToast('เพิ่มประกาศสำเร็จ'); closeModal() } else showToast('เกิดข้อผิดพลาด', 'error');
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

  // Apply general filters to table data
  data = applyFilters(data);
  const total = data.length; const paged = paginate(data);

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
      <thead><tr class="bg-surface text-left"><th class="px-4 py-3 font-semibold">รายวิชา</th><th class="px-4 py-3 font-semibold">ทฤษฎี/ปฏิบัติ</th><th class="px-4 py-3 font-semibold">ชั้นปี</th><th class="px-4 py-3 font-semibold">ภาค</th><th class="px-4 py-3 font-semibold">อ.ประจำชั้นตรวจ</th><th class="px-4 py-3 font-semibold">วิชาการเสนอ</th><th class="px-4 py-3 font-semibold">รอง ผอ.ลงนาม</th><th class="px-4 py-3 font-semibold">วันอนุมัติ</th><th class="px-4 py-3 font-semibold">หมายเหตุ</th>${isAdmin ? '<th class="px-4 py-3"></th>' : ''}</tr></thead>
      <tbody>${paged.length ? paged.map(t => {
    const statusBadge = (s) => s === 'เสร็จสิ้น' ? '<span class="px-2 py-1 rounded-full text-xs bg-green-100 text-green-700">✓ เสร็จสิ้น</span>' : s === 'กำลังดำเนินการ' ? '<span class="px-2 py-1 rounded-full text-xs bg-yellow-100 text-yellow-700">⏳ ดำเนินการ</span>' : '<span class="px-2 py-1 rounded-full text-xs bg-gray-100 text-gray-500">รอ</span>';
    const isLate = t.is_late === 'ใช่' || t.is_late === 'late';
    return `<tr class="border-t hover:bg-gray-50 ${isLate ? 'bg-red-50' : ''}">
        <td class="px-4 py-3 font-medium">${t.subject_name || ''} ${isLate ? '<span class="text-xs text-red-500 font-normal">(ส่งช้า)</span>' : ''}</td><td class="px-4 py-3">${t.theory_practice || ''}</td>
        <td class="px-4 py-3">${t.year_level || ''}</td><td class="px-4 py-3">${t.semester || ''}/${t.academic_year || ''}</td>
        <td class="px-4 py-3">${canApprove ? `<select onchange="updateTrackingField('${t.__backendId}','class_teacher_check',this.value)" class="text-xs border rounded px-1 py-0.5"><option ${t.class_teacher_check === 'รอ' ? 'selected' : ''}>รอ</option><option ${t.class_teacher_check === 'กำลังดำเนินการ' ? 'selected' : ''}>กำลังดำเนินการ</option><option ${t.class_teacher_check === 'เสร็จสิ้น' ? 'selected' : ''}>เสร็จสิ้น</option></select>` : statusBadge(t.class_teacher_check)}</td>
        <td class="px-4 py-3">${canApprove ? `<select onchange="updateTrackingField('${t.__backendId}','academic_propose',this.value)" class="text-xs border rounded px-1 py-0.5"><option ${t.academic_propose === 'รอ' ? 'selected' : ''}>รอ</option><option ${t.academic_propose === 'กำลังดำเนินการ' ? 'selected' : ''}>กำลังดำเนินการ</option><option ${t.academic_propose === 'เสร็จสิ้น' ? 'selected' : ''}>เสร็จสิ้น</option></select>` : statusBadge(t.academic_propose)}</td>
        <td class="px-4 py-3">${canApprove ? `<select onchange="updateTrackingField('${t.__backendId}','deputy_sign',this.value)" class="text-xs border rounded px-1 py-0.5"><option ${t.deputy_sign === 'รอ' ? 'selected' : ''}>รอ</option><option ${t.deputy_sign === 'กำลังดำเนินการ' ? 'selected' : ''}>กำลังดำเนินการ</option><option ${t.deputy_sign === 'เสร็จสิ้น' ? 'selected' : ''}>เสร็จสิ้น</option></select>` : statusBadge(t.deputy_sign)}</td>
        <td class="px-4 py-3">${t.approved_date || '-'}</td>
        <td class="px-4 py-3 text-xs text-gray-500">${t.remarks || ''}</td>
        ${isAdmin ? `<td class="px-4 py-3"><div class="flex gap-1"><button onclick="showEditTrackingModal('${t.__backendId}')" class="text-blue-400 hover:text-blue-600" title="แก้ไข"><i data-lucide="pencil" class="w-4 h-4"></i></button><button onclick="deleteRecord('${t.__backendId}')" class="text-red-400 hover:text-red-600" title="ลบ"><i data-lucide="trash-2" class="w-4 h-4"></i></button></div></td>` : ''}</tr>`
  }).join('') : '<tr><td colspan="10" class="px-4 py-8 text-center text-gray-400">ไม่มีข้อมูล</td></tr>'}</tbody>
    </table></div>
  </div>
  ${paginationHTML(total, APP.pagination.perPage, APP.pagination.page, 'changePage')}` : noYearSelectedMsg('ติดตามการส่งรายละเอียดรายวิชา')}`;
}

function showAddTrackingModal() {
  showModal('เพิ่มรายละเอียดรายวิชา', `
    <form id="addTrackingForm" class="space-y-3">
      <div><label class="block text-xs text-gray-600 mb-1">ชื่อรายวิชา</label><input name="subject_name" required class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-xs text-gray-600 mb-1">ทฤษฎี/ปฏิบัติ</label><select name="theory_practice" class="w-full border rounded-xl px-3 py-2 text-sm"><option>ทฤษฎี</option><option>ปฏิบัติ</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">ชั้นปี</label><select name="year_level" class="w-full border rounded-xl px-3 py-2 text-sm"><option>1</option><option>2</option><option>3</option><option>4</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">ห้อง</label><input name="room" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">ภาคการศึกษา</label><select name="semester" class="w-full border rounded-xl px-3 py-2 text-sm"><option value="1">1</option><option value="2">2</option></select></div>
      </div>
      <div><label class="block text-xs text-gray-600 mb-1">ผู้ประสานงาน</label><input name="coordinator" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึก</button>
    </form>
  `);
  document.getElementById('addTrackingForm').onsubmit = async (e) => {
    e.preventDefault(); const fd = new FormData(e.target);
    const obj = { type: 'tracking', class_teacher_check: 'รอ', academic_propose: 'รอ', deputy_sign: 'รอ', approved_date: '', created_at: new Date().toISOString() };
    fd.forEach((v, k) => obj[k] = v);

    const r = await GSheetDB.create(obj);
    if (r.isOk) { showToast('เพิ่มสำเร็จ'); closeModal() } else showToast('เกิดข้อผิดพลาด', 'error');
  };
}

async function updateTrackingField(id, field, value) {
  const rec = APP.allData.find(d => d.__backendId === id); if (!rec) return;
  rec[field] = value;
  if (field === 'deputy_sign' && value === 'เสร็จสิ้น') rec.approved_date = new Date().toISOString().split('T')[0];

  const r = await GSheetDB.update(rec);
  if (r.isOk) showToast('อัปเดตสำเร็จ'); else showToast('เกิดข้อผิดพลาด', 'error');
}

// ======================== GRADE TRACKING ========================
function gradeTrackingPage() {
  const isAdmin = APP.currentRole === 'admin' || APP.currentRole === 'academic';
  const canEdit = isAdmin || APP.currentRole === 'teacher' || APP.currentRole === 'classTeacher';
  let data = getDataByType('grade_tracking').filter(t => t.subject_name && t.subject_name.trim());
  if (APP.currentRole === 'teacher') data = data.filter(t => t.coordinator === APP.currentUser.name);

  // Year filter for stats
  const selectedYear = APP.filters._gradeTrackingYear || '';
  const allSubjects = getDataByType('subject');
  const subjectsFiltered = selectedYear ? allSubjects.filter(s => norm(s.academic_year) === selectedYear) : allSubjects;
  const dataForStats = selectedYear ? data.filter(t => norm(t.academic_year) === selectedYear) : data;

  // Only show not-submitted and stats when year is selected
  let statsSection = '';
  let notSubmittedSection = '';
  if (selectedYear) {
    const trackedKeys = dataForStats.map(t => `${norm(t.subject_name)}|${norm(t.semester)}|${norm(t.academic_year)}`);
    const notSubmitted = subjectsFiltered.filter(s => !trackedKeys.includes(`${norm(s.subject_name)}|${norm(s.semester)}|${norm(s.academic_year)}`));

    const completed = dataForStats.filter(t => t.deputy_sign === 'เสร็จสิ้น').length;
    const inProgress = dataForStats.filter(t => t.coordinator_check === 'เสร็จสิ้น' && t.deputy_sign !== 'เสร็จสิ้น').length;
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

  // Year options
  const allYears = [...new Set([...allSubjects.map(s => s.academic_year), ...data.map(t => t.academic_year)].filter(Boolean))].sort();

  // Apply general filters to table data
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
      <thead><tr class="bg-surface text-left">
        <th class="px-4 py-3 font-semibold">รายวิชา</th>
        <th class="px-4 py-3 font-semibold">ทฤษฎี/ปฏิบัติ</th>
        <th class="px-4 py-3 font-semibold">ชั้นปี</th>
        <th class="px-4 py-3 font-semibold">ภาค/ปี</th>
        <th class="px-4 py-3 font-semibold">ผู้ประสานงาน</th>
        <th class="px-4 py-3 font-semibold">อ.ประสานงานส่ง</th>
        <th class="px-4 py-3 font-semibold">วิชาการตรวจ</th>
        <th class="px-4 py-3 font-semibold">รอง ผอ.ลงนาม</th>
        <th class="px-4 py-3 font-semibold">วันอนุมัติ</th>
        <th class="px-4 py-3 font-semibold">หมายเหตุ</th>
      </tr></thead>
      <tbody>${paged.length ? paged.map(t => {
    const statusBadge = (s) => s === 'เสร็จสิ้น' ? '<span class="px-2 py-1 rounded-full text-xs bg-green-100 text-green-700">✓ เสร็จสิ้น</span>' : s === 'กำลังดำเนินการ' ? '<span class="px-2 py-1 rounded-full text-xs bg-yellow-100 text-yellow-700">⏳ ดำเนินการ</span>' : '<span class="px-2 py-1 rounded-full text-xs bg-gray-100 text-gray-500">รอ</span>';
    const isLate = t.is_late === 'ใช่' || t.is_late === 'late';
    return `<tr class="border-t hover:bg-gray-50 ${isLate ? 'bg-red-50' : ''}">
        <td class="px-4 py-3 font-medium">${t.subject_name || ''} ${isLate ? '<span class="text-xs text-red-500 font-normal">(ส่งช้า)</span>' : ''}</td>
        <td class="px-4 py-3">${t.theory_practice || ''}</td>
        <td class="px-4 py-3">${t.year_level || ''}</td>
        <td class="px-4 py-3">${t.semester || ''}/${t.academic_year || ''}</td>
        <td class="px-4 py-3">${t.coordinator || ''}</td>
        <td class="px-4 py-3">${statusBadge(t.coordinator_check)}</td>
        <td class="px-4 py-3">${statusBadge(t.academic_check)}</td>
        <td class="px-4 py-3">${statusBadge(t.deputy_sign)}</td>
        <td class="px-4 py-3">${t.approved_date || '-'}</td>
        <td class="px-4 py-3 text-xs text-gray-500">${t.remarks || ''}</td>
      </tr>`}).join('') : '<tr><td colspan="10" class="px-4 py-8 text-center text-gray-400">ไม่มีข้อมูล</td></tr>'}</tbody>
    </table></div>
  </div>
  ${paginationHTML(total, APP.pagination.perPage, APP.pagination.page, 'changePage')}` : noYearSelectedMsg('ติดตามการส่งเกรดรายวิชา')}`;
}

function showAddGradeTrackingModal() {
  showModal('เพิ่มข้อมูลติดตามการส่งเกรดรายวิชา', `
    <form id="addGradeTrackingForm" class="space-y-3">
      <div><label class="block text-xs text-gray-600 mb-1">ชื่อรายวิชา *</label><input name="subject_name" required class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-xs text-gray-600 mb-1">ทฤษฎี/ปฏิบัติ</label><select name="theory_practice" class="w-full border rounded-xl px-3 py-2 text-sm"><option>ทฤษฎี</option><option>ปฏิบัติ</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">ชั้นปี</label><select name="year_level" class="w-full border rounded-xl px-3 py-2 text-sm"><option>1</option><option>2</option><option>3</option><option>4</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">ภาคการศึกษา</label><select name="semester" class="w-full border rounded-xl px-3 py-2 text-sm"><option value="1">1</option><option value="2">2</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">ปีการศึกษา</label><input name="academic_year" value="2568" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      </div>
      <div><label class="block text-xs text-gray-600 mb-1">ผู้ประสานงาน</label><input name="coordinator" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <div><label class="block text-xs text-gray-600 mb-1">หมายเหตุ</label><input name="remarks" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึก</button>
    </form>
  `);
  document.getElementById('addGradeTrackingForm').onsubmit = async (e) => {
    e.preventDefault(); const fd = new FormData(e.target);
    const obj = { type: 'grade_tracking', coordinator_check: 'รอ', academic_check: 'รอ', deputy_sign: 'รอ', approved_date: '', created_at: new Date().toISOString() };
    fd.forEach((v, k) => obj[k] = v);
    const r = await GSheetDB.create(obj);
    if (r.isOk) { showToast('เพิ่มสำเร็จ'); closeModal() } else showToast('เกิดข้อผิดพลาด', 'error');
  };
}

// ======================== FILE TRACKING ========================
function fileTrackingPage() {
  const isAdmin = APP.currentRole === 'admin' || APP.currentRole === 'academic';
  const canEdit = isAdmin || APP.currentRole === 'teacher' || APP.currentRole === 'classTeacher';
  let data = getDataByType('file_tracking').filter(t => t.subject_name && t.subject_name.trim());
  if (APP.currentRole === 'teacher') data = data.filter(t => t.coordinator === APP.currentUser.name);

  // Year filter for stats
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
    const inProgress = dataForStats.filter(t => t.coordinator_check === 'เสร็จสิ้น' && t.deputy_sign !== 'เสร็จสิ้น').length;
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
      <thead><tr class="bg-surface text-left">
        <th class="px-4 py-3 font-semibold">รายวิชา</th>
        <th class="px-4 py-3 font-semibold">ทฤษฎี/ปฏิบัติ</th>
        <th class="px-4 py-3 font-semibold">ชั้นปี</th>
        <th class="px-4 py-3 font-semibold">ภาค/ปี</th>
        <th class="px-4 py-3 font-semibold">ผู้ประสานงาน</th>
        <th class="px-4 py-3 font-semibold">อ.ประสานงานส่ง</th>
        <th class="px-4 py-3 font-semibold">วิชาการตรวจ</th>
        <th class="px-4 py-3 font-semibold">รอง ผอ.ลงนาม</th>
        <th class="px-4 py-3 font-semibold">วันอนุมัติ</th>
        <th class="px-4 py-3 font-semibold">หมายเหตุ</th>
      </tr></thead>
      <tbody>${paged.length ? paged.map(t => {
    const statusBadge = (s) => s === 'เสร็จสิ้น' ? '<span class="px-2 py-1 rounded-full text-xs bg-green-100 text-green-700">✓ เสร็จสิ้น</span>' : s === 'กำลังดำเนินการ' ? '<span class="px-2 py-1 rounded-full text-xs bg-yellow-100 text-yellow-700">⏳ ดำเนินการ</span>' : '<span class="px-2 py-1 rounded-full text-xs bg-gray-100 text-gray-500">รอ</span>';
    const isLate = t.is_late === 'ใช่' || t.is_late === 'late';
    return `<tr class="border-t hover:bg-gray-50 ${isLate ? 'bg-red-50' : ''}">
        <td class="px-4 py-3 font-medium">${t.subject_name || ''} ${isLate ? '<span class="text-xs text-red-500 font-normal">(ส่งช้า)</span>' : ''}</td>
        <td class="px-4 py-3">${t.theory_practice || ''}</td>
        <td class="px-4 py-3">${t.year_level || ''}</td>
        <td class="px-4 py-3">${t.semester || ''}/${t.academic_year || ''}</td>
        <td class="px-4 py-3">${t.coordinator || ''}</td>
        <td class="px-4 py-3">${statusBadge(t.coordinator_check)}</td>
        <td class="px-4 py-3">${statusBadge(t.academic_check)}</td>
        <td class="px-4 py-3">${statusBadge(t.deputy_sign)}</td>
        <td class="px-4 py-3">${t.approved_date || '-'}</td>
        <td class="px-4 py-3 text-xs text-gray-500">${t.remarks || ''}</td>
      </tr>`}).join('') : '<tr><td colspan="10" class="px-4 py-8 text-center text-gray-400">ไม่มีข้อมูล</td></tr>'}</tbody>
    </table></div>
  </div>
  ${paginationHTML(total, APP.pagination.perPage, APP.pagination.page, 'changePage')}` : noYearSelectedMsg('ติดตามการส่งแฟ้มรายวิชา')}`;
}

function showAddFileTrackingModal() {
  showModal('เพิ่มข้อมูลติดตามการส่งแฟ้มรายวิชา', `
    <form id="addFileTrackingForm" class="space-y-3">
      <div><label class="block text-xs text-gray-600 mb-1">ชื่อรายวิชา *</label><input name="subject_name" required class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-xs text-gray-600 mb-1">ทฤษฎี/ปฏิบัติ</label><select name="theory_practice" class="w-full border rounded-xl px-3 py-2 text-sm"><option>ทฤษฎี</option><option>ปฏิบัติ</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">ชั้นปี</label><select name="year_level" class="w-full border rounded-xl px-3 py-2 text-sm"><option>1</option><option>2</option><option>3</option><option>4</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">ภาคการศึกษา</label><select name="semester" class="w-full border rounded-xl px-3 py-2 text-sm"><option value="1">1</option><option value="2">2</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">ปีการศึกษา</label><input name="academic_year" value="2568" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      </div>
      <div><label class="block text-xs text-gray-600 mb-1">ผู้ประสานงาน</label><input name="coordinator" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <div><label class="block text-xs text-gray-600 mb-1">หมายเหตุ</label><input name="remarks" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึก</button>
    </form>
  `);
  document.getElementById('addFileTrackingForm').onsubmit = async (e) => {
    e.preventDefault(); const fd = new FormData(e.target);
    const obj = { type: 'file_tracking', coordinator_check: 'รอ', academic_check: 'รอ', deputy_sign: 'รอ', approved_date: '', created_at: new Date().toISOString() };
    fd.forEach((v, k) => obj[k] = v);
    const r = await GSheetDB.create(obj);
    if (r.isOk) { showToast('เพิ่มสำเร็จ'); closeModal() } else showToast('เกิดข้อผิดพลาด', 'error');
  };
}

// ======================== LEAVE ========================
function leavePage() {
  const isStudent = APP.currentRole === 'student';
  const isAdmin = APP.currentRole === 'admin' || APP.currentRole === 'academic';
  const isTeacher = APP.currentRole === 'teacher';
  const isClassTeacher = APP.currentRole === 'classTeacher';
  const canEdit = isAdmin || isTeacher || isClassTeacher;

  let data = getDataByType('leave');

  if (isStudent && APP.currentUser.data) data = data.filter(l => l.name === APP.currentUser.data.name);
  if (isTeacher) data = data.filter(l => {
    const sub = getDataByType('subject').find(s => s.subject_name === l.subject_name && s.coordinator && s.coordinator.includes(APP.currentUser.name));
    return !!sub;
  });
  if (isClassTeacher) {
    const yr = APP.currentUser.responsible_year || '1';
    const stuNames = getDataByType('student').filter(s => norm(s.year_level) === norm(yr)).map(s => s.name);
    data = data.filter(l => stuNames.includes(l.name));
  }
  data = applyFilters(data);
  const total = data.length; const paged = paginate(data);

  let form = '';
  if (isStudent) {
    const subjects = getDataByType('subject');
    form = `<div class="bg-white rounded-2xl p-5 border border-blue-100 mb-4">
      <h3 class="font-bold mb-3">กรอกข้อมูลการลา</h3>
      <form id="leaveForm" class="space-y-3">
        <div class="grid grid-cols-1 md:grid-cols-2 gap-3">
          <div><label class="block text-xs text-gray-600 mb-1">ชื่อ-สกุล</label><input name="name" value="${APP.currentUser.data?.name || ''}" readonly class="w-full border rounded-xl px-3 py-2 text-sm bg-gray-50"></div>
          <div><label class="block text-xs text-gray-600 mb-1">รายวิชา</label><select name="subject_name" required class="w-full border rounded-xl px-3 py-2 text-sm" onchange="updateLeaveCoordinator(this.value)"><option value="">เลือกรายวิชา</option>${subjects.map(s => `<option value="${s.subject_name}">${s.subject_code ? s.subject_code + ' ' : ''}${s.subject_name}</option>`).join('')}</select></div>
          <div><label class="block text-xs text-gray-600 mb-1">อาจารย์ผู้ประสานรายวิชา</label><input name="coordinator" id="leaveCoordinator" readonly class="w-full border rounded-xl px-3 py-2 text-sm bg-gray-50"></div>
          <div><label class="block text-xs text-gray-600 mb-1">จำนวนชั่วโมง</label><input name="leave_hours" type="number" required class="w-full border rounded-xl px-3 py-2 text-sm"></div>
          <div><label class="block text-xs text-gray-600 mb-1">ภาคการศึกษา</label><select name="semester" class="w-full border rounded-xl px-3 py-2 text-sm"><option value="1">1</option><option value="2">2</option></select></div>
          <div><label class="block text-xs text-gray-600 mb-1">ปีการศึกษา</label><input name="academic_year" value="2568" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
          <div><label class="block text-xs text-gray-600 mb-1">วันที่ลา</label><input name="leave_date" type="date" required class="w-full border rounded-xl px-3 py-2 text-sm" onchange="validateLeaveDate()"></div>
          <div><label class="block text-xs text-gray-600 mb-1">ประเภทการลา</label><select name="leave_type" required class="w-full border rounded-xl px-3 py-2 text-sm" onchange="onLeaveTypeChange(this.value)"><option value="">เลือก</option><option value="ลาป่วย">ลาป่วย</option><option value="ลากิจ">ลากิจ</option><option value="ลาพบแพทย์">ลาพบแพทย์</option></select></div>
        </div>
        <div id="leaveExtra" class="space-y-3"></div>
        <div><label class="block text-xs text-gray-600 mb-1">อาจารย์ประจำชั้น</label><input name="class_teacher" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div id="leaveValidation" class="text-red-500 text-xs hidden"></div>
        <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">ส่งใบลา</button>
      </form>
    </div>`;
  }

  return `<h2 class="text-xl font-bold text-gray-800 mb-4"><i data-lucide="calendar-off" class="w-6 h-6 inline mr-2"></i>ระบบการลาของนักศึกษา</h2>
  ${isAdmin ? `<div class="flex gap-2 mb-4"><button onclick="showAddLeaveModal()" class="flex items-center gap-2 px-4 py-2 bg-primary text-white rounded-xl hover:bg-primaryDark text-sm"><i data-lucide="plus" class="w-4 h-4"></i>เพิ่มข้อมูลการลา</button>${csvUploadBtn('leave', 'name,subject_name,leave_hours,leave_percent,semester,academic_year,leave_date,leave_type')}</div>` : ''}
  ${form}
  ${filterBar()}
  <div class="bg-white rounded-2xl border border-blue-100 overflow-hidden">
    <div class="overflow-x-auto"><table class="w-full text-sm">
      <thead><tr class="bg-surface text-left"><th class="px-4 py-3 font-semibold">ชื่อ-สกุล</th><th class="px-4 py-3 font-semibold">รายวิชา</th><th class="px-4 py-3 font-semibold">ประเภท</th><th class="px-4 py-3 font-semibold">ชม.</th><th class="px-4 py-3 font-semibold">%ลา</th><th class="px-4 py-3 font-semibold">วันที่</th><th class="px-4 py-3 font-semibold">ภาค/ปี</th><th class="px-4 py-3 font-semibold">สถานะ</th>${isTeacher || isClassTeacher ? '<th class="px-4 py-3 font-semibold">การอนุมัติ</th>' : ''}${isAdmin ? '<th class="px-4 py-3"></th>' : ''}</tr></thead>
      <tbody>${paged.length ? paged.map(l => {
    const getStatusBadge = (status) => {
      if (status === 'รออนุมัติ') return '<span class="px-2 py-1 rounded-full text-xs bg-yellow-100 text-yellow-700">⏳ รออนุมัติ</span>';
      if (status === 'อนุมัติแล้ว') return '<span class="px-2 py-1 rounded-full text-xs bg-green-100 text-green-700">✓ อนุมัติแล้ว</span>';
      if (status === 'ปฏิเสธ') return '<span class="px-2 py-1 rounded-full text-xs bg-red-100 text-red-700">✕ ปฏิเสธ</span>';
      return '<span class="px-2 py-1 rounded-full text-xs bg-gray-100 text-gray-700">รอส่ง</span>';
    };
    const approvalButtons = `<div class="flex gap-1">
          <button onclick="approveLeave('${l.__backendId}','${isTeacher ? 'coordinator_approval' : isClassTeacher ? 'class_teacher_approval' : 'deputy_approval'}')" class="px-2 py-1 bg-green-500 hover:bg-green-600 text-white rounded-lg text-xs flex items-center gap-1"><i data-lucide="check" class="w-3 h-3"></i>อนุมัติ</button>
          <button onclick="rejectLeave('${l.__backendId}','${isTeacher ? 'coordinator_approval' : isClassTeacher ? 'class_teacher_approval' : 'deputy_approval'}')" class="px-2 py-1 bg-red-500 hover:bg-red-600 text-white rounded-lg text-xs flex items-center gap-1"><i data-lucide="x" class="w-3 h-3"></i>ปฏิเสธ</button>
        </div>`;
    return `<tr class="border-t hover:bg-gray-50">
        <td class="px-4 py-3">${l.name || ''}</td><td class="px-4 py-3">${l.subject_name || ''}</td>
        <td class="px-4 py-3"><span class="px-2 py-1 rounded-full text-xs ${l.leave_type === 'ลาป่วย' ? 'bg-red-100 text-red-700' : l.leave_type === 'ลาพบแพทย์' ? 'bg-purple-100 text-purple-700' : 'bg-orange-100 text-orange-700'}">${l.leave_type || ''}</span></td>
        <td class="px-4 py-3">${l.leave_hours || ''}</td><td class="px-4 py-3">${l.leave_percent || '-'}%</td>
        <td class="px-4 py-3">${l.leave_date || ''}</td><td class="px-4 py-3">${l.semester || ''}/${l.academic_year || ''}</td>
        <td class="px-4 py-3">${getStatusBadge(l.leave_status)}</td>
        ${(isTeacher || isClassTeacher) ? `<td class="px-4 py-3">${approvalButtons}</td>` : ''} 
        ${isAdmin ? `<td class="px-4 py-3"><div class="flex gap-1"><button onclick="showEditLeaveModal('${l.__backendId}')" class="text-blue-400 hover:text-blue-600" title="แก้ไข"><i data-lucide="pencil" class="w-4 h-4"></i></button><button onclick="deleteRecord('${l.__backendId}')" class="text-red-400 hover:text-red-600" title="ลบ"><i data-lucide="trash-2" class="w-4 h-4"></i></button></div></td>` : ''}</tr>`
  }).join('') : '<tr><td colspan="10" class="px-4 py-8 text-center text-gray-400">ไม่มีข้อมูล</td></tr>'}</tbody>
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
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึก</button>
    </form>
  `);
  document.getElementById('addLeaveForm').onsubmit = async (e) => {
    e.preventDefault(); const fd = new FormData(e.target);
    const obj = { type: 'leave', created_at: new Date().toISOString() }; fd.forEach((v, k) => obj[k] = (k === 'leave_hours' || k === 'leave_percent') ? Number(v) : v);
    obj.leave_status = 'รออนุมัติ';
    obj.coordinator_approval = 'รอ';
    obj.class_teacher_approval = 'รอ';
    obj.deputy_approval = 'รอ';

    const r = await GSheetDB.create(obj);
    if (r.isOk) { showToast('เพิ่มข้อมูลการลาสำเร็จ'); closeModal() } else showToast('เกิดข้อผิดพลาด', 'error');
  };
}

// ======================== SETTINGS ========================
function settingsPage() {
  const roles = ['admin', 'academic', 'executive', 'teacher', 'classTeacher', 'student'];
  const modules = ['dashboard', 'students', 'subjects', 'schedule', 'grades', 'engResults', 'evalTeacher', 'teachers', 'teacherDirectory', 'services', 'tracking', 'gradeTracking', 'fileTracking', 'leave'];
  const moduleLabels = { dashboard: 'หน้าหลัก', students: 'ข้อมูลนักศึกษา', subjects: 'รายวิชา', schedule: 'ตารางเรียน/สอบ', grades: 'ผลการเรียน', engResults: 'ผลสอบ ENG', evalTeacher: 'ประเมินอาจารย์', teachers: 'ข้อมูลอาจารย์', teacherDirectory: 'ทำเนียบอาจารย์', services: 'บริการอื่นๆ', tracking: 'ติดตามการส่งรายละเอียดรายวิชา', gradeTracking: 'ติดตามการส่งเกรดรายวิชา', fileTracking: 'ติดตามส่งแฟ้มรายวิชา', leave: 'ระบบการลาของนักศึกษา' };
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
    extra.innerHTML = `<div class="bg-blue-50 p-3 rounded-xl text-xs text-blue-700"><i data-lucide="info" class="w-4 h-4 inline"></i> ลากิจต้องส่งล่วงหน้า 1-2 วัน หากส่งช้ากรุณาใส่เหตุผล</div>
    <div id="lateReasonDiv" class="hidden"><label class="block text-xs text-gray-600 mb-1">เหตุผลที่ส่งล่าช้า *</label><textarea name="leave_reason" class="w-full border rounded-xl px-3 py-2 text-sm" rows="2"></textarea></div>`;
  } else if (type === 'ลาพบแพทย์') {
    extra.innerHTML = `<div class="bg-purple-50 p-3 rounded-xl text-xs text-purple-700"><i data-lucide="info" class="w-4 h-4 inline"></i> ลาพบแพทย์ต้องส่งล่วงหน้า 3 วันทำการ และแนบใบนัดแพทย์</div>
    <div><label class="block text-xs text-gray-600 mb-1">แนบใบนัดแพทย์ * (.jpg, .pdf, .png)</label><input type="file" accept=".jpg,.pdf,.png" class="w-full text-sm" name="appointment_doc"></div>`;
  }
  lucide.createIcons();
}

function validateLeaveDate() {
  const form = document.getElementById('leaveForm'); if (!form) return;
  const typeSelect = form.querySelector('[name="leave_type"]');
  if (!typeSelect) return;
  const type = typeSelect.value;
  const dateInput = form.querySelector('[name="leave_date"]');
  if (!dateInput || !dateInput.value) return;
  const leaveDate = new Date(dateInput.value);
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const diffDays = Math.ceil((leaveDate - today) / (1000 * 60 * 60 * 24));

  if (type === 'ลากิจ') {
    const lateDiv = document.getElementById('lateReasonDiv');
    if (lateDiv) { lateDiv.classList.toggle('hidden', diffDays >= 1) }
  }
}

function updateLeaveCoordinator(subjectName) {
  const sub = getDataByType('subject').find(s => s.subject_name === subjectName);
  const el = document.getElementById('leaveCoordinator');
  if (el) el.value = sub?.coordinator || '';
}

function updateEvalTeacherOptions(subjectName) {
  const sub = getDataByType('subject').find(s => s.subject_name === subjectName);
  const sel = document.getElementById('evalTeacherSelect');
  if (!sel) return;
  sel.innerHTML = '<option value="">เลือกอาจารย์</option>';
  if (sub && sub.coordinator) {
    sub.coordinator.split(',').map(c => c.trim()).filter(Boolean).forEach(c => {
      sel.innerHTML += `<option value="${c}">${c}</option>`;
    });
  }
  // Also add all teachers
  getDataByType('teacher').forEach(t => {
    if (!sel.querySelector(`option[value="${t.name}"]`)) sel.innerHTML += `<option value="${t.name}">${t.name}</option>`;
  });
}

function setEvalScore(n) {
  document.getElementById('evalScoreInput').value = n;
  document.querySelectorAll('.eval-star').forEach((btn, i) => {
    btn.classList.toggle('bg-yellow-400', i < n);
    btn.classList.toggle('text-white', i < n);
    btn.classList.toggle('border-yellow-400', i < n);
  });
}

// ======================== INIT PAGE SCRIPTS ========================
function initPageScripts(page) {
  if (page === 'dashboard') { renderCalendar('dashCalendar') }
  if (page === 'schedule') { renderCalendar('scheduleCalendar') }

  // Student leave form
  const leaveForm = document.getElementById('leaveForm');
  if (leaveForm) {
    leaveForm.onsubmit = async (e) => {
      e.preventDefault();

      const diff = Math.ceil((new Date(leaveDate) - today) / (1000 * 60 * 60 * 24));
      const valEl = document.getElementById('leaveValidation');

      // Validate sick leave: if >3 days after, check cert
      if (type === 'ลาป่วย' && diff <= -3) {
        const cert = fd.get('medical_cert');
        if (!cert || !cert.name) {
          if (valEl) { valEl.textContent = 'ลาป่วย 3 วันขึ้นไป ต้องแนบใบรับรองแพทย์'; valEl.classList.remove('hidden') }
          return;
        }
      }
      if (type === 'ลากิจ' && diff < 1) {
        const reason = fd.get('leave_reason');
        if (!reason) {
          if (valEl) { valEl.textContent = 'กรุณาใส่เหตุผลที่ส่งล่าช้า'; valEl.classList.remove('hidden') }
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
      const obj = { type: 'leave', created_at: new Date().toISOString() };
      ['name', 'subject_name', 'leave_hours', 'semester', 'academic_year', 'leave_date', 'leave_type', 'leave_reason', 'coordinator'].forEach(k => {
        const v = fd.get(k); if (v) obj[k] = k === 'leave_hours' ? Number(v) : v;
      });
      obj.leave_status = 'รออนุมัติ';
      obj.coordinator_approval = 'รอ';
      obj.class_teacher_approval = 'รอ';
      obj.deputy_approval = 'รอ';
      if (APP.allData.filter(d => d.type === 'leave').length >= 999) { showToast('ข้อมูลเต็ม', 'error'); return }

      const r = await GSheetDB.create(obj);
      if (r.isOk) { showToast('ส่งใบลาสำเร็จ'); leaveForm.reset() } else showToast('เกิดข้อผิดพลาด', 'error');
    };
  }

  // Eval form
  const evalForm = document.getElementById('evalForm');
  if (evalForm) {
    evalForm.onsubmit = async (e) => {
      e.preventDefault();

      const obj = { type: 'evaluation', student_name: APP.currentUser.data?.name || '', created_at: new Date().toISOString() };
      fd.forEach((v, k) => { if (k !== 'medical_cert' && k !== 'appointment_doc') obj[k] = k === 'eval_score' ? Number(v) : v });

      const r = await GSheetDB.create(obj);
      if (r.isOk) { showToast('ส่งผลประเมินสำเร็จ'); evalForm.reset() } else showToast('เกิดข้อผิดพลาด', 'error');
    };
  }
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
async function approveLeave(id, approvalField) {
  const rec = APP.allData.find(d => d.__backendId === id); if (!rec) return;
  rec[approvalField] = 'อนุมัติ';
  // If all approvals done, mark as approved
  if (rec.coordinator_approval === 'อนุมัติ' && rec.class_teacher_approval === 'อนุมัติ' && rec.deputy_approval === 'อนุมัติ') {
    rec.leave_status = 'อนุมัติแล้ว';
  }

  const r = await GSheetDB.update(rec);
  if (r.isOk) { showToast('อนุมัติการลาสำเร็จ'); renderCurrentPage() } else showToast('เกิดข้อผิดพลาด', 'error');
}

async function rejectLeave(id, approvalField) {
  const rec = APP.allData.find(d => d.__backendId === id); if (!rec) return;
  rec[approvalField] = 'ปฏิเสธ';
  rec.leave_status = 'ปฏิเสธ';

  const r = await GSheetDB.update(rec);
  if (r.isOk) { showToast('ปฏิเสธการลาสำเร็จ'); renderCurrentPage() } else showToast('เกิดข้อผิดพลาด', 'error');
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
  showModal('แก้ไขผลสอบภาษาอังกฤษ', `
    <form id="editEngForm" class="space-y-3">
      <div><label class="block text-xs text-gray-600 mb-1">นักศึกษา</label><select name="student_id" class="w-full border rounded-xl px-3 py-2 text-sm">${studentOptionsHTML(e.student_id)}</select></div>
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-xs text-gray-600 mb-1">คะแนน</label><input name="eng_score" type="number" value="${e.eng_score || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">รูปแบบ</label><input name="eng_type" value="${e.eng_type || ''}" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="TOEIC, IELTS..."></div>
      </div>
      <div><label class="block text-xs text-gray-600 mb-1">สถานะ</label><select name="eng_status" class="w-full border rounded-xl px-3 py-2 text-sm"><option ${e.eng_status === 'ผ่าน' ? 'selected' : ''}>ผ่าน</option><option ${e.eng_status === 'ไม่ผ่าน' ? 'selected' : ''}>ไม่ผ่าน</option></select></div>
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึกการแก้ไข</button>
    </form>
  `);
  document.getElementById('editEngForm').onsubmit = (ev) => { ev.preventDefault(); editRecord(id, 'editEngForm') };
}

function showEditEvalFormModal(id) {
  const f = APP.allData.find(d => d.__backendId === id); if (!f) return;
  showModal('แก้ไขแบบประเมิน', `
    <form id="editEvalFormForm" class="space-y-3">
      <div><label class="block text-xs text-gray-600 mb-1">รหัสวิชา</label><input name="subject_code" value="${f.subject_code || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <div><label class="block text-xs text-gray-600 mb-1">รายวิชา</label><input name="subject_name" value="${f.subject_name || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <div><label class="block text-xs text-gray-600 mb-1">อาจารย์ผู้สอน</label><input name="teacher_name" value="${f.teacher_name || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-xs text-gray-600 mb-1">ภาคการศึกษา</label><select name="semester" class="w-full border rounded-xl px-3 py-2 text-sm"><option value="1" ${norm(f.semester) === '1' ? 'selected' : ''}>1</option><option value="2" ${norm(f.semester) === '2' ? 'selected' : ''}>2</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">ปีการศึกษา</label><input name="academic_year" value="${f.academic_year || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      </div>
      <div><label class="block text-xs text-gray-600 mb-1">หัวข้อประเมิน (คั่นด้วย ,)</label><textarea name="eval_items" rows="3" class="w-full border rounded-xl px-3 py-2 text-sm">${f.eval_items || ''}</textarea></div>
      <div><label class="block text-xs text-gray-600 mb-1">สถานะ</label><select name="status" class="w-full border rounded-xl px-3 py-2 text-sm"><option value="เปิด" ${f.status === 'เปิด' ? 'selected' : ''}>เปิดรับประเมิน</option><option value="ปิด" ${f.status === 'ปิด' ? 'selected' : ''}>ปิดรับประเมิน</option></select></div>
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึกการแก้ไข</button>
    </form>
  `);
  document.getElementById('editEvalFormForm').onsubmit = (e) => { e.preventDefault(); editRecord(id, 'editEvalFormForm') };
}

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
