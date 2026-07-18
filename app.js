// ======================== STATE ========================
let APP = {
  currentUser: null, currentRole: null, currentPage: 'dashboard', sidebarOpen: false,
  allData: [],
  config: { system_title: 'ระบบบริหารจัดการงานวิชาการ (EMS-BCNB)', college_name: 'วิทยาลัยพยาบาลบรมราชชนนี กรุงเทพ' },
  permissions: { admin: { dashboard: 1, students: 1, subjects: 1, schedule: 1, grades: 1, engResults: 1, teachers: 1, specialTeachers: 1, alumni: 1, teacherDirectory: 1, services: 1, tracking: 1, resultTracking: 1, gradeTracking: 1, fileTracking: 1, leave: 1, settings: 1, loginLog: 1, advisors: 1, surveyManage: 1 }, academic: { dashboard: 1, students: 1, subjects: 1, schedule: 1, grades: 1, engResults: 1, teachers: 1, specialTeachers: 1, alumni: 1, teacherDirectory: 1, services: 1, tracking: 1, resultTracking: 1, gradeTracking: 1, fileTracking: 1, leave: 1, settings: 1, advisors: 1, survey: 1 }, registrar: { dashboard: 1, students: 1, subjects: 1, schedule: 1, grades: 1, engResults: 1, teachers: 1, specialTeachers: 1, alumni: 1, teacherDirectory: 1, services: 1, leave: 1, advisors: 1, survey: 1 }, deptHead: { dashboard: 1, teacherDirectory: 1, tracking: 1, resultTracking: 1, gradeTracking: 1, fileTracking: 1, survey: 1 }, teacher: { dashboard: 1, students: 1, subjects: 1, grades: 1, engResults: 1, tracking: 1, gradeTracking: 1, fileTracking: 1, leave: 1, survey: 1 }, classTeacher: { dashboard: 1, students: 1, subjects: 1, grades: 1, engResults: 1, tracking: 1, gradeTracking: 1, fileTracking: 1, leave: 1, survey: 1 }, student: { dashboard: 1, students: 1, grades: 1, engResults: 1, leave: 1, survey: 1 }, executive: { dashboard: 1, students: 1, subjects: 1, schedule: 1, grades: 1, engResults: 1, teachers: 1, specialTeachers: 1, alumni: 1, teacherDirectory: 1, tracking: 1, resultTracking: 1, gradeTracking: 1, fileTracking: 1, leave: 1, advisors: 1, survey: 1 } },
  filters: { semester: '', academicYear: '', search: '', yearLevel: '' },
  pagination: { page: 1, perPage: 10 }
};

function getDataByType(t) { return APP.allData.filter(d => d.type === t) }
// ชั้นปีที่อาจารย์ประจำชั้นรับผิดชอบ (ว่าง = ไม่ใช่ classTeacher → ไม่กรอง)
function ctYear() { return APP.currentRole === 'classTeacher' ? norm((APP.currentUser && APP.currentUser.responsible_year) || '') : ''; }

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
  const v = prompt(`กรอกหมายเลขห้องเรียนประจำของชั้นปี ${yr}\nถ้ามีหลายห้องย่อย คั่นด้วยจุลภาค (,) จะแสดงเป็นลำดับลงมา\nเช่น: ห้อง A 101, ห้อง B 202`, current);
  if (v === null) return;
  setHomeroomNumber(yr, v);
  if (typeof renderCurrentPage === 'function') renderCurrentPage();
}

// ======================== LOGIN ACTIVITY LOG ========================
// Save login/logout events to Google Sheet "login_log" tab
const LOGIN_LOG_ROLE_LABEL = { admin: 'ผู้ดูแลระบบ', academic: 'เจ้าหน้าที่งานวิชาการ', teacher: 'อาจารย์', classTeacher: 'อาจารย์ประจำชั้น', student: 'นักศึกษา', executive: 'ผู้บริหาร', registrar: 'เจ้าหน้าที่งานทะเบียน', deptHead: 'ประธานสาขาวิชา' };
// บทบาทที่ใช้ฟีเจอร์ลืม/เปลี่ยนรหัสผ่านผ่านอีเมลได้
const PWD_SELF_ROLES = ['academic', 'executive', 'teacher', 'deptHead', 'classTeacher'];
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
      // Fire and forget - บันทึกแบบไม่อ่าน log กลับ (กันดึง login_log ทั้งหมดมาที่เบราว์เซอร์นักศึกษา)
      const writer = GSheetDB.appendNoRefresh ? GSheetDB.appendNoRefresh(obj) : GSheetDB.create(obj);
      Promise.resolve(writer).catch(err => console.warn('logLoginEvent error:', err));
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
// รายการสาขาวิชา (รวมจากอาจารย์ + รายวิชา) สำหรับ dropdown ที่พิมพ์เองได้
function deptDatalistHTML(listId) {
  const ds = [...new Set([...getDataByType('teacher').map(t => norm(t.department)), ...getDataByType('subject').map(s => norm(s.department))].filter(Boolean))].sort();
  return `<datalist id="${listId}">${ds.map(d => `<option value="${d}"></option>`).join('')}</datalist>`;
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

async function resetGSheetConfig() {
  // ตรวจรหัสผ่านผู้ดูแลฝั่งเซิร์ฟเวอร์ (Apps Script) แทนการเทียบในเบราว์เซอร์
  const p = prompt("กรุณากรอกรหัสผ่านผู้ดูแลระบบ (Admin) 6 หลัก เพื่อเปลี่ยน Google Sheet:");
  if (!p) return;
  const res = await GSheetDB.login({ role: 'admin', identifier: '', password: p.trim() });
  if (!res || !res.isOk || !res.user) {
    alert("รหัสผ่านผู้ดูแลระบบไม่ถูกต้อง");
    return;
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

// กู้รหัสวิชาที่เลข 0 นำหน้าหาย (เช่น 101300101) ให้กลับเป็นรหัสเต็ม (0101300101)
// โดยเทียบกับรหัสจริงในแท็บ subject — แก้เฉพาะหน่วยความจำ (การแสดงผล) ไม่แตะ Google Sheet
function normalizeSubjectCodes(data) {
  const map = {};
  data.forEach(d => {
    if (d.type !== 'subject') return;
    const c = String(d.subject_code || '').trim();
    if (/^\d+$/.test(c)) { const k = c.replace(/^0+/, ''); if (k && !(k in map)) map[k] = c; }
  });
  if (!Object.keys(map).length) return;
  data.forEach(d => {
    if (d.type === 'subject' || !d.subject_code) return;
    const c = String(d.subject_code).trim();
    if (/^\d+$/.test(c)) {
      const k = c.replace(/^0+/, '');
      if (map[k] && map[k] !== c) d.subject_code = map[k];
    }
  });
}

function initGSheet(sheetId, scriptUrl) {
  showScreen('loadingScreen');
  GSheetDB.init({ spreadsheetId: sheetId, scriptUrl: scriptUrl || '' }, (data) => {
    normalizeSubjectCodes(data);
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
  // นักศึกษา: รีเฟรชเฉพาะข้อมูลตัวเอง / บุคลากร: โหลดครบทุกแท็บ
  if (APP.currentRole === 'student') await GSheetDB.studentRefresh();
  else await GSheetDB.refresh();
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
  if (r === 'registrar' || r === 'เจ้าหน้าที่งานทะเบียน' || r === 'งานทะเบียน') return 'registrar';
  if (r === 'depthead' || r === 'ประธานสาขาวิชา' || r === 'ประธานสาขา' || r === 'หัวหน้าสาขาวิชา') return 'deptHead';
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
  } else if (role === 'executive' || role === 'classTeacher' || role === 'registrar' || role === 'deptHead') {
    f.innerHTML = `<div class="mb-4"><label class="block text-sm font-medium text-gray-700 mb-2">Username</label><input type="text" id="loginUsername" class="w-full border border-gray-300 rounded-xl px-4 py-3 focus:ring-2 focus:ring-primary focus:border-primary outline-none" placeholder="กรอก Username"></div><div class="mb-4"><label class="block text-sm font-medium text-gray-700 mb-2">รหัสผ่าน</label><input type="password" id="loginUserPass" class="w-full border border-gray-300 rounded-xl px-4 py-3 focus:ring-2 focus:ring-primary focus:border-primary outline-none" placeholder="รหัสผ่าน" onkeypress="if(event.key==='Enter')handleLogin()"></div>`;
  } else {
    f.innerHTML = `<div class="mb-4"><label class="block text-sm font-medium text-gray-700 mb-2">E-mail</label><input type="email" id="teacherEmail" class="w-full border border-gray-300 rounded-xl px-4 py-3 focus:ring-2 focus:ring-primary focus:border-primary outline-none" placeholder="E-mail"></div><div class="mb-4"><label class="block text-sm font-medium text-gray-700 mb-2">รหัสผ่าน</label><input type="password" id="teacherPass" class="w-full border border-gray-300 rounded-xl px-4 py-3 focus:ring-2 focus:ring-primary focus:border-primary outline-none" placeholder="รหัสผ่าน" onkeypress="if(event.key==='Enter')handleLogin()"></div>`;
  }
  if (PWD_SELF_ROLES.includes(role)) {
    f.innerHTML += `<div class="text-right -mt-1 mb-1"><button type="button" onclick="showPasswordOtpModal('forgot')" class="text-sm text-primary hover:underline">ลืมรหัสผ่าน?</button></div>`;
  }
}
updateLoginFields();

async function handleLogin() {
  const role = document.getElementById('loginRole').value;
  const err = document.getElementById('loginError');
  err.classList.add('hidden');

  // ===== นักศึกษา: ตรวจฝั่งเซิร์ฟเวอร์ (เลขบัตรไม่ถูกส่งมาเทียบในเบราว์เซอร์) =====
  if (role === 'student') {
    const nid = document.getElementById('studentNID').value.trim();
    if (!/^\d{13}$/.test(nid)) { err.textContent = 'กรุณากรอกเลขบัตรประชาชน 13 หลัก'; err.classList.remove('hidden'); return }
    const loginBtn0 = document.querySelector('#loginScreen button[onclick="handleLogin()"]');
    if (loginBtn0) { loginBtn0.disabled = true; loginBtn0.textContent = 'กำลังตรวจสอบ...'; }
    showScreen('loadingScreen');  // แสดงหน้าโหลดข้อมูลระหว่างดึงข้อมูล (เหมือนบทบาทอื่น)
    const sres = await GSheetDB.studentLogin(nid);
    if (loginBtn0) { loginBtn0.disabled = false; loginBtn0.textContent = 'เข้าสู่ระบบ'; }
    if (!sres || !sres.isOk || !sres.student) { showScreen('loginScreen'); err.textContent = 'ไม่พบข้อมูลนักศึกษา กรุณาตรวจสอบเลขบัตรประชาชน'; err.classList.remove('hidden'); return }
    const stu = sres.student;
    if (norm(stu.status) === 'สำเร็จการศึกษา' || norm(stu.year_level) === 'จบ') { showScreen('loginScreen'); err.textContent = 'บัญชีนี้เป็นผู้สำเร็จการศึกษาแล้ว ไม่สามารถเข้าสู่ระบบได้'; err.classList.remove('hidden'); return }
    APP.currentUser = { name: stu.name, role: 'student', data: stu };
  } else {
    // ===== บุคลากร: ตรวจรหัสผ่านฝั่งเซิร์ฟเวอร์ (Apps Script) — รหัสผ่านไม่ถูกเทียบในเบราว์เซอร์ =====
    let identifier = '', password = '';
    if (role === 'admin') {
      password = document.getElementById('adminPass').value;
      if (!/^\d{6}$/.test(password)) { err.textContent = 'กรุณากรอกรหัสผ่าน 6 หลัก (ตัวเลขเท่านั้น)'; err.classList.remove('hidden'); return }
    } else if (role === 'teacher' || role === 'academic') {
      identifier = document.getElementById('teacherEmail').value.trim();
      password = document.getElementById('teacherPass').value;
      if (!identifier) { err.textContent = 'กรุณากรอก E-mail'; err.classList.remove('hidden'); return }
      if (!password) { err.textContent = 'กรุณากรอกรหัสผ่าน'; err.classList.remove('hidden'); return }
    } else {
      identifier = document.getElementById('loginUsername').value.trim();
      password = document.getElementById('loginUserPass').value;
      if (!identifier) { err.textContent = 'กรุณากรอก Username'; err.classList.remove('hidden'); return }
      if (!password) { err.textContent = 'กรุณากรอกรหัสผ่าน'; err.classList.remove('hidden'); return }
    }

    const loginBtn = document.querySelector('#loginScreen button[onclick="handleLogin()"]');
    if (loginBtn) { loginBtn.disabled = true; loginBtn.textContent = 'กำลังตรวจสอบ...'; }
    const res = await GSheetDB.login({ role, identifier, password });
    if (loginBtn) { loginBtn.disabled = false; loginBtn.textContent = 'เข้าสู่ระบบ'; }

    if (!res || !res.isOk || !res.user) {
      err.textContent = (role === 'admin')
        ? 'รหัสผ่านไม่ถูกต้อง'
        : (role === 'teacher' || role === 'academic') ? 'E-mail หรือรหัสผ่านไม่ถูกต้อง' : 'Username หรือรหัสผ่านไม่ถูกต้อง';
      err.classList.remove('hidden'); return;
    }
    const u = res.user;

    // บุคลากร: โหลดข้อมูลครบทุกแท็บหลังล็อกอินสำเร็จ (เดิมโหลดตั้งแต่เปิดหน้า)
    // ต้องโหลดก่อนบรรทัด getDataByType ด้านล่าง เพื่อให้ผูกข้อมูลอาจารย์/สาขาได้ถูกต้อง
    showScreen('loadingScreen');
    await GSheetDB.refresh();

    if (role === 'admin') {
      APP.currentUser = { name: u.name || 'ผู้ดูแลระบบ', role: 'admin' };
    } else if (role === 'teacher') {
      const t = getDataByType('teacher').find(x => (x.email || '') === (u.email || identifier));
      APP.currentUser = t ? { name: t.name, role: 'teacher', data: t } : { name: u.name || identifier, role: 'teacher', email: u.email || identifier };
    } else if (role === 'academic') {
      APP.currentUser = { name: u.name || identifier, role: 'academic', email: u.email || identifier };
    } else if (role === 'classTeacher') {
      const t = getDataByType('teacher').find(x => (x.name || '').trim().toLowerCase() === (u.name || '').trim().toLowerCase() || (x.email || '') === (u.email || ''));
      APP.currentUser = t ? { name: t.name, role: 'classTeacher', data: t, responsible_year: t.responsible_year || u.responsible_year || '1' } : { name: u.name || identifier, role: 'classTeacher', responsible_year: u.responsible_year || '1', email: u.email || '' };
    } else if (role === 'executive') {
      APP.currentUser = { name: u.name || identifier, role: 'executive', email: u.email || '' };
    } else if (role === 'registrar') {
      APP.currentUser = { name: u.name || identifier, role: 'registrar' };
    } else if (role === 'deptHead') {
      const nameKey = (u.name || identifier).trim().toLowerCase();
      const t = getDataByType('teacher').find(x => (x.name || '').trim().toLowerCase() === nameKey || (x.email || '') === (u.email || ''));
      let dept = t ? norm(t.department) : '';
      if (!dept) { const td = getDataByType('teacher_directory').find(x => (x.name || '').trim().toLowerCase() === nameKey); if (td) dept = norm(td.nursing_branch); }
      if (!dept) dept = norm(u.department);
      APP.currentUser = { name: u.name || identifier, role: 'deptHead', department: dept, email: u.email || '' };
    }
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
  document.getElementById('currentUserRole').textContent = { admin: 'ผู้ดูแลระบบ', academic: 'เจ้าหน้าที่งานวิชาการ', teacher: 'อาจารย์', classTeacher: 'อาจารย์ประจำชั้น', student: 'นักศึกษา', executive: 'ผู้บริหาร', registrar: 'เจ้าหน้าที่งานทะเบียน', deptHead: 'ประธานสาขาวิชา' }[APP.currentRole];
  const _cpb = document.getElementById('changePwBtn');
  if (_cpb) _cpb.classList.toggle('hidden', !PWD_SELF_ROLES.includes(APP.currentRole));
  buildSidebar();
  navigateTo('dashboard');
  updateNotifBadge();   // อัปเดตป้ายแจ้งเตือนตามบทบาท/ข้อมูลที่เพิ่งโหลด
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
  APP.allData = [];
  if (GSheetDB.clearSession) GSheetDB.clearSession();   // ล้างเลขบัตร/ข้อมูลใน session
  const _cpb2 = document.getElementById('changePwBtn');
  if (_cpb2) _cpb2.classList.add('hidden');
  showScreen('loginScreen');
}

// ======================== SIDEBAR ========================
function buildSidebar() {
  const r = APP.currentRole;
  const p = APP.permissions[r] || {};
  let items = [];
  if (p.dashboard) items.push({ id: 'dashboard', icon: 'layout-dashboard', label: 'หน้าหลัก' });

  // Registration dropdown — permission-driven for all roles
  // ระบบทะเบียน: 1.ข้อมูลนักศึกษา 2.ข้อมูลอาจารย์ 3.ปฏิทินกิจกรรมวิชาการ (+ รายวิชาที่เปิดสอน)
  let regSub = [];
  if (r === 'student') {
    if (p.students) regSub.push({ id: 'studentInfo', label: 'ข้อมูลนักศึกษา' });
  } else {
    if (p.students) regSub.push({ id: 'students', label: 'ข้อมูลนักศึกษา' });
  }
  if (p.teachers) regSub.push({ id: 'teachers', label: 'ข้อมูลอาจารย์' });
  if (p.advisors) regSub.push({ id: 'advisorInfo', label: 'ข้อมูลอาจารย์ที่ปรึกษา' });
  if (p.specialTeachers) regSub.push({ id: 'specialTeachers', label: 'ข้อมูลอาจารย์พิเศษ' });
  if (p.alumni) regSub.push({ id: 'alumni', label: 'ข้อมูลศิษย์เก่า' });
  if (p.subjects) regSub.push({ id: 'subjects', label: 'รายวิชาที่เปิดสอน' });
  if (regSub.length) items.push({ id: 'registration', icon: 'book-open', label: 'ระบบทะเบียน', sub: regSub });

  // ปฏิทินกิจกรรมวิชาการ — แยกเป็นเมนูหลัก ต่อจากระบบทะเบียน
  if (p.schedule) items.push({ id: 'schedule', icon: 'calendar', label: 'ปฏิทินกิจกรรมวิชาการ' });

  // ผลการศึกษา: 1.ผลการเรียน 2.ผลสอบภาษาอังกฤษ
  let eduSub = [];
  if (p.grades) eduSub.push({ id: 'grades', label: 'ผลการเรียน' });
  if (p.engResults) eduSub.push({ id: 'engResults', label: 'ผลสอบภาษาอังกฤษ' });
  if (eduSub.length) items.push({ id: 'eduResults', icon: 'graduation-cap', label: 'ผลการศึกษา', sub: eduSub });

  if (p.teacherDirectory) items.push({ id: 'teacherDirectory', icon: 'award', label: 'ทำเนียบอาจารย์' });

  // Tracking dropdown
  let trackSub = [];
  if (p.tracking) trackSub.push({ id: 'tracking', label: 'ส่งรายละเอียดรายวิชา' });
  if (p.resultTracking) trackSub.push({ id: 'resultTracking', label: 'ส่งผลการดำเนินงานรายวิชา' });
  if (p.gradeTracking) trackSub.push({ id: 'gradeTracking', label: 'ส่งเกรดรายวิชา' });
  if (p.fileTracking) trackSub.push({ id: 'fileTracking', label: 'ส่งแฟ้มรายวิชา' });
  if (trackSub.length) items.push({ id: 'trackingGroup', icon: 'clipboard-list', label: 'ติดตามการส่ง', sub: trackSub });

  if (p.leave) items.push({ id: 'leave', icon: 'calendar-off', label: 'ระบบการลาของนักศึกษา' });
  // แบบประเมินความพึงพอใจ — ผู้ใช้ทั่วไป (7 บทบาท) ทำแบบประเมิน, admin จัดการ+ดูสรุปผล
  if (p.survey) items.push({ id: 'survey', icon: 'clipboard-check', label: 'แบบประเมินความพึงพอใจ' });
  if (p.surveyManage) items.push({ id: 'surveyManage', icon: 'clipboard-check', label: 'แบบประเมินความพึงพอใจ' });
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
  APP.filters._resultTrackingYear = '';
  APP.filters._gradeTrackingYear = '';
  APP.filters._fileTrackingYear = '';
  APP.filters._trackingDept = '';
  APP.filters._resultTrackingDept = '';
  APP.filters._gradeTrackingDept = '';
  APP.filters._fileTrackingDept = '';
  APP.filters._studentYearLevel = '';
  APP.filters._engAdvisor = '';
  APP.filters._gradeAdvisor = '';
  APP.filters._advisorDept = '';
  APP.filters._advisorSearch = '';
  APP.filters._advisorSelected = '';
  APP.filters._gradeBatch = '';
  APP.filters._engYear = '';
  APP.filters._engYearLevel = '';
  APP.filters._gradeYearLevel = '';
  APP._directoryTab = 'all';
  APP._directoryView = 'list';
  APP.filters._directoryYear = '';
  APP.filters._pageYear = '';
  APP.filters._subjectBatch = '';
  APP.filters._surveyYear = '';
  APP.filters._surveyManageYear = '';
  APP.filters._surveyQRoleFilter = '';
  APP.filters._surveyManageRole = '';
  APP._surveyManageTab = 'config';
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
function norm(v) { return String(v || '').replace(/\.0$/, '').replace(/\s+/g, ' ').trim() }

// สถานะที่ไม่นับรวมเป็นจำนวนนักศึกษา (สำเร็จการศึกษา / พักการศึกษา / ลาออก)
const NON_ACTIVE_STATUS = ['สำเร็จการศึกษา', 'พักการศึกษา', 'ลาออก', 'ขอโอนย้ายสถานศึกษา'];
// นับเฉพาะนักศึกษาที่ "กำลังศึกษาอยู่" — ตัดผู้ที่สถานะไม่ active และผู้ที่สำเร็จการศึกษา (year_level = 'จบ') ออก
function isActiveStudent(s) { return !NON_ACTIVE_STATUS.includes(norm(s && s.status)) && norm(s && s.year_level) !== 'จบ'; }
// คืนเฉพาะนักศึกษาที่ยังกำลังศึกษา (ใช้สำหรับการนับจำนวน)
function activeStudents(list) { return (list || []).filter(isActiveStudent); }

// ======================== ROLE HELPERS (RBAC) ========================
// สิทธิ์ "แก้ไขเต็มที่เหมือน admin" สำหรับหน้าทั่วไป: admin / งานวิชาการ / งานทะเบียน
// + ประธานสาขาวิชา (deptHead) ให้แก้ไขได้เฉพาะหน้าติดตามการส่ง และทำเนียบอาจารย์
function isAdminRole() {
  const r = APP.currentRole;
  if (r === 'admin' || r === 'academic' || r === 'registrar') return true;
  if (r === 'deptHead' && ['tracking', 'resultTracking', 'gradeTracking', 'fileTracking', 'teacherDirectory'].includes(APP.currentPage)) return true;
  return false;
}
// สิทธิ์ระดับ admin-only เดิม (ข้อมูลอาจารย์ ฯลฯ) — ให้งานทะเบียนทำได้เหมือน admin
function isAdminOnlyRole() { const r = APP.currentRole; return r === 'admin' || r === 'registrar'; }

// ======================== DEPARTMENT (สาขาวิชา) HELPERS ========================
function deptEq(a, b) { return norm(a) !== '' && norm(a) === norm(b); }
// รายวิชา 1 วิชารองรับได้หลายสาขา — คั่นด้วย , ; หรือ /
function splitDepts(v) { return norm(v).split(/[,;/]+/).map(x => x.trim()).filter(Boolean); }
function subjectDeptsOf(s) { return splitDepts(s && s.department); }
function subjectHasDept(s, dept) { const d = norm(dept); return d !== '' && subjectDeptsOf(s).some(x => norm(x) === d); }
// หาสาขาวิชา (อาจมีหลายสาขา) ของ tracking record โดยจับคู่กับรายวิชา ด้วยรหัสวิชาก่อน แล้วชื่อวิชา
function trackingDeptsOf(t) {
  const subs = getDataByType('subject');
  const code = norm(t.subject_code), name = norm(t.subject_name);
  let m = code ? subs.find(s => norm(s.subject_code) === code) : null;
  if (!m && name) m = subs.find(s => norm(s.subject_name) === name);
  return m ? subjectDeptsOf(m) : [];
}
function trackingHasDept(t, dept) { const d = norm(dept); return d !== '' && trackingDeptsOf(t).some(x => norm(x) === d); }
// สาขาวิชาของผู้ใช้ปัจจุบัน (ประธานสาขาวิชา) — ผูกจากชื่ออาจารย์ตอน login
function currentDept() { return norm(APP.currentUser && APP.currentUser.department); }
// รายการสาขาวิชาทั้งหมดจากรายวิชา (กระจายกรณีมีหลายสาขาต่อวิชา)
function allSubjectDepts() { const set = new Set(); getDataByType('subject').forEach(s => subjectDeptsOf(s).forEach(d => set.add(d))); return [...set].sort(); }
// ตัวกรองสาขาวิชาในหน้าติดตามการส่ง (admin = dropdown, ประธานสาขา = ป้ายสาขาตัวเอง)
function trackingDeptFilterHTML(filterKey, selectedDept) {
  if (APP.currentRole === 'deptHead') {
    return `<label class="text-sm font-medium text-gray-700 ml-2">สาขาวิชา:</label><span class="text-sm font-semibold text-primary px-3 py-2 bg-primaryLight rounded-lg">${currentDept() || '— ไม่พบสาขา —'}</span>`;
  }
  const depts = allSubjectDepts();
  if (!depts.length) return '';
  return `<label class="text-sm font-medium text-gray-700 ml-2"><i data-lucide="git-branch" class="w-4 h-4 inline mr-1"></i>สาขาวิชา:</label>
    <select onchange="APP.filters.${filterKey}=this.value;renderCurrentPage()" class="border border-gray-200 rounded-xl px-3 py-2 text-sm">
      <option value="">-- ทุกสาขา --</option>
      ${depts.map(d => `<option value="${d}" ${selectedDept === d ? 'selected' : ''}>${d}</option>`).join('')}
    </select>`;
}
// ใช้กรองข้อมูล tracking + รายวิชา ตามสาขา (deptHead ใช้สาขาตัวเอง, admin ใช้ค่าที่เลือก)
function applyDeptFilter(dataArr, subjArr, filterKey) {
  const isHead = APP.currentRole === 'deptHead';
  const dept = isHead ? currentDept() : (APP.filters[filterKey] || '');
  if (!isHead && !dept) return { data: dataArr, subjects: subjArr };
  return {
    data: dataArr.filter(t => trackingHasDept(t, dept)),
    subjects: subjArr.filter(s => subjectHasDept(s, dept))
  };
}

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

// ตรวจว่ารายวิชาๆ หนึ่งมีอาจารย์คนนี้เป็น coordinator หรือไม่
// coordinator ในรายวิชาอาจเก็บเป็นรายชื่อ คั่นด้วย ',' หรือ '/'
function subjectHasCoordinator(subject, teacherName) {
  const name = norm(teacherName);
  if (!name) return false;
  const list = String(subject.coordinator || '').split(/[,\/]/).map(s => norm(s)).filter(Boolean);
  if (!list.length) return false;
  return list.some(c => c === name || c.includes(name) || name.includes(c));
}

// ตัวช่วยจับคู่รายวิชา ↔ tracking record
// คืนค่าฟังก์ชัน isTracked(subject) ที่จะ true ถ้ารายวิชานี้มี tracking record อยู่แล้ว
// กฎ:
//   1) ถ้ามี subject_code ทั้ง 2 ฝั่ง → ใช้ code+sem[+year] (แม่นที่สุด)
//   2) ถ้าไม่มี → ใช้ name+sem[+year]
//   3) ถ้า tracking record ไม่ได้บันทึก academic_year → ใช้ key แบบไม่มีปี (ภายในชุดข้อมูลของปีที่เลือกอยู่แล้ว)
function makeTrackingMatcher(trackingRecords) {
  const codeFull = new Set();
  const codeNoYear = new Set();
  const nameFull = new Set();
  const nameNoYear = new Set();
  trackingRecords.forEach(t => {
    const code = norm(t.subject_code);
    const name = norm(t.subject_name);
    const sem = normSem(t.semester);
    const year = norm(t.academic_year);
    if (code) {
      codeFull.add(`${code}|${sem}|${year}`);
      if (!year) codeNoYear.add(`${code}|${sem}`);
    }
    if (name) {
      nameFull.add(`${name}|${sem}|${year}`);
      if (!year) nameNoYear.add(`${name}|${sem}`);
    }
  });
  return function isTracked(s) {
    const code = norm(s.subject_code);
    const name = norm(s.subject_name);
    const sem = normSem(s.semester);
    const year = norm(s.academic_year);
    if (code) {
      if (codeFull.has(`${code}|${sem}|${year}`)) return true;
      if (codeNoYear.has(`${code}|${sem}`)) return true;
    }
    if (name) {
      if (nameFull.has(`${name}|${sem}|${year}`)) return true;
      if (nameNoYear.has(`${name}|${sem}`)) return true;
    }
    return false;
  };
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
    <img src="https://cdn.jsdelivr.net/gh/JOB-BCNB-P/picture/cat_pose_white.gif" alt="" class="w-16 h-16 object-contain">
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
  const iconHTML = type === 'loading'
    ? `<img src="https://cdn.jsdelivr.net/gh/JOB-BCNB-P/picture/cat_pose_white.gif" style="width:1.7em;height:1.7em;object-fit:contain;flex-shrink:0" alt="">`
    : `<i data-lucide="${icon}" class="w-5 h-5"></i>`;
  d.innerHTML = iconHTML + msg;
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
  if (btn) { btn.disabled = true; btn.innerHTML = '<img src="https://cdn.jsdelivr.net/gh/JOB-BCNB-P/picture/cat_run_transparent.gif" class="cat-run-inline" alt="">กำลังบันทึก...'; lucide.createIcons(); }
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

// คืนรายการใบลาที่ "รอ" การอนุมัติของ user ปัจจุบัน (ตามบทบาท teacher/classTeacher/executive)
// หลังจาก user กดอนุมัติ → leave นั้นจะไม่อยู่ในลิสต์นี้อีก → bell count ลดอัตโนมัติ
function getPendingLeavesForCurrentRole() {
  const role = APP.currentRole;
  if (role !== 'teacher' && role !== 'classTeacher' && role !== 'executive') return [];
  let leaves = getDataByType('leave').filter(l => l.leave_status !== 'ปฏิเสธ');
  if (role === 'teacher') {
    const myName = (APP.currentUser.name || '').trim();
    leaves = leaves.filter(l => {
      if ((l.coordinator_approval || 'รอ') !== 'รอ') return false;
      let coordStr = (l.coordinator || '').trim();
      if (!coordStr) {
        const sub = getDataByType('subject').find(s =>
          s.subject_name === l.subject_name &&
          normSem(s.semester) === normSem(l.semester) &&
          norm(s.academic_year) === norm(l.academic_year)
        );
        if (sub && sub.coordinator) coordStr = sub.coordinator;
      }
      if (!coordStr) {
        const sub = getDataByType('subject').find(s => s.subject_name === l.subject_name);
        if (sub && sub.coordinator) coordStr = sub.coordinator;
      }
      if (!coordStr) return false;
      const coords = String(coordStr).split(/[,;|]|\sและ\s|\sand\s/).map(c => c.trim()).filter(Boolean);
      return coords.some(c => c === myName || c.includes(myName) || myName.includes(c));
    });
  } else if (role === 'classTeacher') {
    const yr = APP.currentUser.responsible_year || '1';
    const stuNameSet = new Set(
      getDataByType('student').filter(s => norm(s.year_level) === norm(yr)).map(s => (s.name || '').trim())
    );
    const myName = (APP.currentUser.name || '').trim();
    leaves = leaves.filter(l => {
      const isMyStudent = stuNameSet.has((l.name || '').trim()) || ((l.class_teacher || '').trim() === myName);
      return isMyStudent && allCoordinatorsApproved(l) && (l.class_teacher_approval || 'รอ') === 'รอ';
    });
  } else if (role === 'executive') {
    leaves = leaves.filter(l =>
      l.coordinator_approval === 'อนุมัติ' &&
      l.class_teacher_approval === 'อนุมัติ' &&
      (l.deputy_approval || 'รอ') === 'รอ'
    );
  }
  return leaves;
}

// คืนรายการ tracking record (รายละเอียด/ผลการดำเนินงาน/เกรด/แฟ้ม)
// ที่ "รอ" การอนุมัติของ user ปัจจุบัน (classTeacher/academic/admin/executive)
// type: tracking/result_tracking ใช้ field class_teacher_check, academic_propose, deputy_sign
// type: grade_tracking/file_tracking ใช้ field coordinator_check, academic_check, deputy_sign
function getPendingTrackingsForCurrentRole() {
  const role = APP.currentRole;
  if (role !== 'classTeacher' && role !== 'admin' && role !== 'academic' && role !== 'executive') return [];
  const types = [
    { type: 'tracking', page: 'tracking', label: 'รายละเอียดรายวิชา', step1: 'class_teacher_check', step2: 'academic_propose', step3: 'deputy_sign' },
    { type: 'result_tracking', page: 'resultTracking', label: 'ผลการดำเนินงาน', step1: 'class_teacher_check', step2: 'academic_propose', step3: 'deputy_sign' },
    { type: 'grade_tracking', page: 'gradeTracking', label: 'เกรด', step1: 'coordinator_check', step2: 'academic_check', step3: 'deputy_sign' },
    { type: 'file_tracking', page: 'fileTracking', label: 'แฟ้มรายวิชา', step1: 'coordinator_check', step2: 'academic_check', step3: 'deputy_sign' },
  ];
  const DONE = 'เสร็จสิ้น';
  const needAct = v => { const x = norm(v); return (!x || x === 'รอ' || x === 'ส่งกลับแก้ไข'); };
  const isPendingForRole = (rec, t) => {
    const s1 = norm(rec[t.step1]), s2 = norm(rec[t.step2]), s3 = norm(rec[t.step3]);
    // ลงนามครบ (ขั้นสุดท้ายเสร็จสิ้น) = จบกระบวนการแล้ว ไม่ต้องแจ้งเตือนใครอีก
    if (s3 === DONE) return false;
    if (role === 'classTeacher') {
      // ถ้าขั้นถัดไปดำเนินการแล้ว = ผ่านขั้นนี้ไปแล้ว ไม่ต้องแจ้ง
      if (s2 === DONE) return false;
      return needAct(s1);
    }
    if (role === 'admin' || role === 'academic') {
      return s1 === DONE && needAct(s2);
    }
    if (role === 'executive') {
      return s2 === DONE && needAct(s3);
    }
    return false;
  };
  const pendings = [];
  types.forEach(t => {
    const records = getDataByType(t.type).filter(r => r.subject_name && r.subject_name.trim());
    records.forEach(r => {
      if (isPendingForRole(r, t)) pendings.push({ rec: r, type: t });
    });
  });
  return pendings;
}

// รายการติดตามที่ "ยังไม่ได้ดำเนินการเลย" ทุกขั้น (ข้อมูลย้อนหลัง/ค้าง) — สำหรับปุ่มปิดแจ้งเตือนย้อนหลัง
function getUnstartedTrackings() {
  const types = [
    { type: 'tracking', steps: ['class_teacher_check', 'academic_propose', 'deputy_sign'] },
    { type: 'result_tracking', steps: ['class_teacher_check', 'academic_propose', 'deputy_sign'] },
    { type: 'grade_tracking', steps: ['coordinator_check', 'academic_check', 'deputy_sign'] },
    { type: 'file_tracking', steps: ['coordinator_check', 'academic_check', 'deputy_sign'] },
  ];
  const out = [];
  types.forEach(t => {
    getDataByType(t.type).filter(r => r.subject_name && r.subject_name.trim()).forEach(r => {
      const anyDone = t.steps.some(k => norm(r[k]) === 'เสร็จสิ้น');
      if (!anyDone) out.push({ rec: r, steps: t.steps });
    });
  });
  return out;
}

// ปิดแจ้งเตือนย้อนหลัง: บันทึกรายการที่ยังไม่ได้ดำเนินการเป็นเสร็จสิ้นครบทุกขั้น
async function clearTrackingBacklog() {
  const backlog = getUnstartedTrackings();
  if (!backlog.length) { showToast('ไม่มีรายการค้าง', 'error'); return; }
  showModal('ยืนยันปิดแจ้งเตือนย้อนหลัง', `<div class="text-center text-gray-600 text-sm space-y-2"><p>บันทึก <b>${backlog.length}</b> รายการติดตามที่ยังไม่ได้ดำเนินการ ให้เป็น "เสร็จสิ้น" ทุกขั้น และปิดแจ้งเตือน?</p><p class="text-xs text-amber-600">⚠️ ใช้สำหรับข้อมูลย้อนหลังที่ดำเนินการเสร็จแล้ว — รายการเหล่านี้จะถูกทำเครื่องหมายว่าเสร็จสิ้นทันที</p></div>`, async () => {
    closeModal();
    showToast('กำลังบันทึก...', 'loading');
    const today = new Date().toISOString().split('T')[0];
    const nowIso = new Date().toISOString();
    const by = (APP.currentUser && APP.currentUser.name) || APP.currentRole || '';
    const toUpdate = [];
    backlog.forEach(b => {
      const rec = b.rec;
      b.steps.forEach(k => { rec[k] = 'เสร็จสิ้น'; rec[k + '_at'] = nowIso; rec[k + '_by'] = by; });
      if (!rec.approved_date) rec.approved_date = today;
      toUpdate.push(rec);
    });
    const r = await GSheetDB.updateMany(toUpdate);
    hideLoadingToast();
    showToast((r && r.ok ? r.ok : toUpdate.length) + ' รายการถูกปิดแจ้งเตือนแล้ว');
    closeNotifications();
    updateNotifBadge();
    renderCurrentPage();
  });
}

function showNotifications() {
  document.getElementById('notifPanel').style.transform = 'translateX(0)';
  renderNotifications();
  // เคลียร์เฉพาะส่วน "ประกาศ" เมื่อเปิดดู — ส่วนใบลา/ติดตามจะค้างไว้จนกว่าจะอนุมัติจริง
  const seenCount = visibleAnnouncements().length;
  try { localStorage.setItem('notifSeenCount', String(seenCount)); } catch (e) { }
  updateNotifBadge();
}
function closeNotifications() { document.getElementById('notifPanel').style.transform = 'translateX(100%)' }
function renderNotifications() {
  const ann = visibleAnnouncements().slice(-10).reverse();
  const pendingLeaves = getPendingLeavesForCurrentRole();
  const pendingTracks = getPendingTrackingsForCurrentRole();
  let html = '';
  // === ปุ่มปิดแจ้งเตือนย้อนหลัง (ผู้ดูแล/งานวิชาการ) ===
  if (APP.currentRole === 'admin' || APP.currentRole === 'academic') {
    const backlog = getUnstartedTrackings();
    if (backlog.length) {
      html += `<div class="mb-3 p-3 bg-orange-50 border border-orange-200 rounded-xl">
        <p class="text-xs text-orange-800 mb-2"><i data-lucide="history" class="w-3 h-3 inline mr-0.5"></i>มีรายการติดตามที่ยังไม่ได้ดำเนินการ ${backlog.length} รายการ (ข้อมูลย้อนหลัง)</p>
        <button onclick="clearTrackingBacklog()" class="w-full text-sm bg-orange-600 hover:bg-orange-700 text-white py-2 rounded-lg flex items-center justify-center gap-1"><i data-lucide="check-check" class="w-4 h-4"></i>ปิดแจ้งเตือนย้อนหลังทั้งหมด</button>
      </div>`;
    }
  }
  // === ส่วน "ใบลารออนุมัติ" ===
  if (pendingLeaves.length) {
    html += `<div class="mb-3">
      <p class="text-xs font-semibold text-amber-700 mb-2 flex items-center gap-1">
        <i data-lucide="alert-circle" class="w-3 h-3"></i>ใบลารออนุมัติของคุณ (${pendingLeaves.length})
      </p>`;
    html += pendingLeaves.slice(0, 15).map(l => {
      const dateText = toBuddhistDateList ? toBuddhistDateList(l.leave_date) : (l.leave_date || '-');
      return `<div class="p-3 bg-amber-50 border border-amber-200 rounded-xl mb-2 cursor-pointer hover:bg-amber-100 transition" onclick="closeNotifications();navigateTo('leave')">
        <p class="font-medium text-sm text-gray-800">${l.name || ''}</p>
        <p class="text-xs text-gray-700 mt-1"><i data-lucide="book-open" class="w-3 h-3 inline mr-0.5"></i>${l.subject_name || ''}</p>
        <p class="text-xs text-gray-600 mt-1"><i data-lucide="calendar" class="w-3 h-3 inline mr-0.5"></i>${dateText} · ${l.leave_type || ''} · ${l.leave_hours || '-'} ชม.</p>
        <p class="text-xs text-amber-700 font-medium mt-1">⏳ คลิกเพื่อไปอนุมัติ</p>
      </div>`;
    }).join('');
    html += `</div>`;
  }
  // === ส่วน "ติดตามการส่งรออนุมัติ" ===
  if (pendingTracks.length) {
    if (pendingLeaves.length) html += `<div class="border-t pt-2 mt-1"></div>`;
    html += `<div class="mb-3">
      <p class="text-xs font-semibold text-orange-700 mb-2 flex items-center gap-1">
        <i data-lucide="file-check" class="w-3 h-3"></i>ติดตามการส่งรออนุมัติของคุณ (${pendingTracks.length})
      </p>`;
    html += pendingTracks.slice(0, 20).map(p => {
      const r = p.rec, t = p.type;
      const semText = r.semester ? ` · ภาค ${r.semester}` : '';
      const yrText = r.academic_year ? `/${r.academic_year}` : '';
      const coordText = r.coordinator ? ` · ${r.coordinator}` : '';
      const subText = `${r.subject_code ? r.subject_code + ' ' : ''}${r.subject_name || ''}`;
      return `<div class="p-3 bg-orange-50 border border-orange-200 rounded-xl mb-2 cursor-pointer hover:bg-orange-100 transition" onclick="closeNotifications();navigateTo('${t.page}')">
        <p class="text-[10px] text-orange-600 font-semibold uppercase">${t.label}</p>
        <p class="font-medium text-sm text-gray-800 mt-0.5"><i data-lucide="book-open" class="w-3 h-3 inline mr-0.5"></i>${subText}</p>
        <p class="text-xs text-gray-600 mt-1">${semText}${yrText}${coordText}</p>
        <p class="text-xs text-orange-700 font-medium mt-1">⏳ คลิกเพื่อไปอนุมัติ</p>
      </div>`;
    }).join('');
    html += `</div>`;
  }
  // === ส่วน "ประกาศ" ===
  if (ann.length) {
    if (pendingLeaves.length || pendingTracks.length) html += `<div class="border-t pt-2 mt-1"></div>`;
    html += `<p class="text-xs font-semibold text-blue-700 mb-2 flex items-center gap-1"><i data-lucide="megaphone" class="w-3 h-3"></i>ประกาศ (${ann.length})</p>`;
    html += ann.map(a => `<div class="p-3 bg-surface rounded-xl mb-2"><p class="font-medium text-sm">${a.announcement_title || ''}</p><p class="text-xs text-gray-500 mt-1">${a.announcement_date || ''}</p><p class="text-xs text-gray-600 mt-1">${a.announcement_content || ''}</p></div>`).join('');
  }
  document.getElementById('notifList').innerHTML = html || '<p class="text-gray-400 text-center text-sm">ไม่มีการแจ้งเตือน</p>';
  if (typeof lucide !== 'undefined' && lucide.createIcons) lucide.createIcons();
}
function updateNotifBadge() {
  const b = document.getElementById('notifBadge');
  if (!b) return;
  const annTotal = visibleAnnouncements().length;
  let seenAnn = 0;
  try { seenAnn = parseInt(localStorage.getItem('notifSeenCount') || '0', 10) || 0; } catch (e) { }
  const unseenAnn = Math.max(0, annTotal - seenAnn);
  // ใบลา + ติดตาม: นับค้างไว้จนกว่าจะอนุมัติจริง (รายการที่เสร็จ/ลงย้อนหลังถูกกรองออกแล้วใน getPending*)
  const pendingLeaves = getPendingLeavesForCurrentRole().length;
  const pendingTracks = getPendingTrackingsForCurrentRole().length;
  const total = unseenAnn + pendingLeaves + pendingTracks;
  if (total > 0) { b.textContent = total > 99 ? '99+' : total; b.classList.remove('hidden'); } else b.classList.add('hidden');
}

function paginationHTML(total, perPage, page, onChange) {
  const pages = Math.ceil(total / perPage) || 1;
  if (pages <= 1) return '';
  page = Math.min(Math.max(1, page), pages);

  // เลือกเลขหน้าที่จะแสดง: หน้าแรก + หน้าสุดท้าย + ช่วงรอบหน้าปัจจุบัน (±2) เสมอ
  let nums;
  if (pages <= 7) {
    nums = Array.from({ length: pages }, (_, i) => i + 1);
  } else {
    const set = new Set([1, pages, page, page - 1, page + 1, page - 2, page + 2]);
    nums = [...set].filter(n => n >= 1 && n <= pages).sort((a, b) => a - b);
  }

  let h = '<div class="flex items-center justify-center gap-1 mt-4 flex-wrap">';
  h += `<button onclick="${onChange}(1)" class="px-3 py-1 rounded-lg border text-sm hover:bg-gray-50 ${page === 1 ? 'opacity-40' : ''}" title="หน้าแรก">«</button>`;
  h += `<button onclick="${onChange}(${Math.max(1, page - 1)})" class="px-3 py-1 rounded-lg border text-sm hover:bg-gray-50 ${page === 1 ? 'opacity-40' : ''}" title="ก่อนหน้า">‹</button>`;
  let prev = 0;
  nums.forEach(n => {
    if (n - prev > 1) h += `<span class="px-2 text-gray-400 text-sm select-none">…</span>`;
    h += `<button onclick="${onChange}(${n})" class="px-3 py-1 rounded-lg text-sm ${n === page ? 'bg-primary text-white' : 'border hover:bg-gray-50'}">${n}</button>`;
    prev = n;
  });
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
    // ดึงปีการศึกษาเฉพาะจากข้อมูลจริงในระบบ (ไม่ hardcode ปีที่ไม่มีข้อมูล)
    const yrSet = new Set();
    const _yrSrc = Array.isArray(opts.yearData) ? opts.yearData : (APP.allData || []);
    _yrSrc.forEach(d => { const y = norm(d.academic_year); if (y) yrSet.add(y); });
    // ถ้ายังไม่มีข้อมูลเลย ให้ใส่ปีปัจจุบันเป็น default
    if (!yrSet.size) {
      const cur = new Date().getFullYear() + 543;
      yrSet.add(String(cur));
    }
    const yrOptions = [...yrSet].sort();
    h += `<select onchange="APP.filters.academicYear=this.value;APP.pagination.page=1;renderCurrentPage()" class="border border-gray-200 rounded-xl px-3 py-2.5 text-sm"><option value="">ทุกปีการศึกษา</option>${yrOptions.map(y => `<option value="${y}" ${yr === y ? 'selected' : ''}>${y}</option>`).join('')}</select>`;
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
  const allYears = [...new Set(allData.map(d => norm(d.academic_year)).filter(Boolean))].sort().reverse();
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
  return `<button onclick="downloadCSVTemplate('${type}','${fields}')" class="flex items-center gap-2 px-4 py-2 border border-emerald-500 text-emerald-600 rounded-xl hover:bg-emerald-50 text-sm" title="ดาวน์โหลดไฟล์ตัวอย่าง (เฉพาะหัวตาราง) สำหรับกรอกแล้ว Upload"><i data-lucide="download" class="w-4 h-4"></i>ตัวอย่าง CSV</button>
  <button onclick="triggerCSVUpload('${type}','${fields}')" class="flex items-center gap-2 px-4 py-2 border border-primary text-primary rounded-xl hover:bg-primaryLight text-sm"><i data-lucide="upload" class="w-4 h-4"></i>Upload CSV</button>
  <input type="file" id="csvInput_${type}" accept=".csv" class="hidden" onchange="handleCSVUpload(event,'${type}','${fields}')">`;
}

// ดาวน์โหลดไฟล์ CSV ตัวอย่าง (หัวตารางตรงกับที่ระบบรองรับ) สำหรับใช้กรอกข้อมูลแล้ว Upload
function downloadCSVTemplate(type, fields) {
  const bom = String.fromCharCode(0xFEFF); // BOM ให้ Excel เปิดภาษาไทยได้ถูกต้อง
  const blob = new Blob([bom + fields + '\r\n'], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = `template_${type}.csv`;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
  if (typeof showToast === 'function') showToast('ดาวน์โหลดไฟล์ตัวอย่าง CSV แล้ว');
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
    case 'advisorInfo': return advisorInfoPage();
    case 'subjects': return subjectsPage();
    case 'schedule': return schedulePage();
    case 'grades': return gradesPage();
    case 'engResults': return engResultsPage();
    case 'teachers': return teachersPage();
    case 'specialTeachers': return specialTeachersPage();
    case 'alumni': return alumniPage();
    case 'teacherDirectory': return teacherDirectoryPage();
    case 'services': return servicesPage();
    case 'tracking': return trackingPage();
    case 'resultTracking': return resultTrackingPage();
    case 'gradeTracking': return gradeTrackingPage();
    case 'fileTracking': return fileTrackingPage();
    case 'leave': return leavePage();
    case 'survey': return surveyPage();
    case 'surveyManage': return surveyManagePage();
    case 'settings': return settingsPage();
    case 'loginLog': return loginLogPage();
    case 'userGuide': return userGuidePage();
    default: return '<p>ไม่พบหน้าที่ต้องการ</p>';
  }
}

// ======================== DASHBOARD ========================
// แสดงห้องเรียนประจำ — ถ้ามีหลายห้องย่อย (คั่นด้วย , ; หรือขึ้นบรรทัดใหม่) ให้แสดงลงมาเป็นลำดับ
function renderHomeroomHTML(homeroom) {
  if (!norm(homeroom)) return '<span class="text-gray-400 font-normal">ยังไม่กำหนด</span>';
  const parts = String(homeroom).split(/[,;\n]+/).map(x => x.trim()).filter(Boolean);
  if (parts.length <= 1) return parts[0] || '';
  return parts.map(p => `<span class="block leading-tight">• ${p}</span>`).join('');
}
// การ์ดนักศึกษารายชั้นปี (ใช้ทั้งแดชบอร์ด admin และประธานสาขา)
function yearLevelCardsHTML(students, engPassRecords, canEdit) {
  const _hr = getHomeroomNumbers();
  return [1, 2, 3, 4].map(yr => {
    const yrStudents = activeStudents(students).filter(s => norm(s.year_level) === String(yr));
    const yrEngPassUnique = [...new Set(engPassRecords.filter(e => yrStudents.some(s => s.student_id === e.student_id)).map(e => e.student_id))];
    const homeroom = _hr[String(yr)] || '';
    return `<div class="bg-white rounded-2xl p-5 border border-blue-100">
      <p class="text-sm font-medium text-gray-500 mb-3">ชั้นปี ${yr}</p>
      <div class="flex items-center gap-4">
        <div class="flex-1"><p class="text-3xl font-bold text-primary leading-none">${yrStudents.length}</p><p class="text-xs text-gray-500 mt-1.5">นักศึกษา</p></div>
        <div class="w-px self-stretch bg-gray-100"></div>
        <div class="flex-1"><p class="text-3xl font-bold text-green-500 leading-none">${yrEngPassUnique.length}</p><p class="text-xs text-gray-500 mt-1.5">ผ่าน ENG</p></div>
      </div>
      <div class="mt-3 pt-3 border-t border-gray-100 flex items-start justify-between gap-2">
        <div class="flex items-start gap-2 min-w-0">
          <i data-lucide="door-open" class="w-4 h-4 text-blue-500 flex-shrink-0 mt-0.5"></i>
          <div class="min-w-0">
            <p class="text-xs text-gray-500 leading-tight">ห้องเรียนประจำ</p>
            <div class="text-sm font-bold text-gray-800">${renderHomeroomHTML(homeroom)}</div>
          </div>
        </div>
        ${canEdit ? `<button onclick="promptEditHomeroom('${yr}')" title="แก้ไขหมายเลขห้องเรียนประจำ" class="text-xs text-blue-600 hover:text-blue-800 hover:bg-blue-50 rounded-lg p-1.5 flex-shrink-0"><i data-lucide="pencil" class="w-3.5 h-3.5"></i></button>` : ''}
      </div>
    </div>`;
  }).join('');
}
function dashboardPage() {
  const students = getDataByType('student');
  const teachers = getDataByType('teacher');
  const allEngResults = getDataByType('eng_result');
  const engPassRecords = allEngResults.filter(e => e.eng_status === 'ผ่าน');
  // นับเฉพาะ "นักศึกษาที่กำลังศึกษา" ซึ่งสอบผ่าน (ให้ตรงกับหน้าผลสอบ — ไม่นับรหัสที่ไม่มีในรายชื่อ/จบ/พัก/ลาออก)
  const engPassIdSet = new Set(engPassRecords.map(e => norm(e.student_id)));
  const engPassStudentIds = activeStudents(students).filter(s => engPassIdSet.has(norm(s.student_id))).map(s => s.student_id);
  const announcements = visibleAnnouncements().slice(-5).reverse();
  const r = APP.currentRole;

  let stats = '';
  if (r === 'admin' || r === 'academic' || r === 'registrar' || r === 'executive') {
    // Teacher breakdown by department (count only active)
    const activeTeachers = teachers.filter(t => (t.teacher_status || 'ปฏิบัติงานอยู่') === 'ปฏิบัติงานอยู่');
    // อาจารย์ลาศึกษาต่อ (teacher_status = 'ลาศึกษาต่อ')
    const studyLeaveCount = teachers.filter(t => norm(t.teacher_status) === 'ลาศึกษาต่อ').length;
    // อาจารย์พิเศษ — แสดง/เลือกดูตามปีการศึกษาได้
    const _specialTeachers = getDataByType('special_teacher');
    const _specialYears = [...new Set(_specialTeachers.map(t => norm(t.academic_year)).filter(Boolean))].sort().reverse();
    const _selSpecialYear = APP.filters._dashSpecialYear || '';
    const specialTeacherCount = _selSpecialYear ? _specialTeachers.filter(t => norm(t.academic_year) === _selSpecialYear).length : _specialTeachers.length;
    const specialTeacherCard = `<div class="card-stat bg-white rounded-2xl p-5 border border-blue-100">
      <div class="flex items-center gap-4">
        <div class="w-12 h-12 bg-indigo-500 rounded-xl flex items-center justify-center"><i data-lucide="user-plus" class="w-6 h-6 text-white"></i></div>
        <div class="min-w-0"><p class="text-sm text-gray-500">จำนวนอาจารย์พิเศษ ${_selSpecialYear ? '(ปี ' + _selSpecialYear + ')' : '(ทุกปี)'}</p><p class="text-2xl font-bold text-gray-800">${specialTeacherCount} <span class="text-sm font-normal text-gray-500">คน</span></p></div>
      </div>
      <div class="mt-3 pt-3 border-t border-gray-100">
        <select onchange="APP.filters._dashSpecialYear=this.value;renderCurrentPage()" class="w-full border border-gray-200 rounded-lg px-2 py-1.5 text-xs bg-white">
          <option value="">ทุกปีการศึกษา</option>
          ${_specialYears.map(y => '<option value="' + y + '"' + (_selSpecialYear === y ? ' selected' : '') + '>ปีการศึกษา ' + y + '</option>').join('')}
        </select>
      </div>
    </div>`;
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

    const _canEditHr = APP.currentUser && (APP.currentUser.role === 'admin' || APP.currentUser.role === 'academic');
    stats = `
    <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4 mb-6">
      ${statCard('users', 'จำนวนนักศึกษาทั้งหมด', activeStudents(students).length, 'คน', 'bg-blue-500')}
      ${statCard('briefcase', 'จำนวนอาจารย์ (ปฏิบัติงาน)', activeTeachers.length, 'คน', 'bg-emerald-500')}
      ${statCard('graduation-cap', 'อาจารย์ลาศึกษาต่อ', studyLeaveCount, 'คน', 'bg-purple-500')}
      ${specialTeacherCard}
      ${statCard('check-circle', 'นักศึกษาสอบผ่านภาษาอังกฤษ', engPassStudentIds.length, 'คน', 'bg-amber-500')}
    </div>
    <h3 class="font-bold mb-3 text-gray-800">จำนวนอาจารย์แยกสาขาวิชา</h3>
    <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3 mb-6">${deptCards || '<p class="text-gray-400 text-sm col-span-3">ไม่มีข้อมูลสาขาวิชา</p>'}</div>
    <h3 class="font-bold mb-3 text-gray-800">นักศึกษารายชั้นปี</h3>
    <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 mb-6">
      ${yearLevelCardsHTML(students, engPassRecords, _canEditHr)}
    </div>`;
  } else if (r === 'deptHead') {
    stats = `
    <div class="grid grid-cols-1 sm:grid-cols-2 gap-4 mb-6">
      ${statCard('users', 'จำนวนนักศึกษาทั้งหมด', activeStudents(students).length, 'คน', 'bg-blue-500')}
      ${statCard('check-circle', 'นักศึกษาสอบผ่านภาษาอังกฤษ', engPassStudentIds.length, 'คน', 'bg-amber-500')}
    </div>
    <h3 class="font-bold mb-3 text-gray-800">นักศึกษารายชั้นปี</h3>
    <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 mb-6">
      ${yearLevelCardsHTML(students, engPassRecords, false)}
    </div>`;
  } else if (r === 'teacher') {
    const myStudents = activeStudents(students).filter(s => s.advisor === APP.currentUser.name);
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
    const myStudents = activeStudents(students).filter(s => norm(s.year_level) === norm(yr));
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
      <h3 class="font-bold mb-3 flex items-center gap-2"><i data-lucide="calendar" class="w-5 h-5 text-primary"></i>ปฏิทินกิจกรรมวิชาการ</h3>
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

// ======================== ANALYTICS HELPERS (เพิ่มเติม) ========================
// เพศของนักศึกษา — ใช้คอลัมน์ gender ถ้ามี ไม่งั้นเดาจากคำนำหน้าชื่อ (นาย/นาง/นางสาว)
function studentGender(s) {
  const g = norm(s && s.gender);
  if (g) {
    if (/^(ช|ชาย|m|male)/i.test(g)) return 'M';
    if (/^(ญ|หญิง|f|female)/i.test(g)) return 'F';
  }
  const n = norm(s && s.name);
  if (/^นางสาว|^น\.ส\.|^เด็กหญิง|^ด\.ญ\./.test(n)) return 'F';
  if (/^นาง(?!สาว)/.test(n)) return 'F';
  if (/^นาย|^เด็กชาย|^ด\.ช\./.test(n)) return 'M';
  return 'U';
}

// ชื่อสำหรับแสดงผล — เติมคำนำหน้า (title_prefix) ถ้าชื่อยังไม่มีคำนำหน้าอยู่แล้ว
function studentDisplayName(s) {
  const name = norm(s && s.name);
  const p = norm(s && s.title_prefix);
  if (p && name && name.indexOf(p) !== 0) return p + name;
  return name || '';
}

// กราฟวงกลม (SVG) — เอาเมาส์ชี้เซกเมนต์เพื่อดูจำนวน/เปอร์เซ็นต์ + เอฟเฟคหมุนตอนโหลด
function svgDonut(segments, centerLabel) {
  const total = segments.reduce((s, x) => s + (parseFloat(x.value) || 0), 0);
  const r = 42, C = 2 * Math.PI * r; let acc = 0;
  const ring = total ? segments.filter(s => (parseFloat(s.value) || 0) > 0).map(s => {
    const v = parseFloat(s.value) || 0; const len = v / total * C; const off = -acc; acc += len;
    const pct = Math.round(v / total * 100);
    return `<circle class="donut-seg" cx="60" cy="60" r="${r}" fill="none" stroke="${s.color}" stroke-width="20" stroke-dasharray="${len.toFixed(2)} ${(C - len).toFixed(2)}" stroke-dashoffset="${off.toFixed(2)}"><title>${s.label}: ${v} คน (${pct}%)</title></circle>`;
  }).join('') : `<circle cx="60" cy="60" r="${r}" fill="none" stroke="#e5e7eb" stroke-width="20"></circle>`;
  const legend = segments.map(s => { const v = parseFloat(s.value) || 0; const pct = total ? Math.round(v / total * 100) : 0; return `<div style="display:flex;align-items:center;gap:8px;margin-bottom:6px;font-size:13px"><span style="width:13px;height:13px;border-radius:3px;background:${s.color};flex-shrink:0"></span><span style="flex:1;color:#475569">${s.label}</span><b style="color:#1e293b">${v}</b><span style="color:#94a3b8;width:40px;text-align:right">${pct}%</span></div>`; }).join('');
  return `<div style="display:flex;align-items:center;gap:18px;flex-wrap:wrap">
    <div class="donut-wrap pie-spin" style="width:150px;height:150px;flex-shrink:0;position:relative">
      <svg viewBox="0 0 120 120" width="150" height="150" style="transform:rotate(-90deg)">${ring}</svg>
      <div style="position:absolute;inset:0;display:flex;flex-direction:column;align-items:center;justify-content:center;pointer-events:none"><span style="font-size:26px;font-weight:700;color:#1e6fba">${total || 0}</span><span style="font-size:11px;color:#94a3b8">${centerLabel || 'รวม'}</span></div>
    </div>
    <div style="flex:1;min-width:180px">${legend}</div>
  </div>`;
}

// กราฟแท่งแนวนอน (มีอนิเมชันโตจากซ้าย→ขวา) — items: {label,value,color}
function animBarRows(items) {
  const max = Math.max(1, ...items.map(i => parseFloat(i.value) || 0));
  return items.map(i => {
    const w = Math.round((parseFloat(i.value) || 0) / max * 100);
    return `<div>
      <div style="display:flex;justify-content:space-between;font-size:12px;margin-bottom:3px"><span style="color:#475569">${i.label}</span><b style="color:#1e293b">${i.value || 0}</b></div>
      <div style="background:#eef2f7;border-radius:7px;height:18px;overflow:hidden"><div class="grow-bar" style="--tw:${w}%;height:100%;background:${i.color || '#1e6fba'};border-radius:7px"></div></div>
    </div>`;
  }).join('');
}

// ===== ตัวการ์ตูนบัณฑิต 3D — ยกมือทักทายเมื่อเอาเมาส์ชี้ที่การ์ด =====
// ทั้งชายและหญิงมี 2 เฟรม (แขนลง ↔ ยกมือ) สลับกันเป็นการโบกมือทักทาย
const GC_M_DOWN = 'data:image/png;base64,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';
const GC_M_UP = 'data:image/png;base64,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';
const GC_F_DOWN = 'data:image/png;base64,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';
const GC_F_UP = 'data:image/png;base64,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';
function genderFigureSVG(sex) {
  const down = sex === 'F' ? GC_F_DOWN : GC_M_DOWN;
  const up = sex === 'F' ? GC_F_UP : GC_M_UP;
  const label = sex === 'F' ? 'นักศึกษาหญิง' : 'นักศึกษาชาย';
  return `<span class="gc-fig">`
    + `<img class="gc-frame gc-down" src="${down}" alt="${label}" draggable="false">`
    + `<img class="gc-frame gc-up" src="${up}" alt="" aria-hidden="true" draggable="false">`
    + `</span>`;
}

// แผงวิเคราะห์อัตราการคงอยู่ของนักศึกษา (หน้าข้อมูลนักศึกษา — สำหรับผู้ดูแล/วิชาการ/ผู้บริหาร)
function studentRetentionAnalyticsHTML() {
  const all = getDataByType('student');
  // --- สัดส่วนการคงอยู่ (ไม่รวมผู้สำเร็จการศึกษา) ---
  const nonGrad = all.filter(s => !isGraduate(s));
  const cActive = nonGrad.filter(isActiveStudent).length;
  const cLeave = nonGrad.filter(s => norm(s.status) === 'พักการศึกษา').length;
  const cResign = nonGrad.filter(s => norm(s.status) === 'ลาออก').length;
  const cTransfer = nonGrad.filter(s => norm(s.status) === 'ขอโอนย้ายสถานศึกษา').length;
  const baseTotal = cActive + cLeave + cResign + cTransfer;
  const retentionPct = baseTotal ? Math.round(cActive / baseTotal * 1000) / 10 : 0;
  const donut = svgDonut([
    { label: 'กำลังศึกษา (คงอยู่)', value: cActive, color: '#22c55e' },
    { label: 'พักการศึกษา', value: cLeave, color: '#f59e0b' },
    { label: 'ลาออก', value: cResign, color: '#ef4444' },
    { label: 'ขอโอนย้ายสถานศึกษา', value: cTransfer, color: '#f97316' }
  ], 'คน');

  // --- สรุปเหตุผลการลาออก/พักการศึกษา (จากคอลัมน์ status_reason) ---
  const attrition = nonGrad.filter(s => ['ลาออก', 'พักการศึกษา', 'ขอโอนย้ายสถานศึกษา'].includes(norm(s.status)));
  const reasonMap = {};
  attrition.forEach(s => { const k = norm(s.status_reason) || 'ไม่ระบุเหตุผล'; reasonMap[k] = (reasonMap[k] || 0) + 1; });
  const reasonPalette = ['#ef4444', '#f97316', '#f59e0b', '#eab308', '#84cc16', '#14b8a6', '#3b82f6', '#a855f7'];
  const reasonItems = Object.entries(reasonMap).sort((a, b) => b[1] - a[1]).slice(0, 8).map(([k, v], i) => ({ label: k, value: v, color: reasonPalette[i % reasonPalette.length] }));
  const reasonCard = `<div class="mt-6 pt-5 border-t border-gray-100">
    <p class="text-sm font-semibold text-gray-600 mb-1"><i data-lucide="clipboard-list" class="w-4 h-4 inline mr-1"></i>สรุปเหตุผลการลาออก/พักการศึกษา</p>
    <p class="text-xs text-gray-400 mb-3">ลาออก ${cResign} · พักการศึกษา ${cLeave} · โอนย้าย ${cTransfer} คน (รวม ${attrition.length} คน)</p>
    ${reasonItems.length ? `<div class="space-y-2.5" style="max-width:640px">${animBarRows(reasonItems)}</div><p class="text-[11px] text-gray-400 mt-2">* กรอกช่อง "เหตุผลการลาออก/พักการศึกษา" ในฟอร์มนักศึกษาให้ครบ กราฟจะละเอียดขึ้น</p>` : '<p class="text-sm text-gray-400">ยังไม่มีข้อมูลการลาออก/พักการศึกษา</p>'}
  </div>`;

  // --- ตารางเปรียบเทียบรับเข้า (รายรุ่น) vs กำลังศึกษา ---
  const batches = [...new Set(all.map(s => norm(s.batch)).filter(Boolean))].sort((a, b) => b.localeCompare(a, undefined, { numeric: true }));
  const cohortRows = batches.map(b => {
    const inB = all.filter(s => norm(s.batch) === b);
    const admitted = inB.length;
    const active = inB.filter(isActiveStudent).length;
    const resign = inB.filter(s => norm(s.status) === 'ลาออก').length;
    const leave = inB.filter(s => norm(s.status) === 'พักการศึกษา').length;
    const transfer = inB.filter(s => norm(s.status) === 'ขอโอนย้ายสถานศึกษา').length;
    const grad = inB.filter(isGraduate).length;
    const stayPct = admitted ? Math.round((admitted - resign - transfer) / admitted * 1000) / 10 : 0;
    const pctColor = stayPct >= 90 ? 'text-green-600' : stayPct >= 80 ? 'text-amber-600' : 'text-red-600';
    return `<tr class="border-t hover:bg-gray-50">
      <td class="px-3 py-2 font-medium">รุ่นที่ ${b}</td>
      <td class="px-3 py-2 text-center font-semibold text-primary">${admitted}</td>
      <td class="px-3 py-2 text-center text-green-600 font-semibold">${active}</td>
      <td class="px-3 py-2 text-center">${leave}</td>
      <td class="px-3 py-2 text-center">${resign}</td>
      <td class="px-3 py-2 text-center">${transfer}</td>
      <td class="px-3 py-2 text-center text-blue-600">${grad}</td>
      <td class="px-3 py-2 text-center font-bold ${pctColor}">${stayPct}%</td>
    </tr>`;
  }).join('');
  const cohortTable = `<div class="overflow-x-auto"><table class="w-full text-sm">
    <thead><tr class="bg-surface text-left">
      <th class="px-3 py-2 font-semibold">รุ่น</th>
      <th class="px-3 py-2 font-semibold text-center">รับเข้า</th>
      <th class="px-3 py-2 font-semibold text-center">กำลังศึกษา</th>
      <th class="px-3 py-2 font-semibold text-center">พักฯ</th>
      <th class="px-3 py-2 font-semibold text-center">ลาออก</th>
      <th class="px-3 py-2 font-semibold text-center">โอนย้าย</th>
      <th class="px-3 py-2 font-semibold text-center">สำเร็จฯ</th>
      <th class="px-3 py-2 font-semibold text-center">คงอยู่ %</th>
    </tr></thead>
    <tbody>${cohortRows || '<tr><td colspan="8" class="px-3 py-6 text-center text-gray-400">ไม่มีข้อมูลรุ่น</td></tr>'}</tbody>
  </table></div>
  <p class="text-xs text-gray-400 mt-2"><i data-lucide="info" class="w-3 h-3 inline mr-1"></i>คงอยู่ % = (รับเข้า − ลาออก − โอนย้าย) ÷ รับเข้า (นับผู้สำเร็จการศึกษาเป็นการคงอยู่จนจบ)</p>`;

  // --- เพศ (ชาย/หญิง) ทั้งหมด + รายชั้นปี (เฉพาะที่กำลังศึกษา) ---
  const activeAll = activeStudents(all);
  const gc = list => ({ M: list.filter(s => studentGender(s) === 'M').length, F: list.filter(s => studentGender(s) === 'F').length, U: list.filter(s => studentGender(s) === 'U').length });
  const tot = gc(activeAll);
  const genderCard = (label, count, total, trim, sex) => {
    const pct = total ? Math.round(count / total * 100) : 0;
    return `<div class="gender-card card-stat bg-white rounded-2xl p-4 border border-blue-100 flex items-center gap-4">
      <div style="flex-shrink:0">${genderFigureSVG(sex)}</div>
      <div class="flex-1 min-w-0">
        <p class="text-sm text-gray-500">${label}</p>
        <p class="text-4xl font-bold leading-none" style="color:${trim}">${count}<span class="text-sm font-normal text-gray-400"> คน</span></p>
        <div class="mt-2 bg-gray-100 rounded-full h-2 overflow-hidden"><div class="grow-bar" style="--tw:${pct}%;height:100%;background:${trim}"></div></div>
        <p class="text-xs text-gray-400 mt-1">${pct}% ของผู้กำลังศึกษา <span class="gc-hint" style="color:${trim}">· ชี้เพื่อทักทาย 👋</span></p>
      </div>
    </div>`;
  };
  const yrGenderRows = [1, 2, 3, 4].map(yr => {
    const g = gc(activeAll.filter(s => norm(s.year_level) === String(yr)));
    const t = g.M + g.F + g.U;
    const mPct = t ? g.M / t * 100 : 0;
    return `<div class="card-stat bg-white rounded-xl p-3 border border-blue-100">
      <p class="text-sm font-semibold text-gray-700 mb-2">ชั้นปี ${yr} <span class="font-normal text-gray-400">(${t} คน)</span></p>
      <div class="flex h-3 rounded-full overflow-hidden bg-gray-100 mb-2"><div class="grow-bar" style="--tw:${mPct}%;background:#3b82f6"></div><div class="grow-bar" style="--tw:${100 - mPct}%;background:#ec4899"></div></div>
      <div class="flex justify-between text-xs"><span class="text-blue-600 font-semibold">♂ ชาย ${g.M}</span><span class="text-pink-600 font-semibold">♀ หญิง ${g.F}</span></div>
    </div>`;
  }).join('');

  return `<details open class="bg-white rounded-2xl border border-blue-100 mb-4">
    <summary class="cursor-pointer select-none p-5 flex items-center justify-between">
      <span class="font-bold text-gray-800 flex items-center gap-2"><i data-lucide="activity" class="w-5 h-5 text-primary"></i>อัตราการคงอยู่ของนักศึกษา <span class="text-sm font-normal text-green-600">(คงอยู่รวม ${retentionPct}%)</span></span>
      <span class="text-xs text-gray-400 flex items-center gap-1">ยุบ/เปิด <i data-lucide="chevron-down" class="chev w-5 h-5"></i></span>
    </summary>
    <div class="px-5 pb-5 fade-in">
      <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <div><p class="text-sm font-semibold text-gray-600 mb-3">สัดส่วนการคงอยู่ (ไม่รวมผู้สำเร็จการศึกษา)</p>${donut}</div>
        <div><p class="text-sm font-semibold text-gray-600 mb-3">เปรียบเทียบรับเข้า vs กำลังศึกษา (รายรุ่น)</p>${cohortTable}</div>
      </div>
      ${reasonCard}
      <div class="mt-6 pt-5 border-t border-gray-100">
        <p class="text-sm font-semibold text-gray-600 mb-3"><i data-lucide="users" class="w-4 h-4 inline mr-1"></i>นักศึกษาที่กำลังศึกษา แยกตามเพศ</p>
        <div class="grid grid-cols-1 sm:grid-cols-2 gap-4 mb-4">
          ${genderCard('นักศึกษาชาย', tot.M, tot.M + tot.F + tot.U, '#2563eb', 'M')}
          ${genderCard('นักศึกษาหญิง', tot.F, tot.M + tot.F + tot.U, '#db2777', 'F')}
        </div>
        ${tot.U ? `<p class="text-xs text-amber-600 mb-3"><i data-lucide="alert-triangle" class="w-3 h-3 inline mr-1"></i>มี ${tot.U} คนที่ระบุเพศไม่ได้จากคำนำหน้าชื่อ — แนะนำเพิ่มคอลัมน์ "gender" ในชีต student</p>` : ''}
        <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">${yrGenderRows}</div>
      </div>
    </div>
  </details>`;
}

// ======================== STUDENTS ========================
function studentsPage() {
  const isAdmin = isAdminRole();
  const isClassTeacher = APP.currentRole === 'classTeacher';
  const canEdit = isAdmin || APP.currentRole === 'teacher' || isClassTeacher;
  const allStudents = getDataByType('student');
  const selectedYearLevel = APP.filters._studentYearLevel || '';

  // ClassTeacher: room-based selector
  if (isClassTeacher) {
    const yr = APP.currentUser.responsible_year || '1';
    const myYrStudents = allStudents.filter(s => norm(s.year_level) === norm(yr));
    const selectedRoom = APP.filters._studentRoom || '';
    // จำนวนนักศึกษานับเฉพาะคนที่กำลังศึกษาอยู่
    const myYrActive = activeStudents(myYrStudents);
    const roomACount = myYrActive.filter(s => norm(s.room).toUpperCase() === 'A').length;
    const roomBCount = myYrActive.filter(s => norm(s.room).toUpperCase() === 'B').length;

    let headerHtml = `<div class="flex flex-wrap items-center justify-between gap-3 mb-4">
      <h2 class="text-xl font-bold text-gray-800"><i data-lucide="users" class="w-6 h-6 inline mr-2"></i>ข้อมูลนักศึกษา ชั้นปีที่ ${yr}</h2>
    </div>
    <div class="grid grid-cols-2 sm:grid-cols-3 gap-3 mb-4">
      <div class="bg-white rounded-2xl p-4 border border-blue-100 text-center"><p class="text-sm text-gray-500">ทั้งหมด</p><p class="text-2xl font-bold text-primary">${myYrActive.length} <span class="text-sm font-normal text-gray-500">คน</span></p></div>
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

    let data = myYrStudents.filter(s => !isGraduate(s));
    if (selectedRoom) data = data.filter(s => norm(s.room).toUpperCase() === selectedRoom.toUpperCase());
    data = applyFilters(data);
    const total = data.length;
    const paged = paginate(data);

    return headerHtml + `
    ${filterBar({ semester: false, year: false, yearLevel: false })}
    <p class="text-xs text-gray-400 mb-2"><i data-lucide="info" class="w-3 h-3 inline mr-1"></i>แสดงผู้ที่ลาออก/พักการศึกษา/โอนย้ายไว้ตามชั้นปีเดิมด้วย (ดูป้ายสถานภาพ) โดยไม่นับรวมในจำนวนนักศึกษา รายชั้นปี และอัตราคงอยู่</p>
    <div class="bg-white rounded-2xl border border-blue-100 overflow-hidden">
      <div class="overflow-x-auto"><table class="w-full text-sm">
        <thead><tr class="bg-surface text-left"><th class="px-4 py-3 font-semibold">รหัสนักศึกษา</th><th class="px-4 py-3 font-semibold">ชื่อ-สกุล</th><th class="px-4 py-3 font-semibold">ห้อง</th><th class="px-4 py-3 font-semibold">รุ่นที่</th><th class="px-4 py-3 font-semibold">สถานภาพ</th><th class="px-4 py-3"></th></tr></thead>
        <tbody>${paged.length ? paged.map(s => `<tr class="border-t hover:bg-gray-50">
          <td class="px-4 py-3">${s.student_id || ''}</td><td class="px-4 py-3 font-medium">${studentDisplayName(s)}</td><td class="px-4 py-3">${s.room || '-'}</td><td class="px-4 py-3">${s.batch || ''}</td>
          <td class="px-4 py-3"><span class="px-2 py-1 rounded-full text-xs ${s.status === 'กำลังศึกษา' ? 'bg-green-100 text-green-700' : s.status === 'สำเร็จการศึกษา' ? 'bg-blue-100 text-blue-700' : s.status === 'ลาออก' ? 'bg-red-100 text-red-700' : s.status === 'ขอโอนย้ายสถานศึกษา' ? 'bg-orange-100 text-orange-700' : 'bg-yellow-100 text-yellow-700'}">${s.status || ''}</span></td>
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
          <button onclick="APP.filters._studentYearLevel='__grad';APP.pagination.page=1;renderCurrentPage()" class="px-4 py-2 rounded-xl text-sm font-medium ${selectedYearLevel === '__grad' ? 'bg-primary text-white' : 'bg-gray-100 text-gray-600 hover:bg-gray-200'}">ผู้สำเร็จการศึกษา</button>
        </div>
      </div>
    </div>`;

    let data = myAllStudents;
    if (selectedYearLevel === '__grad') data = data.filter(s => norm(s.status) === 'สำเร็จการศึกษา' || norm(s.year_level) === 'จบ');
    else if (selectedYearLevel) data = data.filter(s => norm(s.year_level) === selectedYearLevel && !isGraduate(s));
    data = applyFilters(data);
    const total = data.length;
    const paged = paginate(data);

    return headerHtml + `
    ${filterBar({ semester: false, year: false, yearLevel: false })}
    <p class="text-xs text-gray-400 mb-2"><i data-lucide="info" class="w-3 h-3 inline mr-1"></i>แสดงผู้ที่ลาออก/พักการศึกษา/โอนย้ายไว้ตามชั้นปีเดิมด้วย (ดูป้ายสถานภาพ) โดยไม่นับรวมในจำนวนนักศึกษา รายชั้นปี และอัตราคงอยู่</p>
    <div class="bg-white rounded-2xl border border-blue-100 overflow-hidden">
      <div class="overflow-x-auto"><table class="w-full text-sm">
        <thead><tr class="bg-surface text-left"><th class="px-4 py-3 font-semibold">รหัสนักศึกษา</th><th class="px-4 py-3 font-semibold">ชื่อ-สกุล</th><th class="px-4 py-3 font-semibold">ชั้นปี</th><th class="px-4 py-3 font-semibold">รุ่นที่</th><th class="px-4 py-3 font-semibold">สถานภาพ</th><th class="px-4 py-3"></th></tr></thead>
        <tbody>${paged.length ? paged.map(s => `<tr class="border-t hover:bg-gray-50">
          <td class="px-4 py-3">${s.student_id || ''}</td><td class="px-4 py-3 font-medium">${studentDisplayName(s)}</td><td class="px-4 py-3">${s.year_level || ''}</td><td class="px-4 py-3">${s.batch || ''}</td>
          <td class="px-4 py-3"><span class="px-2 py-1 rounded-full text-xs ${s.status === 'กำลังศึกษา' ? 'bg-green-100 text-green-700' : s.status === 'สำเร็จการศึกษา' ? 'bg-blue-100 text-blue-700' : s.status === 'ลาออก' ? 'bg-red-100 text-red-700' : s.status === 'ขอโอนย้ายสถานศึกษา' ? 'bg-orange-100 text-orange-700' : 'bg-yellow-100 text-yellow-700'}">${s.status || ''}</span></td>
          <td class="px-4 py-3"><div class="flex gap-1"><button onclick="showStudentDetail('${s.__backendId}')" class="text-gray-400 hover:text-primary" title="ดูข้อมูล"><i data-lucide="eye" class="w-4 h-4"></i></button></div></td></tr>`).join('') : '<tr><td colspan="6" class="px-4 py-8 text-center text-gray-400">ไม่มีข้อมูล</td></tr>'}</tbody>
      </table></div>
    </div>
    ${paginationHTML(total, APP.pagination.perPage, APP.pagination.page, 'changePage')}`;
  }

  let headerHtml = `<div class="flex flex-wrap items-center justify-between gap-3 mb-4">
    <h2 class="text-xl font-bold text-gray-800"><i data-lucide="users" class="w-6 h-6 inline mr-2"></i>ข้อมูลนักศึกษา</h2>
    ${isAdmin ? `<div class="flex gap-2"><button onclick="showAddStudentModal()" class="flex items-center gap-2 px-4 py-2 bg-primary text-white rounded-xl hover:bg-primaryDark text-sm"><i data-lucide="plus" class="w-4 h-4"></i>เพิ่มนักศึกษา</button>${csvUploadBtn('student', 'name,title_prefix,gender,student_id,batch,status,status_date,status_reason,phone,email,parent_name,parent_phone,advisor,year_level,room,national_id,name_en,birth_date,birth_province,nationality,religion,prev_education,degree,honors,admission_date,graduation_date,comprehensive_exam')}</div>` : ''}
  </div>
  ${['admin', 'academic', 'registrar', 'executive'].includes(APP.currentRole) ? studentRetentionAnalyticsHTML() : ''}
  ${isAdmin ? promotePanelHTML(allStudents) : ''}
  <div class="bg-white rounded-2xl p-4 border border-blue-100 mb-4">
    <div class="flex items-center gap-3">
      <label class="text-sm font-medium text-gray-700">ชั้นปี:</label>
      <select onchange="APP.filters._studentYearLevel=this.value;APP.pagination.page=1;renderCurrentPage()" class="border border-gray-200 rounded-xl px-3 py-2 text-sm">
        <option value="">-- เลือกชั้นปี --</option>
        <option value="1" ${selectedYearLevel === '1' ? 'selected' : ''}>ชั้นปี 1</option>
        <option value="2" ${selectedYearLevel === '2' ? 'selected' : ''}>ชั้นปี 2</option>
        <option value="3" ${selectedYearLevel === '3' ? 'selected' : ''}>ชั้นปี 3</option>
        <option value="4" ${selectedYearLevel === '4' ? 'selected' : ''}>ชั้นปี 4</option>
        <option value="__grad" ${selectedYearLevel === '__grad' ? 'selected' : ''}>ผู้สำเร็จการศึกษา</option>
      </select>
      ${selectedYearLevel ? '<span class="text-xs text-gray-500">' + (selectedYearLevel === '__grad' ? 'แสดงผู้สำเร็จการศึกษา' : 'แสดงข้อมูลชั้นปี ' + selectedYearLevel) + '</span>' : ''}
    </div>
  </div>`;

  if (!selectedYearLevel) return headerHtml + noYearSelectedMsg('นักศึกษา (กรุณาเลือกชั้นปี)');

  let data = selectedYearLevel === '__grad'
    ? allStudents.filter(s => norm(s.status) === 'สำเร็จการศึกษา' || norm(s.year_level) === 'จบ')
    : allStudents.filter(s => norm(s.year_level) === selectedYearLevel && !isGraduate(s));
  data = applyFilters(data);
  const total = data.length;
  const paged = paginate(data);

  return headerHtml + `
  ${filterBar({ semester: false, year: false, yearLevel: false })}
  <p class="text-xs text-gray-400 mb-2"><i data-lucide="info" class="w-3 h-3 inline mr-1"></i>แสดงผู้ที่ลาออก/พักการศึกษา/โอนย้ายไว้ตามชั้นปีเดิมด้วย (ดูป้ายสถานภาพ) โดยไม่นับรวมในจำนวนนักศึกษา รายชั้นปี และอัตราคงอยู่</p>
  <div class="bg-white rounded-2xl border border-blue-100 overflow-hidden">
    <div class="overflow-x-auto"><table class="w-full text-sm">
      <thead><tr class="bg-surface text-left"><th class="px-4 py-3 font-semibold">รหัสนักศึกษา</th><th class="px-4 py-3 font-semibold">ชื่อ-สกุล</th><th class="px-4 py-3 font-semibold">ชั้นปี</th><th class="px-4 py-3 font-semibold">รุ่นที่</th><th class="px-4 py-3 font-semibold">สถานภาพ</th><th class="px-4 py-3"></th></tr></thead>
      <tbody>${paged.length ? paged.map(s => `<tr class="border-t hover:bg-gray-50">
        <td class="px-4 py-3">${s.student_id || ''}</td><td class="px-4 py-3 font-medium">${studentDisplayName(s)}</td><td class="px-4 py-3">${s.year_level || ''}</td><td class="px-4 py-3">${s.batch || ''}</td>
        <td class="px-4 py-3"><span class="px-2 py-1 rounded-full text-xs ${s.status === 'กำลังศึกษา' ? 'bg-green-100 text-green-700' : s.status === 'สำเร็จการศึกษา' ? 'bg-blue-100 text-blue-700' : s.status === 'ลาออก' ? 'bg-red-100 text-red-700' : s.status === 'ขอโอนย้ายสถานศึกษา' ? 'bg-orange-100 text-orange-700' : 'bg-yellow-100 text-yellow-700'}">${s.status || ''}</span></td>
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

// ช่องข้อมูลสำหรับใบระเบียนแสดงผลการเรียน (ใช้ทั้งฟอร์มเพิ่มและแก้ไขนักศึกษา)
function transcriptFieldsHTML(s) {
  s = s || {};
  const v = k => String(s[k] == null ? '' : s[k]).replace(/"/g, '&quot;');
  return `<details class="border border-blue-100 rounded-xl bg-blue-50 mt-1">
    <summary class="cursor-pointer px-3 py-2 text-xs font-medium text-blue-800">ข้อมูลสำหรับใบระเบียนแสดงผลการเรียน (Transcript) <span class="font-normal text-blue-500">— ใช้เมื่อสำเร็จการศึกษา</span></summary>
    <div class="p-3 grid grid-cols-1 md:grid-cols-2 gap-3">
      <div><label class="block text-xs text-gray-600 mb-1">ชื่อ-สกุล (อังกฤษ)</label><input name="name_en" value="${v('name_en')}" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="เช่น Miss KANITA ..."></div>
      <div><label class="block text-xs text-gray-600 mb-1">วันเกิด</label><input name="birth_date" type="date" value="${v('birth_date')}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <div><label class="block text-xs text-gray-600 mb-1">จังหวัดที่เกิด</label><input name="birth_province" value="${v('birth_province')}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <div><label class="block text-xs text-gray-600 mb-1">สัญชาติ</label><input name="nationality" value="${v('nationality')}" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="ไทย"></div>
      <div><label class="block text-xs text-gray-600 mb-1">ศาสนา</label><input name="religion" value="${v('religion')}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <div><label class="block text-xs text-gray-600 mb-1">วุฒิการศึกษาเดิม</label><input name="prev_education" value="${v('prev_education')}" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="มัธยมศึกษาปีที่ 6"></div>
      <div><label class="block text-xs text-gray-600 mb-1">วุฒิการศึกษา</label><input name="degree" value="${v('degree')}" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="พยาบาลศาสตรบัณฑิต"></div>
      <div><label class="block text-xs text-gray-600 mb-1">เกียรตินิยม</label><input name="honors" value="${v('honors')}" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="-"></div>
      <div><label class="block text-xs text-gray-600 mb-1">วันที่เข้ารับการศึกษา</label><input name="admission_date" type="date" value="${v('admission_date')}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <div><label class="block text-xs text-gray-600 mb-1">วันที่สำเร็จการศึกษา</label><input name="graduation_date" type="date" value="${v('graduation_date')}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <div><label class="block text-xs text-gray-600 mb-1">การสอบรวบยอด</label><select name="comprehensive_exam" class="w-full border rounded-xl px-3 py-2 text-sm"><option value="">-</option><option ${norm(s.comprehensive_exam) === 'ผ่าน' ? 'selected' : ''}>ผ่าน</option><option ${norm(s.comprehensive_exam) === 'ไม่ผ่าน' ? 'selected' : ''}>ไม่ผ่าน</option></select></div>
    </div>
  </details>`;
}

function showAddStudentModal() {
  showModal('เพิ่มนักศึกษา', `
    <form id="addStudentForm" class="space-y-3">
      ${titlePrefixField('')}
      <div class="grid grid-cols-1 md:grid-cols-2 gap-3">
        <div><label class="block text-xs text-gray-600 mb-1">รหัสนักศึกษา *</label><input name="student_id" required class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">รุ่นที่</label><input name="batch" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="เช่น 36"></div>
        <div><label class="block text-xs text-gray-600 mb-1">เลขบัตรประชาชน</label><input name="national_id" maxlength="13" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">สถานภาพ</label><select name="status" class="w-full border rounded-xl px-3 py-2 text-sm"><option>กำลังศึกษา</option><option>พักการศึกษา</option><option>ลาออก</option><option>ขอโอนย้ายสถานศึกษา</option><option>สำเร็จการศึกษา</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">ชั้นปี</label><select name="year_level" class="w-full border rounded-xl px-3 py-2 text-sm"><option>1</option><option>2</option><option>3</option><option>4</option><option value="จบ">จบ (สำเร็จการศึกษา)</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">ห้อง</label><input name="room" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">เพศ</label><select name="gender" class="w-full border rounded-xl px-3 py-2 text-sm"><option value="">ไม่ระบุ</option><option>ชาย</option><option>หญิง</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">โทรศัพท์</label><input name="phone" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">E-mail</label><input name="email" type="email" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">ชื่อผู้ปกครอง</label><input name="parent_name" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">โทรผู้ปกครอง</label><input name="parent_phone" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">อาจารย์ที่ปรึกษา</label><input name="advisor" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">วันที่ลาออก/พักการศึกษา</label><input name="status_date" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="เช่น 1 มิ.ย. 2569"></div>
        <div class="md:col-span-2"><label class="block text-xs text-gray-600 mb-1">เหตุผลการลาออก/พักการศึกษา</label><input name="status_reason" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="กรอกเมื่อสถานะเป็น ลาออก / พักการศึกษา / โอนย้าย"></div>
      </div>
      ${transcriptFieldsHTML({})}
      <button type="submit" class="w-full mt-3 bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึก</button>
    </form>
  `);
  document.getElementById('addStudentForm').onsubmit = async (e) => {
    e.preventDefault();
    await withLoading(e.target, async () => {
      const fd = new FormData(e.target);
      const obj = { type: 'student', created_at: new Date().toISOString() };
      fd.forEach((v, k) => obj[k] = v);
      obj.name = combineName(e.target); // เก็บ title_prefix ไว้เป็นคอลัมน์แยกด้วย (ถ้าชีตมีคอลัมน์นี้)
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
    ${infoRow('ชื่อ-สกุล', s.name)}${infoRow('รหัสนักศึกษา', s.student_id)}${infoRow('เลขบัตรประชาชน', maskNationalId(s.national_id))}${infoRow('รุ่นที่', s.batch)}
    ${infoRow('สถานภาพ', s.status)}${infoRow('ชั้นปี', s.year_level)}${infoRow('ห้อง', s.room)}${infoRow('เพศ', s.gender ? s.gender : (studentGender(s) === 'M' ? 'ชาย (จากคำนำหน้า)' : studentGender(s) === 'F' ? 'หญิง (จากคำนำหน้า)' : '-'))}
    ${infoRow('โทร', s.phone)}${infoRow('E-mail', s.email)}${infoRow('ผู้ปกครอง', s.parent_name)}${infoRow('โทรผู้ปกครอง', s.parent_phone)}${infoRow('อาจารย์ที่ปรึกษา', s.advisor)}
    ${norm(s.status_date) ? infoRow('วันที่ลาออก/พักการศึกษา', s.status_date) : ''}${norm(s.status_reason) ? infoRow('เหตุผลลาออก/พักการศึกษา', s.status_reason) : ''}
  </div>`);
}

// คำนวณปีการศึกษาปัจจุบัน (พ.ศ.) — ปีการศึกษาไทยเริ่ม มิ.ย.
function currentAcademicYearBE() {
  const d = new Date();
  let yr = d.getFullYear() + 543;
  if (d.getMonth() < 5) yr -= 1; // ก่อนเดือน มิ.ย. ยังนับเป็นปีการศึกษาก่อน
  return String(yr);
}

// ======================== SUBJECTS ========================
// ---- รหัสหน่วยกิตรายวิชา: น(ท-ป-อ) = หน่วยกิตรวม(ทฤษฎี-ปฏิบัติ/ทดลอง-ศึกษาด้วยตนเอง) ----
// หมายเหตุ: ค่าหน่วยกิตรวม (credits) ยังเป็นตัวเลขเหมือนเดิม ใช้คำนวณ GPA ได้ตามปกติ
function creditCode(s) {
  if (!s) return '';
  const c = norm(s.credits);
  const t = norm(s.hours_theory), l = norm(s.hours_lab), se = norm(s.hours_self);
  if (!c) return '';
  if (t === '' && l === '' && se === '') return c;
  return c + '(' + (t || '0') + '-' + (l || '0') + '-' + (se || '0') + ')';
}

function updateCreditPreview(el) {
  const form = el.closest('form'); if (!form) return;
  const g = n => ((form.querySelector('[name="' + n + '"]') || {}).value || '').trim();
  const c = g('credits'), t = g('hours_theory'), l = g('hours_lab'), se = g('hours_self');
  const prev = form.querySelector('.credit-preview');
  if (prev) prev.textContent = !c ? '-' : ((t === '' && l === '' && se === '') ? c : c + '(' + (t || '0') + '-' + (l || '0') + '-' + (se || '0') + ')');
}

// กลุ่มช่องกรอกหน่วยกิต (ใช้ทั้งฟอร์มเพิ่มและแก้ไขรายวิชา)
function creditFields(s) {
  s = s || {};
  const v = k => String(s[k] == null ? '' : s[k]).replace(/"/g, '&quot;');
  return `<div class="col-span-2 p-3 bg-blue-50 rounded-xl border border-blue-100">
    <label class="block text-xs font-medium text-blue-800 mb-2">หน่วยกิต <span class="font-normal text-blue-600">รูปแบบ น(ท-ป-อ) เช่น 2(1-2-3)</span></label>
    <div class="grid grid-cols-2 sm:grid-cols-4 gap-2">
      <div><label class="block text-[11px] text-gray-600 mb-1">หน่วยกิตรวม</label><input name="credits" type="number" min="0" step="0.5" value="${v('credits')}" oninput="updateCreditPreview(this)" class="w-full border rounded-lg px-2 py-1.5 text-sm"></div>
      <div><label class="block text-[11px] text-gray-600 mb-1">ทฤษฎี (ชม./สัปดาห์)</label><input name="hours_theory" type="number" min="0" value="${v('hours_theory')}" oninput="updateCreditPreview(this)" class="w-full border rounded-lg px-2 py-1.5 text-sm"></div>
      <div><label class="block text-[11px] text-gray-600 mb-1">ปฏิบัติ/ทดลอง (ชม./สัปดาห์)</label><input name="hours_lab" type="number" min="0" value="${v('hours_lab')}" oninput="updateCreditPreview(this)" class="w-full border rounded-lg px-2 py-1.5 text-sm"></div>
      <div><label class="block text-[11px] text-gray-600 mb-1">ศึกษาด้วยตนเอง (ชม./สัปดาห์)</label><input name="hours_self" type="number" min="0" value="${v('hours_self')}" oninput="updateCreditPreview(this)" class="w-full border rounded-lg px-2 py-1.5 text-sm"></div>
    </div>
    <p class="text-xs text-blue-700 mt-2">รหัสหน่วยกิต: <span class="credit-preview font-bold font-mono">${creditCode(s) || '-'}</span></p>
  </div>`;
}

// กล่องอธิบายความหมายของรหัสหน่วยกิต (แสดงในหน้ารายวิชา)
function creditInfoBox() {
  return `<details class="bg-blue-50 border border-blue-100 rounded-2xl mb-4">
    <summary class="cursor-pointer px-4 py-3 text-sm font-medium text-blue-800 flex items-center gap-2"><i data-lucide="info" class="w-4 h-4"></i>ความหมายของรหัสหน่วยกิตรายวิชา <span class="text-xs font-normal text-blue-500">(คลิกเพื่อดู)</span></summary>
    <div class="px-4 pb-4 text-sm text-gray-700 space-y-1.5">
      <p>รหัสหน่วยกิตกำหนดเป็นตัวเลขในรูปแบบ <strong class="font-mono">น(ท-ป-อ)</strong> โดย:</p>
      <p>• <strong>ตัวเลขหน้าวงเล็บ</strong> = จำนวนหน่วยกิตรวมของรายวิชา</p>
      <p>• <strong>ตัวเลขแรกในวงเล็บ</strong> = จำนวนชั่วโมงภาคทฤษฎีต่อสัปดาห์</p>
      <p>• <strong>ตัวเลขที่สองในวงเล็บ</strong> = จำนวนชั่วโมงภาคทดลองในห้องปฏิบัติการ หรือภาคปฏิบัติการพยาบาลในคลินิก/ชุมชนต่อสัปดาห์</p>
      <p>• <strong>ตัวเลขที่สามในวงเล็บ</strong> = จำนวนชั่วโมงการศึกษาด้วยตนเองต่อสัปดาห์</p>
      <div class="mt-2 p-3 bg-white rounded-lg border border-blue-100 text-xs space-y-1.5">
        <p><span class="font-mono font-bold">2(1-2-3)</span> = รายวิชามี 2 หน่วยกิต · ภาคทฤษฎี 1 ชม./สัปดาห์ · ภาคทดลองในห้องปฏิบัติการ 2 ชม./สัปดาห์ · ศึกษาด้วยตนเอง 3 ชม./สัปดาห์</p>
        <p><span class="font-mono font-bold">3(0-9-0)</span> = รายวิชาปฏิบัติมี 3 หน่วยกิต · ภาคปฏิบัติ 9 ชม./สัปดาห์</p>
      </div>
    </div>
  </details>`;
}

function subjectsPage() {
  const isAdmin = isAdminRole();
  const canEdit = isAdmin || APP.currentRole === 'teacher' || APP.currentRole === 'classTeacher';
  const isStudent = APP.currentRole === 'student';
  const allSubjects = getDataByType('subject');
  // นักศึกษา: จำกัดปีการศึกษา/ข้อมูลให้เห็นเฉพาะรุ่นของตนเองเท่านั้น
  const pickerSubjects = (isStudent && APP.currentUser.data && norm(APP.currentUser.data.batch))
    ? allSubjects.filter(s => norm(s.batch) === norm(APP.currentUser.data.batch))
    : allSubjects;
  // นักศึกษา: ถ้ายังไม่เลือกปี ให้ default เป็นปีการศึกษาปัจจุบันโดยอัตโนมัติ
  if (isStudent && !APP.filters._pageYear) {
    const cur = currentAcademicYearBE();
    // ตรวจว่ามีข้อมูลปีนั้นไหม ถ้าไม่มี ใช้ปีล่าสุดที่มีข้อมูลแทน
    const years = [...new Set(pickerSubjects.map(s => norm(s.academic_year)).filter(Boolean))].sort();
    APP.filters._pageYear = years.includes(cur) ? cur : (years[years.length - 1] || cur);
  }
  const selectedYear = APP.filters._pageYear || '';

  let headerHtml = `<div class="flex flex-wrap items-center justify-between gap-3 mb-4">
    <h2 class="text-xl font-bold text-gray-800"><i data-lucide="book-open" class="w-6 h-6 inline mr-2"></i>รายวิชาที่เปิดสอน</h2>
    ${isAdmin ? `<div class="flex gap-2"><button onclick="showAddSubjectModal()" class="flex items-center gap-2 px-4 py-2 bg-primary text-white rounded-xl hover:bg-primaryDark text-sm"><i data-lucide="plus" class="w-4 h-4"></i>เพิ่มรายวิชา</button>${csvUploadBtn('subject', 'subject_code,subject_name,coordinator,department,year_level,batch,room,credits,hours_theory,hours_lab,hours_self,semester,academic_year')}</div>` : ''}
  </div>`;
  headerHtml += yearPickerBar(pickerSubjects, 'ปีการศึกษา');
  if (isAdmin) headerHtml += creditInfoBox();

  // แสดงป้ายบอกบริบทสำหรับนักศึกษา (รุ่น/ชั้นปี/ปีการศึกษา)
  if (isStudent && APP.currentUser.data) {
    const stu = APP.currentUser.data;
    headerHtml += `<div class="bg-blue-50 border border-blue-200 rounded-xl p-3 mb-4 text-xs text-blue-800 flex flex-wrap gap-3">
      <span><i data-lucide="user" class="w-3 h-3 inline mr-1"></i><strong>${stu.name || ''}</strong></span>
      ${stu.batch ? `<span>รุ่นที่ <strong>${stu.batch}</strong></span>` : ''}
      <span>ชั้นปี <strong>${stu.year_level || '-'}</strong></span>
      <span>ปีการศึกษา <strong>${selectedYear || '-'}</strong></span>
    </div>`;
  }

  if (!selectedYear) return headerHtml + noYearSelectedMsg('รายวิชา');

  // Filter subjects for the selected year — apply only the filters shown on THIS page
  // (search, semester, yearLevel). Do NOT use applyFilters() because it also checks
  // APP.filters.academicYear and other global filters that may be leftover from other pages.
  let data = allSubjects.filter(s => norm(s.academic_year) === norm(selectedYear));
  if (APP.filters.search) { const s = APP.filters.search.toLowerCase(); data = data.filter(x => Object.values(x).some(v => String(v).toLowerCase().includes(s))); }
  if (APP.filters.semester) data = data.filter(x => normSem(x.semester) === APP.filters.semester);
  if (APP.filters.yearLevel) data = data.filter(x => norm(x.year_level) === APP.filters.yearLevel);
  if (APP.currentRole === 'classTeacher') data = data.filter(s => norm(s.year_level) === norm(APP.currentUser.responsible_year || '1'));
  if (isStudent && APP.currentUser.data) {
    const stuBatch = norm(APP.currentUser.data.batch);
    const stuYearLevel = norm(APP.currentUser.data.year_level);
    // [DIAGNOSTIC] log ข้อมูลก่อนกรอง — เปิด Console (F12) ดูได้
    try {
      const beforeFilter = data.length;
      const allGeRaw = allSubjects.filter(s => (s.subject_code || '').toUpperCase().startsWith('GE'));
      const allBatch80Raw = allSubjects.filter(s => norm(s.batch) === '80');
      console.log('[Subjects] นักศึกษา:', { name: APP.currentUser.data.name, batch: stuBatch, year_level: stuYearLevel });
      console.log('[Subjects] selectedYear =', JSON.stringify(selectedYear), '| ทั้งหมดในชีต =', allSubjects.length, '| ในปีที่เลือก =', beforeFilter);
      console.log('[Subjects] === GE ทุกตัว (จากทั้งชีต) ===');
      console.table(allGeRaw.map(s => ({
        code: s.subject_code, batch: s.batch, year_level: s.year_level, sem: s.semester,
        academic_year_RAW: s.academic_year,
        academic_year_norm: norm(s.academic_year),
        academic_year_type: typeof s.academic_year,
        academic_year_length: (s.academic_year || '').length,
        matches_2568: norm(s.academic_year) === '2568'
      })));
      console.log('[Subjects] === Batch 80 ทุกตัว (จากทั้งชีต) ===');
      console.table(allBatch80Raw.map(s => ({
        code: s.subject_code, year_level: s.year_level, sem: s.semester,
        academic_year_RAW: s.academic_year,
        academic_year_norm: norm(s.academic_year),
        matches_2568: norm(s.academic_year) === '2568'
      })));
    } catch (e) { console.error('diag err', e); }
    // ขั้น 1) กรองรายวิชาให้ตรงกับนักศึกษา
    //   - ถ้ารายวิชาระบุ batch → batch ต้องตรงกับนักศึกษา
    //   - ถ้ารายวิชาไม่ระบุ batch (เช่นวิชา GE ทั่วไป) → ใช้ year_level ตรงเป็นตัวจับคู่
    if (stuBatch) {
      data = data.filter(s => {
        const sBatch = norm(s.batch);
        const sYear = norm(s.year_level);
        if (sBatch) return sBatch === stuBatch;
        // ไม่ระบุ batch → จับคู่ด้วย year_level (ต้องมีค่าและตรงกัน)
        return sYear && sYear === stuYearLevel;
      });
    } else if (stuYearLevel) {
      data = data.filter(s => norm(s.year_level) === stuYearLevel);
    } else {
      data = [];
    }
    try { console.log('[Subjects] รายวิชาหลังกรองด้วยรุ่น/ชั้นปี:', data.length); } catch (e) { }
    // ขั้น 2) ถ้ามีรายวิชาซ้ำ (รหัสวิชา + ภาค + ปี ตรงกัน) → กรองเพิ่มด้วย year_level
    //         เพื่อเลือกเฉพาะรายวิชาของชั้นปีของนักศึกษา
    if (stuYearLevel) {
      const keyCounts = {};
      data.forEach(s => {
        const key = `${norm(s.subject_code)}|${normSem(s.semester)}|${norm(s.academic_year)}`;
        keyCounts[key] = (keyCounts[key] || 0) + 1;
      });
      data = data.filter(s => {
        const key = `${norm(s.subject_code)}|${normSem(s.semester)}|${norm(s.academic_year)}`;
        if ((keyCounts[key] || 0) > 1) {
          return norm(s.year_level) === stuYearLevel;
        }
        return true;
      });
    }
  }

  // ตัวกรองรุ่น (สำหรับ admin/academic) — รวบรวมรุ่นจากรายวิชาในปีที่เลือก
  const batchFilter = APP.filters._subjectBatch || '';
  if (batchFilter) data = data.filter(s => norm(s.batch) === batchFilter);
  const allBatches = [...new Set(allSubjects.filter(s => norm(s.academic_year) === norm(selectedYear)).map(s => norm(s.batch)).filter(Boolean))].sort();
  const batchSelector = (allBatches.length && !isStudent) ? `<div class="flex flex-wrap gap-3 mb-4 items-center">
    <label class="text-sm text-gray-700">รุ่น:</label>
    <select onchange="APP.filters._subjectBatch=this.value;APP.pagination.page=1;renderCurrentPage()" class="border border-gray-200 rounded-xl px-3 py-2 text-sm">
      <option value="">ทุกรุ่น</option>
      ${allBatches.map(b => `<option value="${b}" ${batchFilter === b ? 'selected' : ''}>รุ่นที่ ${b}</option>`).join('')}
    </select>
  </div>` : '';

  const total = data.length; const paged = paginate(data);

  // Build lookup map of completed tracking records (with PDF link) by subject_name|semester|academic_year
  const trackingRecords = getDataByType('tracking');
  const trackingPdfMap = {};
  trackingRecords.forEach(t => {
    if (t.deputy_sign === 'เสร็จสิ้น' && t.file_link) {
      const key = `${norm(t.subject_name)}|${normSem(t.semester)}|${norm(t.academic_year)}`;
      trackingPdfMap[key] = t.file_link;
    }
  });

  return headerHtml + `
  ${filterBar({ semester: true, year: false, yearLevel: !isStudent })}
  ${batchSelector}
  <div class="flex items-center justify-between mb-2 px-1">
    <span class="text-xs text-gray-500"><i data-lucide="list" class="w-3 h-3 inline mr-1"></i>พบรายวิชาทั้งหมด <strong>${total}</strong> รายวิชา</span>
  </div>
  <div class="bg-white rounded-2xl border border-blue-100 overflow-hidden">
    <div class="overflow-x-auto"><table class="w-full text-sm">
      <thead><tr class="bg-surface text-left"><th class="px-4 py-3 font-semibold">รหัสวิชา</th><th class="px-4 py-3 font-semibold">ชื่อรายวิชา</th><th class="px-4 py-3 font-semibold">ผู้ประสานงาน</th><th class="px-4 py-3 font-semibold">ชั้นปี</th><th class="px-4 py-3 font-semibold">รุ่น</th><th class="px-4 py-3 font-semibold">หน่วยกิต</th><th class="px-4 py-3 font-semibold">ภาค/ปี</th><th class="px-4 py-3 font-semibold text-center">ข้อมูลรายวิชา</th>${isAdmin ? '<th class="px-4 py-3"></th>' : ''}</tr></thead>
      <tbody>${paged.length ? paged.map(s => {
    const pdfKey = `${norm(s.subject_name)}|${normSem(s.semester)}|${norm(s.academic_year)}`;
    const pdfLink = trackingPdfMap[pdfKey];
    const eyeCell = pdfLink
      ? `<a href="${pdfLink}" target="_blank" title="ดูข้อมูลรายวิชา" class="inline-flex items-center justify-center w-8 h-8 rounded-lg bg-blue-100 text-blue-600 hover:bg-blue-200 hover:text-blue-800 transition"><i data-lucide="eye" class="w-4 h-4"></i></a>`
      : `<span class="text-xs text-gray-300" title="ยังไม่มีไฟล์ PDF">-</span>`;
    return `<tr class="border-t hover:bg-gray-50">
        <td class="px-4 py-3 font-mono text-primary">${s.subject_code || ''}</td><td class="px-4 py-3 font-medium">${s.subject_name || ''}</td><td class="px-4 py-3">${s.coordinator || ''}</td>
        <td class="px-4 py-3">${s.year_level || ''}</td><td class="px-4 py-3">${s.batch || '-'}</td><td class="px-4 py-3 font-mono">${creditCode(s)}</td>
        <td class="px-4 py-3">${semLabel(s.semester)}/${s.academic_year || ''}</td>
        <td class="px-4 py-3 text-center">${eyeCell}</td>
        ${isAdmin ? `<td class="px-4 py-3"><div class="flex gap-1"><button onclick="showImportGradesFromSubjectModal('${s.__backendId}')" class="text-green-500 hover:text-green-700" title="นำเข้ารายชื่อสร้างผลการเรียน"><i data-lucide="user-plus" class="w-4 h-4"></i></button><button onclick="showEditSubjectModal('${s.__backendId}')" class="text-blue-400 hover:text-blue-600" title="แก้ไข"><i data-lucide="pencil" class="w-4 h-4"></i></button><button onclick="deleteRecord('${s.__backendId}')" class="text-red-400 hover:text-red-600" title="ลบ"><i data-lucide="trash-2" class="w-4 h-4"></i></button></div></td>` : ''}</tr>`;
  }).join('') : `<tr><td colspan="${isAdmin ? 9 : 8}" class="px-4 py-8 text-center text-gray-400">ไม่มีข้อมูล</td></tr>`}</tbody>
    </table></div>
  </div>
  ${paginationHTML(total, APP.pagination.perPage, APP.pagination.page, 'changePage')}`;
}

function showAddSubjectModal() {
  showModal('เพิ่มรายวิชา', `
    <form id="addSubjectForm" class="space-y-3">
      <div><label class="block text-xs text-gray-600 mb-1">รหัสวิชา</label><input name="subject_code" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="เช่น GE 104"></div>
      <div><label class="block text-xs text-gray-600 mb-1">ชื่อรายวิชา *</label><input name="subject_name" required class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <div><label class="block text-xs text-gray-600 mb-1">อาจารย์ผู้ประสานงาน (คั่นด้วย ,)</label><input name="coordinator" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="อ.ก, อ.ข"></div>
      <div><label class="block text-xs text-gray-600 mb-1">สาขาวิชาที่รับผิดชอบ <span class="text-gray-400">(มี 2 สาขา คั่นด้วย ,)</span></label><input name="department" list="subjectDeptList" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="เลือก/พิมพ์ เช่น การพยาบาลผู้ใหญ่, การพยาบาลชุมชน">${deptDatalistHTML('subjectDeptList')}</div>
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-xs text-gray-600 mb-1">ชั้นปี</label><select name="year_level" class="w-full border rounded-xl px-3 py-2 text-sm"><option>1</option><option>2</option><option>3</option><option>4</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">รุ่นที่</label><input name="batch" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="เช่น 28"></div>
        <div><label class="block text-xs text-gray-600 mb-1">ห้อง</label><input name="room" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        ${creditFields({})}
        <div><label class="block text-xs text-gray-600 mb-1">ภาคการศึกษา</label><select name="semester" class="w-full border rounded-xl px-3 py-2 text-sm"><option value="1">1</option><option value="2">2</option><option value="3">ฤดูร้อน</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">ปีการศึกษา</label><input name="academic_year" class="w-full border rounded-xl px-3 py-2 text-sm" value="2568"></div>
      </div>
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

// ======================== IMPORT GRADES FROM SUBJECT ========================
// สำหรับ admin/academic: นำรายวิชาที่เปิดสอน → สร้างรายการผลการเรียนแบบ "ยังไม่ระบุเกรด"
// ให้กับนักศึกษาที่ตรงรุ่น/ชั้นปี เพื่อให้แค่กรอกเกรดทีหลังได้เลย
function showImportGradesFromSubjectModal(subjectId) {
  const role = APP.currentRole;
  if (role !== 'admin' && role !== 'academic') {
    showToast('เฉพาะผู้ดูแลระบบ/งานวิชาการเท่านั้น', 'error');
    return;
  }
  const subj = APP.allData.find(d => d.__backendId === subjectId);
  if (!subj) { showToast('ไม่พบรายวิชา', 'error'); return; }

  // หานักศึกษาที่ตรง batch/year_level
  // - ถ้ารายวิชาระบุ batch → ใช้ batch จับคู่
  // - ถ้าไม่ระบุ batch แต่ระบุ year_level → ใช้ year_level จับคู่
  // - ถ้าไม่ระบุทั้งคู่ → แสดงทั้งหมด (ให้ผู้ใช้เลือกเอง)
  const subjBatch = norm(subj.batch);
  const subjYear = norm(subj.year_level);
  let students = getDataByType('student');
  if (subjBatch) {
    students = students.filter(s => norm(s.batch) === subjBatch);
  } else if (subjYear) {
    students = students.filter(s => norm(s.year_level) === subjYear);
  }
  students.sort((a, b) => norm(a.student_id).localeCompare(norm(b.student_id)));

  // หาว่านักศึกษาคนไหนมีผลการเรียนสำหรับวิชานี้ใน semester/year นี้แล้วบ้าง
  const semKey = normSem(subj.semester);
  const yearKey = norm(subj.academic_year);
  const codeKey = norm(subj.subject_code);
  const nameKey = norm(subj.subject_name);
  const existingGrades = getDataByType('grade').filter(g =>
    normSem(g.semester) === semKey &&
    norm(g.academic_year) === yearKey &&
    ((codeKey && norm(g.subject_code) === codeKey) || (!codeKey && norm(g.subject_name) === nameKey))
  );
  const existingByStudent = new Set(existingGrades.map(g => norm(g.student_id)));

  const subjectInfo = `
    <div class="bg-blue-50 border border-blue-200 rounded-xl p-3 mb-3 text-sm">
      <div class="grid grid-cols-2 gap-2">
        <div><span class="text-gray-600">รหัสวิชา:</span> <strong class="font-mono text-primary">${subj.subject_code || '-'}</strong></div>
        <div><span class="text-gray-600">หน่วยกิต:</span> <strong class="font-mono">${creditCode(subj) || '-'}</strong></div>
        <div class="col-span-2"><span class="text-gray-600">ชื่อรายวิชา:</span> <strong>${subj.subject_name || '-'}</strong></div>
        <div><span class="text-gray-600">ภาค/ปี:</span> <strong>${semLabel(subj.semester)}/${subj.academic_year || ''}</strong></div>
        <div><span class="text-gray-600">รุ่น/ชั้นปี:</span> <strong>${subj.batch || '-'} / ${subj.year_level || '-'}</strong></div>
      </div>
    </div>`;

  if (students.length === 0) {
    showModal('นำเข้ารายชื่อสร้างผลการเรียน', `${subjectInfo}
      <div class="text-center text-gray-500 py-6">
        <i data-lucide="user-x" class="w-10 h-10 mx-auto mb-2 text-gray-300"></i>
        <p>ไม่พบนักศึกษาที่ตรงกับรุ่น/ชั้นปีของรายวิชานี้</p>
        <p class="text-xs text-gray-400 mt-1">โปรดตรวจสอบ "รุ่น (batch)" หรือ "ชั้นปี (year_level)" ของนักศึกษาและรายวิชา</p>
      </div>
      <button onclick="closeModal()" class="w-full mt-2 py-2.5 rounded-xl border hover:bg-gray-50">ปิด</button>
    `, null, 'max-w-2xl');
    return;
  }

  const rows = students.map(s => {
    const sid = norm(s.student_id);
    const hasGrade = existingByStudent.has(sid);
    return `<tr class="border-t hover:bg-gray-50 ${hasGrade ? 'bg-gray-50' : ''}">
      <td class="px-2 py-2 text-center">
        <input type="checkbox" class="importGradeStu" value="${sid}" ${hasGrade ? 'disabled' : 'checked data-selectable="1"'}>
      </td>
      <td class="px-2 py-2 font-mono text-xs">${sid}</td>
      <td class="px-2 py-2 text-xs">${s.name || ''}</td>
      <td class="px-2 py-2 text-xs text-center">${hasGrade
        ? '<span class="text-amber-600 inline-flex items-center gap-1"><i data-lucide="check-circle" class="w-3 h-3"></i>มีแล้ว</span>'
        : '<span class="text-green-600">พร้อมนำเข้า</span>'}</td>
    </tr>`;
  }).join('');

  const selectableCount = students.length - existingGrades.length;

  showModal('นำเข้ารายชื่อสร้างผลการเรียน', `
    ${subjectInfo}
    <div class="flex flex-wrap items-center justify-between gap-2 mb-2">
      <div class="text-xs text-gray-600">
        พบนักศึกษา <strong>${students.length}</strong> คน
        ${existingGrades.length ? `<span class="ml-2 text-amber-600">(มีผลการเรียนอยู่แล้ว ${existingGrades.length} คน — จะถูกข้าม)</span>` : ''}
      </div>
      <div class="flex gap-2 text-xs">
        <button type="button" onclick="document.querySelectorAll('.importGradeStu:not([disabled])').forEach(c=>c.checked=true)" class="px-2 py-1 border rounded hover:bg-gray-50">เลือกทั้งหมด</button>
        <button type="button" onclick="document.querySelectorAll('.importGradeStu:not([disabled])').forEach(c=>c.checked=false)" class="px-2 py-1 border rounded hover:bg-gray-50">ล้างเลือก</button>
      </div>
    </div>
    <div class="border rounded-xl overflow-hidden mb-3 max-h-[300px] overflow-y-auto">
      <table class="w-full text-sm">
        <thead class="bg-surface sticky top-0"><tr>
          <th class="px-2 py-2 w-10"></th>
          <th class="px-2 py-2 text-left text-xs font-semibold">รหัสนักศึกษา</th>
          <th class="px-2 py-2 text-left text-xs font-semibold">ชื่อ-สกุล</th>
          <th class="px-2 py-2 text-center text-xs font-semibold">สถานะ</th>
        </tr></thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
    <div class="bg-amber-50 border border-amber-200 rounded-xl p-2.5 text-xs text-amber-800 mb-3">
      <i data-lucide="info" class="w-3 h-3 inline mr-1"></i>
      ระบบจะสร้างรายการผลการเรียนแบบ "ยังไม่ระบุเกรด" สำหรับนักศึกษาที่เลือก — กรอกเกรดในภายหลังที่หน้า "ผลการเรียน"
    </div>
    <div class="flex gap-2">
      <button onclick="closeModal()" class="flex-1 py-2.5 rounded-xl border hover:bg-gray-50">ยกเลิก</button>
      <button id="importGradeConfirmBtn" onclick="runImportGradesFromSubject('${subjectId}')" class="flex-1 py-2.5 rounded-xl bg-primary text-white hover:bg-primaryDark disabled:opacity-50" ${selectableCount === 0 ? 'disabled' : ''}>
        <i data-lucide="user-plus" class="w-4 h-4 inline mr-1"></i>สร้างรายการผลการเรียน
      </button>
    </div>
  `, null, 'max-w-2xl');
}

async function runImportGradesFromSubject(subjectId) {
  const subj = APP.allData.find(d => d.__backendId === subjectId);
  if (!subj) { showToast('ไม่พบรายวิชา', 'error'); return; }
  const selected = Array.from(document.querySelectorAll('.importGradeStu:checked:not([disabled])')).map(c => c.value);
  if (selected.length === 0) { showToast('กรุณาเลือกอย่างน้อย 1 คน', 'error'); return; }

  const btn = document.getElementById('importGradeConfirmBtn');
  if (btn) { btn.disabled = true; btn.innerHTML = '<img src="https://cdn.jsdelivr.net/gh/JOB-BCNB-P/picture/cat_run_transparent.gif" class="cat-run-inline" alt="">กำลังสร้าง...'; lucide.createIcons(); }

  showLoading(`กำลังนำเข้า 0/${selected.length}...`);

  let success = 0, failed = 0;
  for (let i = 0; i < selected.length; i++) {
    const sid = selected[i];
    const lblEl = document.getElementById('loadingMsg');
    if (lblEl) lblEl.textContent = `กำลังนำเข้า ${i + 1}/${selected.length}...`;
    const obj = {
      type: 'grade',
      student_id: sid,
      subject_code: subj.subject_code || '',
      subject_name: subj.subject_name || '',
      grade: '',
      credits: Number(subj.credits) || 3,
      semester: normSem(subj.semester) || '1',
      academic_year: norm(subj.academic_year) || '',
      created_at: new Date().toISOString()
    };
    try {
      const r = await GSheetDB.create(obj);
      if (r && r.isOk) success++; else failed++;
    } catch (e) {
      failed++;
    }
  }

  hideLoading();
  closeModal();
  if (failed === 0) {
    showToast(`สร้างผลการเรียนสำเร็จ ${success} รายการ`);
  } else {
    showToast(`สำเร็จ ${success} รายการ | ผิดพลาด ${failed}`, 'error');
  }
}

// ======================== SCHEDULE (ปฏิทินกิจกรรมวิชาการ) ========================
function scheduleTypeBadge(type) {
  const t = (type || '').trim();
  if (!t) return '<span class="px-2 py-1 rounded-full text-xs bg-gray-100 text-gray-600">ไม่ระบุ</span>';
  if (t.includes('สอบ')) return `<span class="px-2 py-1 rounded-full text-xs bg-red-100 text-red-700">${t}</span>`;
  if (t === 'วันหยุด') return `<span class="px-2 py-1 rounded-full text-xs bg-green-100 text-green-700">${t}</span>`;
  if (t === 'กิจกรรม') return `<span class="px-2 py-1 rounded-full text-xs bg-purple-100 text-purple-700">${t}</span>`;
  return `<span class="px-2 py-1 rounded-full text-xs bg-blue-100 text-blue-700">${t}</span>`;
}

function schedulePage() {
  const canManage = APP.currentRole === 'admin' || APP.currentRole === 'academic' || APP.currentRole === 'executive' || APP.currentRole === 'registrar';
  const isStudent = APP.currentRole === 'student';
  const _todayISO = new Date().toISOString().slice(0, 10);
  const allSchedule = filterScheduleForStudent(getDataByType('schedule'))
    .filter(s => { const d = norm(s.schedule_date); return !d || d >= _todayISO; })
    .sort((a, b) => (a.schedule_date || '').localeCompare(b.schedule_date || ''));
  const total = allSchedule.length; const paged = paginate(allSchedule);

  return `<div class="flex flex-wrap items-center justify-between gap-3 mb-4">
    <h2 class="text-xl font-bold text-gray-800"><i data-lucide="calendar" class="w-6 h-6 inline mr-2"></i>ปฏิทินกิจกรรมวิชาการ</h2>
    ${canManage ? `<div class="flex gap-2">
      <button onclick="showAddScheduleModal()" class="flex items-center gap-2 px-4 py-2 bg-primary text-white rounded-xl hover:bg-primaryDark text-sm"><i data-lucide="plus" class="w-4 h-4"></i>เพิ่มรายการ</button>
      ${csvUploadBtn('schedule', 'subject_name,schedule_date,schedule_time,schedule_time_end,schedule_type,room,year_level,exam_round,student_count,proctor,proctor_change_date,exam_split,room2,student_count2,proctor2')}
    </div>` : ''}
  </div>
  <div class="bg-white rounded-2xl border border-blue-100 p-5">
    <div id="scheduleCalendar"></div>
  </div>
  ${isStudent ? '' : `<div class="mt-4 bg-white rounded-2xl border border-blue-100 overflow-hidden">
    <div class="overflow-x-auto"><table class="w-full text-sm">
      <thead><tr class="bg-surface text-left"><th class="px-4 py-3 font-semibold">วันที่</th><th class="px-4 py-3 font-semibold">เวลา</th><th class="px-4 py-3 font-semibold">รายวิชา/กิจกรรม</th><th class="px-4 py-3 font-semibold">ประเภท</th><th class="px-4 py-3 font-semibold">ห้อง</th>${canManage ? '<th class="px-4 py-3"></th>' : ''}</tr></thead>
      <tbody>${paged.length ? paged.map(s => `<tr class="border-t hover:bg-gray-50">
        <td class="px-4 py-3">${toBuddhistDate(s.schedule_date) || s.schedule_date || ''}</td><td class="px-4 py-3 whitespace-nowrap">${schedTimeRange(s)}</td>
        <td class="px-4 py-3">${(s.subject_name || '').replace(/,\s*/g, '<br>')}${s.exam_round ? ` <span class="text-xs text-gray-400">(ครั้งที่ ${s.exam_round})</span>` : ''}${s.proctor ? `<div class="text-xs text-gray-400">ผู้คุมสอบ: ${s.proctor}</div>` : ''}</td>
        <td class="px-4 py-3">${scheduleTypeBadge(s.schedule_type)}</td>
        <td class="px-4 py-3">${s.room || ''}</td>
        ${canManage ? `<td class="px-4 py-3"><div class="flex gap-1"><button onclick="showEditScheduleModal('${s.__backendId}')" class="text-blue-400 hover:text-blue-600" title="แก้ไข"><i data-lucide="pencil" class="w-4 h-4"></i></button><button onclick="deleteRecord('${s.__backendId}')" class="text-red-400 hover:text-red-600" title="ลบ"><i data-lucide="trash-2" class="w-4 h-4"></i></button></div></td>` : ''}
      </tr>`).join('') : `<tr><td colspan="${canManage ? 6 : 5}" class="px-4 py-8 text-center text-gray-400">ไม่มีข้อมูล</td></tr>`}</tbody>
    </table></div>
  </div>
  ${paginationHTML(total, APP.pagination.perPage, APP.pagination.page, 'changePage')}`}`;
}

function scheduleTypeInput(name, selectedValue) {
  const existing = [...new Set(getDataByType('schedule').map(s => (s.schedule_type || '').trim()).filter(Boolean))];
  const defaults = ['สอบกลางภาค', 'สอบซ่อมกลางภาค', 'สอบปลายภาค', 'สอบซ่อมปลายภาค', 'สอบย่อย', 'สอบภาษาอังกฤษสบช.', 'สอบ OSCE', 'สอบรวบยอด', 'สอบซ่อมรวบยอด', 'กิจกรรม', 'วันหยุด'];
  const allTypes = [...new Set([...defaults, ...existing])];
  const listId = 'scheduleTypeList_' + Date.now();
  return `<input name="${name}" list="${listId}" value="${selectedValue || ''}" oninput="if(window.onScheduleTypeChange)onScheduleTypeChange(this)" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="เช่น สอบกลางภาค, กิจกรรม">
    <datalist id="${listId}">${allTypes.map(t => `<option value="${t}">`).join('')}</datalist>`;
}

// ---- ปฏิทินกิจกรรมวิชาการ: ฟอร์ม + ฟิลด์การสอบแบบไดนามิก + เลือกหลายวิชา ----
// ตัวเลือกรายวิชา (จากชีต subject) กรองตามชั้นปีที่เลือก — ไม่ซ้ำชื่อวิชา
function schedSubjectOptionsHTML(yearLevel) {
  let subs = getDataByType('subject');
  const y = norm(yearLevel);
  if (y) subs = subs.filter(s => norm(s.year_level) === y);
  const seen = new Set(); const out = [];
  subs.forEach(s => {
    const name = norm(s.subject_name); if (!name || seen.has(name)) return;
    seen.add(name);
    out.push({ name, label: s.subject_code ? `${norm(s.subject_code)} ${name}` : name });
  });
  out.sort((a, b) => a.label.localeCompare(b.label, 'th'));
  if (!out.length) return '<option value="">— ไม่มีรายวิชา —</option>';
  return '<option value="">— เลือกรายวิชา —</option>' + out.map(o => `<option value="${o.name.replace(/"/g, '&quot;')}">${o.label.replace(/</g, '&lt;')}</option>`).join('');
}

// datalist ชื่ออาจารย์ (จากชีต teacher + ทำเนียบ) สำหรับช่องผู้คุมสอบ
function proctorDatalistHTML() {
  const names = [...new Set([...getDataByType('teacher').map(t => norm(t.name)), ...getDataByType('teacher_directory').map(t => norm(t.name))].filter(Boolean))].sort((a, b) => a.localeCompare(b, 'th'));
  return `<datalist id="proctorList">${names.map(n => `<option value="${n.replace(/"/g, '&quot;')}">`).join('')}</datalist>`;
}

function refreshSchedSubjectOptions() {
  const yEl = document.getElementById('schedYearLevel');
  const sel = document.getElementById('schedSubjectSelect');
  if (!yEl || !sel) return;
  sel.innerHTML = schedSubjectOptionsHTML(yEl.value);
}
function getSchedSubjectsArr() {
  const v = ((document.getElementById('schedSubjectMultiValue') || {}).value || '');
  return v.split(',').map(s => s.trim()).filter(Boolean);
}
function setSchedSubjectsArr(arr) {
  const u = [...new Set(arr.map(s => s.trim()).filter(Boolean))];
  const h = document.getElementById('schedSubjectMultiValue'); if (h) h.value = u.join(', ');
  renderSchedSubjectChips();
}
function renderSchedSubjectChips() {
  const box = document.getElementById('schedSubjectChips'); if (!box) return;
  const arr = getSchedSubjectsArr();
  box.innerHTML = arr.length
    ? arr.map((s, i) => `<span class="inline-flex items-center gap-1 bg-red-50 text-red-700 border border-red-100 rounded-lg px-2 py-1 text-xs">${String(s).replace(/</g, '&lt;')}<button type="button" onclick="removeSchedSubject(${i})" class="text-red-400 hover:text-red-600 font-bold leading-none">×</button></span>`).join('')
    : '<span class="text-xs text-gray-400">ยังไม่ได้เลือกรายวิชา</span>';
}
function addSchedSubject() {
  const sel = document.getElementById('schedSubjectSelect'); if (!sel || !sel.value) return;
  const a = getSchedSubjectsArr(); a.push(sel.value); setSchedSubjectsArr(a); sel.value = '';
}
function removeSchedSubject(i) {
  const a = getSchedSubjectsArr(); a.splice(i, 1); setSchedSubjectsArr(a);
}
// อาจารย์ผู้คุมสอบ — เลือกได้หลายคน ต่อชุด/ห้อง (idx 1 = ห้องสอบ 1 → proctor, idx 2 = ห้องสอบ 2 → proctor2)
function getSchedProctorArr(idx) {
  const v = ((document.getElementById('schedProctorValue' + idx) || {}).value || '');
  return v.split(',').map(s => s.trim()).filter(Boolean);
}
function setSchedProctorArr(idx, arr) {
  const u = [...new Set(arr.map(s => s.trim()).filter(Boolean))];
  const h = document.getElementById('schedProctorValue' + idx); if (h) h.value = u.join(', ');
  renderSchedProctorChips(idx);
}
function renderSchedProctorChips(idx) {
  const box = document.getElementById('schedProctorChips' + idx); if (!box) return;
  const arr = getSchedProctorArr(idx);
  box.innerHTML = arr.length
    ? arr.map((s, i) => `<span class="inline-flex items-center gap-1 bg-amber-50 text-amber-700 border border-amber-100 rounded-lg px-2 py-1 text-xs">${String(s).replace(/</g, '&lt;')}<button type="button" onclick="removeSchedProctor(${idx},${i})" class="text-amber-400 hover:text-red-600 font-bold leading-none">×</button></span>`).join('')
    : '<span class="text-xs text-gray-400">ยังไม่ได้เลือกผู้คุมสอบ</span>';
}
function addSchedProctor(idx) {
  const el = document.getElementById('schedProctorSelect' + idx); if (!el || !el.value.trim()) return;
  const a = getSchedProctorArr(idx); a.push(el.value.trim()); setSchedProctorArr(idx, a); el.value = '';
}
function removeSchedProctor(idx, i) {
  const a = getSchedProctorArr(idx); a.splice(i, 1); setSchedProctorArr(idx, a);
}
// HTML ตัวเลือกผู้คุมสอบ 1 ชุด (idx, ชื่อฟิลด์, ค่าเดิม, เปิดใช้งานไหม)
function schedProctorPickerHTML(idx, name, value, enabled) {
  return `<input type="hidden" name="${name}" id="schedProctorValue${idx}" value="${value}" ${enabled ? '' : 'disabled'}>
    <div class="flex gap-2 items-stretch">
      <input id="schedProctorSelect${idx}" list="proctorList" class="flex-1 min-w-0 border rounded-xl px-3 py-2 text-sm" placeholder="พิมพ์หรือเลือกชื่ออาจารย์" onkeydown="if(event.key==='Enter'){event.preventDefault();addSchedProctor(${idx});}">
      <button type="button" onclick="addSchedProctor(${idx})" class="shrink-0 px-3 py-2 bg-amber-600 text-white rounded-xl text-sm hover:bg-amber-700 whitespace-nowrap flex items-center gap-1"><i data-lucide="plus" class="w-4 h-4"></i>เพิ่ม</button>
    </div>
    <div id="schedProctorChips${idx}" class="flex flex-wrap gap-1.5 mt-2"></div>`;
}
function schedIsExam() {
  const el = document.querySelector('[name="schedule_type"]');
  return ((el && el.value) || '').includes('สอบ');
}
// สลับสถานะห้องสอบที่ 2 ตามการติ๊ก "แบ่งห้องสอบเป็น 2 ห้อง"
function updateSchedSplitState() {
  const isExam = schedIsExam();
  const sp = document.getElementById('schedExamSplit');
  const on = isExam && sp && sp.checked;
  const blk = document.getElementById('schedRoom2Block');
  if (blk) blk.classList.toggle('hidden', !on);
  const hv = document.getElementById('schedExamSplitValue'); if (hv) { hv.value = on ? '✓' : ''; hv.disabled = !isExam; }
  const r2 = document.querySelector('[name="room2"]'); if (r2) r2.disabled = !on;
  const sc2 = document.querySelector('[name="student_count2"]'); if (sc2) sc2.disabled = !on;
  const pv2 = document.getElementById('schedProctorValue2'); if (pv2) pv2.disabled = !on;
  if (on) renderSchedProctorChips(2);
}
// สลับการแสดงฟิลด์เมื่อประเภทเป็น/ไม่เป็น "การสอบ"
function onScheduleTypeChange(el) {
  const isExam = ((el && el.value) || '').includes('สอบ');
  const ex = document.getElementById('schedExamFields');
  const sw = document.getElementById('schedSubjectSingleWrap');
  const mw = document.getElementById('schedSubjectMultiWrap');
  const si = document.getElementById('schedSubjectSingle');
  const mv = document.getElementById('schedSubjectMultiValue');
  const er = document.getElementById('schedExamRound');
  const pv = document.getElementById('schedProctorValue1');
  const pcd = document.getElementById('schedProctorChangeDate');
  const sc1 = document.querySelector('[name="student_count"]');
  const sp = document.getElementById('schedExamSplit');
  if (ex) ex.classList.toggle('hidden', !isExam);
  if (sw) sw.classList.toggle('hidden', isExam);
  if (mw) mw.classList.toggle('hidden', !isExam);
  if (si) si.disabled = isExam;
  if (mv) mv.disabled = !isExam;
  if (er) er.disabled = !isExam;
  if (pv) pv.disabled = !isExam;
  if (pcd) pcd.disabled = !isExam;
  if (sc1) sc1.disabled = !isExam;
  if (sp) sp.disabled = !isExam;
  const rh = document.getElementById('schedRoomHint'); if (rh) rh.textContent = isExam ? '(ห้องสอบที่ 1)' : '';
  updateSchedSplitState();
  if (isExam) { renderSchedSubjectChips(); renderSchedProctorChips(1); if (window.lucide) lucide.createIcons(); }
  // เมื่อเพิ่งเปลี่ยนประเภทเป็น "การสอบ" → เปิดแจ้งเตือน + เลือกบทบาทเริ่มต้นให้ (ผู้ใช้ปรับเองได้)
  const wasExam = window._schedWasExam === true;
  window._schedWasExam = isExam;
  if (isExam && !wasExam) {
    const nt = document.getElementById('schedNotify');
    if (nt && !nt.checked) {
      nt.checked = true;
      const cbs = document.querySelectorAll('.ann-role-cb');
      const anyChecked = Array.prototype.some.call(cbs, c => c.checked);
      if (!anyChecked) ['student', 'teacher', 'classTeacher'].forEach(r => { const c = document.querySelector('.ann-role-cb[value="' + r + '"]'); if (c) c.checked = true; });
      toggleSchedNotify();
    }
  }
}

// จัดสอบหลายวัน — เพิ่มวันที่หลายวันในฟอร์มเพิ่มรายการ
function getSchedExtraDatesArr() {
  const v = ((document.getElementById('schedExtraDates') || {}).value || '');
  return v.split(',').map(x => x.trim()).filter(Boolean);
}
function setSchedExtraDatesArr(arr) {
  const u = [...new Set(arr.filter(Boolean))];
  const h = document.getElementById('schedExtraDates'); if (h) h.value = u.join(',');
  renderSchedExtraDateChips();
}
function renderSchedExtraDateChips() {
  const box = document.getElementById('schedExtraDateChips'); if (!box) return;
  const arr = getSchedExtraDatesArr();
  box.innerHTML = arr.length
    ? arr.map((d, i) => `<span class="inline-flex items-center gap-1 bg-blue-100 text-blue-700 border border-blue-200 rounded-lg px-2 py-1 text-xs">${(typeof toBuddhistDate === 'function' && toBuddhistDate(d)) || d}<button type="button" onclick="removeSchedExtraDate(${i})" class="text-blue-400 hover:text-red-600 font-bold leading-none">×</button></span>`).join('')
    : '<span class="text-xs text-gray-400">ยังไม่ได้เพิ่มวันเพิ่มเติม</span>';
}
function addSchedExtraDate() {
  const el = document.getElementById('schedExtraDateInput'); if (!el || !el.value) return;
  const a = getSchedExtraDatesArr(); a.push(el.value); setSchedExtraDatesArr(a); el.value = '';
}
function removeSchedExtraDate(i) {
  const a = getSchedExtraDatesArr(); a.splice(i, 1); setSchedExtraDatesArr(a);
}

function scheduleFormBody(s, isNew) {
  s = s || {};
  const v = k => String(s[k] == null ? '' : s[k]).replace(/"/g, '&quot;');
  const isExam = norm(s.schedule_type).includes('สอบ');
  const subjVal = v('subject_name');
  const split = norm(s.exam_split) !== '';
  const notifyDefault = !!(isNew && isExam);
  const notifyRolesDefault = isExam ? 'student,teacher,classTeacher' : '';
  return `
    <div><label class="block text-xs text-gray-600 mb-1">ประเภท * <span class="font-normal text-gray-400">(เลือกก่อน)</span></label>${scheduleTypeInput('schedule_type', s.schedule_type || '')}</div>
    <div id="schedSubjectSingleWrap" class="${isExam ? 'hidden' : ''}">
      <label class="block text-xs text-gray-600 mb-1">รายวิชา/กิจกรรม *</label>
      <input name="subject_name" id="schedSubjectSingle" value="${subjVal}" ${isExam ? 'disabled' : ''} class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="ชื่อรายวิชาหรือกิจกรรม">
    </div>
    <div id="schedSubjectMultiWrap" class="${isExam ? '' : 'hidden'}">
      <label class="block text-xs text-gray-600 mb-1">รายวิชาที่สอบ * <span class="font-normal text-gray-400">(เลือกได้หลายวิชา)</span></label>
      <input type="hidden" name="subject_name" id="schedSubjectMultiValue" value="${subjVal}" ${isExam ? '' : 'disabled'}>
      <div class="flex gap-2 items-stretch">
        <select id="schedSubjectSelect" class="flex-1 min-w-0 border rounded-xl px-3 py-2 text-sm">${schedSubjectOptionsHTML(s.year_level)}</select>
        <button type="button" onclick="addSchedSubject()" class="shrink-0 px-3 py-2 bg-primary text-white rounded-xl text-sm hover:bg-primaryDark whitespace-nowrap flex items-center gap-1"><i data-lucide="plus" class="w-4 h-4"></i>เพิ่ม</button>
      </div>
      <div id="schedSubjectChips" class="flex flex-wrap gap-1.5 mt-2"></div>
    </div>
    <div class="grid grid-cols-2 gap-3">
      <div><label class="block text-xs text-gray-600 mb-1">วันที่ *</label><input name="schedule_date" type="date" required value="${v('schedule_date')}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <div><label class="block text-xs text-gray-600 mb-1">ชั้นปี</label><select name="year_level" id="schedYearLevel" onchange="refreshSchedSubjectOptions()" class="w-full border rounded-xl px-3 py-2 text-sm"><option value="">ทุกชั้นปี</option>${['1', '2', '3', '4'].map(y => `<option ${norm(s.year_level) === y ? 'selected' : ''}>${y}</option>`).join('')}</select></div>
      <div><label class="block text-xs text-gray-600 mb-1">เวลา (เริ่ม)</label><input name="schedule_time" type="time" value="${fmtSchedTime(s.schedule_time)}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <div><label class="block text-xs text-gray-600 mb-1">ถึงเวลา</label><input name="schedule_time_end" type="time" value="${fmtSchedTime(s.schedule_time_end)}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <div class="col-span-2"><label class="block text-xs text-gray-600 mb-1">ห้อง <span class="font-normal text-gray-400" id="schedRoomHint">${isExam ? '(ห้องสอบที่ 1)' : ''}</span></label><input name="room" value="${v('room')}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
    </div>
    ${isNew ? `<div class="p-3 bg-blue-50 rounded-xl border border-blue-100">
      <label class="block text-xs text-gray-600 mb-1">จัดหลายวัน (ถ้ามี) <span class="font-normal text-gray-400">— สร้างรายการเหมือนกันในทุกวันที่เพิ่ม</span></label>
      <div class="flex gap-2 items-stretch">
        <input type="date" id="schedExtraDateInput" class="flex-1 min-w-0 border rounded-xl px-3 py-2 text-sm">
        <button type="button" onclick="addSchedExtraDate()" class="shrink-0 px-3 py-2 bg-primary text-white rounded-xl text-sm hover:bg-primaryDark whitespace-nowrap flex items-center gap-1"><i data-lucide="plus" class="w-4 h-4"></i>เพิ่มวันที่</button>
      </div>
      <input type="hidden" id="schedExtraDates" value="">
      <div id="schedExtraDateChips" class="flex flex-wrap gap-1.5 mt-2"></div>
    </div>` : ''}
    ${proctorDatalistHTML()}
    <div id="schedExamFields" class="${isExam ? '' : 'hidden'} space-y-3 p-3 bg-red-50 rounded-xl border border-red-100">
      <div class="text-xs font-semibold text-red-700"><i data-lucide="clipboard-list" class="w-3.5 h-3.5 inline"></i> ข้อมูลการสอบ</div>
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-xs text-gray-600 mb-1">ครั้งที่</label><input name="exam_round" id="schedExamRound" value="${v('exam_round')}" ${isExam ? '' : 'disabled'} class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="เช่น 1, 2"></div>
        <div><label class="block text-xs text-gray-600 mb-1">จำนวนนักศึกษา (ห้องสอบ 1)</label><input name="student_count" type="number" min="0" value="${v('student_count')}" ${isExam ? '' : 'disabled'} class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="เช่น 40"></div>
      </div>
      <div>
        <label class="block text-xs text-gray-600 mb-1">อาจารย์ผู้คุมสอบ (ห้องสอบ 1) <span class="font-normal text-gray-400">(เลือกได้หลายคน)</span></label>
        ${schedProctorPickerHTML(1, 'proctor', v('proctor'), isExam)}
      </div>
      <div><label class="block text-xs text-gray-600 mb-1">วันที่เปลี่ยนผู้คุมสอบ <span class="font-normal text-gray-400">(กรอกเมื่อมีการขอเปลี่ยน)</span></label><input name="proctor_change_date" id="schedProctorChangeDate" type="date" value="${v('proctor_change_date')}" ${isExam ? '' : 'disabled'} class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <label class="flex items-center gap-2 bg-white rounded-xl px-3 py-2 cursor-pointer border border-red-100"><input type="checkbox" id="schedExamSplit" ${split ? 'checked' : ''} ${isExam ? '' : 'disabled'} onchange="updateSchedSplitState()" class="w-4 h-4"><span class="text-sm text-red-700">🏫 แบ่งห้องสอบเป็น 2 ห้อง (ผู้คุมสอบ 2 ชุด)</span></label>
      <input type="hidden" name="exam_split" id="schedExamSplitValue" value="${split ? '✓' : ''}" ${isExam ? '' : 'disabled'}>
      <div id="schedRoom2Block" class="${isExam && split ? '' : 'hidden'} space-y-3 p-3 bg-white rounded-xl border border-red-100">
        <div class="text-xs font-semibold text-red-700">ห้องสอบที่ 2</div>
        <div class="grid grid-cols-2 gap-3">
          <div><label class="block text-xs text-gray-600 mb-1">ห้องสอบที่ 2</label><input name="room2" value="${v('room2')}" ${isExam && split ? '' : 'disabled'} class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="เช่น ห้อง 202"></div>
          <div><label class="block text-xs text-gray-600 mb-1">จำนวนนักศึกษา (ห้องสอบ 2)</label><input name="student_count2" type="number" min="0" value="${v('student_count2')}" ${isExam && split ? '' : 'disabled'} class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="เช่น 40"></div>
        </div>
        <div>
          <label class="block text-xs text-gray-600 mb-1">อาจารย์ผู้คุมสอบ (ห้องสอบ 2) <span class="font-normal text-gray-400">(เลือกได้หลายคน)</span></label>
          ${schedProctorPickerHTML(2, 'proctor2', v('proctor2'), isExam && split)}
        </div>
      </div>
    </div>
    <div class="p-3 bg-green-50 rounded-xl border border-green-100 space-y-2">
      <label class="flex items-center gap-2 cursor-pointer"><input type="checkbox" id="schedNotify" ${notifyDefault ? 'checked' : ''} onchange="toggleSchedNotify()" class="w-4 h-4"><span class="text-sm font-medium text-green-800">🔔 สร้างประกาศแจ้งเตือนจากรายการนี้</span></label>
      <div id="schedNotifyOptions" class="${notifyDefault ? '' : 'hidden'} space-y-2">
        ${annRolesFieldHTML(notifyRolesDefault)}
        <label class="flex items-center gap-2 bg-white rounded-xl px-3 py-2 cursor-pointer border border-green-100"><input type="checkbox" id="schedNotifyLine" checked class="w-4 h-4"><span class="text-sm text-green-700">📢 ส่งประกาศนี้เข้า LINE</span></label>
      </div>
    </div>`;
}
function toggleSchedNotify() {
  const c = document.getElementById('schedNotify');
  const o = document.getElementById('schedNotifyOptions');
  if (o) o.classList.toggle('hidden', !(c && c.checked));
}

// สร้างประกาศแจ้งเตือนจากรายการปฏิทิน — เลือกบทบาทผู้รับ (roles) และเลือกส่ง LINE ได้
async function createScheduleAnnouncement(s, roles, sendLine) {
  const subjects = norm(s.subject_name).replace(/,\s*/g, ', ');
  const yr = norm(s.year_level);
  const type = norm(s.schedule_type);
  const isExam = type.includes('สอบ');
  const dateTh = (typeof toBuddhistDate === 'function' && toBuddhistDate(s.schedule_date)) || s.schedule_date || '';
  const timeRange = schedTimeRange(s);
  const proctors = norm(s.proctor).replace(/,\s*/g, ', ');
  const lines = [];
  lines.push((isExam ? 'รายวิชาที่สอบ: ' : 'รายการ: ') + (subjects || '-'));
  if (type) lines.push('ประเภท: ' + type);
  if (isExam && norm(s.exam_round)) lines.push('ครั้งที่: ' + norm(s.exam_round));
  lines.push('ชั้นปี: ' + (yr ? ('ชั้นปีที่ ' + yr) : 'ทุกชั้นปี'));
  if (dateTh) lines.push('วันที่: ' + dateTh);
  if (timeRange.trim()) lines.push('เวลา: ' + timeRange);
  const isSplit = isExam && norm(s.exam_split) !== '';
  if (isSplit) {
    lines.push('ห้องสอบที่ 1: ' + (norm(s.room) || '-') + (norm(s.student_count) ? ' (' + norm(s.student_count) + ' คน)' : ''));
    if (proctors) lines.push('ผู้คุมสอบ ห้อง 1: ' + proctors);
    lines.push('ห้องสอบที่ 2: ' + (norm(s.room2) || '-') + (norm(s.student_count2) ? ' (' + norm(s.student_count2) + ' คน)' : ''));
    const proctors2 = norm(s.proctor2).replace(/,\s*/g, ', ');
    if (proctors2) lines.push('ผู้คุมสอบ ห้อง 2: ' + proctors2);
  } else {
    if (norm(s.room)) lines.push('ห้อง: ' + norm(s.room));
    if (isExam && norm(s.student_count)) lines.push('จำนวนนักศึกษา: ' + norm(s.student_count) + ' คน');
    if (isExam && proctors) lines.push('อาจารย์ผู้คุมสอบ: ' + proctors);
  }
  const obj = {
    type: 'announcement',
    announcement_title: (isExam ? '📝 แจ้งกำหนดสอบ' : '📅 แจ้งเตือนปฏิทินกิจกรรมวิชาการ') + (yr ? ' — ชั้นปีที่ ' + yr : '') + (subjects ? ' (' + subjects + ')' : ''),
    announcement_content: lines.join('\n'),
    announcement_date: norm(s.schedule_date) || new Date().toISOString().slice(0, 10),
    event_type: isExam ? 'สอบ' : (type || 'ทั่วไป'),
    roles: roles || '',
    target_names: (isExam && [norm(s.proctor), norm(s.proctor2)].filter(Boolean).join(', ')) || '',
    line_notify: sendLine ? '✓' : '',
    created_at: new Date().toISOString()
  };
  try { return await GSheetDB.create(obj); } catch (_) { return { isOk: false }; }
}

function showAddScheduleModal() {
  showModal('เพิ่มรายการปฏิทินกิจกรรมวิชาการ', `
    <form id="addScheduleForm" class="space-y-3">
      ${scheduleFormBody({}, true)}
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึก</button>
    </form>
  `);
  renderSchedSubjectChips();
  renderSchedProctorChips(1);
  renderSchedExtraDateChips();
  updateSchedSplitState();
  window._schedWasExam = false;
  document.getElementById('addScheduleForm').onsubmit = async (e) => {
    e.preventDefault();
    // อ่านค่าการแจ้งเตือนก่อน (เพราะ modal จะถูกปิดหลังบันทึก)
    const notifyEl = document.getElementById('schedNotify');
    const doNotify = !!(notifyEl && notifyEl.checked);
    const roles = doNotify ? annCollectRoles() : '';
    const lineEl = document.getElementById('schedNotifyLine');
    const sendLine = !!(lineEl && lineEl.checked);
    await withLoading(e.target, async () => {
      const fd = new FormData(e.target);
      if (!(fd.get('schedule_type') || '').trim()) { showToast('กรุณาระบุประเภท', 'error'); return; }
      if (!(fd.get('subject_name') || '').trim()) { showToast('กรุณาระบุรายวิชา/กิจกรรม', 'error'); return; }
      const base = { type: 'schedule', created_at: new Date().toISOString() }; fd.forEach((v, k) => base[k] = v);
      // รวมวันที่: วันหลัก + วันที่เพิ่มเติม (จัดหลายวัน)
      const dates = [];
      if (norm(base.schedule_date)) dates.push(norm(base.schedule_date));
      getSchedExtraDatesArr().forEach(d => { d = norm(d); if (d && dates.indexOf(d) === -1) dates.push(d); });
      if (!dates.length) { showToast('กรุณาระบุวันที่', 'error'); return; }
      const objs = dates.map(d => Object.assign({}, base, { schedule_date: d }));
      const r = await GSheetDB.createMany(objs);
      if (r.isOk || r.ok) {
        if (doNotify) { for (const o of objs) { await createScheduleAnnouncement(o, roles, sendLine); } }
        showToast('เพิ่มรายการสำเร็จ ' + (r.ok || objs.length) + ' วัน' + (doNotify ? ' + สร้างประกาศแจ้งเตือน' : ''));
        closeModal();
      } else showToast('เกิดข้อผิดพลาด', 'error');
    });
  };
}

// ======================== GRADES ========================
// แผงวิเคราะห์ GPAx (หน้าผลการเรียน — สำหรับผู้ดูแล/วิชาการ/ผู้บริหาร)
// อิงตัวกรองชั้นปีเดิม (APP.filters._gradeYearLevel): '' = ทุกชั้นปี, '1'-'4', '__grad'
function gpaxAnalyticsHTML() {
  const sel = APP.filters._gradeYearLevel || '';
  let scope;
  if (sel === '__grad') scope = getDataByType('student').filter(isGraduate);
  else { scope = activeStudents(getDataByType('student')); if (sel) scope = scope.filter(s => norm(s.year_level) === sel); }

  const allGrades = getDataByType('grade');
  const gradeMap = { 'A': 4, 'B+': 3.5, 'B': 3, 'C+': 2.5, 'C': 2, 'D+': 1.5, 'D': 1, 'F': 0 };
  const byStu = {};
  allGrades.forEach(g => {
    const gv = gradeMap[norm(g.grade)]; if (gv === undefined) return;
    const sid = norm(g.student_id); const cr = Number(_gradeCredits(g)) || 3;
    if (!byStu[sid]) byStu[sid] = { p: 0, c: 0 };
    byStu[sid].p += gv * cr; byStu[sid].c += cr;
  });
  const round2 = x => Math.round(x * 100) / 100;
  const ranked = scope.map(s => {
    const rec = byStu[norm(s.student_id)];
    const gpax = rec && rec.c ? round2(rec.p / rec.c) : null;
    return { s, gpax };
  });
  const withGpax = ranked.filter(r => r.gpax !== null).sort((a, b) => b.gpax - a.gpax);
  const noGrade = ranked.length - withGpax.length;
  const atRisk = withGpax.filter(r => r.gpax < 2.30).length;

  // 9 ช่วง GPAx (ช่วงไม่ทับซ้อน)
  const buckets = [
    { label: '0.00 – 1.50', color: 'bg-red-600', n: withGpax.filter(r => r.gpax <= 1.50).length },
    { label: '1.51 – 2.00', color: 'bg-red-500', n: withGpax.filter(r => r.gpax > 1.50 && r.gpax <= 2.00).length },
    { label: '2.01 – 2.30', color: 'bg-orange-500', n: withGpax.filter(r => r.gpax > 2.00 && r.gpax <= 2.30).length },
    { label: '2.31 – 2.50', color: 'bg-amber-500', n: withGpax.filter(r => r.gpax > 2.30 && r.gpax <= 2.50).length },
    { label: '2.51 – 2.70', color: 'bg-yellow-500', n: withGpax.filter(r => r.gpax > 2.50 && r.gpax <= 2.70).length },
    { label: '2.71 – 3.00', color: 'bg-lime-500', n: withGpax.filter(r => r.gpax > 2.70 && r.gpax <= 3.00).length },
    { label: '3.01 – 3.50', color: 'bg-green-500', n: withGpax.filter(r => r.gpax > 3.00 && r.gpax <= 3.50).length },
    { label: '3.51 – 3.99', color: 'bg-emerald-500', n: withGpax.filter(r => r.gpax > 3.50 && r.gpax < 4.00).length },
    { label: '4.00 (เต็ม)', color: 'bg-teal-600', n: withGpax.filter(r => r.gpax >= 4.00).length }
  ];
  const bucketCards = buckets.map(b => `<div class="card-stat bg-white rounded-xl border border-gray-100 p-3 text-center">
    <div class="w-full h-1.5 rounded-full ${b.color} mb-2"></div>
    <p class="text-2xl font-bold text-gray-800">${b.n}</p>
    <p class="text-xs text-gray-500 mt-0.5">${b.label}</p>
  </div>`).join('');

  const scopeLabel = sel === '__grad' ? 'ผู้สำเร็จการศึกษา' : (sel ? 'ชั้นปีที่ ' + sel : 'ทุกชั้นปี');
  const rankRows = withGpax.map((r, i) => {
    const g = r.gpax;
    const c = g < 2.30 ? 'text-red-600 font-bold' : g >= 3.50 ? 'text-emerald-600 font-bold' : 'text-gray-800';
    const badge = g < 2.30 ? '<span class="ml-2 px-1.5 py-0.5 rounded text-[10px] bg-red-100 text-red-700">เฝ้าระวัง</span>' : '';
    return `<tr class="border-t hover:bg-gray-50">
      <td class="px-3 py-2 text-center text-gray-400">${i + 1}</td>
      <td class="px-3 py-2 font-mono text-primary">${r.s.student_id || ''}</td>
      <td class="px-3 py-2">${r.s.name || ''}${badge}</td>
      <td class="px-3 py-2 text-center">${r.s.year_level || ''}</td>
      <td class="px-3 py-2 text-center ${c}">${g.toFixed(2)}</td>
    </tr>`;
  }).join('');

  return `<div class="bg-white rounded-2xl border border-blue-100 p-5 mb-4">
    <h3 class="font-bold text-gray-800 mb-4 flex items-center gap-2"><i data-lucide="bar-chart-3" class="w-5 h-5 text-primary"></i>ภาพรวมผลการเรียน (GPAx) <span class="text-sm font-normal text-gray-500">— ${scopeLabel} · ${withGpax.length} คน${noGrade ? ' (ยังไม่มีเกรด ' + noGrade + ' คน)' : ''}</span></h3>
    <div class="grid grid-cols-1 sm:grid-cols-3 gap-3 mb-4">
      ${statCard('alert-triangle', 'GPAx ต่ำกว่า 2.30 (เฝ้าระวัง)', atRisk, 'คน', 'bg-red-500')}
      ${statCard('users', 'มีผลการเรียนแล้ว', withGpax.length, 'คน', 'bg-blue-500')}
      ${statCard('award', 'GPAx 3.50 ขึ้นไป', withGpax.filter(r => r.gpax >= 3.50).length, 'คน', 'bg-emerald-500')}
    </div>
    <p class="text-sm font-semibold text-gray-600 mb-2">จำนวนนักศึกษาแยกตามช่วง GPAx</p>
    <div class="grid grid-cols-3 sm:grid-cols-5 lg:grid-cols-9 gap-2 mb-4">${bucketCards}</div>
    <p class="text-sm font-semibold text-gray-600 mb-2"><i data-lucide="list-ordered" class="w-4 h-4 inline mr-1"></i>รายงาน GPAx เรียงมากไปน้อย <span class="font-normal text-gray-400">(${scopeLabel} — เลือกชั้นปีจากแถบด้านล่างเพื่อดูเฉพาะชั้นปี)</span></p>
    <div class="border border-gray-100 rounded-xl overflow-hidden">
      <div class="overflow-auto" style="max-height:340px"><table class="w-full text-sm">
        <thead class="sticky top-0"><tr class="bg-surface text-left">
          <th class="px-3 py-2 font-semibold text-center">อันดับ</th>
          <th class="px-3 py-2 font-semibold">รหัสนักศึกษา</th>
          <th class="px-3 py-2 font-semibold">ชื่อ-สกุล</th>
          <th class="px-3 py-2 font-semibold text-center">ชั้นปี</th>
          <th class="px-3 py-2 font-semibold text-center">GPAx</th>
        </tr></thead>
        <tbody>${rankRows || '<tr><td colspan="5" class="px-3 py-6 text-center text-gray-400">ยังไม่มีข้อมูลผลการเรียน</td></tr>'}</tbody>
      </table></div>
    </div>
  </div>`;
}

function gradesPage() {
  const isAdmin = isAdminRole();
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
    // ผลการเรียน (อาจารย์/อาจารย์ประจำชั้น): นับ/แสดงเฉพาะนักศึกษาที่กำลังศึกษา
    if (APP.currentRole === 'teacher' || APP.currentRole === 'classTeacher') {
      studentList = activeStudents(studentList);
    }

    // Year level + Advisor filter for admin/academic/executive
    let advisorSelector = '';
    if (canFilterByAdvisor) {
      // Year level filter
      const selectedGradeYear = APP.filters._gradeYearLevel || '';
      const isGradFilter = selectedGradeYear === '__grad';
      if (isGradFilter) {
        studentList = studentList.filter(s => isGraduate(s));
      } else if (selectedGradeYear) {
        studentList = studentList.filter(s => norm(s.year_level) === selectedGradeYear);
      }
      // โหมดที่ไม่ใช่ผู้สำเร็จการศึกษา → นับ/แสดงเฉพาะนักศึกษาที่กำลังศึกษา (ตัดพัก/ลาออก/จบออก)
      if (!isGradFilter) studentList = activeStudents(studentList);
      const yearLevelSelector = `<div class="bg-white rounded-2xl p-4 border border-blue-100 mb-4">
        <label class="block text-sm font-medium text-gray-700 mb-2"><i data-lucide="layers" class="w-4 h-4 inline mr-1"></i>กรองตามชั้นปี</label>
        <div class="flex flex-wrap gap-2">
          <button onclick="APP.filters._gradeYearLevel='';APP.filters._gradeStudent='';APP.filters._gradeSearch='';APP.pagination.page=1;renderCurrentPage()" class="px-4 py-2 rounded-xl text-sm font-medium ${selectedGradeYear === '' ? 'bg-primary text-white' : 'bg-gray-100 text-gray-600 hover:bg-gray-200'}">ทุกชั้นปี</button>
          ${['1', '2', '3', '4'].map(yr => `<button onclick="APP.filters._gradeYearLevel='${yr}';APP.filters._gradeStudent='';APP.filters._gradeSearch='';APP.pagination.page=1;renderCurrentPage()" class="px-4 py-2 rounded-xl text-sm font-medium ${selectedGradeYear === yr ? 'bg-primary text-white' : 'bg-gray-100 text-gray-600 hover:bg-gray-200'}">ชั้นปี ${yr}</button>`).join('')}
          <button onclick="APP.filters._gradeYearLevel='__grad';APP.filters._gradeStudent='';APP.filters._gradeSearch='';APP.filters._gradeAdvisor='';APP.pagination.page=1;renderCurrentPage()" class="px-4 py-2 rounded-xl text-sm font-medium ${isGradFilter ? 'bg-primary text-white' : 'bg-gray-100 text-gray-600 hover:bg-gray-200'}">ผู้สำเร็จการศึกษา</button>
        </div>
        ${selectedGradeYear ? `<p class="text-xs text-gray-500 mt-2"><i data-lucide="info" class="w-3 h-3 inline mr-1"></i>${isGradFilter ? 'แสดงเฉพาะผู้สำเร็จการศึกษา' : 'แสดงเฉพาะนักศึกษาชั้นปีที่ ' + selectedGradeYear} (${studentList.length} คน)</p>` : ''}
      </div>`;

      // กรองตามรุ่น — แสดงเฉพาะตอนเลือก "ผู้สำเร็จการศึกษา" เท่านั้น (ชั้นปี 1-4 ไม่แสดง)
      const gradStudentsAll = getDataByType('student').filter(s => isGraduate(s));
      const allBatches = [...new Set(gradStudentsAll.map(s => norm(s.batch)).filter(Boolean))].sort((a, b) => b.localeCompare(a, undefined, { numeric: true }));
      const selectedBatch = APP.filters._gradeBatch || '';
      if (isGradFilter && selectedBatch) studentList = gradStudentsAll.filter(s => norm(s.batch) === selectedBatch);
      const batchSelector = isGradFilter ? `<div class="bg-white rounded-2xl p-4 border border-blue-100 mb-4">
        <label class="block text-sm font-medium text-gray-700 mb-2"><i data-lucide="users" class="w-4 h-4 inline mr-1"></i>กรองตามรุ่น <span class="font-normal text-gray-400 text-xs">(เฉพาะผู้สำเร็จการศึกษา)</span></label>
        <select onchange="APP.filters._gradeBatch=this.value;APP.filters._gradeStudent='';APP.filters._gradeSearch='';APP.pagination.page=1;renderCurrentPage()" class="w-full border border-gray-200 rounded-xl px-4 py-2.5 text-sm">
          <option value="">-- ทุกรุ่น --</option>
          ${allBatches.map(b => `<option value="${b}" ${selectedBatch === b ? 'selected' : ''}>รุ่นที่ ${b}</option>`).join('')}
        </select>
        ${selectedBatch ? `<p class="text-xs text-gray-500 mt-2"><i data-lucide="info" class="w-3 h-3 inline mr-1"></i>แสดงเฉพาะผู้สำเร็จการศึกษา รุ่นที่ ${selectedBatch} (${studentList.length} คน)</p>` : ''}
      </div>` : '';

      const allAdvisors = [...new Set(studentList.map(s => s.advisor).filter(Boolean))].sort();
      const selectedAdvisor = APP.filters._gradeAdvisor || '';
      if (!isGradFilter && selectedAdvisor) {
        studentList = studentList.filter(s => (s.advisor || '') === selectedAdvisor);
      }
      // ผู้สำเร็จการศึกษา: ไม่แสดงตัวกรองอาจารย์ที่ปรึกษา (คงตัวกรองรุ่นและเลือกนักศึกษาไว้)
      const advisorDiv = isGradFilter ? '' : `<div class="bg-white rounded-2xl p-4 border border-blue-100 mb-4">
        <label class="block text-sm font-medium text-gray-700 mb-2"><i data-lucide="user-check" class="w-4 h-4 inline mr-1"></i>กรองตามอาจารย์ที่ปรึกษา</label>
        <select onchange="APP.filters._gradeAdvisor=this.value;APP.filters._gradeStudent='';APP.filters._gradeSearch='';APP.pagination.page=1;renderCurrentPage()" class="w-full border border-gray-200 rounded-xl px-4 py-2.5 text-sm">
          <option value="">-- แสดงนักศึกษาทั้งหมด --</option>
          ${allAdvisors.map(a => `<option value="${a}" ${selectedAdvisor === a ? 'selected' : ''}>${a}</option>`).join('')}
        </select>
        ${selectedAdvisor ? `<p class="text-xs text-gray-500 mt-2"><i data-lucide="info" class="w-3 h-3 inline mr-1"></i>แสดงเฉพาะนักศึกษาในความดูแลของ ${selectedAdvisor} (${studentList.length} คน)</p>` : ''}
      </div>`;
      advisorSelector = `${yearLevelSelector}${batchSelector}${advisorDiv}`;
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
  const yearScopeGrades = data.slice(); // เกรดของนักศึกษาคนนี้ (ก่อนกรองปี) — ใช้สร้างตัวเลือก "ปีการศึกษา" เฉพาะรุ่น/ปีที่มีจริง
  data = applyFilters(data);
  // เรียงตามปีการศึกษา จากเก่าไปใหม่ แล้วตามภาคการศึกษา (1, 2, ฤดูร้อน=3)
  data.sort((a, b) => {
    const ay = parseInt(norm(a.academic_year), 10) || 0;
    const by = parseInt(norm(b.academic_year), 10) || 0;
    if (ay !== by) return ay - by;
    const as = parseInt(normSem(a.semester), 10) || 0;
    const bs = parseInt(normSem(b.semester), 10) || 0;
    return as - bs;
  });
  const total = data.length; const paged = paginate(data);

  // GPA calc
  let gpaSection = '';
  if (data.length && (isStudent || selectedStudentName)) {
    const gradeMap = { 'A': 4, 'B+': 3.5, 'B': 3, 'C+': 2.5, 'C': 2, 'D+': 1.5, 'D': 1, 'F': 0 };
    let totalCredits = 0, totalPoints = 0;
    data.forEach(g => { const gv = gradeMap[g.grade]; const cr = Number(_gradeCredits(g)) || 3; if (gv !== undefined) { totalPoints += gv * cr; totalCredits += cr } });
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
  ${['admin', 'academic', 'registrar', 'executive'].includes(APP.currentRole) ? gpaxAnalyticsHTML() : ''}
  ${studentSelector}
  ${noSelectionMsg || `${filterBar({ yearData: yearScopeGrades })}
  ${gpaSection}
  <div class="bg-white rounded-2xl border border-blue-100 overflow-hidden">
    <div class="overflow-x-auto"><table class="w-full text-sm">
      <thead><tr class="bg-surface text-left"><th class="px-4 py-3 font-semibold">รหัสวิชา</th><th class="px-4 py-3 font-semibold">รายวิชา</th><th class="px-4 py-3 font-semibold">เกรด</th><th class="px-4 py-3 font-semibold">หน่วยกิต</th><th class="px-4 py-3 font-semibold">ภาค/ปี</th>${isAdmin ? '<th class="px-4 py-3"></th>' : ''}</tr></thead>
      <tbody>${paged.length ? paged.map(g => `<tr class="border-t hover:bg-gray-50">
        <td class="px-4 py-3 font-mono text-primary">${g.subject_code || ''}</td><td class="px-4 py-3">${g.subject_name || ''}</td>
        <td class="px-4 py-3"><span class="px-2 py-1 rounded-full text-xs font-bold ${g.grade === 'F' ? 'bg-red-100 text-red-700' : 'bg-green-100 text-green-700'}">${g.grade || ''}</span></td>
        <td class="px-4 py-3">${_gradeCredits(g) || ''}</td><td class="px-4 py-3">${semLabel(g.semester)}/${g.academic_year || ''}</td>
        ${isAdmin ? `<td class="px-4 py-3"><div class="flex gap-1"><button onclick="showEditGradeModal('${g.__backendId}')" class="text-blue-400 hover:text-blue-600" title="แก้ไข"><i data-lucide="pencil" class="w-4 h-4"></i></button><button onclick="deleteRecord('${g.__backendId}')" class="text-red-400 hover:text-red-600" title="ลบ"><i data-lucide="trash-2" class="w-4 h-4"></i></button></div></td>` : ''}</tr>`).join('') : '<tr><td colspan="5" class="px-4 py-8 text-center text-gray-400">ไม่มีข้อมูล</td></tr>'}</tbody>
    </table></div>
  </div>
  ${paginationHTML(total, APP.pagination.perPage, APP.pagination.page, 'changePage')}`}`;
}

// ======================== Auto-fill รายวิชา จากรหัสวิชา (ฟอร์มผลการเรียน) ========================
// คืนรายวิชาที่ไม่ซ้ำรหัส กรองตาม "ปีการศึกษาที่เลือก" และ/หรือ "ชั้นปีของนักศึกษาที่เลือก"
// ถ้ากรองแล้วไม่เหลือรายการ → fallback ใช้ทั้งหมด (กันกรณีข้อมูล subject กรอกปี/ชั้นปีไม่ครบ)
// หน่วยกิตของแถวเกรด: ใช้ค่าที่บันทึกในเกรด ถ้าว่างให้ดึงจากแท็บ subject ตามรหัสวิชา
function _gradeCredits(g) {
  var c = norm(g && g.credits);
  if (c) return c;
  var sub = getDataByType('subject').find(s => norm(s.subject_code) === norm(g && g.subject_code));
  return sub ? norm(sub.credits) : '';
}

// รหัสหน่วยกิตของแถวเกรด แสดงแบบ น(ท-ป-อ) เหมือนหน้ารายวิชา — ดึงชั่วโมง ท/ป/อ จากแท็บ subject ตามรหัสวิชา
function _gradeCreditCode(g) {
  var sub = getDataByType('subject').find(s => norm(s.subject_code) === norm(g && g.subject_code));
  if (sub) { var cc = creditCode(sub); if (cc) return cc; }
  return norm(g && g.credits) || '';
}

function _gradeSubjectsFor(studentId, academicYear) {
  let subs = getDataByType('subject').filter(s => norm(s.subject_code));
  const stu = norm(studentId) ? getDataByType('student').find(s => norm(s.student_id) === norm(studentId)) : null;
  const stuBatch = stu ? norm(stu.batch) : '';

  if (stuBatch) {
    // รู้รุ่น (batch) ของนักศึกษา → แสดงรายวิชา "ของรุ่นนั้น" ทั้งหมด (รวมวิชาที่ไม่ระบุรุ่น)
    // ไม่กรองด้วยปีการศึกษา/ชั้นปี เพื่อให้เห็นวิชาของรุ่นครบทุกชั้นปี (กันวิชาตกหล่นเหมือนกรณีรุ่น 78)
    const byBatch = subs.filter(s => !norm(s.batch) || norm(s.batch) === stuBatch);
    if (byBatch.length) subs = byBatch;
  } else if (norm(academicYear)) {
    // ยังไม่เลือกนักศึกษา/ไม่รู้รุ่น → ใช้ปีการศึกษาเป็นตัวกรอง (เก็บวิชาที่ปีว่างไว้ด้วย)
    const byYear = subs.filter(s => !norm(s.academic_year) || norm(s.academic_year) === norm(academicYear));
    if (byYear.length) subs = byYear;
  }

  const seen = new Set(); const out = [];
  subs.forEach(s => { const c = norm(s.subject_code); if (!seen.has(c)) { seen.add(c); out.push(s); } });
  return out.sort((a, b) => norm(a.subject_code).localeCompare(norm(b.subject_code), 'th'));
}

// สร้าง <option> ของ select รหัสวิชา (แสดงเฉพาะรหัสวิชา)
function _gradeCodeOptions(subs, selected) {
  return `<option value="">— เลือกรหัสวิชา —</option>` +
    subs.map(s => `<option value="${norm(s.subject_code)}" ${norm(s.subject_code) === norm(selected || '') ? 'selected' : ''}>${norm(s.subject_code)}</option>`).join('');
}

// โหลดรายการรหัสวิชาใหม่ เมื่อเปลี่ยนนักศึกษา/ปีการศึกษา (prefix: 'ag' = เพิ่ม, 'eg' = แก้ไข)
function _rebuildGradeSubjectSelect(prefix) {
  const p = prefix || 'ag';
  const codeSel = document.getElementById(p + 'Code'); if (!codeSel) return;
  const studentId = (document.getElementById(p + 'Student') || {}).value || '';
  const academicYear = (document.getElementById(p + 'Year') || {}).value || '';
  const subs = _gradeSubjectsFor(studentId, academicYear);
  window['_' + p + 'Subjects'] = subs;
  const cur = codeSel.value;
  codeSel.innerHTML = _gradeCodeOptions(subs, cur);
  _onGradeCodeChange(p);
}

// เมื่อเลือกรหัสวิชา → เติมชื่อรายวิชา + หน่วยกิต + ภาคการศึกษา อัตโนมัติ
function _onGradeCodeChange(prefix) {
  const p = prefix || 'ag';
  const codeSel = document.getElementById(p + 'Code'); if (!codeSel) return;
  const subs = window['_' + p + 'Subjects'] || [];
  const subj = subs.find(s => norm(s.subject_code) === norm(codeSel.value));
  const nameEl = document.getElementById(p + 'Name');
  const crEl = document.getElementById(p + 'Credits');
  const semEl = document.getElementById(p + 'Sem');
  if (nameEl) nameEl.value = subj ? norm(subj.subject_name) : '';
  if (subj) {
    if (crEl && norm(subj.credits)) crEl.value = norm(subj.credits);
    if (semEl && norm(subj.semester)) semEl.value = norm(subj.semester);
  }
}

// ===== เพิ่มผลการเรียนแบบหลายรายวิชาพร้อมกัน (multi-row) =====
// โหลดรายการรหัสวิชาตามนักศึกษา/ปีการศึกษา แล้วรีเฟรช dropdown ของทุกแถว
function _agRefreshSubjects() {
  const studentId = (document.getElementById('agStudent') || {}).value || '';
  const academicYear = (document.getElementById('agYear') || {}).value || '';
  window._agSubjects = _gradeSubjectsFor(studentId, academicYear);
  document.querySelectorAll('#agRows [data-agcode]').forEach(sel => {
    const cur = sel.value;
    sel.innerHTML = _gradeCodeOptions(window._agSubjects, cur);
    _agOnCodeChange(sel);
  });
}
// เลือกรหัสวิชาในแถว → เติมชื่อรายวิชา/หน่วยกิต/ภาคเรียน อัตโนมัติ
function _agOnCodeChange(sel) {
  const row = sel.closest('[data-agrow]'); if (!row) return;
  const subj = (window._agSubjects || []).find(s => norm(s.subject_code) === norm(sel.value));
  const nameEl = row.querySelector('[data-agname]');
  const crEl = row.querySelector('[data-agcredits]');
  const semEl = row.querySelector('[data-agsem]');
  if (nameEl) nameEl.value = subj ? norm(subj.subject_name) : '';
  if (subj) {
    if (crEl && norm(subj.credits)) crEl.value = norm(subj.credits);
    if (semEl && norm(subj.semester)) semEl.value = norm(subj.semester);
  }
}
// เพิ่มแถวรายวิชาใหม่
function _agAddRow() {
  const wrap = document.getElementById('agRows'); if (!wrap) return;
  const subs = window._agSubjects || [];
  const div = document.createElement('div');
  div.setAttribute('data-agrow', '');
  div.className = 'grid grid-cols-12 gap-2 items-center';
  div.innerHTML = `
    <div class="col-span-3"><select data-agcode onchange="_agOnCodeChange(this)" class="w-full border rounded-lg px-2 py-1.5 text-sm">${_gradeCodeOptions(subs, '')}</select></div>
    <div class="col-span-4"><input data-agname readonly class="w-full border rounded-lg px-2 py-1.5 text-sm bg-gray-50" placeholder="เลือกรหัสวิชา"></div>
    <div class="col-span-2"><select data-aggrade class="w-full border rounded-lg px-2 py-1.5 text-sm"><option>A</option><option>B+</option><option>B</option><option>C+</option><option>C</option><option>D+</option><option>D</option><option>F</option></select></div>
    <div class="col-span-1"><input data-agcredits type="number" value="3" class="w-full border rounded-lg px-2 py-1.5 text-sm"></div>
    <div class="col-span-1"><select data-agsem class="w-full border rounded-lg px-2 py-1.5 text-sm"><option value="1">1</option><option value="2">2</option></select></div>
    <div class="col-span-1 text-center"><button type="button" onclick="this.closest('[data-agrow]').remove()" class="text-red-400 hover:text-red-600" title="ลบแถว"><i data-lucide="trash-2" class="w-4 h-4"></i></button></div>`;
  wrap.appendChild(div);
  lucide.createIcons();
}

function showAddGradeModal() {
  showModal('เพิ่มผลการเรียน', `
    <form id="addGradeForm" class="space-y-3">
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-xs text-gray-600 mb-1">นักศึกษา * (พิมพ์รหัสหรือเลือก)</label><input list="addGradeStudentList" id="agStudent" required class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="พิมพ์รหัสนักศึกษา..." onchange="_agRefreshSubjects()" oninput="_agRefreshSubjects()">${studentDatalistHTML('addGradeStudentList')}</div>
        <div><label class="block text-xs text-gray-600 mb-1">ปีการศึกษา</label><input id="agYear" value="2568" class="w-full border rounded-xl px-3 py-2 text-sm" onchange="_agRefreshSubjects()" oninput="_agRefreshSubjects()"></div>
      </div>
      <div class="flex items-center justify-between pt-1">
        <label class="text-xs font-medium text-gray-600">รายวิชา <span class="text-gray-400">(เพิ่มได้หลายวิชาพร้อมกัน)</span></label>
        <button type="button" onclick="_agAddRow()" class="flex items-center gap-1 text-xs px-2 py-1 bg-primaryLight text-primary rounded-lg hover:bg-blue-100"><i data-lucide="plus" class="w-3.5 h-3.5"></i>เพิ่มรายวิชา</button>
      </div>
      <div class="grid grid-cols-12 gap-2 text-xs text-gray-400 px-1">
        <div class="col-span-3">รหัสวิชา</div><div class="col-span-4">รายวิชา</div><div class="col-span-2">เกรด</div><div class="col-span-1">นก.</div><div class="col-span-1">ภาค</div><div class="col-span-1"></div>
      </div>
      <div id="agRows" class="space-y-2 max-h-[40vh] overflow-auto"></div>
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึกทั้งหมด</button>
    </form>
  `, null, 'max-w-2xl');
  _agRefreshSubjects();
  _agAddRow();
  document.getElementById('addGradeForm').onsubmit = async (e) => {
    e.preventDefault();
    await withLoading(e.target, async () => {
      const studentId = (document.getElementById('agStudent').value || '').trim();
      const academicYear = (document.getElementById('agYear').value || '').trim();
      if (!studentId) { showToast('กรุณาเลือกนักศึกษา', 'error'); return; }
      const objs = [];
      document.querySelectorAll('#agRows [data-agrow]').forEach(row => {
        const name = (row.querySelector('[data-agname]').value || '').trim();
        if (!name) return; // ข้ามแถวที่ยังไม่ได้เลือกวิชา
        objs.push({
          type: 'grade', student_id: studentId,
          subject_code: row.querySelector('[data-agcode]').value || '',
          subject_name: name,
          grade: row.querySelector('[data-aggrade]').value || '',
          credits: Number(row.querySelector('[data-agcredits]').value) || 0,
          semester: row.querySelector('[data-agsem]').value || '',
          academic_year: academicYear,
          created_at: new Date().toISOString()
        });
      });
      if (!objs.length) { showToast('กรุณาเลือกอย่างน้อย 1 รายวิชา', 'error'); return; }
      const r = await GSheetDB.createMany(objs);
      if (r.isOk) { showToast(`เพิ่มผลการเรียน ${objs.length} รายวิชาสำเร็จ`); closeModal(); }
      else showToast(`บันทึกสำเร็จ ${r.ok || 0} / ผิดพลาด ${r.fail || 0}`, (r.ok ? 'success' : 'error'));
    });
  };
}

// ======================== ใบระเบียนแสดงผลการเรียน (ผู้สำเร็จการศึกษา) ========================
function isGraduate(stu) { return norm(stu && stu.status) === 'สำเร็จการศึกษา' || norm(stu && stu.year_level) === 'จบ'; }

const THAI_MONTHS = ['', 'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];
function toThaiLongDate(dateStr) {
  if (!dateStr) return '';
  const s = String(dateStr).trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) { return parseInt(m[3], 10) + ' ' + THAI_MONTHS[parseInt(m[2], 10)] + ' ' + (parseInt(m[1], 10) + 543); }
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) { const y = parseInt(m[3], 10); return parseInt(m[1], 10) + ' ' + THAI_MONTHS[parseInt(m[2], 10)] + ' ' + (y > 2400 ? y : y + 543); }
  return s;
}

// ชั่วโมงฝึกปฏิบัติการพยาบาลตามหลักสูตร (มาตรฐาน — แก้ได้ที่นี่)
const PRACTICUM_HOURS = [
  ['ปฏิบัติการพยาบาลขั้นพื้นฐาน', '180'],
  ['ปฏิบัติการพยาบาลผู้ใหญ่และผู้สูงอายุ', '270'],
  ['ปฏิบัติการพยาบาลเด็กและวัยรุ่น', '180'],
  ['ปฏิบัติการพยาบาลสุขภาพชุมชน', '135'],
  ['ปฏิบัติการพยาบาลผู้สูงอายุ', '90'],
  ['ปฏิบัติการพยาบาลมารดา ทารกและการผดุงครรภ์', '270'],
  ['ปฏิบัติการพยาบาลสุขภาพจิตและจิตเวช', '135'],
  ['ปฏิบัติการพยาบาลและการบริหารจัดการสุขภาวะชุมชน', '135'],
  ['ปฏิบัติการรักษาโรคเบื้องต้นสำหรับพยาบาล', '135'],
  ['ปฏิบัติการบริหารและการจัดการคุณภาพทางการพยาบาล', '90'],
];

function buildOfficialTranscript(stu, logoSrc) {
  const stuId = norm(stu.student_id);
  const grades = getDataByType('grade').filter(g => norm(g.student_id) === stuId);
  const gradeMap = { 'A': 4, 'B+': 3.5, 'B': 3, 'C+': 2.5, 'C': 2, 'D+': 1.5, 'D': 1, 'F': 0 };
  const acadOf = g => { const y = parseInt(norm(g.academic_year), 10); return isNaN(y) ? null : y; };
  // ปีการศึกษาเริ่มต้น (ใช้ admission_date ก่อน ถ้าไม่มีใช้ปีต่ำสุดของเกรด)
  let baseYear = null;
  if (stu.admission_date) { const m = String(stu.admission_date).match(/^(\d{4})-(\d{1,2})/); if (m) { let y = parseInt(m[1], 10) + 543; if (parseInt(m[2], 10) < 6) y -= 1; baseYear = y; } }
  if (baseYear == null) { const ys = grades.map(acadOf).filter(v => v != null); if (ys.length) baseYear = Math.min(...ys); }

  const yearGroups = {};
  let totalCredits = 0, totalPoints = 0;
  grades.forEach(g => {
    const ay = acadOf(g);
    let sy = (baseYear != null && ay != null) ? (ay - baseYear + 1) : 1;
    if (sy < 1) sy = 1;
    if (!yearGroups[sy]) yearGroups[sy] = { courses: [], credits: 0 };
    yearGroups[sy].courses.push(g);
    const cr = Number(_gradeCredits(g)) || 0; const gv = gradeMap[g.grade];
    yearGroups[sy].credits += cr;
    if (gv !== undefined) { totalPoints += gv * cr; totalCredits += cr; }
  });
  const gpax = totalCredits ? (totalPoints / totalCredits).toFixed(2) : '-';
  const sortedYears = Object.keys(yearGroups).map(Number).sort((a, b) => a - b);

  const courseHead = `<tr style="background:#dceaf7"><th style="border:1px solid #999;padding:2px 4px;font-size:10px;width:22%">รหัสรายวิชา</th><th style="border:1px solid #999;padding:2px 4px;font-size:10px">ชื่อรายวิชา</th><th style="border:1px solid #999;padding:2px 4px;font-size:10px;width:15%">หน่วยกิต</th><th style="border:1px solid #999;padding:2px 4px;font-size:10px;width:10%">เกรด</th></tr>`;
  const yearBlock = sy => {
    const grp = yearGroups[sy]; if (!grp) return '';
    let rows = `<tr><td colspan="4" style="text-align:center;font-weight:700;padding:2px;border:1px solid #999;background:#eef5fb;font-size:10px">ชั้นปีที่ ${sy}</td></tr>`;
    grp.courses.forEach(g => {
      rows += `<tr><td style="border:1px solid #999;padding:2px 4px;font-family:monospace;font-size:9.5px">${g.subject_code || ''}</td><td style="border:1px solid #999;padding:2px 4px;font-size:9.5px">${g.subject_name || ''}</td><td style="border:1px solid #999;padding:2px 4px;text-align:center;font-size:9.5px">${_gradeCreditCode(g) || ''}</td><td style="border:1px solid #999;padding:2px 4px;text-align:center;font-size:9.5px;font-weight:600">${g.grade || ''}</td></tr>`;
    });
    rows += `<tr><td colspan="2" style="border:1px solid #999;padding:2px 4px;text-align:right;font-weight:600;font-size:9.5px">รวม</td><td style="border:1px solid #999;padding:2px 4px;text-align:center;font-weight:700;font-size:9.5px">${grp.credits}</td><td style="border:1px solid #999"></td></tr>`;
    return rows;
  };
  const half = Math.ceil(sortedYears.length / 2) || 1;
  const leftRows = sortedYears.slice(0, half).map(yearBlock).join('');
  const rightRows = sortedYears.slice(half).map(yearBlock).join('');

  const practicumRows = PRACTICUM_HOURS.map(p => `<div style="display:flex;justify-content:space-between"><span>${p[0]}</span><span>${p[1]} ชั่วโมง</span></div>`).join('');
  const engPass = getDataByType('eng_result').some(e => norm(e.student_id) === stuId && e.eng_status === 'ผ่าน');
  const college = (APP.config && APP.config.college_name) || 'วิทยาลัยพยาบาลบรมราชชนนี กรุงเทพ';
  const director = (APP.config && APP.config.director_name) || 'ผู้ช่วยศาสตราจารย์พนารัตน์ วิศวเทพนิมิตร';
  const logoTag = logoSrc ? `<img src="${logoSrc}" style="width:70px;height:auto;margin:0 auto 4px auto;display:block">` : '';

  return `
  <div style="font-family:'Sarabun',sans-serif;color:#000;font-size:11px;width:100%">
    <div style="text-align:center;margin-bottom:6px">
      ${logoTag}
      <div style="font-weight:700;font-size:15px">ระเบียนแสดงผลการเรียน</div>
      <div style="font-size:11px">สถาบันพระบรมราชชนก กระทรวงสาธารณสุข</div>
      <div style="font-size:11px">คณะพยาบาลศาสตร์</div>
      <div style="font-size:11px">${college}</div>
    </div>
    <table style="width:100%;font-size:10.5px;margin-bottom:6px;border-collapse:collapse"><tr>
      <td style="vertical-align:top;width:56%;padding-right:6px">
        <div>รหัสนักศึกษา: <b>${stu.student_id || ''}</b></div>
        <div>ชื่อ-นามสกุล (ไทย): <b>${stu.name || ''}</b></div>
        <div>(อังกฤษ): <b>${stu.name_en || ''}</b></div>
        <div>วันที่เกิด: <b>${toThaiLongDate(stu.birth_date) || '-'}</b></div>
        <div>จังหวัดที่เกิด: <b>${stu.birth_province || '-'}</b></div>
        <div>สัญชาติ: <b>${stu.nationality || 'ไทย'}</b> &nbsp;&nbsp; ศาสนา: <b>${stu.religion || '-'}</b></div>
      </td>
      <td style="vertical-align:top;width:44%">
        <div>วันที่เข้ารับการศึกษา: <b>${toThaiLongDate(stu.admission_date) || '-'}</b></div>
        <div>วันที่สำเร็จการศึกษา: <b>${toThaiLongDate(stu.graduation_date) || '-'}</b></div>
        <div>วุฒิการศึกษา: <b>${stu.degree || 'พยาบาลศาสตรบัณฑิต'}</b></div>
        <div>เกียรตินิยม: <b>${stu.honors || '-'}</b></div>
        <div>วุฒิการศึกษาเดิม: <b>${stu.prev_education || 'มัธยมศึกษาปีที่ 6'}</b></div>
      </td>
    </tr></table>
    <table style="width:100%;border-collapse:collapse"><tr>
      <td style="vertical-align:top;width:50%;padding-right:4px">
        <table style="width:100%;border-collapse:collapse">${courseHead}${leftRows}</table>
      </td>
      <td style="vertical-align:top;width:50%;padding-left:4px">
        <table style="width:100%;border-collapse:collapse">${courseHead}${rightRows}</table>
        <div style="margin-top:6px;font-size:10px;border:1px solid #999;padding:4px">
          <div style="display:flex;justify-content:space-between"><span>จำนวนหน่วยกิตตามหลักสูตร</span><span><b>${stu.curriculum_credits || totalCredits}</b> หน่วย</span></div>
          <div style="display:flex;justify-content:space-between"><span>จำนวนหน่วยกิตที่ลงทะเบียน</span><span><b>${totalCredits}</b> หน่วย</span></div>
          <div style="display:flex;justify-content:space-between"><span>คะแนนเฉลี่ยสะสมตลอดหลักสูตร</span><span><b>${gpax}</b></span></div>
        </div>
        <div style="margin-top:6px;font-size:10px">
          <div style="font-weight:600">จำนวนชั่วโมงฝึกปฏิบัติการพยาบาล:</div>
          ${practicumRows}
        </div>
        <div style="margin-top:4px;font-size:10px">
          <div style="display:flex;justify-content:space-between"><span>การทดสอบภาษาอังกฤษมาตรฐานของสถาบันพระบรมราชชนก</span><span><b>${engPass ? 'ผ่าน' : (stu.eng_test || '-')}</b></span></div>
          <div style="display:flex;justify-content:space-between"><span>การสอบรวบยอดของสถาบันพระบรมราชชนก</span><span><b>${stu.comprehensive_exam || '-'}</b></span></div>
        </div>
      </td>
    </tr></table>
    <div style="margin-top:6px;font-size:9px;color:#333">
      <div style="font-weight:600">ความหมายเกรด</div>
      <div>A : 4.00 (ดีเยี่ยม) &nbsp; B+ : 3.50 (ดีมาก) &nbsp; B : 3.00 (ดี) &nbsp; C+ : 2.50 (ค่อนข้างดี) &nbsp; C : 2.00 (พอใช้) &nbsp; D+ : 1.50 (อ่อน) &nbsp; D : 1.00 (อ่อนมาก)</div>
      <div>F : ตก &nbsp; S : พึงพอใจ &nbsp; U : ไม่พึงพอใจ &nbsp; CP : เทียบโอน &nbsp; AU : ไม่นับหน่วยกิต &nbsp; W : ถอนรายวิชา &nbsp; P : ผ่าน</div>
    </div>
    <div style="margin-top:26px;text-align:center;font-size:10.5px">
      <div>….......................................................................................</div>
      <div>(${director})</div>
      <div>ผู้อำนวยการ${college}</div>
      <div>นายทะเบียน</div>
    </div>
  </div>`;
}

function _renderOfficialTranscript(stu) {
  const stuNameSafe = (stu.name || '').replace(/'/g, "\\'");
  const inner = buildOfficialTranscript(stu, 'https://cdn.jsdelivr.net/gh/JOB-BCNB-P/LOGO/Logo%20Thai.png');
  showModal('ระเบียนแสดงผลการเรียน', `
    <div id="transcriptContent" style="background:#fff;padding:8px;max-width:780px;margin:auto;overflow-x:auto">${inner}</div>
    <div class="flex justify-center mt-4"><button onclick="downloadTranscriptPDF('${stuNameSafe}')" class="flex items-center gap-2 px-6 py-2.5 bg-primary text-white rounded-xl hover:bg-primaryDark text-sm"><i data-lucide="download" class="w-4 h-4"></i>ดาวน์โหลด PDF</button></div>
  `, null, 'max-w-4xl');
  setTimeout(() => lucide.createIcons(), 100);
}

async function _printOfficialTranscript(stu) {
  let logoBase64 = '';
  const srcs = (window.LOGO_SOURCES || ['https://cdn.jsdelivr.net/gh/JOB-BCNB-P/LOGO/Logo%20Thai.png']);
  for (const src of srcs) { try { const resp = await fetch(src); if (resp.ok) { const blob = await resp.blob(); logoBase64 = await new Promise(r => { const fr = new FileReader(); fr.onload = () => r(fr.result); fr.readAsDataURL(blob); }); break; } } catch (e) { } }
  const inner = buildOfficialTranscript(stu, logoBase64);
  const html = `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>ระเบียนแสดงผลการเรียน - ${stu.name || ''}</title>
    <style>@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;600;700&display=swap');*{font-family:'Sarabun',sans-serif;box-sizing:border-box}body{margin:0;padding:8mm}@media print{@page{size:A4;margin:7mm}}</style>
    </head><body>${inner}<script>window.onload=function(){window.print()}<\/script></body></html>`;
  const w = window.open('', '_blank');
  if (w) { w.document.write(html); w.document.close(); } else { showToast('กรุณาอนุญาต Popup เพื่อดาวน์โหลด PDF', 'error'); }
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
  if (isGraduate(stu)) { _renderOfficialTranscript(stu); return; }
  const stuId = norm(stu.student_id);
  let grades = getDataByType('grade').filter(g => norm(g.student_id) === stuId);
  // Respect current filters (academic year / semester) selected on grades page
  const filterYear = APP.filters.academicYear || '';
  const filterSem = APP.filters.semester || '';
  if (filterYear) grades = grades.filter(g => norm(g.academic_year) === filterYear);
  if (filterSem) grades = grades.filter(g => normSem(g.semester) === filterSem);

  const gradeMap = { 'A': 4, 'B+': 3.5, 'B': 3, 'C+': 2.5, 'C': 2, 'D+': 1.5, 'D': 1, 'F': 0 };

  const semesters = {};
  grades.forEach(g => {
    // key: ปี/ภาค เพื่อให้เรียงตามปีก่อน แล้วค่อยภาค
    const key = `${norm(g.academic_year) || '0000'}/${normSem(g.semester) || '0'}`;
    if (!semesters[key]) semesters[key] = { semester: normSem(g.semester) || '1', year: norm(g.academic_year) || '', grades: [] };
    semesters[key].grades.push(g);
  });
  const semKeys = Object.keys(semesters).sort();

  let totalCreditsAll = 0, totalPointsAll = 0;
  let tableRows = '';
  semKeys.forEach(key => {
    const sem = semesters[key];
    const semText = sem.semester === '3' ? 'ภาคฤดูร้อน' : 'ภาคการศึกษาที่ ' + sem.semester;
    tableRows += `<tr class="bg-blue-50"><td colspan="4" class="px-3 py-2 font-bold text-primary text-center">${semText} ปีการศึกษา ${sem.year}</td></tr>`;
    let semCredits = 0, semPoints = 0;
    sem.grades.forEach(g => {
      const cr = Number(_gradeCredits(g)) || 0;
      const gv = gradeMap[g.grade];
      tableRows += `<tr class="border-t border-gray-200"><td class="px-3 py-1.5 font-mono text-xs">${g.subject_code || ''}</td><td class="px-3 py-1.5 text-xs">${g.subject_name || ''}</td><td class="px-3 py-1.5 text-center text-xs">${_gradeCreditCode(g) || ''}</td><td class="px-3 py-1.5 text-center text-xs font-bold">${g.grade || ''}</td></tr>`;
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
      <div class="relative">
      <div class="text-center mb-3">
        <img src="https://cdn.jsdelivr.net/gh/JOB-BCNB-P/LOGO/Logo%20Thai.png" alt="Logo" data-li="0" style="width:60px;height:auto;margin:0 auto 6px auto;display:block;" onerror="if(typeof logoFallback==='function'){logoFallback(this)}else{this.style.display='none'}">
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
        <p>A : ดีเยี่ยม &nbsp; B+ : ดีมาก &nbsp; B : ดี &nbsp; C+ : ค่อนข้างดี &nbsp; C : พอใช้ &nbsp; D+ : อ่อน</p>
        <p>D : อ่อนมาก &nbsp; F : ตก &nbsp; S : พึงพอใจ &nbsp; U : ไม่พึงพอใจ &nbsp; CP : เทียบโอน &nbsp; AU : ไม่นับหน่วยกิต</p>
        <p>W : ถอนรายวิชา &nbsp; I : ยังไม่สมบูรณ์ &nbsp; E : มีเงื่อนไข &nbsp; P : ยังไม่สิ้นสุด &nbsp; X : ยังไม่ส่งเกรด</p>
      </div>
      <div class="mt-10 flex justify-end">
        <div class="text-center text-xs text-gray-700" style="min-width:220px;">
          <p>รัชฎาพร เขษมโตมณี</p>
          <p>หัวหน้างานทะเบียน</p>
        </div>
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
  const key = norm(studentKey);
  const stu = getDataByType('student').find(s => norm(s.student_id) === key || norm(s.name) === key) || (APP.currentUser.data && (norm(APP.currentUser.data.name) === key || norm(APP.currentUser.data.student_id) === key) ? APP.currentUser.data : null);
  if (!stu) return;
  if (isGraduate(stu)) { await _printOfficialTranscript(stu); return; }
  const stuId = norm(stu.student_id);
  let grades = getDataByType('grade').filter(g => norm(g.student_id) === stuId);
  // Respect current filters (academic year / semester) selected on grades page
  const filterYear = APP.filters.academicYear || '';
  const filterSem = APP.filters.semester || '';
  if (filterYear) grades = grades.filter(g => norm(g.academic_year) === filterYear);
  if (filterSem) grades = grades.filter(g => normSem(g.semester) === filterSem);
  const gradeMap = { 'A': 4, 'B+': 3.5, 'B': 3, 'C+': 2.5, 'C': 2, 'D+': 1.5, 'D': 1, 'F': 0 };

  const semesters = {};
  grades.forEach(g => {
    // key: ปี/ภาค เพื่อให้เรียงตามปีก่อน แล้วค่อยภาค
    const key = `${norm(g.academic_year) || '0000'}/${normSem(g.semester) || '0'}`;
    if (!semesters[key]) semesters[key] = { semester: normSem(g.semester) || '1', year: norm(g.academic_year) || '', grades: [] };
    semesters[key].grades.push(g);
  });
  const semKeys = Object.keys(semesters).sort();
  let totalCreditsAll = 0, totalPointsAll = 0;
  let tableHTML = '';
  semKeys.forEach(key => {
    const sem = semesters[key];
    const semText = sem.semester === '3' ? 'ภาคฤดูร้อน' : 'ภาคการศึกษาที่ ' + sem.semester;
    tableHTML += `<tr style="background:#e8f4fd"><td colspan="4" style="padding:4px 8px;font-weight:bold;text-align:center;font-size:12px;border:1px solid #999">${semText} ปีการศึกษา ${sem.year}</td></tr>`;
    let semCredits = 0, semPoints = 0;
    sem.grades.forEach(g => {
      const cr = Number(_gradeCredits(g)) || 0; const gv = gradeMap[g.grade];
      tableHTML += `<tr><td style="padding:3px 8px;font-size:11px;border:1px solid #999;font-family:monospace">${g.subject_code || ''}</td><td style="padding:3px 8px;font-size:11px;border:1px solid #999">${g.subject_name || ''}</td><td style="padding:3px 8px;font-size:11px;text-align:center;border:1px solid #999">${_gradeCreditCode(g) || ''}</td><td style="padding:3px 8px;font-size:11px;text-align:center;font-weight:bold;border:1px solid #999">${g.grade || ''}</td></tr>`;
      if (gv !== undefined) { semPoints += gv * cr; semCredits += cr; }
    });
    const semGpa = semCredits ? (semPoints / semCredits).toFixed(2) : 'N/A';
    totalCreditsAll += semCredits; totalPointsAll += semPoints;
    tableHTML += `<tr style="background:#f9fafb"><td colspan="2" style="padding:3px 8px;font-size:11px;text-align:right;font-weight:600;border:1px solid #999">จำนวนหน่วยกิตรวม: ${semCredits}</td><td colspan="2" style="padding:3px 8px;font-size:11px;text-align:right;font-weight:600;border:1px solid #999">คะแนนเฉลี่ย: ${semGpa}</td></tr>`;
  });
  const gpax = totalCreditsAll ? (totalPointsAll / totalCreditsAll).toFixed(2) : 'N/A';

  let logoBase64 = '';
  const _logoSrcs = (window.LOGO_SOURCES || ['https://cdn.jsdelivr.net/gh/JOB-BCNB-P/LOGO/Logo%20Thai.png']);
  for (const _src of _logoSrcs) {
    try { const resp = await fetch(_src); if (resp.ok) { const blob = await resp.blob(); logoBase64 = await new Promise(r => { const fr = new FileReader(); fr.onload = () => r(fr.result); fr.readAsDataURL(blob); }); break; } } catch (e) { }
  }

  const studentProgram = stu.program || 'หลักสูตรพยาบาลศาสตรบัณฑิต';
  const studentLevel = stu.level || 'ปริญญาตรี';

  const htmlContent = `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>ใบรายงานผลการเรียน - ${stu.name || ''}</title>
    <style>
      @import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;600;700&display=swap');
      *{font-family:'Sarabun',sans-serif;margin:0;padding:0;box-sizing:border-box}
      html,body{position:relative}
      body{padding:20px 40px;font-size:12px;color:#333;min-height:100vh}
      @media print{body{padding:15px 30px}@page{size:A4;margin:15mm 20mm}}
      table{width:100%;border-collapse:collapse}
      .content{position:relative}
      .signature{margin-top:60px;display:flex;justify-content:flex-end}
      .signature-box{text-align:center;min-width:220px;font-size:12px}
    </style></head><body>
    <div class="content">
    <div style="text-align:center;margin-bottom:12px">${logoBase64 ? `<img src="${logoBase64}" style="width:55px;height:auto;margin-bottom:4px">` : ''}<div style="font-weight:700;font-size:14px">${APP.config.college_name}</div><div style="font-size:12px;color:#555">ใบรายงานผลการเรียนนักศึกษารายภาคการศึกษา</div><div style="font-size:11px;color:#555">${studentProgram} ระดับ ${studentLevel}</div>${(filterYear || filterSem) ? `<div style="font-size:11px;color:#1e3a8a;font-weight:600;margin-top:2px">เฉพาะ${filterSem ? ' ภาคการศึกษาที่ ' + filterSem : ''}${filterYear ? ' ปีการศึกษา ' + filterYear : ''}</div>` : ''}</div>
    <div style="display:flex;justify-content:space-between;margin-bottom:10px;font-size:11px"><div>รหัสนักศึกษา: <strong>${stu.student_id || ''}</strong></div><div>ชื่อ-สกุล: <strong>${stu.name || ''}</strong></div></div>
    <table><thead><tr style="background:#e8f4fd"><th style="padding:5px 8px;text-align:left;font-size:11px;border:1px solid #999;width:22%">รหัสวิชา</th><th style="padding:5px 8px;text-align:left;font-size:11px;border:1px solid #999;width:48%">รายวิชา</th><th style="padding:5px 8px;text-align:center;font-size:11px;border:1px solid #999;width:15%">หน่วยกิต</th><th style="padding:5px 8px;text-align:center;font-size:11px;border:1px solid #999;width:15%">ระดับคะแนน</th></tr></thead><tbody>${tableHTML}</tbody></table>
    <div style="margin-top:10px;border:1px solid #999;padding:6px 10px;font-size:11px"><div style="display:flex;justify-content:space-between"><span>รวมหน่วยกิตตลอดปีการศึกษา: <strong>${totalCreditsAll}</strong></span><span>คะแนนเฉลี่ยตลอดปีการศึกษา: <strong>${gpax}</strong></span></div><div style="display:flex;justify-content:space-between;margin-top:3px"><span>รวมหน่วยกิตสะสมตลอดหลักสูตร: <strong>${totalCreditsAll}</strong></span><span>คะแนนเฉลี่ยสะสมตลอดหลักสูตร: <strong>${gpax}</strong></span></div></div>
    <div style="margin-top:10px;font-size:10px;color:#666"><div style="font-weight:600;margin-bottom:3px">หมายเหตุ</div><div>A : ดีเยี่ยม &nbsp; B+ : ดีมาก &nbsp; B : ดี &nbsp; C+ : ค่อนข้างดี &nbsp; C : พอใช้ &nbsp; D+ : อ่อน</div><div>D : อ่อนมาก &nbsp; F : ตก &nbsp; S : พึงพอใจ &nbsp; U : ไม่พึงพอใจ &nbsp; CP : เทียบโอน &nbsp; AU : ไม่นับหน่วยกิต</div><div>W : ถอนรายวิชา &nbsp; I : ยังไม่สมบูรณ์ &nbsp; E : มีเงื่อนไข &nbsp; P : ยังไม่สิ้นสุด &nbsp; X : ยังไม่ส่งเกรด</div></div>
    <div class="signature"><div class="signature-box"><div>รัชฎาพร เขษมโตมณี</div><div>หัวหน้างานทะเบียน</div></div></div>
    </div>
    <script>window.onload=function(){window.print()}<\/script></body></html>`;
  const w = window.open('', '_blank');
  if (w) { w.document.write(htmlContent); w.document.close(); } else { showToast('กรุณาอนุญาต Popup เพื่อดาวน์โหลด PDF', 'error'); }
}

// ======================== ENG RESULTS ========================
// แผงวิเคราะห์ผลสอบภาษาอังกฤษ (PBRI) — กราฟแท่งระดับ + กราฟวงกลมผ่าน/ไม่ผ่าน
// เลือกดูได้ตาม ชั้นปี / ครั้งที่สอบ / ปีการศึกษา (self-contained ในการ์ด)
function engAnalyticsHTML() {
  const allEngRaw = getDataByType('eng_result');
  const students = getDataByType('student');
  const stuById = {}; students.forEach(s => stuById[norm(s.student_id)] = s);

  const selYear = APP.filters._engYear || '';
  const selYrLevel = APP.filters._engYearLevel || '';
  const selAttempt = APP.filters._engAttempt || '';

  // ตัวเลือก
  const engYears = [...new Set(allEngRaw.map(e => norm(e.academic_year)).filter(Boolean))].sort().reverse();
  const attempts = [...new Set(allEngRaw.map(e => norm(e.eng_attempt)).filter(Boolean))].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));

  // ขอบเขตนักศึกษา
  let scopeStudents;
  if (selYrLevel === '__grad') scopeStudents = students.filter(isGraduate);
  else { scopeStudents = activeStudents(students); if (selYrLevel) scopeStudents = scopeStudents.filter(s => norm(s.year_level) === selYrLevel); }
  const scopeIds = new Set(scopeStudents.map(s => norm(s.student_id)));

  // กรองผลสอบตามปี/ครั้งที่/ขอบเขตนักศึกษา
  let eng = allEngRaw.slice();
  if (selYear) eng = eng.filter(e => norm(e.academic_year) === selYear);
  if (selAttempt) eng = eng.filter(e => norm(e.eng_attempt) === selAttempt);
  const engScope = eng.filter(e => scopeIds.has(norm(e.student_id)));

  // --- กราฟแท่ง: ระดับผลสอบ PBRI (เฉพาะข้อสอบ สบช.) ---
  const levelDefs = [
    { key: 'Beginner', label: 'Beginner (0–20)', color: '#ef4444' },
    { key: 'Elementary', label: 'Elementary (21–40)', color: '#f97316' },
    { key: 'Intermediate', label: 'Intermediate (41–60)', color: '#f59e0b' },
    { key: 'Upper Intermediate', label: 'Upper-Intermediate (61–80)', color: '#84cc16' },
    { key: 'Advanced', label: 'Advanced (81–90)', color: '#22c55e' },
    { key: 'Proficiency', label: 'Proficiency (91–100)', color: '#0ea5e9' }
  ];
  const levelCount = {}; levelDefs.forEach(l => levelCount[l.key] = 0);
  const pbriRecs = engScope.filter(e => norm(e.eng_type) === 'สบช.' && norm(e.eng_status) !== 'ไม่เข้าสอบ');
  pbriRecs.forEach(e => { const lv = getEngLevel(Number(e.eng_score) || 0); if (levelCount[lv] !== undefined) levelCount[lv]++; });
  const barItems = animBarRows(levelDefs.map(l => ({ label: l.label, value: levelCount[l.key], color: l.color })));

  // --- กราฟวงกลม: ผ่าน/ไม่ผ่าน + แยกประเภทการผ่าน ---
  const passedPBRI = new Set(engScope.filter(e => norm(e.eng_status) === 'ผ่าน' && norm(e.eng_type) === 'สบช.').map(e => norm(e.student_id)));
  const passedExt = new Set(engScope.filter(e => norm(e.eng_status) === 'ผ่าน' && norm(e.eng_type) !== 'สบช.').map(e => norm(e.student_id)));
  let cPBRI = 0, cExt = 0, cFail = 0;
  scopeStudents.forEach(s => { const id = norm(s.student_id); if (passedPBRI.has(id)) cPBRI++; else if (passedExt.has(id)) cExt++; else cFail++; });
  const passTotal = cPBRI + cExt;
  const passPct = scopeStudents.length ? Math.round(passTotal / scopeStudents.length * 1000) / 10 : 0;
  const donut = svgDonut([
    { label: 'ผ่าน — สอบ PBRI (สบช.)', value: cPBRI, color: '#22c55e' },
    { label: 'ผ่าน — สอบภายนอกสถาบัน', value: cExt, color: '#0ea5e9' },
    { label: 'ยังไม่ผ่าน', value: cFail, color: '#ef4444' }
  ], 'คน');

  // ตัวกรอง
  const yrBtns = `<div class="flex flex-wrap gap-1.5">
    <button onclick="APP.filters._engYearLevel='';renderCurrentPage()" class="px-3 py-1.5 rounded-lg text-xs font-medium ${selYrLevel === '' ? 'bg-primary text-white' : 'bg-gray-100 text-gray-600'}">ทุกชั้นปี</button>
    ${['1', '2', '3', '4'].map(y => `<button onclick="APP.filters._engYearLevel='${y}';renderCurrentPage()" class="px-3 py-1.5 rounded-lg text-xs font-medium ${selYrLevel === y ? 'bg-primary text-white' : 'bg-gray-100 text-gray-600'}">ปี ${y}</button>`).join('')}
  </div>`;
  const yearSel = `<select onchange="APP.filters._engYear=this.value;renderCurrentPage()" class="border border-gray-200 rounded-lg px-2 py-1.5 text-xs bg-white"><option value="">ทุกปีการศึกษา</option>${engYears.map(y => `<option value="${y}" ${selYear === y ? 'selected' : ''}>ปี ${y}</option>`).join('')}</select>`;
  const attemptSel = `<select onchange="APP.filters._engAttempt=this.value;renderCurrentPage()" class="border border-gray-200 rounded-lg px-2 py-1.5 text-xs bg-white"><option value="">ทุกครั้งที่สอบ</option>${attempts.map(a => `<option value="${a}" ${selAttempt === a ? 'selected' : ''}>ครั้งที่ ${a}</option>`).join('')}</select>`;

  return `<div class="bg-white rounded-2xl border border-blue-100 p-5 mb-4">
    <div class="flex flex-wrap items-center justify-between gap-3 mb-4">
      <h3 class="font-bold text-gray-800 flex items-center gap-2"><i data-lucide="bar-chart-3" class="w-5 h-5 text-primary"></i>วิเคราะห์ผลสอบภาษาอังกฤษ (PBRI)</h3>
      <div class="flex flex-wrap items-center gap-2">${yearSel}${attemptSel}</div>
    </div>
    <div class="mb-4">${yrBtns}</div>
    <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
      <div>
        <p class="text-sm font-semibold text-gray-600 mb-3">ระดับผลสอบ PBRI (สบช.) — ${pbriRecs.length} ครั้ง</p>
        <div class="space-y-2.5">${barItems}</div>
      </div>
      <div>
        <p class="text-sm font-semibold text-gray-600 mb-3">สัดส่วนการสอบผ่าน <span class="font-normal text-green-600">(ผ่าน ${passPct}% ของ ${scopeStudents.length} คน)</span></p>
        ${donut}
      </div>
    </div>
  </div>`;
}

function engResultsPage() {
  const isAdmin = isAdminRole();
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
    // ผลสอบภาษาอังกฤษ (อาจารย์/อาจารย์ประจำชั้น): แสดงเฉพาะนักศึกษาที่กำลังศึกษา
    if (APP.currentRole === 'teacher' || APP.currentRole === 'classTeacher') {
      studentList = activeStudents(studentList);
    }

    // Year level + Advisor filter for admin/academic/executive
    let advisorSelector = '';
    if (canFilterByAdvisor) {
      // Year level filter
      const selectedEngYrLevel = APP.filters._engYearLevel || '';
      const isGradFilter = selectedEngYrLevel === '__grad';
      if (isGradFilter) {
        studentList = studentList.filter(s => isGraduate(s));
      } else if (selectedEngYrLevel) {
        studentList = studentList.filter(s => norm(s.year_level) === selectedEngYrLevel);
      }
      // โหมดที่ไม่ใช่ผู้สำเร็จการศึกษา → แสดงเฉพาะนักศึกษาที่กำลังศึกษา
      if (!isGradFilter) studentList = activeStudents(studentList);
      const yearLevelSelector = `<div class="bg-white rounded-2xl p-4 border border-blue-100 mb-4">
        <label class="block text-sm font-medium text-gray-700 mb-2"><i data-lucide="layers" class="w-4 h-4 inline mr-1"></i>กรองตามชั้นปี</label>
        <div class="flex flex-wrap gap-2">
          <button onclick="APP.filters._engYearLevel='';APP.filters._engStudent='';APP.filters._engSearch='';APP.pagination.page=1;renderCurrentPage()" class="px-4 py-2 rounded-xl text-sm font-medium ${selectedEngYrLevel === '' ? 'bg-primary text-white' : 'bg-gray-100 text-gray-600 hover:bg-gray-200'}">ทุกชั้นปี</button>
          ${['1', '2', '3', '4'].map(yr => `<button onclick="APP.filters._engYearLevel='${yr}';APP.filters._engStudent='';APP.filters._engSearch='';APP.pagination.page=1;renderCurrentPage()" class="px-4 py-2 rounded-xl text-sm font-medium ${selectedEngYrLevel === yr ? 'bg-primary text-white' : 'bg-gray-100 text-gray-600 hover:bg-gray-200'}">ชั้นปี ${yr}</button>`).join('')}
          <button onclick="APP.filters._engYearLevel='__grad';APP.filters._engStudent='';APP.filters._engSearch='';APP.filters._engAdvisor='';APP.pagination.page=1;renderCurrentPage()" class="px-4 py-2 rounded-xl text-sm font-medium ${isGradFilter ? 'bg-primary text-white' : 'bg-gray-100 text-gray-600 hover:bg-gray-200'}">ผู้สำเร็จการศึกษา</button>
        </div>
        ${selectedEngYrLevel ? `<p class="text-xs text-gray-500 mt-2"><i data-lucide="info" class="w-3 h-3 inline mr-1"></i>${isGradFilter ? 'แสดงเฉพาะผู้สำเร็จการศึกษา' : 'แสดงเฉพาะนักศึกษาชั้นปีที่ ' + selectedEngYrLevel} (${studentList.length} คน)</p>` : ''}
      </div>`;

      // กรองตามรุ่น — แสดงเฉพาะตอนเลือก "ผู้สำเร็จการศึกษา"
      const gradStudentsAll = getDataByType('student').filter(s => isGraduate(s));
      const allBatches = [...new Set(gradStudentsAll.map(s => norm(s.batch)).filter(Boolean))].sort((a, b) => b.localeCompare(a, undefined, { numeric: true }));
      const selectedBatch = APP.filters._engBatch || '';
      if (isGradFilter && selectedBatch) studentList = gradStudentsAll.filter(s => norm(s.batch) === selectedBatch);
      const batchSelector = isGradFilter ? `<div class="bg-white rounded-2xl p-4 border border-blue-100 mb-4">
        <label class="block text-sm font-medium text-gray-700 mb-2"><i data-lucide="users" class="w-4 h-4 inline mr-1"></i>กรองตามรุ่น <span class="font-normal text-gray-400 text-xs">(เฉพาะผู้สำเร็จการศึกษา)</span></label>
        <select onchange="APP.filters._engBatch=this.value;APP.filters._engStudent='';APP.filters._engSearch='';APP.pagination.page=1;renderCurrentPage()" class="w-full border border-gray-200 rounded-xl px-4 py-2.5 text-sm">
          <option value="">-- ทุกรุ่น --</option>
          ${allBatches.map(b => `<option value="${b}" ${selectedBatch === b ? 'selected' : ''}>รุ่นที่ ${b}</option>`).join('')}
        </select>
        ${selectedBatch ? `<p class="text-xs text-gray-500 mt-2"><i data-lucide="info" class="w-3 h-3 inline mr-1"></i>แสดงเฉพาะผู้สำเร็จการศึกษา รุ่นที่ ${selectedBatch} (${studentList.length} คน)</p>` : ''}
      </div>` : '';

      const allAdvisors = [...new Set(studentList.map(s => s.advisor).filter(Boolean))].sort();
      const selectedAdvisor = APP.filters._engAdvisor || '';
      if (!isGradFilter && selectedAdvisor) {
        studentList = studentList.filter(s => (s.advisor || '') === selectedAdvisor);
      }
      // ผู้สำเร็จการศึกษา: ไม่แสดงตัวกรองอาจารย์ที่ปรึกษา
      const advisorDiv = isGradFilter ? '' : `<div class="bg-white rounded-2xl p-4 border border-blue-100 mb-4">
        <label class="block text-sm font-medium text-gray-700 mb-2"><i data-lucide="user-check" class="w-4 h-4 inline mr-1"></i>กรองตามอาจารย์ที่ปรึกษา</label>
        <select onchange="APP.filters._engAdvisor=this.value;APP.filters._engStudent='';APP.filters._engSearch='';APP.pagination.page=1;renderCurrentPage()" class="w-full border border-gray-200 rounded-xl px-4 py-2.5 text-sm">
          <option value="">-- แสดงนักศึกษาทั้งหมด --</option>
          ${allAdvisors.map(a => `<option value="${a}" ${selectedAdvisor === a ? 'selected' : ''}>${a}</option>`).join('')}
        </select>
        ${selectedAdvisor ? `<p class="text-xs text-gray-500 mt-2"><i data-lucide="info" class="w-3 h-3 inline mr-1"></i>แสดงเฉพาะนักศึกษาในความดูแลของ ${selectedAdvisor} (${studentList.length} คน)</p>` : ''}
      </div>`;
      advisorSelector = `${yearLevelSelector}${batchSelector}${advisorDiv}`;
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
    // Get students list for summary — นับเฉพาะนักศึกษาที่กำลังศึกษาอยู่เท่านั้น
    let summaryStudents = activeStudents(getDataByType('student'));
    if (APP.currentRole === 'classTeacher') {
      const yr = APP.currentUser.responsible_year || '1';
      summaryStudents = summaryStudents.filter(s => norm(s.year_level) === norm(yr));
    }
    if (APP.currentRole === 'teacher') {
      summaryStudents = summaryStudents.filter(s => s.advisor === APP.currentUser.name);
    }
    const selectedAdvisor = APP.filters._engAdvisor || '';
    // สรุปผลให้สอดคล้องกับตัวกรองที่เลือก (ชั้นปี / ผู้สำเร็จการศึกษา / รุ่น / อาจารย์ที่ปรึกษา)
    const sumYrLevel = canFilterByAdvisor ? (APP.filters._engYearLevel || '') : '';
    const sumIsGrad = sumYrLevel === '__grad';
    const sumBatch = APP.filters._engBatch || '';
    if (canFilterByAdvisor) {
      if (sumIsGrad) {
        summaryStudents = getDataByType('student').filter(s => isGraduate(s));
        if (sumBatch) summaryStudents = summaryStudents.filter(s => norm(s.batch) === sumBatch);
      } else {
        if (selectedAdvisor) summaryStudents = summaryStudents.filter(s => (s.advisor || '') === selectedAdvisor);
        if (sumYrLevel) summaryStudents = summaryStudents.filter(s => norm(s.year_level) === sumYrLevel);
      }
    } else if (selectedAdvisor) {
      summaryStudents = summaryStudents.filter(s => (s.advisor || '') === selectedAdvisor);
    }
    const scopeLabel = sumIsGrad ? ('ผู้สำเร็จการศึกษา' + (sumBatch ? ' รุ่นที่ ' + sumBatch : '')) : (sumYrLevel ? 'ชั้นปีที่ ' + sumYrLevel : '');

    // Count unique students who passed / not passed
    const passedIds = new Set(allEng.filter(e => e.eng_status === 'ผ่าน').map(e => norm(e.student_id)));
    const passedCount = summaryStudents.filter(s => passedIds.has(norm(s.student_id))).length;
    const notPassedCount = summaryStudents.length - passedCount;

    // การ์ดสรุปแยกรายชั้นปี (แสดงเมื่อยังไม่เลือกชั้นปีเจาะจง สำหรับ admin/ผู้บริหาร)
    let perYearCardsHtml = '';
    if (canFilterByAdvisor && !sumYrLevel) {
      const baseAll = activeStudents(getDataByType('student'));
      const yrCards = [1, 2, 3, 4].map(yr => {
        const ys = baseAll.filter(s => norm(s.year_level) === String(yr));
        const p = ys.filter(s => passedIds.has(norm(s.student_id))).length;
        const f = ys.length - p;
        return `<div class="bg-gray-50 rounded-xl p-3 border border-gray-100">
          <p class="text-sm font-semibold text-gray-700 mb-1">ชั้นปี ${yr} <span class="font-normal text-gray-400">(${ys.length} คน)</span></p>
          <div class="flex gap-2 text-sm"><span class="px-2 py-0.5 bg-green-50 text-green-700 rounded-lg">ผ่าน <b>${p}</b></span><span class="px-2 py-0.5 bg-red-50 text-red-700 rounded-lg">ไม่ผ่าน <b>${f}</b></span></div>
        </div>`;
      }).join('');
      perYearCardsHtml = `<div class="mt-4 pt-4 border-t border-gray-100"><p class="text-xs font-semibold text-gray-500 mb-2"><i data-lucide="layers" class="w-3.5 h-3.5 inline mr-1"></i>แยกรายชั้นปี</p><div class="grid grid-cols-2 lg:grid-cols-4 gap-2">${yrCards}</div></div>`;
    }

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
      <h3 class="font-bold text-gray-800 mb-3 flex items-center gap-2"><i data-lucide="bar-chart-3" class="w-5 h-5 text-primary"></i>สรุปผลสอบภาษาอังกฤษ${scopeLabel ? ` <span class="text-sm font-normal text-gray-500">— ${scopeLabel}</span>` : ''}</h3>
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
      ${perYearCardsHtml}
    </div>`;
  }

  return `<div class="flex flex-wrap items-center justify-between gap-3 mb-4">
    <h2 class="text-xl font-bold text-gray-800"><i data-lucide="languages" class="w-6 h-6 inline mr-2"></i>ผลสอบภาษาอังกฤษ</h2>
    ${isAdmin ? `<div class="flex gap-2"><button onclick="showAddEngModal()" class="flex items-center gap-2 px-4 py-2 bg-primary text-white rounded-xl hover:bg-primaryDark text-sm"><i data-lucide="plus" class="w-4 h-4"></i>เพิ่มผลสอบ</button>${csvUploadBtn('eng_result', 'student_id,eng_score,eng_type,eng_attempt,eng_date,eng_status,academic_year')}</div>` : ''}
  </div>
  ${['admin', 'academic', 'registrar', 'executive'].includes(APP.currentRole) ? engAnalyticsHTML() : ''}
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
// เกณฑ์ผ่านสอบภาษาอังกฤษ (สบช.):
//   นักศึกษารุ่นที่ 81 เป็นต้นไป หรือ ปีการศึกษา 2569 เป็นต้นไป → ผ่านเมื่อคะแนน ≥ 51
//   รุ่น/ปีก่อนหน้านั้น → ใช้เกณฑ์เดิม ผ่านเมื่อคะแนน ≥ 41
function engIsNewCriterion(studentId, academicYear) {
  const stu = getDataByType('student').find(s => norm(s.student_id) === norm(studentId));
  const b = parseInt(norm(stu && stu.batch), 10);
  const y = parseInt(norm(academicYear), 10);
  return (!isNaN(b) && b >= 81) || (!isNaN(y) && y >= 2569);
}
function engPassThreshold(studentId, academicYear) {
  return engIsNewCriterion(studentId, academicYear) ? 51 : 41;
}

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
  const sid = (document.querySelector('#' + prefix + 'EngForm [name="student_id"]') || {}).value || '';
  const yr = (document.getElementById(prefix + 'EngYear') || {}).value || '';
  const th = engPassThreshold(sid, yr);
  if (levelEl) levelEl.value = total > 0 ? getEngLevel(total) : '';
  if (statusEl) {
    if (total > 0) {
      const pass = total >= th;
      statusEl.textContent = (pass ? 'ผ่าน' : 'ไม่ผ่าน') + ' (เกณฑ์ ≥ ' + th + ')';
      statusEl.className = 'inline-block px-3 py-1 rounded-full text-sm font-semibold ' + (pass ? 'bg-green-100 text-green-700' : 'bg-red-100 text-red-700');
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
      <div><label class="block text-xs text-gray-600 mb-1">นักศึกษา * (พิมพ์รหัสหรือเลือก)</label><input list="addEngStudentList" name="student_id" required oninput="updateEngSbchLevelStatus('add')" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="พิมพ์รหัสนักศึกษา...">${studentDatalistHTML('addEngStudentList')}</div>
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-xs text-gray-600 mb-1">รูปแบบการสอบ *</label>
          <select id="addEngType" onchange="updateEngTypeForm('add')" class="w-full border rounded-xl px-3 py-2 text-sm">${engTypeOptions('')}</select>
        </div>
        <div><label class="block text-xs text-gray-600 mb-1">สอบครั้งที่</label><input id="addEngAttempt" type="number" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="เช่น 1, 2"></div>
        <div><label class="block text-xs text-gray-600 mb-1">วันที่สอบ <span class="text-gray-400">(วว/ดด/ปปปป พ.ศ. หรือ ค.ศ.)</span></label><input id="addEngDate" type="text" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="เช่น 05/04/2568 หรือ 05/04/2025"></div>
        <div><label class="block text-xs text-gray-600 mb-1">ปีการศึกษา</label><input id="addEngYear" type="text" oninput="updateEngSbchLevelStatus('add')" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="เช่น 2568"></div>
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
          <div><label class="block text-xs text-gray-600 mb-1">คะแนนรวม</label><input id="addEngOtherScore" type="number" step="any" min="0" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="กรอกทศนิยมได้ เช่น 85.5"></div>
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
        obj.eng_status = total >= engPassThreshold(studentId, year) ? 'ผ่าน' : 'ไม่ผ่าน';
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
  const isAdmin = isAdminRole();
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
      applyTrackingBackfill(obj);
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
  const isAdmin = isAdminOnlyRole();
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
      const st = t.teacher_status || 'ปฏิบัติงานอยู่'; const stColor = st === 'ปฏิบัติงานอยู่' ? 'bg-green-100 text-green-700' : st === 'ลาออก' ? 'bg-red-100 text-red-700' : 'bg-yellow-100 text-yellow-700'; return `<tr class="border-t hover:bg-gray-50">
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
    const st = t.teacher_status || 'ปฏิบัติงานอยู่'; const stColor = st === 'ปฏิบัติงานอยู่' ? 'bg-green-100 text-green-700' : st === 'ลาออก' ? 'bg-red-100 text-red-700' : 'bg-yellow-100 text-yellow-700'; return `<tr class="border-t hover:bg-gray-50">
        <td class="px-4 py-3 font-medium">${t.name || ''}</td><td class="px-4 py-3">${t.position || ''}</td>
        <td class="px-4 py-3">${t.department || ''}</td>
        <td class="px-4 py-3"><span class="px-2 py-1 rounded-full text-xs ${stColor}">${st}</span></td>
        <td class="px-4 py-3"><div class="flex gap-1"><button onclick="showTeacherDetail('${t.__backendId}')" class="text-gray-400 hover:text-primary" title="ดูข้อมูล"><i data-lucide="eye" class="w-4 h-4"></i></button>${isAdmin ? `<button onclick="showEditTeacherModal('${t.__backendId}')" class="text-blue-400 hover:text-blue-600" title="แก้ไข"><i data-lucide="pencil" class="w-4 h-4"></i></button><button onclick="deleteRecord('${t.__backendId}')" class="text-red-400 hover:text-red-600" title="ลบ"><i data-lucide="trash-2" class="w-4 h-4"></i></button>` : ''}</div></td></tr>`
  }).join('') : '<tr><td colspan="5" class="px-4 py-8 text-center text-gray-400">ไม่มีข้อมูล</td></tr>'}</tbody>
    </table></div>
  </div>
  ${paginationHTML(total, APP.pagination.perPage, APP.pagination.page, 'changePage')}`;
}

// ======================== ข้อมูลอาจารย์ที่ปรึกษา (ระบบทะเบียน) ========================
// รวมข้อมูลอาจารย์ที่ปรึกษา → จำนวน/รายชื่อนักศึกษาในความดูแล
// + ลิงก์เชื่อมไปยัง ผลการเรียน / ผลสอบภาษาอังกฤษ / ข้อมูลนักศึกษา
function advisorInfoPage() {
  const isAdmin = isAdminRole();
  const students = getDataByType('student');
  const teachers = getDataByType('teacher');
  const NO_DEPT = 'ไม่ระบุสาขาวิชา';

  // ----- จับคู่ชื่ออาจารย์ที่ปรึกษากับทะเบียนอาจารย์ (ยืดหยุ่นเรื่องเว้นวรรค/คำนำหน้า) -----
  const TITLES = ['ผศ.ดร.', 'รศ.ดร.', 'ศ.ดร.', 'ผศ.', 'รศ.', 'ศ.', 'ดร.', 'อ.', 'อาจารย์', 'ว่าที่ร.ต.', 'ว่าที่ร้อยตรี', 'น.ส.', 'นางสาว', 'นาง', 'นาย', 'นพ.', 'พญ.'];
  const nameKey = v => norm(v).replace(/\s+/g, '');
  const stripTitle = k => { let n = k, go = true; while (go) { go = false; for (const pr of TITLES) { const pk = pr.replace(/\s+/g, ''); if (pk && n.startsWith(pk)) { n = n.slice(pk.length); go = true; break; } } } return n; };
  // ดัชนีอาจารย์: ทั้งคีย์เต็มและคีย์ที่ตัดคำนำหน้าออก เพื่อจับคู่แม้พิมพ์เว้นวรรค/คำนำหน้าต่างกัน
  const teacherIdx = {};
  teachers.forEach(t => { const k = nameKey(t.name); if (!k) return; teacherIdx[k] = t; const sk = stripTitle(k); if (sk && !(sk in teacherIdx)) teacherIdx[sk] = t; });
  const findTeacher = adv => { const k = nameKey(adv); return teacherIdx[k] || teacherIdx[stripTitle(k)] || null; };

  // รวมกลุ่มนักศึกษาตามอาจารย์ที่ปรึกษา (รวมชื่อที่พิมพ์เว้นวรรคต่างกันให้เป็นกลุ่มเดียว)
  const advisorMap = {};
  students.forEach(s => {
    const adv = norm(s.advisor);
    if (!adv) return;
    const t = findTeacher(adv);
    const key = t ? nameKey(t.name) : nameKey(adv);
    if (!advisorMap[key]) {
      advisorMap[key] = { key: key, name: t ? norm(t.name) : adv, dept: t ? (norm(t.department) || NO_DEPT) : NO_DEPT, phone: t ? norm(t.phone) : '', email: t ? norm(t.email) : '', students: [] };
    }
    advisorMap[key].students.push(s);
  });
  let advisors = Object.values(advisorMap).sort((a, b) => a.name.localeCompare(b.name, 'th'));

  const selectedDept = APP.filters._advisorDept || '';
  const selectedAdvisor = APP.filters._advisorSelected || '';

  // ---------- มุมมองรายละเอียดอาจารย์ที่เลือก ----------
  if (selectedAdvisor && (advisorMap[selectedAdvisor] || teacherIdx[selectedAdvisor])) {
    let a = advisorMap[selectedAdvisor];
    if (!a) { const _t = teacherIdx[selectedAdvisor]; a = { key: selectedAdvisor, name: norm(_t.name), dept: norm(_t.department) || NO_DEPT, phone: norm(_t.phone), email: norm(_t.email), students: [] }; }
    const list = a.students.slice().sort((x, y) => norm(x.year_level).localeCompare(norm(y.year_level)) || norm(x.name).localeCompare(norm(y.name), 'th'));
    const activeCount = activeStudents(a.students).length;
    const totalCount = a.students.length;
    const rows = list.map(s => {
      const sid = (s.student_id || s.name || '').replace(/'/g, "\\'");
      const stColor = s.status === 'กำลังศึกษา' ? 'bg-green-100 text-green-700' : s.status === 'สำเร็จการศึกษา' ? 'bg-blue-100 text-blue-700' : s.status === 'ลาออก' ? 'bg-red-100 text-red-700' : s.status === 'ขอโอนย้ายสถานศึกษา' ? 'bg-orange-100 text-orange-700' : 'bg-yellow-100 text-yellow-700';
      return `<tr class="border-t hover:bg-gray-50">
        <td class="px-4 py-3">${s.student_id || ''}</td>
        <td class="px-4 py-3 font-medium">${studentDisplayName(s)}</td>
        <td class="px-4 py-3">${s.year_level || ''}</td>
        <td class="px-4 py-3">${s.batch || ''}</td>
        <td class="px-4 py-3"><span class="px-2 py-1 rounded-full text-xs ${stColor}">${s.status || ''}</span></td>
        <td class="px-4 py-3">
          <div class="flex flex-wrap gap-1">
            <button onclick="showStudentDetail('${s.__backendId}')" class="inline-flex items-center gap-1 px-2 py-1 rounded-lg text-xs bg-gray-100 text-gray-600 hover:bg-gray-200" title="ข้อมูลนักศึกษา"><i data-lucide="user" class="w-3.5 h-3.5"></i>ข้อมูล</button>
            ${isAdmin ? `<button onclick="advisorRemoveStudent('${s.__backendId}')" class="inline-flex items-center gap-1 px-2 py-1 rounded-lg text-xs bg-red-50 text-red-600 hover:bg-red-100" title="นำออกจากความดูแล"><i data-lucide="user-minus" class="w-3.5 h-3.5"></i>นำออก</button>` : ''}
            <button onclick="advisorGotoGrades('${sid}')" class="inline-flex items-center gap-1 px-2 py-1 rounded-lg text-xs bg-blue-50 text-blue-600 hover:bg-blue-100" title="ผลการเรียน"><i data-lucide="graduation-cap" class="w-3.5 h-3.5"></i>ผลการเรียน</button>
            <button onclick="advisorGotoEng('${sid}')" class="inline-flex items-center gap-1 px-2 py-1 rounded-lg text-xs bg-emerald-50 text-emerald-600 hover:bg-emerald-100" title="ผลสอบภาษาอังกฤษ"><i data-lucide="languages" class="w-3.5 h-3.5"></i>ผลสอบภาษาอังกฤษ</button>
          </div>
        </td>
      </tr>`;
    }).join('');

    window._advisorCtx = { key: a.key, name: a.name, studentIds: a.students.map(x => x.__backendId) };
    return `<div class="flex flex-wrap items-center justify-between gap-3 mb-4">
      <h2 class="text-xl font-bold text-gray-800"><i data-lucide="user-check" class="w-6 h-6 inline mr-2"></i>ข้อมูลอาจารย์ที่ปรึกษา</h2>
      <div class="flex gap-2 flex-wrap">
        ${isAdmin ? `<button onclick="showAdvisorAddStudents()" class="inline-flex items-center gap-1 px-3 py-2 rounded-xl text-sm bg-primary text-white hover:bg-primaryDark"><i data-lucide="user-plus" class="w-4 h-4"></i>เพิ่มนักศึกษาในความดูแล</button>` : ''}
        <button onclick="APP.filters._advisorSelected='';APP.pagination.page=1;renderCurrentPage()" class="inline-flex items-center gap-1 px-3 py-2 rounded-xl text-sm bg-gray-100 text-gray-600 hover:bg-gray-200"><i data-lucide="arrow-left" class="w-4 h-4"></i>เลือกอาจารย์ท่านอื่น</button>
      </div>
    </div>
    <div class="bg-gradient-to-br from-primary to-primaryDark rounded-2xl p-5 text-white mb-4 shadow">
      <div class="flex items-center gap-4">
        <div class="w-14 h-14 bg-white bg-opacity-20 rounded-2xl flex items-center justify-center flex-shrink-0"><i data-lucide="user-check" class="w-7 h-7 text-white"></i></div>
        <div class="flex-1 min-w-0">
          <p class="text-lg font-bold truncate">${a.name}</p>
          <p class="text-sm text-white text-opacity-90 truncate">${a.dept}${a.phone ? ' · โทร. ' + a.phone : ''}${a.email ? ' · ' + a.email : ''}</p>
        </div>
        <div class="text-right flex-shrink-0">
          <p class="text-3xl font-bold">${activeCount}</p>
          <p class="text-xs text-white text-opacity-90">กำลังศึกษา${totalCount > activeCount ? ' / ' + totalCount + ' ทั้งหมด' : ''}</p>
        </div>
      </div>
    </div>
    <div class="bg-white rounded-2xl border border-blue-100 overflow-hidden">
      <div class="overflow-x-auto"><table class="w-full text-sm">
        <thead><tr class="bg-surface text-left"><th class="px-4 py-3 font-semibold">รหัสนักศึกษา</th><th class="px-4 py-3 font-semibold">ชื่อ-สกุล</th><th class="px-4 py-3 font-semibold">ชั้นปี</th><th class="px-4 py-3 font-semibold">รุ่นที่</th><th class="px-4 py-3 font-semibold">สถานภาพ</th><th class="px-4 py-3 font-semibold">ดูข้อมูลเชื่อมโยง</th></tr></thead>
        <tbody>${rows || '<tr><td colspan="6" class="px-4 py-8 text-center text-gray-400">ไม่มีนักศึกษาในความดูแล</td></tr>'}</tbody>
      </table></div>
    </div>`;
  }

  // ---------- มุมมองรายชื่ออาจารย์ที่ปรึกษา ----------
  // ตัวเลือกสาขาวิชา (ดันตัวเลือก "ไม่ระบุสาขาวิชา" ไว้ท้ายสุด)
  const depts = [...new Set(advisors.map(a => a.dept))].sort((a, b) => a === NO_DEPT ? 1 : b === NO_DEPT ? -1 : a.localeCompare(b, 'th'));
  if (selectedDept) advisors = advisors.filter(a => a.dept === selectedDept);
  const searchVal = APP.filters._advisorSearch || '';
  if (searchVal) { const q = searchVal.toLowerCase(); advisors = advisors.filter(a => a.name.toLowerCase().includes(q)); }

  const deptFilter = `<div class="bg-white rounded-2xl p-4 border border-blue-100 mb-4">
    <div class="flex items-center gap-3 flex-wrap">
      <label class="text-sm font-medium text-gray-700"><i data-lucide="filter" class="w-4 h-4 inline mr-1"></i>สาขาวิชา:</label>
      <select onchange="APP.filters._advisorDept=this.value;APP.filters._advisorSearch='';APP.pagination.page=1;renderCurrentPage()" class="border border-gray-200 rounded-xl px-3 py-2 text-sm">
        <option value="">-- ทุกสาขาวิชา --</option>
        ${depts.map(d => `<option value="${d}" ${selectedDept === d ? 'selected' : ''}>${d}</option>`).join('')}
      </select>
      ${selectedDept ? `<span class="text-xs text-gray-500">แสดงสาขา: ${selectedDept}</span>` : ''}
    </div>
  </div>`;

  const searchBox = `<div class="bg-white rounded-2xl p-4 border border-blue-100 mb-4">
    <label class="block text-sm font-medium text-gray-700 mb-2"><i data-lucide="user-search" class="w-4 h-4 inline mr-1"></i>ค้นหาอาจารย์ที่ปรึกษา</label>
    <div class="relative"><i data-lucide="search" class="absolute left-3 top-2.5 w-4 h-4 text-gray-400"></i><input type="text" placeholder="พิมพ์ชื่ออาจารย์ที่ปรึกษา..." value="${searchVal}" class="w-full pl-10 pr-4 py-2 border border-gray-200 rounded-xl text-sm" oninput="clearTimeout(window._advisorSearchTimer);window._advisorSearchTimer=setTimeout(()=>{APP.filters._advisorSearch=this.value;APP.pagination.page=1;renderCurrentPage()},300)"></div>
    <p class="text-xs text-gray-400 mt-2">คลิกการ์ดรายชื่อด้านล่างเพื่อดูนักศึกษาในความดูแล</p>
  </div>`;

  // ----- พาเนล: นักศึกษาที่ยังไม่มีอาจารย์ที่ปรึกษา (เฉพาะที่กำลังศึกษา) -----
  const noAdvisorStudents = activeStudents(students).filter(s => !norm(s.advisor))
    .sort((x, y) => norm(x.year_level).localeCompare(norm(y.year_level)) || norm(x.name).localeCompare(norm(y.name), 'th'));
  const noAdvRows = noAdvisorStudents.map(s => `<tr class="border-t border-amber-100 hover:bg-amber-50">
      <td class="px-4 py-2.5">${s.student_id || ''}</td>
      <td class="px-4 py-2.5 font-medium">${s.name || ''}</td>
      <td class="px-4 py-2.5">${s.year_level || ''}</td>
      <td class="px-4 py-2.5">${s.batch || ''}</td>
      <td class="px-4 py-2.5"><div class="flex flex-wrap gap-1">
        <button onclick="showStudentDetail('${s.__backendId}')" class="inline-flex items-center gap-1 px-2 py-1 rounded-lg text-xs bg-gray-100 text-gray-600 hover:bg-gray-200" title="ข้อมูลนักศึกษา"><i data-lucide="user" class="w-3.5 h-3.5"></i>ข้อมูล</button>
        ${isAdmin ? `<button onclick="showAssignAdvisorToStudent('${s.__backendId}')" class="inline-flex items-center gap-1 px-2 py-1 rounded-lg text-xs bg-primary text-white hover:bg-primaryDark" title="กำหนดอาจารย์ที่ปรึกษา"><i data-lucide="user-plus" class="w-3.5 h-3.5"></i>กำหนดที่ปรึกษา</button>` : ''}
      </div></td>
    </tr>`).join('');
  const noAdvisorPanel = `<details class="bg-amber-50 border border-amber-200 rounded-2xl mb-4">
    <summary class="cursor-pointer px-4 py-3 text-sm font-semibold text-amber-800 flex items-center gap-2"><i data-lucide="user-x" class="w-4 h-4"></i>นักศึกษาที่ยังไม่มีอาจารย์ที่ปรึกษา <span class="px-2 py-0.5 rounded-full bg-amber-200 text-amber-800 text-xs">${noAdvisorStudents.length} คน</span></summary>
    <div class="px-3 pb-3">
      ${noAdvisorStudents.length ? `<div class="bg-white rounded-xl border border-amber-100 overflow-hidden"><div class="overflow-x-auto"><table class="w-full text-sm">
        <thead><tr class="bg-amber-100 text-left text-amber-900"><th class="px-4 py-2 font-semibold">รหัส</th><th class="px-4 py-2 font-semibold">ชื่อ-สกุล</th><th class="px-4 py-2 font-semibold">ชั้นปี</th><th class="px-4 py-2 font-semibold">รุ่นที่</th><th class="px-4 py-2 font-semibold">จัดการ</th></tr></thead>
        <tbody>${noAdvRows}</tbody></table></div></div>` : '<p class="text-sm text-amber-700 px-1 py-2">นักศึกษาที่กำลังศึกษาทุกคนมีอาจารย์ที่ปรึกษาแล้ว</p>'}
    </div>
  </details>`;

  // ----- พาเนล: อาจารย์ที่ยังไม่มีนักศึกษาในความดูแล (ไม่รวมผู้ที่ลาออก, เคารพตัวกรองสาขาวิชา) -----
  let freeTeachers = teachers.filter(t => norm(t.name) && norm(t.teacher_status || '') !== 'ลาออก' && !advisorMap[nameKey(t.name)]);
  if (selectedDept) freeTeachers = freeTeachers.filter(t => (norm(t.department) || NO_DEPT) === selectedDept);
  freeTeachers.sort((a, b) => norm(a.name).localeCompare(norm(b.name), 'th'));
  const freeTeacherCards = freeTeachers.map(t => {
    const tkey = nameKey(t.name).replace(/'/g, "\\'");
    return `<button onclick="APP.filters._advisorSelected='${tkey}';APP.pagination.page=1;renderCurrentPage()" class="text-left bg-white rounded-xl p-3 border border-gray-200 hover:border-primary hover:shadow transition flex items-center gap-3">
      <div class="w-9 h-9 bg-gray-100 rounded-lg flex items-center justify-center flex-shrink-0"><i data-lucide="user" class="w-4 h-4 text-gray-400"></i></div>
      <div class="min-w-0 flex-1"><p class="font-medium text-gray-800 truncate text-sm">${norm(t.name)}</p><p class="text-xs text-gray-500 truncate">${norm(t.department) || NO_DEPT}</p></div>
      ${isAdmin ? '<i data-lucide="user-plus" class="w-4 h-4 text-primary flex-shrink-0"></i>' : ''}
    </button>`;
  }).join('');
  const freeTeacherPanel = `<details class="bg-gray-50 border border-gray-200 rounded-2xl mb-4">
    <summary class="cursor-pointer px-4 py-3 text-sm font-semibold text-gray-700 flex items-center gap-2"><i data-lucide="users" class="w-4 h-4"></i>อาจารย์ที่ยังไม่มีนักศึกษาในความดูแล <span class="px-2 py-0.5 rounded-full bg-gray-200 text-gray-700 text-xs">${freeTeachers.length} ท่าน</span></summary>
    <div class="px-3 pb-3">
      ${freeTeachers.length ? `<div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-2">${freeTeacherCards}</div>${isAdmin ? '<p class="text-xs text-gray-400 mt-2"><i data-lucide="info" class="w-3 h-3 inline"></i> คลิกอาจารย์เพื่อเปิดและเพิ่มนักศึกษาในความดูแล</p>' : ''}` : '<p class="text-sm text-gray-500 px-1 py-2">อาจารย์ทุกท่านตามเงื่อนไขมีนักศึกษาในความดูแลแล้ว</p>'}
    </div>
  </details>`;

  const cards = advisors.map(a => {
    const activeCount = activeStudents(a.students).length;
    const safeName = (a.key || a.name).replace(/'/g, "\\'");
    return `<button onclick="APP.filters._advisorSelected='${safeName}';APP.pagination.page=1;renderCurrentPage()" class="text-left bg-white rounded-2xl p-4 border border-blue-100 hover:border-primary hover:shadow-md transition">
      <div class="flex items-center gap-3">
        <div class="w-11 h-11 bg-primaryLight rounded-xl flex items-center justify-center flex-shrink-0"><i data-lucide="user-check" class="w-5 h-5 text-primary"></i></div>
        <div class="min-w-0 flex-1">
          <p class="font-semibold text-gray-800 truncate">${a.name}</p>
          <p class="text-xs text-gray-500 truncate">${a.dept}</p>
        </div>
        <div class="text-right flex-shrink-0">
          <p class="text-xl font-bold text-primary">${activeCount}</p>
          <p class="text-xs text-gray-400">คน</p>
        </div>
      </div>
    </button>`;
  }).join('');

  return `<div class="flex flex-wrap items-center justify-between gap-3 mb-4">
    <h2 class="text-xl font-bold text-gray-800"><i data-lucide="user-check" class="w-6 h-6 inline mr-2"></i>ข้อมูลอาจารย์ที่ปรึกษา</h2>
    <span class="text-sm text-gray-500">ทั้งหมด ${advisors.length} ท่าน</span>
  </div>
  ${deptFilter}
  ${searchBox}
  ${noAdvisorPanel}
  ${freeTeacherPanel}
  ${advisors.length ? `<div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">${cards}</div>` : '<div class="bg-white rounded-2xl p-8 text-center border border-blue-100"><p class="text-gray-400">ไม่พบอาจารย์ที่ปรึกษาตามเงื่อนไข</p></div>'}`;
}

// เชื่อมไปหน้าผลการเรียน / ผลสอบภาษาอังกฤษ โดยเลือกนักศึกษาไว้ล่วงหน้า
function advisorGotoGrades(sid) { navigateTo('grades'); APP.filters._gradeStudent = sid; renderCurrentPage(); }
function advisorGotoEng(sid) { navigateTo('engResults'); APP.filters._engStudent = sid; renderCurrentPage(); }

// ----- กำหนดอาจารย์ที่ปรึกษาให้นักศึกษา (สำหรับนักศึกษาที่ยังไม่มีที่ปรึกษา) -----
function showAssignAdvisorToStudent(id) {
  if (!(GSheetDB.hasWriteAccess && GSheetDB.hasWriteAccess())) { showToast('ระบบอยู่ในโหมดอ่านอย่างเดียว — ตั้งค่า Apps Script URL ก่อน', 'error'); return; }
  const s = APP.allData.find(d => d.__backendId === id); if (!s) return;
  const tList = getDataByType('teacher').filter(t => norm(t.name) && norm(t.teacher_status || '') !== 'ลาออก')
    .sort((a, b) => norm(a.name).localeCompare(norm(b.name), 'th'));
  const rows = tList.map(t => {
    const nm = norm(t.name).replace(/'/g, "\\'");
    return `<button onclick="advisorDoAssign('${s.__backendId}','${nm}')" class="adv-asg-row w-full text-left px-3 py-2 border-b hover:bg-blue-50 text-sm" data-search="${(norm(t.name) + ' ' + norm(t.department)).toLowerCase().replace(/"/g, '')}"><span class="font-medium text-gray-800">${norm(t.name)}</span> <span class="text-xs text-gray-400">· ${norm(t.department) || 'ไม่ระบุสาขาวิชา'}</span></button>`;
  }).join('');
  showModal('กำหนดอาจารย์ที่ปรึกษา', `
    <div class="space-y-3">
      <p class="text-sm text-gray-600">เลือกอาจารย์ที่ปรึกษาให้ <strong class="text-primary">${s.name || ''}</strong></p>
      <div class="relative"><i data-lucide="search" class="absolute left-3 top-2.5 w-4 h-4 text-gray-400"></i><input type="text" placeholder="ค้นหาชื่อ/สาขาวิชา..." class="w-full pl-10 pr-4 py-2 border border-gray-200 rounded-xl text-sm" oninput="advisorFilterAssignList(this.value)"></div>
      <div class="border border-gray-200 rounded-xl max-h-80 overflow-y-auto" id="advAsgList">${rows || '<p class="px-3 py-6 text-center text-gray-400 text-sm">ไม่มีรายชื่ออาจารย์</p>'}</div>
    </div>`, null, 'max-w-lg');
  setTimeout(() => lucide.createIcons(), 50);
}

function advisorFilterAssignList(q) {
  q = (q || '').toLowerCase();
  document.querySelectorAll('#advAsgList .adv-asg-row').forEach(el => { el.style.display = (!q || (el.dataset.search || '').includes(q)) ? '' : 'none'; });
}

async function advisorDoAssign(id, teacherName) {
  const s = APP.allData.find(d => d.__backendId === id); if (!s) return;
  s.advisor = teacherName;
  closeModal();
  showToast('กำลังบันทึก...', 'loading');
  const r = await GSheetDB.update(s);
  hideLoadingToast();
  if (r.isOk) { showToast('กำหนดอาจารย์ที่ปรึกษาแล้ว'); renderCurrentPage(); }
  else showToast('เกิดข้อผิดพลาด: ' + (r.error || ''), 'error');
}

// ----- จัดการรายชื่อนักศึกษาในความดูแลของอาจารย์ที่ปรึกษา (เพิ่ม/นำออก) -----
async function advisorRemoveStudent(id) {
  if (!(GSheetDB.hasWriteAccess && GSheetDB.hasWriteAccess())) { showToast('ระบบอยู่ในโหมดอ่านอย่างเดียว — ตั้งค่า Apps Script URL ก่อน', 'error'); return; }
  const s = APP.allData.find(d => d.__backendId === id); if (!s) return;
  const advName = (window._advisorCtx && window._advisorCtx.name) || s.advisor || '';
  showModal('นำนักศึกษาออกจากความดูแล', `
    <div class="space-y-2 text-sm">
      <p>ต้องการนำ <strong class="text-primary">${s.name || ''}</strong> ออกจากความดูแลของ <strong>${advName}</strong> ใช่หรือไม่?</p>
      <p class="text-xs text-gray-500">ระบบจะล้างค่า "อาจารย์ที่ปรึกษา" ของนักศึกษาคนนี้ (ข้อมูลอื่นไม่เปลี่ยน) ภายหลังกำหนดที่ปรึกษาใหม่ได้</p>
    </div>`, async () => {
    closeModal();
    s.advisor = '';
    showToast('กำลังบันทึก...', 'loading');
    const r = await GSheetDB.update(s);
    hideLoadingToast();
    if (r.isOk) { showToast('นำออกจากความดูแลแล้ว'); renderCurrentPage(); }
    else showToast('เกิดข้อผิดพลาด: ' + (r.error || ''), 'error');
  });
}

function showAdvisorAddStudents() {
  const ctx = window._advisorCtx; if (!ctx || !ctx.name) return;
  if (!(GSheetDB.hasWriteAccess && GSheetDB.hasWriteAccess())) { showToast('ระบบอยู่ในโหมดอ่านอย่างเดียว — ตั้งค่า Apps Script URL ก่อน', 'error'); return; }
  const inSet = new Set(ctx.studentIds || []);
  const pool = activeStudents(getDataByType('student'))
    .filter(s => !inSet.has(s.__backendId))
    .sort((x, y) => norm(x.year_level).localeCompare(norm(y.year_level)) || norm(x.name).localeCompare(norm(y.name), 'th'));
  const rowsHTML = pool.map(s => `<label class="adv-add-row flex items-center gap-3 px-3 py-2 border-b hover:bg-gray-50 cursor-pointer" data-search="${((s.student_id || '') + ' ' + (s.name || '') + ' ' + (s.advisor || '')).toLowerCase().replace(/"/g, '')}">
      <input type="checkbox" class="adv-add-chk w-4 h-4" value="${s.__backendId}">
      <span class="flex-1 text-sm"><span class="font-mono text-xs text-gray-500">${s.student_id || ''}</span> ${s.name || ''} <span class="text-xs text-gray-400">· ชั้นปี ${s.year_level || '-'} · รุ่น ${s.batch || '-'}</span>${norm(s.advisor) ? `<span class="text-xs text-amber-600"> · ที่ปรึกษาปัจจุบัน: ${s.advisor}</span>` : ''}</span>
    </label>`).join('');
  showModal('เพิ่มนักศึกษาในความดูแล', `
    <div class="space-y-3">
      <p class="text-sm text-gray-600">เลือกนักศึกษาเพื่อกำหนดให้อยู่ในความดูแลของ <strong class="text-primary">${ctx.name}</strong></p>
      <div class="relative"><i data-lucide="search" class="absolute left-3 top-2.5 w-4 h-4 text-gray-400"></i><input type="text" placeholder="ค้นหาชื่อ/รหัส..." class="w-full pl-10 pr-4 py-2 border border-gray-200 rounded-xl text-sm" oninput="advisorFilterAddList(this.value)"></div>
      <p class="text-xs text-amber-600"><i data-lucide="info" class="w-3 h-3 inline"></i> นักศึกษาที่มีที่ปรึกษาอยู่แล้ว หากเลือก ระบบจะย้ายมาอยู่ในความดูแลของอาจารย์ท่านนี้แทน</p>
      <div class="border border-gray-200 rounded-xl max-h-80 overflow-y-auto" id="advAddList">${rowsHTML || '<p class="px-3 py-6 text-center text-gray-400 text-sm">ไม่มีนักศึกษาที่กำลังศึกษาให้เพิ่ม</p>'}</div>
      <button onclick="advisorAssignSelected()" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark text-sm font-medium"><i data-lucide="user-plus" class="w-4 h-4 inline mr-1"></i>เพิ่มนักศึกษาที่เลือก</button>
    </div>`, null, 'max-w-2xl');
  setTimeout(() => lucide.createIcons(), 50);
}

function advisorFilterAddList(q) {
  q = (q || '').toLowerCase();
  document.querySelectorAll('#advAddList .adv-add-row').forEach(el => {
    el.style.display = (!q || (el.dataset.search || '').includes(q)) ? '' : 'none';
  });
}

async function advisorAssignSelected() {
  const ctx = window._advisorCtx; if (!ctx || !ctx.name) return;
  const ids = [...document.querySelectorAll('.adv-add-chk:checked')].map(b => b.value);
  if (!ids.length) { showToast('ยังไม่ได้เลือกนักศึกษา', 'error'); return; }
  const list = ids.map(id => APP.allData.find(d => d.__backendId === id)).filter(Boolean);
  list.forEach(s => s.advisor = ctx.name);
  closeModal();
  showToast('กำลังบันทึก ' + list.length + ' คน...', 'loading');
  const r = await GSheetDB.updateMany(list);
  hideLoadingToast();
  if (r.isOk) showToast('เพิ่มเข้าความดูแลแล้ว ' + (r.ok != null ? r.ok : list.length) + ' คน');
  else showToast('บันทึกเสร็จ ' + (r.ok || 0) + ' คน · ผิดพลาด ' + (r.fail || 0) + ' คน', (r.ok ? 'success' : 'error'));
  renderCurrentPage();
}

// ======================== เลื่อนชั้นปี (Promote) ========================
// แผงปุ่มเลื่อนชั้น แยกต่อชั้นปี (admin เท่านั้น) — เลื่อนเฉพาะผู้ที่กำลังศึกษา
function promotePanelHTML(allStudents) {
  const cnt = y => (allStudents || []).filter(s => isActiveStudent(s) && norm(s.year_level) === y).length;
  const btn = (from, label, color) => `<button onclick="confirmPromoteYear('${from}')" class="flex items-center gap-2 px-3 py-2 rounded-xl text-sm font-medium ${color}"><i data-lucide="arrow-up" class="w-4 h-4"></i>${label} <span class="opacity-80">(${cnt(from)} คน)</span></button>`;
  return `<div class="bg-amber-50 rounded-2xl p-4 border border-amber-200 mb-4">
    <div class="flex items-center gap-2 mb-2 flex-wrap"><i data-lucide="arrow-up-circle" class="w-5 h-5 text-amber-600"></i><h3 class="font-semibold text-amber-800 text-sm">เลื่อนชั้นปี</h3><span class="text-xs text-amber-600">(เลื่อนเฉพาะผู้ที่กำลังศึกษา · ควรสำรองแท็บ student ก่อน)</span></div>
    <div class="flex gap-2 flex-wrap">
      ${btn('1', 'ปี 1 → ปี 2', 'bg-white text-gray-700 border border-amber-300 hover:bg-amber-100')}
      ${btn('2', 'ปี 2 → ปี 3', 'bg-white text-gray-700 border border-amber-300 hover:bg-amber-100')}
      ${btn('3', 'ปี 3 → ปี 4', 'bg-white text-gray-700 border border-amber-300 hover:bg-amber-100')}
      ${btn('4', 'ปี 4 → สำเร็จการศึกษา', 'bg-amber-600 text-white hover:bg-amber-700')}
    </div>
  </div>`;
}

function confirmPromoteYear(fromYear) {
  if (!(GSheetDB.hasWriteAccess && GSheetDB.hasWriteAccess())) { showToast('ระบบอยู่ในโหมดอ่านอย่างเดียว — ตั้งค่า Apps Script URL ก่อน', 'error'); return; }
  const list = getDataByType('student').filter(s => isActiveStudent(s) && norm(s.year_level) === String(fromYear));
  if (!list.length) { showToast('ไม่มีนักศึกษาที่กำลังศึกษาในชั้นปีที่ ' + fromYear, 'error'); return; }
  const isGrad = String(fromYear) === '4';
  const targetLabel = isGrad ? 'สำเร็จการศึกษา (ชั้นปี "จบ")' : 'ชั้นปีที่ ' + (Number(fromYear) + 1);
  showModal('ยืนยันการเลื่อนชั้น', `
    <div class="space-y-3 text-sm">
      <p>กำลังจะเลื่อนนักศึกษา <strong class="text-primary">${list.length} คน</strong> จาก <strong>ชั้นปีที่ ${fromYear}</strong> → <strong>${targetLabel}</strong></p>
      <div class="bg-amber-50 border border-amber-200 rounded-xl p-3 text-amber-800 text-xs space-y-1">
        <p class="font-semibold">⚠ ข้อควรระวัง</p>
        <p>• เลื่อนเฉพาะผู้ที่ "กำลังศึกษา" (ข้ามผู้ที่พักการศึกษา/ลาออก/จบแล้ว)</p>
        <p>• ผลการเรียนและผลสอบเดิมไม่ได้รับผลกระทบ (ผูกด้วยรหัสนักศึกษา)</p>
        <p>• การเลื่อนชั้นย้อนกลับอัตโนมัติไม่ได้ — แนะนำให้คัดลอกแท็บ student สำรองก่อน</p>
        ${isGrad ? '<p>• ปี 4 จะเปลี่ยนสถานะเป็น "สำเร็จการศึกษา" และชั้นปีเป็น "จบ"</p><p>• ระบบจะบันทึกรายชื่อเข้า "ข้อมูลศิษย์เก่า" ให้อัตโนมัติ</p>' : ''}
      </div>
    </div>`, () => doPromoteYear(fromYear));
}

async function doPromoteYear(fromYear) {
  const isGrad = String(fromYear) === '4';
  const list = getDataByType('student').filter(s => isActiveStudent(s) && norm(s.year_level) === String(fromYear));
  if (!list.length) { closeModal(); return; }
  // เก็บข้อมูลสำหรับสร้างศิษย์เก่า (ก่อนอัปเดต/รีเฟรช)
  const gradSnapshot = isGrad ? list.map(s => ({ name: s.name || '', batch: s.batch || '', student_id: s.student_id || '' })) : [];
  list.forEach(s => {
    if (isGrad) { s.year_level = 'จบ'; s.status = 'สำเร็จการศึกษา'; }
    else { s.year_level = String(Number(fromYear) + 1); }
  });
  closeModal();
  showToast('กำลังเลื่อนชั้น ' + list.length + ' คน...', 'loading');
  const r = await GSheetDB.updateMany(list);
  hideLoadingToast();

  // ปี 4 → จบ: บันทึกเข้า "ข้อมูลศิษย์เก่า" อัตโนมัติ (ข้ามคนที่มีในศิษย์เก่าอยู่แล้ว)
  let alumniMsg = '';
  if (isGrad && r.ok) {
    const existing = new Set(getDataByType('alumni').map(a => norm(a.student_id)).filter(Boolean));
    const today = new Date().toISOString().slice(0, 10);
    const newAlumni = gradSnapshot
      .filter(s => { const sid = norm(s.student_id); return !sid || !existing.has(sid); })
      .map(s => ({ type: 'alumni', name: s.name, batch: s.batch, admission_date: '', graduation_date: today, alumni_status: '', workplace: '', recorded_date: today, student_id: s.student_id, created_at: new Date().toISOString() }));
    if (newAlumni.length) {
      showToast('กำลังบันทึกข้อมูลศิษย์เก่า ' + newAlumni.length + ' คน...', 'loading');
      const ar = await GSheetDB.createMany(newAlumni);
      hideLoadingToast();
      alumniMsg = ' · เพิ่มศิษย์เก่า ' + ar.ok + ' คน';
    }
  }

  if (r.isOk) showToast('เลื่อนชั้นสำเร็จ ' + r.ok + ' คน' + alumniMsg);
  else showToast('เลื่อนชั้นเสร็จ ' + r.ok + ' คน · ผิดพลาด ' + r.fail + ' คน' + alumniMsg, r.ok ? 'success' : 'error');
  renderCurrentPage();
}

// ======================== ข้อมูลอาจารย์พิเศษ (ระบบทะเบียน) ========================
// เก็บในแท็บ special_teacher: ปีการศึกษา→academic_year, ชื่อ(รวมคำนำหน้า)→name,
// ตำแหน่ง→academic_position, หน่วยงาน→agency, ระดับวุฒิ→edu_level
function specialTeachersPage() {
  const isAdmin = isAdminOnlyRole();
  const all = getDataByType('special_teacher');
  const years = [...new Set(all.map(t => norm(t.academic_year)).filter(Boolean))].sort((a, b) => b.localeCompare(a));
  const selYear = APP.filters._specialTeacherYear || '';
  const kw = norm(APP.filters._specialTeacherSearch || '').toLowerCase();
  let data = selYear ? all.filter(t => norm(t.academic_year) === selYear) : all.slice();
  if (kw) data = data.filter(t => [t.name, t.academic_position, t.agency, t.subjects, t.edu_level].map(v => norm(v).toLowerCase()).join(' ').includes(kw));
  data.sort((a, b) => norm(b.academic_year).localeCompare(norm(a.academic_year)) || (a.name || '').localeCompare(b.name || ''));
  const total = data.length;
  const paged = paginate(data);

  return `<div class="flex flex-wrap items-center justify-between gap-3 mb-4">
    <h2 class="text-xl font-bold text-gray-800"><i data-lucide="user-plus" class="w-6 h-6 inline mr-2"></i>ข้อมูลอาจารย์พิเศษ</h2>
    ${isAdmin ? `<div class="flex gap-2"><button onclick="showAddSpecialTeacherRegModal()" class="flex items-center gap-2 px-4 py-2 bg-primary text-white rounded-xl hover:bg-primaryDark text-sm"><i data-lucide="plus" class="w-4 h-4"></i>เพิ่มอาจารย์พิเศษ</button>${csvUploadBtn('special_teacher', 'academic_year,name,academic_position,agency,subjects')}</div>` : ''}
  </div>
  <div class="bg-white rounded-2xl p-4 border border-blue-100 mb-4">
    <div class="flex items-center gap-3 flex-wrap">
      <label class="text-sm font-medium text-gray-700">ปีการศึกษา:</label>
      <select onchange="APP.filters._specialTeacherYear=this.value;APP.pagination.page=1;renderCurrentPage()" class="border border-gray-200 rounded-xl px-3 py-2 text-sm">
        <option value="">-- ทุกปีการศึกษา --</option>
        ${years.map(y => `<option value="${y}" ${selYear === y ? 'selected' : ''}>${y}</option>`).join('')}
      </select>
      ${selYear ? `<span class="text-xs text-gray-500">แสดงปีการศึกษา ${selYear}</span>` : ''}
      <div class="relative flex-1 min-w-[200px] max-w-xs ml-auto">
        <i data-lucide="search" class="absolute left-3 top-2.5 w-4 h-4 text-gray-400"></i>
        <input type="text" value="${(APP.filters._specialTeacherSearch || '').replace(/"/g, '&quot;')}" placeholder="ค้นหาชื่อ / ตำแหน่ง / หน่วยงาน..." class="w-full pl-10 pr-4 py-2 border border-gray-200 rounded-xl text-sm" oninput="clearTimeout(window._specialTeacherSearchTimer);window._specialTeacherSearchTimer=setTimeout(()=>{APP.filters._specialTeacherSearch=this.value;APP.pagination.page=1;renderCurrentPage()},300)">
      </div>
    </div>
    ${kw ? `<p class="text-xs text-gray-500 mt-2">พบ ${total} รายการจากคำค้น "${kw.replace(/</g, '&lt;')}"</p>` : ''}
  </div>
  <div class="bg-white rounded-2xl border border-blue-100 overflow-hidden">
    <div class="overflow-x-auto"><table class="w-full text-sm">
      <thead><tr class="bg-surface text-left"><th class="px-4 py-3 font-semibold">ปีการศึกษา</th><th class="px-4 py-3 font-semibold">ชื่อ-สกุล</th><th class="px-4 py-3 font-semibold">ตำแหน่ง</th><th class="px-4 py-3 font-semibold">หน่วยงาน</th>${isAdmin ? '<th class="px-4 py-3"></th>' : ''}</tr></thead>
      <tbody>${paged.length ? paged.map(t => `<tr class="border-t hover:bg-gray-50">
        <td class="px-4 py-3 align-top">${t.academic_year || ''}</td>
        <td class="px-4 py-3 font-medium align-top">${t.name || ''}</td>
        <td class="px-4 py-3 align-top">${t.academic_position || ''}</td>
        <td class="px-4 py-3 align-top">${t.agency || ''}</td>
        ${isAdmin ? `<td class="px-4 py-3 align-top"><div class="flex gap-1"><button onclick="showEditSpecialTeacherRegModal('${t.__backendId}')" class="text-blue-400 hover:text-blue-600" title="แก้ไข"><i data-lucide="pencil" class="w-4 h-4"></i></button><button onclick="deleteRecord('${t.__backendId}')" class="text-red-400 hover:text-red-600" title="ลบ"><i data-lucide="trash-2" class="w-4 h-4"></i></button></div></td>` : ''}</tr>`).join('') : `<tr><td colspan="${isAdmin ? 5 : 4}" class="px-4 py-8 text-center text-gray-400">ไม่มีข้อมูล</td></tr>`}</tbody>
    </table></div>
  </div>
  ${paginationHTML(total, APP.pagination.perPage, APP.pagination.page, 'changePage')}`;
}

// แยกรายการรายวิชา 1 รายการ → { code, name } โดยเทียบกับชีต subject (ตามชื่อ + ปีการศึกษา)
function splitSubjectEntry(entry, academicYear) {
  const e = norm(entry);
  if (!e) return { code: '', name: '' };
  const subs = getDataByType('subject');
  const y = norm(academicYear);
  // 1) ตรงชื่อวิชา (เลือกปีเดียวกันก่อน)
  let m = subs.find(s => norm(s.subject_name) === e && (!y || norm(s.academic_year) === y))
    || subs.find(s => norm(s.subject_name) === e);
  if (m) return { code: norm(m.subject_code), name: norm(m.subject_name) };
  // 2) ตรงรูปแบบ "รหัส ชื่อวิชา"
  m = subs.find(s => norm(s.subject_code) && norm(norm(s.subject_code) + ' ' + norm(s.subject_name)) === e);
  if (m) return { code: norm(m.subject_code), name: norm(m.subject_name) };
  // 3) ขึ้นต้นด้วยรหัสวิชาที่รู้จัก → ตัดรหัสออกจากชื่อ
  m = subs.find(s => norm(s.subject_code) && e.indexOf(norm(s.subject_code)) === 0);
  if (m) { const code = norm(m.subject_code); return { code, name: e.slice(code.length).trim() }; }
  // 4) ไม่พบ → ใส่ทั้งหมดในช่องชื่อวิชา
  return { code: '', name: e };
}

// สร้าง 2 เซลล์ (รหัสวิชา | รายวิชาที่สอน) สำหรับตารางอาจารย์พิเศษ
function specialSubjectCells(t) {
  const raw = norm(t.subjects || t.edu_level);
  if (!raw) return '<td class="px-4 py-3 text-gray-400 align-top">-</td><td class="px-4 py-3 text-gray-400 align-top">-</td>';
  const items = raw.split(',').map(s => s.trim()).filter(Boolean).map(s => splitSubjectEntry(s, t.academic_year));
  const codeHtml = items.map(it => `<div class="py-0.5">${(it.code || '-').replace(/</g, '&lt;')}</div>`).join('');
  const nameHtml = items.map(it => `<div class="py-0.5">${(it.name || '-').replace(/</g, '&lt;')}</div>`).join('');
  return `<td class="px-4 py-3 whitespace-nowrap align-top font-mono text-xs text-gray-600">${codeHtml}</td><td class="px-4 py-3 align-top">${nameHtml}</td>`;
}

function specialTeacherRegFormBody(t) {
  t = t || {};
  const year = norm(t.academic_year) || currentAcademicYearBE();
  return `
    <div><label class="block text-xs text-gray-600 mb-1">ปีการศึกษา *</label><input name="academic_year" id="specialAcademicYear" required value="${String(t.academic_year || currentAcademicYearBE()).replace(/"/g, '&quot;')}" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="เช่น 2568"></div>
    ${titlePrefixField(t.name || '')}
    <div class="grid grid-cols-2 gap-3">
      <div><label class="block text-xs text-gray-600 mb-1">ตำแหน่ง</label><input name="academic_position" value="${(t.academic_position || '').replace(/"/g, '&quot;')}" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="เช่น นายแพทย์ชำนาญการ"></div>
      <div><label class="block text-xs text-gray-600 mb-1">หน่วยงาน</label><input name="agency" value="${(t.agency || '').replace(/"/g, '&quot;')}" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="เช่น รพ.ราชวิถี"></div>
    </div>
    <p class="text-[11px] text-gray-400">การเลือก "รายวิชาที่สอน" ย้ายไปอยู่ที่ระบบทำเนียบอาจารย์ (เมนูเพิ่มอาจารย์พิเศษ)</p>`;
}

// สร้างตัวเลือกรายวิชา (จากชีต subject) กรองตามปีการศึกษา — ไม่ซ้ำชื่อวิชา
function specialSubjectOptionsForYear(year) {
  const y = norm(year);
  let subs = getDataByType('subject');
  if (y) subs = subs.filter(s => norm(s.academic_year) === y);
  const seen = new Set(); const out = [];
  subs.forEach(s => {
    const name = norm(s.subject_name); if (!name || seen.has(name)) return;
    seen.add(name);
    out.push({ name, label: s.subject_code ? `${norm(s.subject_code)} ${name}` : name });
  });
  out.sort((a, b) => a.label.localeCompare(b.label, 'th'));
  return out;
}

function specialSubjectDropdownOptionsHTML(year) {
  const opts = specialSubjectOptionsForYear(year);
  if (!opts.length) return '<option value="">— ไม่มีรายวิชาในปีนี้ —</option>';
  return '<option value="">— เลือกรายวิชา —</option>' + opts.map(o => `<option value="${o.name.replace(/"/g, '&quot;')}">${o.label.replace(/</g, '&lt;')}</option>`).join('');
}

function refreshSpecialSubjectOptions() {
  const yearEl = document.getElementById('specialAcademicYear');
  const sel = document.getElementById('specialSubjectSelect');
  if (!yearEl || !sel) return;
  sel.innerHTML = specialSubjectDropdownOptionsHTML(yearEl.value);
}

function getSpecialSubjectsArr() {
  const v = ((document.getElementById('specialSubjectsValue') || {}).value || '');
  return v.split(',').map(s => s.trim()).filter(Boolean);
}

function setSpecialSubjectsArr(arr) {
  const uniq = [...new Set(arr.map(s => s.trim()).filter(Boolean))];
  const hid = document.getElementById('specialSubjectsValue'); if (hid) hid.value = uniq.join(', ');
  renderSpecialSubjectChips();
}

function renderSpecialSubjectChips() {
  const box = document.getElementById('specialSubjectChips'); if (!box) return;
  const arr = getSpecialSubjectsArr();
  box.innerHTML = arr.length
    ? arr.map((s, i) => `<span class="inline-flex items-center gap-1 bg-blue-50 text-blue-700 border border-blue-100 rounded-lg px-2 py-1 text-xs">${String(s).replace(/</g, '&lt;')}<button type="button" onclick="removeSpecialSubject(${i})" class="text-blue-400 hover:text-red-500 font-bold leading-none">×</button></span>`).join('')
    : '<span class="text-xs text-gray-400">ยังไม่ได้เลือกรายวิชา</span>';
}

function addSpecialSubject() {
  const sel = document.getElementById('specialSubjectSelect'); if (!sel || !sel.value) return;
  const arr = getSpecialSubjectsArr(); arr.push(sel.value); setSpecialSubjectsArr(arr); sel.value = '';
}

function addSpecialSubjectManual() {
  const el = document.getElementById('specialSubjectManual'); if (!el || !el.value.trim()) return;
  const arr = getSpecialSubjectsArr(); arr.push(el.value.trim()); setSpecialSubjectsArr(arr); el.value = '';
}

function removeSpecialSubject(i) {
  const arr = getSpecialSubjectsArr(); arr.splice(i, 1); setSpecialSubjectsArr(arr);
}

function collectSpecialTeacherReg(form, obj) {
  obj.academic_year = form.querySelector('[name="academic_year"]').value;
  obj.name = combineName(form);
  obj.academic_position = form.querySelector('[name="academic_position"]').value;
  obj.agency = form.querySelector('[name="agency"]').value;
  return obj;
}

function showAddSpecialTeacherRegModal() {
  showModal('เพิ่มอาจารย์พิเศษ', `
    <form id="addSpecialTeacherRegForm" class="space-y-3">
      ${specialTeacherRegFormBody({})}
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึก</button>
    </form>
  `);
  renderSpecialSubjectChips();
  document.getElementById('addSpecialTeacherRegForm').onsubmit = async (e) => {
    e.preventDefault();
    await withLoading(e.target, async () => {
      const obj = collectSpecialTeacherReg(e.target, { type: 'special_teacher', created_at: new Date().toISOString() });
      const r = await GSheetDB.create(obj);
      if (r.isOk) { showToast('เพิ่มอาจารย์พิเศษสำเร็จ'); closeModal(); } else showToast('เกิดข้อผิดพลาด', 'error');
    });
  };
}

function showEditSpecialTeacherRegModal(id) {
  const t = APP.allData.find(d => d.__backendId === id); if (!t) return;
  showModal('แก้ไขข้อมูลอาจารย์พิเศษ', `
    <form id="editSpecialTeacherRegForm" class="space-y-3">
      ${specialTeacherRegFormBody(t)}
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึกการแก้ไข</button>
    </form>
  `);
  renderSpecialSubjectChips();
  document.getElementById('editSpecialTeacherRegForm').onsubmit = async (e) => {
    e.preventDefault();
    await withLoading(e.target, async () => {
      const rec = APP.allData.find(d => d.__backendId === id); if (!rec) return;
      collectSpecialTeacherReg(e.target, rec);
      const r = await GSheetDB.update(rec);
      if (r.isOk) { showToast('แก้ไขข้อมูลสำเร็จ'); closeModal(); renderCurrentPage(); } else showToast('เกิดข้อผิดพลาด: ' + (r.error || ''), 'error');
    });
  };
}

// ======================== ข้อมูลศิษย์เก่า (ระบบทะเบียน) ========================
// แท็บ alumni: name(รวมคำนำหน้า), batch(รุ่น), alumni_status(สถานภาพ), workplace(สถานที่ปฏิบัติงาน), recorded_date(วันที่บันทึก), student_id
const ALUMNI_STATUS_OPTIONS = ['ปฏิบัติงาน', 'กำลังศึกษาต่อ', 'ยังไม่ได้ปฏิบัติงาน', 'อื่นๆ'];

function alumniPage() {
  const isAdmin = isAdminRole();
  const all = getDataByType('alumni');
  const batches = [...new Set(all.map(a => norm(a.batch)).filter(Boolean))].sort((a, b) => b.localeCompare(a, undefined, { numeric: true }));
  const selBatch = APP.filters._alumniBatch || '';
  let data = selBatch ? all.filter(a => norm(a.batch) === selBatch) : all.slice();
  data.sort((a, b) => norm(b.batch).localeCompare(norm(a.batch), undefined, { numeric: true }) || (a.name || '').localeCompare(b.name || ''));
  const total = data.length;
  const paged = paginate(data);

  return `<div class="flex flex-wrap items-center justify-between gap-3 mb-4">
    <h2 class="text-xl font-bold text-gray-800"><i data-lucide="graduation-cap" class="w-6 h-6 inline mr-2"></i>ข้อมูลศิษย์เก่า</h2>
    ${isAdmin ? `<div class="flex gap-2"><button onclick="showAddAlumniModal()" class="flex items-center gap-2 px-4 py-2 bg-primary text-white rounded-xl hover:bg-primaryDark text-sm"><i data-lucide="plus" class="w-4 h-4"></i>เพิ่มศิษย์เก่า</button>${csvUploadBtn('alumni', 'name,batch,admission_date,graduation_date,alumni_status,workplace,recorded_date,student_id')}</div>` : ''}
  </div>
  <div class="bg-white rounded-2xl p-4 border border-blue-100 mb-4">
    <div class="flex items-center gap-3 flex-wrap">
      <label class="text-sm font-medium text-gray-700">รุ่นที่:</label>
      <select onchange="APP.filters._alumniBatch=this.value;APP.pagination.page=1;renderCurrentPage()" class="border border-gray-200 rounded-xl px-3 py-2 text-sm">
        <option value="">-- ทุกรุ่น --</option>
        ${batches.map(b => `<option value="${b}" ${selBatch === b ? 'selected' : ''}>รุ่นที่ ${b}</option>`).join('')}
      </select>
      ${selBatch ? `<span class="text-xs text-gray-500">แสดงรุ่นที่ ${selBatch}</span>` : ''}
      <span class="text-xs text-gray-400 ml-auto"><i data-lucide="users" class="w-3 h-3 inline mr-1"></i>${total} คน</span>
    </div>
  </div>
  <div class="bg-white rounded-2xl border border-blue-100 overflow-hidden">
    <div class="overflow-x-auto"><table class="w-full text-sm">
      <thead><tr class="bg-surface text-left"><th class="px-4 py-3 font-semibold">คำนำหน้า</th><th class="px-4 py-3 font-semibold">ชื่อ-สกุล</th><th class="px-4 py-3 font-semibold">รุ่นที่</th><th class="px-4 py-3 font-semibold">เข้าศึกษาวันที่</th><th class="px-4 py-3 font-semibold">จบการศึกษาวันที่</th><th class="px-4 py-3 font-semibold">สถานภาพ</th><th class="px-4 py-3 font-semibold">สถานที่ปฏิบัติงาน</th>${isAdmin ? '<th class="px-4 py-3"></th>' : ''}</tr></thead>
      <tbody>${paged.length ? paged.map(a => {
    const pp = parseTitlePrefix(a.name || '');
    return `<tr class="border-t hover:bg-gray-50">
        <td class="px-4 py-3">${pp.prefix || ''}</td>
        <td class="px-4 py-3 font-medium">${pp.rest || a.name || ''}</td>
        <td class="px-4 py-3">${a.batch || ''}</td>
        <td class="px-4 py-3">${a.admission_date ? toBuddhistDate(a.admission_date) : ''}</td>
        <td class="px-4 py-3">${a.graduation_date ? toBuddhistDate(a.graduation_date) : ''}</td>
        <td class="px-4 py-3">${a.alumni_status || ''}</td>
        <td class="px-4 py-3">${a.workplace || ''}</td>
        ${isAdmin ? `<td class="px-4 py-3"><div class="flex gap-1"><button onclick="showEditAlumniModal('${a.__backendId}')" class="text-blue-400 hover:text-blue-600" title="แก้ไข"><i data-lucide="pencil" class="w-4 h-4"></i></button><button onclick="deleteRecord('${a.__backendId}')" class="text-red-400 hover:text-red-600" title="ลบ"><i data-lucide="trash-2" class="w-4 h-4"></i></button></div></td>` : ''}</tr>`;
  }).join('') : `<tr><td colspan="${isAdmin ? 8 : 7}" class="px-4 py-8 text-center text-gray-400">ไม่มีข้อมูล</td></tr>`}</tbody>
    </table></div>
  </div>
  ${paginationHTML(total, APP.pagination.perPage, APP.pagination.page, 'changePage')}`;
}

function alumniFormBody(a) {
  a = a || {};
  const v = k => String(a[k] == null ? '' : a[k]).replace(/"/g, '&quot;');
  const statusOpts = ALUMNI_STATUS_OPTIONS.map(o => `<option ${norm(a.alumni_status) === o ? 'selected' : ''}>${o}</option>`).join('');
  return `
    ${titlePrefixField(a.name || '')}
    <div class="grid grid-cols-2 gap-3">
      <div><label class="block text-xs text-gray-600 mb-1">รุ่นที่</label><input name="batch" value="${v('batch')}" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="เช่น 36"></div>
      <div><label class="block text-xs text-gray-600 mb-1">สถานภาพ</label><select name="alumni_status" class="w-full border rounded-xl px-3 py-2 text-sm"><option value="">-- เลือก --</option>${statusOpts}</select></div>
      <div><label class="block text-xs text-gray-600 mb-1">เข้าศึกษาวันที่</label><input name="admission_date" type="date" value="${v('admission_date')}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <div><label class="block text-xs text-gray-600 mb-1">จบการศึกษาวันที่</label><input name="graduation_date" type="date" value="${v('graduation_date')}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
    </div>
    <div><label class="block text-xs text-gray-600 mb-1">สถานที่ปฏิบัติงาน</label><input name="workplace" value="${v('workplace')}" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="เช่น รพ.ราชวิถี"></div>
    <div><label class="block text-xs text-gray-600 mb-1">วันที่บันทึกข้อมูล</label><input name="recorded_date" type="date" value="${v('recorded_date')}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
    <input type="hidden" name="student_id" value="${v('student_id')}">`;
}

function collectAlumni(form, obj) {
  obj.name = combineName(form);
  obj.batch = form.querySelector('[name="batch"]').value;
  obj.alumni_status = form.querySelector('[name="alumni_status"]').value;
  obj.admission_date = form.querySelector('[name="admission_date"]').value;
  obj.graduation_date = form.querySelector('[name="graduation_date"]').value;
  obj.workplace = form.querySelector('[name="workplace"]').value;
  obj.recorded_date = form.querySelector('[name="recorded_date"]').value;
  const sid = form.querySelector('[name="student_id"]'); if (sid) obj.student_id = sid.value;
  return obj;
}

function showAddAlumniModal() {
  showModal('เพิ่มศิษย์เก่า', `
    <form id="addAlumniForm" class="space-y-3">
      ${alumniFormBody({ recorded_date: new Date().toISOString().slice(0, 10) })}
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึก</button>
    </form>
  `);
  document.getElementById('addAlumniForm').onsubmit = async (e) => {
    e.preventDefault();
    await withLoading(e.target, async () => {
      const obj = collectAlumni(e.target, { type: 'alumni', created_at: new Date().toISOString() });
      const r = await GSheetDB.create(obj);
      if (r.isOk) { showToast('เพิ่มศิษย์เก่าสำเร็จ'); closeModal(); } else showToast('เกิดข้อผิดพลาด', 'error');
    });
  };
}

function showEditAlumniModal(id) {
  const a = APP.allData.find(d => d.__backendId === id); if (!a) return;
  showModal('แก้ไขข้อมูลศิษย์เก่า', `
    <form id="editAlumniForm" class="space-y-3">
      ${alumniFormBody(a)}
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึกการแก้ไข</button>
    </form>
  `);
  document.getElementById('editAlumniForm').onsubmit = async (e) => {
    e.preventDefault();
    await withLoading(e.target, async () => {
      const rec = APP.allData.find(d => d.__backendId === id); if (!rec) return;
      collectAlumni(e.target, rec);
      const r = await GSheetDB.update(rec);
      if (r.isOk) { showToast('แก้ไขข้อมูลสำเร็จ'); closeModal(); renderCurrentPage(); } else showToast('เกิดข้อผิดพลาด: ' + (r.error || ''), 'error');
    });
  };
}

function showTeacherDetail(id) {
  const t = APP.allData.find(d => d.__backendId === id); if (!t) return;
  const isAdmin = isAdminOnlyRole();
  const _tst = t.teacher_status || 'ปฏิบัติงานอยู่';
  const stBadge = _tst === 'ปฏิบัติงานอยู่' ? '<span class="px-2 py-1 rounded-full text-xs bg-green-100 text-green-700">ปฏิบัติงานอยู่</span>' : _tst === 'ลาออก' ? '<span class="px-2 py-1 rounded-full text-xs bg-red-100 text-red-700">ลาออก</span>' : '<span class="px-2 py-1 rounded-full text-xs bg-yellow-100 text-yellow-700">ลาศึกษาต่อ</span>';
  showModal('ข้อมูลอาจารย์', `<div class="grid grid-cols-2 gap-3">
    ${infoRow('ชื่อ-สกุล', t.name)}${infoRow('ตำแหน่ง', t.position)}${infoRow('สาขาวิชา', t.department)}
    <div><p class="text-xs text-gray-500">สถานะ</p><p class="font-medium mt-1">${stBadge}</p></div>
    ${infoRow('โทร', t.phone)}${infoRow('E-mail', t.email)}${infoRow('ชั้นปีที่รับผิดชอบ', t.responsible_year)}
    ${isAdmin ? infoRow('เลขบัญชีธนาคาร', t.bank_account) : ''}
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
        <div><label class="block text-xs text-gray-600 mb-1">สถานะ</label><select name="teacher_status" class="w-full border rounded-xl px-3 py-2 text-sm"><option>ปฏิบัติงานอยู่</option><option>ลาศึกษาต่อ</option><option>ลาออก</option></select></div>
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

// ---- ผลงานวิชาการ: แต่ละรายการมี ประเภท (บทความวิชาการ/บทความวิจัย/อื่นๆ) + รายละเอียด ----
// เก็บเป็นข้อความ "ประเภท: รายละเอียด" คั่นรายการด้วย ||
function awParse(values) {
  const items = values ? values.split('||').map(v => v.trim()).filter(Boolean) : [];
  return items.map(item => {
    const idx = item.indexOf(': ');
    if (idx > 0 && idx <= 20) return { type: item.slice(0, idx).trim(), detail: item.slice(idx + 2).trim() };
    return { type: '', detail: item };
  });
}
function awRowHTML(r) {
  r = r || { type: '', detail: '' };
  return `<div class="aw-row flex gap-2 items-start"><input name="aw_type__multi" list="awTypeList" autocomplete="off" value="${(r.type || '').replace(/"/g, '&quot;')}" class="w-36 border rounded-xl px-2 py-2 text-sm flex-shrink-0" placeholder="เลือก/พิมพ์เอง"><input name="aw_detail__multi" value="${(r.detail || '').replace(/"/g, '&quot;')}" class="flex-1 border rounded-xl px-3 py-2 text-sm" placeholder="รายละเอียดผลงาน"><button type="button" onclick="this.parentElement.remove()" class="text-red-400 hover:text-red-600 px-1" title="ลบ">✕</button></div>`;
}
function multiAcademicWorkField(values) {
  const rows = awParse(values); if (!rows.length) rows.push({ type: '', detail: '' });
  return `<div>
    <label class="block text-xs text-gray-600 mb-1">ผลงานวิชาการ (ย้อนหลัง 5 ปี) <span class="text-gray-400">(ช่องประเภทเลือกหรือพิมพ์เองก็ได้ · กดปุ่ม + เพื่อเพิ่ม)</span></label>
    <datalist id="awTypeList"><option value="บทความวิชาการ"></option><option value="บทความวิจัย"></option><option value="อื่นๆ"></option></datalist>
    <div id="multi_academic_work" class="space-y-2">${rows.map(awRowHTML).join('')}</div>
    <button type="button" onclick="addAcademicWorkInput()" class="mt-1 text-xs text-primary hover:underline flex items-center gap-1"><span>＋</span> เพิ่มผลงาน</button>
  </div>`;
}
function addAcademicWorkInput() {
  const c = document.getElementById('multi_academic_work'); if (!c) return;
  const div = document.createElement('div');
  div.className = 'aw-row flex gap-2 items-start';
  div.innerHTML = `<input name="aw_type__multi" list="awTypeList" autocomplete="off" value="" class="w-36 border rounded-xl px-2 py-2 text-sm flex-shrink-0" placeholder="เลือก/พิมพ์เอง"><input name="aw_detail__multi" value="" class="flex-1 border rounded-xl px-3 py-2 text-sm" placeholder="รายละเอียดผลงาน"><button type="button" onclick="this.parentElement.remove()" class="text-red-400 hover:text-red-600 px-1" title="ลบ">✕</button>`;
  c.appendChild(div);
}
function collectAcademicWork(form) {
  const rows = form.querySelectorAll('#multi_academic_work .aw-row');
  const out = [];
  rows.forEach(row => {
    const typeEl = row.querySelector('[name="aw_type__multi"]');
    const detailEl = row.querySelector('[name="aw_detail__multi"]');
    const t = (typeEl ? typeEl.value : '').trim();
    const d = (detailEl ? detailEl.value : '').trim();
    if (!t && !d) return;
    out.push(t ? (t + ': ' + d) : d);
  });
  return out.join(' || ');
}

// ---- ประสบการณ์: แต่ละแถวมี ปี/เดือน + รายละเอียด (เพิ่มได้หลายแถว) ----
// เก็บเป็น "ปี::เดือน::รายละเอียด" คั่นด้วย || และคำนวณรวมเก็บใน *_years/_months
function expParse(values) {
  const items = values ? values.split('||').map(v => v.trim()).filter(Boolean) : [];
  return items.map(item => {
    const m = item.match(/^(\d*)\s*::\s*(\d*)\s*::\s*([\s\S]*)$/);
    if (m) return { y: m[1], mo: m[2], detail: m[3].trim() };
    return { y: '', mo: '', detail: item };
  });
}
function expRowHTML(name, ph, r) {
  r = r || { y: '', mo: '', detail: '' };
  return `<div class="${name}-row flex gap-1 items-center">
    <input name="${name}_y" type="number" min="0" value="${r.y || ''}" class="w-14 border rounded-lg px-2 py-2 text-sm" placeholder="ปี"><span class="text-xs text-gray-500">ปี</span>
    <input name="${name}_mo" type="number" min="0" max="11" value="${r.mo || ''}" class="w-16 border rounded-lg px-2 py-2 text-sm" placeholder="เดือน"><span class="text-xs text-gray-500">เดือน</span>
    <input name="${name}_d" value="${(r.detail || '').replace(/"/g, '&quot;')}" class="flex-1 border rounded-xl px-3 py-2 text-sm" placeholder="${ph}">
    <button type="button" onclick="this.parentElement.remove()" class="text-red-400 hover:text-red-600 px-1" title="ลบ">✕</button>
  </div>`;
}
function multiExpField(name, label, ph, values, legacyY, legacyMo) {
  const rows = expParse(values); if (!rows.length) rows.push({ y: '', mo: '', detail: '' });
  if ((legacyY || legacyMo) && rows.length && !rows[0].y && !rows[0].mo) { rows[0].y = legacyY || ''; rows[0].mo = legacyMo || ''; }
  return `<div>
    <label class="block text-xs text-gray-600 mb-1">${label} <span class="text-gray-400">(แต่ละรายการระบุ ปี/เดือน ได้ · กดปุ่ม + เพื่อเพิ่ม)</span></label>
    <div id="multi_${name}" class="space-y-2">${rows.map(r => expRowHTML(name, ph, r)).join('')}</div>
    <button type="button" onclick="addExpRow('${name}','${ph}')" class="mt-1 text-xs text-primary hover:underline flex items-center gap-1"><span>＋</span> เพิ่มรายการ</button>
  </div>`;
}
function addExpRow(name, ph) {
  const c = document.getElementById('multi_' + name); if (!c) return;
  const div = document.createElement('div');
  div.className = name + '-row flex gap-1 items-center';
  div.innerHTML = `<input name="${name}_y" type="number" min="0" value="" class="w-14 border rounded-lg px-2 py-2 text-sm" placeholder="ปี"><span class="text-xs text-gray-500">ปี</span><input name="${name}_mo" type="number" min="0" max="11" value="" class="w-16 border rounded-lg px-2 py-2 text-sm" placeholder="เดือน"><span class="text-xs text-gray-500">เดือน</span><input name="${name}_d" value="" class="flex-1 border rounded-xl px-3 py-2 text-sm" placeholder="${ph}"><button type="button" onclick="this.parentElement.remove()" class="text-red-400 hover:text-red-600 px-1" title="ลบ">✕</button>`;
  c.appendChild(div);
}
function collectExp(form, name) {
  const rows = form.querySelectorAll('#multi_' + name + ' .' + name + '-row');
  const out = []; let totalY = 0, totalMo = 0;
  rows.forEach(row => {
    const y = ((row.querySelector('[name="' + name + '_y"]') || {}).value || '').trim();
    const mo = ((row.querySelector('[name="' + name + '_mo"]') || {}).value || '').trim();
    const d = ((row.querySelector('[name="' + name + '_d"]') || {}).value || '').trim();
    if (!y && !mo && !d) return;
    out.push(y + '::' + mo + '::' + d);
    totalY += parseInt(y, 10) || 0; totalMo += parseInt(mo, 10) || 0;
  });
  totalY += Math.floor(totalMo / 12); totalMo = totalMo % 12;
  return { exp: out.join(' || '), years: out.length ? String(totalY) : '', months: out.length ? String(totalMo) : '' };
}
// แสดงรายการประสบการณ์พร้อมระยะเวลาต่อรายการ
function expDetailList(values) {
  const rows = expParse(values);
  if (!rows.length) return '<span class="text-gray-400">-</span>';
  return '<ul class="list-disc list-inside text-sm space-y-1">' + rows.map(r => {
    const dur = expYM(r.y, r.mo);
    return `<li>${r.detail || ''}${dur ? ` <span class="text-primary">(${dur})</span>` : ''}</li>`;
  }).join('') + '</ul>';
}

// สาขาวิชามาตรฐาน (สำหรับแดชบอร์ดสรุปตามสาขา)
const NURSING_BRANCHES = [
  'การพยาบาลผู้ใหญ่และผู้สูงอายุ',
  'การพยาบาลเด็ก',
  'การพยาบาลอนามัยชุมชน',
  'การพยาบาลมารดา ทารก และผดุงครรภ์',
  'การพยาบาลจิตเวชและสุขภาพจิต'
];

// ฟิลด์เสริมสำหรับป้อนแดชบอร์ดสรุป: สาขาวิชา + ระดับวุฒิ + กลุ่มสาขาวุฒิ
function branchEduFields(t) {
  t = t || {};
  const opt = (v, cur) => `<option ${norm(cur) === v ? 'selected' : ''}>${v}</option>`;
  return `
  <div class="p-3 bg-blue-50 rounded-xl border border-blue-100 space-y-2">
    <p class="text-xs font-semibold text-primary"><i data-lucide="bar-chart-3" class="w-3 h-3 inline"></i> ข้อมูลสำหรับแดชบอร์ดสรุป</p>
    <div><label class="block text-xs text-gray-600 mb-1">สาขาวิชา</label>
      <input name="nursing_branch" list="branchList" value="${(t.nursing_branch || '').replace(/"/g, '&quot;')}" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="เลือกหรือพิมพ์สาขา">
      <datalist id="branchList">${NURSING_BRANCHES.map(b => `<option value="${b}"></option>`).join('')}</datalist>
    </div>
    <div class="grid grid-cols-2 gap-3">
      <div><label class="block text-xs text-gray-600 mb-1">ระดับวุฒิ (สูงสุด)</label>
        <select name="edu_level" class="w-full border rounded-xl px-3 py-2 text-sm"><option value="">-- เลือก --</option>${['ปริญญาเอก', 'ปริญญาโท', 'ปริญญาตรี', 'อื่นๆ'].map(v => opt(v, t.edu_level)).join('')}</select>
      </div>
      <div><label class="block text-xs text-gray-600 mb-1">กลุ่มสาขาวุฒิ</label>
        <select name="edu_field" class="w-full border rounded-xl px-3 py-2 text-sm"><option value="">-- เลือก --</option>${['สาขาการพยาบาล', 'สาขาที่สัมพันธ์ทางการพยาบาล'].map(v => opt(v, t.edu_field)).join('')}</select>
      </div>
    </div>
  </div>`;
}

// ---- คำนำหน้าชื่อ: ปุ่มเลือก (นาง/นางสาว/นาย) + พิมพ์เองได้ — เก็บรวมในชื่อ ----
const TITLE_PREFIXES = ['นางสาว', 'นายแพทย์', 'แพทย์หญิง', 'ว่าที่ร้อยตรีหญิง', 'ว่าที่ร้อยตรี', 'นพ.', 'พญ.', 'ผศ.ดร.', 'รศ.ดร.', 'ผศ.', 'รศ.', 'ดร.', 'ศ.', 'นาง', 'นาย'];
function parseTitlePrefix(fullName) {
  const s = String(fullName || '').trim();
  for (const p of TITLE_PREFIXES) { if (s.startsWith(p)) return { prefix: p, rest: s.slice(p.length).trim() }; }
  return { prefix: '', rest: s };
}
function titlePrefixField(fullName) {
  const { prefix, rest } = parseTitlePrefix(fullName || '');
  const chip = v => `<button type="button" onclick="setTitlePrefix(this,'${v}')" class="tp-chip px-3 py-1.5 rounded-lg border text-sm transition ${prefix === v ? 'bg-primary text-white border-primary' : 'bg-white text-gray-700 border-gray-300 hover:bg-surface'}">${v}</button>`;
  return `
  <div class="title-prefix-wrap">
    <label class="block text-xs text-gray-600 mb-1">คำนำหน้า</label>
    <div class="flex gap-2 flex-wrap items-center mb-2">
      ${['นาง', 'นางสาว', 'นาย'].map(chip).join('')}
      <input name="title_prefix" value="${prefix.replace(/"/g, '&quot;')}" oninput="syncTitlePrefixChips(this)" class="w-28 border rounded-lg px-2 py-1.5 text-sm" placeholder="หรือพิมพ์เอง">
    </div>
    <label class="block text-xs text-gray-600 mb-1">ชื่อ-สกุล *</label>
    <input name="name" required value="${rest.replace(/"/g, '&quot;')}" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="เช่น อุบล สุทธิเนียม">
  </div>`;
}
function setTitlePrefix(btn, val) {
  const wrap = btn.closest('.title-prefix-wrap'); if (!wrap) return;
  const inp = wrap.querySelector('[name="title_prefix"]');
  if (inp) inp.value = val;
  wrap.querySelectorAll('.tp-chip').forEach(c => { c.classList.remove('bg-primary', 'text-white', 'border-primary'); c.classList.add('bg-white', 'text-gray-700', 'border-gray-300'); });
  btn.classList.remove('bg-white', 'text-gray-700', 'border-gray-300');
  btn.classList.add('bg-primary', 'text-white', 'border-primary');
}
function syncTitlePrefixChips(inp) {
  const wrap = inp.closest('.title-prefix-wrap'); if (!wrap) return;
  wrap.querySelectorAll('.tp-chip').forEach(c => {
    const on = c.textContent.trim() === inp.value.trim();
    c.classList.toggle('bg-primary', on); c.classList.toggle('text-white', on); c.classList.toggle('border-primary', on);
    c.classList.toggle('bg-white', !on); c.classList.toggle('text-gray-700', !on); c.classList.toggle('border-gray-300', !on);
  });
}
// รวมคำนำหน้า + ชื่อ เป็นชื่อเต็ม
function combineName(form) {
  const p = ((form.querySelector('[name="title_prefix"]') || {}).value || '').trim();
  const n = ((form.querySelector('[name="name"]') || {}).value || '').trim();
  return (p + n).trim();
}

// Helper: mask national ID — show first 9 digits, last 4 as xxxx
function maskNationalId(nid) {
  if (!nid) return '-';
  const s = String(nid).trim();
  if (s.length < 5) return s;
  return s.substring(0, s.length - 4) + 'xxxx';
}

// แสดงระยะเวลาเป็น "X ปี Y เดือน" (รองรับกรณีไม่ถึงปี)
function expYM(years, months) {
  const y = parseInt(years, 10); const m = parseInt(months, 10);
  const parts = [];
  if (!isNaN(y) && y > 0) parts.push(y + ' ปี');
  if (!isNaN(m) && m > 0) parts.push(m + ' เดือน');
  return parts.join(' ');
}
// แปลงระยะเวลาเป็นจำนวนปี (ทศนิยม) สำหรับคำนวณ
function expToYears(years, months) {
  const y = parseFloat(years); const m = parseFloat(months);
  let v = 0; if (!isNaN(y)) v += y; if (!isNaN(m)) v += m / 12;
  return v;
}

// ---- Teacher category metadata (5 ประเภท) ----
const TEACHER_CAT_META = {
  'อาจารย์ประจำหลักสูตร': { tab: 'curriculum', badge: 'bg-purple-100 text-purple-700', num: 'text-purple-600', ring: 'ring-purple-500' },
  'อาจารย์ประจำ': { tab: 'regular', badge: 'bg-blue-100 text-blue-700', num: 'text-blue-600', ring: 'ring-blue-500' },
  'อาจารย์ผู้รับผิดชอบหลักสูตร': { tab: 'responsible', badge: 'bg-indigo-100 text-indigo-700', num: 'text-indigo-600', ring: 'ring-indigo-500' },
  'อาจารย์ที่ลาศึกษาต่อ': { tab: 'studyleave', badge: 'bg-amber-100 text-amber-700', num: 'text-amber-600', ring: 'ring-amber-500' },
  'อาจารย์พิเศษ': { tab: 'special', badge: 'bg-emerald-100 text-emerald-700', num: 'text-emerald-600', ring: 'ring-emerald-500' },
};
// Ordered category list for tabs/cards/dropdowns
const TEACHER_CATEGORIES = Object.keys(TEACHER_CAT_META);
// Categories that use the full directory form (everything except special)
const TEACHER_FULL_CATEGORIES = TEACHER_CATEGORIES.filter(c => c !== 'อาจารย์พิเศษ');
function catBadge(cat) { return (TEACHER_CAT_META[cat] || {}).badge || 'bg-gray-100 text-gray-600'; }
function catFromTab(tabId) { for (const k in TEACHER_CAT_META) { if (TEACHER_CAT_META[k].tab === tabId) return k; } return ''; }
function isSpecialTeacher(t) { return (t.teacher_category || '') === 'อาจารย์พิเศษ'; }

// Build <option> list for the subject datalist (รหัส + ชื่อวิชา จากแท็บ subject)
function subjectDatalistOptions() {
  const subs = [...new Set(getDataByType('subject').map(s => {
    const c = norm(s.subject_code); const n = norm(s.subject_name);
    return [c, n].filter(Boolean).join(' ');
  }).filter(Boolean))].sort();
  return subs.map(s => `<option value="${s.replace(/"/g, '&quot;')}"></option>`).join('');
}

// Multi-row subject picker: เลือกจาก datalist หรือพิมพ์เองได้ (เก็บเป็น || )
function multiSubjectField(name, label, values) {
  const items = values ? values.split('||').map(v => v.trim()).filter(Boolean) : [''];
  if (!items.length) items.push('');
  return `
    <div>
      <label class="block text-xs text-gray-600 mb-1">${label} <span class="text-gray-400">(เลือกจากรายการหรือพิมพ์เอง · กดปุ่ม + เพื่อเพิ่มวิชา)</span></label>
      <datalist id="subjectDatalist">${subjectDatalistOptions()}</datalist>
      <div id="multi_${name}" class="space-y-2">
        ${items.map((v, i) => `<div class="flex gap-2"><input name="${name}__multi" list="subjectDatalist" value="${String(v).replace(/"/g, '&quot;')}" class="flex-1 border rounded-xl px-3 py-2 text-sm" placeholder="เลือกหรือพิมพ์ชื่อวิชา">${i > 0 ? `<button type="button" onclick="this.parentElement.remove()" class="text-red-400 hover:text-red-600 px-2" title="ลบ">✕</button>` : '<div class="w-8"></div>'}</div>`).join('')}
      </div>
      <button type="button" onclick="addMultiSubjectInput('${name}')" class="mt-1 text-xs text-primary hover:underline flex items-center gap-1"><span>＋</span> เพิ่มวิชา</button>
    </div>`;
}

function addMultiSubjectInput(name) {
  const container = document.getElementById('multi_' + name);
  if (!container) return;
  const div = document.createElement('div');
  div.className = 'flex gap-2';
  div.innerHTML = `<input name="${name}__multi" list="subjectDatalist" value="" class="flex-1 border rounded-xl px-3 py-2 text-sm" placeholder="เลือกหรือพิมพ์ชื่อวิชา"><button type="button" onclick="this.parentElement.remove()" class="text-red-400 hover:text-red-600 px-2" title="ลบ">✕</button>`;
  container.appendChild(div);
}

// Export to PDF via print
function exportTeacherDirectoryPDF() {
  let data = directoryRecords();
  const activeTab = APP._directoryTab || 'all';
  const selectedYear = APP.filters._directoryYear || '';
  if (selectedYear) data = data.filter(d => (d.academic_year || '') === selectedYear);
  if (activeTab !== 'all') {
    const cat = catFromTab(activeTab);
    if (cat) data = data.filter(d => (d.teacher_category || '') === cat);
  }

  if (!data.length) { showToast('ไม่มีข้อมูลสำหรับส่งออก', 'error'); return; }

  const tabLabel = activeTab === 'all' ? 'ทั้งหมด' : (catFromTab(activeTab) || 'ทั้งหมด');
  const title = `ทำเนียบอาจารย์ — ${tabLabel}` + (selectedYear ? ` ปีการศึกษา ${selectedYear}` : '');

  // อาจารย์พิเศษ — ตารางแบบย่อ (ชื่อ · ตำแหน่ง · หน่วยงาน · รายวิชาที่สอน)
  const specialOnly = activeTab === 'special';
  let headers, tableRows = '';
  if (specialOnly) {
    headers = ['#', 'ชื่อ-สกุล', 'ตำแหน่ง', 'หน่วยงาน', 'รายวิชาที่สอน'];
    data.forEach((row, i) => {
      tableRows += `<tr>
        <td style="text-align:center">${i + 1}</td>
        <td>${(row.name || '').replace(/\|\|/g, ', ')}</td>
        <td>${row.academic_position || ''}</td>
        <td>${row.agency || ''}</td>
        <td>${(row.subjects_taught || '').replace(/\|\|/g, '<br>')}</td>
      </tr>`;
    });
  } else {
    headers = ['#', 'ชื่อ-สกุล', 'เลขบัตรประชาชน', 'เลขใบประกอบวิชาชีพ', 'ตำแหน่ง/ตำแหน่งวิชาการ', 'วุฒิการศึกษา', 'ประสบการณ์สอน (ปี)', 'ประสบการณ์ปฏิบัติการ (ปี)', 'ผลงานวิชาการ (ย้อนหลัง 5 ปี)', 'ประเภทอาจารย์'];
    data.forEach((row, i) => {
      const special = isSpecialTeacher(row);
      tableRows += `<tr>
        <td style="text-align:center">${i + 1}</td>
        <td>${(row.name || '').replace(/\|\|/g, ', ')}</td>
        <td style="text-align:center">${special ? '' : maskNationalId(row.national_id)}</td>
        <td style="text-align:center">${special ? '' : (row.license_no || '')}</td>
        <td>${row.academic_position || ''}</td>
        <td>${special ? ('หน่วยงาน: ' + (row.agency || '-')) : (row.education || '').replace(/\|\|/g, '<br>')}</td>
        <td style="text-align:center">${special ? '' : (row.nursing_teaching_years || '')}</td>
        <td style="text-align:center">${special ? '' : (row.nursing_practice_years || '')}</td>
        <td>${special ? ('รายวิชาที่สอน: ' + (row.subjects_taught || '').replace(/\|\|/g, ', ')) : (row.academic_work || '').replace(/\|\|/g, '<br>')}</td>
        <td>${row.teacher_category || ''}</td>
      </tr>`;
    });
  }

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

// รวมข้อมูลทำเนียบอาจารย์ + อาจารย์พิเศษจากระบบทะเบียน (special_teacher)
// ที่ยังไม่มีในทำเนียบ (เทียบด้วยชื่อ + ปีการศึกษา) — เพื่อให้อาจารย์พิเศษที่ลงทะเบียนไว้แสดงในทำเนียบด้วย
function directoryRecords() {
  // แสดงเฉพาะข้อมูลที่ลงในระบบทำเนียบอาจารย์เท่านั้น (ไม่ดึงอาจารย์พิเศษจากระบบทะเบียนมาแสดงอัตโนมัติ)
  return getDataByType('teacher_directory');
}

function renderDirectoryDataSection(paged, total, counts, activeTab, isAdmin) {
  // Stat cards: ทั้งหมด + 5 ประเภท
  let html = '<div class="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3 mb-4">';
  html += '<div class="card-stat bg-white rounded-2xl p-4 border border-blue-100 text-center cursor-pointer ' + (activeTab === 'all' ? 'ring-2 ring-primary' : '') + '" onclick="APP._directoryTab=\'all\';APP.pagination.page=1;renderCurrentPage()"><p class="text-2xl font-bold text-primary">' + (counts.all || 0) + '</p><p class="text-xs text-gray-500">ทั้งหมด</p></div>';
  TEACHER_CATEGORIES.forEach(cat => {
    const m = TEACHER_CAT_META[cat];
    html += '<div class="card-stat bg-white rounded-2xl p-4 border border-gray-100 text-center cursor-pointer ' + (activeTab === m.tab ? 'ring-2 ' + m.ring : '') + '" onclick="APP._directoryTab=\'' + m.tab + '\';APP.pagination.page=1;renderCurrentPage()"><p class="text-2xl font-bold ' + m.num + '">' + (counts[m.tab] || 0) + '</p><p class="text-xs text-gray-500 leading-tight">' + cat + '</p></div>';
  });
  html += '</div>';
  html += filterBar({ semester: false, year: false });
  const isSpecialTab = activeTab === 'special';
  html += '<div class="bg-white rounded-2xl border border-blue-100 overflow-hidden"><div class="overflow-x-auto"><table class="w-full text-sm">';
  if (isSpecialTab) {
    html += '<thead><tr class="bg-surface text-left"><th class="px-4 py-3 font-semibold">ชื่อ-สกุล</th><th class="px-4 py-3 font-semibold">ตำแหน่ง</th><th class="px-4 py-3 font-semibold">หน่วยงาน</th><th class="px-4 py-3 font-semibold">รายวิชาที่สอน</th><th class="px-4 py-3"></th></tr></thead>';
  } else {
    html += '<thead><tr class="bg-surface text-left"><th class="px-4 py-3 font-semibold">ชื่อ-สกุล</th><th class="px-4 py-3 font-semibold">ตำแหน่ง</th><th class="px-4 py-3 font-semibold">ประเภท</th><th class="px-4 py-3"></th></tr></thead>';
  }
  html += '<tbody>';
  const colspan = isSpecialTab ? 5 : 4;
  if (paged.length) {
    paged.forEach(t => {
      const special = isSpecialTeacher(t);
      const fromReg = !!t.__fromRegistry;
      const editFn = fromReg ? 'showEditSpecialTeacherRegModal' : (special ? 'showEditSpecialTeacherModal' : 'showEditTeacherDirectoryModal');
      const regBadge = fromReg ? ' <span class="ml-1 px-1.5 py-0.5 rounded-md text-[10px] bg-emerald-50 text-emerald-600 border border-emerald-100" title="ดึงจากระบบทะเบียนอาจารย์พิเศษ">ทะเบียน</span>' : '';
      html += '<tr class="border-t hover:bg-gray-50">';
      html += '<td class="px-4 py-3 font-medium">' + (t.name || '') + regBadge + '</td>';
      html += '<td class="px-4 py-3">' + (t.academic_position || '-') + '</td>';
      if (isSpecialTab) {
        html += '<td class="px-4 py-3">' + (t.agency || '-') + '</td>';
        html += '<td class="px-4 py-3 text-xs">' + multiValList(t.subjects_taught) + '</td>';
      } else {
        html += '<td class="px-4 py-3"><span class="px-2 py-1 rounded-full text-xs ' + catBadge(t.teacher_category) + '">' + (t.teacher_category || '-') + '</span></td>';
      }
      html += '<td class="px-4 py-3"><div class="flex gap-1">';
      html += '<button onclick="showTeacherDirectoryDetail(\'' + t.__backendId + '\')" class="text-gray-400 hover:text-primary" title="ดูข้อมูล"><i data-lucide="eye" class="w-4 h-4"></i></button>';
      if (isAdmin) {
        html += '<button onclick="' + editFn + '(\'' + t.__backendId + '\')" class="text-blue-400 hover:text-blue-600" title="' + (fromReg ? 'แก้ไขที่ระบบทะเบียนอาจารย์พิเศษ' : 'แก้ไข') + '"><i data-lucide="pencil" class="w-4 h-4"></i></button>';
        if (!fromReg) html += '<button onclick="deleteRecord(\'' + t.__backendId + '\')" class="text-red-400 hover:text-red-600" title="ลบ"><i data-lucide="trash-2" class="w-4 h-4"></i></button>';
      }
      html += '</div></td></tr>';
    });
  } else {
    html += '<tr><td colspan="' + colspan + '" class="px-4 py-8 text-center text-gray-400">ไม่มีข้อมูล</td></tr>';
  }
  html += '</tbody></table></div></div>';
  html += paginationHTML(total, APP.pagination.perPage, APP.pagination.page, 'changePage');
  return html;
}

function teacherDirectoryPage() {
  const isAdmin = isAdminRole();
  let allData = applyFilters(directoryRecords());
  // ประธานสาขาวิชา: เห็นเฉพาะอาจารย์ในสาขาเดียวกัน (จับคู่ด้วย nursing_branch)
  if (APP.currentRole === 'deptHead') { const _d = currentDept(); allData = allData.filter(x => deptEq(norm(x.nursing_branch), _d)); }

  // Academic year filter
  const allYears = [...new Set(allData.map(d => d.academic_year).filter(Boolean))].sort().reverse();
  const selectedYear = APP.filters._directoryYear || '';
  if (selectedYear) allData = allData.filter(d => (d.academic_year || '') === selectedYear);
  let data = allData;

  // Counts by category (keyed by tab id)
  const counts = { all: data.length };
  TEACHER_CATEGORIES.forEach(cat => { counts[TEACHER_CAT_META[cat].tab] = data.filter(d => (d.teacher_category || '') === cat).length; });

  // Tab filter
  const activeTab = APP._directoryTab || 'all';
  if (activeTab !== 'all') {
    const cat = catFromTab(activeTab);
    if (cat) data = data.filter(d => (d.teacher_category || '') === cat);
  }

  const total = data.length; const paged = paginate(data);
  const view = APP._directoryView || 'list';

  function viewBtn(id, label, icon) {
    const active = view === id;
    return `<button onclick="APP._directoryView='${id}';APP.pagination.page=1;renderCurrentPage()" class="flex items-center gap-2 px-4 py-2 rounded-xl text-sm font-medium transition ${active ? 'bg-primary text-white shadow' : 'bg-white text-gray-600 hover:bg-surface border border-gray-200'}"><i data-lucide="${icon}" class="w-4 h-4"></i>${label}</button>`;
  }

  return `<div class="flex flex-wrap items-center justify-between gap-3 mb-4">
    <h2 class="text-xl font-bold text-gray-800"><i data-lucide="award" class="w-6 h-6 inline mr-2"></i>ทำเนียบอาจารย์</h2>
    <div class="flex flex-wrap gap-2">
      ${view === 'list' ? `<button onclick="exportTeacherDirectoryPDF()" class="flex items-center gap-2 px-4 py-2 bg-red-500 text-white rounded-xl hover:bg-red-600 text-sm"><i data-lucide="file-text" class="w-4 h-4"></i>ส่งออก PDF</button>
      ${isAdmin ? `<button onclick="showAddTeacherDirectoryModal()" class="flex items-center gap-2 px-4 py-2 bg-primary text-white rounded-xl hover:bg-primaryDark text-sm"><i data-lucide="plus" class="w-4 h-4"></i>เพิ่มอาจารย์</button>
      <button onclick="showAddSpecialTeacherModal()" class="flex items-center gap-2 px-4 py-2 bg-emerald-600 text-white rounded-xl hover:bg-emerald-700 text-sm"><i data-lucide="user-plus" class="w-4 h-4"></i>เพิ่มอาจารย์พิเศษ</button>${csvUploadBtn('teacher_directory', 'name,national_id,license_no,academic_position,note,nursing_branch,edu_level,edu_field,teaching_type,agency,subjects_taught,education,nursing_teaching_years,nursing_teaching_months,nursing_teaching_exp,nursing_practice_years,nursing_practice_months,nursing_practice_exp,academic_work,academic_year,teacher_category')}` : ''}`
      : view === 'summary' ? (isAdmin && selectedYear ? `<button onclick="showEditDirectorySummaryModal('${selectedYear}')" class="flex items-center gap-2 px-4 py-2 bg-primary text-white rounded-xl hover:bg-primaryDark text-sm"><i data-lucide="pencil" class="w-4 h-4"></i>แก้ไขตัวเลขสรุป</button>
      <button onclick="printDirectorySummary('${selectedYear}')" class="flex items-center gap-2 px-4 py-2 bg-red-500 text-white rounded-xl hover:bg-red-600 text-sm"><i data-lucide="printer" class="w-4 h-4"></i>พิมพ์</button>` : '')
        : (isAdmin && selectedYear ? `<button onclick="showEditBranchSummaryModal('${selectedYear}')" class="flex items-center gap-2 px-4 py-2 bg-primary text-white rounded-xl hover:bg-primaryDark text-sm"><i data-lucide="pencil" class="w-4 h-4"></i>แก้ไขตัวเลขตามสาขา</button>
      <button onclick="printBranchSummary('${selectedYear}')" class="flex items-center gap-2 px-4 py-2 bg-red-500 text-white rounded-xl hover:bg-red-600 text-sm"><i data-lucide="printer" class="w-4 h-4"></i>พิมพ์</button>` : '')}
    </div>
  </div>

  <div class="flex flex-wrap items-center gap-2 mb-4">
    ${viewBtn('list', 'รายชื่ออาจารย์', 'list')}
    ${viewBtn('summary', 'สรุปอัตรากำลัง', 'bar-chart-3')}
    ${viewBtn('branch', 'สรุปตามสาขาวิชา', 'git-branch')}
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

  ${!selectedYear ? noYearSelectedMsg('ทำเนียบอาจารย์')
      : (view === 'summary' ? directorySummaryView(selectedYear, isAdmin)
        : view === 'branch' ? directoryBranchView(selectedYear, isAdmin)
          : renderDirectoryDataSection(paged, total, counts, activeTab, isAdmin))}`;
}

function showTeacherDirectoryDetail(id) {
  const t = APP.allData.find(d => d.__backendId === id); if (!t) return;
  const catColor = catBadge(t.teacher_category);

  function detailList(val) {
    if (!val) return '<span class="text-gray-400">-</span>';
    const items = val.split('||').map(v => v.trim()).filter(Boolean);
    if (!items.length) return '<span class="text-gray-400">-</span>';
    if (items.length === 1) return `<span class="text-sm">${items[0]}</span>`;
    return '<ul class="list-disc list-inside text-sm space-y-1">' + items.map(v => `<li>${v}</li>`).join('') + '</ul>';
  }

  // อาจารย์พิเศษ — รายละเอียดแบบย่อ (รองรับทั้งทำเนียบ และที่ดึงมาจากระบบทะเบียน special_teacher)
  if (isSpecialTeacher(t) || t.type === 'special_teacher') {
    const cat = t.teacher_category || 'อาจารย์พิเศษ';
    const subj = t.subjects_taught || norm(t.subjects || '').split(/[,;/]/).map(s => s.trim()).filter(Boolean).join('||');
    showModal('ข้อมูลอาจารย์พิเศษ', `
    <div class="space-y-3">
      <div class="flex items-center gap-3 mb-2"><div class="w-12 h-12 bg-emerald-600 rounded-full flex items-center justify-center"><i data-lucide="user" class="w-6 h-6 text-white"></i></div><div><p class="font-bold text-lg">${t.name || '-'}</p><span class="px-2 py-1 rounded-full text-xs ${catBadge(cat)}">${cat}</span></div></div>
      <div class="grid grid-cols-2 gap-3">
        ${infoRow('ตำแหน่ง', t.academic_position)}
        ${infoRow('หน่วยงาน', t.agency)}
        ${infoRow('ประเภทการสอน', t.teaching_type)}
        ${infoRow('ระดับวุฒิ', t.edu_level)}
        ${infoRow('ปีการศึกษา', t.academic_year)}
      </div>
      <div class="bg-surface rounded-xl p-3"><p class="text-xs text-gray-500 mb-1 font-semibold">รายวิชาที่สอน</p>${detailList(subj)}</div>
    </div>
  `);
    return;
  }

  showModal('ข้อมูลทำเนียบอาจารย์', `
    <div class="space-y-3">
      <div class="flex items-center gap-3 mb-2"><div class="w-12 h-12 bg-primary rounded-full flex items-center justify-center"><i data-lucide="user" class="w-6 h-6 text-white"></i></div><div><p class="font-bold text-lg">${t.name || '-'}</p><span class="px-2 py-1 rounded-full text-xs ${catColor}">${t.teacher_category || '-'}</span></div></div>
      <div class="grid grid-cols-2 gap-3">
        ${infoRow('เลขบัตรประชาชน', maskNationalId(t.national_id))}
        ${infoRow('เลขใบประกอบวิชาชีพ', t.license_no)}
        ${infoRow('ตำแหน่งทางวิชาการ', t.academic_position)}
        ${infoRow('สาขาวิชา', t.nursing_branch)}
        ${infoRow('ระดับวุฒิ', [t.edu_level, t.edu_field].filter(Boolean).join(' · '))}
        ${infoRow('ปีการศึกษา', t.academic_year)}
      </div>
      ${t.note ? `<div class="bg-amber-50 rounded-xl p-3 border border-amber-100"><p class="text-xs text-gray-500 mb-1 font-semibold">หมายเหตุ</p><p class="text-sm">${t.note}</p></div>` : ''}
      <div class="bg-surface rounded-xl p-3"><p class="text-xs text-gray-500 mb-1 font-semibold">วุฒิการศึกษา</p>${detailList(t.education)}</div>
      <div class="bg-surface rounded-xl p-3"><p class="text-xs text-gray-500 mb-1 font-semibold">ประสบการณ์สอนทางการพยาบาล ${expYM(t.nursing_teaching_years, t.nursing_teaching_months) ? '<span class="text-primary font-bold">(รวม ' + expYM(t.nursing_teaching_years, t.nursing_teaching_months) + ')</span>' : ''}</p>${expDetailList(t.nursing_teaching_exp)}</div>
      <div class="bg-surface rounded-xl p-3"><p class="text-xs text-gray-500 mb-1 font-semibold">ประสบการณ์ปฏิบัติการพยาบาล ${expYM(t.nursing_practice_years, t.nursing_practice_months) ? '<span class="text-primary font-bold">(รวม ' + expYM(t.nursing_practice_years, t.nursing_practice_months) + ')</span>' : ''}</p>${expDetailList(t.nursing_practice_exp)}</div>
      <div class="bg-surface rounded-xl p-3"><p class="text-xs text-gray-500 mb-1 font-semibold">ผลงานวิชาการ (ย้อนหลัง 5 ปี)</p>${detailList(t.academic_work)}</div>
    </div>
  `);
}

function showAddTeacherDirectoryModal() {
  showModal('เพิ่มอาจารย์ (ทำเนียบ)', `
    <form id="addTeacherDirForm" class="space-y-3" style="max-height:70vh;overflow-y:auto;padding-right:4px">
      ${titlePrefixField('')}
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-xs text-gray-600 mb-1">เลขบัตรประชาชน</label><input name="national_id" maxlength="13" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="13 หลัก"></div>
        <div><label class="block text-xs text-gray-600 mb-1">เลขใบประกอบวิชาชีพ</label><input name="license_no" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      </div>
      <div><label class="block text-xs text-gray-600 mb-1">ตำแหน่งทางวิชาการ</label><input name="academic_position" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="เช่น ผศ.ดร., รศ."></div>
      <div><label class="block text-xs text-gray-600 mb-1">หมายเหตุ</label><textarea name="note" rows="2" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="หมายเหตุเพิ่มเติม (ถ้ามี)"></textarea></div>
      ${branchEduFields({})}
      ${multiInputField('education', 'วุฒิการศึกษา', 'เช่น พย.บ., พย.ม., ปร.ด.', '')}
      ${multiExpField('nursing_teaching_exp', 'ประสบการณ์สอนทางการพยาบาล', 'เช่น สอนวิชาการพยาบาลผู้ใหญ่', '')}
      ${multiExpField('nursing_practice_exp', 'ประสบการณ์ปฏิบัติการพยาบาล', 'เช่น พยาบาลวิชาชีพ รพ.รามาธิบดี', '')}
      ${multiAcademicWorkField('')}
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-xs text-gray-600 mb-1">ปีการศึกษา *</label><input name="academic_year" required class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="เช่น 2568" value="2568"></div>
        <div><label class="block text-xs text-gray-600 mb-1">ประเภทอาจารย์ *</label>
          <select name="teacher_category" required class="w-full border rounded-xl px-3 py-2 text-sm">
            <option value="">-- เลือกประเภท --</option>
            ${TEACHER_FULL_CATEGORIES.map(c => `<option>${c}</option>`).join('')}
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
      obj.name = combineName(form);
      obj.national_id = form.querySelector('[name="national_id"]').value;
      obj.license_no = form.querySelector('[name="license_no"]').value;
      obj.academic_position = form.querySelector('[name="academic_position"]').value;
      obj.note = form.querySelector('[name="note"]').value;
      obj.nursing_branch = form.querySelector('[name="nursing_branch"]').value;
      obj.edu_level = form.querySelector('[name="edu_level"]').value;
      obj.edu_field = form.querySelector('[name="edu_field"]').value;
      obj.education = collectMultiInputs(form, 'education');
      const _te = collectExp(form, 'nursing_teaching_exp');
      obj.nursing_teaching_exp = _te.exp; obj.nursing_teaching_years = _te.years; obj.nursing_teaching_months = _te.months;
      const _pe = collectExp(form, 'nursing_practice_exp');
      obj.nursing_practice_exp = _pe.exp; obj.nursing_practice_years = _pe.years; obj.nursing_practice_months = _pe.months;
      obj.academic_work = collectAcademicWork(form);
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
      ${titlePrefixField(t.name || '')}
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-xs text-gray-600 mb-1">เลขบัตรประชาชน</label><input name="national_id" value="${t.national_id || ''}" maxlength="13" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">เลขใบประกอบวิชาชีพ</label><input name="license_no" value="${t.license_no || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      </div>
      <div><label class="block text-xs text-gray-600 mb-1">ตำแหน่งทางวิชาการ</label><input name="academic_position" value="${t.academic_position || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <div><label class="block text-xs text-gray-600 mb-1">หมายเหตุ</label><textarea name="note" rows="2" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="หมายเหตุเพิ่มเติม (ถ้ามี)">${t.note || ''}</textarea></div>
      ${branchEduFields(t)}
      ${multiInputField('education', 'วุฒิการศึกษา', 'เช่น พย.บ., พย.ม., ปร.ด.', t.education || '')}
      ${multiExpField('nursing_teaching_exp', 'ประสบการณ์สอนทางการพยาบาล', 'เช่น สอนวิชาการพยาบาลผู้ใหญ่', t.nursing_teaching_exp || '', t.nursing_teaching_years || '', t.nursing_teaching_months || '')}
      ${multiExpField('nursing_practice_exp', 'ประสบการณ์ปฏิบัติการพยาบาล', 'เช่น พยาบาลวิชาชีพ รพ.รามาธิบดี', t.nursing_practice_exp || '', t.nursing_practice_years || '', t.nursing_practice_months || '')}
      ${multiAcademicWorkField(t.academic_work || '')}
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-xs text-gray-600 mb-1">ปีการศึกษา</label><input name="academic_year" value="${t.academic_year || ''}" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="เช่น 2568"></div>
        <div><label class="block text-xs text-gray-600 mb-1">ประเภทอาจารย์</label>
          <select name="teacher_category" class="w-full border rounded-xl px-3 py-2 text-sm">
            <option value="">-- เลือกประเภท --</option>
            ${TEACHER_FULL_CATEGORIES.map(c => `<option ${(t.teacher_category || '') === c ? 'selected' : ''}>${c}</option>`).join('')}
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
      rec.name = combineName(form);
      rec.national_id = form.querySelector('[name="national_id"]').value;
      rec.license_no = form.querySelector('[name="license_no"]').value;
      rec.academic_position = form.querySelector('[name="academic_position"]').value;
      rec.note = form.querySelector('[name="note"]').value;
      rec.nursing_branch = form.querySelector('[name="nursing_branch"]').value;
      rec.edu_level = form.querySelector('[name="edu_level"]').value;
      rec.edu_field = form.querySelector('[name="edu_field"]').value;
      rec.education = collectMultiInputs(form, 'education');
      const _te = collectExp(form, 'nursing_teaching_exp');
      rec.nursing_teaching_exp = _te.exp; rec.nursing_teaching_years = _te.years; rec.nursing_teaching_months = _te.months;
      const _pe = collectExp(form, 'nursing_practice_exp');
      rec.nursing_practice_exp = _pe.exp; rec.nursing_practice_years = _pe.years; rec.nursing_practice_months = _pe.months;
      rec.academic_work = collectAcademicWork(form);
      rec.academic_year = form.querySelector('[name="academic_year"]').value;
      rec.teacher_category = form.querySelector('[name="teacher_category"]').value;
      const r = await GSheetDB.update(rec);
      if (r.isOk) { showToast('แก้ไขข้อมูลสำเร็จ'); closeModal(); renderCurrentPage(); } else showToast('เกิดข้อผิดพลาด: ' + (r.error || ''), 'error');
    });
  };
}

// ---- อาจารย์พิเศษ (ฟอร์มย่อ: ชื่อ-สกุล · ตำแหน่ง · หน่วยงาน · รายวิชาที่สอน) ----
// เก็บในแท็บ teacher_directory เดิม: ตำแหน่ง→academic_position, หน่วยงาน→agency, รายวิชา→subjects_taught
function specialTeacherFormBody(t) {
  t = t || {};
  return `
    ${titlePrefixField(t.name || '')}
    <div class="grid grid-cols-2 gap-3">
      <div><label class="block text-xs text-gray-600 mb-1">ตำแหน่ง</label><input name="academic_position" value="${(t.academic_position || '').replace(/"/g, '&quot;')}" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="เช่น นายแพทย์ชำนาญการ"></div>
      <div><label class="block text-xs text-gray-600 mb-1">หน่วยงาน</label><input name="agency" value="${(t.agency || '').replace(/"/g, '&quot;')}" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="เช่น รพ.ราชวิถี"></div>
    </div>
    ${multiSubjectField('subjects_taught', 'รายวิชาที่สอน (เลือกได้หลายวิชา)', t.subjects_taught || '')}
    <div class="grid grid-cols-2 gap-3 p-3 bg-blue-50 rounded-xl border border-blue-100">
      <div><label class="block text-xs text-gray-600 mb-1">ประเภทการสอน <span class="text-gray-400">(สำหรับสรุป)</span></label>
        <select name="teaching_type" class="w-full border rounded-xl px-3 py-2 text-sm"><option value="">-- เลือก --</option>${['ภาคทฤษฎี', 'ภาคปฏิบัติ'].map(v => `<option ${norm(t.teaching_type) === v ? 'selected' : ''}>${v}</option>`).join('')}</select>
      </div>
      <div><label class="block text-xs text-gray-600 mb-1">ระดับวุฒิ <span class="text-gray-400">(สำหรับสรุป)</span></label>
        <select name="edu_level" class="w-full border rounded-xl px-3 py-2 text-sm"><option value="">-- เลือก --</option>${['ปริญญาเอก', 'ปริญญาโท', 'ปริญญาตรี'].map(v => `<option ${norm(t.edu_level) === v ? 'selected' : ''}>${v}</option>`).join('')}</select>
      </div>
    </div>
    <div><label class="block text-xs text-gray-600 mb-1">ปีการศึกษา *</label><input name="academic_year" required value="${(t.academic_year || '2568').replace(/"/g, '&quot;')}" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="เช่น 2568"></div>
    <input type="hidden" name="teacher_category" value="อาจารย์พิเศษ">`;
}

function collectSpecialTeacher(form, obj) {
  obj.name = combineName(form);
  obj.academic_position = form.querySelector('[name="academic_position"]').value;
  obj.agency = form.querySelector('[name="agency"]').value;
  obj.subjects_taught = collectMultiInputs(form, 'subjects_taught');
  obj.teaching_type = form.querySelector('[name="teaching_type"]').value;
  obj.edu_level = form.querySelector('[name="edu_level"]').value;
  obj.academic_year = form.querySelector('[name="academic_year"]').value;
  obj.teacher_category = 'อาจารย์พิเศษ';
  return obj;
}

// ดึงรายการอาจารย์พิเศษจากระบบทะเบียน (special_teacher) มาเป็นตัวเลือกเติมข้อมูลอัตโนมัติ
function specialRegOptionHTML(t) {
  return `<option value="${t.__backendId}">${(t.name || '').replace(/"/g, '&quot;')}${t.academic_position ? ' · ' + (t.academic_position || '').replace(/"/g, '&quot;') : ''}${t.agency ? ' · ' + (t.agency || '').replace(/"/g, '&quot;') : ''}</option>`;
}
function specialTeacherRegPickerHTML() {
  const list = getDataByType('special_teacher');
  if (!list.length) return '';
  const sorted = list.slice()
    .sort((a, b) => norm(b.academic_year).localeCompare(norm(a.academic_year)) || (a.name || '').localeCompare(b.name || ''));
  window._specialRegPickerList = sorted;
  const opts = sorted.map(specialRegOptionHTML).join('');
  return `<div class="p-3 bg-emerald-50 rounded-xl border border-emerald-100">
    <label class="block text-xs font-medium text-emerald-800 mb-1"><i data-lucide="download" class="w-3.5 h-3.5 inline mr-1"></i>ดึงข้อมูลจาก "ข้อมูลอาจารย์พิเศษ" (ระบบทะเบียน)</label>
    <input type="text" id="specialRegPickerSearch" oninput="filterSpecialTeacherReg(this.value)" placeholder="พิมพ์ค้นหาชื่อ-สกุล / ตำแหน่ง / หน่วยงาน..." class="w-full border rounded-xl px-3 py-2 text-sm mb-2">
    <select id="specialRegPickerSelect" onchange="fillFromSpecialTeacherReg(this)" class="w-full border rounded-xl px-3 py-2 text-sm">
      <option value="">-- เลือกเพื่อกรอกข้อมูลอัตโนมัติ --</option>${opts}
    </select>
  </div>`;
}
function filterSpecialTeacherReg(q) {
  const sel = document.getElementById('specialRegPickerSelect'); if (!sel) return;
  const kw = norm(q).toLowerCase();
  const list = window._specialRegPickerList || [];
  const matched = kw ? list.filter(t => (norm(t.name) + ' ' + norm(t.academic_position) + ' ' + norm(t.agency)).toLowerCase().includes(kw)) : list;
  sel.innerHTML = `<option value="">-- ${matched.length ? 'เลือกเพื่อกรอกข้อมูลอัตโนมัติ' : 'ไม่พบรายชื่อที่ค้นหา'} --</option>` + matched.map(specialRegOptionHTML).join('');
  // ถ้าเหลือรายชื่อเดียว เลือกให้อัตโนมัติ
  if (matched.length === 1) { sel.value = matched[0].__backendId; fillFromSpecialTeacherReg(sel); }
}

function fillFromSpecialTeacherReg(sel) {
  const id = sel.value; if (!id) return;
  const t = getDataByType('special_teacher').find(x => x.__backendId === id); if (!t) return;
  const form = sel.closest('form'); if (!form) return;
  const setv = (n, v) => { const el = form.querySelector('[name="' + n + '"]'); if (el) el.value = v || ''; };
  const { prefix, rest } = parseTitlePrefix(t.name || '');
  setv('title_prefix', prefix);
  setv('name', rest);
  setv('academic_position', t.academic_position);
  setv('agency', t.agency);
  // ดึงเฉพาะ ชื่อ-สกุล/ตำแหน่ง/หน่วยงาน — ไม่ดึงปีการศึกษา (เพิ่มในปีอื่นได้)
  const tp = form.querySelector('[name="title_prefix"]'); if (tp) syncTitlePrefixChips(tp);
}

function showAddSpecialTeacherModal() {
  showModal('เพิ่มอาจารย์พิเศษ', `
    <form id="addSpecialTeacherForm" class="space-y-3" style="max-height:70vh;overflow-y:auto;padding-right:4px">
      ${specialTeacherRegPickerHTML()}
      ${specialTeacherFormBody({})}
      <button type="submit" class="w-full bg-emerald-600 text-white py-2.5 rounded-xl hover:bg-emerald-700">บันทึก</button>
    </form>
  `);
  document.getElementById('addSpecialTeacherForm').onsubmit = async (e) => {
    e.preventDefault();
    await withLoading(e.target, async () => {
      const obj = collectSpecialTeacher(e.target, { type: 'teacher_directory', created_at: new Date().toISOString() });
      const r = await GSheetDB.create(obj);
      if (r.isOk) { showToast('เพิ่มอาจารย์พิเศษสำเร็จ'); closeModal(); } else showToast('เกิดข้อผิดพลาด', 'error');
    });
  };
}

function showEditSpecialTeacherModal(id) {
  const t = APP.allData.find(d => d.__backendId === id); if (!t) return;
  showModal('แก้ไขข้อมูลอาจารย์พิเศษ', `
    <form id="editSpecialTeacherForm" class="space-y-3" style="max-height:70vh;overflow-y:auto;padding-right:4px">
      ${specialTeacherFormBody(t)}
      <button type="submit" class="w-full bg-emerald-600 text-white py-2.5 rounded-xl hover:bg-emerald-700">บันทึกการแก้ไข</button>
    </form>
  `);
  document.getElementById('editSpecialTeacherForm').onsubmit = async (e) => {
    e.preventDefault();
    await withLoading(e.target, async () => {
      const rec = APP.allData.find(d => d.__backendId === id); if (!rec) return;
      collectSpecialTeacher(e.target, rec);
      const r = await GSheetDB.update(rec);
      if (r.isOk) { showToast('แก้ไขข้อมูลสำเร็จ'); closeModal(); renderCurrentPage(); } else showToast('เกิดข้อผิดพลาด: ' + (r.error || ''), 'error');
    });
  };
}

// ======================== TEACHER DIRECTORY — สรุป/แดชบอร์ด ========================
// ตัวเลขที่กรอกเองเก็บเป็น JSON ในแท็บ directory_summary (1 แถว/ปีการศึกษา)

// จัดกลุ่มตำแหน่งทางวิชาการจากข้อความอิสระ
function classifyRank(pos) {
  const s = norm(pos);
  if (/รองศาสตราจารย์|(^|[\s(.])รศ/.test(s)) return 'assoc';
  if (/ผู้ช่วยศาสตราจารย์|(^|[\s(.])ผศ/.test(s)) return 'asst';
  if (/ศาสตราจารย์|(^|[\s(.])ศ\.?\s?ด|(^|[\s(.])ศ\s/.test(s)) return 'prof';
  return 'ajarn';
}

// จัดวุฒิอาจารย์ (เอก/โท × สาขาพยาบาล/สัมพันธ์) จากข้อความวุฒิ
function classifyTeacherEdu(eduStr) {
  const items = (eduStr || '').split('||').map(v => norm(v)).filter(Boolean);
  const isPhd = s => /ปร\.?ด|ปริญญาเอก|ph\.?\s?d|ดุษฎี|d\.?n\.?s|ด\.\s?$/i.test(s);
  const isMaster = s => /ปริญญาโท|พย\.?ม|วท\.?ม|ศศ\.?ม|ค\.?ม|กศ\.?ม|สม\.|m\.?\s?(s|sc|a|ed|ns|ph)/i.test(s);
  const phd = items.find(isPhd);
  const master = items.find(isMaster);
  if (phd) return { level: 'phd', nursing: /พยาบาล|nurs/i.test(phd) };
  if (master) return { level: 'master', nursing: /พยาบาล|nurs/i.test(master) };
  return { level: '', nursing: false };
}

// จัดวุฒิ: ใช้ฟิลด์ที่เลือกไว้ (edu_level + edu_field) ก่อน ถ้าไม่มีค่อย fallback เดาจากข้อความ
function teacherEduClass(t) {
  const lvl = norm(t.edu_level);
  if (lvl) {
    let level = ''; if (/เอก/.test(lvl)) level = 'phd'; else if (/โท/.test(lvl)) level = 'master'; else if (/ตรี/.test(lvl)) level = 'bachelor';
    const fld = norm(t.edu_field);
    const nursing = fld ? (fld === 'สาขาการพยาบาล' || (/พยาบาล/.test(fld) && !/สัมพันธ์/.test(fld))) : /พยาบาล/.test(norm(t.education));
    return { level, nursing };
  }
  return classifyTeacherEdu(t.education);
}
// นับระดับวุฒิ (เอก/โท/ตรี) จาก edu_level
function eduLevelKey(t) {
  const lvl = norm(t.edu_level);
  if (/เอก/.test(lvl)) return 'phd'; if (/โท/.test(lvl)) return 'master'; if (/ตรี/.test(lvl)) return 'bachelor';
  const e = classifyTeacherEdu(t.education); return e.level || '';
}

// คำนวณตัวเลขอัตโนมัติจากข้อมูล teacher_directory ของปีที่เลือก
function computeDirectoryAuto(year) {
  let data = getDataByType('teacher_directory');
  if (year) data = data.filter(d => norm(d.academic_year) === norm(year));
  const byCat = cat => data.filter(d => (d.teacher_category || '') === cat);
  const cResp = byCat('อาจารย์ผู้รับผิดชอบหลักสูตร');
  const cCurr = byCat('อาจารย์ประจำหลักสูตร');
  const cReg = byCat('อาจารย์ประจำ');
  const cLeave = byCat('อาจารย์ที่ลาศึกษาต่อ');
  const cSpec = byCat('อาจารย์พิเศษ');
  const rankCount = (arr, r) => arr.filter(t => classifyRank(t.academic_position) === r).length;

  const edu = { phd_nursing: 0, phd_related: 0, master_nursing: 0, master_related: 0 };
  cCurr.forEach(t => {
    const e = teacherEduClass(t);
    if (e.level === 'phd') { e.nursing ? edu.phd_nursing++ : edu.phd_related++; }
    else if (e.level === 'master') { e.nursing ? edu.master_nursing++ : edu.master_related++; }
  });

  // อาจารย์ภายนอก/พิเศษ — แยกทฤษฎี/ปฏิบัติ × ระดับวุฒิ (จากข้อมูล อ.พิเศษ)
  const specTheory = cSpec.filter(t => /ทฤษฎี/.test(norm(t.teaching_type)));
  const specPractice = cSpec.filter(t => /ปฏิบัติ/.test(norm(t.teaching_type)));
  const lvlCount = (arr, key) => arr.filter(t => eduLevelKey(t) === key).length;

  const internal = [...cResp, ...cCurr, ...cReg, ...cLeave];
  const exp = { le5: 0, b6_10: 0, b10_15: 0, b15_20: 0, gt20: 0 };
  internal.forEach(t => {
    if (!norm(t.nursing_teaching_years) && !norm(t.nursing_teaching_months)) return;
    const y = expToYears(t.nursing_teaching_years, t.nursing_teaching_months);
    if (y <= 5) exp.le5++; else if (y <= 10) exp.b6_10++; else if (y <= 15) exp.b10_15++; else if (y <= 20) exp.b15_20++; else exp.gt20++;
  });

  const works = cCurr.map(t => (t.academic_work || '').split('||').map(v => v.trim()).filter(Boolean).length).filter(n => n > 0);
  const workMin = works.length ? Math.min(...works) : '';
  const workMax = works.length ? Math.max(...works) : '';

  return {
    internal_total: cResp.length + cCurr.length + cReg.length + cLeave.length,
    responsible_total: cResp.length,
    responsible_ajarn: rankCount(cResp, 'ajarn'), responsible_asst: rankCount(cResp, 'asst'),
    responsible_assoc: rankCount(cResp, 'assoc'), responsible_prof: rankCount(cResp, 'prof'),
    curriculum_total: cCurr.length,
    curriculum_ajarn: rankCount(cCurr, 'ajarn'), curriculum_asst: rankCount(cCurr, 'asst'),
    curriculum_assoc: rankCount(cCurr, 'assoc'), curriculum_prof: rankCount(cCurr, 'prof'),
    edu_phd_nursing: edu.phd_nursing, edu_phd_related: edu.phd_related,
    edu_master_nursing: edu.master_nursing, edu_master_related: edu.master_related,
    work_min: workMin, work_max: workMax,
    regular_total: cReg.length, studyleave_total: cLeave.length, special_total: cSpec.length,
    exp_le5: exp.le5, exp_6_10: exp.b6_10, exp_10_15: exp.b10_15, exp_15_20: exp.b15_20, exp_gt20: exp.gt20,
    external_total: cSpec.length,
    theory_total: specTheory.length, theory_phd: lvlCount(specTheory, 'phd'), theory_master: lvlCount(specTheory, 'master'), theory_bachelor: lvlCount(specTheory, 'bachelor'),
    practice_total: specPractice.length, practice_master: lvlCount(specPractice, 'master'), practice_bachelor: lvlCount(specPractice, 'bachelor')
  };
}

// ฟิลด์ทั้งหมดของหน้าสรุป (ขับฟอร์มแก้ไข + การแสดงผล)
const DS_FIELDS = [
  { sec: 'อาจารย์ผู้สอนภายใน', key: 'internal_total', label: 'อาจารย์ภายใน รวมทั้งหมด (คน)' },
  { sec: 'อาจารย์ผู้รับผิดชอบหลักสูตร', key: 'responsible_total', label: 'รวมทั้งหมด' },
  { key: 'responsible_ajarn', label: '• ตำแหน่งอาจารย์' },
  { key: 'responsible_asst', label: '• ผู้ช่วยศาสตราจารย์' },
  { key: 'responsible_assoc', label: '• รองศาสตราจารย์' },
  { key: 'responsible_prof', label: '• ศาสตราจารย์' },
  { sec: 'อาจารย์ประจำหลักสูตร', key: 'curriculum_total', label: 'รวมทั้งหมด (รวมผู้รับผิดชอบหลักสูตร)' },
  { key: 'curriculum_asst', label: '• ผู้ช่วยศาสตราจารย์' },
  { key: 'curriculum_ajarn', label: '• ตำแหน่งอาจารย์' },
  { key: 'curriculum_assoc', label: '• รองศาสตราจารย์' },
  { key: 'curriculum_prof', label: '• ศาสตราจารย์' },
  { key: 'edu_phd_nursing', label: 'วุฒิ ป.เอก สาขาการพยาบาล' },
  { key: 'edu_phd_related', label: 'วุฒิ ป.เอก สาขาที่สัมพันธ์ฯ' },
  { key: 'edu_master_nursing', label: 'วุฒิ ป.โท สาขาการพยาบาล' },
  { key: 'edu_master_related', label: 'วุฒิ ป.โท สาขาที่สัมพันธ์ฯ' },
  { key: 'work_min', label: 'ผลงานวิชาการ ต่ำสุด (ผลงาน)' },
  { key: 'work_max', label: 'ผลงานวิชาการ สูงสุด (ผลงาน)' },
  { sec: 'ประเภทอื่น', key: 'regular_total', label: 'อาจารย์ประจำ' },
  { key: 'studyleave_total', label: 'อาจารย์ที่ลาศึกษาต่อ' },
  { key: 'special_total', label: 'อาจารย์พิเศษ' },
  { sec: 'ประสบการณ์การสอนทางการพยาบาล', key: 'exp_le5', label: 'น้อยกว่าหรือเท่ากับ 5 ปี' },
  { key: 'exp_6_10', label: 'มากกว่า 6 ถึง 10 ปี' },
  { key: 'exp_10_15', label: 'มากกว่า 10 ถึง 15 ปี' },
  { key: 'exp_15_20', label: 'มากกว่า 15 ถึง 20 ปี' },
  { key: 'exp_gt20', label: 'มากกว่า 20 ปีขึ้นไป' },
  { sec: 'อาจารย์ผู้สอนภายนอก/อาจารย์พิเศษ', key: 'external_total', label: 'รวมทั้งหมด' },
  { key: 'theory_total', label: 'ภาคทฤษฎี รวม' },
  { key: 'theory_phd', label: '• ระดับปริญญาเอก' },
  { key: 'theory_master', label: '• ระดับปริญญาโท' },
  { key: 'theory_bachelor', label: '• ระดับปริญญาตรี' },
  { key: 'practice_total', label: 'ภาคปฏิบัติ รวม' },
  { key: 'practice_master', label: '• ระดับปริญญาโท' },
  { key: 'practice_bachelor', label: '• ระดับปริญญาตรี' }
];

function getDirectorySummary(year) {
  const rec = getDataByType('directory_summary').find(d => norm(d.academic_year) === norm(year));
  if (!rec) return {};
  try { return JSON.parse(rec.summary_json || '{}'); } catch (e) { return {}; }
}

// บันทึก JSON เต็มลงแถวของปีนั้น (สร้างใหม่ถ้ายังไม่มี)
function persistDirectorySummary(year, fullObj) {
  const rec = APP.allData.find(d => d.type === 'directory_summary' && norm(d.academic_year) === norm(year));
  if (rec) { rec.summary_json = JSON.stringify(fullObj); return GSheetDB.update(rec); }
  return GSheetDB.create({ type: 'directory_summary', academic_year: year, summary_json: JSON.stringify(fullObj), created_at: new Date().toISOString() });
}
// บันทึกตัวเลขแดชบอร์ด 1 (flat keys) โดยคงค่า branch_summary ของแดชบอร์ด 2 ไว้
function saveDirectorySummaryFields(year, fieldsObj) {
  const existing = getDirectorySummary(year);
  const merged = Object.assign({}, fieldsObj);
  if (existing.branch_summary) merged.branch_summary = existing.branch_summary;
  return persistDirectorySummary(year, merged);
}
// บันทึกตัวเลขแดชบอร์ด 2 (ตามสาขา) โดยคงตัวเลขแดชบอร์ด 1 ไว้
function saveDirectoryBranchSummary(year, branchObj) {
  const existing = getDirectorySummary(year);
  existing.branch_summary = branchObj;
  return persistDirectorySummary(year, existing);
}

// ค่าแสดงผล: ใช้ค่าที่บันทึกเองก่อน ถ้าไม่มีใช้ค่าที่คำนวณอัตโนมัติ
function dsResolve(saved, auto, key) {
  if (saved && saved[key] !== undefined && String(saved[key]) !== '') return saved[key];
  if (auto && auto[key] !== undefined && String(auto[key]) !== '') return auto[key];
  return '';
}

// กราฟโดนัท (CSS conic-gradient)
function dsDonut(segments, centerLabel) {
  const total = segments.reduce((s, x) => s + (parseFloat(x.value) || 0), 0);
  let acc = 0;
  const stops = segments.map(s => {
    const start = total ? acc / total * 360 : 0; acc += (parseFloat(s.value) || 0);
    const end = total ? acc / total * 360 : 0; return `${s.color} ${start}deg ${end}deg`;
  }).join(', ');
  return `<div style="display:flex;align-items:center;gap:18px;flex-wrap:wrap">
    <div style="width:150px;height:150px;border-radius:50%;background:conic-gradient(${stops || '#e5e7eb 0deg 360deg'});position:relative;flex-shrink:0">
      <div style="position:absolute;inset:26px;background:#fff;border-radius:50%;display:flex;flex-direction:column;align-items:center;justify-content:center"><span style="font-size:26px;font-weight:700;color:#1e6fba">${total || 0}</span><span style="font-size:11px;color:#94a3b8">${centerLabel || 'รวม'}</span></div>
    </div>
    <div style="flex:1;min-width:180px">${segments.map(s => `<div style="display:flex;align-items:center;gap:8px;margin-bottom:6px;font-size:13px"><span style="width:13px;height:13px;border-radius:3px;background:${s.color};flex-shrink:0"></span><span style="flex:1;color:#475569">${s.label}</span><b style="color:#1e293b">${s.value || 0}</b></div>`).join('')}</div>
  </div>`;
}

// กราฟแท่งแนวนอน (CSS)
function dsBars(items, color) {
  const max = Math.max(1, ...items.map(i => parseFloat(i.value) || 0));
  return `<div class="space-y-3">${items.map(i => {
    const w = Math.round((parseFloat(i.value) || 0) / max * 100);
    return `<div><div style="display:flex;justify-content:space-between;font-size:13px;margin-bottom:3px"><span style="color:#475569">${i.label}</span><b style="color:#1e293b">${i.value || 0}</b></div><div style="background:#eef2f7;border-radius:7px;height:16px;overflow:hidden"><div style="width:${w}%;height:100%;background:${color};border-radius:7px;transition:width .3s"></div></div></div>`;
  }).join('')}</div>`;
}

function dsBigCard(num, label, color, sub) {
  return `<div class="bg-white rounded-2xl p-5 border border-gray-100 shadow-sm">
    <p class="text-3xl font-bold" style="color:${color}">${num === '' || num === undefined ? '-' : num}</p>
    <p class="text-sm text-gray-600 mt-1">${label}</p>
    ${sub ? `<p class="text-xs text-gray-400 mt-1">${sub}</p>` : ''}
  </div>`;
}

function directorySummaryView(year, isAdmin) {
  const saved = getDirectorySummary(year);
  const auto = computeDirectoryAuto(year);
  const v = k => { const r = dsResolve(saved, auto, k); return r === '' ? '' : r; };
  const n = k => { const r = v(k); return r === '' ? '-' : r; };
  const hasSaved = saved && Object.keys(saved).length;

  // การ์ดภาพรวม
  let html = `<div class="grid grid-cols-2 md:grid-cols-4 gap-3 mb-5">
    ${dsBigCard(n('internal_total'), 'อาจารย์ผู้สอนภายใน', '#1e6fba', 'รวมทุกประเภทภายใน')}
    ${dsBigCard(n('external_total'), 'อาจารย์ผู้สอนภายนอก/พิเศษ', '#0ea5e9', 'ทฤษฎี + ปฏิบัติ')}
    ${dsBigCard(n('curriculum_total'), 'อาจารย์ประจำหลักสูตร', '#7c3aed', 'รวมผู้รับผิดชอบหลักสูตร')}
    ${dsBigCard(n('responsible_total'), 'อาจารย์ผู้รับผิดชอบหลักสูตร', '#4f46e5', '')}
  </div>`;

  // การ์ดประเภทย่อย
  html += `<div class="grid grid-cols-2 md:grid-cols-4 gap-3 mb-5">
    ${dsBigCard(n('regular_total'), 'อาจารย์ประจำ', '#2563eb', '')}
    ${dsBigCard(n('studyleave_total'), 'อาจารย์ที่ลาศึกษาต่อ', '#d97706', '')}
    ${dsBigCard(n('special_total'), 'อาจารย์พิเศษ (ภายใน)', '#059669', '')}
    ${dsBigCard(n('external_total'), 'อาจารย์ภายนอกทั้งหมด', '#0891b2', '')}
  </div>`;

  // กราฟ: วุฒิ (โดนัท) + ประสบการณ์ (แท่ง)
  html += `<div class="grid grid-cols-1 lg:grid-cols-2 gap-4 mb-5">
    <div class="bg-white rounded-2xl p-5 border border-blue-100">
      <h3 class="font-bold text-gray-800 mb-3 text-sm"><i data-lucide="graduation-cap" class="w-4 h-4 inline mr-1"></i>วุฒิการศึกษา — อาจารย์ประจำหลักสูตร</h3>
      ${dsDonut([
    { label: 'ป.เอก สาขาการพยาบาล', value: v('edu_phd_nursing'), color: '#7c3aed' },
    { label: 'ป.เอก สาขาที่สัมพันธ์ฯ', value: v('edu_phd_related'), color: '#a78bfa' },
    { label: 'ป.โท สาขาการพยาบาล', value: v('edu_master_nursing'), color: '#2563eb' },
    { label: 'ป.โท สาขาที่สัมพันธ์ฯ', value: v('edu_master_related'), color: '#60a5fa' }
  ], 'คน')}
    </div>
    <div class="bg-white rounded-2xl p-5 border border-blue-100">
      <h3 class="font-bold text-gray-800 mb-3 text-sm"><i data-lucide="clock" class="w-4 h-4 inline mr-1"></i>ประสบการณ์การสอนทางการพยาบาล</h3>
      ${dsBars([
    { label: '≤ 5 ปี', value: v('exp_le5') },
    { label: '> 6–10 ปี', value: v('exp_6_10') },
    { label: '> 10–15 ปี', value: v('exp_10_15') },
    { label: '> 15–20 ปี', value: v('exp_15_20') },
    { label: '> 20 ปีขึ้นไป', value: v('exp_gt20') }
  ], '#1e6fba')}
    </div>
  </div>`;

  // กราฟ: อาจารย์ภายนอก ทฤษฎี/ปฏิบัติ
  html += `<div class="grid grid-cols-1 lg:grid-cols-2 gap-4 mb-5">
    <div class="bg-white rounded-2xl p-5 border border-blue-100">
      <h3 class="font-bold text-gray-800 mb-3 text-sm"><i data-lucide="book-open" class="w-4 h-4 inline mr-1"></i>อาจารย์ภายนอก — ภาคทฤษฎี (${n('theory_total')} คน)</h3>
      ${dsBars([
    { label: 'ปริญญาเอก', value: v('theory_phd') },
    { label: 'ปริญญาโท', value: v('theory_master') },
    { label: 'ปริญญาตรี', value: v('theory_bachelor') }
  ], '#0ea5e9')}
    </div>
    <div class="bg-white rounded-2xl p-5 border border-blue-100">
      <h3 class="font-bold text-gray-800 mb-3 text-sm"><i data-lucide="activity" class="w-4 h-4 inline mr-1"></i>อาจารย์ภายนอก — ภาคปฏิบัติ (${n('practice_total')} คน)</h3>
      ${dsBars([
    { label: 'ปริญญาโท', value: v('practice_master') },
    { label: 'ปริญญาตรี', value: v('practice_bachelor') }
  ], '#06b6d4')}
    </div>
  </div>`;

  // รายงานข้อความ (เหมือนรูปต้นฉบับ)
  html += dsTextReport(year, v, n);

  if (!hasSaved) {
    html += `<div class="mt-4 p-3 bg-amber-50 border border-amber-200 rounded-xl text-sm text-amber-800"><i data-lucide="info" class="w-4 h-4 inline mr-1"></i>ตัวเลขที่แสดงมาจากการคำนวณอัตโนมัติ ${isAdmin ? 'กดปุ่ม "แก้ไขตัวเลขสรุป" ด้านบนเพื่อปรับ/เติมค่าที่ระบบคำนวณไม่ได้ (เช่น อาจารย์ภายนอก)' : ''}</div>`;
  }
  return html;
}

// รายงานข้อความตามรูปแบบเอกสารต้นฉบับ
function dsTextReport(year, v, n) {
  const row = (label, num) => `<div style="display:flex;justify-content:space-between;gap:12px;padding:2px 0"><span>${label}</span><span style="white-space:nowrap"><b>${num}</b> คน</span></div>`;
  return `<div class="bg-white rounded-2xl p-5 border border-blue-100" id="dsTextReport">
    <h3 class="font-bold text-gray-800 mb-3"><i data-lucide="file-text" class="w-5 h-5 inline mr-1"></i>สรุปอัตรากำลังอาจารย์ผู้สอน ปีการศึกษา ${year}</h3>
    <div class="text-sm text-gray-700 leading-relaxed space-y-3">
      <div>
        <p class="font-bold text-primary">อาจารย์ผู้สอนภายใน มีจำนวน ${n('internal_total')} คน โดยแบ่งได้ดังนี้</p>
        <div class="pl-3 mt-2 space-y-2">
          <div>
            <p class="font-semibold">- อาจารย์ผู้รับผิดชอบหลักสูตร จำนวนทั้งหมด ${n('responsible_total')} คน</p>
            <p class="text-gray-500 text-xs pl-3">แยกตามตำแหน่งทางวิชาการ: อาจารย์ ${n('responsible_ajarn')} คน · ผศ. ${n('responsible_asst')} คน · รศ. ${n('responsible_assoc')} คน · ศ. ${n('responsible_prof')} คน</p>
          </div>
          <div>
            <p class="font-semibold">- อาจารย์ประจำหลักสูตร จำนวนทั้งหมด ${n('curriculum_total')} คน (รวมอาจารย์ผู้รับผิดชอบหลักสูตร)</p>
            <p class="text-gray-500 text-xs pl-3">แยกตามตำแหน่งทางวิชาการ: ผศ. ${n('curriculum_asst')} คน · อาจารย์ ${n('curriculum_ajarn')} คน · รศ. ${n('curriculum_assoc')} คน · ศ. ${n('curriculum_prof')} คน</p>
            <div class="pl-3 mt-1 max-w-md">
              <p class="text-xs text-gray-500 mb-1">แยกตามวุฒิการศึกษา:</p>
              ${row('ปริญญาเอกในสาขาการพยาบาล', n('edu_phd_nursing'))}
              ${row('ปริญญาเอกในสาขาที่สัมพันธ์ทางการพยาบาล', n('edu_phd_related'))}
              ${row('ปริญญาโทในสาขาการพยาบาล', n('edu_master_nursing'))}
              ${row('ปริญญาโทในสาขาที่สัมพันธ์ทางการพยาบาล', n('edu_master_related'))}
              <p class="text-xs text-gray-400 mt-1">(ผลงานทางวิชาการของอาจารย์ประจำหลักสูตรอยู่ระหว่าง ${n('work_min')} – ${n('work_max')} ผลงาน)</p>
            </div>
          </div>
          <p class="font-semibold">- อาจารย์ประจำ จำนวนทั้งหมด ${n('regular_total')} คน</p>
          <p class="font-semibold">- อาจารย์ที่ลาศึกษาต่อ จำนวนทั้งหมด ${n('studyleave_total')} คน</p>
          <p class="font-semibold">- อาจารย์พิเศษ จำนวนทั้งหมด ${n('special_total')} คน</p>
        </div>
      </div>
      <div>
        <p class="font-bold text-primary">ประสบการณ์การสอนทางการพยาบาล</p>
        <div class="pl-3 mt-1 max-w-md">
          ${row('น้อยกว่าหรือเท่ากับ 5 ปี', n('exp_le5'))}
          ${row('มากกว่า 6 ถึง 10 ปี', n('exp_6_10'))}
          ${row('มากกว่า 10 ปี ถึง 15 ปี', n('exp_10_15'))}
          ${row('มากกว่า 15 ปี ถึง 20 ปี', n('exp_15_20'))}
          ${row('มากกว่า 20 ปีขึ้นไป', n('exp_gt20'))}
        </div>
      </div>
      <div>
        <p class="font-bold text-primary">อาจารย์ผู้สอนภายนอก/อาจารย์พิเศษ มีจำนวน ${n('external_total')} คน โดยแบ่งได้ดังนี้</p>
        <div class="pl-3 mt-1 space-y-1">
          <p class="font-semibold">- ภาคทฤษฎี จำนวน ${n('theory_total')} คน</p>
          <div class="pl-3 max-w-md">
            ${row('ระดับปริญญาเอก', n('theory_phd'))}
            ${row('ระดับปริญญาโท', n('theory_master'))}
            ${row('ระดับปริญญาตรี', n('theory_bachelor'))}
          </div>
          <p class="font-semibold">- ภาคปฏิบัติ จำนวน ${n('practice_total')} คน</p>
          <div class="pl-3 max-w-md">
            ${row('ระดับปริญญาโท', n('practice_master'))}
            ${row('ระดับปริญญาตรี', n('practice_bachelor'))}
          </div>
        </div>
      </div>
    </div>
  </div>`;
}

function showEditDirectorySummaryModal(year) {
  const saved = getDirectorySummary(year);
  const auto = computeDirectoryAuto(year);
  window.__dsAuto = auto;
  let body = '';
  DS_FIELDS.forEach(f => {
    if (f.sec) body += `<p class="text-sm font-bold text-primary mt-3 mb-1 border-b border-gray-200 pb-1">${f.sec}</p>`;
    const cur = (saved && saved[f.key] !== undefined) ? saved[f.key] : '';
    const hint = (auto[f.key] !== undefined && String(auto[f.key]) !== '') ? 'auto: ' + auto[f.key] : 'กรอกเอง';
    body += `<div class="flex items-center gap-2 mb-1"><label class="text-xs text-gray-600 flex-1">${f.label}</label><input name="${f.key}" value="${cur}" inputmode="decimal" class="w-24 border rounded-lg px-2 py-1 text-sm text-right" placeholder="${hint}"></div>`;
  });
  showModal('แก้ไขตัวเลขสรุป — ปีการศึกษา ' + year, `
    <form id="dsForm" style="max-height:70vh;overflow-y:auto;padding-right:4px">
      <div class="bg-blue-50 rounded-xl p-3 text-xs text-blue-800 mb-3">ช่องที่เว้นว่าง = ใช้ค่าที่ระบบคำนวณอัตโนมัติ (ตัวเลข auto ในช่อง placeholder) · กรอกตัวเลขเพื่อกำหนดเอง รองรับทศนิยม เช่น 48.5</div>
      <button type="button" onclick="dsAutoFill()" class="mb-3 text-sm px-3 py-1.5 bg-emerald-600 text-white rounded-lg hover:bg-emerald-700"><i data-lucide="wand-2" class="w-4 h-4 inline"></i> เติมค่าอัตโนมัติทั้งหมด</button>
      ${body}
      <button type="submit" class="w-full mt-3 bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึก</button>
    </form>
  `);
  lucide.createIcons();
  document.getElementById('dsForm').onsubmit = async (e) => {
    e.preventDefault();
    await withLoading(e.target, async () => {
      const obj = {};
      DS_FIELDS.forEach(f => { const el = e.target.querySelector('[name="' + f.key + '"]'); if (el && el.value.trim() !== '') obj[f.key] = el.value.trim(); });
      const r = await saveDirectorySummaryFields(year, obj);
      if (r.isOk) { showToast('บันทึกข้อมูลสรุปสำเร็จ'); closeModal(); renderCurrentPage(); } else showToast('เกิดข้อผิดพลาด: ' + (r.error || ''), 'error');
    });
  };
}

function dsAutoFill() {
  const auto = window.__dsAuto || {};
  const form = document.getElementById('dsForm'); if (!form) return;
  DS_FIELDS.forEach(f => {
    const el = form.querySelector('[name="' + f.key + '"]');
    if (el && auto[f.key] !== undefined && String(auto[f.key]) !== '') el.value = auto[f.key];
  });
  showToast('เติมค่าอัตโนมัติแล้ว — ปรับแก้ได้ตามต้องการ');
}

function printDirectorySummary(year) {
  const el = document.getElementById('dsTextReport');
  if (!el) { showToast('ไม่พบรายงาน', 'error'); return; }
  const printWin = window.open('', '_blank');
  printWin.document.write(`<!DOCTYPE html><html lang="th"><head><meta charset="UTF-8"><title>สรุปอัตรากำลังอาจารย์ ${year}</title>
    <link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;600;700&display=swap" rel="stylesheet">
    <style>*{margin:0;padding:0;box-sizing:border-box}body{font-family:'Sarabun',sans-serif;font-size:14px;color:#222;padding:24px;line-height:1.7}@page{size:A4;margin:18mm}.font-bold{font-weight:700}.text-primary{color:#1e6fba}b{font-weight:600}.pl-3{padding-left:14px}.space-y-1>*+*,.space-y-2>*+*,.space-y-3>*+*{margin-top:6px}h3{font-size:17px;color:#1e6fba;margin-bottom:10px}.max-w-md{max-width:520px}[data-lucide]{display:none}</style>
    </head><body>${el.innerHTML}</body></html>`);
  printWin.document.close();
  printWin.onload = () => setTimeout(() => printWin.print(), 400);
  showToast('เปิดหน้าต่างพิมพ์แล้ว — เลือก "Save as PDF" ได้');
}

// ======================== แดชบอร์ดที่ 2 — สรุปตามสาขาวิชา ========================
const DB_EDU_KEYS = [
  { key: 'phd_nursing', label: 'ปริญญาเอกในสาขาการพยาบาล', short: 'ป.เอก พยาบาล', color: '#7c3aed' },
  { key: 'phd_related', label: 'ปริญญาเอกในสาขาที่สัมพันธ์ทางการพยาบาล', short: 'ป.เอก สัมพันธ์ฯ', color: '#a78bfa' },
  { key: 'master_nursing', label: 'ปริญญาโทในสาขาการพยาบาล', short: 'ป.โท พยาบาล', color: '#2563eb' },
  { key: 'master_related', label: 'ปริญญาโทในสาขาที่สัมพันธ์ทางการพยาบาล', short: 'ป.โท สัมพันธ์ฯ', color: '#60a5fa' }
];

// คำนวณอัตโนมัติ: อาจารย์ประจำหลักสูตร แยกตามสาขา × วุฒิ
function computeBranchAuto(year) {
  let data = getDataByType('teacher_directory').filter(d => (d.teacher_category || '') === 'อาจารย์ประจำหลักสูตร');
  if (year) data = data.filter(d => norm(d.academic_year) === norm(year));
  const out = {};
  const ensure = b => { if (!out[b]) out[b] = { phd_nursing: 0, phd_related: 0, master_nursing: 0, master_related: 0 }; return out[b]; };
  data.forEach(t => {
    const b = norm(t.nursing_branch) || 'ไม่ระบุสาขา';
    const o = ensure(b);
    const e = teacherEduClass(t);
    if (e.level === 'phd') e.nursing ? o.phd_nursing++ : o.phd_related++;
    else if (e.level === 'master') e.nursing ? o.master_nursing++ : o.master_related++;
  });
  return out;
}

// แท่งสัดส่วน (stacked) สำหรับ 1 สาขา
function dbStack(parts) {
  const total = parts.reduce((s, p) => s + (p.value || 0), 0) || 1;
  return `<div style="display:flex;height:18px;border-radius:6px;overflow:hidden;background:#eef2f7">${parts.map(p => p.value ? `<div title="${p.label}: ${p.value}" style="width:${p.value / total * 100}%;background:${p.color}"></div>` : '').join('')}</div>`;
}

function directoryBranchView(year, isAdmin) {
  const saved = (getDirectorySummary(year).branch_summary) || {};
  const auto = computeBranchAuto(year);
  const names = [...new Set([...NURSING_BRANCHES, ...Object.keys(auto), ...Object.keys(saved)])];
  const bget = (branch, key) => {
    const s = saved[branch]; if (s && s[key] !== undefined && String(s[key]) !== '') return parseFloat(s[key]) || 0;
    const a = auto[branch]; if (a && a[key] !== undefined) return a[key];
    return 0;
  };
  const bTotal = branch => DB_EDU_KEYS.reduce((s, k) => s + bget(branch, k.key), 0);
  const shown = names.filter(b => bTotal(b) > 0);
  const grand = shown.reduce((s, b) => s + bTotal(b), 0);

  if (!shown.length) {
    return `<div class="bg-white rounded-2xl p-8 border border-blue-100 text-center text-gray-400">
      <i data-lucide="git-branch" class="w-10 h-10 mx-auto mb-2"></i>
      <p>ยังไม่มีข้อมูลสาขาวิชาของอาจารย์ประจำหลักสูตร ปีการศึกษา ${year}</p>
      <p class="text-xs mt-1">กรุณาระบุ "สาขาวิชา" และ "ระดับวุฒิ/กลุ่มสาขาวุฒิ" ในข้อมูลอาจารย์ ${isAdmin ? 'หรือกดปุ่ม "แก้ไขตัวเลขตามสาขา" เพื่อกรอกเอง' : ''}</p>
    </div>`;
  }

  // การ์ดรวม + โดนัทภาพรวม
  const eduTotals = DB_EDU_KEYS.map(k => ({ label: k.short, value: shown.reduce((s, b) => s + bget(b, k.key), 0), color: k.color }));
  let html = `<div class="grid grid-cols-1 lg:grid-cols-3 gap-4 mb-5">
    <div class="bg-white rounded-2xl p-5 border border-gray-100 shadow-sm flex flex-col justify-center">
      <p class="text-3xl font-bold text-primary">${grand}</p>
      <p class="text-sm text-gray-600 mt-1">อาจารย์ประจำหลักสูตร</p>
      <p class="text-xs text-gray-400 mt-1">จำแนกได้ ${shown.length} สาขาวิชา</p>
    </div>
    <div class="bg-white rounded-2xl p-5 border border-blue-100 lg:col-span-2">
      <h3 class="font-bold text-gray-800 mb-3 text-sm"><i data-lucide="graduation-cap" class="w-4 h-4 inline mr-1"></i>ภาพรวมวุฒิการศึกษา (ทุกสาขา)</h3>
      ${dsDonut(eduTotals, 'คน')}
    </div>
  </div>`;

  // การ์ดต่อสาขา + แท่งสัดส่วน
  html += '<div class="grid grid-cols-1 md:grid-cols-2 gap-4 mb-5">';
  shown.forEach(b => {
    const parts = DB_EDU_KEYS.map(k => ({ label: k.short, value: bget(b, k.key), color: k.color }));
    html += `<div class="bg-white rounded-2xl p-4 border border-blue-100">
      <div class="flex items-center justify-between mb-2"><h3 class="font-bold text-gray-800 text-sm">${b}</h3><span class="px-2 py-1 rounded-full text-xs bg-primaryLight text-primary font-bold">${bTotal(b)} คน</span></div>
      ${dbStack(parts)}
      <div class="grid grid-cols-2 gap-x-3 gap-y-1 mt-3 text-xs">
        ${DB_EDU_KEYS.map(k => bget(b, k.key) ? `<div class="flex items-center gap-1"><span style="width:10px;height:10px;border-radius:2px;background:${k.color};display:inline-block"></span><span class="text-gray-600 flex-1">${k.short}</span><b>${bget(b, k.key)}</b></div>` : '').join('')}
      </div>
    </div>`;
  });
  html += '</div>';

  // รายงานข้อความตามรูปต้นฉบับ
  html += `<div class="bg-white rounded-2xl p-5 border border-blue-100" id="dbTextReport">
    <h3 class="font-bold text-gray-800 mb-3"><i data-lucide="file-text" class="w-5 h-5 inline mr-1"></i>สรุปจำนวนอาจารย์ประจำหลักสูตรจำแนกตามสาขาวิชา ปีการศึกษา ${year}</h3>
    <div class="text-sm text-gray-700 leading-relaxed">
      <p class="mb-2">ปีการศึกษา ${year} วิทยาลัยฯ มีอาจารย์ประจำหลักสูตร จำนวนทั้งหมด <b>${grand}</b> คน แยกตามสาขาวิชาและวุฒิการศึกษา มีดังนี้</p>
      <div class="space-y-3 pl-2">
        ${shown.map(b => `<div>
          <p class="font-semibold">- สาขา${b} จำนวนทั้งหมด ${bTotal(b)} คน</p>
          <div class="pl-4 max-w-lg">
            ${DB_EDU_KEYS.map(k => bget(b, k.key) ? `<div style="display:flex;justify-content:space-between;gap:12px;padding:1px 0"><span>${k.label}</span><span><b>${bget(b, k.key)}</b> คน</span></div>` : '').join('')}
          </div>
        </div>`).join('')}
      </div>
    </div>
  </div>`;

  const hasSaved = Object.keys(saved).length;
  if (!hasSaved) html += `<div class="mt-4 p-3 bg-amber-50 border border-amber-200 rounded-xl text-sm text-amber-800"><i data-lucide="info" class="w-4 h-4 inline mr-1"></i>ตัวเลขมาจากการคำนวณอัตโนมัติตามข้อมูลในระบบ ${isAdmin ? 'กดปุ่ม "แก้ไขตัวเลขตามสาขา" เพื่อปรับแก้เอง' : ''}</div>`;
  return html;
}

function showEditBranchSummaryModal(year) {
  const saved = (getDirectorySummary(year).branch_summary) || {};
  const auto = computeBranchAuto(year);
  window.__dbAuto = auto;
  const names = [...new Set([...NURSING_BRANCHES, ...Object.keys(auto), ...Object.keys(saved)])];

  const inpCell = (idx, branch, key) => {
    const sv = saved[branch] || {}; const av = auto[branch] || {};
    const val = sv[key] !== undefined ? sv[key] : '';
    const ph = av[key] !== undefined ? String(av[key]) : '0';
    return `<input name="b_${idx}_${key}" data-branch="${(branch || '').replace(/"/g, '&quot;')}" data-key="${key}" value="${val}" inputmode="numeric" class="w-14 border rounded-lg px-1 py-1 text-sm text-right" placeholder="${ph}">`;
  };
  const rowHTML = (idx, branch, editable) => `<tr class="border-t">
    <td class="px-1 py-1">${editable ? `<input name="bname_${idx}" value="${(branch || '').replace(/"/g, '&quot;')}" class="w-40 border rounded-lg px-2 py-1 text-sm" placeholder="ชื่อสาขา">` : `<input name="bname_${idx}" type="hidden" value="${(branch || '').replace(/"/g, '&quot;')}"><span class="text-xs">${branch}</span>`}</td>
    ${DB_EDU_KEYS.map(k => `<td class="px-1 py-1 text-center">${inpCell(idx, branch, k.key)}</td>`).join('')}
  </tr>`;

  const rows = names.map((b, i) => rowHTML(i, b, false)).join('');
  showModal('แก้ไขตัวเลขตามสาขา — ปีการศึกษา ' + year, `
    <form id="dbForm" style="max-height:70vh;overflow-y:auto;padding-right:4px">
      <div class="bg-blue-50 rounded-xl p-3 text-xs text-blue-800 mb-3">ช่องว่าง = ใช้ค่าที่คำนวณอัตโนมัติ (เลข auto ใน placeholder) · กรอกเพื่อกำหนดเอง</div>
      <button type="button" onclick="dbBranchAutoFill()" class="mb-3 text-sm px-3 py-1.5 bg-emerald-600 text-white rounded-lg hover:bg-emerald-700"><i data-lucide="wand-2" class="w-4 h-4 inline"></i> เติมค่าอัตโนมัติ</button>
      <div class="overflow-x-auto"><table class="w-full text-xs"><thead><tr class="text-gray-500"><th class="px-1 py-1 text-left">สาขาวิชา</th>${DB_EDU_KEYS.map(k => `<th class="px-1 py-1">${k.short}</th>`).join('')}</tr></thead>
      <tbody id="dbRows">${rows}</tbody></table></div>
      <button type="button" onclick="dbAddBranchRow()" class="mt-2 text-xs text-primary hover:underline">＋ เพิ่มสาขาใหม่</button>
      <button type="submit" class="w-full mt-3 bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึก</button>
    </form>
  `);
  lucide.createIcons();
  window.__dbRowIdx = names.length;
  document.getElementById('dbForm').onsubmit = async (e) => {
    e.preventDefault();
    await withLoading(e.target, async () => {
      const form = e.target;
      const result = {};
      form.querySelectorAll('[name^="bname_"]').forEach(nameEl => {
        const idx = nameEl.name.substring('bname_'.length);
        const branch = (nameEl.value || '').trim(); if (!branch) return;
        const obj = {};
        DB_EDU_KEYS.forEach(k => { const el = form.querySelector('[name="b_' + idx + '_' + k.key + '"]'); if (el && el.value.trim() !== '') obj[k.key] = el.value.trim(); });
        if (Object.keys(obj).length) result[branch] = obj;
      });
      const r = await saveDirectoryBranchSummary(year, result);
      if (r.isOk) { showToast('บันทึกข้อมูลตามสาขาสำเร็จ'); closeModal(); renderCurrentPage(); } else showToast('เกิดข้อผิดพลาด: ' + (r.error || ''), 'error');
    });
  };
}

function dbBranchAutoFill() {
  const auto = window.__dbAuto || {};
  document.querySelectorAll('#dbForm input[data-key]').forEach(el => {
    const b = el.getAttribute('data-branch'); const k = el.getAttribute('data-key');
    if (auto[b] && auto[b][k] !== undefined) el.value = auto[b][k];
  });
  showToast('เติมค่าอัตโนมัติแล้ว — ปรับแก้ได้');
}

function dbAddBranchRow() {
  const tbody = document.getElementById('dbRows'); if (!tbody) return;
  const idx = 'c' + (window.__dbRowIdx = (window.__dbRowIdx || 0) + 1);
  const tr = document.createElement('tr');
  tr.className = 'border-t';
  tr.innerHTML = `<td class="px-1 py-1"><input name="bname_${idx}" value="" class="w-40 border rounded-lg px-2 py-1 text-sm" placeholder="ชื่อสาขา"></td>` +
    DB_EDU_KEYS.map(k => `<td class="px-1 py-1 text-center"><input name="b_${idx}_${k.key}" value="" inputmode="numeric" class="w-14 border rounded-lg px-1 py-1 text-sm text-right" placeholder="0"></td>`).join('');
  tbody.appendChild(tr);
}

function printBranchSummary(year) {
  const el = document.getElementById('dbTextReport');
  if (!el) { showToast('ไม่พบรายงาน', 'error'); return; }
  const printWin = window.open('', '_blank');
  printWin.document.write(`<!DOCTYPE html><html lang="th"><head><meta charset="UTF-8"><title>สรุปอาจารย์ตามสาขา ${year}</title>
    <link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;600;700&display=swap" rel="stylesheet">
    <style>*{margin:0;padding:0;box-sizing:border-box}body{font-family:'Sarabun',sans-serif;font-size:14px;color:#222;padding:24px;line-height:1.7}@page{size:A4;margin:18mm}.font-bold{font-weight:700}b{font-weight:600}.pl-2{padding-left:8px}.pl-4{padding-left:18px}.space-y-3>*+*{margin-top:10px}h3{font-size:17px;color:#1e6fba;margin-bottom:10px}.max-w-lg{max-width:560px}[data-lucide]{display:none}</style>
    </head><body>${el.innerHTML}</body></html>`);
  printWin.document.close();
  printWin.onload = () => setTimeout(() => printWin.print(), 400);
  showToast('เปิดหน้าต่างพิมพ์แล้ว — เลือก "Save as PDF" ได้');
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
        <div><p class="font-medium text-sm">${a.announcement_title || ''}</p><p class="text-xs text-gray-500">${a.announcement_date || ''} · ${a.event_type || 'ทั่วไป'} · <span class="${annParseRoles(a.roles).length ? 'text-teal-600' : 'text-gray-400'}">${annParseRoles(a.roles).length ? annParseRoles(a.roles).map(r => ANN_ROLE_LABEL[r] || r).join('/') : 'ทุกบทบาท'}</span></p><p class="text-xs text-gray-600 mt-1">${(a.announcement_content || '').substring(0, 80)}</p></div>
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

// ======================== ANNOUNCEMENT ROLE TARGETING ========================
const ANN_ROLES = ['admin', 'academic', 'registrar', 'deptHead', 'executive', 'teacher', 'classTeacher', 'student'];
const ANN_ROLE_LABEL = { admin: 'ผู้ดูแลระบบ', academic: 'งานวิชาการ', registrar: 'งานทะเบียน', deptHead: 'ประธานสาขา', executive: 'ผู้บริหาร', teacher: 'อาจารย์', classTeacher: 'อ.ประจำชั้น', student: 'นักศึกษา' };
function annParseRoles(s) { return String(s == null ? '' : s).split(/[,|]/).map(x => x.trim()).filter(Boolean); }
// ประกาศนี้ผู้ใช้บทบาท role เห็นไหม — ว่าง = ทุกบทบาท
// แยกรายชื่อผู้รับเจาะจง (target_names) และทำ key เทียบชื่อ (ตัดช่องว่าง/คำนำหน้า)
function annParseNames(v) { return String(v == null ? '' : v).split(/[,;|]/).map(x => x.trim()).filter(Boolean); }
function annNameKey(name) {
  let n = norm(name).toLowerCase().replace(/\s+/g, '');
  let changed = true;
  while (changed) {
    changed = false;
    for (const pr of TITLE_PREFIXES) {
      const pk = String(pr).toLowerCase().replace(/\s+/g, '');
      if (pk && n.startsWith(pk)) { n = n.slice(pk.length); changed = true; break; }
    }
  }
  return n;
}
function annVisibleTo(a, role) {
  // ประกาศที่เจาะจงรายบุคคล (เช่น แจ้งผู้คุมสอบ) — เห็นเฉพาะคนที่มีชื่อ + ผู้ดูแล/งานวิชาการ
  const targets = annParseNames(a && a.target_names);
  if (targets.length) {
    if (role === 'admin' || role === 'academic') return true;
    const myKey = annNameKey((APP.currentUser && APP.currentUser.name) || '');
    return !!myKey && targets.some(n => annNameKey(n) === myKey);
  }
  const rs = annParseRoles(a && a.roles);
  return rs.length === 0 || rs.indexOf(role) !== -1;
}
// ประกาศที่บทบาทผู้ใช้ปัจจุบันมีสิทธิ์เห็น (ใช้กับกระดิ่ง + หน้าหลัก + badge)
function visibleAnnouncements() { const role = APP.currentRole; return getDataByType('announcement').filter(a => annVisibleTo(a, role)); }
// ช่องเลือกบทบาทผู้รับในฟอร์มประกาศ (ไม่ติ๊กเลย = ทุกบทบาท)
function annRolesFieldHTML(selected) {
  const sel = annParseRoles(selected);
  return `<div><label class="block text-xs text-gray-600 mb-1">แจ้งให้บทบาท <span class="text-gray-400">(ไม่เลือกเลย = ทุกบทบาท)</span></label>
    <div class="grid grid-cols-2 sm:grid-cols-4 gap-1.5 p-2 bg-gray-50 rounded-xl">
      ${ANN_ROLES.map(r => `<label class="flex items-center gap-1.5 text-sm text-gray-700"><input type="checkbox" class="ann-role-cb accent-primary" value="${r}" ${sel.indexOf(r) !== -1 ? 'checked' : ''}> ${ANN_ROLE_LABEL[r]}</label>`).join('')}
    </div></div>`;
}
function annCollectRoles() {
  const arr = Array.prototype.map.call(document.querySelectorAll('.ann-role-cb:checked'), el => el.value);
  return (arr.length === 0 || arr.length === ANN_ROLES.length) ? '' : arr.join(',');
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
      ${annRolesFieldHTML('')}
      <label class="flex items-center gap-2 bg-green-50 rounded-xl px-3 py-2 cursor-pointer"><input type="checkbox" name="line_notify" value="✓" class="w-4 h-4"><span class="text-sm text-green-700">📢 ส่งประกาศนี้เข้า LINE</span></label>
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึก</button>
    </form>
  `);
  document.getElementById('addAnnForm').onsubmit = async (e) => {
    e.preventDefault();
    await withLoading(e.target, async () => {
      const fd = new FormData(e.target);
      const obj = { type: 'announcement', created_at: new Date().toISOString() }; fd.forEach((v, k) => obj[k] = v);
      obj.roles = annCollectRoles();
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
  const isAdmin = isAdminRole();
  const isExecutive = APP.currentRole === 'executive';
  const canApprove = isAdmin || isExecutive;
  const canEdit = isAdmin || APP.currentRole === 'teacher' || APP.currentRole === 'classTeacher';
  let data = getDataByType('tracking').filter(t => t.subject_name && t.subject_name.trim());
  if (APP.currentRole === 'teacher') data = data.filter(t => subjectHasCoordinator({ coordinator: t.coordinator }, APP.currentUser.name));
  if (ctYear()) data = data.filter(t => norm(t.year_level) === ctYear());

  // Year filter for stats
  const selectedYear = APP.filters._trackingYear || '';
  const allSubjects = getDataByType('subject');
  let subjectsFiltered = selectedYear ? allSubjects.filter(s => norm(s.academic_year) === selectedYear) : allSubjects;
  // teacher: เห็นเฉพาะวิชาที่ตัวเองเป็นผู้ประสานงาน
  if (APP.currentRole === 'teacher') subjectsFiltered = subjectsFiltered.filter(s => subjectHasCoordinator(s, APP.currentUser.name));
  if (ctYear()) subjectsFiltered = subjectsFiltered.filter(s => norm(s.year_level) === ctYear());
  { const _df = applyDeptFilter(data, subjectsFiltered, '_' + APP.currentPage + 'Dept'); data = _df.data; subjectsFiltered = _df.subjects; }
  // dataForStats: รวม record ที่ academic_year ตรง หรือ academic_year ว่าง (บันทึกโดยไม่ระบุปี)
  const dataForStats = selectedYear ? data.filter(t => norm(t.academic_year) === selectedYear || !norm(t.academic_year)) : data;

  // Only show not-submitted and stats when year is selected
  let statsSection = '';
  let notSubmittedSection = '';
  if (selectedYear) {
    // Find subjects not yet in tracking
    // จับคู่ด้วย subject_code เป็นหลัก (ถ้ามี) — fallback เป็น subject_name + semester [+ academic_year]
    const isTracked = makeTrackingMatcher(dataForStats);
    const notSubmitted = subjectsFiltered.filter(s => !isTracked(s));

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
    const submitted = subjectsFiltered.filter(s => isTracked(s));
    if (notSubmitted.length) {
      notSubmittedSection = `<div class="bg-red-50 rounded-2xl p-4 border border-red-200 mb-4">
        <h3 onclick="this.parentElement.querySelector('.tracking-list-body').classList.toggle('hidden')" class="font-bold text-red-700 mb-2 text-sm flex items-center gap-2 cursor-pointer select-none"><i data-lucide="alert-triangle" class="w-4 h-4"></i>รายวิชาที่ยังไม่ส่งรายละเอียด (${notSubmitted.length} วิชา) <i data-lucide="chevron-down" class="w-4 h-4 ml-auto"></i></h3>
        <div class="flex flex-wrap gap-2 tracking-list-body">${notSubmitted.map(s => `<span class="px-3 py-1 bg-white border border-red-200 rounded-lg text-xs text-red-700">${s.subject_code ? s.subject_code + ' ' : ''}${s.subject_name || ''} <span class="text-gray-400">(ภาค ${s.semester || ''})</span></span>`).join('')}</div>
      </div>`;
    }
    if (submitted.length) {
      notSubmittedSection += `<div class="bg-green-50 rounded-2xl p-4 border border-green-200 mb-4">
        <h3 onclick="this.parentElement.querySelector('.tracking-list-body').classList.toggle('hidden')" class="font-bold text-green-700 mb-2 text-sm flex items-center gap-2 cursor-pointer select-none"><i data-lucide="check-circle" class="w-4 h-4"></i>รายวิชาที่ส่งรายละเอียดแล้ว (${submitted.length} วิชา) <i data-lucide="chevron-down" class="w-4 h-4 ml-auto"></i></h3>
        <div class="flex flex-wrap gap-2 tracking-list-body hidden">${submitted.map(s => `<span class="px-3 py-1 bg-white border border-green-200 rounded-lg text-xs text-green-700">${s.subject_code ? s.subject_code + ' ' : ''}${s.subject_name || ''} <span class="text-gray-400">(ภาค ${s.semester || ''})</span></span>`).join('')}</div>
      </div>`;
    }
  }

  // Year options
  const allYears = [...new Set([...allSubjects.map(s => norm(s.academic_year)), ...data.map(t => norm(t.academic_year))].filter(Boolean))].sort();

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
      ${trackingDeptFilterHTML('_trackingDept', APP.filters._trackingDept || '')}
    </div>
    ${statsSection}
  </div>
  ${selectedYear ? `${notSubmittedSection}
  ${filterBar({ yearLevel: true, semester: false, year: false })}
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
        <td class="px-4 py-3">${ctCell}${actionTimeText(t, 'class_teacher_check')}</td>
        <td class="px-4 py-3">${acCell}${actionTimeText(t, 'academic_propose')}</td>
        <td class="px-4 py-3">${dpCell}${actionTimeText(t, 'deputy_sign')}</td>
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

// ตัวเลือก "ลงข้อมูลย้อนหลัง" ในฟอร์มเพิ่มข้อมูลติดตาม — ติ๊กแล้วบันทึกเป็นเสร็จสิ้นครบทุกขั้น จึงไม่ไปสร้างแจ้งเตือน
function trackingBackfillCheckboxHTML() {
  return `<label class="flex items-center gap-2 bg-amber-50 border border-amber-100 rounded-xl px-3 py-2 cursor-pointer"><input type="checkbox" id="trackingBackfill" class="w-4 h-4"><span class="text-sm text-amber-800">🕘 ลงข้อมูลย้อนหลัง (เสร็จเรียบร้อยแล้ว — ไม่ต้องแจ้งเตือน)</span></label>`;
}
function applyTrackingBackfill(obj) {
  const cb = document.getElementById('trackingBackfill');
  if (!(cb && cb.checked)) return obj;
  const steps = (obj.type === 'grade_tracking' || obj.type === 'file_tracking')
    ? ['coordinator_check', 'academic_check', 'deputy_sign']
    : ['class_teacher_check', 'academic_propose', 'deputy_sign'];
  steps.forEach(k => obj[k] = 'เสร็จสิ้น');
  if (!obj.approved_date) obj.approved_date = new Date().toISOString().split('T')[0];
  return obj;
}

function showAddTrackingModal() {
  const subjects = getDataByType('subject');
  const subjectOptions = [...new Set(subjects.map(s => s.subject_name).filter(Boolean))].sort()
    .map(name => {
      const s = subjects.find(x => x.subject_name === name) || {};
      return `<option value="${name.replace(/"/g, '&quot;')}" data-code="${(s.subject_code || '').replace(/"/g, '&quot;')}" data-year="${(s.academic_year || '').replace(/"/g, '&quot;')}">${s.subject_code ? s.subject_code + ' ' : ''}${name}</option>`;
    }).join('');
  const teachers = getDataByType('teacher');
  const teacherList = [...new Set(teachers.map(t => (t.name || '').trim()).filter(Boolean))].sort();
  const myName = (APP.currentUser && APP.currentUser.name || '').trim();
  const coordCheckboxes = teacherList.map(name => {
    const isMe = name === myName;
    return `<label class="flex items-center gap-2 p-1.5 rounded-lg hover:bg-gray-50 cursor-pointer ${isMe ? 'bg-blue-50' : ''}"><input type="checkbox" class="coord-check w-4 h-4 rounded border-gray-300 text-primary focus:ring-primary" value="${name.replace(/"/g, '&quot;')}" ${isMe ? 'checked' : ''}><span class="text-sm">${name}${isMe ? ' <span class="text-xs text-blue-500">(คุณ)</span>' : ''}</span></label>`;
  }).join('');
  const currentYear = (APP.filters && APP.filters._trackingYear) || '2568';
  showModal('เพิ่มรายละเอียดรายวิชา', `
    <form id="addTrackingForm" class="space-y-3">
      <div><label class="block text-xs text-gray-600 mb-1">ชื่อรายวิชา *</label><select name="subject_name" id="trackingSubjectSelect" required class="w-full border rounded-xl px-3 py-2 text-sm"><option value="">-- เลือกรายวิชา --</option>${subjectOptions}</select></div>
      <input type="hidden" name="subject_code" id="trackingSubjectCode">
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-xs text-gray-600 mb-1">ทฤษฎี/ปฏิบัติ</label><select name="theory_practice" class="w-full border rounded-xl px-3 py-2 text-sm"><option>ทฤษฎี</option><option>ปฏิบัติ</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">ชั้นปี</label><select name="year_level" class="w-full border rounded-xl px-3 py-2 text-sm"><option>1</option><option>2</option><option>3</option><option>4</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">ห้อง</label><input name="room" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">ภาคการศึกษา</label><select name="semester" class="w-full border rounded-xl px-3 py-2 text-sm"><option value="1">1</option><option value="2">2</option><option value="3">ฤดูร้อน</option></select></div>
        <div class="col-span-2"><label class="block text-xs text-gray-600 mb-1">ปีการศึกษา</label><input name="academic_year" value="${currentYear}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      </div>
      <div>
        <label class="block text-xs text-gray-600 mb-1">ผู้ประสานงานรายวิชา (เลือกจากรายชื่อ หรือพิมพ์เพิ่มเติม)</label>
        <div class="max-h-40 overflow-y-auto border rounded-xl p-2 bg-gray-50 space-y-0.5">${coordCheckboxes || '<p class="text-xs text-gray-400 p-2">ยังไม่มีข้อมูลอาจารย์ในระบบ</p>'}</div>
        <input type="text" id="trackingCoordExtra" class="w-full border rounded-xl px-3 py-2 text-sm mt-2" placeholder="พิมพ์ชื่อเพิ่มเติม (คั่นด้วยเครื่องหมาย ,)">
        <input type="hidden" name="coordinator" id="trackingCoordHidden">
      </div>
      ${trackingBackfillCheckboxHTML()}
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึก</button>
    </form>
  `);
  // Auto-fill subject_code เมื่อเลือกรายวิชา
  const sel = document.getElementById('trackingSubjectSelect');
  if (sel) sel.addEventListener('change', () => {
    const opt = sel.options[sel.selectedIndex];
    const codeEl = document.getElementById('trackingSubjectCode');
    if (codeEl) codeEl.value = (opt && opt.dataset.code) || '';
  });
  document.getElementById('addTrackingForm').onsubmit = async (e) => {
    e.preventDefault();
    // Sync coordinator checkboxes to hidden field
    const checked = [...e.target.querySelectorAll('.coord-check:checked')].map(cb => cb.value);
    const extraText = (document.getElementById('trackingCoordExtra') || {}).value || '';
    const extraNames = extraText.split(',').map(s => s.trim()).filter(Boolean);
    const allNames = [...checked, ...extraNames].filter(Boolean);
    const coordHidden = document.getElementById('trackingCoordHidden');
    if (coordHidden) coordHidden.value = allNames.join(', ');
    // Ensure subject_code is populated (fallback ถ้าผู้ใช้ไม่ได้กดเปลี่ยน)
    const ss = document.getElementById('trackingSubjectSelect');
    const cc = document.getElementById('trackingSubjectCode');
    if (ss && cc && !cc.value) {
      const opt = ss.options[ss.selectedIndex];
      cc.value = (opt && opt.dataset.code) || '';
    }
    await withLoading(e.target, async () => {
      const fd = new FormData(e.target);
      const obj = { type: 'tracking', class_teacher_check: 'รอ', academic_propose: 'รอ', deputy_sign: 'รอ', approved_date: '', created_at: new Date().toISOString() };
      fd.forEach((v, k) => obj[k] = v);
      applyTrackingBackfill(obj);
      const r = await GSheetDB.create(obj);
      if (r.isOk) { showToast('เพิ่มสำเร็จ'); closeModal() } else showToast('เกิดข้อผิดพลาด', 'error');
    });
  };
}

// แสดงวัน-เวลาที่คลิกอนุมัติ/ส่งกลับแก้ไข (รูปแบบ พ.ศ.)
// ใช้กับ field: coordinator_check, class_teacher_check, academic_propose, academic_check, deputy_sign
function actionTimeText(rec, field) {
  const ts = rec && rec[field + '_at'];
  if (!ts) return '';
  const d = new Date(ts);
  if (isNaN(d)) return '';
  const dd = String(d.getDate()).padStart(2, '0');
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const yyyy = d.getFullYear() + 543;
  const hh = String(d.getHours()).padStart(2, '0');
  const mn = String(d.getMinutes()).padStart(2, '0');
  const by = rec[field + '_by'] ? ' โดย ' + rec[field + '_by'] : '';
  return `<div class="text-[10px] text-gray-400 mt-1 leading-tight" title="คลิกเมื่อ ${dd}/${mm}/${yyyy} ${hh}:${mn}${by}">🕒 ${dd}/${mm}/${yyyy} ${hh}:${mn}${by ? '<br>' + by : ''}</div>`;
}

async function updateTrackingField(id, field, value) {
  const el = window.event ? (window.event.target.closest('button') || window.event.target.closest('select')) : null;
  await withLoading(el, async () => {
    const rec = APP.allData.find(d => d.__backendId === id); if (!rec) return;
    rec[field] = value;
    // บันทึกวัน-เวลา และผู้คลิก ทุกครั้งที่มีการอนุมัติ/ส่งกลับแก้ไข
    rec[field + '_at'] = new Date().toISOString();
    rec[field + '_by'] = (APP.currentUser && APP.currentUser.name) || APP.currentRole || '';
    if (field === 'deputy_sign' && value === 'เสร็จสิ้น') rec.approved_date = new Date().toISOString().split('T')[0];

    const r = await GSheetDB.update(rec);
    if (r.isOk) { showToast('อัปเดตสำเร็จ'); renderCurrentPage(); updateNotifBadge(); } else showToast('เกิดข้อผิดพลาด', 'error');
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
  const isAdmin = isAdminRole();
  const isExecutive = APP.currentRole === 'executive';
  const isClassTeacher = APP.currentRole === 'classTeacher';
  const canEdit = isAdmin || APP.currentRole === 'teacher' || APP.currentRole === 'classTeacher';
  let data = getDataByType('result_tracking').filter(t => t.subject_name && t.subject_name.trim());
  if (APP.currentRole === 'teacher') data = data.filter(t => subjectHasCoordinator({ coordinator: t.coordinator }, APP.currentUser.name));
  if (ctYear()) data = data.filter(t => norm(t.year_level) === ctYear());

  const selectedYear = APP.filters._resultTrackingYear || '';
  const allSubjects = getDataByType('subject');
  let subjectsFiltered = selectedYear ? allSubjects.filter(s => norm(s.academic_year) === selectedYear) : allSubjects;
  if (APP.currentRole === 'teacher') subjectsFiltered = subjectsFiltered.filter(s => subjectHasCoordinator(s, APP.currentUser.name));
  if (ctYear()) subjectsFiltered = subjectsFiltered.filter(s => norm(s.year_level) === ctYear());
  { const _df = applyDeptFilter(data, subjectsFiltered, '_' + APP.currentPage + 'Dept'); data = _df.data; subjectsFiltered = _df.subjects; }
  const dataForStats = selectedYear ? data.filter(t => norm(t.academic_year) === selectedYear || !norm(t.academic_year)) : data;

  let statsSection = '';
  let notSubmittedSection = '';
  if (selectedYear) {
    const isTracked = makeTrackingMatcher(dataForStats);
    const notSubmitted = subjectsFiltered.filter(s => !isTracked(s));

    const completed = dataForStats.filter(t => t.deputy_sign === 'เสร็จสิ้น').length;
    const inProgress = dataForStats.filter(t => (t.class_teacher_check === 'เสร็จสิ้น' || t.academic_propose === 'เสร็จสิ้น') && t.deputy_sign !== 'เสร็จสิ้น').length;
    const pending = dataForStats.length - completed - inProgress;

    statsSection = `<div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3 mt-3">
      ${statCard('alert-circle', 'ยังไม่ส่ง', notSubmitted.length, 'วิชา', 'bg-red-500')}
      ${statCard('clock', 'รอดำเนินการ', pending, 'วิชา', 'bg-yellow-500')}
      ${statCard('loader', 'กำลังดำเนินการ', inProgress, 'วิชา', 'bg-blue-500')}
      ${statCard('check-circle', 'เสร็จสิ้น', completed, 'วิชา', 'bg-green-500')}
    </div>`;
    const submitted = subjectsFiltered.filter(s => isTracked(s));
    if (notSubmitted.length) {
      notSubmittedSection = `<div class="bg-red-50 rounded-2xl p-4 border border-red-200 mb-4">
        <h3 onclick="this.parentElement.querySelector('.tracking-list-body').classList.toggle('hidden')" class="font-bold text-red-700 mb-2 text-sm flex items-center gap-2 cursor-pointer select-none"><i data-lucide="alert-triangle" class="w-4 h-4"></i>รายวิชาที่ยังไม่ส่งผลการดำเนินงาน (${notSubmitted.length} วิชา) <i data-lucide="chevron-down" class="w-4 h-4 ml-auto"></i></h3>
        <div class="flex flex-wrap gap-2 tracking-list-body">${notSubmitted.map(s => `<span class="px-3 py-1 bg-white border border-red-200 rounded-lg text-xs text-red-700">${s.subject_code ? s.subject_code + ' ' : ''}${s.subject_name || ''} <span class="text-gray-400">(ภาค ${s.semester || ''})</span></span>`).join('')}</div>
      </div>`;
    }
    if (submitted.length) {
      notSubmittedSection += `<div class="bg-green-50 rounded-2xl p-4 border border-green-200 mb-4">
        <h3 onclick="this.parentElement.querySelector('.tracking-list-body').classList.toggle('hidden')" class="font-bold text-green-700 mb-2 text-sm flex items-center gap-2 cursor-pointer select-none"><i data-lucide="check-circle" class="w-4 h-4"></i>รายวิชาที่ส่งผลการดำเนินงานแล้ว (${submitted.length} วิชา) <i data-lucide="chevron-down" class="w-4 h-4 ml-auto"></i></h3>
        <div class="flex flex-wrap gap-2 tracking-list-body hidden">${submitted.map(s => `<span class="px-3 py-1 bg-white border border-green-200 rounded-lg text-xs text-green-700">${s.subject_code ? s.subject_code + ' ' : ''}${s.subject_name || ''} <span class="text-gray-400">(ภาค ${s.semester || ''})</span></span>`).join('')}</div>
      </div>`;
    }
  }

  const allYears = [...new Set([...allSubjects.map(s => norm(s.academic_year)), ...data.map(t => norm(t.academic_year))].filter(Boolean))].sort();

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
      ${trackingDeptFilterHTML('_resultTrackingDept', APP.filters._resultTrackingDept || '')}
    </div>
    ${statsSection}
  </div>
  ${selectedYear ? `${notSubmittedSection}
  ${filterBar({ yearLevel: true, semester: false, year: false })}
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
        <td class="px-4 py-3">${ctCell}${actionTimeText(t, 'class_teacher_check')}</td>
        <td class="px-4 py-3">${acCell}${actionTimeText(t, 'academic_propose')}</td>
        <td class="px-4 py-3">${dpCell}${actionTimeText(t, 'deputy_sign')}</td>
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
    .map(name => {
      const s = subjects.find(x => x.subject_name === name) || {};
      return `<option value="${name.replace(/"/g, '&quot;')}" data-code="${(s.subject_code || '').replace(/"/g, '&quot;')}">${s.subject_code ? s.subject_code + ' ' : ''}${name}</option>`;
    }).join('');
  const teachers = getDataByType('teacher');
  const teacherList = [...new Set(teachers.map(t => (t.name || '').trim()).filter(Boolean))].sort();
  const myName = (APP.currentUser && APP.currentUser.name || '').trim();
  const coordCheckboxes = teacherList.map(name => {
    const isMe = name === myName;
    return `<label class="flex items-center gap-2 p-1.5 rounded-lg hover:bg-gray-50 cursor-pointer ${isMe ? 'bg-blue-50' : ''}"><input type="checkbox" class="coord-check w-4 h-4 rounded border-gray-300 text-primary focus:ring-primary" value="${name.replace(/"/g, '&quot;')}" ${isMe ? 'checked' : ''}><span class="text-sm">${name}${isMe ? ' <span class="text-xs text-blue-500">(คุณ)</span>' : ''}</span></label>`;
  }).join('');
  showModal('เพิ่มข้อมูลติดตามการส่งผลการดำเนินงานรายวิชา', `
    <form id="addResultTrackingForm" class="space-y-3">
      <div><label class="block text-xs text-gray-600 mb-1">ชื่อรายวิชา *</label><select name="subject_name" id="resultTrackingSubjectSelect" required class="w-full border rounded-xl px-3 py-2 text-sm"><option value="">-- เลือกรายวิชา --</option>${subjectOptions}</select></div>
      <input type="hidden" name="subject_code" id="resultTrackingSubjectCode">
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
      ${trackingBackfillCheckboxHTML()}
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึก</button>
    </form>
  `);
  // Auto-fill subject_code
  const rSel = document.getElementById('resultTrackingSubjectSelect');
  if (rSel) rSel.addEventListener('change', () => {
    const opt = rSel.options[rSel.selectedIndex];
    const codeEl = document.getElementById('resultTrackingSubjectCode');
    if (codeEl) codeEl.value = (opt && opt.dataset.code) || '';
  });
  document.getElementById('addResultTrackingForm').onsubmit = async (e) => {
    e.preventDefault();
    const checked = [...e.target.querySelectorAll('.coord-check:checked')].map(cb => cb.value);
    const extraText = (document.getElementById('resultTrackingCoordExtra') || {}).value || '';
    const extraNames = extraText.split(',').map(s => s.trim()).filter(Boolean);
    const allNames = [...checked, ...extraNames].filter(Boolean);
    const coordHidden = document.getElementById('resultTrackingCoordHidden');
    if (coordHidden) coordHidden.value = allNames.join(', ');
    const ss = document.getElementById('resultTrackingSubjectSelect');
    const cc = document.getElementById('resultTrackingSubjectCode');
    if (ss && cc && !cc.value) {
      const opt = ss.options[ss.selectedIndex];
      cc.value = (opt && opt.dataset.code) || '';
    }
    await withLoading(e.target, async () => {
      const fd = new FormData(e.target);
      const obj = { type: 'result_tracking', class_teacher_check: 'รอ', academic_propose: 'รอ', deputy_sign: 'รอ', approved_date: '', created_at: new Date().toISOString() };
      fd.forEach((v, k) => obj[k] = v);
      applyTrackingBackfill(obj);
      const r = await GSheetDB.create(obj);
      if (r.isOk) { showToast('เพิ่มสำเร็จ'); closeModal() } else showToast('เกิดข้อผิดพลาด', 'error');
    });
  };
}
// ======================== GRADE TRACKING ========================
function gradeTrackingPage() {
  const isAdmin = isAdminRole();
  const isExecutive = APP.currentRole === 'executive';
  const isClassTeacher = APP.currentRole === 'classTeacher';
  const canEdit = isAdmin || APP.currentRole === 'teacher' || APP.currentRole === 'classTeacher';
  let data = getDataByType('grade_tracking').filter(t => t.subject_name && t.subject_name.trim());
  if (APP.currentRole === 'teacher') data = data.filter(t => subjectHasCoordinator({ coordinator: t.coordinator }, APP.currentUser.name));
  if (ctYear()) data = data.filter(t => norm(t.year_level) === ctYear());

  const selectedYear = APP.filters._gradeTrackingYear || '';
  const allSubjects = getDataByType('subject');
  let subjectsFiltered = selectedYear ? allSubjects.filter(s => norm(s.academic_year) === selectedYear) : allSubjects;
  if (APP.currentRole === 'teacher') subjectsFiltered = subjectsFiltered.filter(s => subjectHasCoordinator(s, APP.currentUser.name));
  if (ctYear()) subjectsFiltered = subjectsFiltered.filter(s => norm(s.year_level) === ctYear());
  { const _df = applyDeptFilter(data, subjectsFiltered, '_' + APP.currentPage + 'Dept'); data = _df.data; subjectsFiltered = _df.subjects; }
  const dataForStats = selectedYear ? data.filter(t => norm(t.academic_year) === selectedYear || !norm(t.academic_year)) : data;

  let statsSection = '';
  let notSubmittedSection = '';
  if (selectedYear) {
    const isTracked = makeTrackingMatcher(dataForStats);
    const notSubmitted = subjectsFiltered.filter(s => !isTracked(s));

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

  const allYears = [...new Set([...allSubjects.map(s => norm(s.academic_year)), ...data.map(t => norm(t.academic_year))].filter(Boolean))].sort();

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
      ${trackingDeptFilterHTML('_gradeTrackingDept', APP.filters._gradeTrackingDept || '')}
    </div>
    ${statsSection}
  </div>
  ${selectedYear ? `${notSubmittedSection}
  ${filterBar({ yearLevel: true, semester: false, year: false })}
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
        <td class="px-4 py-3">${ctCell}${actionTimeText(t, 'coordinator_check')}</td>
        <td class="px-4 py-3">${acCell}${actionTimeText(t, 'academic_check')}</td>
        <td class="px-4 py-3">${dpCell}${actionTimeText(t, 'deputy_sign')}</td>
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
    .map(name => {
      const s = subjects.find(x => x.subject_name === name) || {};
      return `<option value="${name.replace(/"/g, '&quot;')}" data-code="${(s.subject_code || '').replace(/"/g, '&quot;')}">${s.subject_code ? s.subject_code + ' ' : ''}${name}</option>`;
    }).join('');
  const teachers = getDataByType('teacher');
  const teacherList = [...new Set(teachers.map(t => (t.name || '').trim()).filter(Boolean))].sort();
  const myName = (APP.currentUser && APP.currentUser.name || '').trim();
  const coordCheckboxes = teacherList.map(name => {
    const isMe = name === myName;
    return `<label class="flex items-center gap-2 p-1.5 rounded-lg hover:bg-gray-50 cursor-pointer ${isMe ? 'bg-blue-50' : ''}"><input type="checkbox" class="coord-check w-4 h-4 rounded border-gray-300 text-primary focus:ring-primary" value="${name.replace(/"/g, '&quot;')}" ${isMe ? 'checked' : ''}><span class="text-sm">${name}${isMe ? ' <span class="text-xs text-blue-500">(คุณ)</span>' : ''}</span></label>`;
  }).join('');
  showModal('เพิ่มข้อมูลติดตามการส่งเกรดรายวิชา', `
    <form id="addGradeTrackingForm" class="space-y-3">
      <div><label class="block text-xs text-gray-600 mb-1">ชื่อรายวิชา *</label><select name="subject_name" id="gradeTrackingSubjectSelect" required class="w-full border rounded-xl px-3 py-2 text-sm"><option value="">-- เลือกรายวิชา --</option>${subjectOptions}</select></div>
      <input type="hidden" name="subject_code" id="gradeTrackingSubjectCode">
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
      ${trackingBackfillCheckboxHTML()}
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึก</button>
    </form>
  `);
  // Auto-fill subject_code
  const gSel = document.getElementById('gradeTrackingSubjectSelect');
  if (gSel) gSel.addEventListener('change', () => {
    const opt = gSel.options[gSel.selectedIndex];
    const codeEl = document.getElementById('gradeTrackingSubjectCode');
    if (codeEl) codeEl.value = (opt && opt.dataset.code) || '';
  });
  document.getElementById('addGradeTrackingForm').onsubmit = async (e) => {
    e.preventDefault();
    const checked = [...e.target.querySelectorAll('.coord-check:checked')].map(cb => cb.value);
    const extraText = (document.getElementById('gradeTrackingCoordExtra') || {}).value || '';
    const extraNames = extraText.split(',').map(s => s.trim()).filter(Boolean);
    const allNames = [...checked, ...extraNames].filter(Boolean);
    const coordHidden = document.getElementById('gradeTrackingCoordHidden');
    if (coordHidden) coordHidden.value = allNames.join(', ');
    const ss = document.getElementById('gradeTrackingSubjectSelect');
    const cc = document.getElementById('gradeTrackingSubjectCode');
    if (ss && cc && !cc.value) {
      const opt = ss.options[ss.selectedIndex];
      cc.value = (opt && opt.dataset.code) || '';
    }
    await withLoading(e.target, async () => {
      const fd = new FormData(e.target);
      const obj = { type: 'grade_tracking', coordinator_check: 'รอ', academic_check: 'รอ', deputy_sign: 'รอ', approved_date: '', created_at: new Date().toISOString() };
      fd.forEach((v, k) => obj[k] = v);
      applyTrackingBackfill(obj);
      const r = await GSheetDB.create(obj);
      if (r.isOk) { showToast('เพิ่มสำเร็จ'); closeModal() } else showToast('เกิดข้อผิดพลาด', 'error');
    });
  };
}

// ======================== FILE TRACKING ========================
function fileTrackingPage() {
  const isAdmin = isAdminRole();
  const isExecutive = APP.currentRole === 'executive';
  const isClassTeacher = APP.currentRole === 'classTeacher';
  const canEdit = isAdmin || APP.currentRole === 'teacher' || APP.currentRole === 'classTeacher';
  let data = getDataByType('file_tracking').filter(t => t.subject_name && t.subject_name.trim());
  if (APP.currentRole === 'teacher') data = data.filter(t => subjectHasCoordinator({ coordinator: t.coordinator }, APP.currentUser.name));
  if (ctYear()) data = data.filter(t => norm(t.year_level) === ctYear());

  const selectedYear = APP.filters._fileTrackingYear || '';
  const allSubjects = getDataByType('subject');
  let subjectsFiltered = selectedYear ? allSubjects.filter(s => norm(s.academic_year) === selectedYear) : allSubjects;
  if (APP.currentRole === 'teacher') subjectsFiltered = subjectsFiltered.filter(s => subjectHasCoordinator(s, APP.currentUser.name));
  if (ctYear()) subjectsFiltered = subjectsFiltered.filter(s => norm(s.year_level) === ctYear());
  { const _df = applyDeptFilter(data, subjectsFiltered, '_' + APP.currentPage + 'Dept'); data = _df.data; subjectsFiltered = _df.subjects; }
  const dataForStats = selectedYear ? data.filter(t => norm(t.academic_year) === selectedYear || !norm(t.academic_year)) : data;

  let statsSection = '';
  let notSubmittedSection = '';
  if (selectedYear) {
    const isTracked = makeTrackingMatcher(dataForStats);
    const notSubmitted = subjectsFiltered.filter(s => !isTracked(s));

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

  const allYears = [...new Set([...allSubjects.map(s => norm(s.academic_year)), ...data.map(t => norm(t.academic_year))].filter(Boolean))].sort();

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
      ${trackingDeptFilterHTML('_fileTrackingDept', APP.filters._fileTrackingDept || '')}
    </div>
    ${statsSection}
  </div>
  ${selectedYear ? `${notSubmittedSection}
  ${filterBar({ yearLevel: true, semester: false, year: false })}
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
        <td class="px-4 py-3">${ctCell}${actionTimeText(t, 'coordinator_check')}</td>
        <td class="px-4 py-3">${acCell}${actionTimeText(t, 'academic_check')}</td>
        <td class="px-4 py-3">${dpCell}${actionTimeText(t, 'deputy_sign')}</td>
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
    .map(name => {
      const s = subjects.find(x => x.subject_name === name) || {};
      return `<option value="${name.replace(/"/g, '&quot;')}" data-code="${(s.subject_code || '').replace(/"/g, '&quot;')}">${s.subject_code ? s.subject_code + ' ' : ''}${name}</option>`;
    }).join('');
  const teachers = getDataByType('teacher');
  const teacherList = [...new Set(teachers.map(t => (t.name || '').trim()).filter(Boolean))].sort();
  const myName = (APP.currentUser && APP.currentUser.name || '').trim();
  const coordCheckboxes = teacherList.map(name => {
    const isMe = name === myName;
    return `<label class="flex items-center gap-2 p-1.5 rounded-lg hover:bg-gray-50 cursor-pointer ${isMe ? 'bg-blue-50' : ''}"><input type="checkbox" class="coord-check w-4 h-4 rounded border-gray-300 text-primary focus:ring-primary" value="${name.replace(/"/g, '&quot;')}" ${isMe ? 'checked' : ''}><span class="text-sm">${name}${isMe ? ' <span class="text-xs text-blue-500">(คุณ)</span>' : ''}</span></label>`;
  }).join('');
  showModal('เพิ่มข้อมูลติดตามการส่งแฟ้มรายวิชา', `
    <form id="addFileTrackingForm" class="space-y-3">
      <div><label class="block text-xs text-gray-600 mb-1">ชื่อรายวิชา *</label><select name="subject_name" id="fileTrackingSubjectSelect" required class="w-full border rounded-xl px-3 py-2 text-sm"><option value="">-- เลือกรายวิชา --</option>${subjectOptions}</select></div>
      <input type="hidden" name="subject_code" id="fileTrackingSubjectCode">
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
      ${trackingBackfillCheckboxHTML()}
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึก</button>
    </form>
  `);
  // Auto-fill subject_code
  const fSel = document.getElementById('fileTrackingSubjectSelect');
  if (fSel) fSel.addEventListener('change', () => {
    const opt = fSel.options[fSel.selectedIndex];
    const codeEl = document.getElementById('fileTrackingSubjectCode');
    if (codeEl) codeEl.value = (opt && opt.dataset.code) || '';
  });
  document.getElementById('addFileTrackingForm').onsubmit = async (e) => {
    e.preventDefault();
    const checked = [...e.target.querySelectorAll('.coord-check:checked')].map(cb => cb.value);
    const extraText = (document.getElementById('fileTrackingCoordExtra') || {}).value || '';
    const extraNames = extraText.split(',').map(s => s.trim()).filter(Boolean);
    const allNames = [...checked, ...extraNames].filter(Boolean);
    const coordHidden = document.getElementById('fileTrackingCoordHidden');
    if (coordHidden) coordHidden.value = allNames.join(', ');
    const ss = document.getElementById('fileTrackingSubjectSelect');
    const cc = document.getElementById('fileTrackingSubjectCode');
    if (ss && cc && !cc.value) {
      const opt = ss.options[ss.selectedIndex];
      cc.value = (opt && opt.dataset.code) || '';
    }
    await withLoading(e.target, async () => {
      const fd = new FormData(e.target);
      const obj = { type: 'file_tracking', coordinator_check: 'รอ', academic_check: 'รอ', deputy_sign: 'รอ', approved_date: '', created_at: new Date().toISOString() };
      fd.forEach((v, k) => obj[k] = v);
      applyTrackingBackfill(obj);
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
  const isAdmin = isAdminRole();
  const isTeacher = APP.currentRole === 'teacher';
  const isClassTeacher = APP.currentRole === 'classTeacher';
  const isExecutive = APP.currentRole === 'executive';
  const canApprove = isTeacher || isClassTeacher || isExecutive;
  const canEdit = isAdmin || isTeacher || isClassTeacher;

  let data = getDataByType('leave');

  // ไม่แสดงข้อมูลการลาของนักศึกษาที่ลาออก หรือพักการศึกษา
  const _leaveExcludedNames = new Set(getDataByType('student')
    .filter(s => { const st = norm(s.status); return st === 'ลาออก' || st === 'พักการศึกษา'; })
    .map(s => norm(s.name)).filter(Boolean));
  if (_leaveExcludedNames.size) data = data.filter(l => !_leaveExcludedNames.has(norm(l.name)));

  if (isStudent && APP.currentUser.data) data = data.filter(l => l.name === APP.currentUser.data.name);
  if (isTeacher) {
    const myName = (APP.currentUser.name || '').trim();
    data = data.filter(l => {
      // 1) ใช้ coordinator ที่บันทึกไว้ในใบลาก่อน (ตรงตามช่วงเวลาที่นักศึกษาส่งลา)
      let coordStr = (l.coordinator || '').trim();
      // 2) fallback: ค้นจากตาราง subject — match ทั้ง subject_name + semester + academic_year
      //    เพื่อกันกรณีรายวิชาเดียวกันแต่คนละปี/คนละ batch มีผู้ประสานคนละคน
      if (!coordStr) {
        const sub = getDataByType('subject').find(s =>
          s.subject_name === l.subject_name &&
          normSem(s.semester) === normSem(l.semester) &&
          norm(s.academic_year) === norm(l.academic_year)
        );
        if (sub && sub.coordinator) coordStr = sub.coordinator;
      }
      // 3) สุดท้าย fallback หา subject ใดก็ได้ที่ชื่อตรง (กันใบลาเก่าที่ไม่มี coordinator)
      if (!coordStr) {
        const sub = getDataByType('subject').find(s => s.subject_name === l.subject_name);
        if (sub && sub.coordinator) coordStr = sub.coordinator;
      }
      if (!coordStr) return false;
      // Support multiple coordinators separated by comma / และ / and
      const coords = String(coordStr).split(/[,;|]|\sและ\s|\sand\s/).map(c => c.trim()).filter(Boolean);
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

  // Admin/Academic/Executive/ClassTeacher/Teacher: ใช้ student picker
  // - admin/academic/executive: เลือก year ได้อิสระ + เลือกนักศึกษา
  // - classTeacher: บังคับใช้ year ที่ตัวเองรับผิดชอบ (เลือกได้แค่นักศึกษาในชั้นปีนั้น)
  // - teacher: เลือกนักศึกษาจากที่มีใบลาในรายวิชาที่ตนเองสอน (ไม่ต้องเลือก year)
  const showStudentPicker = isAdmin || isExecutive || isClassTeacher || isTeacher;
  let leaveYearLevel = '';
  let leaveStudentName = '';
  if (showStudentPicker) {
    if (isClassTeacher) {
      // อาจารย์ประจำชั้น: ล็อก year ไว้ที่ชั้นปีที่รับผิดชอบ
      leaveYearLevel = String(APP.currentUser.responsible_year || '1');
    } else if (!isTeacher) {
      leaveYearLevel = APP.filters._leaveYearLevel || '';
    }
    leaveStudentName = APP.filters._leaveStudent || '';
    // ใช้ year filter เฉพาะ admin/executive (classTeacher/teacher ถูกกรองตั้งแต่ขั้นต้นแล้ว)
    if (leaveYearLevel && !isClassTeacher && !isTeacher) {
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
  // ถ้าเป็น role ที่ต้อง "เลือกนักศึกษา" ก่อน แต่ยังไม่ได้เลือก → ซ่อนข้อมูล
  // - admin / academic / executive / classTeacher / teacher → ต้องเลือกนักศึกษา
  //   (teacher จะเห็นเฉพาะนักศึกษาที่ลาในวิชาของตนเองใน dropdown)
  // - student → ไม่เกี่ยว (เห็นแค่ของตัวเอง)
  const requireStudentSelection = showStudentPicker && !leaveStudentName;
  if (requireStudentSelection) {
    data = []; // ไม่แสดงข้อมูลใบลาในตาราง / summary จนกว่าจะเลือกนักศึกษา
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
    const stuBatch = APP.currentUser.data?.batch || '';
    const allSubjects = getDataByType('subject');
    // กรองรายวิชาตามนักศึกษา (batch + year_level)
    //   - ถ้ารายวิชาระบุ batch → batch ต้องตรงกับนักศึกษา
    //   - ถ้ารายวิชาไม่ระบุ batch (เช่นวิชา GE ทั่วไป) → ใช้ year_level ตรงเป็นตัวจับคู่
    let subjects = allSubjects;
    if (norm(stuBatch)) {
      subjects = subjects.filter(s => {
        const sBatch = norm(s.batch);
        if (sBatch) return sBatch === norm(stuBatch);
        return norm(s.year_level) === norm(stuYearLevel);
      });
    } else if (stuYearLevel) {
      subjects = subjects.filter(s => norm(s.year_level) === norm(stuYearLevel));
    }
    // ขจัดรายวิชาซ้ำ: ถ้ามี subject_code+ภาค+ปี ซ้ำกัน → เลือกเฉพาะที่ year_level ตรงกับนักศึกษา
    if (stuYearLevel) {
      const keyCounts = {};
      subjects.forEach(s => {
        const k = `${norm(s.subject_code)}|${normSem(s.semester)}|${norm(s.academic_year)}`;
        keyCounts[k] = (keyCounts[k] || 0) + 1;
      });
      subjects = subjects.filter(s => {
        const k = `${norm(s.subject_code)}|${normSem(s.semester)}|${norm(s.academic_year)}`;
        if ((keyCounts[k] || 0) > 1) return norm(s.year_level) === norm(stuYearLevel);
        return true;
      });
    }
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

  // Admin/Academic/Executive/ClassTeacher/Teacher: year-level + student filter card + pie chart
  let adminFilterCard = '';
  let pieChartCard = '';
  if (showStudentPicker) {
    // Build student dropdown
    //   - admin/executive: นักศึกษาทั้งหมด (กรองตาม year ที่เลือก ถ้ามี)
    //   - classTeacher: นักศึกษาในชั้นปีที่ตัวเองรับผิดชอบ
    //   - teacher: นักศึกษาที่มีใบลาในรายวิชาที่ตนเองสอน (ดึงจาก leave records)
    let studentNamesForDropdown = [];
    if (isTeacher) {
      // หาจาก leave records ของวิชาที่ teacher คนนี้ดูแล
      const myName = (APP.currentUser.name || '').trim();
      const allLeavesForTeacher = getDataByType('leave').filter(l => {
        if (l.leave_status === 'ปฏิเสธ') return false;
        let coordStr = (l.coordinator || '').trim();
        if (!coordStr) {
          const sub = getDataByType('subject').find(s => s.subject_name === l.subject_name);
          if (sub) coordStr = sub.coordinator || '';
        }
        if (!coordStr) return false;
        const coords = String(coordStr).split(/[,;|]|\sและ\s|\sand\s/).map(c => c.trim()).filter(Boolean);
        return coords.some(c => c === myName || c.includes(myName) || myName.includes(c));
      });
      studentNamesForDropdown = [...new Set(allLeavesForTeacher.map(l => (l.name || '').trim()).filter(Boolean))].sort();
    } else {
      const allStudents = getDataByType('student');
      const studentsForDropdown = leaveYearLevel
        ? allStudents.filter(s => norm(s.year_level) === norm(leaveYearLevel))
        : allStudents;
      studentNamesForDropdown = studentsForDropdown.map(s => (s.name || '').trim()).filter(Boolean).sort();
    }
    const studentOptions = studentNamesForDropdown
      .map(name => `<option value="${name.replace(/"/g, '&quot;')}" ${leaveStudentName === name ? 'selected' : ''}>${name}</option>`)
      .join('');

    // Year selector — locked for classTeacher, hidden for teacher, dropdown for admin/executive
    let yearSelectorHTML = '';
    if (isClassTeacher) {
      yearSelectorHTML = `<div>
          <label class="block text-xs text-gray-600 mb-1">ชั้นปี</label>
          <div class="border border-amber-200 bg-amber-50 rounded-xl px-3 py-2 text-sm text-amber-800 font-medium flex items-center gap-2">
            <i data-lucide="lock" class="w-3 h-3"></i>
            ชั้นปีที่ ${leaveYearLevel}
            <span class="text-[10px] text-amber-600 font-normal">(ที่คุณรับผิดชอบ)</span>
          </div>
        </div>`;
    } else if (!isTeacher) {
      yearSelectorHTML = `<div>
          <label class="block text-xs text-gray-600 mb-1">ชั้นปี</label>
          <select onchange="setLeaveYearLevel(this.value)" class="border border-gray-200 rounded-xl px-3 py-2 text-sm">
            <option value="">ทุกชั้นปี</option>
            <option value="1" ${leaveYearLevel === '1' ? 'selected' : ''}>ชั้นปี 1</option>
            <option value="2" ${leaveYearLevel === '2' ? 'selected' : ''}>ชั้นปี 2</option>
            <option value="3" ${leaveYearLevel === '3' ? 'selected' : ''}>ชั้นปี 3</option>
            <option value="4" ${leaveYearLevel === '4' ? 'selected' : ''}>ชั้นปี 4</option>
          </select>
        </div>`;
    }

    const showClearBtn = isTeacher ? !!leaveStudentName : (isClassTeacher ? !!leaveStudentName : (leaveYearLevel || leaveStudentName));
    const dropdownPlaceholder = isTeacher
      ? '-- กรุณาเลือกนักศึกษา (เฉพาะคนที่ลาในวิชาของคุณ) --'
      : (leaveYearLevel ? '-- เลือกนักศึกษาในชั้นปี ' + leaveYearLevel + ' --' : '-- กรุณาเลือกนักศึกษา --');
    // Label: บังคับเลือกนักศึกษาทุก role (มีดอกจันแดง + ข้อความบังคับ)
    const studentLabelHTML = `นักศึกษา <span class="text-red-500">*</span> <span class="text-[10px] text-gray-400">(ต้องเลือกเพื่อดูข้อมูล)</span>`;

    adminFilterCard = `<div class="bg-white rounded-2xl p-4 border border-blue-100 mb-4">
      <div class="flex flex-wrap items-end gap-3">
        ${yearSelectorHTML}
        <div class="flex-1 min-w-[200px]">
          <label class="block text-xs text-gray-600 mb-1">${studentLabelHTML}</label>
          <select onchange="setLeaveStudent(this.value)" class="w-full border border-gray-200 rounded-xl px-3 py-2 text-sm">
            <option value="">${dropdownPlaceholder}</option>
            ${studentOptions}
          </select>
        </div>
        ${showClearBtn ? `<button onclick="clearLeaveFilters()" class="px-3 py-2 text-xs text-gray-600 hover:bg-gray-100 rounded-xl border border-gray-200"><i data-lucide="x" class="w-3 h-3 inline"></i> ล้างตัวกรอง</button>` : ''}
      </div>
    </div>`;
    if (leaveStudentName) {
      pieChartCard = renderLeavePieChart(data, leaveStudentName);
    }
  }

  // ถ้า user เป็น role ที่ต้องเลือกนักศึกษา แต่ยังไม่ได้เลือก → แสดง empty state แทนตาราง+summary
  const emptyStateMsg = requireStudentSelection
    ? `<div class="bg-white rounded-2xl border border-blue-100 p-8 text-center text-gray-400">
        <i data-lucide="user-search" class="w-12 h-12 mx-auto mb-3 text-gray-300"></i>
        <p class="text-sm">กรุณาเลือกนักศึกษาจากตัวกรองด้านบน เพื่อดูข้อมูลการลา</p>
      </div>`
    : '';

  return `<h2 class="text-xl font-bold text-gray-800 mb-4"><i data-lucide="calendar-off" class="w-6 h-6 inline mr-2"></i>ระบบการลาของนักศึกษา</h2>
  ${isAdmin ? `<div class="flex gap-2 mb-4"><button onclick="showAddLeaveModal()" class="flex items-center gap-2 px-4 py-2 bg-primary text-white rounded-xl hover:bg-primaryDark text-sm"><i data-lucide="plus" class="w-4 h-4"></i>เพิ่มข้อมูลการลา</button>${csvUploadBtn('leave', 'name,subject_name,leave_hours,leave_percent,semester,academic_year,leave_date,leave_type')}</div>` : ''}
  ${pendingBanner}
  ${form}
  ${adminFilterCard}
  ${requireStudentSelection ? emptyStateMsg : `${pieChartCard}
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

    // หมายเหตุจากแต่ละขั้นการอนุมัติ (แสดงใต้ workflow steps)
    const notesList = [];
    if (l.coordinator_note) notesList.push(`<div class="text-[11px] text-gray-700 bg-gray-50 border-l-2 border-amber-300 px-2 py-1 rounded mt-1" title="หมายเหตุจากอาจารย์ผู้ประสาน"><strong class="text-amber-700">ปสน.:</strong> ${l.coordinator_note}</div>`);
    if (l.class_teacher_note) notesList.push(`<div class="text-[11px] text-gray-700 bg-gray-50 border-l-2 border-blue-300 px-2 py-1 rounded mt-1" title="หมายเหตุจากอาจารย์ประจำชั้น"><strong class="text-blue-700">ปจช.:</strong> ${l.class_teacher_note}</div>`);
    if (l.deputy_note) notesList.push(`<div class="text-[11px] text-gray-700 bg-gray-50 border-l-2 border-purple-300 px-2 py-1 rounded mt-1" title="หมายเหตุจากผู้บริหาร"><strong class="text-purple-700">ผบห.:</strong> ${l.deputy_note}</div>`);
    const notesHTML = notesList.length ? `<div class="mt-1.5">${notesList.join('')}</div>` : '';

    const approvalField = isTeacher ? 'coordinator_approval' : isClassTeacher ? 'class_teacher_approval' : isExecutive ? 'deputy_approval' : '';
    const myApprovalStatus = approvalField ? (l[approvalField] || 'รอ') : '';

    // Sequential rule: classTeacher waits for ALL coordinators (all subjects same student+date); deputy waits for classTeacher
    const allCoordsOk = allCoordinatorsApproved(l);
    const cannotActYet = (isClassTeacher && !allCoordsOk) || (isExecutive && (!coordApproved || !classApproved));
    const waitingForLabel = isClassTeacher ? (allCoordsOk ? 'รออาจารย์ผู้ประสานอนุมัติก่อน' : 'รอ coordinator วิชาอื่นอนุมัติครบก่อน') : 'รออาจารย์ประจำชั้นอนุมัติก่อน';

    const renderApprovalCell = () => {
      // Student/admin view: show only the workflow + current stage + notes
      if (!canApprove) {
        return `${workflowSteps}${currentStage || '<span class="px-2 py-1 rounded-full text-xs bg-gray-100 text-gray-700">รอส่ง</span>'}${notesHTML}`;
      }
      if (isRejected && myApprovalStatus !== 'ปฏิเสธ' && myApprovalStatus !== 'อนุมัติ') {
        return `${workflowSteps}<span class="px-2 py-1 rounded-full text-xs bg-red-100 text-red-700">ใบลาถูกปฏิเสธแล้ว</span>${notesHTML}`;
      }
      if (myApprovalStatus === 'อนุมัติ') return `${workflowSteps}<span class="px-2 py-1 rounded-full text-xs bg-green-100 text-green-700"><i data-lucide="check" class="w-3 h-3 inline"></i> คุณอนุมัติแล้ว</span>${notesHTML}`;
      if (myApprovalStatus === 'ปฏิเสธ') return `${workflowSteps}<span class="px-2 py-1 rounded-full text-xs bg-red-100 text-red-700"><i data-lucide="x" class="w-3 h-3 inline"></i> คุณปฏิเสธแล้ว</span>${notesHTML}`;
      if (cannotActYet) {
        return `${workflowSteps}<span class="px-2 py-1 rounded-full text-xs bg-amber-50 text-amber-700 border border-amber-200">⏳ ${waitingForLabel}</span>${notesHTML}`;
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
        </div>${notesHTML}`;
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
  ${paginationHTML(total, APP.pagination.perPage, APP.pagination.page, 'changePage')}`}`;
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
  ${paginationHTML(total, APP.pagination.perPage, APP.pagination.page, 'changePage')}`;
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
// ======================== คู่มือการใช้งาน ========================
function userGuidePage() {
  const sec = (icon, title, body) => `<details class="bg-white rounded-2xl border border-blue-100 mb-3 group" open>
    <summary class="cursor-pointer px-5 py-4 font-bold text-gray-800 flex items-center gap-2"><i data-lucide="${icon}" class="w-5 h-5 text-primary"></i>${title}</summary>
    <div class="px-5 pb-5 text-sm text-gray-700 leading-relaxed space-y-2">${body}</div>
  </details>`;
  const li = items => `<ul class="list-disc pl-5 space-y-1">${items.map(i => `<li>${i}</li>`).join('')}</ul>`;

  return `<div class="max-w-4xl">
  <div class="flex items-center gap-3 mb-2"><i data-lucide="help-circle" class="w-7 h-7 text-primary"></i><h2 class="text-2xl font-bold text-gray-800">คู่มือการใช้งานระบบ EMS-BCNB</h2></div>
  <p class="text-sm text-gray-500 mb-5">ระบบบริหารจัดการงานวิชาการ · วิทยาลัยพยาบาลบรมราชชนนี กรุงเทพ</p>

  ${sec('info', '1. ภาพรวมระบบ', `
    <p>EMS-BCNB เป็นระบบบริหารจัดการงานวิชาการที่ทำงานผ่านเว็บเบราว์เซอร์ โดยใช้ <b>Google Sheet</b> เป็นฐานข้อมูล (อ่านข้อมูลแบบสาธารณะ และเขียน/แก้ไขผ่าน Apps Script) ข้อมูลทั้งหมดจึงอัปเดตแบบเรียลไทม์ร่วมกัน</p>
    <p>เมนูหลักแบ่งเป็นกลุ่ม: <b>หน้าหลัก</b>, <b>ระบบทะเบียน</b>, <b>ผลการศึกษา</b>, <b>ทำเนียบอาจารย์</b>, <b>ติดตามการส่ง</b>, <b>ระบบการลาของนักศึกษา</b>, <b>บริการอื่นๆ</b>, <b>ตั้งค่าระบบ</b> และ <b>บันทึกการเข้าใช้ระบบ</b> (เมนูที่เห็นขึ้นอยู่กับบทบาทผู้ใช้)</p>
    <p>มุมขวาบนมีปุ่ม <b>รีเฟรชข้อมูล</b> (โหลดข้อมูลล่าสุดจาก Sheet), <b>แจ้งเตือน</b> (กระดิ่ง) และ <b>ออกจากระบบ</b></p>`)}

  ${sec('log-in', '2. การเข้าสู่ระบบ (6 บทบาท)', `
    <p>หน้าเข้าสู่ระบบให้เลือกบทบาทก่อน แล้วกรอกข้อมูลตามแต่ละบทบาท:</p>
    ${li([
    '<b>ผู้ดูแลระบบ (Admin):</b> รหัสผ่าน 6 หลัก (ตัวเลข)',
    '<b>เจ้าหน้าที่งานวิชาการ:</b> Email + รหัสผ่าน',
    '<b>ผู้บริหาร:</b> Username + รหัสผ่าน',
    '<b>อาจารย์:</b> Email + รหัสผ่าน',
    '<b>อาจารย์ประจำชั้น:</b> Username + รหัสผ่าน',
    '<b>นักศึกษา:</b> เลขบัตรประชาชน 13 หลัก',
  ])}
    <p class="text-amber-700 bg-amber-50 border border-amber-200 rounded-lg p-2"><b>หมายเหตุ:</b> นักศึกษาที่ <b>สำเร็จการศึกษา</b> แล้วจะไม่สามารถเข้าสู่ระบบได้อีก</p>`)}

  ${sec('shield', '3. บทบาทและสิทธิ์การเข้าถึง', `
    <p>แต่ละบทบาทเห็นเมนูและทำสิ่งต่างๆ ได้ไม่เท่ากัน (ผู้ดูแลปรับสิทธิ์ได้ที่ "ตั้งค่าระบบ"):</p>
    ${li([
    '<b>ผู้ดูแลระบบ / งานวิชาการ:</b> เข้าถึงและแก้ไขได้เกือบทุกส่วน รวมถึงเพิ่ม/แก้/ลบข้อมูลและจัดการผู้ใช้',
    '<b>ผู้บริหาร:</b> ดูข้อมูลภาพรวมเป็นหลัก (อ่านอย่างเดียวในหลายส่วน) และอนุมัติใบลา',
    '<b>อาจารย์:</b> ดูนักศึกษาในที่ปรึกษา ผลการเรียน/ENG ติดตามการส่งงานรายวิชาที่รับผิดชอบ และอนุมัติใบลา',
    '<b>อาจารย์ประจำชั้น:</b> ดูแลนักศึกษาในชั้นปีที่รับผิดชอบ (แยกห้อง A/B) และอนุมัติใบลา',
    '<b>นักศึกษา:</b> ดูข้อมูลตนเอง ผลการเรียน ผลสอบภาษาอังกฤษ และส่งใบลา',
  ])}`)}

  ${sec('book-open', '4. ระบบทะเบียน', `
    <p><b>ข้อมูลนักศึกษา</b> — เลือกชั้นปี (1-4) หรือ "ผู้สำเร็จการศึกษา" เพื่อดูรายชื่อ; ผู้ดูแลเพิ่ม/แก้ไข/ลบ และนำเข้าด้วย CSV ได้ มีปุ่ม <b>เลื่อนชั้นปี</b> (ดูข้อ 5) สถานภาพมี กำลังศึกษา / พักการศึกษา / ลาออก / สำเร็จการศึกษา (ระบบจะนับเฉพาะผู้ที่กำลังศึกษาในหน้าหลัก)</p>
    <p><b>ข้อมูลอาจารย์</b> — เลือกสาขาวิชาเพื่อดูรายชื่อ ผู้ดูแลเพิ่ม/แก้ไข/ลบได้</p>
    <p><b>ข้อมูลอาจารย์พิเศษ</b> — บันทึกอาจารย์พิเศษ (ปีการศึกษา, คำนำหน้า, ชื่อ, ตำแหน่ง, หน่วยงาน, ระดับวุฒิ) กรองตามปีการศึกษาได้ และทำเนียบอาจารย์สามารถดึงข้อมูลจากที่นี่ไปใช้ได้</p>
    <p><b>ข้อมูลศิษย์เก่า</b> — รายชื่อผู้สำเร็จการศึกษา (คำนำหน้า, ชื่อ, รุ่น, สถานภาพ, สถานที่ปฏิบัติงาน, วันเข้า/จบการศึกษา, วันที่บันทึก) กรองตามรุ่นได้ ระบบจะเพิ่มให้อัตโนมัติเมื่อเลื่อนชั้นปี 4 เป็นสำเร็จการศึกษา</p>
    <p><b>ปฏิทินกิจกรรมวิชาการ</b> — รายการกิจกรรม/กำหนดการของวิทยาลัย</p>
    <p><b>รายวิชาที่เปิดสอน</b> — เลือกปีการศึกษาเพื่อดูรายวิชา แสดงรหัสหน่วยกิตแบบ น(ท-ป-อ) (ดูข้อ 6) ผู้ดูแลเพิ่ม/แก้ไข/นำเข้า CSV และ "นำเข้ารายชื่อสร้างผลการเรียน" ได้</p>`)}

  ${sec('graduation-cap', '5. การเลื่อนชั้นปี และผู้สำเร็จการศึกษา', `
    <p>ในหน้า <b>ข้อมูลนักศึกษา</b> (ผู้ดูแล) มีแผง "เลื่อนชั้นปี" แยกปุ่มต่อชั้นปี: ปี 1→2, 2→3, 3→4 และ <b>ปี 4 → สำเร็จการศึกษา</b></p>
    ${li([
    'เลื่อนเฉพาะผู้ที่ "กำลังศึกษา" (ข้ามผู้ที่พัก/ลาออก/จบแล้ว)',
    'ปี 4 จะเปลี่ยนสถานะเป็น "สำเร็จการศึกษา" และชั้นปีเป็น "จบ" พร้อมบันทึกเข้า "ข้อมูลศิษย์เก่า" อัตโนมัติ',
    'ผลการเรียน/ผลสอบเดิมไม่ได้รับผลกระทบ (ผูกด้วยรหัสนักศึกษา)',
    'มีหน้ายืนยันก่อนทุกครั้ง — แนะนำให้สำรอง (คัดลอก) แท็บ student ก่อน เพราะย้อนกลับอัตโนมัติไม่ได้',
  ])}`)}

  ${sec('file-text', '6. ผลการศึกษา และใบ Transcript', `
    <p><b>ผลการเรียน</b> — ดูเกรดรายวิชา GPAX และพิมพ์ใบแสดงผลการเรียน; ผู้ดูแลเพิ่มเกรด/นำเข้า CSV ได้</p>
    <p><b>ผลสอบภาษาอังกฤษ</b> — บันทึก/ดูผลสอบและสถานะผ่าน-ไม่ผ่าน กรองตามปีการศึกษา/ชั้นปี/อาจารย์ที่ปรึกษาได้</p>
    <p><b>ใบระเบียนแสดงผลการเรียน (Transcript):</b> สำหรับ <b>ผู้สำเร็จการศึกษา</b> เมื่อกด "ใบแสดงผลการเรียน" จะได้รูปแบบทางการ (โลโก้, ข้อมูลส่วนตัว, รายวิชาแยกชั้นปี, สรุปหน่วยกิต/GPAX, ชั่วโมงฝึกปฏิบัติ, ผลสอบภาษาอังกฤษ/สอบรวบยอด, ลายเซ็นนายทะเบียน) และดาวน์โหลด PDF ได้ ข้อมูลส่วนตัวสำหรับ Transcript กรอกได้ในฟอร์มนักศึกษา (กล่อง "ข้อมูลสำหรับใบ Transcript")</p>`)}

  ${sec('hash', '7. ความหมายรหัสหน่วยกิต น(ท-ป-อ)', `
    ${li([
    '<b>ตัวเลขหน้าวงเล็บ</b> = จำนวนหน่วยกิตรวม',
    '<b>ตัวแรกในวงเล็บ</b> = ชั่วโมงทฤษฎี/สัปดาห์',
    '<b>ตัวที่สองในวงเล็บ</b> = ชั่วโมงปฏิบัติ/ทดลอง/ฝึกในคลินิกหรือชุมชน/สัปดาห์',
    '<b>ตัวที่สามในวงเล็บ</b> = ชั่วโมงศึกษาด้วยตนเอง/สัปดาห์',
  ])}
    <p>ตัวอย่าง <span class="font-mono font-bold">2(1-2-3)</span> = 2 หน่วยกิต · ทฤษฎี 1 · ทดลอง 2 · ศึกษาด้วยตนเอง 3 ชม./สัปดาห์</p>`)}

  ${sec('clipboard-list', '8. ติดตามการส่ง และระบบการลา', `
    <p><b>ติดตามการส่ง</b> มี 4 ประเภท: ส่งรายละเอียดรายวิชา, ส่งผลการดำเนินงานรายวิชา, ส่งเกรดรายวิชา และส่งแฟ้มรายวิชา — แสดงสถานะส่ง/ยังไม่ส่งของแต่ละรายวิชา (กรองตามปีการศึกษาได้)</p>
    <p><b>ระบบการลาของนักศึกษา</b> — นักศึกษาส่งใบลา (เลือกรายวิชา/ชั่วโมง/ประเภท) อาจารย์ผู้ประสานงานรายวิชา/อาจารย์ประจำชั้น/ผู้บริหารพิจารณาอนุมัติ มีสถานะ รออนุมัติ / อนุมัติแล้ว / ปฏิเสธ และนับชั่วโมงการลาให้อัตโนมัติ</p>`)}

  ${sec('grid', '9. บริการอื่นๆ และตั้งค่าระบบ', `
    <p><b>บริการอื่นๆ</b> — ข่าวสาร/ประกาศ และคำร้องขอเอกสาร (อัปเดตสถานะ รอ/กำลังดำเนินการ/เสร็จสิ้น)</p>
    <p><b>ทำเนียบอาจารย์</b> — ฐานข้อมูลอาจารย์แบบละเอียดแยกประเภท (ประจำหลักสูตร/ประจำ/รับผิดชอบหลักสูตร/ลาศึกษาต่อ/พิเศษ) พร้อมสรุปจำนวน</p>
    <p><b>ตั้งค่าระบบ</b> (ผู้ดูแล) — จัดการผู้ใช้งาน (เพิ่ม/แก้/ลบ) และตารางกำหนด <b>สิทธิ์การเข้าถึง</b> ของแต่ละบทบาทต่อแต่ละโมดูล รวมถึงตั้งค่าการเชื่อมต่อ Google Sheet</p>
    <p><b>บันทึกการเข้าใช้ระบบ</b> (ผู้ดูแล) — ประวัติการเข้าใช้งาน</p>`)}

  ${sec('alert-triangle', '10. การแก้ปัญหาเบื้องต้น', `
    ${li([
    '<b>ข้อมูล/เมนูใหม่ไม่ขึ้น:</b> กด Ctrl+F5 เพื่อล้างแคชและโหลดใหม่',
    '<b>กดเพิ่ม/แก้ไขแล้วบันทึกไม่ได้:</b> ตรวจว่าตั้งค่า Apps Script URL (โหมดเขียน) แล้ว — ถ้าเป็นโหมดอ่านอย่างเดียวจะแก้ข้อมูลไม่ได้',
    '<b>ดาวน์โหลด PDF ไม่ขึ้น:</b> อนุญาต Popup ของเบราว์เซอร์',
    '<b>ข้อมูลไม่อัปเดต:</b> กดปุ่มรีเฟรช (มุมขวาบน) หรือโหลดหน้าใหม่',
  ])}`)}

  ${sec('phone', '11. ติดต่อสอบถาม', `
    <p>โทร. 02 354 2320 · งานวิชาการ (admin) ต่อ 310 · งานบริหารหลักสูตร ต่อ 311 · งานข้อสอบและวัดประเมินผล ต่อ 320 · งานดิจิทัล ต่อ 322 · งานเลขา ต่อ 330 · งานทะเบียน ต่อ 340</p>`)}

  <p class="text-xs text-gray-400 mt-4">จัดทำโดยงานดิจิทัล · ระบบ EMS-BCNB</p>
  </div>`;
}

// ======================== PASSWORD OTP (ลืม/เปลี่ยนรหัสผ่านผ่านอีเมล) ========================
// mode: 'forgot' (หน้า login) | 'change' (ผู้ใช้ที่ล็อกอินอยู่)
function showPasswordOtpModal(mode) {
  const cu = APP.currentUser || {};
  const prefEmail = (mode === 'change') ? (cu.email || (cu.data && cu.data.email) || '') : '';
  const lockEmail = (mode === 'change' && prefEmail) ? 'readonly' : '';
  const title = (mode === 'change') ? 'เปลี่ยนรหัสผ่าน (ยืนยันด้วย OTP ทางอีเมล)' : 'ลืมรหัสผ่าน / ตั้งรหัสผ่านใหม่';
  showModal(title, `
    <div class="space-y-3 text-sm">
      <p class="text-gray-600">ระบบจะส่งรหัส OTP 6 หลักไปยังอีเมลของคุณ เพื่อยืนยันตัวตนก่อนตั้งรหัสผ่านใหม่</p>
      <div>
        <label class="block text-xs text-gray-600 mb-1">อีเมลที่ลงทะเบียนไว้</label>
        <input id="pwOtpEmail" type="email" value="${prefEmail}" ${lockEmail} class="w-full border rounded-xl px-3 py-2" placeholder="name@bcn.ac.th">
      </div>
      <button type="button" id="pwOtpSendBtn" data-no-loading onclick="pwOtpSend('${mode}')" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">ขอรหัส OTP</button>
      <div id="pwOtpStep2" class="hidden space-y-3 pt-3 border-t">
        <div>
          <label class="block text-xs text-gray-600 mb-1">รหัส OTP (6 หลักจากอีเมล)</label>
          <input id="pwOtpCode" maxlength="6" inputmode="numeric" class="w-full border rounded-xl px-3 py-2 text-center tracking-[0.5em] text-lg" placeholder="______">
        </div>
        <div>
          <label class="block text-xs text-gray-600 mb-1">รหัสผ่านใหม่ (อย่างน้อย 6 ตัวอักษร)</label>
          <input id="pwOtpNew" type="password" class="w-full border rounded-xl px-3 py-2">
        </div>
        <div>
          <label class="block text-xs text-gray-600 mb-1">ยืนยันรหัสผ่านใหม่</label>
          <input id="pwOtpNew2" type="password" class="w-full border rounded-xl px-3 py-2">
        </div>
        <button type="button" onclick="pwOtpConfirm('${mode}')" class="w-full bg-emerald-600 text-white py-2.5 rounded-xl hover:bg-emerald-700">ตั้งรหัสผ่านใหม่</button>
      </div>
      <div id="pwOtpMsg" class="text-sm hidden"></div>
    </div>
  `);
}

function _pwOtpMsg(text, ok) {
  const el = document.getElementById('pwOtpMsg');
  if (!el) return;
  el.textContent = text;
  el.className = 'text-sm ' + (ok ? 'text-emerald-600' : 'text-red-500');
  el.classList.remove('hidden');
}

async function pwOtpSend(mode) {
  const email = (document.getElementById('pwOtpEmail').value || '').trim();
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) { _pwOtpMsg('กรุณากรอกอีเมลให้ถูกต้อง', false); return; }
  const btn = document.getElementById('pwOtpSendBtn');
  btn.disabled = true; btn.textContent = 'กำลังส่ง...';
  const r = await GSheetDB.requestPasswordOtp(email);
  btn.disabled = false; btn.textContent = 'ขอรหัส OTP อีกครั้ง';
  if (r && r.isOk) {
    document.getElementById('pwOtpStep2').classList.remove('hidden');
    _pwOtpMsg('ส่งรหัส OTP ไปที่อีเมลแล้ว (หากอีเมลมีอยู่ในระบบ) — โปรดตรวจสอบกล่องจดหมาย/สแปม', true);
  } else {
    _pwOtpMsg((r && r.error) || 'ส่งรหัสไม่สำเร็จ', false);
  }
}

async function pwOtpConfirm(mode) {
  const email = (document.getElementById('pwOtpEmail').value || '').trim();
  const code = (document.getElementById('pwOtpCode').value || '').trim();
  const np = document.getElementById('pwOtpNew').value || '';
  const np2 = document.getElementById('pwOtpNew2').value || '';
  if (!/^\d{6}$/.test(code)) { _pwOtpMsg('กรุณากรอกรหัส OTP 6 หลัก', false); return; }
  if (np.length < 6) { _pwOtpMsg('รหัสผ่านใหม่ต้องยาวอย่างน้อย 6 ตัวอักษร', false); return; }
  if (np !== np2) { _pwOtpMsg('รหัสผ่านใหม่ทั้งสองช่องไม่ตรงกัน', false); return; }
  const r = await GSheetDB.resetPasswordOtp({ email, code, newPassword: np, source: (mode === 'change' ? 'change' : 'forgot') });
  if (r && r.isOk) {
    _pwOtpMsg('ตั้งรหัสผ่านใหม่สำเร็จ! กรุณาเข้าสู่ระบบด้วยรหัสใหม่', true);
    showToast('ตั้งรหัสผ่านใหม่สำเร็จ');
    setTimeout(closeModal, 1400);
    if (mode === 'change') setTimeout(handleLogout, 1500);
  } else {
    _pwOtpMsg((r && r.error) || 'ตั้งรหัสผ่านไม่สำเร็จ', false);
  }
}

// แท็บในหน้าตั้งค่า (ผู้ดูแลระบบ): จัดการผู้ใช้ / บันทึกการเปลี่ยนรหัสผ่าน
function changeSettingsTab(t) { APP._settingsTab = t; renderCurrentPage(); }

function passwordLogSection() {
  const roleLabels = { admin: 'ผู้ดูแลระบบ', academic: 'เจ้าหน้าที่งานวิชาการ', executive: 'ผู้บริหาร', teacher: 'อาจารย์', classTeacher: 'อาจารย์ประจำชั้น', deptHead: 'ประธานสาขาวิชา', registrar: 'เจ้าหน้าที่งานทะเบียน', student: 'นักศึกษา' };
  const actionLabels = { forgot: 'ลืมรหัสผ่าน (รีเซ็ตผ่านอีเมล)', reset: 'ลืมรหัสผ่าน (รีเซ็ตผ่านอีเมล)', change: 'เปลี่ยนรหัสผ่านในระบบ' };
  let logs = getDataByType('password_log').slice();
  logs.sort((a, b) => String(b.created_at || b.timestamp || '').localeCompare(String(a.created_at || a.timestamp || '')));
  const rows = logs.map(l => `<tr class="border-t hover:bg-gray-50">
    <td class="px-4 py-3 text-sm whitespace-nowrap">${l.timestamp || ''}</td>
    <td class="px-4 py-3 text-sm">${l.user_name || ''}</td>
    <td class="px-4 py-3 text-sm">${l.email || ''}</td>
    <td class="px-4 py-3 text-sm">${roleLabels[l.role] || l.role || ''}</td>
    <td class="px-4 py-3 text-sm"><span class="px-2 py-1 rounded-full text-xs ${l.action === 'change' ? 'bg-blue-50 text-blue-600' : 'bg-amber-50 text-amber-700'}">${actionLabels[l.action] || l.action || ''}</span></td>
  </tr>`).join('');
  return `<div class="bg-white rounded-2xl p-5 border border-blue-100">
    <h3 class="font-bold mb-1 flex items-center gap-2"><i data-lucide="key-round" class="w-5 h-5 text-primary"></i>บันทึกการเปลี่ยนรหัสผ่าน</h3>
    <p class="text-xs text-gray-500 mb-4">บันทึกทุกครั้งที่มีการตั้ง/เปลี่ยนรหัสผ่านผ่านอีเมล (ล่าสุดอยู่บนสุด) — รวม ${logs.length} รายการ</p>
    <div class="overflow-x-auto"><table class="w-full text-sm">
      <thead><tr class="bg-surface text-left"><th class="px-4 py-3 font-semibold">วันที่-เวลา</th><th class="px-4 py-3 font-semibold">ชื่อ-สกุล</th><th class="px-4 py-3 font-semibold">อีเมล</th><th class="px-4 py-3 font-semibold">บทบาท</th><th class="px-4 py-3 font-semibold">ประเภท</th></tr></thead>
      <tbody>${rows || '<tr><td colspan="5" class="px-4 py-8 text-center text-gray-400">ยังไม่มีบันทึกการเปลี่ยนรหัสผ่าน</td></tr>'}</tbody>
    </table></div>
  </div>`;
}

function settingsPage() {
  const roles = ['admin', 'academic', 'registrar', 'deptHead', 'executive', 'teacher', 'classTeacher', 'student'];
  const modules = ['dashboard', 'students', 'teachers', 'advisors', 'specialTeachers', 'alumni', 'schedule', 'subjects', 'grades', 'engResults', 'teacherDirectory', 'services', 'tracking', 'resultTracking', 'gradeTracking', 'fileTracking', 'leave', 'survey'];
  const moduleLabels = { dashboard: 'หน้าหลัก', students: 'ข้อมูลนักศึกษา', teachers: 'ข้อมูลอาจารย์', advisors: 'ข้อมูลอาจารย์ที่ปรึกษา', specialTeachers: 'ข้อมูลอาจารย์พิเศษ', alumni: 'ข้อมูลศิษย์เก่า', schedule: 'ปฏิทินกิจกรรมวิชาการ', subjects: 'รายวิชาที่เปิดสอน', grades: 'ผลการเรียน', engResults: 'ผลสอบ ENG', teacherDirectory: 'ทำเนียบอาจารย์', services: 'บริการอื่นๆ', tracking: 'ติดตามการส่งรายละเอียดรายวิชา', resultTracking: 'ติดตามการส่งผลการดำเนินงานรายวิชา', gradeTracking: 'ติดตามการส่งเกรดรายวิชา', fileTracking: 'ติดตามส่งแฟ้มรายวิชา', leave: 'ระบบการลาของนักศึกษา', survey: 'แบบประเมินความพึงพอใจ' };
  const roleLabels = { admin: 'ผู้ดูแลระบบ', academic: 'เจ้าหน้าที่งานวิชาการ', registrar: 'งานทะเบียน', deptHead: 'ประธานสาขา', executive: 'ผู้บริหาร', teacher: 'อาจารย์', classTeacher: 'อ.ประจำชั้น', student: 'นักศึกษา' };

  const users = applyFilters(getDataByType('user'));
  const total = users.length; const paged = paginate(users);
  const usersTable = paged.map(u => `<tr class="border-t hover:bg-gray-50">
    <td class="px-4 py-3">${u.name || ''}</td>
    <td class="px-4 py-3">${u.email || u.national_id || ''}</td>
    <td class="px-4 py-3"><span class="px-2 py-1 rounded-full text-xs bg-surface">${roleLabels[u.role] || u.role}</span></td>
    <td class="px-4 py-3"><div class="flex gap-1"><button onclick="showEditUserModal('${u.__backendId}')" class="text-blue-400 hover:text-blue-600" title="แก้ไข"><i data-lucide="pencil" class="w-4 h-4"></i></button><button onclick="deleteRecord('${u.__backendId}')" class="text-red-400 hover:text-red-600" title="ลบ"><i data-lucide="trash-2" class="w-4 h-4"></i></button></div></td>
  </tr>`).join('');

  const _stab = APP._settingsTab || 'users';
  const _tabBar = `<h2 class="text-xl font-bold text-gray-800 mb-6"><i data-lucide="settings" class="w-6 h-6 inline mr-2"></i>ตั้งค่าระบบ</h2>
  <div class="flex gap-1 mb-5 border-b">
    <button onclick="changeSettingsTab('users')" class="px-4 py-2 text-sm font-medium ${_stab === 'users' ? 'border-b-2 border-primary text-primary' : 'text-gray-500 hover:text-gray-700'}"><i data-lucide="users" class="w-4 h-4 inline mr-1"></i>จัดการผู้ใช้งาน</button>
    <button onclick="changeSettingsTab('pwlog')" class="px-4 py-2 text-sm font-medium ${_stab === 'pwlog' ? 'border-b-2 border-primary text-primary' : 'text-gray-500 hover:text-gray-700'}"><i data-lucide="key-round" class="w-4 h-4 inline mr-1"></i>บันทึกการเปลี่ยนรหัสผ่าน</button>
  </div>`;
  if (_stab === 'pwlog') return _tabBar + passwordLogSection();
  return _tabBar + `
  
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
    <p class="text-xs text-gray-400 mb-3">หมายเหตุ: ช่อง "แบบประเมินความพึงพอใจ" ของผู้ดูแลระบบ = สิทธิ์จัดการ/สรุปผลแบบประเมิน · บทบาทอื่น = สิทธิ์ทำแบบประเมิน</p>
    <div class="overflow-x-auto"><table class="w-full text-sm" style="min-width:980px">
      <thead><tr class="bg-surface"><th class="px-3 py-2 text-left font-semibold whitespace-nowrap">โมดูล</th>${roles.map(r => `<th class="px-3 py-2 text-center font-semibold whitespace-nowrap">${roleLabels[r]}</th>`).join('')}</tr></thead>
      <tbody>${modules.map(m => `<tr class="border-t hover:bg-gray-50"><td class="px-3 py-2 font-medium whitespace-nowrap">${moduleLabels[m]}</td>${roles.map(r => { const mk = (m === 'survey' && r === 'admin') ? 'surveyManage' : m; return `<td class="px-3 py-2 text-center"><label class="inline-flex"><input type="checkbox" ${APP.permissions[r]?.[mk] ? 'checked' : ''} onchange="togglePermission('${r}','${mk}',this.checked)" class="w-4 h-4 rounded border-gray-300 text-primary focus:ring-primary"></label></td>`; }).join('')}</tr>`).join('')}</tbody>
    </table></div>
  `;
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
          <option value="registrar">เจ้าหน้าที่งานทะเบียน</option>
          <option value="deptHead">ประธานสาขาวิชา</option>
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
    else if (role === 'registrar' || role === 'deptHead') { obj.username = fd.get('username') || fd.get('name'); obj.password = fd.get('password') || '123456'; if (role === 'deptHead' && fd.get('department')) obj.department = fd.get('department'); }
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
  } else if (role === 'registrar' || role === 'deptHead') {
    fieldsDiv.innerHTML = `<div><label class="block text-xs text-gray-600 mb-1">Username * <span class="text-gray-400">(ใช้ Login)</span></label>
      <input name="username" required class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="ตั้ง Username"></div>
      <div><label class="block text-xs text-gray-600 mb-1">รหัสผ่าน (ถ้าไม่ระบุจะเป็น 123456)</label>
      <input name="password" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="กรุณาตั้งรหัสผ่าน"></div>
      ${role === 'deptHead' ? `<div><label class="block text-xs text-gray-600 mb-1">สาขาวิชา <span class="text-gray-400">(เว้นว่างได้ ระบบจะอิงชื่อ-สกุลให้ตรงกับอาจารย์)</span></label>
      <input name="department" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="เช่น การพยาบาลผู้ใหญ่"></div>` : ''}`;
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
// แปลงค่าเวลาให้เป็น HH:MM (กัน Google Sheet แปลงเป็น serial date เช่น 1899-12-30T..)
function fmtSchedTime(v) {
  const s = String(v == null ? '' : v).trim();
  if (!s) return '';
  let m = s.match(/^(\d{1,2}):(\d{2})/); if (m) return m[1].padStart(2, '0') + ':' + m[2];
  m = s.match(/T(\d{2}):(\d{2})/); if (m) return m[1] + ':' + m[2];
  m = s.match(/\b(\d{1,2}):(\d{2}):\d{2}\b/); if (m) return m[1].padStart(2, '0') + ':' + m[2];
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return '';
  return s;
}
function schedTimeRange(e) {
  const st = fmtSchedTime(e.schedule_time), et = fmtSchedTime(e.schedule_time_end);
  return st + (et ? ' - ' + et : '');
}

// นักศึกษา: เห็นเฉพาะ "วันสอบ" ของชั้นปีตัวเอง (วันหยุด/กิจกรรม/ประกาศ เห็นทุกคน) — บทบาทอื่นเห็นทุกอย่าง
function filterScheduleForStudent(records) {
  if (APP.currentRole !== 'student') return records;
  const yr = norm((APP.currentUser && APP.currentUser.data && APP.currentUser.data.year_level) || '');
  return records.filter(e => {
    if (!norm(e.schedule_date)) return true;               // ประกาศ (ไม่ใช่รายการปฏิทิน)
    const isExam = norm(e.schedule_type).includes('สอบ');
    if (!isExam) return true;                               // วันหยุด/กิจกรรม แสดงทุกคน
    if (!yr) return true;                                   // ไม่ทราบชั้นปี → แสดงไว้ก่อน
    const yl = norm(e.year_level);
    return !yl || yl === yr;                                // สอบ: เฉพาะชั้นปีตัวเอง หรือ ทุกชั้นปี
  });
}

// ประกาศที่ระบบสร้างอัตโนมัติจากรายการปฏิทิน — ไม่ต้องโชว์ซ้ำบนปฏิทิน (มีรายการปฏิทินอยู่แล้ว)
function isAutoScheduleAnnouncement(e) {
  if (norm(e.schedule_date)) return false;
  const t = norm(e.announcement_title);
  return t.indexOf('แจ้งกำหนดสอบ') !== -1 || t.indexOf('แจ้งเตือนปฏิทิน') !== -1;
}

// สีจุดตามประเภทกิจกรรม (บนปฏิทิน)
function calEventDot(e) {
  const st = (e.schedule_type || e.event_type || '').trim();
  if (st.includes('สอบ')) return 'bg-red-500';
  if (st.includes('วันหยุด')) return 'bg-green-500';
  if (st === 'กิจกรรม') return 'bg-purple-500';
  return 'bg-blue-500';
}
function calEventTitle(e) {
  const st = (e.schedule_type || e.event_type || '').trim();
  const name = e.subject_name || e.announcement_title || '';
  return (st ? '[' + st + '] ' : '') + String(name).replace(/"/g, '');
}
function calChangeMonth(delta, containerId) {
  const now = new Date();
  if (APP._calYear == null) { APP._calYear = now.getFullYear(); APP._calMonth = now.getMonth(); }
  let m = APP._calMonth + delta, y = APP._calYear;
  if (m < 0) { m = 11; y--; } else if (m > 11) { m = 0; y++; }
  APP._calMonth = m; APP._calYear = y;
  renderCalendar(containerId || 'scheduleCalendar');
}
function calToday(containerId) {
  const now = new Date();
  APP._calYear = now.getFullYear(); APP._calMonth = now.getMonth();
  renderCalendar(containerId || 'scheduleCalendar');
}

function renderCalendar(containerId) {
  const el = document.getElementById(containerId); if (!el) return;
  const now = new Date();
  if (APP._calYear == null || APP._calMonth == null) { APP._calYear = now.getFullYear(); APP._calMonth = now.getMonth(); }
  const year = APP._calYear, month = APP._calMonth;
  const firstDay = new Date(year, month, 1).getDay();
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const monthNames = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];
  const events = filterScheduleForStudent([...getDataByType('announcement'), ...getDataByType('schedule')].filter(e => !isAutoScheduleAnnouncement(e)));
  const isThisMonth = (year === now.getFullYear() && month === now.getMonth());

  let h = `<div class="flex items-center justify-between mb-3">
      <button onclick="calChangeMonth(-1,'${containerId}')" class="p-1.5 rounded-lg border hover:bg-surface" title="เดือนก่อนหน้า"><i data-lucide="chevron-left" class="w-4 h-4"></i></button>
      <div class="flex items-center gap-2"><span class="font-bold">${monthNames[month]} ${year + 543}</span>${!isThisMonth ? `<button onclick="calToday('${containerId}')" class="text-xs px-2 py-0.5 rounded-lg border text-primary hover:bg-surface">วันนี้</button>` : ''}</div>
      <button onclick="calChangeMonth(1,'${containerId}')" class="p-1.5 rounded-lg border hover:bg-surface" title="เดือนถัดไป"><i data-lucide="chevron-right" class="w-4 h-4"></i></button>
    </div>`;
  h += '<div class="grid grid-cols-7 gap-px text-center text-xs">';
  ['อา', 'จ', 'อ', 'พ', 'พฤ', 'ศ', 'ส'].forEach((d, idx) => { const c = idx === 0 ? 'text-red-500' : idx === 6 ? 'text-indigo-500' : 'text-gray-500'; h += `<div class="py-1 font-semibold ${c}">${d}</div>`; });
  for (let i = 0; i < firstDay; i++)h += '<div></div>';
  for (let d = 1; d <= daysInMonth; d++) {
    const dateStr = `${year}-${String(month + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
    const dayEvents = events.filter(e => (e.schedule_date || e.announcement_date || '').startsWith(dateStr));
    const isToday = isThisMonth && d === now.getDate();
    const dow = new Date(year, month, d).getDay();
    const isHoliday = dayEvents.some(e => (e.schedule_type || e.event_type || '').includes('วันหยุด'));
    const cellCls = isToday ? 'bg-primary text-white'
      : isHoliday ? 'bg-green-100 ring-1 ring-green-300'
        : dow === 0 ? 'bg-red-50'
          : dow === 6 ? 'bg-indigo-50' : '';
    const numCls = isToday ? 'font-bold'
      : isHoliday ? 'text-green-700 font-semibold'
        : dow === 0 ? 'text-red-500 font-medium'
          : dow === 6 ? 'text-indigo-500 font-medium' : '';
    const clickable = dayEvents.length ? `onclick="showCalendarDayModal('${dateStr}')" style="cursor:pointer"` : '';
    const dots = dayEvents.slice(0, 8).map(e => `<span class="inline-block w-2 h-2 rounded-full ${calEventDot(e)}" title="${calEventTitle(e)}"></span>`).join('');
    h += `<div class="cal-day p-1 min-h-[44px] rounded-lg ${cellCls}" ${clickable}>
      <div class="text-xs ${numCls}">${d}</div>
      ${dayEvents.length ? `<div class="flex flex-wrap gap-0.5 justify-center mt-1">${dots}${dayEvents.length > 8 ? `<span class="text-[9px] text-gray-400 leading-none">+${dayEvents.length - 8}</span>` : ''}</div>` : ''}
    </div>`;
  }
  h += '</div>';
  h += '<div class="flex flex-wrap items-center justify-center gap-x-3 gap-y-1 mt-3 text-xs text-gray-500">'
    + '<span class="inline-flex items-center gap-1"><span class="w-2 h-2 rounded-full bg-red-500"></span>สอบ</span>'
    + '<span class="inline-flex items-center gap-1"><span class="w-2 h-2 rounded-full bg-purple-500"></span>กิจกรรม</span>'
    + '<span class="inline-flex items-center gap-1"><span class="w-2 h-2 rounded-full bg-green-500"></span>วันหยุด</span>'
    + '<span class="inline-flex items-center gap-1"><span class="w-2 h-2 rounded-full bg-blue-500"></span>ประกาศ/อื่นๆ</span>'
    + '<span class="mx-1 text-gray-300">|</span>'
    + '<span class="inline-flex items-center gap-1"><span class="w-3 h-3 rounded bg-red-50 border border-red-200"></span>อาทิตย์</span>'
    + '<span class="inline-flex items-center gap-1"><span class="w-3 h-3 rounded bg-indigo-50 border border-indigo-200"></span>เสาร์</span>'
    + '<span class="inline-flex items-center gap-1"><span class="w-3 h-3 rounded bg-primary"></span>วันนี้</span>'
    + '</div>';
  el.innerHTML = h;
  if (window.lucide) lucide.createIcons();
}

// การ์ดรายละเอียด 1 กิจกรรม (ใช้ในป็อปอัพวันที่)
function scheduleEventCardHTML(e, canManage) {
  const isAnn = !norm(e.schedule_date) && norm(e.announcement_date);
  if (isAnn) {
    return `<div class="border border-blue-100 rounded-xl p-3 bg-blue-50">
      <div class="font-semibold text-sm">📢 ${e.announcement_title || '(ประกาศ)'}</div>
      ${e.announcement_content ? `<div class="text-xs text-gray-600 mt-1 whitespace-pre-line">${e.announcement_content}</div>` : ''}
    </div>`;
  }
  const type = norm(e.schedule_type);
  const isExam = type.includes('สอบ');
  const timeRange = schedTimeRange(e);
  const subjects = norm(e.subject_name).replace(/,\s*/g, ', ');
  const pcd = norm(e.proctor_change_date);
  const isSplit = norm(e.exam_split) !== '';
  const room1Label = isExam && isSplit ? 'ห้องสอบที่ 1' : 'ห้อง';
  const proc1 = norm(e.proctor).replace(/,\s*/g, ', ');
  const proc2 = norm(e.proctor2).replace(/,\s*/g, ', ');
  return `<div class="border border-gray-100 rounded-xl p-3 bg-surface">
    <div class="flex items-start justify-between gap-2">
      <div class="font-semibold text-sm">${(subjects || '-').replace(/,\s*/g, '<br>')}</div>
      <span class="shrink-0">${scheduleTypeBadge(e.schedule_type)}</span>
    </div>
    <div class="grid grid-cols-2 gap-2 mt-2">
      ${infoRow('เวลา', timeRange)}
      ${infoRow(room1Label, e.room)}
      ${infoRow('ชั้นปี', norm(e.year_level) ? 'ชั้นปีที่ ' + norm(e.year_level) : 'ทุกชั้นปี')}
      ${isExam ? infoRow('ครั้งที่', e.exam_round) : ''}
      ${isExam ? infoRow('จำนวนนักศึกษา' + (isSplit ? ' (ห้อง 1)' : ''), e.student_count) : ''}
      ${isExam ? infoRow('อาจารย์ผู้คุมสอบ' + (isSplit ? ' (ห้อง 1)' : ''), proc1) : ''}
      ${isExam && isSplit ? infoRow('ห้องสอบที่ 2', e.room2) : ''}
      ${isExam && isSplit ? infoRow('จำนวนนักศึกษา (ห้อง 2)', e.student_count2) : ''}
      ${isExam && isSplit ? infoRow('อาจารย์ผู้คุมสอบ (ห้อง 2)', proc2) : ''}
      ${isExam && pcd ? infoRow('วันที่เปลี่ยนผู้คุมสอบ', (typeof toBuddhistDate === 'function' && toBuddhistDate(pcd)) || pcd) : ''}
    </div>
    ${canManage ? `<div class="flex justify-end gap-2 mt-2">
      <button onclick="showEditScheduleModal('${e.__backendId}')" class="text-xs px-3 py-1.5 rounded-lg bg-blue-50 text-blue-600 hover:bg-blue-100"><i data-lucide="pencil" class="w-3.5 h-3.5 inline"></i> แก้ไข</button>
      <button onclick="deleteRecord('${e.__backendId}')" class="text-xs px-3 py-1.5 rounded-lg bg-red-50 text-red-600 hover:bg-red-100"><i data-lucide="trash-2" class="w-3.5 h-3.5 inline"></i> ลบ</button>
    </div>` : ''}
  </div>`;
}

// ป็อปอัพแสดงกิจกรรมของวันที่เลือก — ถ้าหลายรายการ มีปุ่มเลื่อนหน้า
function showCalendarDayModal(dateStr, idx) {
  idx = idx || 0;
  const canManage = APP.currentRole === 'admin' || APP.currentRole === 'academic' || APP.currentRole === 'executive' || APP.currentRole === 'registrar';
  const events = filterScheduleForStudent([...getDataByType('schedule'), ...getDataByType('announcement')]
    .filter(e => (e.schedule_date || e.announcement_date || '').startsWith(dateStr))
    .filter(e => !isAutoScheduleAnnouncement(e)));
  const dateTh = (typeof toBuddhistDate === 'function' && toBuddhistDate(dateStr)) || dateStr;
  if (!events.length) { showModal('รายการวันที่ ' + dateTh, '<p class="text-center text-gray-400 py-6">ไม่มีรายการในวันนี้</p>'); return; }
  if (idx < 0) idx = 0; if (idx >= events.length) idx = events.length - 1;
  const card = scheduleEventCardHTML(events[idx], canManage);
  const nav = events.length > 1 ? `<div class="flex items-center justify-between mt-3">
      <button ${idx === 0 ? 'disabled' : ''} onclick="showCalendarDayModal('${dateStr}',${idx - 1})" class="px-3 py-1.5 rounded-lg border text-sm ${idx === 0 ? 'opacity-40 cursor-not-allowed' : 'hover:bg-surface'}"><i data-lucide="chevron-left" class="w-4 h-4 inline"></i> ก่อนหน้า</button>
      <span class="text-xs text-gray-500">${idx + 1} / ${events.length}</span>
      <button ${idx === events.length - 1 ? 'disabled' : ''} onclick="showCalendarDayModal('${dateStr}',${idx + 1})" class="px-3 py-1.5 rounded-lg border text-sm ${idx === events.length - 1 ? 'opacity-40 cursor-not-allowed' : 'hover:bg-surface'}">ถัดไป <i data-lucide="chevron-right" class="w-4 h-4 inline"></i></button>
    </div>` : '';
  showModal('รายการวันที่ ' + dateTh + (events.length > 1 ? ' (' + events.length + ' รายการ)' : ''), card + nav);
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
  if (r.isOk) { showToast('อนุมัติการลาสำเร็จ'); renderCurrentPage(); updateNotifBadge(); } else showToast('เกิดข้อผิดพลาด: ' + (r.error || ''), 'error');
}

async function rejectLeave(id, approvalField) {
  const rec = APP.allData.find(d => d.__backendId === id); if (!rec) return;
  if (!confirm('ยืนยันการปฏิเสธใบลานี้?')) return;
  rec[approvalField] = 'ปฏิเสธ';
  rec.leave_status = 'ปฏิเสธ';

  showToast('กำลังบันทึก...', 'loading');
  const r = await GSheetDB.update(rec);
  hideLoadingToast && hideLoadingToast();
  if (r.isOk) { showToast('ปฏิเสธการลาสำเร็จ'); renderCurrentPage(); updateNotifBadge(); } else showToast('เกิดข้อผิดพลาด: ' + (r.error || ''), 'error');
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
    // ลบรายการปฏิทิน → ลบประกาศแจ้งเตือนที่ระบบสร้างให้ด้วย (ถ้ามี)
    if (r.isOk && rec.type === 'schedule') {
      const linked = findLinkedScheduleAnnouncements(rec).sort((a, b) => (b.__rowIndex || 0) - (a.__rowIndex || 0));
      for (const a of linked) { try { await GSheetDB.delete(a); } catch (_) { } }
    }
    hideLoadingToast();
    if (r.isOk) { showToast('ลบสำเร็จ'); closeModal(); renderCurrentPage() } else { showToast('เกิดข้อผิดพลาด', 'error'); closeModal() }
  });
}

// หาประกาศแจ้งเตือนที่ระบบสร้างจากรายการปฏิทิน (จับคู่ด้วยวันที่ + ชื่อวิชา)
function findLinkedScheduleAnnouncements(rec) {
  const date = norm(rec.schedule_date);
  const subj = norm(rec.subject_name);
  return getDataByType('announcement').filter(a => {
    if (!isAutoScheduleAnnouncement(a)) return false;
    if (date && norm(a.announcement_date) !== date) return false;
    if (subj) { const hay = norm(a.announcement_title) + ' ' + norm(a.announcement_content); return subj.split(/,\s*/).some(x => x.trim() && hay.indexOf(x.trim()) !== -1); }
    return true;
  });
}

// ======================== GENERIC EDIT HELPER ========================
async function editRecord(id, formId) {
  const rec = APP.allData.find(d => d.__backendId === id); if (!rec) return;
  const form = document.getElementById(formId); if (!form) return;
  const btn = form.querySelector('[type="submit"]');
  const origText = btn ? btn.innerHTML : '';
  if (btn) { btn.disabled = true; btn.innerHTML = '<span class="flex items-center justify-center gap-2"><img src="https://cdn.jsdelivr.net/gh/JOB-BCNB-P/picture/cat_run_transparent.gif" class="cat-run-inline" alt="">กำลังบันทึก...</span>'; lucide.createIcons() }
  const fd = new FormData(form);
  fd.forEach((v, k) => { if (k !== '__backendId') rec[k] = v });
  // ฟอร์มที่ใช้ตัวเลือกคำนำหน้า → รวมคำนำหน้า + ชื่อ เป็นชื่อเต็ม
  if (form.querySelector('[name="title_prefix"]')) { rec.name = combineName(form); } // เก็บ title_prefix แยกด้วย (ถ้าชีตมีคอลัมน์นี้ — ถ้าไม่มีจะถูกละเว้นอัตโนมัติ)
  const r = await GSheetDB.update(rec);
  if (btn) { btn.disabled = false; btn.innerHTML = origText; lucide.createIcons() }
  if (r.isOk) { showToast('แก้ไขข้อมูลสำเร็จ'); closeModal(); renderCurrentPage() } else showToast('เกิดข้อผิดพลาด: ' + (r.error || ''), 'error');
}

// ======================== EDIT MODALS ========================
function showEditStudentModal(id) {
  const s = APP.allData.find(d => d.__backendId === id); if (!s) return;
  showModal('แก้ไขข้อมูลนักศึกษา', `
    <form id="editStudentForm" class="space-y-3">
      ${titlePrefixField(s.name || '')}
      <div class="grid grid-cols-1 md:grid-cols-2 gap-3">
        <div><label class="block text-xs text-gray-600 mb-1">รหัสนักศึกษา</label><input name="student_id" value="${s.student_id || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">รุ่นที่</label><input name="batch" value="${s.batch || ''}" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="เช่น 36"></div>
        <div><label class="block text-xs text-gray-600 mb-1">เลขบัตรประชาชน <span class="text-gray-400">(เว้นว่าง = ไม่เปลี่ยน)</span></label><input name="national_id" value="" maxlength="13" placeholder="กรอกเฉพาะเมื่อต้องการแก้ไข" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">สถานภาพ</label><select name="status" class="w-full border rounded-xl px-3 py-2 text-sm"><option ${s.status === 'กำลังศึกษา' ? 'selected' : ''}>กำลังศึกษา</option><option ${s.status === 'พักการศึกษา' ? 'selected' : ''}>พักการศึกษา</option><option ${s.status === 'ลาออก' ? 'selected' : ''}>ลาออก</option><option ${s.status === 'ขอโอนย้ายสถานศึกษา' ? 'selected' : ''}>ขอโอนย้ายสถานศึกษา</option><option ${s.status === 'สำเร็จการศึกษา' ? 'selected' : ''}>สำเร็จการศึกษา</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">ชั้นปี</label><select name="year_level" class="w-full border rounded-xl px-3 py-2 text-sm"><option ${norm(s.year_level) === '1' ? 'selected' : ''}>1</option><option ${norm(s.year_level) === '2' ? 'selected' : ''}>2</option><option ${norm(s.year_level) === '3' ? 'selected' : ''}>3</option><option ${norm(s.year_level) === '4' ? 'selected' : ''}>4</option><option value="จบ" ${norm(s.year_level) === 'จบ' ? 'selected' : ''}>จบ (สำเร็จการศึกษา)</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">ห้อง</label><input name="room" value="${s.room || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">เพศ</label><select name="gender" class="w-full border rounded-xl px-3 py-2 text-sm"><option value="" ${!norm(s.gender) ? 'selected' : ''}>ไม่ระบุ</option><option ${norm(s.gender) === 'ชาย' ? 'selected' : ''}>ชาย</option><option ${norm(s.gender) === 'หญิง' ? 'selected' : ''}>หญิง</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">โทรศัพท์</label><input name="phone" value="${s.phone || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">E-mail</label><input name="email" value="${s.email || ''}" type="email" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">ชื่อผู้ปกครอง</label><input name="parent_name" value="${s.parent_name || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">โทรผู้ปกครอง</label><input name="parent_phone" value="${s.parent_phone || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">อาจารย์ที่ปรึกษา</label><input name="advisor" value="${s.advisor || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">วันที่ลาออก/พักการศึกษา</label><input name="status_date" value="${(s.status_date || '').replace(/"/g, '&quot;')}" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="เช่น 1 มิ.ย. 2569"></div>
        <div class="md:col-span-2"><label class="block text-xs text-gray-600 mb-1">เหตุผลการลาออก/พักการศึกษา</label><input name="status_reason" value="${(s.status_reason || '').replace(/"/g, '&quot;')}" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="กรอกเมื่อสถานะเป็น ลาออก / พักการศึกษา / โอนย้าย"></div>
      </div>
      ${transcriptFieldsHTML(s)}
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
      <div><label class="block text-xs text-gray-600 mb-1">สาขาวิชาที่รับผิดชอบ <span class="text-gray-400">(มี 2 สาขา คั่นด้วย ,)</span></label><input name="department" list="editSubjectDeptList" value="${(s.department || '').replace(/"/g, '&quot;')}" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="เลือก/พิมพ์ คั่นด้วย , ถ้ามี 2 สาขา">${deptDatalistHTML('editSubjectDeptList')}</div>
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-xs text-gray-600 mb-1">ชั้นปี</label><select name="year_level" class="w-full border rounded-xl px-3 py-2 text-sm"><option ${norm(s.year_level) === '1' ? 'selected' : ''}>1</option><option ${norm(s.year_level) === '2' ? 'selected' : ''}>2</option><option ${norm(s.year_level) === '3' ? 'selected' : ''}>3</option><option ${norm(s.year_level) === '4' ? 'selected' : ''}>4</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">รุ่นที่</label><input name="batch" value="${s.batch || ''}" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="เช่น 28"></div>
        <div><label class="block text-xs text-gray-600 mb-1">ห้อง</label><input name="room" value="${s.room || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        ${creditFields(s)}
        <div><label class="block text-xs text-gray-600 mb-1">ภาคการศึกษา</label><select name="semester" class="w-full border rounded-xl px-3 py-2 text-sm"><option value="1" ${normSem(s.semester) === '1' ? 'selected' : ''}>1</option><option value="2" ${normSem(s.semester) === '2' ? 'selected' : ''}>2</option><option value="3" ${normSem(s.semester) === '3' ? 'selected' : ''}>ฤดูร้อน</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">ปีการศึกษา</label><input name="academic_year" value="${s.academic_year || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      </div>
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึกการแก้ไข</button>
    </form>
  `);
  document.getElementById('editSubjectForm').onsubmit = (e) => { e.preventDefault(); editRecord(id, 'editSubjectForm') };
}

function showEditScheduleModal(id) {
  const s = APP.allData.find(d => d.__backendId === id); if (!s) return;
  showModal('แก้ไขรายการปฏิทินกิจกรรมวิชาการ', `
    <form id="editScheduleForm" class="space-y-3">
      ${scheduleFormBody(s, false)}
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึกการแก้ไข</button>
    </form>
  `);
  renderSchedSubjectChips();
  renderSchedProctorChips(1);
  renderSchedProctorChips(2);
  updateSchedSplitState();
  window._schedWasExam = norm(s.schedule_type).includes('สอบ');
  document.getElementById('editScheduleForm').onsubmit = async (e) => {
    e.preventDefault();
    const fd = new FormData(e.target);
    if (!(fd.get('subject_name') || '').trim()) { showToast('กรุณาระบุรายวิชา/กิจกรรม', 'error'); return; }
    // อ่านค่าการแจ้งเตือนก่อนปิด modal
    const notifyEl = document.getElementById('schedNotify');
    const doNotify = !!(notifyEl && notifyEl.checked);
    const roles = doNotify ? annCollectRoles() : '';
    const lineEl = document.getElementById('schedNotifyLine');
    const sendLine = !!(lineEl && lineEl.checked);
    await editRecord(id, 'editScheduleForm');
    if (doNotify) {
      const rec = APP.allData.find(d => d.__backendId === id);
      if (rec) { await createScheduleAnnouncement(rec, roles, sendLine); showToast('บันทึกและสร้างประกาศแจ้งเตือนแล้ว'); }
    }
  };
}

function showEditGradeModal(id) {
  const g = APP.allData.find(d => d.__backendId === id); if (!g) return;
  showModal('แก้ไขผลการเรียน', `
    <form id="editGradeForm" class="space-y-3">
      <div><label class="block text-xs text-gray-600 mb-1">นักศึกษา</label><select id="egStudent" name="student_id" class="w-full border rounded-xl px-3 py-2 text-sm" onchange="_rebuildGradeSubjectSelect('eg')">${studentOptionsHTML(g.student_id)}</select></div>
      <div><label class="block text-xs text-gray-600 mb-1">ปีการศึกษา</label><input id="egYear" name="academic_year" value="${g.academic_year || ''}" class="w-full border rounded-xl px-3 py-2 text-sm" onchange="_rebuildGradeSubjectSelect('eg')" oninput="_rebuildGradeSubjectSelect('eg')"></div>
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-xs text-gray-600 mb-1">รหัสวิชา</label><select id="egCode" name="subject_code" class="w-full border rounded-xl px-3 py-2 text-sm" onchange="_onGradeCodeChange('eg')"></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">รายวิชา <span class="text-gray-400">(อัตโนมัติ)</span></label><input id="egName" name="subject_name" required readonly class="w-full border rounded-xl px-3 py-2 text-sm bg-gray-50" placeholder="เลือกรหัสวิชาก่อน"></div>
        <div><label class="block text-xs text-gray-600 mb-1">เกรด</label><select name="grade" class="w-full border rounded-xl px-3 py-2 text-sm"><option ${g.grade === 'A' ? 'selected' : ''}>A</option><option ${g.grade === 'B+' ? 'selected' : ''}>B+</option><option ${g.grade === 'B' ? 'selected' : ''}>B</option><option ${g.grade === 'C+' ? 'selected' : ''}>C+</option><option ${g.grade === 'C' ? 'selected' : ''}>C</option><option ${g.grade === 'D+' ? 'selected' : ''}>D+</option><option ${g.grade === 'D' ? 'selected' : ''}>D</option><option ${g.grade === 'F' ? 'selected' : ''}>F</option></select></div>
        <div><label class="block text-xs text-gray-600 mb-1">หน่วยกิต</label><input id="egCredits" name="credits" type="number" value="${_gradeCredits(g) || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">ภาคการศึกษา</label><select id="egSem" name="semester" class="w-full border rounded-xl px-3 py-2 text-sm"><option value="1" ${norm(g.semester) === '1' ? 'selected' : ''}>1</option><option value="2" ${norm(g.semester) === '2' ? 'selected' : ''}>2</option></select></div>
      </div>
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึกการแก้ไข</button>
    </form>
  `);
  // เตรียมรายการรหัสวิชา + คงค่ารหัส/ชื่อเดิมไว้ (เผื่อรหัสเดิมไม่มีในชีต subject)
  (function () {
    const subs = _gradeSubjectsFor(g.student_id, g.academic_year);
    const code = norm(g.subject_code);
    if (code && !subs.some(s => norm(s.subject_code) === code)) {
      subs.unshift({ subject_code: code, subject_name: g.subject_name, credits: g.credits, semester: g.semester });
    }
    window._egSubjects = subs;
    const codeSel = document.getElementById('egCode');
    if (codeSel) codeSel.innerHTML = _gradeCodeOptions(subs, code);
    const nameEl = document.getElementById('egName');
    if (nameEl) nameEl.value = norm(g.subject_name);
  })();
  document.getElementById('editGradeForm').onsubmit = (e) => { e.preventDefault(); editRecord(id, 'editGradeForm') };
}

function showEditEngModal(id) {
  const e = APP.allData.find(d => d.__backendId === id); if (!e) return;
  const isSbch = e.eng_type === 'สบช.';
  const initTotal = Number(e.eng_score) || 0;
  const initLevel = e.eng_level || (isSbch ? getEngLevel(initTotal) : '');
  const initStatus = e.eng_status || (isSbch ? (initTotal >= engPassThreshold(e.student_id, e.academic_year) ? 'ผ่าน' : 'ไม่ผ่าน') : '');
  showModal('แก้ไขผลสอบภาษาอังกฤษ', `
    <form id="editEngForm" class="space-y-3">
      <div><label class="block text-xs text-gray-600 mb-1">นักศึกษา</label><select name="student_id" onchange="updateEngSbchLevelStatus('edit')" class="w-full border rounded-xl px-3 py-2 text-sm">${studentOptionsHTML(e.student_id)}</select></div>
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-xs text-gray-600 mb-1">รูปแบบการสอบ *</label>
          <select id="editEngType" onchange="updateEngTypeForm('edit')" class="w-full border rounded-xl px-3 py-2 text-sm">${engTypeOptions(e.eng_type || '')}</select>
        </div>
        <div><label class="block text-xs text-gray-600 mb-1">สอบครั้งที่</label><input id="editEngAttempt" type="number" value="${e.eng_attempt || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
        <div><label class="block text-xs text-gray-600 mb-1">วันที่สอบ <span class="text-gray-400">(วว/ดด/ปปปป พ.ศ. หรือ ค.ศ.)</span></label><input id="editEngDate" type="text" value="${formatDate(e.eng_date)}" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="เช่น 05/04/2568 หรือ 05/04/2025"></div>
        <div><label class="block text-xs text-gray-600 mb-1">ปีการศึกษา</label><input id="editEngYear" type="text" oninput="updateEngSbchLevelStatus('edit')" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="เช่น 2568" value="${e.academic_year || ''}"></div>
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
          <div><label class="block text-xs text-gray-600 mb-1">คะแนนรวม</label><input id="editEngOtherScore" type="number" step="any" min="0" value="${!isSbch ? (e.eng_score || '') : ''}" class="w-full border rounded-xl px-3 py-2 text-sm" placeholder="กรอกทศนิยมได้ เช่น 85.5"></div>
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
        obj.eng_status = total >= engPassThreshold(studentId, year) ? 'ผ่าน' : 'ไม่ผ่าน';
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
        <div><label class="block text-xs text-gray-600 mb-1">สถานะ</label><select name="teacher_status" class="w-full border rounded-xl px-3 py-2 text-sm"><option ${(t.teacher_status || 'ปฏิบัติงานอยู่') === 'ปฏิบัติงานอยู่' ? 'selected' : ''}>ปฏิบัติงานอยู่</option><option ${t.teacher_status === 'ลาศึกษาต่อ' ? 'selected' : ''}>ลาศึกษาต่อ</option><option ${t.teacher_status === 'ลาออก' ? 'selected' : ''}>ลาออก</option></select></div>
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
      ${annRolesFieldHTML(a.roles || '')}
      <label class="flex items-center gap-2 bg-green-50 rounded-xl px-3 py-2 cursor-pointer"><input type="checkbox" name="line_notify" value="✓" class="w-4 h-4" ${['✓', '✔', 'true', 'yes', 'y', '1', 'ส่ง', 'แจ้ง'].includes(String(a.line_notify || '').trim().toLowerCase()) ? 'checked' : ''}><span class="text-sm text-green-700">📢 ส่งประกาศนี้เข้า LINE</span></label>
      <button type="submit" class="w-full bg-primary text-white py-2.5 rounded-xl hover:bg-primaryDark">บันทึกการแก้ไข</button>
    </form>
  `);
  document.getElementById('editAnnForm').onsubmit = (e) => { e.preventDefault(); a.line_notify = e.target.querySelector('[name="line_notify"]').checked ? '✓' : ''; a.roles = annCollectRoles(); editRecord(id, 'editAnnForm') };
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
  const roleLabels = { admin: 'ผู้ดูแลระบบ', academic: 'เจ้าหน้าที่งานวิชาการ', registrar: 'เจ้าหน้าที่งานทะเบียน', deptHead: 'ประธานสาขาวิชา', executive: 'ผู้บริหาร', teacher: 'อาจารย์', classTeacher: 'อาจารย์ประจำชั้น', student: 'นักศึกษา' };
  showModal('แก้ไขผู้ใช้งาน', `
    <form id="editUserForm" class="space-y-3">
      <div><label class="block text-xs text-gray-600 mb-1">ชื่อ-สกุล</label><input name="name" value="${u.name || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <div><label class="block text-xs text-gray-600 mb-1">บทบาท</label><select name="role" class="w-full border rounded-xl px-3 py-2 text-sm"><option value="admin" ${u.role === 'admin' ? 'selected' : ''}>ผู้ดูแลระบบ</option><option value="academic" ${u.role === 'academic' ? 'selected' : ''}>เจ้าหน้าที่งานวิชาการ</option><option value="registrar" ${u.role === 'registrar' ? 'selected' : ''}>เจ้าหน้าที่งานทะเบียน</option><option value="deptHead" ${u.role === 'deptHead' ? 'selected' : ''}>ประธานสาขาวิชา</option><option value="executive" ${u.role === 'executive' ? 'selected' : ''}>ผู้บริหาร</option><option value="teacher" ${u.role === 'teacher' ? 'selected' : ''}>อาจารย์</option><option value="classTeacher" ${u.role === 'classTeacher' ? 'selected' : ''}>อาจารย์ประจำชั้น</option><option value="student" ${u.role === 'student' ? 'selected' : ''}>นักศึกษา</option></select></div>
      <div><label class="block text-xs text-gray-600 mb-1">Username <span class="text-gray-400">(งานทะเบียน/ประธานสาขา)</span></label><input name="username" value="${u.username || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
      <div><label class="block text-xs text-gray-600 mb-1">สาขาวิชา <span class="text-gray-400">(ประธานสาขา)</span></label><input name="department" value="${u.department || ''}" class="w-full border rounded-xl px-3 py-2 text-sm"></div>
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

// ======================== GLOBAL BUTTON LOADING STATE ========================
// กดปุ่มใดก็แสดงสปินเนอร์ + กันกดซ้ำ จนกว่าจะเรนเดอร์ใหม่/หมดเวลา
// ข้ามปุ่มที่ทำงานทันที (เมนู, dropdown, กระดิ่ง, popup, ออกจากระบบ, เปิด modal) และปุ่มที่ใส่ data-no-loading
(function () {
  var SKIP = /toggleSidebar|toggleDropdown|Notifications|contactPopup|handleLogout|closeModal|classList\.|Modal\(/i;
  var active = new Set();
  function clearBtn(btn) {
    btn.classList.remove('btn-spin');
    btn.style.removeProperty('min-width');
    btn.style.removeProperty('min-height');
    if (btn._loadTimer) { clearTimeout(btn._loadTimer); btn._loadTimer = null; }
    active.delete(btn);
  }
  function clearAll() { Array.prototype.slice.call(active).forEach(clearBtn); }
  document.addEventListener('click', function (e) {
    var btn = e.target.closest ? e.target.closest('button') : null;
    if (!btn) return;
    if (btn.disabled || btn.classList.contains('btn-spin')) return;
    if (btn.hasAttribute('data-no-loading')) return;
    if (SKIP.test(btn.getAttribute('onclick') || '')) return;
    var r = btn.getBoundingClientRect();
    btn.style.minWidth = r.width + 'px';
    btn.style.minHeight = r.height + 'px';
    btn.style.setProperty('--spin-color', getComputedStyle(btn).color);
    btn.classList.add('btn-spin');
    active.add(btn);
    btn._loadTimer = setTimeout(function () { clearBtn(btn); }, 1500);
  }, true);
  function observe(id) {
    var el = document.getElementById(id);
    if (el) new MutationObserver(clearAll).observe(el, { childList: true, subtree: true });
  }
  function setup() { observe('mainContent'); observe('modalContainer'); }
  if (document.readyState !== 'loading') setup();
  else document.addEventListener('DOMContentLoaded', setup);
})();

// ============================================================================
// ======================== แบบประเมินความพึงพอใจการใช้งานระบบ ========================
// ผู้ใช้ทั่วไป (7 บทบาท) ทำแบบประเมินได้ครั้งเดียวต่อปีการศึกษา (เซิร์ฟเวอร์บังคับ)
// admin: จัดการคำถาม/เปิด-ปิดแบบประเมินรายปี + ดูสรุปผล (μ, S.D., ร้อยละ, AUN-QA)
// ============================================================================

const SURVEY_DEVICES = ['คอมพิวเตอร์ตั้งโต๊ะ / โน้ตบุ๊ก', 'แท็บเล็ต', 'โทรศัพท์มือถือ'];
const SURVEY_FREQ = ['ทุกวัน', '2-3 ครั้งต่อสัปดาห์', 'สัปดาห์ละครั้ง', 'เดือนละครั้ง', 'นานๆ ครั้ง / ตามที่จำเป็น'];
const SURVEY_RATING_LABELS = { 5: 'มากที่สุด', 4: 'มาก', 3: 'ปานกลาง', 2: 'น้อย', 1: 'น้อยที่สุด' };
const SURVEY_ROLE_LABEL = { admin: 'ผู้ดูแลระบบ', academic: 'เจ้าหน้าที่งานวิชาการ', registrar: 'เจ้าหน้าที่งานทะเบียน', deptHead: 'ประธานสาขาวิชา', executive: 'ผู้บริหาร', teacher: 'อาจารย์ / อาจารย์ที่ปรึกษา', classTeacher: 'อาจารย์ประจำชั้น', student: 'นักศึกษา' };

// ชุดคำถามเริ่มต้น (อิงแบบประเมินที่ร่างไว้) — admin กดสร้างให้ปีการศึกษาที่เลือกได้
const SURVEY_DEFAULT_QUESTIONS = [
  { section: 'ด้านการเข้าถึงและการเข้าสู่ระบบ', q_type: 'rating', question_text: 'การเข้าสู่ระบบด้วยบทบาทของท่าน (username/รหัสผ่าน หรือรหัสนักศึกษา) สะดวกและใช้งานง่าย' },
  { section: 'ด้านการเข้าถึงและการเข้าสู่ระบบ', q_type: 'rating', question_text: 'ขั้นตอนการเปลี่ยน/รีเซ็ตรหัสผ่านผ่าน OTP มีความสะดวกและเข้าใจง่าย' },
  { section: 'ด้านการเข้าถึงและการเข้าสู่ระบบ', q_type: 'rating', question_text: 'เมนูและสิทธิ์การใช้งานที่ปรากฏ ตรงกับบทบาทและหน้าที่ของท่าน' },
  { section: 'ด้านการเข้าถึงและการเข้าสู่ระบบ', q_type: 'rating', question_text: 'ท่านสามารถเข้าถึงข้อมูลและฟังก์ชันที่ต้องใช้ได้โดยไม่ติดขัด' },

  { section: 'ด้านความง่ายในการใช้งาน (UI/UX)', q_type: 'rating', question_text: 'การจัดวางเมนูและการแบ่งหมวดหมู่เข้าใจง่าย' },
  { section: 'ด้านความง่ายในการใช้งาน (UI/UX)', q_type: 'rating', question_text: 'การสลับไปมาระหว่างหน้าต่างๆ ทำได้รวดเร็วและไม่สับสน' },
  { section: 'ด้านความง่ายในการใช้งาน (UI/UX)', q_type: 'rating', question_text: 'รูปแบบหน้าจอ สีสัน ตัวอักษร อ่านง่ายและสบายตา' },
  { section: 'ด้านความง่ายในการใช้งาน (UI/UX)', q_type: 'rating', question_text: 'ปุ่ม ฟอร์ม และตัวกรองข้อมูล ใช้งานง่ายและเข้าใจได้ทันที' },
  { section: 'ด้านความง่ายในการใช้งาน (UI/UX)', q_type: 'rating', question_text: 'สัญลักษณ์/ข้อความแจ้งสถานะการทำงาน ช่วยให้ทราบว่าระบบกำลังทำงานอยู่' },
  { section: 'ด้านความง่ายในการใช้งาน (UI/UX)', q_type: 'rating', question_text: 'ระบบใช้งานได้ดีบนอุปกรณ์ของท่าน (คอมพิวเตอร์/แท็บเล็ต/มือถือ)' },

  { section: 'ด้านความถูกต้องและครบถ้วนของข้อมูล', q_type: 'rating', question_text: 'ข้อมูลในระบบ (นักศึกษา อาจารย์ รายวิชา) มีความถูกต้องและเป็นปัจจุบัน' },
  { section: 'ด้านความถูกต้องและครบถ้วนของข้อมูล', q_type: 'rating', question_text: 'ผลการเรียนและการคำนวณเกรดเฉลี่ยสะสม (GPAX) มีความถูกต้อง' },
  { section: 'ด้านความถูกต้องและครบถ้วนของข้อมูล', q_type: 'rating', question_text: 'เอกสารที่ออกจากระบบ (ใบรายงานผลการเรียน/Transcript/รายงาน PDF) ถูกต้องครบถ้วน' },
  { section: 'ด้านความถูกต้องและครบถ้วนของข้อมูล', q_type: 'rating', question_text: 'ข้อมูลที่บันทึก/แก้ไข ถูกจัดเก็บได้อย่างถูกต้องและแสดงผลทันที' },
  { section: 'ด้านความถูกต้องและครบถ้วนของข้อมูล', q_type: 'rating', question_text: 'ตัวกรองและการค้นหา (ชั้นปี รุ่น สาขาวิชา ปีการศึกษา) แสดงผลตรงตามที่ต้องการ' },

  { section: 'ด้านฟังก์ชันและความสามารถของระบบ', q_type: 'rating', question_text: 'ระบบทะเบียน (นักศึกษา/อาจารย์/อาจารย์พิเศษ/ศิษย์เก่า/ปฏิทิน/รายวิชา) ตอบโจทย์การทำงาน' },
  { section: 'ด้านฟังก์ชันและความสามารถของระบบ', q_type: 'rating', question_text: 'ระบบข้อมูลอาจารย์ที่ปรึกษา (ดูนักศึกษาในความดูแล เชื่อมผลการเรียน/ผลสอบ) ใช้งานสะดวก' },
  { section: 'ด้านฟังก์ชันและความสามารถของระบบ', q_type: 'rating', question_text: 'ฟังก์ชันเลื่อนชั้นปีและบันทึกศิษย์เก่าอัตโนมัติ ทำงานถูกต้องและช่วยลดงาน' },
  { section: 'ด้านฟังก์ชันและความสามารถของระบบ', q_type: 'rating', question_text: 'ระบบติดตามการส่งงานรายวิชา ช่วยติดตามงานได้มีประสิทธิภาพ' },
  { section: 'ด้านฟังก์ชันและความสามารถของระบบ', q_type: 'rating', question_text: 'ระบบการลาของนักศึกษา (บันทึก/อนุมัติ/สรุปผล) ใช้งานได้ครบถ้วน' },
  { section: 'ด้านฟังก์ชันและความสามารถของระบบ', q_type: 'rating', question_text: 'ระบบผลสอบภาษาอังกฤษ และทำเนียบอาจารย์ มีข้อมูลครบถ้วนตามที่ต้องการ' },
  { section: 'ด้านฟังก์ชันและความสามารถของระบบ', q_type: 'rating', question_text: 'การนำเข้าข้อมูลด้วยไฟล์ CSV สะดวกและช่วยประหยัดเวลา' },

  { section: 'ด้านการแจ้งเตือนและการสื่อสาร', q_type: 'rating', question_text: 'ประกาศ/ข่าวสารในระบบ ช่วยให้ได้รับข้อมูลที่จำเป็นอย่างทันท่วงที' },
  { section: 'ด้านการแจ้งเตือนและการสื่อสาร', q_type: 'rating', question_text: 'การแจ้งเตือนผ่าน LINE มีประโยชน์และทันเวลา' },
  { section: 'ด้านการแจ้งเตือนและการสื่อสาร', q_type: 'rating', question_text: 'ข้อความแจ้งผลการทำงาน (สำเร็จ/กำลังทำงาน/ผิดพลาด) ชัดเจนและเข้าใจง่าย' },

  { section: 'ด้านความเร็วและเสถียรภาพ', q_type: 'rating', question_text: 'ระบบโหลดข้อมูลและแสดงผลได้รวดเร็ว' },
  { section: 'ด้านความเร็วและเสถียรภาพ', q_type: 'rating', question_text: 'ระบบทำงานต่อเนื่อง ไม่ค้างหรือเกิดข้อผิดพลาดบ่อย' },
  { section: 'ด้านความเร็วและเสถียรภาพ', q_type: 'rating', question_text: 'เมื่อเกิดปัญหา ระบบแสดงข้อความที่ช่วยให้แก้ไข/ดำเนินการต่อได้' },

  { section: 'ด้านความปลอดภัยของข้อมูล', q_type: 'rating', question_text: 'มั่นใจว่าข้อมูลส่วนบุคคลและข้อมูลทางวิชาการได้รับการคุ้มครองอย่างเหมาะสม' },
  { section: 'ด้านความปลอดภัยของข้อมูล', q_type: 'rating', question_text: 'การกำหนดสิทธิ์การเข้าถึงตามบทบาท ช่วยให้ข้อมูลปลอดภัยและเหมาะสม' },
  { section: 'ด้านความปลอดภัยของข้อมูล', q_type: 'rating', question_text: 'มั่นใจในความปลอดภัยของการเข้าสู่ระบบและการจัดการรหัสผ่าน' },

  { section: 'ความพึงพอใจในภาพรวม', q_type: 'rating', question_text: 'โดยภาพรวม ท่านพึงพอใจต่อระบบ EMS-BCNB' },
  { section: 'ความพึงพอใจในภาพรวม', q_type: 'rating', question_text: 'ระบบช่วยให้การทำงาน/การเข้าถึงข้อมูลสะดวกและมีประสิทธิภาพมากขึ้น' },
  { section: 'ความพึงพอใจในภาพรวม', q_type: 'rating', question_text: 'ท่านจะแนะนำให้ผู้อื่นใช้งานระบบนี้' },

  { section: 'ข้อเสนอแนะเพิ่มเติม', q_type: 'text', question_text: 'สิ่งที่ท่านชอบหรือประทับใจมากที่สุดในระบบ' },
  { section: 'ข้อเสนอแนะเพิ่มเติม', q_type: 'text', question_text: 'ปัญหาหรืออุปสรรคที่พบระหว่างการใช้งาน' },
  { section: 'ข้อเสนอแนะเพิ่มเติม', q_type: 'text', question_text: 'ฟังก์ชันหรือสิ่งที่อยากให้เพิ่มเติม/ปรับปรุง' }
];

// ---------- ตัวช่วยอ่านข้อมูล ----------
function surveyEsc(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;'); }
function surveyIsActive(q) { const v = String(q && q.active).trim().toLowerCase(); return v === '1' || v === 'true' || v === 'ใช่' || v === 'yes'; }
// ตัวเลือกของคำถามชนิด "ตัวเลือก" — เก็บในชีตคั่นด้วย | หรือขึ้นบรรทัดใหม่
function surveyParseOptions(s) { return String(s == null ? '' : s).split(/\r?\n|\|/).map(x => x.trim()).filter(Boolean); }
function surveyTypeLabel(q) { return q && q.q_type === 'text' ? 'ข้อความ' : (q && q.q_type === 'choice' ? 'ตัวเลือก (' + surveyParseOptions(q.options).length + ')' : 'มาตรวัด 1-5'); }
function surveyToggleOptionsField() { const t = (document.getElementById('surveyQType') || {}).value; const w = document.getElementById('surveyQOptionsWrap'); if (w) w.classList.toggle('hidden', t !== 'choice'); }
// บทบาทผู้ตอบที่ผูกกับคำถามได้ (ไม่รวม admin ซึ่งเป็นผู้จัดการ)
const SURVEY_EVAL_ROLES = ['academic', 'registrar', 'deptHead', 'executive', 'teacher', 'classTeacher', 'student'];

// สีประจำแต่ละด้าน (section) ของแบบประเมิน — วนซ้ำตามลำดับด้าน
const SURVEY_SECTION_COLORS = [
  { border: 'border-blue-200', head: 'text-blue-700', dot: 'bg-blue-500' },
  { border: 'border-emerald-200', head: 'text-emerald-700', dot: 'bg-emerald-500' },
  { border: 'border-purple-200', head: 'text-purple-700', dot: 'bg-purple-500' },
  { border: 'border-amber-200', head: 'text-amber-700', dot: 'bg-amber-500' },
  { border: 'border-rose-200', head: 'text-rose-700', dot: 'bg-rose-500' },
  { border: 'border-cyan-200', head: 'text-cyan-700', dot: 'bg-cyan-500' },
  { border: 'border-indigo-200', head: 'text-indigo-700', dot: 'bg-indigo-500' },
  { border: 'border-teal-200', head: 'text-teal-700', dot: 'bg-teal-500' }
];
function surveySectionColor(i) { return SURVEY_SECTION_COLORS[((i % SURVEY_SECTION_COLORS.length) + SURVEY_SECTION_COLORS.length) % SURVEY_SECTION_COLORS.length]; }

// กลุ่มบทบาทสำหรับสรุปผล: บุคลากร vs นักศึกษา
const SURVEY_ROLE_GROUPS = [
  { key: 'student', label: 'นักศึกษา', full: 'นักศึกษา', roles: ['student'], color: 'bg-emerald-500' },
  { key: 'staff', label: 'บุคลากร', full: 'บุคลากร (อาจารย์ · ประธานสาขา · อ.ประจำชั้น · งานทะเบียน · งานวิชาการ · ผู้บริหาร)', roles: ['teacher', 'deptHead', 'classTeacher', 'registrar', 'academic', 'executive'], color: 'bg-primary' }
];
function surveyParseRoles(s) { return String(s == null ? '' : s).split(/[,|]/).map(x => x.trim()).filter(Boolean); }
// คำถามนี้ใช้กับบทบาทใดบ้าง — ว่าง = ทุกบทบาท
function surveyQuestionAppliesTo(q, role) { const rs = surveyParseRoles(q && q.roles); return rs.length === 0 || rs.indexOf(role) !== -1; }
function surveyRolesLabel(q) { const rs = surveyParseRoles(q && q.roles); if (!rs.length || rs.length === SURVEY_EVAL_ROLES.length) return 'ทุกบทบาท'; return rs.map(r => SURVEY_ROLE_LABEL[r] || r).join(', '); }
function surveyQuestionsForRole(year, role, onlyActive) { return surveyQuestionsForYear(year, onlyActive).filter(q => surveyQuestionAppliesTo(q, role)); }
// บทบาทที่คำถามนี้มีผลจริง (ว่าง = ทุกบทบาท)
function surveyEffectiveRoles(q) { const rs = surveyParseRoles(q && q.roles); return rs.length ? rs : SURVEY_EVAL_ROLES.slice(); }
// บทบาทที่กำลังทำงานอยู่ในหน้าจัดการคำถาม (active role)
function surveyActiveManageRole() { return norm(APP.filters._surveyManageRole) || SURVEY_EVAL_ROLES[0]; }
// คำถามนี้เป็นของ "เฉพาะบทบาทนี้บทบาทเดียว" หรือไม่
function surveyIsExclusiveTo(q, role) { const eff = surveyEffectiveRoles(q); return eff.length === 1 && eff[0] === role; }
// แก้ไข/ลบคำถาม "เฉพาะบทบาทที่เลือก" โดยไม่กระทบบทบาทอื่น
//   newFields = null  → เอาคำถามนี้ออกจากบทบาทนี้ (ถ้าผูกบทบาทเดียว = ลบจริง)
//   newFields = {...} → ใช้ค่าที่แก้กับบทบาทนี้เท่านั้น
// ถ้าคำถามใช้ร่วมหลายบทบาท ระบบจะแตกเป็นสำเนาเฉพาะบทบาทนี้ และคงข้อเดิมไว้ให้บทบาทอื่น
async function surveyApplyToRole(q, role, newFields) {
  if (surveyIsExclusiveTo(q, role)) {
    if (newFields === null) return GSheetDB.delete(q);
    return GSheetDB.update({ ...q, ...newFields, roles: role });
  }
  const remaining = surveyEffectiveRoles(q).filter(r => r !== role);
  const r1 = await GSheetDB.update({ ...q, roles: remaining.join(',') });
  if (newFields === null) return r1;
  const pick = (k) => (newFields[k] != null ? newFields[k] : q[k]);
  return GSheetDB.create({
    type: 'survey_question', q_id: 'Q' + Date.now(), academic_year: q.academic_year,
    section: pick('section'), q_order: pick('q_order'), question_text: pick('question_text'),
    q_type: pick('q_type'), options: pick('options'), active: pick('active'), roles: role
  });
}
function surveyConfigs() { return getDataByType('survey_config'); }
function surveyConfigForYear(y) { return surveyConfigs().find(c => norm(c.academic_year) === norm(y)) || null; }
function surveyQuestionsAll() { return getDataByType('survey_question'); }
function surveyQuestionsForYear(y, onlyActive) {
  let qs = surveyQuestionsAll().filter(q => norm(q.academic_year) === norm(y));
  if (onlyActive) qs = qs.filter(surveyIsActive);
  return qs.sort((a, b) => (Number(a.q_order) || 0) - (Number(b.q_order) || 0));
}
function surveyResponsesForYear(y) { return getDataByType('survey_response').filter(r => norm(r.academic_year) === norm(y)); }
function surveyRespondentKey() {
  const u = APP.currentUser; if (!u) return '';
  if (u.role === 'student') return 'STU:' + ((u.data && u.data.student_id) || u.name || '');
  return u.role + ':' + (u.email || (u.data && u.data.email) || u.name || '');
}
function surveyMyResponseForYear(y) {
  const key = norm(surveyRespondentKey());
  return surveyResponsesForYear(y).find(r => norm(r.respondent_key) === key) || null;
}
function surveyOpenYears() {
  return surveyConfigs().filter(c => String(c.status).trim() === 'open').map(c => norm(c.academic_year)).filter(Boolean).sort().reverse();
}
// บทบาทที่เปิดรับของปีนั้น (ว่าง = เปิดทุกบทบาท) — ใช้เปิดแบบประเมินทีละบทบาท
function surveyOpenRolesFor(year) {
  const cfg = surveyConfigForYear(year);
  return cfg ? surveyParseRoles(cfg.open_roles) : [];
}
// ปีนี้เปิดรับสำหรับบทบาทนี้ไหม
function surveyIsOpenForRole(year, role) {
  const cfg = surveyConfigForYear(year);
  if (!cfg || String(cfg.status).trim() !== 'open') return false;
  const rs = surveyParseRoles(cfg.open_roles);
  return rs.length === 0 || rs.indexOf(role) !== -1;
}
function surveyAllYears() {
  const set = new Set();
  surveyConfigs().forEach(c => { if (norm(c.academic_year)) set.add(norm(c.academic_year)); });
  surveyQuestionsAll().forEach(q => { if (norm(q.academic_year)) set.add(norm(q.academic_year)); });
  getDataByType('survey_response').forEach(r => { if (norm(r.academic_year)) set.add(norm(r.academic_year)); });
  return [...set].sort().reverse();
}
function surveyCurrentThaiYear() { const d = new Date(); let y = d.getFullYear() + 543; if (d.getMonth() < 5) y -= 1; return String(y); }

// ---------- สถิติ ----------
function surveyMeanSD(vals) {
  const n = vals.length; if (!n) return { n: 0, mean: 0, sd: 0 };
  const mean = vals.reduce((a, b) => a + b, 0) / n;
  let sd = 0; if (n > 1) sd = Math.sqrt(vals.reduce((a, b) => a + (b - mean) * (b - mean), 0) / (n - 1));
  return { n, mean, sd };
}
function surveyInterpret(m) {
  if (m > 4.50) return { t: 'มากที่สุด', c: 'bg-green-100 text-green-700' };
  if (m > 3.50) return { t: 'มาก', c: 'bg-emerald-100 text-emerald-700' };
  if (m > 2.50) return { t: 'ปานกลาง', c: 'bg-amber-100 text-amber-700' };
  if (m > 1.50) return { t: 'น้อย', c: 'bg-orange-100 text-orange-700' };
  return { t: 'น้อยที่สุด', c: 'bg-red-100 text-red-700' };
}

// ======================== หน้าทำแบบประเมิน (ผู้ใช้ทั่วไป) ========================
function surveyPage() {
  const _myRole = (APP.currentUser && APP.currentUser.role) || APP.currentRole;
  const openYears = surveyOpenYears().filter(y => surveyIsOpenForRole(y, _myRole));
  let year = norm(APP.filters._surveyYear);
  if (!year && openYears.length) year = openYears[0];

  let h = `<div class="flex items-center gap-3 mb-5"><i data-lucide="clipboard-check" class="w-7 h-7 text-primary"></i>
    <div><h2 class="text-xl font-bold text-gray-800">แบบประเมินความพึงพอใจการใช้งานระบบ</h2>
    <p class="text-sm text-gray-500">ระบบบริหารจัดการงานวิชาการ (EMS-BCNB)</p></div></div>`;

  if (!openYears.length) {
    return h + `<div class="bg-white rounded-2xl p-8 border border-blue-100 text-center text-gray-500">
      <i data-lucide="inbox" class="w-10 h-10 mx-auto mb-3 text-gray-300"></i>
      <p>ขณะนี้ยังไม่มีแบบประเมินที่เปิดให้ทำ</p><p class="text-sm mt-1">โปรดติดต่อผู้ดูแลระบบ หรือกลับมาใหม่ภายหลัง</p></div>`;
  }

  // ตัวเลือกปีการศึกษา (เฉพาะปีที่เปิด)
  h += `<div class="bg-white rounded-2xl p-4 border border-blue-100 mb-4 flex flex-wrap items-center gap-3">
    <label class="text-sm font-medium text-gray-700">ปีการศึกษาที่ต้องการประเมิน:</label>
    <select onchange="APP.filters._surveyYear=this.value;renderCurrentPage()" class="border border-gray-200 rounded-xl px-3 py-2 text-sm">
      ${openYears.map(y => `<option value="${y}" ${y === year ? 'selected' : ''}>ปีการศึกษา ${y}</option>`).join('')}
    </select></div>`;

  const cfg = surveyConfigForYear(year);
  const mine = surveyMyResponseForYear(year);
  if (mine) {
    return h + `<div class="bg-white rounded-2xl p-8 border border-green-200 text-center">
      <div class="w-16 h-16 mx-auto mb-3 bg-green-100 rounded-full flex items-center justify-center"><i data-lucide="check-circle" class="w-9 h-9 text-green-600"></i></div>
      <p class="text-lg font-bold text-gray-800">ขอบคุณค่ะ ท่านได้ทำแบบประเมินของปีการศึกษา ${year} แล้ว</p>
      <p class="text-sm text-gray-500 mt-1">เมื่อ ${surveyEsc(mine.submitted_at || mine.created_at || '')}</p>
      <p class="text-sm text-gray-500 mt-3">ระบบอนุญาตให้ทำแบบประเมินได้เพียงครั้งเดียวต่อปีการศึกษา</p></div>`;
  }

  const u = APP.currentUser;
  const qs = surveyQuestionsForRole(year, u.role, true);
  if (!qs.length) {
    return h + `<div class="bg-white rounded-2xl p-8 border border-amber-200 text-center text-gray-600">
      <i data-lucide="alert-triangle" class="w-9 h-9 mx-auto mb-2 text-amber-500"></i>
      <p>แบบประเมินของปีการศึกษานี้ยังไม่มีข้อคำถามสำหรับบทบาทของท่าน โปรดติดต่อผู้ดูแลระบบ</p></div>`;
  }

  const isStudent = u.role === 'student';
  const yearLevel = isStudent && u.data ? (u.data.year_level || '') : '';

  // คำชี้แจง
  if (cfg && norm(cfg.description)) {
    h += `<div class="bg-primaryLight border border-blue-100 rounded-2xl p-4 mb-4 text-sm text-gray-700 whitespace-pre-line">${surveyEsc(cfg.description)}</div>`;
  }
  h += `<div class="bg-blue-50 border border-blue-100 rounded-xl p-3 mb-4 text-sm text-blue-800">เกณฑ์การให้คะแนน: 5 = มากที่สุด, 4 = มาก, 3 = ปานกลาง, 2 = น้อย, 1 = น้อยที่สุด</div>`;

  h += `<form id="surveyForm" data-year="${year}" onsubmit="submitSurvey(event)" class="space-y-4">`;

  // ส่วนข้อมูลผู้ตอบ
  h += `<div class="bg-white rounded-2xl p-5 border border-blue-100">
    <h3 class="font-bold text-gray-800 mb-3 flex items-center gap-2"><i data-lucide="user" class="w-5 h-5 text-primary"></i>ข้อมูลผู้ตอบ</h3>
    <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
      <div><p class="text-xs text-gray-500">บทบาท</p><p class="font-semibold text-gray-800">${SURVEY_ROLE_LABEL[u.role] || u.role}</p></div>
      ${isStudent ? `<div><p class="text-xs text-gray-500">ชั้นปี</p><p class="font-semibold text-gray-800">${surveyEsc(yearLevel) || '-'}</p></div>` : ''}
      <div>
        <label class="text-xs text-gray-500">อุปกรณ์ที่ใช้งานระบบเป็นหลัก <span class="text-red-500">*</span></label>
        <select name="device" required class="w-full border border-gray-300 rounded-xl px-3 py-2 text-sm mt-1">
          <option value="">-- เลือก --</option>${SURVEY_DEVICES.map(d => `<option value="${d}">${d}</option>`).join('')}</select>
      </div>
      <div>
        <label class="text-xs text-gray-500">ความถี่ในการใช้งานระบบ <span class="text-red-500">*</span></label>
        <select name="frequency" required class="w-full border border-gray-300 rounded-xl px-3 py-2 text-sm mt-1">
          <option value="">-- เลือก --</option>${SURVEY_FREQ.map(d => `<option value="${d}">${d}</option>`).join('')}</select>
      </div>
    </div></div>`;

  // คำถามจัดกลุ่มตาม section
  const sections = [];
  qs.forEach(q => { if (!sections.includes(q.section)) sections.push(q.section); });
  let runningNo = 0;
  sections.forEach((sec, secIdx) => {
    const secQs = qs.filter(q => q.section === sec);
    const col = surveySectionColor(secIdx);
    h += `<div class="bg-white rounded-2xl p-5 border-2 ${col.border}">
      <h3 class="font-bold ${col.head} mb-1 flex items-center gap-2"><span class="w-3 h-3 rounded-full ${col.dot}"></span>${surveyEsc(sec)}</h3><div class="space-y-4 mt-3">`;
    secQs.forEach(q => {
      runningNo++;
      if (q.q_type === 'text') {
        h += `<div><p class="text-sm text-gray-700">${runningNo}. ${surveyEsc(q.question_text)}</p>
          <textarea name="q_${q.q_id}" rows="2" class="w-full border border-gray-300 rounded-xl px-3 py-2 text-sm mt-2" placeholder="ความคิดเห็น (ไม่บังคับ)"></textarea></div>`;
      } else if (q.q_type === 'choice') {
        const opts = surveyParseOptions(q.options);
        h += `<div><p class="text-sm text-gray-700">${runningNo}. ${surveyEsc(q.question_text)} <span class="text-red-500">*</span></p>
          <div class="flex flex-col gap-2 mt-2">${opts.map(o => `
            <label class="flex items-center gap-2 px-3 py-1.5 rounded-lg border border-gray-200 cursor-pointer hover:bg-surface text-sm">
              <input type="radio" name="q_${q.q_id}" value="${surveyEsc(o)}" required class="accent-primary"> ${surveyEsc(o)}</label>`).join('') || '<span class="text-xs text-amber-500">ยังไม่ได้กำหนดตัวเลือก</span>'}</div></div>`;
      } else {
        h += `<div><p class="text-sm text-gray-700">${runningNo}. ${surveyEsc(q.question_text)} <span class="text-red-500">*</span></p>
          <div class="flex flex-wrap gap-2 mt-2">${[5, 4, 3, 2, 1].map(v => `
            <label class="flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-gray-200 cursor-pointer hover:bg-surface text-sm">
              <input type="radio" name="q_${q.q_id}" value="${v}" required class="accent-primary"> ${v} <span class="text-gray-400 text-xs">(${SURVEY_RATING_LABELS[v]})</span></label>`).join('')}</div></div>`;
      }
    });
    h += `</div></div>`;
  });

  h += `<div class="flex justify-end pt-2 pb-8"><button type="submit" class="px-6 py-3 bg-primary hover:bg-primaryDark text-white font-semibold rounded-xl shadow-md">ส่งแบบประเมิน</button></div></form>`;
  return h;
}

async function submitSurvey(ev) {
  ev.preventDefault();
  const form = ev.target;
  const year = form.dataset.year;
  const qs = surveyQuestionsForRole(year, APP.currentUser.role, true);
  const device = (form.querySelector('[name="device"]') || {}).value || '';
  const frequency = (form.querySelector('[name="frequency"]') || {}).value || '';
  if (!device) { showToast('กรุณาเลือกอุปกรณ์ที่ใช้งาน', 'error'); return; }
  if (!frequency) { showToast('กรุณาเลือกความถี่ในการใช้งาน', 'error'); return; }

  const answers = {}; let sum = 0, cnt = 0;
  for (const q of qs) {
    if (q.q_type === 'rating') {
      const sel = form.querySelector(`input[name="q_${q.q_id}"]:checked`);
      if (!sel) { showToast('กรุณาตอบคำถามให้ครบทุกข้อ', 'error'); return; }
      const v = Number(sel.value); answers[q.q_id] = v; sum += v; cnt++;
    } else if (q.q_type === 'choice') {
      const sel = form.querySelector(`input[name="q_${q.q_id}"]:checked`);
      if (!sel) { showToast('กรุณาตอบคำถามให้ครบทุกข้อ', 'error'); return; }
      answers[q.q_id] = sel.value;
    } else {
      const ta = form.querySelector(`[name="q_${q.q_id}"]`);
      const t = ta ? ta.value.trim() : '';
      if (t) answers[q.q_id] = t;
    }
  }
  const overall = cnt ? (sum / cnt) : '';
  const u = APP.currentUser;
  const payload = {
    academic_year: year, role: u.role, role_label: SURVEY_ROLE_LABEL[u.role] || u.role,
    respondent_key: surveyRespondentKey(), respondent_name: u.name || '',
    year_level: (u.role === 'student' && u.data) ? (u.data.year_level || '') : '',
    device, frequency, answers, overall_avg: overall === '' ? '' : overall.toFixed(2)
  };

  await withLoading(form.querySelector('[type="submit"]'), async () => {
    const res = await GSheetDB.surveySubmit(payload);
    if (res && res.isOk) {
      // อัปเดต cache ในเครื่องให้หน้าจอแสดง "ประเมินแล้ว" ทันที (ไม่ต้องรีโหลด)
      APP.allData.push({
        type: 'survey_response', __backendId: 'survey_response_local_' + Date.now(),
        respondent_key: payload.respondent_key, academic_year: year,
        role: payload.role, role_label: payload.role_label, respondent_name: payload.respondent_name,
        year_level: payload.year_level, device: device, frequency: frequency,
        answers_json: JSON.stringify(answers), overall_avg: payload.overall_avg,
        submitted_at: new Date().toLocaleString('th-TH')
      });
      showToast('ขอบคุณค่ะ บันทึกแบบประเมินเรียบร้อยแล้ว', 'success');
      renderCurrentPage();
    } else {
      const msg = (res && res.error) || 'บันทึกแบบประเมินไม่สำเร็จ';
      showToast(msg, 'error');
      if (msg.indexOf('ไปแล้ว') >= 0) {
        APP.allData.push({ type: 'survey_response', __backendId: 'survey_response_dup_' + Date.now(), respondent_key: payload.respondent_key, academic_year: year, submitted_at: '' });
        renderCurrentPage();
      }
    }
  });
}

// ======================== หน้าจัดการแบบประเมิน (admin) ========================
function surveyManagePage() {
  if (APP.currentRole !== 'admin') return '<p class="text-gray-500">เฉพาะผู้ดูแลระบบเท่านั้น</p>';
  const years = surveyAllYears();
  let year = norm(APP.filters._surveyManageYear);
  if (!year) year = years[0] || surveyCurrentThaiYear();
  const tab = APP._surveyManageTab || 'config';

  let h = `<div class="flex items-center gap-3 mb-5"><i data-lucide="clipboard-check" class="w-7 h-7 text-primary"></i>
    <h2 class="text-xl font-bold text-gray-800">จัดการแบบประเมินความพึงพอใจ</h2></div>`;

  // เลือก/เพิ่มปีการศึกษา
  h += `<div class="bg-white rounded-2xl p-4 border border-blue-100 mb-4 flex flex-wrap items-center gap-3">
    <label class="text-sm font-medium text-gray-700">ปีการศึกษา:</label>
    <select onchange="APP.filters._surveyManageYear=this.value;renderCurrentPage()" class="border border-gray-200 rounded-xl px-3 py-2 text-sm">
      ${(years.length ? years : [year]).map(y => `<option value="${y}" ${y === year ? 'selected' : ''}>${y}</option>`).join('')}
    </select>
    <span class="text-gray-300">|</span>
    <input id="surveyNewYear" placeholder="เพิ่มปีใหม่ เช่น ${surveyCurrentThaiYear()}" class="border border-gray-200 rounded-xl px-3 py-2 text-sm w-44">
    <button onclick="surveyGotoNewYear()" class="px-3 py-2 bg-primary text-white rounded-xl text-sm hover:bg-primaryDark">ไปยังปีนี้</button>
  </div>`;

  // แท็บ
  const tabs = [['config', 'ตั้งค่าแบบประเมิน'], ['questions', 'จัดการคำถาม'], ['results', 'สรุปผล']];
  h += `<div class="flex gap-2 mb-4 border-b border-gray-200">${tabs.map(([id, label]) => `
    <button onclick="APP._surveyManageTab='${id}';renderCurrentPage()" class="px-4 py-2 text-sm font-medium border-b-2 ${tab === id ? 'border-primary text-primary' : 'border-transparent text-gray-500 hover:text-gray-700'}">${label}</button>`).join('')}</div>`;

  if (tab === 'config') h += surveyConfigTabHTML(year);
  else if (tab === 'questions') h += surveyQuestionsTabHTML(year);
  else h += surveyResultsTabHTML(year);
  return h;
}

function surveyGotoNewYear() {
  const v = norm((document.getElementById('surveyNewYear') || {}).value || '');
  if (!v) { showToast('กรุณากรอกปีการศึกษา', 'error'); return; }
  APP.filters._surveyManageYear = v; renderCurrentPage();
}

function surveyConfigTabHTML(year) {
  const cfg = surveyConfigForYear(year);
  const status = cfg ? String(cfg.status).trim() : 'closed';
  const isOpen = status === 'open';
  const qCount = surveyQuestionsForYear(year, false).length;
  const respCount = surveyResponsesForYear(year).length;

  const openRoles = surveyOpenRolesFor(year);           // [] = ทุกบทบาท (เมื่อเปิด)
  const isRoleOpen = r => isOpen && (openRoles.length === 0 || openRoles.indexOf(r) !== -1);
  const statusText = !isOpen ? '○ ปิดรับการประเมิน'
    : (openRoles.length === 0 ? '● เปิดรับ (ทุกบทบาท)' : '● เปิดรับ: ' + openRoles.map(r => SURVEY_ROLE_LABEL[r] || r).join(', '));

  return `<div class="bg-white rounded-2xl p-5 border border-blue-100 max-w-2xl">
    <div class="flex items-center justify-between mb-4">
      <div><p class="text-sm text-gray-500">สถานะแบบประเมินปีการศึกษา ${year}</p>
        <p class="font-bold text-lg ${isOpen ? 'text-green-600' : 'text-gray-500'}">${statusText}</p></div>
      <div class="text-right text-sm text-gray-500"><p>คำถาม: <b class="text-gray-800">${qCount}</b> ข้อ</p><p>ผู้ตอบแล้ว: <b class="text-gray-800">${respCount}</b> คน</p></div>
    </div>
    <label class="block text-sm font-medium text-gray-700 mb-1">ชื่อแบบประเมิน (ไม่บังคับ)</label>
    <input id="surveyCfgTitle" value="${surveyEsc(cfg ? cfg.title : '')}" class="w-full border border-gray-300 rounded-xl px-3 py-2 text-sm mb-3" placeholder="แบบประเมินความพึงพอใจการใช้งานระบบ EMS-BCNB">
    <label class="block text-sm font-medium text-gray-700 mb-1">คำชี้แจง (แสดงให้ผู้ตอบเห็นด้านบนแบบประเมิน)</label>
    <textarea id="surveyCfgDesc" rows="3" class="w-full border border-gray-300 rounded-xl px-3 py-2 text-sm mb-4" placeholder="คำชี้แจงการทำแบบประเมิน...">${surveyEsc(cfg ? cfg.description : '')}</textarea>

    <div class="p-3 bg-green-50 rounded-xl border border-green-100 mb-4">
      <div class="flex items-center justify-between mb-2">
        <label class="text-sm font-semibold text-green-800">เปิดรับการประเมินเฉพาะบทบาทที่เลือก</label>
        <div class="flex gap-2 text-xs">
          <button type="button" onclick="surveyToggleAllRoles(true)" class="px-2 py-1 rounded-lg border border-green-200 text-green-700 hover:bg-green-100">เลือกทุกบทบาท</button>
          <button type="button" onclick="surveyToggleAllRoles(false)" class="px-2 py-1 rounded-lg border border-gray-200 text-gray-600 hover:bg-gray-100">ไม่เลือกเลย</button>
        </div>
      </div>
      <div class="grid grid-cols-2 sm:grid-cols-3 gap-1.5">
        ${SURVEY_EVAL_ROLES.map(r => `<label class="flex items-center gap-2 text-sm text-gray-700 bg-white rounded-lg px-2 py-1.5 border border-green-100"><input type="checkbox" class="survey-open-role accent-green-600" value="${r}" ${isRoleOpen(r) ? 'checked' : ''}> ${SURVEY_ROLE_LABEL[r] || r}</label>`).join('')}
      </div>
      <p class="text-[11px] text-gray-500 mt-2">ติ๊กบทบาทที่ต้องการให้ทำแบบประเมิน แล้วกด "บันทึกการตั้งค่า" — บทบาทที่ไม่ติ๊กจะยังทำไม่ได้ (ไม่ติ๊กเลย = ปิดรับทั้งหมด)</p>
    </div>

    <div class="flex flex-wrap gap-2">
      <button onclick="surveyPreview('${year}')" class="px-4 py-2 bg-white border border-primary text-primary rounded-xl text-sm hover:bg-primaryLight flex items-center gap-1"><i data-lucide="eye" class="w-4 h-4"></i>แสดงตัวอย่างแบบประเมิน</button>
      <button id="surveyCfgSaveBtn" onclick="surveySaveOpenRoles('${year}')" class="px-4 py-2 bg-green-600 text-white rounded-xl text-sm hover:bg-green-700 flex items-center gap-1"><i data-lucide="save" class="w-4 h-4"></i>บันทึกการตั้งค่า (ชื่อ/คำชี้แจง/การเปิดรับ)</button>
      ${qCount === 0 ? `<button onclick="surveyCreateDefaultQuestions('${year}')" class="px-4 py-2 bg-primary text-white rounded-xl text-sm hover:bg-primaryDark">สร้างชุดคำถามเริ่มต้น (ใช้ร่วมทุกบทบาท)</button>` : ''}
    </div>
  </div>`;
}

function surveyToggleAllRoles(on) {
  document.querySelectorAll('.survey-open-role').forEach(el => { el.checked = !!on; });
}

// ---------- ตัวอย่างแบบประเมิน (พรีวิว ก่อนเปิดรับ) ----------
function surveyPreview(year) {
  const role = SURVEY_EVAL_ROLES[0];
  showModal('ตัวอย่างแบบประเมิน · ปีการศึกษา ' + year, `
    <div class="mb-3">
      <label class="text-xs text-gray-500">ดูตัวอย่างสำหรับบทบาท</label>
      <select id="surveyPreviewRole" onchange="surveyPreviewSetRole('${year}')" class="w-full border rounded-xl px-3 py-2 text-sm mt-1">
        ${SURVEY_EVAL_ROLES.map(r => `<option value="${r}" ${r === role ? 'selected' : ''}>${SURVEY_ROLE_LABEL[r] || r}</option>`).join('')}
      </select>
    </div>
    <div id="surveyPreviewInner" class="max-h-[62vh] overflow-y-auto pr-1">${surveyPreviewInner(year, role)}</div>
  `, null, 'max-w-2xl');
}
function surveyPreviewSetRole(year) {
  const role = (document.getElementById('surveyPreviewRole') || {}).value || SURVEY_EVAL_ROLES[0];
  const box = document.getElementById('surveyPreviewInner');
  if (box) { box.innerHTML = surveyPreviewInner(year, role); if (window.lucide) lucide.createIcons(); }
}
function surveyPreviewInner(year, role) {
  const cfg = surveyConfigForYear(year);
  const qs = surveyQuestionsForRole(year, role, true);
  let h = '';
  if (cfg && norm(cfg.description)) h += `<div class="bg-primaryLight border border-blue-100 rounded-2xl p-3 mb-3 text-sm text-gray-700 whitespace-pre-line">${surveyEsc(cfg.description)}</div>`;
  h += `<div class="bg-blue-50 border border-blue-100 rounded-xl p-2.5 mb-3 text-xs text-blue-800">เกณฑ์การให้คะแนน: 5 = มากที่สุด, 4 = มาก, 3 = ปานกลาง, 2 = น้อย, 1 = น้อยที่สุด</div>`;
  if (!qs.length) return h + '<p class="text-center text-amber-600 py-6 text-sm">ยังไม่มีคำถามสำหรับบทบาทนี้</p>';
  const sections = [];
  qs.forEach(q => { if (!sections.includes(q.section)) sections.push(q.section); });
  let no = 0;
  h += '<fieldset disabled class="space-y-3">';
  sections.forEach((sec, secIdx) => {
    const secQs = qs.filter(q => q.section === sec);
    const col = surveySectionColor(secIdx);
    h += `<div class="bg-white rounded-xl p-4 border-2 ${col.border}"><h3 class="font-bold ${col.head} mb-2 text-sm flex items-center gap-2"><span class="w-2.5 h-2.5 rounded-full ${col.dot}"></span>${surveyEsc(sec)}</h3><div class="space-y-3">`;
    secQs.forEach(q => {
      no++;
      if (q.q_type === 'text') {
        h += `<div><p class="text-sm text-gray-700">${no}. ${surveyEsc(q.question_text)}</p><textarea rows="2" class="w-full border border-gray-200 rounded-xl px-3 py-2 text-sm mt-1 bg-gray-50" placeholder="ความคิดเห็น (ไม่บังคับ)"></textarea></div>`;
      } else if (q.q_type === 'choice') {
        const opts = surveyParseOptions(q.options);
        h += `<div><p class="text-sm text-gray-700">${no}. ${surveyEsc(q.question_text)} <span class="text-red-500">*</span></p><div class="flex flex-col gap-1.5 mt-1">${opts.map(o => `<label class="flex items-center gap-2 px-3 py-1.5 rounded-lg border border-gray-200 text-sm"><input type="radio" class="accent-primary"> ${surveyEsc(o)}</label>`).join('') || '<span class="text-xs text-amber-500">ยังไม่ได้กำหนดตัวเลือก</span>'}</div></div>`;
      } else {
        h += `<div><p class="text-sm text-gray-700">${no}. ${surveyEsc(q.question_text)} <span class="text-red-500">*</span></p><div class="flex flex-wrap gap-2 mt-1">${[5, 4, 3, 2, 1].map(v => `<label class="flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-gray-200 text-sm"><input type="radio" class="accent-primary"> ${v} <span class="text-gray-400 text-xs">(${SURVEY_RATING_LABELS[v]})</span></label>`).join('')}</div></div>`;
      }
    });
    h += `</div></div>`;
  });
  h += '</fieldset>';
  h += `<p class="text-xs text-gray-400 mt-2 text-center">ตัวอย่างเท่านั้น — ไม่บันทึกคำตอบ · มี ${qs.length} ข้อสำหรับบทบาทนี้</p>`;
  return h;
}

async function surveySaveOpenRoles(year) {
  const roles = [...document.querySelectorAll('.survey-open-role:checked')].map(el => el.value);
  const title = (document.getElementById('surveyCfgTitle') || {}).value || '';
  const desc = (document.getElementById('surveyCfgDesc') || {}).value || '';
  const existing = surveyConfigForYear(year);
  const wasOpen = existing && String(existing.status).trim() === 'open';
  const status = roles.length ? 'open' : 'closed';
  // ติ๊กครบทุกบทบาท → เก็บว่าง (= ทุกบทบาท) ไม่งั้นเก็บรายการบทบาท
  const openRolesStr = (roles.length === SURVEY_EVAL_ROLES.length) ? '' : roles.join(',');
  const now = new Date().toISOString();
  await withLoading(document.getElementById('surveyCfgSaveBtn'), async () => {
    const payload = { status, open_roles: openRolesStr, title, description: desc, updated_at: now };
    let res;
    if (existing) res = await GSheetDB.update({ ...existing, ...payload });
    else res = await GSheetDB.create({ type: 'survey_config', academic_year: year, created_at: now, ...payload });
    if (res && res.isOk) {
      if (status === 'open' && !wasOpen) await surveyCreateOpenAnnouncement(year, title, openRolesStr);
      const openLabel = !roles.length ? 'ปิดรับทั้งหมด' : (openRolesStr === '' ? 'เปิดรับทุกบทบาท' : 'เปิดรับ: ' + roles.map(r => SURVEY_ROLE_LABEL[r] || r).join(', '));
      showToast('บันทึกแล้ว — ' + openLabel, 'success');
      if (typeof updateNotifBadge === 'function') updateNotifBadge();
      renderCurrentPage();
    } else showToast((res && res.error) || 'บันทึกไม่สำเร็จ', 'error');
  });
}

async function surveySaveConfig(year, status) {
  const title = (document.getElementById('surveyCfgTitle') || {}).value || '';
  const desc = (document.getElementById('surveyCfgDesc') || {}).value || '';
  const existing = surveyConfigForYear(year);
  const wasOpen = existing && String(existing.status).trim() === 'open';
  const now = new Date().toISOString();
  const btn = document.getElementById('surveyCfgSaveBtn');
  await withLoading(btn, async () => {
    let res;
    if (existing) res = await GSheetDB.update({ ...existing, status, title, description: desc, updated_at: now });
    else res = await GSheetDB.create({ type: 'survey_config', academic_year: year, status, title, description: desc, updated_at: now });
    if (res && res.isOk) {
      // เปลี่ยนเป็น "เปิดรับ" ครั้งใหม่ → สร้างประกาศแจ้งเตือนในระบบ (กระดิ่ง + การ์ดหน้าหลัก) โดยไม่ส่ง LINE
      if (status === 'open' && !wasOpen) await surveyCreateOpenAnnouncement(year, title);
      showToast('บันทึกการตั้งค่าแล้ว', 'success');
      if (typeof updateNotifBadge === 'function') updateNotifBadge();
      renderCurrentPage();
    }
    else showToast((res && res.error) || 'บันทึกไม่สำเร็จ', 'error');
  });
}

// สร้างประกาศแจ้งเตือน "เปิดให้ทำแบบประเมิน" เข้าระบบ (กระดิ่ง + การ์ดหน้าหลัก)
// ตั้ง line_sent ไว้ล่วงหน้า เพื่อกันไม่ให้ตัวแจ้งเตือน LINE (line-announcement-notify.gs) หยิบไปส่ง
async function surveyCreateOpenAnnouncement(year, title, roles) {
  try {
    const pad = n => String(n).padStart(2, '0');
    const d = new Date();
    const dateStr = `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
    const t = (title && norm(title)) || 'แบบประเมินความพึงพอใจการใช้งานระบบ EMS-BCNB';
    await GSheetDB.create({
      type: 'announcement',
      announcement_title: 'เปิดให้ทำแบบประเมินความพึงพอใจ ปีการศึกษา ' + year,
      announcement_content: 'ขอเชิญผู้ใช้งานร่วมทำ "' + t + '" ประจำปีการศึกษา ' + year + ' ได้ที่เมนู "แบบประเมินความพึงพอใจ" (ทำได้ครั้งเดียวต่อปีการศึกษา)',
      announcement_date: dateStr,
      roles: roles || '',
      line_sent: 'ไม่ส่ง LINE (แจ้งเฉพาะในระบบ)',
      line_notify: ''
    });
  } catch (e) { /* ไม่ให้การสร้างประกาศที่ผิดพลาดทำให้การเปิดแบบประเมินล้มเหลว */ }
}

async function surveyCreateDefaultQuestions(year, role) {
  const label = role ? ('บทบาท "' + (SURVEY_ROLE_LABEL[role] || role) + '"') : 'ทุกบทบาท (ใช้ร่วมกัน)';
  if (!confirm('สร้างชุดคำถามเริ่มต้น (' + SURVEY_DEFAULT_QUESTIONS.length + ' ข้อ) สำหรับ ' + label + ' ปีการศึกษา ' + year + ' หรือไม่?')) return;
  const base = Date.now();
  const objs = SURVEY_DEFAULT_QUESTIONS.map((q, i) => ({
    type: 'survey_question', q_id: 'Q' + base + '_' + i, academic_year: year,
    section: q.section, q_order: (i + 1) * 10, question_text: q.question_text, q_type: q.q_type, roles: role || '', active: '1'
  }));
  showToast('กำลังสร้างชุดคำถาม...', 'loading');
  const res = await GSheetDB.createMany(objs);
  hideLoadingToast();
  if (res && res.isOk) { showToast('สร้างชุดคำถามเรียบร้อย ' + (res.ok || objs.length) + ' ข้อ', 'success'); APP._surveyManageTab = 'questions'; renderCurrentPage(); }
  else showToast('สร้างคำถามไม่สำเร็จ' + (res && res.fail ? ' (สำเร็จ ' + res.ok + ' / ล้มเหลว ' + res.fail + ')' : ''), 'error');
}

function surveyQuestionsTabHTML(year) {
  const role = surveyActiveManageRole();
  const roleName = SURVEY_ROLE_LABEL[role] || role;
  const qs = surveyQuestionsForRole(year, role, false);
  let h = `<div class="flex flex-wrap items-center justify-between gap-2 mb-2">
    <div class="flex items-center gap-2 flex-wrap">
      <label class="text-sm font-medium text-gray-700">ชุดคำถามของบทบาท:</label>
      <select onchange="APP.filters._surveyManageRole=this.value;renderCurrentPage()" class="border border-gray-200 rounded-lg px-2 py-1.5 text-sm">
        ${SURVEY_EVAL_ROLES.map(r => `<option value="${r}" ${role === r ? 'selected' : ''}>${SURVEY_ROLE_LABEL[r] || r}</option>`).join('')}
      </select>
      <span class="text-sm text-gray-500">${qs.length} ข้อ</span>
    </div>
    <div class="flex gap-2 flex-wrap">
      <button onclick="surveyPickBaseQuestionsModal('${year}')" class="px-4 py-2 bg-white border border-primary text-primary rounded-xl text-sm hover:bg-primaryLight flex items-center gap-1"><i data-lucide="list-checks" class="w-4 h-4"></i>เลือกจากคำถามพื้นฐาน</button>
      <button onclick="surveyAddQuestionModal('${year}')" class="px-4 py-2 bg-primary text-white rounded-xl text-sm hover:bg-primaryDark flex items-center gap-1"><i data-lucide="plus" class="w-4 h-4"></i>เพิ่มคำถามเอง</button>
    </div></div>
  <p class="text-xs text-gray-400 mb-3">การแก้ไข/ลบ/เปิด-ปิด จะมีผลเฉพาะบทบาท <b>"${roleName}"</b> เท่านั้น · ข้อที่ยัง "ใช้ร่วมทุกบทบาท" เมื่อแก้จะถูกแยกเป็นชุดเฉพาะบทบาทนี้ให้อัตโนมัติ (บทบาทอื่นไม่เปลี่ยน)</p>`;
  if (!qs.length) {
    return h + `<div class="bg-white rounded-2xl p-8 border border-blue-100 text-center text-gray-500">ยังไม่มีคำถามสำหรับบทบาท "${roleName}"
      <div class="mt-3 flex gap-2 justify-center flex-wrap">
        <button onclick="surveyCreateDefaultQuestions('${year}','${role}')" class="px-4 py-2 bg-primary text-white rounded-xl text-sm hover:bg-primaryDark">สร้างชุดคำถามเริ่มต้นทั้งหมด</button>
        <button onclick="surveyPickBaseQuestionsModal('${year}')" class="px-4 py-2 bg-white border border-primary text-primary rounded-xl text-sm hover:bg-primaryLight">เลือกจากคำถามพื้นฐาน</button>
      </div></div>`;
  }

  const sections = [];
  qs.forEach(q => { if (!sections.includes(q.section)) sections.push(q.section); });
  sections.forEach(sec => {
    h += `<div class="bg-white rounded-2xl p-4 border border-blue-100 mb-3"><h4 class="font-bold text-gray-700 mb-2 text-sm">${surveyEsc(sec)}</h4><div class="space-y-2">`;
    qs.filter(q => q.section === sec).forEach(q => {
      const active = surveyIsActive(q);
      const shared = !surveyIsExclusiveTo(q, role);
      h += `<div class="flex items-start gap-2 p-2 rounded-lg ${active ? '' : 'opacity-50'} hover:bg-gray-50">
        <span class="text-xs text-gray-400 mt-1 w-8">#${surveyEsc(q.q_order)}</span>
        <div class="flex-1 min-w-0"><p class="text-sm text-gray-800">${surveyEsc(q.question_text)}</p>
          <span class="text-xs ${q.q_type === 'text' ? 'text-purple-500' : (q.q_type === 'choice' ? 'text-indigo-500' : 'text-blue-500')}">${surveyTypeLabel(q)}${active ? '' : ' · ปิดใช้งาน'}</span>
          ${shared ? `<span class="text-xs text-amber-600"> · ใช้ร่วมทุกบทบาท</span>` : `<span class="text-xs text-teal-600"> · เฉพาะบทบาทนี้</span>`}
          ${q.q_type === 'choice' ? `<span class="text-xs text-gray-400 block truncate">ตัวเลือก: ${surveyEsc(surveyParseOptions(q.options).join(' / '))}</span>` : ''}</div>
        <button onclick="surveyToggleQuestionActive('${q.q_id}')" title="${active ? 'ปิดใช้งาน' : 'เปิดใช้งาน'}" class="p-1.5 rounded hover:bg-gray-100 text-gray-500"><i data-lucide="${active ? 'eye' : 'eye-off'}" class="w-4 h-4"></i></button>
        <button onclick="surveyEditQuestionModal('${q.q_id}')" class="p-1.5 rounded hover:bg-blue-50 text-blue-600"><i data-lucide="pencil" class="w-4 h-4"></i></button>
        <button onclick="surveyDeleteQuestion('${q.q_id}')" class="p-1.5 rounded hover:bg-red-50 text-red-500" title="${shared ? 'นำออกจากบทบาทนี้' : 'ลบ'}"><i data-lucide="trash-2" class="w-4 h-4"></i></button>
      </div>`;
    });
    h += `</div></div>`;
  });
  return h;
}

function surveyQuestionFormHTML(year, q, role) {
  const roleName = SURVEY_ROLE_LABEL[role] || role;
  const willSplit = q && !surveyIsExclusiveTo(q, role);
  return `<div class="space-y-3">
    <div class="bg-blue-50 border border-blue-100 rounded-xl px-3 py-2 text-xs text-blue-800">คำถามนี้สำหรับบทบาท: <b>${roleName}</b>${willSplit ? ' · เดิมใช้ร่วมหลายบทบาท — เมื่อบันทึกจะถูกแยกเป็นชุดเฉพาะบทบาทนี้ (บทบาทอื่นไม่เปลี่ยน)' : ''}</div>
    <div><label class="text-sm font-medium text-gray-700">หมวด (section)</label>
      <input id="surveyQSection" value="${surveyEsc(q ? q.section : '')}" class="w-full border border-gray-300 rounded-xl px-3 py-2 text-sm mt-1" placeholder="เช่น ด้านความง่ายในการใช้งาน"></div>
    <div><label class="text-sm font-medium text-gray-700">ข้อคำถาม</label>
      <textarea id="surveyQText" rows="2" class="w-full border border-gray-300 rounded-xl px-3 py-2 text-sm mt-1">${surveyEsc(q ? q.question_text : '')}</textarea></div>
    <div class="grid grid-cols-2 gap-3">
      <div><label class="text-sm font-medium text-gray-700">ชนิดคำถาม</label>
        <select id="surveyQType" onchange="surveyToggleOptionsField()" class="w-full border border-gray-300 rounded-xl px-3 py-2 text-sm mt-1">
          <option value="rating" ${(!q || q.q_type === 'rating' || (q.q_type !== 'text' && q.q_type !== 'choice')) ? 'selected' : ''}>มาตรวัด 1-5 (คิดค่าเฉลี่ย)</option>
          <option value="choice" ${q && q.q_type === 'choice' ? 'selected' : ''}>ตัวเลือก (กำหนดเอง)</option>
          <option value="text" ${q && q.q_type === 'text' ? 'selected' : ''}>ข้อความ (ข้อเสนอแนะ)</option></select></div>
      <div><label class="text-sm font-medium text-gray-700">ลำดับ (q_order)</label>
        <input id="surveyQOrder" type="number" value="${surveyEsc(q ? q.q_order : (surveyQuestionsForYear(year, false).length + 1) * 10)}" class="w-full border border-gray-300 rounded-xl px-3 py-2 text-sm mt-1"></div>
    </div>
    <div id="surveyQOptionsWrap" class="${q && q.q_type === 'choice' ? '' : 'hidden'}">
      <label class="text-sm font-medium text-gray-700">ตัวเลือก <span class="text-xs text-gray-400">(หนึ่งบรรทัดต่อหนึ่งตัวเลือก — ใช้กับชนิด "ตัวเลือก" เท่านั้น)</span></label>
      <textarea id="surveyQOptions" rows="4" class="w-full border border-gray-300 rounded-xl px-3 py-2 text-sm mt-1" placeholder="เช่น&#10;พึงพอใจมาก&#10;พึงพอใจ&#10;ควรปรับปรุง">${q ? surveyEsc(surveyParseOptions(q.options).join('\n')) : ''}</textarea>
    </div>
    <label class="flex items-center gap-2 text-sm text-gray-700"><input id="surveyQActive" type="checkbox" ${!q || surveyIsActive(q) ? 'checked' : ''} class="accent-primary">เปิดใช้งานคำถามนี้</label>
  </div>`;
}

function surveyAddQuestionModal(year) {
  const role = surveyActiveManageRole();
  showModal('เพิ่มคำถาม (' + (SURVEY_ROLE_LABEL[role] || role) + ') — ปีการศึกษา ' + year, surveyQuestionFormHTML(year, null, role), () => surveySaveQuestion(year, null), 'max-w-xl');
}
function surveyEditQuestionModal(qid) {
  const q = surveyQuestionsAll().find(x => x.q_id === qid);
  if (!q) { showToast('ไม่พบคำถาม', 'error'); return; }
  const role = surveyActiveManageRole();
  showModal('แก้ไขคำถาม (' + (SURVEY_ROLE_LABEL[role] || role) + ')', surveyQuestionFormHTML(q.academic_year, q, role), () => surveySaveQuestion(q.academic_year, q), 'max-w-xl');
}
async function surveySaveQuestion(year, q) {
  const gv = id => { const e = document.getElementById(id); return e ? e.value : ''; };
  const section = gv('surveyQSection').trim();
  const text = gv('surveyQText').trim();
  const qtype = gv('surveyQType') || 'rating';
  const order = gv('surveyQOrder') || '0';
  const active = (document.getElementById('surveyQActive') || {}).checked ? '1' : '0';
  const optList = surveyParseOptions(gv('surveyQOptions'));
  const options = qtype === 'choice' ? optList.join('|') : '';
  const role = surveyActiveManageRole();
  if (!text) { showToast('กรุณากรอกข้อคำถาม', 'error'); return; }
  if (qtype === 'choice' && optList.length < 2) { showToast('คำถามชนิดตัวเลือกต้องมีอย่างน้อย 2 ตัวเลือก', 'error'); return; }
  let res;
  if (q) res = await surveyApplyToRole(q, role, { section, question_text: text, q_type: qtype, options, q_order: order, active });
  else res = await GSheetDB.create({ type: 'survey_question', q_id: 'Q' + Date.now(), academic_year: year, section, question_text: text, q_type: qtype, options, roles: role, q_order: order, active });
  if (res && res.isOk) { closeModal(); showToast('บันทึกคำถามแล้ว', 'success'); renderCurrentPage(); }
  else showToast((res && res.error) || 'บันทึกไม่สำเร็จ', 'error');
}
async function surveyToggleQuestionActive(qid) {
  const q = surveyQuestionsAll().find(x => x.q_id === qid); if (!q) return;
  const role = surveyActiveManageRole();
  const res = await surveyApplyToRole(q, role, { active: surveyIsActive(q) ? '0' : '1' });
  if (res && res.isOk) { showToast('อัปเดตแล้ว', 'success'); renderCurrentPage(); }
  else showToast((res && res.error) || 'อัปเดตไม่สำเร็จ', 'error');
}
async function surveyDeleteQuestion(qid) {
  const q = surveyQuestionsAll().find(x => x.q_id === qid); if (!q) return;
  const role = surveyActiveManageRole();
  const shared = !surveyIsExclusiveTo(q, role);
  const msg = shared
    ? ('นำคำถามนี้ออกจากบทบาท "' + (SURVEY_ROLE_LABEL[role] || role) + '"?\n(บทบาทอื่นยังคงเห็นคำถามนี้)\n\n')
    : ('ลบคำถามนี้?\n\n');
  if (!confirm(msg + (q.question_text || ''))) return;
  const res = await surveyApplyToRole(q, role, null);
  if (res && res.isOk) { showToast(shared ? 'นำออกจากบทบาทนี้แล้ว' : 'ลบคำถามแล้ว', 'success'); renderCurrentPage(); }
  else showToast((res && res.error) || 'ลบไม่สำเร็จ', 'error');
}

// เลือกจากคลังคำถามพื้นฐาน (SURVEY_DEFAULT_QUESTIONS) มาเพิ่มให้บทบาทที่กำลังทำงาน
function surveyPickBaseQuestionsModal(year) {
  const role = surveyActiveManageRole();
  const roleName = SURVEY_ROLE_LABEL[role] || role;
  const existingTexts = {}; surveyQuestionsForRole(year, role, false).forEach(q => { existingTexts[norm(q.question_text)] = true; });
  const sections = []; SURVEY_DEFAULT_QUESTIONS.forEach(q => { if (!sections.includes(q.section)) sections.push(q.section); });
  let body = `<div class="flex items-center justify-between mb-2 gap-2">
    <p class="text-sm text-gray-600">เลือกข้อคำถามเพื่อเพิ่มให้บทบาท <b>"${roleName}"</b></p>
    <button type="button" onclick="document.querySelectorAll('.survey-base-q:not(:disabled)').forEach(function(c){c.checked=true})" class="text-xs text-primary hover:underline whitespace-nowrap">เลือกทั้งหมด</button>
  </div><div class="max-h-[55vh] overflow-auto space-y-3 pr-1">`;
  sections.forEach(sec => {
    body += `<div><p class="font-semibold text-sm text-gray-700 mb-1">${surveyEsc(sec)}</p><div class="space-y-1">`;
    SURVEY_DEFAULT_QUESTIONS.forEach((q, idx) => {
      if (q.section !== sec) return;
      const added = !!existingTexts[norm(q.question_text)];
      const typeTag = q.q_type === 'text' ? 'ข้อความ' : (q.q_type === 'choice' ? 'ตัวเลือก' : 'มาตรวัด 1-5');
      body += `<label class="flex items-start gap-2 text-sm p-1.5 rounded-lg ${added ? 'opacity-50' : 'hover:bg-gray-50 cursor-pointer'}">
        <input type="checkbox" class="survey-base-q accent-primary mt-0.5" value="${idx}" ${added ? 'disabled' : ''}>
        <span class="flex-1">${surveyEsc(q.question_text)} <span class="text-xs text-gray-400">(${typeTag}${added ? ' · เพิ่มแล้ว' : ''})</span></span></label>`;
    });
    body += `</div></div>`;
  });
  body += `</div>`;
  showModal('เลือกจากคำถามพื้นฐาน — บทบาท ' + roleName, body, () => surveyAddBaseQuestions(year), 'max-w-2xl');
}

async function surveyAddBaseQuestions(year) {
  const role = surveyActiveManageRole();
  const idxs = Array.prototype.map.call(document.querySelectorAll('.survey-base-q:checked'), el => Number(el.value));
  if (!idxs.length) { showToast('กรุณาเลือกอย่างน้อย 1 ข้อ', 'error'); return; }
  const existing = surveyQuestionsForRole(year, role, false);
  let order = existing.reduce((m, q) => Math.max(m, Number(q.q_order) || 0), 0);
  const base = Date.now();
  const objs = idxs.map((idx, i) => {
    const dq = SURVEY_DEFAULT_QUESTIONS[idx]; order += 10;
    return { type: 'survey_question', q_id: 'Q' + base + '_' + i, academic_year: year, section: dq.section, q_order: order, question_text: dq.question_text, q_type: dq.q_type, options: '', roles: role, active: '1' };
  });
  closeModal();
  showToast('กำลังเพิ่มคำถาม...', 'loading');
  const res = await GSheetDB.createMany(objs);
  hideLoadingToast();
  if (res && res.isOk) { showToast('เพิ่มคำถามแล้ว ' + (res.ok || objs.length) + ' ข้อ', 'success'); renderCurrentPage(); }
  else showToast('เพิ่มคำถามไม่สำเร็จ' + (res && res.fail ? ' (สำเร็จ ' + res.ok + ' / ล้มเหลว ' + res.fail + ')' : ''), 'error');
}

// ======================== สรุปผล (admin) ========================
function surveyResultsTabHTML(year) {
  const resps = surveyResponsesForYear(year);
  const qs = surveyQuestionsForYear(year, false);
  if (!resps.length) return `<div class="bg-white rounded-2xl p-8 border border-blue-100 text-center text-gray-500">ยังไม่มีผู้ตอบแบบประเมินของปีการศึกษา ${year}</div>`;

  const groups = SURVEY_ROLE_GROUPS.map(g => ({ ...g, resps: resps.filter(r => g.roles.indexOf(norm(r.role)) !== -1) }));
  const other = resps.filter(r => !SURVEY_ROLE_GROUPS.some(g => g.roles.indexOf(norm(r.role)) !== -1));

  let h = `<div class="flex justify-end mb-3"><button onclick="downloadSurveyResultsPDF('${year}')" class="flex items-center gap-2 px-4 py-2 bg-primary text-white rounded-xl hover:bg-primaryDark text-sm"><i data-lucide="download" class="w-4 h-4"></i>ดาวน์โหลด PDF</button></div>`;

  // กราฟแท่งจำนวนผู้ตอบแยกตามกลุ่ม
  const maxC = Math.max(1, ...groups.map(g => g.resps.length));
  h += `<div class="bg-white rounded-2xl p-5 border border-blue-100 mb-5">
    <h4 class="font-bold text-gray-800 mb-3 flex items-center gap-2"><i data-lucide="bar-chart-3" class="w-5 h-5 text-primary"></i>จำนวนผู้ตอบแบบประเมิน แยกตามกลุ่ม</h4>
    <div class="space-y-3">${groups.map(g => {
    const c = g.resps.length; const pct = (c / maxC * 100);
    return `<div class="flex items-center gap-3 text-sm">
        <span class="w-28 sm:w-36 shrink-0 text-gray-600">${g.label}</span>
        <div class="flex-1 bg-gray-100 rounded-lg h-6 overflow-hidden"><div class="h-6 rounded-lg ${g.color} flex items-center justify-end pr-2 text-white text-xs font-bold" style="width:${Math.max(pct, c ? 8 : 0)}%">${c ? c : ''}</div></div>
        <span class="w-10 text-right font-bold text-gray-800">${c}</span></div>`;
  }).join('')}</div>
    <p class="text-xs text-gray-400 mt-3">รวมผู้ตอบทั้งหมด ${resps.length} คน${other.length ? ' · อื่นๆ ' + other.length + ' คน' : ''}</p>
  </div>`;

  // ผลรายด้าน — แยกตามกลุ่ม (เทียบข้าง ๆ กัน)
  h += surveySectionByGroupHTML(groups, qs);

  // ผลรายข้อ — เลือกดูตามกลุ่ม
  const selKey = APP._surveyResultGroup || groups[0].key;
  const selGroup = groups.find(g => g.key === selKey) || groups[0];
  h += `<div class="mt-6 mb-3 flex flex-wrap items-center gap-2">
    <span class="text-sm font-semibold text-gray-700"><i data-lucide="list-checks" class="w-4 h-4 inline text-primary"></i> ผลรายข้อ — เลือกกลุ่มที่ต้องการดู:</span>
    ${groups.map(g => `<button onclick="APP._surveyResultGroup='${g.key}';renderCurrentPage()" class="px-4 py-1.5 rounded-xl text-sm transition ${g.key === selKey ? g.color + ' text-white shadow' : 'bg-white border border-gray-200 text-gray-600 hover:bg-surface'}">${g.label} (${g.resps.length})</button>`).join('')}
  </div>`;
  h += `<div class="mb-2"><h3 class="text-base font-bold text-gray-800 flex items-center gap-2"><span class="w-3 h-3 rounded-full ${selGroup.color}"></span>รายละเอียดกลุ่ม — ${selGroup.full} <span class="text-sm font-normal text-gray-400">(${selGroup.resps.length} คน)</span></h3></div>`;
  h += selGroup.resps.length ? surveyAnalysisHTML(selGroup.resps, qs) : `<div class="bg-white rounded-2xl p-6 border border-blue-100 text-center text-gray-400 text-sm">ยังไม่มีผู้ตอบในกลุ่มนี้</div>`;

  h += `<div class="bg-blue-50 border border-blue-100 rounded-xl p-3 mt-4 text-xs text-blue-800">เกณฑ์แปลผล (AUN-QA): 4.51-5.00 มากที่สุด · 3.51-4.50 มาก · 2.51-3.50 ปานกลาง · 1.51-2.50 น้อย · 1.00-1.50 น้อยที่สุด &nbsp;|&nbsp; S.D. คำนวณแบบ n-1 (sample)</div>`;
  return h;
}

// ค่าเฉลี่ยของ "ด้าน" หนึ่ง สำหรับชุดผู้ตอบชุดหนึ่ง
function surveySectionStatFor(resps, qs, section) {
  const secQs = qs.filter(q => q.q_type === 'rating' && q.section === section);
  const parsed = resps.map(r => { let a = {}; try { a = JSON.parse(r.answers_json || '{}'); } catch (_) { } return a; });
  let vals = [];
  secQs.forEach(q => parsed.forEach(a => { const v = Number(a[q.q_id]); if (!isNaN(v) && v >= 1 && v <= 5) vals.push(v); }));
  return surveyMeanSD(vals);
}

// ตารางผลรายด้าน เทียบระหว่างกลุ่ม (นักศึกษา / บุคลากร)
function surveySectionByGroupHTML(groups, qs) {
  const ratingQs = qs.filter(q => q.q_type === 'rating');
  const sections = [];
  ratingQs.forEach(q => { if (!sections.includes(q.section)) sections.push(q.section); });
  if (!sections.length) return '';
  let h = `<div class="bg-white rounded-2xl border border-blue-100 overflow-hidden mb-5">
    <div class="p-4 border-b"><h4 class="font-bold text-gray-800 flex items-center gap-2"><i data-lucide="layers" class="w-5 h-5 text-primary"></i>ผลรายด้าน — แยกตามกลุ่ม</h4></div>
    <div class="overflow-x-auto"><table class="w-full text-sm"><thead><tr class="bg-gray-50 text-gray-600 text-left">
      <th class="px-4 py-2">ด้าน</th>${groups.map(g => `<th class="px-3 py-2 text-center border-l border-gray-100">${g.label}<br><span class="text-[11px] font-normal text-gray-400">μ (n) · แปลผล</span></th>`).join('')}</tr></thead><tbody>`;
  sections.forEach((sec, idx) => {
    const col = surveySectionColor(idx);
    h += `<tr class="border-t border-gray-100"><td class="px-4 py-2 font-medium text-gray-700"><span class="inline-block w-2.5 h-2.5 rounded-full ${col.dot} mr-1 align-middle"></span>${surveyEsc(sec)}</td>`;
    groups.forEach(g => {
      const s = surveySectionStatFor(g.resps, qs, sec); const it = surveyInterpret(s.mean);
      h += `<td class="px-3 py-2 text-center border-l border-gray-100">${s.n ? `<span class="font-bold text-gray-800">${s.mean.toFixed(2)}</span> <span class="text-gray-400 text-xs">(${s.n})</span><br><span class="px-2 py-0.5 rounded-full text-[11px] ${it.c}">${it.t}</span>` : '<span class="text-gray-300">—</span>'}</td>`;
    });
    h += `</tr>`;
  });
  // แถวรวมทุกด้าน
  h += `<tr class="bg-blue-50 border-t"><td class="px-4 py-2 font-bold text-primary">รวมทุกด้าน</td>`;
  groups.forEach(g => {
    const parsed = g.resps.map(r => { let a = {}; try { a = JSON.parse(r.answers_json || '{}'); } catch (_) { } return a; });
    let vals = []; ratingQs.forEach(q => parsed.forEach(a => { const v = Number(a[q.q_id]); if (!isNaN(v) && v >= 1 && v <= 5) vals.push(v); }));
    const s = surveyMeanSD(vals);
    h += `<td class="px-3 py-2 text-center border-l border-blue-100 font-bold text-primary">${s.n ? `${s.mean.toFixed(2)} <span class="text-xs font-normal">(${s.n})</span>` : '—'}</td>`;
  });
  h += `</tr></tbody></table></div></div>`;
  return h;
}

// สรุปผลของชุดผู้ตอบชุดหนึ่ง (ใช้กับแต่ละกลุ่มบทบาท)
function surveyAnalysisHTML(resps, qs) {
  const parsed = resps.map(r => { let a = {}; try { a = JSON.parse(r.answers_json || '{}'); } catch (_) { } return { r, a }; });

  let allVals = [];
  parsed.forEach(p => Object.keys(p.a).forEach(k => { const v = Number(p.a[k]); if (!isNaN(v) && v >= 1 && v <= 5) allVals.push(v); }));
  const grand = surveyMeanSD(allVals);
  const grandInt = surveyInterpret(grand.mean);

  const byRole = {}, byDevice = {}, byYear = {};
  resps.forEach(r => {
    const rl = r.role_label || SURVEY_ROLE_LABEL[r.role] || r.role || '-';
    byRole[rl] = (byRole[rl] || 0) + 1;
    const dv = r.device || '-'; byDevice[dv] = (byDevice[dv] || 0) + 1;
    if (norm(r.role) === 'student') { const y = r.year_level || 'ไม่ระบุ'; byYear[y] = (byYear[y] || 0) + 1; }
  });
  const chip = (obj) => Object.entries(obj).sort((a, b) => b[1] - a[1]).map(([k, v]) => `<span class="inline-flex items-center gap-1 px-3 py-1 bg-gray-100 rounded-full text-sm text-gray-700">${surveyEsc(k)} <b class="text-primary">${v}</b></span>`).join(' ');

  let h = `<div class="grid grid-cols-1 sm:grid-cols-3 gap-4 mb-4">
    ${statCard('users', 'จำนวนผู้ตอบ', resps.length, 'คน', 'bg-blue-500')}
    ${statCard('bar-chart-3', 'ค่าเฉลี่ยรวม (μ)', grand.mean.toFixed(2), '/ 5.00', 'bg-emerald-500')}
    <div class="bg-white rounded-2xl p-4 border border-blue-100 flex items-center gap-3">
      <div class="w-11 h-11 rounded-xl ${grandInt.c} flex items-center justify-center"><i data-lucide="award" class="w-6 h-6"></i></div>
      <div><p class="text-xs text-gray-500">ระดับความพึงพอใจ (AUN-QA)</p><p class="text-lg font-bold text-gray-800">${grandInt.t}</p><p class="text-xs text-gray-500">S.D. = ${grand.sd.toFixed(2)} · ${(grand.mean / 5 * 100).toFixed(1)}%</p></div></div>
  </div>`;

  h += `<div class="bg-white rounded-2xl p-4 border border-blue-100 mb-4">
    <h4 class="font-bold text-gray-700 text-sm mb-2">ข้อมูลผู้ตอบ — แยกตามบทบาท</h4><div class="flex flex-wrap gap-2 mb-3">${chip(byRole)}</div>
    <h4 class="font-bold text-gray-700 text-sm mb-2">อุปกรณ์ที่ใช้งานเป็นหลัก</h4><div class="flex flex-wrap gap-2 mb-3">${chip(byDevice)}</div>
    ${Object.keys(byYear).length ? `<h4 class="font-bold text-gray-700 text-sm mb-2">นักศึกษา — แยกตามชั้นปี</h4><div class="flex flex-wrap gap-2">${chip(byYear)}</div>` : ''}
  </div>`;

  // ตารางรายข้อ (เฉพาะมาตรวัด 1-5) — แยกแต่ละด้าน คลิกเปิด-ยุบได้
  const ratingQs = qs.filter(q => q.q_type === 'rating');
  const sections = [];
  ratingQs.forEach(q => { if (!sections.includes(q.section)) sections.push(q.section); });
  h += `<div class="bg-white rounded-2xl border border-blue-100 overflow-hidden mb-4"><div class="p-4 border-b flex items-center justify-between"><h4 class="font-bold text-gray-800">ผลรายข้อ (μ, S.D., ร้อยละ, แปลผล)</h4><span class="text-xs text-gray-400">คลิกที่ด้านเพื่อเปิด/ยุบ</span></div>`;
  sections.forEach((sec, idx) => {
    const col = surveySectionColor(idx);
    const secQs = ratingQs.filter(q => q.section === sec);
    let secVals = [];
    const rows = secQs.map(q => {
      const vals = parsed.map(p => Number(p.a[q.q_id])).filter(v => !isNaN(v) && v >= 1 && v <= 5);
      secVals = secVals.concat(vals);
      const s = surveyMeanSD(vals); const it = surveyInterpret(s.mean);
      return `<tr class="border-t border-gray-100"><td class="px-4 py-2 text-gray-700">${surveyEsc(q.question_text)}</td>
        <td class="px-3 py-2 text-center text-gray-500">${s.n}</td><td class="px-3 py-2 text-center font-semibold">${s.mean.toFixed(2)}</td>
        <td class="px-3 py-2 text-center text-gray-500">${s.sd.toFixed(2)}</td><td class="px-3 py-2 text-center text-gray-500">${(s.mean / 5 * 100).toFixed(1)}</td>
        <td class="px-3 py-2 text-center"><span class="px-2 py-0.5 rounded-full text-xs ${it.c}">${it.t}</span></td></tr>`;
    }).join('');
    const ss = surveyMeanSD(secVals); const sit = surveyInterpret(ss.mean);
    h += `<details class="border-t border-gray-100">
      <summary class="cursor-pointer select-none px-4 py-3 flex items-center justify-between gap-3 hover:bg-gray-50">
        <span class="font-bold ${col.head} flex items-center gap-2 min-w-0"><span class="w-2.5 h-2.5 rounded-full ${col.dot} shrink-0"></span><span class="truncate">${surveyEsc(sec)}</span></span>
        <span class="flex items-center gap-2 shrink-0 text-sm"><span class="font-bold text-gray-800">μ ${ss.mean.toFixed(2)}</span><span class="px-2 py-0.5 rounded-full text-xs ${sit.c}">${sit.t}</span><i data-lucide="chevron-down" class="chev w-4 h-4 text-gray-400"></i></span>
      </summary>
      <div class="overflow-x-auto border-t border-gray-100"><table class="w-full text-sm"><thead><tr class="bg-gray-50 text-gray-600 text-left">
        <th class="px-4 py-2">ข้อคำถาม</th><th class="px-3 py-2 text-center">n</th><th class="px-3 py-2 text-center">μ</th><th class="px-3 py-2 text-center">S.D.</th><th class="px-3 py-2 text-center">ร้อยละ</th><th class="px-3 py-2 text-center">แปลผล</th></tr></thead><tbody>${rows}</tbody></table></div>
    </details>`;
  });
  h += `</div>`;

  // คำถามแบบตัวเลือก — แสดงการกระจายคำตอบ (distribution)
  const choiceQs = qs.filter(q => q.q_type === 'choice');
  if (choiceQs.length) {
    h += `<div class="bg-white rounded-2xl p-4 border border-blue-100 mb-4"><h4 class="font-bold text-gray-800 mb-3">ผลคำถามแบบตัวเลือก</h4>`;
    choiceQs.forEach(q => {
      const counts = {};
      parsed.forEach(p => { const v = (p.a[q.q_id] || '').toString().trim(); if (v) counts[v] = (counts[v] || 0) + 1; });
      const total = Object.values(counts).reduce((a, b) => a + b, 0);
      const opts = surveyParseOptions(q.options);
      const keys = opts.length ? opts.slice() : Object.keys(counts);
      Object.keys(counts).forEach(k => { if (keys.indexOf(k) === -1) keys.push(k); });
      h += `<div class="mb-3"><p class="text-sm font-medium text-gray-700 mb-1">${surveyEsc(q.question_text)} <span class="text-xs text-gray-400">(${total})</span></p>
        <div class="space-y-1">${keys.map(k => {
        const c = counts[k] || 0; const pct = total ? (c / total * 100) : 0;
        return `<div class="flex items-center gap-2 text-sm"><span class="w-40 truncate text-gray-600" title="${surveyEsc(k)}">${surveyEsc(k)}</span>
            <div class="flex-1 bg-gray-100 rounded-full h-3"><div class="bg-primary h-3 rounded-full" style="width:${pct.toFixed(0)}%"></div></div>
            <span class="w-20 text-right text-gray-500">${c} (${pct.toFixed(0)}%)</span></div>`;
      }).join('')}</div></div>`;
    });
    h += `</div>`;
  }

  // ข้อเสนอแนะ (text)
  const textQs = qs.filter(q => q.q_type === 'text');
  if (textQs.length) {
    h += `<div class="bg-white rounded-2xl p-4 border border-blue-100"><h4 class="font-bold text-gray-800 mb-3">ข้อเสนอแนะเพิ่มเติม (เชิงคุณภาพ)</h4>`;
    textQs.forEach(q => {
      const items = parsed.map(p => ({ txt: (p.a[q.q_id] || '').toString().trim(), role: p.r.role_label || p.r.role || '' })).filter(x => x.txt);
      h += `<div class="mb-3"><p class="text-sm font-medium text-gray-700 mb-1">${surveyEsc(q.question_text)} <span class="text-xs text-gray-400">(${items.length})</span></p>`;
      if (!items.length) h += `<p class="text-sm text-gray-400">— ไม่มีผู้ตอบ —</p>`;
      else h += `<ul class="space-y-1">${items.map(x => `<li class="text-sm text-gray-600 bg-gray-50 rounded-lg px-3 py-2">"${surveyEsc(x.txt)}" <span class="text-xs text-gray-400">— ${surveyEsc(x.role)}</span></li>`).join('')}</ul>`;
      h += `</div>`;
    });
    h += `</div>`;
  }

  return h;
}

// ดาวน์โหลดสรุปผลแบบประเมินเป็น PDF — เปิดหน้าจัดรูปแบบ A4 (ฟอนต์ Sarabun) แล้วสั่งพิมพ์/บันทึกเป็น PDF
// (ใช้แนวทางเดียวกับการพิมพ์ใบรายงานผลการเรียน/ใบ Transcript ของระบบเดิม)
function downloadSurveyResultsPDF(year) {
  const resps = surveyResponsesForYear(year);
  if (!resps.length) { showToast('ยังไม่มีผู้ตอบแบบประเมินของปีการศึกษานี้', 'error'); return; }
  const qs = surveyQuestionsForYear(year, false);
  const parsed = resps.map(r => { let a = {}; try { a = JSON.parse(r.answers_json || '{}'); } catch (_) { } return { r, a }; });

  // ค่าเฉลี่ยรวม
  let allVals = [];
  parsed.forEach(p => Object.keys(p.a).forEach(k => { const v = Number(p.a[k]); if (!isNaN(v) && v >= 1 && v <= 5) allVals.push(v); }));
  const grand = surveyMeanSD(allVals); const gi = surveyInterpret(grand.mean);

  // ข้อมูลผู้ตอบ
  const byRole = {}, byDevice = {}, byFreq = {}, byYear = {};
  resps.forEach(r => {
    const rl = r.role_label || SURVEY_ROLE_LABEL[r.role] || r.role || '-'; byRole[rl] = (byRole[rl] || 0) + 1;
    byDevice[r.device || '-'] = (byDevice[r.device || '-'] || 0) + 1;
    byFreq[r.frequency || '-'] = (byFreq[r.frequency || '-'] || 0) + 1;
    if (norm(r.role) === 'student') { const y = r.year_level || 'ไม่ระบุ'; byYear[y] = (byYear[y] || 0) + 1; }
  });
  const kv = (obj) => { const ent = Object.entries(obj).sort((a, b) => b[1] - a[1]); const tot = ent.reduce((s, e) => s + e[1], 0) || 1; return `<table class="mini"><tbody>${ent.map(([k, v]) => `<tr><td>${surveyEsc(k)}</td><td class="c">${v} (${(v / tot * 100).toFixed(1)}%)</td></tr>`).join('')}</tbody></table>`; };

  // ตารางผลรายข้อ (rating) แยกตามด้าน
  const ratingQs = qs.filter(q => q.q_type === 'rating');
  const sections = [];
  ratingQs.forEach(q => { if (!sections.includes(q.section)) sections.push(q.section); });
  let tableRows = '';
  sections.forEach(sec => {
    const secQs = ratingQs.filter(q => q.section === sec);
    let secVals = [];
    const rows = secQs.map(q => {
      const vals = parsed.map(p => Number(p.a[q.q_id])).filter(v => !isNaN(v) && v >= 1 && v <= 5);
      secVals = secVals.concat(vals);
      const s = surveyMeanSD(vals); const it = surveyInterpret(s.mean);
      return `<tr><td>${surveyEsc(q.question_text)}</td><td class="c">${s.n}</td><td class="c">${s.mean.toFixed(2)}</td><td class="c">${s.sd.toFixed(2)}</td><td class="c">${(s.mean / 5 * 100).toFixed(1)}</td><td class="c">${it.t}</td></tr>`;
    }).join('');
    const ss = surveyMeanSD(secVals); const sit = surveyInterpret(ss.mean);
    tableRows += `<tr class="sec"><td>▸ ${surveyEsc(sec)}</td><td class="c">-</td><td class="c">${ss.mean.toFixed(2)}</td><td class="c">${ss.sd.toFixed(2)}</td><td class="c">${(ss.mean / 5 * 100).toFixed(1)}</td><td class="c">${sit.t}</td></tr>` + rows;
  });

  // คำถามแบบตัวเลือก
  const choiceQs = qs.filter(q => q.q_type === 'choice');
  let choiceHTML = '';
  if (choiceQs.length) {
    choiceHTML = '<h3>ผลคำถามแบบตัวเลือก</h3>';
    choiceQs.forEach(q => {
      const counts = {}; parsed.forEach(p => { const v = (p.a[q.q_id] || '').toString().trim(); if (v) counts[v] = (counts[v] || 0) + 1; });
      const tot = Object.values(counts).reduce((a, b) => a + b, 0) || 1;
      const opts = surveyParseOptions(q.options); const keys = opts.length ? opts.slice() : Object.keys(counts);
      Object.keys(counts).forEach(k => { if (keys.indexOf(k) === -1) keys.push(k); });
      choiceHTML += `<div class="q"><b>${surveyEsc(q.question_text)}</b><table class="mini"><tbody>${keys.map(k => { const c = counts[k] || 0; return `<tr><td>${surveyEsc(k)}</td><td class="c">${c} (${(c / tot * 100).toFixed(1)}%)</td></tr>`; }).join('')}</tbody></table></div>`;
    });
  }

  // ข้อเสนอแนะ
  const textQs = qs.filter(q => q.q_type === 'text');
  let sugHTML = '';
  if (textQs.length) {
    sugHTML = '<h3>ข้อเสนอแนะเพิ่มเติม (เชิงคุณภาพ)</h3>';
    textQs.forEach(q => {
      const items = parsed.map(p => ({ txt: (p.a[q.q_id] || '').toString().trim(), role: p.r.role_label || p.r.role || '' })).filter(x => x.txt);
      sugHTML += `<div class="q"><b>${surveyEsc(q.question_text)} (${items.length})</b>`;
      sugHTML += items.length ? `<ul>${items.map(x => `<li>${surveyEsc(x.txt)} <span class="muted">— ${surveyEsc(x.role)}</span></li>`).join('')}</ul>` : '<p class="muted">— ไม่มีผู้ตอบ —</p>';
      sugHTML += '</div>';
    });
  }

  const cfg = surveyConfigForYear(year);
  const title = (cfg && norm(cfg.title)) || 'แบบประเมินความพึงพอใจการใช้งานระบบ EMS-BCNB';
  const college = (APP.config && APP.config.college_name) || 'วิทยาลัยพยาบาลบรมราชชนนี กรุงเทพ';
  let today = ''; try { today = new Date().toLocaleDateString('th-TH', { year: 'numeric', month: 'long', day: 'numeric' }); } catch (_) { today = new Date().toLocaleDateString(); }

  const html = `<!DOCTYPE html><html lang="th"><head><meta charset="UTF-8"><title>สรุปผลแบบประเมิน ปีการศึกษา ${year}</title>
<style>@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;600;700&display=swap');
*{font-family:'Sarabun',sans-serif;box-sizing:border-box}
body{margin:0;padding:10mm;color:#1f2937;font-size:12px}
h1{font-size:18px;text-align:center;margin:0}
h2{font-size:13px;text-align:center;margin:2px 0 0;font-weight:400;color:#374151}
h3{font-size:13px;margin:14px 0 5px;border-left:4px solid #1e6fba;padding-left:8px}
.meta{text-align:center;color:#6b7280;font-size:11px;margin:6px 0 10px}
table{width:100%;border-collapse:collapse;margin-top:4px}
th,td{border:1px solid #cbd5e1;padding:5px 7px;font-size:11px;vertical-align:top}
th{background:#eaf2fb}
.c{text-align:center;white-space:nowrap}
tr.sec td{background:#f0f7ff;font-weight:600;color:#14507f}
.summary{display:flex;gap:8px;margin:8px 0}
.box{flex:1;border:1px solid #cbd5e1;border-radius:8px;padding:8px;text-align:center}
.box .n{font-size:19px;font-weight:700;color:#1e6fba}
.box .l{font-size:10px;color:#6b7280}
.muted{color:#6b7280}
ul{margin:4px 0;padding-left:18px}
li{margin:2px 0}
.q{margin:6px 0}
.mini{width:60%}
.foot{margin-top:14px;font-size:10px;color:#6b7280;border-top:1px solid #e5e7eb;padding-top:6px}
@media print{@page{size:A4;margin:8mm}h3,.q{page-break-inside:avoid}tr{page-break-inside:avoid}}
</style></head><body>
<h1>${surveyEsc(title)}</h1>
<h2>${surveyEsc(college)}</h2>
<div class="meta">สรุปผลการประเมิน ปีการศึกษา ${year} · จำนวนผู้ตอบ ${resps.length} คน · พิมพ์เมื่อ ${today}</div>
<div class="summary">
  <div class="box"><div class="n">${resps.length}</div><div class="l">จำนวนผู้ตอบ (คน)</div></div>
  <div class="box"><div class="n">${grand.mean.toFixed(2)}</div><div class="l">ค่าเฉลี่ยรวม μ (เต็ม 5)</div></div>
  <div class="box"><div class="n">${grand.sd.toFixed(2)}</div><div class="l">ส่วนเบี่ยงเบนมาตรฐาน S.D.</div></div>
  <div class="box"><div class="n">${(grand.mean / 5 * 100).toFixed(1)}%</div><div class="l">ระดับ: ${gi.t}</div></div>
</div>
<h3>ข้อมูลผู้ตอบ — แยกตามบทบาท</h3>${kv(byRole)}
<h3>อุปกรณ์ที่ใช้งานเป็นหลัก</h3>${kv(byDevice)}
<h3>ความถี่ในการใช้งานระบบ</h3>${kv(byFreq)}
${Object.keys(byYear).length ? `<h3>นักศึกษา — แยกตามชั้นปี</h3>${kv(byYear)}` : ''}
${tableRows ? `<h3>ผลรายข้อ (μ, S.D., ร้อยละ, แปลผล)</h3>
<table><thead><tr><th>ข้อคำถาม</th><th class="c">n</th><th class="c">μ</th><th class="c">S.D.</th><th class="c">ร้อยละ</th><th class="c">แปลผล</th></tr></thead><tbody>${tableRows}</tbody></table>` : ''}
${choiceHTML}
${sugHTML}
<div class="foot">เกณฑ์แปลผล (AUN-QA): 4.51-5.00 มากที่สุด · 3.51-4.50 มาก · 2.51-3.50 ปานกลาง · 1.51-2.50 น้อย · 1.00-1.50 น้อยที่สุด | S.D. คำนวณแบบ n-1 (sample) | ระบบบริหารจัดการงานวิชาการ EMS-BCNB</div>
<script>window.onload=function(){window.print()}<\/script></body></html>`;

  const w = window.open('', '_blank');
  if (w) { w.document.write(html); w.document.close(); }
  else { showToast('กรุณาอนุญาต Popup เพื่อดาวน์โหลด PDF', 'error'); }
}
