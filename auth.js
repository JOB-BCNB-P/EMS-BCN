/* ===================================================================
   auth.js — ด่านล็อกอิน Google เฉพาะบัญชี @bcn.ac.th
   วิทยาลัยพยาบาลบรมราชชนนี กรุงเทพ (EMS-BCNB)
   -------------------------------------------------------------------
   ทำงาน 2 ชั้น:
   (1) ฝั่งหน้าเว็บ  : บังคับลงชื่อเข้าด้วย Google ก่อนเข้าระบบ และรับ ID token
   (2) ฝั่งเซิร์ฟเวอร์ (code.gs): ตรวจ token ทุก request ว่าเป็นโดเมน bcn.ac.th จริง
       → ชั้นนี้คือด่านที่กันคนนอกได้จริง (เพราะไฟล์เว็บเป็น public บน GitHub Pages)

   ▼▼▼ ตั้งค่าครั้งเดียว: ใส่ OAuth Client ID (Web application) ที่สร้างจาก Google Cloud Console ▼▼▼
   =================================================================== */
var DEFAULT_CLIENT_ID = '284853368755-27oi5mn1po4aog33vgbbfprh52auq4e5.apps.googleusercontent.com';   // OAuth Client ID (Web) โปรเจกต์ EMSbcn
var ALLOWED_DOMAIN = 'bcn.ac.th';

(function (global) {
  'use strict';

  var LS_TOKEN = 'ems_google_idtoken';
  var LS_CLIENT = 'ems_google_client_id';
  var _onSuccess = null;
  var _refreshTimer = null;

  function clientId() {
    return (localStorage.getItem(LS_CLIENT) || DEFAULT_CLIENT_ID || '').trim();
  }

  /* ---------- token store ---------- */
  function saveToken(tok, payload) {
    localStorage.setItem(LS_TOKEN, JSON.stringify({ t: tok, exp: (payload.exp || 0) * 1000, email: payload.email || '' }));
  }
  function readToken() {
    try { return JSON.parse(localStorage.getItem(LS_TOKEN) || 'null'); } catch (e) { return null; }
  }
  function clearToken() { localStorage.removeItem(LS_TOKEN); }

  function isValid() {
    var s = readToken();
    return !!(s && s.t && s.exp && s.exp - 60000 > Date.now());
  }
  function getToken() { var s = readToken(); return s ? s.t : ''; }
  function getEmail() { var s = readToken(); return s ? s.email : ''; }

  /* ---------- JWT decode + domain check ---------- */
  function decodeJwt(t) {
    try {
      var p = t.split('.')[1].replace(/-/g, '+').replace(/_/g, '/');
      var pad = p.length % 4; if (pad) p += new Array(5 - pad).join('=');
      return JSON.parse(decodeURIComponent(escape(atob(p))));
    } catch (e) { return null; }
  }
  function domainOk(payload) {
    if (!payload) return false;
    var email = String(payload.email || '').toLowerCase();
    var hd = String(payload.hd || '').toLowerCase();
    return (hd === ALLOWED_DOMAIN) || new RegExp('@' + ALLOWED_DOMAIN.replace('.', '\\.') + '$').test(email);
  }

  /* ---------- fetch interceptor: แนบ id_token ทุก request ไป Apps Script ---------- */
  var _origFetch = global.fetch.bind(global);
  global.fetch = function (input, opts) {
    try {
      var url = (typeof input === 'string') ? input : (input && input.url) || '';
      if (/script\.google\.com\/macros\//.test(url)) {
        var tok = getToken();
        if (tok && url.indexOf('id_token=') < 0) {
          var sep = url.indexOf('?') >= 0 ? '&' : '?';
          url += sep + 'id_token=' + encodeURIComponent(tok);
          if (typeof input === 'string') input = url; else input = new Request(url, input);
        }
      }
    } catch (e) { /* ไม่ขวางการทำงานปกติ */ }
    return _origFetch(input, opts);
  };

  /* ---------- GIS ---------- */
  function loadGIS(cb) {
    if (global.google && google.accounts && google.accounts.id) { cb(); return; }
    var s = document.createElement('script');
    s.src = 'https://accounts.google.com/gsi/client';
    s.async = true; s.defer = true; s.onload = cb;
    s.onerror = function () { gateError('โหลด Google Sign-In ไม่สำเร็จ ตรวจสอบการเชื่อมต่ออินเทอร์เน็ต'); };
    document.head.appendChild(s);
  }

  function onCredential(resp) {
    var tok = resp && resp.credential;
    var payload = tok && decodeJwt(tok);
    if (!payload) { gateError('รับข้อมูลบัญชีไม่สำเร็จ ลองใหม่อีกครั้ง'); return; }
    if (!domainOk(payload)) {
      gateError('อนุญาตเฉพาะบัญชีอีเมล @' + ALLOWED_DOMAIN + ' ของวิทยาลัยเท่านั้น<br><span class="text-xs">บัญชีที่ใช้: ' + (payload.email || '-') + '</span>');
      try { google.accounts.id.disableAutoSelect(); } catch (e) {}
      return;
    }
    saveToken(tok, payload);
    scheduleRefresh(payload.exp);
    removeGate();
    if (_onSuccess) { var cb = _onSuccess; _onSuccess = null; cb(); }
  }

  // ต่ออายุ token อัตโนมัติก่อนหมดอายุ ~5 นาที (เงียบ ๆ ถ้ายังล็อกอิน Google อยู่)
  function scheduleRefresh(expSec) {
    if (_refreshTimer) clearTimeout(_refreshTimer);
    if (!expSec) return;
    var ms = expSec * 1000 - Date.now() - 5 * 60000;
    if (ms < 10000) ms = 10000;
    _refreshTimer = setTimeout(function () {
      try { google.accounts.id.prompt(); } catch (e) {}
    }, Math.min(ms, 2147483000));
  }

  /* ---------- Gate UI ---------- */
  function gateError(html) {
    var el = document.getElementById('authGateError');
    if (el) { el.innerHTML = html; el.classList.remove('hidden'); }
  }
  function removeGate() {
    var g = document.getElementById('authGate');
    if (g) g.parentNode.removeChild(g);
  }
  function buildGate() {
    if (document.getElementById('authGate')) return;
    var wrap = document.createElement('div');
    wrap.id = 'authGate';
    wrap.style.cssText = 'position:fixed;inset:0;z-index:9999;display:flex;align-items:center;justify-content:center;background:linear-gradient(135deg,#f0f7ff,#ffffff 45%,#e6f1ff);';
    var needCid = !clientId();
    wrap.innerHTML =
      '<div style="width:100%;max-width:26rem;margin:1rem;" class="fade-in">' +
      '<div style="text-align:center;margin-bottom:1.25rem;">' +
      '<img src="https://cdn.jsdelivr.net/gh/JOB-BCNB-P/LOGO/Logo%20Thai.png" alt="Logo" style="width:88px;height:88px;object-fit:contain;margin:0 auto .75rem;">' +
      '<h1 style="font-size:1.35rem;font-weight:700;color:#1f2937;">ระบบบริหารจัดการงานวิชาการ</h1>' +
      '<p style="color:#6b7280;margin-top:.25rem;">วิทยาลัยพยาบาลบรมราชชนนี กรุงเทพ</p></div>' +
      '<div style="background:#fff;border:1px solid #dbeafe;border-radius:1rem;box-shadow:0 10px 40px rgba(30,111,186,.12);padding:1.75rem;text-align:center;">' +
      '<p style="color:#374151;font-weight:600;margin-bottom:.35rem;">เข้าสู่ระบบด้วยบัญชีวิทยาลัย</p>' +
      '<p style="color:#6b7280;font-size:.85rem;margin-bottom:1rem;">อนุญาตเฉพาะอีเมล <b>@' + ALLOWED_DOMAIN + '</b></p>' +
      (needCid ? clientIdSetupHTML() :
        '<div id="gsiButton" style="display:flex;justify-content:center;"></div>' +
        '<p style="color:#9ca3af;font-size:.75rem;margin-top:1rem;">หากหน้าต่างไม่ขึ้น ให้กดปุ่มด้านบน หรืออนุญาตป็อปอัปของเบราว์เซอร์</p>') +
      '<div id="authGateError" class="hidden" style="margin-top:1rem;color:#e11d48;font-size:.85rem;background:#fff1f2;border:1px solid #fecdd3;border-radius:.6rem;padding:.6rem;"></div>' +
      '</div>' +
      '<p style="text-align:center;color:#9ca3af;font-size:10px;margin-top:1rem;">พัฒนาระบบ โดย Oranit.R นักวิชาการศึกษา BCNB</p>' +
      '</div>';
    document.body.appendChild(wrap);
  }

  function clientIdSetupHTML() {
    return '<div style="text-align:left;background:#eff6ff;border:1px solid #bfdbfe;border-radius:.6rem;padding:.75rem;font-size:.8rem;color:#1e40af;">' +
      '<p style="font-weight:600;margin-bottom:.35rem;">ยังไม่ได้ตั้งค่า Google Client ID</p>' +
      '<p>วาง OAuth Client ID (Web) จาก Google Cloud Console เพื่อเปิดใช้ครั้งแรก (ดูวิธีใน README):</p>' +
      '<input id="cidInput" type="text" placeholder="xxxx.apps.googleusercontent.com" style="width:100%;margin-top:.5rem;border:1px solid #cbd5e1;border-radius:.5rem;padding:.5rem;font-size:.8rem;">' +
      '<button onclick="EMSAuth.setClientId(document.getElementById(\'cidInput\').value)" style="width:100%;margin-top:.5rem;background:#1e6fba;color:#fff;font-weight:600;padding:.5rem;border-radius:.5rem;">บันทึกและเริ่มใช้งาน</button>' +
      '<p style="margin-top:.4rem;font-size:.72rem;color:#3b82f6;">* แนะนำให้ฝังไว้ในไฟล์ auth.js (ตัวแปร DEFAULT_CLIENT_ID) เพื่อให้ทุกเครื่องใช้ได้เลย</p></div>';
  }

  function setClientId(v) {
    v = (v || '').trim();
    if (!/\.apps\.googleusercontent\.com$/.test(v)) { gateError('รูปแบบ Client ID ไม่ถูกต้อง'); return; }
    localStorage.setItem(LS_CLIENT, v);
    removeGate();
    showGate(_onSuccess);
  }

  function renderButton() {
    loadGIS(function () {
      try {
        google.accounts.id.initialize({
          client_id: clientId(),
          callback: onCredential,
          auto_select: true,
          hd: ALLOWED_DOMAIN,
          cancel_on_tap_outside: false
        });
        var btn = document.getElementById('gsiButton');
        if (btn) google.accounts.id.renderButton(btn, { theme: 'filled_blue', size: 'large', width: 280, text: 'signin_with', shape: 'pill' });
        google.accounts.id.prompt();
      } catch (e) { gateError('เริ่ม Google Sign-In ไม่สำเร็จ: ' + e.message + '<br><span class="text-xs">ตรวจสอบว่า Client ID และ Authorized JavaScript origins ตั้งค่าถูกต้อง</span>'); }
    });
  }

  /* ---------- public API ---------- */
  function showGate(onSuccess) {
    _onSuccess = onSuccess || _onSuccess;
    // token ยังไม่หมดอายุ → ผ่านเลย
    if (isValid()) { var s = readToken(); scheduleRefresh(s.exp / 1000); if (_onSuccess) { var cb = _onSuccess; _onSuccess = null; cb(); } return; }
    clearToken();
    buildGate();
    if (clientId()) renderButton();
  }

  function signOut() {
    clearToken();
    try { if (global.google && google.accounts) google.accounts.id.disableAutoSelect(); } catch (e) {}
  }

  global.EMSAuth = {
    ALLOWED_DOMAIN: ALLOWED_DOMAIN,
    showGate: showGate,
    isValid: isValid,
    getToken: getToken,
    getEmail: getEmail,
    signOut: signOut,
    setClientId: setClientId
  };
})(window);
