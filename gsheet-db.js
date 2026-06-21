// ================================================================
// gsheet-db.js — Google Sheets data layer (Read + Write)
//
// READ:  via public gviz endpoint (no auth needed)
// WRITE: via Google Apps Script Web App (deployed from the Sheet)
// ================================================================

const GSheetDB = (() => {
    let _spreadsheetId = null;
    let _scriptUrl = null;
    let _onDataChanged = null;
    let _allData = [];
    let _studentNID = null;   // เก็บเลขบัตรของนักศึกษาไว้ในหน่วยความจำ (เฉพาะ session) เพื่อใช้รีเฟรช — ไม่บันทึกลง localStorage

    const SHEET_TABS = [
        'student', 'teacher', 'subject', 'schedule',
        'grade', 'eng_result', 'leave',
        'tracking', 'result_tracking', 'grade_tracking', 'file_tracking', 'announcement', 'user', 'doc_request', 'permission', 'teacher_directory', 'directory_summary', 'login_log', 'special_teacher', 'alumni', 'password_log'
    ];



    // ---------- Config ----------
    function getStoredConfig() {
        try { const r = localStorage.getItem('gsheetConfig'); return r ? JSON.parse(r) : null; } catch { return null; }
    }
    function storeConfig(c) { localStorage.setItem('gsheetConfig', JSON.stringify(c)); }
    function clearConfig() { localStorage.removeItem('gsheetConfig'); }

    function extractSheetId(input) {
        if (!input) return null;
        input = input.trim();
        const m = input.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
        if (m) return m[1];
        if (/^[a-zA-Z0-9_-]{20,}$/.test(input)) return input;
        return null;
    }

    // ---------- READ via Apps Script (เฟส 2 — ชีตเป็นส่วนตัวได้) ----------
    // เรียกข้อมูลผ่าน Apps Script แทนการอ่านชีตสาธารณะ (gviz)
    // ทำให้ปิดแชร์ชีตได้ และเซิร์ฟเวอร์ตัดคอลัมน์อ่อนไหว (รหัสผ่าน/เลขบัตร) ออกก่อนส่ง
    async function _readFromScript(tabName) {
        if (!_scriptUrl) throw new Error('ยังไม่ได้ตั้งค่า Apps Script URL — อ่านข้อมูลไม่ได้');
        const qs = new URLSearchParams({ action: 'read' });
        if (tabName) qs.set('sheet', tabName);
        qs.set('_', String(Date.now()));
        const resp = await fetch(_scriptUrl + '?' + qs.toString(), {
            method: 'POST', redirect: 'follow',
            headers: { 'Content-Type': 'text/plain' },
            body: '{}'
        });
        const json = await resp.json();
        if (!json || !json.isOk) throw new Error((json && json.error) || 'อ่านข้อมูลไม่สำเร็จ');
        return json.data || {};
    }

    // แปลง {tab:[rows]} เป็น array แบน พร้อมใส่ type + __backendId
    function _mergeTabData(out, tabName, rows) {
        (rows || []).forEach((row, idx) => {
            row.type = tabName;
            row.__backendId = `${tabName}_${idx}`;
            out.push(row);
        });
    }

    // ---------- READ (เดิม via gviz — เลิกใช้แล้ว เก็บไว้อ้างอิง) ----------
    async function fetchTab(tabName) {
        // ใช้ range=A1:AZ200000 บังคับให้ gviz อ่านข้อมูลในช่วงนี้ทั้งหมด
        // (ป้องกัน gviz auto-detect range ผิดเมื่อมี blank row คั่นกลาง)
        // ขยายจาก 50000 เป็น 200000 เพื่อรองรับข้อมูลที่อาจมีหลายแสน row
        // + cache buster (timestamp) กันเบราว์เซอร์ cache
        const url = `https://docs.google.com/spreadsheets/d/${_spreadsheetId}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(tabName)}&headers=1&range=A1:AZ200000&_=${Date.now()}`;
        try {
            const resp = await fetch(url, { cache: 'no-store' });
            if (!resp.ok) return [];
            let text = await resp.text();
            const jsonStr = text.match(/google\.visualization\.Query\.setResponse\((.+)\);?$/s);
            if (!jsonStr) return [];
            const json = JSON.parse(jsonStr[1]);
            if (json.status === 'error') {
                console.warn(`[GSheet] Tab "${tabName}" error:`, json.errors);
                return [];
            }
            if (!json.table || !json.table.cols || !json.table.rows) return [];

            let cols = json.table.cols.map(c => (c.label || '').trim());
            if (cols.every(c => c === '')) {
                const fr = json.table.rows[0];
                if (fr && fr.c) { cols = fr.c.map(cell => cell ? String(cell.v || '').trim() : ''); json.table.rows.shift(); }
            }

            const rawRowCount = (json.table.rows || []).length;
            const result = (json.table.rows || []).map((row, idx) => {
                const obj = { type: tabName, __backendId: `${tabName}_${idx}`, __rowIndex: idx + 2 };
                if (row && Array.isArray(row.c)) {
                    row.c.forEach((cell, i) => {
                        if (cols[i]) {
                            let val = cell ? (cell.v !== null && cell.v !== undefined ? String(cell.v) : '') : '';
                            // Fallback: ถ้า cell.v ว่าง/null แต่ cell.f (formatted/displayed value) มีค่า — ใช้แทน
                            // กรณีนี้เกิดเมื่อ cell type ในชีตไม่ตรงกับ column type ที่ gviz auto-detect
                            // (เช่น column เป็น Number แต่ cell เก็บเป็น Text — gviz return null)
                            if (!val && cell && cell.f !== undefined && cell.f !== null && String(cell.f).trim() !== '') {
                                val = String(cell.f).trim();
                            }
                            if (val.match(/^\d+\.0$/)) val = val.replace('.0', '');
                            // Handle scientific notation from Google Sheets (e.g. 1.40932E5)
                            if (val.match(/^[\d.]+[Ee][+\-]?\d+$/)) { try { val = String(BigInt(Math.round(Number(val)))); } catch (_) { val = String(Math.round(Number(val))); } }
                            // ลบอักขระเพี้ยนหน้ารหัสที่เป็นตัวเลขล้วน (เช่น 0'0101300101 หรือ '0101300101)
                            // เกิดจากการบังคับให้ Google Sheets เก็บเลข 0 นำหน้าเป็นข้อความ — เหลือเฉพาะรหัสจริง
                            if (/^0?'\d+$/.test(val)) val = val.replace(/^0?'/, '');
                            obj[cols[i]] = val;
                        }
                    });
                }
                return obj;
            }).filter(obj => Object.entries(obj).some(([k, v]) => !['type', '__backendId', '__rowIndex'].includes(k) && v !== ''));
            console.log(`[GSheet] Tab "${tabName}" — raw rows: ${rawRowCount}, valid rows: ${result.length}`);
            return result;
        } catch (err) { console.warn(`Tab "${tabName}":`, err.message); return []; }
    }

    async function fetchAllData() {
        // ดึงทีละแท็บแบบ "ขนาน" ผ่าน Apps Script (เร็วกว่าเรียกครั้งเดียวอ่านทุกแท็บเรียงกัน)
        const results = await Promise.allSettled(SHEET_TABS.map(tab => _readFromScript(tab)));
        _allData = [];
        results.forEach((res, i) => {
            const tab = SHEET_TABS[i];
            if (res.status === 'fulfilled') _mergeTabData(_allData, tab, res.value[tab]);
            else console.warn(`[GSheet] อ่านแท็บ "${tab}" ไม่สำเร็จ:`, res.reason && res.reason.message);
        });
        if (_onDataChanged) _onDataChanged(_allData);
        return _allData;
    }

    // ---------- WRITE via Apps Script ----------
    // Refresh only one tab and merge into _allData
    async function refreshTab(tabName) {
        const data = await _readFromScript(tabName);
        _allData = _allData.filter(d => d.type !== tabName);
        _mergeTabData(_allData, tabName, data[tabName]);
        if (_onDataChanged) _onDataChanged(_allData);
    }

    async function _callScript(params) {
        if (!_scriptUrl) return { isOk: false, error: 'ยังไม่ได้ตั้งค่า Apps Script URL — ไปที่ตั้งค่าระบบเพื่อกรอก URL' };
        try {
            const qs = new URLSearchParams({ action: params.action });
            if (params.sheet) qs.set('sheet', params.sheet);
            if (params.rowIndex) qs.set('rowIndex', String(params.rowIndex));
            const resp = await fetch(_scriptUrl + '?' + qs.toString(), {
                method: 'POST', redirect: 'follow',
                headers: { 'Content-Type': 'text/plain' },
                body: JSON.stringify(params.data || {})
            });
            const result = await resp.json();

            if (result.isOk && params.sheet) {
                // Refresh data in background
                await refreshTab(params.sheet);
            }
            return result;
        } catch (err) {
            return { isOk: false, error: err.message };
        }
    }

    // ---------- LOGIN (ตรวจรหัสผ่านฝั่งเซิร์ฟเวอร์) ----------
    // payload = { role, identifier, password }
    // รหัสผ่านถูกส่งไปตรวจที่ Apps Script — ไม่ถูกเทียบในเบราว์เซอร์ และไม่ส่งแฮชกลับมา
    async function login(payload) {
        if (!_scriptUrl) return { isOk: false, error: 'ยังไม่ได้ตั้งค่า Apps Script URL — ไม่สามารถเข้าสู่ระบบได้' };
        try {
            const resp = await fetch(_scriptUrl + '?' + new URLSearchParams({ action: 'login' }).toString(), {
                method: 'POST', redirect: 'follow',
                headers: { 'Content-Type': 'text/plain' },
                body: JSON.stringify(payload || {})
            });
            return await resp.json();
        } catch (err) {
            return { isOk: false, error: err.message };
        }
    }

    async function create(obj) {
        const sheet = obj.type;
        if (!sheet) return { isOk: false, error: 'No type' };
        const data = { ...obj };
        delete data.type; delete data.__backendId; delete data.__rowIndex;
        if (!data.created_at) data.created_at = new Date().toISOString();
        return _callScript({ action: 'create', sheet, data });
    }

    async function update(obj) {
        const sheet = obj.type; const rowIndex = obj.__rowIndex;
        if (!sheet || !rowIndex) return { isOk: false, error: 'Missing type/__rowIndex' };
        const data = { ...obj };
        delete data.type; delete data.__backendId; delete data.__rowIndex;
        return _callScript({ action: 'update', sheet, rowIndex, data });
    }

    // อัปเดตหลายแถวรวดเดียว (เขียนทีละแถว แต่ refresh แค่ครั้งเดียวตอนจบ — เร็วกว่า update() วนลูป)
    async function updateMany(objs) {
        if (!_scriptUrl) return { isOk: false, ok: 0, fail: (objs || []).length, error: 'ยังไม่ได้ตั้งค่า Apps Script URL' };
        let ok = 0, fail = 0, sheet = null;
        for (const obj of (objs || [])) {
            sheet = obj.type; const rowIndex = obj.__rowIndex;
            if (!sheet || !rowIndex) { fail++; continue; }
            const data = { ...obj }; delete data.type; delete data.__backendId; delete data.__rowIndex;
            try {
                const resp = await fetch(_scriptUrl + '?' + new URLSearchParams({ action: 'update', sheet, rowIndex: String(rowIndex) }).toString(), {
                    method: 'POST', redirect: 'follow',
                    headers: { 'Content-Type': 'text/plain' },
                    body: JSON.stringify(data)
                });
                const result = await resp.json();
                if (result && result.isOk) ok++; else fail++;
            } catch (err) { fail++; }
        }
        if (sheet) await refreshTab(sheet);
        return { isOk: fail === 0, ok, fail };
    }

    // สร้างหลายแถวรวดเดียว (refresh แค่ครั้งเดียวตอนจบ)
    async function createMany(objs) {
        if (!_scriptUrl) return { isOk: false, ok: 0, fail: (objs || []).length, error: 'ยังไม่ได้ตั้งค่า Apps Script URL' };
        let ok = 0, fail = 0, sheet = null;
        for (const obj of (objs || [])) {
            sheet = obj.type; if (!sheet) { fail++; continue; }
            const data = { ...obj }; delete data.type; delete data.__backendId; delete data.__rowIndex;
            if (!data.created_at) data.created_at = new Date().toISOString();
            try {
                const resp = await fetch(_scriptUrl + '?' + new URLSearchParams({ action: 'create', sheet }).toString(), {
                    method: 'POST', redirect: 'follow',
                    headers: { 'Content-Type': 'text/plain' },
                    body: JSON.stringify(data)
                });
                const result = await resp.json();
                if (result && result.isOk) ok++; else fail++;
            } catch (err) { fail++; }
        }
        if (sheet) await refreshTab(sheet);
        return { isOk: fail === 0, ok, fail };
    }

    async function remove(obj) {
        const sheet = obj.type; const rowIndex = obj.__rowIndex;
        if (!sheet || !rowIndex) return { isOk: false, error: 'Missing type/__rowIndex' };
        return _callScript({ action: 'delete', sheet, rowIndex });
    }

    // ---------- Init ----------
    async function init(config, onDataChanged) {
        _onDataChanged = onDataChanged;
        _spreadsheetId = extractSheetId(config.spreadsheetId || config);
        _scriptUrl = config.scriptUrl || null;
        if (!_spreadsheetId) return { isOk: false, error: 'Invalid Spreadsheet ID' };
        // ไม่ preload ข้อมูลทั้ง 20 แท็บตอนเปิดหน้าอีกต่อไป — โหลดหลังล็อกอินตามบทบาท
        // (นักศึกษา = โหลดเฉพาะข้อมูลตัวเอง, บุคลากร = โหลดครบ) เพื่อลดโหลดช่วงคนเข้าพร้อมกันเยอะ
        return { isOk: true };
    }

    async function refresh() { return fetchAllData(); }       // บุคลากร: โหลดครบทุกแท็บ
    function destroy() { _studentNID = null; _allData = []; }
    function clearSession() { _studentNID = null; _allData = []; }
    function hasWriteAccess() { return !!_scriptUrl; }

    // ล็อกอินนักศึกษาผ่านเซิร์ฟเวอร์ (เลขบัตรไม่ถูกส่งมาที่เบราว์เซอร์)
    // เซิร์ฟเวอร์คืน "เฉพาะข้อมูลของนักศึกษาคนนี้" + ตารางที่ใช้ร่วมกัน มาในครั้งเดียว
    async function studentLogin(nationalId) {
        if (!_scriptUrl) return { isOk: false, error: 'ยังไม่ได้ตั้งค่า Apps Script URL' };
        try {
            _studentNID = nationalId;
            const resp = await fetch(_scriptUrl + '?' + new URLSearchParams({ action: 'studentLogin' }).toString(), {
                method: 'POST', redirect: 'follow',
                headers: { 'Content-Type': 'text/plain' },
                body: JSON.stringify({ national_id: nationalId })
            });
            const result = await resp.json();
            // ถ้ามี bundle ข้อมูลแนบมา ให้สร้าง _allData จากชุดนี้ (ไม่ต้องเรียกอ่านทุกแท็บอีก)
            if (result && result.isOk && result.data) {
                _allData = [];
                Object.keys(result.data).forEach(tab => _mergeTabData(_allData, tab, result.data[tab]));
                if (_onDataChanged) _onDataChanged(_allData);
            }
            return result;
        } catch (err) { return { isOk: false, error: err.message }; }
    }

    // รีเฟรชข้อมูลนักศึกษา (ใช้เลขบัตรที่เก็บไว้ใน session) — ดึงเฉพาะข้อมูลตัวเองเหมือนตอนล็อกอิน
    async function studentRefresh() {
        if (!_studentNID) return { isOk: false, error: 'ไม่มี session นักศึกษา' };
        return studentLogin(_studentNID);
    }

    // เขียนแถวใหม่แบบ "ไม่อ่านกลับ" — ใช้กับ login_log เพื่อไม่ดึง log ทั้งหมดมาที่เบราว์เซอร์
    async function appendNoRefresh(obj) {
        if (!_scriptUrl || !obj || !obj.type) return { isOk: false };
        const sheet = obj.type;
        const data = { ...obj };
        delete data.type; delete data.__backendId; delete data.__rowIndex;
        if (!data.created_at) data.created_at = new Date().toISOString();
        try {
            const resp = await fetch(_scriptUrl + '?' + new URLSearchParams({ action: 'create', sheet }).toString(), {
                method: 'POST', redirect: 'follow',
                headers: { 'Content-Type': 'text/plain' },
                body: JSON.stringify(data)
            });
            return await resp.json();
        } catch (err) { return { isOk: false, error: err.message }; }
    }

    // ---------- PASSWORD OTP (ลืม/เปลี่ยนรหัสผ่าน ยืนยันด้วยอีเมล) ----------
    // ขั้นที่ 1: ขอรหัส OTP ส่งไปอีเมล
    async function requestPasswordOtp(email) {
        if (!_scriptUrl) return { isOk: false, error: 'ยังไม่ได้ตั้งค่า Apps Script URL' };
        try {
            const resp = await fetch(_scriptUrl + '?' + new URLSearchParams({ action: 'requestPasswordOtp' }).toString(), {
                method: 'POST', redirect: 'follow',
                headers: { 'Content-Type': 'text/plain' },
                body: JSON.stringify({ email })
            });
            return await resp.json();
        } catch (err) { return { isOk: false, error: err.message }; }
    }

    // ขั้นที่ 2: ยืนยัน OTP + ตั้งรหัสผ่านใหม่ — payload = { email, code, newPassword, source }
    async function resetPasswordOtp(payload) {
        if (!_scriptUrl) return { isOk: false, error: 'ยังไม่ได้ตั้งค่า Apps Script URL' };
        try {
            const resp = await fetch(_scriptUrl + '?' + new URLSearchParams({ action: 'resetPasswordOtp' }).toString(), {
                method: 'POST', redirect: 'follow',
                headers: { 'Content-Type': 'text/plain' },
                body: JSON.stringify(payload || {})
            });
            return await resp.json();
        } catch (err) { return { isOk: false, error: err.message }; }
    }

    async function debugTab(tabName) {
        try {
            const data = await _readFromScript(tabName);
            const rows = data[tabName] || [];
            return { status: 'ok', rowCount: rows.length, firstRow: rows[0] || null };
        } catch (err) { return { error: err.message }; }
    }

    return {
        getStoredConfig, storeConfig, clearConfig, extractSheetId,
        init, refresh, destroy, clearSession, debugTab, hasWriteAccess, login,
        studentLogin, studentRefresh, appendNoRefresh,
        requestPasswordOtp, resetPasswordOtp,
        create, createMany, update, updateMany, delete: remove, SHEET_TABS
    };
})();
