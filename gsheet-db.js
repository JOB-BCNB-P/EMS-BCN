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

    const SHEET_TABS = [
        'student', 'teacher', 'subject', 'schedule',
        'grade', 'eng_result', 'leave',
        'tracking', 'result_tracking', 'grade_tracking', 'file_tracking', 'announcement', 'user', 'doc_request', 'permission', 'teacher_directory', 'directory_summary', 'login_log', 'special_teacher'
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

    // ---------- READ ----------
    async function fetchTab(tabName) {
        // ใช้ range=A1:AZ50000 บังคับให้ gviz อ่านข้อมูลในช่วงนี้ทั้งหมด
        // (ป้องกัน gviz auto-detect range ผิดเมื่อมี blank row คั่นกลาง)
        // ขยายจาก 5000 เป็น 50000 เพื่อรองรับข้อมูล grade ที่อาจมีหลายหมื่น row
        // + cache buster (timestamp) กันเบราว์เซอร์ cache
        const url = `https://docs.google.com/spreadsheets/d/${_spreadsheetId}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(tabName)}&headers=1&range=A1:AZ50000&_=${Date.now()}`;
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
        const results = await Promise.allSettled(SHEET_TABS.map(tab => fetchTab(tab)));
        _allData = [];
        results.forEach(r => { if (r.status === 'fulfilled') _allData.push(...r.value); });
        if (_onDataChanged) _onDataChanged(_allData);
        return _allData;
    }

    // ---------- WRITE via Apps Script ----------
    // Refresh only one tab and merge into _allData
    async function refreshTab(tabName) {
        const newRows = await fetchTab(tabName);
        _allData = _allData.filter(d => d.type !== tabName).concat(newRows);
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
        try {
            await fetchAllData();
            return { isOk: true };
        } catch (err) { return { isOk: false, error: err.message }; }
    }

    async function refresh() { return fetchAllData(); }
    function destroy() { }
    function hasWriteAccess() { return !!_scriptUrl; }

    async function debugTab(tabName) {
        const url = `https://docs.google.com/spreadsheets/d/${_spreadsheetId}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(tabName)}&headers=1`;
        try {
            const resp = await fetch(url); const text = await resp.text();
            const m = text.match(/google\.visualization\.Query\.setResponse\((.+)\);?$/s);
            if (!m) return { error: 'Cannot parse', raw: text.substring(0, 500) };
            const json = JSON.parse(m[1]);
            return { status: json.status, cols: json.table?.cols, rowCount: (json.table?.rows || []).length, firstRow: json.table?.rows?.[0]?.c };
        } catch (err) { return { error: err.message }; }
    }

    return {
        getStoredConfig, storeConfig, clearConfig, extractSheetId,
        init, refresh, destroy, debugTab, hasWriteAccess,
        create, update, updateMany, delete: remove, SHEET_TABS
    };
})();
