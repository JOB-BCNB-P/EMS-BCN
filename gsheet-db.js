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
  let _refreshTimer = null;
  let _allData = [];

  const SHEET_TABS = [
    'student', 'teacher', 'subject', 'schedule',
    'grade', 'eng_result', 'eval_form', 'evaluation', 'leave',
    'tracking', 'grade_tracking', 'announcement', 'user', 'doc_request'
  ];

  const REFRESH_INTERVAL = 60000;

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
    const url = `https://docs.google.com/spreadsheets/d/${_spreadsheetId}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(tabName)}&headers=1`;
    try {
      const resp = await fetch(url);
      if (!resp.ok) return [];
      let text = await resp.text();
      const jsonStr = text.match(/google\.visualization\.Query\.setResponse\((.+)\);?$/s);
      if (!jsonStr) return [];
      const json = JSON.parse(jsonStr[1]);
      if (json.status === 'error') return [];
      if (!json.table || !json.table.cols || !json.table.rows) return [];

      let cols = json.table.cols.map(c => (c.label || '').trim());
      if (cols.every(c => c === '')) {
        const fr = json.table.rows[0];
        if (fr && fr.c) { cols = fr.c.map(cell => cell ? String(cell.v || '').trim() : ''); json.table.rows.shift(); }
      }

      return (json.table.rows || []).map((row, idx) => {
        const obj = { type: tabName, __backendId: `${tabName}_${idx}`, __rowIndex: idx + 2 };
        row.c.forEach((cell, i) => {
          if (cols[i]) {
            let val = cell ? (cell.v !== null && cell.v !== undefined ? String(cell.v) : '') : '';
            if (val.match(/^\d+\.0$/)) val = val.replace('.0', '');
            obj[cols[i]] = val;
          }
        });
        return obj;
      }).filter(obj => Object.entries(obj).some(([k, v]) => !['type','__backendId','__rowIndex'].includes(k) && v !== ''));
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
  async function _callScript(params) {
    if (!_scriptUrl) return { isOk: false, error: 'ยังไม่ได้ตั้งค่า Apps Script URL — ไปที่ตั้งค่าระบบเพื่อกรอก URL' };
    try {
      if (typeof showLoading === 'function') showLoading(params.action === 'delete' ? 'กำลังลบข้อมูล...' : params.action === 'update' ? 'กำลังอัปเดต...' : 'กำลังบันทึก...');
      const qs = new URLSearchParams({ action: params.action });
      if (params.sheet) qs.set('sheet', params.sheet);
      if (params.rowIndex) qs.set('rowIndex', String(params.rowIndex));
      const resp = await fetch(_scriptUrl + '?' + qs.toString(), {
        method: 'POST', redirect: 'follow',
        headers: { 'Content-Type': 'text/plain' },
        body: JSON.stringify(params.data || {})
      });
      const result = await resp.json();
      if (result.isOk) {
        await fetchAllData();
      }
      if (typeof hideLoading === 'function') hideLoading();
      return result;
    } catch (err) {
      if (typeof hideLoading === 'function') hideLoading();
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
      if (_refreshTimer) clearInterval(_refreshTimer);
      _refreshTimer = setInterval(fetchAllData, REFRESH_INTERVAL);
      return { isOk: true };
    } catch (err) { return { isOk: false, error: err.message }; }
  }

  async function refresh() { return fetchAllData(); }
  function destroy() { if (_refreshTimer) { clearInterval(_refreshTimer); _refreshTimer = null; } }
  function hasWriteAccess() { return !!_scriptUrl; }

  async function debugTab(tabName) {
    const url = `https://docs.google.com/spreadsheets/d/${_spreadsheetId}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(tabName)}&headers=1`;
    try {
      const resp = await fetch(url); const text = await resp.text();
      const m = text.match(/google\.visualization\.Query\.setResponse\((.+)\);?$/s);
      if (!m) return { error: 'Cannot parse', raw: text.substring(0, 500) };
      const json = JSON.parse(m[1]);
      return { status: json.status, cols: json.table?.cols, rowCount: (json.table?.rows||[]).length, firstRow: json.table?.rows?.[0]?.c };
    } catch (err) { return { error: err.message }; }
  }

  return {
    getStoredConfig, storeConfig, clearConfig, extractSheetId,
    init, refresh, destroy, debugTab, hasWriteAccess,
    create, update, delete: remove, SHEET_TABS
  };
})();
