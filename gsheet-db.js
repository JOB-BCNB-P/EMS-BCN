// ================================================================
// gsheet-db.js — Google Sheets read-only data layer
// Replaces FirebaseDB (Firebase) with public Google Sheets API
//
// HOW IT WORKS:
// 1. User creates a Google Sheet with tabs matching data types
// 2. Sheet is published as public (Anyone with link → Viewer)
// 3. This module fetches data from each tab via Google Sheets JSON API
// 4. Data auto-refreshes every 60 seconds
// ================================================================

const GSheetDB = (() => {
  let _spreadsheetId = null;
  let _onDataChanged = null;
  let _refreshTimer = null;
  let _allData = [];

  // Tab names in Google Sheet → type field in app
  // Each tab = one data type. Tab name must match exactly.
  const SHEET_TABS = [
    'student', 'teacher', 'subject', 'schedule',
    'grade', 'eng_result', 'evaluation', 'leave',
    'tracking', 'grade_tracking', 'announcement', 'user', 'doc_request'
  ];

  const REFRESH_INTERVAL = 60000; // 60 seconds

  // ---------- Config ----------
  function getStoredConfig() {
    try {
      const raw = localStorage.getItem('gsheetConfig');
      return raw ? JSON.parse(raw) : null;
    } catch { return null; }
  }

  function storeConfig(config) {
    localStorage.setItem('gsheetConfig', JSON.stringify(config));
  }

  function clearConfig() {
    localStorage.removeItem('gsheetConfig');
  }

  // ---------- Extract Spreadsheet ID from URL or ID ----------
  function extractSheetId(input) {
    if (!input) return null;
    input = input.trim();
    // Full URL: https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/...
    const match = input.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
    if (match) return match[1];
    // Just the ID
    if (/^[a-zA-Z0-9_-]{20,}$/.test(input)) return input;
    return null;
  }

  // ---------- Fetch one tab ----------
  async function fetchTab(tabName) {
    // Use Google Sheets API v4 public endpoint (no API key needed for public sheets)
    const url = `https://docs.google.com/spreadsheets/d/${_spreadsheetId}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(tabName)}`;
    try {
      const resp = await fetch(url);
      if (!resp.ok) { console.warn(`Tab "${tabName}": HTTP ${resp.status}`); return []; }
      let text = await resp.text();
      // Response is wrapped in google.visualization.Query.setResponse({...})
      const jsonStr = text.match(/google\.visualization\.Query\.setResponse\((.+)\);?$/s);
      if (!jsonStr) { console.warn(`Tab "${tabName}": cannot parse gviz response`); return []; }
      const json = JSON.parse(jsonStr[1]);

      // Check for errors from Google
      if (json.status === 'error') {
        console.warn(`Tab "${tabName}": gviz error`, json.errors);
        return [];
      }

      if (!json.table || !json.table.cols || !json.table.rows) return [];

      // Get column headers — prefer label, fallback to id
      let cols = json.table.cols.map(c => (c.label || '').trim());

      // If all labels are empty, the first row IS the header — use first data row as headers
      if (cols.every(c => c === '')) {
        const firstRow = json.table.rows[0];
        if (firstRow && firstRow.c) {
          cols = firstRow.c.map(cell => cell ? String(cell.v || '').trim() : '');
          // Remove first row from data since it's headers
          json.table.rows.shift();
        }
      }

      const rows = json.table.rows || [];

      return rows.map((row, idx) => {
        const obj = { type: tabName, __backendId: `${tabName}_${idx}` };
        row.c.forEach((cell, i) => {
          if (cols[i]) {
            let val = cell ? (cell.v !== null && cell.v !== undefined ? String(cell.v) : '') : '';
            // Google Sheets sends numbers like 123456.0 — strip trailing .0
            if (val.match(/^\d+\.0$/)) val = val.replace('.0', '');
            obj[cols[i]] = val;
          }
        });
        return obj;
      }).filter(obj => {
        // Skip rows where all fields (except type/__backendId) are empty
        return Object.entries(obj).some(([k, v]) => k !== 'type' && k !== '__backendId' && v !== '');
      });
    } catch (err) {
      console.warn(`Failed to fetch tab "${tabName}":`, err.message);
      return [];
    }
  }

  // ---------- Debug: fetch raw response for a tab ----------
  async function debugTab(tabName) {
    const url = `https://docs.google.com/spreadsheets/d/${_spreadsheetId}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(tabName)}`;
    try {
      const resp = await fetch(url);
      const text = await resp.text();
      const jsonStr = text.match(/google\.visualization\.Query\.setResponse\((.+)\);?$/s);
      if (!jsonStr) return { error: 'Cannot parse response', raw: text.substring(0, 500) };
      const json = JSON.parse(jsonStr[1]);
      return {
        status: json.status,
        errors: json.errors,
        cols: json.table ? json.table.cols : null,
        rowCount: json.table ? (json.table.rows || []).length : 0,
        firstRow: json.table && json.table.rows && json.table.rows[0] ? json.table.rows[0].c : null
      };
    } catch (err) { return { error: err.message }; }
  }

  // ---------- Fetch all tabs ----------
  async function fetchAllData() {
    const results = await Promise.allSettled(
      SHEET_TABS.map(tab => fetchTab(tab))
    );
    _allData = [];
    results.forEach(r => {
      if (r.status === 'fulfilled') _allData.push(...r.value);
    });
    if (_onDataChanged) _onDataChanged(_allData);
    return _allData;
  }

  // ---------- Init ----------
  async function init(config, onDataChanged) {
    _onDataChanged = onDataChanged;
    _spreadsheetId = extractSheetId(config.spreadsheetId || config);

    if (!_spreadsheetId) {
      return { isOk: false, error: 'Invalid Spreadsheet ID or URL' };
    }

    try {
      await fetchAllData();
      // Auto-refresh
      if (_refreshTimer) clearInterval(_refreshTimer);
      _refreshTimer = setInterval(fetchAllData, REFRESH_INTERVAL);
      return { isOk: true };
    } catch (err) {
      return { isOk: false, error: err.message };
    }
  }

  // ---------- Manual refresh ----------
  async function refresh() {
    return fetchAllData();
  }

  // ---------- Cleanup ----------
  function destroy() {
    if (_refreshTimer) { clearInterval(_refreshTimer); _refreshTimer = null; }
  }

  return {
    getStoredConfig, storeConfig, clearConfig, extractSheetId,
    init, refresh, destroy, debugTab, SHEET_TABS
  };
})();
