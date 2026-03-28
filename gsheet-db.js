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
    'tracking', 'announcement', 'user', 'doc_request'
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
      if (!resp.ok) return [];
      let text = await resp.text();
      // Response is wrapped in google.visualization.Query.setResponse({...})
      const jsonStr = text.match(/google\.visualization\.Query\.setResponse\((.+)\);?$/s);
      if (!jsonStr) return [];
      const json = JSON.parse(jsonStr[1]);
      if (!json.table || !json.table.cols || !json.table.rows) return [];

      const cols = json.table.cols.map(c => c.label || '');
      const rows = json.table.rows || [];

      return rows.map((row, idx) => {
        const obj = { type: tabName, __backendId: `${tabName}_${idx}` };
        row.c.forEach((cell, i) => {
          if (cols[i]) {
            obj[cols[i]] = cell ? (cell.v !== null && cell.v !== undefined ? String(cell.v) : '') : '';
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
    init, refresh, destroy, SHEET_TABS
  };
})();
