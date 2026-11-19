(function () {
  'use strict';

  if (typeof window === 'undefined' || !window.__ModuleLoader__ || typeof window.__ModuleLoader__.load !== 'function') {
    return;
  }

  var PLUGIN_ID = 'deepseek-office-plugin';
  var PANEL_ID = 'excel-studio';
  var STYLE_ID = 'dsh-excel-studio-styles';

  var CSS_TEXT = [
    '.dsh-xs-root {',
    '  display: flex;',
    '  flex-direction: column;',
    '  height: 100%;',
    '  width: 100%;',
    '  min-height: 0;',
    '  min-width: 0;',
    '  overflow: hidden;',
    '  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;',
    '  font-size: 13px;',
    '  color: var(--dsw-alias-text-primary, #1f2328);',
    '  background: var(--dsw-alias-bg-primary, #ffffff);',
    '}',
    '.dsh-xs-topbar {',
    '  display: flex;',
    '  align-items: center;',
    '  gap: 8px;',
    '  padding: 10px 14px;',
    '  border-bottom: 1px solid var(--dsw-alias-border-subtle, rgba(127,127,127,0.22));',
    '  background: var(--dsw-alias-bg-secondary, rgba(127,127,127,0.05));',
    '  flex-wrap: wrap;',
    '}',
    '.dsh-xs-brand {',
    '  display: inline-flex;',
    '  align-items: center;',
    '  gap: 6px;',
    '  font-weight: 600;',
    '  font-size: 14px;',
    '  margin-right: 4px;',
    '}',
    '.dsh-xs-badge {',
    '  font-size: 11px;',
    '  padding: 1px 6px;',
    '  border-radius: 999px;',
    '  background: rgba(16, 124, 65, 0.14);',
    '  color: #107c41;',
    '  font-weight: 600;',
    '}',
    '.dsh-xs-select, .dsh-xs-input {',
    '  height: 28px;',
    '  padding: 0 8px;',
    '  border-radius: 6px;',
    '  border: 1px solid var(--dsw-alias-border-default, rgba(127,127,127,0.3));',
    '  background: var(--dsw-alias-bg-primary, #ffffff);',
    '  color: inherit;',
    '  font-size: 12px;',
    '  outline: none;',
    '}',
    '.dsh-xs-select:focus, .dsh-xs-input:focus {',
    '  border-color: #107c41;',
    '}',
    '.dsh-xs-btn {',
    '  height: 28px;',
    '  padding: 0 10px;',
    '  border-radius: 6px;',
    '  border: 1px solid var(--dsw-alias-border-default, rgba(127,127,127,0.28));',
    '  background: var(--dsw-alias-bg-primary, transparent);',
    '  color: inherit;',
    '  font-size: 12px;',
    '  font-weight: 500;',
    '  cursor: pointer;',
    '  display: inline-flex;',
    '  align-items: center;',
    '  gap: 5px;',
    '  white-space: nowrap;',
    '}',
    '.dsh-xs-btn:hover:not(:disabled) {',
    '  background: var(--dsw-alias-bg-hover, rgba(127,127,127,0.12));',
    '}',
    '.dsh-xs-btn:disabled {',
    '  opacity: 0.45;',
    '  cursor: not-allowed;',
    '}',
    '.dsh-xs-btn-primary {',
    '  background: #107c41;',
    '  border-color: #107c41;',
    '  color: #ffffff;',
    '}',
    '.dsh-xs-btn-primary:hover:not(:disabled) {',
    '  background: #0d6635;',
    '}',
    '.dsh-xs-btn-active {',
    '  background: rgba(16, 124, 65, 0.16);',
    '  border-color: #107c41;',
    '  color: #107c41;',
    '}',
    '.dsh-xs-btn-danger {',
    '  color: #cf222e;',
    '}',
    '.dsh-xs-ribbon {',
    '  display: flex;',
    '  align-items: center;',
    '  gap: 6px;',
    '  padding: 6px 14px;',
    '  border-bottom: 1px solid var(--dsw-alias-border-subtle, rgba(127,127,127,0.2));',
    '  background: var(--dsw-alias-bg-primary, transparent);',
    '  flex-wrap: wrap;',
    '}',
    '.dsh-xs-sep {',
    '  width: 1px;',
    '  height: 18px;',
    '  background: var(--dsw-alias-border-subtle, rgba(127,127,127,0.25));',
    '  margin: 0 2px;',
    '}',
    '.dsh-xs-swatch {',
    '  width: 18px;',
    '  height: 18px;',
    '  border-radius: 4px;',
    '  border: 1px solid rgba(127,127,127,0.4);',
    '  cursor: pointer;',
    '  padding: 0;',
    '}',
    '.dsh-xs-formulabar {',
    '  display: flex;',
    '  align-items: center;',
    '  gap: 8px;',
    '  padding: 6px 14px;',
    '  border-bottom: 1px solid var(--dsw-alias-border-subtle, rgba(127,127,127,0.2));',
    '  background: var(--dsw-alias-bg-secondary, rgba(127,127,127,0.03));',
    '}',
    '.dsh-xs-coord {',
    '  min-width: 56px;',
    '  text-align: center;',
    '  font-family: ui-monospace, SFMono-Regular, Menlo, monospace;',
    '  font-weight: 600;',
    '  font-size: 12px;',
    '  padding: 3px 6px;',
    '  border-radius: 4px;',
    '  border: 1px solid var(--dsw-alias-border-subtle, rgba(127,127,127,0.25));',
    '}',
    '.dsh-xs-fx {',
    '  font-style: italic;',
    '  font-family: Georgia, serif;',
    '  opacity: 0.65;',
    '  font-weight: 600;',
    '}',
    '.dsh-xs-grid-wrap {',
    '  flex: 1;',
    '  min-height: 0;',
    '  overflow: auto;',
    '  position: relative;',
    '}',
    '.dsh-xs-table {',
    '  border-collapse: separate;',
    '  border-spacing: 0;',
    '  table-layout: fixed;',
    '  width: max-content;',
    '  min-width: 100%;',
    '}',
    '.dsh-xs-th {',
    '  position: sticky;',
    '  top: 0;',
    '  z-index: 2;',
    '  background: var(--dsw-alias-bg-secondary, #f4f5f7);',
    '  color: var(--dsw-alias-text-secondary, #57606a);',
    '  font-weight: 600;',
    '  font-size: 11px;',
    '  text-align: center;',
    '  height: 24px;',
    '  min-width: 96px;',
    '  width: 104px;',
    '  border-right: 1px solid var(--dsw-alias-border-subtle, rgba(127,127,127,0.25));',
    '  border-bottom: 1px solid var(--dsw-alias-border-subtle, rgba(127,127,127,0.3));',
    '  user-select: none;',
    '}',
    '.dsh-xs-corner {',
    '  position: sticky;',
    '  top: 0;',
    '  left: 0;',
    '  z-index: 4;',
    '  width: 44px;',
    '  min-width: 44px;',
    '  background: var(--dsw-alias-bg-secondary, #f4f5f7);',
    '  border-right: 1px solid var(--dsw-alias-border-subtle, rgba(127,127,127,0.3));',
    '  border-bottom: 1px solid var(--dsw-alias-border-subtle, rgba(127,127,127,0.3));',
    '}',
    '.dsh-xs-rowhdr {',
    '  position: sticky;',
    '  left: 0;',
    '  z-index: 1;',
    '  width: 44px;',
    '  min-width: 44px;',
    '  text-align: center;',
    '  font-size: 11px;',
    '  color: var(--dsw-alias-text-secondary, #57606a);',
    '  background: var(--dsw-alias-bg-secondary, #f4f5f7);',
    '  border-right: 1px solid var(--dsw-alias-border-subtle, rgba(127,127,127,0.28));',
    '  border-bottom: 1px solid var(--dsw-alias-border-subtle, rgba(127,127,127,0.2));',
    '  user-select: none;',
    '}',
    '.dsh-xs-cell {',
    '  height: 25px;',
    '  padding: 0 6px;',
    '  border-right: 1px solid var(--dsw-alias-border-subtle, rgba(127,127,127,0.18));',
    '  border-bottom: 1px solid var(--dsw-alias-border-subtle, rgba(127,127,127,0.18));',
    '  white-space: nowrap;',
    '  overflow: hidden;',
    '  text-overflow: ellipsis;',
    '  cursor: cell;',
    '  position: relative;',
    '  vertical-align: middle;',
    '}',
    '.dsh-xs-cell-selected {',
    '  outline: 2px solid #107c41;',
    '  outline-offset: -2px;',
    '  z-index: 1;',
    '}',
    '.dsh-xs-cell-dirty::after {',
    '  content: "";',
    '  position: absolute;',
    '  top: 2px;',
    '  right: 2px;',
    '  width: 5px;',
    '  height: 5px;',
    '  border-radius: 50%;',
    '  background: #d97706;',
    '}',
    '.dsh-xs-cell-editor {',
    '  width: 100%;',
    '  height: 100%;',
    '  border: none;',
    '  outline: none;',
    '  padding: 0;',
    '  margin: 0;',
    '  background: transparent;',
    '  color: inherit;',
    '  font: inherit;',
    '}',
    '.dsh-xs-tabs {',
    '  display: flex;',
    '  align-items: center;',
    '  gap: 4px;',
    '  padding: 6px 12px;',
    '  border-top: 1px solid var(--dsw-alias-border-subtle, rgba(127,127,127,0.22));',
    '  background: var(--dsw-alias-bg-secondary, rgba(127,127,127,0.05));',
    '  overflow-x: auto;',
    '}',
    '.dsh-xs-tab {',
    '  padding: 4px 12px;',
    '  border-radius: 6px;',
    '  border: 1px solid transparent;',
    '  background: transparent;',
    '  color: inherit;',
    '  font-size: 12px;',
    '  cursor: pointer;',
    '  white-space: nowrap;',
    '}',
    '.dsh-xs-tab-active {',
    '  background: var(--dsw-alias-bg-primary, #ffffff);',
    '  border-color: var(--dsw-alias-border-default, rgba(127,127,127,0.3));',
    '  font-weight: 600;',
    '  color: #107c41;',
    '}',
    '.dsh-xs-status {',
    '  display: flex;',
    '  align-items: center;',
    '  justify-content: space-between;',
    '  gap: 12px;',
    '  padding: 5px 14px;',
    '  font-size: 11px;',
    '  opacity: 0.85;',
    '  border-top: 1px solid var(--dsw-alias-border-subtle, rgba(127,127,127,0.16));',
    '}',
    '.dsh-xs-banner {',
    '  padding: 8px 14px;',
    '  font-size: 12px;',
    '  display: flex;',
    '  align-items: center;',
    '  justify-content: space-between;',
    '  gap: 8px;',
    '}',
    '.dsh-xs-banner-err {',
    '  background: rgba(207, 34, 46, 0.1);',
    '  color: #cf222e;',
    '  border-bottom: 1px solid rgba(207, 34, 46, 0.25);',
    '}',
    '.dsh-xs-banner-info {',
    '  background: rgba(16, 124, 65, 0.1);',
    '  color: #107c41;',
    '  border-bottom: 1px solid rgba(16, 124, 65, 0.22);',
    '}',
    '.dsh-xs-empty {',
    '  flex: 1;',
    '  display: flex;',
    '  flex-direction: column;',
    '  align-items: center;',
    '  justify-content: center;',
    '  gap: 12px;',
    '  padding: 32px;',
    '  text-align: center;',
    '  opacity: 0.85;',
    '}',
    '.dsh-xs-summary {',
    '  padding: 10px 14px;',
    '  border-bottom: 1px solid var(--dsw-alias-border-subtle, rgba(127,127,127,0.22));',
    '  background: var(--dsw-alias-bg-secondary, rgba(127,127,127,0.04));',
    '  max-height: 180px;',
    '  overflow: auto;',
    '}',
  ].join('\n');

  function ensureStyles() {
    if (typeof document === 'undefined' || !document.head) return;
    if (document.getElementById(STYLE_ID)) return;
    var style = document.createElement('style');
    style.id = STYLE_ID;
    style.textContent = CSS_TEXT;
    document.head.appendChild(style);
  }

  function colIndexToLetter(index) {
    var col = '';
    var n = index;
    while (n > 0) {
      var rem = (n - 1) % 26;
      col = String.fromCharCode(65 + rem) + col;
      n = Math.floor((n - 1) / 26);
    }
    return col;
  }

  function makeCoord(row, col) {
    return colIndexToLetter(col) + String(row);
  }

  async function apiRequest(url, options) {
    var res = await fetch(url, Object.assign({ credentials: 'same-origin' }, options || {}));
    var data = null;
    try {
      data = await res.json();
    } catch (_e) {
      data = { ok: false, error: 'HTTP ' + res.status };
    }
    if (!res.ok || !data || !data.ok) {
      throw new Error((data && data.error) || 'Request failed (HTTP ' + res.status + ')');
    }
    return data.data;
  }

  try {
    window.__ModuleLoader__.load({
      id: PLUGIN_ID,
      factory: function (require) {
        var module = { exports: {} };
        var exports = module.exports;
        var React = require('react');
        var h = React.createElement;

        function ExcelStudioIcon(props) {
          var size = (props && props.size) || 18;
          return h(
            'svg',
            {
              width: size,
              height: size,
              viewBox: '0 0 24 24',
              fill: 'none',
              stroke: 'currentColor',
              strokeWidth: '1.75',
              strokeLinecap: 'round',
              strokeLinejoin: 'round',
              'aria-hidden': 'true',
            },
            h('rect', { x: '3', y: '3', width: '18', height: '18', rx: '2.5' }),
            h('line', { x1: '3', y1: '9', x2: '21', y2: '9' }),
            h('line', { x1: '3', y1: '15', x2: '21', y2: '15' }),
            h('line', { x1: '9', y1: '9', x2: '9', y2: '21' }),
            h('line', { x1: '15', y1: '9', x2: '15', y2: '21' })
          );
        }

        function ExcelStudioPanel() {
          var _useState = React.useState;
          var _useEffect = React.useEffect;
          var _useCallback = React.useCallback;

          var workbooksState = _useState([]);
          var workbooks = workbooksState[0];
          var setWorkbooks = workbooksState[1];

          var allowedDirsState = _useState([]);
          var allowedDirs = allowedDirsState[0];
          var setAllowedDirs = allowedDirsState[1];

          var sheetState = _useState(null);
          var sheet = sheetState[0];
          var setSheet = sheetState[1];

          var loadingState = _useState(false);
          var loading = loadingState[0];
          var setLoading = loadingState[1];

          var savingState = _useState(false);
          var saving = savingState[0];
          var setSaving = savingState[1];

          var errorState = _useState('');
          var error = errorState[0];
          var setError = errorState[1];

          var noticeState = _useState('');
          var notice = noticeState[0];
          var setNotice = noticeState[1];

          var selectedState = _useState({ row: 1, col: 1 });
          var selected = selectedState[0];
          var setSelected = selectedState[1];

          var editingState = _useState(false);
          var editing = editingState[0];
          var setEditing = editingState[1];

          var draftValueState = _useState('');
          var draftValue = draftValueState[0];
          var setDraftValue = draftValueState[1];

          var pendingEditsState = _useState({});
          var pendingEdits = pendingEditsState[0];
          var setPendingEdits = pendingEditsState[1];

          var pendingFormatsState = _useState({});
          var pendingFormats = pendingFormatsState[0];
          var setPendingFormats = pendingFormatsState[1];

          var showNewFormState = _useState(false);
          var showNewForm = showNewFormState[0];
          var setShowNewForm = showNewFormState[1];

          var newFilenameState = _useState('untitled.xlsx');
          var newFilename = newFilenameState[0];
          var setNewFilename = newFilenameState[1];

          var openPathInputState = _useState('');
          var openPathInput = openPathInputState[0];
          var setOpenPathInput = openPathInputState[1];

          var summaryState = _useState(null);
          var summary = summaryState[0];
          var setSummary = summaryState[1];

          _useEffect(function () {
            ensureStyles();
          }, []);

          var dirtyCount = Object.keys(pendingEdits).length + Object.keys(pendingFormats).length;
          var selectedCoord = makeCoord(selected.row, selected.col);

          var getCellRaw = function (coord) {
            if (Object.prototype.hasOwnProperty.call(pendingEdits, coord)) {
              return pendingEdits[coord];
            }
            var cell = sheet && sheet.cells && sheet.cells[coord];
            if (!cell) return '';
            return cell.formula !== undefined ? cell.formula : (cell.raw !== undefined ? String(cell.raw) : '');
          };

          var getCellDisplay = function (coord) {
            if (Object.prototype.hasOwnProperty.call(pendingEdits, coord)) {
              return pendingEdits[coord];
            }
            var cell = sheet && sheet.cells && sheet.cells[coord];
            if (!cell) return '';
            return cell.display !== undefined ? String(cell.display) : '';
          };

          var getCellStyle = function (coord) {
            var base = (sheet && sheet.cells && sheet.cells[coord] && sheet.cells[coord].style) || {};
            var patch = pendingFormats[coord] || {};
            return Object.assign({}, base, patch);
          };

          var openWorkbook = _useCallback(function (filePath, sheetName) {
            if (!filePath) return;
            setLoading(true);
            setError('');
            setSummary(null);
            apiRequest('/api/excel-studio/open', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ file_path: filePath, sheet_name: sheetName }),
            })
              .then(function (data) {
                setSheet(data);
                setPendingEdits({});
                setPendingFormats({});
                setSelected({ row: 1, col: 1 });
                setEditing(false);
                var firstCell = data.cells && data.cells.A1;
                setDraftValue(firstCell ? (firstCell.formula || String(firstCell.raw || '')) : '');
                setOpenPathInput(data.file_path || filePath);
              })
              .catch(function (err) {
                setError(err.message || String(err));
              })
              .finally(function () {
                setLoading(false);
              });
          }, []);

          var refreshWorkbooks = _useCallback(function (autoOpenFirst) {
            setLoading(true);
            setError('');
            apiRequest('/api/excel-studio/workbooks', { method: 'GET' })
              .then(function (data) {
                var list = (data && data.workbooks) || [];
                setWorkbooks(list);
                setAllowedDirs((data && data.allowed_directories) || []);
                if (autoOpenFirst && list.length > 0) {
                  openWorkbook(list[0].path);
                }
              })
              .catch(function (err) {
                setError(err.message || String(err));
              })
              .finally(function () {
                setLoading(false);
              });
          }, [openWorkbook]);

          _useEffect(function () {
            refreshWorkbooks(true);
          }, [refreshWorkbooks]);

          var selectCell = function (row, col, startEdit) {
            var maxR = (sheet && sheet.row_count) || 40;
            var maxC = (sheet && sheet.col_count) || 16;
            var r = Math.max(1, Math.min(maxR, row));
            var c = Math.max(1, Math.min(maxC, col));
            var coord = makeCoord(r, c);
            setSelected({ row: r, col: c });
            setDraftValue(getCellRaw(coord));
            setEditing(Boolean(startEdit));
          };

          var commitDraft = function (coord, value) {
            var prev = (sheet && sheet.cells && sheet.cells[coord])
              ? (sheet.cells[coord].formula !== undefined ? sheet.cells[coord].formula : String(sheet.cells[coord].raw || ''))
              : '';
            setPendingEdits(function (cur) {
              var next = Object.assign({}, cur);
              if (String(value) === String(prev)) {
                delete next[coord];
              } else {
                next[coord] = value;
              }
              return next;
            });
          };

          var handleSave = _useCallback(function () {
            if (!sheet || saving) return;
            var editList = Object.keys(pendingEdits).map(function (coord) {
              return { cell: coord, value: pendingEdits[coord] };
            });
            var formatList = Object.keys(pendingFormats).map(function (coord) {
              var f = pendingFormats[coord];
              return Object.assign({ cell: coord }, f);
            });
            if (editList.length === 0 && formatList.length === 0) {
              setNotice('No unsaved changes.');
              return;
            }
            setSaving(true);
            setError('');
            setNotice('');
            apiRequest('/api/excel-studio/save', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({
                file_path: sheet.file_path,
                sheet_name: sheet.active_sheet,
                edits: editList,
                formats: formatList,
              }),
            })
              .then(function (data) {
                setSheet(data);
                setPendingEdits({});
                setPendingFormats({});
                setEditing(false);
                var msg = 'Saved ' + editList.length + ' cell edit(s)';
                if (formatList.length > 0) msg += ' & ' + formatList.length + ' format(s)';
                if (data.backup_path) msg += ' (backup created)';
                setNotice(msg + '.');
              })
              .catch(function (err) {
                setError(err.message || String(err));
              })
              .finally(function () {
                setSaving(false);
              });
          }, [sheet, saving, pendingEdits, pendingFormats]);

          _useEffect(function () {
            var onKeyDown = function (e) {
              if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
                if (sheet) {
                  e.preventDefault();
                  handleSave();
                }
              }
            };
            window.addEventListener('keydown', onKeyDown);
            return function () { window.removeEventListener('keydown', onKeyDown); };
          }, [sheet, handleSave]);

          var patchSelectedFormat = function (patch) {
            if (!sheet) return;
            setPendingFormats(function (cur) {
              var next = Object.assign({}, cur);
              next[selectedCoord] = Object.assign({}, cur[selectedCoord] || {}, patch);
              return next;
            });
          };

          var runSheetOp = function (op, extra) {
            if (!sheet) return;
            setSaving(true);
            setError('');
            setNotice('');
            var body = Object.assign(
              { file_path: sheet.file_path, sheet_name: sheet.active_sheet, op: op },
              extra || {}
            );
            apiRequest('/api/excel-studio/sheet-op', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify(body),
            })
              .then(function (data) {
                setSheet(data);
                setPendingEdits({});
                setPendingFormats({});
                setNotice('Completed ' + op.replace(/_/g, ' ') + (data.backup_path ? ' (backup created).' : '.'));
              })
              .catch(function (err) {
                setError(err.message || String(err));
              })
              .finally(function () {
                setSaving(false);
              });
          };

          var handleCreateWorkbook = function (e) {
            e.preventDefault();
            var nameTrim = newFilename.trim();
            if (!nameTrim) return;
            setSaving(true);
            setError('');
            apiRequest('/api/excel-studio/create', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({
                directory: allowedDirs[0] || undefined,
                filename: nameTrim,
              }),
            })
              .then(function (data) {
                setSheet(data);
                setPendingEdits({});
                setPendingFormats({});
                setShowNewForm(false);
                setNotice('Created ' + data.filename);
                refreshWorkbooks(false);
              })
              .catch(function (err) {
                setError(err.message || String(err));
              })
              .finally(function () {
                setSaving(false);
              });
          };

          var handleSummary = function () {
            if (!sheet) return;
            setError('');
            apiRequest('/api/excel-studio/summary', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ file_path: sheet.file_path, sheet_name: sheet.active_sheet }),
            })
              .then(function (data) {
                setSummary(data);
              })
              .catch(function (err) {
                setError(err.message || String(err));
              });
          };

          var currentStyle = getCellStyle(selectedCoord);
          var rowCount = (sheet && sheet.row_count) || 40;
          var colCount = (sheet && sheet.col_count) || 16;

          var colIndices = [];
          for (var ci = 1; ci <= colCount; ci++) colIndices.push(ci);
          var rowIndices = [];
          for (var ri = 1; ri <= rowCount; ri++) rowIndices.push(ri);

          var fillSwatches = [
            { label: 'Clear fill', value: '', color: 'transparent' },
            { label: 'Soft yellow', value: '#FFF3CD', color: '#FFF3CD' },
            { label: 'Soft green', value: '#D1E7DD', color: '#D1E7DD' },
            { label: 'Soft blue', value: '#CFE2FF', color: '#CFE2FF' },
            { label: 'Soft red', value: '#F8D7DA', color: '#F8D7DA' },
            { label: 'Header navy', value: '#1F4E79', color: '#1F4E79' },
          ];

          return h(
            'div',
            { className: 'dsh-xs-root', 'data-testid': 'excel-studio-panel' },

            // Top bar
            h(
              'div',
              { className: 'dsh-xs-topbar' },
              h(
                'div',
                { className: 'dsh-xs-brand' },
                h(ExcelStudioIcon, { size: 17 }),
                h('span', null, 'Excel Studio'),
                h('span', { className: 'dsh-xs-badge' }, 'Local')
              ),
              h(
                'select',
                {
                  className: 'dsh-xs-select',
                  style: { minWidth: '210px', maxWidth: '300px' },
                  value: (sheet && sheet.file_path) || '',
                  onChange: function (e) {
                    if (e.target.value) openWorkbook(e.target.value);
                  },
                  'aria-label': 'Select workbook',
                },
                h('option', { value: '' }, workbooks.length ? 'Select a workbook...' : 'No workbooks found'),
                workbooks.map(function (wb) {
                  return h('option', { key: wb.path, value: wb.path }, wb.name + (wb.is_locked ? ' (locked)' : ''));
                })
              ),
              h(
                'button',
                {
                  type: 'button',
                  className: 'dsh-xs-btn',
                  onClick: function () { setShowNewForm(function (v) { return !v; }); },
                },
                '+ New'
              ),
              h(
                'button',
                {
                  type: 'button',
                  className: 'dsh-xs-btn',
                  disabled: loading,
                  onClick: function () {
                    refreshWorkbooks(false);
                    if (sheet) openWorkbook(sheet.file_path, sheet.active_sheet);
                  },
                },
                loading ? 'Loading...' : 'Refresh'
              ),
              h(
                'button',
                {
                  type: 'button',
                  className: 'dsh-xs-btn ' + (dirtyCount > 0 ? 'dsh-xs-btn-primary' : ''),
                  disabled: !sheet || saving || dirtyCount === 0,
                  onClick: handleSave,
                },
                saving ? 'Saving...' : (dirtyCount > 0 ? 'Save (' + dirtyCount + ')' : 'Saved')
              ),
              h(
                'form',
                {
                  style: { display: 'inline-flex', gap: '4px', marginLeft: 'auto' },
                  onSubmit: function (e) {
                    e.preventDefault();
                    if (openPathInput.trim()) openWorkbook(openPathInput.trim());
                  },
                },
                h('input', {
                  className: 'dsh-xs-input',
                  style: { width: '210px' },
                  placeholder: 'Open .xlsx path in allowed folder...',
                  value: openPathInput,
                  onChange: function (e) { setOpenPathInput(e.target.value); },
                }),
                h('button', { type: 'submit', className: 'dsh-xs-btn' }, 'Open')
              )
            ),

            // Optional New Workbook inline form
            showNewForm
              ? h(
                  'form',
                  {
                    className: 'dsh-xs-ribbon',
                    onSubmit: handleCreateWorkbook,
                  },
                  h('span', { style: { fontWeight: 600, fontSize: '12px' } }, 'New Workbook:'),
                  h('input', {
                    className: 'dsh-xs-input',
                    style: { width: '220px' },
                    value: newFilename,
                    onChange: function (e) { setNewFilename(e.target.value); },
                    placeholder: 'budget_2026.xlsx',
                  }),
                  h('span', { style: { fontSize: '11px', opacity: 0.7 } }, allowedDirs[0] ? 'in ' + allowedDirs[0] : ''),
                  h('button', { type: 'submit', className: 'dsh-xs-btn dsh-xs-btn-primary', disabled: saving }, 'Create'),
                  h(
                    'button',
                    { type: 'button', className: 'dsh-xs-btn', onClick: function () { setShowNewForm(false); } },
                    'Cancel'
                  )
                )
              : null,

            // Error / Notice banners
            error
              ? h(
                  'div',
                  { className: 'dsh-xs-banner dsh-xs-banner-err', role: 'alert' },
                  h('span', null, error),
                  h('button', { type: 'button', className: 'dsh-xs-btn', onClick: function () { setError(''); } }, 'Dismiss')
                )
              : null,
            notice
              ? h(
                  'div',
                  { className: 'dsh-xs-banner dsh-xs-banner-info' },
                  h('span', null, notice),
                  h('button', { type: 'button', className: 'dsh-xs-btn', onClick: function () { setNotice(''); } }, 'OK')
                )
              : null,

            // Formatting & Structural Ribbon
            sheet
              ? h(
                  'div',
                  { className: 'dsh-xs-ribbon' },
                  h(
                    'button',
                    {
                      type: 'button',
                      className: 'dsh-xs-btn ' + (currentStyle.bold ? 'dsh-xs-btn-active' : ''),
                      title: 'Bold',
                      onClick: function () { patchSelectedFormat({ bold: !currentStyle.bold }); },
                    },
                    h('strong', null, 'B')
                  ),
                  h(
                    'button',
                    {
                      type: 'button',
                      className: 'dsh-xs-btn ' + (currentStyle.italic ? 'dsh-xs-btn-active' : ''),
                      title: 'Italic',
                      onClick: function () { patchSelectedFormat({ italic: !currentStyle.italic }); },
                    },
                    h('em', null, 'I')
                  ),
                  h('span', { className: 'dsh-xs-sep' }),
                  ['left', 'center', 'right'].map(function (al) {
                    return h(
                      'button',
                      {
                        key: al,
                        type: 'button',
                        className: 'dsh-xs-btn ' + (currentStyle.align === al ? 'dsh-xs-btn-active' : ''),
                        title: 'Align ' + al,
                        onClick: function () { patchSelectedFormat({ align: al }); },
                      },
                      al[0].toUpperCase()
                    );
                  }),
                  h('span', { className: 'dsh-xs-sep' }),
                  h('span', { style: { fontSize: '11px', opacity: 0.7 } }, 'Fill:'),
                  fillSwatches.map(function (sw) {
                    return h('button', {
                      key: sw.label,
                      type: 'button',
                      title: sw.label,
                      className: 'dsh-xs-swatch',
                      style: {
                        background: sw.color,
                        backgroundImage: sw.value ? undefined : 'linear-gradient(135deg, transparent 45%, #cf222e 50%, transparent 55%)',
                      },
                      onClick: function () { patchSelectedFormat({ fill_color: sw.value }); },
                    });
                  }),
                  h('span', { className: 'dsh-xs-sep' }),
                  h(
                    'select',
                    {
                      className: 'dsh-xs-select',
                      title: 'Number format',
                      value: currentStyle.numberFormat || currentStyle.number_format || 'General',
                      onChange: function (e) { patchSelectedFormat({ number_format: e.target.value }); },
                    },
                    h('option', { value: 'General' }, 'General'),
                    h('option', { value: '#,##0' }, 'Number (#,##0)'),
                    h('option', { value: '$#,##0.00' }, 'Currency ($0.00)'),
                    h('option', { value: '0.0%' }, 'Percent (0.0%)'),
                    h('option', { value: 'yyyy-mm-dd' }, 'Date (YYYY-MM-DD)')
                  ),
                  h('span', { className: 'dsh-xs-sep' }),
                  h(
                    'button',
                    {
                      type: 'button',
                      className: 'dsh-xs-btn',
                      disabled: saving,
                      onClick: function () { runSheetOp('insert_rows', { index: selected.row, amount: 1 }); },
                    },
                    '+ Row'
                  ),
                  h(
                    'button',
                    {
                      type: 'button',
                      className: 'dsh-xs-btn dsh-xs-btn-danger',
                      disabled: saving,
                      onClick: function () {
                        if (window.confirm('Delete row ' + selected.row + '? A backup will be created first.')) {
                          runSheetOp('delete_rows', { index: selected.row, amount: 1 });
                        }
                      },
                    },
                    '- Row'
                  ),
                  h(
                    'button',
                    {
                      type: 'button',
                      className: 'dsh-xs-btn',
                      disabled: saving,
                      onClick: function () { runSheetOp('insert_columns', { index: selected.col, amount: 1 }); },
                    },
                    '+ Col'
                  ),
                  h(
                    'button',
                    {
                      type: 'button',
                      className: 'dsh-xs-btn dsh-xs-btn-danger',
                      disabled: saving,
                      onClick: function () {
                        if (window.confirm('Delete column ' + colIndexToLetter(selected.col) + '? A backup will be created first.')) {
                          runSheetOp('delete_columns', { index: selected.col, amount: 1 });
                        }
                      },
                    },
                    '- Col'
                  ),
                  h('span', { className: 'dsh-xs-sep' }),
                  h(
                    'button',
                    {
                      type: 'button',
                      className: 'dsh-xs-btn',
                      onClick: function () {
                        if (summary) setSummary(null);
                        else handleSummary();
                      },
                    },
                    summary ? 'Hide Stats' : 'Summary Stats'
                  )
                )
              : null,

            // Summary Stats drawer
            summary
              ? h(
                  'div',
                  { className: 'dsh-xs-summary' },
                  h(
                    'div',
                    { style: { display: 'flex', justifyContent: 'space-between', marginBottom: '6px', fontWeight: 600 } },
                    h('span', null, 'Dataset Summary (' + (summary.total_rows || 0) + ' rows)'),
                    h('button', { type: 'button', className: 'dsh-xs-btn', onClick: function () { setSummary(null); } }, 'Close')
                  ),
                  h(
                    'pre',
                    { style: { margin: 0, fontSize: '11px', whiteSpace: 'pre-wrap' } },
                    JSON.stringify(summary.numeric_summary || summary, null, 2)
                  )
                )
              : null,

            // Formula bar
            sheet
              ? h(
                  'div',
                  { className: 'dsh-xs-formulabar' },
                  h('span', { className: 'dsh-xs-coord' }, selectedCoord),
                  h('span', { className: 'dsh-xs-fx' }, 'fx'),
                  h('input', {
                    className: 'dsh-xs-input',
                    style: { flex: 1, fontFamily: 'ui-monospace, SFMono-Regular, Menlo, monospace' },
                    value: draftValue,
                    placeholder: 'Enter value or formula (e.g. =SUM(B2:B10))',
                    onChange: function (e) {
                      var val = e.target.value;
                      setDraftValue(val);
                      commitDraft(selectedCoord, val);
                    },
                    onKeyDown: function (e) {
                      if (e.key === 'Enter') {
                        e.preventDefault();
                        commitDraft(selectedCoord, draftValue);
                        selectCell(selected.row + 1, selected.col, false);
                      } else if (e.key === 'Tab') {
                        e.preventDefault();
                        commitDraft(selectedCoord, draftValue);
                        selectCell(selected.row, selected.col + 1, false);
                      } else if (e.key === 'Escape') {
                        setDraftValue(getCellRaw(selectedCoord));
                        setEditing(false);
                      }
                    },
                  })
                )
              : null,

            // Main Spreadsheet Grid or Empty State
            sheet
              ? h(
                  'div',
                  { className: 'dsh-xs-grid-wrap' },
                  h(
                    'table',
                    { className: 'dsh-xs-table' },
                    h(
                      'thead',
                      null,
                      h(
                        'tr',
                        null,
                        h('th', { className: 'dsh-xs-corner' }),
                        colIndices.map(function (c) {
                          var letter = colIndexToLetter(c);
                          var w = sheet.column_widths && sheet.column_widths[letter];
                          var px = w ? Math.max(70, Math.min(320, Math.round(w * 8))) : 104;
                          return h(
                            'th',
                            { key: letter, className: 'dsh-xs-th', style: { width: px + 'px', minWidth: px + 'px' } },
                            letter
                          );
                        })
                      )
                    ),
                    h(
                      'tbody',
                      null,
                      rowIndices.map(function (r) {
                        return h(
                          'tr',
                          { key: r },
                          h('td', { className: 'dsh-xs-rowhdr' }, r),
                          colIndices.map(function (c) {
                            var coord = makeCoord(r, c);
                            var isSel = selected.row === r && selected.col === c;
                            var isDirty =
                              Object.prototype.hasOwnProperty.call(pendingEdits, coord) ||
                              Object.prototype.hasOwnProperty.call(pendingFormats, coord);
                            var st = getCellStyle(coord);
                            var cellCss = {
                              fontWeight: st.bold ? '700' : undefined,
                              fontStyle: st.italic ? 'italic' : undefined,
                              color: st.font_color || st.color || undefined,
                              backgroundColor: st.fill_color || st.fill || undefined,
                              textAlign: st.align || undefined,
                            };
                            var cls =
                              'dsh-xs-cell' +
                              (isSel ? ' dsh-xs-cell-selected' : '') +
                              (isDirty ? ' dsh-xs-cell-dirty' : '');

                            return h(
                              'td',
                              {
                                key: coord,
                                className: cls,
                                style: cellCss,
                                onClick: function () {
                                  if (!isSel) selectCell(r, c, false);
                                },
                                onDoubleClick: function () {
                                  selectCell(r, c, true);
                                },
                              },
                              isSel && editing
                                ? h('input', {
                                    className: 'dsh-xs-cell-editor',
                                    autoFocus: true,
                                    value: draftValue,
                                    onChange: function (e) {
                                      setDraftValue(e.target.value);
                                      commitDraft(coord, e.target.value);
                                    },
                                    onBlur: function () {
                                      setEditing(false);
                                    },
                                    onKeyDown: function (e) {
                                      if (e.key === 'Enter') {
                                        e.preventDefault();
                                        commitDraft(coord, draftValue);
                                        selectCell(r + 1, c, false);
                                      } else if (e.key === 'Tab') {
                                        e.preventDefault();
                                        commitDraft(coord, draftValue);
                                        selectCell(r, c + 1, false);
                                      } else if (e.key === 'Escape') {
                                        setDraftValue(getCellRaw(coord));
                                        setEditing(false);
                                      }
                                    },
                                  })
                                : getCellDisplay(coord)
                            );
                          })
                        );
                      })
                    )
                  )
                )
              : h(
                  'div',
                  { className: 'dsh-xs-empty' },
                  h(ExcelStudioIcon, { size: 36 }),
                  h('div', { style: { fontWeight: 600, fontSize: '15px' } }, 'DeepSeek Excel Studio'),
                  h(
                    'div',
                    { style: { maxWidth: '420px', fontSize: '12px', lineHeight: '1.5' } },
                    'Open a local .xlsx or .xlsm workbook from your allowed folders, or create a new spreadsheet to edit directly inside DeepSeek Harness.'
                  ),
                  h(
                    'button',
                    {
                      type: 'button',
                      className: 'dsh-xs-btn dsh-xs-btn-primary',
                      onClick: function () { setShowNewForm(true); },
                    },
                    '+ Create New Workbook'
                  )
                ),

            // Bottom Sheet Tabs
            sheet
              ? h(
                  'div',
                  { className: 'dsh-xs-tabs' },
                  (sheet.sheets || []).map(function (name) {
                    var active = name === sheet.active_sheet;
                    return h(
                      'button',
                      {
                        key: name,
                        type: 'button',
                        className: 'dsh-xs-tab ' + (active ? 'dsh-xs-tab-active' : ''),
                        onClick: function () {
                          if (!active) openWorkbook(sheet.file_path, name);
                        },
                      },
                      name
                    );
                  }),
                  h(
                    'button',
                    {
                      type: 'button',
                      className: 'dsh-xs-btn',
                      title: 'Add worksheet',
                      onClick: function () {
                        var name = window.prompt('New worksheet name:', 'Sheet' + ((sheet.sheets || []).length + 1));
                        if (name && name.trim()) runSheetOp('create_sheet', { new_name: name.trim() });
                      },
                    },
                    '+ Sheet'
                  ),
                  h(
                    'button',
                    {
                      type: 'button',
                      className: 'dsh-xs-btn',
                      title: 'Rename current worksheet',
                      onClick: function () {
                        var name = window.prompt('Rename worksheet "' + sheet.active_sheet + '" to:', sheet.active_sheet);
                        if (name && name.trim() && name.trim() !== sheet.active_sheet) {
                          runSheetOp('rename_sheet', { new_name: name.trim() });
                        }
                      },
                    },
                    'Rename'
                  ),
                  (sheet.sheets || []).length > 1
                    ? h(
                        'button',
                        {
                          type: 'button',
                          className: 'dsh-xs-btn dsh-xs-btn-danger',
                          title: 'Delete current worksheet',
                          onClick: function () {
                            if (window.confirm('Delete worksheet "' + sheet.active_sheet + '"? A backup will be created first.')) {
                              runSheetOp('delete_sheet');
                            }
                          },
                        },
                        'Delete Sheet'
                      )
                    : null
                )
              : null,

            // Status bar
            h(
              'div',
              { className: 'dsh-xs-status' },
              h('span', null, sheet ? sheet.file_path : (allowedDirs.length ? 'Allowed: ' + allowedDirs.join(' | ') : 'Ready')),
              h(
                'span',
                null,
                sheet
                  ? (sheet.has_vba ? 'VBA preserved · ' : '') +
                      (sheet.is_locked ? 'Excel lock detected · ' : '') +
                      'Rows: ' +
                      sheet.max_row +
                      ' · Cols: ' +
                      sheet.max_column +
                      (dirtyCount > 0 ? ' · ' + dirtyCount + ' unsaved' : '')
                  : 'v0.3.0'
              )
            )
          );
        }

        exports.name = PLUGIN_ID;
        exports.inject = ['slots'];
        exports.apply = function apply(ctx) {
          try {
            if (!ctx || !ctx.slots || typeof ctx.slots.inject !== 'function') {
              return;
            }
            ctx.slots.inject('main', function () {
              try {
                return ctx.slots.register(
                  {
                    name: 'main',
                    key: PANEL_ID,
                  },
                  ExcelStudioPanel
                );
              } catch (_e) {
                return undefined;
              }
            });
            ctx.slots.inject('sidebar.panellist', function () {
              try {
                return ctx.slots.register(
                  {
                    name: 'sidebar.panellist',
                    id: PANEL_ID,
                    order: 45,
                    label: 'Excel Studio',
                  },
                  ExcelStudioIcon
                );
              } catch (_e) {
                return undefined;
              }
            });
          } catch (_err) {
            // Never block DeepSeek Harness Web UI boot if slots differ across builds
          }
        };

        return module.exports;
      },
    });
  } catch (_err) {
    // Ignore duplicate registration in combo-script re-evaluation
  }
})();
