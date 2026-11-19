(function () {
  'use strict';

  if (typeof window === 'undefined' || !window.__ModuleLoader__ || typeof window.__ModuleLoader__.load !== 'function') {
    return;
  }

  var PLUGIN_ID = 'deepseek-office-plugin';
  var PANEL_ID = 'excel-studio';
  var STYLE_ID = 'dsh-excel-studio-styles-v4';

  var CSS_TEXT = [
    '.dsh-xl-root {',
    '  display: flex;',
    '  flex-direction: column;',
    '  height: 100%;',
    '  width: 100%;',
    '  min-height: 0;',
    '  min-width: 0;',
    '  overflow: hidden;',
    '  font-family: "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, sans-serif;',
    '  font-size: 12px;',
    '  color: var(--dsw-alias-text-primary, #1f2328);',
    '  background: var(--dsw-alias-bg-primary, #ffffff);',
    '  position: relative;',
    '  user-select: none;',
    '}',
    '/* Excel Green Title / Quick Access Bar */',
    '.dsh-xl-titlebar {',
    '  display: flex;',
    '  align-items: center;',
    '  justify-content: space-between;',
    '  height: 34px;',
    '  padding: 0 12px;',
    '  background: #107c41;',
    '  color: #ffffff;',
    '  gap: 10px;',
    '  flex-shrink: 0;',
    '}',
    '.dsh-xl-qat {',
    '  display: inline-flex;',
    '  align-items: center;',
    '  gap: 2px;',
    '}',
    '.dsh-xl-qat-btn {',
    '  height: 24px;',
    '  min-width: 26px;',
    '  padding: 0 6px;',
    '  border: none;',
    '  border-radius: 3px;',
    '  background: transparent;',
    '  color: #ffffff;',
    '  font-size: 12px;',
    '  cursor: pointer;',
    '  display: inline-flex;',
    '  align-items: center;',
    '  justify-content: center;',
    '  gap: 4px;',
    '}',
    '.dsh-xl-qat-btn:hover:not(:disabled) { background: rgba(255,255,255,0.18); }',
    '.dsh-xl-qat-btn:disabled { opacity: 0.4; cursor: not-allowed; }',
    '.dsh-xl-title-center {',
    '  display: inline-flex;',
    '  align-items: center;',
    '  gap: 8px;',
    '  font-weight: 600;',
    '  font-size: 12.5px;',
    '  min-width: 0;',
    '}',
    '.dsh-xl-wb-select {',
    '  height: 24px;',
    '  max-width: 260px;',
    '  padding: 0 8px;',
    '  border-radius: 4px;',
    '  border: 1px solid rgba(255,255,255,0.32);',
    '  background: rgba(0,0,0,0.16);',
    '  color: #ffffff;',
    '  font-size: 12px;',
    '  font-weight: 600;',
    '  outline: none;',
    '}',
    '.dsh-xl-wb-select option { color: #1f2328; background: #ffffff; }',
    '.dsh-xl-pill {',
    '  font-size: 11px;',
    '  padding: 1px 7px;',
    '  border-radius: 999px;',
    '  background: rgba(255,255,255,0.18);',
    '  font-weight: 500;',
    '}',
    '/* Excel Menu Tabs Bar */',
    '.dsh-xl-tabs-bar {',
    '  display: flex;',
    '  align-items: flex-end;',
    '  gap: 2px;',
    '  padding: 2px 10px 0;',
    '  background: var(--dsw-alias-bg-secondary, #f3f3f3);',
    '  border-bottom: 1px solid var(--dsw-alias-border-subtle, #d9d9d9);',
    '  flex-shrink: 0;',
    '}',
    '.dsh-xl-menu-tab {',
    '  height: 27px;',
    '  padding: 0 13px;',
    '  border: none;',
    '  border-bottom: 2px solid transparent;',
    '  background: transparent;',
    '  color: inherit;',
    '  font-size: 12.5px;',
    '  cursor: pointer;',
    '  border-radius: 4px 4px 0 0;',
    '}',
    '.dsh-xl-menu-tab:hover { background: rgba(127,127,127,0.1); }',
    '.dsh-xl-menu-tab-active {',
    '  font-weight: 600;',
    '  color: #107c41;',
    '  border-bottom-color: #107c41;',
    '  background: var(--dsw-alias-bg-primary, #ffffff);',
    '}',
    '.dsh-xl-file-tab {',
    '  background: #107c41;',
    '  color: #ffffff !important;',
    '  font-weight: 600;',
    '  margin-right: 4px;',
    '  border-bottom-color: #107c41;',
    '}',
    '.dsh-xl-file-tab:hover { background: #0d6635 !important; }',
    '/* Excel Ribbon Panel */',
    '.dsh-xl-ribbon {',
    '  display: flex;',
    '  align-items: stretch;',
    '  gap: 0;',
    '  padding: 5px 10px 3px;',
    '  background: var(--dsw-alias-bg-primary, #ffffff);',
    '  border-bottom: 1px solid var(--dsw-alias-border-subtle, #d9d9d9);',
    '  overflow-x: auto;',
    '  flex-shrink: 0;',
    '  min-height: 76px;',
    '}',
    '.dsh-xl-rgroup {',
    '  display: flex;',
    '  flex-direction: column;',
    '  justify-content: space-between;',
    '  padding: 0 10px;',
    '  border-right: 1px solid var(--dsw-alias-border-subtle, rgba(127,127,127,0.22));',
    '  flex-shrink: 0;',
    '}',
    '.dsh-xl-rgroup:first-child { padding-left: 2px; }',
    '.dsh-xl-rgroup:last-child { border-right: none; }',
    '.dsh-xl-rgroup-body {',
    '  display: flex;',
    '  flex-direction: column;',
    '  gap: 4px;',
    '  justify-content: center;',
    '  flex: 1;',
    '}',
    '.dsh-xl-rgroup-row {',
    '  display: flex;',
    '  align-items: center;',
    '  gap: 3px;',
    '}',
    '.dsh-xl-rgroup-label {',
    '  font-size: 10px;',
    '  text-align: center;',
    '  opacity: 0.62;',
    '  margin-top: 3px;',
    '  letter-spacing: 0.2px;',
    '}',
    '.dsh-xl-rbtn {',
    '  height: 24px;',
    '  min-width: 24px;',
    '  padding: 0 7px;',
    '  border: 1px solid transparent;',
    '  border-radius: 3px;',
    '  background: transparent;',
    '  color: inherit;',
    '  font-size: 12px;',
    '  cursor: pointer;',
    '  display: inline-flex;',
    '  align-items: center;',
    '  justify-content: center;',
    '  gap: 4px;',
    '  white-space: nowrap;',
    '}',
    '.dsh-xl-rbtn:hover:not(:disabled) {',
    '  background: var(--dsw-alias-bg-hover, rgba(127,127,127,0.12));',
    '  border-color: rgba(127,127,127,0.22);',
    '}',
    '.dsh-xl-rbtn:disabled { opacity: 0.4; cursor: not-allowed; }',
    '.dsh-xl-rbtn-active {',
    '  background: rgba(16,124,65,0.15) !important;',
    '  border-color: rgba(16,124,65,0.5) !important;',
    '  color: #107c41;',
    '}',
    '.dsh-xl-rbtn-tall {',
    '  height: 48px;',
    '  flex-direction: column;',
    '  padding: 4px 9px;',
    '  font-size: 11px;',
    '  gap: 2px;',
    '}',
    '.dsh-xl-rselect, .dsh-xl-rinput {',
    '  height: 24px;',
    '  padding: 0 6px;',
    '  border-radius: 3px;',
    '  border: 1px solid var(--dsw-alias-border-default, rgba(127,127,127,0.32));',
    '  background: var(--dsw-alias-bg-primary, #ffffff);',
    '  color: inherit;',
    '  font-size: 12px;',
    '  outline: none;',
    '}',
    '.dsh-xl-rselect:focus, .dsh-xl-rinput:focus { border-color: #107c41; }',
    '.dsh-xl-swatch {',
    '  width: 16px;',
    '  height: 16px;',
    '  border-radius: 2px;',
    '  border: 1px solid rgba(127,127,127,0.45);',
    '  cursor: pointer;',
    '  padding: 0;',
    '}',
    '/* File Backstage Dropdown Overlay */',
    '.dsh-xl-file-menu {',
    '  position: absolute;',
    '  top: 63px;',
    '  left: 10px;',
    '  width: 320px;',
    '  z-index: 40;',
    '  background: var(--dsw-alias-bg-primary, #ffffff);',
    '  border: 1px solid var(--dsw-alias-border-default, #c8c8c8);',
    '  border-radius: 6px;',
    '  box-shadow: 0 10px 28px rgba(0,0,0,0.18);',
    '  padding: 8px;',
    '  display: flex;',
    '  flex-direction: column;',
    '  gap: 4px;',
    '}',
    '.dsh-xl-menu-item {',
    '  display: flex;',
    '  align-items: center;',
    '  justify-content: space-between;',
    '  padding: 7px 10px;',
    '  border: none;',
    '  border-radius: 4px;',
    '  background: transparent;',
    '  color: inherit;',
    '  font-size: 12.5px;',
    '  cursor: pointer;',
    '  text-align: left;',
    '  width: 100%;',
    '}',
    '.dsh-xl-menu-item:hover:not(:disabled) { background: rgba(16,124,65,0.1); }',
    '.dsh-xl-menu-item:disabled { opacity: 0.45; cursor: not-allowed; }',
    '.dsh-xl-menu-sep { height: 1px; background: rgba(127,127,127,0.2); margin: 3px 0; }',
    '/* Formula Bar */',
    '.dsh-xl-fbar {',
    '  display: flex;',
    '  align-items: center;',
    '  gap: 6px;',
    '  padding: 4px 10px;',
    '  background: var(--dsw-alias-bg-secondary, #f8f9fa);',
    '  border-bottom: 1px solid var(--dsw-alias-border-subtle, #d9d9d9);',
    '  flex-shrink: 0;',
    '}',
    '.dsh-xl-namebox {',
    '  width: 78px;',
    '  height: 24px;',
    '  text-align: center;',
    '  font-family: ui-monospace, SFMono-Regular, Menlo, monospace;',
    '  font-weight: 600;',
    '  font-size: 12px;',
    '  border: 1px solid var(--dsw-alias-border-default, #c8c8c8);',
    '  border-radius: 3px;',
    '  background: var(--dsw-alias-bg-primary, #ffffff);',
    '  color: inherit;',
    '  outline: none;',
    '}',
    '.dsh-xl-finput {',
    '  flex: 1;',
    '  height: 24px;',
    '  padding: 0 8px;',
    '  border: 1px solid var(--dsw-alias-border-default, #c8c8c8);',
    '  border-radius: 3px;',
    '  background: var(--dsw-alias-bg-primary, #ffffff);',
    '  color: inherit;',
    '  font-family: ui-monospace, SFMono-Regular, Menlo, monospace;',
    '  font-size: 12.5px;',
    '  outline: none;',
    '  user-select: text;',
    '}',
    '.dsh-xl-finput:focus { border-color: #107c41; }',
    '/* Grid */',
    '.dsh-xl-grid-wrap {',
    '  flex: 1;',
    '  min-height: 0;',
    '  overflow: auto;',
    '  position: relative;',
    '  background: var(--dsw-alias-bg-primary, #ffffff);',
    '}',
    '.dsh-xl-table {',
    '  border-collapse: separate;',
    '  border-spacing: 0;',
    '  table-layout: fixed;',
    '  width: max-content;',
    '  min-width: 100%;',
    '}',
    '.dsh-xl-corner {',
    '  position: sticky;',
    '  top: 0;',
    '  left: 0;',
    '  z-index: 6;',
    '  width: 44px;',
    '  min-width: 44px;',
    '  height: 24px;',
    '  background: var(--dsw-alias-bg-secondary, #f0f0f0);',
    '  border-right: 1px solid #c6c6c6;',
    '  border-bottom: 1px solid #c6c6c6;',
    '}',
    '.dsh-xl-colhdr {',
    '  position: sticky;',
    '  top: 0;',
    '  z-index: 4;',
    '  height: 24px;',
    '  background: var(--dsw-alias-bg-secondary, #f0f0f0);',
    '  color: var(--dsw-alias-text-secondary, #444444);',
    '  font-weight: 600;',
    '  font-size: 11px;',
    '  text-align: center;',
    '  border-right: 1px solid #d0d0d0;',
    '  border-bottom: 1px solid #c6c6c6;',
    '  position: sticky;',
    '}',
    '.dsh-xl-colhdr-active {',
    '  background: rgba(16,124,65,0.16);',
    '  color: #107c41;',
    '  border-bottom: 2px solid #107c41;',
    '}',
    '.dsh-xl-resizer {',
    '  position: absolute;',
    '  top: 0;',
    '  right: 0;',
    '  width: 5px;',
    '  height: 100%;',
    '  cursor: col-resize;',
    '}',
    '.dsh-xl-rowhdr {',
    '  position: sticky;',
    '  left: 0;',
    '  z-index: 3;',
    '  width: 44px;',
    '  min-width: 44px;',
    '  text-align: center;',
    '  font-size: 11px;',
    '  color: var(--dsw-alias-text-secondary, #555555);',
    '  background: var(--dsw-alias-bg-secondary, #f0f0f0);',
    '  border-right: 1px solid #c6c6c6;',
    '  border-bottom: 1px solid #d8d8d8;',
    '}',
    '.dsh-xl-rowhdr-active {',
    '  background: rgba(16,124,65,0.16);',
    '  color: #107c41;',
    '  font-weight: 700;',
    '  border-right: 2px solid #107c41;',
    '}',
    '.dsh-xl-cell {',
    '  height: 24px;',
    '  padding: 0 5px;',
    '  white-space: nowrap;',
    '  overflow: hidden;',
    '  text-overflow: ellipsis;',
    '  cursor: cell;',
    '  position: relative;',
    '  vertical-align: middle;',
    '  box-sizing: border-box;',
    '}',
    '.dsh-xl-gridlines .dsh-xl-cell {',
    '  border-right: 1px solid var(--dsw-alias-border-subtle, #e2e2e2);',
    '  border-bottom: 1px solid var(--dsw-alias-border-subtle, #e2e2e2);',
    '}',
    '.dsh-xl-cell-inrange {',
    '  background-color: rgba(16, 124, 65, 0.1) !important;',
    '}',
    '.dsh-xl-cell-active {',
    '  outline: 2px solid #107c41;',
    '  outline-offset: -2px;',
    '  z-index: 2;',
    '}',
    '.dsh-xl-cell-dirty::before {',
    '  content: "";',
    '  position: absolute;',
    '  top: 1px;',
    '  right: 1px;',
    '  width: 4px;',
    '  height: 4px;',
    '  border-radius: 50%;',
    '  background: #d97706;',
    '}',
    '.dsh-xl-fillhandle {',
    '  position: absolute;',
    '  right: -3px;',
    '  bottom: -3px;',
    '  width: 7px;',
    '  height: 7px;',
    '  background: #107c41;',
    '  border: 1px solid #ffffff;',
    '  cursor: crosshair;',
    '  z-index: 5;',
    '}',
    '.dsh-xl-editor-input {',
    '  width: 100%;',
    '  height: 100%;',
    '  border: none;',
    '  outline: none;',
    '  padding: 0;',
    '  margin: 0;',
    '  background: var(--dsw-alias-bg-primary, #ffffff);',
    '  color: inherit;',
    '  font: inherit;',
    '  user-select: text;',
    '}',
    '.dsh-xl-frozen-row td {',
    '  position: sticky;',
    '  top: 24px;',
    '  z-index: 3;',
    '  background: var(--dsw-alias-bg-primary, #ffffff);',
    '  box-shadow: 0 2px 0 rgba(16,124,65,0.35);',
    '}',
    '.dsh-xl-frozen-row .dsh-xl-rowhdr { z-index: 5; }',
    '/* Context Menu */',
    '.dsh-xl-ctxmenu {',
    '  position: fixed;',
    '  z-index: 60;',
    '  min-width: 185px;',
    '  background: var(--dsw-alias-bg-primary, #ffffff);',
    '  border: 1px solid var(--dsw-alias-border-default, #c8c8c8);',
    '  border-radius: 6px;',
    '  box-shadow: 0 8px 24px rgba(0,0,0,0.2);',
    '  padding: 4px;',
    '}',
    '/* Bottom Sheet Bar & Excel Status Bar */',
    '.dsh-xl-sheetbar {',
    '  display: flex;',
    '  align-items: center;',
    '  gap: 2px;',
    '  padding: 0 10px;',
    '  height: 30px;',
    '  background: var(--dsw-alias-bg-secondary, #f3f3f3);',
    '  border-top: 1px solid var(--dsw-alias-border-subtle, #d9d9d9);',
    '  overflow-x: auto;',
    '  flex-shrink: 0;',
    '}',
    '.dsh-xl-sheettab {',
    '  height: 26px;',
    '  padding: 0 14px;',
    '  border: none;',
    '  border-bottom: 2px solid transparent;',
    '  background: transparent;',
    '  color: inherit;',
    '  font-size: 12px;',
    '  cursor: pointer;',
    '  white-space: nowrap;',
    '}',
    '.dsh-xl-sheettab:hover { background: rgba(127,127,127,0.1); }',
    '.dsh-xl-sheettab-active {',
    '  background: var(--dsw-alias-bg-primary, #ffffff);',
    '  color: #107c41;',
    '  font-weight: 600;',
    '  border-bottom-color: #107c41;',
    '}',
    '.dsh-xl-addsheet {',
    '  width: 24px;',
    '  height: 24px;',
    '  border-radius: 50%;',
    '  border: 1px solid rgba(127,127,127,0.35);',
    '  background: transparent;',
    '  color: inherit;',
    '  font-weight: 700;',
    '  cursor: pointer;',
    '  margin-left: 4px;',
    '}',
    '.dsh-xl-addsheet:hover { background: rgba(16,124,65,0.14); color: #107c41; }',
    '.dsh-xl-statusbar {',
    '  display: flex;',
    '  align-items: center;',
    '  justify-content: space-between;',
    '  height: 24px;',
    '  padding: 0 12px;',
    '  background: #107c41;',
    '  color: #ffffff;',
    '  font-size: 11px;',
    '  gap: 12px;',
    '  flex-shrink: 0;',
    '}',
    '.dsh-xl-banner {',
    '  padding: 6px 12px;',
    '  font-size: 12px;',
    '  display: flex;',
    '  align-items: center;',
    '  justify-content: space-between;',
    '  gap: 8px;',
    '  flex-shrink: 0;',
    '}',
    '.dsh-xl-banner-err { background: #fdecea; color: #b42318; border-bottom: 1px solid #f5c2c7; }',
    '.dsh-xl-banner-info { background: #e8f5ee; color: #107c41; border-bottom: 1px solid #badbcc; }',
    '.dsh-xl-findbar {',
    '  display: flex;',
    '  align-items: center;',
    '  gap: 6px;',
    '  padding: 5px 12px;',
    '  background: var(--dsw-alias-bg-secondary, #f3f3f3);',
    '  border-bottom: 1px solid var(--dsw-alias-border-subtle, #d9d9d9);',
    '  flex-shrink: 0;',
    '}',
    '.dsh-xl-empty {',
    '  flex: 1;',
    '  display: flex;',
    '  flex-direction: column;',
    '  align-items: center;',
    '  justify-content: center;',
    '  gap: 12px;',
    '  padding: 32px;',
    '  text-align: center;',
    '}',
    '.dsh-xl-btn-primary {',
    '  height: 30px;',
    '  padding: 0 14px;',
    '  border-radius: 4px;',
    '  border: none;',
    '  background: #107c41;',
    '  color: #ffffff;',
    '  font-weight: 600;',
    '  font-size: 12.5px;',
    '  cursor: pointer;',
    '}',
    '.dsh-xl-btn-primary:hover { background: #0d6635; }',
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

  function letterToColIndex(letters) {
    var s = String(letters || '').toUpperCase();
    var n = 0;
    for (var i = 0; i < s.length; i++) {
      n = n * 26 + (s.charCodeAt(i) - 64);
    }
    return n;
  }

  function makeCoord(row, col) {
    return colIndexToLetter(col) + String(row);
  }

  function parseCoord(str) {
    var m = /^\$?([A-Za-z]+)\$?(\d+)$/.exec(String(str || '').trim());
    if (!m) return null;
    return { col: letterToColIndex(m[1]), row: parseInt(m[2], 10) };
  }

  function normalizeRange(a, b) {
    var r1 = Math.min(a.row, b.row);
    var r2 = Math.max(a.row, b.row);
    var c1 = Math.min(a.col, b.col);
    var c2 = Math.max(a.col, b.col);
    return { r1: r1, r2: r2, c1: c1, c2: c2 };
  }

  function formatRangeLabel(a, b) {
    var norm = normalizeRange(a, b);
    if (norm.r1 === norm.r2 && norm.c1 === norm.c2) {
      return makeCoord(norm.r1, norm.c1);
    }
    return makeCoord(norm.r1, norm.c1) + ':' + makeCoord(norm.r2, norm.c2);
  }

  // Lightweight live Excel formula evaluator for immediate grid feedback
  function evalExcelFormula(expr, getRawFn, depth) {
    if (depth > 8) return '#REF!';
    var body = String(expr || '').trim();
    if (!body.startsWith('=')) return body;
    var f = body.slice(1).trim();

    function resolveNum(coord) {
      var v = getRawFn(coord);
      if (typeof v === 'string' && v.startsWith('=')) {
        v = evalExcelFormula(v, getRawFn, depth + 1);
      }
      var n = Number(v);
      return Number.isFinite(n) && String(v).trim() !== '' ? n : 0;
    }

    function expandRangeNums(rangeText) {
      var parts = rangeText.split(':');
      var nums = [];
      if (parts.length === 2) {
        var p1 = parseCoord(parts[0]);
        var p2 = parseCoord(parts[1]);
        if (p1 && p2) {
          var norm = normalizeRange(p1, p2);
          for (var r = norm.r1; r <= norm.r2; r++) {
            for (var c = norm.c1; c <= norm.c2; c++) {
              var raw = getRawFn(makeCoord(r, c));
              if (raw !== '' && raw !== null && raw !== undefined) {
                var val = typeof raw === 'string' && raw.startsWith('=')
                  ? evalExcelFormula(raw, getRawFn, depth + 1)
                  : raw;
                var n = Number(val);
                if (Number.isFinite(n)) nums.push(n);
              }
            }
          }
          return nums;
        }
      }
      var items = rangeText.split(',');
      for (var i = 0; i < items.length; i++) {
        var tok = items[i].trim();
        if (!tok) continue;
        if (tok.indexOf(':') !== -1) {
          nums = nums.concat(expandRangeNums(tok));
        } else if (parseCoord(tok)) {
          nums.push(resolveNum(tok.toUpperCase()));
        } else {
          var lit = Number(tok);
          if (Number.isFinite(lit)) nums.push(lit);
        }
      }
      return nums;
    }

    try {
      var fnMatch = /^(SUM|AVERAGE|AVG|COUNT|MIN|MAX)\((.+)\)$/i.exec(f);
      if (fnMatch) {
        var fnName = fnMatch[1].toUpperCase();
        var nums = expandRangeNums(fnMatch[2]);
        if (fnName === 'COUNT') return nums.length;
        if (nums.length === 0) return 0;
        if (fnName === 'SUM') {
          var s = 0;
          for (var k = 0; k < nums.length; k++) s += nums[k];
          return Math.round(s * 1e6) / 1e6;
        }
        if (fnName === 'AVERAGE' || fnName === 'AVG') {
          var sa = 0;
          for (var ka = 0; ka < nums.length; ka++) sa += nums[ka];
          return Math.round((sa / nums.length) * 1e6) / 1e6;
        }
        if (fnName === 'MIN') return Math.min.apply(null, nums);
        if (fnName === 'MAX') return Math.max.apply(null, nums);
      }

      // Simple arithmetic with cell references like =B2*2+C2
      var replaced = f.replace(/\b([A-Za-z]{1,3}\d+)\b/g, function (_m, ref) {
        return String(resolveNum(ref.toUpperCase()));
      });
      if (/^[0-9+\-*/().\s]+$/.test(replaced)) {
        // Safe numeric expression only (no identifiers)
        var computed = Function('"use strict"; return (' + replaced + ');')();
        if (typeof computed === 'number' && Number.isFinite(computed)) {
          return Math.round(computed * 1e6) / 1e6;
        }
      }
    } catch (_e) {
      return '#ERROR!';
    }
    return body;
  }

  function formatDisplayValue(val, numFmt) {
    if (val === '' || val === null || val === undefined) return '';
    if (!numFmt || numFmt === 'General') return String(val);
    var n = Number(val);
    if (!Number.isFinite(n)) return String(val);
    if (numFmt.indexOf('%') !== -1) {
      return (n * 100).toFixed(1) + '%';
    }
    if (numFmt.indexOf('$') !== -1) {
      return '$' + n.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }
    if (numFmt === '#,##0') {
      return Math.round(n).toLocaleString();
    }
    if (numFmt === '#,##0.00') {
      return n.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }
    return String(val);
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
          var useState = React.useState;
          var useEffect = React.useEffect;
          var useCallback = React.useCallback;
          var useRef = React.useRef;

          var sWorkbooks = useState([]);
          var workbooks = sWorkbooks[0], setWorkbooks = sWorkbooks[1];
          var sAllowedDirs = useState([]);
          var allowedDirs = sAllowedDirs[0], setAllowedDirs = sAllowedDirs[1];
          var sSheet = useState(null);
          var sheet = sSheet[0], setSheet = sSheet[1];

          var sRibbonTab = useState('Home');
          var ribbonTab = sRibbonTab[0], setRibbonTab = sRibbonTab[1];
          var sFileMenu = useState(false);
          var fileMenuOpen = sFileMenu[0], setFileMenuOpen = sFileMenu[1];

          var sLoading = useState(false);
          var loading = sLoading[0], setLoading = sLoading[1];
          var sSaving = useState(false);
          var saving = sSaving[0], setSaving = sSaving[1];
          var sError = useState('');
          var error = sError[0], setError = sError[1];
          var sNotice = useState('');
          var notice = sNotice[0], setNotice = sNotice[1];

          // Selection range (anchor -> focus)
          var sSelStart = useState({ row: 1, col: 1 });
          var selStart = sSelStart[0], setSelStart = sSelStart[1];
          var sSelEnd = useState({ row: 1, col: 1 });
          var selEnd = sSelEnd[0], setSelEnd = sSelEnd[1];
          var isDraggingRef = useRef(false);
          var isFillingRef = useRef(false);

          var sEditing = useState(false);
          var editing = sEditing[0], setEditing = sEditing[1];
          var sDraft = useState('');
          var draftValue = sDraft[0], setDraftValue = sDraft[1];
          var sNameBox = useState('A1');
          var nameBoxText = sNameBox[0], setNameBoxText = sNameBox[1];

          var sPendingEdits = useState({});
          var pendingEdits = sPendingEdits[0], setPendingEdits = sPendingEdits[1];
          var sPendingFormats = useState({});
          var pendingFormats = sPendingFormats[0], setPendingFormats = sPendingFormats[1];
          var sColWidths = useState({});
          var pendingColWidths = sColWidths[0], setPendingColWidths = sColWidths[1];

          // Undo / Redo stacks
          var sUndo = useState([]);
          var undoStack = sUndo[0], setUndoStack = sUndo[1];
          var sRedo = useState([]);
          var redoStack = sRedo[0], setRedoStack = sRedo[1];

          // Clipboard (2D matrix)
          var clipboardRef = useRef(null);

          // View options
          var sGridlines = useState(true);
          var showGridlines = sGridlines[0], setShowGridlines = sGridlines[1];
          var sFormulaBar = useState(true);
          var showFormulaBar = sFormulaBar[0], setShowFormulaBar = sFormulaBar[1];
          var sHeadings = useState(true);
          var showHeadings = sHeadings[0], setShowHeadings = sHeadings[1];
          var sFreezeTop = useState(false);
          var freezeTopRow = sFreezeTop[0], setFreezeTopRow = sFreezeTop[1];
          var sShowFormulas = useState(false);
          var showFormulas = sShowFormulas[0], setShowFormulas = sShowFormulas[1];
          var sZoom = useState(100);
          var zoom = sZoom[0], setZoom = sZoom[1];

          // Find & Replace
          var sFindOpen = useState(false);
          var findOpen = sFindOpen[0], setFindOpen = sFindOpen[1];
          var sFindText = useState('');
          var findText = sFindText[0], setFindText = sFindText[1];
          var sReplaceText = useState('');
          var replaceText = sReplaceText[0], setReplaceText = sReplaceText[1];

          // Context menu
          var sCtxMenu = useState(null);
          var ctxMenu = sCtxMenu[0], setCtxMenu = sCtxMenu[1];

          // Summary drawer
          var sSummary = useState(null);
          var summary = sSummary[0], setSummary = sSummary[1];

          // New workbook prompt
          var sShowNew = useState(false);
          var showNewForm = sShowNew[0], setShowNewForm = sShowNew[1];
          var sNewName = useState('Book1.xlsx');
          var newFilename = sNewName[0], setNewFilename = sNewName[1];
          var sOpenPath = useState('');
          var openPathInput = sOpenPath[0], setOpenPathInput = sOpenPath[1];

          useEffect(function () {
            ensureStyles();
          }, []);

          useEffect(function () {
            var onMouseUp = function () {
              isDraggingRef.current = false;
              isFillingRef.current = false;
            };
            var onClickOutside = function () {
              setCtxMenu(null);
            };
            window.addEventListener('mouseup', onMouseUp);
            window.addEventListener('click', onClickOutside);
            return function () {
              window.removeEventListener('mouseup', onMouseUp);
              window.removeEventListener('click', onClickOutside);
            };
          }, []);

          var activeCoord = makeCoord(selStart.row, selStart.col);
          var selNorm = normalizeRange(selStart, selEnd);
          var dirtyCount =
            Object.keys(pendingEdits).length +
            Object.keys(pendingFormats).length +
            Object.keys(pendingColWidths).length;

          useEffect(function () {
            setNameBoxText(formatRangeLabel(selStart, selEnd));
          }, [selStart, selEnd]);

          var getCellRaw = useCallback(function (coord) {
            if (Object.prototype.hasOwnProperty.call(pendingEdits, coord)) {
              return pendingEdits[coord];
            }
            var cell = sheet && sheet.cells && sheet.cells[coord];
            if (!cell) return '';
            return cell.formula !== undefined
              ? cell.formula
              : (cell.raw !== undefined ? String(cell.raw) : '');
          }, [sheet, pendingEdits]);

          var getCellStyle = useCallback(function (coord) {
            var base = (sheet && sheet.cells && sheet.cells[coord] && sheet.cells[coord].style) || {};
            var patch = pendingFormats[coord] || {};
            return Object.assign({}, base, patch);
          }, [sheet, pendingFormats]);

          var getCellDisplay = useCallback(function (coord) {
            var raw = getCellRaw(coord);
            var st = getCellStyle(coord);
            if (showFormulas) return String(raw);
            if (typeof raw === 'string' && raw.startsWith('=')) {
              var computed = evalExcelFormula(raw, getCellRaw, 0);
              return formatDisplayValue(computed, st.numberFormat || st.number_format);
            }
            return formatDisplayValue(raw, st.numberFormat || st.number_format);
          }, [getCellRaw, getCellStyle, showFormulas]);

          var pushUndoSnapshot = useCallback(function () {
            setUndoStack(function (stack) {
              var next = stack.slice(-29);
              next.push({
                edits: Object.assign({}, pendingEdits),
                formats: JSON.parse(JSON.stringify(pendingFormats)),
              });
              return next;
            });
            setRedoStack([]);
          }, [pendingEdits, pendingFormats]);

          var handleUndo = useCallback(function () {
            if (undoStack.length === 0) return;
            var prev = undoStack[undoStack.length - 1];
            setRedoStack(function (r) {
              return r.concat([{ edits: Object.assign({}, pendingEdits), formats: JSON.parse(JSON.stringify(pendingFormats)) }]);
            });
            setUndoStack(undoStack.slice(0, -1));
            setPendingEdits(prev.edits);
            setPendingFormats(prev.formats);
          }, [undoStack, pendingEdits, pendingFormats]);

          var handleRedo = useCallback(function () {
            if (redoStack.length === 0) return;
            var next = redoStack[redoStack.length - 1];
            setUndoStack(function (u) {
              return u.concat([{ edits: Object.assign({}, pendingEdits), formats: JSON.parse(JSON.stringify(pendingFormats)) }]);
            });
            setRedoStack(redoStack.slice(0, -1));
            setPendingEdits(next.edits);
            setPendingFormats(next.formats);
          }, [redoStack, pendingEdits, pendingFormats]);

          var openWorkbook = useCallback(function (filePath, sheetName) {
            if (!filePath) return;
            setLoading(true);
            setError('');
            setSummary(null);
            setFileMenuOpen(false);
            apiRequest('/api/excel-studio/open', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ file_path: filePath, sheet_name: sheetName }),
            })
              .then(function (data) {
                setSheet(data);
                setPendingEdits({});
                setPendingFormats({});
                setPendingColWidths({});
                setUndoStack([]);
                setRedoStack([]);
                setSelStart({ row: 1, col: 1 });
                setSelEnd({ row: 1, col: 1 });
                setEditing(false);
                var first = data.cells && data.cells.A1;
                setDraftValue(first ? (first.formula || String(first.raw || '')) : '');
                setOpenPathInput(data.file_path || filePath);
              })
              .catch(function (err) {
                setError(err.message || String(err));
              })
              .finally(function () {
                setLoading(false);
              });
          }, []);

          var refreshWorkbooks = useCallback(function (autoOpenFirst) {
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

          useEffect(function () {
            refreshWorkbooks(true);
          }, [refreshWorkbooks]);

          var selectCell = function (row, col, startEdit, extend) {
            var maxR = (sheet && sheet.row_count) || 40;
            var maxC = (sheet && sheet.col_count) || 16;
            var r = Math.max(1, Math.min(maxR, row));
            var c = Math.max(1, Math.min(maxC, col));
            if (extend) {
              setSelEnd({ row: r, col: c });
            } else {
              setSelStart({ row: r, col: c });
              setSelEnd({ row: r, col: c });
              setDraftValue(getCellRaw(makeCoord(r, c)));
            }
            setEditing(Boolean(startEdit));
          };

          var setCellValuesBatch = useCallback(function (updates) {
            pushUndoSnapshot();
            setPendingEdits(function (cur) {
              var next = Object.assign({}, cur);
              for (var coord in updates) {
                if (Object.prototype.hasOwnProperty.call(updates, coord)) {
                  next[coord] = updates[coord];
                }
              }
              return next;
            });
          }, [pushUndoSnapshot]);

          var applyFormatToSelection = useCallback(function (patch) {
            if (!sheet) return;
            pushUndoSnapshot();
            setPendingFormats(function (cur) {
              var next = Object.assign({}, cur);
              for (var r = selNorm.r1; r <= selNorm.r2; r++) {
                for (var c = selNorm.c1; c <= selNorm.c2; c++) {
                  var coord = makeCoord(r, c);
                  next[coord] = Object.assign({}, cur[coord] || {}, patch);
                }
              }
              return next;
            });
          }, [sheet, selNorm, pushUndoSnapshot]);

          var handleCopy = useCallback(function (isCut) {
            if (!sheet) return;
            var rows = [];
            var clearMap = {};
            for (var r = selNorm.r1; r <= selNorm.r2; r++) {
              var rowVals = [];
              for (var c = selNorm.c1; c <= selNorm.c2; c++) {
                var coord = makeCoord(r, c);
                rowVals.push(getCellRaw(coord));
                if (isCut) clearMap[coord] = '';
              }
              rows.push(rowVals);
            }
            clipboardRef.current = rows;
            if (isCut) setCellValuesBatch(clearMap);
            setNotice((isCut ? 'Cut ' : 'Copied ') + formatRangeLabel(selStart, selEnd));
          }, [sheet, selNorm, selStart, selEnd, getCellRaw, setCellValuesBatch]);

          var handlePaste = useCallback(function () {
            if (!sheet || !clipboardRef.current) return;
            var matrix = clipboardRef.current;
            var updates = {};
            for (var r = 0; r < matrix.length; r++) {
              for (var c = 0; c < matrix[r].length; c++) {
                updates[makeCoord(selStart.row + r, selStart.col + c)] = matrix[r][c];
              }
            }
            setCellValuesBatch(updates);
          }, [sheet, selStart, setCellValuesBatch]);

          var handleClearSelection = function (mode) {
            if (!sheet) return;
            pushUndoSnapshot();
            if (mode === 'contents' || mode === 'all') {
              var updates = {};
              for (var r = selNorm.r1; r <= selNorm.r2; r++) {
                for (var c = selNorm.c1; c <= selNorm.c2; c++) {
                  updates[makeCoord(r, c)] = '';
                }
              }
              setPendingEdits(function (cur) { return Object.assign({}, cur, updates); });
              setDraftValue('');
            }
            if (mode === 'formats' || mode === 'all') {
              setPendingFormats(function (cur) {
                var next = Object.assign({}, cur);
                for (var r = selNorm.r1; r <= selNorm.r2; r++) {
                  for (var c = selNorm.c1; c <= selNorm.c2; c++) {
                    next[makeCoord(r, c)] = {
                      bold: false,
                      italic: false,
                      underline: false,
                      strike: false,
                      fill_color: '',
                      font_color: '',
                      align: '',
                      border_preset: 'none',
                      number_format: 'General',
                    };
                  }
                }
                return next;
              });
            }
          };

          var handleAutoSum = function (fnName) {
            if (!sheet) return;
            var fn = fnName || 'SUM';
            var r = selStart.row;
            var c = selStart.col;
            var formula;
            if (r > 1) {
              var startR = r - 1;
              while (startR > 1 && getCellRaw(makeCoord(startR - 1, c)) !== '') {
                var prevVal = getCellRaw(makeCoord(startR - 1, c));
                if (!Number.isFinite(Number(prevVal))) break;
                startR--;
              }
              formula = '=' + fn + '(' + makeCoord(startR, c) + ':' + makeCoord(r - 1, c) + ')';
            } else {
              formula = '=' + fn + '(A1:A10)';
            }
            var obj = {};
            obj[activeCoord] = formula;
            setCellValuesBatch(obj);
            setDraftValue(formula);
          };

          var handleSave = useCallback(function (targetPath) {
            if (!sheet || saving) return;
            var editList = Object.keys(pendingEdits).map(function (coord) {
              return { cell: coord, value: pendingEdits[coord] };
            });
            var formatList = Object.keys(pendingFormats).map(function (coord) {
              return Object.assign({ cell: coord }, pendingFormats[coord]);
            });
            if (!targetPath && editList.length === 0 && formatList.length === 0 && Object.keys(pendingColWidths).length === 0) {
              setNotice('All changes are already saved.');
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
                target_path: targetPath || undefined,
                sheet_name: sheet.active_sheet,
                edits: editList,
                formats: formatList,
                column_widths: pendingColWidths,
              }),
            })
              .then(function (data) {
                setSheet(data);
                setPendingEdits({});
                setPendingFormats({});
                setPendingColWidths({});
                setEditing(false);
                if (targetPath) refreshWorkbooks(false);
                setNotice('Saved ' + data.filename + (data.backup_path ? ' (backup snapshot created)' : ''));
              })
              .catch(function (err) {
                setError(err.message || String(err));
              })
              .finally(function () {
                setSaving(false);
              });
          }, [sheet, saving, pendingEdits, pendingFormats, pendingColWidths, refreshWorkbooks]);

          var handleExportCsv = function () {
            if (!sheet) return;
            var lines = [];
            var maxR = Math.max(sheet.max_row || 1, 1);
            var maxC = Math.max(sheet.max_column || 1, 1);
            for (var r = 1; r <= maxR; r++) {
              var cols = [];
              for (var c = 1; c <= maxC; c++) {
                var val = String(getCellDisplay(makeCoord(r, c)) || '');
                if (val.indexOf(',') !== -1 || val.indexOf('"') !== -1 || val.indexOf('\n') !== -1) {
                  val = '"' + val.replace(/"/g, '""') + '"';
                }
                cols.push(val);
              }
              lines.push(cols.join(','));
            }
            var blob = new Blob([lines.join('\n')], { type: 'text/csv;charset=utf-8;' });
            var url = URL.createObjectURL(blob);
            var a = document.createElement('a');
            a.href = url;
            a.download = (sheet.active_sheet || 'sheet') + '.csv';
            a.click();
            URL.revokeObjectURL(url);
            setFileMenuOpen(false);
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
                setNotice('Completed ' + op.replace(/_/g, ' '));
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
                setFileMenuOpen(false);
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
              .then(function (data) { setSummary(data); })
              .catch(function (err) { setError(err.message || String(err)); });
          };

          var handleFindNext = function () {
            if (!sheet || !findText.trim()) return;
            var needle = findText.toLowerCase();
            var maxR = sheet.row_count || 40;
            var maxC = sheet.col_count || 16;
            for (var r = 1; r <= maxR; r++) {
              for (var c = 1; c <= maxC; c++) {
                if (r < selStart.row || (r === selStart.row && c <= selStart.col)) continue;
                var val = String(getCellDisplay(makeCoord(r, c))).toLowerCase();
                if (val.indexOf(needle) !== -1) {
                  selectCell(r, c, false, false);
                  return;
                }
              }
            }
            // Wrap around from top
            for (var r2 = 1; r2 <= maxR; r2++) {
              for (var c2 = 1; c2 <= maxC; c2++) {
                var val2 = String(getCellDisplay(makeCoord(r2, c2))).toLowerCase();
                if (val2.indexOf(needle) !== -1) {
                  selectCell(r2, c2, false, false);
                  return;
                }
              }
            }
            setNotice('No matches for "' + findText + '"');
          };

          var handleReplaceAll = function () {
            if (!sheet || !findText) return;
            var updates = {};
            var count = 0;
            var maxR = sheet.row_count || 40;
            var maxC = sheet.col_count || 16;
            for (var r = 1; r <= maxR; r++) {
              for (var c = 1; c <= maxC; c++) {
                var coord = makeCoord(r, c);
                var raw = String(getCellRaw(coord));
                if (raw.indexOf(findText) !== -1) {
                  updates[coord] = raw.split(findText).join(replaceText);
                  count++;
                }
              }
            }
            if (count > 0) {
              setCellValuesBatch(updates);
              setNotice('Replaced ' + count + ' cell(s). Click Save to write to disk.');
            } else {
              setNotice('No matches found to replace.');
            }
          };

          // Global keyboard shortcuts
          useEffect(function () {
            var onKeyDown = function (e) {
              var tag = (e.target && e.target.tagName) || '';
              var inInput = tag === 'INPUT' || tag === 'SELECT' || tag === 'TEXTAREA';
              var ctrl = e.ctrlKey || e.metaKey;

              if (ctrl && e.key.toLowerCase() === 's' && sheet) {
                e.preventDefault();
                handleSave();
                return;
              }
              if (ctrl && e.key.toLowerCase() === 'f' && sheet) {
                e.preventDefault();
                setFindOpen(true);
                return;
              }
              if (inInput) return;

              if (ctrl && e.key.toLowerCase() === 'z') { e.preventDefault(); handleUndo(); return; }
              if (ctrl && e.key.toLowerCase() === 'y') { e.preventDefault(); handleRedo(); return; }
              if (ctrl && e.key.toLowerCase() === 'c') { e.preventDefault(); handleCopy(false); return; }
              if (ctrl && e.key.toLowerCase() === 'x') { e.preventDefault(); handleCopy(true); return; }
              if (ctrl && e.key.toLowerCase() === 'v') { e.preventDefault(); handlePaste(); return; }
              if (ctrl && e.key.toLowerCase() === 'b') { e.preventDefault(); applyFormatToSelection({ bold: !getCellStyle(activeCoord).bold }); return; }
              if (ctrl && e.key.toLowerCase() === 'i') { e.preventDefault(); applyFormatToSelection({ italic: !getCellStyle(activeCoord).italic }); return; }
              if (ctrl && e.key.toLowerCase() === 'u') { e.preventDefault(); applyFormatToSelection({ underline: !getCellStyle(activeCoord).underline }); return; }

              if (!sheet) return;
              if (e.key === 'ArrowDown') { e.preventDefault(); selectCell(selStart.row + 1, selStart.col, false, e.shiftKey); }
              else if (e.key === 'ArrowUp') { e.preventDefault(); selectCell(selStart.row - 1, selStart.col, false, e.shiftKey); }
              else if (e.key === 'ArrowRight') { e.preventDefault(); selectCell(selStart.row, selStart.col + 1, false, e.shiftKey); }
              else if (e.key === 'ArrowLeft') { e.preventDefault(); selectCell(selStart.row, selStart.col - 1, false, e.shiftKey); }
              else if (e.key === 'Tab') { e.preventDefault(); selectCell(selStart.row, selStart.col + (e.shiftKey ? -1 : 1), false, false); }
              else if (e.key === 'Enter') { e.preventDefault(); selectCell(selStart.row + (e.shiftKey ? -1 : 1), selStart.col, false, false); }
              else if (e.key === 'F2') { e.preventDefault(); setEditing(true); }
              else if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); handleClearSelection('contents'); }
              else if (e.key.length === 1 && !ctrl && !e.altKey) {
                setEditing(true);
                setDraftValue(e.key);
                var upd = {};
                upd[activeCoord] = e.key;
                setCellValuesBatch(upd);
              }
            };
            window.addEventListener('keydown', onKeyDown);
            return function () { window.removeEventListener('keydown', onKeyDown); };
          }, [sheet, selStart, activeCoord, handleSave, handleUndo, handleRedo, handleCopy, handlePaste, applyFormatToSelection, getCellStyle, setCellValuesBatch]);

          // Compute live selection aggregations for Excel Status Bar
          var selStats = (function () {
            if (!sheet) return null;
            var count = 0;
            var nums = [];
            for (var r = selNorm.r1; r <= selNorm.r2; r++) {
              for (var c = selNorm.c1; c <= selNorm.c2; c++) {
                var disp = getCellDisplay(makeCoord(r, c));
                if (disp !== '') {
                  count++;
                  var n = Number(String(disp).replace(/[$,%]/g, ''));
                  if (Number.isFinite(n)) nums.push(n);
                }
              }
            }
            if (count <= 1) return null;
            var sum = nums.reduce(function (a, b) { return a + b; }, 0);
            return {
              count: count,
              sum: Math.round(sum * 1e4) / 1e4,
              avg: nums.length ? Math.round((sum / nums.length) * 1e4) / 1e4 : null,
            };
          })();

          var curStyle = getCellStyle(activeCoord);
          var rowCount = (sheet && sheet.row_count) || 40;
          var colCount = (sheet && sheet.col_count) || 16;
          var colIndices = [];
          for (var ci = 1; ci <= colCount; ci++) colIndices.push(ci);
          var rowIndices = [];
          for (var ri = 1; ri <= rowCount; ri++) rowIndices.push(ri);

          var fillPalette = ['', '#FFF2CC', '#D9EAD3', '#CFE2F3', '#F4CCCC', '#D9D2E9', '#1F4E79'];
          var fontPalette = ['#000000', '#107C41', '#1F4E79', '#C00000', '#7030A0', '#FFFFFF'];

          var startColResize = function (letter, startX, currentPx) {
            var onMove = function (ev) {
              var newPx = Math.max(48, Math.min(420, currentPx + (ev.clientX - startX)));
              var excelWidth = Math.round((newPx / 8) * 100) / 100;
              setPendingColWidths(function (cur) {
                var next = Object.assign({}, cur);
                next[letter] = excelWidth;
                return next;
              });
            };
            var onUp = function () {
              window.removeEventListener('mousemove', onMove);
              window.removeEventListener('mouseup', onUp);
            };
            window.addEventListener('mousemove', onMove);
            window.addEventListener('mouseup', onUp);
          };

          var handleFillDragEnter = function (targetRow, targetCol) {
            if (!isFillingRef.current || !sheet) return;
            var baseRaw = getCellRaw(activeCoord);
            var baseNum = Number(baseRaw);
            var isNum = baseRaw !== '' && Number.isFinite(baseNum);
            var updates = {};
            if (targetRow > selStart.row) {
              for (var r = selStart.row + 1; r <= targetRow; r++) {
                updates[makeCoord(r, selStart.col)] = isNum ? String(baseNum + (r - selStart.row)) : baseRaw;
              }
              setSelEnd({ row: targetRow, col: selStart.col });
            } else if (targetCol > selStart.col) {
              for (var c = selStart.col + 1; c <= targetCol; c++) {
                updates[makeCoord(selStart.row, c)] = isNum ? String(baseNum + (c - selStart.col)) : baseRaw;
              }
              setSelEnd({ row: selStart.row, col: targetCol });
            }
            if (Object.keys(updates).length > 0) setCellValuesBatch(updates);
          };

          function ribbonGroup(label, content) {
            return h(
              'div',
              { className: 'dsh-xl-rgroup' },
              h('div', { className: 'dsh-xl-rgroup-body' }, content),
              h('div', { className: 'dsh-xl-rgroup-label' }, label)
            );
          }

          return h(
            'div',
            { className: 'dsh-xl-root', 'data-testid': 'excel-studio-panel' },

            // 1. Excel Green Title & Quick Access Bar
            h(
              'div',
              { className: 'dsh-xl-titlebar' },
              h(
                'div',
                { className: 'dsh-xl-qat' },
                h(ExcelStudioIcon, { size: 16 }),
                h(
                  'button',
                  {
                    type: 'button',
                    className: 'dsh-xl-qat-btn',
                    title: 'Save (Ctrl+S)',
                    disabled: !sheet || saving,
                    onClick: function () { handleSave(); },
                  },
                  saving ? 'Saving...' : 'Save'
                ),
                h(
                  'button',
                  {
                    type: 'button',
                    className: 'dsh-xl-qat-btn',
                    title: 'Undo (Ctrl+Z)',
                    disabled: undoStack.length === 0,
                    onClick: handleUndo,
                  },
                  '↶'
                ),
                h(
                  'button',
                  {
                    type: 'button',
                    className: 'dsh-xl-qat-btn',
                    title: 'Redo (Ctrl+Y)',
                    disabled: redoStack.length === 0,
                    onClick: handleRedo,
                  },
                  '↷'
                )
              ),
              h(
                'div',
                { className: 'dsh-xl-title-center' },
                h(
                  'select',
                  {
                    className: 'dsh-xl-wb-select',
                    value: (sheet && sheet.file_path) || '',
                    onChange: function (e) { if (e.target.value) openWorkbook(e.target.value); },
                    'aria-label': 'Select workbook',
                  },
                  h('option', { value: '' }, workbooks.length ? 'Select Workbook...' : 'No Workbooks Found'),
                  workbooks.map(function (wb) {
                    return h('option', { key: wb.path, value: wb.path }, wb.name + (wb.is_locked ? ' (locked)' : ''));
                  })
                ),
                dirtyCount > 0 ? h('span', { className: 'dsh-xl-pill' }, dirtyCount + ' unsaved') : null
              ),
              h(
                'div',
                { className: 'dsh-xl-qat' },
                h(
                  'button',
                  {
                    type: 'button',
                    className: 'dsh-xl-qat-btn',
                    title: 'Find & Replace (Ctrl+F)',
                    disabled: !sheet,
                    onClick: function () { setFindOpen(function (v) { return !v; }); },
                  },
                  'Find'
                ),
                h(
                  'button',
                  {
                    type: 'button',
                    className: 'dsh-xl-qat-btn',
                    title: 'Refresh Workbooks',
                    onClick: function () {
                      refreshWorkbooks(false);
                      if (sheet) openWorkbook(sheet.file_path, sheet.active_sheet);
                    },
                  },
                  loading ? '...' : '↻'
                )
              )
            ),

            // 2. Real Excel Menu Tab Bar (File | Home | Insert | Formulas | Data | View)
            h(
              'div',
              { className: 'dsh-xl-tabs-bar', role: 'tablist' },
              h(
                'button',
                {
                  type: 'button',
                  className: 'dsh-xl-menu-tab dsh-xl-file-tab',
                  onClick: function (e) {
                    e.stopPropagation();
                    setFileMenuOpen(function (v) { return !v; });
                  },
                },
                'File'
              ),
              ['Home', 'Insert', 'Formulas', 'Data', 'View'].map(function (tab) {
                return h(
                  'button',
                  {
                    key: tab,
                    type: 'button',
                    role: 'tab',
                    'aria-selected': ribbonTab === tab,
                    className: 'dsh-xl-menu-tab ' + (ribbonTab === tab ? 'dsh-xl-menu-tab-active' : ''),
                    onClick: function () {
                      setRibbonTab(tab);
                      setFileMenuOpen(false);
                    },
                  },
                  tab
                );
              })
            ),

            // File Backstage Dropdown Menu
            fileMenuOpen
              ? h(
                  'div',
                  {
                    className: 'dsh-xl-file-menu',
                    onClick: function (e) { e.stopPropagation(); },
                  },
                  h(
                    'button',
                    {
                      type: 'button',
                      className: 'dsh-xl-menu-item',
                      onClick: function () { setShowNewForm(true); setFileMenuOpen(false); },
                    },
                    h('span', null, 'New Blank Workbook'),
                    h('span', { style: { opacity: 0.55 } }, 'Ctrl+N')
                  ),
                  h(
                    'button',
                    {
                      type: 'button',
                      className: 'dsh-xl-menu-item',
                      disabled: !sheet,
                      onClick: function () { handleSave(); setFileMenuOpen(false); },
                    },
                    h('span', null, 'Save'),
                    h('span', { style: { opacity: 0.55 } }, 'Ctrl+S')
                  ),
                  h(
                    'button',
                    {
                      type: 'button',
                      className: 'dsh-xl-menu-item',
                      disabled: !sheet,
                      onClick: function () {
                        var dest = window.prompt('Save As (full .xlsx path in allowed directory):', sheet.file_path.replace(/\.xlsx$/i, '_copy.xlsx'));
                        if (dest && dest.trim()) handleSave(dest.trim());
                        setFileMenuOpen(false);
                      },
                    },
                    h('span', null, 'Save As...'),
                    h('span', { style: { opacity: 0.55 } }, '.xlsx')
                  ),
                  h(
                    'button',
                    {
                      type: 'button',
                      className: 'dsh-xl-menu-item',
                      disabled: !sheet,
                      onClick: handleExportCsv,
                    },
                    h('span', null, 'Export Active Sheet as CSV'),
                    h('span', { style: { opacity: 0.55 } }, '.csv')
                  ),
                  h('div', { className: 'dsh-xl-menu-sep' }),
                  h(
                    'form',
                    {
                      style: { display: 'flex', gap: '4px', padding: '4px' },
                      onSubmit: function (e) {
                        e.preventDefault();
                        if (openPathInput.trim()) openWorkbook(openPathInput.trim());
                      },
                    },
                    h('input', {
                      className: 'dsh-xl-rinput',
                      style: { flex: 1 },
                      placeholder: 'Open .xlsx path...',
                      value: openPathInput,
                      onChange: function (e) { setOpenPathInput(e.target.value); },
                    }),
                    h('button', { type: 'submit', className: 'dsh-xl-rbtn dsh-xl-rbtn-active' }, 'Open')
                  )
                )
              : null,

            // 3. Real Excel Grouped Ribbon
            h(
              'div',
              { className: 'dsh-xl-ribbon' },
              ribbonTab === 'Home'
                ? [
                    ribbonGroup(
                      'Clipboard',
                      h(
                        'div',
                        { className: 'dsh-xl-rgroup-row' },
                        h('button', { type: 'button', className: 'dsh-xl-rbtn dsh-xl-rbtn-tall', disabled: !sheet, onClick: handlePaste }, h('span', { style: { fontSize: '15px' } }, '📋'), h('span', null, 'Paste')),
                        h(
                          'div',
                          { style: { display: 'flex', flexDirection: 'column', gap: '2px' } },
                          h('button', { type: 'button', className: 'dsh-xl-rbtn', disabled: !sheet, onClick: function () { handleCopy(true); } }, '✂ Cut'),
                          h('button', { type: 'button', className: 'dsh-xl-rbtn', disabled: !sheet, onClick: function () { handleCopy(false); } }, '⧉ Copy')
                        )
                      )
                    ),
                    ribbonGroup(
                      'Font',
                      [
                        h(
                          'div',
                          { key: 'r1', className: 'dsh-xl-rgroup-row' },
                          h(
                            'select',
                            {
                              className: 'dsh-xl-rselect',
                              style: { width: '112px' },
                              disabled: !sheet,
                              value: curStyle.fontFamily || curStyle.font_family || 'Calibri',
                              onChange: function (e) { applyFormatToSelection({ font_family: e.target.value }); },
                            },
                            ['Calibri', 'Arial', 'Segoe UI', 'Times New Roman', 'Consolas', 'Verdana'].map(function (f) {
                              return h('option', { key: f, value: f }, f);
                            })
                          ),
                          h(
                            'select',
                            {
                              className: 'dsh-xl-rselect',
                              style: { width: '50px' },
                              disabled: !sheet,
                              value: String(curStyle.fontSize || curStyle.font_size || 11),
                              onChange: function (e) { applyFormatToSelection({ font_size: Number(e.target.value) }); },
                            },
                            [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36].map(function (sz) {
                              return h('option', { key: sz, value: String(sz) }, String(sz));
                            })
                          )
                        ),
                        h(
                          'div',
                          { key: 'r2', className: 'dsh-xl-rgroup-row' },
                          h('button', { type: 'button', className: 'dsh-xl-rbtn ' + (curStyle.bold ? 'dsh-xl-rbtn-active' : ''), disabled: !sheet, title: 'Bold (Ctrl+B)', onClick: function () { applyFormatToSelection({ bold: !curStyle.bold }); } }, h('strong', null, 'B')),
                          h('button', { type: 'button', className: 'dsh-xl-rbtn ' + (curStyle.italic ? 'dsh-xl-rbtn-active' : ''), disabled: !sheet, title: 'Italic (Ctrl+I)', onClick: function () { applyFormatToSelection({ italic: !curStyle.italic }); } }, h('em', null, 'I')),
                          h('button', { type: 'button', className: 'dsh-xl-rbtn ' + (curStyle.underline ? 'dsh-xl-rbtn-active' : ''), disabled: !sheet, title: 'Underline (Ctrl+U)', onClick: function () { applyFormatToSelection({ underline: !curStyle.underline }); } }, h('u', null, 'U')),
                          h('button', { type: 'button', className: 'dsh-xl-rbtn ' + (curStyle.strike ? 'dsh-xl-rbtn-active' : ''), disabled: !sheet, title: 'Strikethrough', onClick: function () { applyFormatToSelection({ strike: !curStyle.strike }); } }, h('s', null, 'S')),
                          h(
                            'select',
                            {
                              className: 'dsh-xl-rselect',
                              style: { width: '76px' },
                              disabled: !sheet,
                              title: 'Borders',
                              value: '',
                              onChange: function (e) {
                                if (e.target.value) applyFormatToSelection({ border_preset: e.target.value });
                              },
                            },
                            h('option', { value: '' }, 'Borders'),
                            h('option', { value: 'all' }, 'All Borders'),
                            h('option', { value: 'bottom' }, 'Bottom'),
                            h('option', { value: 'top_bottom' }, 'Top & Double Bottom'),
                            h('option', { value: 'none' }, 'No Border')
                          ),
                          fillPalette.map(function (c) {
                            return h('button', {
                              key: 'bg-' + (c || 'none'),
                              type: 'button',
                              disabled: !sheet,
                              title: c ? 'Fill ' + c : 'No Fill',
                              className: 'dsh-xl-swatch',
                              style: {
                                background: c || 'transparent',
                                backgroundImage: c ? undefined : 'linear-gradient(135deg, transparent 45%, #c00 50%, transparent 55%)',
                              },
                              onClick: function () { applyFormatToSelection({ fill_color: c }); },
                            });
                          }),
                          fontPalette.map(function (fc) {
                            return h('button', {
                              key: 'fg-' + fc,
                              type: 'button',
                              disabled: !sheet,
                              title: 'Font Color ' + fc,
                              className: 'dsh-xl-rbtn',
                              style: { padding: '0 4px', color: fc === '#FFFFFF' ? '#888' : fc, fontWeight: 700 },
                              onClick: function () { applyFormatToSelection({ font_color: fc }); },
                            }, 'A');
                          })
                        ),
                      ]
                    ),
                    ribbonGroup(
                      'Alignment',
                      [
                        h(
                          'div',
                          { key: 'a1', className: 'dsh-xl-rgroup-row' },
                          ['top', 'middle', 'bottom'].map(function (va) {
                            return h('button', {
                              key: va,
                              type: 'button',
                              disabled: !sheet,
                              title: 'Vertical ' + va,
                              className: 'dsh-xl-rbtn ' + (curStyle.valign === va ? 'dsh-xl-rbtn-active' : ''),
                              onClick: function () { applyFormatToSelection({ valign: va }); },
                            }, va === 'top' ? '⤒' : va === 'middle' ? '⇳' : '⤓');
                          }),
                          h('button', {
                            type: 'button',
                            disabled: !sheet,
                            className: 'dsh-xl-rbtn ' + (curStyle.wrap ? 'dsh-xl-rbtn-active' : ''),
                            onClick: function () { applyFormatToSelection({ wrap: !curStyle.wrap }); },
                          }, 'Wrap Text')
                        ),
                        h(
                          'div',
                          { key: 'a2', className: 'dsh-xl-rgroup-row' },
                          ['left', 'center', 'right'].map(function (ha) {
                            return h('button', {
                              key: ha,
                              type: 'button',
                              disabled: !sheet,
                              title: 'Align ' + ha,
                              className: 'dsh-xl-rbtn ' + (curStyle.align === ha ? 'dsh-xl-rbtn-active' : ''),
                              onClick: function () { applyFormatToSelection({ align: ha }); },
                            }, ha === 'left' ? '≡L' : ha === 'center' ? '≡C' : '≡R');
                          })
                        ),
                      ]
                    ),
                    ribbonGroup(
                      'Number',
                      [
                        h(
                          'div',
                          { key: 'n1', className: 'dsh-xl-rgroup-row' },
                          h(
                            'select',
                            {
                              className: 'dsh-xl-rselect',
                              style: { width: '112px' },
                              disabled: !sheet,
                              value: curStyle.numberFormat || curStyle.number_format || 'General',
                              onChange: function (e) { applyFormatToSelection({ number_format: e.target.value }); },
                            },
                            h('option', { value: 'General' }, 'General'),
                            h('option', { value: '#,##0' }, 'Number'),
                            h('option', { value: '#,##0.00' }, 'Number (0.00)'),
                            h('option', { value: '$#,##0.00' }, 'Currency'),
                            h('option', { value: '0.0%' }, 'Percentage'),
                            h('option', { value: 'yyyy-mm-dd' }, 'Short Date'),
                            h('option', { value: '@' }, 'Text')
                          )
                        ),
                        h(
                          'div',
                          { key: 'n2', className: 'dsh-xl-rgroup-row' },
                          h('button', { type: 'button', className: 'dsh-xl-rbtn', disabled: !sheet, title: 'Currency', onClick: function () { applyFormatToSelection({ number_format: '$#,##0.00' }); } }, '$'),
                          h('button', { type: 'button', className: 'dsh-xl-rbtn', disabled: !sheet, title: 'Percent', onClick: function () { applyFormatToSelection({ number_format: '0.0%' }); } }, '%'),
                          h('button', { type: 'button', className: 'dsh-xl-rbtn', disabled: !sheet, title: 'Comma Style', onClick: function () { applyFormatToSelection({ number_format: '#,##0' }); } }, ',')
                        ),
                      ]
                    ),
                    ribbonGroup(
                      'Cells',
                      [
                        h(
                          'div',
                          { key: 'c1', className: 'dsh-xl-rgroup-row' },
                          h('button', { type: 'button', className: 'dsh-xl-rbtn', disabled: !sheet || saving, onClick: function () { runSheetOp('insert_rows', { index: selStart.row, amount: 1 }); } }, '+ Row'),
                          h('button', { type: 'button', className: 'dsh-xl-rbtn', disabled: !sheet || saving, onClick: function () { runSheetOp('insert_columns', { index: selStart.col, amount: 1 }); } }, '+ Col')
                        ),
                        h(
                          'div',
                          { key: 'c2', className: 'dsh-xl-rgroup-row' },
                          h('button', { type: 'button', className: 'dsh-xl-rbtn', disabled: !sheet || saving, onClick: function () { if (window.confirm('Delete row ' + selStart.row + '?')) runSheetOp('delete_rows', { index: selStart.row, amount: 1 }); } }, '− Row'),
                          h('button', { type: 'button', className: 'dsh-xl-rbtn', disabled: !sheet || saving, onClick: function () { if (window.confirm('Delete column ' + colIndexToLetter(selStart.col) + '?')) runSheetOp('delete_columns', { index: selStart.col, amount: 1 }); } }, '− Col')
                        ),
                      ]
                    ),
                    ribbonGroup(
                      'Editing',
                      [
                        h(
                          'div',
                          { key: 'e1', className: 'dsh-xl-rgroup-row' },
                          h('button', { type: 'button', className: 'dsh-xl-rbtn', disabled: !sheet, onClick: function () { handleAutoSum('SUM'); } }, 'Σ AutoSum'),
                          h('button', { type: 'button', className: 'dsh-xl-rbtn', disabled: !sheet, onClick: function () { setFindOpen(true); } }, '🔍 Find')
                        ),
                        h(
                          'div',
                          { key: 'e2', className: 'dsh-xl-rgroup-row' },
                          h('button', { type: 'button', className: 'dsh-xl-rbtn', disabled: !sheet, onClick: function () { handleClearSelection('contents'); } }, 'Clear Text'),
                          h('button', { type: 'button', className: 'dsh-xl-rbtn', disabled: !sheet, onClick: function () { handleClearSelection('formats'); } }, 'Clear Format')
                        ),
                      ]
                    ),
                  ]
                : ribbonTab === 'Insert'
                ? [
                    ribbonGroup(
                      'Structure',
                      h(
                        'div',
                        { className: 'dsh-xl-rgroup-row' },
                        h('button', { type: 'button', className: 'dsh-xl-rbtn dsh-xl-rbtn-tall', disabled: !sheet, onClick: function () { runSheetOp('insert_rows', { index: selStart.row, amount: 1 }); } }, h('span', null, '➕'), h('span', null, 'Row Above')),
                        h('button', { type: 'button', className: 'dsh-xl-rbtn dsh-xl-rbtn-tall', disabled: !sheet, onClick: function () { runSheetOp('insert_columns', { index: selStart.col, amount: 1 }); } }, h('span', null, '➕'), h('span', null, 'Col Left')),
                        h('button', {
                          type: 'button',
                          className: 'dsh-xl-rbtn dsh-xl-rbtn-tall',
                          disabled: !sheet,
                          onClick: function () {
                            var nm = window.prompt('New worksheet name:', 'Sheet' + (((sheet && sheet.sheets) || []).length + 1));
                            if (nm && nm.trim()) runSheetOp('create_sheet', { new_name: nm.trim() });
                          },
                        }, h('span', null, '📄'), h('span', null, 'New Sheet'))
                      )
                    ),
                    ribbonGroup(
                      'Table Styles',
                      h(
                        'div',
                        { className: 'dsh-xl-rgroup-row' },
                        h('button', {
                          type: 'button',
                          className: 'dsh-xl-rbtn dsh-xl-rbtn-tall',
                          disabled: !sheet,
                          onClick: function () {
                            applyFormatToSelection({ bold: true, fill_color: '#1F4E79', font_color: '#FFFFFF', align: 'center', border_preset: 'all' });
                          },
                        }, h('span', null, '▦'), h('span', null, 'Header Row')),
                        h('button', {
                          type: 'button',
                          className: 'dsh-xl-rbtn dsh-xl-rbtn-tall',
                          disabled: !sheet,
                          onClick: function () {
                            applyFormatToSelection({ bold: true, fill_color: '#D9EAD3', border_preset: 'top_bottom' });
                          },
                        }, h('span', null, 'Σ'), h('span', null, 'Total Row'))
                      )
                    ),
                  ]
                : ribbonTab === 'Formulas'
                ? [
                    ribbonGroup(
                      'Function Library',
                      h(
                        'div',
                        { className: 'dsh-xl-rgroup-row' },
                        ['SUM', 'AVERAGE', 'COUNT', 'MIN', 'MAX'].map(function (fn) {
                          return h('button', {
                            key: fn,
                            type: 'button',
                            className: 'dsh-xl-rbtn dsh-xl-rbtn-tall',
                            disabled: !sheet,
                            onClick: function () { handleAutoSum(fn); },
                          }, h('span', { style: { fontWeight: 700, color: '#107c41' } }, 'fx'), h('span', null, fn));
                        })
                      )
                    ),
                    ribbonGroup(
                      'Formula Auditing',
                      h(
                        'div',
                        { className: 'dsh-xl-rgroup-row' },
                        h('button', {
                          type: 'button',
                          className: 'dsh-xl-rbtn dsh-xl-rbtn-tall ' + (showFormulas ? 'dsh-xl-rbtn-active' : ''),
                          disabled: !sheet,
                          onClick: function () { setShowFormulas(function (v) { return !v; }); },
                        }, h('span', null, 'fx'), h('span', null, 'Show Formulas'))
                      )
                    ),
                  ]
                : ribbonTab === 'Data'
                ? [
                    ribbonGroup(
                      'Sort',
                      h(
                        'div',
                        { className: 'dsh-xl-rgroup-row' },
                        h('button', { type: 'button', className: 'dsh-xl-rbtn dsh-xl-rbtn-tall', disabled: !sheet || saving, onClick: function () { runSheetOp('sort_column', { index: selStart.col, ascending: true }); } }, h('span', null, 'A↓Z'), h('span', null, 'Sort A→Z')),
                        h('button', { type: 'button', className: 'dsh-xl-rbtn dsh-xl-rbtn-tall', disabled: !sheet || saving, onClick: function () { runSheetOp('sort_column', { index: selStart.col, ascending: false }); } }, h('span', null, 'Z↓A'), h('span', null, 'Sort Z→A'))
                      )
                    ),
                    ribbonGroup(
                      'Analysis',
                      h(
                        'div',
                        { className: 'dsh-xl-rgroup-row' },
                        h('button', {
                          type: 'button',
                          className: 'dsh-xl-rbtn dsh-xl-rbtn-tall ' + (summary ? 'dsh-xl-rbtn-active' : ''),
                          disabled: !sheet,
                          onClick: function () { if (summary) setSummary(null); else handleSummary(); },
                        }, h('span', null, '📊'), h('span', null, 'Summary Stats'))
                      )
                    ),
                  ]
                : [
                    ribbonGroup(
                      'Show',
                      h(
                        'div',
                        { className: 'dsh-xl-rgroup-row' },
                        h('button', { type: 'button', className: 'dsh-xl-rbtn ' + (showGridlines ? 'dsh-xl-rbtn-active' : ''), onClick: function () { setShowGridlines(function (v) { return !v; }); } }, '☑ Gridlines'),
                        h('button', { type: 'button', className: 'dsh-xl-rbtn ' + (showFormulaBar ? 'dsh-xl-rbtn-active' : ''), onClick: function () { setShowFormulaBar(function (v) { return !v; }); } }, '☑ Formula Bar'),
                        h('button', { type: 'button', className: 'dsh-xl-rbtn ' + (showHeadings ? 'dsh-xl-rbtn-active' : ''), onClick: function () { setShowHeadings(function (v) { return !v; }); } }, '☑ Headings')
                      )
                    ),
                    ribbonGroup(
                      'Window',
                      h(
                        'div',
                        { className: 'dsh-xl-rgroup-row' },
                        h('button', { type: 'button', className: 'dsh-xl-rbtn dsh-xl-rbtn-tall ' + (freezeTopRow ? 'dsh-xl-rbtn-active' : ''), onClick: function () { setFreezeTopRow(function (v) { return !v; }); } }, h('span', null, '❄'), h('span', null, 'Freeze Top Row'))
                      )
                    ),
                    ribbonGroup(
                      'Zoom',
                      h(
                        'div',
                        { className: 'dsh-xl-rgroup-row' },
                        [75, 100, 125, 150].map(function (z) {
                          return h('button', { key: z, type: 'button', className: 'dsh-xl-rbtn ' + (zoom === z ? 'dsh-xl-rbtn-active' : ''), onClick: function () { setZoom(z); } }, z + '%');
                        })
                      )
                    ),
                  ]
            ),

            // Optional New Workbook inline bar
            showNewForm
              ? h(
                  'form',
                  { className: 'dsh-xl-findbar', onSubmit: handleCreateWorkbook },
                  h('strong', null, 'New Workbook:'),
                  h('input', {
                    className: 'dsh-xl-rinput',
                    style: { width: '220px' },
                    value: newFilename,
                    onChange: function (e) { setNewFilename(e.target.value); },
                    placeholder: 'Book1.xlsx',
                  }),
                  h('span', { style: { opacity: 0.65 } }, allowedDirs[0] ? 'in ' + allowedDirs[0] : ''),
                  h('button', { type: 'submit', className: 'dsh-xl-btn-primary', disabled: saving }, 'Create'),
                  h('button', { type: 'button', className: 'dsh-xl-rbtn', onClick: function () { setShowNewForm(false); } }, 'Cancel')
                )
              : null,

            // Optional Find & Replace bar
            findOpen && sheet
              ? h(
                  'div',
                  { className: 'dsh-xl-findbar' },
                  h('strong', null, 'Find:'),
                  h('input', { className: 'dsh-xl-rinput', style: { width: '160px' }, value: findText, onChange: function (e) { setFindText(e.target.value); }, onKeyDown: function (e) { if (e.key === 'Enter') handleFindNext(); }, placeholder: 'Search text...' }),
                  h('button', { type: 'button', className: 'dsh-xl-rbtn', onClick: handleFindNext }, 'Find Next'),
                  h('strong', { style: { marginLeft: '6px' } }, 'Replace:'),
                  h('input', { className: 'dsh-xl-rinput', style: { width: '160px' }, value: replaceText, onChange: function (e) { setReplaceText(e.target.value); }, placeholder: 'Replace with...' }),
                  h('button', { type: 'button', className: 'dsh-xl-rbtn', onClick: handleReplaceAll }, 'Replace All'),
                  h('button', { type: 'button', className: 'dsh-xl-rbtn', style: { marginLeft: 'auto' }, onClick: function () { setFindOpen(false); } }, '✕')
                )
              : null,

            // Error / Notice banners
            error ? h('div', { className: 'dsh-xl-banner dsh-xl-banner-err', role: 'alert' }, h('span', null, error), h('button', { type: 'button', className: 'dsh-xl-rbtn', onClick: function () { setError(''); } }, 'Dismiss')) : null,
            notice ? h('div', { className: 'dsh-xl-banner dsh-xl-banner-info' }, h('span', null, notice), h('button', { type: 'button', className: 'dsh-xl-rbtn', onClick: function () { setNotice(''); } }, 'OK')) : null,

            // Summary drawer
            summary
              ? h(
                  'div',
                  { style: { padding: '8px 14px', borderBottom: '1px solid rgba(127,127,127,0.22)', maxHeight: '160px', overflow: 'auto', background: 'var(--dsw-alias-bg-secondary, #f8f9fa)' } },
                  h('div', { style: { display: 'flex', justifyContent: 'space-between', fontWeight: 600, marginBottom: '4px' } }, h('span', null, 'Dataset Summary (' + (summary.total_rows || 0) + ' rows)'), h('button', { type: 'button', className: 'dsh-xl-rbtn', onClick: function () { setSummary(null); } }, 'Close')),
                  h('pre', { style: { margin: 0, fontSize: '11px' } }, JSON.stringify(summary.summary || summary, null, 2))
                )
              : null,

            // 4. Excel Name Box & Formula Bar
            sheet && showFormulaBar
              ? h(
                  'div',
                  { className: 'dsh-xl-fbar' },
                  h('input', {
                    className: 'dsh-xl-namebox',
                    value: nameBoxText,
                    title: 'Name Box (type cell coordinate and press Enter)',
                    onChange: function (e) { setNameBoxText(e.target.value); },
                    onKeyDown: function (e) {
                      if (e.key === 'Enter') {
                        var parsed = parseCoord(nameBoxText);
                        if (parsed) selectCell(parsed.row, parsed.col, false, false);
                      }
                    },
                  }),
                  h('button', { type: 'button', className: 'dsh-xl-rbtn', title: 'Cancel', onClick: function () { setDraftValue(getCellRaw(activeCoord)); setEditing(false); } }, '✕'),
                  h('button', {
                    type: 'button',
                    className: 'dsh-xl-rbtn',
                    title: 'Enter',
                    onClick: function () {
                      var upd = {};
                      upd[activeCoord] = draftValue;
                      setCellValuesBatch(upd);
                      setEditing(false);
                    },
                  }, '✓'),
                  h('button', { type: 'button', className: 'dsh-xl-rbtn', title: 'Insert SUM', style: { fontStyle: 'italic', fontFamily: 'Georgia, serif', fontWeight: 700 }, onClick: function () { handleAutoSum('SUM'); } }, 'fx'),
                  h('input', {
                    className: 'dsh-xl-finput',
                    value: draftValue,
                    placeholder: 'Enter value or formula (e.g. =SUM(B2:B10))',
                    onChange: function (e) {
                      var val = e.target.value;
                      setDraftValue(val);
                      var upd = {};
                      upd[activeCoord] = val;
                      setCellValuesBatch(upd);
                    },
                    onKeyDown: function (e) {
                      if (e.key === 'Enter') {
                        e.preventDefault();
                        selectCell(selStart.row + 1, selStart.col, false, false);
                      } else if (e.key === 'Tab') {
                        e.preventDefault();
                        selectCell(selStart.row, selStart.col + 1, false, false);
                      } else if (e.key === 'Escape') {
                        setDraftValue(getCellRaw(activeCoord));
                        setEditing(false);
                      }
                    },
                  })
                )
              : null,

            // 5. Excel Interactive Grid or Empty Start Screen
            sheet
              ? h(
                  'div',
                  { className: 'dsh-xl-grid-wrap ' + (showGridlines ? 'dsh-xl-gridlines' : '') },
                  h(
                    'table',
                    { className: 'dsh-xl-table', style: { zoom: zoom / 100 } },
                    showHeadings
                      ? h(
                          'thead',
                          null,
                          h(
                            'tr',
                            null,
                            h('th', { className: 'dsh-xl-corner' }),
                            colIndices.map(function (c) {
                              var letter = colIndexToLetter(c);
                              var w = pendingColWidths[letter] || (sheet.column_widths && sheet.column_widths[letter]);
                              var px = w ? Math.max(52, Math.min(420, Math.round(w * 8))) : 100;
                              var colActive = c >= selNorm.c1 && c <= selNorm.c2;
                              return h(
                                'th',
                                {
                                  key: letter,
                                  className: 'dsh-xl-colhdr ' + (colActive ? 'dsh-xl-colhdr-active' : ''),
                                  style: { width: px + 'px', minWidth: px + 'px' },
                                  onClick: function () {
                                    setSelStart({ row: 1, col: c });
                                    setSelEnd({ row: rowCount, col: c });
                                  },
                                },
                                letter,
                                h('span', {
                                  className: 'dsh-xl-resizer',
                                  onMouseDown: function (e) {
                                    e.stopPropagation();
                                    startColResize(letter, e.clientX, px);
                                  },
                                })
                              );
                            })
                          )
                        )
                      : null,
                    h(
                      'tbody',
                      null,
                      rowIndices.map(function (r) {
                        var rowActive = r >= selNorm.r1 && r <= selNorm.r2;
                        return h(
                          'tr',
                          { key: r, className: freezeTopRow && r === 1 ? 'dsh-xl-frozen-row' : '' },
                          showHeadings
                            ? h(
                                'td',
                                {
                                  className: 'dsh-xl-rowhdr ' + (rowActive ? 'dsh-xl-rowhdr-active' : ''),
                                  onClick: function () {
                                    setSelStart({ row: r, col: 1 });
                                    setSelEnd({ row: r, col: colCount });
                                  },
                                },
                                r
                              )
                            : null,
                          colIndices.map(function (c) {
                            var coord = makeCoord(r, c);
                            var isActive = selStart.row === r && selStart.col === c;
                            var inRange = r >= selNorm.r1 && r <= selNorm.r2 && c >= selNorm.c1 && c <= selNorm.c2 && !(selNorm.r1 === selNorm.r2 && selNorm.c1 === selNorm.c2);
                            var isBottomRight = r === selNorm.r2 && c === selNorm.c2;
                            var isDirty =
                              Object.prototype.hasOwnProperty.call(pendingEdits, coord) ||
                              Object.prototype.hasOwnProperty.call(pendingFormats, coord);
                            var st = getCellStyle(coord);
                            var borderPreset = st.border_preset;
                            var cellCss = {
                              fontFamily: st.fontFamily || st.font_family || undefined,
                              fontSize: (st.fontSize || st.font_size) ? (st.fontSize || st.font_size) + 'px' : undefined,
                              fontWeight: st.bold ? '700' : undefined,
                              fontStyle: st.italic ? 'italic' : undefined,
                              textDecoration: [st.underline ? 'underline' : '', st.strike ? 'line-through' : ''].filter(Boolean).join(' ') || undefined,
                              color: st.font_color || st.color || undefined,
                              backgroundColor: st.fill_color || st.fill || undefined,
                              textAlign: st.align || undefined,
                              verticalAlign: st.valign || undefined,
                              whiteSpace: st.wrap ? 'normal' : 'nowrap',
                              border: borderPreset === 'all' ? '1px solid #666' : undefined,
                              borderBottom: borderPreset === 'bottom' ? '1px solid #444' : borderPreset === 'top_bottom' ? '3px double #444' : undefined,
                              borderTop: borderPreset === 'top_bottom' ? '1px solid #444' : undefined,
                            };
                            var cls =
                              'dsh-xl-cell' +
                              (inRange ? ' dsh-xl-cell-inrange' : '') +
                              (isActive ? ' dsh-xl-cell-active' : '') +
                              (isDirty ? ' dsh-xl-cell-dirty' : '');

                            return h(
                              'td',
                              {
                                key: coord,
                                className: cls,
                                style: cellCss,
                                onMouseDown: function (e) {
                                  if (e.button === 2) return;
                                  setFileMenuOpen(false);
                                  isDraggingRef.current = true;
                                  selectCell(r, c, false, e.shiftKey);
                                },
                                onMouseEnter: function () {
                                  if (isDraggingRef.current) setSelEnd({ row: r, col: c });
                                  else if (isFillingRef.current) handleFillDragEnter(r, c);
                                },
                                onDoubleClick: function () {
                                  selectCell(r, c, true, false);
                                },
                                onContextMenu: function (e) {
                                  e.preventDefault();
                                  selectCell(r, c, false, false);
                                  setCtxMenu({ x: e.clientX, y: e.clientY, row: r, col: c });
                                },
                              },
                              isActive && editing
                                ? h('input', {
                                    className: 'dsh-xl-editor-input',
                                    autoFocus: true,
                                    value: draftValue,
                                    onChange: function (e) {
                                      var v = e.target.value;
                                      setDraftValue(v);
                                      var upd = {};
                                      upd[coord] = v;
                                      setCellValuesBatch(upd);
                                    },
                                    onBlur: function () { setEditing(false); },
                                    onKeyDown: function (e) {
                                      if (e.key === 'Enter') {
                                        e.preventDefault();
                                        selectCell(r + 1, c, false, false);
                                      } else if (e.key === 'Tab') {
                                        e.preventDefault();
                                        selectCell(r, c + 1, false, false);
                                      } else if (e.key === 'Escape') {
                                        setDraftValue(getCellRaw(coord));
                                        setEditing(false);
                                      }
                                    },
                                  })
                                : getCellDisplay(coord),
                              isBottomRight && !editing
                                ? h('span', {
                                    className: 'dsh-xl-fillhandle',
                                    title: 'Drag to AutoFill',
                                    onMouseDown: function (e) {
                                      e.stopPropagation();
                                      isFillingRef.current = true;
                                    },
                                  })
                                : null
                            );
                          })
                        );
                      })
                    )
                  )
                )
              : h(
                  'div',
                  { className: 'dsh-xl-empty' },
                  h(ExcelStudioIcon, { size: 42 }),
                  h('div', { style: { fontWeight: 700, fontSize: '16px' } }, 'DeepSeek Excel Studio'),
                  h('div', { style: { maxWidth: '420px', opacity: 0.78, lineHeight: '1.5' } }, 'Create a new blank workbook or open a local .xlsx / .xlsm file from the File menu above.'),
                  h('button', { type: 'button', className: 'dsh-xl-btn-primary', onClick: function () { setShowNewForm(true); } }, '+ New Blank Workbook')
                ),

            // Right-Click Cell Context Menu
            ctxMenu
              ? h(
                  'div',
                  {
                    className: 'dsh-xl-ctxmenu',
                    style: { left: ctxMenu.x + 'px', top: ctxMenu.y + 'px' },
                    onClick: function (e) { e.stopPropagation(); },
                  },
                  h('button', { type: 'button', className: 'dsh-xl-menu-item', onClick: function () { handleCopy(true); setCtxMenu(null); } }, h('span', null, 'Cut'), h('span', { style: { opacity: 0.5 } }, 'Ctrl+X')),
                  h('button', { type: 'button', className: 'dsh-xl-menu-item', onClick: function () { handleCopy(false); setCtxMenu(null); } }, h('span', null, 'Copy'), h('span', { style: { opacity: 0.5 } }, 'Ctrl+C')),
                  h('button', { type: 'button', className: 'dsh-xl-menu-item', onClick: function () { handlePaste(); setCtxMenu(null); } }, h('span', null, 'Paste'), h('span', { style: { opacity: 0.5 } }, 'Ctrl+V')),
                  h('div', { className: 'dsh-xl-menu-sep' }),
                  h('button', { type: 'button', className: 'dsh-xl-menu-item', onClick: function () { runSheetOp('insert_rows', { index: ctxMenu.row, amount: 1 }); setCtxMenu(null); } }, 'Insert Row Above'),
                  h('button', { type: 'button', className: 'dsh-xl-menu-item', onClick: function () { if (window.confirm('Delete row ' + ctxMenu.row + '?')) runSheetOp('delete_rows', { index: ctxMenu.row, amount: 1 }); setCtxMenu(null); } }, 'Delete Row'),
                  h('button', { type: 'button', className: 'dsh-xl-menu-item', onClick: function () { runSheetOp('insert_columns', { index: ctxMenu.col, amount: 1 }); setCtxMenu(null); } }, 'Insert Column Left'),
                  h('button', { type: 'button', className: 'dsh-xl-menu-item', onClick: function () { if (window.confirm('Delete column ' + colIndexToLetter(ctxMenu.col) + '?')) runSheetOp('delete_columns', { index: ctxMenu.col, amount: 1 }); setCtxMenu(null); } }, 'Delete Column'),
                  h('div', { className: 'dsh-xl-menu-sep' }),
                  h('button', { type: 'button', className: 'dsh-xl-menu-item', onClick: function () { handleClearSelection('contents'); setCtxMenu(null); } }, 'Clear Contents'),
                  h('button', { type: 'button', className: 'dsh-xl-menu-item', onClick: function () { applyFormatToSelection({ bold: !curStyle.bold }); setCtxMenu(null); } }, 'Toggle Bold')
                )
              : null,

            // 6. Excel Bottom Sheet Tab Bar
            sheet
              ? h(
                  'div',
                  { className: 'dsh-xl-sheetbar' },
                  (sheet.sheets || []).map(function (name) {
                    var active = name === sheet.active_sheet;
                    return h(
                      'button',
                      {
                        key: name,
                        type: 'button',
                        className: 'dsh-xl-sheettab ' + (active ? 'dsh-xl-sheettab-active' : ''),
                        title: 'Double-click to rename',
                        onClick: function () { if (!active) openWorkbook(sheet.file_path, name); },
                        onDoubleClick: function () {
                          var nm = window.prompt('Rename worksheet "' + name + '" to:', name);
                          if (nm && nm.trim() && nm.trim() !== name) {
                            runSheetOp('rename_sheet', { sheet_name: name, new_name: nm.trim() });
                          }
                        },
                      },
                      name
                    );
                  }),
                  h('button', {
                    type: 'button',
                    className: 'dsh-xl-addsheet',
                    title: 'New sheet',
                    onClick: function () {
                      var nm = window.prompt('New worksheet name:', 'Sheet' + ((sheet.sheets || []).length + 1));
                      if (nm && nm.trim()) runSheetOp('create_sheet', { new_name: nm.trim() });
                    },
                  }, '+'),
                  h('button', {
                    type: 'button',
                    className: 'dsh-xl-rbtn',
                    style: { marginLeft: 'auto' },
                    onClick: function () { runSheetOp('duplicate_sheet'); },
                  }, 'Duplicate Sheet'),
                  (sheet.sheets || []).length > 1
                    ? h('button', {
                        type: 'button',
                        className: 'dsh-xl-rbtn',
                        style: { color: '#c00000' },
                        onClick: function () {
                          if (window.confirm('Delete worksheet "' + sheet.active_sheet + '"?')) runSheetOp('delete_sheet');
                        },
                      }, 'Delete Sheet')
                    : null
                )
              : null,

            // 7. Authentic Excel Green Status Bar (with live Selection Average/Count/Sum & Zoom)
            h(
              'div',
              { className: 'dsh-xl-statusbar' },
              h('span', null, editing ? 'Edit' : (sheet ? 'Ready · ' + sheet.file_path : 'Ready')),
              h(
                'div',
                { style: { display: 'inline-flex', alignItems: 'center', gap: '14px' } },
                selStats
                  ? h(
                      'span',
                      null,
                      (selStats.avg !== null ? 'Average: ' + selStats.avg + '   ' : '') +
                        'Count: ' + selStats.count +
                        (selStats.avg !== null ? '   Sum: ' + selStats.sum : '')
                    )
                  : null,
                h(
                  'span',
                  { style: { display: 'inline-flex', alignItems: 'center', gap: '4px' } },
                  h('button', { type: 'button', className: 'dsh-xl-qat-btn', style: { height: '18px' }, onClick: function () { setZoom(function (z) { return Math.max(50, z - 10); }); } }, '−'),
                  h('span', null, zoom + '%'),
                  h('button', { type: 'button', className: 'dsh-xl-qat-btn', style: { height: '18px' }, onClick: function () { setZoom(function (z) { return Math.min(200, z + 10); }); } }, '+')
                ),
                h('span', null, 'v0.4.0')
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
