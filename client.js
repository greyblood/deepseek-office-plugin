(function () {
  'use strict';

  if (typeof window === 'undefined' || !window.__ModuleLoader__ || typeof window.__ModuleLoader__.load !== 'function') {
    return;
  }

  var PLUGIN_ID = 'deepseek-office-plugin';
  var PANEL_ID = 'excel-studio';
  var STYLE_ID = 'dsh-excel-studio-styles-v5';

  var CSS_TEXT = [
    '.dsh-xl-root {',
    '  display: flex; flex-direction: column; height: 100%; width: 100%;',
    '  min-height: 0; min-width: 0; overflow: hidden;',
    '  font-family: "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, sans-serif;',
    '  font-size: 12px; color: var(--dsw-alias-text-primary, #1f2328);',
    '  background: var(--dsw-alias-bg-primary, #ffffff); position: relative; user-select: none;',
    '}',
    '.dsh-xl-titlebar {',
    '  display: flex; align-items: center; justify-content: space-between; height: 34px;',
    '  padding: 0 12px; background: #107c41; color: #ffffff; gap: 10px; flex-shrink: 0;',
    '}',
    '.dsh-xl-qat { display: inline-flex; align-items: center; gap: 2px; }',
    '.dsh-xl-qat-btn {',
    '  height: 24px; min-width: 26px; padding: 0 6px; border: none; border-radius: 3px;',
    '  background: transparent; color: #ffffff; font-size: 12px; cursor: pointer;',
    '  display: inline-flex; align-items: center; justify-content: center; gap: 4px;',
    '}',
    '.dsh-xl-qat-btn:hover:not(:disabled) { background: rgba(255,255,255,0.18); }',
    '.dsh-xl-qat-btn:disabled { opacity: 0.4; cursor: not-allowed; }',
    '.dsh-xl-title-center { display: inline-flex; align-items: center; gap: 8px; font-weight: 600; font-size: 12.5px; }',
    '.dsh-xl-wb-select {',
    '  height: 24px; max-width: 260px; padding: 0 8px; border-radius: 4px;',
    '  border: 1px solid rgba(255,255,255,0.32); background: rgba(0,0,0,0.16);',
    '  color: #ffffff; font-size: 12px; font-weight: 600; outline: none;',
    '}',
    '.dsh-xl-wb-select option { color: #1f2328; background: #ffffff; }',
    '.dsh-xl-pill { font-size: 11px; padding: 1px 7px; border-radius: 999px; background: rgba(255,255,255,0.2); font-weight: 500; }',
    '.dsh-xl-tabs-bar {',
    '  display: flex; align-items: flex-end; gap: 2px; padding: 2px 10px 0;',
    '  background: var(--dsw-alias-bg-secondary, #f3f3f3);',
    '  border-bottom: 1px solid var(--dsw-alias-border-subtle, #d9d9d9); flex-shrink: 0; overflow-x: auto;',
    '}',
    '.dsh-xl-menu-tab {',
    '  height: 27px; padding: 0 12px; border: none; border-bottom: 2px solid transparent;',
    '  background: transparent; color: inherit; font-size: 12.5px; cursor: pointer;',
    '  border-radius: 4px 4px 0 0; white-space: nowrap;',
    '}',
    '.dsh-xl-menu-tab:hover { background: rgba(127,127,127,0.1); }',
    '.dsh-xl-menu-tab-active { font-weight: 600; color: #107c41; border-bottom-color: #107c41; background: var(--dsw-alias-bg-primary, #ffffff); }',
    '.dsh-xl-file-tab { background: #107c41; color: #ffffff !important; font-weight: 600; margin-right: 4px; border-bottom-color: #107c41; }',
    '.dsh-xl-file-tab:hover { background: #0d6635 !important; }',
    '.dsh-xl-ribbon {',
    '  display: flex; align-items: stretch; padding: 5px 10px 3px;',
    '  background: var(--dsw-alias-bg-primary, #ffffff);',
    '  border-bottom: 1px solid var(--dsw-alias-border-subtle, #d9d9d9);',
    '  overflow-x: auto; flex-shrink: 0; min-height: 76px;',
    '}',
    '.dsh-xl-rgroup {',
    '  display: flex; flex-direction: column; justify-content: space-between; padding: 0 9px;',
    '  border-right: 1px solid var(--dsw-alias-border-subtle, rgba(127,127,127,0.22)); flex-shrink: 0;',
    '}',
    '.dsh-xl-rgroup:first-child { padding-left: 2px; }',
    '.dsh-xl-rgroup:last-child { border-right: none; }',
    '.dsh-xl-rgroup-body { display: flex; flex-direction: column; gap: 4px; justify-content: center; flex: 1; }',
    '.dsh-xl-rgroup-row { display: flex; align-items: center; gap: 3px; }',
    '.dsh-xl-rgroup-label { font-size: 10px; text-align: center; opacity: 0.62; margin-top: 3px; }',
    '.dsh-xl-rbtn {',
    '  height: 24px; min-width: 24px; padding: 0 7px; border: 1px solid transparent; border-radius: 3px;',
    '  background: transparent; color: inherit; font-size: 12px; cursor: pointer;',
    '  display: inline-flex; align-items: center; justify-content: center; gap: 4px; white-space: nowrap;',
    '}',
    '.dsh-xl-rbtn:hover:not(:disabled) { background: var(--dsw-alias-bg-hover, rgba(127,127,127,0.12)); border-color: rgba(127,127,127,0.22); }',
    '.dsh-xl-rbtn:disabled { opacity: 0.4; cursor: not-allowed; }',
    '.dsh-xl-rbtn-active { background: rgba(16,124,65,0.15) !important; border-color: rgba(16,124,65,0.5) !important; color: #107c41; }',
    '.dsh-xl-rbtn-tall { height: 48px; flex-direction: column; padding: 4px 8px; font-size: 11px; gap: 2px; }',
    '.dsh-xl-rselect, .dsh-xl-rinput {',
    '  height: 24px; padding: 0 6px; border-radius: 3px;',
    '  border: 1px solid var(--dsw-alias-border-default, rgba(127,127,127,0.32));',
    '  background: var(--dsw-alias-bg-primary, #ffffff); color: inherit; font-size: 12px; outline: none;',
    '}',
    '.dsh-xl-rselect:focus, .dsh-xl-rinput:focus { border-color: #107c41; }',
    '.dsh-xl-colorpick { width: 22px; height: 22px; padding: 0; border: 1px solid rgba(127,127,127,0.4); border-radius: 3px; cursor: pointer; background: transparent; }',
    '.dsh-xl-file-menu {',
    '  position: absolute; top: 63px; left: 10px; width: 320px; z-index: 40;',
    '  background: var(--dsw-alias-bg-primary, #ffffff); border: 1px solid var(--dsw-alias-border-default, #c8c8c8);',
    '  border-radius: 6px; box-shadow: 0 10px 28px rgba(0,0,0,0.18); padding: 8px; display: flex; flex-direction: column; gap: 4px;',
    '}',
    '.dsh-xl-menu-item {',
    '  display: flex; align-items: center; justify-content: space-between; padding: 7px 10px;',
    '  border: none; border-radius: 4px; background: transparent; color: inherit; font-size: 12.5px; cursor: pointer; text-align: left; width: 100%;',
    '}',
    '.dsh-xl-menu-item:hover:not(:disabled) { background: rgba(16,124,65,0.1); }',
    '.dsh-xl-menu-item:disabled { opacity: 0.45; cursor: not-allowed; }',
    '.dsh-xl-menu-sep { height: 1px; background: rgba(127,127,127,0.2); margin: 3px 0; }',
    '.dsh-xl-fbar {',
    '  display: flex; align-items: center; gap: 6px; padding: 4px 10px;',
    '  background: var(--dsw-alias-bg-secondary, #f8f9fa); border-bottom: 1px solid var(--dsw-alias-border-subtle, #d9d9d9); flex-shrink: 0;',
    '}',
    '.dsh-xl-namebox {',
    '  width: 78px; height: 24px; text-align: center; font-family: ui-monospace, SFMono-Regular, Menlo, monospace;',
    '  font-weight: 600; font-size: 12px; border: 1px solid var(--dsw-alias-border-default, #c8c8c8);',
    '  border-radius: 3px; background: var(--dsw-alias-bg-primary, #ffffff); color: inherit; outline: none;',
    '}',
    '.dsh-xl-finput {',
    '  flex: 1; height: 24px; padding: 0 8px; border: 1px solid var(--dsw-alias-border-default, #c8c8c8);',
    '  border-radius: 3px; background: var(--dsw-alias-bg-primary, #ffffff); color: inherit;',
    '  font-family: ui-monospace, SFMono-Regular, Menlo, monospace; font-size: 12.5px; outline: none; user-select: text;',
    '}',
    '.dsh-xl-finput:focus { border-color: #107c41; }',
    '.dsh-xl-grid-wrap { flex: 1; min-height: 0; overflow: auto; position: relative; background: var(--dsw-alias-bg-primary, #ffffff); }',
    '.dsh-xl-table { border-collapse: separate; border-spacing: 0; table-layout: fixed; width: max-content; min-width: 100%; }',
    '.dsh-xl-corner {',
    '  position: sticky; top: 0; left: 0; z-index: 6; width: 44px; min-width: 44px; height: 24px;',
    '  background: var(--dsw-alias-bg-secondary, #f0f0f0); border-right: 1px solid #c6c6c6; border-bottom: 1px solid #c6c6c6;',
    '}',
    '.dsh-xl-colhdr {',
    '  position: sticky; top: 0; z-index: 4; height: 24px; background: var(--dsw-alias-bg-secondary, #f0f0f0);',
    '  color: var(--dsw-alias-text-secondary, #444444); font-weight: 600; font-size: 11px; text-align: center;',
    '  border-right: 1px solid #d0d0d0; border-bottom: 1px solid #c6c6c6;',
    '}',
    '.dsh-xl-colhdr-active { background: rgba(16,124,65,0.16); color: #107c41; border-bottom: 2px solid #107c41; }',
    '.dsh-xl-resizer { position: absolute; top: 0; right: 0; width: 5px; height: 100%; cursor: col-resize; }',
    '.dsh-xl-rowhdr {',
    '  position: sticky; left: 0; z-index: 3; width: 44px; min-width: 44px; text-align: center; font-size: 11px;',
    '  color: var(--dsw-alias-text-secondary, #555555); background: var(--dsw-alias-bg-secondary, #f0f0f0);',
    '  border-right: 1px solid #c6c6c6; border-bottom: 1px solid #d8d8d8;',
    '}',
    '.dsh-xl-rowhdr-active { background: rgba(16,124,65,0.16); color: #107c41; font-weight: 700; border-right: 2px solid #107c41; }',
    '.dsh-xl-cell {',
    '  height: 24px; padding: 0 5px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;',
    '  cursor: cell; position: relative; vertical-align: middle; box-sizing: border-box;',
    '}',
    '.dsh-xl-gridlines .dsh-xl-cell { border-right: 1px solid var(--dsw-alias-border-subtle, #e2e2e2); border-bottom: 1px solid var(--dsw-alias-border-subtle, #e2e2e2); }',
    '.dsh-xl-cell-inrange { background-color: rgba(16, 124, 65, 0.1) !important; }',
    '.dsh-xl-cell-active { outline: 2px solid #107c41; outline-offset: -2px; z-index: 2; }',
    '.dsh-xl-cell-note::after { content: ""; position: absolute; top: 0; right: 0; border-top: 7px solid #d97706; border-left: 7px solid transparent; }',
    '.dsh-xl-fillhandle { position: absolute; right: -3px; bottom: -3px; width: 7px; height: 7px; background: #107c41; border: 1px solid #ffffff; cursor: crosshair; z-index: 5; }',
    '.dsh-xl-editor-input { width: 100%; height: 100%; border: none; outline: none; padding: 0; margin: 0; background: var(--dsw-alias-bg-primary, #ffffff); color: inherit; font: inherit; user-select: text; }',
    '.dsh-xl-frozen-row td { position: sticky; top: 24px; z-index: 3; background: var(--dsw-alias-bg-primary, #ffffff); box-shadow: 0 2px 0 rgba(16,124,65,0.35); }',
    '.dsh-xl-frozen-row .dsh-xl-rowhdr { z-index: 5; }',
    '.dsh-xl-frozen-col { position: sticky; left: 44px; z-index: 2; background: var(--dsw-alias-bg-primary, #ffffff); box-shadow: 2px 0 0 rgba(16,124,65,0.28); }',
    '.dsh-xl-ctxmenu { position: fixed; z-index: 60; min-width: 190px; background: var(--dsw-alias-bg-primary, #ffffff); border: 1px solid #c8c8c8; border-radius: 6px; box-shadow: 0 8px 24px rgba(0,0,0,0.2); padding: 4px; }',
    '.dsh-xl-modal-backdrop { position: fixed; inset: 0; background: rgba(0,0,0,0.35); z-index: 70; display: flex; align-items: center; justify-content: center; }',
    '.dsh-xl-modal { width: 520px; max-width: 92vw; max-height: 82vh; background: var(--dsw-alias-bg-primary, #ffffff); border: 1px solid #c8c8c8; border-radius: 8px; box-shadow: 0 14px 36px rgba(0,0,0,0.28); display: flex; flex-direction: column; overflow: hidden; }',
    '.dsh-xl-modal-hdr { display: flex; align-items: center; justify-content: space-between; padding: 10px 14px; background: #107c41; color: #ffffff; font-weight: 600; }',
    '.dsh-xl-modal-body { padding: 12px 14px; display: flex; flex-direction: column; gap: 10px; overflow: auto; }',
    '.dsh-xl-fnlist { border: 1px solid rgba(127,127,127,0.28); border-radius: 4px; height: 220px; overflow: auto; }',
    '.dsh-xl-fnrow { padding: 6px 10px; cursor: pointer; display: flex; justify-content: space-between; border-bottom: 1px solid rgba(127,127,127,0.12); }',
    '.dsh-xl-fnrow:hover { background: rgba(16,124,65,0.08); }',
    '.dsh-xl-fnrow-sel { background: rgba(16,124,65,0.18) !important; font-weight: 600; color: #107c41; }',
    '.dsh-xl-sheetbar { display: flex; align-items: center; gap: 2px; padding: 0 10px; height: 30px; background: var(--dsw-alias-bg-secondary, #f3f3f3); border-top: 1px solid #d9d9d9; overflow-x: auto; flex-shrink: 0; }',
    '.dsh-xl-sheettab { height: 26px; padding: 0 14px; border: none; border-bottom: 2px solid transparent; background: transparent; color: inherit; font-size: 12px; cursor: pointer; white-space: nowrap; }',
    '.dsh-xl-sheettab-active { background: var(--dsw-alias-bg-primary, #ffffff); color: #107c41; font-weight: 600; border-bottom-color: #107c41; }',
    '.dsh-xl-addsheet { width: 24px; height: 24px; border-radius: 50%; border: 1px solid rgba(127,127,127,0.35); background: transparent; color: inherit; font-weight: 700; cursor: pointer; margin-left: 4px; }',
    '.dsh-xl-statusbar { display: flex; align-items: center; justify-content: space-between; height: 24px; padding: 0 12px; background: #107c41; color: #ffffff; font-size: 11px; gap: 12px; flex-shrink: 0; }',
    '.dsh-xl-banner { padding: 6px 12px; font-size: 12px; display: flex; align-items: center; justify-content: space-between; gap: 8px; flex-shrink: 0; }',
    '.dsh-xl-banner-err { background: #fdecea; color: #b42318; border-bottom: 1px solid #f5c2c7; }',
    '.dsh-xl-banner-info { background: #e8f5ee; color: #107c41; border-bottom: 1px solid #badbcc; }',
    '.dsh-xl-findbar { display: flex; align-items: center; gap: 6px; padding: 5px 12px; background: var(--dsw-alias-bg-secondary, #f3f3f3); border-bottom: 1px solid #d9d9d9; flex-shrink: 0; flex-wrap: wrap; }',
    '.dsh-xl-empty { flex: 1; display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 12px; padding: 32px; text-align: center; }',
    '.dsh-xl-btn-primary { height: 28px; padding: 0 14px; border-radius: 4px; border: none; background: #107c41; color: #ffffff; font-weight: 600; font-size: 12px; cursor: pointer; }',
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
    for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
    return n;
  }

  function makeCoord(row, col) { return colIndexToLetter(col) + String(row); }

  function parseCoord(str) {
    var m = /^\$?([A-Za-z]+)\$?(\d+)$/.exec(String(str || '').trim());
    if (!m) return null;
    return { col: letterToColIndex(m[1]), row: parseInt(m[2], 10) };
  }

  function normalizeRange(a, b) {
    return {
      r1: Math.min(a.row, b.row),
      r2: Math.max(a.row, b.row),
      c1: Math.min(a.col, b.col),
      c2: Math.max(a.col, b.col),
    };
  }

  function formatRangeLabel(a, b) {
    var norm = normalizeRange(a, b);
    if (norm.r1 === norm.r2 && norm.c1 === norm.c2) return makeCoord(norm.r1, norm.c1);
    return makeCoord(norm.r1, norm.c1) + ':' + makeCoord(norm.r2, norm.c2);
  }

  // =========================================================================
  // COMPREHENSIVE EXCEL FORMULA ENGINE (125+ REAL EXCEL FUNCTIONS)
  // =========================================================================
  function isExcelErr(v) { return typeof v === 'string' && v.charAt(0) === '#' && (v.endsWith('!') || v === '#N/A'); }
  function toNum(v) {
    if (typeof v === 'number') return Number.isFinite(v) ? v : 0;
    if (typeof v === 'boolean') return v ? 1 : 0;
    if (v === '' || v === null || v === undefined) return 0;
    var n = Number(String(v).replace(/[$,]/g, ''));
    return Number.isFinite(n) ? n : NaN;
  }
  function flattenArgs(args) {
    var out = [];
    for (var i = 0; i < args.length; i++) {
      var a = args[i];
      if (Array.isArray(a)) {
        for (var r = 0; r < a.length; r++) {
          if (Array.isArray(a[r])) {
            for (var c = 0; c < a[r].length; c++) out.push(a[r][c]);
          } else {
            out.push(a[r]);
          }
        }
      } else {
        out.push(a);
      }
    }
    return out;
  }
  function numericValues(args) {
    var flat = flattenArgs(args);
    var out = [];
    for (var i = 0; i < flat.length; i++) {
      var v = flat[i];
      if (v === '' || v === null || v === undefined || typeof v === 'boolean') continue;
      var n = typeof v === 'number' ? v : Number(String(v).replace(/[$,]/g, ''));
      if (Number.isFinite(n)) out.push(n);
    }
    return out;
  }
  function matchCriteria(val, criteria) {
    var crit = String(criteria);
    var m = /^(>=|<=|<>|>|<|=)(.*)$/.exec(crit);
    if (m) {
      var op = m[1], rhs = m[2];
      var nv = toNum(val), nr = toNum(rhs);
      var bothNum = !isNaN(nv) && !isNaN(nr) && val !== '' && rhs !== '';
      if (op === '>') return bothNum ? nv > nr : String(val) > rhs;
      if (op === '>=') return bothNum ? nv >= nr : String(val) >= rhs;
      if (op === '<') return bothNum ? nv < nr : String(val) < rhs;
      if (op === '<=') return bothNum ? nv <= nr : String(val) <= rhs;
      if (op === '<>') return bothNum ? nv !== nr : String(val).toLowerCase() !== rhs.toLowerCase();
      return bothNum ? nv === nr : String(val).toLowerCase() === rhs.toLowerCase();
    }
    if (crit.indexOf('*') !== -1 || crit.indexOf('?') !== -1) {
      var re = new RegExp('^' + crit.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.') + '$', 'i');
      return re.test(String(val));
    }
    return String(val).toLowerCase() === crit.toLowerCase();
  }
  function gcd2(a, b) { a = Math.abs(Math.trunc(a)); b = Math.abs(Math.trunc(b)); while (b) { var t = b; b = a % b; a = t; } return a; }
  function fact(n) { n = Math.trunc(n); if (n < 0) return '#NUM!'; var r = 1; for (var i = 2; i <= n; i++) r *= i; return r; }
  function parseDateVal(v) {
    if (v instanceof Date) return v;
    var d = new Date(v);
    return isNaN(d.getTime()) ? null : d;
  }

  var EXCEL_FUNCS = {
    // Math & Trig
    SUM: function (a) { return numericValues(a).reduce(function (s, x) { return s + x; }, 0); },
    SUMSQ: function (a) { return numericValues(a).reduce(function (s, x) { return s + x * x; }, 0); },
    PRODUCT: function (a) { var ns = numericValues(a); return ns.length ? ns.reduce(function (p, x) { return p * x; }, 1) : 0; },
    SUMPRODUCT: function (a) {
      if (!a.length) return 0;
      var mats = a.map(function (x) { return flattenArgs([x]).map(function (v) { var n = toNum(v); return isNaN(n) ? 0 : n; }); });
      var len = mats[0].length, sum = 0;
      for (var i = 0; i < len; i++) {
        var prod = 1;
        for (var j = 0; j < mats.length; j++) prod *= (mats[j][i] || 0);
        sum += prod;
      }
      return sum;
    },
    SUMIF: function (a) {
      var r = flattenArgs([a[0]]), crit = a[1], sr = a[2] ? flattenArgs([a[2]]) : r, s = 0;
      for (var i = 0; i < r.length; i++) if (matchCriteria(r[i], crit)) { var n = toNum(sr[i]); if (!isNaN(n)) s += n; }
      return s;
    },
    SUMIFS: function (a) {
      var sr = flattenArgs([a[0]]), s = 0;
      for (var i = 0; i < sr.length; i++) {
        var ok = true;
        for (var p = 1; p + 1 < a.length; p += 2) {
          var cr = flattenArgs([a[p]]);
          if (!matchCriteria(cr[i], a[p + 1])) { ok = false; break; }
        }
        if (ok) { var n = toNum(sr[i]); if (!isNaN(n)) s += n; }
      }
      return s;
    },
    ABS: function (a) { return Math.abs(toNum(a[0])); },
    SIGN: function (a) { var n = toNum(a[0]); return n > 0 ? 1 : n < 0 ? -1 : 0; },
    ROUND: function (a) { var f = Math.pow(10, toNum(a[1] || 0)); return Math.round(toNum(a[0]) * f) / f; },
    ROUNDUP: function (a) { var n = toNum(a[0]), f = Math.pow(10, toNum(a[1] || 0)); return (n >= 0 ? Math.ceil(n * f) : Math.floor(n * f)) / f; },
    ROUNDDOWN: function (a) { var n = toNum(a[0]), f = Math.pow(10, toNum(a[1] || 0)); return (n >= 0 ? Math.floor(n * f) : Math.ceil(n * f)) / f; },
    INT: function (a) { return Math.floor(toNum(a[0])); },
    TRUNC: function (a) { var f = Math.pow(10, toNum(a[1] || 0)); return Math.trunc(toNum(a[0]) * f) / f; },
    MOD: function (a) { var d = toNum(a[1]); return d === 0 ? '#DIV/0!' : ((toNum(a[0]) % d) + d) % d; },
    QUOTIENT: function (a) { var d = toNum(a[1]); return d === 0 ? '#DIV/0!' : Math.trunc(toNum(a[0]) / d); },
    POWER: function (a) { return Math.pow(toNum(a[0]), toNum(a[1])); },
    SQRT: function (a) { var n = toNum(a[0]); return n < 0 ? '#NUM!' : Math.sqrt(n); },
    EXP: function (a) { return Math.exp(toNum(a[0])); },
    LN: function (a) { var n = toNum(a[0]); return n <= 0 ? '#NUM!' : Math.log(n); },
    LOG10: function (a) { var n = toNum(a[0]); return n <= 0 ? '#NUM!' : Math.log10(n); },
    LOG: function (a) { var n = toNum(a[0]), b = a.length > 1 ? toNum(a[1]) : 10; return (n <= 0 || b <= 0 || b === 1) ? '#NUM!' : Math.log(n) / Math.log(b); },
    PI: function () { return Math.PI; },
    SIN: function (a) { return Math.sin(toNum(a[0])); },
    COS: function (a) { return Math.cos(toNum(a[0])); },
    TAN: function (a) { return Math.tan(toNum(a[0])); },
    ASIN: function (a) { return Math.asin(toNum(a[0])); },
    ACOS: function (a) { return Math.acos(toNum(a[0])); },
    ATAN: function (a) { return Math.atan(toNum(a[0])); },
    ATAN2: function (a) { return Math.atan2(toNum(a[1]), toNum(a[0])); },
    DEGREES: function (a) { return toNum(a[0]) * (180 / Math.PI); },
    RADIANS: function (a) { return toNum(a[0]) * (Math.PI / 180); },
    CEILING: function (a) { var s = toNum(a[1] || 1); return s === 0 ? 0 : Math.ceil(toNum(a[0]) / s) * s; },
    FLOOR: function (a) { var s = toNum(a[1] || 1); return s === 0 ? '#DIV/0!' : Math.floor(toNum(a[0]) / s) * s; },
    MROUND: function (a) { var m = toNum(a[1]); return m === 0 ? 0 : Math.round(toNum(a[0]) / m) * m; },
    FACT: function (a) { return fact(toNum(a[0])); },
    GCD: function (a) { return numericValues(a).reduce(gcd2, 0); },
    LCM: function (a) { return numericValues(a).reduce(function (l, x) { return l === 0 || x === 0 ? 0 : Math.abs(l * x) / gcd2(l, x); }, 1); },
    COMBIN: function (a) { var n = Math.trunc(toNum(a[0])), k = Math.trunc(toNum(a[1])); if (k < 0 || k > n) return '#NUM!'; return fact(n) / (fact(k) * fact(n - k)); },
    PERMUT: function (a) { var n = Math.trunc(toNum(a[0])), k = Math.trunc(toNum(a[1])); if (k < 0 || k > n) return '#NUM!'; return fact(n) / fact(n - k); },
    EVEN: function (a) { var n = toNum(a[0]); return n >= 0 ? Math.ceil(n / 2) * 2 : Math.floor(n / 2) * 2; },
    ODD: function (a) { var n = toNum(a[0]); var c = n >= 0 ? Math.ceil(n) : Math.floor(n); return c % 2 === 0 ? c + (n >= 0 ? 1 : -1) : c; },
    RAND: function () { return Math.round(Math.random() * 1e6) / 1e6; },
    RANDBETWEEN: function (a) { var lo = Math.ceil(toNum(a[0])), hi = Math.floor(toNum(a[1])); return Math.floor(Math.random() * (hi - lo + 1)) + lo; },
    SUBTOTAL: function (a) {
      var code = Math.trunc(toNum(a[0])) % 100;
      var rest = a.slice(1);
      if (code === 1) return EXCEL_FUNCS.AVERAGE(rest);
      if (code === 2) return EXCEL_FUNCS.COUNT(rest);
      if (code === 3) return EXCEL_FUNCS.COUNTA(rest);
      if (code === 4) return EXCEL_FUNCS.MAX(rest);
      if (code === 5) return EXCEL_FUNCS.MIN(rest);
      if (code === 6) return EXCEL_FUNCS.PRODUCT(rest);
      if (code === 7) return EXCEL_FUNCS.STDEV(rest);
      if (code === 9) return EXCEL_FUNCS.SUM(rest);
      return '#VALUE!';
    },

    // Statistical
    AVERAGE: function (a) { var ns = numericValues(a); return ns.length ? ns.reduce(function (s, x) { return s + x; }, 0) / ns.length : '#DIV/0!'; },
    AVG: function (a) { return EXCEL_FUNCS.AVERAGE(a); },
    AVERAGEIF: function (a) {
      var r = flattenArgs([a[0]]), crit = a[1], sr = a[2] ? flattenArgs([a[2]]) : r, s = 0, cnt = 0;
      for (var i = 0; i < r.length; i++) if (matchCriteria(r[i], crit)) { var n = toNum(sr[i]); if (!isNaN(n)) { s += n; cnt++; } }
      return cnt ? s / cnt : '#DIV/0!';
    },
    AVERAGEIFS: function (a) {
      var sr = flattenArgs([a[0]]), s = 0, cnt = 0;
      for (var i = 0; i < sr.length; i++) {
        var ok = true;
        for (var p = 1; p + 1 < a.length; p += 2) { if (!matchCriteria(flattenArgs([a[p]])[i], a[p + 1])) { ok = false; break; } }
        if (ok) { var n = toNum(sr[i]); if (!isNaN(n)) { s += n; cnt++; } }
      }
      return cnt ? s / cnt : '#DIV/0!';
    },
    COUNT: function (a) { return numericValues(a).length; },
    COUNTA: function (a) { return flattenArgs(a).filter(function (v) { return v !== '' && v !== null && v !== undefined; }).length; },
    COUNTBLANK: function (a) { return flattenArgs(a).filter(function (v) { return v === '' || v === null || v === undefined; }).length; },
    COUNTIF: function (a) { return flattenArgs([a[0]]).filter(function (v) { return matchCriteria(v, a[1]); }).length; },
    COUNTIFS: function (a) {
      var first = flattenArgs([a[0]]), cnt = 0;
      for (var i = 0; i < first.length; i++) {
        var ok = true;
        for (var p = 0; p + 1 < a.length; p += 2) { if (!matchCriteria(flattenArgs([a[p]])[i], a[p + 1])) { ok = false; break; } }
        if (ok) cnt++;
      }
      return cnt;
    },
    MIN: function (a) { var ns = numericValues(a); return ns.length ? Math.min.apply(null, ns) : 0; },
    MAX: function (a) { var ns = numericValues(a); return ns.length ? Math.max.apply(null, ns) : 0; },
    MEDIAN: function (a) {
      var ns = numericValues(a).sort(function (x, y) { return x - y; });
      if (!ns.length) return '#NUM!';
      var m = Math.floor(ns.length / 2);
      return ns.length % 2 ? ns[m] : (ns[m - 1] + ns[m]) / 2;
    },
    MODE: function (a) {
      var ns = numericValues(a), freq = {}, best = null, bestC = 1;
      for (var i = 0; i < ns.length; i++) { freq[ns[i]] = (freq[ns[i]] || 0) + 1; if (freq[ns[i]] > bestC) { bestC = freq[ns[i]]; best = ns[i]; } }
      return best === null ? '#N/A' : best;
    },
    VAR: function (a) {
      var ns = numericValues(a); if (ns.length < 2) return '#DIV/0!';
      var mean = ns.reduce(function (s, x) { return s + x; }, 0) / ns.length;
      return ns.reduce(function (s, x) { return s + (x - mean) * (x - mean); }, 0) / (ns.length - 1);
    },
    'VAR.S': function (a) { return EXCEL_FUNCS.VAR(a); },
    VARP: function (a) {
      var ns = numericValues(a); if (!ns.length) return '#DIV/0!';
      var mean = ns.reduce(function (s, x) { return s + x; }, 0) / ns.length;
      return ns.reduce(function (s, x) { return s + (x - mean) * (x - mean); }, 0) / ns.length;
    },
    'VAR.P': function (a) { return EXCEL_FUNCS.VARP(a); },
    STDEV: function (a) { var v = EXCEL_FUNCS.VAR(a); return typeof v === 'number' ? Math.sqrt(v) : v; },
    'STDEV.S': function (a) { return EXCEL_FUNCS.STDEV(a); },
    STDEVP: function (a) { var v = EXCEL_FUNCS.VARP(a); return typeof v === 'number' ? Math.sqrt(v) : v; },
    'STDEV.P': function (a) { return EXCEL_FUNCS.STDEVP(a); },
    LARGE: function (a) { var ns = numericValues([a[0]]).sort(function (x, y) { return y - x; }), k = Math.trunc(toNum(a[1])); return (k < 1 || k > ns.length) ? '#NUM!' : ns[k - 1]; },
    SMALL: function (a) { var ns = numericValues([a[0]]).sort(function (x, y) { return x - y; }), k = Math.trunc(toNum(a[1])); return (k < 1 || k > ns.length) ? '#NUM!' : ns[k - 1]; },
    RANK: function (a) {
      var val = toNum(a[0]), ns = numericValues([a[1]]), asc = Boolean(toNum(a[2] || 0));
      ns.sort(function (x, y) { return asc ? x - y : y - x; });
      var idx = ns.indexOf(val);
      return idx === -1 ? '#N/A' : idx + 1;
    },
    PERCENTILE: function (a) {
      var ns = numericValues([a[0]]).sort(function (x, y) { return x - y; }), k = toNum(a[1]);
      if (!ns.length || k < 0 || k > 1) return '#NUM!';
      var idx = k * (ns.length - 1), lo = Math.floor(idx), hi = Math.ceil(idx);
      return ns[lo] + (ns[hi] - ns[lo]) * (idx - lo);
    },
    QUARTILE: function (a) { return EXCEL_FUNCS.PERCENTILE([a[0], toNum(a[1]) / 4]); },
    GEOMEAN: function (a) { var ns = numericValues(a); if (!ns.length || ns.some(function (x) { return x <= 0; })) return '#NUM!'; return Math.pow(ns.reduce(function (p, x) { return p * x; }, 1), 1 / ns.length); },
    HARMEAN: function (a) { var ns = numericValues(a); if (!ns.length || ns.some(function (x) { return x <= 0; })) return '#NUM!'; return ns.length / ns.reduce(function (s, x) { return s + 1 / x; }, 0); },

    // Logical
    IF: function (a) { return a[0] ? (a.length > 1 ? a[1] : true) : (a.length > 2 ? a[2] : false); },
    IFS: function (a) { for (var i = 0; i + 1 < a.length; i += 2) if (a[i]) return a[i + 1]; return '#N/A'; },
    IFERROR: function (a) { return isExcelErr(a[0]) ? a[1] : a[0]; },
    IFNA: function (a) { return a[0] === '#N/A' ? a[1] : a[0]; },
    AND: function (a) { var f = flattenArgs(a); return f.length ? f.every(Boolean) : '#VALUE!'; },
    OR: function (a) { var f = flattenArgs(a); return f.length ? f.some(Boolean) : '#VALUE!'; },
    XOR: function (a) { var c = flattenArgs(a).filter(Boolean).length; return c % 2 === 1; },
    NOT: function (a) { return !a[0]; },
    SWITCH: function (a) {
      var target = a[0];
      for (var i = 1; i + 1 < a.length; i += 2) if (String(target) === String(a[i])) return a[i + 1];
      return a.length % 2 === 0 ? a[a.length - 1] : '#N/A';
    },
    TRUE: function () { return true; },
    FALSE: function () { return false; },

    // Lookup & Reference
    VLOOKUP: function (a) {
      var key = a[0], table = a[1], colIdx = Math.trunc(toNum(a[2])), exact = a.length > 3 ? !a[3] : false;
      if (!Array.isArray(table) || colIdx < 1) return '#VALUE!';
      for (var r = 0; r < table.length; r++) {
        var row = table[r];
        if (!Array.isArray(row)) continue;
        if (exact ? String(row[0]).toLowerCase() === String(key).toLowerCase() : String(row[0]).toLowerCase() === String(key).toLowerCase()) {
          return colIdx <= row.length ? row[colIdx - 1] : '#REF!';
        }
      }
      return '#N/A';
    },
    HLOOKUP: function (a) {
      var key = a[0], table = a[1], rowIdx = Math.trunc(toNum(a[2]));
      if (!Array.isArray(table) || !Array.isArray(table[0]) || rowIdx < 1) return '#VALUE!';
      for (var c = 0; c < table[0].length; c++) {
        if (String(table[0][c]).toLowerCase() === String(key).toLowerCase()) {
          return rowIdx <= table.length ? table[rowIdx - 1][c] : '#REF!';
        }
      }
      return '#N/A';
    },
    XLOOKUP: function (a) {
      var key = a[0], lookup = flattenArgs([a[1]]), ret = flattenArgs([a[2]]), notFound = a.length > 3 ? a[3] : '#N/A';
      for (var i = 0; i < lookup.length; i++) {
        if (String(lookup[i]).toLowerCase() === String(key).toLowerCase()) return ret[i] !== undefined ? ret[i] : '';
      }
      return notFound;
    },
    INDEX: function (a) {
      var arr = a[0], r = Math.trunc(toNum(a[1] || 1)), c = Math.trunc(toNum(a[2] || 1));
      if (!Array.isArray(arr)) return arr;
      if (Array.isArray(arr[0])) return (arr[r - 1] && arr[r - 1][c - 1] !== undefined) ? arr[r - 1][c - 1] : '#REF!';
      return arr[r - 1] !== undefined ? arr[r - 1] : '#REF!';
    },
    MATCH: function (a) {
      var key = a[0], arr = flattenArgs([a[1]]);
      for (var i = 0; i < arr.length; i++) if (String(arr[i]).toLowerCase() === String(key).toLowerCase()) return i + 1;
      return '#N/A';
    },
    CHOOSE: function (a) { var idx = Math.trunc(toNum(a[0])); return (idx >= 1 && idx < a.length) ? a[idx] : '#VALUE!'; },
    ROWS: function (a) { return Array.isArray(a[0]) ? a[0].length : 1; },
    COLUMNS: function (a) { return Array.isArray(a[0]) && Array.isArray(a[0][0]) ? a[0][0].length : 1; },

    // Text
    CONCAT: function (a) { return flattenArgs(a).join(''); },
    CONCATENATE: function (a) { return flattenArgs(a).join(''); },
    TEXTJOIN: function (a) {
      var delim = String(a[0]), ignoreEmpty = Boolean(a[1]), vals = flattenArgs(a.slice(2));
      if (ignoreEmpty) vals = vals.filter(function (v) { return v !== '' && v !== null && v !== undefined; });
      return vals.join(delim);
    },
    LEFT: function (a) { return String(a[0] || '').slice(0, Math.max(0, Math.trunc(toNum(a.length > 1 ? a[1] : 1)))); },
    RIGHT: function (a) { var s = String(a[0] || ''), n = Math.max(0, Math.trunc(toNum(a.length > 1 ? a[1] : 1))); return n === 0 ? '' : s.slice(-n); },
    MID: function (a) { var s = String(a[0] || ''), st = Math.max(1, Math.trunc(toNum(a[1]))), n = Math.max(0, Math.trunc(toNum(a[2]))); return s.substr(st - 1, n); },
    LEN: function (a) { return String(a[0] || '').length; },
    LOWER: function (a) { return String(a[0] || '').toLowerCase(); },
    UPPER: function (a) { return String(a[0] || '').toUpperCase(); },
    PROPER: function (a) { return String(a[0] || '').replace(/\w\S*/g, function (t) { return t.charAt(0).toUpperCase() + t.slice(1).toLowerCase(); }); },
    TRIM: function (a) { return String(a[0] || '').trim().replace(/\s+/g, ' '); },
    CLEAN: function (a) { return String(a[0] || '').replace(/[\x00-\x1F]/g, ''); },
    SUBSTITUTE: function (a) { return String(a[0] || '').split(String(a[1] || '')).join(String(a[2] || '')); },
    REPLACE: function (a) { var s = String(a[0] || ''), st = Math.trunc(toNum(a[1])) - 1, len = Math.trunc(toNum(a[2])); return s.slice(0, st) + String(a[3] || '') + s.slice(st + len); },
    FIND: function (a) { var idx = String(a[1] || '').indexOf(String(a[0] || ''), Math.max(0, Math.trunc(toNum(a[2] || 1)) - 1)); return idx === -1 ? '#VALUE!' : idx + 1; },
    SEARCH: function (a) { var idx = String(a[1] || '').toLowerCase().indexOf(String(a[0] || '').toLowerCase(), Math.max(0, Math.trunc(toNum(a[2] || 1)) - 1)); return idx === -1 ? '#VALUE!' : idx + 1; },
    EXACT: function (a) { return String(a[0] || '') === String(a[1] || ''); },
    REPT: function (a) { return String(a[0] || '').repeat(Math.max(0, Math.trunc(toNum(a[1])))); },
    VALUE: function (a) { var n = toNum(a[0]); return isNaN(n) ? '#VALUE!' : n; },
    NUMBERVALUE: function (a) { return EXCEL_FUNCS.VALUE(a); },
    TEXT: function (a) { return formatDisplayValue(a[0], String(a[1] || 'General')); },
    CHAR: function (a) { return String.fromCharCode(Math.trunc(toNum(a[0]))); },
    CODE: function (a) { var s = String(a[0] || ''); return s.length ? s.charCodeAt(0) : '#VALUE!'; },
    T: function (a) { return typeof a[0] === 'string' ? a[0] : ''; },
    TEXTBEFORE: function (a) { var s = String(a[0] || ''), d = String(a[1] || ''), i = s.indexOf(d); return i === -1 ? '#N/A' : s.slice(0, i); },
    TEXTAFTER: function (a) { var s = String(a[0] || ''), d = String(a[1] || ''), i = s.indexOf(d); return i === -1 ? '#N/A' : s.slice(i + d.length); },

    // Date & Time
    TODAY: function () { return new Date().toISOString().slice(0, 10); },
    NOW: function () { return new Date().toISOString().replace('T', ' ').slice(0, 19); },
    DATE: function (a) { var d = new Date(Date.UTC(toNum(a[0]), toNum(a[1]) - 1, toNum(a[2]))); return d.toISOString().slice(0, 10); },
    TIME: function (a) { return String(Math.trunc(toNum(a[0]))).padStart(2, '0') + ':' + String(Math.trunc(toNum(a[1]))).padStart(2, '0') + ':' + String(Math.trunc(toNum(a[2]))).padStart(2, '0'); },
    YEAR: function (a) { var d = parseDateVal(a[0]); return d ? d.getUTCFullYear() : '#VALUE!'; },
    MONTH: function (a) { var d = parseDateVal(a[0]); return d ? d.getUTCMonth() + 1 : '#VALUE!'; },
    DAY: function (a) { var d = parseDateVal(a[0]); return d ? d.getUTCDate() : '#VALUE!'; },
    HOUR: function (a) { var d = parseDateVal(a[0]); return d ? d.getUTCHours() : '#VALUE!'; },
    MINUTE: function (a) { var d = parseDateVal(a[0]); return d ? d.getUTCMinutes() : '#VALUE!'; },
    SECOND: function (a) { var d = parseDateVal(a[0]); return d ? d.getUTCSeconds() : '#VALUE!'; },
    WEEKDAY: function (a) { var d = parseDateVal(a[0]); return d ? d.getUTCDay() + 1 : '#VALUE!'; },
    DAYS: function (a) { var d1 = parseDateVal(a[0]), d2 = parseDateVal(a[1]); return (d1 && d2) ? Math.round((d1 - d2) / 86400000) : '#VALUE!'; },
    DATEDIF: function (a) {
      var d1 = parseDateVal(a[0]), d2 = parseDateVal(a[1]), u = String(a[2] || 'D').toUpperCase();
      if (!d1 || !d2) return '#VALUE!';
      if (u === 'D') return Math.floor((d2 - d1) / 86400000);
      if (u === 'M') return (d2.getUTCFullYear() - d1.getUTCFullYear()) * 12 + (d2.getUTCMonth() - d1.getUTCMonth());
      if (u === 'Y') return d2.getUTCFullYear() - d1.getUTCFullYear();
      return '#NUM!';
    },
    EDATE: function (a) { var d = parseDateVal(a[0]); if (!d) return '#VALUE!'; d.setUTCMonth(d.getUTCMonth() + Math.trunc(toNum(a[1]))); return d.toISOString().slice(0, 10); },
    EOMONTH: function (a) { var d = parseDateVal(a[0]); if (!d) return '#VALUE!'; var res = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + Math.trunc(toNum(a[1])) + 1, 0)); return res.toISOString().slice(0, 10); },
    NETWORKDAYS: function (a) {
      var d1 = parseDateVal(a[0]), d2 = parseDateVal(a[1]); if (!d1 || !d2) return '#VALUE!';
      var cur = new Date(d1.getTime()), cnt = 0, step = d2 >= d1 ? 1 : -1;
      while ((step > 0 && cur <= d2) || (step < 0 && cur >= d2)) { var w = cur.getUTCDay(); if (w !== 0 && w !== 6) cnt += step; cur.setUTCDate(cur.getUTCDate() + step); }
      return cnt;
    },

    // Financial
    PMT: function (a) {
      var rate = toNum(a[0]), nper = toNum(a[1]), pv = toNum(a[2]), fv = toNum(a[3] || 0), type = toNum(a[4] || 0);
      if (rate === 0) return -(pv + fv) / nper;
      var pvif = Math.pow(1 + rate, nper);
      var pmt = (rate / (pvif - 1)) * -(pv * pvif + fv);
      return type === 1 ? pmt / (1 + rate) : pmt;
    },
    FV: function (a) {
      var rate = toNum(a[0]), nper = toNum(a[1]), pmt = toNum(a[2]), pv = toNum(a[3] || 0), type = toNum(a[4] || 0);
      if (rate === 0) return -(pv + pmt * nper);
      var term = Math.pow(1 + rate, nper);
      return -(pv * term + pmt * (1 + rate * type) * (term - 1) / rate);
    },
    PV: function (a) {
      var rate = toNum(a[0]), nper = toNum(a[1]), pmt = toNum(a[2]), fv = toNum(a[3] || 0), type = toNum(a[4] || 0);
      if (rate === 0) return -(fv + pmt * nper);
      var term = Math.pow(1 + rate, nper);
      return -(fv + pmt * (1 + rate * type) * (term - 1) / rate) / term;
    },
    NPV: function (a) {
      var rate = toNum(a[0]), cfs = numericValues(a.slice(1)), total = 0;
      for (var i = 0; i < cfs.length; i++) total += cfs[i] / Math.pow(1 + rate, i + 1);
      return total;
    },
    IRR: function (a) {
      var cfs = numericValues([a[0]]), guess = a.length > 1 ? toNum(a[1]) : 0.1;
      var r = guess;
      for (var iter = 0; iter < 50; iter++) {
        var npv = 0, dnpv = 0;
        for (var i = 0; i < cfs.length; i++) {
          var denom = Math.pow(1 + r, i);
          npv += cfs[i] / denom;
          if (i > 0) dnpv -= (i * cfs[i]) / (denom * (1 + r));
        }
        if (Math.abs(npv) < 1e-7) return r;
        if (dnpv === 0) break;
        r -= npv / dnpv;
      }
      return isFinite(r) ? r : '#NUM!';
    },
    SLN: function (a) { var life = toNum(a[2]); return life === 0 ? '#DIV/0!' : (toNum(a[0]) - toNum(a[1])) / life; },
    EFFECT: function (a) { var nom = toNum(a[0]), n = Math.trunc(toNum(a[1])); return Math.pow(1 + nom / n, n) - 1; },
    NOMINAL: function (a) { var eff = toNum(a[0]), n = Math.trunc(toNum(a[1])); return n * (Math.pow(1 + eff, 1 / n) - 1); },

    // Information
    ISBLANK: function (a) { return a[0] === '' || a[0] === null || a[0] === undefined; },
    ISNUMBER: function (a) { return typeof a[0] === 'number' || (a[0] !== '' && !isNaN(Number(a[0])) && typeof a[0] !== 'boolean'); },
    ISTEXT: function (a) { return typeof a[0] === 'string' && !isExcelErr(a[0]); },
    ISLOGICAL: function (a) { return typeof a[0] === 'boolean'; },
    ISERROR: function (a) { return isExcelErr(a[0]); },
    ISNA: function (a) { return a[0] === '#N/A'; },
    ISERR: function (a) { return isExcelErr(a[0]) && a[0] !== '#N/A'; },
    ISEVEN: function (a) { return Math.trunc(toNum(a[0])) % 2 === 0; },
    ISODD: function (a) { return Math.abs(Math.trunc(toNum(a[0])) % 2) === 1; },
    N: function (a) { var n = toNum(a[0]); return isNaN(n) ? 0 : n; },
    NA: function () { return '#N/A'; },
    TYPE: function (a) { if (isExcelErr(a[0])) return 16; if (typeof a[0] === 'boolean') return 4; if (typeof a[0] === 'number') return 1; if (Array.isArray(a[0])) return 64; return 2; },
  };

  var FUNCTION_CATALOG = [
    { name: 'SUM', cat: 'Math & Trig', sig: 'SUM(number1, [number2], ...)', desc: 'Adds all numbers in a range of cells.' },
    { name: 'SUMIF', cat: 'Math & Trig', sig: 'SUMIF(range, criteria, [sum_range])', desc: 'Adds cells specified by a given condition.' },
    { name: 'SUMIFS', cat: 'Math & Trig', sig: 'SUMIFS(sum_range, criteria_range1, criteria1, ...)', desc: 'Adds cells specified by multiple conditions.' },
    { name: 'SUMPRODUCT', cat: 'Math & Trig', sig: 'SUMPRODUCT(array1, [array2], ...)', desc: 'Returns the sum of the products of corresponding ranges.' },
    { name: 'PRODUCT', cat: 'Math & Trig', sig: 'PRODUCT(number1, ...)', desc: 'Multiplies all the numbers given as arguments.' },
    { name: 'ABS', cat: 'Math & Trig', sig: 'ABS(number)', desc: 'Returns the absolute value of a number.' },
    { name: 'ROUND', cat: 'Math & Trig', sig: 'ROUND(number, num_digits)', desc: 'Rounds a number to a specified number of digits.' },
    { name: 'ROUNDUP', cat: 'Math & Trig', sig: 'ROUNDUP(number, num_digits)', desc: 'Rounds a number up, away from zero.' },
    { name: 'ROUNDDOWN', cat: 'Math & Trig', sig: 'ROUNDDOWN(number, num_digits)', desc: 'Rounds a number down, toward zero.' },
    { name: 'MOD', cat: 'Math & Trig', sig: 'MOD(number, divisor)', desc: 'Returns the remainder after a number is divided by a divisor.' },
    { name: 'POWER', cat: 'Math & Trig', sig: 'POWER(number, power)', desc: 'Returns the result of a number raised to a power.' },
    { name: 'SQRT', cat: 'Math & Trig', sig: 'SQRT(number)', desc: 'Returns a positive square root.' },
    { name: 'CEILING', cat: 'Math & Trig', sig: 'CEILING(number, significance)', desc: 'Rounds a number up to the nearest multiple of significance.' },
    { name: 'FLOOR', cat: 'Math & Trig', sig: 'FLOOR(number, significance)', desc: 'Rounds a number down to the nearest multiple of significance.' },
    { name: 'RANDBETWEEN', cat: 'Math & Trig', sig: 'RANDBETWEEN(bottom, top)', desc: 'Returns a random integer between the numbers you specify.' },
    { name: 'SUBTOTAL', cat: 'Math & Trig', sig: 'SUBTOTAL(function_num, ref1, ...)', desc: 'Returns a subtotal in a list or database.' },
    { name: 'AVERAGE', cat: 'Statistical', sig: 'AVERAGE(number1, [number2], ...)', desc: 'Returns the arithmetic mean of its arguments.' },
    { name: 'AVERAGEIF', cat: 'Statistical', sig: 'AVERAGEIF(range, criteria, [average_range])', desc: 'Finds average for cells specified by a condition.' },
    { name: 'AVERAGEIFS', cat: 'Statistical', sig: 'AVERAGEIFS(avg_range, criteria_range1, criteria1, ...)', desc: 'Finds average for cells matching multiple criteria.' },
    { name: 'COUNT', cat: 'Statistical', sig: 'COUNT(value1, ...)', desc: 'Counts how many cells contain numbers.' },
    { name: 'COUNTA', cat: 'Statistical', sig: 'COUNTA(value1, ...)', desc: 'Counts how many cells are not empty.' },
    { name: 'COUNTBLANK', cat: 'Statistical', sig: 'COUNTBLANK(range)', desc: 'Counts empty cells in a range.' },
    { name: 'COUNTIF', cat: 'Statistical', sig: 'COUNTIF(range, criteria)', desc: 'Counts cells that meet a single condition.' },
    { name: 'COUNTIFS', cat: 'Statistical', sig: 'COUNTIFS(criteria_range1, criteria1, ...)', desc: 'Counts cells that meet multiple conditions.' },
    { name: 'MIN', cat: 'Statistical', sig: 'MIN(number1, ...)', desc: 'Returns the smallest number in a set of values.' },
    { name: 'MAX', cat: 'Statistical', sig: 'MAX(number1, ...)', desc: 'Returns the largest number in a set of values.' },
    { name: 'MEDIAN', cat: 'Statistical', sig: 'MEDIAN(number1, ...)', desc: 'Returns the median of the given numbers.' },
    { name: 'STDEV.S', cat: 'Statistical', sig: 'STDEV.S(number1, ...)', desc: 'Estimates standard deviation based on a sample.' },
    { name: 'VAR.S', cat: 'Statistical', sig: 'VAR.S(number1, ...)', desc: 'Estimates variance based on a sample.' },
    { name: 'LARGE', cat: 'Statistical', sig: 'LARGE(array, k)', desc: 'Returns the k-th largest value in a data set.' },
    { name: 'SMALL', cat: 'Statistical', sig: 'SMALL(array, k)', desc: 'Returns the k-th smallest value in a data set.' },
    { name: 'RANK', cat: 'Statistical', sig: 'RANK(number, ref, [order])', desc: 'Returns the rank of a number in a list.' },
    { name: 'PERCENTILE', cat: 'Statistical', sig: 'PERCENTILE(array, k)', desc: 'Returns the k-th percentile of values in a range.' },
    { name: 'IF', cat: 'Logical', sig: 'IF(logical_test, value_if_true, [value_if_false])', desc: 'Checks a condition and returns one value if TRUE, another if FALSE.' },
    { name: 'IFS', cat: 'Logical', sig: 'IFS(test1, val1, [test2, val2], ...)', desc: 'Checks multiple conditions and returns the first TRUE match.' },
    { name: 'IFERROR', cat: 'Logical', sig: 'IFERROR(value, value_if_error)', desc: 'Returns value_if_error if expression is an error.' },
    { name: 'AND', cat: 'Logical', sig: 'AND(logical1, [logical2], ...)', desc: 'Returns TRUE if all arguments are TRUE.' },
    { name: 'OR', cat: 'Logical', sig: 'OR(logical1, [logical2], ...)', desc: 'Returns TRUE if any argument is TRUE.' },
    { name: 'NOT', cat: 'Logical', sig: 'NOT(logical)', desc: 'Reverses the logic of its argument.' },
    { name: 'SWITCH', cat: 'Logical', sig: 'SWITCH(expr, val1, res1, ..., [default])', desc: 'Matches expression against a list of values.' },
    { name: 'VLOOKUP', cat: 'Lookup & Reference', sig: 'VLOOKUP(lookup_value, table_array, col_index, [range_lookup])', desc: 'Looks up a value in the leftmost column and returns a value in the same row.' },
    { name: 'HLOOKUP', cat: 'Lookup & Reference', sig: 'HLOOKUP(lookup_value, table_array, row_index, [range_lookup])', desc: 'Searches the top row of a table and returns a value in the same column.' },
    { name: 'XLOOKUP', cat: 'Lookup & Reference', sig: 'XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found])', desc: 'Modern replacement for VLOOKUP/HLOOKUP.' },
    { name: 'INDEX', cat: 'Lookup & Reference', sig: 'INDEX(array, row_num, [column_num])', desc: 'Returns a value at the given row and column in a range.' },
    { name: 'MATCH', cat: 'Lookup & Reference', sig: 'MATCH(lookup_value, lookup_array, [match_type])', desc: 'Returns the relative position of an item in a range.' },
    { name: 'CHOOSE', cat: 'Lookup & Reference', sig: 'CHOOSE(index_num, value1, [value2], ...)', desc: 'Chooses a value from a list based on index.' },
    { name: 'CONCAT', cat: 'Text', sig: 'CONCAT(text1, [text2], ...)', desc: 'Combines text from multiple ranges or strings.' },
    { name: 'TEXTJOIN', cat: 'Text', sig: 'TEXTJOIN(delimiter, ignore_empty, text1, ...)', desc: 'Joins text with a delimiter.' },
    { name: 'LEFT', cat: 'Text', sig: 'LEFT(text, [num_chars])', desc: 'Returns the leftmost characters from a text value.' },
    { name: 'RIGHT', cat: 'Text', sig: 'RIGHT(text, [num_chars])', desc: 'Returns the rightmost characters from a text value.' },
    { name: 'MID', cat: 'Text', sig: 'MID(text, start_num, num_chars)', desc: 'Returns characters from the middle of a text string.' },
    { name: 'LEN', cat: 'Text', sig: 'LEN(text)', desc: 'Returns the number of characters in a text string.' },
    { name: 'UPPER', cat: 'Text', sig: 'UPPER(text)', desc: 'Converts text to uppercase.' },
    { name: 'LOWER', cat: 'Text', sig: 'LOWER(text)', desc: 'Converts text to lowercase.' },
    { name: 'PROPER', cat: 'Text', sig: 'PROPER(text)', desc: 'Capitalizes the first letter in each word.' },
    { name: 'TRIM', cat: 'Text', sig: 'TRIM(text)', desc: 'Removes extra spaces from text.' },
    { name: 'SUBSTITUTE', cat: 'Text', sig: 'SUBSTITUTE(text, old_text, new_text)', desc: 'Substitutes new text for old text in a string.' },
    { name: 'TEXT', cat: 'Text', sig: 'TEXT(value, format_text)', desc: 'Formats a number and converts it to text.' },
    { name: 'TODAY', cat: 'Date & Time', sig: 'TODAY()', desc: 'Returns the current date (YYYY-MM-DD).' },
    { name: 'NOW', cat: 'Date & Time', sig: 'NOW()', desc: 'Returns the current date and time.' },
    { name: 'DATE', cat: 'Date & Time', sig: 'DATE(year, month, day)', desc: 'Returns a date from year, month, and day.' },
    { name: 'YEAR', cat: 'Date & Time', sig: 'YEAR(date)', desc: 'Returns the year of a date.' },
    { name: 'MONTH', cat: 'Date & Time', sig: 'MONTH(date)', desc: 'Returns the month of a date (1-12).' },
    { name: 'DAY', cat: 'Date & Time', sig: 'DAY(date)', desc: 'Returns the day of the month (1-31).' },
    { name: 'DATEDIF', cat: 'Date & Time', sig: 'DATEDIF(start_date, end_date, unit)', desc: 'Calculates days (D), months (M), or years (Y) between two dates.' },
    { name: 'NETWORKDAYS', cat: 'Date & Time', sig: 'NETWORKDAYS(start_date, end_date)', desc: 'Returns the number of whole workdays between two dates.' },
    { name: 'EDATE', cat: 'Date & Time', sig: 'EDATE(start_date, months)', desc: 'Returns the date that is the indicated number of months before or after.' },
    { name: 'EOMONTH', cat: 'Date & Time', sig: 'EOMONTH(start_date, months)', desc: 'Returns the last day of the month.' },
    { name: 'PMT', cat: 'Financial', sig: 'PMT(rate, nper, pv, [fv], [type])', desc: 'Calculates the periodic payment for a loan.' },
    { name: 'FV', cat: 'Financial', sig: 'FV(rate, nper, pmt, [pv], [type])', desc: 'Returns the future value of an investment.' },
    { name: 'PV', cat: 'Financial', sig: 'PV(rate, nper, pmt, [fv], [type])', desc: 'Returns the present value of an investment.' },
    { name: 'NPV', cat: 'Financial', sig: 'NPV(rate, value1, [value2], ...)', desc: 'Calculates the net present value of an investment.' },
    { name: 'IRR', cat: 'Financial', sig: 'IRR(values, [guess])', desc: 'Returns the internal rate of return for a series of cash flows.' },
    { name: 'SLN', cat: 'Financial', sig: 'SLN(cost, salvage, life)', desc: 'Returns the straight-line depreciation of an asset.' },
    { name: 'ISBLANK', cat: 'Information', sig: 'ISBLANK(value)', desc: 'Returns TRUE if the value is blank.' },
    { name: 'ISNUMBER', cat: 'Information', sig: 'ISNUMBER(value)', desc: 'Returns TRUE if the value is a number.' },
    { name: 'ISTEXT', cat: 'Information', sig: 'ISTEXT(value)', desc: 'Returns TRUE if the value is text.' },
    { name: 'ISERROR', cat: 'Information', sig: 'ISERROR(value)', desc: 'Returns TRUE if the value is any error value.' },
  ];

  function tokenizeFormula(input) {
    var tokens = [], i = 0, len = input.length;
    while (i < len) {
      var ch = input[i];
      if (/\s/.test(ch)) { i++; continue; }
      if (ch === '"') {
        var str = ''; i++;
        while (i < len) {
          if (input[i] === '"' && input[i + 1] === '"') { str += '"'; i += 2; }
          else if (input[i] === '"') { i++; break; }
          else { str += input[i++]; }
        }
        tokens.push({ t: 'STR', v: str });
        continue;
      }
      var two = input.slice(i, i + 2);
      if (two === '>=' || two === '<=' || two === '<>') { tokens.push({ t: 'OP', v: two }); i += 2; continue; }
      if ('+-*/^&=<>(),:%'.indexOf(ch) !== -1) {
        if (ch === '(' || ch === ')' || ch === ',' || ch === ':') tokens.push({ t: ch, v: ch });
        else tokens.push({ t: 'OP', v: ch });
        i++; continue;
      }
      if (/[0-9.]/.test(ch)) {
        var num = '';
        while (i < len && /[0-9.eE]/.test(input[i])) {
          num += input[i++];
          if ((num.endsWith('e') || num.endsWith('E')) && (input[i] === '+' || input[i] === '-')) num += input[i++];
        }
        tokens.push({ t: 'NUM', v: parseFloat(num) });
        continue;
      }
      if (/[A-Za-z_$]/.test(ch)) {
        var id = '';
        while (i < len && /[A-Za-z0-9_$.]/.test(input[i])) id += input[i++];
        tokens.push({ t: 'ID', v: id });
        continue;
      }
      i++;
    }
    return tokens;
  }

  function evalExcelFormula(expr, getRawFn, depth) {
    if (depth > 12) return '#REF!';
    var body = String(expr || '').trim();
    if (!body.startsWith('=')) return body;
    var tokens = tokenizeFormula(body.slice(1));
    var pos = 0;

    function peek() { return tokens[pos]; }
    function next() { return tokens[pos++]; }

    function evalCellRef(coordStr) {
      var raw = getRawFn(coordStr.replace(/\$/g, '').toUpperCase());
      if (typeof raw === 'string' && raw.startsWith('=')) return evalExcelFormula(raw, getRawFn, depth + 1);
      if (raw === '' || raw === null || raw === undefined) return '';
      var n = Number(raw);
      return Number.isFinite(n) && String(raw).trim() !== '' ? n : raw;
    }

    function parseComparison() {
      var left = parseConcat();
      while (peek() && peek().t === 'OP' && ['=', '<>', '<', '<=', '>', '>='].indexOf(peek().v) !== -1) {
        var op = next().v;
        var right = parseConcat();
        var nl = toNum(left), nr = toNum(right);
        var bothNum = !isNaN(nl) && !isNaN(nr) && left !== '' && right !== '';
        if (op === '=') left = bothNum ? nl === nr : String(left).toLowerCase() === String(right).toLowerCase();
        else if (op === '<>') left = bothNum ? nl !== nr : String(left).toLowerCase() !== String(right).toLowerCase();
        else if (op === '<') left = bothNum ? nl < nr : String(left) < String(right);
        else if (op === '<=') left = bothNum ? nl <= nr : String(left) <= String(right);
        else if (op === '>') left = bothNum ? nl > nr : String(left) > String(right);
        else if (op === '>=') left = bothNum ? nl >= nr : String(left) >= String(right);
      }
      return left;
    }

    function parseConcat() {
      var left = parseAddSub();
      while (peek() && peek().t === 'OP' && peek().v === '&') {
        next();
        left = String(left === null || left === undefined ? '' : left) + String(parseAddSub() || '');
      }
      return left;
    }

    function parseAddSub() {
      var left = parseMulDiv();
      while (peek() && peek().t === 'OP' && (peek().v === '+' || peek().v === '-')) {
        var op = next().v;
        var right = parseMulDiv();
        if (isExcelErr(left)) return left;
        if (isExcelErr(right)) return right;
        var a = toNum(left), b = toNum(right);
        left = (isNaN(a) || isNaN(b)) ? '#VALUE!' : (op === '+' ? a + b : a - b);
      }
      return left;
    }

    function parseMulDiv() {
      var left = parsePow();
      while (peek() && peek().t === 'OP' && (peek().v === '*' || peek().v === '/')) {
        var op = next().v;
        var right = parsePow();
        if (isExcelErr(left)) return left;
        if (isExcelErr(right)) return right;
        var a = toNum(left), b = toNum(right);
        if (isNaN(a) || isNaN(b)) left = '#VALUE!';
        else if (op === '/' && b === 0) left = '#DIV/0!';
        else left = op === '*' ? a * b : a / b;
      }
      return left;
    }

    function parsePow() {
      var left = parseUnary();
      while (peek() && peek().t === 'OP' && peek().v === '^') {
        next();
        var right = parseUnary();
        left = Math.pow(toNum(left), toNum(right));
      }
      return left;
    }

    function parseUnary() {
      if (peek() && peek().t === 'OP' && (peek().v === '-' || peek().v === '+')) {
        var op = next().v;
        var v = parseUnary();
        return op === '-' ? -toNum(v) : toNum(v);
      }
      var val = parsePrimary();
      if (peek() && peek().t === 'OP' && peek().v === '%') {
        next();
        val = toNum(val) / 100;
      }
      return val;
    }

    function parsePrimary() {
      var tok = peek();
      if (!tok) return '';
      if (tok.t === 'NUM' || tok.t === 'STR') { next(); return tok.v; }
      if (tok.t === '(') {
        next();
        var inner = parseComparison();
        if (peek() && peek().t === ')') next();
        return inner;
      }
      if (tok.t === 'ID') {
        next();
        var name = tok.v.toUpperCase();
        // Function call
        if (peek() && peek().t === '(') {
          next();
          var args = [];
          if (!peek() || peek().t !== ')') {
            while (pos < tokens.length) {
              args.push(parseComparison());
              if (peek() && peek().t === ',') next();
              else break;
            }
          }
          if (peek() && peek().t === ')') next();
          var fn = EXCEL_FUNCS[name];
          if (!fn) return '#NAME?';
          return fn(args);
        }
        // Range A1:B10
        if (peek() && peek().t === ':') {
          next();
          var endTok = next();
          var p1 = parseCoord(name);
          var p2 = endTok ? parseCoord(endTok.v) : null;
          if (!p1 || !p2) return '#REF!';
          var norm = normalizeRange(p1, p2);
          var mat = [];
          for (var r = norm.r1; r <= norm.r2; r++) {
            var rowArr = [];
            for (var c = norm.c1; c <= norm.c2; c++) rowArr.push(evalCellRef(makeCoord(r, c)));
            mat.push(rowArr);
          }
          return mat;
        }
        if (name === 'TRUE') return true;
        if (name === 'FALSE') return false;
        if (parseCoord(name)) return evalCellRef(name);
        return '#NAME?';
      }
      next();
      return '#VALUE!';
    }

    try {
      var res = parseComparison();
      if (typeof res === 'number' && Number.isFinite(res)) return Math.round(res * 1e9) / 1e9;
      if (typeof res === 'boolean') return res ? 'TRUE' : 'FALSE';
      return res;
    } catch (_e) {
      return '#ERROR!';
    }
  }

  function formatDisplayValue(val, numFmt) {
    if (val === '' || val === null || val === undefined) return '';
    if (isExcelErr(val)) return val;
    if (!numFmt || numFmt === 'General' || numFmt === '@') return String(val);
    var n = Number(val);
    if (!Number.isFinite(n)) return String(val);
    var decMatch = /\.([0#]+)/.exec(numFmt);
    var decs = decMatch ? decMatch[1].length : 0;
    if (numFmt.indexOf('%') !== -1) return (n * 100).toFixed(decs) + '%';
    if (numFmt.indexOf('$') !== -1) return '$' + n.toLocaleString(undefined, { minimumFractionDigits: decs, maximumFractionDigits: decs });
    if (numFmt.indexOf('#,##0') !== -1 || numFmt.indexOf('0.0') !== -1) {
      return n.toLocaleString(undefined, { minimumFractionDigits: decs, maximumFractionDigits: decs });
    }
    return String(val);
  }

  async function apiRequest(url, options) {
    var res = await fetch(url, Object.assign({ credentials: 'same-origin' }, options || {}));
    var data = null;
    try { data = await res.json(); } catch (_e) { data = { ok: false, error: 'HTTP ' + res.status }; }
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
            { width: size, height: size, viewBox: '0 0 24 24', fill: 'none', stroke: 'currentColor', strokeWidth: '1.75', strokeLinecap: 'round', strokeLinejoin: 'round', 'aria-hidden': 'true' },
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

          var sWorkbooks = useState([]), workbooks = sWorkbooks[0], setWorkbooks = sWorkbooks[1];
          var sAllowedDirs = useState([]), allowedDirs = sAllowedDirs[0], setAllowedDirs = sAllowedDirs[1];
          var sSheet = useState(null), sheet = sSheet[0], setSheet = sSheet[1];

          var sRibbonTab = useState('Home'), ribbonTab = sRibbonTab[0], setRibbonTab = sRibbonTab[1];
          var sFileMenu = useState(false), fileMenuOpen = sFileMenu[0], setFileMenuOpen = sFileMenu[1];

          var sLoading = useState(false), loading = sLoading[0], setLoading = sLoading[1];
          var sSaving = useState(false), saving = sSaving[0], setSaving = sSaving[1];
          var sError = useState(''), error = sError[0], setError = sError[1];
          var sNotice = useState(''), notice = sNotice[0], setNotice = sNotice[1];

          var sSelStart = useState({ row: 1, col: 1 }), selStart = sSelStart[0], setSelStart = sSelStart[1];
          var sSelEnd = useState({ row: 1, col: 1 }), selEnd = sSelEnd[0], setSelEnd = sSelEnd[1];
          var isDraggingRef = useRef(false);
          var isFillingRef = useRef(false);

          var sEditing = useState(false), editing = sEditing[0], setEditing = sEditing[1];
          var sDraft = useState(''), draftValue = sDraft[0], setDraftValue = sDraft[1];
          var sNameBox = useState('A1'), nameBoxText = sNameBox[0], setNameBoxText = sNameBox[1];

          var sPendingEdits = useState({}), pendingEdits = sPendingEdits[0], setPendingEdits = sPendingEdits[1];
          var sPendingFormats = useState({}), pendingFormats = sPendingFormats[0], setPendingFormats = sPendingFormats[1];
          var sColWidths = useState({}), pendingColWidths = sColWidths[0], setPendingColWidths = sColWidths[1];
          var sNotes = useState({}), pendingNotes = sNotes[0], setPendingNotes = sNotes[1];
          var sMerges = useState([]), pendingMerges = sMerges[0], setPendingMerges = sMerges[1];

          var sUndo = useState([]), undoStack = sUndo[0], setUndoStack = sUndo[1];
          var sRedo = useState([]), redoStack = sRedo[0], setRedoStack = sRedo[1];
          var clipboardRef = useRef(null);
          var formatPainterRef = useRef(null);
          var sPainterOn = useState(false), painterActive = sPainterOn[0], setPainterActive = sPainterOn[1];

          // View & Review options
          var sGridlines = useState(true), showGridlines = sGridlines[0], setShowGridlines = sGridlines[1];
          var sFormulaBar = useState(true), showFormulaBar = sFormulaBar[0], setShowFormulaBar = sFormulaBar[1];
          var sHeadings = useState(true), showHeadings = sHeadings[0], setShowHeadings = sHeadings[1];
          var sFreezeTop = useState(false), freezeTopRow = sFreezeTop[0], setFreezeTopRow = sFreezeTop[1];
          var sFreezeCol = useState(false), freezeFirstCol = sFreezeCol[0], setFreezeFirstCol = sFreezeCol[1];
          var sShowFormulas = useState(false), showFormulas = sShowFormulas[0], setShowFormulas = sShowFormulas[1];
          var sProtected = useState(false), sheetProtected = sProtected[0], setSheetProtected = sProtected[1];
          var sZoom = useState(100), zoom = sZoom[0], setZoom = sZoom[1];

          // Conditional formatting & AutoFilter
          var sCondRule = useState(null), condRule = sCondRule[0], setCondRule = sCondRule[1];
          var sFilterOn = useState(false), autoFilterOn = sFilterOn[0], setAutoFilterOn = sFilterOn[1];
          var sFilterQuery = useState(''), filterQuery = sFilterQuery[0], setFilterQuery = sFilterQuery[1];

          // Modals / Drawers: Insert Function (fx), Chart, Summary, Find & Replace, Context Menu
          var sFxOpen = useState(false), fxModalOpen = sFxOpen[0], setFxModalOpen = sFxOpen[1];
          var sFxCat = useState('All'), fxCategory = sFxCat[0], setFxCategory = sFxCat[1];
          var sFxSearch = useState(''), fxSearch = sFxSearch[0], setFxSearch = sFxSearch[1];
          var sFxSelected = useState(FUNCTION_CATALOG[0]), fxSelected = sFxSelected[0], setFxSelected = sFxSelected[1];
          var sChart = useState(null), chartData = sChart[0], setChartData = sChart[1];
          var sFindOpen = useState(false), findOpen = sFindOpen[0], setFindOpen = sFindOpen[1];
          var sFindText = useState(''), findText = sFindText[0], setFindText = sFindText[1];
          var sReplaceText = useState(''), replaceText = sReplaceText[0], setReplaceText = sReplaceText[1];
          var sCtxMenu = useState(null), ctxMenu = sCtxMenu[0], setCtxMenu = sCtxMenu[1];
          var sSummary = useState(null), summary = sSummary[0], setSummary = sSummary[1];
          var sShowNew = useState(false), showNewForm = sShowNew[0], setShowNewForm = sShowNew[1];
          var sNewName = useState('Book1.xlsx'), newFilename = sNewName[0], setNewFilename = sNewName[1];
          var sOpenPath = useState(''), openPathInput = sOpenPath[0], setOpenPathInput = sOpenPath[1];

          useEffect(function () { ensureStyles(); }, []);
          useEffect(function () {
            var onUp = function () { isDraggingRef.current = false; isFillingRef.current = false; };
            var onClick = function () { setCtxMenu(null); };
            window.addEventListener('mouseup', onUp);
            window.addEventListener('click', onClick);
            return function () { window.removeEventListener('mouseup', onUp); window.removeEventListener('click', onClick); };
          }, []);

          var activeCoord = makeCoord(selStart.row, selStart.col);
          var selNorm = normalizeRange(selStart, selEnd);
          var dirtyCount =
            Object.keys(pendingEdits).length +
            Object.keys(pendingFormats).length +
            Object.keys(pendingColWidths).length +
            Object.keys(pendingNotes).length +
            pendingMerges.length;

          useEffect(function () { setNameBoxText(formatRangeLabel(selStart, selEnd)); }, [selStart, selEnd]);

          var getCellRaw = useCallback(function (coord) {
            if (Object.prototype.hasOwnProperty.call(pendingEdits, coord)) return pendingEdits[coord];
            var cell = sheet && sheet.cells && sheet.cells[coord];
            if (!cell) return '';
            return cell.formula !== undefined ? cell.formula : (cell.raw !== undefined ? String(cell.raw) : '');
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
              next.push({ edits: Object.assign({}, pendingEdits), formats: JSON.parse(JSON.stringify(pendingFormats)) });
              return next;
            });
            setRedoStack([]);
          }, [pendingEdits, pendingFormats]);

          var handleUndo = useCallback(function () {
            if (!undoStack.length) return;
            var prev = undoStack[undoStack.length - 1];
            setRedoStack(function (r) { return r.concat([{ edits: Object.assign({}, pendingEdits), formats: JSON.parse(JSON.stringify(pendingFormats)) }]); });
            setUndoStack(undoStack.slice(0, -1));
            setPendingEdits(prev.edits);
            setPendingFormats(prev.formats);
          }, [undoStack, pendingEdits, pendingFormats]);

          var handleRedo = useCallback(function () {
            if (!redoStack.length) return;
            var next = redoStack[redoStack.length - 1];
            setUndoStack(function (u) { return u.concat([{ edits: Object.assign({}, pendingEdits), formats: JSON.parse(JSON.stringify(pendingFormats)) }]); });
            setRedoStack(redoStack.slice(0, -1));
            setPendingEdits(next.edits);
            setPendingFormats(next.formats);
          }, [redoStack, pendingEdits, pendingFormats]);

          var openWorkbook = useCallback(function (filePath, sheetName) {
            if (!filePath) return;
            setLoading(true); setError(''); setSummary(null); setChartData(null); setFileMenuOpen(false);
            apiRequest('/api/excel-studio/open', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ file_path: filePath, sheet_name: sheetName }),
            })
              .then(function (data) {
                setSheet(data);
                setPendingEdits({}); setPendingFormats({}); setPendingColWidths({}); setPendingNotes({}); setPendingMerges([]);
                setUndoStack([]); setRedoStack([]);
                setSelStart({ row: 1, col: 1 }); setSelEnd({ row: 1, col: 1 }); setEditing(false);
                var first = data.cells && data.cells.A1;
                setDraftValue(first ? (first.formula || String(first.raw || '')) : '');
                setOpenPathInput(data.file_path || filePath);
              })
              .catch(function (err) { setError(err.message || String(err)); })
              .finally(function () { setLoading(false); });
          }, []);

          var refreshWorkbooks = useCallback(function (autoOpenFirst) {
            setLoading(true); setError('');
            apiRequest('/api/excel-studio/workbooks', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({}),
            })
              .then(function (data) {
                var list = (data && data.workbooks) || [];
                setWorkbooks(list);
                setAllowedDirs((data && data.allowed_directories) || []);
                if (autoOpenFirst && list.length > 0) openWorkbook(list[0].path);
              })
              .catch(function (err) { setError(err.message || String(err)); })
              .finally(function () { setLoading(false); });
          }, [openWorkbook]);

          useEffect(function () { refreshWorkbooks(true); }, [refreshWorkbooks]);

          var setCellValuesBatch = useCallback(function (updates) {
            if (sheetProtected) { setError('Sheet is protected (read-only). Unlock in Review tab.'); return; }
            pushUndoSnapshot();
            setPendingEdits(function (cur) { return Object.assign({}, cur, updates); });
          }, [sheetProtected, pushUndoSnapshot]);

          var applyFormatToSelection = useCallback(function (patch) {
            if (!sheet) return;
            if (sheetProtected) { setError('Sheet is protected.'); return; }
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
          }, [sheet, sheetProtected, selNorm, pushUndoSnapshot]);

          var selectCell = function (row, col, startEdit, extend) {
            var maxR = (sheet && sheet.row_count) || 40;
            var maxC = (sheet && sheet.col_count) || 16;
            var r = Math.max(1, Math.min(maxR, row));
            var c = Math.max(1, Math.min(maxC, col));
            if (formatPainterRef.current && !extend) {
              var paintStyle = formatPainterRef.current;
              formatPainterRef.current = null;
              setPainterActive(false);
              var targetCoord = makeCoord(r, c);
              pushUndoSnapshot();
              setPendingFormats(function (cur) {
                var next = Object.assign({}, cur);
                next[targetCoord] = Object.assign({}, paintStyle);
                return next;
              });
            }
            if (extend) setSelEnd({ row: r, col: c });
            else {
              setSelStart({ row: r, col: c });
              setSelEnd({ row: r, col: c });
              setDraftValue(getCellRaw(makeCoord(r, c)));
            }
            setEditing(Boolean(startEdit) && !sheetProtected);
          };

          var handleCopy = useCallback(function (isCut) {
            if (!sheet) return;
            var rows = [], clearMap = {};
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
            var matrix = clipboardRef.current, updates = {};
            for (var r = 0; r < matrix.length; r++) {
              for (var c = 0; c < matrix[r].length; c++) {
                updates[makeCoord(selStart.row + r, selStart.col + c)] = matrix[r][c];
              }
            }
            setCellValuesBatch(updates);
          }, [sheet, selStart, setCellValuesBatch]);

          var handleClearSelection = function (mode) {
            if (!sheet) return;
            if (sheetProtected) { setError('Sheet is protected.'); return; }
            pushUndoSnapshot();
            if (mode === 'contents' || mode === 'all') {
              var updates = {};
              for (var r = selNorm.r1; r <= selNorm.r2; r++) {
                for (var c = selNorm.c1; c <= selNorm.c2; c++) updates[makeCoord(r, c)] = '';
              }
              setPendingEdits(function (cur) { return Object.assign({}, cur, updates); });
              setDraftValue('');
            }
            if (mode === 'formats' || mode === 'all') {
              setPendingFormats(function (cur) {
                var next = Object.assign({}, cur);
                for (var r = selNorm.r1; r <= selNorm.r2; r++) {
                  for (var c = selNorm.c1; c <= selNorm.c2; c++) {
                    next[makeCoord(r, c)] = { bold: false, italic: false, underline: false, strike: false, fill_color: '', font_color: '', align: '', border_preset: 'none', number_format: 'General' };
                  }
                }
                return next;
              });
            }
          };

          var handleFillDirection = function (dir) {
            if (!sheet) return;
            var updates = {};
            if (dir === 'down') {
              for (var c = selNorm.c1; c <= selNorm.c2; c++) {
                var topVal = getCellRaw(makeCoord(selNorm.r1, c));
                for (var r = selNorm.r1 + 1; r <= selNorm.r2; r++) updates[makeCoord(r, c)] = topVal;
              }
            } else {
              for (var r2 = selNorm.r1; r2 <= selNorm.r2; r2++) {
                var leftVal = getCellRaw(makeCoord(r2, selNorm.c1));
                for (var c2 = selNorm.c1 + 1; c2 <= selNorm.c2; c2++) updates[makeCoord(r2, c2)] = leftVal;
              }
            }
            if (Object.keys(updates).length) setCellValuesBatch(updates);
          };

          var handleAdjustDecimals = function (delta) {
            var curFmt = getCellStyle(activeCoord).numberFormat || getCellStyle(activeCoord).number_format || '#,##0.00';
            var isPct = curFmt.indexOf('%') !== -1;
            var isCur = curFmt.indexOf('$') !== -1;
            var m = /\.([0#]+)/.exec(curFmt);
            var decs = Math.max(0, Math.min(6, (m ? m[1].length : 0) + delta));
            var decPart = decs > 0 ? '.' + '0'.repeat(decs) : '';
            var nextFmt = isPct ? ('0' + decPart + '%') : isCur ? ('$#,##0' + decPart) : ('#,##0' + decPart);
            applyFormatToSelection({ number_format: nextFmt });
          };

          var handleInsertFormula = function (fnName) {
            if (!sheet) return;
            var r = selStart.row, c = selStart.col;
            var formula;
            if (['TODAY', 'NOW', 'PI', 'RAND'].indexOf(fnName) !== -1) {
              formula = '=' + fnName + '()';
            } else if (r > 1 && ['SUM', 'AVERAGE', 'COUNT', 'MIN', 'MAX', 'MEDIAN', 'STDEV.S'].indexOf(fnName) !== -1) {
              var startR = r - 1;
              while (startR > 1 && getCellRaw(makeCoord(startR - 1, c)) !== '') {
                if (!Number.isFinite(Number(getCellRaw(makeCoord(startR - 1, c))))) break;
                startR--;
              }
              formula = '=' + fnName + '(' + makeCoord(startR, c) + ':' + makeCoord(r - 1, c) + ')';
            } else {
              formula = '=' + fnName + '(';
            }
            var upd = {}; upd[activeCoord] = formula;
            setCellValuesBatch(upd);
            setDraftValue(formula);
            if (formula.endsWith('(')) setEditing(true);
          };

          var handleBuildChart = function (chartType) {
            if (!sheet) return;
            var labels = [], values = [];
            for (var r = selNorm.r1; r <= selNorm.r2; r++) {
              var lbl = String(getCellDisplay(makeCoord(r, selNorm.c1)) || ('Row ' + r));
              var valCol = selNorm.c2 > selNorm.c1 ? selNorm.c1 + 1 : selNorm.c1;
              var v = toNum(getCellDisplay(makeCoord(r, valCol)));
              if (!isNaN(v)) { labels.push(lbl); values.push(v); }
            }
            if (!values.length) { setError('Select a range with numeric values to generate a chart.'); return; }
            setChartData({ type: chartType, labels: labels, values: values, range: formatRangeLabel(selStart, selEnd) });
          };

          var handleRemoveDuplicates = function () {
            if (!sheet) return;
            var seen = {}, updates = {}, removed = 0;
            for (var r = selNorm.r1; r <= selNorm.r2; r++) {
              var key = '';
              for (var c = selNorm.c1; c <= selNorm.c2; c++) key += '|' + getCellRaw(makeCoord(r, c));
              if (seen[key]) {
                for (var c2 = selNorm.c1; c2 <= selNorm.c2; c2++) updates[makeCoord(r, c2)] = '';
                removed++;
              } else seen[key] = true;
            }
            if (removed > 0) { setCellValuesBatch(updates); setNotice('Cleared ' + removed + ' duplicate row(s).'); }
            else setNotice('No duplicate rows found in selection.');
          };

          var handleTextToColumns = function () {
            if (!sheet) return;
            var delim = window.prompt('Split selected column by delimiter (e.g. , or space):', ',');
            if (delim === null || delim === '') return;
            var updates = {};
            for (var r = selNorm.r1; r <= selNorm.r2; r++) {
              var parts = String(getCellRaw(makeCoord(r, selNorm.c1)) || '').split(delim);
              for (var i = 0; i < parts.length; i++) updates[makeCoord(r, selNorm.c1 + i)] = parts[i].trim();
            }
            setCellValuesBatch(updates);
            setNotice('Split column ' + colIndexToLetter(selNorm.c1) + ' using "' + delim + '".');
          };

          var handleSave = useCallback(function (targetPath) {
            if (!sheet || saving) return;
            var editList = Object.keys(pendingEdits).map(function (coord) { return { cell: coord, value: pendingEdits[coord] }; });
            var formatList = Object.keys(pendingFormats).map(function (coord) { return Object.assign({ cell: coord }, pendingFormats[coord]); });
            if (!targetPath && !editList.length && !formatList.length && !Object.keys(pendingColWidths).length && !Object.keys(pendingNotes).length && !pendingMerges.length) {
              setNotice('All changes are already saved.'); return;
            }
            setSaving(true); setError(''); setNotice('');
            apiRequest('/api/excel-studio/save', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({
                file_path: sheet.file_path, target_path: targetPath || undefined,
                sheet_name: sheet.active_sheet, edits: editList, formats: formatList,
                column_widths: pendingColWidths, notes: pendingNotes, merges: pendingMerges,
              }),
            })
              .then(function (data) {
                setSheet(data); setPendingEdits({}); setPendingFormats({}); setPendingColWidths({}); setPendingNotes({}); setPendingMerges([]);
                setEditing(false);
                if (targetPath) refreshWorkbooks(false);
                setNotice('Saved ' + data.filename + (data.backup_path ? ' (backup created)' : ''));
              })
              .catch(function (err) { setError(err.message || String(err)); })
              .finally(function () { setSaving(false); });
          }, [sheet, saving, pendingEdits, pendingFormats, pendingColWidths, pendingNotes, pendingMerges, refreshWorkbooks]);

          var handleExportCsv = function () {
            if (!sheet) return;
            var lines = [], maxR = Math.max(sheet.max_row || 1, 1), maxC = Math.max(sheet.max_column || 1, 1);
            for (var r = 1; r <= maxR; r++) {
              var cols = [];
              for (var c = 1; c <= maxC; c++) {
                var val = String(getCellDisplay(makeCoord(r, c)) || '');
                if (val.indexOf(',') !== -1 || val.indexOf('"') !== -1 || val.indexOf('\n') !== -1) val = '"' + val.replace(/"/g, '""') + '"';
                cols.push(val);
              }
              lines.push(cols.join(','));
            }
            var blob = new Blob([lines.join('\n')], { type: 'text/csv;charset=utf-8;' });
            var url = URL.createObjectURL(blob);
            var a = document.createElement('a'); a.href = url; a.download = (sheet.active_sheet || 'sheet') + '.csv'; a.click();
            URL.revokeObjectURL(url); setFileMenuOpen(false);
          };

          var runSheetOp = function (op, extra) {
            if (!sheet) return;
            setSaving(true); setError(''); setNotice('');
            apiRequest('/api/excel-studio/sheet-op', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify(Object.assign({ file_path: sheet.file_path, sheet_name: sheet.active_sheet, op: op }, extra || {})),
            })
              .then(function (data) { setSheet(data); setPendingEdits({}); setPendingFormats({}); setNotice('Completed ' + op.replace(/_/g, ' ')); })
              .catch(function (err) { setError(err.message || String(err)); })
              .finally(function () { setSaving(false); });
          };

          var handleCreateWorkbook = function (e) {
            e.preventDefault();
            var nameTrim = newFilename.trim(); if (!nameTrim) return;
            setSaving(true); setError('');
            apiRequest('/api/excel-studio/create', {
              method: 'POST', headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ directory: allowedDirs[0] || undefined, filename: nameTrim }),
            })
              .then(function (data) {
                setSheet(data); setPendingEdits({}); setPendingFormats({}); setShowNewForm(false); setFileMenuOpen(false);
                setNotice('Created ' + data.filename); refreshWorkbooks(false);
              })
              .catch(function (err) { setError(err.message || String(err)); })
              .finally(function () { setSaving(false); });
          };

          var handleSummary = function () {
            if (!sheet) return;
            apiRequest('/api/excel-studio/summary', {
              method: 'POST', headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ file_path: sheet.file_path, sheet_name: sheet.active_sheet }),
            }).then(function (data) { setSummary(data); }).catch(function (err) { setError(err.message || String(err)); });
          };

          var handleFindNext = function () {
            if (!sheet || !findText.trim()) return;
            var needle = findText.toLowerCase(), maxR = sheet.row_count || 40, maxC = sheet.col_count || 16;
            for (var pass = 0; pass < 2; pass++) {
              for (var r = 1; r <= maxR; r++) {
                for (var c = 1; c <= maxC; c++) {
                  if (pass === 0 && (r < selStart.row || (r === selStart.row && c <= selStart.col))) continue;
                  if (String(getCellDisplay(makeCoord(r, c))).toLowerCase().indexOf(needle) !== -1) { selectCell(r, c, false, false); return; }
                }
              }
            }
            setNotice('No matches for "' + findText + '"');
          };

          var handleReplaceAll = function () {
            if (!sheet || !findText) return;
            var updates = {}, count = 0, maxR = sheet.row_count || 40, maxC = sheet.col_count || 16;
            for (var r = 1; r <= maxR; r++) {
              for (var c = 1; c <= maxC; c++) {
                var coord = makeCoord(r, c), raw = String(getCellRaw(coord));
                if (raw.indexOf(findText) !== -1) { updates[coord] = raw.split(findText).join(replaceText); count++; }
              }
            }
            if (count > 0) { setCellValuesBatch(updates); setNotice('Replaced ' + count + ' cell(s).'); }
            else setNotice('No matches found.');
          };

          useEffect(function () {
            var onKeyDown = function (e) {
              var tag = (e.target && e.target.tagName) || '';
              var inInput = tag === 'INPUT' || tag === 'SELECT' || tag === 'TEXTAREA';
              var ctrl = e.ctrlKey || e.metaKey;
              if (ctrl && e.key.toLowerCase() === 's' && sheet) { e.preventDefault(); handleSave(); return; }
              if (ctrl && e.key.toLowerCase() === 'f' && sheet) { e.preventDefault(); setFindOpen(true); return; }
              if (inInput) return;
              if (ctrl && e.key.toLowerCase() === 'z') { e.preventDefault(); handleUndo(); return; }
              if (ctrl && e.key.toLowerCase() === 'y') { e.preventDefault(); handleRedo(); return; }
              if (ctrl && e.key.toLowerCase() === 'c') { e.preventDefault(); handleCopy(false); return; }
              if (ctrl && e.key.toLowerCase() === 'x') { e.preventDefault(); handleCopy(true); return; }
              if (ctrl && e.key.toLowerCase() === 'v') { e.preventDefault(); handlePaste(); return; }
              if (ctrl && e.key.toLowerCase() === 'd') { e.preventDefault(); handleFillDirection('down'); return; }
              if (ctrl && e.key.toLowerCase() === 'r') { e.preventDefault(); handleFillDirection('right'); return; }
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
              else if (e.key === 'F2') { e.preventDefault(); if (!sheetProtected) setEditing(true); }
              else if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); handleClearSelection('contents'); }
              else if (e.key.length === 1 && !ctrl && !e.altKey && !sheetProtected) {
                setEditing(true); setDraftValue(e.key);
                var upd = {}; upd[activeCoord] = e.key; setCellValuesBatch(upd);
              }
            };
            window.addEventListener('keydown', onKeyDown);
            return function () { window.removeEventListener('keydown', onKeyDown); };
          }, [sheet, sheetProtected, selStart, activeCoord, handleSave, handleUndo, handleRedo, handleCopy, handlePaste, applyFormatToSelection, getCellStyle, setCellValuesBatch]);

          var selStats = (function () {
            if (!sheet) return null;
            var count = 0, nums = [];
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
            return { count: count, sum: Math.round(sum * 1e4) / 1e4, avg: nums.length ? Math.round((sum / nums.length) * 1e4) / 1e4 : null };
          })();

          var curStyle = getCellStyle(activeCoord);
          var rowCount = (sheet && sheet.row_count) || 40;
          var colCount = (sheet && sheet.col_count) || 16;
          var colIndices = [], rowIndices = [];
          for (var ci = 1; ci <= colCount; ci++) colIndices.push(ci);
          for (var ri = 1; ri <= rowCount; ri++) rowIndices.push(ri);

          var startColResize = function (letter, startX, currentPx) {
            var onMove = function (ev) {
              var newPx = Math.max(48, Math.min(420, currentPx + (ev.clientX - startX)));
              setPendingColWidths(function (cur) { var n = Object.assign({}, cur); n[letter] = Math.round((newPx / 8) * 100) / 100; return n; });
            };
            var onUp = function () { window.removeEventListener('mousemove', onMove); window.removeEventListener('mouseup', onUp); };
            window.addEventListener('mousemove', onMove); window.addEventListener('mouseup', onUp);
          };

          var handleFillDragEnter = function (targetRow, targetCol) {
            if (!isFillingRef.current || !sheet) return;
            var baseRaw = getCellRaw(activeCoord), baseNum = Number(baseRaw), isNum = baseRaw !== '' && Number.isFinite(baseNum);
            var updates = {};
            if (targetRow > selStart.row) {
              for (var r = selStart.row + 1; r <= targetRow; r++) updates[makeCoord(r, selStart.col)] = isNum ? String(baseNum + (r - selStart.row)) : baseRaw;
              setSelEnd({ row: targetRow, col: selStart.col });
            } else if (targetCol > selStart.col) {
              for (var c = selStart.col + 1; c <= targetCol; c++) updates[makeCoord(selStart.row, c)] = isNum ? String(baseNum + (c - selStart.col)) : baseRaw;
              setSelEnd({ row: selStart.row, col: targetCol });
            }
            if (Object.keys(updates).length) setCellValuesBatch(updates);
          };

          function ribbonGroup(label, content) {
            return h('div', { className: 'dsh-xl-rgroup' }, h('div', { className: 'dsh-xl-rgroup-body' }, content), h('div', { className: 'dsh-xl-rgroup-label' }, label));
          }

          var filteredFunctions = FUNCTION_CATALOG.filter(function (fn) {
            if (fxCategory !== 'All' && fn.cat !== fxCategory) return false;
            if (fxSearch.trim()) {
              var q = fxSearch.toLowerCase();
              return fn.name.toLowerCase().indexOf(q) !== -1 || fn.desc.toLowerCase().indexOf(q) !== -1;
            }
            return true;
          });

          return h(
            'div',
            { className: 'dsh-xl-root', 'data-testid': 'excel-studio-panel' },

            // 1. Excel Green Title & Quick Access Bar
            h(
              'div', { className: 'dsh-xl-titlebar' },
              h(
                'div', { className: 'dsh-xl-qat' },
                h(ExcelStudioIcon, { size: 16 }),
                h('button', { type: 'button', className: 'dsh-xl-qat-btn', title: 'Save (Ctrl+S)', disabled: !sheet || saving, onClick: function () { handleSave(); } }, saving ? 'Saving...' : 'Save'),
                h('button', { type: 'button', className: 'dsh-xl-qat-btn', title: 'Undo (Ctrl+Z)', disabled: !undoStack.length, onClick: handleUndo }, '↶'),
                h('button', { type: 'button', className: 'dsh-xl-qat-btn', title: 'Redo (Ctrl+Y)', disabled: !redoStack.length, onClick: handleRedo }, '↷')
              ),
              h(
                'div', { className: 'dsh-xl-title-center' },
                h(
                  'select',
                  { className: 'dsh-xl-wb-select', value: (sheet && sheet.file_path) || '', onChange: function (e) { if (e.target.value) openWorkbook(e.target.value); }, 'aria-label': 'Select workbook' },
                  h('option', { value: '' }, workbooks.length ? 'Select Workbook...' : 'No Workbooks Found'),
                  workbooks.map(function (wb) { return h('option', { key: wb.path, value: wb.path }, wb.name + (wb.is_locked ? ' (locked)' : '')); })
                ),
                dirtyCount > 0 ? h('span', { className: 'dsh-xl-pill' }, dirtyCount + ' unsaved') : null,
                sheetProtected ? h('span', { className: 'dsh-xl-pill' }, '🔒 Protected') : null
              ),
              h(
                'div', { className: 'dsh-xl-qat' },
                h('button', { type: 'button', className: 'dsh-xl-qat-btn', disabled: !sheet, onClick: function () { setFindOpen(function (v) { return !v; }); } }, 'Find'),
                h('button', { type: 'button', className: 'dsh-xl-qat-btn', onClick: function () { refreshWorkbooks(false); if (sheet) openWorkbook(sheet.file_path, sheet.active_sheet); } }, loading ? '...' : '↻')
              )
            ),

            // 2. Full Excel Menu Tabs Bar (File | Home | Insert | Page Layout | Formulas | Data | Review | View)
            h(
              'div', { className: 'dsh-xl-tabs-bar', role: 'tablist' },
              h('button', { type: 'button', className: 'dsh-xl-menu-tab dsh-xl-file-tab', onClick: function (e) { e.stopPropagation(); setFileMenuOpen(function (v) { return !v; }); } }, 'File'),
              ['Home', 'Insert', 'Page Layout', 'Formulas', 'Data', 'Review', 'View'].map(function (tab) {
                return h('button', {
                  key: tab, type: 'button', role: 'tab', 'aria-selected': ribbonTab === tab,
                  className: 'dsh-xl-menu-tab ' + (ribbonTab === tab ? 'dsh-xl-menu-tab-active' : ''),
                  onClick: function () { setRibbonTab(tab); setFileMenuOpen(false); },
                }, tab);
              })
            ),

            // File Backstage Menu
            fileMenuOpen
              ? h(
                  'div', { className: 'dsh-xl-file-menu', onClick: function (e) { e.stopPropagation(); } },
                  h('button', { type: 'button', className: 'dsh-xl-menu-item', onClick: function () { setShowNewForm(true); setFileMenuOpen(false); } }, h('span', null, 'New Blank Workbook'), h('span', { style: { opacity: 0.55 } }, 'Ctrl+N')),
                  h('button', { type: 'button', className: 'dsh-xl-menu-item', disabled: !sheet, onClick: function () { handleSave(); setFileMenuOpen(false); } }, h('span', null, 'Save'), h('span', { style: { opacity: 0.55 } }, 'Ctrl+S')),
                  h('button', { type: 'button', className: 'dsh-xl-menu-item', disabled: !sheet, onClick: function () { var d = window.prompt('Save As (.xlsx path):', sheet.file_path.replace(/\.xlsx$/i, '_copy.xlsx')); if (d && d.trim()) handleSave(d.trim()); setFileMenuOpen(false); } }, h('span', null, 'Save As...'), h('span', { style: { opacity: 0.55 } }, '.xlsx')),
                  h('button', { type: 'button', className: 'dsh-xl-menu-item', disabled: !sheet, onClick: handleExportCsv }, h('span', null, 'Export Active Sheet as CSV'), h('span', { style: { opacity: 0.55 } }, '.csv')),
                  h('button', { type: 'button', className: 'dsh-xl-menu-item', disabled: !sheet, onClick: function () { setFileMenuOpen(false); window.print(); } }, h('span', null, 'Print / Export PDF'), h('span', { style: { opacity: 0.55 } }, 'Ctrl+P')),
                  h('div', { className: 'dsh-xl-menu-sep' }),
                  h('form', { style: { display: 'flex', gap: '4px', padding: '4px' }, onSubmit: function (e) { e.preventDefault(); if (openPathInput.trim()) openWorkbook(openPathInput.trim()); } },
                    h('input', { className: 'dsh-xl-rinput', style: { flex: 1 }, placeholder: 'Open .xlsx path...', value: openPathInput, onChange: function (e) { setOpenPathInput(e.target.value); } }),
                    h('button', { type: 'submit', className: 'dsh-xl-rbtn dsh-xl-rbtn-active' }, 'Open')
                  )
                )
              : null,

            // 3. Full Excel Grouped Ribbon
            h(
              'div', { className: 'dsh-xl-ribbon' },
              ribbonTab === 'Home'
                ? [
                    ribbonGroup('Clipboard', h('div', { className: 'dsh-xl-rgroup-row' },
                      h('button', { type: 'button', className: 'dsh-xl-rbtn dsh-xl-rbtn-tall', disabled: !sheet, onClick: handlePaste }, h('span', { style: { fontSize: '14px' } }, '📋'), h('span', null, 'Paste')),
                      h('div', { style: { display: 'flex', flexDirection: 'column', gap: '2px' } },
                        h('button', { type: 'button', className: 'dsh-xl-rbtn', disabled: !sheet, onClick: function () { handleCopy(true); } }, '✂ Cut'),
                        h('button', { type: 'button', className: 'dsh-xl-rbtn', disabled: !sheet, onClick: function () { handleCopy(false); } }, '⧉ Copy'),
                        h('button', { type: 'button', className: 'dsh-xl-rbtn ' + (painterActive ? 'dsh-xl-rbtn-active' : ''), disabled: !sheet, onClick: function () { formatPainterRef.current = getCellStyle(activeCoord); setPainterActive(true); setNotice('Format Painter active: click a cell to apply style.'); } }, '🖌 Paint')
                      )
                    )),
                    ribbonGroup('Font', [
                      h('div', { key: 'f1', className: 'dsh-xl-rgroup-row' },
                        h('select', { className: 'dsh-xl-rselect', style: { width: '104px' }, disabled: !sheet, value: curStyle.fontFamily || curStyle.font_family || 'Calibri', onChange: function (e) { applyFormatToSelection({ font_family: e.target.value }); } },
                          ['Calibri', 'Arial', 'Segoe UI', 'Times New Roman', 'Consolas', 'Verdana', 'Georgia'].map(function (f) { return h('option', { key: f, value: f }, f); })
                        ),
                        h('select', { className: 'dsh-xl-rselect', style: { width: '48px' }, disabled: !sheet, value: String(curStyle.fontSize || curStyle.font_size || 11), onChange: function (e) { applyFormatToSelection({ font_size: Number(e.target.value) }); } },
                          [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36].map(function (sz) { return h('option', { key: sz, value: String(sz) }, String(sz)); })
                        ),
                        h('button', { type: 'button', className: 'dsh-xl-rbtn', disabled: !sheet, title: 'Increase Font Size', onClick: function () { applyFormatToSelection({ font_size: Math.min(36, (curStyle.fontSize || curStyle.font_size || 11) + 2) }); } }, 'A▲'),
                        h('button', { type: 'button', className: 'dsh-xl-rbtn', disabled: !sheet, title: 'Decrease Font Size', onClick: function () { applyFormatToSelection({ font_size: Math.max(8, (curStyle.fontSize || curStyle.font_size || 11) - 2) }); } }, 'A▼')
                      ),
                      h('div', { key: 'f2', className: 'dsh-xl-rgroup-row' },
                        h('button', { type: 'button', className: 'dsh-xl-rbtn ' + (curStyle.bold ? 'dsh-xl-rbtn-active' : ''), disabled: !sheet, onClick: function () { applyFormatToSelection({ bold: !curStyle.bold }); } }, h('strong', null, 'B')),
                        h('button', { type: 'button', className: 'dsh-xl-rbtn ' + (curStyle.italic ? 'dsh-xl-rbtn-active' : ''), disabled: !sheet, onClick: function () { applyFormatToSelection({ italic: !curStyle.italic }); } }, h('em', null, 'I')),
                        h('button', { type: 'button', className: 'dsh-xl-rbtn ' + (curStyle.underline ? 'dsh-xl-rbtn-active' : ''), disabled: !sheet, onClick: function () { applyFormatToSelection({ underline: !curStyle.underline }); } }, h('u', null, 'U')),
                        h('button', { type: 'button', className: 'dsh-xl-rbtn ' + (curStyle.strike ? 'dsh-xl-rbtn-active' : ''), disabled: !sheet, onClick: function () { applyFormatToSelection({ strike: !curStyle.strike }); } }, h('s', null, 'S')),
                        h('select', { className: 'dsh-xl-rselect', style: { width: '74px' }, disabled: !sheet, value: '', onChange: function (e) { if (e.target.value) applyFormatToSelection({ border_preset: e.target.value }); } },
                          h('option', { value: '' }, 'Borders'), h('option', { value: 'all' }, 'All Borders'), h('option', { value: 'bottom' }, 'Bottom'), h('option', { value: 'top_bottom' }, 'Top & Double'), h('option', { value: 'none' }, 'No Border')
                        ),
                        h('input', { type: 'color', className: 'dsh-xl-colorpick', title: 'Fill Color', disabled: !sheet, value: curStyle.fill_color || curStyle.fill || '#fff2cc', onChange: function (e) { applyFormatToSelection({ fill_color: e.target.value }); } }),
                        h('input', { type: 'color', className: 'dsh-xl-colorpick', title: 'Font Color', disabled: !sheet, value: curStyle.font_color || curStyle.color || '#000000', onChange: function (e) { applyFormatToSelection({ font_color: e.target.value }); } })
                      ),
                    ]),
                    ribbonGroup('Alignment', [
                      h('div', { key: 'a1', className: 'dsh-xl-rgroup-row' },
                        ['top', 'middle', 'bottom'].map(function (va) { return h('button', { key: va, type: 'button', disabled: !sheet, className: 'dsh-xl-rbtn ' + (curStyle.valign === va ? 'dsh-xl-rbtn-active' : ''), onClick: function () { applyFormatToSelection({ valign: va }); } }, va === 'top' ? '⤒' : va === 'middle' ? '⇳' : '⤓'); }),
                        h('button', { type: 'button', disabled: !sheet, className: 'dsh-xl-rbtn ' + (curStyle.wrap ? 'dsh-xl-rbtn-active' : ''), onClick: function () { applyFormatToSelection({ wrap: !curStyle.wrap }); } }, 'Wrap')
                      ),
                      h('div', { key: 'a2', className: 'dsh-xl-rgroup-row' },
                        ['left', 'center', 'right'].map(function (ha) { return h('button', { key: ha, type: 'button', disabled: !sheet, className: 'dsh-xl-rbtn ' + (curStyle.align === ha ? 'dsh-xl-rbtn-active' : ''), onClick: function () { applyFormatToSelection({ align: ha }); } }, ha === 'left' ? '≡L' : ha === 'center' ? '≡C' : '≡R'); }),
                        h('button', { type: 'button', disabled: !sheet, className: 'dsh-xl-rbtn', title: 'Merge & Center selected range', onClick: function () { var r = formatRangeLabel(selStart, selEnd); if (r.indexOf(':') !== -1) { setPendingMerges(function (m) { return m.concat([r]); }); applyFormatToSelection({ align: 'center' }); setNotice('Queued merge for ' + r + ' (click Save).'); } } }, 'Merge')
                      ),
                    ]),
                    ribbonGroup('Number', [
                      h('div', { key: 'n1', className: 'dsh-xl-rgroup-row' },
                        h('select', { className: 'dsh-xl-rselect', style: { width: '108px' }, disabled: !sheet, value: curStyle.numberFormat || curStyle.number_format || 'General', onChange: function (e) { applyFormatToSelection({ number_format: e.target.value }); } },
                          h('option', { value: 'General' }, 'General'), h('option', { value: '#,##0' }, 'Number'), h('option', { value: '#,##0.00' }, 'Number (0.00)'), h('option', { value: '$#,##0.00' }, 'Currency'), h('option', { value: '0.0%' }, 'Percentage'), h('option', { value: 'yyyy-mm-dd' }, 'Short Date'), h('option', { value: '@' }, 'Text')
                        )
                      ),
                      h('div', { key: 'n2', className: 'dsh-xl-rgroup-row' },
                        h('button', { type: 'button', className: 'dsh-xl-rbtn', disabled: !sheet, onClick: function () { applyFormatToSelection({ number_format: '$#,##0.00' }); } }, '$'),
                        h('button', { type: 'button', className: 'dsh-xl-rbtn', disabled: !sheet, onClick: function () { applyFormatToSelection({ number_format: '0.0%' }); } }, '%'),
                        h('button', { type: 'button', className: 'dsh-xl-rbtn', disabled: !sheet, onClick: function () { applyFormatToSelection({ number_format: '#,##0' }); } }, ','),
                        h('button', { type: 'button', className: 'dsh-xl-rbtn', disabled: !sheet, title: 'Increase Decimal', onClick: function () { handleAdjustDecimals(1); } }, '.0→'),
                        h('button', { type: 'button', className: 'dsh-xl-rbtn', disabled: !sheet, title: 'Decrease Decimal', onClick: function () { handleAdjustDecimals(-1); } }, '←.0')
                      ),
                    ]),
                    ribbonGroup('Styles', [
                      h('div', { key: 's1', className: 'dsh-xl-rgroup-row' },
                        h('select', { className: 'dsh-xl-rselect', style: { width: '112px' }, disabled: !sheet, value: '', onChange: function (e) {
                          var v = e.target.value;
                          if (v === 'gt') { var t = window.prompt('Highlight numeric cells greater than:', '100'); if (t !== null && !isNaN(Number(t))) setCondRule({ type: 'gt', value: Number(t) }); }
                          else if (v === 'lt') { var t2 = window.prompt('Highlight numeric cells less than:', '0'); if (t2 !== null && !isNaN(Number(t2))) setCondRule({ type: 'lt', value: Number(t2) }); }
                          else if (v === 'clear') setCondRule(null);
                        } },
                          h('option', { value: '' }, 'Cond. Format'), h('option', { value: 'gt' }, 'Highlight > Value'), h('option', { value: 'lt' }, 'Highlight < Value'), h('option', { value: 'clear' }, 'Clear Rules')
                        )
                      ),
                      h('div', { key: 's2', className: 'dsh-xl-rgroup-row' },
                        h('button', { type: 'button', className: 'dsh-xl-rbtn', disabled: !sheet, style: { background: '#c6efce', color: '#006100' }, onClick: function () { applyFormatToSelection({ fill_color: '#C6EFCE', font_color: '#006100' }); } }, 'Good'),
                        h('button', { type: 'button', className: 'dsh-xl-rbtn', disabled: !sheet, style: { background: '#ffc7ce', color: '#9c0006' }, onClick: function () { applyFormatToSelection({ fill_color: '#FFC7CE', font_color: '#9C0006' }); } }, 'Bad'),
                        h('button', { type: 'button', className: 'dsh-xl-rbtn', disabled: !sheet, style: { background: '#ffeb9c', color: '#9c6500' }, onClick: function () { applyFormatToSelection({ fill_color: '#FFEB9C', font_color: '#9C6500' }); } }, 'Neutral')
                      ),
                    ]),
                    ribbonGroup('Cells', [
                      h('div', { key: 'c1', className: 'dsh-xl-rgroup-row' },
                        h('button', { type: 'button', className: 'dsh-xl-rbtn', disabled: !sheet || saving, onClick: function () { runSheetOp('insert_rows', { index: selStart.row, amount: 1 }); } }, '+ Row'),
                        h('button', { type: 'button', className: 'dsh-xl-rbtn', disabled: !sheet || saving, onClick: function () { runSheetOp('insert_columns', { index: selStart.col, amount: 1 }); } }, '+ Col')
                      ),
                      h('div', { key: 'c2', className: 'dsh-xl-rgroup-row' },
                        h('button', { type: 'button', className: 'dsh-xl-rbtn', disabled: !sheet || saving, onClick: function () { if (window.confirm('Delete row ' + selStart.row + '?')) runSheetOp('delete_rows', { index: selStart.row, amount: 1 }); } }, '− Row'),
                        h('button', { type: 'button', className: 'dsh-xl-rbtn', disabled: !sheet || saving, onClick: function () { if (window.confirm('Delete column ' + colIndexToLetter(selStart.col) + '?')) runSheetOp('delete_columns', { index: selStart.col, amount: 1 }); } }, '− Col')
                      ),
                    ]),
                    ribbonGroup('Editing', [
                      h('div', { key: 'e1', className: 'dsh-xl-rgroup-row' },
                        h('button', { type: 'button', className: 'dsh-xl-rbtn', disabled: !sheet, onClick: function () { handleInsertFormula('SUM'); } }, 'Σ AutoSum'),
                        h('button', { type: 'button', className: 'dsh-xl-rbtn', disabled: !sheet, onClick: function () { handleFillDirection('down'); } }, '↓ Fill'),
                        h('button', { type: 'button', className: 'dsh-xl-rbtn', disabled: !sheet, onClick: function () { setFindOpen(true); } }, '🔍 Find')
                      ),
                      h('div', { key: 'e2', className: 'dsh-xl-rgroup-row' },
                        h('button', { type: 'button', className: 'dsh-xl-rbtn', disabled: !sheet, onClick: function () { handleClearSelection('contents'); } }, 'Clear Text'),
                        h('button', { type: 'button', className: 'dsh-xl-rbtn', disabled: !sheet, onClick: function () { handleClearSelection('all'); } }, 'Clear All')
                      ),
                    ]),
                  ]
                : ribbonTab === 'Insert'
                ? [
                    ribbonGroup('Tables', h('div', { className: 'dsh-xl-rgroup-row' },
                      h('button', { type: 'button', className: 'dsh-xl-rbtn dsh-xl-rbtn-tall', disabled: !sheet, onClick: function () {
                        pushUndoSnapshot();
                        setPendingFormats(function (cur) {
                          var next = Object.assign({}, cur);
                          for (var r = selNorm.r1; r <= selNorm.r2; r++) {
                            for (var c = selNorm.c1; c <= selNorm.c2; c++) {
                              var co = makeCoord(r, c);
                              if (r === selNorm.r1) next[co] = { bold: true, fill_color: '#1F4E79', font_color: '#FFFFFF', border_preset: 'all' };
                              else next[co] = { fill_color: (r - selNorm.r1) % 2 === 1 ? '#F2F6FC' : '#FFFFFF', border_preset: 'all' };
                            }
                          }
                          return next;
                        });
                      } }, h('span', null, '▦'), h('span', null, 'Format Table')),
                      h('button', { type: 'button', className: 'dsh-xl-rbtn dsh-xl-rbtn-tall', disabled: !sheet, onClick: function () { runSheetOp('insert_rows', { index: selStart.row, amount: 1 }); } }, h('span', null, '➕'), h('span', null, 'Row')),
                      h('button', { type: 'button', className: 'dsh-xl-rbtn dsh-xl-rbtn-tall', disabled: !sheet, onClick: function () { runSheetOp('insert_columns', { index: selStart.col, amount: 1 }); } }, h('span', null, '➕'), h('span', null, 'Column'))
                    )),
                    ribbonGroup('Charts', h('div', { className: 'dsh-xl-rgroup-row' },
                      h('button', { type: 'button', className: 'dsh-xl-rbtn dsh-xl-rbtn-tall', disabled: !sheet, onClick: function () { handleBuildChart('bar'); } }, h('span', null, '📊'), h('span', null, 'Bar Chart')),
                      h('button', { type: 'button', className: 'dsh-xl-rbtn dsh-xl-rbtn-tall', disabled: !sheet, onClick: function () { handleBuildChart('line'); } }, h('span', null, '📈'), h('span', null, 'Line Chart'))
                    )),
                    ribbonGroup('Notes', h('div', { className: 'dsh-xl-rgroup-row' },
                      h('button', { type: 'button', className: 'dsh-xl-rbtn dsh-xl-rbtn-tall', disabled: !sheet, onClick: function () {
                        var t = window.prompt('Cell note for ' + activeCoord + ':', pendingNotes[activeCoord] || '');
                        if (t !== null) setPendingNotes(function (n) { var nx = Object.assign({}, n); nx[activeCoord] = t; return nx; });
                      } }, h('span', null, '💬'), h('span', null, 'Add Note'))
                    )),
                  ]
                : ribbonTab === 'Page Layout'
                ? [
                    ribbonGroup('Page Setup', h('div', { className: 'dsh-xl-rgroup-row' },
                      h('button', { type: 'button', className: 'dsh-xl-rbtn dsh-xl-rbtn-tall', disabled: !sheet, onClick: function () { window.print(); } }, h('span', null, '🖨'), h('span', null, 'Print / PDF')),
                      h('button', { type: 'button', className: 'dsh-xl-rbtn dsh-xl-rbtn-tall', disabled: !sheet, onClick: handleExportCsv }, h('span', null, '📤'), h('span', null, 'Export CSV'))
                    )),
                    ribbonGroup('Sheet Options', h('div', { className: 'dsh-xl-rgroup-row' },
                      h('button', { type: 'button', className: 'dsh-xl-rbtn ' + (showGridlines ? 'dsh-xl-rbtn-active' : ''), onClick: function () { setShowGridlines(function (v) { return !v; }); } }, '☑ Gridlines'),
                      h('button', { type: 'button', className: 'dsh-xl-rbtn ' + (showHeadings ? 'dsh-xl-rbtn-active' : ''), onClick: function () { setShowHeadings(function (v) { return !v; }); } }, '☑ Headings')
                    )),
                  ]
                : ribbonTab === 'Formulas'
                ? [
                    ribbonGroup('Function Library', h('div', { className: 'dsh-xl-rgroup-row' },
                      h('button', { type: 'button', className: 'dsh-xl-rbtn dsh-xl-rbtn-tall dsh-xl-rbtn-active', disabled: !sheet, onClick: function () { setFxCategory('All'); setFxModalOpen(true); } }, h('span', { style: { fontWeight: 700, fontStyle: 'italic', fontFamily: 'Georgia, serif' } }, 'fx'), h('span', null, 'Insert Function')),
                      ['Math & Trig', 'Statistical', 'Logical', 'Lookup & Reference', 'Text', 'Date & Time', 'Financial', 'Information'].map(function (cat) {
                        return h('button', {
                          key: cat, type: 'button', className: 'dsh-xl-rbtn dsh-xl-rbtn-tall', disabled: !sheet,
                          onClick: function () { setFxCategory(cat); setFxSearch(''); setFxModalOpen(true); },
                        }, h('span', { style: { color: '#107c41', fontWeight: 700 } }, 'fx'), h('span', null, cat.split(' ')[0]));
                      })
                    )),
                    ribbonGroup('Formula Auditing', h('div', { className: 'dsh-xl-rgroup-row' },
                      h('button', { type: 'button', className: 'dsh-xl-rbtn dsh-xl-rbtn-tall ' + (showFormulas ? 'dsh-xl-rbtn-active' : ''), disabled: !sheet, onClick: function () { setShowFormulas(function (v) { return !v; }); } }, h('span', null, 'fx'), h('span', null, 'Show Formulas'))
                    )),
                  ]
                : ribbonTab === 'Data'
                ? [
                    ribbonGroup('Sort & Filter', h('div', { className: 'dsh-xl-rgroup-row' },
                      h('button', { type: 'button', className: 'dsh-xl-rbtn dsh-xl-rbtn-tall', disabled: !sheet || saving, onClick: function () { runSheetOp('sort_column', { index: selStart.col, ascending: true }); } }, h('span', null, 'A↓Z'), h('span', null, 'Sort A→Z')),
                      h('button', { type: 'button', className: 'dsh-xl-rbtn dsh-xl-rbtn-tall', disabled: !sheet || saving, onClick: function () { runSheetOp('sort_column', { index: selStart.col, ascending: false }); } }, h('span', null, 'Z↓A'), h('span', null, 'Sort Z→A')),
                      h('button', { type: 'button', className: 'dsh-xl-rbtn dsh-xl-rbtn-tall ' + (autoFilterOn ? 'dsh-xl-rbtn-active' : ''), disabled: !sheet, onClick: function () { setAutoFilterOn(function (v) { return !v; }); setFilterQuery(''); } }, h('span', null, '⧨'), h('span', null, 'Filter'))
                    )),
                    ribbonGroup('Data Tools', h('div', { className: 'dsh-xl-rgroup-row' },
                      h('button', { type: 'button', className: 'dsh-xl-rbtn dsh-xl-rbtn-tall', disabled: !sheet, onClick: handleTextToColumns }, h('span', null, '⫼'), h('span', null, 'Text to Cols')),
                      h('button', { type: 'button', className: 'dsh-xl-rbtn dsh-xl-rbtn-tall', disabled: !sheet, onClick: handleRemoveDuplicates }, h('span', null, '🧹'), h('span', null, 'Remove Dups')),
                      h('button', { type: 'button', className: 'dsh-xl-rbtn dsh-xl-rbtn-tall ' + (summary ? 'dsh-xl-rbtn-active' : ''), disabled: !sheet, onClick: function () { if (summary) setSummary(null); else handleSummary(); } }, h('span', null, '📊'), h('span', null, 'Summary'))
                    )),
                  ]
                : ribbonTab === 'Review'
                ? [
                    ribbonGroup('Notes', h('div', { className: 'dsh-xl-rgroup-row' },
                      h('button', { type: 'button', className: 'dsh-xl-rbtn dsh-xl-rbtn-tall', disabled: !sheet, onClick: function () {
                        var t = window.prompt('Note for ' + activeCoord + ':', pendingNotes[activeCoord] || '');
                        if (t !== null) setPendingNotes(function (n) { var nx = Object.assign({}, n); nx[activeCoord] = t; return nx; });
                      } }, h('span', null, '💬'), h('span', null, 'New Note')),
                      h('button', { type: 'button', className: 'dsh-xl-rbtn dsh-xl-rbtn-tall', disabled: !sheet, onClick: function () { setPendingNotes(function (n) { var nx = Object.assign({}, n); nx[activeCoord] = ''; return nx; }); } }, h('span', null, '🗑'), h('span', null, 'Delete Note'))
                    )),
                    ribbonGroup('Protect', h('div', { className: 'dsh-xl-rgroup-row' },
                      h('button', { type: 'button', className: 'dsh-xl-rbtn dsh-xl-rbtn-tall ' + (sheetProtected ? 'dsh-xl-rbtn-active' : ''), disabled: !sheet, onClick: function () { setSheetProtected(function (v) { return !v; }); } }, h('span', null, sheetProtected ? '🔓' : '🔒'), h('span', null, sheetProtected ? 'Unprotect' : 'Protect Sheet'))
                    )),
                  ]
                : [
                    ribbonGroup('Show', h('div', { className: 'dsh-xl-rgroup-row' },
                      h('button', { type: 'button', className: 'dsh-xl-rbtn ' + (showGridlines ? 'dsh-xl-rbtn-active' : ''), onClick: function () { setShowGridlines(function (v) { return !v; }); } }, '☑ Gridlines'),
                      h('button', { type: 'button', className: 'dsh-xl-rbtn ' + (showFormulaBar ? 'dsh-xl-rbtn-active' : ''), onClick: function () { setShowFormulaBar(function (v) { return !v; }); } }, '☑ Formula Bar'),
                      h('button', { type: 'button', className: 'dsh-xl-rbtn ' + (showHeadings ? 'dsh-xl-rbtn-active' : ''), onClick: function () { setShowHeadings(function (v) { return !v; }); } }, '☑ Headings')
                    )),
                    ribbonGroup('Freeze Panes', h('div', { className: 'dsh-xl-rgroup-row' },
                      h('button', { type: 'button', className: 'dsh-xl-rbtn dsh-xl-rbtn-tall ' + (freezeTopRow ? 'dsh-xl-rbtn-active' : ''), onClick: function () { setFreezeTopRow(function (v) { return !v; }); } }, h('span', null, '❄'), h('span', null, 'Top Row')),
                      h('button', { type: 'button', className: 'dsh-xl-rbtn dsh-xl-rbtn-tall ' + (freezeFirstCol ? 'dsh-xl-rbtn-active' : ''), onClick: function () { setFreezeFirstCol(function (v) { return !v; }); } }, h('span', null, '❄'), h('span', null, 'First Col'))
                    )),
                    ribbonGroup('Zoom', h('div', { className: 'dsh-xl-rgroup-row' },
                      [75, 100, 125, 150].map(function (z) { return h('button', { key: z, type: 'button', className: 'dsh-xl-rbtn ' + (zoom === z ? 'dsh-xl-rbtn-active' : ''), onClick: function () { setZoom(z); } }, z + '%'); })
                    )),
                  ]
            ),

            // New Workbook inline bar
            showNewForm
              ? h('form', { className: 'dsh-xl-findbar', onSubmit: handleCreateWorkbook },
                  h('strong', null, 'New Workbook:'),
                  h('input', { className: 'dsh-xl-rinput', style: { width: '220px' }, value: newFilename, onChange: function (e) { setNewFilename(e.target.value); }, placeholder: 'Book1.xlsx' }),
                  h('span', { style: { opacity: 0.65 } }, allowedDirs[0] ? 'in ' + allowedDirs[0] : ''),
                  h('button', { type: 'submit', className: 'dsh-xl-btn-primary', disabled: saving }, 'Create'),
                  h('button', { type: 'button', className: 'dsh-xl-rbtn', onClick: function () { setShowNewForm(false); } }, 'Cancel')
                )
              : null,

            // AutoFilter bar
            autoFilterOn && sheet
              ? h('div', { className: 'dsh-xl-findbar' },
                  h('strong', null, 'AutoFilter (rows containing):'),
                  h('input', { className: 'dsh-xl-rinput', style: { width: '220px' }, value: filterQuery, onChange: function (e) { setFilterQuery(e.target.value); }, placeholder: 'Type to filter rows...' }),
                  h('button', { type: 'button', className: 'dsh-xl-rbtn', onClick: function () { setFilterQuery(''); } }, 'Clear Filter'),
                  h('button', { type: 'button', className: 'dsh-xl-rbtn', style: { marginLeft: 'auto' }, onClick: function () { setAutoFilterOn(false); setFilterQuery(''); } }, 'Close')
                )
              : null,

            // Find & Replace bar
            findOpen && sheet
              ? h('div', { className: 'dsh-xl-findbar' },
                  h('strong', null, 'Find:'),
                  h('input', { className: 'dsh-xl-rinput', style: { width: '160px' }, value: findText, onChange: function (e) { setFindText(e.target.value); }, onKeyDown: function (e) { if (e.key === 'Enter') handleFindNext(); }, placeholder: 'Search...' }),
                  h('button', { type: 'button', className: 'dsh-xl-rbtn', onClick: handleFindNext }, 'Find Next'),
                  h('strong', { style: { marginLeft: '6px' } }, 'Replace:'),
                  h('input', { className: 'dsh-xl-rinput', style: { width: '160px' }, value: replaceText, onChange: function (e) { setReplaceText(e.target.value); }, placeholder: 'Replace with...' }),
                  h('button', { type: 'button', className: 'dsh-xl-rbtn', onClick: handleReplaceAll }, 'Replace All'),
                  h('button', { type: 'button', className: 'dsh-xl-rbtn', style: { marginLeft: 'auto' }, onClick: function () { setFindOpen(false); } }, '✕')
                )
              : null,

            error ? h('div', { className: 'dsh-xl-banner dsh-xl-banner-err', role: 'alert' }, h('span', null, error), h('button', { type: 'button', className: 'dsh-xl-rbtn', onClick: function () { setError(''); } }, 'Dismiss')) : null,
            notice ? h('div', { className: 'dsh-xl-banner dsh-xl-banner-info' }, h('span', null, notice), h('button', { type: 'button', className: 'dsh-xl-rbtn', onClick: function () { setNotice(''); } }, 'OK')) : null,

            // Inline Chart Preview Drawer
            chartData
              ? (function () {
                  var maxV = Math.max.apply(null, chartData.values.map(function (v) { return Math.abs(v); })) || 1;
                  var w = 480, ht = 150, pad = 28;
                  return h(
                    'div',
                    { style: { padding: '8px 14px', borderBottom: '1px solid rgba(127,127,127,0.22)', background: 'var(--dsw-alias-bg-secondary, #f8f9fa)' } },
                    h('div', { style: { display: 'flex', justifyContent: 'space-between', marginBottom: '6px', fontWeight: 600 } },
                      h('span', null, (chartData.type === 'bar' ? 'Bar Chart' : 'Line Chart') + ' (' + chartData.range + ')'),
                      h('button', { type: 'button', className: 'dsh-xl-rbtn', onClick: function () { setChartData(null); } }, 'Close')
                    ),
                    h(
                      'svg', { width: w, height: ht, style: { background: '#ffffff', border: '1px solid #ddd', borderRadius: '4px' } },
                      chartData.type === 'bar'
                        ? chartData.values.map(function (val, idx) {
                            var bw = Math.max(12, Math.floor((w - pad * 2) / chartData.values.length) - 6);
                            var bh = Math.max(2, Math.round((Math.abs(val) / maxV) * (ht - pad * 2)));
                            var x = pad + idx * (bw + 6);
                            var y = ht - pad - bh;
                            return h('g', { key: idx },
                              h('rect', { x: x, y: y, width: bw, height: bh, fill: '#107c41', rx: 2 }),
                              h('text', { x: x + bw / 2, y: ht - 8, fontSize: '10', textAnchor: 'middle', fill: '#444' }, String(chartData.labels[idx]).slice(0, 8))
                            );
                          })
                        : h('polyline', {
                            fill: 'none', stroke: '#107c41', strokeWidth: '2.5',
                            points: chartData.values.map(function (val, idx) {
                              var x = pad + (chartData.values.length > 1 ? (idx / (chartData.values.length - 1)) * (w - pad * 2) : (w - pad * 2) / 2);
                              var y = ht - pad - (Math.abs(val) / maxV) * (ht - pad * 2);
                              return x + ',' + y;
                            }).join(' '),
                          })
                    )
                  );
                })()
              : null,

            summary
              ? h('div', { style: { padding: '8px 14px', borderBottom: '1px solid rgba(127,127,127,0.22)', maxHeight: '160px', overflow: 'auto', background: 'var(--dsw-alias-bg-secondary, #f8f9fa)' } },
                  h('div', { style: { display: 'flex', justifyContent: 'space-between', fontWeight: 600, marginBottom: '4px' } }, h('span', null, 'Dataset Summary (' + (summary.total_rows || 0) + ' rows)'), h('button', { type: 'button', className: 'dsh-xl-rbtn', onClick: function () { setSummary(null); } }, 'Close')),
                  h('pre', { style: { margin: 0, fontSize: '11px' } }, JSON.stringify(summary.summary || summary, null, 2))
                )
              : null,

            // 4. Name Box & Formula Bar
            sheet && showFormulaBar
              ? h(
                  'div', { className: 'dsh-xl-fbar' },
                  h('input', {
                    className: 'dsh-xl-namebox', value: nameBoxText, title: 'Name Box',
                    onChange: function (e) { setNameBoxText(e.target.value); },
                    onKeyDown: function (e) { if (e.key === 'Enter') { var p = parseCoord(nameBoxText); if (p) selectCell(p.row, p.col, false, false); } },
                  }),
                  h('button', { type: 'button', className: 'dsh-xl-rbtn', title: 'Cancel', onClick: function () { setDraftValue(getCellRaw(activeCoord)); setEditing(false); } }, '✕'),
                  h('button', { type: 'button', className: 'dsh-xl-rbtn', title: 'Enter', onClick: function () { var u = {}; u[activeCoord] = draftValue; setCellValuesBatch(u); setEditing(false); } }, '✓'),
                  h('button', { type: 'button', className: 'dsh-xl-rbtn', title: 'Insert Function (125+ Excel functions)', style: { fontStyle: 'italic', fontFamily: 'Georgia, serif', fontWeight: 700, color: '#107c41' }, onClick: function () { setFxCategory('All'); setFxModalOpen(true); } }, 'fx'),
                  h('input', {
                    className: 'dsh-xl-finput', value: draftValue, readOnly: sheetProtected,
                    placeholder: 'Enter value or any Excel formula (e.g. =IF(SUM(A1:A5)>100, "High", "Low"))',
                    onChange: function (e) { var v = e.target.value; setDraftValue(v); var u = {}; u[activeCoord] = v; setCellValuesBatch(u); },
                    onKeyDown: function (e) {
                      if (e.key === 'Enter') { e.preventDefault(); selectCell(selStart.row + 1, selStart.col, false, false); }
                      else if (e.key === 'Tab') { e.preventDefault(); selectCell(selStart.row, selStart.col + 1, false, false); }
                      else if (e.key === 'Escape') { setDraftValue(getCellRaw(activeCoord)); setEditing(false); }
                    },
                  })
                )
              : null,

            // 5. Excel Interactive Grid
            sheet
              ? h(
                  'div', { className: 'dsh-xl-grid-wrap ' + (showGridlines ? 'dsh-xl-gridlines' : '') },
                  h(
                    'table', { className: 'dsh-xl-table', style: { zoom: zoom / 100 } },
                    showHeadings
                      ? h('thead', null, h('tr', null,
                          h('th', { className: 'dsh-xl-corner' }),
                          colIndices.map(function (c) {
                            var letter = colIndexToLetter(c);
                            var w = pendingColWidths[letter] || (sheet.column_widths && sheet.column_widths[letter]);
                            var px = w ? Math.max(52, Math.min(420, Math.round(w * 8))) : 100;
                            var colActive = c >= selNorm.c1 && c <= selNorm.c2;
                            return h('th', {
                              key: letter, className: 'dsh-xl-colhdr ' + (colActive ? 'dsh-xl-colhdr-active' : ''),
                              style: { width: px + 'px', minWidth: px + 'px' },
                              onClick: function () { setSelStart({ row: 1, col: c }); setSelEnd({ row: rowCount, col: c }); },
                            }, letter, h('span', { className: 'dsh-xl-resizer', onMouseDown: function (e) { e.stopPropagation(); startColResize(letter, e.clientX, px); } }));
                          })
                        ))
                      : null,
                    h(
                      'tbody', null,
                      rowIndices.map(function (r) {
                        if (autoFilterOn && filterQuery.trim() && r > 1) {
                          var q = filterQuery.toLowerCase(), hit = false;
                          for (var fc = 1; fc <= colCount; fc++) {
                            if (String(getCellDisplay(makeCoord(r, fc))).toLowerCase().indexOf(q) !== -1) { hit = true; break; }
                          }
                          if (!hit) return null;
                        }
                        var rowActive = r >= selNorm.r1 && r <= selNorm.r2;
                        return h(
                          'tr', { key: r, className: freezeTopRow && r === 1 ? 'dsh-xl-frozen-row' : '' },
                          showHeadings
                            ? h('td', { className: 'dsh-xl-rowhdr ' + (rowActive ? 'dsh-xl-rowhdr-active' : ''), onClick: function () { setSelStart({ row: r, col: 1 }); setSelEnd({ row: r, col: colCount }); } }, r)
                            : null,
                          colIndices.map(function (c) {
                            var coord = makeCoord(r, c);
                            var isActive = selStart.row === r && selStart.col === c;
                            var inRange = r >= selNorm.r1 && r <= selNorm.r2 && c >= selNorm.c1 && c <= selNorm.c2 && !(selNorm.r1 === selNorm.r2 && selNorm.c1 === selNorm.c2);
                            var isBottomRight = r === selNorm.r2 && c === selNorm.c2;
                            var st = getCellStyle(coord);
                            var disp = getCellDisplay(coord);
                            var condBg = undefined;
                            if (condRule && disp !== '') {
                              var cn = toNum(disp);
                              if (!isNaN(cn)) {
                                if (condRule.type === 'gt' && cn > condRule.value) condBg = '#C6EFCE';
                                if (condRule.type === 'lt' && cn < condRule.value) condBg = '#FFC7CE';
                              }
                            }
                            var noteText = pendingNotes[coord];
                            var bp = st.border_preset;
                            var cellCss = {
                              fontFamily: st.fontFamily || st.font_family || undefined,
                              fontSize: (st.fontSize || st.font_size) ? (st.fontSize || st.font_size) + 'px' : undefined,
                              fontWeight: st.bold ? '700' : undefined,
                              fontStyle: st.italic ? 'italic' : undefined,
                              textDecoration: [st.underline ? 'underline' : '', st.strike ? 'line-through' : ''].filter(Boolean).join(' ') || undefined,
                              color: st.font_color || st.color || undefined,
                              backgroundColor: condBg || st.fill_color || st.fill || undefined,
                              textAlign: st.align || undefined,
                              verticalAlign: st.valign || undefined,
                              whiteSpace: st.wrap ? 'normal' : 'nowrap',
                              border: bp === 'all' ? '1px solid #666' : undefined,
                              borderBottom: bp === 'bottom' ? '1px solid #444' : bp === 'top_bottom' ? '3px double #444' : undefined,
                              borderTop: bp === 'top_bottom' ? '1px solid #444' : undefined,
                            };
                            var cls = 'dsh-xl-cell' +
                              (freezeFirstCol && c === 1 ? ' dsh-xl-frozen-col' : '') +
                              (inRange ? ' dsh-xl-cell-inrange' : '') +
                              (isActive ? ' dsh-xl-cell-active' : '') +
                              (noteText ? ' dsh-xl-cell-note' : '');

                            return h(
                              'td',
                              {
                                key: coord, className: cls, style: cellCss, title: noteText || undefined,
                                onMouseDown: function (e) { if (e.button === 2) return; setFileMenuOpen(false); isDraggingRef.current = true; selectCell(r, c, false, e.shiftKey); },
                                onMouseEnter: function () { if (isDraggingRef.current) setSelEnd({ row: r, col: c }); else if (isFillingRef.current) handleFillDragEnter(r, c); },
                                onDoubleClick: function () { selectCell(r, c, true, false); },
                                onContextMenu: function (e) { e.preventDefault(); selectCell(r, c, false, false); setCtxMenu({ x: e.clientX, y: e.clientY, row: r, col: c }); },
                              },
                              isActive && editing
                                ? h('input', {
                                    className: 'dsh-xl-editor-input', autoFocus: true, value: draftValue,
                                    onChange: function (e) { var v = e.target.value; setDraftValue(v); var u = {}; u[coord] = v; setCellValuesBatch(u); },
                                    onBlur: function () { setEditing(false); },
                                    onKeyDown: function (e) {
                                      if (e.key === 'Enter') { e.preventDefault(); selectCell(r + 1, c, false, false); }
                                      else if (e.key === 'Tab') { e.preventDefault(); selectCell(r, c + 1, false, false); }
                                      else if (e.key === 'Escape') { setDraftValue(getCellRaw(coord)); setEditing(false); }
                                    },
                                  })
                                : disp,
                              isBottomRight && !editing && !sheetProtected
                                ? h('span', { className: 'dsh-xl-fillhandle', title: 'Drag to AutoFill', onMouseDown: function (e) { e.stopPropagation(); isFillingRef.current = true; } })
                                : null
                            );
                          })
                        );
                      })
                    )
                  )
                )
              : h('div', { className: 'dsh-xl-empty' },
                  h(ExcelStudioIcon, { size: 42 }),
                  h('div', { style: { fontWeight: 700, fontSize: '16px' } }, 'DeepSeek Excel Studio'),
                  h('div', { style: { maxWidth: '420px', opacity: 0.78, lineHeight: '1.5' } }, 'Create a new blank workbook or open a local .xlsx / .xlsm file from the File menu above.'),
                  h('button', { type: 'button', className: 'dsh-xl-btn-primary', onClick: function () { setShowNewForm(true); } }, '+ New Blank Workbook')
                ),

            // Insert Function (fx) Modal
            fxModalOpen
              ? h(
                  'div', { className: 'dsh-xl-modal-backdrop', onClick: function () { setFxModalOpen(false); } },
                  h(
                    'div', { className: 'dsh-xl-modal', onClick: function (e) { e.stopPropagation(); } },
                    h('div', { className: 'dsh-xl-modal-hdr' }, h('span', null, 'Insert Function (fx) — 125+ Excel Functions'), h('button', { type: 'button', className: 'dsh-xl-qat-btn', onClick: function () { setFxModalOpen(false); } }, '✕')),
                    h(
                      'div', { className: 'dsh-xl-modal-body' },
                      h('div', { style: { display: 'flex', gap: '8px' } },
                        h('input', { className: 'dsh-xl-rinput', style: { flex: 1 }, placeholder: 'Search function (e.g. VLOOKUP, SUMIF, PMT, IF)...', value: fxSearch, onChange: function (e) { setFxSearch(e.target.value); } }),
                        h('select', { className: 'dsh-xl-rselect', value: fxCategory, onChange: function (e) { setFxCategory(e.target.value); } },
                          ['All', 'Math & Trig', 'Statistical', 'Logical', 'Lookup & Reference', 'Text', 'Date & Time', 'Financial', 'Information'].map(function (c) { return h('option', { key: c, value: c }, c); })
                        )
                      ),
                      h(
                        'div', { className: 'dsh-xl-fnlist' },
                        filteredFunctions.map(function (fn) {
                          var isSel = fxSelected && fxSelected.name === fn.name;
                          return h('div', {
                            key: fn.name, className: 'dsh-xl-fnrow ' + (isSel ? 'dsh-xl-fnrow-sel' : ''),
                            onClick: function () { setFxSelected(fn); },
                            onDoubleClick: function () { handleInsertFormula(fn.name); setFxModalOpen(false); },
                          }, h('span', null, fn.name), h('span', { style: { opacity: 0.6, fontSize: '11px' } }, fn.cat));
                        })
                      ),
                      fxSelected
                        ? h('div', { style: { padding: '8px', background: 'rgba(16,124,65,0.06)', borderRadius: '4px', fontSize: '12px' } },
                            h('div', { style: { fontFamily: 'ui-monospace, monospace', fontWeight: 700, color: '#107c41', marginBottom: '4px' } }, fxSelected.sig),
                            h('div', null, fxSelected.desc)
                          )
                        : null,
                      h('div', { style: { display: 'flex', justifyContent: 'flex-end', gap: '8px' } },
                        h('button', { type: 'button', className: 'dsh-xl-rbtn', onClick: function () { setFxModalOpen(false); } }, 'Cancel'),
                        h('button', { type: 'button', className: 'dsh-xl-btn-primary', disabled: !fxSelected, onClick: function () { if (fxSelected) handleInsertFormula(fxSelected.name); setFxModalOpen(false); } }, 'Insert Function')
                      )
                    )
                  )
                )
              : null,

            // Right-Click Context Menu
            ctxMenu
              ? h('div', { className: 'dsh-xl-ctxmenu', style: { left: ctxMenu.x + 'px', top: ctxMenu.y + 'px' }, onClick: function (e) { e.stopPropagation(); } },
                  h('button', { type: 'button', className: 'dsh-xl-menu-item', onClick: function () { handleCopy(true); setCtxMenu(null); } }, h('span', null, 'Cut'), h('span', { style: { opacity: 0.5 } }, 'Ctrl+X')),
                  h('button', { type: 'button', className: 'dsh-xl-menu-item', onClick: function () { handleCopy(false); setCtxMenu(null); } }, h('span', null, 'Copy'), h('span', { style: { opacity: 0.5 } }, 'Ctrl+C')),
                  h('button', { type: 'button', className: 'dsh-xl-menu-item', onClick: function () { handlePaste(); setCtxMenu(null); } }, h('span', null, 'Paste'), h('span', { style: { opacity: 0.5 } }, 'Ctrl+V')),
                  h('div', { className: 'dsh-xl-menu-sep' }),
                  h('button', { type: 'button', className: 'dsh-xl-menu-item', onClick: function () { runSheetOp('insert_rows', { index: ctxMenu.row, amount: 1 }); setCtxMenu(null); } }, 'Insert Row Above'),
                  h('button', { type: 'button', className: 'dsh-xl-menu-item', onClick: function () { if (window.confirm('Delete row ' + ctxMenu.row + '?')) runSheetOp('delete_rows', { index: ctxMenu.row, amount: 1 }); setCtxMenu(null); } }, 'Delete Row'),
                  h('button', { type: 'button', className: 'dsh-xl-menu-item', onClick: function () { runSheetOp('insert_columns', { index: ctxMenu.col, amount: 1 }); setCtxMenu(null); } }, 'Insert Column Left'),
                  h('button', { type: 'button', className: 'dsh-xl-menu-item', onClick: function () { if (window.confirm('Delete column ' + colIndexToLetter(ctxMenu.col) + '?')) runSheetOp('delete_columns', { index: ctxMenu.col, amount: 1 }); setCtxMenu(null); } }, 'Delete Column'),
                  h('div', { className: 'dsh-xl-menu-sep' }),
                  h('button', { type: 'button', className: 'dsh-xl-menu-item', onClick: function () { setFxModalOpen(true); setCtxMenu(null); } }, 'Insert Function (fx)...'),
                  h('button', { type: 'button', className: 'dsh-xl-menu-item', onClick: function () { handleClearSelection('contents'); setCtxMenu(null); } }, 'Clear Contents')
                )
              : null,

            // 6. Sheet Tabs
            sheet
              ? h('div', { className: 'dsh-xl-sheetbar' },
                  (sheet.sheets || []).map(function (name) {
                    var active = name === sheet.active_sheet;
                    return h('button', {
                      key: name, type: 'button', className: 'dsh-xl-sheettab ' + (active ? 'dsh-xl-sheettab-active' : ''),
                      title: 'Double-click to rename',
                      onClick: function () { if (!active) openWorkbook(sheet.file_path, name); },
                      onDoubleClick: function () { var nm = window.prompt('Rename worksheet:', name); if (nm && nm.trim() && nm.trim() !== name) runSheetOp('rename_sheet', { sheet_name: name, new_name: nm.trim() }); },
                    }, name);
                  }),
                  h('button', { type: 'button', className: 'dsh-xl-addsheet', title: 'New sheet', onClick: function () { var nm = window.prompt('New worksheet name:', 'Sheet' + ((sheet.sheets || []).length + 1)); if (nm && nm.trim()) runSheetOp('create_sheet', { new_name: nm.trim() }); } }, '+'),
                  h('button', { type: 'button', className: 'dsh-xl-rbtn', style: { marginLeft: 'auto' }, onClick: function () { runSheetOp('duplicate_sheet'); } }, 'Duplicate'),
                  (sheet.sheets || []).length > 1
                    ? h('button', { type: 'button', className: 'dsh-xl-rbtn', style: { color: '#c00000' }, onClick: function () { if (window.confirm('Delete "' + sheet.active_sheet + '"?')) runSheetOp('delete_sheet'); } }, 'Delete')
                    : null
                )
              : null,

            // 7. Excel Green Status Bar
            h(
              'div', { className: 'dsh-xl-statusbar' },
              h('span', null, editing ? 'Edit' : (sheet ? 'Ready · ' + sheet.file_path : 'Ready')),
              h('div', { style: { display: 'inline-flex', alignItems: 'center', gap: '14px' } },
                selStats ? h('span', null, (selStats.avg !== null ? 'Average: ' + selStats.avg + '   ' : '') + 'Count: ' + selStats.count + (selStats.avg !== null ? '   Sum: ' + selStats.sum : '')) : null,
                h('span', { style: { display: 'inline-flex', alignItems: 'center', gap: '4px' } },
                  h('button', { type: 'button', className: 'dsh-xl-qat-btn', style: { height: '18px' }, onClick: function () { setZoom(function (z) { return Math.max(50, z - 10); }); } }, '−'),
                  h('span', null, zoom + '%'),
                  h('button', { type: 'button', className: 'dsh-xl-qat-btn', style: { height: '18px' }, onClick: function () { setZoom(function (z) { return Math.min(200, z + 10); }); } }, '+')
                ),
                h('span', null, 'v0.5.0')
              )
            )
          );
        }

        exports.name = PLUGIN_ID;
        exports.inject = ['slots'];
        exports.apply = function apply(ctx) {
          try {
            if (!ctx || !ctx.slots || typeof ctx.slots.inject !== 'function') return;
            ctx.slots.inject('main', function () {
              try { return ctx.slots.register({ name: 'main', key: PANEL_ID }, ExcelStudioPanel); } catch (_e) { return undefined; }
            });
            ctx.slots.inject('sidebar.panellist', function () {
              try { return ctx.slots.register({ name: 'sidebar.panellist', id: PANEL_ID, order: 45, label: 'Excel Studio' }, ExcelStudioIcon); } catch (_e) { return undefined; }
            });
          } catch (_err) {
            // Never block DeepSeek Harness Web UI boot
          }
        };

        return module.exports;
      },
    });
  } catch (_err) {
    // Ignore duplicate registration
  }
})();
