import * as ExcelJSRaw from 'exceljs'
import { copyFileSync, existsSync, mkdirSync, readdirSync, statSync } from 'node:fs'
import { createRequire } from 'node:module'
import path from 'node:path'

function resolveExcelWorkbookCtor() {
  const candidates = [
    ExcelJSRaw?.Workbook,
    ExcelJSRaw?.default?.Workbook,
    ExcelJSRaw?.default?.default?.Workbook,
  ]
  for (const Ctor of candidates) {
    if (typeof Ctor === 'function') return Ctor
  }
  try {
    const req = createRequire(import.meta.url)
    const mod = req('exceljs')
    if (typeof mod?.Workbook === 'function') return mod.Workbook
    if (typeof mod?.default?.Workbook === 'function') return mod.default.Workbook
  } catch (_err) {
    // Fallback failed; throw descriptive error below
  }
  throw new Error('Failed to resolve ExcelJS.Workbook constructor')
}

const ExcelWorkbook = resolveExcelWorkbookCtor()

const ALLOWED_EXTENSIONS = new Set(['.xlsx', '.xlsm'])
const DEFAULT_GRID_ROWS = 40
const DEFAULT_GRID_COLS = 16
const MAX_GRID_ROWS = 200
const MAX_GRID_COLS = 50

export class SecurityError extends Error {
  constructor(message) {
    super(message)
    this.name = 'SecurityError'
    this.code = 'SECURITY_ERROR'
  }
}

export class OperationError extends Error {
  constructor(message) {
    super(message)
    this.name = 'OperationError'
    this.code = 'OPERATION_ERROR'
  }
}

export function parseAllowedDirs(raw) {
  if (Array.isArray(raw)) {
    return raw.map((d) => path.resolve(String(d).trim())).filter(Boolean)
  }
  const str = String(raw || '').trim()
  if (!str) return [path.resolve(process.cwd())]
  const sep = process.platform === 'win32' || str.includes(';') ? ';' : ':'
  return str
    .split(sep)
    .map((d) => d.trim())
    .filter(Boolean)
    .map((d) => path.resolve(d))
}

function isWithinDir(targetPath, baseDir) {
  const rel = path.relative(baseDir, targetPath)
  return rel === '' || (!rel.startsWith('..') && !path.isAbsolute(rel))
}

export function validateDirectoryPath(dirPath, allowedDirs) {
  if (!dirPath || typeof dirPath !== 'string') {
    throw new SecurityError('Directory path is required.')
  }
  const resolved = path.resolve(dirPath)
  const dirs = parseAllowedDirs(allowedDirs)
  if (!dirs.some((base) => isWithinDir(resolved, base))) {
    throw new SecurityError(`Access denied: '${resolved}' is outside allowed directories.`)
  }
  if (!existsSync(resolved) || !statSync(resolved).isDirectory()) {
    throw new OperationError(`Directory does not exist: '${resolved}'.`)
  }
  return resolved
}

export function validateFilePath(filePath, allowedDirs, { mustExist = true } = {}) {
  if (!filePath || typeof filePath !== 'string' || !filePath.trim()) {
    throw new SecurityError('File path must be a non-empty string.')
  }
  const resolved = path.resolve(filePath.trim())
  const dirs = parseAllowedDirs(allowedDirs)
  if (!dirs.some((base) => isWithinDir(resolved, base))) {
    throw new SecurityError(`Access denied: '${resolved}' is outside allowed directories.`)
  }
  const ext = path.extname(resolved).toLowerCase()
  if (!ALLOWED_EXTENSIONS.has(ext)) {
    throw new SecurityError(`Unsupported file extension '${ext}'. Only .xlsx and .xlsm are supported.`)
  }
  if (mustExist && !existsSync(resolved)) {
    throw new OperationError(`Workbook not found: '${resolved}'.`)
  }
  return resolved
}

export function isExcelLockPresent(filePath) {
  const dir = path.dirname(filePath)
  const base = path.basename(filePath)
  return existsSync(path.join(dir, `~$${base}`))
}

export function createBackup(filePath, autoBackup = true) {
  if (!autoBackup || !existsSync(filePath)) return null
  const dir = path.dirname(filePath)
  const ext = path.extname(filePath)
  const stem = path.basename(filePath, ext)
  const backupDir = path.join(dir, '.dsh_backup')
  mkdirSync(backupDir, { recursive: true })
  const stamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19)
  const backupPath = path.join(backupDir, `${stem}_${stamp}${ext}`)
  copyFileSync(filePath, backupPath)
  return backupPath
}

export function colIndexToLetter(index) {
  let col = ''
  let n = index
  while (n > 0) {
    const rem = (n - 1) % 26
    col = String.fromCharCode(65 + rem) + col
    n = Math.floor((n - 1) / 26)
  }
  return col
}

export function letterToColIndex(letters) {
  const s = String(letters || '').toUpperCase()
  let n = 0
  for (let i = 0; i < s.length; i++) {
    n = n * 26 + (s.charCodeAt(i) - 64)
  }
  return n
}

export function parseCellCoord(coord) {
  const m = /^\$?([A-Za-z]+)\$?(\d+)$/.exec(String(coord || '').trim())
  if (!m) throw new OperationError(`Invalid cell coordinate: '${coord}'`)
  return { col: letterToColIndex(m[1]), row: parseInt(m[2], 10) }
}

export function parseRangeCoords(rangeStr) {
  const trimmed = String(rangeStr || '').trim()
  if (!trimmed) throw new OperationError('Range string is required.')
  const parts = trimmed.split(':')
  const start = parseCellCoord(parts[0])
  const end = parts[1] ? parseCellCoord(parts[1]) : start
  return {
    minRow: Math.min(start.row, end.row),
    maxRow: Math.max(start.row, end.row),
    minCol: Math.min(start.col, end.col),
    maxCol: Math.max(start.col, end.col),
  }
}

function normalizeArgbToHex(colorObj) {
  if (!colorObj || typeof colorObj !== 'object') return null
  const argb = colorObj.argb
  if (typeof argb !== 'string') return null
  const cleaned = argb.trim().toUpperCase()
  if (cleaned.length === 8) {
    if (cleaned === '00000000') return null
    return `#${cleaned.slice(2)}`
  }
  if (cleaned.length === 6) {
    return `#${cleaned}`
  }
  return null
}

function toArgb(hexColor) {
  const hex = String(hexColor || '').replace(/^#/, '').trim().toUpperCase()
  if (hex.length === 6) return `FF${hex}`
  if (hex.length === 8) return hex
  return null
}

function serializePrimitive(val) {
  if (val === null || val === undefined) return ''
  if (val instanceof Date) return val.toISOString()
  if (typeof val === 'number' || typeof val === 'boolean' || typeof val === 'string') return val
  if (typeof val === 'object') {
    if ('richText' in val && Array.isArray(val.richText)) {
      return val.richText.map((r) => r.text || '').join('')
    }
    if ('text' in val && typeof val.text === 'string') {
      return val.text
    }
    if ('result' in val && val.result !== undefined) {
      return serializePrimitive(val.result)
    }
    if ('formula' in val && typeof val.formula === 'string') {
      return `=${val.formula}`
    }
    if ('error' in val) {
      return String(val.error)
    }
  }
  return String(val)
}

function extractCellInfo(cell) {
  const val = cell.value
  const style = {}

  if (cell.font) {
    if (cell.font.name) style.fontFamily = cell.font.name
    if (cell.font.size) style.fontSize = cell.font.size
    if (cell.font.bold) style.bold = true
    if (cell.font.italic) style.italic = true
    if (cell.font.underline) style.underline = true
    if (cell.font.strike) style.strike = true
    const color = normalizeArgbToHex(cell.font.color)
    if (color && color !== '#000000') style.color = color
  }
  if (cell.fill && cell.fill.type === 'pattern' && cell.fill.pattern === 'solid') {
    const bg = normalizeArgbToHex(cell.fill.fgColor) || normalizeArgbToHex(cell.fill.bgColor)
    if (bg && bg !== '#000000') style.fill = bg
  }
  if (cell.alignment) {
    if (cell.alignment.horizontal) style.align = cell.alignment.horizontal
    if (cell.alignment.vertical) style.valign = cell.alignment.vertical
    if (cell.alignment.wrapText) style.wrap = true
    if (cell.alignment.indent) style.indent = cell.alignment.indent
  }
  if (cell.border) {
    const b = cell.border
    if (b.top || b.bottom || b.left || b.right) {
      style.border = {
        top: Boolean(b.top?.style),
        bottom: Boolean(b.bottom?.style),
        left: Boolean(b.left?.style),
        right: Boolean(b.right?.style),
      }
    }
  }
  if (cell.numFmt && cell.numFmt !== 'General') {
    style.numberFormat = cell.numFmt
  }

  let noteText = ''
  if (typeof cell.note === 'string') {
    noteText = cell.note
  } else if (cell.note && Array.isArray(cell.note.texts)) {
    noteText = cell.note.texts.map((t) => t.text || '').join('')
  }

  let hyperlinkUrl = ''
  if (val && typeof val === 'object' && typeof val.hyperlink === 'string') {
    hyperlinkUrl = val.hyperlink
  }

  const hasStyle = Object.keys(style).length > 0
  if ((val === null || val === undefined) && !hasStyle && !noteText && !hyperlinkUrl) {
    return null
  }

  const entry = {}
  if (val && typeof val === 'object' && typeof val.formula === 'string') {
    const formulaStr = val.formula.startsWith('=') ? val.formula : `=${val.formula}`
    entry.formula = formulaStr
    entry.raw = formulaStr
    entry.display = val.result !== undefined ? serializePrimitive(val.result) : formulaStr
  } else if (hyperlinkUrl) {
    const txt = typeof val.text === 'string' ? val.text : hyperlinkUrl
    entry.raw = txt
    entry.display = txt
    entry.hyperlink = hyperlinkUrl
  } else {
    const serialized = serializePrimitive(val)
    entry.raw = serialized
    entry.display = serialized
  }
  if (hasStyle) entry.style = style
  if (noteText) entry.note = noteText
  return entry
}

function coerceInputValue(raw) {
  if (raw === null || raw === undefined) return null
  if (typeof raw !== 'string') return raw
  const stripped = raw.trim()
  if (!stripped) return null
  if (stripped.startsWith('=')) {
    return { formula: stripped.slice(1) }
  }
  const low = stripped.toLowerCase()
  if (low === 'true') return true
  if (low === 'false') return false
  if (/^-?\d+(\.\d+)?$/.test(stripped)) {
    const num = Number(stripped)
    if (Number.isFinite(num)) return num
  }
  return raw
}

function applyStylePatch(cell, fmt) {
  if (
    'bold' in fmt ||
    'italic' in fmt ||
    'underline' in fmt ||
    'strike' in fmt ||
    'font_color' in fmt ||
    'font_size' in fmt ||
    'font_family' in fmt
  ) {
    const prev = cell.font || {}
    const nextFont = { ...prev }
    if ('font_family' in fmt && fmt.font_family) nextFont.name = String(fmt.font_family)
    if ('font_size' in fmt && fmt.font_size) nextFont.size = Number(fmt.font_size)
    if ('bold' in fmt) nextFont.bold = Boolean(fmt.bold)
    if ('italic' in fmt) nextFont.italic = Boolean(fmt.italic)
    if ('underline' in fmt) nextFont.underline = Boolean(fmt.underline)
    if ('strike' in fmt) nextFont.strike = Boolean(fmt.strike)
    if ('font_color' in fmt) {
      const argb = toArgb(fmt.font_color)
      nextFont.color = argb ? { argb } : undefined
    }
    cell.font = nextFont
  }

  if ('fill_color' in fmt) {
    const argb = toArgb(fmt.fill_color)
    if (argb) {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb },
      }
    } else {
      cell.fill = { type: 'pattern', pattern: 'none' }
    }
  }

  if ('align' in fmt || 'valign' in fmt || 'wrap' in fmt || 'indent' in fmt) {
    const prev = cell.alignment || {}
    cell.alignment = {
      ...prev,
      horizontal: 'align' in fmt ? (fmt.align || undefined) : prev.horizontal,
      vertical: 'valign' in fmt ? (fmt.valign || undefined) : prev.vertical,
      wrapText: 'wrap' in fmt ? Boolean(fmt.wrap) : prev.wrapText,
      indent: 'indent' in fmt ? Math.max(0, Number(fmt.indent) || 0) : prev.indent,
    }
  }

  if ('border_preset' in fmt) {
    const thin = { style: 'thin', color: { argb: 'FF808080' } }
    const p = fmt.border_preset
    if (p === 'all') {
      cell.border = { top: thin, bottom: thin, left: thin, right: thin }
    } else if (p === 'bottom') {
      cell.border = { ...(cell.border || {}), bottom: thin }
    } else if (p === 'top_bottom') {
      cell.border = { ...(cell.border || {}), top: thin, bottom: { style: 'double', color: { argb: 'FF404040' } } }
    } else if (p === 'none') {
      cell.border = {}
    }
  }

  if ('number_format' in fmt && fmt.number_format) {
    cell.numFmt = String(fmt.number_format)
  }
}

async function openWorkbookFile(resolvedPath) {
  const wb = new ExcelWorkbook()
  await wb.xlsx.readFile(resolvedPath)
  return wb
}

function scanDirectoryForWorkbooks(dir, recursive, seen, out) {
  let entries
  try {
    entries = readdirSync(dir, { withFileTypes: true })
  } catch {
    return
  }
  for (const entry of entries) {
    if (entry.name === '.dsh_backup' || entry.name === 'node_modules' || entry.name === '.venv') {
      continue
    }
    const full = path.join(dir, entry.name)
    if (entry.isDirectory()) {
      if (recursive) scanDirectoryForWorkbooks(full, recursive, seen, out)
      continue
    }
    if (!entry.isFile() || entry.name.startsWith('~$')) continue
    const ext = path.extname(entry.name).toLowerCase()
    if (!ALLOWED_EXTENSIONS.has(ext)) continue
    const resolved = path.resolve(full)
    if (seen.has(resolved)) continue
    seen.add(resolved)
    const st = statSync(resolved)
    out.push({
      name: path.basename(resolved),
      path: resolved,
      directory: path.dirname(resolved),
      size_bytes: st.size,
      modified: st.mtime.toISOString().slice(0, 19),
      has_vba: ext === '.xlsm',
      is_locked: isExcelLockPresent(resolved),
    })
  }
}

export async function listWorkbooks({ allowedDirs, directory, recursive = true }) {
  const dirs = directory
    ? [validateDirectoryPath(directory, allowedDirs)]
    : parseAllowedDirs(allowedDirs).filter((d) => existsSync(d))
  const seen = new Set()
  const workbooks = []
  for (const dir of dirs) {
    scanDirectoryForWorkbooks(dir, Boolean(recursive), seen, workbooks)
  }
  workbooks.sort((a, b) => a.name.localeCompare(b.name))
  return {
    workbooks,
    allowed_directories: parseAllowedDirs(allowedDirs),
  }
}

export async function createWorkbook({
  allowedDirs,
  directory,
  filename,
  file_path,
  sheet_name = 'Sheet1',
  overwrite = false,
}) {
  let target = file_path
  if (!target) {
    const cleanName = String(filename || '').trim()
    if (!cleanName) throw new OperationError('filename or file_path is required.')
    const finalName = cleanName.toLowerCase().endsWith('.xlsx') ? cleanName : `${cleanName}.xlsx`
    const baseDir = directory
      ? validateDirectoryPath(directory, allowedDirs)
      : parseAllowedDirs(allowedDirs)[0]
    target = path.join(baseDir, finalName)
  }
  const resolved = validateFilePath(target, allowedDirs, { mustExist: false })
  if (existsSync(resolved) && !overwrite) {
    throw new OperationError(`Workbook '${path.basename(resolved)}' already exists.`)
  }
  mkdirSync(path.dirname(resolved), { recursive: true })
  const wb = new ExcelWorkbook()
  const title = String(sheet_name || 'Sheet1').trim() || 'Sheet1'
  wb.addWorksheet(title)
  await wb.xlsx.writeFile(resolved)
  return loadSheet({ allowedDirs, file_path: resolved, sheet_name: title })
}

export async function getWorkbookInfo({ allowedDirs, file_path }) {
  const resolved = validateFilePath(file_path, allowedDirs, { mustExist: true })
  const wb = await openWorkbookFile(resolved)
  const st = statSync(resolved)
  return {
    file_path: resolved,
    filename: path.basename(resolved),
    size_bytes: st.size,
    has_vba: path.extname(resolved).toLowerCase() === '.xlsm',
    is_locked: isExcelLockPresent(resolved),
    sheets: wb.worksheets.map((ws) => ({
      name: ws.name,
      max_row: ws.actualRowCount || ws.rowCount || 0,
      max_column: ws.actualColumnCount || ws.columnCount || 0,
      state: ws.state || 'visible',
    })),
  }
}

export async function loadSheet({
  allowedDirs,
  file_path,
  sheet_name,
  start_row = 1,
  start_col = 1,
  row_count,
  col_count,
}) {
  const resolved = validateFilePath(file_path, allowedDirs, { mustExist: true })
  const wb = await openWorkbookFile(resolved)
  const sheets = wb.worksheets.map((w) => w.name)
  const ws = (sheet_name && wb.getWorksheet(sheet_name)) || wb.worksheets[0]
  if (!ws) throw new OperationError('Workbook contains no worksheets.')

  const maxRow = ws.rowCount || 0
  const maxCol = ws.columnCount || 0
  const sRow = Math.max(1, Number(start_row) || 1)
  const sCol = Math.max(1, Number(start_col) || 1)
  const rCount = Math.min(MAX_GRID_ROWS, Math.max(1, Number(row_count) || Math.max(DEFAULT_GRID_ROWS, maxRow + 10)))
  const cCount = Math.min(MAX_GRID_COLS, Math.max(1, Number(col_count) || Math.max(DEFAULT_GRID_COLS, maxCol + 4)))

  const endRow = sRow + rCount - 1
  const endCol = sCol + cCount - 1

  const cells = {}
  for (let r = sRow; r <= endRow; r++) {
    const rowObj = ws.findRow(r)
    if (!rowObj) continue
    for (let c = sCol; c <= endCol; c++) {
      const cell = rowObj.findCell(c)
      if (!cell) continue
      const info = extractCellInfo(cell)
      if (info) {
        cells[`${colIndexToLetter(c)}${r}`] = info
      }
    }
  }

  const column_widths = {}
  for (let c = sCol; c <= endCol; c++) {
    const colObj = ws.getColumn(c)
    if (colObj && typeof colObj.width === 'number' && colObj.width > 0) {
      column_widths[colIndexToLetter(c)] = Math.round(colObj.width * 100) / 100
    }
  }

  const row_heights = {}
  for (let r = sRow; r <= endRow; r++) {
    const rowObj = ws.findRow(r)
    if (rowObj && typeof rowObj.height === 'number' && rowObj.height > 0) {
      row_heights[r] = Math.round(rowObj.height * 1.333)
    }
  }

  const merges = Array.isArray(ws.model && ws.model.merges) ? ws.model.merges.slice() : []
  const frozenView = Array.isArray(ws.views) ? ws.views.find((v) => v && v.state === 'frozen') : null
  const freeze = frozenView ? { xSplit: frozenView.xSplit || 0, ySplit: frozenView.ySplit || 0 } : { xSplit: 0, ySplit: 0 }

  return {
    file_path: resolved,
    filename: path.basename(resolved),
    sheets,
    active_sheet: ws.name,
    has_vba: path.extname(resolved).toLowerCase() === '.xlsm',
    is_locked: isExcelLockPresent(resolved),
    max_row: maxRow,
    max_column: maxCol,
    start_row: sRow,
    start_col: sCol,
    row_count: rCount,
    col_count: cCount,
    column_widths,
    row_heights,
    merges,
    freeze,
    cells,
  }
}

export async function saveEdits({
  allowedDirs,
  autoBackup = true,
  file_path,
  target_path,
  sheet_name,
  edits = [],
  formats = [],
  column_widths = {},
  row_heights = {},
  notes = {},
  hyperlinks = {},
  merges = [],
  unmerges = [],
  freeze = null,
}) {
  const resolved = validateFilePath(file_path, allowedDirs, { mustExist: true })
  const destPath = target_path
    ? validateFilePath(target_path, allowedDirs, { mustExist: false })
    : resolved
  const wb = await openWorkbookFile(resolved)
  const ws = (sheet_name && wb.getWorksheet(sheet_name)) || wb.worksheets[0]
  if (!ws) throw new OperationError('Worksheet not found.')

  for (const edit of edits) {
    const coord = String(edit.cell || '').trim().toUpperCase()
    if (!coord) continue
    ws.getCell(coord).value = coerceInputValue(edit.value)
  }

  for (const fmt of formats) {
    const target = String(fmt.cell || fmt.range || '').trim().toUpperCase()
    if (!target) continue
    if (target.includes(':')) {
      const { minRow, maxRow, minCol, maxCol } = parseRangeCoords(target)
      for (let r = minRow; r <= maxRow; r++) {
        for (let c = minCol; c <= maxCol; c++) {
          applyStylePatch(ws.getCell(r, c), fmt)
        }
      }
    } else {
      applyStylePatch(ws.getCell(target), fmt)
    }
  }

  if (column_widths && typeof column_widths === 'object') {
    for (const [letter, width] of Object.entries(column_widths)) {
      const w = Number(width)
      if (Number.isFinite(w) && w > 0) {
        const colIdx = letterToColIndex(letter)
        if (colIdx >= 1) ws.getColumn(colIdx).width = w
      }
    }
  }

  if (row_heights && typeof row_heights === 'object') {
    for (const [rowKey, px] of Object.entries(row_heights)) {
      const rIdx = Number(rowKey)
      const hPx = Number(px)
      if (Number.isInteger(rIdx) && rIdx >= 1 && Number.isFinite(hPx) && hPx >= 12) {
        ws.getRow(rIdx).height = Math.round((hPx / 1.333) * 10) / 10
      }
    }
  }

  if (notes && typeof notes === 'object') {
    for (const [coord, text] of Object.entries(notes)) {
      if (!/^[A-Za-z]+\d+$/.test(coord)) continue
      const cell = ws.getCell(coord.toUpperCase())
      cell.note = text ? String(text) : undefined
    }
  }

  if (hyperlinks && typeof hyperlinks === 'object') {
    for (const [coord, url] of Object.entries(hyperlinks)) {
      if (!/^[A-Za-z]+\d+$/.test(coord)) continue
      const cell = ws.getCell(coord.toUpperCase())
      const curText = cell.value && typeof cell.value === 'object' && cell.value.text ? cell.value.text : String(cell.value ?? url ?? '')
      if (url) {
        cell.value = { text: curText || String(url), hyperlink: String(url) }
      } else if (cell.value && typeof cell.value === 'object' && cell.value.hyperlink) {
        cell.value = curText
      }
    }
  }

  if (Array.isArray(unmerges)) {
    for (const range of unmerges) {
      try {
        if (typeof range === 'string' && range.includes(':')) {
          ws.unMergeCells(range)
        }
      } catch {
        // Ignore if not merged
      }
    }
  }

  if (Array.isArray(merges)) {
    for (const range of merges) {
      try {
        if (typeof range === 'string' && range.includes(':')) {
          ws.mergeCells(range)
        }
      } catch {
        // Ignore already-merged ranges
      }
    }
  }

  if (freeze && typeof freeze === 'object') {
    const xSplit = Math.max(0, Number(freeze.xSplit) || 0)
    const ySplit = Math.max(0, Number(freeze.ySplit) || 0)
    ws.views = xSplit > 0 || ySplit > 0 ? [{ state: 'frozen', xSplit, ySplit }] : [{ state: 'normal' }]
  }

  mkdirSync(path.dirname(destPath), { recursive: true })
  const backup_path = existsSync(destPath) ? createBackup(destPath, autoBackup) : null
  await wb.xlsx.writeFile(destPath)

  const data = await loadSheet({ allowedDirs, file_path: destPath, sheet_name: ws.name })
  return {
    ...data,
    backup_path,
    edits_applied: edits.length,
    formats_applied: formats.length,
  }
}

export async function sheetOp({
  allowedDirs,
  autoBackup = true,
  file_path,
  sheet_name,
  op,
  new_name,
  index = 1,
  amount = 1,
  ascending = true,
}) {
  const resolved = validateFilePath(file_path, allowedDirs, { mustExist: true })
  const wb = await openWorkbookFile(resolved)
  const currentWs = (sheet_name && wb.getWorksheet(sheet_name)) || wb.worksheets[0]
  let targetSheet = currentWs ? currentWs.name : wb.worksheets[0]?.name

  if (op === 'create_sheet') {
    const clean = String(new_name || '').trim()
    if (!clean) throw new OperationError('new_name is required for create_sheet.')
    if (wb.getWorksheet(clean)) throw new OperationError(`Worksheet '${clean}' already exists.`)
    wb.addWorksheet(clean)
    targetSheet = clean
  } else if (op === 'duplicate_sheet') {
    if (!currentWs) throw new OperationError(`Worksheet '${sheet_name}' not found.`)
    let copyName = String(new_name || `${currentWs.name} (Copy)`).trim()
    let suffix = 2
    while (wb.getWorksheet(copyName)) {
      copyName = `${currentWs.name} (${suffix++})`
    }
    const copyWs = wb.addWorksheet(copyName)
    currentWs.eachRow({ includeEmpty: false }, (row, rIdx) => {
      row.eachCell({ includeEmpty: false }, (cell, cIdx) => {
        const dest = copyWs.getCell(rIdx, cIdx)
        dest.value = cell.value
        if (cell.style) dest.style = JSON.parse(JSON.stringify(cell.style))
      })
    })
    targetSheet = copyName
  } else if (op === 'rename_sheet') {
    const clean = String(new_name || '').trim()
    if (!clean) throw new OperationError('new_name is required for rename_sheet.')
    if (!currentWs) throw new OperationError(`Worksheet '${sheet_name}' not found.`)
    if (wb.getWorksheet(clean)) throw new OperationError(`Worksheet '${clean}' already exists.`)
    currentWs.name = clean
    targetSheet = clean
  } else if (op === 'delete_sheet') {
    if (!currentWs) throw new OperationError(`Worksheet '${sheet_name}' not found.`)
    if (wb.worksheets.length <= 1) {
      throw new OperationError('Cannot delete the only worksheet in a workbook.')
    }
    wb.removeWorksheet(currentWs.id)
    targetSheet = wb.worksheets[0].name
  } else if (op === 'insert_rows') {
    if (!currentWs) throw new OperationError('Worksheet not found.')
    const idx = Math.max(1, Number(index) || 1)
    const count = Math.max(1, Math.min(100, Number(amount) || 1))
    const emptyRows = Array.from({ length: count }, () => [])
    currentWs.spliceRows(idx, 0, ...emptyRows)
  } else if (op === 'delete_rows') {
    if (!currentWs) throw new OperationError('Worksheet not found.')
    const idx = Math.max(1, Number(index) || 1)
    const count = Math.max(1, Math.min(100, Number(amount) || 1))
    currentWs.spliceRows(idx, count)
  } else if (op === 'insert_columns') {
    if (!currentWs) throw new OperationError('Worksheet not found.')
    const idx = Math.max(1, Number(index) || 1)
    const count = Math.max(1, Math.min(50, Number(amount) || 1))
    const emptyCols = Array.from({ length: count }, () => [])
    currentWs.spliceColumns(idx, 0, ...emptyCols)
  } else if (op === 'delete_columns') {
    if (!currentWs) throw new OperationError('Worksheet not found.')
    const idx = Math.max(1, Number(index) || 1)
    const count = Math.max(1, Math.min(50, Number(amount) || 1))
    currentWs.spliceColumns(idx, count)
  } else if (op === 'sort_column') {
    if (!currentWs) throw new OperationError('Worksheet not found.')
    const colIdx = Math.max(1, Number(index) || 1)
    const maxR = currentWs.rowCount || 0
    const maxC = currentWs.columnCount || 0
    if (maxR > 2 && maxC >= 1) {
      const dataRows = []
      for (let r = 2; r <= maxR; r++) {
        const vals = []
        let hasAny = false
        for (let c = 1; c <= maxC; c++) {
          const v = currentWs.getCell(r, c).value
          if (v !== null && v !== undefined && v !== '') hasAny = true
          vals.push(v)
        }
        if (hasAny) dataRows.push(vals)
      }
      const dir = ascending ? 1 : -1
      dataRows.sort((a, b) => {
        const va = serializePrimitive(a[colIdx - 1])
        const vb = serializePrimitive(b[colIdx - 1])
        const na = Number(va)
        const nb = Number(vb)
        if (va !== '' && vb !== '' && Number.isFinite(na) && Number.isFinite(nb)) {
          return (na - nb) * dir
        }
        return String(va).localeCompare(String(vb)) * dir
      })
      for (let i = 0; i < dataRows.length; i++) {
        for (let c = 1; c <= maxC; c++) {
          currentWs.getCell(2 + i, c).value = dataRows[i][c - 1]
        }
      }
    }
  } else {
    throw new OperationError(`Unsupported sheet operation: ${op}`)
  }

  const backup_path = createBackup(resolved, autoBackup)
  await wb.xlsx.writeFile(resolved)

  const data = await loadSheet({ allowedDirs, file_path: resolved, sheet_name: targetSheet })
  return { ...data, backup_path }
}

export async function readRange({ allowedDirs, file_path, sheet_name, range_str }) {
  const resolved = validateFilePath(file_path, allowedDirs, { mustExist: true })
  const wb = await openWorkbookFile(resolved)
  const ws = (sheet_name && wb.getWorksheet(sheet_name)) || wb.worksheets[0]
  if (!ws) throw new OperationError('Worksheet not found.')

  const { minRow, maxRow, minCol, maxCol } = range_str
    ? parseRangeCoords(range_str)
    : {
        minRow: 1,
        maxRow: Math.max(1, ws.rowCount || 1),
        minCol: 1,
        maxCol: Math.max(1, ws.columnCount || 1),
      }

  const rows = []
  for (let r = minRow; r <= maxRow; r++) {
    const rowVals = []
    for (let c = minCol; c <= maxCol; c++) {
      rowVals.push(serializePrimitive(ws.getCell(r, c).value))
    }
    rows.push(rowVals)
  }

  return {
    file_path: resolved,
    sheet: ws.name,
    range: `${colIndexToLetter(minCol)}${minRow}:${colIndexToLetter(maxCol)}${maxRow}`,
    rows,
  }
}

export async function writeRange({
  allowedDirs,
  autoBackup = true,
  file_path,
  sheet_name,
  start_cell = 'A1',
  values = [],
}) {
  const resolved = validateFilePath(file_path, allowedDirs, { mustExist: true })
  const wb = await openWorkbookFile(resolved)
  const ws = (sheet_name && wb.getWorksheet(sheet_name)) || wb.worksheets[0]
  if (!ws) throw new OperationError('Worksheet not found.')

  const { row: startRow, col: startCol } = parseCellCoord(start_cell)
  let updated = 0
  for (let r = 0; r < values.length; r++) {
    const rowArr = Array.isArray(values[r]) ? values[r] : [values[r]]
    for (let c = 0; c < rowArr.length; c++) {
      ws.getCell(startRow + r, startCol + c).value = coerceInputValue(rowArr[c])
      updated++
    }
  }

  const backup_path = createBackup(resolved, autoBackup)
  await wb.xlsx.writeFile(resolved)
  return { file_path: resolved, sheet: ws.name, cells_updated: updated, backup_path }
}

export async function clearRange({
  allowedDirs,
  autoBackup = true,
  file_path,
  sheet_name,
  range_str,
  confirm = false,
}) {
  const resolved = validateFilePath(file_path, allowedDirs, { mustExist: true })
  const { minRow, maxRow, minCol, maxCol } = parseRangeCoords(range_str)
  const totalCells = (maxRow - minRow + 1) * (maxCol - minCol + 1)
  if (!confirm) {
    return {
      confirmation_required: true,
      action: 'clear_range',
      range: range_str,
      affected_cells: totalCells,
      message: `Pass confirm=true to clear ${totalCells} cells in ${range_str}.`,
    }
  }
  const wb = await openWorkbookFile(resolved)
  const ws = (sheet_name && wb.getWorksheet(sheet_name)) || wb.worksheets[0]
  if (!ws) throw new OperationError('Worksheet not found.')

  for (let r = minRow; r <= maxRow; r++) {
    for (let c = minCol; c <= maxCol; c++) {
      ws.getCell(r, c).value = null
    }
  }
  const backup_path = createBackup(resolved, autoBackup)
  await wb.xlsx.writeFile(resolved)
  return { file_path: resolved, sheet: ws.name, cleared_cells: totalCells, backup_path }
}

export async function findValue({
  allowedDirs,
  file_path,
  sheet_name,
  query,
  case_sensitive = false,
}) {
  const resolved = validateFilePath(file_path, allowedDirs, { mustExist: true })
  const wb = await openWorkbookFile(resolved)
  const sheets = sheet_name ? [wb.getWorksheet(sheet_name)].filter(Boolean) : wb.worksheets
  const needle = case_sensitive ? String(query) : String(query).toLowerCase()
  const matches = []

  for (const ws of sheets) {
    ws.eachRow((row, rIdx) => {
      row.eachCell((cell, cIdx) => {
        const text = String(serializePrimitive(cell.value))
        const hay = case_sensitive ? text : text.toLowerCase()
        if (hay.includes(needle)) {
          matches.push({
            sheet: ws.name,
            cell: `${colIndexToLetter(cIdx)}${rIdx}`,
            value: text,
          })
        }
      })
    })
  }
  return { query, match_count: matches.length, matches }
}

export async function replaceValue({
  allowedDirs,
  autoBackup = true,
  file_path,
  sheet_name,
  find_text,
  replace_text,
  case_sensitive = false,
}) {
  const resolved = validateFilePath(file_path, allowedDirs, { mustExist: true })
  const wb = await openWorkbookFile(resolved)
  const sheets = sheet_name ? [wb.getWorksheet(sheet_name)].filter(Boolean) : wb.worksheets
  let replaced = 0
  const flags = case_sensitive ? 'g' : 'gi'
  const escaped = String(find_text).replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  const regex = new RegExp(escaped, flags)

  for (const ws of sheets) {
    ws.eachRow((row) => {
      row.eachCell((cell) => {
        if (typeof cell.value === 'string' && regex.test(cell.value)) {
          cell.value = cell.value.replace(regex, String(replace_text))
          replaced++
        }
      })
    })
  }

  const backup_path = replaced > 0 ? createBackup(resolved, autoBackup) : null
  if (replaced > 0) await wb.xlsx.writeFile(resolved)
  return { file_path: resolved, replaced_count: replaced, backup_path }
}

export async function calculateSummary({ allowedDirs, file_path, sheet_name, range_str }) {
  const rangeData = await readRange({ allowedDirs, file_path, sheet_name, range_str })
  const rows = rangeData.rows
  if (rows.length < 2) {
    return {
      sheet: rangeData.sheet,
      range: rangeData.range,
      total_rows: Math.max(0, rows.length - 1),
      summary: {},
    }
  }
  const headers = rows[0].map((h, i) => (String(h || '').trim() || `Column_${i + 1}`))
  const body = rows.slice(1)
  const summary = {}

  for (let c = 0; c < headers.length; c++) {
    const nums = []
    const nonEmpty = []
    for (const row of body) {
      const v = row[c]
      if (v === '' || v === null || v === undefined) continue
      nonEmpty.push(v)
      const n = typeof v === 'number' ? v : Number(v)
      if (Number.isFinite(n) && String(v).trim() !== '') nums.push(n)
    }
    if (nums.length > 0) {
      const sum = nums.reduce((a, b) => a + b, 0)
      const mean = sum / nums.length
      const sorted = [...nums].sort((a, b) => a - b)
      const mid = Math.floor(sorted.length / 2)
      const median =
        sorted.length % 2 === 0 ? (sorted[mid - 1] + sorted[mid]) / 2 : sorted[mid]
      summary[headers[c]] = {
        count: nums.length,
        sum: Math.round(sum * 1e6) / 1e6,
        mean: Math.round(mean * 1e6) / 1e6,
        median: Math.round(median * 1e6) / 1e6,
        min: sorted[0],
        max: sorted[sorted.length - 1],
      }
    } else {
      summary[headers[c]] = {
        count: nonEmpty.length,
        unique: new Set(nonEmpty.map(String)).size,
      }
    }
  }

  return {
    sheet: rangeData.sheet,
    range: rangeData.range,
    total_rows: body.length,
    columns_analyzed: headers.length,
    summary,
  }
}
