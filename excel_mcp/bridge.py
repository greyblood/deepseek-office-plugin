"""
Structured JSON-over-STDIO bridge for the DeepSeek Harness in-app Excel Studio Editor.

Reads a single JSON request from stdin, executes validated workbook operations
using the existing excel_mcp security & core modules, and writes a JSON response
to stdout. Never executes arbitrary code.
"""

from __future__ import annotations

import json
import sys
from datetime import date, datetime, time
from pathlib import Path
from typing import Any, Dict, List, Optional

import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.styles import Alignment, Font, PatternFill

from excel_mcp.config import config
from excel_mcp.core.analysis import AnalysisManager
from excel_mcp.core.workbook import WorkbookManager
from excel_mcp.security.backup import BackupManager
from excel_mcp.security.validator import (
    SecurityError,
    validate_directory_path,
    validate_file_path,
)
from excel_mcp.utils.excel_helpers import is_excel_lock_present


DEFAULT_GRID_ROWS = 40
DEFAULT_GRID_COLS = 16
MAX_GRID_ROWS = 200
MAX_GRID_COLS = 50


def _serialize_value(val: Any) -> Any:
    """Convert cell values to JSON-safe primitives."""
    if val is None:
        return ""
    if isinstance(val, (datetime, date, time)):
        return val.isoformat()
    if isinstance(val, (int, float, bool, str)):
        return val
    return str(val)


def _normalize_argb_to_hex(color_obj: Any) -> Optional[str]:
    """Extract a CSS #RRGGBB hex string from an openpyxl Color object when available."""
    if color_obj is None:
        return None
    if getattr(color_obj, "type", None) != "rgb":
        return None
    rgb = getattr(color_obj, "rgb", None)
    if not isinstance(rgb, str):
        return None
    cleaned = rgb.strip().upper()
    if len(cleaned) == 8:
        # ARGB format; skip fully transparent or default 00000000
        if cleaned.startswith("00"):
            return None
        return f"#{cleaned[2:]}"
    if len(cleaned) == 6:
        return f"#{cleaned}"
    return None


def _extract_cell_style(cell: Any) -> Dict[str, Any]:
    """Return lightweight CSS-friendly style metadata for a cell."""
    style: Dict[str, Any] = {}
    font = getattr(cell, "font", None)
    if font is not None:
        if font.bold:
            style["bold"] = True
        if font.italic:
            style["italic"] = True
        color = _normalize_argb_to_hex(font.color)
        if color and color != "#000000":
            style["color"] = color
    fill = getattr(cell, "fill", None)
    if fill is not None and getattr(fill, "fill_type", None) == "solid":
        bg = _normalize_argb_to_hex(fill.start_color) or _normalize_argb_to_hex(fill.fgColor)
        if bg and bg != "#000000":
            style["fill"] = bg
    alignment = getattr(cell, "alignment", None)
    if alignment is not None and alignment.horizontal:
        style["align"] = alignment.horizontal
    num_fmt = getattr(cell, "number_format", None)
    if num_fmt and num_fmt != "General":
        style["numberFormat"] = num_fmt
    return style


def _handle_list_workbooks(payload: Dict[str, Any]) -> Dict[str, Any]:
    directory = payload.get("directory")
    recursive = bool(payload.get("recursive", True))
    workbooks: List[Dict[str, Any]] = []
    seen: set[Path] = set()

    dirs_to_scan: List[Path]
    if directory:
        dirs_to_scan = [validate_directory_path(directory)]
    else:
        dirs_to_scan = [d for d in config.allowed_directories if d.exists()]

    for root in dirs_to_scan:
        pattern = "**/*" if recursive else "*"
        for path in sorted(root.glob(pattern)):
            if not path.is_file():
                continue
            if path.suffix.lower() not in config.allowed_extensions:
                continue
            if path.name.startswith("~$") or ".dsh_backup" in path.parts:
                continue
            resolved = path.resolve()
            if resolved in seen:
                continue
            seen.add(resolved)
            stat = resolved.stat()
            workbooks.append(
                {
                    "name": resolved.name,
                    "path": str(resolved),
                    "directory": str(resolved.parent),
                    "size_bytes": stat.st_size,
                    "modified": datetime.fromtimestamp(stat.st_mtime).isoformat(timespec="seconds"),
                    "has_vba": resolved.suffix.lower() == ".xlsm",
                    "is_locked": is_excel_lock_present(resolved),
                }
            )

    return {
        "workbooks": workbooks,
        "allowed_directories": [str(d) for d in config.allowed_directories],
    }


def _handle_create_workbook(payload: Dict[str, Any]) -> Dict[str, Any]:
    raw_path = payload.get("file_path")
    if not raw_path:
        directory = payload.get("directory")
        filename = str(payload.get("filename", "")).strip()
        if not filename:
            raise ValueError("A filename or file_path is required to create a workbook.")
        if not filename.lower().endswith(".xlsx"):
            filename = f"{filename}.xlsx"
        base_dir = (
            validate_directory_path(directory)
            if directory
            else config.allowed_directories[0]
        )
        raw_path = str(base_dir / filename)

    resolved = validate_file_path(raw_path, must_exist=False)
    overwrite = bool(payload.get("overwrite", False))
    if resolved.exists() and not overwrite:
        raise ValueError(f"Workbook '{resolved.name}' already exists.")

    resolved.parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()
    sheet_name = str(payload.get("sheet_name", "Sheet1")).strip() or "Sheet1"
    wb.active.title = sheet_name
    wb.save(resolved)
    wb.close()

    return _handle_load_sheet({"file_path": str(resolved), "sheet_name": sheet_name})


def _handle_load_sheet(payload: Dict[str, Any]) -> Dict[str, Any]:
    file_path = payload.get("file_path")
    if not file_path:
        raise ValueError("file_path is required.")
    resolved = validate_file_path(file_path, must_exist=True)
    keep_vba = resolved.suffix.lower() == ".xlsm"

    wb_formulas = openpyxl.load_workbook(resolved, keep_vba=keep_vba, data_only=False)
    try:
        wb_values = openpyxl.load_workbook(resolved, read_only=True, data_only=True)
    except Exception:
        wb_values = None

    try:
        sheet_names = list(wb_formulas.sheetnames)
        requested_sheet = payload.get("sheet_name")
        if requested_sheet and requested_sheet in sheet_names:
            ws = wb_formulas[requested_sheet]
        else:
            ws = wb_formulas.active

        ws_val = wb_values[ws.title] if (wb_values and ws.title in wb_values.sheetnames) else None

        start_row = max(1, int(payload.get("start_row", 1)))
        start_col = max(1, int(payload.get("start_col", 1)))

        min_rows = max(DEFAULT_GRID_ROWS, (ws.max_row or 0) + 10)
        min_cols = max(DEFAULT_GRID_COLS, (ws.max_column or 0) + 4)

        row_count = min(MAX_GRID_ROWS, max(1, int(payload.get("row_count", min_rows))))
        col_count = min(MAX_GRID_COLS, max(1, int(payload.get("col_count", min_cols))))

        end_row = start_row + row_count - 1
        end_col = start_col + col_count - 1

        # Build cached display values lookup when available
        cached_values: Dict[tuple[int, int], Any] = {}
        if ws_val is not None:
            for r_idx, row_vals in enumerate(
                ws_val.iter_rows(
                    min_row=start_row,
                    max_row=end_row,
                    min_col=start_col,
                    max_col=end_col,
                    values_only=True,
                ),
                start=start_row,
            ):
                for offset, val in enumerate(row_vals):
                    if val is not None:
                        cached_values[(r_idx, start_col + offset)] = val

        cells: Dict[str, Dict[str, Any]] = {}
        for r in range(start_row, end_row + 1):
            for c in range(start_col, end_col + 1):
                cell = ws.cell(row=r, column=c)
                raw = cell.value
                style = _extract_cell_style(cell)
                if raw is None and not style:
                    continue

                coord = f"{get_column_letter(c)}{r}"
                is_formula = isinstance(raw, str) and raw.startswith("=")
                display_raw = cached_values.get((r, c), raw) if is_formula else raw
                entry: Dict[str, Any] = {
                    "raw": _serialize_value(raw),
                    "display": _serialize_value(display_raw),
                }
                if is_formula:
                    entry["formula"] = raw
                if style:
                    entry["style"] = style
                cells[coord] = entry

        column_widths: Dict[str, float] = {}
        for c in range(start_col, end_col + 1):
            letter = get_column_letter(c)
            dim = ws.column_dimensions.get(letter)
            if dim and dim.width:
                column_widths[letter] = round(float(dim.width), 2)

        return {
            "file_path": str(resolved),
            "filename": resolved.name,
            "sheets": sheet_names,
            "active_sheet": ws.title,
            "has_vba": keep_vba,
            "is_locked": is_excel_lock_present(resolved),
            "max_row": ws.max_row or 0,
            "max_column": ws.max_column or 0,
            "start_row": start_row,
            "start_col": start_col,
            "row_count": row_count,
            "col_count": col_count,
            "column_widths": column_widths,
            "cells": cells,
        }
    finally:
        wb_formulas.close()
        if wb_values is not None:
            wb_values.close()


def _save_with_backup(wb: openpyxl.Workbook, resolved: Path) -> Optional[str]:
    backup_path = BackupManager.create_backup(resolved)
    wb.save(resolved)
    return str(backup_path) if backup_path else None


def _apply_style_patch(cell: Any, fmt: Dict[str, Any]) -> None:
    """Apply selective formatting attributes to an openpyxl cell while preserving others."""
    if any(k in fmt for k in ("bold", "italic", "font_color", "font_size")):
        existing = cell.font or Font()
        color = existing.color
        if "font_color" in fmt:
            raw_color = fmt["font_color"]
            if raw_color:
                hex_c = raw_color.lstrip("#").upper()
                color = f"FF{hex_c}" if len(hex_c) == 6 else hex_c
            else:
                color = None
        cell.font = Font(
            name=existing.name,
            size=fmt.get("font_size", existing.size),
            bold=fmt.get("bold", existing.bold),
            italic=fmt.get("italic", existing.italic),
            underline=existing.underline,
            color=color,
        )

    if "fill_color" in fmt:
        raw_fill = fmt["fill_color"]
        if raw_fill:
            hex_fill = raw_fill.lstrip("#").upper()
            argb_fill = f"FF{hex_fill}" if len(hex_fill) == 6 else hex_fill
            cell.fill = PatternFill(start_color=argb_fill, end_color=argb_fill, fill_type="solid")
        else:
            cell.fill = PatternFill(fill_type=None)

    if "align" in fmt:
        existing_align = cell.alignment or Alignment()
        cell.alignment = Alignment(
            horizontal=fmt["align"] or None,
            vertical=existing_align.vertical,
            wrap_text=existing_align.wrap_text,
        )

    if "number_format" in fmt and fmt["number_format"]:
        cell.number_format = str(fmt["number_format"])


def _handle_save_edits(payload: Dict[str, Any]) -> Dict[str, Any]:
    file_path = payload.get("file_path")
    sheet_name = payload.get("sheet_name")
    edits = payload.get("edits", [])
    formats = payload.get("formats", [])

    if not file_path:
        raise ValueError("file_path is required.")

    resolved = validate_file_path(file_path, must_exist=True)
    keep_vba = resolved.suffix.lower() == ".xlsm"
    wb = openpyxl.load_workbook(resolved, keep_vba=keep_vba, data_only=False)

    try:
        ws = wb[sheet_name] if (sheet_name and sheet_name in wb.sheetnames) else wb.active

        for edit in edits:
            coord = str(edit.get("cell", "")).strip().upper()
            if not coord:
                continue
            raw_val = edit.get("value", "")
            if raw_val == "" or raw_val is None:
                ws[coord].value = None
            elif isinstance(raw_val, str):
                ws[coord].value = _coerce_input_value(raw_val)
            else:
                ws[coord].value = raw_val

        for fmt in formats:
            target = str(fmt.get("cell") or fmt.get("range") or "").strip().upper()
            if not target:
                continue
            if ":" in target:
                for row in ws[target]:
                    for c in row:
                        _apply_style_patch(c, fmt)
            else:
                _apply_style_patch(ws[target], fmt)

        backup_path = _save_with_backup(wb, resolved)
        active_title = ws.title
    finally:
        wb.close()

    # Invalidate WorkbookManager cache if it was open in the same process
    WorkbookManager._open_workbooks.pop(resolved, None)

    result = _handle_load_sheet({"file_path": str(resolved), "sheet_name": active_title})
    result["backup_path"] = backup_path
    result["edits_applied"] = len(edits)
    result["formats_applied"] = len(formats)
    return result


def _coerce_input_value(raw: str) -> Any:
    """Convert plain text input from the editor into formula, int, float, bool, or str."""
    stripped = raw.strip()
    if not stripped:
        return None
    if stripped.startswith("="):
        return stripped
    low = stripped.lower()
    if low == "true":
        return True
    if low == "false":
        return False
    try:
        if "." not in stripped and "e" not in low:
            return int(stripped)
        return float(stripped)
    except ValueError:
        return raw


def _handle_sheet_op(payload: Dict[str, Any]) -> Dict[str, Any]:
    file_path = payload.get("file_path")
    op = str(payload.get("op", "")).strip()
    if not file_path or not op:
        raise ValueError("file_path and op are required.")

    resolved = validate_file_path(file_path, must_exist=True)
    keep_vba = resolved.suffix.lower() == ".xlsm"
    wb = openpyxl.load_workbook(resolved, keep_vba=keep_vba, data_only=False)

    try:
        sheet_name = payload.get("sheet_name") or wb.active.title
        target_sheet = sheet_name

        if op == "create_sheet":
            new_name = str(payload.get("new_name", "")).strip()
            if not new_name:
                raise ValueError("new_name is required for create_sheet.")
            if new_name in wb.sheetnames:
                raise ValueError(f"Worksheet '{new_name}' already exists.")
            wb.create_sheet(title=new_name)
            target_sheet = new_name

        elif op == "rename_sheet":
            new_name = str(payload.get("new_name", "")).strip()
            if not new_name:
                raise ValueError("new_name is required for rename_sheet.")
            if sheet_name not in wb.sheetnames:
                raise ValueError(f"Worksheet '{sheet_name}' not found.")
            if new_name in wb.sheetnames:
                raise ValueError(f"Worksheet '{new_name}' already exists.")
            wb[sheet_name].title = new_name
            target_sheet = new_name

        elif op == "delete_sheet":
            if sheet_name not in wb.sheetnames:
                raise ValueError(f"Worksheet '{sheet_name}' not found.")
            if len(wb.sheetnames) <= 1:
                raise ValueError("Cannot delete the only worksheet in a workbook.")
            del wb[sheet_name]
            target_sheet = wb.sheetnames[0]

        elif op == "insert_rows":
            ws = wb[sheet_name]
            idx = max(1, int(payload.get("index", 1)))
            amount = max(1, min(100, int(payload.get("amount", 1))))
            ws.insert_rows(idx, amount=amount)

        elif op == "delete_rows":
            ws = wb[sheet_name]
            idx = max(1, int(payload.get("index", 1)))
            amount = max(1, min(100, int(payload.get("amount", 1))))
            ws.delete_rows(idx, amount=amount)

        elif op == "insert_columns":
            ws = wb[sheet_name]
            idx = max(1, int(payload.get("index", 1)))
            amount = max(1, min(50, int(payload.get("amount", 1))))
            ws.insert_cols(idx, amount=amount)

        elif op == "delete_columns":
            ws = wb[sheet_name]
            idx = max(1, int(payload.get("index", 1)))
            amount = max(1, min(50, int(payload.get("amount", 1))))
            ws.delete_cols(idx, amount=amount)

        else:
            raise ValueError(f"Unsupported sheet operation: {op}")

        backup_path = _save_with_backup(wb, resolved)
    finally:
        wb.close()

    WorkbookManager._open_workbooks.pop(resolved, None)
    result = _handle_load_sheet({"file_path": str(resolved), "sheet_name": target_sheet})
    result["backup_path"] = backup_path
    return result


def _handle_summary(payload: Dict[str, Any]) -> Dict[str, Any]:
    file_path = payload.get("file_path")
    sheet_name = payload.get("sheet_name")
    range_str = payload.get("range_str")
    if not file_path or not sheet_name:
        raise ValueError("file_path and sheet_name are required.")
    resolved = validate_file_path(file_path, must_exist=True)
    if not range_str:
        wb = openpyxl.load_workbook(resolved, read_only=True, data_only=True)
        try:
            ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active
            max_r = max(1, ws.max_row or 1)
            max_c = max(1, ws.max_column or 1)
            range_str = f"A1:{get_column_letter(max_c)}{max_r}"
        finally:
            wb.close()
    return AnalysisManager.calculate_summary(
        file_path_str=str(resolved),
        range_str=str(range_str),
        sheet_name=str(sheet_name),
    )


HANDLERS = {
    "list_workbooks": _handle_list_workbooks,
    "create_workbook": _handle_create_workbook,
    "load_sheet": _handle_load_sheet,
    "save_edits": _handle_save_edits,
    "sheet_op": _handle_sheet_op,
    "summary": _handle_summary,
}


def dispatch(request: Dict[str, Any]) -> Dict[str, Any]:
    """Execute a single structured bridge request and return a JSON-serializable envelope."""
    action = str(request.get("action", "")).strip()
    handler = HANDLERS.get(action)
    if handler is None:
        return {"ok": False, "error": f"Unknown bridge action: {action}"}
    try:
        data = handler(request.get("payload") or {})
        return {"ok": True, "data": data}
    except SecurityError as exc:
        return {"ok": False, "code": "SECURITY_ERROR", "error": str(exc)}
    except Exception as exc:
        return {"ok": False, "code": "OPERATION_ERROR", "error": str(exc)}


def main() -> None:
    raw = sys.stdin.read()
    try:
        req = json.loads(raw) if raw.strip() else {}
    except json.JSONDecodeError as exc:
        sys.stdout.write(json.dumps({"ok": False, "code": "BAD_JSON", "error": str(exc)}))
        return
    response = dispatch(req)
    sys.stdout.write(json.dumps(response, default=str))


if __name__ == "__main__":
    main()
