"""
Formatting and layout operations engine.
Supports typography, fills, alignments, borders, number formatting presets,
and dynamic column auto-fitting.
"""

from pathlib import Path
from typing import Any, Dict, Optional
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.cell import get_column_letter, column_index_from_string

from excel_mcp.core.workbook import WorkbookManager
from excel_mcp.core.worksheet import WorksheetManager
from excel_mcp.logging_config import logger
from excel_mcp.security.validator import validate_file_path
from excel_mcp.utils.excel_helpers import (
    parse_range_coordinates,
    split_sheet_and_range,
)


class FormattingManager:
    """Provides visual formatting, number styling, and dimension adjustments."""

    # Pre-configured number format presets
    NUMBER_FORMAT_PRESETS = {
        "currency": "$#,##0.00",
        "currency_eur": "€#,##0.00",
        "currency_gbp": "£#,##0.00",
        "percentage": "0.00%",
        "percentage_int": "0%",
        "integer": "#,##0",
        "decimal": "#,##0.00",
        "date": "YYYY-MM-DD",
        "datetime": "YYYY-MM-DD HH:MM:SS",
        "scientific": "0.00E+00",
        "text": "@",
    }

    @classmethod
    def format_cells(
        cls,
        file_path_str: str | Path,
        range_str: str,
        sheet_name: Optional[str] = None,
        font_style: Optional[Dict[str, Any]] = None,
        fill_style: Optional[Dict[str, Any]] = None,
        alignment_style: Optional[Dict[str, Any]] = None,
        border_style: Optional[Dict[str, Any]] = None,
    ) -> Dict[str, Any]:
        """
        Applies styling (font, fill, alignment, borders) to a range of cells.
        """
        resolved_path = validate_file_path(file_path_str, must_exist=True)
        embedded_sheet, clean_range = split_sheet_and_range(range_str)
        target_sheet_name = embedded_sheet or sheet_name

        wb = WorkbookManager.get_or_open(resolved_path, data_only=False)
        ws = WorksheetManager.get_sheet(wb, target_sheet_name)

        min_row, min_col, max_row, max_col = parse_range_coordinates(clean_range)

        # Build Font object
        openpyxl_font = None
        if font_style:
            openpyxl_font = Font(
                name=font_style.get("name"),
                size=font_style.get("size"),
                bold=font_style.get("bold"),
                italic=font_style.get("italic"),
                underline="single" if font_style.get("underline") else None,
                color=font_style.get("color", "").replace("#", "") or None,
            )

        # Build PatternFill object
        openpyxl_fill = None
        if fill_style and fill_style.get("bg_color"):
            color_hex = fill_style["bg_color"].replace("#", "")
            fill_type = fill_style.get("fill_type", "solid")
            openpyxl_fill = PatternFill(start_color=color_hex, end_color=color_hex, fill_type=fill_type)

        # Build Alignment object
        openpyxl_align = None
        if alignment_style:
            openpyxl_align = Alignment(
                horizontal=alignment_style.get("horizontal"),
                vertical=alignment_style.get("vertical"),
                wrap_text=alignment_style.get("wrap_text"),
            )

        # Build Border object
        openpyxl_border = None
        if border_style:
            b_style = border_style.get("style", "thin")
            b_color = border_style.get("color", "000000").replace("#", "")
            side = Side(style=b_style, color=b_color)
            openpyxl_border = Border(left=side, right=side, top=side, bottom=side)

        cells_modified = 0
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                cell = ws.cell(row=r, column=c)
                if openpyxl_font:
                    cell.font = openpyxl_font
                if openpyxl_fill:
                    cell.fill = openpyxl_fill
                if openpyxl_align:
                    cell.alignment = openpyxl_align
                if openpyxl_border:
                    cell.border = openpyxl_border
                cells_modified += 1

        WorkbookManager.save_workbook(resolved_path)

        logger.info("Formatted %s cells in range %s (%s)", cells_modified, clean_range, ws.title)
        return {
            "status": "success",
            "sheet": ws.title,
            "range": clean_range,
            "cells_modified": cells_modified,
        }

    @classmethod
    def apply_number_format(
        cls,
        file_path_str: str | Path,
        range_str: str,
        format_code: str,
        sheet_name: Optional[str] = None,
    ) -> Dict[str, Any]:
        """
        Applies a number/currency/date format string or preset name to a range.
        Presets: 'currency', 'percentage', 'integer', 'decimal', 'date', 'datetime'.
        """
        resolved_path = validate_file_path(file_path_str, must_exist=True)
        embedded_sheet, clean_range = split_sheet_and_range(range_str)
        target_sheet_name = embedded_sheet or sheet_name

        wb = WorkbookManager.get_or_open(resolved_path, data_only=False)
        ws = WorksheetManager.get_sheet(wb, target_sheet_name)

        min_row, min_col, max_row, max_col = parse_range_coordinates(clean_range)

        # Check preset dictionary or use raw code
        actual_format = cls.NUMBER_FORMAT_PRESETS.get(format_code.lower(), format_code)

        count = 0
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                ws.cell(row=r, column=c).number_format = actual_format
                count += 1

        WorkbookManager.save_workbook(resolved_path)

        return {
            "status": "success",
            "sheet": ws.title,
            "range": clean_range,
            "format_applied": actual_format,
            "cells_formatted": count,
        }

    @classmethod
    def set_column_width(
        cls,
        file_path_str: str | Path,
        column: str | int,
        width: Optional[float] = None,
        auto_fit: bool = False,
        sheet_name: Optional[str] = None,
    ) -> Dict[str, Any]:
        """
        Sets explicit column width or auto-fits based on cell contents.
        """
        resolved_path = validate_file_path(file_path_str, must_exist=True)
        wb = WorkbookManager.get_or_open(resolved_path, data_only=False)
        ws = WorksheetManager.get_sheet(wb, sheet_name)

        if isinstance(column, int):
            col_letter = get_column_letter(column)
        else:
            col_letter = column.strip().upper()

        if auto_fit:
            col_idx = column_index_from_string(col_letter)
            max_len = 0
            for r in range(1, ws.max_row + 1):
                val = ws.cell(row=r, column=col_idx).value
                if val is not None:
                    max_len = max(max_len, len(str(val)))
            effective_width = max(max_len + 3, 10.0)
        else:
            if width is None or width <= 0:
                raise ValueError("Width must be a positive number when auto_fit is False.")
            effective_width = float(width)

        ws.column_dimensions[col_letter].width = effective_width
        WorkbookManager.save_workbook(resolved_path)

        logger.info("Set column %s width to %.2f in %s", col_letter, effective_width, ws.title)
        return {
            "status": "success",
            "sheet": ws.title,
            "column": col_letter,
            "width": effective_width,
            "auto_fit": auto_fit,
        }

    @classmethod
    def set_row_height(
        cls,
        file_path_str: str | Path,
        row: int,
        height: float,
        sheet_name: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Sets the height in points for a specific row."""
        if row < 1 or height <= 0:
            raise ValueError(f"Invalid row or height: row={row}, height={height}")

        resolved_path = validate_file_path(file_path_str, must_exist=True)
        wb = WorkbookManager.get_or_open(resolved_path, data_only=False)
        ws = WorksheetManager.get_sheet(wb, sheet_name)

        ws.row_dimensions[row].height = float(height)
        WorkbookManager.save_workbook(resolved_path)

        logger.info("Set row %s height to %.2f in %s", row, height, ws.title)
        return {
            "status": "success",
            "sheet": ws.title,
            "row": row,
            "height": float(height),
        }
