"""
Unit tests for FormattingManager operations.
"""

from pathlib import Path
import pytest
from excel_mcp.core.formatting import FormattingManager
from excel_mcp.core.workbook import WorkbookManager


def test_format_cells_font_and_fill(sample_workbook: Path):
    """Tests applying font styling and background fill."""
    res = FormattingManager.format_cells(
        sample_workbook,
        range_str="A1:D1",
        font_style={"bold": True, "size": 12, "color": "FFFFFF"},
        fill_style={"bg_color": "1F4E78"},
        alignment_style={"horizontal": "center"},
    )
    assert res["status"] == "success"
    assert res["cells_modified"] == 4

    # Verify directly on openpyxl cell
    wb = WorkbookManager.get_or_open(sample_workbook)
    ws = wb.active
    assert ws["A1"].font.bold is True
    assert ws["A1"].font.size == 12
    assert ws["A1"].fill.start_color.rgb == "001F4E78" or "1F4E78" in str(ws["A1"].fill.start_color.rgb)
    assert ws["A1"].alignment.horizontal == "center"


def test_apply_number_format(sample_workbook: Path):
    """Tests applying number formatting presets."""
    res = FormattingManager.apply_number_format(
        sample_workbook,
        range_str="D2:D5",
        format_code="currency",
    )
    assert res["status"] == "success"
    assert res["cells_formatted"] == 4
    assert res["format_applied"] == "$#,##0.00"

    wb = WorkbookManager.get_or_open(sample_workbook)
    assert wb.active["D2"].number_format == "$#,##0.00"


def test_set_column_width_explicit_and_auto(sample_workbook: Path):
    """Tests setting explicit and auto-calculated column widths."""
    # Explicit width
    res_exp = FormattingManager.set_column_width(sample_workbook, column="B", width=25.0)
    assert res_exp["width"] == 25.0

    wb = WorkbookManager.get_or_open(sample_workbook)
    assert wb.active.column_dimensions["B"].width == 25.0

    # Auto-fit width
    res_auto = FormattingManager.set_column_width(sample_workbook, column="C", auto_fit=True)
    assert res_auto["auto_fit"] is True
    assert res_auto["width"] >= 10.0


def test_set_row_height(sample_workbook: Path):
    """Tests setting row height."""
    res = FormattingManager.set_row_height(sample_workbook, row=1, height=30.0)
    assert res["status"] == "success"
    assert res["height"] == 30.0

    wb = WorkbookManager.get_or_open(sample_workbook)
    assert wb.active.row_dimensions[1].height == 30.0
