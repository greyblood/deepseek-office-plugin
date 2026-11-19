"""
Unit tests for AdvancedManager operations (row/col shifts and charts).
"""

from pathlib import Path
import pytest
from excel_mcp.core.advanced import AdvancedManager
from excel_mcp.core.cells import CellManager
from excel_mcp.core.workbook import WorkbookManager
from excel_mcp.security.guard import ConfirmationRequiredError


def test_insert_rows(sample_workbook: Path):
    """Tests inserting blank rows."""
    res = AdvancedManager.insert_rows(sample_workbook, row_index=2, count=2)
    assert res["status"] == "success"
    assert res["rows_inserted"] == 2

    # Row 1 is still header, row 2 and 3 are empty, row 4 is Alice
    read_back = CellManager.read_range(sample_workbook, "A1:B4")
    assert read_back["values"][0] == ["ID", "Name"]
    assert read_back["values"][1] == [None, None]
    assert read_back["values"][2] == [None, None]
    assert read_back["values"][3] == [1, "Alice"]


def test_delete_rows_confirmation(sample_workbook: Path):
    """Tests that deleting rows requires confirmation."""
    with pytest.raises(ConfirmationRequiredError):
        AdvancedManager.delete_rows(sample_workbook, row_index=2, count=1, confirm=False)

    res = AdvancedManager.delete_rows(sample_workbook, row_index=2, count=1, confirm=True)
    assert res["status"] == "success"
    assert res["rows_deleted"] == 1


def test_insert_and_delete_columns(sample_workbook: Path):
    """Tests column insertion and confirmed deletion."""
    # Insert column at B (index 2)
    res_ins = AdvancedManager.insert_columns(sample_workbook, col_index="B", count=1)
    assert res_ins["status"] == "success"
    assert res_ins["column_letter"] == "B"

    # Header in A is ID, in B is None, in C is Name
    read_back = CellManager.read_range(sample_workbook, "A1:C1")
    assert read_back["values"][0] == ["ID", None, "Name"]

    # Delete column B (requires confirm)
    with pytest.raises(ConfirmationRequiredError):
        AdvancedManager.delete_columns(sample_workbook, col_index="B", count=1, confirm=False)

    res_del = AdvancedManager.delete_columns(sample_workbook, col_index="B", count=1, confirm=True)
    assert res_del["status"] == "success"

    # Restored
    read_back_after = CellManager.read_range(sample_workbook, "A1:B1")
    assert read_back_after["values"][0] == ["ID", "Name"]


def test_create_chart(sample_workbook: Path):
    """Tests creating and embedding a native openpyxl chart."""
    res = AdvancedManager.create_chart(
        sample_workbook,
        chart_type="bar",
        data_range="D1:D5",
        categories_range="B2:B5",
        title="Department Salaries",
        target_cell="F2",
    )
    assert res["status"] == "success"
    assert res["chart_type"] == "bar"
    assert res["target_cell"] == "F2"

    wb = WorkbookManager.get_or_open(sample_workbook)
    ws = wb.active
    assert len(ws._charts) == 1
    assert "Department Salaries" in str(ws._charts[0].title)
