"""
Unit tests for CellManager operations.
"""

from pathlib import Path
import pytest
from excel_mcp.core.cells import CellManager
from excel_mcp.security.guard import ConfirmationRequiredError


def test_read_range(sample_workbook: Path):
    """Tests reading a multi-row rectangular range."""
    res = CellManager.read_range(sample_workbook, "A1:D3")
    assert res["sheet"] == "Sheet1"
    assert res["rows_count"] == 3
    assert res["columns_count"] == 4
    assert res["values"][0] == ["ID", "Name", "Department", "Salary"]
    assert res["values"][1] == [1, "Alice", "Engineering", 95000]
    assert res["values"][2] == [2, "Bob", "Design", 80000]


def test_read_range_embedded_sheet(sample_workbook: Path):
    """Tests reading with sheet prefix in range string."""
    res = CellManager.read_range(sample_workbook, "Sheet1!B2:B3")
    assert res["values"] == [["Alice"], ["Bob"]]


def test_write_range(sample_workbook: Path):
    """Tests writing a 2D matrix into the sheet."""
    new_data = [
        [5, "Eve", "Marketing", 88000],
        [6, "Frank", "Sales", 75000],
    ]
    res = CellManager.write_range(sample_workbook, start_cell="A6", data=new_data)
    assert res["status"] == "success"
    assert res["rows_written"] == 2
    assert res["written_range"] == "A6:D7"

    # Verify via read
    read_back = CellManager.read_range(sample_workbook, "A6:D7")
    assert read_back["values"] == new_data


def test_update_cell(sample_workbook: Path):
    """Tests atomic single cell update."""
    res = CellManager.update_cell(sample_workbook, cell="D2", value=99000)
    assert res["status"] == "success"
    assert res["old_value"] == 95000
    assert res["new_value"] == 99000

    read_back = CellManager.read_range(sample_workbook, "D2")
    assert read_back["values"] == [[99000]]


def test_clear_range_confirmation(sample_workbook: Path):
    """Tests clearing large range triggers confirmation check."""
    # Write 15 cells
    large_data = [[i] * 5 for i in range(3)]  # 15 cells
    CellManager.write_range(sample_workbook, "A10", large_data)

    with pytest.raises(ConfirmationRequiredError):
        CellManager.clear_range(sample_workbook, "A10:E12", confirm=False)

    # Confirmed clear
    res = CellManager.clear_range(sample_workbook, "A10:E12", confirm=True)
    assert res["status"] == "success"
    assert res["cells_cleared"] == 15


def test_find_value(sample_workbook: Path):
    """Tests searching cells for text values."""
    matches = CellManager.find_value(sample_workbook, query="Engineering")
    assert len(matches) == 2
    assert matches[0]["coordinate"] == "C2"  # Alice
    assert matches[1]["coordinate"] == "C5"  # Diana


def test_replace_value(sample_workbook: Path):
    """Tests replacing values in sheet."""
    res = CellManager.replace_value(
        sample_workbook,
        old_value="Engineering",
        new_value="R&D",
        match_case=True,
    )
    assert res["status"] == "success"
    assert res["replacements_count"] == 2

    # Verify replacement took effect
    matches = CellManager.find_value(sample_workbook, query="R&D")
    assert len(matches) == 2
