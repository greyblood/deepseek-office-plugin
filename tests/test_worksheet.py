"""
Unit tests for WorksheetManager operations.
"""

from pathlib import Path
import pytest
from excel_mcp.core.worksheet import WorksheetManager
from excel_mcp.security.guard import ConfirmationRequiredError


def test_list_worksheets(sample_workbook: Path):
    """Tests listing all sheets in a workbook."""
    sheets = WorksheetManager.list_worksheets(sample_workbook)
    assert len(sheets) == 1
    assert sheets[0]["name"] == "Sheet1"
    assert sheets[0]["is_active"] is True


def test_create_worksheet(sample_workbook: Path):
    """Tests creating a new worksheet."""
    res = WorksheetManager.create_worksheet(sample_workbook, "Analytics")
    assert res["status"] == "success"
    assert res["sheet_name"] == "Analytics"
    assert res["total_sheets"] == 2

    sheets = WorksheetManager.list_worksheets(sample_workbook)
    assert len(sheets) == 2
    assert any(s["name"] == "Analytics" for s in sheets)


def test_create_worksheet_invalid_name(sample_workbook: Path):
    """Tests rejection of invalid characters in sheet names."""
    with pytest.raises(ValueError) as exc:
        WorksheetManager.create_worksheet(sample_workbook, "Bad/Name")
    assert "Invalid character" in str(exc.value)

    with pytest.raises(ValueError) as exc:
        WorksheetManager.create_worksheet(sample_workbook, "A" * 32)
    assert "cannot exceed 31 characters" in str(exc.value)


def test_create_worksheet_duplicate(sample_workbook: Path):
    """Tests rejection of duplicate worksheet names."""
    with pytest.raises(ValueError) as exc:
        WorksheetManager.create_worksheet(sample_workbook, "Sheet1")
    assert "already exists" in str(exc.value)


def test_delete_worksheet_requires_confirm(sample_workbook: Path):
    """Tests that deleting a worksheet requires confirmation."""
    WorksheetManager.create_worksheet(sample_workbook, "ToDelete")
    
    with pytest.raises(ConfirmationRequiredError) as exc_info:
        WorksheetManager.delete_worksheet(sample_workbook, "ToDelete", confirm=False)
    
    preview = exc_info.value.to_dict()
    assert preview["status"] == "confirmation_required"
    assert preview["impact_preview"]["sheet_name"] == "ToDelete"


def test_delete_worksheet_confirmed(sample_workbook: Path):
    """Tests deleting a worksheet when confirm=True."""
    WorksheetManager.create_worksheet(sample_workbook, "TempSheet")
    res = WorksheetManager.delete_worksheet(sample_workbook, "TempSheet", confirm=True)
    assert res["status"] == "success"
    assert "TempSheet" not in res["remaining_sheets"]


def test_cannot_delete_only_sheet(sample_workbook: Path):
    """Tests that deleting the only remaining sheet is prohibited."""
    with pytest.raises(ValueError) as exc:
        WorksheetManager.delete_worksheet(sample_workbook, "Sheet1", confirm=True)
    assert "Cannot delete the only remaining worksheet" in str(exc.value)


def test_rename_worksheet(sample_workbook: Path):
    """Tests renaming an existing worksheet."""
    res = WorksheetManager.rename_worksheet(sample_workbook, "Sheet1", "Overview")
    assert res["status"] == "success"
    assert res["new_name"] == "Overview"
    assert "Overview" in res["sheet_names"]
    assert "Sheet1" not in res["sheet_names"]


def test_rename_worksheet_conflict(sample_workbook: Path):
    """Tests that renaming to an existing sheet name raises ValueError."""
    WorksheetManager.create_worksheet(sample_workbook, "Second")
    with pytest.raises(ValueError) as exc:
        WorksheetManager.rename_worksheet(sample_workbook, "Sheet1", "Second")
    assert "already exists" in str(exc.value)
