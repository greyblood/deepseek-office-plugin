"""
Unit tests for WorkbookManager operations.
"""

from pathlib import Path
import pytest
from excel_mcp.core.workbook import WorkbookManager
from excel_mcp.security.guard import ConfirmationRequiredError


def test_open_and_get_info(sample_workbook: Path):
    """Tests opening a workbook and retrieving metadata."""
    info = WorkbookManager.get_workbook_info(sample_workbook)
    assert info["file_name"] == "sample.xlsx"
    assert info["active_sheet"] == "Sheet1"
    assert "Sheet1" in info["sheet_names"]
    assert info["total_sheets"] == 1
    assert not info["has_vba"]
    assert not info["is_locked"]


def test_list_workbooks(temp_workspace: Path, sample_workbook: Path):
    """Tests listing workbooks in workspace directory."""
    files = WorkbookManager.list_workbooks(temp_workspace)
    assert len(files) == 1
    assert files[0]["name"] == "sample.xlsx"
    assert files[0]["extension"] == ".xlsx"


def test_save_workbook_in_place(sample_workbook: Path):
    """Tests saving modifications to the same file."""
    wb = WorkbookManager.get_or_open(sample_workbook)
    ws = wb.active
    ws["A1"] = "Modified"

    res = WorkbookManager.save_workbook(sample_workbook)
    assert res["status"] == "success"
    assert res["backup_created"] is not None

    # Reload from disk and verify
    WorkbookManager.close_workbook(sample_workbook)
    fresh_wb = WorkbookManager.open_workbook(sample_workbook)
    assert fresh_wb.active["A1"].value == "Modified"
    fresh_wb.close()


def test_save_as_new_workbook(temp_workspace: Path, sample_workbook: Path):
    """Tests saving as a new file name."""
    new_path = temp_workspace / "new_copy.xlsx"
    res = WorkbookManager.save_workbook(sample_workbook, target_path_str=str(new_path))
    assert res["status"] == "success"
    assert new_path.exists()


def test_overwrite_requires_confirm(temp_workspace: Path, sample_workbook: Path):
    """Tests that overwriting an existing destination requires confirm=True."""
    other_path = temp_workspace / "other.xlsx"
    other_path.write_bytes(b"dummy")

    with pytest.raises(ConfirmationRequiredError):
        WorkbookManager.save_workbook(sample_workbook, target_path_str=str(other_path), overwrite=False)

    # With overwrite=True, succeeds
    res = WorkbookManager.save_workbook(sample_workbook, target_path_str=str(other_path), overwrite=True)
    assert res["status"] == "success"


def test_close_workbook(sample_workbook: Path):
    """Tests closing an active session workbook."""
    WorkbookManager.get_or_open(sample_workbook)
    assert sample_workbook in WorkbookManager._open_workbooks

    close_res = WorkbookManager.close_workbook(sample_workbook)
    assert close_res["status"] == "closed"
    assert sample_workbook not in WorkbookManager._open_workbooks


def test_corrupted_file_detection(temp_workspace: Path):
    """Tests detection of corrupt or invalid Excel archives."""
    corrupt_file = temp_workspace / "bad.xlsx"
    corrupt_file.write_bytes(b"Not a zip file")

    with pytest.raises(ValueError) as exc_info:
        WorkbookManager.open_workbook(corrupt_file)
    assert "not a valid or readable Excel archive" in str(exc_info.value)
