"""
Security and validation unit tests.
Verifies path traversal defense, extension restrictions, automated backup snapshots,
and the confirmation/preview guard.
"""

from pathlib import Path
import pytest
from excel_mcp.security.validator import (
    validate_file_path,
    validate_directory_path,
    SecurityError,
)
from excel_mcp.security.backup import BackupManager
from excel_mcp.security.guard import ConfirmationGuard, ConfirmationRequiredError
from excel_mcp.utils.excel_helpers import (
    parse_cell_coordinate,
    format_cell_coordinate,
    parse_range_coordinates,
    split_sheet_and_range,
    is_excel_lock_present,
)


def test_validate_file_path_success(temp_workspace: Path, sample_workbook: Path):
    """Verifies that an authorized file path resolves successfully."""
    resolved = validate_file_path(sample_workbook)
    assert resolved == sample_workbook.resolve()
    assert resolved.exists()


def test_validate_file_path_traversal_blocked(temp_workspace: Path, tmp_path: Path):
    """Verifies that directory traversal outside allowed roots is blocked."""
    outside_dir = tmp_path / "forbidden"
    outside_dir.mkdir()
    outside_file = outside_dir / "secret.xlsx"
    outside_file.write_text("data")

    # Attempting to access outside_file via relative traversal
    traversal_path = temp_workspace / ".." / "forbidden" / "secret.xlsx"
    
    with pytest.raises(SecurityError) as exc_info:
        validate_file_path(traversal_path)
    assert "outside permitted directories" in str(exc_info.value)


def test_validate_file_path_invalid_extension(temp_workspace: Path):
    """Verifies that unauthorized file extensions are rejected."""
    bad_files = ["script.py", "sheet.csv", "macro.xlsm.exe", "data.json"]
    for bad in bad_files:
        p = temp_workspace / bad
        with pytest.raises(SecurityError) as exc_info:
            validate_file_path(p, must_exist=False)
        assert "Unsupported file format" in str(exc_info.value)


def test_validate_file_path_empty():
    """Verifies that empty paths raise ValueError."""
    with pytest.raises(ValueError):
        validate_file_path("")


def test_validate_file_path_nonexistent(temp_workspace: Path):
    """Verifies that nonexistent files raise FileNotFoundError when must_exist=True."""
    non_existent = temp_workspace / "missing.xlsx"
    with pytest.raises(FileNotFoundError):
        validate_file_path(non_existent, must_exist=True)
    
    # But when must_exist=False (e.g. for saving a new file), it succeeds
    valid_new = validate_file_path(non_existent, must_exist=False)
    assert valid_new == non_existent.resolve()


def test_validate_directory_path(temp_workspace: Path, tmp_path: Path):
    """Verifies directory sandboxing and traversal checks."""
    assert validate_directory_path(temp_workspace) == temp_workspace.resolve()

    outside = tmp_path / "outside_dir"
    outside.mkdir()
    with pytest.raises(SecurityError):
        validate_directory_path(outside)


def test_automated_backup_creation(sample_workbook: Path):
    """Verifies automatic snapshot creation in .dsh_backup directory."""
    backup_path = BackupManager.create_backup(sample_workbook)
    assert backup_path is not None
    assert backup_path.exists()
    assert backup_path.parent.name == ".dsh_backup"
    assert backup_path.stat().st_size == sample_workbook.stat().st_size

    # List backups
    backups = BackupManager.list_backups(sample_workbook)
    assert len(backups) >= 1
    assert backups[0] == backup_path

    # Modify the original file
    sample_workbook.write_bytes(b"corrupted")
    assert sample_workbook.read_bytes() == b"corrupted"

    # Restore from backup
    BackupManager.restore_backup(backup_path, sample_workbook)
    assert sample_workbook.stat().st_size == backup_path.stat().st_size


def test_confirmation_guard():
    """Verifies destructive operation guard and preview error structure."""
    impact = {"rows_to_delete": 5, "sheet": "Sheet1"}
    
    # When confirmed=False, should raise ConfirmationRequiredError with preview
    with pytest.raises(ConfirmationRequiredError) as exc_info:
        ConfirmationGuard.check_confirmation(
            action_name="delete_rows",
            confirmed=False,
            impact_details=impact,
        )
    err = exc_info.value
    err_dict = err.to_dict()
    assert err_dict["status"] == "confirmation_required"
    assert err_dict["action"] == "delete_rows"
    assert err_dict["impact_preview"] == impact

    # When confirmed=True, should pass silently
    ConfirmationGuard.check_confirmation(
        action_name="delete_rows",
        confirmed=True,
        impact_details=impact,
    )


def test_excel_lock_detection(sample_workbook: Path):
    """Verifies detection of Excel lock files (~$filename.xlsx)."""
    assert not is_excel_lock_present(sample_workbook)

    # Simulate Excel opening the file by creating the lock file
    lock_file = sample_workbook.parent / f"~${sample_workbook.name}"
    lock_file.write_text("lock")

    assert is_excel_lock_present(sample_workbook)
    lock_file.unlink()


def test_excel_coordinate_helpers():
    """Verifies coordinate conversion and range parsing."""
    # A1 -> (1, 1)
    assert parse_cell_coordinate("A1") == (1, 1)
    # B5 -> (5, 2)
    assert parse_cell_coordinate("B5") == (5, 2)
    assert parse_cell_coordinate("$C$10") == (10, 3)
    
    # (5, 2) -> "B5"
    assert format_cell_coordinate(5, 2) == "B5"
    assert format_cell_coordinate(1, 1) == "A1"

    # Split sheet and range
    assert split_sheet_and_range("Sheet1!A1:B10") == ("Sheet1", "A1:B10")
    assert split_sheet_and_range("'My Sheet'!C5") == ("My Sheet", "C5")
    assert split_sheet_and_range("A1:B10") == (None, "A1:B10")

    # Range coordinates
    assert parse_range_coordinates("A1:C10") == (1, 1, 10, 3)
    assert parse_range_coordinates("B5") == (5, 2, 5, 2)
    assert parse_range_coordinates("C10:A1") == (1, 1, 10, 3)
