"""
Workbook lifecycle and session manager.
Handles opening, inspecting, saving, closing, and listing Excel workbooks with
VBA preservation, file-lock awareness, and automated backup creation.
"""

from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional
import zipfile
import openpyxl
from openpyxl.workbook import Workbook

from excel_mcp.config import config
from excel_mcp.logging_config import logger
from excel_mcp.security.backup import BackupManager
from excel_mcp.security.guard import ConfirmationGuard
from excel_mcp.security.validator import (
    validate_file_path,
    validate_directory_path,
    SecurityError,
)
from excel_mcp.utils.excel_helpers import is_excel_lock_present


class WorkbookManager:
    """Manages active Excel workbook instances and disk operations."""

    # In-memory cache for open workbooks: Path -> openpyxl.Workbook
    _open_workbooks: Dict[Path, Workbook] = {}

    @classmethod
    def get_or_open(
        cls,
        file_path_str: str | Path,
        read_only: bool = False,
        data_only: bool = False,
    ) -> Workbook:
        """
        Retrieves a workbook from the session cache or opens it from disk.
        - data_only=True evaluates formula results where cached by Excel.
        - data_only=False retains the actual formula strings (default).
        """
        resolved_path = validate_file_path(file_path_str, must_exist=True)

        if resolved_path in cls._open_workbooks:
            return cls._open_workbooks[resolved_path]

        return cls.open_workbook(
            resolved_path,
            read_only=read_only,
            data_only=data_only,
        )

    @classmethod
    def open_workbook(
        cls,
        file_path_str: str | Path,
        read_only: bool = False,
        data_only: bool = False,
    ) -> Workbook:
        """
        Opens an Excel workbook (.xlsx or .xlsm) with VBA preservation.
        Caches the opened instance in memory.
        """
        resolved_path = validate_file_path(file_path_str, must_exist=True)

        if is_excel_lock_present(resolved_path):
            logger.warning(
                "Workbook '%s' is currently open in Microsoft Excel (lock file detected).",
                resolved_path.name,
            )

        keep_vba = resolved_path.suffix.lower() == ".xlsm"

        try:
            wb = openpyxl.load_workbook(
                resolved_path,
                read_only=read_only,
                keep_vba=keep_vba,
                data_only=data_only,
            )
            cls._open_workbooks[resolved_path] = wb
            logger.info("Opened workbook: %s (VBA: %s)", resolved_path.name, keep_vba)
            return wb
        except zipfile.BadZipFile as e:
            raise ValueError(
                f"File '{resolved_path.name}' is not a valid or readable Excel archive (corrupted)."
            ) from e
        except Exception as e:
            logger.error("Failed to open workbook %s: %s", resolved_path, e)
            raise IOError(f"Could not open workbook '{resolved_path.name}': {e}") from e

    @classmethod
    def save_workbook(
        cls,
        file_path_str: str | Path,
        target_path_str: Optional[str | Path] = None,
        overwrite: bool = False,
    ) -> Dict[str, Any]:
        """
        Saves changes to an open or specified workbook.
        - If target_path is specified, saves as a new workbook.
        - Checks for file overwrite confirmation.
        - Creates an automatic backup before overwriting existing files.
        """
        source_path = validate_file_path(file_path_str, must_exist=True)
        wb = cls.get_or_open(source_path)

        if target_path_str:
            dest_path = validate_file_path(target_path_str, must_exist=False)
        else:
            dest_path = source_path

        # If saving over an existing file different from itself or overwriting
        if dest_path.exists() and dest_path != source_path and not overwrite:
            ConfirmationGuard.check_confirmation(
                action_name="overwrite_workbook",
                confirmed=overwrite,
                impact_details={
                    "target_file": str(dest_path),
                    "size_bytes": dest_path.stat().st_size,
                },
                custom_message=f"Target file '{dest_path.name}' already exists. Confirm overwrite to replace.",
            )

        # Create defensive backup before write
        backup_path = BackupManager.create_backup(dest_path)

        try:
            wb.save(dest_path)
            logger.info("Saved workbook to %s", dest_path.name)
            
            # If saved as a new file, update cache entry
            if dest_path != source_path:
                cls._open_workbooks[dest_path] = wb

            return {
                "status": "success",
                "message": f"Workbook saved successfully to '{dest_path.name}'",
                "path": str(dest_path),
                "backup_created": str(backup_path) if backup_path else None,
            }
        except PermissionError as e:
            raise PermissionError(
                f"Cannot save '{dest_path.name}'. The file is open and locked by another application."
            ) from e
        except Exception as e:
            logger.error("Failed to save workbook %s: %s", dest_path, e)
            raise IOError(f"Failed to save workbook: {e}") from e

    @classmethod
    def close_workbook(
        cls,
        file_path_str: str | Path,
        save_changes: bool = False,
    ) -> Dict[str, Any]:
        """Closes an active workbook and removes it from the session cache."""
        resolved_path = validate_file_path(file_path_str, must_exist=True)

        if resolved_path not in cls._open_workbooks:
            return {
                "status": "not_open",
                "message": f"Workbook '{resolved_path.name}' was not active in session.",
                "path": str(resolved_path),
            }

        wb = cls._open_workbooks[resolved_path]
        save_result = None

        if save_changes:
            save_result = cls.save_workbook(resolved_path)

        try:
            wb.close()
        finally:
            del cls._open_workbooks[resolved_path]

        logger.info("Closed workbook: %s", resolved_path.name)
        return {
            "status": "closed",
            "message": f"Workbook '{resolved_path.name}' closed successfully.",
            "path": str(resolved_path),
            "saved": save_changes,
            "save_details": save_result,
        }

    @classmethod
    def get_workbook_info(cls, file_path_str: str | Path) -> Dict[str, Any]:
        """Inspects workbook structure, metadata, sheets, dimensions, and file attributes."""
        resolved_path = validate_file_path(file_path_str, must_exist=True)
        wb = cls.get_or_open(resolved_path)

        stat = resolved_path.stat()
        sheets_info = []

        for name in wb.sheetnames:
            ws = wb[name]
            sheets_info.append({
                "name": name,
                "is_active": ws == wb.active,
                "max_row": ws.max_row,
                "max_column": ws.max_column,
                "dimensions": ws.dimensions or "A1",
                "state": ws.sheet_state,  # visible, hidden, veryHidden
            })

        return {
            "file_name": resolved_path.name,
            "file_path": str(resolved_path),
            "file_size_bytes": stat.st_size,
            "last_modified": datetime.fromtimestamp(stat.st_mtime).isoformat(),
            "active_sheet": wb.active.title if wb.active else None,
            "total_sheets": len(wb.sheetnames),
            "sheet_names": wb.sheetnames,
            "sheets": sheets_info,
            "has_vba": resolved_path.suffix.lower() == ".xlsm",
            "is_locked": is_excel_lock_present(resolved_path),
        }

    @classmethod
    def list_workbooks(cls, directory_str: Optional[str | Path] = None) -> List[Dict[str, Any]]:
        """
        Lists all accessible Excel files in the specified directory or configured allowed roots.
        """
        if directory_str:
            target_dir = validate_directory_path(directory_str, must_exist=True)
            search_dirs = [target_dir]
        else:
            search_dirs = config.allowed_directories

        results = []
        for s_dir in search_dirs:
            if not s_dir.exists():
                continue
            for ext in config.allowed_extensions:
                for file_path in s_dir.glob(f"*{ext}"):
                    # Ignore temporary/hidden Excel lock files starting with ~$
                    if file_path.name.startswith("~$"):
                        continue
                    try:
                        stat = file_path.stat()
                        results.append({
                            "name": file_path.name,
                            "path": str(file_path.resolve()),
                            "size_bytes": stat.st_size,
                            "last_modified": datetime.fromtimestamp(stat.st_mtime).isoformat(),
                            "extension": file_path.suffix.lower(),
                            "is_locked": is_excel_lock_present(file_path),
                        })
                    except Exception as e:
                        logger.warning("Could not stat %s: %s", file_path, e)

        return sorted(results, key=lambda x: x["name"])
