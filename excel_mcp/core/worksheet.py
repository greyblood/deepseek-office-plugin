"""
Worksheet management module.
Handles worksheet listing, creation, deletion with safety confirmation, and renaming.
"""

from pathlib import Path
from typing import Any, Dict, List, Optional
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.core.workbook import WorkbookManager
from excel_mcp.logging_config import logger
from excel_mcp.security.guard import ConfirmationGuard
from excel_mcp.security.validator import validate_file_path


class WorksheetManager:
    """Provides worksheet operations over an open or referenced Excel workbook."""

    @classmethod
    def get_sheet(
        cls,
        wb: openpyxl.Workbook,
        sheet_name: Optional[str] = None,
    ) -> Worksheet:
        """Retrieves a worksheet by name, or falls back to the active sheet."""
        if not sheet_name:
            if wb.active is None:
                raise ValueError("Workbook has no active worksheet.")
            return wb.active

        if sheet_name not in wb.sheetnames:
            available = ", ".join(f"'{s}'" for s in wb.sheetnames)
            raise ValueError(
                f"Worksheet '{sheet_name}' not found. Available sheets: {available}"
            )
        return wb[sheet_name]

    @classmethod
    def list_worksheets(cls, file_path_str: str | Path) -> List[Dict[str, Any]]:
        """Lists all worksheets and their properties."""
        resolved_path = validate_file_path(file_path_str, must_exist=True)
        wb = WorkbookManager.get_or_open(resolved_path)

        sheets = []
        for name in wb.sheetnames:
            ws = wb[name]
            sheets.append({
                "name": name,
                "is_active": ws == wb.active,
                "state": ws.sheet_state,
                "max_row": ws.max_row,
                "max_column": ws.max_column,
                "dimensions": ws.dimensions or "A1",
            })
        return sheets

    @classmethod
    def create_worksheet(
        cls,
        file_path_str: str | Path,
        sheet_name: str,
        index: Optional[int] = None,
    ) -> Dict[str, Any]:
        """
        Creates a new worksheet with the specified name and optional position.
        Persists changes and updates active session.
        """
        clean_name = sheet_name.strip()
        if not clean_name:
            raise ValueError("Worksheet name cannot be empty.")

        # Excel forbidden sheet characters
        invalid_chars = [":", "\\", "/", "?", "*", "[", "]"]
        for ch in invalid_chars:
            if ch in clean_name:
                raise ValueError(
                    f"Invalid character '{ch}' in worksheet name. Excel forbids: : \\ / ? * [ ]"
                )

        if len(clean_name) > 31:
            raise ValueError("Worksheet names in Excel cannot exceed 31 characters.")

        resolved_path = validate_file_path(file_path_str, must_exist=True)
        wb = WorkbookManager.get_or_open(resolved_path)

        if clean_name in wb.sheetnames:
            raise ValueError(f"Worksheet '{clean_name}' already exists in this workbook.")

        ws = wb.create_sheet(title=clean_name, index=index)
        WorkbookManager.save_workbook(resolved_path)

        logger.info("Created worksheet '%s' in %s", clean_name, resolved_path.name)
        return {
            "status": "success",
            "message": f"Worksheet '{clean_name}' created successfully.",
            "sheet_name": clean_name,
            "index": wb.sheetnames.index(clean_name),
            "total_sheets": len(wb.sheetnames),
        }

    @classmethod
    def delete_worksheet(
        cls,
        file_path_str: str | Path,
        sheet_name: str,
        confirm: bool = False,
    ) -> Dict[str, Any]:
        """
        Deletes a worksheet. Enforces confirmation guard since deletion is destructive.
        Prevents deleting the last remaining worksheet in a workbook.
        """
        resolved_path = validate_file_path(file_path_str, must_exist=True)
        wb = WorkbookManager.get_or_open(resolved_path)

        target_name = sheet_name.strip()
        if target_name not in wb.sheetnames:
            raise ValueError(f"Worksheet '{target_name}' does not exist.")

        if len(wb.sheetnames) <= 1:
            raise ValueError("Cannot delete the only remaining worksheet in the workbook.")

        ws = wb[target_name]
        has_content = ws.max_row > 1 or ws.max_column > 1 or ws.cell(1, 1).value is not None

        # Check confirmation
        ConfirmationGuard.check_confirmation(
            action_name="delete_worksheet",
            confirmed=confirm,
            impact_details={
                "sheet_name": target_name,
                "has_content": has_content,
                "max_row": ws.max_row,
                "max_column": ws.max_column,
                "file": resolved_path.name,
            },
            custom_message=(
                f"Deleting worksheet '{target_name}' will permanently remove all its data ({ws.max_row} rows). "
                f"Confirm by re-running with confirm=True."
            ),
        )

        wb.remove(ws)
        WorkbookManager.save_workbook(resolved_path)

        logger.info("Deleted worksheet '%s' from %s", target_name, resolved_path.name)
        return {
            "status": "success",
            "message": f"Worksheet '{target_name}' deleted successfully.",
            "remaining_sheets": wb.sheetnames,
        }

    @classmethod
    def rename_worksheet(
        cls,
        file_path_str: str | Path,
        old_name: str,
        new_name: str,
    ) -> Dict[str, Any]:
        """Renames an existing worksheet."""
        resolved_path = validate_file_path(file_path_str, must_exist=True)
        wb = WorkbookManager.get_or_open(resolved_path)

        src_name = old_name.strip()
        dst_name = new_name.strip()

        if not dst_name:
            raise ValueError("New worksheet name cannot be empty.")
        if dst_name == src_name:
            return {"status": "unchanged", "sheet_name": dst_name}

        if src_name not in wb.sheetnames:
            raise ValueError(f"Worksheet '{src_name}' does not exist.")
        if dst_name in wb.sheetnames:
            raise ValueError(f"Worksheet with name '{dst_name}' already exists.")

        ws = wb[src_name]
        ws.title = dst_name
        WorkbookManager.save_workbook(resolved_path)

        logger.info("Renamed worksheet '%s' -> '%s' in %s", src_name, dst_name, resolved_path.name)
        return {
            "status": "success",
            "message": f"Worksheet '{src_name}' renamed to '{dst_name}'.",
            "old_name": src_name,
            "new_name": dst_name,
            "sheet_names": wb.sheetnames,
        }
