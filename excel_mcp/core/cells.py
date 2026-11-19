"""
Cell and range operations engine.
Provides reading, writing, atomic updates, clearing, searching, and replacing
with type serialization, formula preservation, and memory guardrails.
"""

from pathlib import Path
from typing import Any, Dict, List, Optional
import openpyxl

from excel_mcp.config import config
from excel_mcp.core.workbook import WorkbookManager
from excel_mcp.core.worksheet import WorksheetManager
from excel_mcp.logging_config import logger
from excel_mcp.security.guard import ConfirmationGuard
from excel_mcp.security.validator import validate_file_path
from excel_mcp.utils.excel_helpers import (
    parse_cell_coordinate,
    format_cell_coordinate,
    parse_range_coordinates,
    split_sheet_and_range,
    to_serializable_value,
)


class CellManager:
    """Manages cell and range operations on Excel worksheets."""

    @classmethod
    def read_range(
        cls,
        file_path_str: str | Path,
        range_str: str,
        sheet_name: Optional[str] = None,
        data_only: bool = False,
    ) -> Dict[str, Any]:
        """
        Reads values from a specified range (e.g. 'A1:D20' or single cell 'B5').
        - data_only=True returns evaluated formula values (where cached by Excel).
        - data_only=False returns formula strings (e.g. '=SUM(A1:A10)').
        """
        resolved_path = validate_file_path(file_path_str, must_exist=True)
        embedded_sheet, clean_range = split_sheet_and_range(range_str)
        target_sheet_name = embedded_sheet or sheet_name

        wb = WorkbookManager.get_or_open(resolved_path, data_only=data_only)
        ws = WorksheetManager.get_sheet(wb, target_sheet_name)

        min_row, min_col, max_row, max_col = parse_range_coordinates(clean_range)

        # Enforce memory safety limits
        rows_requested = max_row - min_row + 1
        cols_requested = max_col - min_col + 1

        if rows_requested > config.max_rows_read:
            raise ValueError(
                f"Requested range spans {rows_requested} rows, which exceeds safety limit of {config.max_rows_read}."
            )
        if cols_requested > config.max_cols_read:
            raise ValueError(
                f"Requested range spans {cols_requested} columns, which exceeds safety limit of {config.max_cols_read}."
            )

        rows_data: List[List[Any]] = []
        for r in range(min_row, max_row + 1):
            row_vals = []
            for c in range(min_col, max_col + 1):
                val = ws.cell(row=r, column=c).value
                row_vals.append(to_serializable_value(val))
            rows_data.append(row_vals)

        effective_range = f"{format_cell_coordinate(min_row, min_col)}:{format_cell_coordinate(max_row, max_col)}"
        return {
            "sheet": ws.title,
            "range": effective_range,
            "rows_count": rows_requested,
            "columns_count": cols_requested,
            "values": rows_data,
        }

    @classmethod
    def write_range(
        cls,
        file_path_str: str | Path,
        start_cell: str,
        data: List[List[Any]],
        sheet_name: Optional[str] = None,
    ) -> Dict[str, Any]:
        """
        Writes a 2D array of values/formulas starting at start_cell (e.g. 'B2').
        Formulas starting with '=' are written directly and recognized by Excel.
        """
        if not data or not isinstance(data, list):
            raise ValueError("Data must be a non-empty 2D list of values.")

        resolved_path = validate_file_path(file_path_str, must_exist=True)
        embedded_sheet, clean_start = split_sheet_and_range(start_cell)
        target_sheet_name = embedded_sheet or sheet_name

        wb = WorkbookManager.get_or_open(resolved_path, data_only=False)
        ws = WorksheetManager.get_sheet(wb, target_sheet_name)

        start_row, start_col = parse_cell_coordinate(clean_start)
        max_col_offset = 0

        for r_idx, row_values in enumerate(data):
            if not isinstance(row_values, list):
                raise ValueError(f"Row {r_idx} must be a list of values.")
            for c_idx, val in enumerate(row_values):
                ws.cell(
                    row=start_row + r_idx,
                    column=start_col + c_idx,
                    value=val,
                )
            if len(row_values) > max_col_offset:
                max_col_offset = len(row_values)

        end_row = start_row + len(data) - 1
        end_col = start_col + max(0, max_col_offset - 1)
        written_range = f"{format_cell_coordinate(start_row, start_col)}:{format_cell_coordinate(end_row, end_col)}"

        WorkbookManager.save_workbook(resolved_path)

        logger.info("Wrote %s rows to %s (%s)", len(data), ws.title, written_range)
        return {
            "status": "success",
            "sheet": ws.title,
            "start_cell": clean_start,
            "written_range": written_range,
            "rows_written": len(data),
            "columns_written": max_col_offset,
        }

    @classmethod
    def update_cell(
        cls,
        file_path_str: str | Path,
        cell: str,
        value: Any,
        sheet_name: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Updates a single cell value or formula."""
        resolved_path = validate_file_path(file_path_str, must_exist=True)
        embedded_sheet, clean_cell = split_sheet_and_range(cell)
        target_sheet_name = embedded_sheet or sheet_name

        wb = WorkbookManager.get_or_open(resolved_path, data_only=False)
        ws = WorksheetManager.get_sheet(wb, target_sheet_name)

        row, col = parse_cell_coordinate(clean_cell)
        old_val = ws.cell(row=row, column=col).value
        ws.cell(row=row, column=col, value=value)

        WorkbookManager.save_workbook(resolved_path)

        logger.info("Updated cell %s in %s: %s -> %s", clean_cell, ws.title, old_val, value)
        return {
            "status": "success",
            "sheet": ws.title,
            "cell": clean_cell,
            "old_value": to_serializable_value(old_val),
            "new_value": to_serializable_value(value),
        }

    @classmethod
    def clear_range(
        cls,
        file_path_str: str | Path,
        range_str: str,
        sheet_name: Optional[str] = None,
        confirm: bool = False,
    ) -> Dict[str, Any]:
        """
        Clears values in the specified cell range.
        If more than 10 cells are affected, confirmation is required.
        """
        resolved_path = validate_file_path(file_path_str, must_exist=True)
        embedded_sheet, clean_range = split_sheet_and_range(range_str)
        target_sheet_name = embedded_sheet or sheet_name

        wb = WorkbookManager.get_or_open(resolved_path, data_only=False)
        ws = WorksheetManager.get_sheet(wb, target_sheet_name)

        min_row, min_col, max_row, max_col = parse_range_coordinates(clean_range)
        cells_count = (max_row - min_row + 1) * (max_col - min_col + 1)

        if cells_count > 10:
            ConfirmationGuard.check_confirmation(
                action_name="clear_range",
                confirmed=confirm,
                impact_details={
                    "sheet": ws.title,
                    "range": clean_range,
                    "cells_affected": cells_count,
                    "file": resolved_path.name,
                },
                custom_message=f"Clearing range '{clean_range}' will erase {cells_count} cells. Confirm with confirm=True.",
            )

        cleared_count = 0
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                cell = ws.cell(row=r, column=c)
                if cell.value is not None:
                    cell.value = None
                    cleared_count += 1

        WorkbookManager.save_workbook(resolved_path)

        return {
            "status": "success",
            "sheet": ws.title,
            "range": clean_range,
            "cells_cleared": cleared_count,
        }

    @classmethod
    def find_value(
        cls,
        file_path_str: str | Path,
        query: str,
        sheet_name: Optional[str] = None,
        match_case: bool = False,
        exact_match: bool = False,
        max_results: int = 100,
    ) -> List[Dict[str, Any]]:
        """Searches cells across the worksheet for matching values."""
        resolved_path = validate_file_path(file_path_str, must_exist=True)
        wb = WorkbookManager.get_or_open(resolved_path, data_only=False)
        ws = WorksheetManager.get_sheet(wb, sheet_name)

        search_term = query if match_case else query.lower()
        matches: List[Dict[str, Any]] = []

        for row in ws.iter_rows():
            for cell in row:
                if cell.value is None:
                    continue
                cell_str = str(cell.value)
                target_str = cell_str if match_case else cell_str.lower()

                is_match = (target_str == search_term) if exact_match else (search_term in target_str)
                if is_match:
                    matches.append({
                        "sheet": ws.title,
                        "coordinate": cell.coordinate,
                        "row": cell.row,
                        "column": cell.column,
                        "value": to_serializable_value(cell.value),
                    })
                    if len(matches) >= max_results:
                        return matches

        return matches

    @classmethod
    def replace_value(
        cls,
        file_path_str: str | Path,
        old_value: Any,
        new_value: Any,
        sheet_name: Optional[str] = None,
        range_str: Optional[str] = None,
        match_case: bool = False,
        exact_match: bool = True,
        max_replacements: int = 1000,
    ) -> Dict[str, Any]:
        """Replaces occurrences of old_value with new_value in sheet or bounded range."""
        resolved_path = validate_file_path(file_path_str, must_exist=True)
        wb = WorkbookManager.get_or_open(resolved_path, data_only=False)
        ws = WorksheetManager.get_sheet(wb, sheet_name)

        old_str = str(old_value) if match_case else str(old_value).lower()
        replacements_count = 0

        if range_str:
            min_row, min_col, max_row, max_col = parse_range_coordinates(range_str)
            row_iter = ws.iter_rows(
                min_row=min_row, min_col=min_col, max_row=max_row, max_col=max_col
            )
        else:
            row_iter = ws.iter_rows()

        for row in row_iter:
            for cell in row:
                if cell.value is None:
                    continue
                cell_str = str(cell.value)
                target_str = cell_str if match_case else cell_str.lower()

                matched = (target_str == old_str) if exact_match else (old_str in target_str)
                if matched:
                    if exact_match:
                        cell.value = new_value
                    else:
                        # Substring replacement
                        if match_case:
                            cell.value = cell_str.replace(str(old_value), str(new_value))
                        else:
                            # Case-insensitive substring replacement
                            pattern = re.compile(re.escape(str(old_value)), re.IGNORECASE)
                            cell.value = pattern.sub(str(new_value), cell_str)

                    replacements_count += 1
                    if replacements_count >= max_replacements:
                        break
            if replacements_count >= max_replacements:
                break

        if replacements_count > 0:
            WorkbookManager.save_workbook(resolved_path)

        return {
            "status": "success",
            "sheet": ws.title,
            "replacements_count": replacements_count,
            "range": range_str or "entire_sheet",
        }
