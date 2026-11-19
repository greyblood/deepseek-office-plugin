"""
Advanced structural modifications and chart generation engine.
Supports shifting rows/columns with safety confirmations and generating
native openpyxl charts (Bar, Line, Pie, Area).
"""

from pathlib import Path
from typing import Any, Dict, Optional, Union
import openpyxl
from openpyxl.chart import BarChart, LineChart, PieChart, AreaChart, Reference
from openpyxl.utils.cell import column_index_from_string, get_column_letter

from excel_mcp.core.workbook import WorkbookManager
from excel_mcp.core.worksheet import WorksheetManager
from excel_mcp.logging_config import logger
from excel_mcp.security.guard import ConfirmationGuard
from excel_mcp.security.validator import validate_file_path
from excel_mcp.utils.excel_helpers import (
    parse_cell_coordinate,
    parse_range_coordinates,
    split_sheet_and_range,
)


class AdvancedManager:
    """Manages structural mutations (rows/columns) and chart generation."""

    CHART_CLASSES = {
        "bar": BarChart,
        "col": BarChart,
        "column": BarChart,
        "line": LineChart,
        "pie": PieChart,
        "area": AreaChart,
    }

    @classmethod
    def insert_rows(
        cls,
        file_path_str: str | Path,
        row_index: int,
        count: int = 1,
        sheet_name: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Inserts count empty rows before row_index."""
        if row_index < 1 or count < 1:
            raise ValueError(f"Invalid row_index ({row_index}) or count ({count}). Must be >= 1.")

        resolved_path = validate_file_path(file_path_str, must_exist=True)
        wb = WorkbookManager.get_or_open(resolved_path, data_only=False)
        ws = WorksheetManager.get_sheet(wb, sheet_name)

        ws.insert_rows(idx=row_index, amount=count)
        WorkbookManager.save_workbook(resolved_path)

        logger.info("Inserted %s rows at index %s in %s", count, row_index, ws.title)
        return {
            "status": "success",
            "sheet": ws.title,
            "row_index": row_index,
            "rows_inserted": count,
            "new_max_row": ws.max_row,
        }

    @classmethod
    def delete_rows(
        cls,
        file_path_str: str | Path,
        row_index: int,
        count: int = 1,
        sheet_name: Optional[str] = None,
        confirm: bool = False,
    ) -> Dict[str, Any]:
        """Deletes count rows starting at row_index. Requires confirmation."""
        if row_index < 1 or count < 1:
            raise ValueError(f"Invalid row_index ({row_index}) or count ({count}). Must be >= 1.")

        resolved_path = validate_file_path(file_path_str, must_exist=True)
        wb = WorkbookManager.get_or_open(resolved_path, data_only=False)
        ws = WorksheetManager.get_sheet(wb, sheet_name)

        ConfirmationGuard.check_confirmation(
            action_name="delete_rows",
            confirmed=confirm,
            impact_details={
                "sheet": ws.title,
                "start_row": row_index,
                "count": count,
                "file": resolved_path.name,
            },
            custom_message=f"Deleting {count} row(s) starting from row {row_index}. Confirm with confirm=True.",
        )

        ws.delete_rows(idx=row_index, amount=count)
        WorkbookManager.save_workbook(resolved_path)

        logger.info("Deleted %s rows at index %s from %s", count, row_index, ws.title)
        return {
            "status": "success",
            "sheet": ws.title,
            "row_index": row_index,
            "rows_deleted": count,
            "new_max_row": ws.max_row,
        }

    @classmethod
    def insert_columns(
        cls,
        file_path_str: str | Path,
        col_index: Union[str, int],
        count: int = 1,
        sheet_name: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Inserts count empty columns before col_index."""
        if isinstance(col_index, str):
            c_idx = column_index_from_string(col_index.strip().upper())
        else:
            c_idx = int(col_index)

        if c_idx < 1 or count < 1:
            raise ValueError(f"Invalid col_index ({col_index}) or count ({count}). Must be >= 1.")

        resolved_path = validate_file_path(file_path_str, must_exist=True)
        wb = WorkbookManager.get_or_open(resolved_path, data_only=False)
        ws = WorksheetManager.get_sheet(wb, sheet_name)

        ws.insert_cols(idx=c_idx, amount=count)
        WorkbookManager.save_workbook(resolved_path)

        logger.info("Inserted %s columns at index %s in %s", count, c_idx, ws.title)
        return {
            "status": "success",
            "sheet": ws.title,
            "column_index": c_idx,
            "column_letter": get_column_letter(c_idx),
            "columns_inserted": count,
            "new_max_column": ws.max_column,
        }

    @classmethod
    def delete_columns(
        cls,
        file_path_str: str | Path,
        col_index: Union[str, int],
        count: int = 1,
        sheet_name: Optional[str] = None,
        confirm: bool = False,
    ) -> Dict[str, Any]:
        """Deletes count columns starting at col_index. Requires confirmation."""
        if isinstance(col_index, str):
            c_idx = column_index_from_string(col_index.strip().upper())
        else:
            c_idx = int(col_index)

        if c_idx < 1 or count < 1:
            raise ValueError(f"Invalid col_index ({col_index}) or count ({count}). Must be >= 1.")

        resolved_path = validate_file_path(file_path_str, must_exist=True)
        wb = WorkbookManager.get_or_open(resolved_path, data_only=False)
        ws = WorksheetManager.get_sheet(wb, sheet_name)

        ConfirmationGuard.check_confirmation(
            action_name="delete_columns",
            confirmed=confirm,
            impact_details={
                "sheet": ws.title,
                "start_col": get_column_letter(c_idx),
                "count": count,
                "file": resolved_path.name,
            },
            custom_message=f"Deleting {count} column(s) starting from column {get_column_letter(c_idx)}. Confirm with confirm=True.",
        )

        ws.delete_cols(idx=c_idx, amount=count)
        WorkbookManager.save_workbook(resolved_path)

        logger.info("Deleted %s columns at index %s from %s", count, c_idx, ws.title)
        return {
            "status": "success",
            "sheet": ws.title,
            "column_letter": get_column_letter(c_idx),
            "columns_deleted": count,
            "new_max_column": ws.max_column,
        }

    @classmethod
    def create_chart(
        cls,
        file_path_str: str | Path,
        chart_type: str,
        data_range: str,
        title: str,
        target_cell: str = "E2",
        categories_range: Optional[str] = None,
        sheet_name: Optional[str] = None,
        width: float = 15.0,
        height: float = 10.0,
    ) -> Dict[str, Any]:
        """
        Creates a native Excel chart (bar, line, pie, area) and embeds it in the sheet.
        """
        norm_type = chart_type.strip().lower()
        if norm_type not in cls.CHART_CLASSES:
            supported = ", ".join(f"'{k}'" for k in cls.CHART_CLASSES.keys())
            raise ValueError(f"Unsupported chart type '{chart_type}'. Supported: {supported}")

        resolved_path = validate_file_path(file_path_str, must_exist=True)
        embedded_sheet, clean_range = split_sheet_and_range(data_range)
        target_sheet_name = embedded_sheet or sheet_name

        wb = WorkbookManager.get_or_open(resolved_path, data_only=False)
        ws = WorksheetManager.get_sheet(wb, target_sheet_name)

        # Parse numerical data range
        min_row, min_col, max_row, max_col = parse_range_coordinates(clean_range)
        chart_obj = cls.CHART_CLASSES[norm_type]()
        chart_obj.title = title
        chart_obj.width = width
        chart_obj.height = height

        data_ref = Reference(
            ws,
            min_col=min_col,
            min_row=min_row,
            max_col=max_col,
            max_row=max_row,
        )
        chart_obj.add_data(data_ref, titles_from_data=True)

        if categories_range:
            _, cat_clean = split_sheet_and_range(categories_range)
            c_min_row, c_min_col, c_max_row, c_max_col = parse_range_coordinates(cat_clean)
            cats_ref = Reference(
                ws,
                min_col=c_min_col,
                min_row=c_min_row,
                max_col=c_max_col,
                max_row=c_max_row,
            )
            chart_obj.set_categories(cats_ref)

        ws.add_chart(chart_obj, target_cell.upper())
        WorkbookManager.save_workbook(resolved_path)

        logger.info("Added %s chart '%s' to %s at %s", norm_type, title, ws.title, target_cell)
        return {
            "status": "success",
            "sheet": ws.title,
            "chart_type": norm_type,
            "title": title,
            "target_cell": target_cell.upper(),
            "data_range": clean_range,
        }
