"""
Helper routines for Excel coordinates, ranges, conversions, and lock-file detection.
"""

from datetime import date, datetime, time
from decimal import Decimal
from pathlib import Path
import re
from typing import Any, Optional, Tuple
from openpyxl.utils.cell import (
    coordinate_to_tuple,
    get_column_letter,
    column_index_from_string,
)


def split_sheet_and_range(range_expr: str) -> Tuple[Optional[str], str]:
    """
    Splits an optional sheet prefix from a range expression.
    Examples:
      - "Sheet1!A1:B10" -> ("Sheet1", "A1:B10")
      - "'Sales 2024'!C5" -> ("Sales 2024", "C5")
      - "A1:B10" -> (None, "A1:B10")
    """
    clean_expr = range_expr.strip()
    if "!" in clean_expr:
        sheet_part, cell_part = clean_expr.split("!", 1)
        sheet_name = sheet_part.strip().strip("'").strip('"')
        return sheet_name, cell_part.strip()
    return None, clean_expr


def parse_cell_coordinate(cell_ref: str) -> Tuple[int, int]:
    """
    Parses a single cell reference (e.g. 'B5' or '$B$5') into 1-based (row, column) tuple.
    Example: 'B5' -> (5, 2)
    """
    clean_ref = cell_ref.replace("$", "").strip().upper()
    try:
        row, col = coordinate_to_tuple(clean_ref)
        return row, col
    except Exception as e:
        raise ValueError(f"Invalid cell reference format: '{cell_ref}'") from e


def format_cell_coordinate(row: int, column: int) -> str:
    """
    Converts 1-based (row, column) tuple into A1-notation cell string.
    Example: (5, 2) -> 'B5'
    """
    if row < 1 or column < 1:
        raise ValueError(f"Coordinates must be >= 1. Got row={row}, col={column}")
    col_letter = get_column_letter(column)
    return f"{col_letter}{row}"


def parse_range_coordinates(range_str: str) -> Tuple[int, int, int, int]:
    """
    Parses a range string into (min_row, min_col, max_row, max_col).
    Supports:
      - Bounded ranges: "A1:C10" -> (1, 1, 10, 3)
      - Single cell: "B5" -> (5, 2, 5, 2)
    """
    _, cell_range = split_sheet_and_range(range_str)
    parts = cell_range.split(":")
    
    if len(parts) == 1:
        row, col = parse_cell_coordinate(parts[0])
        return row, col, row, col
    elif len(parts) == 2:
        start_row, start_col = parse_cell_coordinate(parts[0])
        end_row, end_col = parse_cell_coordinate(parts[1])
        min_row = min(start_row, end_row)
        max_row = max(start_row, end_row)
        min_col = min(start_col, end_col)
        max_col = max(start_col, end_col)
        return min_row, min_col, max_row, max_col
    else:
        raise ValueError(f"Malformed range expression: '{range_str}'")


def is_excel_lock_present(file_path: Path) -> bool:
    """
    Checks if Microsoft Excel's lock file (~$filename.xlsx) exists,
    indicating the workbook is actively open in desktop Microsoft Excel.
    """
    lock_file = file_path.parent / f"~${file_path.name}"
    return lock_file.exists()


def to_serializable_value(val: Any) -> Any:
    """
    Converts Excel cell values (datetime, Decimal, numpy types, formulas)
    into standard JSON-serializable Python primitives.
    """
    if val is None:
        return None
    if isinstance(val, (datetime, date)):
        return val.isoformat()
    if isinstance(val, time):
        return val.strftime("%H:%M:%S")
    if isinstance(val, Decimal):
        return float(val)
    # Check for pandas/numpy NaN or NaT
    if hasattr(val, "__class__") and "NA" in val.__class__.__name__:
        return None
    try:
        import numpy as np
        if isinstance(val, (np.integer, np.floating)):
            return val.item()
        if isinstance(val, np.ndarray):
            return val.tolist()
    except ImportError:
        pass

    return val
