"""
Data analysis engine powered by Pandas and Openpyxl.
Provides range sorting, conditional filtering, duplicate detection,
and statistical summary computations.
"""

from pathlib import Path
from typing import Any, Dict, List, Optional, Union
import pandas as pd
import openpyxl

from excel_mcp.core.cells import CellManager
from excel_mcp.core.workbook import WorkbookManager
from excel_mcp.core.worksheet import WorksheetManager
from excel_mcp.logging_config import logger
from excel_mcp.security.validator import validate_file_path
from excel_mcp.utils.excel_helpers import (
    parse_range_coordinates,
    format_cell_coordinate,
    split_sheet_and_range,
    to_serializable_value,
)


class AnalysisManager:
    """Performs analytical transformations and statistical aggregations on Excel data."""

    @classmethod
    def _range_to_dataframe(
        cls,
        file_path: Path,
        range_str: str,
        sheet_name: Optional[str] = None,
        has_headers: bool = True,
    ) -> tuple[pd.DataFrame, int, int]:
        """Loads a range into a Pandas DataFrame and returns (df, min_row, min_col)."""
        read_res = CellManager.read_range(
            file_path,
            range_str=range_str,
            sheet_name=sheet_name,
            data_only=True,
        )
        values = read_res["values"]
        if not values or not values[0]:
            raise ValueError(f"Range '{range_str}' contains no data.")

        min_row, min_col, _, _ = parse_range_coordinates(read_res["range"])

        if has_headers:
            headers = [str(h) if h is not None else f"Col_{i}" for i, h in enumerate(values[0])]
            data_rows = values[1:]
            df = pd.DataFrame(data_rows, columns=headers)
        else:
            headers = [f"Col_{i+1}" for i in range(len(values[0]))]
            df = pd.DataFrame(values, columns=headers)

        return df, min_row, min_col

    @classmethod
    def sort_range(
        cls,
        file_path_str: str | Path,
        range_str: str,
        sort_by_col: Union[str, int],
        ascending: bool = True,
        has_headers: bool = True,
        sheet_name: Optional[str] = None,
    ) -> Dict[str, Any]:
        """
        Sorts the tabular range by a specified column and writes the sorted rows back.
        """
        resolved_path = validate_file_path(file_path_str, must_exist=True)
        df, min_row, min_col = cls._range_to_dataframe(
            resolved_path, range_str, sheet_name=sheet_name, has_headers=has_headers
        )

        col_to_sort: str
        if isinstance(sort_by_col, int):
            # 1-based index
            if 1 <= sort_by_col <= len(df.columns):
                col_to_sort = df.columns[sort_by_col - 1]
            else:
                raise ValueError(f"Column index {sort_by_col} is out of bounds (1-{len(df.columns)}).")
        else:
            if sort_by_col not in df.columns:
                available = ", ".join(f"'{c}'" for c in df.columns)
                raise ValueError(f"Column '{sort_by_col}' not found. Available: {available}")
            col_to_sort = sort_by_col

        df_sorted = df.sort_values(by=col_to_sort, ascending=ascending)

        # Write sorted rows back
        start_row = min_row + (1 if has_headers else 0)
        start_cell = format_cell_coordinate(start_row, min_col)
        sorted_rows = df_sorted.values.tolist()

        CellManager.write_range(
            resolved_path,
            start_cell=start_cell,
            data=sorted_rows,
            sheet_name=sheet_name,
        )

        logger.info("Sorted range %s by %s (asc=%s)", range_str, col_to_sort, ascending)
        return {
            "status": "success",
            "sorted_by": col_to_sort,
            "ascending": ascending,
            "rows_sorted": len(df_sorted),
            "range": range_str,
        }

    @classmethod
    def filter_data(
        cls,
        file_path_str: str | Path,
        range_str: str,
        column: Union[str, int],
        operator: str,
        value: Any,
        has_headers: bool = True,
        sheet_name: Optional[str] = None,
    ) -> Dict[str, Any]:
        """
        Filters tabular rows in range based on condition.
        Operators: '==', '!=', '>', '<', '>=', '<=', 'contains', 'not_contains', 'in'
        """
        resolved_path = validate_file_path(file_path_str, must_exist=True)
        df, _, _ = cls._range_to_dataframe(
            resolved_path, range_str, sheet_name=sheet_name, has_headers=has_headers
        )

        col_name = df.columns[column - 1] if isinstance(column, int) else column
        if col_name not in df.columns:
            raise ValueError(f"Column '{col_name}' not found.")

        series = df[col_name]
        op = operator.strip().lower()

        if op in ("==", "equals", "="):
            mask = series == value
        elif op in ("!=", "not_equals", "<>"):
            mask = series != value
        elif op == ">":
            mask = pd.to_numeric(series, errors="coerce") > float(value)
        elif op == "<":
            mask = pd.to_numeric(series, errors="coerce") < float(value)
        elif op == ">=":
            mask = pd.to_numeric(series, errors="coerce") >= float(value)
        elif op == "<=":
            mask = pd.to_numeric(series, errors="coerce") <= float(value)
        elif op == "contains":
            mask = series.astype(str).str.contains(str(value), case=False, na=False)
        elif op == "not_contains":
            mask = ~series.astype(str).str.contains(str(value), case=False, na=False)
        elif op == "in":
            items = value if isinstance(value, list) else [value]
            mask = series.isin(items)
        else:
            raise ValueError(f"Unsupported filter operator: '{operator}'")

        filtered_df = df[mask]
        results = [
            {k: to_serializable_value(v) for k, v in row.items()}
            for row in filtered_df.to_dict(orient="records")
        ]

        return {
            "total_rows": len(df),
            "matched_rows": len(filtered_df),
            "column": col_name,
            "operator": operator,
            "filter_value": value,
            "records": results,
        }

    @classmethod
    def find_duplicates(
        cls,
        file_path_str: str | Path,
        range_str: str,
        columns: Optional[List[Union[str, int]]] = None,
        has_headers: bool = True,
        sheet_name: Optional[str] = None,
    ) -> Dict[str, Any]:
        """
        Finds duplicate rows in tabular data.
        """
        resolved_path = validate_file_path(file_path_str, must_exist=True)
        df, min_row, _ = cls._range_to_dataframe(
            resolved_path, range_str, sheet_name=sheet_name, has_headers=has_headers
        )

        subset_cols = None
        if columns:
            subset_cols = [
                df.columns[c - 1] if isinstance(c, int) else c for c in columns
            ]

        dup_mask = df.duplicated(subset=subset_cols, keep=False)
        dup_df = df[dup_mask].copy()
        
        # Add original spreadsheet row number
        offset = min_row + (1 if has_headers else 0)
        dup_df["_row_number"] = dup_df.index + offset

        records = [
            {k: to_serializable_value(v) for k, v in row.items()}
            for row in dup_df.to_dict(orient="records")
        ]

        return {
            "total_rows": len(df),
            "duplicate_count": len(dup_df),
            "checked_columns": subset_cols or list(df.columns),
            "duplicates": records,
        }

    @classmethod
    def calculate_summary(
        cls,
        file_path_str: str | Path,
        range_str: str,
        columns: Optional[List[Union[str, int]]] = None,
        has_headers: bool = True,
        sheet_name: Optional[str] = None,
    ) -> Dict[str, Any]:
        """
        Calculates summary statistics (count, sum, mean, median, min, max, std) for columns.
        """
        resolved_path = validate_file_path(file_path_str, must_exist=True)
        df, _, _ = cls._range_to_dataframe(
            resolved_path, range_str, sheet_name=sheet_name, has_headers=has_headers
        )

        target_cols = df.columns
        if columns:
            target_cols = [
                df.columns[c - 1] if isinstance(c, int) else c for c in columns
            ]

        summaries = {}
        for col in target_cols:
            series = pd.to_numeric(df[col], errors="coerce")
            valid = series.dropna()
            
            if not valid.empty:
                summaries[col] = {
                    "count": int(valid.count()),
                    "sum": float(valid.sum()),
                    "mean": float(valid.mean()),
                    "median": float(valid.median()),
                    "min": float(valid.min()),
                    "max": float(valid.max()),
                    "std": float(valid.std()) if len(valid) > 1 else 0.0,
                }
            else:
                # Text / non-numeric column summary
                summaries[col] = {
                    "count": int(df[col].count()),
                    "unique": int(df[col].nunique()),
                    "top": to_serializable_value(df[col].mode().iloc[0]) if not df[col].empty else None,
                }

        return {
            "sheet": sheet_name or "active",
            "range": range_str,
            "total_rows": len(df),
            "columns_analyzed": len(summaries),
            "summary": summaries,
        }
