"""
DeepSeek Excel Studio - MCP Server implementation.
Registers production-ready Excel automation tools for DeepSeek Harness Desktop.
"""

from pathlib import Path
from typing import Any, Dict, List, Optional, Union
import json
import traceback

try:
    from mcp.server.mcpserver import MCPServer
except ImportError:
    from mcp.server.fastmcp import FastMCP as MCPServer

from excel_mcp.config import config
from excel_mcp.core.advanced import AdvancedManager
from excel_mcp.core.analysis import AnalysisManager
from excel_mcp.core.cells import CellManager
from excel_mcp.core.formatting import FormattingManager
from excel_mcp.core.workbook import WorkbookManager
from excel_mcp.core.worksheet import WorksheetManager
from excel_mcp.logging_config import setup_logging, logger
from excel_mcp.security.guard import ConfirmationRequiredError
from excel_mcp.security.validator import SecurityError


# Initialize MCP Server instance
app = MCPServer("excel-mcp")


def handle_tool_errors(func):
    """Decorator to catch and format exceptions into structured JSON results."""
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except ConfirmationRequiredError as e:
            return e.to_dict()
        except SecurityError as e:
            logger.warning("Security violation in %s: %s", func.__name__, e)
            return {
                "status": "error",
                "error_type": "SecurityError",
                "message": str(e),
            }
        except (ValueError, FileNotFoundError, PermissionError) as e:
            logger.warning("Operational error in %s: %s", func.__name__, e)
            return {
                "status": "error",
                "error_type": type(e).__name__,
                "message": str(e),
            }
        except Exception as e:
            logger.error("Unhandled error in %s: %s\n%s", func.__name__, e, traceback.format_exc())
            return {
                "status": "error",
                "error_type": "InternalServerError",
                "message": f"Unexpected error executing {func.__name__}: {str(e)}",
            }
    wrapper.__name__ = func.__name__
    wrapper.__doc__ = func.__doc__
    return wrapper


# ============================================================================
# 1. WORKBOOK TOOLS
# ============================================================================

@app.tool()
@handle_tool_errors
def list_workbooks(directory: Optional[str] = None) -> List[Dict[str, Any]]:
    """
    List Excel workbooks (.xlsx, .xlsm) in a directory or in configured allowed directories.
    
    Args:
        directory: Optional directory path. Defaults to allowed directories root.
    """
    return WorkbookManager.list_workbooks(directory)


@app.tool()
@handle_tool_errors
def open_workbook(
    file_path: str,
    read_only: bool = False,
    data_only: bool = False,
) -> Dict[str, Any]:
    """
    Open an Excel workbook and cache it in the active session.
    
    Args:
        file_path: Absolute or workspace path to the .xlsx or .xlsm file.
        read_only: If True, opens in read-only mode for large file performance.
        data_only: If True, reads cached formula values rather than formula strings.
    """
    WorkbookManager.open_workbook(file_path, read_only=read_only, data_only=data_only)
    return WorkbookManager.get_workbook_info(file_path)


@app.tool()
@handle_tool_errors
def get_workbook_info(file_path: str) -> Dict[str, Any]:
    """
    Inspect an Excel workbook to retrieve its sheets, active sheet, dimensions, and file stats.
    
    Args:
        file_path: Path to the Excel workbook.
    """
    return WorkbookManager.get_workbook_info(file_path)


@app.tool()
@handle_tool_errors
def save_workbook(
    file_path: str,
    target_path: Optional[str] = None,
    overwrite: bool = False,
) -> Dict[str, Any]:
    """
    Save the current workbook changes to disk.
    
    Args:
        file_path: Path to the open workbook.
        target_path: Optional new destination path to save a copy.
        overwrite: Set to True to allow overwriting an existing destination file.
    """
    return WorkbookManager.save_workbook(file_path, target_path_str=target_path, overwrite=overwrite)


@app.tool()
@handle_tool_errors
def close_workbook(file_path: str, save_changes: bool = False) -> Dict[str, Any]:
    """
    Close an active workbook and release any memory or file locks.
    
    Args:
        file_path: Path to the workbook to close.
        save_changes: If True, saves any pending changes before closing.
    """
    return WorkbookManager.close_workbook(file_path, save_changes=save_changes)


# ============================================================================
# 2. WORKSHEET TOOLS
# ============================================================================

@app.tool()
@handle_tool_errors
def list_worksheets(file_path: str) -> List[Dict[str, Any]]:
    """
    List all worksheets in an Excel workbook with their dimensions and visibility state.
    
    Args:
        file_path: Path to the Excel workbook.
    """
    return WorksheetManager.list_worksheets(file_path)


@app.tool()
@handle_tool_errors
def create_worksheet(
    file_path: str,
    sheet_name: str,
    index: Optional[int] = None,
) -> Dict[str, Any]:
    """
    Create a new worksheet in the workbook.
    
    Args:
        file_path: Path to the Excel workbook.
        sheet_name: Unique name for the new worksheet (max 31 characters).
        index: Optional zero-based position index to insert the sheet.
    """
    return WorksheetManager.create_worksheet(file_path, sheet_name=sheet_name, index=index)


@app.tool()
@handle_tool_errors
def delete_worksheet(
    file_path: str,
    sheet_name: str,
    confirm: bool = False,
) -> Dict[str, Any]:
    """
    Delete a worksheet from the workbook. Destructive: requires confirm=True.
    
    Args:
        file_path: Path to the Excel workbook.
        sheet_name: Name of the worksheet to permanently delete.
        confirm: Confirmation flag. If False, returns an impact preview first.
    """
    return WorksheetManager.delete_worksheet(file_path, sheet_name=sheet_name, confirm=confirm)


@app.tool()
@handle_tool_errors
def rename_worksheet(
    file_path: str,
    old_name: str,
    new_name: str,
) -> Dict[str, Any]:
    """
    Rename an existing worksheet in the workbook.
    
    Args:
        file_path: Path to the Excel workbook.
        old_name: Current name of the worksheet.
        new_name: New name for the worksheet.
    """
    return WorksheetManager.rename_worksheet(file_path, old_name=old_name, new_name=new_name)


# ============================================================================
# 3. CELL AND RANGE TOOLS
# ============================================================================

@app.tool()
@handle_tool_errors
def read_range(
    file_path: str,
    range: str,
    sheet: Optional[str] = None,
    data_only: bool = False,
) -> Dict[str, Any]:
    """
    Read cell values and formulas from a 2D range (e.g. 'A1:D20' or 'B5').
    
    Args:
        file_path: Path to the Excel workbook.
        range: Range coordinate expression in A1 notation (e.g. 'A1:C10' or 'Sheet1!B2:D5').
        sheet: Optional sheet name if not specified in the range string.
        data_only: If True, returns cached calculation values instead of formula strings.
    """
    return CellManager.read_range(file_path, range_str=range, sheet_name=sheet, data_only=data_only)


@app.tool()
@handle_tool_errors
def write_range(
    file_path: str,
    start_cell: str,
    data: List[List[Any]],
    sheet: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Write a 2D matrix of values or formulas into a sheet starting at start_cell.
    Formulas starting with '=' are properly stored as Excel formulas.
    
    Args:
        file_path: Path to the Excel workbook.
        start_cell: Upper-left cell coordinate to start writing (e.g. 'B2').
        data: 2D list of values or formulas (e.g. [['Revenue', 1000], ['Tax', '=B1*0.2']]).
        sheet: Optional sheet name.
    """
    return CellManager.write_range(file_path, start_cell=start_cell, data=data, sheet_name=sheet)


@app.tool()
@handle_tool_errors
def update_cell(
    file_path: str,
    cell: str,
    value: Any,
    sheet: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Update a single cell value or formula.
    
    Args:
        file_path: Path to the Excel workbook.
        cell: Cell coordinate (e.g. 'C10').
        value: Value or formula string (e.g. 4500 or '=SUM(C2:C9)').
        sheet: Optional sheet name.
    """
    return CellManager.update_cell(file_path, cell=cell, value=value, sheet_name=sheet)


@app.tool()
@handle_tool_errors
def clear_range(
    file_path: str,
    range: str,
    sheet: Optional[str] = None,
    confirm: bool = False,
) -> Dict[str, Any]:
    """
    Clear cell contents in a specified range. Large ranges (>10 cells) require confirm=True.
    
    Args:
        file_path: Path to the Excel workbook.
        range: Range coordinate (e.g. 'A10:D20').
        sheet: Optional sheet name.
        confirm: Confirmation flag for large ranges.
    """
    return CellManager.clear_range(file_path, range_str=range, sheet_name=sheet, confirm=confirm)


@app.tool()
@handle_tool_errors
def find_value(
    file_path: str,
    query: str,
    sheet: Optional[str] = None,
    match_case: bool = False,
    exact_match: bool = False,
    max_results: int = 100,
) -> List[Dict[str, Any]]:
    """
    Search for text or numbers across an Excel worksheet.
    
    Args:
        file_path: Path to the Excel workbook.
        query: String or number to search for.
        sheet: Optional sheet name.
        match_case: Whether search is case-sensitive.
        exact_match: If True, matches entire cell value; otherwise matches substrings.
        max_results: Limit number of returned coordinates (default 100).
    """
    return CellManager.find_value(
        file_path,
        query=query,
        sheet_name=sheet,
        match_case=match_case,
        exact_match=exact_match,
        max_results=max_results,
    )


@app.tool()
@handle_tool_errors
def replace_value(
    file_path: str,
    old_value: Any,
    new_value: Any,
    sheet: Optional[str] = None,
    range: Optional[str] = None,
    match_case: bool = False,
    exact_match: bool = True,
    max_replacements: int = 1000,
) -> Dict[str, Any]:
    """
    Find and replace values across an entire sheet or within a bounded range.
    
    Args:
        file_path: Path to the Excel workbook.
        old_value: Value to find.
        new_value: Value to substitute.
        sheet: Optional sheet name.
        range: Optional range boundary to limit replacements.
        match_case: Whether string comparison is case-sensitive.
        exact_match: If True, replaces entire cell values; if False, replaces substrings.
        max_replacements: Maximum number of cells to modify (default 1000).
    """
    return CellManager.replace_value(
        file_path,
        old_value=old_value,
        new_value=new_value,
        sheet_name=sheet,
        range_str=range,
        match_case=match_case,
        exact_match=exact_match,
        max_replacements=max_replacements,
    )


# ============================================================================
# 4. FORMATTING TOOLS
# ============================================================================

@app.tool()
@handle_tool_errors
def format_cells(
    file_path: str,
    range: str,
    sheet: Optional[str] = None,
    font: Optional[Dict[str, Any]] = None,
    fill: Optional[Dict[str, Any]] = None,
    alignment: Optional[Dict[str, Any]] = None,
    border: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """
    Apply typography, background fills, alignment, and borders to a range.
    
    Args:
        file_path: Path to the Excel workbook.
        range: Target cell range (e.g. 'A1:D1').
        sheet: Optional sheet name.
        font: Dict with font settings (e.g. {'bold': True, 'size': 12, 'color': 'FF0000', 'name': 'Calibri'}).
        fill: Dict with fill settings (e.g. {'bg_color': 'E0E0E0', 'fill_type': 'solid'}).
        alignment: Dict with alignment (e.g. {'horizontal': 'center', 'vertical': 'center', 'wrap_text': True}).
        border: Dict with border settings (e.g. {'style': 'thin', 'color': '000000'}).
    """
    return FormattingManager.format_cells(
        file_path,
        range_str=range,
        sheet_name=sheet,
        font_style=font,
        fill_style=fill,
        alignment_style=alignment,
        border_style=border,
    )


@app.tool()
@handle_tool_errors
def set_column_width(
    file_path: str,
    column: str,
    width: Optional[float] = None,
    auto_fit: bool = False,
    sheet: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Set an explicit column width or auto-fit column based on cell contents.
    
    Args:
        file_path: Path to the Excel workbook.
        column: Column letter (e.g. 'A', 'B', 'AA').
        width: Explicit width value.
        auto_fit: If True, calculates width automatically from cell contents.
        sheet: Optional sheet name.
    """
    return FormattingManager.set_column_width(
        file_path, column=column, width=width, auto_fit=auto_fit, sheet_name=sheet
    )


@app.tool()
@handle_tool_errors
def set_row_height(
    file_path: str,
    row: int,
    height: float,
    sheet: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Set the height of a specific row in points.
    
    Args:
        file_path: Path to the Excel workbook.
        row: 1-based row number (e.g. 1).
        height: Row height in points (e.g. 25.0).
        sheet: Optional sheet name.
    """
    return FormattingManager.set_row_height(file_path, row=row, height=height, sheet_name=sheet)


@app.tool()
@handle_tool_errors
def apply_number_format(
    file_path: str,
    range: str,
    format_code: str,
    sheet: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Apply a number, currency, date, or percentage format to a cell range.
    
    Args:
        file_path: Path to the Excel workbook.
        range: Target cell range (e.g. 'D2:D100').
        format_code: Preset name ('currency', 'percentage', 'integer', 'decimal', 'date') or custom Excel code.
        sheet: Optional sheet name.
    """
    return FormattingManager.apply_number_format(
        file_path, range_str=range, format_code=format_code, sheet_name=sheet
    )


# ============================================================================
# 5. DATA OPERATIONS (PANDAS ENGINE)
# ============================================================================

@app.tool()
@handle_tool_errors
def sort_range(
    file_path: str,
    range: str,
    sort_by_col: Union[str, int],
    ascending: bool = True,
    has_headers: bool = True,
    sheet: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Sort a tabular range by a specified column and update the worksheet.
    
    Args:
        file_path: Path to the Excel workbook.
        range: Tabular data range (e.g. 'A1:D50').
        sort_by_col: Column header name (if has_headers=True) or 1-based column index.
        ascending: True for ascending order, False for descending.
        has_headers: Whether the first row in range contains headers.
        sheet: Optional sheet name.
    """
    return AnalysisManager.sort_range(
        file_path,
        range_str=range,
        sort_by_col=sort_by_col,
        ascending=ascending,
        has_headers=has_headers,
        sheet_name=sheet,
    )


@app.tool()
@handle_tool_errors
def filter_data(
    file_path: str,
    range: str,
    column: Union[str, int],
    operator: str,
    value: Any,
    has_headers: bool = True,
    sheet: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Filter tabular rows matching condition (e.g. 'Salary' > 90000 or 'Status' == 'Active').
    
    Args:
        file_path: Path to the Excel workbook.
        range: Tabular data range.
        column: Column header name or 1-based index to filter by.
        operator: Filter operator: '==', '!=', '>', '<', '>=', '<=', 'contains', 'in'.
        value: Value to compare against.
        has_headers: Whether the first row contains headers.
        sheet: Optional sheet name.
    """
    return AnalysisManager.filter_data(
        file_path,
        range_str=range,
        column=column,
        operator=operator,
        value=value,
        has_headers=has_headers,
        sheet_name=sheet,
    )


@app.tool()
@handle_tool_errors
def find_duplicates(
    file_path: str,
    range: str,
    columns: Optional[List[Union[str, int]]] = None,
    has_headers: bool = True,
    sheet: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Find duplicate rows in a dataset and return their row numbers.
    
    Args:
        file_path: Path to the Excel workbook.
        range: Tabular data range.
        columns: Optional list of column names or indices to check for duplicates.
        has_headers: Whether the first row contains headers.
        sheet: Optional sheet name.
    """
    return AnalysisManager.find_duplicates(
        file_path,
        range_str=range,
        columns=columns,
        has_headers=has_headers,
        sheet_name=sheet,
    )


@app.tool()
@handle_tool_errors
def calculate_summary(
    file_path: str,
    range: str,
    columns: Optional[List[Union[str, int]]] = None,
    has_headers: bool = True,
    sheet: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Calculate statistical summaries (count, sum, mean, median, min, max, std) for columns in range.
    
    Args:
        file_path: Path to the Excel workbook.
        range: Tabular data range.
        columns: Optional list of column names/indices to analyze.
        has_headers: Whether the first row contains headers.
        sheet: Optional sheet name.
    """
    return AnalysisManager.calculate_summary(
        file_path,
        range_str=range,
        columns=columns,
        has_headers=has_headers,
        sheet_name=sheet,
    )


# ============================================================================
# 6. ADVANCED (ROWS, COLUMNS, CHARTS)
# ============================================================================

@app.tool()
@handle_tool_errors
def insert_rows(
    file_path: str,
    row_index: int,
    count: int = 1,
    sheet: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Insert blank rows at row_index, shifting existing content downwards.
    
    Args:
        file_path: Path to the Excel workbook.
        row_index: 1-based row index where rows will be inserted.
        count: Number of blank rows to insert (default 1).
        sheet: Optional sheet name.
    """
    return AdvancedManager.insert_rows(file_path, row_index=row_index, count=count, sheet_name=sheet)


@app.tool()
@handle_tool_errors
def delete_rows(
    file_path: str,
    row_index: int,
    count: int = 1,
    sheet: Optional[str] = None,
    confirm: bool = False,
) -> Dict[str, Any]:
    """
    Delete rows starting at row_index. Destructive: requires confirm=True.
    
    Args:
        file_path: Path to the Excel workbook.
        row_index: 1-based row index to begin deletion.
        count: Number of rows to delete.
        sheet: Optional sheet name.
        confirm: Confirmation flag.
    """
    return AdvancedManager.delete_rows(
        file_path, row_index=row_index, count=count, sheet_name=sheet, confirm=confirm
    )


@app.tool()
@handle_tool_errors
def insert_columns(
    file_path: str,
    col_index: Union[str, int],
    count: int = 1,
    sheet: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Insert blank columns at col_index, shifting existing content rightwards.
    
    Args:
        file_path: Path to the Excel workbook.
        col_index: Column letter (e.g. 'C') or 1-based index (e.g. 3).
        count: Number of blank columns to insert (default 1).
        sheet: Optional sheet name.
    """
    return AdvancedManager.insert_columns(
        file_path, col_index=col_index, count=count, sheet_name=sheet
    )


@app.tool()
@handle_tool_errors
def delete_columns(
    file_path: str,
    col_index: Union[str, int],
    count: int = 1,
    sheet: Optional[str] = None,
    confirm: bool = False,
) -> Dict[str, Any]:
    """
    Delete columns starting at col_index. Destructive: requires confirm=True.
    
    Args:
        file_path: Path to the Excel workbook.
        col_index: Column letter (e.g. 'C') or 1-based index.
        count: Number of columns to delete.
        sheet: Optional sheet name.
        confirm: Confirmation flag.
    """
    return AdvancedManager.delete_columns(
        file_path, col_index=col_index, count=count, sheet_name=sheet, confirm=confirm
    )


@app.tool()
@handle_tool_errors
def create_chart(
    file_path: str,
    chart_type: str,
    data_range: str,
    title: str,
    target_cell: str = "E2",
    categories_range: Optional[str] = None,
    sheet: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Create a native Excel chart (bar, line, pie, area) and embed it in the sheet.
    
    Args:
        file_path: Path to the Excel workbook.
        chart_type: Type of chart: 'bar', 'line', 'pie', 'area'.
        data_range: Cell range containing numerical series data (e.g. 'D1:D10').
        title: Title string displayed on chart.
        target_cell: Upper-left cell where the chart will be placed (default 'E2').
        categories_range: Optional range for axis category labels (e.g. 'A2:A10').
        sheet: Optional sheet name.
    """
    return AdvancedManager.create_chart(
        file_path,
        chart_type=chart_type,
        data_range=data_range,
        title=title,
        target_cell=target_cell,
        categories_range=categories_range,
        sheet_name=sheet,
    )


# ============================================================================
# SERVER RUNNER
# ============================================================================

def main():
    """Main process entry point for stdio MCP execution."""
    setup_logging(level=config.log_level)
    logger.info("Starting DeepSeek Excel Studio MCP Server (transport: stdio)...")
    app.run(transport="stdio")


if __name__ == "__main__":
    main()
