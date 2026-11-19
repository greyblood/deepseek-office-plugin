"""
Unit tests for AnalysisManager operations.
"""

from pathlib import Path
import pytest
from excel_mcp.core.analysis import AnalysisManager
from excel_mcp.core.cells import CellManager


def test_sort_range(sample_workbook: Path):
    """Tests sorting rows by column."""
    res = AnalysisManager.sort_range(
        sample_workbook,
        range_str="A1:D5",
        sort_by_col="Salary",
        ascending=False,
    )
    assert res["status"] == "success"
    assert res["sorted_by"] == "Salary"

    # Read back to verify order: Diana (105k), Alice (95k), Charlie (90k), Bob (80k)
    read_back = CellManager.read_range(sample_workbook, "B2:D5")
    names = [row[0] for row in read_back["values"]]
    assert names == ["Diana", "Alice", "Charlie", "Bob"]


def test_filter_data(sample_workbook: Path):
    """Tests filtering rows by numeric threshold and string condition."""
    # Filter Salary > 85000
    res_num = AnalysisManager.filter_data(
        sample_workbook,
        range_str="A1:D5",
        column="Salary",
        operator=">",
        value=85000,
    )
    assert res_num["matched_rows"] == 3
    matched_names = [r["Name"] for r in res_num["records"]]
    assert set(matched_names) == {"Alice", "Charlie", "Diana"}

    # Filter Department == 'Engineering'
    res_str = AnalysisManager.filter_data(
        sample_workbook,
        range_str="A1:D5",
        column="Department",
        operator="==",
        value="Engineering",
    )
    assert res_str["matched_rows"] == 2


def test_find_duplicates(sample_workbook: Path):
    """Tests identifying duplicate rows."""
    # Add a duplicate department row
    CellManager.write_range(sample_workbook, "A6", [[5, "Eve", "Engineering", 95000]])

    res = AnalysisManager.find_duplicates(
        sample_workbook,
        range_str="A1:D6",
        columns=["Department"],
    )
    # Alice, Diana, and Eve all have Department == "Engineering"
    assert res["duplicate_count"] == 3


def test_calculate_summary(sample_workbook: Path):
    """Tests statistical summary generation."""
    res = AnalysisManager.calculate_summary(
        sample_workbook,
        range_str="A1:D5",
        columns=["Salary"],
    )
    sal = res["summary"]["Salary"]
    assert sal["count"] == 4
    assert sal["sum"] == 95000 + 80000 + 90000 + 105000
    assert sal["min"] == 80000
    assert sal["max"] == 105000
    assert sal["mean"] == 92500
