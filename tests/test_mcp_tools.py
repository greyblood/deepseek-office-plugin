"""
Unit and integration tests for MCP tool endpoints in excel_mcp.server.
Verifies that all 24 tools execute correctly, return structured responses,
and properly contain exceptions into structured error payloads.
"""

from pathlib import Path
import pytest
from excel_mcp.server import (
    list_workbooks,
    open_workbook,
    get_workbook_info,
    save_workbook,
    close_workbook,
    list_worksheets,
    create_worksheet,
    delete_worksheet,
    rename_worksheet,
    read_range,
    write_range,
    update_cell,
    clear_range,
    find_value,
    replace_value,
    format_cells,
    set_column_width,
    set_row_height,
    apply_number_format,
    sort_range,
    filter_data,
    find_duplicates,
    calculate_summary,
    insert_rows,
    delete_rows,
    insert_columns,
    delete_columns,
    create_chart,
)


def test_mcp_workbook_lifecycle(temp_workspace: Path, sample_workbook: Path):
    """Verifies workbook tool endpoints."""
    # list_workbooks
    wb_list = list_workbooks(str(temp_workspace))
    assert len(wb_list) == 1
    assert wb_list[0]["name"] == "sample.xlsx"

    # get_workbook_info
    info = get_workbook_info(str(sample_workbook))
    assert info["file_name"] == "sample.xlsx"
    assert info["total_sheets"] == 1

    # save_workbook
    new_dest = temp_workspace / "saved_copy.xlsx"
    save_res = save_workbook(str(sample_workbook), target_path=str(new_dest))
    assert save_res["status"] == "success"
    assert new_dest.exists()

    # close_workbook
    close_res = close_workbook(str(sample_workbook))
    assert close_res["status"] == "closed"


def test_mcp_worksheet_tools(sample_workbook: Path):
    """Verifies worksheet tool endpoints."""
    # list_worksheets
    sheets = list_worksheets(str(sample_workbook))
    assert len(sheets) == 1
    assert sheets[0]["name"] == "Sheet1"

    # create_worksheet
    create_res = create_worksheet(str(sample_workbook), sheet_name="Q1_Report")
    assert create_res["status"] == "success"

    # rename_worksheet
    rename_res = rename_worksheet(str(sample_workbook), old_name="Q1_Report", new_name="Q1_Final")
    assert rename_res["status"] == "success"

    # delete_worksheet without confirm -> confirmation_required
    del_unconfirmed = delete_worksheet(str(sample_workbook), sheet_name="Q1_Final", confirm=False)
    assert del_unconfirmed["status"] == "confirmation_required"

    # delete_worksheet with confirm=True
    del_confirmed = delete_worksheet(str(sample_workbook), sheet_name="Q1_Final", confirm=True)
    assert del_confirmed["status"] == "success"


def test_mcp_cell_tools(sample_workbook: Path):
    """Verifies cell and range tool endpoints."""
    # read_range
    res_read = read_range(str(sample_workbook), range="A1:B2")
    assert res_read["values"] == [["ID", "Name"], [1, "Alice"]]

    # write_range
    res_write = write_range(str(sample_workbook), start_cell="A6", data=[[5, "Grace"]])
    assert res_write["status"] == "success"

    # update_cell
    res_update = update_cell(str(sample_workbook), cell="B6", value="Grace Hopper")
    assert res_update["status"] == "success"
    assert res_update["new_value"] == "Grace Hopper"

    # find_value
    find_res = find_value(str(sample_workbook), query="Grace")
    assert len(find_res) == 1
    assert find_res[0]["coordinate"] == "B6"

    # replace_value
    rep_res = replace_value(str(sample_workbook), old_value="Grace Hopper", new_value="Admiral Hopper")
    assert rep_res["status"] == "success"
    assert rep_res["replacements_count"] == 1

    # clear_range
    clear_res = clear_range(str(sample_workbook), range="A6:B6", confirm=True)
    assert clear_res["status"] == "success"


def test_mcp_formatting_tools(sample_workbook: Path):
    """Verifies formatting tool endpoints."""
    fmt_res = format_cells(
        str(sample_workbook),
        range="A1:D1",
        font={"bold": True, "color": "FFFFFF"},
        fill={"bg_color": "002060"},
    )
    assert fmt_res["status"] == "success"

    width_res = set_column_width(str(sample_workbook), column="B", width=22.5)
    assert width_res["status"] == "success"

    height_res = set_row_height(str(sample_workbook), row=1, height=28.0)
    assert height_res["status"] == "success"

    num_res = apply_number_format(str(sample_workbook), range="D2:D5", format_code="currency")
    assert num_res["status"] == "success"


def test_mcp_analysis_tools(sample_workbook: Path):
    """Verifies pandas data analysis tool endpoints."""
    sort_res = sort_range(str(sample_workbook), range="A1:D5", sort_by_col="Salary", ascending=True)
    assert sort_res["status"] == "success"

    filt_res = filter_data(str(sample_workbook), range="A1:D5", column="Department", operator="==", value="Design")
    assert filt_res["matched_rows"] == 1
    assert filt_res["records"][0]["Name"] == "Bob"

    dup_res = find_duplicates(str(sample_workbook), range="A1:D5", columns=["Department"])
    assert dup_res["duplicate_count"] == 2  # Alice and Diana

    sum_res = calculate_summary(str(sample_workbook), range="A1:D5", columns=["Salary"])
    assert sum_res["summary"]["Salary"]["count"] == 4


def test_mcp_advanced_tools(sample_workbook: Path):
    """Verifies advanced structural and chart tools."""
    ins_r = insert_rows(str(sample_workbook), row_index=2, count=1)
    assert ins_r["status"] == "success"

    del_r = delete_rows(str(sample_workbook), row_index=2, count=1, confirm=True)
    assert del_r["status"] == "success"

    ins_c = insert_columns(str(sample_workbook), col_index="C", count=1)
    assert ins_c["status"] == "success"

    del_c = delete_columns(str(sample_workbook), col_index="C", count=1, confirm=True)
    assert del_c["status"] == "success"

    chart_res = create_chart(
        str(sample_workbook),
        chart_type="bar",
        data_range="D1:D5",
        categories_range="B2:B5",
        title="Department Salaries",
        target_cell="G2",
    )
    assert chart_res["status"] == "success"


def test_mcp_error_handling(temp_workspace: Path):
    """Verifies that tool errors are cleanly captured into structured error payloads."""
    # Accessing outside path -> SecurityError formatted as error dict
    outside_file = temp_workspace / ".." / "forbidden.xlsx"
    res = get_workbook_info(str(outside_file))
    assert res["status"] == "error"
    assert res["error_type"] == "SecurityError"

    # Nonexistent file -> FileNotFoundError formatted as error dict
    res_fnf = get_workbook_info(str(temp_workspace / "nonexistent.xlsx"))
    assert res_fnf["status"] == "error"
    assert res_fnf["error_type"] == "FileNotFoundError"
