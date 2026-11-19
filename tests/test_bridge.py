"""Tests for excel_mcp.bridge JSON-over-STDIO editor backend."""

from pathlib import Path
import pytest

from excel_mcp.bridge import dispatch
from excel_mcp.config import config


@pytest.fixture
def allowed_tmp(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    monkeypatch.setattr(config, "allowed_directories", [tmp_path.resolve()])
    monkeypatch.setattr(config, "auto_backup", True)
    return tmp_path


def test_create_list_load_and_save_edits(allowed_tmp: Path) -> None:
    created = dispatch(
        {
            "action": "create_workbook",
            "payload": {"directory": str(allowed_tmp), "filename": "budget.xlsx", "sheet_name": "Q1"},
        }
    )
    assert created["ok"] is True
    file_path = created["data"]["file_path"]
    assert created["data"]["active_sheet"] == "Q1"

    listed = dispatch({"action": "list_workbooks", "payload": {}})
    assert listed["ok"] is True
    assert any(wb["name"] == "budget.xlsx" for wb in listed["data"]["workbooks"])

    saved = dispatch(
        {
            "action": "save_edits",
            "payload": {
                "file_path": file_path,
                "sheet_name": "Q1",
                "edits": [
                    {"cell": "A1", "value": "Item"},
                    {"cell": "B1", "value": "Amount"},
                    {"cell": "A2", "value": "Cloud"},
                    {"cell": "B2", "value": "1250.5"},
                    {"cell": "B3", "value": "=B2*2"},
                ],
                "formats": [
                    {"cell": "A1", "bold": True, "fill_color": "#1F4E79", "font_color": "#FFFFFF"},
                    {"cell": "B2", "number_format": "$#,##0.00", "align": "right"},
                ],
            },
        }
    )
    assert saved["ok"] is True
    data = saved["data"]
    assert data["backup_path"] is not None
    assert data["cells"]["A1"]["display"] == "Item"
    assert data["cells"]["A1"]["style"]["bold"] is True
    assert data["cells"]["A1"]["style"]["fill"] == "#1F4E79"
    assert data["cells"]["B2"]["raw"] == 1250.5
    assert data["cells"]["B2"]["style"]["numberFormat"] == "$#,##0.00"
    assert data["cells"]["B3"]["formula"] == "=B2*2"


def test_sheet_operations_and_summary(allowed_tmp: Path) -> None:
    created = dispatch(
        {
            "action": "create_workbook",
            "payload": {"directory": str(allowed_tmp), "filename": "ops.xlsx"},
        }
    )
    file_path = created["data"]["file_path"]

    add_res = dispatch(
        {
            "action": "sheet_op",
            "payload": {"file_path": file_path, "op": "create_sheet", "new_name": "Revenue"},
        }
    )
    assert add_res["ok"] is True
    assert "Revenue" in add_res["data"]["sheets"]

    ren_res = dispatch(
        {
            "action": "sheet_op",
            "payload": {
                "file_path": file_path,
                "sheet_name": "Revenue",
                "op": "rename_sheet",
                "new_name": "Rev2026",
            },
        }
    )
    assert ren_res["ok"] is True
    assert "Rev2026" in ren_res["data"]["sheets"]

    # Populate numeric column and compute summary
    dispatch(
        {
            "action": "save_edits",
            "payload": {
                "file_path": file_path,
                "sheet_name": "Rev2026",
                "edits": [
                    {"cell": "A1", "value": "Sales"},
                    {"cell": "A2", "value": "100"},
                    {"cell": "A3", "value": "200"},
                ],
            },
        }
    )
    summary = dispatch(
        {
            "action": "summary",
            "payload": {"file_path": file_path, "sheet_name": "Rev2026"},
        }
    )
    assert summary["ok"] is True
    assert summary["data"]["total_rows"] == 2

    del_res = dispatch(
        {
            "action": "sheet_op",
            "payload": {"file_path": file_path, "sheet_name": "Rev2026", "op": "delete_sheet"},
        }
    )
    assert del_res["ok"] is True
    assert "Rev2026" not in del_res["data"]["sheets"]


def test_bridge_blocks_outside_allowed_directory(allowed_tmp: Path) -> None:
    res = dispatch(
        {
            "action": "load_sheet",
            "payload": {"file_path": "/etc/passwd"},
        }
    )
    assert res["ok"] is False
    assert res["code"] == "SECURITY_ERROR"
