"""
Pytest configuration and shared fixtures for DeepSeek Excel Studio tests.
"""

import shutil
from pathlib import Path
import pytest
import openpyxl
from excel_mcp.config import config, ServerConfig


@pytest.fixture
def temp_workspace(tmp_path: Path) -> Path:
    """Creates an isolated temporary workspace directory and configures allowed_directories."""
    workspace = tmp_path / "workspace"
    workspace.mkdir()
    
    # Configure global config for testing
    old_allowed = config.allowed_directories
    old_auto_backup = config.auto_backup
    config.allowed_directories = [workspace.resolve()]
    config.auto_backup = True
    config.require_confirmation = True

    yield workspace

    # Restore config
    config.allowed_directories = old_allowed
    config.auto_backup = old_auto_backup


@pytest.fixture
def sample_workbook(temp_workspace: Path) -> Path:
    """Creates a standard test Excel workbook with dummy data."""
    file_path = temp_workspace / "sample.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    
    # Headers
    ws.append(["ID", "Name", "Department", "Salary"])
    # Rows
    ws.append([1, "Alice", "Engineering", 95000])
    ws.append([2, "Bob", "Design", 80000])
    ws.append([3, "Charlie", "Product", 90000])
    ws.append([4, "Diana", "Engineering", 105000])
    
    wb.save(file_path)
    wb.close()
    return file_path
