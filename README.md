# DeepSeek Excel Studio – Production MCP Server

Production-quality Excel automation integration for **DeepSeek Harness Desktop**, powered by the Model Context Protocol (MCP), `openpyxl`, and `pandas`.

Allows DeepSeek AI agents to directly open, inspect, manipulate, analyze, format, and save local `.xlsx` and `.xlsm` workbooks through natural language instructions—without executing unvetted code on your system.

---

## 1. Key Features

- **24 Specialized Tools**: Complete coverage of workbooks, worksheets, ranges, cells, formatting, filtering, sorting, summaries, and charts.
- **Dual Format Support**: Full support for `.xlsx` and `.xlsm` (with VBA macro preservation via `keep_vba=True`).
- **Strict Sandboxing**: Path traversal defense, directory confinement (`EXCEL_MCP_ALLOWED_DIRS`), and file extension verification.
- **Automated Defensive Backups**: Timestamped snapshots saved under `.dsh_backup/` before any destructive modification.
- **Preview & Confirmation Guards**: Dangerous operations (e.g. deleting sheets, rows, columns, or bulk clearing) require explicit confirmation or return an impact preview first.
- **Zero Raw Code Execution**: The agent invokes structured, strongly-typed JSON-RPC tools rather than running arbitrary Python scripts.
- **Lock-File Awareness**: Probes for Microsoft Excel temporary lock files (`~$filename.xlsx`) and warns the user if a file is currently locked by desktop Excel.

---

## 2. Architecture & Modules

```
deepseek-office-plugin/
├── pyproject.toml              # Build & dependency metadata
├── requirements.txt            # Pinned dependency freeze
├── dsh_config_example.yml      # Ready-to-copy DSH cordis.patch.yml snippet
├── excel_mcp/                  # Core MCP Server Package
│   ├── __init__.py             # Version definition
│   ├── __main__.py             # Entry point (python -m excel_mcp)
│   ├── server.py               # MCP Server setup & 24 tool definitions
│   ├── config.py               # Environment configuration loader
│   ├── logging_config.py       # Stderr-only logging (preserves stdout JSON-RPC)
│   ├── core/
│   │   ├── workbook.py         # Open, inspect, save, close, list files
│   │   ├── worksheet.py        # List, create, delete, rename sheets
│   │   ├── cells.py            # Read/write ranges, update, clear, find, replace
│   │   ├── formatting.py       # Typography, fills, alignment, borders, column/row sizes
│   │   ├── analysis.py         # Sort, filter, duplicates, calculate summary
│   │   └── advanced.py         # Insert/delete rows & columns, chart generation
│   ├── security/
│   │   ├── validator.py        # Path canonicalization & sandboxing
│   │   ├── backup.py           # Snapshot & backup manager (.dsh_backup/)
│   │   └── guard.py            # Safety confirmation & preview guard
│   └── utils/
│       └── excel_helpers.py    # Coordinate math (A1 <-> 1,1) & range parsers
└── tests/                      # Automated Test Suite (52 unit & integration tests)
```

---

## 3. Tool Reference (24 MCP Tools)

### A. Workbook Operations
- `list_workbooks(directory?)`: List available workbooks in allowed directories.
- `open_workbook(file_path, read_only?, data_only?)`: Load workbook into session cache.
- `get_workbook_info(file_path)`: Inspect sheets, active sheet, dimensions, and file attributes.
- `save_workbook(file_path, target_path?, overwrite?)`: Save modifications or export to a new file.
- `close_workbook(file_path, save_changes?)`: Release memory cache and close workbook.

### B. Worksheet Operations
- `list_worksheets(file_path)`: Return list of all sheets with visibility and dimensions.
- `create_worksheet(file_path, sheet_name, index?)`: Create a new worksheet.
- `delete_worksheet(file_path, sheet_name, confirm?)`: Delete a sheet (destructive; requires `confirm=True`).
- `rename_worksheet(file_path, old_name, new_name)`: Rename an existing worksheet.

### C. Cell and Range Operations
- `read_range(file_path, range, sheet?, data_only?)`: Read 2D values and formulas from rectangular range.
- `write_range(file_path, start_cell, data, sheet?)`: Write 2D array of values or formulas.
- `update_cell(file_path, cell, value, sheet?)`: Atomically update a single cell value or formula.
- `clear_range(file_path, range, sheet?, confirm?)`: Clear cells in range (>10 cells requires confirmation).
- `find_value(file_path, query, sheet?, match_case?, exact_match?, max_results?)`: Search for matching values.
- `replace_value(file_path, old_value, new_value, sheet?, range?, match_case?, exact_match?)`: Batch search & replace.

### D. Formatting & Dimensions
- `format_cells(file_path, range, sheet?, font?, fill?, alignment?, border?)`: Apply styles to a range.
- `set_column_width(file_path, column, width?, auto_fit?, sheet?)`: Set width or calculate auto-fit width.
- `set_row_height(file_path, row, height, sheet?)`: Set row height in points.
- `apply_number_format(file_path, range, format_code, sheet?)`: Apply number/currency/date/percentage presets.

### E. Tabular Data Analysis (Pandas)
- `sort_range(file_path, range, sort_by_col, ascending?, has_headers?, sheet?)`: Sort rows by column.
- `filter_data(file_path, range, column, operator, value, has_headers?, sheet?)`: Filter rows by condition.
- `find_duplicates(file_path, range, columns?, has_headers?, sheet?)`: Find duplicate records.
- `calculate_summary(file_path, range, columns?, has_headers?, sheet?)`: Compute summary statistics.

### F. Advanced Structure & Charts
- `insert_rows(file_path, row_index, count?, sheet?)`: Insert blank rows at index.
- `delete_rows(file_path, row_index, count?, sheet?, confirm?)`: Delete rows (requires `confirm=True`).
- `insert_columns(file_path, col_index, count?, sheet?)`: Insert blank columns at index.
- `delete_columns(file_path, col_index, count?, sheet?, confirm?)`: Delete columns (requires `confirm=True`).
- `create_chart(file_path, chart_type, data_range, title, target_cell?, categories_range?, sheet?)`: Embed native Excel chart (`bar`, `line`, `pie`, `area`).

---

## 4. Installation & Setup

### Prerequisites
- Python 3.11+
- Git

### Quick Setup (using `uv` or `venv`)
```bash
# Clone the repository
git clone https://github.com/greyblood/deepseek-office-plugin.git
cd deepseek-office-plugin

# Create virtual environment and install dependencies
python -m venv .venv

# On Windows:
.venv\Scripts\activate
# On Linux / macOS:
source .venv/bin/activate

# Install dependencies
pip install -e ".[dev]"
```

---

## 5. DeepSeek Harness Desktop Integration

DeepSeek Harness connects to local MCP servers through its built-in `@deepseek-ai/dsh-mcp-client` plugin.

### Configuration Steps

1. Locate your DeepSeek Harness profile patch file:
   - **Windows**: `%USERPROFILE%\.dsh\profiles\desktop\cordis.patch.yml`
   - **macOS / Linux**: `~/.dsh/profiles/desktop/cordis.patch.yml`

2. Open the file and append the following entry:
```yaml
- insert:
    - id: mcp-excel-studio
      name: '@deepseek-ai/dsh-mcp-client'
      config:
        serverName: excel
        transport: stdio
        command: C:\path\to\deepseek-office-plugin\.venv\Scripts\python.exe
        args:
          - -m
          - excel_mcp
        env:
          EXCEL_MCP_ALLOWED_DIRS: 'C:\Users\YourName\Documents;C:\Projects'
          EXCEL_MCP_AUTO_BACKUP: 'true'
          EXCEL_MCP_REQUIRE_CONFIRMATION: 'true'
          EXCEL_MCP_LOG_LEVEL: 'INFO'
```

3. Launch (or restart) **DeepSeek Harness Desktop**.
4. The tools will now be available to the AI agent with the prefix:
   `mcp__excel__<tool_name>` (e.g. `mcp__excel__read_range`, `mcp__excel__create_chart`).

---

## 6. Manual Testing

### A. Run Automated Tests
```bash
pytest -v
```
All 52 tests verify path sandboxing, backups, confirmation guards, and all 24 MCP tool operations.

### B. Manual Verification via Python CLI
You can verify the server by interacting directly with the core engine:
```python
from excel_mcp.core.cells import CellManager
from excel_mcp.core.analysis import AnalysisManager

# Read from example workbook
data = CellManager.read_range("examples/sample_financials.xlsx", "Summary!A3:D8")
print("Quarterly Revenue Table:", data["values"])

# Calculate statistical summary
summary = AnalysisManager.calculate_summary("examples/sample_financials.xlsx", "Summary!B3:D7")
print("Statistics:", summary["summary"])
```

---

## 7. Security Documentation & Safeguards

1. **Path Traversal Defense**:
   - Every input file path is canonicalized (`resolve()`) to expand symlinks and eliminate `../` traversal tricks.
   - It is checked against the configured `EXCEL_MCP_ALLOWED_DIRS`. Accessing unauthorized directories raises `SecurityError` and halts the operation.
2. **Defensive Snapshots**:
   - Before modifying any file on disk, `BackupManager` creates a copy in `.dsh_backup/<basename>_<timestamp>_<uuid>.bak.xlsx`.
   - In case of an unexpected operation, earlier versions can be retrieved from `.dsh_backup/`.
3. **No Arbitrary Code Execution**:
   - The server only exposes strongly-typed, vetted domain operations.
   - The agent cannot pass raw Python scripts, `exec()`, or `eval()`.
4. **Stderr-Only Logging**:
   - All logging is directed to `sys.stderr`, preventing stdout corruption of the MCP JSON-RPC protocol transport.
   - Cell contents are omitted from log entries to prevent leaking private or financial data.

---

## 8. Troubleshooting Guide

| Issue | Likely Cause | Resolution |
| :--- | :--- | :--- |
| **"Access denied: file is outside permitted directories"** | The requested file is not inside `EXCEL_MCP_ALLOWED_DIRS`. | Add the directory containing your spreadsheets to `EXCEL_MCP_ALLOWED_DIRS` in `cordis.patch.yml`. |
| **"Cannot save: file is open and locked"** | The workbook is currently open and saved in Microsoft Excel. | Save and close the file in Excel, or save changes to a new file name using `save_workbook(target_path="...")`. |
| **"Action requires confirmation"** | Destructive operation called without `confirm=True`. | Review the impact preview returned by the tool and re-run with `confirm=True`. |
| **Tools do not appear in DeepSeek Harness** | DeepSeek Harness was not restarted or Python path is invalid. | Verify the `command:` path in `cordis.patch.yml` points to the Python binary inside your virtual environment. Check DSH logs for connection errors. |
