# DeepSeek Excel Studio – In-App Excel Editor & MCP Plugin

Production-quality **in-app Excel spreadsheet editor** and **24-tool MCP automation server** for **DeepSeek Harness Desktop**, powered by Cordis, `openpyxl`, and `pandas`.

Open, view, edit, format, and save local `.xlsx` and `.xlsm` workbooks directly inside DeepSeek Harness Desktop's main panel—while also giving your DeepSeek AI agent 24 structured MCP tools to automate the exact same workbooks safely.

---

## 1. Key Features

- **In-App Office / Excel Editor UI (`Excel Studio` sidebar panel)**:
  - Native left-sidebar entry (`sidebar.panellist`) opening a full spreadsheet editor inside DeepSeek Harness (`main` panel).
  - Interactive grid with formula bar (`fx`), inline cell editing, workbook switcher, `+ New` workbook creator, sheet tabs (`+ Sheet`, Rename, Delete), formatting ribbon (Bold, Italic, Alignment, Fill swatches, Number formats), Row/Column insert & delete, and one-click Dataset Summary Stats.
  - Backed by authenticated same-origin Host API routes (`/api/excel-studio/*`) with automatic `.dsh_backup/` snapshots on every save.
- **24 Specialized AI MCP Tools**: Complete coverage of workbooks, worksheets, ranges, cells, formatting, filtering, sorting, summaries, and charts.
- **Dual Format Support**: Full support for `.xlsx` and `.xlsm` (with VBA macro preservation via `keep_vba=True`).
- **Strict Sandboxing**: Path traversal defense, directory confinement (`EXCEL_MCP_ALLOWED_DIRS`), and file extension verification.
- **Automated Defensive Backups**: Timestamped snapshots saved under `.dsh_backup/` before any destructive modification.
- **Preview & Confirmation Guards**: Dangerous operations (e.g. deleting sheets, rows, columns, or bulk clearing) require explicit confirmation or return an impact preview first.
- **Zero Raw Code Execution**: Both the UI editor and AI agent invoke structured, strongly-typed commands rather than running arbitrary Python scripts.
- **Lock-File Awareness**: Probes for Microsoft Excel temporary lock files (`~$filename.xlsx`) and warns if a file is currently locked by desktop Excel.

---

## 2. Architecture & Modules

```
deepseek-office-plugin/
├── package.json                # Dual-face DSH manifest (dsh.bundle + dsh.client)
├── cordis.patch.yml            # DSH bundle patch (mounts host plugin + MCP client)
├── index.js                    # Host Cordis plugin (runtime config + /api/excel-studio/* routes)
├── client.js                   # Zero-build DSH Web UI plugin (sidebar icon + Excel Studio editor)
├── pyproject.toml              # Python build & dependency metadata
├── requirements.txt            # Pinned Python dependencies
├── dsh_config_example.yml      # Manual DSH cordis.patch.yml snippet
├── excel_mcp/                  # Core Python Excel Engine & MCP Server
│   ├── __init__.py             # Version definition
│   ├── __main__.py             # MCP entry point (python -m excel_mcp)
│   ├── bridge.py               # JSON-over-STDIO bridge for the in-app Excel Studio Editor
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
└── tests/                      # Automated Test Suite (55 Python + 3 Node.js tests)
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

DeepSeek Excel Studio ships as an installable **DeepSeek Harness Bundle** (`package.json` with `dsh.bundle.patch` + `cordis.patch.yml`) that connects to the local Python MCP server via `@deepseek-ai/dsh-mcp-client`.

### Method A: Install from DeepSeek Harness Desktop (Recommended)

1. Install the Python dependencies once in your system Python (Python 3.11+):
   ```bash
   pip install mcp openpyxl pandas pydantic
   ```
2. Open **DeepSeek Harness Desktop** → **Plugins** → **Add plugin**.
3. Enter the Git repository URL:
   ```text
   https://github.com/greyblood/deepseek-office-plugin
   ```
4. Click **Install**, then **Enable** the plugin (and restart DeepSeek Harness Desktop if prompted).
5. By default, allowed directories include your `Documents` folder and the bundled `examples/` folder. To customize directories or Python path, add an override to `%USERPROFILE%\.dsh\profiles\desktop\cordis.patch.yml`:
   ```yaml
   - id: excel-studio-runtime
     config:
       pythonCommand: 'C:\path\to\python.exe'
       allowedDirs:
         - 'C:\Users\YourName\Documents'
         - 'D:\Projects'
       autoBackup: true
       requireConfirmation: true
       logLevel: 'INFO'
   ```

### Method B: Manual `cordis.patch.yml` Configuration

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
