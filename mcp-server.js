#!/usr/bin/env node
/**
 * Bundled Node.js STDIO MCP server for DeepSeek Excel Studio.
 * Runs directly on DeepSeek Harness's Node.js runtime with zero Python setup.
 */

import { Server } from '@modelcontextprotocol/sdk/server/index.js'
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js'
import { CallToolRequestSchema, ListToolsRequestSchema } from '@modelcontextprotocol/sdk/types.js'

import {
  calculateSummary,
  clearRange,
  createWorkbook,
  findValue,
  getWorkbookInfo,
  listWorkbooks,
  parseAllowedDirs,
  readRange,
  replaceValue,
  saveEdits,
  sheetOp,
  writeRange,
} from './lib/excel-engine.js'

const allowedDirs = parseAllowedDirs(process.env.EXCEL_MCP_ALLOWED_DIRS || process.cwd())
const autoBackup = String(process.env.EXCEL_MCP_AUTO_BACKUP ?? 'true').toLowerCase() !== 'false'

const TOOLS = [
  {
    name: 'list_workbooks',
    description: 'List local .xlsx and .xlsm Excel workbooks inside allowed directories.',
    inputSchema: {
      type: 'object',
      properties: {
        directory: { type: 'string', description: 'Optional subdirectory to scan.' },
        recursive: { type: 'boolean', description: 'Scan subdirectories recursively (default true).' },
      },
    },
  },
  {
    name: 'create_workbook',
    description: 'Create a new blank .xlsx workbook inside an allowed directory.',
    inputSchema: {
      type: 'object',
      properties: {
        file_path: { type: 'string', description: 'Absolute path for the new .xlsx file.' },
        filename: { type: 'string', description: 'Filename when creating inside default allowed directory.' },
        directory: { type: 'string', description: 'Target allowed directory.' },
        sheet_name: { type: 'string', description: 'Initial worksheet name (default Sheet1).' },
        overwrite: { type: 'boolean', description: 'Allow overwriting an existing file (default false).' },
      },
    },
  },
  {
    name: 'get_workbook_info',
    description: 'Inspect workbook metadata, worksheets, dimensions, VBA status, and lock status.',
    inputSchema: {
      type: 'object',
      required: ['file_path'],
      properties: {
        file_path: { type: 'string' },
      },
    },
  },
  {
    name: 'list_worksheets',
    description: 'List all worksheets in an Excel workbook.',
    inputSchema: {
      type: 'object',
      required: ['file_path'],
      properties: {
        file_path: { type: 'string' },
      },
    },
  },
  {
    name: 'create_worksheet',
    description: 'Add a new worksheet to an existing workbook.',
    inputSchema: {
      type: 'object',
      required: ['file_path', 'sheet_name'],
      properties: {
        file_path: { type: 'string' },
        sheet_name: { type: 'string' },
      },
    },
  },
  {
    name: 'rename_worksheet',
    description: 'Rename an existing worksheet.',
    inputSchema: {
      type: 'object',
      required: ['file_path', 'old_name', 'new_name'],
      properties: {
        file_path: { type: 'string' },
        old_name: { type: 'string' },
        new_name: { type: 'string' },
      },
    },
  },
  {
    name: 'delete_worksheet',
    description: 'Delete a worksheet (destructive: requires confirm=true).',
    inputSchema: {
      type: 'object',
      required: ['file_path', 'sheet_name'],
      properties: {
        file_path: { type: 'string' },
        sheet_name: { type: 'string' },
        confirm: { type: 'boolean', description: 'Must be true to execute deletion.' },
      },
    },
  },
  {
    name: 'read_range',
    description: 'Read cell values from a worksheet range (e.g. A1:E20).',
    inputSchema: {
      type: 'object',
      required: ['file_path'],
      properties: {
        file_path: { type: 'string' },
        sheet_name: { type: 'string' },
        range_str: { type: 'string', description: 'Optional A1-notation range (defaults to used range).' },
      },
    },
  },
  {
    name: 'write_range',
    description: 'Write a 2D array of values or formulas starting at start_cell (creates automatic backup).',
    inputSchema: {
      type: 'object',
      required: ['file_path', 'values'],
      properties: {
        file_path: { type: 'string' },
        sheet_name: { type: 'string' },
        start_cell: { type: 'string', description: 'Top-left cell coordinate (default A1).' },
        values: { type: 'array', items: { type: 'array' } },
      },
    },
  },
  {
    name: 'update_cell',
    description: 'Update a single cell value or formula (e.g. =SUM(B2:B10)).',
    inputSchema: {
      type: 'object',
      required: ['file_path', 'cell', 'value'],
      properties: {
        file_path: { type: 'string' },
        sheet_name: { type: 'string' },
        cell: { type: 'string' },
        value: {},
      },
    },
  },
  {
    name: 'clear_range',
    description: 'Clear cell values in a range (destructive: requires confirm=true).',
    inputSchema: {
      type: 'object',
      required: ['file_path', 'range_str'],
      properties: {
        file_path: { type: 'string' },
        sheet_name: { type: 'string' },
        range_str: { type: 'string' },
        confirm: { type: 'boolean' },
      },
    },
  },
  {
    name: 'find_value',
    description: 'Search for text across a worksheet or entire workbook.',
    inputSchema: {
      type: 'object',
      required: ['file_path', 'query'],
      properties: {
        file_path: { type: 'string' },
        sheet_name: { type: 'string' },
        query: { type: 'string' },
        case_sensitive: { type: 'boolean' },
      },
    },
  },
  {
    name: 'replace_value',
    description: 'Find and replace text in cells (creates automatic backup).',
    inputSchema: {
      type: 'object',
      required: ['file_path', 'find_text', 'replace_text'],
      properties: {
        file_path: { type: 'string' },
        sheet_name: { type: 'string' },
        find_text: { type: 'string' },
        replace_text: { type: 'string' },
        case_sensitive: { type: 'boolean' },
      },
    },
  },
  {
    name: 'format_cells',
    description: 'Apply bold, italic, font color, fill color, alignment, or number format to a cell or range.',
    inputSchema: {
      type: 'object',
      required: ['file_path', 'range_str'],
      properties: {
        file_path: { type: 'string' },
        sheet_name: { type: 'string' },
        range_str: { type: 'string' },
        bold: { type: 'boolean' },
        italic: { type: 'boolean' },
        font_color: { type: 'string', description: 'Hex color like #FFFFFF' },
        fill_color: { type: 'string', description: 'Hex color like #1F4E79' },
        align: { type: 'string', enum: ['left', 'center', 'right'] },
        number_format: { type: 'string', description: 'Excel number format like $#,##0.00' },
      },
    },
  },
  {
    name: 'insert_rows',
    description: 'Insert blank rows at a 1-based row index.',
    inputSchema: {
      type: 'object',
      required: ['file_path', 'row_index'],
      properties: {
        file_path: { type: 'string' },
        sheet_name: { type: 'string' },
        row_index: { type: 'integer' },
        amount: { type: 'integer' },
      },
    },
  },
  {
    name: 'delete_rows',
    description: 'Delete rows starting at a 1-based row index (destructive: requires confirm=true).',
    inputSchema: {
      type: 'object',
      required: ['file_path', 'row_index'],
      properties: {
        file_path: { type: 'string' },
        sheet_name: { type: 'string' },
        row_index: { type: 'integer' },
        amount: { type: 'integer' },
        confirm: { type: 'boolean' },
      },
    },
  },
  {
    name: 'insert_columns',
    description: 'Insert blank columns at a 1-based column index.',
    inputSchema: {
      type: 'object',
      required: ['file_path', 'col_index'],
      properties: {
        file_path: { type: 'string' },
        sheet_name: { type: 'string' },
        col_index: { type: 'integer' },
        amount: { type: 'integer' },
      },
    },
  },
  {
    name: 'delete_columns',
    description: 'Delete columns starting at a 1-based column index (destructive: requires confirm=true).',
    inputSchema: {
      type: 'object',
      required: ['file_path', 'col_index'],
      properties: {
        file_path: { type: 'string' },
        sheet_name: { type: 'string' },
        col_index: { type: 'integer' },
        amount: { type: 'integer' },
        confirm: { type: 'boolean' },
      },
    },
  },
  {
    name: 'calculate_summary',
    description: 'Compute count, sum, mean, median, min, max statistics for numeric columns in a worksheet.',
    inputSchema: {
      type: 'object',
      required: ['file_path'],
      properties: {
        file_path: { type: 'string' },
        sheet_name: { type: 'string' },
        range_str: { type: 'string' },
      },
    },
  },
]

const server = new Server(
  { name: 'deepseek-excel-studio', version: '0.3.0' },
  { capabilities: { tools: {} } }
)

server.setRequestHandler(ListToolsRequestSchema, async () => ({ tools: TOOLS }))

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args = {} } = request.params
  try {
    let result
    switch (name) {
      case 'list_workbooks':
        result = await listWorkbooks({ allowedDirs, ...args })
        break
      case 'create_workbook':
        result = await createWorkbook({ allowedDirs, ...args })
        break
      case 'get_workbook_info':
        result = await getWorkbookInfo({ allowedDirs, ...args })
        break
      case 'list_worksheets': {
        const info = await getWorkbookInfo({ allowedDirs, file_path: args.file_path })
        result = { file_path: info.file_path, worksheets: info.sheets }
        break
      }
      case 'create_worksheet':
        result = await sheetOp({
          allowedDirs,
          autoBackup,
          file_path: args.file_path,
          op: 'create_sheet',
          new_name: args.sheet_name,
        })
        break
      case 'rename_worksheet':
        result = await sheetOp({
          allowedDirs,
          autoBackup,
          file_path: args.file_path,
          sheet_name: args.old_name,
          op: 'rename_sheet',
          new_name: args.new_name,
        })
        break
      case 'delete_worksheet':
        if (!args.confirm) {
          result = {
            confirmation_required: true,
            action: 'delete_worksheet',
            sheet_name: args.sheet_name,
            message: `Pass confirm=true to delete worksheet '${args.sheet_name}'.`,
          }
        } else {
          result = await sheetOp({
            allowedDirs,
            autoBackup,
            file_path: args.file_path,
            sheet_name: args.sheet_name,
            op: 'delete_sheet',
          })
        }
        break
      case 'read_range':
        result = await readRange({ allowedDirs, ...args })
        break
      case 'write_range':
        result = await writeRange({ allowedDirs, autoBackup, ...args })
        break
      case 'update_cell':
        result = await saveEdits({
          allowedDirs,
          autoBackup,
          file_path: args.file_path,
          sheet_name: args.sheet_name,
          edits: [{ cell: args.cell, value: args.value }],
        })
        break
      case 'clear_range':
        result = await clearRange({ allowedDirs, autoBackup, ...args })
        break
      case 'find_value':
        result = await findValue({ allowedDirs, ...args })
        break
      case 'replace_value':
        result = await replaceValue({ allowedDirs, autoBackup, ...args })
        break
      case 'format_cells':
        result = await saveEdits({
          allowedDirs,
          autoBackup,
          file_path: args.file_path,
          sheet_name: args.sheet_name,
          formats: [{ range: args.range_str, ...args }],
        })
        break
      case 'insert_rows':
        result = await sheetOp({
          allowedDirs,
          autoBackup,
          file_path: args.file_path,
          sheet_name: args.sheet_name,
          op: 'insert_rows',
          index: args.row_index,
          amount: args.amount || 1,
        })
        break
      case 'delete_rows':
        if (!args.confirm) {
          result = {
            confirmation_required: true,
            action: 'delete_rows',
            row_index: args.row_index,
            amount: args.amount || 1,
            message: `Pass confirm=true to delete ${args.amount || 1} row(s) starting at row ${args.row_index}.`,
          }
        } else {
          result = await sheetOp({
            allowedDirs,
            autoBackup,
            file_path: args.file_path,
            sheet_name: args.sheet_name,
            op: 'delete_rows',
            index: args.row_index,
            amount: args.amount || 1,
          })
        }
        break
      case 'insert_columns':
        result = await sheetOp({
          allowedDirs,
          autoBackup,
          file_path: args.file_path,
          sheet_name: args.sheet_name,
          op: 'insert_columns',
          index: args.col_index,
          amount: args.amount || 1,
        })
        break
      case 'delete_columns':
        if (!args.confirm) {
          result = {
            confirmation_required: true,
            action: 'delete_columns',
            col_index: args.col_index,
            amount: args.amount || 1,
            message: `Pass confirm=true to delete ${args.amount || 1} column(s) starting at column ${args.col_index}.`,
          }
        } else {
          result = await sheetOp({
            allowedDirs,
            autoBackup,
            file_path: args.file_path,
            sheet_name: args.sheet_name,
            op: 'delete_columns',
            index: args.col_index,
            amount: args.amount || 1,
          })
        }
        break
      case 'calculate_summary':
        result = await calculateSummary({ allowedDirs, ...args })
        break
      default:
        throw new Error(`Unknown tool: ${name}`)
    }

    return {
      content: [{ type: 'text', text: JSON.stringify({ status: 'ok', ...result }, null, 2) }],
    }
  } catch (err) {
    return {
      isError: true,
      content: [
        {
          type: 'text',
          text: JSON.stringify({ status: 'error', code: err.code || 'ERROR', error: err.message }),
        },
      ],
    }
  }
})

const transport = new StdioServerTransport()
await server.connect(transport)
