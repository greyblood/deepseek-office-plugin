import { spawn } from 'node:child_process'
import { existsSync } from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

export const name = 'deepseek-office-plugin'

const BUNDLE_DIR = path.dirname(fileURLToPath(import.meta.url))
const BRIDGE_TIMEOUT_MS = 30_000

function resolvePythonCommand(explicitCommand) {
  if (typeof explicitCommand === 'string' && explicitCommand.trim().length > 0) {
    return explicitCommand.trim()
  }

  const envPython = process.env.EXCEL_MCP_PYTHON?.trim()
  if (envPython) {
    return envPython
  }

  const localVenvPython =
    process.platform === 'win32'
      ? path.join(BUNDLE_DIR, '.venv', 'Scripts', 'python.exe')
      : path.join(BUNDLE_DIR, '.venv', 'bin', 'python')

  if (existsSync(localVenvPython)) {
    return localVenvPython
  }

  return process.platform === 'win32' ? 'python' : 'python3'
}

function resolveAllowedDirs(configAllowedDirs) {
  const sep = process.platform === 'win32' ? ';' : ':'

  if (Array.isArray(configAllowedDirs) && configAllowedDirs.length > 0) {
    return configAllowedDirs.map((dir) => String(dir).trim()).filter(Boolean).join(sep)
  }

  if (typeof configAllowedDirs === 'string' && configAllowedDirs.trim().length > 0) {
    return configAllowedDirs.trim()
  }

  const envDirs = process.env.EXCEL_MCP_ALLOWED_DIRS?.trim()
  if (envDirs) {
    return envDirs
  }

  const defaults = []
  const documentsDir = path.join(os.homedir(), 'Documents')
  if (existsSync(documentsDir)) {
    defaults.push(documentsDir)
  }
  const examplesDir = path.join(BUNDLE_DIR, 'examples')
  if (existsSync(examplesDir)) {
    defaults.push(examplesDir)
  }
  if (defaults.length === 0) {
    defaults.push(os.homedir())
  }

  return defaults.join(sep)
}

export function invokePythonBridge(runtime, action, payload = {}) {
  return new Promise((resolve) => {
    const existingPythonPath = process.env.PYTHONPATH
    const sep = process.platform === 'win32' ? ';' : ':'
    const pythonPath = existingPythonPath
      ? `${runtime.bundleDir}${sep}${existingPythonPath}`
      : runtime.bundleDir

    const child = spawn(runtime.pythonCommand, ['-m', 'excel_mcp.bridge'], {
      cwd: runtime.bundleDir,
      shell: false,
      windowsHide: true,
      env: {
        ...process.env,
        PYTHONPATH: pythonPath,
        EXCEL_MCP_ALLOWED_DIRS: runtime.allowedDirs,
        EXCEL_MCP_AUTO_BACKUP: runtime.autoBackup,
        EXCEL_MCP_REQUIRE_CONFIRMATION: runtime.requireConfirmation,
        EXCEL_MCP_LOG_LEVEL: runtime.logLevel,
        PYTHONIOENCODING: 'utf-8',
      },
    })

    let stdout = ''
    let stderr = ''
    let settled = false

    const timer = setTimeout(() => {
      if (!settled) {
        settled = true
        try {
          child.kill()
        } catch {
          // ignore
        }
        resolve({
          ok: false,
          code: 'TIMEOUT',
          error: `Excel bridge timed out after ${BRIDGE_TIMEOUT_MS / 1000}s.`,
        })
      }
    }, BRIDGE_TIMEOUT_MS)

    child.stdout.setEncoding('utf8')
    child.stdout.on('data', (chunk) => {
      stdout += chunk
    })

    child.stderr.setEncoding('utf8')
    child.stderr.on('data', (chunk) => {
      stderr += chunk
    })

    child.on('error', (err) => {
      if (settled) return
      settled = true
      clearTimeout(timer)
      resolve({
        ok: false,
        code: 'SPAWN_ERROR',
        error: `Failed to start Python (${runtime.pythonCommand}): ${err.message}. Install Python 3.11+ and run "pip install -r requirements.txt" inside the plugin directory.`,
      })
    })

    child.on('close', (code) => {
      if (settled) return
      settled = true
      clearTimeout(timer)
      const trimmed = stdout.trim()
      if (!trimmed) {
        resolve({
          ok: false,
          code: 'BRIDGE_ERROR',
          error:
            stderr.trim() ||
            `Python bridge exited with code ${code}. Ensure openpyxl and pydantic are installed for ${runtime.pythonCommand}.`,
        })
        return
      }
      try {
        resolve(JSON.parse(trimmed))
      } catch (err) {
        resolve({
          ok: false,
          code: 'PARSE_ERROR',
          error: `Invalid JSON from Python bridge: ${err.message}`,
        })
      }
    })

    try {
      child.stdin.write(JSON.stringify({ action, payload }))
      child.stdin.end()
    } catch (err) {
      if (!settled) {
        settled = true
        clearTimeout(timer)
        resolve({ ok: false, code: 'STDIN_ERROR', error: err.message })
      }
    }
  })
}

function toHttpResponse(result) {
  if (result && result.ok) {
    return Response.json(result, { status: 200 })
  }
  const code = result?.code
  const status = code === 'SECURITY_ERROR' ? 403 : code === 'OPERATION_ERROR' ? 400 : 500
  return Response.json(result || { ok: false, error: 'Unknown error' }, { status })
}

async function readJsonBody(request) {
  try {
    const text = await request.text()
    return text ? JSON.parse(text) : {}
  } catch {
    return {}
  }
}

export function registerEditorRoutes(connCtx, runtime) {
  const fetchRegistry = connCtx?.connection?.fetch
  if (!fetchRegistry || typeof fetchRegistry.register !== 'function') {
    return []
  }

  const routes = [
    {
      path: '/api/excel-studio/workbooks',
      methods: ['GET'],
      fetch: async (request) => {
        const url = new URL(request.url)
        const directory = url.searchParams.get('directory') || undefined
        const result = await invokePythonBridge(runtime, 'list_workbooks', { directory })
        return toHttpResponse(result)
      },
    },
    {
      path: '/api/excel-studio/open',
      methods: ['POST'],
      requestBody: 'buffered',
      fetch: async (request) => {
        const payload = await readJsonBody(request)
        const result = await invokePythonBridge(runtime, 'load_sheet', payload)
        return toHttpResponse(result)
      },
    },
    {
      path: '/api/excel-studio/create',
      methods: ['POST'],
      requestBody: 'buffered',
      fetch: async (request) => {
        const payload = await readJsonBody(request)
        const result = await invokePythonBridge(runtime, 'create_workbook', payload)
        return toHttpResponse(result)
      },
    },
    {
      path: '/api/excel-studio/save',
      methods: ['POST'],
      requestBody: 'buffered',
      fetch: async (request) => {
        const payload = await readJsonBody(request)
        const result = await invokePythonBridge(runtime, 'save_edits', payload)
        return toHttpResponse(result)
      },
    },
    {
      path: '/api/excel-studio/sheet-op',
      methods: ['POST'],
      requestBody: 'buffered',
      fetch: async (request) => {
        const payload = await readJsonBody(request)
        const result = await invokePythonBridge(runtime, 'sheet_op', payload)
        return toHttpResponse(result)
      },
    },
    {
      path: '/api/excel-studio/summary',
      methods: ['POST'],
      requestBody: 'buffered',
      fetch: async (request) => {
        const payload = await readJsonBody(request)
        const result = await invokePythonBridge(runtime, 'summary', payload)
        return toHttpResponse(result)
      },
    },
  ]

  return routes.map((route) => fetchRegistry.register(route))
}

export function apply(ctx, config = {}) {
  const runtime = {
    bundleDir: BUNDLE_DIR,
    pythonCommand: resolvePythonCommand(config.pythonCommand),
    allowedDirs: resolveAllowedDirs(config.allowedDirs),
    autoBackup: String(config.autoBackup ?? true),
    requireConfirmation: String(config.requireConfirmation ?? true),
    logLevel: String(config.logLevel ?? 'INFO'),
  }

  ctx.provide('excelStudioRuntime', runtime)

  // Optionally mount authenticated HTTP endpoints for the in-app Excel Studio Editor
  // whenever the host profile provides the `connection` service (Web / Desktop GUI).
  if (typeof ctx.inject === 'function') {
    ctx.inject(['connection'], (connCtx) => {
      if (typeof connCtx.effect === 'function') {
        connCtx.effect(() => {
          const disposers = registerEditorRoutes(connCtx, runtime)
          return () => {
            for (const dispose of disposers) {
              if (typeof dispose === 'function') dispose()
            }
          }
        }, 'deepseek-office-plugin: editor routes')
      } else {
        registerEditorRoutes(connCtx, runtime)
      }
    })
  }
}
