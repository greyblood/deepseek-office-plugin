import { existsSync } from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

import {
  calculateSummary,
  createWorkbook,
  listWorkbooks,
  loadSheet,
  parseAllowedDirs,
  saveEdits,
  sheetOp,
} from './lib/excel-engine.js'

export const name = 'deepseek-office-plugin'

const BUNDLE_DIR = path.dirname(fileURLToPath(import.meta.url))

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

async function readJsonBody(request) {
  try {
    const text = await request.text()
    return text ? JSON.parse(text) : {}
  } catch {
    return {}
  }
}

async function handleAction(fn) {
  try {
    const data = await fn()
    return Response.json({ ok: true, data }, { status: 200 })
  } catch (err) {
    const code = err?.code || 'OPERATION_ERROR'
    const status = code === 'SECURITY_ERROR' ? 403 : code === 'OPERATION_ERROR' ? 400 : 500
    return Response.json(
      { ok: false, code, error: err?.message || String(err) },
      { status }
    )
  }
}

export function registerEditorRoutes(connCtx, runtime) {
  const fetchRegistry = connCtx?.connection?.fetch
  if (!fetchRegistry || typeof fetchRegistry.register !== 'function') {
    return []
  }

  const allowedDirs = parseAllowedDirs(runtime.allowedDirs)
  const autoBackup = runtime.autoBackup !== 'false'

  const routes = [
    {
      path: '/api/excel-studio/workbooks',
      methods: ['GET'],
      fetch: async (request) => {
        const url = new URL(request.url)
        const directory = url.searchParams.get('directory') || undefined
        return handleAction(() => listWorkbooks({ allowedDirs, directory }))
      },
    },
    {
      path: '/api/excel-studio/open',
      methods: ['POST'],
      requestBody: 'buffered',
      fetch: async (request) => {
        const payload = await readJsonBody(request)
        return handleAction(() => loadSheet({ allowedDirs, ...payload }))
      },
    },
    {
      path: '/api/excel-studio/create',
      methods: ['POST'],
      requestBody: 'buffered',
      fetch: async (request) => {
        const payload = await readJsonBody(request)
        return handleAction(() => createWorkbook({ allowedDirs, ...payload }))
      },
    },
    {
      path: '/api/excel-studio/save',
      methods: ['POST'],
      requestBody: 'buffered',
      fetch: async (request) => {
        const payload = await readJsonBody(request)
        return handleAction(() => saveEdits({ allowedDirs, autoBackup, ...payload }))
      },
    },
    {
      path: '/api/excel-studio/sheet-op',
      methods: ['POST'],
      requestBody: 'buffered',
      fetch: async (request) => {
        const payload = await readJsonBody(request)
        return handleAction(() => sheetOp({ allowedDirs, autoBackup, ...payload }))
      },
    },
    {
      path: '/api/excel-studio/summary',
      methods: ['POST'],
      requestBody: 'buffered',
      fetch: async (request) => {
        const payload = await readJsonBody(request)
        return handleAction(() => calculateSummary({ allowedDirs, ...payload }))
      },
    },
  ]

  return routes.map((route) => fetchRegistry.register(route))
}

export function apply(ctx, config = {}) {
  const runtime = {
    bundleDir: BUNDLE_DIR,
    nodeCommand: process.execPath,
    mcpServerScript: path.join(BUNDLE_DIR, 'mcp-server.js'),
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
