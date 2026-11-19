import { existsSync } from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

export const name = 'deepseek-office-plugin'

const BUNDLE_DIR = path.dirname(fileURLToPath(import.meta.url))

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
}
