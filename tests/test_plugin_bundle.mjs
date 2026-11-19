import assert from 'node:assert/strict'
import { spawn } from 'node:child_process'
import { mkdtempSync, readFileSync, rmSync } from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import test from 'node:test'
import vm from 'node:vm'
import { fileURLToPath } from 'node:url'

import { apply as applyHost } from '../index.js'

const ROOT_DIR = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..')

test('package.json declares valid DSH dual-face bundle, client exports, and bundled dependencies', () => {
  const pkg = JSON.parse(readFileSync(path.join(ROOT_DIR, 'package.json'), 'utf8'))
  assert.equal(pkg.name, 'deepseek-office-plugin')
  assert.equal(pkg.dsh?.bundle?.patch, './cordis.patch.yml')
  assert.equal(pkg.dsh?.client?.platform, 'web')
  assert.equal(pkg.exports?.['.'], './index.js')
  assert.equal(pkg.exports?.['./client'], './client.js')
  assert.equal(pkg.exports?.['./package.json'], './package.json')
  assert.ok(pkg.files.includes('client.js'))
  assert.ok(pkg.files.includes('mcp-server.js'))
  assert.ok(pkg.files.includes('lib'))
  assert.ok(pkg.dependencies?.exceljs)
  assert.ok(pkg.dependencies?.['@modelcontextprotocol/sdk'])
})

test('index.js mounts /api/excel-studio/* routes with bundled in-process ExcelJS engine', async () => {
  const tmpDir = mkdtempSync(path.join(os.tmpdir(), 'dsh-excel-studio-'))
  try {
    const provided = new Map()
    const routes = new Map()

    const mockConnCtx = {
      connection: {
        fetch: {
          register(route) {
            routes.set(route.path, route)
            return () => routes.delete(route.path)
          },
        },
      },
      effect(fn) {
        fn()
      },
    }

    const mockCtx = {
      provide(key, value) {
        provided.set(key, value)
      },
      inject(deps, cb) {
        if (deps.includes('connection')) cb(mockConnCtx)
      },
    }

    applyHost(mockCtx, {
      allowedDirs: [tmpDir, path.join(ROOT_DIR, 'examples')],
    })

    assert.ok(provided.has('excelStudioRuntime'))
    assert.ok(routes.has('/api/excel-studio/workbooks'))
    assert.ok(routes.has('/api/excel-studio/open'))
    assert.ok(routes.has('/api/excel-studio/create'))
    assert.ok(routes.has('/api/excel-studio/save'))
    assert.ok(routes.has('/api/excel-studio/sheet-op'))
    assert.ok(routes.has('/api/excel-studio/summary'))

    // Every DSH fetch route must declare requestBody: 'buffered' to avoid http-bridge.ts HTTP 400
    for (const [routePath, route] of routes.entries()) {
      assert.equal(route.requestBody, 'buffered', `Route ${routePath} must set requestBody: 'buffered'`)
    }
    const wbRoute = routes.get('/api/excel-studio/workbooks')
    assert.ok(wbRoute.methods.includes('GET') && wbRoute.methods.includes('POST'))

    // 0. Verify both GET and POST /api/excel-studio/workbooks succeed on first open
    const listGetRes = await wbRoute.fetch(new Request('http://localhost/api/excel-studio/workbooks', { method: 'GET' }))
    assert.equal(listGetRes.status, 200)
    assert.equal((await listGetRes.json()).ok, true)

    const listPostRes = await wbRoute.fetch(
      new Request('http://localhost/api/excel-studio/workbooks', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: '{}',
      })
    )
    assert.equal(listPostRes.status, 200)
    assert.equal((await listPostRes.json()).ok, true)

    // 1. Create a new workbook in tmpDir
    const createRes = await routes.get('/api/excel-studio/create').fetch(
      new Request('http://localhost/api/excel-studio/create', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ directory: tmpDir, filename: 'budget.xlsx', sheet_name: 'Q1' }),
      })
    )
    assert.equal(createRes.status, 200)
    const createJson = await createRes.json()
    assert.equal(createJson.ok, true)
    assert.equal(createJson.data.active_sheet, 'Q1')
    const createdPath = createJson.data.file_path

    // 2. Save edits + formatting + notes + merges (with automatic .dsh_backup snapshot)
    const saveRes = await routes.get('/api/excel-studio/save').fetch(
      new Request('http://localhost/api/excel-studio/save', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          file_path: createdPath,
          sheet_name: 'Q1',
          edits: [
            { cell: 'A1', value: 'Item' },
            { cell: 'B1', value: 'Amount' },
            { cell: 'A2', value: 'Hosting' },
            { cell: 'B2', value: '450.25' },
            { cell: 'B3', value: '=B2*2' },
          ],
          formats: [
            { cell: 'A1', bold: true, fill_color: '#1F4E79', font_color: '#FFFFFF' },
            { cell: 'B2', number_format: '$#,##0.00', align: 'right' },
          ],
          notes: { A1: 'Header note' },
          merges: ['C1:D1'],
        }),
      })
    )
    assert.equal(saveRes.status, 200)
    const saveJson = await saveRes.json()
    assert.equal(saveJson.ok, true)
    assert.ok(saveJson.data.backup_path)
    assert.equal(saveJson.data.cells.A1.style.bold, true)
    assert.equal(saveJson.data.cells.A1.style.fill, '#1F4E79')
    assert.equal(saveJson.data.cells.B2.raw, 450.25)
    assert.equal(saveJson.data.cells.B3.formula, '=B2*2')

    // 3. Summary stats
    const sumRes = await routes.get('/api/excel-studio/summary').fetch(
      new Request('http://localhost/api/excel-studio/summary', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ file_path: createdPath, sheet_name: 'Q1' }),
      })
    )
    assert.equal(sumRes.status, 200)
    const sumJson = await sumRes.json()
    assert.equal(sumJson.ok, true)
    assert.equal(sumJson.data.summary.Amount.sum, 450.25)

    // 4. Security sandbox blocks paths outside allowed directories
    const blockedRes = await routes.get('/api/excel-studio/open').fetch(
      new Request('http://localhost/api/excel-studio/open', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ file_path: '/etc/passwd' }),
      })
    )
    assert.equal(blockedRes.status, 403)
  } finally {
    rmSync(tmpDir, { recursive: true, force: true })
  }
})

test('bundled Node.js mcp-server.js responds to MCP initialize and tools/list over STDIO', async () => {
  const child = spawn(process.execPath, [path.join(ROOT_DIR, 'mcp-server.js')], {
    cwd: ROOT_DIR,
    env: { ...process.env, EXCEL_MCP_ALLOWED_DIRS: path.join(ROOT_DIR, 'examples') },
    stdio: ['pipe', 'pipe', 'pipe'],
  })

  const responses = []
  let buffer = ''

  const done = new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      child.kill()
      reject(new Error('MCP server timed out'))
    }, 8000)

    child.stdout.setEncoding('utf8')
    child.stdout.on('data', (chunk) => {
      buffer += chunk
      const lines = buffer.split('\n')
      buffer = lines.pop() || ''
      for (const line of lines) {
        if (!line.trim()) continue
        responses.push(JSON.parse(line))
        if (responses.length === 2) {
          clearTimeout(timer)
          child.kill()
          resolve()
        }
      }
    })
  })

  child.stdin.write(
    JSON.stringify({
      jsonrpc: '2.0',
      id: 1,
      method: 'initialize',
      params: {
        protocolVersion: '2024-11-05',
        capabilities: {},
        clientInfo: { name: 'test', version: '1.0.0' },
      },
    }) + '\n'
  )
  child.stdin.write(
    JSON.stringify({
      jsonrpc: '2.0',
      id: 2,
      method: 'tools/list',
      params: {},
    }) + '\n'
  )

  await done
  const listResp = responses.find((r) => r.id === 2)
  assert.ok(listResp?.result?.tools?.length >= 15)
  assert.ok(listResp.result.tools.some((t) => t.name === 'list_workbooks'))
  assert.ok(listResp.result.tools.some((t) => t.name === 'format_cells'))
})

test('client.js registers cleanly via __ModuleLoader__, survives double-eval, and registers sidebar + main slots', () => {
  const clientCode = readFileSync(path.join(ROOT_DIR, 'client.js'), 'utf8')
  const loaded = []

  const sandbox = {
    window: {
      __ModuleLoader__: {
        load(entry) {
          loaded.push(entry)
        },
      },
      addEventListener() {},
      removeEventListener() {},
    },
    document: {
      head: { appendChild() {} },
      getElementById() { return null },
      createElement() { return {} },
    },
    fetch: async () => ({ ok: true, json: async () => ({ ok: true, data: {} }) }),
  }
  sandbox.globalThis = sandbox

  // Evaluate twice to verify zero top-level const/let SyntaxError in combo scripts
  vm.runInNewContext(clientCode, sandbox)
  vm.runInNewContext(clientCode, sandbox)

  assert.equal(loaded.length, 2)
  assert.equal(loaded[0].id, 'deepseek-office-plugin')

  const mockReact = {
    createElement: (type, props, ...children) => ({ type, props, children }),
    useState: (init) => [init, () => {}],
    useEffect: () => {},
    useCallback: (fn) => fn,
  }

  const pluginExports = loaded[0].factory((id) => {
    if (id === 'react') return mockReact
    throw new Error(`Unexpected require: ${id}`)
  })

  assert.equal(pluginExports.name, 'deepseek-office-plugin')
  assert.deepEqual(Array.from(pluginExports.inject), ['slots'])
  assert.equal(typeof pluginExports.apply, 'function')
  assert.equal('default' in pluginExports, false, 'Client module must not export default')

  const registeredSlots = []
  const mockClientCtx = {
    slots: {
      inject(_name, cb) {
        cb()
      },
      register(meta, component) {
        registeredSlots.push({ meta, component })
        return () => {}
      },
    },
  }

  pluginExports.apply(mockClientCtx)

  const mainSlot = registeredSlots.find((s) => s.meta.name === 'main')
  const sidebarSlot = registeredSlots.find((s) => s.meta.name === 'sidebar.panellist')

  assert.ok(mainSlot, 'Expected main slot registration')
  assert.equal(mainSlot.meta.key, 'excel-studio')
  assert.ok(sidebarSlot, 'Expected sidebar.panellist slot registration')
  assert.equal(sidebarSlot.meta.id, 'excel-studio')
  assert.equal(sidebarSlot.meta.label, 'Excel Studio')

  const iconEl = sidebarSlot.component({ size: 20 })
  assert.equal(iconEl.type, 'svg')
  assert.equal(iconEl.props.width, 20)
})
