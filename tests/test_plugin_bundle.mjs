import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import path from 'node:path'
import test from 'node:test'
import vm from 'node:vm'
import { fileURLToPath } from 'node:url'

import { apply as applyHost } from '../index.js'

const ROOT_DIR = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..')

test('package.json declares valid DSH dual-face bundle and client exports', () => {
  const pkg = JSON.parse(readFileSync(path.join(ROOT_DIR, 'package.json'), 'utf8'))
  assert.equal(pkg.name, 'deepseek-office-plugin')
  assert.equal(pkg.dsh?.bundle?.patch, './cordis.patch.yml')
  assert.equal(pkg.dsh?.client?.platform, 'web')
  assert.equal(pkg.exports?.['.'], './index.js')
  assert.equal(pkg.exports?.['./client'], './client.js')
  assert.equal(pkg.exports?.['./package.json'], './package.json')
  assert.ok(pkg.files.includes('client.js'))
})

test('index.js mounts authenticated /api/excel-studio/* routes and talks to Python bridge', async () => {
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
      if (deps.includes('connection')) {
        cb(mockConnCtx)
      }
    },
  }

  applyHost(mockCtx, {
    allowedDirs: [path.join(ROOT_DIR, 'examples')],
  })

  assert.ok(provided.has('excelStudioRuntime'))
  assert.ok(routes.has('/api/excel-studio/workbooks'))
  assert.ok(routes.has('/api/excel-studio/open'))
  assert.ok(routes.has('/api/excel-studio/create'))
  assert.ok(routes.has('/api/excel-studio/save'))
  assert.ok(routes.has('/api/excel-studio/sheet-op'))
  assert.ok(routes.has('/api/excel-studio/summary'))

  // Live round-trip through Python bridge on examples/sample_financials.xlsx
  const listRes = await routes.get('/api/excel-studio/workbooks').fetch(
    new Request('http://localhost/api/excel-studio/workbooks', { method: 'GET' })
  )
  assert.equal(listRes.status, 200)
  const listJson = await listRes.json()
  assert.equal(listJson.ok, true)
  const sample = listJson.data.workbooks.find((w) => w.name === 'sample_financials.xlsx')
  assert.ok(sample, 'Expected sample_financials.xlsx in allowed examples dir')

  const openRes = await routes.get('/api/excel-studio/open').fetch(
    new Request('http://localhost/api/excel-studio/open', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ file_path: sample.path }),
    })
  )
  assert.equal(openRes.status, 200)
  const openJson = await openRes.json()
  assert.equal(openJson.ok, true)
  assert.equal(openJson.data.filename, 'sample_financials.xlsx')
  assert.ok(openJson.data.sheets.length >= 1)
  assert.ok(Object.keys(openJson.data.cells).length > 0)
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

  // Verify icon component renders an SVG element
  const iconEl = sidebarSlot.component({ size: 20 })
  assert.equal(iconEl.type, 'svg')
  assert.equal(iconEl.props.width, 20)
})
