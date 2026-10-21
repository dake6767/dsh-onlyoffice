/**
 * Save-back route tests: the plugin's only writer.
 *
 * A fake DSH context collects the prefix route; the workspace is a temp
 * directory; the request body carries the replacement document bytes.
 */
import test from 'node:test'
import assert from 'node:assert/strict'
import { mkdir, readFile, readdir, stat, writeFile } from 'node:fs/promises'
import { basename, join } from 'node:path'

import { apply } from '../../host/index.js'
import { createFakeContext, invoke, makeTempDir } from './helpers.mjs'

const SAVE_URL = '/dsh-onlyoffice/save'

/**
 * Mount the plugin and prepare a workspace with one document.
 * @param {import('node:test').TestContext} t
 */
async function mountWorkspace(t) {
  const scratch = await makeTempDir(t)
  const workspace = join(scratch, 'workspace')
  const docs = join(workspace, 'docs')
  await mkdir(docs, { recursive: true })
  const docPath = join(docs, 'report.docx')
  const original = Buffer.from('original-document-bytes')
  await writeFile(docPath, original)

  const cacheDir = join(scratch, 'cache')
  const fake = createFakeContext()
  apply(fake.ctx, { cacheDir, upstreamFallbacks: [], fetch: async () => new Response(Buffer.alloc(0)) })
  const route = fake.routes.find((candidate) => candidate.path === '/dsh-onlyoffice')
  assert.ok(route, 'the /dsh-onlyoffice prefix route must be registered')
  return { route, workspace, docs, docPath, original, scratch }
}

function saveRequest(workspace, extra = {}, body = Buffer.from('new-document-bytes')) {
  const params = new URLSearchParams({ sessionId: 'session-1', path: 'docs/report.docx', cwd: workspace, ...extra })
  return invoke(routeOf(), { url: `${SAVE_URL}?${params.toString()}`, method: 'POST', body })
}

// `saveRequest` needs the route; tests bind it per call through this shim.
let current = null
function routeOf() {
  assert.ok(current, 'mountWorkspace must run first')
  return current.route
}

test('saves bytes over the existing document atomically', async (t) => {
  const app = await mountWorkspace(t)
  current = app
  const replacement = Buffer.from('PK-new-document-bytes')
  const response = await saveRequest(app.workspace, {}, replacement)

  assert.equal(response.statusCode, 200, 'a valid save succeeds')
  const payload = JSON.parse(response.body)
  assert.equal(payload.ok, true)
  assert.equal(payload.bytes, replacement.length)

  const onDisk = await readFile(app.docPath)
  assert.deepEqual(onDisk, replacement, 'the file is replaced in place')
  assert.equal(basename(app.docPath), 'report.docx', 'the target name is preserved')

  const leftovers = (await readdir(app.docs)).filter((name) => name.includes('.~dsh-save-'))
  assert.deepEqual(leftovers, [], 'no temp file survives the atomic write')
  current = null
})

test('refuses unsupported extensions, empty bodies and absent files', async (t) => {
  const app = await mountWorkspace(t)
  current = app
  await writeFile(join(app.docs, 'notes.txt'), 'keep me')

  const wrongExtension = await invoke(routeOf(), {
    url: `${SAVE_URL}?${new URLSearchParams({ sessionId: 'session-1', path: 'docs/notes.txt', cwd: app.workspace })}`,
    method: 'POST',
    body: Buffer.from('nope'),
  })
  assert.equal(wrongExtension.statusCode, 415)
  assert.equal(JSON.parse(wrongExtension.body).error.code, 'unsupported-extension')

  // Native-engine output only: a legacy/OOXML-alternative source is export-only.
  await writeFile(join(app.docs, 'legacy.odt'), 'odt-bytes')
  const legacy = await invoke(routeOf(), {
    url: `${SAVE_URL}?${new URLSearchParams({ sessionId: 'session-1', path: 'docs/legacy.odt', cwd: app.workspace })}`,
    method: 'POST',
    body: Buffer.from('docx-bytes'),
  })
  assert.equal(legacy.statusCode, 415)
  assert.equal(await readFile(join(app.docs, 'legacy.odt'), 'utf8'), 'odt-bytes')

  const empty = await saveRequest(app.workspace, {}, Buffer.alloc(0))
  assert.equal(empty.statusCode, 400)
  assert.equal(JSON.parse(empty.body).error.code, 'empty-body')

  const missing = await invoke(routeOf(), {
    url: `${SAVE_URL}?${new URLSearchParams({ sessionId: 'session-1', path: 'docs/missing.docx', cwd: app.workspace })}`,
    method: 'POST',
    body: Buffer.from('x'),
  })
  assert.equal(missing.statusCode, 409)

  assert.equal(await readFile(join(app.docs, 'notes.txt'), 'utf8'), 'keep me', 'refused writes change nothing')
  current = null
})

test('refuses a target outside the session workspace', async (t) => {
  const app = await mountWorkspace(t)
  current = app
  const outside = join(app.scratch, 'outside.docx')
  await writeFile(outside, 'outside-original')

  const response = await invoke(routeOf(), {
    url: `${SAVE_URL}?${new URLSearchParams({ sessionId: 'session-1', path: '../outside.docx', cwd: app.workspace })}`,
    method: 'POST',
    body: Buffer.from('escape'),
  })
  assert.equal(response.statusCode, 403)
  assert.equal(JSON.parse(response.body).error.code, 'outside-workspace')
  assert.equal(await readFile(outside, 'utf8'), 'outside-original')
  current = null
})

test('an mtime baseline refuses to clobber a changed file', async (t) => {
  const app = await mountWorkspace(t)
  current = app
  const info = await stat(app.docPath)

  const stale = await saveRequest(app.workspace, { mtimeMs: String(info.mtimeMs - 5000) })
  assert.equal(stale.statusCode, 409)
  assert.equal(JSON.parse(stale.body).error.code, 'conflict')

  const fresh = await saveRequest(app.workspace, { mtimeMs: String(info.mtimeMs) })
  assert.equal(fresh.statusCode, 200, 'the unchanged file accepts the save')
  current = null
})

test('only POST is accepted', async (t) => {
  const app = await mountWorkspace(t)
  current = app
  const response = await invoke(routeOf(), { url: `${SAVE_URL}?sessionId=s&path=docs/report.docx`, method: 'GET' })
  assert.equal(response.statusCode, 405)
  current = null
})

test('reads document bytes for the renderer loading mode', async (t) => {
  const app = await mountWorkspace(t)
  current = app
  const readUrl = (extra = {}) =>
    `/dsh-onlyoffice/read?${new URLSearchParams({ sessionId: 'session-1', path: 'docs/report.docx', cwd: app.workspace, ...extra })}`

  const ok = await invoke(routeOf(), { url: readUrl() })
  assert.equal(ok.statusCode, 200)
  assert.deepEqual(ok.body, app.original, 'the read returns the file bytes')

  const outside = await invoke(routeOf(), {
    url: `/dsh-onlyoffice/read?${new URLSearchParams({ sessionId: 'session-1', path: '../outside.docx', cwd: app.workspace })}`,
  })
  assert.equal(outside.statusCode, 403)

  const wrongMethod = await invoke(routeOf(), { url: readUrl(), method: 'POST', body: Buffer.from('x') })
  assert.equal(wrongMethod.statusCode, 405)
  current = null
})
