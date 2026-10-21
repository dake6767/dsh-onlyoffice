/**
 * Route-level tests for host/index.js: a fake DSH context collects the one
 * prefix route, an injectable fetch stands in for the upstream, and the cache
 * lives in a temp directory.
 */
import test from 'node:test'
import assert from 'node:assert/strict'
import { readFile, stat, writeFile } from 'node:fs/promises'
import { join } from 'node:path'

import { apply, inject } from '../../host/index.js'
import {
  createFakeContext,
  invoke,
  isDirectory,
  listFiles,
  listTempFiles,
  makeFetch,
  makeTempDir,
  waitFor,
} from './helpers.mjs'

const VERSION = '9.4.0-develop'
const SDK_JS = '/dsh-onlyoffice/sdk/onlyoffice/9.4.0-develop/web-apps/apps/api/documents/api.js'

/**
 * Mount the plugin on a fake ctx with an isolated cache dir.
 * @param {import('node:test').TestContext} t
 * @param {object} [options]
 */
async function mount(t, options = {}) {
  const scratch = await makeTempDir(t)
  const cacheDir = options.cacheDir ?? join(scratch, 'cache')
  const fetchImpl = options.fetch ?? makeFetch(() => ({ status: 200, body: Buffer.alloc(0) }))
  const fake = createFakeContext({ rejection: options.rejection })
  apply(fake.ctx, {
    version: options.version ?? VERSION,
    upstreamBase: options.upstreamBase ?? 'http://upstream.test/packages',
    // Isolate tests from the real default mirrors: failover tests opt in.
    upstreamFallbacks: options.upstreamFallbacks ?? [],
    cacheDir,
    fetch: fetchImpl,
    ...options.config,
  })
  const route = fake.routes.find((candidate) => candidate.path === '/dsh-onlyoffice')
  assert.ok(route, 'the /dsh-onlyoffice prefix route must be registered')
  return { ...fake, route, cacheDir, fetchImpl, scratch }
}

function getStatus(app) {
  return invoke(app.route, { url: '/dsh-onlyoffice/status' })
}

test('exports the documented plugin face', () => {
  assert.deepEqual(inject, ['webServer', 'connection'])
  assert.equal(typeof apply, 'function')
})

test('registers exactly one prefix route plus owned effects', async (t) => {
  const app = await mount(t)
  assert.equal(app.routes.length, 1)
  assert.equal(app.routes[0].kind, 'prefix')
  assert.equal(app.routes[0].path, '/dsh-onlyoffice')
  assert.equal(typeof app.routes[0].handler, 'function')

  const labels = app.effects.map((effect) => effect.label).sort()
  assert.deepEqual(labels, ['@dake6767/dsh-onlyoffice: cache directory', '@dake6767/dsh-onlyoffice: routes'])

  // The routing effect's disposer must unregister the route on unload.
  const routesEffect = app.effects.find((effect) => effect.label === '@dake6767/dsh-onlyoffice: routes')
  assert.ok(routesEffect, 'routes effect must exist')
  assert.equal(typeof routesEffect.dispose, 'function')
  routesEffect.dispose()
  assert.equal(app.routes.length, 0)
})

test('the cache directory is created by its own effect', async (t) => {
  const app = await mount(t, { cacheDir: join(await makeTempDir(t), 'deep', 'nested', 'cache') })
  await waitFor(() => isDirectory(app.cacheDir))
  assert.equal((await stat(app.cacheDir)).isDirectory(), true)
})

test('status returns exactly the documented shape', async (t) => {
  const app = await mount(t)
  const res = await getStatus(app)
  assert.equal(res.statusCode, 200)
  assert.equal(res.getHeader('content-type'), 'application/json; charset=utf-8')
  assert.equal(res.getHeader('cache-control'), 'no-store')

  const json = res.json()
  assert.deepEqual(Object.keys(json).sort(), [
    'bytes', 'cacheDir', 'cacheable', 'files', 'lastError', 'ok', 'sdkBaseUrl', 'upstreamBase', 'upstreams', 'version',
  ])
  assert.deepEqual(json, {
    ok: true,
    version: VERSION,
    sdkBaseUrl: '/dsh-onlyoffice/sdk',
    cacheDir: app.cacheDir,
    upstreamBase: 'http://upstream.test/packages',
    upstreams: ['http://upstream.test/packages'],
    files: 0,
    bytes: 0,
    lastError: null,
    cacheable: true,
  })
})

test('cache miss fetches upstream, streams exact bytes and caches them atomically', async (t) => {
  // Deliberately not valid UTF-8: binary safety must survive the proxy.
  const payload = Buffer.from([0xef, 0xbb, 0xbf, 0x2f, 0x2f, 0xc3, 0xa9, 0x00, 0x01, 0xff, 0xfe])
  const fetchImpl = makeFetch(() => ({ status: 200, body: payload }))
  const app = await mount(t, { fetch: fetchImpl })

  const res = await invoke(app.route, { url: SDK_JS })
  assert.equal(res.statusCode, 200)
  assert.equal(res.getHeader('content-type'), 'text/javascript')
  assert.equal(res.getHeader('cache-control'), 'public, max-age=31536000, immutable')
  assert.equal(res.getHeader('x-content-type-options'), 'nosniff')
  assert.equal(res.getHeader('content-length'), String(payload.length))
  assert.deepEqual(res.body, payload)

  assert.equal(fetchImpl.calls.length, 1)
  assert.equal(
    fetchImpl.calls[0].url,
    'http://upstream.test/packages/onlyoffice/9.4.0-develop/web-apps/apps/api/documents/api.js',
  )
  assert.equal(fetchImpl.calls[0].init.headers['user-agent'], '@dake6767/dsh-onlyoffice/0.1')
  assert.equal(typeof fetchImpl.calls[0].init.signal.aborted, 'boolean')

  const cacheFile = join(app.cacheDir, VERSION, 'web-apps/apps/api/documents/api.js')
  assert.deepEqual(await readFile(cacheFile), payload)
  assert.deepEqual(await listTempFiles(app.cacheDir), [])

  assert.ok(
    app.logs.some((log) => log.level === 'info' && String(log.args[0]).includes('cache miss')),
    'a miss must be logged',
  )

  const status = (await getStatus(app)).json()
  assert.equal(status.files, 1)
  assert.equal(status.bytes, payload.length)
  assert.equal(status.lastError, null)
})

test('cache hit is served from disk without contacting upstream', async (t) => {
  const payload = Buffer.from('cached sdk asset')
  const fetchImpl = makeFetch((url, callNumber) => {
    if (callNumber > 1) throw new Error(`unexpected second upstream fetch: ${url}`)
    return { status: 200, body: payload }
  })
  const app = await mount(t, { fetch: fetchImpl })

  const first = await invoke(app.route, { url: SDK_JS })
  assert.deepEqual(first.body, payload)
  assert.equal(fetchImpl.calls.length, 1)

  const second = await invoke(app.route, { url: SDK_JS })
  assert.equal(second.statusCode, 200)
  assert.equal(second.getHeader('content-length'), String(payload.length))
  assert.deepEqual(second.body, payload)
  assert.equal(fetchImpl.calls.length, 1, 'cache hit must not fetch')
  assert.ok(
    app.logs.some((log) => log.level === 'info' && String(log.args[0]).includes('cache hit')),
    'a cache hit must be logged',
  )
})

test('nested asset paths map to nested cache paths', async (t) => {
  const fetchImpl = makeFetch(() => ({ status: 200, body: Buffer.from('wasm-bytes') }))
  const app = await mount(t, { fetch: fetchImpl })

  const res = await invoke(app.route, {
    url: '/dsh-onlyoffice/sdk/onlyoffice/9.4.0-develop/sdkjs/word/sdk-all.js',
  })
  assert.equal(res.statusCode, 200)
  assert.equal(res.getHeader('content-type'), 'text/javascript')
  assert.equal(fetchImpl.calls[0].url, 'http://upstream.test/packages/onlyoffice/9.4.0-develop/sdkjs/word/sdk-all.js')
  assert.deepEqual(await listFiles(app.cacheDir), [join(VERSION, 'sdkjs', 'word', 'sdk-all.js')])

  const wasm = await invoke(app.route, {
    url: '/dsh-onlyoffice/sdk/onlyoffice/9.4.0-develop/x2t/x2t.wasm',
  })
  assert.equal(wasm.statusCode, 200)
  assert.equal(wasm.getHeader('content-type'), 'application/wasm')
})

test('a definitive 404 is remembered briefly, a transient failure is not', async (t) => {
  const good = Buffer.from('good bytes')
  let mode = '404'
  const fetchImpl = makeFetch(() => {
    if (mode === '404') return { status: 404, body: Buffer.from('not here') }
    if (mode === 'network') return { error: new TypeError('connect ECONNRESET') }
    return { status: 200, body: good }
  })
  const app = await mount(t, { fetch: fetchImpl })

  // Every configured mirror answered 404: a definitive miss.
  const missing = await invoke(app.route, { url: SDK_JS })
  assert.equal(missing.statusCode, 404)
  const missingJson = missing.json()
  assert.equal(missingJson.ok, false)
  assert.match(missingJson.error.message, /HTTP 404/)
  assert.deepEqual(await listFiles(app.cacheDir), [], 'a failed fetch must not create a cache file')
  assert.equal((await getStatus(app)).json().files, 0)

  const remembered = await invoke(app.route, { url: SDK_JS })
  assert.equal(remembered.statusCode, 404)
  assert.equal(fetchImpl.calls.length, 1, 'a remembered 404 costs no upstream call')

  // A TRANSIENT failure (network) must not be remembered: the next attempt
  // reaches the upstream again and succeeds. A different asset is used, so the
  // remembered 404 above cannot mask this.
  const otherAsset = '/dsh-onlyoffice/sdk/onlyoffice/9.4.0-develop/web-apps/apps/documenteditor/main/app.js'
  mode = 'network'
  const transient = await invoke(app.route, { url: otherAsset })
  assert.equal(transient.statusCode, 502)
  mode = 'ok'
  const retried = await invoke(app.route, { url: otherAsset })
  assert.equal(retried.statusCode, 200)
  assert.deepEqual(retried.body, good)
  const statusAfterRetry = (await getStatus(app)).json()
  assert.equal(statusAfterRetry.lastError, null)
  assert.equal(statusAfterRetry.files, 1)
})

test('upstream 500 is passed through as 500', async (t) => {
  const fetchImpl = makeFetch(() => ({ status: 500, body: Buffer.from('boom') }))
  const app = await mount(t, { fetch: fetchImpl })
  const res = await invoke(app.route, { url: SDK_JS })
  assert.equal(res.statusCode, 500)
  const json = res.json()
  assert.equal(json.error.code, 'upstream-error')
  assert.match(json.error.message, /HTTP 500/)
  assert.deepEqual(await listFiles(app.cacheDir), [])
})

test('network failure becomes 502 and lastError', async (t) => {
  const fetchImpl = makeFetch(() => ({ error: new TypeError('connect ECONNREFUSED') }))
  const app = await mount(t, { fetch: fetchImpl })
  const res = await invoke(app.route, { url: SDK_JS })
  assert.equal(res.statusCode, 502)
  const json = res.json()
  assert.equal(json.error.code, 'upstream-error')
  assert.match(json.error.message, /network error/)
  assert.deepEqual(await listFiles(app.cacheDir), [])

  const status = (await getStatus(app)).json()
  assert.match(status.lastError, /network error/)
})

test('upstream timeout becomes 504', async (t) => {
  const timeoutError = new Error('timed out')
  timeoutError.name = 'TimeoutError'
  const fetchImpl = makeFetch(() => ({ error: timeoutError }))
  const app = await mount(t, { fetch: fetchImpl })
  const res = await invoke(app.route, { url: SDK_JS })
  assert.equal(res.statusCode, 504)
  assert.match(res.json().error.message, /timeout/i)
  assert.match((await getStatus(app)).json().lastError, /timeout/i)
})

test('concurrent misses for one asset coalesce into a single upstream fetch', async (t) => {
  const payload = Buffer.from('one download, two clients')
  const fetchImpl = makeFetch(() => ({ status: 200, body: payload, delayMs: 30 }))
  const app = await mount(t, { fetch: fetchImpl })

  const [first, second] = await Promise.all([
    invoke(app.route, { url: SDK_JS }),
    invoke(app.route, { url: SDK_JS }),
  ])
  assert.equal(first.statusCode, 200)
  assert.equal(second.statusCode, 200)
  assert.deepEqual(first.body, payload)
  assert.deepEqual(second.body, payload)
  assert.equal(fetchImpl.calls.length, 1, 'in-flight dedupe must collapse the two misses')
  assert.deepEqual(await listFiles(app.cacheDir), [join(VERSION, 'web-apps/apps/api/documents/api.js')])
  assert.deepEqual(await listTempFiles(app.cacheDir), [])
})

test('traversal and unsafe asset paths are rejected before any fetch', async (t) => {
  const fetchImpl = makeFetch(() => ({ status: 200, body: Buffer.from('never') }))
  const app = await mount(t, { fetch: fetchImpl })

  // Dot segments — literal or percent-encoded — are collapsed by WHATWG URL
  // normalization BEFORE the handler runs, leaving a path that simply names no
  // asset. Traversal is impossible through this route; assert it never fetches.
  const normalized = [
    '/dsh-onlyoffice/sdk/onlyoffice/9.4.0-develop/%2e%2e/secret.js',
    '/dsh-onlyoffice/sdk/onlyoffice/9.4.0-develop/%2E%2E/secret.js',
    '/dsh-onlyoffice/sdk/onlyoffice/9.4.0-develop/a/%2e%2e/%2e%2e/etc/passwd',
    '/dsh-onlyoffice/sdk/onlyoffice/9.4.0-develop/../etc/passwd',
  ]
  for (const url of normalized) {
    const res = await invoke(app.route, { url })
    assert.equal(res.statusCode, 404, url)
    assert.equal(res.json().ok, false, url)
  }

  // Segments that survive URL parsing yet are unsafe must be rejected 400 by
  // the helper's own validation.
  const unsafe = [
    '/dsh-onlyoffice/sdk/onlyoffice/9.4.0-develop/a%2Fb.js',
    '/dsh-onlyoffice/sdk/onlyoffice/9.4.0-develop/a%5Cb.js',
    '/dsh-onlyoffice/sdk/onlyoffice/9.4.0-develop/a%00b.js',
    '/dsh-onlyoffice/sdk/onlyoffice/9.4.0-develop//etc/passwd',
    '/dsh-onlyoffice/sdk/onlyoffice/9.4.0-develop/C%3A/evil.js',
  ]
  for (const url of unsafe) {
    const res = await invoke(app.route, { url })
    assert.equal(res.statusCode, 400, url)
    assert.equal(res.json().error.code, 'bad-path', url)
  }
  assert.equal(fetchImpl.calls.length, 0)
  assert.deepEqual(await listFiles(app.cacheDir), [])
})

test('requests for a different SDK version are refused', async (t) => {
  const fetchImpl = makeFetch(() => ({ status: 200, body: Buffer.from('never') }))
  const app = await mount(t, { fetch: fetchImpl })
  const res = await invoke(app.route, { url: '/dsh-onlyoffice/sdk/onlyoffice/8.2.0/sdkjs/word/sdk-all.js' })
  assert.equal(res.statusCode, 404)
  assert.equal(res.json().error.code, 'version-mismatch')
  assert.equal(fetchImpl.calls.length, 0)
})

test('configured version and upstream base drive both status and fetches', async (t) => {
  const fetchImpl = makeFetch(() => ({ status: 200, body: Buffer.from('v') }))
  const app = await mount(t, {
    version: '9.5.0-test',
    upstreamBase: 'https://mirror.test/packages/',
    fetch: fetchImpl,
  })
  const status = (await getStatus(app)).json()
  assert.equal(status.version, '9.5.0-test')
  assert.equal(status.upstreamBase, 'https://mirror.test/packages')

  const res = await invoke(app.route, { url: '/dsh-onlyoffice/sdk/onlyoffice/9.5.0-test/x2t/x2t.js' })
  assert.equal(res.statusCode, 200)
  assert.equal(fetchImpl.calls[0].url, 'https://mirror.test/packages/onlyoffice/9.5.0-test/x2t/x2t.js')
  assert.ok(await isDirectory(join(app.cacheDir, '9.5.0-test')))
})

test('unknown routes are 404 and non-GET methods are 405', async (t) => {
  const app = await mount(t)

  for (const url of ['/dsh-onlyoffice/', '/dsh-onlyoffice/nope', '/dsh-onlyoffice/sdk', '/dsh-onlyoffice/statuses']) {
    const res = await invoke(app.route, { url })
    assert.equal(res.statusCode, 404, url)
    assert.equal(res.json().error.code, 'not-found', url)
  }

  for (const url of ['/dsh-onlyoffice/status', SDK_JS]) {
    const res = await invoke(app.route, { url, method: 'POST' })
    assert.equal(res.statusCode, 405, url)
    assert.equal(res.getHeader('allow'), 'GET, HEAD')
    assert.equal(res.json().error.code, 'method-error')
  }
  const put = await invoke(app.route, { url: '/dsh-onlyoffice/status', method: 'PUT' })
  assert.equal(put.statusCode, 405)
})

test('the connection fence rejects unfenced requests before anything else', async (t) => {
  const fetchImpl = makeFetch(() => ({ status: 200, body: Buffer.from('never') }))
  const app = await mount(t, { fetch: fetchImpl, rejection: { code: 'unauthenticated' } })

  for (const url of ['/dsh-onlyoffice/status', SDK_JS]) {
    const res = await invoke(app.route, { url })
    assert.equal(res.statusCode, 403)
    assert.equal(res.text, 'forbidden')
    assert.equal(res.getHeader('content-type'), 'text/plain; charset=utf-8')
  }
  assert.equal(fetchImpl.calls.length, 0)
})

test('status counts every cached file and its bytes', async (t) => {
  const assets = [
    ['a/one.js', 5],
    ['b/two.ttf', 9],
    ['c/three.wasm', 12],
  ]
  const fetchImpl = makeFetch((url) => {
    const hit = assets.find(([relative]) => url.endsWith(`/${relative}`))
    return { status: 200, body: Buffer.alloc(hit === undefined ? 1 : hit[1], 0x41) }
  })
  const app = await mount(t, { fetch: fetchImpl })
  for (const [relative, size] of assets) {
    const res = await invoke(app.route, { url: `/dsh-onlyoffice/sdk/onlyoffice/${VERSION}/${relative}` })
    assert.equal(res.statusCode, 200, relative)
    assert.equal(res.body.length, size)
  }
  const status = (await getStatus(app)).json()
  assert.equal(status.files, 3)
  assert.equal(status.bytes, 26)
  assert.equal(status.cacheable, true)
})

test('HEAD requests carry headers and no body', async (t) => {
  const payload = Buffer.alloc(23, 0x7a)
  const fetchImpl = makeFetch(() => ({ status: 200, body: payload }))
  const app = await mount(t, { fetch: fetchImpl })

  // Prime the cache, then HEAD the asset.
  await invoke(app.route, { url: SDK_JS })
  const headAsset = await invoke(app.route, { url: SDK_JS, method: 'HEAD' })
  assert.equal(headAsset.statusCode, 200)
  assert.equal(headAsset.getHeader('content-length'), String(payload.length))
  assert.equal(headAsset.body.length, 0)

  const headStatus = await invoke(app.route, { url: '/dsh-onlyoffice/status', method: 'HEAD' })
  assert.equal(headStatus.statusCode, 200)
  assert.ok(Number(headStatus.getHeader('content-length')) > 0)
  assert.equal(headStatus.body.length, 0)
  assert.equal(fetchImpl.calls.length, 1)
})

test('a cache-directory failure surfaces as 500 with lastError', async (t) => {
  const scratch = await makeTempDir(t)
  const filePath = join(scratch, 'not-a-directory')
  await writeFile(filePath, 'x')
  const fetchImpl = makeFetch(() => ({ status: 200, body: Buffer.from('never') }))
  const app = await mount(t, { cacheDir: join(filePath, 'cache'), fetch: fetchImpl })

  const res = await invoke(app.route, { url: SDK_JS })
  assert.equal(res.statusCode, 500)
  assert.equal(res.json().error.code, 'internal-error')
  const status = (await getStatus(app)).json()
  assert.ok(status.lastError !== null)
  assert.equal(status.files, 0)
})

test('fetched file bytes are the exact upstream bytes (no text coercion)', async (t) => {
  const payload = Buffer.alloc(4096)
  for (let i = 0; i < payload.length; i += 1) payload[i] = i % 256
  const fetchImpl = makeFetch(() => ({ status: 200, body: payload }))
  const app = await mount(t, { fetch: fetchImpl })
  const res = await invoke(app.route, { url: '/dsh-onlyoffice/sdk/onlyoffice/9.4.0-develop/fonts/font-no-ext' })
  assert.equal(res.statusCode, 200)
  assert.equal(res.getHeader('content-type'), 'application/octet-stream')
  assert.equal(res.getHeader('content-length'), '4096')
  assert.deepEqual(res.body, payload)
  const cached = await readFile(join(app.cacheDir, VERSION, 'fonts', 'font-no-ext'))
  assert.deepEqual(cached, payload)
})

// ── Multi-upstream failover ───────────────────────────────────────────────
const MIRROR_A = 'http://mirror-a.test/packages'
const MIRROR_B = 'http://mirror-b.test/packages'
const MIRROR_C = 'http://mirror-c.test/packages'

/** Mount with three ordered mirrors (A = primary). */
function mountMirrors(t, fetchImpl) {
  return mount(t, {
    upstreamBase: MIRROR_A,
    upstreamFallbacks: [MIRROR_B, MIRROR_C],
    fetch: fetchImpl,
  })
}

test('a hit on the first mirror contacts no other mirror', async (t) => {
  const payload = Buffer.from('mirror-a bytes')
  const fetchImpl = makeFetch(() => ({ status: 200, body: payload }))
  const app = await mountMirrors(t, fetchImpl)

  const res = await invoke(app.route, { url: SDK_JS })
  assert.equal(res.statusCode, 200)
  assert.deepEqual(res.body, payload)
  assert.equal(fetchImpl.calls.length, 1)
  assert.ok(fetchImpl.calls[0].url.startsWith(`${MIRROR_A}/onlyoffice/`))
  assert.equal((await getStatus(app)).json().lastError, null)
})

test('a network error on one mirror fails over to the next', async (t) => {
  const payload = Buffer.from('served by mirror-b')
  const fetchImpl = makeFetch((url) => {
    if (url.startsWith(MIRROR_A)) throw new TypeError('connect ECONNRESET')
    if (url.startsWith(MIRROR_B)) return { status: 200, body: payload }
    throw new Error(`mirror C must not be contacted: ${url}`)
  })
  const app = await mountMirrors(t, fetchImpl)

  const res = await invoke(app.route, { url: SDK_JS })
  assert.equal(res.statusCode, 200)
  assert.deepEqual(res.body, payload)
  assert.equal(fetchImpl.calls.length, 2)
  assert.ok(fetchImpl.calls[0].url.startsWith(MIRROR_A))
  assert.ok(fetchImpl.calls[1].url.startsWith(MIRROR_B))
  assert.equal((await getStatus(app)).json().lastError, null, 'a recovered mirror must clear lastError')
  assert.ok(
    app.logs.some((log) => String(log.args[0]).includes(`mirror fail ${MIRROR_A}`) && String(log.args[0]).includes('-> next')),
    'a failed mirror must be logged with the [dsh-onlyoffice] prefix and "-> next"',
  )
})

test('a 404 on one mirror falls through to the next', async (t) => {
  const payload = Buffer.from('mirror-b has it')
  const fetchImpl = makeFetch((url) => {
    if (url.startsWith(MIRROR_A)) return { status: 404, body: Buffer.from('missing') }
    if (url.startsWith(MIRROR_B)) return { status: 200, body: payload }
    throw new Error(`mirror C must not be contacted: ${url}`)
  })
  const app = await mountMirrors(t, fetchImpl)

  const res = await invoke(app.route, { url: SDK_JS })
  assert.equal(res.statusCode, 200)
  assert.deepEqual(res.body, payload)
  assert.equal(fetchImpl.calls.length, 2)
  assert.equal((await getStatus(app)).json().lastError, null)
})

test('a non-404 upstream error also falls through to the next mirror', async (t) => {
  const payload = Buffer.from('mirror-b survives a 500')
  const fetchImpl = makeFetch((url) => {
    if (url.startsWith(MIRROR_A)) return { status: 500, body: Buffer.from('boom') }
    if (url.startsWith(MIRROR_B)) return { status: 200, body: payload }
    throw new Error(`mirror C must not be contacted: ${url}`)
  })
  const app = await mountMirrors(t, fetchImpl)

  const res = await invoke(app.route, { url: SDK_JS })
  assert.equal(res.statusCode, 200)
  assert.deepEqual(res.body, payload)
  assert.equal(fetchImpl.calls.length, 2)
})

test('all mirrors returning 404 aggregate to 404 and name every mirror', async (t) => {
  const fetchImpl = makeFetch(() => ({ status: 404, body: Buffer.from('missing') }))
  const app = await mountMirrors(t, fetchImpl)

  const res = await invoke(app.route, { url: SDK_JS })
  assert.equal(res.statusCode, 404)
  const json = res.json()
  assert.equal(json.error.code, 'upstream-error')
  for (const mirror of [MIRROR_A, MIRROR_B, MIRROR_C]) {
    assert.ok(json.error.message.includes(mirror), `message must name ${mirror}`)
  }
  assert.equal(fetchImpl.calls.length, 3)
  assert.deepEqual(await listFiles(app.cacheDir), [])

  const status = (await getStatus(app)).json()
  assert.ok(status.lastError.includes(MIRROR_A))
  assert.ok(status.lastError.includes(MIRROR_B))
  assert.ok(status.lastError.includes(MIRROR_C))
  assert.match(status.lastError, /HTTP 404/)
})

test('all mirrors failing at the network level aggregate to 502', async (t) => {
  const fetchImpl = makeFetch(() => ({ error: new TypeError('socket hang up') }))
  const app = await mountMirrors(t, fetchImpl)

  const res = await invoke(app.route, { url: SDK_JS })
  assert.equal(res.statusCode, 502)
  assert.equal(fetchImpl.calls.length, 3)
  for (const mirror of [MIRROR_A, MIRROR_B, MIRROR_C]) {
    assert.ok(res.json().error.message.includes(mirror))
  }
  assert.deepEqual(await listFiles(app.cacheDir), [])
})

test('status reports the full ordered mirror list', async (t) => {
  const app = await mountMirrors(t, makeFetch(() => ({ status: 200, body: Buffer.from('x') })))
  const status = (await getStatus(app)).json()
  assert.deepEqual(status.upstreams, [MIRROR_A, MIRROR_B, MIRROR_C])
  assert.equal(status.upstreamBase, status.upstreams[0])
})

test('browser diagnostics land in the rotating request log', async (t) => {
  const app = await mount(t)
  const posted = await invoke(app.route, {
    url: '/dsh-onlyoffice/diagnostic',
    method: 'POST',
    body: JSON.stringify({ text: '[iframe-error] boom' }),
  })
  assert.equal(posted.statusCode, 204, 'a report is accepted with no body')

  const logFile = join(app.cacheDir, '.requests.log')
  await waitFor(async () => {
    try {
      return (await readFile(logFile, 'utf8')).includes('client [iframe-error] boom')
    } catch {
      return false
    }
  })
  assert.match(await readFile(logFile, 'utf8'), /client \[iframe-error\] boom/)

  const wrongMethod = await invoke(app.route, { url: '/dsh-onlyoffice/diagnostic', method: 'GET' })
  assert.equal(wrongMethod.statusCode, 405, 'only POST reports diagnostics')

  // Cache statistics ignore the dotfile log.
  const status = await getStatus(app)
  assert.equal(JSON.parse(status.body).files, 0)
})
