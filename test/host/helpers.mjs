/**
 * Test doubles for the Host half: a fake Cordis context, a Writable response,
 * an injectable fetch, and a temp cache dir. DSH services never load here —
 * `host/index.js` receives everything through `ctx` and `config`.
 */
import { mkdtemp, readdir, rm, stat } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { Writable } from 'node:stream'

/**
 * Message of a rejected `send_json`-style body: Node's ServerResponse is a
 * Writable, so a Writable subclass is the closest faithful double.
 */
export class MockResponse extends Writable {
  constructor() {
    super()
    this.statusCode = 200
    this._headers = {}
    this._headersSent = false
    this.chunks = []
    this.destroyedByRoute = false
  }

  get headersSent() {
    return this._headersSent
  }

  writeHead(status, headers) {
    this.statusCode = status
    this._headersSent = true
    if (headers !== null && headers !== undefined) {
      for (const [key, value] of Object.entries(headers)) {
        this._headers[String(key).toLowerCase()] = value
      }
    }
    return this
  }

  setHeader(key, value) {
    this._headers[String(key).toLowerCase()] = value
    return this
  }

  getHeader(key) {
    return this._headers[String(key).toLowerCase()]
  }

  _write(chunk, _encoding, callback) {
    this.chunks.push(Buffer.from(chunk))
    callback()
  }

  get body() {
    return Buffer.concat(this.chunks)
  }

  get text() {
    return this.body.toString('utf8')
  }

  json() {
    return JSON.parse(this.text)
  }

  destroy(error) {
    this.destroyedByRoute = true
    return super.destroy(error)
  }
}

/**
 * @param {object} [options]
 * @param {unknown} [options.rejection] - value `connection.requestRejection` returns
 */
export function createFakeContext(options = {}) {
  const routes = []
  const effects = []
  const logs = []
  const logger = {
    debug: (...args) => logs.push({ level: 'debug', args }),
    info: (...args) => logs.push({ level: 'info', args }),
    warn: (...args) => logs.push({ level: 'warn', args }),
    error: (...args) => logs.push({ level: 'error', args }),
  }

  const ctx = {
    logger,
    webServer: {
      register(route) {
        routes.push(route)
        return () => {
          const index = routes.indexOf(route)
          if (index >= 0) routes.splice(index, 1)
        }
      },
    },
    connection: {
      requestRejection() {
        return options.rejection
      },
    },
    effect(callback, label) {
      const dispose = callback()
      const record = { label, dispose }
      effects.push(record)
      return () => {
        const index = effects.indexOf(record)
        if (index >= 0) effects.splice(index, 1)
        if (typeof dispose === 'function') dispose()
      }
    },
  }

  return { ctx, routes, effects, logs, logger }
}

/**
 * Invoke a registered route handler and resolve once it finishes writing.
 * @param {{ handler: (req: object, res: MockResponse) => unknown }} route
 * @param {{ url: string, method?: string, headers?: Record<string, string> }} request
 */
export async function invoke(route, request) {
  const res = new MockResponse()
  const body =
    request.body === undefined || request.body === null
      ? []
      : [Buffer.isBuffer(request.body) ? request.body : Buffer.from(String(request.body), 'utf8')]
  const req = {
    url: request.url,
    method: request.method ?? 'GET',
    headers: request.headers ?? {},
    async *[Symbol.asyncIterator]() {
      // Most routes ignore the body; the diagnostic route reads it.
      for (const chunk of body) yield chunk
    },
  }
  await route.handler(req, res)
  return res
}

/**
 * Fetch double. `handler(url, callNumber)` returns a spec:
 * `{ status?, body?, headers?, error?, delayMs? }`.
 * @param {(url: string, callNumber: number) => object | Promise<object>} handler
 */
export function makeFetch(handler) {
  const calls = []
  const impl = async (url, init) => {
    calls.push({ url, init })
    const spec = await handler(url, calls.length)
    if (spec === null || spec === undefined) throw new Error(`unexpected upstream fetch #${calls.length}: ${url}`)
    if (spec.delayMs) await new Promise((resolveDelay) => setTimeout(resolveDelay, spec.delayMs))
    if (spec.error) throw spec.error
    return new Response(spec.body ?? Buffer.alloc(0), {
      status: spec.status ?? 200,
      headers: spec.headers ?? {},
    })
  }
  impl.calls = calls
  return impl
}

/** Create a temp scratch directory, removed when the test finishes. */
export async function makeTempDir(testContext, prefix = 'dsh-onlyoffice-test-') {
  const dir = await mkdtemp(join(tmpdir(), prefix))
  // The plugin appends diagnostics asynchronously (`.requests.log`), so the
  // directory can gain a file while cleanup runs: retry instead of failing.
  testContext.after(async () => {
    for (let attempt = 0; attempt < 8; attempt += 1) {
      try {
        await rm(dir, { recursive: true, force: true, maxRetries: 3, retryDelay: 20 })
        return
      } catch (error) {
        if (attempt === 7) {
          if (error && (error.code === 'ENOTEMPTY' || error.code === 'ENOENT')) return
          throw error
        }
        await new Promise((resolve) => setTimeout(resolve, 25))
      }
    }
  })
  return dir
}

/** All files under `dir` (recursive), as `${relativePath}` strings. */
export async function listFiles(dir) {
  let entries
  try {
    entries = await readdir(dir, { recursive: true, withFileTypes: true })
  } catch (error) {
    if (error !== null && typeof error === 'object' && error.code === 'ENOENT') return []
    throw error
  }
  return entries
    .filter((entry) => entry.isFile())
    // Dotfiles are the plugin's own bookkeeping (`.requests.log`), not cache.
    .filter((entry) => !entry.name.startsWith('.'))
    .map((entry) => join(entry.parentPath ?? dir, entry.name).slice(dir.length + 1))
}

/** List temp files a crashed/interrupted cache write would leave behind. */
export async function listTempFiles(dir) {
  return (await listFiles(dir)).filter((name) => name.includes('.~dsh-tmp-'))
}

/** Poll `check` until it returns truthy or the budget expires. */
export async function waitFor(check, timeoutMs = 2000, intervalMs = 10) {
  const deadline = Date.now() + timeoutMs
  let lastError
  while (Date.now() < deadline) {
    try {
      const value = await check()
      if (value) return value
    } catch (error) {
      lastError = error
    }
    await new Promise((resolveDelay) => setTimeout(resolveDelay, intervalMs))
  }
  if (lastError !== undefined) throw lastError
  throw new Error(`condition not met within ${timeoutMs} ms`)
}

/** Convenience: does `path` exist and is it a directory? */
export async function isDirectory(path) {
  try {
    return (await stat(path)).isDirectory()
  } catch {
    return false
  }
}
