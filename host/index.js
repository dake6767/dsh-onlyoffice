// SPDX-License-Identifier: AGPL-3.0-only
// dsh-onlyoffice — see LICENSE and THIRD_PARTY_NOTICES.md.

/**
 * dsh-onlyoffice — Host half.
 *
 * Registers exactly one fenced prefix route, `/dsh-onlyoffice`, that serves:
 *
 *   GET /dsh-onlyoffice/status
 *     JSON: { ok, version, sdkBaseUrl, cacheDir, upstreamBase, upstreams,
 *             files, bytes, lastError, cacheable } — the client's single source
 *     of truth for the SDK version/base before it boots the vendored
 *     ONLYOFFICE editor. `upstreams` is the ordered failover candidate list;
 *     `upstreamBase` is its first entry (kept for older readers).
 *
 *   GET /dsh-onlyoffice/sdk/onlyoffice/<version>/<rest...>
 *     The ONLYOFFICE static SDK, served same-origin. Cache hit → stream from
 *     disk. Cache miss → try each configured upstream mirror in order, write
 *     the first successful body atomically to the cache (`tmp` + `rename`) and
 *     then stream it from disk, so the client always receives a body with an
 *     exact `Content-Length` and a failed upstream can never leave a partial
 *     file. A 404 is passed through only after every mirror returned 404.
 *
 * A prepared read-only shared SDK is checked before the writable member cache.
 * Missing assets are fetched on first request into the member cache.
 *
 * Plain ESM JavaScript with only `node:*` imports (plus the sibling helper
 * module): the package is installed as a linked directory outside the DSH
 * profile's `node_modules`, where no bare specifier resolves.
 */
import { createReadStream, createWriteStream, promises as fsp, readFileSync } from 'node:fs'
import { once } from 'node:events'
import { dirname, join } from 'node:path'
import { Readable } from 'node:stream'
import { pipeline } from 'node:stream/promises'

import { handleRead, handleSave, SAVE_ROUTE } from './save.js'
import {
  CACHE_CONTROL,
  PLUGIN_NAME,
  ROUTE_PREFIX,
  SDK_BASE_URL,
  SDK_PATH_PREFIX,
  TMP_MARKER,
  cachePathFor,
  contentTypeForPath,
  describeCache,
  isPathInside,
  parseSdkPath,
  resolveConfig,
  upstreamUrlFor,
} from './sdk-cache.js'

/** Plugin identity for the Loader row (must match the package name). */
export const name = PLUGIN_NAME

/** Routes need the webserver and the request-admission fence. */
export const inject = ['webServer', 'connection']

export const LOG_PREFIX = '[dsh-onlyoffice]'

/**
 * This plugin's own package version, reported by `/status`. Distinguishes the
 * plugin version from the OnlyOffice SDK version (the same two numbers people
 * conflate when debugging a remote deployment).
 */
export const PLUGIN_VERSION = (() => {
  try {
    const manifest = JSON.parse(readFileSync(new URL('../package.json', import.meta.url), 'utf8'))
    return typeof manifest.version === 'string' ? manifest.version : 'unknown'
  } catch {
    return 'unknown'
  }
})()

/** How long a definitive upstream 404 stays remembered (10 minutes). */
const NEGATIVE_CACHE_TTL_MS = 10 * 60_000
/** Upper bound on remembered misses (oldest evicted first). */
const NEGATIVE_CACHE_MAX = 512

/** Whole-request upstream budget (covers headers + body). */
const DEFAULT_FETCH_TIMEOUT_MS = 300_000
const MIN_FETCH_TIMEOUT_MS = 5_000
const MAX_FETCH_TIMEOUT_MS = 3_600_000

/**
 * Mount the plugin. Everything registered here is owned by `ctx` and removed
 * when the plugin unloads.
 *
 * Accepted `config` fields (all optional). `autoSave` defaults to false: manual
 * saves (Cmd/Ctrl+S, the editor save button) are always available, the
 * debounced background writer must be opted into.
 *   version            SDK version segment, default `9.4.0-develop`
 *   upstreamBase       first upstream candidate, default the canonical
 *                      onlyoffice-web-comp Vercel deployment
 *   upstreamFallbacks  extra candidate bases tried in order after
 *                      `upstreamBase`; defaults to the jsDelivr and
 *                      raw.githubusercontent mirrors of the same tree
 *   cacheDir           disk cache root, default `$DSH_HOME/onlyoffice-sdk`
 *   fetchTimeoutMs     per-attempt upstream budget in ms (advanced)
 *   fetch              fetch-compatible function (advanced; used by the tests)
 *
 * @param {object} ctx - Cordis/DSH plugin context (`webServer`, `connection`)
 * @param {unknown} [config]
 */
export function apply(ctx, config) {
  const resolved = resolveConfig(config)
  /** @type {{ lastError: string | null }} */
  const state = { lastError: null }
  const runtime = {
    resolved,
    state,
    fetchImpl: pickFetch(config),
    timeoutMs: pickTimeout(config),
    /** @type {Map<string, Promise<FetchResult>>} in-flight cache fills */
    inflight: new Map(),
    tmpSeq: 0,
    log: createLogger(ctx),
    ensureCacheDir: null,
    nextTmpId: () => {
      runtime.tmpSeq += 1
      return `${process.pid}-${runtime.tmpSeq}`
    },
    /** Serialized request-log appends (diagnostics only; failures are ignored). */
    logChain: Promise.resolve(),
    /** Definitive upstream 404s, cached briefly so optional SDK assets (spell
     *  dictionaries, help images) do not re-walk the mirror chain every time. */
    negative: new Map(),
    /** Save-back knobs (the write route is the plugin's only writer). */
    save: {
      enabled:
        config === null || config === undefined || typeof config !== 'object' || config.saveEnabled !== false,
      // Autosave stays opt-in: an editor that silently writes documents is a
      // surprise, so manual saves (Cmd/Ctrl+S and the editor's own button) are
      // the default and `autoSave: true` restores the debounced writer.
      autoSave: config !== null && config !== undefined && typeof config === 'object' && config.autoSave === true,
      maxSaveBytes: pickMaxSaveBytes(config),
    },
  }

  // Log invalid config values instead of failing activation: the routes still
  // work against the documented defaults.
  for (const warning of resolved.warnings) runtime.log('warn', warning)

  // ── Cache directory ──────────────────────────────────────────────────────
  // One memoized mkdir shared by every miss. The effect kicks it off eagerly
  // so /status can report a real cacheDir before the first asset is requested;
  // a failure is retried by the next miss and surfaces through `lastError`.
  let cacheDirReady = null
  runtime.ensureCacheDir = () => {
    if (cacheDirReady === null) {
      cacheDirReady = fsp.mkdir(resolved.cacheDir, { recursive: true })
        .then(() => resolved.cacheDir)
        .catch((error) => {
          cacheDirReady = null
          throw error
        })
    }
    return cacheDirReady
  }

  ctx.effect(() => {
    void runtime.ensureCacheDir().catch((error) => {
      state.lastError = `cannot create cache directory "${resolved.cacheDir}": ${formatError(error)}`
      runtime.log('warn', state.lastError)
    })
    return () => {
      cacheDirReady = null
      runtime.inflight.clear()
    }
  }, `${PLUGIN_NAME}: cache directory`)

  // ── The single route ─────────────────────────────────────────────────────
  ctx.effect(() => ctx.webServer.register({
    kind: 'prefix',
    path: ROUTE_PREFIX,
    handler: (req, res) => handleRequest(ctx, runtime, req, res),
  }), `${PLUGIN_NAME}: routes`)
}

/** @typedef {{ ok: true, size: number, cached?: boolean } | { ok: false, status: number, message: string }} FetchResult */

/**
 * The same fence `/sidebar/*` uses: the connection service owns trust and
 * authentication, so a route never infers admission by itself.
 * @param {object} ctx
 * @param {object} req
 * @returns {boolean}
 */
function isAdmitted(ctx, req) {
  try {
    return ctx.connection.requestRejection(req) === undefined
  } catch {
    // Fail closed if the fence itself misbehaves.
    return false
  }
}

/**
 * @param {object} ctx
 * @param {object} runtime
 * @param {object} req
 * @param {object} res
 */
async function handleRequest(ctx, runtime, req, res) {
  const method = typeof req.method === 'string' ? req.method.toUpperCase() : 'GET'
  const head = method === 'HEAD'
  const startedAt = Date.now()
  try {
    if (!isAdmitted(ctx, req)) {
      sendPlain(res, 403, 'forbidden')
      return
    }

    let pathname
    try {
      // The SDK itself requests some assets with a duplicated slash
      // (`sdkjs/slide/themes//themes.js`); collapse those instead of rejecting
      // the path, otherwise the presentation theme bundle 400s and never caches.
      pathname = new URL(typeof req.url === 'string' ? req.url : '/', 'http://dsh.internal').pathname.replace(/\/{2,}/g, '/')
    } catch {
      sendJson(res, 400, errorEnvelope('bad-request', 'malformed request URL'), {}, head)
      return
    }

    // The browser reports its own failures here (see src/client/diagnostics.js);
    // this is the only non-GET route.
    if (pathname === `${ROUTE_PREFIX}/diagnostic` || pathname === `${ROUTE_PREFIX}/diagnostic/`) {
      await receiveDiagnostic(runtime, req, res)
      return
    }
    // Save-back: the plugin's only write path (POST, raw document bytes).
    if (pathname === SAVE_ROUTE || pathname === `${SAVE_ROUTE}/`) {
      if (runtime.save.enabled !== true) {
        sendJson(res, 403, errorEnvelope('save-disabled', 'save-back is disabled by configuration'), {}, head)
        return
      }
      await handleSave(ctx, req, res, { maxSaveBytes: runtime.save.maxSaveBytes })
      return
    }
    // Document bytes for the built-in preview's `renderer` loading mode.
    if (pathname === `${ROUTE_PREFIX}/read` || pathname === `${ROUTE_PREFIX}/read/`) {
      await handleRead(ctx, req, res)
      return
    }
    if (method !== 'GET' && method !== 'HEAD') {
      sendJson(res, 405, errorEnvelope('method-error', 'method not allowed'), { allow: 'GET, HEAD' }, head)
      return
    }

    if (pathname === `${ROUTE_PREFIX}/status` || pathname === `${ROUTE_PREFIX}/status/`) {
      await sendStatus(res, runtime, head)
      return
    }
    if (pathname.startsWith(SDK_PATH_PREFIX)) {
      await sendSdkAsset(res, pathname, runtime, head)
      return
    }
    sendJson(res, 404, errorEnvelope('not-found', `no route for ${pathname}`), {}, head)
  } catch (error) {
    const message = `${method} ${typeof req.url === 'string' ? req.url : ''}: ${formatError(error)}`
    runtime.state.lastError = message
    runtime.log('error', message)
    if (res.headersSent === true) {
      // The status line is already on the wire; truncation is the only honest
      // signal left (same policy as dsh-better-sidebar's streamed routes).
      res.destroy?.()
      return
    }
    try {
      sendJson(res, 500, errorEnvelope('internal-error', formatError(error)), {}, head)
    } catch {
      res.destroy?.()
    }
  } finally {
    // Every request (fence rejections included) lands in one small rotating
    // log, so a stuck browser session can be diagnosed from the Host side.
    recordRequest(runtime, method, typeof req.url === 'string' ? req.url : '', res.statusCode, Date.now() - startedAt)
  }
}

/**
 * Accept one browser diagnostic report (`POST /dsh-onlyoffice/diagnostic`).
 * Bounded body, no response payload: the client fires and forgets.
 * @param {object} runtime
 * @param {object} req
 * @param {object} res
 */
async function receiveDiagnostic(runtime, req, res) {
  if (typeof req.method === 'string' && req.method.toUpperCase() !== 'POST') {
    res.writeHead(405, { allow: 'POST', 'content-type': 'text/plain; charset=utf-8', 'content-length': '18' })
    res.end('method not allowed')
    return
  }

  const MAX_BODY = 64 * 1024
  let size = 0
  const chunks = []
  try {
    for await (const chunk of req) {
      size += chunk.length
      if (size > MAX_BODY) {
        res.writeHead(413, { 'content-type': 'text/plain; charset=utf-8', 'content-length': '8' })
        res.end('too big\n')
        return
      }
      chunks.push(chunk)
    }
  } catch {
    res.writeHead(400, { 'content-type': 'text/plain; charset=utf-8', 'content-length': '0' })
    res.end()
    return
  }

  let text = ''
  try {
    const raw = Buffer.concat(chunks).toString('utf8')
    const parsed = raw ? JSON.parse(raw) : undefined
    text = parsed && typeof parsed.text === 'string' ? parsed.text : raw
  } catch {
    text = Buffer.concat(chunks).toString('utf8')
  }

  appendLogLine(runtime, `client ${String(text).slice(0, 4000).replace(/\s*\n\s*/g, ' | ')}`)
  res.writeHead(204, { 'cache-control': 'no-store' })
  res.end()
}

/**
 * Append one diagnostic request line to `<cacheDir>/.requests.log`. The log is
 * bookkeeping only: it never blocks a response and it is excluded from the
 * cache statistics. Rotated at 2 MiB so it cannot grow without bound.
 * @param {object} runtime
 * @param {string} method
 * @param {string} url
 * @param {number | undefined} status
 * @param {number} durationMs
 */
function recordRequest(runtime, method, url, status, durationMs) {
  appendLogLine(
    runtime,
    `${new Date().toISOString()} ${method} ${url} ${Number.isInteger(status) ? status : 0} ${durationMs}ms`,
  )
}

/**
 * Append one line to the rotating diagnostics log. Serialized through the
 * runtime's chain, capped at 2 MiB, and never allowed to affect serving.
 * @param {object} runtime
 * @param {string} line
 */
function appendLogLine(runtime, line) {
  runtime.logChain = runtime.logChain.then(async () => {
    try {
      const file = join(runtime.resolved.cacheDir, '.requests.log')
      const info = await fsp.stat(file).catch(() => null)
      if (info !== null && info.size > 2 * 1024 * 1024) {
        await fsp.rename(file, `${file}.1`).catch(() => {})
      }
      await fsp.appendFile(file, `${line}\n`)
    } catch {
      // Diagnostics must never affect serving.
    }
  }, () => {})
}

/**
 * @param {object} res
 * @param {object} runtime
 * @param {boolean} head
 */
async function sendStatus(res, runtime, head) {
  let stats = { files: 0, bytes: 0 }
  try {
    stats = await describeCache(runtime.resolved.cacheDir)
  } catch (error) {
    runtime.state.lastError = `cannot scan cache directory "${runtime.resolved.cacheDir}": ${formatError(error)}`
    runtime.log('warn', runtime.state.lastError)
  }
  sendJson(res, 200, {
    ok: true,
    pluginVersion: PLUGIN_VERSION,
    version: runtime.resolved.version,
    sdkBaseUrl: SDK_BASE_URL,
    cacheDir: runtime.resolved.cacheDir,
    // `upstreamBase` stays the first (canonical) candidate for older readers;
    // `upstreams` is the complete ordered failover list.
    upstreamBase: runtime.resolved.upstreamBase,
    upstreams: runtime.resolved.upstreams,
    autoSave: runtime.save.autoSave,
    saveEnabled: runtime.save.enabled,
    files: stats.files,
    bytes: stats.bytes,
    lastError: runtime.state.lastError,
    cacheable: true,
  }, {}, head)
}

/**
 * @param {object} res
 * @param {string} pathname
 * @param {object} runtime
 * @param {boolean} head
 */
async function sendSdkAsset(res, pathname, runtime, head) {
  const parsed = parseSdkPath(pathname)
  if (parsed.ok !== true) {
    sendJson(res, parsed.status, errorEnvelope(parsed.code, parsed.message), {}, head)
    return
  }
  const { version, segments } = parsed
  if (version !== runtime.resolved.version) {
    sendJson(res, 404, errorEnvelope(
      'version-mismatch',
      `SDK version "${version}" is not the configured version "${runtime.resolved.version}"`,
    ), {}, head)
    return
  }

  const target = cachePathFor(runtime.resolved.cacheDir, version, segments)
  if (target === null) {
    sendJson(res, 400, errorEnvelope('bad-path', 'unsafe SDK asset path'), {}, head)
    return
  }

  const key = segments.join('/')
  const sharedRoot = runtime.resolved.sharedCacheDir
  const sharedTarget = sharedRoot === null ? null : cachePathFor(sharedRoot, version, segments)
  if (sharedTarget !== null) {
    const shared = await sharedFile(sharedRoot, sharedTarget)
    if (shared !== null) {
      runtime.negative.delete(key)
      runtime.log('info', `shared cache hit ${version}/${key}`)
      await sendFile(res, shared.path, shared.info, head)
      return
    }
  }
  const cached = await statFile(target)
  if (cached !== null) {
    runtime.negative.delete(key)
    runtime.log('info', `cache hit ${version}/${key}`)
    await sendFile(res, target, cached, head)
    return
  }

  const remembered = runtime.negative.get(key)
  if (remembered !== undefined && Date.now() - remembered.at < NEGATIVE_CACHE_TTL_MS) {
    sendJson(res, remembered.status, errorEnvelope('not-found', `upstream has no ${version}/${key}`), {}, head)
    return
  }

  const upstreams = runtime.resolved.upstreams.map((base) => upstreamUrlFor(base, version, segments))
  runtime.log('info', `cache miss ${version}/${key} -> ${upstreams.join(' | ')}`)
  const result = await claimFetch(runtime, target, () => fetchIntoTarget(runtime, target, upstreams, `${version}/${key}`))
  if (result.ok !== true) {
    const status = normalizeErrorStatus(result.status)
    if (status === 404) rememberMissing(runtime, key)
    sendJson(res, status, errorEnvelope('upstream-error', result.message), {}, head)
    return
  }
  runtime.negative.delete(key)

  const info = await statFile(target)
  if (info === null) {
    sendJson(res, 502, errorEnvelope('cache-error', `upstream succeeded but "${target}" is missing`), {}, head)
    return
  }
  await sendFile(res, target, info, head)
}

/**
 * Remember a definitive 404 so the next request answers immediately instead of
 * walking the mirror chain again.
 * @param {object} runtime
 * @param {string} key SDK-relative asset path
 */
function rememberMissing(runtime, key) {
  runtime.negative.set(key, { status: 404, at: Date.now() })
  if (runtime.negative.size <= NEGATIVE_CACHE_MAX) return
  const oldest = runtime.negative.keys().next()
  if (oldest.done !== true) runtime.negative.delete(oldest.value)
}

/**
 * Coalesce concurrent fills of the same cache path into one upstream fetch.
 * @param {object} runtime
 * @param {string} key
 * @param {() => Promise<FetchResult>} task
 * @returns {Promise<FetchResult>}
 */
function claimFetch(runtime, key, task) {
  const existing = runtime.inflight.get(key)
  if (existing !== undefined) return existing
  const tracked = Promise.resolve()
    .then(task)
    .finally(() => {
      if (runtime.inflight.get(key) === tracked) runtime.inflight.delete(key)
    })
  runtime.inflight.set(key, tracked)
  return tracked
}

/**
 * @typedef {{
 *   ok: false,
 *   url: string,
 *   kind: 'http' | 'network' | 'timeout' | 'write',
 *   status: number,
 *   message: string,
 * }} UpstreamAttempt
 */

/**
 * Fetch one asset, trying every upstream candidate in order until one returns
 * 200 and its bytes are safely renamed into the cache.
 *
 * The mirrors serve the same static tree, so a 404/5xx/network failure on one
 * is retried on the next. A definitive 404 is reported only after every
 * candidate answered 404; otherwise the last meaningful failure is reported
 * and `lastError` records every mirror that was tried.
 *
 * @param {object} runtime
 * @param {string} target - absolute cache path
 * @param {string[]} upstreamUrls - ordered candidate URLs for this asset
 * @param {string} label - human-readable asset key for messages
 * @returns {Promise<FetchResult>}
 */
async function fetchIntoTarget(runtime, target, upstreamUrls, label) {
  await runtime.ensureCacheDir()

  // Another process (or an earlier request in this one) may have populated the
  // cache between our stat and claiming the fill.
  const raced = await statFile(target)
  if (raced !== null) return { ok: true, size: raced.size, cached: true }

  /** @type {UpstreamAttempt[]} */
  const attempts = []
  for (let index = 0; index < upstreamUrls.length; index += 1) {
    const url = upstreamUrls[index]
    const attempt = await fetchOne(runtime, target, url)
    if (attempt.ok === true) {
      runtime.state.lastError = null
      runtime.log('info', `cached ${url} (${attempt.size} bytes)`)
      return attempt
    }
    attempts.push(attempt)
    const next = index + 1 < upstreamUrls.length ? ' -> next' : ''
    runtime.log('warn', `mirror fail ${url}${next}: ${attempt.message}`)
  }

  const failure = aggregateAttempts(attempts, label)
  runtime.state.lastError = failure.message
  runtime.log('warn', failure.message)
  return failure
}

/**
 * One upstream attempt. The bytes land in a temp sibling first and only a
 * fully-written file is renamed into place, so a crash, a network failure or a
 * concurrent request can never expose a partial cache file.
 * @param {object} runtime
 * @param {string} target
 * @param {string} url
 * @returns {Promise<{ ok: true, size: number } | UpstreamAttempt>}
 */
async function fetchOne(runtime, target, url) {
  const timeoutError = new Error(`timed out after ${runtime.timeoutMs} ms fetching ${url}`)
  timeoutError.name = 'TimeoutError'
  const controller = new AbortController()
  const timer = setTimeout(() => controller.abort(timeoutError), runtime.timeoutMs)
  timer.unref?.()

  try {
    let response
    try {
      response = await runtime.fetchImpl(url, {
        redirect: 'follow',
        signal: controller.signal,
        headers: { accept: '*/*', 'user-agent': `${PLUGIN_NAME}/0.1` },
      })
    } catch (error) {
      const timeout = isTimeoutError(error, controller)
      return {
        ok: false,
        url,
        kind: timeout ? 'timeout' : 'network',
        status: timeout ? 504 : 502,
        message: `${timeout ? 'timeout' : 'network error'} fetching ${url}: ${formatError(error)}`,
      }
    }

    if (response.status !== 200) {
      try {
        await response.body?.cancel?.()
      } catch {
        // A body that cannot be cancelled is garbage-collected; nothing to do.
      }
      return {
        ok: false,
        url,
        kind: 'http',
        status: response.status,
        message: `upstream ${url} returned HTTP ${response.status}`,
      }
    }

    await fsp.mkdir(dirname(target), { recursive: true })
    const tmp = `${target}${TMP_MARKER}${runtime.nextTmpId()}`
    try {
      if (response.body === null || response.body === undefined) {
        await fsp.writeFile(tmp, Buffer.alloc(0))
      } else {
        await pipeline(Readable.fromWeb(response.body), createWriteStream(tmp))
      }
      if (controller.signal.aborted) throw controller.signal.reason ?? timeoutError
      await fsp.rename(tmp, target)
    } catch (error) {
      await fsp.rm(tmp, { force: true }).catch(() => {})
      const timeout = isTimeoutError(error, controller)
      return {
        ok: false,
        url,
        kind: timeout ? 'timeout' : 'write',
        status: timeout ? 504 : 502,
        message: timeout
          ? timeoutError.message
          : `failed to cache ${url} at "${target}": ${formatError(error)}`,
      }
    }

    const info = await statFile(target)
    return { ok: true, size: info === null ? 0 : info.size }
  } finally {
    clearTimeout(timer)
  }
}

/**
 * Decide the single status/message for "every mirror failed". Keeps the
 * contract's "404 stays 404" rule only when *all* mirrors agreed on 404.
 * @param {UpstreamAttempt[]} attempts
 * @param {string} label
 * @returns {{ ok: false, status: number, message: string }}
 */
function aggregateAttempts(attempts, label) {
  const summary = attempts.map((attempt) => {
    let outcome
    if (attempt.kind === 'http') outcome = `HTTP ${attempt.status}`
    else if (attempt.kind === 'timeout') outcome = 'timeout'
    else if (attempt.kind === 'network') outcome = 'network error'
    else outcome = 'cache write error'
    return `${attempt.url} -> ${outcome}`
  }).join('; ')
  const message = `all upstream mirrors failed for ${label}: ${summary}`

  if (attempts.length > 0
    && attempts.every((attempt) => attempt.kind === 'http')
    && attempts.every((attempt) => attempt.status === 404)) {
    return { ok: false, status: 404, message }
  }
  if (attempts.length > 0 && attempts.every((attempt) => attempt.kind === 'timeout')) {
    return { ok: false, status: 504, message }
  }
  if (attempts.length > 0
    && attempts.every((attempt) => attempt.kind === 'network' || attempt.kind === 'timeout')) {
    return { ok: false, status: 502, message }
  }
  const lastDefinitive = [...attempts].reverse().find((attempt) => attempt.kind === 'http' && attempt.status !== 404)
  if (lastDefinitive !== undefined) return { ok: false, status: lastDefinitive.status, message }
  return { ok: false, status: 502, message }
}

/**
 * Stream one cached file with the immutable-asset headers. Stream open is
 * awaited before `writeHead`, so a filesystem race still becomes a clean 500
 * instead of a 200 with a truncated body.
 * @param {object} res
 * @param {string} target
 * @param {import('node:fs').Stats} info
 * @param {boolean} head
 */
async function sendFile(res, target, info, head) {
  const headers = {
    'content-type': contentTypeForPath(target),
    'content-length': String(info.size),
    'cache-control': CACHE_CONTROL,
    'x-content-type-options': 'nosniff',
  }
  if (head) {
    res.writeHead(200, headers)
    res.end()
    return
  }
  const stream = createReadStream(target)
  try {
    await once(stream, 'open')
  } catch (error) {
    stream.destroy()
    throw error
  }
  res.writeHead(200, headers)
  try {
    await pipeline(stream, res)
  } catch {
    // The status line is already committed; destroy and let the client's
    // loader report the truncation.
    res.destroy?.()
  }
}

/** Read only ordinary files whose resolved path stays inside the shared root. */
async function sharedFile(root, target) {
  try {
    const base = await fsp.realpath(root)
    const path = await fsp.realpath(target)
    if (!isPathInside(base, path)) return null
    const info = await fsp.stat(path)
    return info.isFile() ? { path, info } : null
  } catch {
    // Missing, inaccessible or unsuitable shared entries use the member cache.
    return null
  }
}

/**
 * @param {string} filePath
 * @returns {Promise<import('node:fs').Stats | null>}
 */
async function statFile(filePath) {
  try {
    const info = await fsp.stat(filePath)
    return info.isFile() === true ? info : null
  } catch (error) {
    if (error !== null && typeof error === 'object' && (error.code === 'ENOENT' || error.code === 'ENOTDIR')) {
      return null
    }
    throw error
  }
}

/** @param {unknown} status @returns {number} */
function normalizeErrorStatus(status) {
  return Number.isInteger(status) && status >= 400 && status <= 599 ? status : 502
}

/**
 * @param {unknown} error
 * @param {AbortController} controller
 * @returns {boolean}
 */
function isTimeoutError(error, controller) {
  if (controller.signal.aborted === true) return true
  if (error === null || typeof error !== 'object') return false
  return error.name === 'TimeoutError' || error.name === 'AbortError'
}

/** @param {unknown} error @returns {string} */
function formatError(error) {
  if (error instanceof Error) return `${error.name}: ${error.message}`
  return String(error)
}

/** @param {string} code @param {string} message */
function errorEnvelope(code, message) {
  return { ok: false, error: { code, message } }
}

/**
 * @param {object} res
 * @param {number} status
 * @param {unknown} value
 * @param {Record<string, string>} [extraHeaders]
 * @param {boolean} [head]
 */
function sendJson(res, status, value, extraHeaders = {}, head = false) {
  const body = JSON.stringify(value)
  res.writeHead(status, {
    'content-type': 'application/json; charset=utf-8',
    'content-length': String(Buffer.byteLength(body)),
    'cache-control': 'no-store',
    'x-content-type-options': 'nosniff',
    ...extraHeaders,
  })
  res.end(head ? undefined : body)
}

/**
 * @param {object} res
 * @param {number} status
 * @param {string} text
 */
function sendPlain(res, status, text) {
  res.writeHead(status, {
    'content-type': 'text/plain; charset=utf-8',
    'content-length': String(Buffer.byteLength(text)),
    'cache-control': 'no-store',
    'x-content-type-options': 'nosniff',
  })
  res.end(text)
}

/**
 * Route logging. Prefers the host logger (which prefixes console output) and
 * falls back to `console.log` so a DSH context without a logger still records
 * each miss / served-from-cache line with the `[dsh-onlyoffice]` prefix.
 * @param {object} ctx
 * @returns {(level: 'debug' | 'info' | 'warn' | 'error', message: string) => void}
 */
function createLogger(ctx) {
  return (level, message) => {
    const line = `${LOG_PREFIX} ${message}`
    const logger = ctx === null || ctx === undefined ? undefined : ctx.logger
    const emit = logger === null || logger === undefined ? undefined : logger[level]
    if (typeof emit === 'function') {
      try {
        emit.call(logger, line)
        return
      } catch {
        // A broken logger must never break asset serving.
      }
    }
    console.log(line)
  }
}

/**
 * @param {unknown} config
 * @returns {(...args: unknown[]) => Promise<Response>}
 */
function pickFetch(config) {
  if (config !== null && typeof config === 'object' && typeof config.fetch === 'function') {
    return config.fetch
  }
  return (...args) => globalThis.fetch(...args)
}

/**
 * @param {unknown} config
 * @returns {number}
 */
function pickMaxSaveBytes(config) {
  const configured = config !== null && typeof config === 'object' ? config.maxSaveBytes : undefined
  if (!Number.isFinite(configured)) return 256 * 1024 * 1024
  return Math.min(1024 * 1024 * 1024, Math.max(1024, Math.trunc(configured)))
}

/**
 * @param {unknown} config
 * @returns {number}
 */
function pickTimeout(config) {
  const configured = config !== null && typeof config === 'object' ? config.fetchTimeoutMs : undefined
  if (!Number.isFinite(configured)) return DEFAULT_FETCH_TIMEOUT_MS
  return Math.min(MAX_FETCH_TIMEOUT_MS, Math.max(MIN_FETCH_TIMEOUT_MS, Math.trunc(configured)))
}
