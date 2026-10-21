// SPDX-License-Identifier: AGPL-3.0-only
// dsh-onlyoffice — see LICENSE and THIRD_PARTY_NOTICES.md.

/**
 * dsh-onlyoffice — save-back route.
 *
 * `POST /dsh-onlyoffice/save?sessionId=…&path=…[&cwd=…][&mtimeMs=…][&ext=…]`
 * with the **raw document bytes** as the request body writes the edited
 * document back over the workspace file the user opened.
 *
 * Safety properties (this route is the only writer in the plugin):
 *  - fenced exactly like the asset routes (the caller checks admission first);
 *  - the body is capped and the target must ALREADY exist as a regular file;
 *  - the target must stay inside the session's workspace directory
 *    (`fs.contains` when the filesystem provider offers it, realpath
 *    containment otherwise);
 *  - only the document extensions this plugin opens may be written;
 *  - the write is atomic (same-directory temp file + rename) and preserves the
 *    original file mode;
 *  - an optional `mtimeMs` is an optimistic-concurrency baseline: if the file
 *    changed on disk since the editor loaded it, the write is refused with
 *    `409 conflict` instead of clobbering someone else's edit.
 *
 * Only `node:*` builtins are imported: the package is installed as a linked
 * directory outside the profile's `node_modules`.
 */
import { promises as fsp } from 'node:fs'
import { dirname, extname, isAbsolute, join, relative, resolve, sep } from 'node:path'

/**
 * Extensions this plugin may write back (lower case, no dot).
 *
 * Only the NATIVE output formats of the embedded engine: it always exports
 * docx/xlsx/pptx bytes, so writing those bytes over a `.doc`, `.odt`, `.csv`
 * or `.ppt` file would silently corrupt it. Such files stay open + export only.
 */
export const SAVEABLE_EXTENSIONS = new Set(['docx', 'xlsx', 'pptx'])

/** Extensions the renderer may read through the host route (all openable). */
export const READABLE_EXTENSIONS = new Set([
  ...SAVEABLE_EXTENSIONS,
  'doc',
  'odt',
  'rtf',
  'xls',
  'ods',
  'csv',
  'ppt',
  'odp',
])

/** Hard cap on one saved body (a document, not an upload endpoint). */
export const DEFAULT_MAX_SAVE_BYTES = 256 * 1024 * 1024

/** Query + body limits that keep a malformed request cheap to reject. */
const MAX_PATH_LENGTH = 4096

class SaveError extends Error {
  /**
   * @param {string} code stable machine code
   * @param {string} message human message
   * @param {number} [status] HTTP status
   */
  constructor(code, message, status = 400) {
    super(message)
    this.name = 'SaveError'
    this.code = code
    this.status = status
  }
}

/** @param {string} text @returns {boolean} */
function isSafeSegment(text) {
  return typeof text === 'string' && text.length > 0 && text.length <= MAX_PATH_LENGTH && !text.includes('\u0000')
}

/** Read the whole request body, refusing anything over `maxBytes`. */
async function readBody(req, maxBytes) {
  const chunks = []
  let size = 0
  for await (const chunk of req) {
    const buffer = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk)
    size += buffer.length
    if (size > maxBytes) throw new SaveError('too-large', `the document exceeds the ${maxBytes}-byte save limit`, 413)
    chunks.push(buffer)
  }
  return Buffer.concat(chunks)
}

/** Realpath of the nearest existing ancestor, joined with the unresolved tail. */
async function realpathLoose(target) {
  const absolute = resolve(target)
  try {
    return await fsp.realpath(absolute)
  } catch {
    const parent = dirname(absolute)
    if (parent === absolute) return absolute
    const resolvedParent = await realpathLoose(parent)
    return join(resolvedParent, relative(parent, absolute))
  }
}

/** Whether `child` is `parent` or inside it (both already canonical). */
function isInside(parent, child) {
  const rel = relative(parent, child)
  return rel === '' || (!rel.startsWith('..') && !isAbsolute(rel))
}

/**
 * Resolve the host path a save should write to.
 * Uses the composed filesystem provider when available so a future non-local
 * backend is at least visible (and refused rather than written to a guess).
 *
 * @param {object} ctx plugin context
 * @param {{ sessionId: string, path: string, cwd?: string }} input
 * @returns {Promise<{ hostPath: string, rootPath: string, provider: boolean, version?: unknown, target?: unknown }>}
 */
async function resolveTarget(ctx, input) {
  const fs = typeof ctx.get === 'function' ? ctx.get('fs') : undefined
  const provider = Boolean(fs && typeof fs.resolve === 'function' && typeof fs.processPath === 'function')

  if (provider) {
    const target = await fs.resolve(input.path, input.cwd ? { cwd: input.cwd } : undefined)
    const info = await fs.stat(target)
    if (!info || info.type !== 'file') {
      throw new SaveError('not-a-file', 'the target is not an existing regular file', 409)
    }
    if (typeof fs.contains === 'function') {
      const root = await fs.resolve(input.cwd)
      if (!fs.contains(root, target)) {
        throw new SaveError('outside-workspace', 'the file is outside the session workspace', 403)
      }
    }
    return {
      hostPath: fs.processPath(target),
      rootPath: fs.processPath(await fs.resolve(input.cwd)),
      provider: true,
      version: info.version,
      target,
    }
  }

  // No provider (a stripped-down host): plain filesystem resolution with the
  // same containment rule.
  const cwd = input.cwd ? resolve(input.cwd) : undefined
  if (!cwd) throw new SaveError('no-cwd', 'the session directory is unknown; cannot validate the target', 409)
  const absolute = isAbsolute(input.path) ? resolve(input.path) : resolve(cwd, input.path)
  const hostPath = await realpathLoose(absolute)
  const rootPath = await realpathLoose(cwd)
  if (!isInside(rootPath, hostPath)) {
    throw new SaveError('outside-workspace', 'the file is outside the session workspace', 403)
  }
  return { hostPath, rootPath, provider: false }
}

/** Resolve the session directory for a save request. */
async function resolveCwd(ctx, sessionId, explicit) {
  if (explicit) return explicit
  const sessions = typeof ctx.get === 'function' ? ctx.get('sessions') : undefined
  const workingDirectory = typeof ctx.get === 'function' ? ctx.get('workingDirectory') : undefined
  if (!sessions || typeof sessions.get !== 'function' || !workingDirectory || typeof workingDirectory.get !== 'function') {
    return undefined
  }
  const session = sessions.get(sessionId)
  if (!session) return undefined
  try {
    return workingDirectory.get(session)
  } catch {
    return undefined
  }
}

/** Original file mode so the replaced file does not lose its permissions. */
async function modeOf(filePath) {
  try {
    const info = await fsp.stat(filePath)
    return info.mode & 0o777
  } catch {
    return undefined
  }
}

/**
 * Atomically replace one existing file.
 * @param {string} hostPath absolute path
 * @param {Uint8Array} bytes
 * @returns {Promise<{ bytes: number, mtimeMs: number }>}
 */
async function writeAtomic(hostPath, bytes) {
  const mode = await modeOf(hostPath)
  const tmp = `${hostPath}.~dsh-save-${process.pid}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`
  try {
    await fsp.writeFile(tmp, bytes, mode === undefined ? undefined : { mode })
    await fsp.rename(tmp, hostPath)
  } catch (error) {
    await fsp.rm(tmp, { force: true }).catch(() => {})
    throw error
  }
  const info = await fsp.stat(hostPath)
  return { bytes: info.size, mtimeMs: info.mtimeMs }
}

/**
 * Handle one save request. Admission (the connection fence) is the caller's
 * job; everything else is here.
 * @param {object} ctx plugin context
 * @param {object} req node request
 * @param {object} res node response
 * @param {{ maxSaveBytes?: number }} [options]
 */
export async function handleSave(ctx, req, res, options = {}) {
  const maxSaveBytes = Number.isInteger(options.maxSaveBytes) ? options.maxSaveBytes : DEFAULT_MAX_SAVE_BYTES
  try {
    if (typeof req.method === 'string' && req.method.toUpperCase() !== 'POST') {
      throw new SaveError('method-error', 'method not allowed', 405)
    }

    let url
    try {
      url = new URL(typeof req.url === 'string' ? req.url : '/', 'http://dsh.internal')
    } catch {
      throw new SaveError('bad-request', 'malformed request URL')
    }

    const sessionId = url.searchParams.get('sessionId')
    const path = url.searchParams.get('path')
    const explicitCwd = url.searchParams.get('cwd') || undefined
    if (!isSafeSegment(sessionId) || !isSafeSegment(path)) {
      throw new SaveError('bad-request', 'sessionId and path are required')
    }

    const extension = extname(path).slice(1).toLowerCase()
    if (!SAVEABLE_EXTENSIONS.has(extension)) {
      throw new SaveError('unsupported-extension', `refusing to write ".${extension}" (not an Office document this plugin opens)`, 415)
    }

    const cwd = await resolveCwd(ctx, sessionId, explicitCwd)
    if (!cwd) {
      throw new SaveError('no-cwd', 'the session directory is unknown, so the target cannot be validated', 409)
    }

    const resolvedTarget = await resolveTarget(ctx, { sessionId, path, cwd })
    const hostPath = resolvedTarget.hostPath
    if (typeof hostPath !== 'string' || !isAbsolute(hostPath)) {
      throw new SaveError('no-host-path', 'the workspace filesystem did not yield a writable host path', 409)
    }

    // Optimistic concurrency: the editor loaded bytes at `mtimeMs`.
    const rawMtime = url.searchParams.get('mtimeMs')
    const expectedMtimeMs = rawMtime === null ? undefined : Number(rawMtime)
    let before = null
    try {
      before = await fsp.stat(hostPath)
    } catch {
      before = null
    }
    if (before === null) {
      throw new SaveError('not-a-file', 'the target file no longer exists', 409)
    }
    if (Number.isFinite(expectedMtimeMs) && Number.isFinite(before.mtimeMs)) {
      // Tolerate sub-millisecond formatting differences from the client.
      if (Math.abs(before.mtimeMs - expectedMtimeMs) > 1.5) {
        throw new SaveError(
          'conflict',
          'the file changed on disk since it was opened; reload the document before saving',
          409,
        )
      }
    }

    const body = await readBody(req, maxSaveBytes)
    if (body.length === 0) throw new SaveError('empty-body', 'refusing to replace the document with an empty body')

    const written = await writeAtomic(hostPath, body)

    // The version the document preview tracks must be the POST-write one, so
    // the renderer can report it as rendered and the preview does not treat our
    // own save as an external change (which would reload the editor).
    let version = resolvedTarget.version
    const fs = typeof ctx.get === 'function' ? ctx.get('fs') : undefined
    if (resolvedTarget.provider === true && resolvedTarget.target !== undefined && fs && typeof fs.stat === 'function') {
      try {
        const after = await fs.stat(resolvedTarget.target)
        if (after && after.version !== undefined) version = after.version
      } catch {
        // Keep the pre-write version rather than failing a successful write.
      }
    }

    sendJson(res, 200, {
      ok: true,
      path: hostPath,
      bytes: written.bytes,
      mtimeMs: written.mtimeMs,
      version: version === undefined ? null : String(version),
    })
  } catch (error) {
    const saveError = error instanceof SaveError
      ? error
      : new SaveError('internal-error', error instanceof Error ? `${error.name}: ${error.message}` : String(error), 500)
    sendJson(res, saveError.status, { ok: false, error: { code: saveError.code, message: saveError.message } })
  }
}

/** @param {object} res @param {number} status @param {unknown} value */
function sendJson(res, status, value) {
  const body = JSON.stringify(value)
  res.writeHead(status, {
    'content-type': 'application/json; charset=utf-8',
    'content-length': String(Buffer.byteLength(body)),
    'cache-control': 'no-store',
    'x-content-type-options': 'nosniff',
  })
  res.end(body)
}

/**
 * Handle one document read (`GET /dsh-onlyoffice/read`), used by the built-in
 * preview's `renderer` loading mode: the renderer owns loading, so it reads the
 * bytes (and the file version the preview tracks) through this route.
 *
 * @param {object} ctx plugin context
 * @param {object} req node request
 * @param {object} res node response
 * @param {{ maxReadBytes?: number }} [options]
 */
export async function handleRead(ctx, req, res, options = {}) {
  const maxReadBytes = Number.isInteger(options.maxReadBytes) ? options.maxReadBytes : 512 * 1024 * 1024
  try {
    if (typeof req.method === 'string' && req.method.toUpperCase() !== 'GET') {
      throw new SaveError('method-error', 'method not allowed', 405)
    }
    let url
    try {
      url = new URL(typeof req.url === 'string' ? req.url : '/', 'http://dsh.internal')
    } catch {
      throw new SaveError('bad-request', 'malformed request URL')
    }

    const sessionId = url.searchParams.get('sessionId')
    const path = url.searchParams.get('path')
    const explicitCwd = url.searchParams.get('cwd') || undefined
    if (!isSafeSegment(sessionId) || !isSafeSegment(path)) {
      throw new SaveError('bad-request', 'sessionId and path are required')
    }
    const extension = extname(path).slice(1).toLowerCase()
    if (!READABLE_EXTENSIONS.has(extension)) {
      throw new SaveError('unsupported-extension', `refusing to read ".${extension}" as a document`, 415)
    }
    const cwd = await resolveCwd(ctx, sessionId, explicitCwd)
    if (!cwd) throw new SaveError('no-cwd', 'the session directory is unknown', 409)

    const resolvedTarget = await resolveTarget(ctx, { sessionId, path, cwd })
    const hostPath = resolvedTarget.hostPath
    if (typeof hostPath !== 'string' || !isAbsolute(hostPath)) {
      throw new SaveError('no-host-path', 'the workspace filesystem did not yield a readable host path', 409)
    }

    const info = await fsp.stat(hostPath)
    if (!info.isFile()) throw new SaveError('not-a-file', 'the target is not a regular file', 409)
    if (info.size > maxReadBytes) throw new SaveError('too-large', 'the document exceeds the read limit', 413)

    const buffer = await fsp.readFile(hostPath)
    res.writeHead(200, {
      'content-type': 'application/octet-stream',
      'content-length': String(buffer.length),
      'cache-control': 'no-store',
      'x-content-type-options': 'nosniff',
      ...(resolvedTarget.version === undefined
        ? {}
        : { 'x-dsh-file-version': Buffer.from(String(resolvedTarget.version), 'utf8').toString('base64url') }),
    })
    res.end(buffer)
  } catch (error) {
    const readError = error instanceof SaveError
      ? error
      : new SaveError('internal-error', error instanceof Error ? `${error.name}: ${error.message}` : String(error), 500)
    sendJson(res, readError.status, { ok: false, error: { code: readError.code, message: readError.message } })
  }
}

export { SaveError, isInside, realpathLoose, resolveCwd, resolveTarget, writeAtomic }
export const SAVE_ROUTE = '/dsh-onlyoffice/save'
export const SAVE_TMP_MARKER = '.~dsh-save-'
export const SEP = sep
