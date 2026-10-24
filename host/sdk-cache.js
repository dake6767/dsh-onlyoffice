// SPDX-License-Identifier: AGPL-3.0-only
// dsh-onlyoffice — see LICENSE and THIRD_PARTY_NOTICES.md.

/**
 * dsh-onlyoffice — pure helpers for the Host half.
 *
 * Everything in this module is a plain function over strings/paths (plus one
 * read-only cache-directory walk) so `node --test` can exercise the whole
 * contract without a DSH runtime. `host/index.js` is the only consumer that
 * touches HTTP / DSH services.
 *
 * No bare package imports: the package is installed as a linked directory
 * outside the DSH profile's `node_modules`, so only `node:*` builtins resolve.
 */
import { promises as fsp } from 'node:fs'
import { homedir } from 'node:os'
import { isAbsolute, join, relative, resolve, sep } from 'node:path'

export const PLUGIN_NAME = '@dake6767/dsh-onlyoffice'
/** Every route this plugin owns lives under this prefix. */
export const ROUTE_PREFIX = '/dsh-onlyoffice'
/** The client appends `/onlyoffice/<version>` to this (see docs/BRIEF.md §4). */
export const SDK_BASE_URL = '/dsh-onlyoffice/sdk'
/** `GET <SDK_PATH_PREFIX><version>/<rest...>` serves one SDK asset. */
export const SDK_PATH_PREFIX = '/dsh-onlyoffice/sdk/onlyoffice/'

export const DEFAULT_VERSION = '9.4.0-develop'
/**
 * Mirror candidates, in order. The first is the canonical Vercel deployment
 * (kept as `upstreamBase`); the jsDelivr and raw.githubusercontent mirrors
 * serve the same `public/packages` tree of github.com/electroluxcode/onlyoffice-web-comp
 * and were verified reachable where the Vercel deployment was not.
 */
export const DEFAULT_UPSTREAM_BASES = [
  'https://onlyoffice-web-comp.vercel.app/packages',
  'https://cdn.jsdelivr.net/gh/electroluxcode/onlyoffice-web-comp@main/public/packages',
  'https://raw.githubusercontent.com/electroluxcode/onlyoffice-web-comp/main/public/packages',
]
export const DEFAULT_UPSTREAM_BASE = DEFAULT_UPSTREAM_BASES[0]
export const DEFAULT_CACHE_DIRNAME = 'onlyoffice-sdk'
export const DEFAULT_DSH_HOME = join(homedir(), '.dsh')
export const DEFAULT_CACHE_DIR = join(DEFAULT_DSH_HOME, DEFAULT_CACHE_DIRNAME)

/** Long-lived immutable caching is safe: SDK assets never change per version. */
export const CACHE_CONTROL = 'public, max-age=31536000, immutable'

/** Infix of the in-progress temp file a cache write uses before its rename. */
export const TMP_MARKER = '.~dsh-tmp-'

/** Defensive bounds on a request path (a real SDK asset path is < 100 chars). */
const MAX_PATH_LENGTH = 4096
const MAX_SEGMENTS = 64
const MAX_SEGMENT_LENGTH = 512

/**
 * Expand a leading `~` against `home`; return `null` for "not provided".
 * @param {unknown} input
 * @param {string} [home]
 * @returns {string | null}
 */
export function expandHome(input, home = homedir()) {
  if (typeof input !== 'string') return null
  const value = input.trim()
  if (value === '') return null
  if (value === '~') return home
  if (value.startsWith('~/') || value.startsWith(`~${sep}`)) return join(home, value.slice(2))
  return value
}

/** Environment override shared by every profile on a host (multi-user setups). */
export const CACHE_DIR_ENV = 'DSH_ONLYOFFICE_CACHE_DIR'

/**
 * Resolve the on-disk cache root.
 *
 * Order: explicit plugin config > `DSH_ONLYOFFICE_CACHE_DIR` >
 * `$DSH_HOME/onlyoffice-sdk` (with `$DSH_HOME` defaulting to `~/.dsh`).
 * The environment variable exists for multi-user hosts: every platform user
 * then shares one downloaded SDK instead of paying the first-open download
 * (roughly 100 MB) once per profile.
 *
 * @param {unknown} cacheDir - explicit config value (may be `~`-prefixed/relative)
 * @param {Record<string, string | undefined>} [env]
 * @param {string} [home]
 * @returns {string} absolute path
 */
export function resolveCacheDir(cacheDir, env = process.env, home = homedir()) {
  const explicit = expandHome(cacheDir, home)
  if (explicit !== null) return resolve(explicit)
  const fromEnv = expandHome(env === null || env === undefined ? undefined : env[CACHE_DIR_ENV], home)
  if (fromEnv !== null) return resolve(fromEnv)
  const dshHome = expandHome(env === null || env === undefined ? undefined : env.DSH_HOME, home)
  const base = dshHome === null ? join(home, '.dsh') : dshHome
  return resolve(base, DEFAULT_CACHE_DIRNAME)
}

/**
 * A version becomes a path segment under the cache dir, so it must be one
 * plain, safe segment.
 * @param {unknown} value
 * @returns {boolean}
 */
export function isSafeVersion(value) {
  return typeof value === 'string'
    && value.length > 0
    && value.length <= 128
    && value !== '.'
    && value !== '..'
    && !value.includes('/')
    && !value.includes('\\')
    && !value.includes('\u0000')
    && /^[A-Za-z0-9][A-Za-z0-9._-]*$/.test(value)
}

/**
 * Strict variant of {@link normalizeUpstreamBase}: returns `null` when the
 * value is not a usable absolute http(s) base, so callers can tell
 * "invalid" from "explicitly the default".
 * @param {unknown} value
 * @returns {string | null}
 */
export function normalizeUpstreamBaseStrict(value) {
  if (typeof value !== 'string' || value.trim() === '') return null
  let url
  try {
    url = new URL(value.trim())
  } catch {
    return null
  }
  if (url.protocol !== 'http:' && url.protocol !== 'https:') return null
  if (url.username !== '' || url.password !== '') return null
  return `${url.origin}${url.pathname.replace(/\/+$/, '')}`
}

/**
 * Normalize the upstream package base: absolute http(s), no query/hash, no
 * trailing slash. Anything invalid falls back to the documented default.
 * @param {unknown} value
 * @returns {string}
 */
export function normalizeUpstreamBase(value) {
  return normalizeUpstreamBaseStrict(value) ?? DEFAULT_UPSTREAM_BASE
}

/**
 * Normalize an ordered candidate list, dropping invalid/duplicate entries.
 * @param {unknown} value
 * @returns {string[]}
 */
export function normalizeUpstreamList(value) {
  if (!Array.isArray(value)) return []
  const seen = new Set()
  const list = []
  for (const entry of value) {
    const normalized = normalizeUpstreamBaseStrict(entry)
    if (normalized === null || seen.has(normalized)) continue
    seen.add(normalized)
    list.push(normalized)
  }
  return list
}

/**
 * @param {string[]} values
 * @returns {string[]}
 */
function dedupe(values) {
  const seen = new Set()
  return values.filter((value) => {
    if (seen.has(value)) return false
    seen.add(value)
    return true
  })
}

/**
 * @typedef {{
 *   version: string,
 *   upstreamBase: string,
 *   upstreams: string[],
 *   cacheDir: string,
 *   sharedCacheDir: string | null,
 *   warnings: string[],
 * }} ResolvedConfig
 */

/**
 * Normalize the plugin row's `config` (all fields optional):
 *
 *   version            SDK version segment
 *   upstreamBase       first upstream candidate
 *   upstreamFallbacks  extra candidate bases, tried in order after
 *                      `upstreamBase`; defaults to {@link DEFAULT_UPSTREAM_BASES}
 *   cacheDir           disk cache root
 *
 * Invalid values fall back to the documented defaults (and are reported in
 * `warnings`) rather than breaking plugin activation.
 * @param {unknown} config
 * @param {Record<string, string | undefined>} [env]
 * @param {string} [home]
 * @returns {ResolvedConfig}
 */
export function resolveConfig(config, env = process.env, home = homedir()) {
  const raw = config !== null && typeof config === 'object' ? config : {}
  const warnings = []

  const versionCandidate = typeof raw.version === 'string' ? raw.version.trim() : raw.version
  let version = DEFAULT_VERSION
  if (isSafeVersion(versionCandidate)) version = versionCandidate
  else if (versionCandidate !== undefined && versionCandidate !== null) {
    warnings.push(`invalid version ${JSON.stringify(versionCandidate)}; using "${DEFAULT_VERSION}"`)
  }

  const primary = normalizeUpstreamBaseStrict(raw.upstreamBase)
  if (primary === null && raw.upstreamBase !== undefined && raw.upstreamBase !== null) {
    warnings.push(`invalid upstreamBase ${JSON.stringify(raw.upstreamBase)}; using the default mirror list`)
  }
  const upstreamBase = primary ?? DEFAULT_UPSTREAM_BASE

  let fallbacks
  if (raw.upstreamFallbacks === undefined || raw.upstreamFallbacks === null) {
    fallbacks = DEFAULT_UPSTREAM_BASES
  } else if (Array.isArray(raw.upstreamFallbacks)) {
    fallbacks = raw.upstreamFallbacks
    if (fallbacks.length > 0 && normalizeUpstreamList(fallbacks).length === 0) {
      warnings.push('upstreamFallbacks contained no usable http(s) bases; ignoring them')
    }
  } else {
    warnings.push('upstreamFallbacks must be an array; using the default mirror list')
    fallbacks = DEFAULT_UPSTREAM_BASES
  }
  const upstreams = dedupe([upstreamBase, ...normalizeUpstreamList(fallbacks)])
  if (upstreams.length === 0) upstreams.push(DEFAULT_UPSTREAM_BASE)

  return {
    version,
    upstreamBase: upstreams[0],
    upstreams,
    cacheDir: resolveCacheDir(raw.cacheDir, env, home),
    sharedCacheDir: expandHome(env?.DSH_ONLYOFFICE_SHARED_CACHE_DIR, home),
    warnings,
  }
}

/**
 * @param {number} status
 * @param {string} code
 * @param {string} message
 */
function failure(status, code, message) {
  return { ok: false, status, code, message }
}

/** @param {string} raw @returns {string | null} */
function decodeSegment(raw) {
  try {
    return decodeURIComponent(raw)
  } catch {
    return null
  }
}

/**
 * Validate one already-split raw path segment. Rejects traversal (`.`/`..`,
 * encoded or not), embedded separators (`%2f`, `%5c`), NULs and drive letters.
 * @param {string} raw
 * @returns {{ ok: true, value: string } | { ok: false, message: string }}
 */
function validateSegment(raw) {
  const decoded = decodeSegment(raw)
  if (decoded === null) return { ok: false, message: `invalid percent-encoding in "${raw}"` }
  if (decoded === '') return { ok: false, message: 'empty path segment' }
  if (decoded === '.' || decoded === '..') return { ok: false, message: `path traversal segment "${decoded}"` }
  if (decoded.includes('/') || decoded.includes('\\') || decoded.includes('\u0000')) {
    return { ok: false, message: `illegal character in path segment "${raw}"` }
  }
  if (/^[A-Za-z]:/.test(decoded)) return { ok: false, message: 'absolute drive path is not allowed' }
  if (decoded.length > MAX_SEGMENT_LENGTH) return { ok: false, message: 'path segment too long' }
  return { ok: true, value: decoded }
}

/**
 * Parse `GET /dsh-onlyoffice/sdk/onlyoffice/<version>/<rest...>`.
 *
 * Note: WHATWG URL normalization already collapses *literal* `.`/`..`
 * segments before we see `pathname`; percent-encoded forms (`%2e%2e`) survive
 * and are rejected below, which is where traversal can actually reach us.
 *
 * @param {string} pathname
 * @param {string} [prefix]
 * @returns {{ ok: true, version: string, segments: string[] }
 *   | { ok: false, status: number, code: string, message: string }}
 */
export function parseSdkPath(pathname, prefix = SDK_PATH_PREFIX) {
  if (typeof pathname !== 'string' || pathname === '') return failure(400, 'bad-path', 'missing request path')
  if (!pathname.startsWith(prefix)) return failure(404, 'not-found', 'not an OnlyOffice SDK asset path')
  if (pathname.length > MAX_PATH_LENGTH) return failure(400, 'bad-path', 'request path too long')

  const tail = pathname.slice(prefix.length)
  if (tail === '') return failure(400, 'bad-path', 'missing SDK asset path')

  const rawSegments = tail.split('/')
  if (rawSegments.length < 2) return failure(404, 'not-found', 'missing SDK asset path')
  if (rawSegments.length > MAX_SEGMENTS + 1) return failure(400, 'bad-path', 'too many path segments')

  const versionSegment = validateSegment(rawSegments[0])
  if (versionSegment.ok !== true) return failure(400, 'bad-path', `invalid SDK version: ${versionSegment.message}`)
  const version = versionSegment.value
  if (!isSafeVersion(version)) return failure(400, 'bad-path', `invalid SDK version "${version}"`)

  const segments = []
  for (const raw of rawSegments.slice(1)) {
    const checked = validateSegment(raw)
    if (checked.ok !== true) return failure(400, 'bad-path', checked.message)
    segments.push(checked.value)
  }

  return { ok: true, version, segments }
}

/** @param {string} segment @returns {boolean} */
function isSafeSegment(segment) {
  return typeof segment === 'string'
    && segment !== ''
    && segment !== '.'
    && segment !== '..'
    && !segment.includes('/')
    && !segment.includes('\\')
    && !segment.includes('\u0000')
    && !/^[A-Za-z]:/.test(segment)
}

/**
 * Cache destination for one asset. Returns `null` if the mapping would leave
 * `<cacheDir>/<version>` (defense in depth on top of {@link parseSdkPath}).
 * @param {string} cacheDir
 * @param {string} version
 * @param {string[]} segments
 * @returns {string | null}
 */
export function cachePathFor(cacheDir, version, segments) {
  if (typeof cacheDir !== 'string' || cacheDir === '') return null
  if (!isSafeVersion(version)) return null
  if (!Array.isArray(segments) || segments.length === 0) return null
  if (segments.some((segment) => !isSafeSegment(segment))) return null
  const root = join(cacheDir, version)
  let target
  try {
    target = join(root, ...segments)
  } catch {
    return null
  }
  return isPathInside(root, target) ? target : null
}

/**
 * True when `candidate` is strictly below the `root` directory.
 * @param {string} root
 * @param {string} candidate
 * @returns {boolean}
 */
export function isPathInside(root, candidate) {
  if (typeof root !== 'string' || typeof candidate !== 'string') return false
  const base = resolve(root)
  const target = resolve(candidate)
  if (target === base) return false
  const rel = relative(base, target)
  return rel !== '' && !rel.startsWith('..') && !isAbsolute(rel)
}

/**
 * Upstream URL for one asset:
 * `<upstreamBase>/onlyoffice/<version>/<rest...>`.
 * @param {string} upstreamBase
 * @param {string} version
 * @param {string[]} segments
 * @returns {string}
 */
export function upstreamUrlFor(upstreamBase, version, segments) {
  const base = typeof upstreamBase === 'string' && upstreamBase !== ''
    ? upstreamBase.replace(/\/+$/, '')
    : DEFAULT_UPSTREAM_BASE
  const path = [version, ...segments].map((segment) => encodeURIComponent(segment)).join('/')
  return `${base}/onlyoffice/${path}`
}

/**
 * Extension → response Content-Type, per docs/BRIEF.md §2.
 * Values are intentionally the brief's bare types (no charset suffix); the
 * upstream's own `application/javascript` is normalized to `text/javascript`.
 * @type {Record<string, string>}
 */
const CONTENT_TYPES = {
  '.js': 'text/javascript',
  '.mjs': 'text/javascript',
  '.cjs': 'text/javascript',
  '.html': 'text/html',
  '.htm': 'text/html',
  '.css': 'text/css',
  '.json': 'application/json',
  '.map': 'application/json',
  '.wasm': 'application/wasm',
  '.xml': 'text/xml',
  '.xsl': 'text/xml',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.webp': 'image/webp',
  '.bmp': 'image/bmp',
  '.ico': 'image/x-icon',
  '.cur': 'image/x-icon',
  '.ttf': 'font/ttf',
  '.otf': 'font/otf',
  '.ttc': 'font/collection',
  '.woff': 'font/woff',
  '.woff2': 'font/woff2',
  '.eot': 'application/vnd.ms-fontobject',
  '.txt': 'text/plain',
  '.csv': 'text/csv',
  '.pdf': 'application/pdf',
  '.mp4': 'video/mp4',
}

/**
 * Content-Type for a served asset path. Extension-less SDK font files under
 * `fonts/` (and anything unknown) fall back to `application/octet-stream`.
 * @param {string} filePath
 * @returns {string}
 */
export function contentTypeForPath(filePath) {
  if (typeof filePath !== 'string' || filePath === '') return 'application/octet-stream'
  const slash = Math.max(filePath.lastIndexOf('/'), filePath.lastIndexOf(sep))
  const name = filePath.slice(slash + 1)
  const dot = name.lastIndexOf('.')
  if (dot <= 0) return 'application/octet-stream'
  return CONTENT_TYPES[name.slice(dot).toLowerCase()] ?? 'application/octet-stream'
}

/**
 * Count the files and bytes currently cached. Understands a missing cache dir
 * (zeroes) and ignores in-progress temp files.
 * @param {string} cacheDir
 * @returns {Promise<{ files: number, bytes: number }>}
 */
export async function describeCache(cacheDir) {
  let entries
  try {
    entries = await fsp.readdir(cacheDir, { recursive: true, withFileTypes: true })
  } catch (error) {
    if (error !== null && typeof error === 'object' && (error.code === 'ENOENT' || error.code === 'ENOTDIR')) {
      return { files: 0, bytes: 0 }
    }
    throw error
  }

  let files = 0
  let bytes = 0
  for (const entry of entries) {
    if (entry.isFile() !== true) continue
    if (entry.name.includes(TMP_MARKER)) continue
    // Dotfiles are the plugin's own bookkeeping (the request log), not cache.
    if (entry.name.startsWith('.')) continue
    const parent = typeof entry.parentPath === 'string' ? entry.parentPath : cacheDir
    try {
      const info = await fsp.stat(join(parent, entry.name))
      if (info.isFile() !== true) continue
      files += 1
      bytes += info.size
    } catch {
      // Raced with a cleanup; an entry that vanished contributes nothing.
    }
  }
  return { files, bytes }
}
