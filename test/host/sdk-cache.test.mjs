/**
 * Unit tests for the pure helpers in host/sdk-cache.js.
 * No DSH runtime, no network: every input is a string or a temp directory.
 */
import test from 'node:test'
import assert from 'node:assert/strict'
import { mkdir, writeFile } from 'node:fs/promises'
import { homedir } from 'node:os'
import { isAbsolute, join, resolve } from 'node:path'

import {
  CACHE_CONTROL,
  DEFAULT_CACHE_DIR,
  DEFAULT_UPSTREAM_BASE,
  DEFAULT_UPSTREAM_BASES,
  DEFAULT_VERSION,
  PLUGIN_NAME,
  SDK_BASE_URL,
  SDK_PATH_PREFIX,
  TMP_MARKER,
  cachePathFor,
  contentTypeForPath,
  describeCache,
  expandHome,
  isPathInside,
  isSafeVersion,
  normalizeUpstreamBase,
  normalizeUpstreamBaseStrict,
  normalizeUpstreamList,
  parseSdkPath,
  resolveCacheDir,
  resolveConfig,
  upstreamUrlFor,
} from '../../host/sdk-cache.js'
import { listFiles, makeTempDir, waitFor } from './helpers.mjs'

test('documented constants match the brief', () => {
  assert.equal(PLUGIN_NAME, '@dake6767/dsh-onlyoffice')
  assert.equal(SDK_BASE_URL, '/dsh-onlyoffice/sdk')
  assert.equal(SDK_PATH_PREFIX, '/dsh-onlyoffice/sdk/onlyoffice/')
  assert.equal(DEFAULT_VERSION, '9.4.0-develop')
  assert.equal(DEFAULT_UPSTREAM_BASE, 'https://onlyoffice-web-comp.vercel.app/packages')
  assert.deepEqual(DEFAULT_UPSTREAM_BASES, [
    'https://onlyoffice-web-comp.vercel.app/packages',
    'https://cdn.jsdelivr.net/gh/electroluxcode/onlyoffice-web-comp@main/public/packages',
    'https://raw.githubusercontent.com/electroluxcode/onlyoffice-web-comp/main/public/packages',
  ])
  assert.equal(DEFAULT_UPSTREAM_BASE, DEFAULT_UPSTREAM_BASES[0])
  assert.equal(CACHE_CONTROL, 'public, max-age=31536000, immutable')
  assert.equal(DEFAULT_CACHE_DIR, join(homedir(), '.dsh', 'onlyoffice-sdk'))
})

test('expandHome handles ~, ~/paths and ordinary values', () => {
  assert.equal(expandHome('~', '/home/t'), '/home/t')
  assert.equal(expandHome('~/x', '/home/t'), '/home/t/x')
  assert.equal(expandHome('  ~/x  ', '/home/t'), '/home/t/x')
  assert.equal(expandHome('/abs/path', '/home/t'), '/abs/path')
  assert.equal(expandHome('relative/path', '/home/t'), 'relative/path')
  assert.equal(expandHome('', '/home/t'), null)
  assert.equal(expandHome('   ', '/home/t'), null)
  assert.equal(expandHome(undefined, '/home/t'), null)
  assert.equal(expandHome(42, '/home/t'), null)
})

test('resolveCacheDir honours DSH_HOME, config and home defaults', () => {
  assert.equal(resolveCacheDir(undefined, {}, '/home/t'), '/home/t/.dsh/onlyoffice-sdk')
  assert.equal(resolveCacheDir(undefined, { DSH_HOME: '/var/dsh' }, '/home/t'), '/var/dsh/onlyoffice-sdk')
  assert.equal(resolveCacheDir(undefined, { DSH_HOME: '~/dsh' }, '/home/t'), '/home/t/dsh/onlyoffice-sdk')
  assert.equal(resolveCacheDir('/explicit/cache', { DSH_HOME: '/var/dsh' }, '/home/t'), '/explicit/cache')
  assert.equal(resolveCacheDir('~/cache', {}, '/home/t'), '/home/t/cache')
  const relative = resolveCacheDir('relative-cache', {}, '/home/t')
  assert.ok(isAbsolute(relative))
  assert.equal(relative, resolve('relative-cache'))
})

test('isSafeVersion accepts version tags and rejects path-shaped values', () => {
  assert.equal(isSafeVersion('9.4.0-develop'), true)
  assert.equal(isSafeVersion('8.2.0.1'), true)
  assert.equal(isSafeVersion('v1'), true)
  for (const bad of ['', ' ', '.', '..', '../evil', 'a/b', 'a\\b', '/abs', 'a\u0000b', 42, null, undefined, {}]) {
    assert.equal(isSafeVersion(bad), false, `expected ${JSON.stringify(bad)} to be unsafe`)
  }
})

test('normalizeUpstreamBase strips trailing slash/query and validates protocol', () => {
  assert.equal(normalizeUpstreamBase('https://x.test/packages/'), 'https://x.test/packages')
  assert.equal(normalizeUpstreamBase('https://x.test/packages'), 'https://x.test/packages')
  assert.equal(normalizeUpstreamBase('https://x.test/'), 'https://x.test')
  assert.equal(normalizeUpstreamBase('https://x.test/p/a///'), 'https://x.test/p/a')
  assert.equal(normalizeUpstreamBase('https://x.test/p/?q=1#f'), 'https://x.test/p')
  assert.equal(normalizeUpstreamBase('  https://x.test/p  '), 'https://x.test/p')
  assert.equal(normalizeUpstreamBase('ftp://x.test/p'), DEFAULT_UPSTREAM_BASE)
  assert.equal(normalizeUpstreamBase('not a url'), DEFAULT_UPSTREAM_BASE)
  assert.equal(normalizeUpstreamBase(''), DEFAULT_UPSTREAM_BASE)
  assert.equal(normalizeUpstreamBase(undefined), DEFAULT_UPSTREAM_BASE)
})

test('resolveConfig falls back to defaults for invalid/missing values', () => {
  const defaults = resolveConfig(undefined, {}, '/home/t')
  assert.deepEqual(defaults, {
    version: DEFAULT_VERSION,
    upstreamBase: DEFAULT_UPSTREAM_BASE,
    upstreams: DEFAULT_UPSTREAM_BASES,
    cacheDir: '/home/t/.dsh/onlyoffice-sdk',
    warnings: [],
  })
  assert.deepEqual(resolveConfig(null, {}, '/home/t'), defaults)
  assert.deepEqual(resolveConfig('not-an-object', {}, '/home/t'), defaults)

  assert.deepEqual(resolveConfig({
    version: '9.4.0-develop',
    upstreamBase: 'https://mirror.test/packages/',
    upstreamFallbacks: [],
    cacheDir: '~/sdk-cache',
  }, {}, '/home/t'), {
    version: '9.4.0-develop',
    upstreamBase: 'https://mirror.test/packages',
    upstreams: ['https://mirror.test/packages'],
    cacheDir: '/home/t/sdk-cache',
    warnings: [],
  })

  const invalid = resolveConfig({
    version: '../evil',
    upstreamBase: 'ftp://nope',
    cacheDir: '   ',
  }, { DSH_HOME: '/var/dsh' }, '/home/t')
  assert.equal(invalid.version, DEFAULT_VERSION)
  assert.equal(invalid.upstreamBase, DEFAULT_UPSTREAM_BASE)
  assert.deepEqual(invalid.upstreams, DEFAULT_UPSTREAM_BASES)
  assert.equal(invalid.cacheDir, '/var/dsh/onlyoffice-sdk')
  assert.equal(invalid.warnings.length, 2)
})

test('upstream candidates: primary first, configured fallbacks after it', () => {
  const custom = resolveConfig({
    upstreamBase: 'https://mirror.test/packages',
    upstreamFallbacks: ['https://b.test/p', 'https://c.test/p'],
  }, {}, '/home/t')
  assert.deepEqual(custom.upstreams, [
    'https://mirror.test/packages',
    'https://b.test/p',
    'https://c.test/p',
  ])
  assert.equal(custom.upstreamBase, 'https://mirror.test/packages')

  // Explicit primary + default fallbacks: primary first, then the three
  // documented mirrors (the default primary is not the configured one here).
  const primaryOnly = resolveConfig({ upstreamBase: 'https://mirror.test/packages' }, {}, '/home/t')
  assert.deepEqual(primaryOnly.upstreams, [
    'https://mirror.test/packages',
    ...DEFAULT_UPSTREAM_BASES,
  ])

  // An explicitly empty fallback list disables failover.
  const noFallbacks = resolveConfig({ upstreamFallbacks: [] }, {}, '/home/t')
  assert.deepEqual(noFallbacks.upstreams, [DEFAULT_UPSTREAM_BASE])

  // Duplicate candidates collapse; invalid entries are dropped; a non-array
  // value is warned about and replaced by the defaults.
  const messy = resolveConfig({
    upstreamFallbacks: ['https://b.test/p', 'not a url', 'https://b.test/p', 'ftp://x/p'],
  }, {}, '/home/t')
  assert.deepEqual(messy.upstreams, [DEFAULT_UPSTREAM_BASE, 'https://b.test/p'])

  const notArray = resolveConfig({ upstreamFallbacks: 'https://b.test/p' }, {}, '/home/t')
  assert.deepEqual(notArray.upstreams, DEFAULT_UPSTREAM_BASES)
  assert.equal(notArray.warnings.length, 1)
})

test('normalizeUpstreamList and the strict normalizer', () => {
  assert.equal(normalizeUpstreamBaseStrict('https://a.test/p/'), 'https://a.test/p')
  assert.equal(normalizeUpstreamBaseStrict('ftp://a.test/p'), null)
  assert.equal(normalizeUpstreamBaseStrict('https://user:pass@a.test/p'), null)
  assert.equal(normalizeUpstreamBaseStrict(42), null)
  assert.deepEqual(
    normalizeUpstreamList(['https://a.test/p', 'https://a.test/p/', 'nope', 'https://b.test/p']),
    ['https://a.test/p', 'https://b.test/p'],
  )
  assert.deepEqual(normalizeUpstreamList('not-an-array'), [])
})

test('parseSdkPath maps the documented SDK URL shape', () => {
  const parsed = parseSdkPath('/dsh-onlyoffice/sdk/onlyoffice/9.4.0-develop/sdkjs/word/sdk-all.js')
  assert.deepEqual(parsed, {
    ok: true,
    version: '9.4.0-develop',
    segments: ['sdkjs', 'word', 'sdk-all.js'],
  })

  const nested = parseSdkPath('/dsh-onlyoffice/sdk/onlyoffice/9.4.0-develop/web-apps/apps/api/documents/api.js')
  assert.equal(nested.ok, true)
  if (nested.ok) assert.equal(nested.segments.join('/'), 'web-apps/apps/api/documents/api.js')
})

test('parseSdkPath decodes percent-encoded segments', () => {
  const parsed = parseSdkPath('/dsh-onlyoffice/sdk/onlyoffice/9.4.0-develop/fonts/Open%20Sans.ttf')
  assert.deepEqual(parsed, {
    ok: true,
    version: '9.4.0-develop',
    segments: ['fonts', 'Open Sans.ttf'],
  })
})

test('parseSdkPath rejects traversal and unsafe segments with 400', () => {
  const unsafe = [
    '/dsh-onlyoffice/sdk/onlyoffice/9.4.0-develop/%2e%2e/secret.js',
    '/dsh-onlyoffice/sdk/onlyoffice/9.4.0-develop/%2E%2E/secret.js',
    '/dsh-onlyoffice/sdk/onlyoffice/9.4.0-develop/a/%2e%2e/%2e%2e/secret.js',
    '/dsh-onlyoffice/sdk/onlyoffice/9.4.0-develop/../secret.js',
    '/dsh-onlyoffice/sdk/onlyoffice/9.4.0-develop/./sdk-all.js',
    '/dsh-onlyoffice/sdk/onlyoffice/9.4.0-develop/a//b.js',
    '/dsh-onlyoffice/sdk/onlyoffice/9.4.0-develop/a%2Fb.js',
    '/dsh-onlyoffice/sdk/onlyoffice/9.4.0-develop/a%5Cb.js',
    '/dsh-onlyoffice/sdk/onlyoffice/9.4.0-develop/a%00b.js',
    '/dsh-onlyoffice/sdk/onlyoffice/9.4.0-develop/C%3A/evil.js',
    '/dsh-onlyoffice/sdk/onlyoffice/9.4.0-develop/%zz/bad.js',
    '/dsh-onlyoffice/sdk/onlyoffice/%2e%2e/sdk-all.js',
  ]
  for (const pathname of unsafe) {
    const parsed = parseSdkPath(pathname)
    assert.equal(parsed.ok, false, `expected rejection: ${pathname}`)
    if (parsed.ok !== true) {
      assert.equal(parsed.status, 400, `expected 400 for ${pathname}`)
      assert.equal(parsed.code, 'bad-path')
      assert.ok(parsed.message.length > 0)
    }
  }
})

test('parseSdkPath reports missing assets and foreign paths', () => {
  const missingRest = parseSdkPath('/dsh-onlyoffice/sdk/onlyoffice/9.4.0-develop')
  assert.equal(missingRest.ok, false)
  if (missingRest.ok !== true) assert.equal(missingRest.status, 404)

  const trailingSlash = parseSdkPath('/dsh-onlyoffice/sdk/onlyoffice/9.4.0-develop/')
  assert.equal(trailingSlash.ok, false)
  if (trailingSlash.ok !== true) assert.equal(trailingSlash.status, 400)

  const foreign = parseSdkPath('/dsh-onlyoffice/status')
  assert.equal(foreign.ok, false)
  if (foreign.ok !== true) {
    assert.equal(foreign.status, 404)
    assert.equal(foreign.code, 'not-found')
  }

  const tooLong = parseSdkPath(`/dsh-onlyoffice/sdk/onlyoffice/9.4.0-develop/${'a'.repeat(5000)}.js`)
  assert.equal(tooLong.ok, false)
  if (tooLong.ok !== true) assert.equal(tooLong.status, 400)
})

test('cachePathFor maps a request into <cacheDir>/<version>/...', () => {
  assert.equal(
    cachePathFor('/cache', '9.4.0-develop', ['sdkjs', 'word', 'sdk-all.js']),
    join('/cache', '9.4.0-develop', 'sdkjs', 'word', 'sdk-all.js'),
  )
  assert.equal(cachePathFor('/cache', '../evil', ['a.js']), null)
  assert.equal(cachePathFor('/cache', '9.4.0-develop', []), null)
  assert.equal(cachePathFor('/cache', '9.4.0-develop', ['']), null)
  assert.equal(cachePathFor('/cache', '9.4.0-develop', ['..', 'evil']), null)
  assert.equal(cachePathFor('/cache', '9.4.0-develop', ['/abs']), null)
})

test('isPathInside only accepts strict descendants', () => {
  assert.equal(isPathInside('/root', '/root/a/b'), true)
  assert.equal(isPathInside('/root', '/root/a/../b'), true)
  assert.equal(isPathInside('/root', '/root'), false)
  assert.equal(isPathInside('/root', '/rooted'), false)
  assert.equal(isPathInside('/root', '/root/../escape'), false)
  assert.equal(isPathInside('/root', '/absolute/elsewhere'), false)
})

test('upstreamUrlFor builds the documented upstream URL', () => {
  assert.equal(
    upstreamUrlFor('https://onlyoffice-web-comp.vercel.app/packages', '9.4.0-develop', ['x2t', 'x2t.wasm']),
    'https://onlyoffice-web-comp.vercel.app/packages/onlyoffice/9.4.0-develop/x2t/x2t.wasm',
  )
  assert.equal(
    upstreamUrlFor('https://x.test/packages/', '9.4.0-develop', ['fonts', 'Open Sans.ttf']),
    'https://x.test/packages/onlyoffice/9.4.0-develop/fonts/Open%20Sans.ttf',
  )
})

test('contentTypeForPath follows the brief table', () => {
  const cases = [
    ['a.js', 'text/javascript'],
    ['a.min.js', 'text/javascript'],
    ['a.mjs', 'text/javascript'],
    ['a.cjs', 'text/javascript'],
    ['index.html', 'text/html'],
    ['index.htm', 'text/html'],
    ['a.css', 'text/css'],
    ['a.json', 'application/json'],
    ['a.js.map', 'application/json'],
    ['x2t.wasm', 'application/wasm'],
    ['a.xml', 'text/xml'],
    ['logo.svg', 'image/svg+xml'],
    ['logo.png', 'image/png'],
    ['photo.jpg', 'image/jpeg'],
    ['photo.jpeg', 'image/jpeg'],
    ['a.gif', 'image/gif'],
    ['a.webp', 'image/webp'],
    ['a.ico', 'image/x-icon'],
    ['font.ttf', 'font/ttf'],
    ['font.otf', 'font/otf'],
    ['font.woff', 'font/woff'],
    ['font.woff2', 'font/woff2'],
    ['font.eot', 'application/vnd.ms-fontobject'],
    ['notes.txt', 'text/plain'],
    ['data.csv', 'text/csv'],
    ['a.pdf', 'application/pdf'],
  ]
  for (const [file, expected] of cases) {
    assert.equal(contentTypeForPath(`/cache/9.4.0-develop/${file}`), expected, file)
  }
  assert.equal(contentTypeForPath('/cache/fonts/font-without-extension'), 'application/octet-stream')
  assert.equal(contentTypeForPath('/cache/a.unknownext'), 'application/octet-stream')
  assert.equal(contentTypeForPath('/cache/.hidden'), 'application/octet-stream')
  assert.equal(contentTypeForPath(''), 'application/octet-stream')
  assert.equal(contentTypeForPath(undefined), 'application/octet-stream')
  // Case-insensitive extension matching.
  assert.equal(contentTypeForPath('/cache/FONT.TTF'), 'font/ttf')
  assert.equal(contentTypeForPath('/cache/APP.JS'), 'text/javascript')
})

test('describeCache walks the tree, sums bytes and ignores temp files', async (t) => {
  const dir = await makeTempDir(t)
  assert.deepEqual(await describeCache(dir), { files: 0, bytes: 0 })
  assert.deepEqual(await describeCache(join(dir, 'missing')), { files: 0, bytes: 0 })

  await mkdir(join(dir, '9.4.0-develop', 'sdkjs', 'word'), { recursive: true })
  await writeFile(join(dir, '9.4.0-develop', 'sdkjs', 'word', 'sdk-all.js'), Buffer.alloc(11))
  await writeFile(join(dir, '9.4.0-develop', 'x2t.wasm'), Buffer.alloc(7))
  const leftover = join(dir, '9.4.0-develop', `partial.js${TMP_MARKER}123-1`)
  await writeFile(leftover, Buffer.alloc(99))

  const stats = await describeCache(dir)
  assert.deepEqual(stats, { files: 2, bytes: 18 })

  const listed = await listFiles(dir)
  assert.ok(listed.some((name) => name.includes(TMP_MARKER)))
})

test('describeCache eventually sees atomically renamed files', async (t) => {
  const dir = await makeTempDir(t)
  await writeFile(join(dir, 'a.js'), 'payload')
  await waitFor(async () => (await describeCache(dir)).files === 1)
  assert.deepEqual(await describeCache(dir), { files: 1, bytes: 7 })
})

test('DSH_ONLYOFFICE_CACHE_DIR lets a multi-user host share one cache', async () => {
  const shared = '/var/lib/dsh-phalanx/data/onlyoffice-sdk'
  assert.equal(resolveCacheDir(undefined, { DSH_ONLYOFFICE_CACHE_DIR: shared }, '/home/u'), shared)
  // explicit plugin config still wins over the environment
  assert.equal(resolveCacheDir('/tmp/plugin-cache', { DSH_ONLYOFFICE_CACHE_DIR: shared }, '/home/u'), '/tmp/plugin-cache')
})
