/**
 * Static packaging guarantees for the Host half:
 *  - only `node:*` builtins and relative sibling imports (the package is a
 *    linked directory outside the profile's node_modules);
 *  - no CommonJS (`require`) sneaks in;
 *  - the entry point parses and exposes the documented plugin face.
 */
import test from 'node:test'
import assert from 'node:assert/strict'
import { readFile, readdir } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const HERE = dirname(fileURLToPath(import.meta.url))
const HOST_DIR = join(HERE, '..', '..', 'host')

const IMPORT_PATTERNS = [
  // Anchored statements only: prose inside comments also contains the word
  // "from", and that must not be mistaken for a module specifier.
  /^import\b[^'"]*?['"]([^'"]+)['"]/gms,
  /^export\b[^'"]*?\bfrom\s+['"]([^'"]+)['"]/gms,
  /import\(\s*['"]([^'"]+)['"]\s*\)/g,
]

function importedSpecifiers(source) {
  const specifiers = []
  for (const pattern of IMPORT_PATTERNS) {
    for (const match of source.matchAll(pattern)) specifiers.push(match[1])
  }
  return specifiers
}

test('host modules only import node:* builtins and relative siblings', async () => {
  const files = (await readdir(HOST_DIR)).filter((name) => name.endsWith('.js'))
  assert.ok(files.includes('index.js'), 'host/index.js must exist')
  assert.ok(files.includes('sdk-cache.js'), 'host/sdk-cache.js must exist')

  for (const file of files) {
    const source = await readFile(join(HOST_DIR, file), 'utf8')
    for (const specifier of importedSpecifiers(source)) {
      assert.ok(
        specifier.startsWith('node:') || specifier.startsWith('./') || specifier.startsWith('../'),
        `${file}: bare/unsupported import "${specifier}" would not resolve in the linked install`,
      )
    }
    assert.ok(!/\brequire\s*\(/.test(source), `${file} must not use require()`)
  }
})

test('the entry point exposes inject + apply', async () => {
  const entry = await import('../../host/index.js')
  assert.deepEqual(entry.inject, ['webServer', 'connection'])
  assert.equal(typeof entry.apply, 'function')
  assert.equal(entry.inject.includes('webServer'), true)
  assert.equal(entry.inject.includes('connection'), true)
  // version-independent: the SDK version comes from config/status, never the module.
  assert.equal(Object.hasOwn(entry, 'Config'), false)
})
