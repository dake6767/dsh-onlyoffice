// SPDX-License-Identifier: AGPL-3.0-only
// dsh-onlyoffice — see LICENSE and THIRD_PARTY_NOTICES.md.

/**
 * Host-half build step.
 *
 * `host/**` is authored as plain ESM (no bundler, only `node:*` imports and
 * sibling modules) and copied verbatim into `lib/`, which package.json's
 * `main`/`exports` advertise. A copy — not a bundle — so the Host process
 * loads exactly the reviewed source.
 */
import { copyFile, mkdir, readdir } from 'node:fs/promises'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const from = resolve(root, 'host')
const to = resolve(root, 'lib')

await mkdir(to, { recursive: true })
const entries = await readdir(from, { withFileTypes: true })
const copied = []
for (const entry of entries) {
  if (!entry.isFile() || !entry.name.endsWith('.js')) continue
  await copyFile(resolve(from, entry.name), resolve(to, entry.name))
  copied.push(entry.name)
}
if (!copied.includes('index.js')) {
  throw new Error(`host/index.js is missing (found: ${copied.join(', ') || 'nothing'})`)
}
await copyFile(resolve(from, 'index.js'), resolve(to, 'index.js'))
console.log(`[dsh-onlyoffice] host half copied -> lib/ (${copied.sort().join(', ')})`)
