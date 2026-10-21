/** Page-relative route resolution: root, reverse-proxy prefix, desktop shell. */
import test from 'node:test'
import assert from 'node:assert/strict'

import { resolvePageRelative } from '../../src/client/page-url.js'

test('origin-root deployment', () => {
  assert.equal(
    resolvePageRelative('dsh-onlyoffice/status', 'http://127.0.0.1:53803/'),
    'http://127.0.0.1:53803/dsh-onlyoffice/status',
  )
  assert.equal(
    resolvePageRelative('/dsh-onlyoffice/sdk', 'http://127.0.0.1:53803/'),
    'http://127.0.0.1:53803/dsh-onlyoffice/sdk',
  )
})

test('reverse-proxy directory keeps the prefix (the server deployment case)', () => {
  const base = 'http://118.193.35.18:18080/app/9333880fd9deb9962689420d314ad43e/'
  assert.equal(
    resolvePageRelative('dsh-onlyoffice/sdk', base),
    'http://118.193.35.18:18080/app/9333880fd9deb9962689420d314ad43e/dsh-onlyoffice/sdk',
  )
  // A leading slash must NOT escape back to the origin root.
  assert.equal(
    resolvePageRelative('/dsh-onlyoffice/sdk', base),
    'http://118.193.35.18:18080/app/9333880fd9deb9962689420d314ad43e/dsh-onlyoffice/sdk',
  )
})

test('a base without its trailing slash still appends, not replaces', () => {
  assert.equal(
    resolvePageRelative('dsh-onlyoffice/status', 'http://host/app/hash'),
    'http://host/app/hash/dsh-onlyoffice/status',
  )
})

test('the desktop shell scheme is preserved', () => {
  assert.equal(
    resolvePageRelative('dsh-onlyoffice/status', 'dsh-app://app/'),
    'dsh-app://app/dsh-onlyoffice/status',
  )
})

test('no usable base falls back to origin, then to a root-relative path', () => {
  assert.equal(
    resolvePageRelative('dsh-onlyoffice/sdk', undefined, 'http://host:1'),
    'http://host:1/dsh-onlyoffice/sdk',
  )
  assert.equal(resolvePageRelative('dsh-onlyoffice/sdk', undefined, 'null'), '/dsh-onlyoffice/sdk')
})

test('query and fragment in the base are dropped, trailing slash trimmed', () => {
  assert.equal(
    resolvePageRelative('dsh-onlyoffice/sdk', 'http://host/app/?token=abc#x'),
    'http://host/app/dsh-onlyoffice/sdk',
  )
})
