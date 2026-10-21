/** Byte round-trip detection used to avoid reloading after a save. */
import test from 'node:test'
import assert from 'node:assert/strict'

import { sameBytes, toUint8Array } from '../../src/client/bytes.js'

test('toUint8Array normalizes the shapes a preview can hand us', () => {
  const bytes = new Uint8Array([1, 2, 3])
  assert.equal(toUint8Array(bytes), bytes)
  assert.deepEqual([...toUint8Array(bytes.buffer)], [1, 2, 3])
  assert.deepEqual([...toUint8Array(Buffer.from([4, 5]))], [4, 5])
  assert.equal(toUint8Array(null), null)
  assert.equal(toUint8Array(undefined), null)
  assert.equal(toUint8Array('nope'), null)
})

test('sameBytes only accepts identical content', () => {
  assert.equal(sameBytes(new Uint8Array([1, 2, 3]), new Uint8Array([1, 2, 3])), true)
  assert.equal(sameBytes(new Uint8Array([1, 2, 3]).buffer, new Uint8Array([1, 2, 3])), true)
  assert.equal(sameBytes(new Uint8Array([1, 2, 3]), new Uint8Array([1, 2, 4])), false)
  assert.equal(sameBytes(new Uint8Array([1, 2]), new Uint8Array([1, 2, 3])), false)
  assert.equal(sameBytes(new Uint8Array(0), new Uint8Array(0)), true)
  assert.equal(sameBytes(null, new Uint8Array([1])), false)
  assert.equal(sameBytes(new Uint8Array([1]), undefined), false)
})
