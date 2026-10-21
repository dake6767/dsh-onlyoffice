/** Save-shortcut predicate: Cmd/Ctrl+S and nothing else. */
import test from 'node:test'
import assert from 'node:assert/strict'

import { isSaveShortcut, saveShortcutLabel } from '../../src/client/shortcut.js'

test('Cmd+S and Ctrl+S are the save shortcut', () => {
  assert.equal(isSaveShortcut({ key: 's', metaKey: true }), true)
  assert.equal(isSaveShortcut({ key: 'S', metaKey: true }), true)
  assert.equal(isSaveShortcut({ key: 's', ctrlKey: true }), true)
  assert.equal(isSaveShortcut({ key: 's', metaKey: true, ctrlKey: true }), true)
})

test('other chords and plain keys are left to the host', () => {
  for (const event of [
    { key: 's' },
    { key: 'S', shiftKey: true, metaKey: true },
    { key: 's', altKey: true, metaKey: true },
    { key: 'p', metaKey: true },
    { key: 'Enter', ctrlKey: true },
    null,
    undefined,
    {},
  ]) {
    assert.equal(isSaveShortcut(event), false, JSON.stringify(event))
  }
})

test('the hint label names the platform chord', () => {
  assert.ok(['⌘S', 'Ctrl+S'].includes(saveShortcutLabel()))
})
