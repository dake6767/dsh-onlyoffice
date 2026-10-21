/** Debounced auto-save: bursts collapse into one save, flush is immediate. */
import test from 'node:test'
import assert from 'node:assert/strict'

import { AUTOSAVE_DELAY_MS, createAutoSaver } from '../../src/client/autosave.js'

function fakeTimers() {
  let now = 0
  let nextId = 1
  const timers = new Map()
  return {
    setTimer(fn, ms) {
      const id = nextId++
      timers.set(id, { fn, at: now + ms })
      return id
    },
    clearTimer(id) {
      timers.delete(id)
    },
    advance(ms) {
      now += ms
      for (const [id, timer] of [...timers]) {
        if (timer.at <= now) {
          timers.delete(id)
          timer.fn()
        }
      }
    },
    get size() {
      return timers.size
    },
  }
}

test('a burst of changes produces exactly one save', async () => {
  const timers = fakeTimers()
  let saves = 0
  const saver = createAutoSaver({
    delayMs: 1000,
    run: () => {
      saves += 1
    },
    setTimer: timers.setTimer,
    clearTimer: timers.clearTimer,
  })

  saver.schedule()
  timers.advance(300)
  saver.schedule()
  timers.advance(300)
  saver.schedule()
  assert.equal(saves, 0, 'nothing runs while the user keeps editing')

  timers.advance(1000)
  assert.equal(saves, 1, 'one save after the pause')
  assert.equal(saver.isPending(), false)
})

test('flush saves immediately and cancels a pending schedule', async () => {
  const timers = fakeTimers()
  let saves = 0
  const saver = createAutoSaver({
    delayMs: AUTOSAVE_DELAY_MS,
    run: () => {
      saves += 1
    },
    setTimer: timers.setTimer,
    clearTimer: timers.clearTimer,
  })

  saver.schedule()
  await saver.flush()
  assert.equal(saves, 1)
  assert.equal(timers.size, 0, 'the pending timer is cleared')

  timers.advance(AUTOSAVE_DELAY_MS)
  assert.equal(saves, 1, 'the cancelled timer never fires twice')
})

test('dispose cancels future work', async () => {
  const timers = fakeTimers()
  let saves = 0
  const saver = createAutoSaver({
    delayMs: 500,
    run: () => {
      saves += 1
    },
    setTimer: timers.setTimer,
    clearTimer: timers.clearTimer,
  })

  saver.schedule()
  saver.dispose()
  timers.advance(5000)
  assert.equal(saves, 0)
  saver.schedule()
  assert.equal(timers.size, 0, 'a disposed saver does not schedule again')
})
