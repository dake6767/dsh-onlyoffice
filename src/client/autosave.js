// SPDX-License-Identifier: AGPL-3.0-only
// dsh-onlyoffice — see LICENSE and THIRD_PARTY_NOTICES.md.

/**
 * Debounced auto-save.
 *
 * The editor reports every content change; saving on each one would run an
 * export (x2t conversion) per keystroke. This collapses bursts into one save
 * after the user pauses, while `flush()` lets the explicit save shortcut fire
 * immediately and `cancel()` drops a pending save (e.g. document closed).
 */

/** Default idle delay before an automatic save. */
export const AUTOSAVE_DELAY_MS = 1800

/**
 * @param {{
 *   run: () => Promise<void> | void,
 *   delayMs?: number,
 *   setTimer?: (fn: () => void, ms: number) => unknown,
 *   clearTimer?: (handle: unknown) => void,
 * }} options
 */
export function createAutoSaver(options) {
  const run = options && typeof options.run === 'function' ? options.run : () => {}
  const delayMs = Number.isFinite(options && options.delayMs) ? Math.max(100, Math.trunc(options.delayMs)) : AUTOSAVE_DELAY_MS
  const setTimer = (options && options.setTimer) || ((fn, ms) => setTimeout(fn, ms))
  const clearTimer = (options && options.clearTimer) || ((handle) => clearTimeout(handle))

  let timer = null
  let running = false
  let disposed = false

  const cancel = () => {
    if (timer !== null) {
      clearTimer(timer)
      timer = null
    }
  }

  const flush = async () => {
    cancel()
    if (disposed || running) return
    running = true
    try {
      await run()
    } finally {
      running = false
    }
  }

  const schedule = () => {
    if (disposed) return
    cancel()
    timer = setTimer(() => {
      timer = null
      void flush()
    }, delayMs)
  }

  return {
    schedule,
    flush,
    cancel,
    dispose() {
      disposed = true
      cancel()
    },
    isPending: () => timer !== null,
    isRunning: () => running,
  }
}
