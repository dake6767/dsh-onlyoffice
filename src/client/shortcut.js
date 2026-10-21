// SPDX-License-Identifier: AGPL-3.0-only
// dsh-onlyoffice — see LICENSE and THIRD_PARTY_NOTICES.md.

/**
 * Save shortcut predicate.
 *
 * Cmd+S on macOS, Ctrl+S elsewhere — and nothing else: Alt / Shift variants
 * belong to browser or application shortcuts (Save As, view source), so they
 * are deliberately not captured. Kept pure so it is unit-testable without a DOM.
 */

/**
 * @param {KeyboardEvent | { key?: string, metaKey?: boolean, ctrlKey?: boolean, altKey?: boolean, shiftKey?: boolean }} event
 * @returns {boolean} whether this key event is the save shortcut
 */
export function isSaveShortcut(event) {
  if (!event) return false;
  if (!(event.metaKey === true || event.ctrlKey === true)) return false;
  if (event.altKey === true || event.shiftKey === true) return false;
  return event.key === 's' || event.key === 'S';
}

/** Platform-appropriate label for the hint (`⌘S` or `Ctrl+S`). */
export function saveShortcutLabel() {
  try {
    const platform = typeof navigator === 'undefined' ? '' : String(navigator.platform || navigator.userAgent || '');
    if (/mac|iphone|ipad|ipod/i.test(platform)) return '⌘S';
  } catch {
    /* ignore */
  }
  return 'Ctrl+S';
}
