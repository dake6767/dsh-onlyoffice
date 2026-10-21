// SPDX-License-Identifier: AGPL-3.0-only
// dsh-onlyoffice — see LICENSE and THIRD_PARTY_NOTICES.md.

/**
 * Byte helpers used to detect a save round-trip.
 *
 * After a successful write-back the document preview's file watcher sees the
 * new version and re-reads the file. Those bytes are exactly the ones we just
 * saved, so the viewer can keep its live editor instead of rebuilding it (a
 * visible reload after every Cmd/Ctrl+S).
 */

/**
 * Normalize an ArrayBuffer / typed array / Buffer into a Uint8Array view.
 * @param {unknown} value
 * @returns {Uint8Array | null}
 */
export function toUint8Array(value) {
  if (!value) return null;
  if (value instanceof Uint8Array) return value;
  if (typeof ArrayBuffer !== 'undefined' && value instanceof ArrayBuffer) return new Uint8Array(value);
  if (typeof ArrayBuffer !== 'undefined' && ArrayBuffer.isView && ArrayBuffer.isView(value)) {
    return new Uint8Array(value.buffer, value.byteOffset, value.byteLength);
  }
  return null;
}

/**
 * Compare two byte sources without allocating. Used only on a save round-trip,
 * so the linear scan is fine (a few MB at most).
 * @param {unknown} left
 * @param {unknown} right
 * @returns {boolean} whether both are present and byte-identical
 */
export function sameBytes(left, right) {
  const a = toUint8Array(left);
  const b = toUint8Array(right);
  if (!a || !b) return false;
  if (a.byteLength !== b.byteLength) return false;
  if (a.byteLength === 0) return true;
  for (let index = 0; index < a.byteLength; index += 1) {
    if (a[index] !== b[index]) return false;
  }
  return true;
}
