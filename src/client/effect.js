// SPDX-License-Identifier: AGPL-3.0-only
// dsh-onlyoffice — see LICENSE and THIRD_PARTY_NOTICES.md.

/**
 * Run one registration as a context effect (auto-disposed on plugin unload /
 * HMR). Falls back to a plain call on a context without `effect`, so an
 * unexpected host shape still registers.
 */

export function registerEffect(ctx, factory, label) {
  try {
    if (ctx && typeof ctx.effect === 'function') {
      ctx.effect(factory, label);
      return;
    }
  } catch (error) {
    console.warn(`[dsh-onlyoffice] ${label} failed`, error);
    return;
  }

  try {
    factory();
  } catch (error) {
    console.warn(`[dsh-onlyoffice] ${label} failed`, error);
  }
}
