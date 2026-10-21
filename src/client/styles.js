// SPDX-License-Identifier: AGPL-3.0-only
// dsh-onlyoffice — see LICENSE and THIRD_PARTY_NOTICES.md.

/**
 * Viewer chrome stylesheet.
 *
 * Every color comes from a `--dsw-alias-*` theme token, so light/dark follow the
 * host automatically and no literal color is used for chrome. Only spacing,
 * sizing and motion values are literal. The sheet is installed once in
 * `apply()` and removed with the context's effect disposer.
 */

const STYLE_ATTRIBUTE = 'data-dsh-onlyoffice-styles';

const CSS = `
.dshoo-root {
  position: relative;
  display: flex;
  flex-direction: column;
  width: 100%;
  height: 100%;
  min-width: 0;
  min-height: 0;
  background: var(--dsw-alias-bg-base);
  color: var(--dsw-alias-label-primary);
  font: inherit;
  font-size: 12px;
  line-height: 1.4;
}

.dshoo-toolbar {
  flex: 0 0 auto;
  display: flex;
  align-items: center;
  gap: 6px;
  min-height: 34px;
  padding: 4px 8px;
  box-sizing: border-box;
  background: var(--dsw-alias-bg-layer-1);
  border-bottom: 1px solid var(--dsw-alias-border-l1);
  overflow-x: auto;
  scrollbar-width: none;
}
.dshoo-toolbar::-webkit-scrollbar { display: none; }

.dshoo-title {
  flex: 0 1 auto;
  min-width: 0;
  max-width: 42%;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  color: var(--dsw-alias-label-secondary);
  font-size: 12px;
}

.dshoo-spacer { flex: 1 1 auto; }

/* The save action lives on Cmd/Ctrl+S; this is its discoverability hint. */
.dshoo-hint-inline {
  flex: 0 0 auto;
  padding: 0 4px;
  color: var(--dsw-alias-label-secondary);
  font-size: 10px;
  line-height: 16px;
  white-space: nowrap;
  opacity: 0.75;
  user-select: none;
}

.dshoo-badge {
  flex: 0 0 auto;
  padding: 0 6px;
  border: 1px solid var(--dsw-alias-border-l2);
  border-radius: 999px;
  color: var(--dsw-alias-label-secondary);
  font-size: 10px;
  line-height: 16px;
  white-space: nowrap;
}

.dshoo-btn {
  flex: 0 0 auto;
  display: inline-flex;
  align-items: center;
  gap: 4px;
  height: 24px;
  padding: 0 8px;
  box-sizing: border-box;
  border: 1px solid var(--dsw-alias-border-l2);
  border-radius: 6px;
  background: transparent;
  color: var(--dsw-alias-label-secondary);
  font: inherit;
  font-size: 12px;
  line-height: 1;
  white-space: nowrap;
  cursor: pointer;
  user-select: none;
}
.dshoo-btn:hover:not(:disabled) {
  background: var(--dsw-alias-bg-layer-2);
  color: var(--dsw-alias-label-primary);
}
.dshoo-btn:focus-visible {
  outline: 2px solid var(--dsw-alias-brand-primary);
  outline-offset: 1px;
}
.dshoo-btn:disabled {
  opacity: 0.5;
  cursor: default;
}
.dshoo-btn[aria-pressed="true"] {
  border-color: transparent;
  background: var(--dsw-alias-brand-primary);
  color: var(--dsw-alias-bg-base);
}

.dshoo-notice {
  flex: 0 0 auto;
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 4px 10px;
  border-bottom: 1px solid var(--dsw-alias-border-l1);
  background: var(--dsw-alias-bg-layer-2);
  color: var(--dsw-alias-state-warn-primary);
  font-size: 11px;
}
.dshoo-notice > span { flex: 1 1 auto; min-width: 0; }

/* The external-change banner has its own action button. */
.dshoo-notice-action {
  color: var(--dsw-alias-label-primary);
}
.dshoo-notice-action .dshoo-btn {
  height: 22px;
  padding: 0 8px;
  font-size: 11px;
}
.dshoo-notice button {
  flex: 0 0 auto;
  border: 0;
  background: transparent;
  color: var(--dsw-alias-label-secondary);
  font: inherit;
  cursor: pointer;
  padding: 0 2px;
}
.dshoo-notice button:hover { color: var(--dsw-alias-label-primary); }

.dshoo-stage {
  position: relative;
  flex: 1 1 auto;
  min-width: 0;
  min-height: 0;
  background: var(--dsw-alias-bg-base);
}

.dshoo-frame {
  position: absolute;
  inset: 0;
  overflow: hidden;
}

.dshoo-state {
  position: absolute;
  inset: 0;
  z-index: 2;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: 10px;
  padding: 20px;
  box-sizing: border-box;
  text-align: center;
  background: var(--dsw-alias-bg-base);
  color: var(--dsw-alias-label-secondary);
}

.dshoo-state-title {
  color: var(--dsw-alias-label-primary);
  font-size: 13px;
  font-weight: 600;
}
.dshoo-state-title[data-tone="error"] { color: var(--dsw-alias-state-error-primary); }

.dshoo-hint {
  max-width: 46ch;
  color: var(--dsw-alias-label-secondary);
  font-size: 11px;
  overflow-wrap: anywhere;
}

.dshoo-state-actions {
  display: flex;
  flex-wrap: wrap;
  gap: 8px;
  justify-content: center;
  margin-top: 2px;
}

.dshoo-detail {
  max-width: 100%;
  max-height: 120px;
  overflow: auto;
  margin: 0;
  padding: 6px 8px;
  border: 1px solid var(--dsw-alias-border-l1);
  border-radius: 6px;
  background: var(--dsw-alias-bg-layer-1);
  color: var(--dsw-alias-label-secondary);
  font-family: ui-monospace, SFMono-Regular, Menlo, monospace;
  font-size: 10px;
  text-align: left;
  white-space: pre-wrap;
  overflow-wrap: anywhere;
}

.dshoo-spinner {
  width: 18px;
  height: 18px;
  border: 2px solid var(--dsw-alias-border-l2);
  border-top-color: var(--dsw-alias-brand-primary);
  border-radius: 50%;
  animation: dshoo-spin 0.9s linear infinite;
}

@keyframes dshoo-spin {
  to { transform: rotate(360deg); }
}

@media (prefers-reduced-motion: reduce) {
  .dshoo-spinner { animation-duration: 2.4s; }
}
`;

/**
 * Install the sheet into the host document once. Returns a disposer that
 * removes only the tag this function created.
 */
export function installStyles(doc = typeof document === 'undefined' ? undefined : document) {
  if (!doc || !doc.head) return () => {};
  if (typeof doc.querySelector === 'function' && doc.querySelector(`style[${STYLE_ATTRIBUTE}]`)) {
    return () => {};
  }

  const style = doc.createElement('style');
  style.setAttribute(STYLE_ATTRIBUTE, 'dsh-onlyoffice');
  style.textContent = CSS;
  doc.head.appendChild(style);

  return () => {
    try {
      style.remove();
    } catch {
      /* ignore */
    }
  };
}
