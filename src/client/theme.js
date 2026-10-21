// SPDX-License-Identifier: AGPL-3.0-only
// dsh-onlyoffice — see LICENSE and THIRD_PARTY_NOTICES.md.

/**
 * Host light/dark detection for the embedded OnlyOffice UI theme.
 *
 * DSH toggles light/dark by swapping `--dsw-alias-*` token VALUES, so the
 * document does not necessarily carry a `dark` class we can rely on. The
 * authoritative, implementation-agnostic signal is the resolved value of the
 * base background token; the remaining checks are fallbacks for hosts that
 * style with color-scheme / explicit theme attributes.
 *
 * Nothing here paints anything: the values are only inspected to decide
 * `editorConfig.customization.uiTheme` (OFFICE_THEME.DARK vs OFFICE_THEME.WHITE).
 */

const DARK = 'dark';
const LIGHT = 'light';

const THEME_ATTRIBUTES = ['data-theme', 'data-color-scheme', 'data-dsw-theme', 'data-mode'];

function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

function srgbToLuminance(r, g, b) {
  const channel = (value) => {
    const v = clamp(value, 0, 255) / 255;
    return v <= 0.04045 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4;
  };
  return 0.2126 * channel(r) + 0.7152 * channel(g) + 0.0722 * channel(b);
}

function parseHexColor(text) {
  const match = /^#([0-9a-f]{3,8})$/i.exec(text);
  if (!match) return null;
  let hex = match[1];
  if (hex.length === 3 || hex.length === 4) {
    hex = hex
      .split('')
      .map((ch) => ch + ch)
      .join('');
  }
  if (hex.length !== 6 && hex.length !== 8) return null;
  return {
    r: parseInt(hex.slice(0, 2), 16),
    g: parseInt(hex.slice(2, 4), 16),
    b: parseInt(hex.slice(4, 6), 16),
  };
}

function parseRgbColor(text) {
  const match = /^rgba?\(([^)]+)\)$/i.exec(text);
  if (!match) return null;
  const parts = match[1]
    .replace(/\//g, ' ')
    .split(/[\s,]+/)
    .filter(Boolean);
  if (parts.length < 3) return null;
  const channel = (raw) => {
    if (raw.endsWith('%')) return (parseFloat(raw) / 100) * 255;
    return parseFloat(raw);
  };
  const r = channel(parts[0]);
  const g = channel(parts[1]);
  const b = channel(parts[2]);
  if (![r, g, b].every((value) => Number.isFinite(value))) return null;
  return { r, g, b };
}

function parseHslColor(text) {
  const match = /^hsla?\(([^)]+)\)$/i.exec(text);
  if (!match) return null;
  const parts = match[1]
    .replace(/\//g, ' ')
    .replace(/,/g, ' ')
    .split(/\s+/)
    .filter(Boolean);
  if (parts.length < 3) return null;
  const h = parseFloat(parts[0]);
  const s = parseFloat(parts[1]) / 100;
  const l = parseFloat(parts[2]) / 100;
  if (![h, s, l].every((value) => Number.isFinite(value))) return null;
  // A quick HSL → RGB conversion is enough to rank lightness.
  const c = (1 - Math.abs(2 * l - 1)) * s;
  const hp = (((h % 360) + 360) % 360) / 60;
  const x = c * (1 - Math.abs((hp % 2) - 1));
  const [r1, g1, b1] =
    hp < 1
      ? [c, x, 0]
      : hp < 2
        ? [x, c, 0]
        : hp < 3
          ? [0, c, x]
          : hp < 4
            ? [0, x, c]
            : hp < 5
              ? [x, 0, c]
              : [c, 0, x];
  const m = l - c / 2;
  return { r: (r1 + m) * 255, g: (g1 + m) * 255, b: (b1 + m) * 255 };
}

/** Perceptual lightness of a modern color function (0..1), when parseable. */
function parsePerceptualLightness(text) {
  const match = /^(?:oklch|oklab|lab|lch)\(\s*([\d.]+)(%?)/i.exec(text);
  if (!match) return null;
  let value = parseFloat(match[1]);
  if (!Number.isFinite(value)) return null;
  if (match[2] === '%') value /= 100;
  // oklch/oklab: 0..1; CSS lab/lch L is 0..100.
  if (value > 1) value /= 100;
  return value;
}

/** Returns `true`/`false` for a parsed CSS color, or `null` when unparseable. */
function colorIsDark(raw) {
  const text = String(raw || '')
    .trim()
    .toLowerCase();
  if (!text) return null;

  if (text === 'white' || text === 'transparent') return false;
  if (text === 'black') return true;

  const rgb =
    parseHexColor(text) || parseRgbColor(text) || parseHslColor(text);
  if (rgb) return srgbToLuminance(rgb.r, rgb.g, rgb.b) < 0.4;

  const lightness = parsePerceptualLightness(text);
  if (lightness !== null) return lightness < 0.55;

  return null;
}

function readTokenStyle(win, doc, token) {
  const roots = [doc.documentElement, doc.body].filter(Boolean);
  for (const root of roots) {
    try {
      const style = win.getComputedStyle(root);
      const value = style.getPropertyValue(token);
      if (value && value.trim()) return value.trim();
    } catch {
      /* keep looking */
    }
  }
  return '';
}

function readColorScheme(win, doc) {
  const roots = [doc.documentElement, doc.body].filter(Boolean);
  for (const root of roots) {
    try {
      const scheme = String(win.getComputedStyle(root).colorScheme || '').toLowerCase();
      if (scheme.includes('dark')) return DARK;
      if (scheme.includes('light')) return LIGHT;
    } catch {
      /* keep looking */
    }
  }
  return '';
}

function readAttributeSignal(doc) {
  const elements = [doc.documentElement, doc.body].filter(Boolean);
  for (const element of elements) {
    for (const name of THEME_ATTRIBUTES) {
      let value;
      try {
        value = element.getAttribute(name);
      } catch {
        continue;
      }
      if (!value) continue;
      const normalized = String(value).toLowerCase();
      if (normalized === DARK || normalized.endsWith('-dark')) return DARK;
      if (normalized === LIGHT || normalized.endsWith('-light')) return LIGHT;
    }

    try {
      const classes = element.classList;
      if (classes && typeof classes.contains === 'function') {
        if (classes.contains('dark') || classes.contains('theme-dark')) return DARK;
        if (classes.contains('light') || classes.contains('theme-light')) return LIGHT;
      }
    } catch {
      /* ignore */
    }
  }
  return '';
}

function readMediaSignal(win) {
  try {
    const query = win.matchMedia?.('(prefers-color-scheme: dark)');
    if (query && typeof query.matches === 'boolean') return query.matches ? DARK : LIGHT;
  } catch {
    /* ignore */
  }
  return '';
}

/** `true` when the host is currently in dark mode. */
export function isHostDark(win = typeof window === 'undefined' ? undefined : window) {
  if (!win || !win.document) return false;
  const doc = win.document;

  const tokenDark = colorIsDark(readTokenStyle(win, doc, '--dsw-alias-bg-base'));
  if (tokenDark !== null) return tokenDark;

  // A second token is cheap insurance when the base token is a color-mix().
  const labelDark = colorIsDark(readTokenStyle(win, doc, '--dsw-alias-label-primary'));
  if (labelDark !== null) return !labelDark;

  return (
    readColorScheme(win, doc) === DARK ||
    readAttributeSignal(doc) === DARK ||
    readMediaSignal(win) === DARK
  );
}

/**
 * Watch the host theme. `listener(dark: boolean)` fires only when the resolved
 * mode actually changes. Returns an unsubscribe function.
 */
export function subscribeHostTheme(listener, win = typeof window === 'undefined' ? undefined : window) {
  if (!win || !win.document) return () => {};

  const doc = win.document;
  let current = isHostDark(win);
  let scheduled = false;

  const notify = () => {
    scheduled = false;
    const next = isHostDark(win);
    if (next === current) return;
    current = next;
    try {
      listener(next);
    } catch {
      /* a listener must not break the theme watcher */
    }
  };

  const schedule = () => {
    if (scheduled) return;
    scheduled = true;
    try {
      win.requestAnimationFrame(notify);
    } catch {
      win.setTimeout(notify, 16);
    }
  };

  let observer;
  try {
    observer = new win.MutationObserver(schedule);
    const targets = [doc.documentElement, doc.body].filter(Boolean);
    for (const target of targets) {
      observer.observe(target, {
        attributes: true,
        attributeFilter: ['class', 'style', ...THEME_ATTRIBUTES],
      });
    }
  } catch {
    observer = undefined;
  }

  let mediaQuery;
  let mediaListener;
  try {
    mediaQuery = win.matchMedia?.('(prefers-color-scheme: dark)');
    if (mediaQuery && typeof mediaQuery.addEventListener === 'function') {
      mediaListener = () => schedule();
      mediaQuery.addEventListener('change', mediaListener);
    }
  } catch {
    mediaQuery = undefined;
  }

  return () => {
    try {
      observer?.disconnect();
    } catch {
      /* ignore */
    }
    try {
      if (mediaQuery && mediaListener) mediaQuery.removeEventListener('change', mediaListener);
    } catch {
      /* ignore */
    }
  };
}
