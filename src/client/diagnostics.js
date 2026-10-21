// SPDX-License-Identifier: AGPL-3.0-only
// dsh-onlyoffice — see LICENSE and THIRD_PARTY_NOTICES.md.

/**
 * Browser-side diagnostics for the OnlyOffice views.
 *
 * A failing editor is almost always the SDK's own iframe (its scripts run in a
 * same-origin iframe we own), and those errors never reach our `catch`. This
 * module listens at both levels, keeps a small ring buffer, and lets the error
 * UI show the underlying cause instead of a bare "engine failed".
 *
 * It also reports, on a strict budget, to the Host route
 * `POST /dsh-onlyoffice/diagnostic`, which appends the lines to the same
 * `.requests.log` the Host writes assets requests to — so a stuck session can
 * be diagnosed from the Host side.
 */

const MAX_ENTRIES = 16;
const MAX_POSTS = 24;
const TEXT_LIMIT = 900;

/** @type {{ at: number, kind: string, text: string }[]} */
const entries = [];
let installed = false;
let posts = 0;
let poster = null;
let reportChain = Promise.resolve();

function describeErrorEvent(event) {
  if (!event) return 'unknown error';
  const parts = [];
  if (event.message) parts.push(String(event.message));
  if (event.filename) parts.push(`@ ${event.filename}${event.lineno ? `:${event.lineno}` : ''}${event.colno ? `:${event.colno}` : ''}`);
  if (event.error && event.error.stack) parts.push(String(event.error.stack));
  return parts.join(' ') || 'unknown error';
}

function describeRejection(event) {
  const reason = event ? event.reason : undefined;
  if (!reason) return 'unhandled rejection';
  if (typeof reason === 'string') return reason;
  return reason.stack ? String(reason.stack) : String(reason.message || reason);
}

function describeResourceError(event) {
  const target = event ? event.target : undefined;
  if (!target || target === window) return '';
  const url = target.src || target.href || '';
  if (!url) return '';
  return `${target.tagName || 'resource'} failed to load: ${url}`;
}

/**
 * Record one diagnostic line. Safe to call from anywhere, never throws.
 * @param {string} kind short tag (`page-error`, `iframe-error`, `resource`, …)
 * @param {string} text human-readable detail
 */
export function recordDiagnostic(kind, text) {
  try {
    const entry = { at: Date.now(), kind: String(kind), text: String(text).slice(0, TEXT_LIMIT) };
    entries.push(entry);
    if (entries.length > MAX_ENTRIES) entries.shift();
    const payload = `[${entry.kind}] ${entry.text}`;
    if (poster && posts < MAX_POSTS) {
      posts += 1;
      reportChain = reportChain.then(() => poster(payload)).catch(() => {});
    }
  } catch {
    /* diagnostics must never break the view */
  }
}

/** @returns {{ at: number, kind: string, text: string }[]} recent lines, oldest first */
export function recentDiagnostics() {
  return entries.slice();
}

/** One short block for the error UI (`<pre>` detail). */
export function diagnosticsDetail(extra) {
  const lines = [];
  if (extra) lines.push(String(extra));
  for (const entry of entries) lines.push(`[${entry.kind}] ${entry.text}`);
  return lines.join('\n');
}

/**
 * Install page-level listeners once. `post` receives one formatted string per
 * diagnostic and must never throw.
 * @param {(text: string) => Promise<void> | void} [post]
 */
export function installDiagnostics(post) {
  if (typeof post === 'function') poster = post;
  if (installed) return () => {};
  if (typeof window === 'undefined' || typeof window.addEventListener !== 'function') return () => {};
  installed = true;

  const onError = (event) => {
    const resource = describeResourceError(event);
    if (resource) {
      recordDiagnostic('resource', resource);
      return;
    }
    recordDiagnostic('page-error', describeErrorEvent(event));
  };
  const onRejection = (event) => recordDiagnostic('page-rejection', describeRejection(event));

  // DocsAPI reports a refused editor creation through console.error, not an
  // exception, so mirror the page console into the ring buffer.
  const consoleLikes = ['error', 'warn'].map((level) => {
    const original = typeof console !== 'undefined' ? console[level] : undefined;
    if (typeof original !== 'function') return null;
    const wrapped = (...args) => {
      try {
        recordDiagnostic(
          `console-${level}`,
          args
            .map((value) => {
              if (value instanceof Error) return value.stack || value.message;
              if (typeof value === 'string') return value;
              try {
                return JSON.stringify(value);
              } catch {
                return String(value);
              }
            })
            .join(' ')
            .slice(0, TEXT_LIMIT),
        );
      } catch {
        /* ignore */
      }
      return original.apply(console, args);
    };
    console[level] = wrapped;
    return () => {
      try {
        console[level] = original;
      } catch {
        /* ignore */
      }
    };
  });

  const consoleRestore = [];
  for (const restore of consoleLikes) if (typeof restore === 'function') consoleRestore.push(restore);

  try {
    // Capture phase: resource load failures do not bubble to window.
    window.addEventListener('error', onError, true);
    window.addEventListener('unhandledrejection', onRejection);
  } catch {
    installed = false;
    for (const restore of consoleRestore) restore();
    return () => {};
  }

  return () => {
    try {
      window.removeEventListener('error', onError, true);
      window.removeEventListener('unhandledrejection', onRejection);
    } catch {
      /* ignore */
    }
    for (const restore of consoleRestore) restore();
    installed = false;
  };
}

/**
 * A compact snapshot of the DOM / SDK state a failed open ended in. This is
 * what turns "engine failed" into an actionable line for the user.
 * @param {string} containerId
 * @returns {string}
 */
export function environmentReport(containerId) {
  const lines = [];
  try {
    lines.push(`page: ${location.href} (origin=${location.origin})`);
  } catch {
    lines.push('page: unavailable');
  }
  try {
    const api = typeof window !== 'undefined' ? window.DocsAPI : undefined;
    lines.push(`DocsAPI: ${api ? 'present' : 'absent'}`);
  } catch {
    lines.push('DocsAPI: unreadable');
  }
  try {
    const host = document.getElementById(containerId);
    const frames = host ? host.querySelectorAll('iframe') : [];
    lines.push(`container: ${host ? 'present' : 'MISSING'} (${frames.length} iframe(s))`);
    for (const frame of Array.from(frames).slice(0, 4)) lines.push(`  frame: ${frame.src || '(no src)'}`);
  } catch (error) {
    lines.push(`container: probe failed (${error && error.message ? error.message : String(error)})`);
  }
  try {
    const preload = document.querySelector('iframe[data-onlyoffice-preload]');
    lines.push(`preload iframe: ${preload ? preload.src : 'none'}`);
    lines.push(`document iframes: ${document.querySelectorAll('iframe').length}`);
  } catch {
    /* ignore */
  }
  return lines.join('\n');
}

/**
 * Attach listeners to the editor iframe as soon as it exists inside our
 * container (the SDK creates it asynchronously). The iframe is same-origin, so
 * its errors are readable; a cross-origin CDN mode simply fails the try/catch.
 * @param {string} containerId
 * @returns {() => void} cleanup
 */
export function attachEditorDiagnostics(containerId, options = {}) {
  if (typeof document === 'undefined') return () => {};
  let timer = null;
  let ticks = 0;
  let attachedWindow = null;
  const listeners = [];
  const keyHandler = typeof options.onKeyDown === 'function' ? options.onKeyDown : null;

  const onIframeError = (event) => recordDiagnostic('iframe-error', describeErrorEvent(event));
  const onIframeRejection = (event) => recordDiagnostic('iframe-rejection', describeRejection(event));
  const onIframeResource = (event) => {
    const text = describeResourceError(event);
    if (text) recordDiagnostic('iframe-resource', text);
  };

  const tryAttach = () => {
    if (attachedWindow) return true;
    let frame = null;
    try {
      const host = document.getElementById(containerId);
      frame = host ? host.querySelector('iframe') : null;
    } catch {
      frame = null;
    }
    if (!frame) return false;
    let win = null;
    try {
      win = frame.contentWindow;
    } catch {
      win = null;
    }
    if (!win || typeof win.addEventListener !== 'function') return false;
    try {
      win.addEventListener('error', onIframeError, true);
      win.addEventListener('unhandledrejection', onIframeRejection);
      win.addEventListener('error', onIframeResource, true);
      // The editor iframe owns keyboard focus while editing, and the parent
      // page never sees those events: the save shortcut must be captured HERE,
      // before the editor's own Ctrl/Cmd+S handler.
      if (keyHandler) win.addEventListener('keydown', keyHandler, true);
      listeners.push(() => {
        try {
          win.removeEventListener('error', onIframeError, true);
          win.removeEventListener('unhandledrejection', onIframeRejection);
          win.removeEventListener('error', onIframeResource, true);
          if (keyHandler) win.removeEventListener('keydown', keyHandler, true);
        } catch {
          /* ignore */
        }
      });
      attachedWindow = win;
      return true;
    } catch {
      return false;
    }
  };

  if (!tryAttach()) {
    timer = setInterval(() => {
      ticks += 1;
      if (tryAttach() || ticks > 480) {
        clearInterval(timer);
        timer = null;
      }
    }, 250);
    if (timer && typeof timer === 'object' && typeof timer.unref === 'function') timer.unref();
  }

  return () => {
    if (timer !== null) clearInterval(timer);
    timer = null;
    for (const remove of listeners.splice(0)) remove();
    attachedWindow = null;
  };
}
