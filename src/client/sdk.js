// SPDX-License-Identifier: AGPL-3.0-only
// dsh-onlyoffice — see LICENSE and THIRD_PARTY_NOTICES.md.

/**
 * SDK bootstrap for the dsh-onlyoffice viewers.
 *
 * Responsibilities:
 * - resolve the host route prefix for `GET /dsh-onlyoffice/status`
 * - fetch that status once per page and register the same-origin static SDK
 *   base BEFORE any editor is created (the vendored component derives
 *   `<origin><sdkBaseUrl>/onlyoffice/<version>` from it, so the version never
 *   has to be hardcoded here)
 * - install the x2t conversion worker factory from the inlined IIFE source
 *   (a single-file CJS bundle has no module URL; see vendor/PATCHES.md)
 * - serialize document conversions: the vendored x2t worker is a page-wide
 *   singleton that shares one virtual filesystem, so overlapping conversions
 *   would interleave files
 */

import { registerOnlyOfficeStaticResource } from '../../vendor/onlyoffice-web-comp/index';
import { resolvePageRelative } from './page-url.js';
import { setX2tWorkerFactory } from '../../vendor/onlyoffice-web-comp/internal/editor/x2t';
import { X2T_WORKER_SOURCE } from '../../build/generated/x2t-worker-source.js';

const STATUS_PATH = 'dsh-onlyoffice/status';
const STATUS_TIMEOUT_MS = 20000;
const DEFAULT_SDK_BASE = '/dsh-onlyoffice/sdk';
const SDK_MARKER = '/onlyoffice/';

/** A viewer failure with a stable `stage` so the UI can explain it. */
export class OfficeViewError extends Error {
  constructor(stage, message, detail) {
    super(message);
    this.name = 'OfficeViewError';
    this.stage = stage;
    this.detail = detail;
  }
}

let sdkPromise = null;
let sdkConfig = null;

/**
 * Resolve one host route.
 *
 * `ctx.get(name)` is the cordis-safe way to look at an OPTIONAL service: a
 * direct `ctx.betterSidebar` read throws `cannot get property "betterSidebar"
 * without inject` when the service was never injected. Better-sidebar's
 * `hostRouteUrl` (when present) keeps the request under a reverse-proxy
 * directory and on the desktop shell's custom protocol; otherwise resolve the
 * relative path against the page's own base.
 */
/**
 * Base URL of this plugin's Host routes, discovered once per page.
 *
 * DSH deployments disagree about where the Host mounts `/dsh-onlyoffice`:
 * a space workspace may expose the page at `/app/<spaceId>/` while the plugin
 * routes stay at the origin root, and other pages use a `<base href>` that
 * really is the deployment prefix. Guessing either one produced a hard failure
 * on the other deployment, so `ensureSdkReady()` PROBES both candidates with
 * `GET <base>/status` and remembers the first one that answers with JSON.
 *
 * @type {string | null}
 */
let hostBase = null;

/**
 * Candidate ORIGIN-level bases for the Host routes ("origin" here means
 * "everything before `/dsh-onlyoffice`"), most specific first, de-duplicated.
 * The last candidate is always the empty string, i.e. a root-relative URL.
 *
 * @returns {string[]}
 */
export function hostBaseCandidates() {
  const out = [];
  const push = (value) => {
    if (typeof value !== 'string') return;
    const clean = value.replace(/\/+$/, '');
    if (!out.includes(clean)) out.push(clean);
  };
  const origin = typeof location !== 'undefined' ? location.origin : undefined;
  const base = typeof document !== 'undefined' ? document.baseURI : undefined;

  // 1. The page directory itself (`…/app/<spaceId>/` on a space page).
  try {
    if (typeof base === 'string' && base !== '') {
      const dir = new URL('.', base);
      dir.search = '';
      dir.hash = '';
      push(dir.href);
    }
  } catch {
    /* fall through */
  }
  // 2. The origin root (plain web deployment).
  if (typeof origin === 'string' && origin !== '' && origin !== 'null' && origin !== 'undefined') {
    push(origin);
  }
  // 3. Root-relative: also valid for opaque origins like `dsh-app://app/`.
  push('');
  return out;
}

/** @param {string | null} base */
export function setHostBase(base) {
  hostBase = typeof base === 'string' && base !== '' ? base.replace(/\/+$/, '') : null;
}

/** @returns {string | null} */
export function getHostBase() {
  return hostBase;
}

export function resolveHostUrl(ctx, path) {
  try {
    const service =
      ctx && typeof ctx.get === 'function' ? ctx.get('betterSidebar') : undefined;
    if (service && typeof service.hostRouteUrl === 'function') {
      return service.hostRouteUrl(path);
    }
  } catch {
    /* fall through to the probed host base */
  }
  const clean = String(path).replace(/^\/+/, '');
  const base = hostBase !== null ? hostBase : hostBaseCandidates()[0];
  return base === '' ? `/${clean}` : `${base}/${clean}`;
}

function timeoutSignal(timeoutMs) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  return {
    signal: controller.signal,
    dispose: () => clearTimeout(timer),
  };
}

async function fetchStatus(statusPath, url) {
  const timeout = timeoutSignal(STATUS_TIMEOUT_MS);

  let response;
  try {
    response = await fetch(url, {
      method: 'GET',
      credentials: 'same-origin',
      headers: { accept: 'application/json' },
      signal: timeout.signal,
    });
  } catch (error) {
    const reason = error && error.name === 'AbortError' ? 'timeout' : error?.message || String(error);
    throw new OfficeViewError('status', `GET ${statusPath} failed`, reason);
  } finally {
    timeout.dispose();
  }

  if (!response.ok) {
    throw new OfficeViewError('status', `GET ${statusPath} returned ${response.status}`, `HTTP ${response.status}`);
  }

  let status;
  try {
    status = await response.json();
  } catch (error) {
    throw new OfficeViewError('status', 'Status response is not JSON', error?.message || String(error));
  }

  return status;
}

/**
 * Fetch `/status` from every candidate base and keep the one that answers.
 * The URL that worked becomes `hostBase` for reads, saves, diagnostics and the
 * SDK asset base, so a page served from `/app/<spaceId>/` and a page served
 * from the origin root both work without a deployment-specific guess.
 */
async function readStatus(ctx) {
  let explicit = null;
  try {
    const service =
      ctx && typeof ctx.get === 'function' ? ctx.get('betterSidebar') : undefined;
    if (service && typeof service.hostRouteUrl === 'function') {
      explicit = service.hostRouteUrl(STATUS_PATH);
    }
  } catch {
    /* fall through to probing */
  }

  const suffix = '/' + String(STATUS_PATH).replace(/^\/+/, '');
  const urls = explicit
    ? [explicit]
    : hostBaseCandidates().map((base) => (base === '' ? suffix : `${base}${suffix}`));
  const failures = [];
  for (const url of urls) {
    try {
      const status = await fetchStatus(STATUS_PATH, url);
      const base = url.slice(0, url.length - suffix.length);
      setHostBase(base);
      return { status, url };
    } catch (error) {
      failures.push({ url, error });
    }
  }

  const first = failures[0];
  if (urls.length > 1) {
    throw new OfficeViewError(
      'status',
      `host route unreachable (${urls.length} bases tried)`,
      failures.map((failure) => `${failure.url}: ${failure.error?.message || failure.error}`).join(' | '),
    );
  }
  throw first ? first.error : new OfficeViewError('status', 'no host route candidate', urls.join(', '));
}

function sdkOriginFrom(status) {
  const raw =
    status && typeof status.sdkBaseUrl === 'string' && status.sdkBaseUrl.trim()
      ? status.sdkBaseUrl.trim()
      : DEFAULT_SDK_BASE;
  const clean = raw.replace(/^\/+/, '');
  // Prefer the base that actually answered /status; both the asset URLs and the
  // API bridge must live on that same base or the iframe 404s.
  const base = hostBase !== null ? hostBase : '';
  return base === '' ? `/${clean}` : `${base}/${clean}`;
}

/**
 * Fetch the host status once and register the static SDK base.
 * Concurrent callers share the same promise; a failure clears it so a retry
 * can run again.
 */
export function ensureSdkReady(ctx) {
  if (sdkConfig) return Promise.resolve(sdkConfig);
  if (sdkPromise) return sdkPromise;

  sdkPromise = (async () => {
    const { status } = await readStatus(ctx);

    if (status && status.ok === false && !status.version) {
      throw new OfficeViewError(
        'status',
        'The host status route reported a failure',
        typeof status.lastError === 'string' && status.lastError ? status.lastError : 'ok: false',
      );
    }

    const version =
      status && typeof status.version === 'string' && status.version.trim()
        ? status.version.trim()
        : undefined;

    const cdnOrigin = sdkOriginFrom(status);

    // Must run before the first DocsAPI / x2t initialization.
    registerOnlyOfficeStaticResource({
      cdnOrigin,
      ...(version ? { onlyofficeVersion: version } : {}),
    });

    const config = {
      status,
      version,
      cdnOrigin,
      sdkRoot: version ? `${cdnOrigin}${SDK_MARKER}${version}` : cdnOrigin,
    };
    sdkConfig = config;
    return config;
  })();

  sdkPromise = sdkPromise.catch((error) => {
    sdkPromise = null;
    sdkConfig = null;
    throw error;
  });

  return sdkPromise;
}

/** Drop the cached status/SDK base (used when a viewer reports an engine bug). */
export function resetSdkCache() {
  sdkPromise = null;
  sdkConfig = null;
}

/**
 * Write edited document bytes back to the workspace file the viewer opened.
 *
 * The Host route (`POST /dsh-onlyoffice/save`) validates the path against the
 * session workspace, refuses unknown extensions and stale baselines, and
 * replaces the file atomically. Errors carry the Host's machine code so the UI
 * can distinguish a conflict (reload first) from a generic failure.
 *
 * @param {object} ctx client context (host-route resolution)
 * @param {{ sessionId: string, path: string, cwd?: string, mtimeMs?: number, bytes: ArrayBuffer }} request
 * @returns {Promise<{ bytes: number, mtimeMs: number, path?: string }>}
 */
export async function saveDocumentFile(ctx, request) {
  const { sessionId, path, cwd, mtimeMs, bytes } = request || {};
  if (!sessionId || !path) {
    throw new OfficeViewError('save', 'the document has no workspace path to save to', 'missing sessionId/path');
  }

  const url = new URL(resolveHostUrl(ctx, 'dsh-onlyoffice/save'));
  url.searchParams.set('sessionId', sessionId);
  url.searchParams.set('path', path);
  if (cwd) url.searchParams.set('cwd', cwd);
  if (Number.isFinite(mtimeMs)) url.searchParams.set('mtimeMs', String(mtimeMs));

  let response;
  try {
    response = await fetch(url.href, {
      method: 'POST',
      credentials: 'same-origin',
      headers: { 'content-type': 'application/octet-stream' },
      body: bytes,
    });
  } catch (error) {
    throw new OfficeViewError('save', 'the save request failed', error?.message || String(error));
  }

  let payload = null;
  try {
    payload = await response.json();
  } catch {
    payload = null;
  }

  if (!response.ok || !payload || payload.ok !== true) {
    const code = payload && payload.error ? payload.error.code : `http-${response.status}`;
    const message = payload && payload.error ? payload.error.message : `HTTP ${response.status}`;
    throw new OfficeViewError('save', message, `${code}`);
  }

  return payload;
}

/** Decode the base64url file version the read route reports. */
function decodeVersionHeader(raw) {
  if (typeof raw !== 'string' || raw === '') return undefined;
  try {
    const padded = raw.replace(/-/g, '+').replace(/_/g, '/');
    const binary = atob(padded + '='.repeat((4 - (padded.length % 4)) % 4));
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
    return new TextDecoder().decode(bytes);
  } catch {
    return undefined;
  }
}

/**
 * Read one document's bytes for the built-in preview's `renderer` loading mode.
 *
 * The renderer owns loading, so it asks this plugin's Host route (which does
 * the same workspace validation as save-back) instead of relying on the
 * preview's internal content plumbing. The route reports the file's version, so
 * the renderer can mark the document rendered and prevent the preview from
 * treating our own save as an external change.
 *
 * @param {object} ctx client context (host-route resolution)
 * @param {{ sessionId: string, path: string, cwd?: string, signal?: AbortSignal }} request
 * @returns {Promise<{ bytes: ArrayBuffer, version?: string }>}
 */
export async function readDocumentFile(ctx, request) {
  const { sessionId, path, cwd, signal } = request || {};
  if (!sessionId || !path) {
    throw new OfficeViewError('read', 'the document has no workspace path to read', 'missing sessionId/path');
  }

  const url = new URL(resolveHostUrl(ctx, 'dsh-onlyoffice/read'));
  url.searchParams.set('sessionId', sessionId);
  url.searchParams.set('path', path);
  if (cwd) url.searchParams.set('cwd', cwd);

  let response;
  try {
    response = await fetch(url.href, { credentials: 'same-origin', signal });
  } catch (error) {
    if (error && error.name === 'AbortError') throw error;
    throw new OfficeViewError('read', 'the document read failed', error?.message || String(error));
  }

  if (!response.ok) {
    let detail = `HTTP ${response.status}`;
    try {
      const payload = await response.json();
      if (payload && payload.error) detail = `${payload.error.code}: ${payload.error.message}`;
    } catch {
      /* keep the status */
    }
    throw new OfficeViewError('read', 'the document could not be read', detail);
  }

  const bytes = await response.arrayBuffer();
  if (!bytes || bytes.byteLength === 0) {
    throw new OfficeViewError('read', 'the document is empty', 'empty body');
  }

  return {
    bytes,
    version: decodeVersionHeader(response.headers.get('x-dsh-file-version')),
  };
}

let workerFactoryInstalled = false;

/**
 * Build the x2t worker from the inlined IIFE bundle: a classic worker from a
 * blob URL. The blob URL is revoked once the worker answers (or fails), with a
 * timer as a backstop.
 */
function buildX2tWorker() {
  if (typeof Worker === 'undefined' || typeof Blob === 'undefined' || typeof URL === 'undefined') {
    throw new OfficeViewError('engine', 'Web Workers are not available', 'Worker/Blob/URL missing');
  }

  const blob = new Blob([X2T_WORKER_SOURCE], { type: 'text/javascript' });
  const blobUrl = URL.createObjectURL(blob);
  const worker = new Worker(blobUrl);

  let revoked = false;
  const revoke = () => {
    if (revoked) return;
    revoked = true;
    clearTimeout(timer);
    try {
      URL.revokeObjectURL(blobUrl);
    } catch {
      /* ignore */
    }
  };
  const timer = setTimeout(revoke, 30000);
  worker.addEventListener('message', revoke, { once: true });
  worker.addEventListener('error', revoke, { once: true });

  return worker;
}

/** Register the worker factory with the vendored converter. Returns a disposer. */
export function installX2tWorkerFactory() {
  if (workerFactoryInstalled) return () => {};
  workerFactoryInstalled = true;
  setX2tWorkerFactory(buildX2tWorker);

  return () => {
    workerFactoryInstalled = false;
    setX2tWorkerFactory(null);
  };
}

let conversionChain = Promise.resolve();

/**
 * How long a queued conversion waits for the queue head before it stops
 * honoring the serialization guarantee. A head that never settles must not
 * deadlock every later open (and the viewer's retry button).
 */
const CONVERSION_WAIT_TIMEOUT_MS = 240000;

/**
 * Serialize x2t conversions (document opens and exports). The vendored worker
 * keeps a single Emscripten filesystem at `/working`, so two overlapping
 * conversions can read each other's files.
 *
 * The wait is bounded: after `waitTimeoutMs` the task runs anyway (with a
 * warning), which is strictly better than leaving the sidebar permanently
 * stuck behind one hung conversion.
 */
export function withConversionLock(task, waitTimeoutMs = CONVERSION_WAIT_TIMEOUT_MS) {
  let timer = null;
  const head = conversionChain.catch(() => undefined);
  const ready = head.then(() => 'ready');
  const stalled = new Promise((resolve) => {
    timer = setTimeout(() => resolve('stalled'), waitTimeoutMs);
  });

  const run = Promise.race([ready, stalled]).then((state) => {
    if (state === 'stalled') {
      console.warn(
        '[dsh-onlyoffice] the document conversion queue did not drain in time; running this conversion without the serialization guarantee',
      );
    }
    return task();
  });

  run.then(
    () => clearTimeout(timer),
    () => clearTimeout(timer),
  );

  conversionChain = run.then(
    () => undefined,
    () => undefined,
  );
  return run;
}

/** Short, user-facing detail for an unknown error. */
export function errorText(error) {
  if (!error) return 'Unknown error';
  if (typeof error === 'string') return error;
  const message = typeof error.message === 'string' && error.message ? error.message : String(error);
  return message;
}
