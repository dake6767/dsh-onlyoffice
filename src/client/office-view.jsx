// SPDX-License-Identifier: AGPL-3.0-only
// dsh-onlyoffice — see LICENSE and THIRD_PARTY_NOTICES.md.

/**
 * The OnlyOffice viewer rendered inside the dsh-better-sidebar file panel.
 *
 * Bytes come from the host-provided `mediaUrl` (same-origin raw file route), so
 * the viewer does: fetch(mediaUrl) → ArrayBuffer → File → the vendored
 * `OnlyOfficeManager.createWithFile(...)`. The manager owns the editor iframe,
 * the mock collaboration server and the x2t conversion bridge; it is destroyed
 * whenever this component unmounts or reloads.
 *
 * Failure policy: every stage (host status route, file bytes, engine boot,
 * conversion, editor render) maps to a visible error with a retry button. A
 * missing/failed better-sidebar status route can never throw out of the
 * component and blank the sidebar.
 */

import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  FILE_TYPE,
  OFFICE_THEME,
  ONLYOFFICE_EVENT_KEYS,
  OnlyOfficeManager,
  editorManagerFactory,
  onlyofficeEventbus,
} from '../../vendor/onlyoffice-web-comp/index';
import {
  OfficeViewError,
  ensureSdkReady,
  errorText,
  saveDocumentFile,
  withConversionLock,
  isAutoSaveEnabled,
} from './sdk.js';
import { isHostDark, subscribeHostTheme } from './theme.js';
import { t } from './i18n.js';
import {
  attachEditorDiagnostics,
  diagnosticsDetail,
  environmentReport,
  recordDiagnostic,
} from './diagnostics.js';
import { isSaveShortcut, saveShortcutLabel } from './shortcut.js';
import { sameBytes, toUint8Array } from './bytes.js';
import { AUTOSAVE_DELAY_MS, createAutoSaver } from './autosave.js';

/** Long enough for a cold first open (SDK download + x2t conversion). */
const DOCUMENT_READY_TIMEOUT_MS = 180000;
/** Hard stop for the whole open pipeline (a hung host route must surface an error + retry). */
const LOAD_WATCHDOG_MS = 240000;
const SLOW_HINT_DELAY_MS = 12000;
const NOTICE_TIMEOUT_MS = 8000;

const KIND_TO_FILE_TYPE = {
  docx: FILE_TYPE.DOCX,
  xlsx: FILE_TYPE.XLSX,
  pptx: FILE_TYPE.PPTX,
};

const KIND_TO_EXT = {
  docx: 'docx',
  xlsx: 'xlsx',
  pptx: 'pptx',
};

const MIME_BY_EXT = {
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  doc: 'application/msword',
  odt: 'application/vnd.oasis.opendocument.text',
  rtf: 'application/rtf',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  xls: 'application/vnd.ms-excel',
  ods: 'application/vnd.oasis.opendocument.spreadsheet',
  csv: 'text/csv',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  ppt: 'application/vnd.ms-powerpoint',
  odp: 'application/vnd.oasis.opendocument.presentation',
};

let containerSeq = 0;

/**
 * Bytes written back per document target, keyed `sessionId\0path`.
 *
 * A file watcher re-read arrives as new `bytes` props — possibly after a
 * remount, which would clear component-local state. Keeping the last written
 * bytes here lets both cases recognise "this is my own save" and keep the live
 * editor instead of rebuilding it.
 */
const savedBytesByTarget = new Map();
const SAVED_BYTES_LIMIT = 8;

function savedBytesKey(sessionId, path) {
  return `${String(sessionId || '')}\u0000${String(path || '')}`;
}

function rememberSavedBytes(sessionId, path, bytes) {
  if (!sessionId || !path) return;
  const key = savedBytesKey(sessionId, path);
  savedBytesByTarget.delete(key);
  savedBytesByTarget.set(key, bytes);
  while (savedBytesByTarget.size > SAVED_BYTES_LIMIT) {
    const oldest = savedBytesByTarget.keys().next();
    if (oldest.done === true) break;
    savedBytesByTarget.delete(oldest.value);
  }
}

function knownSavedBytes(sessionId, path) {
  if (!sessionId || !path) return null;
  return savedBytesByTarget.get(savedBytesKey(sessionId, path)) ?? null;
}

/**
 * Right after a successful save the preview watcher re-reads the file, which
 * briefly surfaces as "waiting for bytes" again. Inside this window the
 * viewer must keep its live editor instead of rebuilding during the gap.
 */
const WATCHER_RELOAD_WINDOW_MS = 10_000;
const reloadExpectedUntil = new Map();

function markWatcherReloadExpected(sessionId, path) {
  if (!sessionId || !path) return;
  reloadExpectedUntil.set(savedBytesKey(sessionId, path), Date.now() + WATCHER_RELOAD_WINDOW_MS);
  while (reloadExpectedUntil.size > SAVED_BYTES_LIMIT) {
    const oldest = reloadExpectedUntil.keys().next();
    if (oldest.done === true) break;
    reloadExpectedUntil.delete(oldest.value);
  }
}

function isWatcherReloadExpected(sessionId, path) {
  if (!sessionId || !path) return false;
  const until = reloadExpectedUntil.get(savedBytesKey(sessionId, path));
  if (until === undefined) return false;
  if (Date.now() > until) {
    reloadExpectedUntil.delete(savedBytesKey(sessionId, path));
    return false;
  }
  return true;
}

function clearWatcherReloadExpected(sessionId, path) {
  if (!sessionId || !path) return;
  reloadExpectedUntil.delete(savedBytesKey(sessionId, path));
}

function nextContainerId() {
  containerSeq += 1;
  const random = Math.random().toString(36).slice(2, 8);
  return `dsh-onlyoffice-editor-${containerSeq}-${random}`;
}

function pathBaseName(path) {
  const text = String(path || '');
  const parts = text.split(/[\\/]/);
  return parts[parts.length - 1] || '';
}

function fileExtension(name) {
  const match = /\.([A-Za-z0-9]{1,8})$/.exec(String(name || ''));
  return match ? match[1].toLowerCase() : '';
}

/** Keep the real extension (the host file type may be .doc/.csv/.odp), else append the viewer's default. */
function resolveFileName(title, path, kind) {
  const fallbackExt = KIND_TO_EXT[kind] || 'docx';
  const fromTitle = typeof title === 'string' ? title.trim() : '';
  const fromPath = pathBaseName(path);
  const raw = fromTitle || fromPath || `document.${fallbackExt}`;
  return fileExtension(raw) ? raw : `${raw}.${fallbackExt}`;
}

function mimeFor(fileName) {
  return MIME_BY_EXT[fileExtension(fileName)] || 'application/octet-stream';
}

function disposeManager(manager, containerId) {
  try {
    manager?.setSaveShortcutHandler?.(null);
    manager?.setChangeHandler?.(null);
  } catch {
    /* the editor is going away anyway */
  }
  try {
    manager?.destroy?.();
  } catch (error) {
    console.warn('[dsh-onlyoffice] failed to destroy the Office editor', error);
  }
  try {
    editorManagerFactory.destroy(containerId);
  } catch (error) {
    console.warn('[dsh-onlyoffice] failed to release the editor manager', error);
  }
}

/** Wrap bytes the document owner already read (DSH document preview) as a File. */
function fileFromBytes(bytes, fileName) {
  if (typeof File !== 'function') {
    throw new OfficeViewError('fetch', t('errorUnsupported'));
  }
  // A File copies its parts, so the borrowed Host byte view can be released.
  return new File([bytes], fileName, { type: mimeFor(fileName) });
}

async function fetchOfficeFile(mediaUrl, fileName, signal) {
  if (!mediaUrl) {
    throw new OfficeViewError('fetch', t('errorNoMedia'));
  }
  if (typeof File !== 'function') {
    throw new OfficeViewError('fetch', t('errorUnsupported'));
  }

  let response;
  try {
    response = await fetch(mediaUrl, { credentials: 'same-origin', signal });
  } catch (error) {
    if (error && error.name === 'AbortError') throw error;
    throw new OfficeViewError('fetch', t('errorFetch'), errorText(error));
  }

  if (!response.ok) {
    throw new OfficeViewError('fetch', t('errorFetch'), `HTTP ${response.status} ${response.statusText || ''}`.trim());
  }

  const buffer = await response.arrayBuffer();
  if (!buffer || buffer.byteLength === 0) {
    throw new OfficeViewError('fetch', t('errorFetch'), 'empty response body');
  }

  return new File([buffer], fileName, { type: mimeFor(fileName) });
}

function waitForDocumentReady(containerId, timeoutMs, signal) {
  return new Promise((resolve, reject) => {
    if (signal && signal.aborted) {
      reject(new OfficeViewError('aborted', 'aborted'));
      return;
    }

    let settled = false;
    let timer = null;

    const cleanup = () => {
      onlyofficeEventbus.off(ONLYOFFICE_EVENT_KEYS.DOCUMENT_READY, handler);
      if (timer !== null) clearTimeout(timer);
      try {
        signal?.removeEventListener?.('abort', onAbort);
      } catch {
        /* ignore */
      }
    };

    const finish = (settle, value) => {
      if (settled) return;
      settled = true;
      cleanup();
      settle(value);
    };

    const handler = (data) => {
      if (data && data.instanceId && data.instanceId !== containerId) return;
      finish(resolve, data);
    };

    const onAbort = () => finish(reject, new OfficeViewError('aborted', 'aborted'));

    onlyofficeEventbus.on(ONLYOFFICE_EVENT_KEYS.DOCUMENT_READY, handler);
    timer = setTimeout(
      () => finish(reject, new OfficeViewError('ready', t('errorReady'))),
      timeoutMs,
    );
    try {
      signal?.addEventListener?.('abort', onAbort, { once: true });
    } catch {
      /* ignore */
    }
  });
}

/** What the editor iframe looks like right now — the useful half of a timeout. */
function editorFrameInfo(containerId) {
  return environmentReport(containerId);
}

function toViewerError(error) {
  if (error instanceof OfficeViewError) {
    return { stage: error.stage, message: error.message, detail: error.detail };
  }
  return { stage: 'engine', message: t('errorEngine'), detail: errorText(error) };
}

function stageHint(stage) {
  if (stage === 'status') return t('errorStatusHint');
  if (stage === 'timeout') return t('errorTimeoutHint');
  return t('errorRetryHint');
}

function stageMessage(stage, fallback) {
  if (stage === 'status') return t('errorStatus');
  return fallback;
}

/** True when the editor iframe exists inside our container (the readiness fallback). */
/**
 * Whether the OnlyOffice editor iframe exists for this container.
 *
 * DocsAPI may REPLACE the `#containerId` element while mounting: the editor
 * iframe then lives as a SIBLING under our `.onlyoffice-container` wrapper and
 * `getElementById(containerId)` returns null. Checking only inside the id
 * element (the original implementation) therefore missed a perfectly mounted
 * editor and turned a slow-but-alive cold start into a "render timed out"
 * error on remote deployments.
 */
function hasEditorFrame(containerId) {
  try {
    const host = document.getElementById(containerId);
    if (host && host.querySelector('iframe')) return true;

    const wrapper = document.querySelector(`[data-onlyoffice-container-id="${containerId}"]`);
    if (wrapper && wrapper.querySelector('iframe')) return true;

    // Last resort: a frameEditor that carries our container id in its URL.
    return Array.from(document.querySelectorAll('iframe[name="frameEditor"]')).some((frame) =>
      String(frame.src || '').includes(containerId),
    );
  } catch {
    return false;
  }
}

/**
 * @param props FileViewerProps (ctx / scope / path / title / mediaUrl / …) plus `kind`.
 */
export function OfficeView(props) {
  const { path, title, mediaUrl, kind, bytes, waitingForBytes } = props;
  const ctxRef = useRef(props.ctx);
  ctxRef.current = props.ctx;

  const [containerId] = useState(() => nextContainerId());
  const [phase, setPhase] = useState('loading');
  const [error, setError] = useState(null);
  const [readOnly, setReadOnly] = useState(false);
  const [busy, setBusy] = useState('');
  const [notice, setNotice] = useState('');
  const [reloadToken, setReloadToken] = useState(0);
  const [slowHint, setSlowHint] = useState(false);
  const [dark, setDark] = useState(() => isHostDark());
  /** { version, bytes } of an external disk change awaiting the user's reload. */
  const [externalUpdate, setExternalUpdate] = useState(null);
  const externalUpdateRef = useRef(null);
  externalUpdateRef.current = externalUpdate;

  const managerRef = useRef(null);
  const mountedRef = useRef(true);
  const readOnlyRef = useRef(readOnly);
  const themeRef = useRef(dark);
  const phaseRef = useRef(phase);
  const noticeTimerRef = useRef(null);

  readOnlyRef.current = readOnly;
  themeRef.current = dark;
  phaseRef.current = phase;

  const fileName = useMemo(
    () => resolveFileName(title, path, kind),
    [title, path, kind],
  );
  const fileNameRef = useRef(fileName);
  fileNameRef.current = fileName;

  // The workspace coordinates this document came from. The built-in preview
  // supplies `sessionId` + parsed `sourcePath`; better-sidebar supplies
  // `scope`/`path`/`mtimeMs`. Both are optional: without a session+path there
  // is nothing to write back to, so the save control stays disabled.
  const ownerScope = props.scope || {};
  const saveSessionId = props.sessionId || ownerScope.sessionId || '';
  const savePath = props.sourcePath || props.path || '';
  const saveCwd = props.cwd || ownerScope.cwd || '';
  // The engine exports native bytes only, so a `.doc`/`.odt`/`.csv`/`.ppt`
  // source must not be overwritten with them (see host/save.js).
  const nativeSaveable = /^(docx|xlsx|pptx)$/i.test((savePath.match(/\.([A-Za-z0-9]+)$/) || [])[1] || '');
  const hasSaveTarget = Boolean(saveSessionId) && Boolean(savePath);
  const canSaveTarget = hasSaveTarget && nativeSaveable;
  const initialMtime = Number.isFinite(props.mtimeMs) ? props.mtimeMs : undefined;
  const savedMtimeRef = useRef(initialMtime);
  /** Latest save action, so the (stable) key handlers always call it. */
  const saveRef = useRef(() => {});
  /** One write-back at a time: the editor's save button and the shortcut can
   *  both arrive for the same edit. */
  const savingRef = useRef(false);
  /** Bytes we last wrote back: a watcher-driven re-read of exactly these is
   *  our own save round-tripping, not a new document. */
  const lastSavedBytesRef = useRef(null);
  /** Latest "save automatically if allowed" action (refreshed every render). */
  const autoSaveRef = useRef(() => {});
  /** Debounced auto-save driven by the editor's document-change events. */
  const autoSaverRef = useRef(null);
  if (autoSaverRef.current === null) {
    autoSaverRef.current = createAutoSaver({
      delayMs: AUTOSAVE_DELAY_MS,
      run: () => autoSaveRef.current(),
    });
  }
  /** Ignore change events caused by our own export right after a save. */
  const ignoreChangesUntilRef = useRef(0);

  /** The file version this editor currently shows (renderer loading mode). */
  const lastRenderedVersionRef = useRef(undefined);
  /** Pending "is this really our own write?" re-check (renderer mode). */
  const deferredExternalCheckRef = useRef(null);
  /** Set when an in-place refresh failed and a full rebuild is required. */
  const forceRebuildRef = useRef(false);
  /** Renderer-loading-mode hooks from the built-in preview (revision + reports). */
  const rendererRef = useRef(props.renderer);
  rendererRef.current = props.renderer;
  const rendererRevision = props.renderer ? props.renderer.revision : undefined;

  /** Capture Cmd/Ctrl+S from the editor iframe (it owns keyboard focus). */
  const interceptSaveKey = useCallback((event) => {
    if (!isSaveShortcut(event)) return;
    // Always swallow the browser's Save-Page dialog and the editor's own
    // mock-server save, then run the real write-back.
    event.preventDefault();
    event.stopImmediatePropagation?.();
    event.stopPropagation();
    saveRef.current();
  }, []);

  const clearNoticeTimer = useCallback(() => {
    if (noticeTimerRef.current !== null) {
      clearTimeout(noticeTimerRef.current);
      noticeTimerRef.current = null;
    }
  }, []);

  const showNotice = useCallback(
    (text) => {
      clearNoticeTimer();
      setNotice(text);
      noticeTimerRef.current = setTimeout(() => {
        noticeTimerRef.current = null;
        if (mountedRef.current) setNotice('');
      }, NOTICE_TIMEOUT_MS);
    },
    [clearNoticeTimer],
  );

  useEffect(() => {
    mountedRef.current = true;
    return () => {
      mountedRef.current = false;
      clearNoticeTimer();
      autoSaverRef.current?.dispose();
      if (deferredExternalCheckRef.current !== null) clearTimeout(deferredExternalCheckRef.current);
      deferredExternalCheckRef.current = null;
    };
  }, [clearNoticeTimer]);

  // Host light/dark → OnlyOffice uiTheme.
  useEffect(() => subscribeHostTheme((nextDark) => setDark(nextDark)), []);

  // The editor's own Save affordances (Quick Access floppy / File ▸ Save)
  // reach us as ONSAVE. The vendor hooks DocsAPI's onSave into it, so one
  // handler covers both the editor button and any future menu entry.
  useEffect(() => {
    const onSaved = (data) => {
      if (data && data.instanceId && data.instanceId !== containerId) return;
      saveRef.current();
    };
    onlyofficeEventbus.on(ONLYOFFICE_EVENT_KEYS.ONSAVE, onSaved);
    return () => {
      onlyofficeEventbus.off(ONLYOFFICE_EVENT_KEYS.ONSAVE, onSaved);
    };
  }, [containerId]);

  // Load (and reload) the document.
  useEffect(() => {
    // Our own write-back makes the preview watcher re-read the file. When the
    // incoming bytes are byte-identical to what we just saved, keep the live
    // editor: rebuilding it here is the visible "reload after saving" flicker.
    const rendererMode = rendererRevision !== undefined;
    const incomingBytes = rendererMode ? null : toUint8Array(bytes);
    const knownSaved =
      lastSavedBytesRef.current ?? knownSavedBytes(saveSessionId, savePath);
    const sameAsSaved =
      incomingBytes !== null && knownSaved !== null && sameBytes(incomingBytes, knownSaved);
    const ownWatcherReload =
      managerRef.current !== null &&
      !rendererMode &&
      waitingForBytes === true &&
      isWatcherReloadExpected(saveSessionId, savePath);

    // Renderer mode: the preview bumps the revision when its watcher sees the
    // file change. Right after our own save that is our write coming back —
    // keep the live editor and re-report the version we already render.
    const ownRendererReload =
      rendererMode &&
      managerRef.current !== null &&
      lastRenderedVersionRef.current !== undefined &&
      isWatcherReloadExpected(saveSessionId, savePath);

    if (ownRendererReload) {
      try {
        rendererRef.current?.loaded?.(lastRenderedVersionRef.current);
      } catch {
        /* ignore */
      }
      // The write could also be an EXTERNAL one that landed inside our save
      // window. Re-check the real version shortly after: if it differs from the
      // one this editor shows, clear the window and force a reload.
      if (deferredExternalCheckRef.current === null) {
        deferredExternalCheckRef.current = setTimeout(async () => {
          deferredExternalCheckRef.current = null;
          try {
            const loaded = await rendererRef.current?.load?.();
            if (loaded && loaded.version && loaded.version !== lastRenderedVersionRef.current) {
              clearWatcherReloadExpected(saveSessionId, savePath);
              rendererRef.current?.reload?.();
            }
          } catch {
            /* a failed probe simply leaves the editor as is */
          }
        }, 1500);
      }
      const detach = attachEditorDiagnostics(containerId, { onKeyDown: interceptSaveKey });
      return () => detach();
    }

    if (managerRef.current !== null && (sameAsSaved || ownWatcherReload)) {
      // Our save round-tripping through the file watcher: keep the editor.
      if (sameAsSaved) {
        lastSavedBytesRef.current = incomingBytes;
        clearWatcherReloadExpected(saveSessionId, savePath);
      }
      const detach = attachEditorDiagnostics(containerId, { onKeyDown: interceptSaveKey });
      return () => detach();
    }
    clearWatcherReloadExpected(saveSessionId, savePath);

    let cancelled = false;
    const controller = new AbortController();

    // ── External change while a live editor is mounted ──────────────────
    // Refresh the DOCUMENT in place (`refreshFile`) instead of destroying and
    // rebuilding the editor: the OnlyOffice chrome and layout stay put and only
    // the content reloads. Falls back to a full rebuild on failure.
    if (
      rendererMode &&
      !forceRebuildRef.current &&
      managerRef.current !== null &&
      lastRenderedVersionRef.current !== undefined &&
      !waitingForBytes
    ) {
      const inPlaceController = new AbortController();
      let inPlaceCancelled = false;
      const detachInPlace = attachEditorDiagnostics(containerId, {
        onKeyDown: interceptSaveKey,
      });

      const refreshInPlace = async () => {
        try {
          const loaded = await rendererRef.current.load(inPlaceController.signal);
          if (inPlaceCancelled) return;
          if (loaded.version !== undefined && loaded.version === lastRenderedVersionRef.current) {
            // Our own save round-trip: nothing to show.
            rendererRef.current?.loaded?.(loaded.version);
            return;
          }
          // External change: never swap bytes under the user. Stop the preview
          // from retrying (report the version) and offer an explicit reload.
          if (loaded.version !== undefined) {
            try {
              rendererRef.current?.loaded?.(loaded.version);
            } catch {
              /* ignore */
            }
          }
          clearWatcherReloadExpected(saveSessionId, savePath);
          if (!inPlaceCancelled) {
            setExternalUpdate({ version: loaded.version, bytes: loaded.bytes });
          }
        } catch {
          /* a failed probe leaves the editor as it was */
        }
      };

      void refreshInPlace();
      return () => {
        inPlaceCancelled = true;
        inPlaceController.abort();
        detachInPlace();
      };
    }

    forceRebuildRef.current = false;
    setExternalUpdate(null);
    autoSaverRef.current?.cancel();
    disposeManager(managerRef.current, containerId);
    managerRef.current = null;
    savedMtimeRef.current = initialMtime;

    setPhase('loading');
    setError(null);
    setNotice('');
    setBusy('');
    setSlowHint(false);

    const slowTimer = setTimeout(() => {
      if (!cancelled && mountedRef.current) setSlowHint(true);
    }, SLOW_HINT_DELAY_MS);

    // If the host SDK route hangs (a fetch that never settles), the pipeline
    // would spin forever. Surface an error with retry instead; a late success
    // still flips the view back to ready.
    const watchdogTimer = setTimeout(() => {
      if (cancelled || !mountedRef.current) return;
      const current = phaseRef.current;
      if (current === 'ready' || current === 'error') return;
      recordDiagnostic('viewer-timeout', editorFrameInfo(containerId));
      setError({
        stage: 'timeout',
        message: t('errorTimeout'),
        detail: diagnosticsDetail(`${t('errorTimeoutHint')}\n${editorFrameInfo(containerId)}`),
      });
      setPhase('error');
    }, LOAD_WATCHDOG_MS);

    // The SDK's iframe errors never reach our catch: listen inside it too.
    const detachEditorDiagnostics = attachEditorDiagnostics(containerId, {
      onKeyDown: interceptSaveKey,
    });

    const run = async () => {
      try {
        if (waitingForBytes) {
          // The document owner (DSH document preview) is still reading the
          // file. Keep the loading state; this effect re-runs when bytes land.
          setPhase('loading');
          return;
        }
        await ensureSdkReady(ctxRef.current);
        if (cancelled) return;

        setPhase('opening');
        let file;
        let loadedVersion;
        if (rendererMode) {
          const loaded = await rendererRef.current.load(controller.signal);
          if (cancelled) return;
          file = fileFromBytes(loaded.bytes, fileNameRef.current);
          loadedVersion = loaded.version;
        } else if (bytes) {
          file = fileFromBytes(bytes, fileNameRef.current);
        } else {
          file = await fetchOfficeFile(mediaUrl, fileNameRef.current, controller.signal);
        }
        if (cancelled) return;

        setPhase('converting');
        const manager = await withConversionLock(async () => {
          const created = await OnlyOfficeManager.createWithFile(
            {
              containerId,
              fileType: KIND_TO_FILE_TYPE[kind] || FILE_TYPE.DOCX,
              defaultFileName: fileNameRef.current,
              readOnly: readOnlyRef.current,
              theme: themeRef.current ? OFFICE_THEME.DARK : OFFICE_THEME.WHITE,
            },
            file,
          );
          if (cancelled) {
            disposeManager(created, containerId);
            return null;
          }
          return created;
        });
        if (!manager || cancelled) return;

        managerRef.current = manager;
        // The editor's Ctrl/Cmd+S blocker now delegates to this view's save.
        try {
          manager.setSaveShortcutHandler?.(() => saveRef.current());
          // DSH patch: the debounced write-back is opt-in (host `autoSave`).
          // Manual saves (Cmd/Ctrl+S, the editor save button) always work.
          manager.setChangeHandler?.(() => {
            if (!isAutoSaveEnabled()) return;
            if (ignoreChangesUntilRef.current > Date.now()) return;
            autoSaverRef.current?.schedule();
          });
        } catch (error) {
          console.warn('[dsh-onlyoffice] could not hook the editor save affordances', error);
        }
        setReadOnly(Boolean(manager.getReadOnly?.()));

        try {
          await waitForDocumentReady(containerId, DOCUMENT_READY_TIMEOUT_MS, controller.signal);
        } catch (readyError) {
          if (cancelled || controller.signal.aborted) return;
          // A slow-but-alive editor: the readiness event was missed, but the
          // frame is there and the document is rendering.
          if (hasEditorFrame(containerId)) {
            showNotice(t('errorReady'));
          } else {
            throw readyError;
          }
        }
        if (cancelled) return;

        setPhase('ready');
        setSlowHint(false);
        // Tell the preview which file version this editor shows: our own save
        // reports the post-write version here, so a save never looks like an
        // external change (which would reload the editor).
        if (rendererMode && loadedVersion !== undefined) {
          try {
            rendererRef.current?.loaded?.(loadedVersion);
          } catch (error) {
            console.warn('[dsh-onlyoffice] could not report the rendered version', error);
          }
          lastRenderedVersionRef.current = loadedVersion;
        }
      } catch (loadError) {
        if (cancelled || controller.signal.aborted) return;
        // Release anything a partially-failed open left behind (the manager
        // factory keeps per-container entries even when `create` throws).
        disposeManager(managerRef.current, containerId);
        managerRef.current = null;
        const viewerError = toViewerError(loadError);
        recordDiagnostic(
          'viewer-error',
          `${viewerError.stage}: ${viewerError.message}${viewerError.detail ? ` — ${viewerError.detail}` : ''} (${editorFrameInfo(containerId)})`,
        );
        setError({ ...viewerError, detail: diagnosticsDetail(viewerError.detail) });
        setPhase('error');
      } finally {
        clearTimeout(slowTimer);
        clearTimeout(watchdogTimer);
      }
    };

    run();

    return () => {
      cancelled = true;
      controller.abort();
      clearTimeout(slowTimer);
      clearTimeout(watchdogTimer);
      detachEditorDiagnostics();
    };
  }, [containerId, kind, mediaUrl, bytes, waitingForBytes, reloadToken, showNotice, initialMtime, interceptSaveKey, rendererRevision]);

  // Unmount: never leave an editor iframe (or its manager) behind.
  useEffect(() => {
    return () => {
      disposeManager(managerRef.current, containerId);
      managerRef.current = null;
    };
  }, [containerId]);

  // Apply a theme switch to an already-mounted editor.
  useEffect(() => {
    const manager = managerRef.current;
    if (!manager || typeof manager.setTheme !== 'function') return;
    const next = dark ? OFFICE_THEME.DARK : OFFICE_THEME.WHITE;
    let current;
    try {
      current = manager.getTheme?.();
    } catch {
      current = undefined;
    }
    if (current === next) return;
    try {
      const pending = manager.setTheme(next);
      pending?.catch?.((themeError) => {
        console.warn('[dsh-onlyoffice] failed to sync the editor theme', themeError);
      });
    } catch (themeError) {
      console.warn('[dsh-onlyoffice] failed to sync the editor theme', themeError);
    }
  }, [dark]);

  const toggleReadOnly = useCallback(async () => {
    const manager = managerRef.current;
    if (!manager) return;
    setBusy('readonly');
    try {
      await manager.setReadOnly(!manager.getReadOnly());
      if (mountedRef.current) setReadOnly(Boolean(manager.getReadOnly()));
    } catch (toggleError) {
      if (mountedRef.current) showNotice(`${t('readonlyFailed')}: ${errorText(toggleError)}`);
    } finally {
      if (mountedRef.current) setBusy('');
    }
  }, [showNotice]);

  const exportDocument = useCallback(async () => {
    const manager = managerRef.current;
    if (!manager) return;
    setBusy('export');
    try {
      await withConversionLock(() => manager.downloadExport());
    } catch (exportError) {
      if (mountedRef.current) showNotice(`${t('exportFailed')}: ${errorText(exportError)}`);
    } finally {
      if (mountedRef.current) setBusy('');
    }
  }, [showNotice]);

  /**
   * Export the edited document through the engine (Editor.bin → x2t → Office
   * bytes) and write it back over the original workspace file. The Host does
   * the authorization and the atomic replace; this only calls it.
   */
  const saveToFile = useCallback(async (silent = false) => {
    const manager = managerRef.current;
    if (!manager || !canSaveTarget || savingRef.current) return;
    savingRef.current = true;
    setBusy('save');
    try {
      const exported = await withConversionLock(() => manager.exportAsBlob());
      if (!exported || !exported.blob) {
        throw new OfficeViewError('save', t('saveFailed'), 'the engine produced no document bytes');
      }
      const bytes = await exported.blob.arrayBuffer();
      const result = await saveDocumentFile(ctxRef.current, {
        sessionId: saveSessionId,
        path: savePath,
        cwd: saveCwd || undefined,
        mtimeMs: savedMtimeRef.current,
        bytes,
      });
      // Remember exactly what the file now contains: the watcher's re-read is
      // recognised as our own round-trip and does not rebuild the editor.
      const savedSnapshot = new Uint8Array(bytes.slice(0));
      lastSavedBytesRef.current = savedSnapshot;
      rememberSavedBytes(saveSessionId, savePath, savedSnapshot);
      markWatcherReloadExpected(saveSessionId, savePath);
      // Renderer mode: adopt the post-write version as rendered. The preview
      // then sees no external change and keeps this editor mounted.
      if (result.version !== undefined && result.version !== null && result.version !== '') {
        const version = String(result.version);
        lastRenderedVersionRef.current = version;
        try {
          rendererRef.current?.loaded?.(version);
        } catch (error) {
          console.warn('[dsh-onlyoffice] could not report the saved version', error);
        }
      }
      if (Number.isFinite(result.mtimeMs)) savedMtimeRef.current = result.mtimeMs;
      // Our own export can emit a change event; do not let it re-schedule.
      ignoreChangesUntilRef.current = Date.now() + 800;
      if (!silent && mountedRef.current) showNotice(t('savedNotice'));
    } catch (saveError) {
      const code = saveError && saveError.detail ? String(saveError.detail) : '';
      const isConflict = code.includes('conflict');
      if (mountedRef.current) {
        showNotice(`${isConflict ? t('saveConflict') : t('saveFailed')}: ${errorText(saveError)}`);
      }
    } finally {
      savingRef.current = false;
      if (mountedRef.current) setBusy('');
    }
  }, [canSaveTarget, saveSessionId, savePath, saveCwd, showNotice]);

  /**
   * Load the external update the banner announced.
   *
   * This deliberately rebuilds the document the same way reopening it does:
   * the in-place `refreshFile` path was unreliable in the static SDK (the x2t
   * refresh conversion could stall or produce no output). The reload is
   * user-initiated, so a normal document open is the honest behaviour.
   */
  const applyExternalUpdate = useCallback(() => {
    const pending = externalUpdateRef.current;
    if (!pending) return;
    autoSaverRef.current?.cancel();
    // Any watcher expectation belongs to the old editor; the rebuild is a
    // deliberate user action, not our own save round-trip.
    clearWatcherReloadExpected(saveSessionId, savePath);
    forceRebuildRef.current = true;
    setExternalUpdate(null);
    setReloadToken((value) => value + 1);
  }, [saveSessionId, savePath]);

  const reload = useCallback(() => {
    // Renderer mode: the preview owns reload (it bumps the revision we load).
    if (rendererRef.current && typeof rendererRef.current.reload === 'function') {
      try {
        rendererRef.current.reload();
        return;
      } catch {
        /* fall through to a local retry */
      }
    }
    setReloadToken((value) => value + 1);
  }, []);

  const openExternal = useCallback(() => {
    if (!mediaUrl) return;
    try {
      window.open(mediaUrl, '_blank', 'noopener,noreferrer');
    } catch {
      /* a blocked popup is not fatal */
    }
  }, [mediaUrl]);

  const loading = phase === 'loading' || phase === 'opening' || phase === 'converting';
  const failed = phase === 'error' && Boolean(error);
  const ready = phase === 'ready';
  const canUseEditor = ready && Boolean(managerRef.current);
  // The host renders its own file header; without it, keep our own title row.
  const showOwnTitle = props.toolbar !== 'host';

  /**
   * The single entry point behind Cmd/Ctrl+S: explains why nothing can be
   * written when the document has no native save target, otherwise saves.
   *
   * Defined AFTER `canUseEditor`: its dependency array is evaluated during
   * render, so referencing a later `const` here is a temporal-dead-zone crash.
   */
  const requestSave = useCallback(() => {
    if (!canUseEditor || !managerRef.current || savingRef.current) return;
    if (readOnly) {
      showNotice(t('saveReadOnly'));
      return;
    }
    if (!hasSaveTarget) {
      showNotice(t('saveNoTarget'));
      return;
    }
    if (!nativeSaveable) {
      showNotice(t('saveUnsupported'));
      return;
    }
    autoSaverRef.current?.cancel();
    void saveToFile(false);
  }, [canUseEditor, readOnly, hasSaveTarget, nativeSaveable, saveToFile, showNotice]);

  /** Auto-save entry: same guards, but it reports nothing on success. */
  const autoSave = useCallback(() => {
    if (ignoreChangesUntilRef.current > Date.now()) return;
    if (!canUseEditor || !managerRef.current || savingRef.current) return;
    if (readOnly || !hasSaveTarget || !nativeSaveable) return;
    void saveToFile(true);
  }, [canUseEditor, readOnly, hasSaveTarget, nativeSaveable, saveToFile]);

  autoSaveRef.current = autoSave;

  saveRef.current = requestSave;

  const loadingLabel =
    phase === 'opening' ? t('openingDocument') : phase === 'converting' ? t('converting') : t('loadingEngine');

  return (
    <div
      className="dshoo-root"
      title={canSaveTarget ? `${fileName} — ${saveShortcutLabel()} ${t('save')}` : fileName}
      onKeyDown={interceptSaveKey}
    >
      <div className="dshoo-toolbar" role="toolbar" aria-label="OnlyOffice">
        {showOwnTitle ? (
          <span className="dshoo-title" title={fileName}>
            {fileName}
          </span>
        ) : null}
        <span className="dshoo-badge">
          {busy === 'save' ? t('saving') : readOnly ? t('readOnly') : t('editable')}
        </span>
        <span className="dshoo-hint-inline" title={t('saveHint')}>
          {saveShortcutLabel()} {t('save')}
        </span>
        <span className="dshoo-spacer" />
        <button
          type="button"
          className="dshoo-btn"
          aria-pressed={readOnly}
          title={t('toggleReadOnly')}
          disabled={!canUseEditor || busy !== ''}
          onClick={toggleReadOnly}
        >
          {readOnly ? t('readOnly') : t('editable')}
        </button>
        <button
          type="button"
          className="dshoo-btn"
          disabled={!canUseEditor || busy !== ''}
          onClick={exportDocument}
          title={t('export')}
        >
          {busy === 'export' ? t('exporting') : t('export')}
        </button>
      </div>

      {externalUpdate ? (
        <div className="dshoo-notice dshoo-notice-action" role="status">
          <span>{t('externalChanged')}</span>
          <button type="button" className="dshoo-btn" onClick={applyExternalUpdate} disabled={busy !== ''}>
            {t('externalReload')}
          </button>
        </div>
      ) : null}

      {notice ? (
        <div className="dshoo-notice" role="status">
          <span>{notice}</span>
          <button type="button" onClick={() => setNotice('')} title={t('dismiss')}>
            ×
          </button>
        </div>
      ) : null}

      <div className="dshoo-stage">
        <div
          className="onlyoffice-container dshoo-frame"
          data-onlyoffice-container-id={containerId}
        >
          <div id={containerId} style={{ position: 'absolute', top: 0, right: 0, bottom: 0, left: 0 }} />
        </div>

        {loading ? (
          <div className="dshoo-state" role="status">
            <div className="dshoo-spinner" aria-hidden="true" />
            <div className="dshoo-state-title">{loadingLabel}</div>
            {slowHint ? <div className="dshoo-hint">{t('slowHint')}</div> : null}
          </div>
        ) : null}

        {failed ? (
          <div className="dshoo-state" role="alert">
            <div className="dshoo-state-title" data-tone="error">
              {t('loadFailed')}
            </div>
            <div className="dshoo-hint">{stageMessage(error.stage, error.message)}</div>
            {error.detail ? <pre className="dshoo-detail">{error.detail}</pre> : null}
            <div className="dshoo-hint">{stageHint(error.stage)}</div>
            <div className="dshoo-state-actions">
              <button type="button" className="dshoo-btn" onClick={reload}>
                {t('retry')}
              </button>
              <button type="button" className="dshoo-btn" onClick={openExternal} disabled={!mediaUrl}>
                {t('openExternal')}
              </button>
            </div>
          </div>
        ) : null}
      </div>
    </div>
  );
}
