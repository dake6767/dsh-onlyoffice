// SPDX-License-Identifier: AGPL-3.0-only
// dsh-onlyoffice — see LICENSE and THIRD_PARTY_NOTICES.md.

/**
 * DSH-native document-preview integration.
 *
 * The built-in right-sidebar document preview is extensible through two public
 * contracts of `@deepseek-ai/dsh-client-ui-sidebar-documentpreview`:
 *
 *   1. `ctx.documentPreviews.register(definition)` — metadata: which file
 *      suffixes this implementation claims. `priority: 'extension'` outranks
 *      the product implementations (including the built-in PDF-conversion
 *      Office preview), so .docx/.xlsx/.pptx automatically select us.
 *   2. the keyed `sidebar.right.tab.document` slot, registered under the
 *      definition's `id` — the body component. With
 *      `loading: 'bytes-complete'` the document owner reads the whole file
 *      through its authorized Host Remote and hands the bytes to the body as
 *      `content = { kind: 'bytes', data }`.
 *
 * This is the integration that works with DSH's own file manager and does not
 * depend on any third-party sidebar plugin.
 */

import { createElement } from 'react';
import { OfficeView } from './office-view.jsx';
import { readDocumentFile } from './sdk.js';
import { t } from './i18n.js';
import { registerEffect } from './effect.js';

/** Implementation id; also the keyed document-slot key. */
export const NATIVE_DOCUMENT_ID = 'dsh-onlyoffice:document';

/** Every suffix this implementation claims (declared binary, so the preview
 *  never offers its plain-text fallback for them). */
export const NATIVE_EXTENSIONS = Object.freeze([
  'docx',
  'doc',
  'odt',
  'rtf',
  'xlsx',
  'xls',
  'ods',
  'csv',
  'pptx',
  'ppt',
  'odp',
]);

/** Extension → vendored editor macro (`FILE_TYPE`). */
const EXT_TO_KIND = Object.freeze({
  docx: 'docx',
  doc: 'docx',
  odt: 'docx',
  rtf: 'docx',
  xlsx: 'xlsx',
  xls: 'xlsx',
  ods: 'xlsx',
  csv: 'xlsx',
  pptx: 'pptx',
  ppt: 'pptx',
  odp: 'pptx',
});

function extensionOf(name) {
  const match = /\.([A-Za-z0-9]{1,8})$/.exec(String(name || ''));
  return match ? match[1].toLowerCase() : '';
}

function baseNameOf(filePath) {
  const parts = String(filePath || '').split(/[\\/]/);
  return parts[parts.length - 1] || '';
}

/**
 * Decode the path a `dsh-resource://file/<session>/<path>` tab address names.
 * Purely string work: the bytes themselves are delivered by the owner.
 */
/**
 * Parse a `dsh-resource://file/…` address exactly like the shipped document
 * preview does:
 *
 *   dsh-resource://file/session/<sessionId>/<path segments…>
 *   dsh-resource://file/absolute/<absolute path segments…>
 *
 * Returns `{ scope, sessionId, path }`; `sessionId` is empty for an absolute
 * address (there is no session to validate a write against).
 * @param {string} address
 * @returns {{ scope: 'session' | 'absolute' | '', sessionId: string, path: string }}
 */
export function parseFileResourceAddress(address) {
  const empty = { scope: '', sessionId: '', path: '' };
  const text = String(address || '');
  const prefix = 'dsh-resource://file/';
  if (!text.startsWith(prefix)) return empty;
  const cut = text.search(/[?#]/);
  const body = text.slice(prefix.length, cut === -1 ? undefined : cut);
  const [scope, ...rest] = body.split('/');

  if (scope === 'session') {
    const [rawSession, ...segments] = rest;
    if (!rawSession || segments.length === 0) return empty;
    try {
      return {
        scope: 'session',
        sessionId: decodeURIComponent(rawSession),
        path: segments.map(decodeURIComponent).join('/'),
      };
    } catch {
      return empty;
    }
  }

  if (scope === 'absolute') {
    const unc = rest[0] === '' && rest.length > 1;
    const segments = (unc ? rest.slice(1) : rest).map((segment) => {
      try {
        return decodeURIComponent(segment);
      } catch {
        return segment;
      }
    });
    if (segments.length === 0 || segments[0] === '') return empty;
    const windowsStyle = /^[A-Za-z]:$/.test(segments[0]);
    return {
      scope: 'absolute',
      sessionId: '',
      path: unc ? `//${segments.join('/')}` : windowsStyle ? segments.join('/') : `/${segments.join('/')}`,
    };
  }

  return empty;
}

/** The path a `dsh-resource://file/…` address names ('' when unparsable). */
export function pathFromResourceAddress(address) {
  return parseFileResourceAddress(address).path;
}

/** The session a `dsh-resource://file/…` address names ('' for absolute). */
export function sessionIdFromResourceAddress(address) {
  return parseFileResourceAddress(address).sessionId;
}

/** The body registered in `sidebar.right.tab.document` (keyed by the id). */
export function NativeDocumentBody(props) {
  const address = props && props.resourceAddress ? props.resourceAddress : '';
  const parsedAddress = parseFileResourceAddress(address);
  const filePath = parsedAddress.path;
  const fileName = baseNameOf(filePath);
  const kind = EXT_TO_KIND[extensionOf(fileName)] || EXT_TO_KIND[extensionOf(filePath)] || 'docx';
  const content = props ? props.content : undefined;
  const owner = content && content.kind === 'renderer' ? content : undefined;
  const pluginCtx = props ? props.ctx : undefined;

  // `loading: 'renderer'`: the preview hands us a revision and completion
  // callbacks; the bytes come from this plugin's Host read route, and reporting
  // `loaded(version)` after a save keeps the preview from reloading the editor.
  const renderer = owner
    ? {
        revision: owner.revision,
        loaded: owner.loaded,
        failed: owner.failed,
        reload: owner.reload,
        load: (signal) =>
          readDocumentFile(pluginCtx, {
            sessionId: parsedAddress.sessionId,
            path: filePath,
            signal,
          }),
      }
    : undefined;

  return createElement(OfficeView, {
    ctx: pluginCtx,
    kind,
    // `path` is what the editor shows; `sourcePath` + `sessionId` are what the
    // save-back route writes to (path may be absolute or session-relative).
    path: filePath,
    sourcePath: filePath,
    sessionId: parsedAddress.sessionId,
    title: fileName || `document.${kind}`,
    renderer,
  });
}

/**
 * Register the metadata + body. Returns false when this DSH build does not
 * expose the document-preview extension points, so the caller can fall back to
 * a third-party sidebar integration.
 */
export function installNativeDocumentPreview(ctx) {
  const previews = ctx && ctx.documentPreviews;
  const slots = ctx && ctx.slots;
  if (
    !previews ||
    typeof previews.register !== 'function' ||
    !slots ||
    typeof slots.inject !== 'function' ||
    typeof slots.register !== 'function'
  ) {
    return false;
  }

  registerEffect(
    ctx,
    () =>
      previews.register({
        id: NATIVE_DOCUMENT_ID,
        extensions: NATIVE_EXTENSIONS,
        binaryExtensions: NATIVE_EXTENSIONS,
        priority: 'extension',
        title: () => t('viewerOnlyOffice'),
        loading: 'renderer',
      }),
    'dsh-onlyoffice: document preview implementation',
  );

  // Slot components receive the owner's props, not the plugin context: close
  // over it so the body can reach the host status route and the locale.
  const Body = (props) => createElement(NativeDocumentBody, Object.assign({}, props, { ctx }));

  registerEffect(
    ctx,
    () =>
      slots.inject('sidebar.right.tab.document', () =>
        slots.register(
          { name: 'sidebar.right.tab.document', key: NATIVE_DOCUMENT_ID },
          Body,
        ),
      ),
    'dsh-onlyoffice: document preview body',
  );

  return true;
}
