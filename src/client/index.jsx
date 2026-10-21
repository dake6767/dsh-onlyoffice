// SPDX-License-Identifier: AGPL-3.0-only
// dsh-onlyoffice — see LICENSE and THIRD_PARTY_NOTICES.md.

/**
 * Client half of the dsh-onlyoffice plugin.
 *
 * Two integration paths, because DSH can host the file tree in more than one
 * sidebar implementation:
 *
 *   1. **DSH built-in document preview** (primary) — the shipped
 *      `@deepseek-ai/dsh-client-ui-sidebar-documentpreview` exposes
 *      `ctx.documentPreviews` plus the keyed `sidebar.right.tab.document` slot.
 *      Registering there makes .docx/.xlsx/.pptx open in the built-in right
 *      sidebar's file manager, with no third-party dependency.
 *   2. **dsh-better-sidebar** (optional) — when that plugin is installed and
 *      active, `ctx.betterSidebar.registerFileViewer` is used as well. The
 *      service is optional: its absence never blocks path 1.
 *
 * Both paths render the same vendored-onlyoffice engine from `office-view.jsx`;
 * the heavy engine lives only in this bundle.
 */

import { createElement } from 'react';
import { OfficeView } from './office-view.jsx';
import { IconDocx, IconPptx, IconXlsx } from './icons.jsx';
import { attachLocale, t } from './i18n.js';
import { installX2tWorkerFactory } from './sdk.js';
import { installStyles } from './styles.js';
import { installNativeDocumentPreview } from './native.jsx';
import { registerEffect } from './effect.js';
import { installDiagnostics } from './diagnostics.js';
import { resolveHostUrl } from './sdk.js';

/** Hard requirements: the built-in document preview registry and the slot
 *  registry it renders through. */
export const inject = ['documentPreviews', 'slots', 'locale'];

const VIEWER_EXTS = {
  docx: ['docx', 'doc', 'odt', 'rtf'],
  xlsx: ['xlsx', 'xls', 'ods', 'csv'],
  pptx: ['pptx', 'ppt', 'odp'],
};

function viewerComponent(kind) {
  const Component = function DshOnlyOfficeFileView(props) {
    return createElement(OfficeView, Object.assign({}, props, { kind }));
  };
  Component.displayName = `DshOnlyOfficeFileView(${kind})`;
  return Component;
}

const COMPONENTS = {
  docx: viewerComponent('docx'),
  xlsx: viewerComponent('xlsx'),
  pptx: viewerComponent('pptx'),
};

const ICONS = {
  docx: IconDocx,
  xlsx: IconXlsx,
  pptx: IconPptx,
};

const VIEWER_TITLES = {
  docx: () => t('viewerDocx'),
  xlsx: () => t('viewerXlsx'),
  pptx: () => t('viewerPptx'),
};

/** The three better-sidebar viewer descriptors (ids / exts / strategy / priority). */
export const viewerDescriptors = ['docx', 'xlsx', 'pptx'].map((kind) => ({
  id: `dsh-onlyoffice:${kind}`,
  title: VIEWER_TITLES[kind],
  icon: (size) => createElement(ICONS[kind], { size }),
  exts: VIEWER_EXTS[kind],
  priority: 10,
  fetchStrategy: 'mediaUrl',
  component: COMPONENTS[kind],
}));

/** Resolve `ctx.betterSidebar` defensively. `ctx.get` is the only safe look at
 *  an OPTIONAL service: a direct property read throws cordis's
 *  "cannot get property … without inject" guard when it was never injected. */
function readSidebarService(ctx) {
  try {
    const viaGet = ctx && typeof ctx.get === 'function' ? ctx.get('betterSidebar') : undefined;
    if (viaGet && typeof viaGet.registerFileViewer === 'function') return viaGet;
  } catch {
    /* the inject guard: the service is not provided in this host */
  }

  // Non-cordis / test hosts expose services as plain properties. Guarded,
  // because on a cordis context this exact read is what throws the guard.
  try {
    const direct = ctx && ctx.betterSidebar;
    if (direct && typeof direct.registerFileViewer === 'function') return direct;
  } catch {
    /* ignore */
  }

  try {
    const viaGet = ctx && typeof ctx.get === 'function' ? ctx.get('betterSidebar') : undefined;
    if (viaGet && typeof viaGet.registerFileViewer === 'function') return viaGet;
  } catch {
    /* ignore */
  }

  return undefined;
}

/** How long to wait for a late `betterSidebar` service (load-order race). */
const BETTER_SIDEBAR_POLL_MS = 400;
const BETTER_SIDEBAR_POLL_TICKS = 75; // ≈ 30 s

/**
 * Register the optional better-sidebar viewers once its service exists.
 *
 * The service may appear AFTER this plugin (the two client bundles are
 * independent), so three complementary paths are used:
 *   1. it is already resolvable → register now;
 *   2. cordis optional injection (waits for the service, disposal-safe);
 *   3. a bounded poll that covers hosts where `ctx.inject` is unavailable or
 *      the service is provided imperatively (`ctx.provide`) later on.
 * The `attached` latch makes repeated paths idempotent.
 */
function installBetterSidebarViewers(ctx) {
  let attached = false;

  const attach = (service) => {
    if (attached) return true;
    if (!service || typeof service.registerFileViewer !== 'function') return false;
    attached = true;
    for (const descriptor of viewerDescriptors) {
      registerEffect(
        ctx,
        () => service.registerFileViewer(descriptor),
        `dsh-onlyoffice: viewer ${descriptor.id}`,
      );
    }
    return true;
  };

  if (attach(readSidebarService(ctx))) return true;

  // 2. Optional cordis injection. Even the FEATURE CHECK is guarded: reading an
  // undeclared service name throws cordis's inject guard on some versions.
  let canInject = false;
  try {
    canInject = Boolean(ctx) && typeof ctx.inject === 'function';
  } catch {
    canInject = false;
  }
  if (canInject) {
    try {
      ctx.inject(['betterSidebar'], (scoped) => {
        attach(readSidebarService(scoped));
      });
    } catch (error) {
      console.warn('[dsh-onlyoffice] optional betterSidebar injection failed', error);
    }
  }

  // 3. Bounded poll for imperative/delayed provision; cleaned up with the
  // plugin fiber and stopped as soon as the service shows up.
  registerEffect(
    ctx,
    () => {
      let ticks = 0;
      const timer = setInterval(() => {
        ticks += 1;
        if (attach(readSidebarService(ctx)) || ticks >= BETTER_SIDEBAR_POLL_TICKS) {
          clearInterval(timer);
        }
      }, BETTER_SIDEBAR_POLL_MS);
      return () => clearInterval(timer);
    },
    'dsh-onlyoffice: betterSidebar waiter',
  );

  return attached;
}

/**
 * Client plugin body.
 * @param ctx the client cordis context (`ctx.documentPreviews`, `ctx.slots`,
 *   optional `ctx.betterSidebar` and `ctx.locale`).
 */
export function apply(ctx) {
  registerEffect(ctx, () => installStyles(), 'dsh-onlyoffice: viewer styles');
  registerEffect(
    ctx,
    () =>
      installDiagnostics(async (text) => {
        try {
          await fetch(resolveHostUrl(ctx, 'dsh-onlyoffice/diagnostic'), {
            method: 'POST',
            credentials: 'same-origin',
            headers: { 'content-type': 'application/json' },
            body: JSON.stringify({ text }),
            keepalive: true,
          });
        } catch {
          // The Host log is best-effort; the same text is always shown in the UI.
        }
      }),
    'dsh-onlyoffice: diagnostics',
  );
  registerEffect(ctx, () => installX2tWorkerFactory(), 'dsh-onlyoffice: x2t worker factory');
  registerEffect(ctx, () => attachLocale(ctx), 'dsh-onlyoffice: locale tracking');

  const native = installNativeDocumentPreview(ctx);
  const betterSidebar = installBetterSidebarViewers(ctx);

  if (!native && !betterSidebar) {
    console.warn(
      '[dsh-onlyoffice] neither the built-in document preview nor betterSidebar is available; Office viewers were not registered.',
    );
  }
}
