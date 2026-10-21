# Local patches to `vendor/onlyoffice-web-comp`

The vendored OnlyOffice static-SDK component is upstream
`onlyoffice-web-comp` (vanilla TS, no React, no external packages). It was
vendored unmodified except for the four patches below. Each patch is the
minimum needed for the DSH environment: a **single-file browser bundle**, no
module files, no package manager, and the DSH host page instead of a Tailwind
demo page.

Re-apply these patches when updating the vendored tree.

---

## Patch 1 — `internal/editor/x2t.ts`: worker factory hook instead of `import.meta.url`

**Why.** Upstream creates the x2t conversion worker with

```ts
this.worker = new Worker(new URL("./x2t.worker.ts", import.meta.url), {
  type: "module",
});
```

That only works when the sources are served as real ES modules. The DSH plugin
ships ONE browser file (`lib/client.js`, a CJS closure factory), so
`import.meta` is unavailable and there is no worker file URL to point at.

**Patched to.** A module-level factory hook plus a required factory:

```ts
export type X2tWorkerFactory = () => Worker;

let workerFactory: X2tWorkerFactory | null = null;

export function setX2tWorkerFactory(factory: X2tWorkerFactory | null): void {
  workerFactory = factory;
}

export function getX2tWorkerFactory(): X2tWorkerFactory | null {
  return workerFactory;
}

function createX2tWorker(): Worker {
  if (!workerFactory) {
    throw new Error(
      "x2t worker factory is not registered (call setX2tWorkerFactory before opening documents)",
    );
  }
  return workerFactory();
}
```

and inside `X2tConverter.init()`:

```ts
this.worker = createX2tWorker();
```

The plugin registers the factory in `apply()`
(`src/client/sdk.js → installX2tWorkerFactory`). It builds a **classic** worker
from the inlined IIFE source:

```ts
const blob = new Blob([X2T_WORKER_SOURCE], { type: "text/javascript" });
const blobUrl = URL.createObjectURL(blob);
const worker = new Worker(blobUrl); // classic worker: importScripts works
```

`build/build-client.mjs` compiles `x2t.worker.ts` separately (with the vendored
brotli decoder) and writes `build/generated/x2t-worker-source.js` containing that
IIFE source as a string.

**Behavior preserved.** `X2tConverter`'s public surface is unchanged
(`init` / `convert` / `terminate` / `isInitialized`), the same worker protocol is
used, and a classic worker is strictly more capable here than the upstream module
worker (the Emscripten script is loaded with `importScripts`, the upstream
preferred path). The only difference: a build that forgets to register the
factory now fails with an explicit message instead of an unusable URL.

---

## Patch 2 — `internal/editor/office-format.ts`: drop the external `exceljs` CSV fallback

**Why.** Upstream converts "complex" CSV to XLSX through a lazy ExcelJS import:

```ts
async function loadExcelJS() {
  const mod = await import("exceljs/dist/exceljs.min");
  return mod.default ?? mod;
}

export async function convertCsvBufferToXlsxBuffer(buffer: ArrayBuffer) {
  const ExcelJS = await loadExcelJS();
  …
}
```

- It is the only bare-package import in the vendored tree, so an esbuild bundle
  fails to resolve it.
- ExcelJS is not vendored and cannot be fetched from the DSH module table.

**Patched to.** The fallback now uses the file's own minimal XLSX writer:

```ts
export async function convertCsvBufferToXlsxBuffer(buffer: ArrayBuffer) {
  return createXlsxFromRows(parseCsvBuffer(buffer));
}
```

`createXlsxFromRows(rows)` and a small `xlsxColumnName()` helper were added
beside `createXlsxFromText`; `createXlsxFromText` now delegates to
`createXlsxFromRows([[text]])` (same one-cell output as before, no duplicated
package XML). Cells are written as `inlineStr`, so numbers/dates arrive as their
CSV text rather than typed cells — acceptable for a crash-avoidance fallback that
hands the workbook straight to x2t.

**Behavior preserved.** `convertCsvBufferToXlsxBuffer` still resolves to an XLSX
`ArrayBuffer`, `createXlsxFromText` output shape is unchanged, and the CSV path
(multi-line CSV, or x2t CSV conversion failure → `convertCsvViaXlsx`) keeps
working — verified in `test/client` by parsing the produced ZIP and its
`xl/worksheets/sheet1.xml`.

---

## Patch 3 — `util/initialize.ts`: hide the SDK preload iframe with inline styles

**Why.** Upstream styles the preload iframe with Tailwind utility classes:

```ts
iframe.className = "w-0 h-0 hidden absolute -z-10";
```

The DSH page has no Tailwind, so the classes are inert and the iframe renders as
a visible 300×150 blank box appended to `document.body` — a visible defect in the
host page.

**Patched to.** Inline styles/attributes that cannot depend on a stylesheet:

```ts
iframe.style.position = "absolute";
iframe.style.width = "0";
iframe.style.height = "0";
iframe.style.border = "0";
iframe.style.left = "-9999px";
iframe.style.top = "-9999px";
iframe.style.visibility = "hidden";
iframe.style.pointerEvents = "none";
iframe.setAttribute("aria-hidden", "true");
iframe.setAttribute("tabindex", "-1");
```

**Behavior preserved.** The preload iframe still warms up the editor HTML at the
same URL before the first document opens; it is just invisible and inert.

---

## Patch 4 — `core/editor-manager.ts`: no third-party logo asset in the editor config

**Why.** Upstream's `buildEditorCustomization()` points the editor logo at a
jsDelivr URL:

```ts
logo: {
  image: OFFICE_EDITOR_LOGO.image,       // https://cdn.jsdelivr.net/npm/simple-icons@9.21.0/...
  imageDark: OFFICE_EDITOR_LOGO.imageDark,
},
```

The editor iframe would then request a third-party host (broken/invisible logo
when offline, and an outbound request the DSH plugin does not need). The plugin
ships no branding asset and must stay self-contained.

**Patched to.** The `logo` block (and the now-unused `OFFICE_EDITOR_LOGO`
import) were removed; the editor uses OnlyOffice's built-in logo.

**Behavior preserved.** Everything else in the customization is byte-identical
(`uiTheme`, `autosave:false`, header/toolbar save suppression, review flags,
spellcheck); only the cosmetic logo source changes.

---

## Known unpatched upstream leftovers (not reachable in this plugin)

- `internal/ui/office-xml-size-limit-overlay.ts` hardcodes light-theme colors for
  its "document too large" overlay. It only renders when
  `officeXmlEvent.isEnable === true`; this plugin leaves the vendored default
  (`OFFICE_XML_EVENT_CONFIG.default.isEnable === false`), so it never appears.
- `const/index.ts` still exports `OFFICE_EDITOR_LOGO` (a jsDelivr URL) but no code
  path used by this plugin references it after patch 4.
- `x2t-assets.ts` may fetch PDF-export fonts (`x2t-fonts/*.ttf`) only when
  exporting to PDF; the DSH viewer exports Office formats (docx/xlsx/pptx), so
  the host SDK route never needs those paths for the current feature set.

---

## Patch 5 — `const/index.ts`: `resolveSiteUrl` accepts any URL scheme

**Why.** Upstream treats only `http(s)://` as absolute:

```ts
export function resolveSiteUrl(origin: string, path: string): string {
  if (/^https?:\/\//i.test(path)) return path;
  return `${origin}${path.startsWith("/") ? path : `/${path}`}`;
}
```

The DSH desktop shell serves the GUI from the custom `dsh-app://app` origin.
The SDK base is then absolute (`dsh-app://app/dsh-onlyoffice/sdk/…`, produced by
`registerOnlyOfficeStaticResource`), but `resolveSiteUrl` re-prefixed it — once
on the main thread (`x2t.ts` → worker payload) and again inside the blob-URL
worker (`x2t.worker.ts`, whose `self.location.origin` is the page origin) —
yielding `dsh-app://app/dsh-app://app/dsh-app://app/dsh-onlyoffice/…`. The
worker's `fetch(x2t/x2t.js)` 404'd, so document conversion never ran and the
editor never became ready (observed: "编辑器渲染超时" plus
`Failed to fetch dsh-app://app/dsh-app://app/…/x2t/x2t.js: 404`).

**Patched to.** Any scheme-qualified URL is absolute:

```ts
if (/^[a-z][a-z0-9+.-]*:\/\//i.test(path)) return path;
```

`isOnlyOfficeCdnMode()` deliberately keeps its `http(s)` test: a `dsh-app://`
root must stay in local (same-origin) mode, not switch to CDN mode.

---

## Patch 6 — `core/editor-manager.ts`: save shortcut hook + `onSave` bridge

**Why.** Upstream deliberately disables the editor's internal save
(`customization.autosave: false`, `layout.header.save: false`, and
*"不注册 onSave/onSaveDocument：内部保存已禁用"*), while
`installSaveShortcutBlocker()` still swallows Ctrl/Cmd+S inside the editor
(`preventDefault` + `stopPropagation` + `stopImmediatePropagation`). In the DSH
plugin that swallow made the shortcut a no-op: the browser dialog is suppressed
and nothing else happens.

**Patched to.** Two small wires, no behaviour change when unused:

1. `installSaveShortcutBlocker` now calls `this.saveShortcutHandler?.()` after
   stopping propagation, and the class exposes
   `setSaveShortcutHandler(handler | null)`. The plugin's viewer registers its
   write-back there, so Ctrl/Cmd+S inside the editor saves to the workspace.
2. `mountDocEditor`'s `events` gains `onSave: () => this.notifyUserSave()`.
   The editor's own save affordance (Quick Access floppy / File ▸ Save) now
   emits the existing `SAVE_DOCUMENT` + `ONSAVE` events, which the viewer
   listens to and turns into the same write-back.

`OnlyOfficeManager` forwards `setSaveShortcutHandler()` to its editor, alongside
the other facade methods. Upstream's `export()` path is untouched (it emits
`SAVE_DOCUMENT` only), so the viewer's own export never re-enters the save hook.

---

## Patch 7 — `core/onlyoffice-manager.ts`: façade for the shortcut hook

`setSaveShortcutHandler(handler)` delegating to `this.editor`, so the viewer can
register/clear the handler without reaching into the internal editor manager.

---

## Patch 8 — `core/editor-manager.ts`: document-change hook for auto-save

**Why.** The editor already reports content changes through DocsAPI's
`onDocumentStateChange` (upstream only sets an internal `dirty` flag from it),
and the DSH plugin needs an event to drive debounced write-back. Without it the
only save triggers were the explicit shortcut and the editor's save control.

**Patched to.** A `setChangeHandler(handler | null)` on the editor manager
(forwarded by `OnlyOfficeManager`); the existing
`onDocumentStateChange` callback now also calls it when the document becomes
dirty. The viewer schedules a debounced save from it (`src/client/autosave.js`,
1.8 s idle) and cancels/flushes it around explicit saves and editor disposal.
No upstream behaviour changes when no handler is registered.

---

## Patch 9 — in-place document refresh for external changes

**Why.** A file changed on disk re-enters the viewer as new bytes. Reusing the
document-open path (`destroy()` + `mountDocEditor()`) destroys and recreates the
whole DocEditor instance, so the sidebar's editor (toolbar included) visibly
reloads. DocsAPI's `refreshFile(config)` re-opens just the document inside the
existing instance.

**Patched to.** `EditorManager.refreshDocument(file, fileType?)`:

1. `server.open(file, …)` swaps the mock server's document (new id/key/url),
2. `editor.refreshFile(buildRefreshFileConfig(true))` makes the live editor load
   the new document.

`OnlyOfficeManager.refreshDocument(file)` forwards it. The viewer calls it when
the preview reports an external revision change, waits for `DOCUMENT_READY`,
then reports the new version via `loaded(version)`. If the in-place path is
unavailable (no editor, old SDK) it falls back to the full rebuild.
