# dsh-onlyoffice

**English** | [中文](./README.zh.md)

[![npm version](https://img.shields.io/npm/v/@dake6767/dsh-onlyoffice.svg)](https://www.npmjs.com/package/@dake6767/dsh-onlyoffice)
[![license: AGPL-3.0](https://img.shields.io/badge/license-AGPL--3.0-blue.svg)](./LICENSE)
[![CI](https://github.com/dake6767/dsh-onlyoffice/actions/workflows/ci.yml/badge.svg)](https://github.com/dake6767/dsh-onlyoffice/actions/workflows/ci.yml)


A DeepSeek Harness plugin that opens **.docx / .xlsx / .pptx** files from the
Harness sidebar file managers and edits them in an embedded **OnlyOffice**
editor — client-side static SDK, **no Document Server, no backend conversion
service**. Documents never leave the machine; the only network traffic is
fetching the OpenOffice-format engine assets once (cached on disk).

| | |
|---|---|
| Package | `@dake6767/dsh-onlyoffice` (npm) — linked directory install in the `desktop` profile |
| Host half | `lib/index.js` — `/dsh-onlyoffice/*` routes (SDK proxy + cache, read/write, diagnostics) |
| Client half | `lib/client.js` — one module-loader bundle, three consumers of the vendored engine |
| Engine | [`onlyoffice-web-comp`](https://github.com/electroluxcode/onlyoffice-web-comp) vendored under `vendor/onlyoffice-web-comp` (patched: `vendor/PATCHES.md`) |

## What it does

- **Two integration paths, one engine**
  - **DSH built-in document preview** (primary, no third-party dependency):
    registers a document implementation through `ctx.documentPreviews` plus the
    keyed `sidebar.right.tab.document` slot (`loading: 'renderer'`).
  - **dsh-better-sidebar** (optional): registers three file viewers through
    `ctx.betterSidebar.registerFileViewer`. The service is awaited across load
    order (get → optional inject → bounded poll), so the two plugins compose in
    either boot order.
- **Edit + save back to the original file**
  - `⌘S` / `Ctrl+S` inside the editor (the vendor's save-shortcut blocker is
    re-pointed at this plugin) and the editor's own save affordance
    (`onSave` → `ONSAVE`) both write the document back.
  - **Auto-save**: every document change schedules a debounced (1.8 s idle)
    write-back; manual saves flush immediately. Saves are silent.
- **External changes are detected, never swapped underneath you**: the preview
  watcher reports a new file version; the viewer shows
  *“文件已在磁盘上更新。［重新加载］”* and you decide when to reload.
- **Export** the edited document as a download; toggle **read-only/edit**;
  editor UI theme follows DSH light/dark.
- **Diagnostics left in the product**: viewer failures render the captured
  console/iframe errors and an environment snapshot, and report to
  `POST /dsh-onlyoffice/diagnostic` (appended to `<cacheDir>/.requests.log`).

## Requirements

- A DSH build whose client exposes `documentPreviews` + `slots` (the shipped
  `@deepseek-ai/dsh-client-ui-sidebar-documentpreview`) and, for the optional
  path, `dsh-better-sidebar`.
- Network access to one of the SDK mirrors on first open (see below).

## Install
> Deployments behind a reverse proxy are supported: the plugin resolves its routes against the page directory, so a GUI served at `https://host/app/<id>/` works the same as `http://127.0.0.1:53803/` or the desktop shell `dsh-app://app/`.


The bundle is a linked directory; `plugin_manager install_bundle` with
`target: /Users/wuliqi/projects/dsh-onlyoffice` adds the dependency and the
profile bundle row. A page refresh (`Cmd+R`) loads client changes; **host-half
changes need a full DSH restart** (HMR is not reliable for the linked host
module).

```bash
node build/build-client.mjs   # vendor + src/client -> lib/client.js (582 KB)
node build/build-host.mjs     # host/*.js -> lib/*.js
node --test "test/**/*.test.mjs"
```

## OnlyOffice static SDK: where the engine comes from

The SDK tree is ~19.6k files / 1.1 GB and is distributed separately. The Host
reads a prepared shared copy when configured, then falls back to its writable
member cache and upstream mirrors.

```
GET /dsh-onlyoffice/sdk/onlyoffice/<version>/<rest>
  shared hit -> <DSH_ONLYOFFICE_SHARED_CACHE_DIR>/<version>/<rest>
  member hit -> <cacheDir>/<version>/<rest>
  cache miss -> <mirror>/onlyoffice/<version>/<rest>  (first working mirror)
```

Default ordered mirrors (failover on network error / timeout / 404 / 5xx):

1. `https://onlyoffice-web-comp.vercel.app/packages`
2. `https://cdn.jsdelivr.net/gh/electroluxcode/onlyoffice-web-comp@main/public/packages`
3. `https://raw.githubusercontent.com/electroluxcode/onlyoffice-web-comp/main/public/packages`

Definitive 404s (e.g. `dictionaries/en_US/*`, absent from every mirror) are
remembered for 10 minutes so they do not re-walk the mirror chain on every
render.

### Configuration (`cordis.patch.yml` row `onlyoffice`)

```yaml
- id: onlyoffice
  name: '@dake6767/dsh-onlyoffice'
  config:
    version: '9.4.0-develop'          # SDK directory segment (client reads it from /status)
    upstreamBase: 'https://onlyoffice-web-comp.vercel.app/packages'
    upstreamFallbacks: []             # [] = single source; omit for the defaults above
    cacheDir: '~/.dsh/onlyoffice-sdk'
    fetchTimeoutMs: 300000            # upstream budget per asset
    saveEnabled: true                 # false disables the write route
    autoSave: false                   # true enables the debounced background writer
                                      #   (manual Cmd/Ctrl+S and the editor save button always work)
    maxSaveBytes: 268435456           # 256 MiB write cap
```

Offline / intranet: pre-populate `cacheDir` with the same directory layout, or
point `upstreamBase` at a self-hosted static copy.

### Shared SDK with a writable member fallback

Set `DSH_ONLYOFFICE_SHARED_CACHE_DIR` (0.1.9+) to a prepared SDK root visible to
each Host process:

```bash
export DSH_ONLYOFFICE_SHARED_CACHE_DIR=/opt/onlyoffice-sdk
# Expected asset: /opt/onlyoffice-sdk/9.4.0-develop/web-apps/apps/api/documents/api.js
```

Mount this directory read-only for members. Shared hits contact no upstream and
write nothing into the shared tree. Missing, inaccessible or non-file entries
fall back to the member cache; symlinks resolving outside the shared root are
ignored. A newly available shared file takes priority over both a member copy
and a remembered upstream 404.

Downloads and request logs stay in the writable cache. Its existing precedence
is unchanged: explicit plugin `cacheDir`, then `DSH_ONLYOFFICE_CACHE_DIR`, then
`$DSH_HOME/onlyoffice-sdk`. The older `DSH_ONLYOFFICE_CACHE_DIR` setting still
selects a writable cache; use the separate shared setting for a read-only mount.
`/status` file and byte counts describe only the writable cache.

## HTTP routes (all behind the connection fence, same as `/sidebar/*`)

| Route | Purpose |
|---|---|
| `GET /status` | `{version, sdkBaseUrl, upstreams, cacheDir, files, bytes, lastError}` |
| `GET /sdk/onlyoffice/<ver>/<rest>` | SDK proxy + disk cache (immutable headers) |
| `GET /read?sessionId&path[&cwd]` | document bytes + `x-dsh-file-version` (renderer mode) |
| `POST /save?sessionId&path[&cwd][&mtimeMs]` | atomic write-back of raw document bytes |
| `POST /diagnostic` | bounded browser diagnostic line → `.requests.log` |

**Write safety** (the only writer in the plugin): target must already exist and
be a regular file inside the session workspace (`fs.contains`, realpath fallback);
extension must be `docx/xlsx/pptx` (the engine's native output — writing docx
bytes over `.odt/.doc/.csv` would corrupt them, those stay export-only); optional
`mtimeMs` is an optimistic lock (409 instead of clobbering a changed file);
same-directory temp file + rename, original mode preserved.

## Known limitations
### Troubleshooting: "host SDK route unreachable" on a space page

DSH spaces can serve the workspace at `/app/<spaceId>/`. On such a page the
plugin routes are reachable **only under that prefix**; the same routes at the
origin root are rejected by the connection fence with `403 Forbidden`. The
client therefore probes the candidates at startup (page directory → origin →
root-relative) and reuses whichever answers `GET …/status` with JSON.

Run this in the page console to see exactly what each base returns:

```js
await window.__dshooProbe()
```

It reports HTTP status, content-type, timing and a body prefix for both
`/status` and the SDK `api.js`, so `403 / text/plain` (fence) is easy to tell
apart from `200 / text/html` (SPA fallback) or a timeout.

## File formats

| What | Formats |
|---|---|
| Opened in OnlyOffice (built-in preview & better-sidebar) | `docx` `doc` `odt` `rtf` · `xlsx` `xls` `ods` `csv` · `pptx` `ppt` `odp` |
| Written back to the original file | **only** `docx` `xlsx` `pptx` (the engine's native output; writing docx bytes over `.odt/.doc/.csv` would corrupt them) |
| Other formats (still previewed/edited by DSH's own viewers) | text, markdown, code, images, **pdf**, html… |
| Parsed by the engine but **not usable in this SDK export** | `pdf`/`pdfa`/`djvu`/`xps` (needs the **`pdfeditor`** app), Visio `vsdx/vssx/vstx/...` (needs **`visioeditor`**) — the shipped static SDK contains only documenteditor/spreadsheeteditor/presentationeditor, and `api.js` maps `documentType: 'pdf'` to `pdfeditor`, so those 404 |
| Engine input formats not registered here | `txt` `html` `mht` `epub` `fb2` `mobi` `docm` `dotx` `dotm` `fodt` `ott` `oform` `docxf` `xlsm` `xltx` `xltm` `xlsb` `fods` `ots` `ppsx` `pptm` `ppsm` `potx` `potm` `fodp` `otp` — add the extension to `NATIVE_EXTENSIONS` / `VIEWER_EXTS` and a MIME entry to enable, but keep them read-only/export-only |

The extension must match the actual content (x2t picks its parser by extension).
PDF deliberately stays on DSH's native viewer: it renders the embedded glyphs
directly, while routing it through OnlyOffice would require a separate
`pdfeditor` application and a conversion round-trip.



- **No real-time co-editing**: the static SDK has no collaboration channel.
  True character-level sync needs a Document Server / WOPI backend.
- External changes reload the **document** (on your click), not individual
  characters; the in-place `refreshFile` path is unreliable in this SDK build
  and is only a documented vendor patch, not the default.
- Save-back supports native OOXML formats only; legacy/ODF/CSV files open and
  export but cannot be written back.
- `dsh-resource://file/absolute/...` addresses carry no session, so save-back is
  unavailable for them (viewing works).
- First open of a document downloads its slice of the SDK (tens of MB); later
  opens are local. x2t conversion is a page-global worker — conversions are
  serialized by a lock.
- Viewer chrome is our own toolbar row; the reload control at the top-right
  belongs to the built-in preview, and the external-update banner is the only
  additional reload entry this plugin adds.

## Troubleshooting

| Symptom | Check |
|---|---|
| First open very slow / `loading engine` | expected cold download; watch `<cacheDir>` grow |
| `编辑器渲染超时` | viewer error detail shows console errors + iframe state |
| 404 for `dictionaries/en_US/*` | harmless (spellcheck dictionaries absent upstream), negative-cached |
| Save fails with `conflict` | file changed on disk since load — reload first |
| Save fails with `method not allowed` | the running Host predates the save route — restart DSH |
| Blank sidebar after a plugin update | client bundle crash — the viewer reports `[console-*]` lines; check `.requests.log` |

`<cacheDir>/.requests.log` records every SDK/read/save/diagnostic request with
status and duration (2 MiB rotation) — the fastest way to see what the browser
actually did.

## License

This package is distributed under the **GNU Affero General Public License v3.0**
(see [`LICENSE`](./LICENSE)); first-party sources carry
`SPDX-License-Identifier: AGPL-3.0-only` headers.

- The vendored editor engine `vendor/onlyoffice-web-comp` is upstream
  [`electroluxcode/onlyoffice-web-comp`](https://github.com/electroluxcode/onlyoffice-web-comp),
  also AGPL-3.0, modified per [`vendor/PATCHES.md`](./vendor/PATCHES.md); its
  license text is kept at `vendor/onlyoffice-web-comp/LICENSE`.
- OnlyOffice SDK assets are **not** part of this package — they are fetched at
  runtime from the configured mirrors and cached locally. Their terms (AGPL-3.0
  community code, commercial Developer Edition builds, an unclear `x2t` WASM
  build) are recorded in [`THIRD_PARTY_NOTICES.md`](./THIRD_PARTY_NOTICES.md);
  read that before publishing, self-hosting a mirror, or shipping commercially.
- The vendored brotli decoder keeps its Apache-2.0 headers.

OnlyOffice is a trademark of Ascensio System SIA; this plugin is an independent
integration.
