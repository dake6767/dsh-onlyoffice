# Third-party notices and license situation

Package: `@dake6767/dsh-onlyoffice`.

This file records what this plugin embeds, what it only downloads at runtime, and
which upstream terms apply. It is a factual inventory, not legal advice.

## What this package contains

| Component | Where | License status |
|---|---|---|
| First-party Host/Client code (`host/**`, `src/client/**`, `build/**`) | this repo | **no LICENSE declared yet** — add one before distributing |
| Vendored `onlyoffice-web-comp` component source | `vendor/onlyoffice-web-comp/**` (patched, see `vendor/PATCHES.md`) | **AGPL-3.0** (upstream repo `electroluxcode/onlyoffice-web-comp`, `LICENSE` + GitHub license metadata) |
| Vendored brotli decoder (inside the component, `internal/vendor/brotli-dec`) | `vendor/onlyoffice-web-comp/internal/vendor/brotli-dec/**` | **Apache-2.0** (Copyright 2013 Google Inc., per-file headers) |
| Build output `lib/client.js` | this repo | a combined work that includes the AGPL-3.0 component above |
| OnlyOffice static SDK assets (`web-apps/`, `sdkjs/`, `fonts/`, `x2t/`, `x2t-fonts/`, `plugins.json`, `themes.json`) | **NOT bundled** — fetched at runtime by `GET /dsh-onlyoffice/sdk/...` and cached under `~/.dsh/onlyoffice-sdk` | see “OnlyOffice assets” below |

OnlyOffice, the OnlyOffice logo and document-format names are trademarks of
Ascensio System SIA. This plugin is not affiliated with or endorsed by
Ascensio System SIA.

## Upstream project: `onlyoffice-web-comp`

- Repository: <https://github.com/electroluxcode/onlyoffice-web-comp>
- License: **GNU Affero General Public License v3.0** (root `LICENSE`).
- The component sources carry no per-file copyright headers; the only
  third-party code inside them is the brotli decoder noted above.
- This plugin vendors that source under `vendor/onlyoffice-web-comp/` and
  modifies it (9 documented patches). Those modifications are therefore
  derivative of an AGPL-3.0 work.

## OnlyOffice static SDK assets

The upstream project ships an exported static SDK under
`public/packages/onlyoffice/9.4.0-develop/` (~20,274 files; README calls it the
“Developer Edition Docker-exported 9.4 SDK”). That directory contains **no
OnlyOffice license file** — only unrelated third-party notices (Monaco editor,
webpack module notices, etc.). The applicable terms therefore come from
OnlyOffice itself and the referenced source projects:

- `ONLYOFFICE/web-apps` — AGPL-3.0 (the editor frontend).
- `ONLYOFFICE/sdkjs` — AGPL-3.0 (the document engine).
- `x2t` / `x2t.wasm` — the converter; the upstream README points at
  `cryptpad/onlyoffice-x2t-wasm`, which declares **no license** in its
  repository metadata, so the WASM build's redistribution terms are unclear.
- ONLYOFFICE Docs **Developer Edition** is a commercial edition; embedding or
  redistributing its exported assets in a distributed product normally requires
  a Developer Edition license.

**What this plugin does:** it does not redistribute those assets. It downloads
them on first use from the configured mirrors, writes them to a local cache, and
serves them to the local browser. The mirrors’ availability and their
redistribution rights are their operators’ responsibility.

**What that means for you:**

- *Private / internal use on your own machine* (the current setup): you are not
  distributing the AGPL work or the OnlyOffice assets; keep the license texts
  for reference and you are done.
- *Publishing this plugin*: the bundle embeds an AGPL-3.0 derivative, so the
  published package must be licensed under AGPL-3.0 (or a compatible license)
  with its source available. MIT/closed licensing of the whole bundle would not
  be compliant.
- *Self-hosting the SDK mirror* (pre-populating `cacheDir` or pointing
  `upstreamBase` at your own static host): you then become a redistributor of the
  OnlyOffice assets and must satisfy their terms — AGPL-3.0 source-offer
  obligations for the community code, or a Developer Edition license for the
  Developer Edition build, plus clarity on the `x2t` WASM build.
- *Commercial embedding*: obtain an ONLYOFFICE Developer Edition license, or
  build the SDK from the AGPL-3.0 sources and comply with AGPL for your own
  application.

## Recommended files before distribution

1. `LICENSE` — AGPL-3.0 (matches the vendored component; the simplest compliant
   choice for the combined bundle).
2. Keep this notice file in the published package.
3. Keep `vendor/PATCHES.md` (documents the modifications, as AGPL requires).
4. Do not describe the plugin as “OnlyOffice” or use its marks as your product
   identity; refer to it as an integration.

## References

- <https://github.com/electroluxcode/onlyoffice-web-comp/blob/main/LICENSE>
- <https://github.com/ONLYOFFICE/web-apps>
- <https://github.com/ONLYOFFICE/sdkjs>
- <https://github.com/cryptpad/onlyoffice-x2t-wasm>
- <https://www.onlyoffice.com/>
