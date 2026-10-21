/**
 * Smoke checks for the built DSH client artifact (`lib/client.js`).
 *
 * No browser is available in this environment, so the bundle is verified by:
 *  1. `node --check` (the file is a plain CJS script),
 *  2. static assertions on the module-loader wrapper / externals / inlined bits,
 *  3. actually evaluating the bundle in a VM sandbox with a fake
 *     `window.__ModuleLoader__`, executing the factory, and calling `apply()`
 *     against a stub better-sidebar service to confirm the three viewer
 *     descriptors are registered, and that a missing service does not throw.
 *
 * Run with: node --test test/client
 */

import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { readFile, stat } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import test from 'node:test';
import vm from 'node:vm';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..', '..');
const BUNDLE_PATH = join(ROOT, 'lib', 'client.js');
const CLIENT_ID = '@dake6767/dsh-onlyoffice';
const MAX_BUNDLE_BYTES = 3 * 1024 * 1024;

const EXPECTED_NATIVE_EXTS = [
  'docx', 'doc', 'odt', 'rtf',
  'xlsx', 'xls', 'ods', 'csv',
  'pptx', 'ppt', 'odp',
];

const EXPECTED_VIEWERS = {
  docx: ['docx', 'doc', 'odt', 'rtf'],
  xlsx: ['xlsx', 'xls', 'ods', 'csv'],
  pptx: ['pptx', 'ppt', 'odp'],
};

function readBundle() {
  return readFile(BUNDLE_PATH, 'utf8');
}

// ── 1. syntax ────────────────────────────────────────────────────────────────

test('lib/client.js parses as a CommonJS script (node --check)', () => {
  execFileSync(process.execPath, ['--check', BUNDLE_PATH], { stdio: 'pipe' });
});

// ── 2. wrapper / externals / inlined worker ─────────────────────────────────

test('bundle declares exactly one module-loader registration for dsh-onlyoffice', async () => {
  const bundle = await readBundle();

  const loads = bundle.match(/__ModuleLoader__\.load\s*\(/g) || [];
  assert.equal(loads.length, 1, 'expected exactly one __ModuleLoader__.load call');

  assert.match(
    bundle,
    /^window\.__ModuleLoader__\.load\(\{\s*id:\s*"@dake6767\/dsh-onlyoffice",\s*factory:\s*\(require\)\s*=>\s*\{/,
    'bundle must start with the expected banner',
  );
  assert.ok(
    bundle.trimEnd().endsWith('return module.exports; } });'),
    'bundle must end with the expected factory footer',
  );
  assert.ok(bundle.includes(`id: "${CLIENT_ID}"`), 'bundle id must equal the package name');
});

test('react / react-dom / @deepseek-ai stay external (not inlined)', async () => {
  const bundle = await readBundle();

  // The factory receives the module-table `require`; these calls prove the
  // externals were kept out of the bundle body.
  assert.match(bundle, /require\("react"\)/, 'react must stay an external require');
  assert.match(bundle, /require\("react\/jsx-runtime"\)/, 'the automatic JSX runtime must stay external');

  // Any of these markers means a client/Harness package (or the React runtime
  // itself) leaked into the bundle.
  for (const marker of [
    'node_modules/react',
    'react-dom.development.js',
    'react-jsx-runtime.development.js',
    'Minified React error',
    'node_modules/@deepseek-ai/',
    'createRoot(',
  ]) {
    assert.ok(!bundle.includes(marker), `bundle must not inline ${marker}`);
  }

  // No dependency of ours should have dragged a node_modules tree in.
  assert.ok(!bundle.includes('node_modules/'), 'bundle must not embed node_modules paths');
});

test('bundle has no import.meta / import.meta.url and no worker URL file', async () => {
  const bundle = await readBundle();

  assert.ok(!/\bimport\.meta\b/.test(bundle), 'no import.meta may remain');
  assert.ok(!bundle.includes('new URL("./x2t.worker'), 'the module-URL worker construction must be gone');
  assert.ok(!bundle.includes('type: "module"'), 'the bundle must not construct a module worker from a URL');

  // The x2t worker is inlined as an IIFE string and started from a Blob URL.
  assert.ok(bundle.includes('[x2t.worker]'), 'the inlined x2t worker source must be present');
  assert.ok(bundle.includes('URL.createObjectURL'), 'the worker must be created from a Blob URL');
});

test('the editor save affordances are wired into the bundle', async () => {
  const bundle = await readBundle();

  // Vendor patch 6/7: the shortcut blocker and DocsAPI onSave must both exist,
  // otherwise Cmd/Ctrl+S and the editor's own save button silently do nothing.
  assert.match(bundle, /setSaveShortcutHandler/, 'the save-shortcut hook must be present');
  assert.match(bundle, /setChangeHandler/, 'the auto-save change hook must be present');
  assert.match(bundle, /onSave:\s*\(\)\s*=>/, 'DocsAPI onSave must be registered');
  assert.match(bundle, /ONSAVE/, 'the viewer must listen for the ONSAVE event');
});

test('bundle stays under the 3 MiB budget', async () => {
  const info = await stat(BUNDLE_PATH);
  assert.ok(info.size > 0, 'bundle must not be empty');
  assert.ok(
    info.size < MAX_BUNDLE_BYTES,
    `bundle is ${info.size} bytes, expected < ${MAX_BUNDLE_BYTES}`,
  );
});

// ── 3. sandbox execution ────────────────────────────────────────────────────

function createFakeDocument() {
  const appended = [];

  const createElement = (tagName) => ({
    tagName: String(tagName).toUpperCase(),
    style: {},
    dataset: {},
    textContent: '',
    attributes: {},
    setAttribute(name, value) {
      this.attributes[name] = value;
    },
    getAttribute(name) {
      return this.attributes[name] ?? null;
    },
    appendChild() {},
    remove() {
      this.removed = true;
    },
    querySelector() {
      return null;
    },
  });

  const document = {
    baseURI: 'http://127.0.0.1:53803/',
    documentElement: {
      classList: { contains: () => false },
      getAttribute: () => null,
      style: {},
    },
    body: {
      classList: { contains: () => false },
      getAttribute: () => null,
    },
    head: {
      appendChild(node) {
        appended.push(node);
      },
    },
    createElement,
    querySelector() {
      return null;
    },
    querySelectorAll() {
      return [];
    },
    getElementById() {
      return null;
    },
    appended,
  };

  return document;
}

function makeCreateElement() {
  return (type, props, ...children) => ({
    $$element: true,
    type,
    props: { ...(props || {}), children: children.length ? children : undefined },
  });
}

function createSandbox() {
  const createElement = makeCreateElement();
  const fakeReact = {
    createElement,
    useState: (initial) => [typeof initial === 'function' ? initial() : initial, () => {}],
    useEffect: () => {},
    useCallback: (fn) => fn,
    useMemo: (factory) => factory(),
    useRef: (value) => ({ current: value }),
  };
  const fakeJsxRuntime = {
    Fragment: Symbol.for('react.fragment'),
    jsx: (type, props) => createElement(type, props),
    jsxs: (type, props) => createElement(type, props),
  };

  const document = createFakeDocument();

  const sandbox = {
    console,
    document,
    navigator: { language: 'zh-CN' },
    location: { origin: 'http://127.0.0.1:53803', href: 'http://127.0.0.1:53803/' },
    setTimeout,
    clearTimeout,
    setInterval,
    clearInterval,
    requestAnimationFrame: (fn) => setTimeout(fn, 0),
    cancelAnimationFrame: (handle) => clearTimeout(handle),
    MutationObserver: class {
      observe() {}
      disconnect() {}
    },
    matchMedia: () => ({ matches: false, addEventListener() {}, removeEventListener() {} }),
    getComputedStyle: () => ({
      getPropertyValue: () => '',
      colorScheme: 'light',
    }),
    URL,
    Blob,
    File: class File {},
    TextEncoder,
    TextDecoder,
    atob,
    btoa,
    AbortController,
    Event: class Event {},
    Response: class Response {},
    fetch: async () => {
      throw new Error('network disabled in the smoke test');
    },
  };

  sandbox.window = sandbox;
  sandbox.self = sandbox;
  sandbox.globalThis = sandbox;
  sandbox.__ModuleLoader__ = {
    load(entry) {
      sandbox.__loadedEntry = entry;
    },
  };

  const requireStub = (specifier) => {
    if (specifier === 'react') return fakeReact;
    if (specifier === 'react/jsx-runtime') return fakeJsxRuntime;
    throw new Error(`unexpected external module in the smoke test: ${specifier}`);
  };

  return { sandbox, requireStub };
}

async function loadClientExports() {
  const bundle = await readBundle();
  const { sandbox, requireStub } = createSandbox();

  vm.runInNewContext(bundle, sandbox, { filename: BUNDLE_PATH });

  const entry = sandbox.__loadedEntry;
  assert.ok(entry, 'the bundle must register a module-loader entry');
  assert.equal(entry.id, CLIENT_ID, 'module-loader entry id must be the package name');
  assert.equal(typeof entry.factory, 'function', 'the module-loader entry must expose a factory');

  const exports = entry.factory(requireStub);
  assert.ok(exports && typeof exports === 'object', 'the factory must return the module exports');
  return { exports, sandbox };
}

test('factory returns the cordis plugin shape (inject / apply)', async () => {
  const { exports } = await loadClientExports();
  // The array comes from the VM realm, so copy it before a strict deep compare.
  assert.deepEqual([...exports.inject], ['documentPreviews', 'slots', 'locale']);
  assert.equal(typeof exports.apply, 'function');
});

test('apply registers the three viewers on ctx.betterSidebar', async () => {
  const { exports } = await loadClientExports();

  const registered = [];
  const disposers = [];
  const ctx = {
    betterSidebar: {
      registerFileViewer(descriptor) {
        registered.push(descriptor);
        const dispose = () => {};
        disposers.push(dispose);
        return dispose;
      },
      hostRouteUrl(path) {
        return `http://127.0.0.1:53803/${String(path).replace(/^\/+/, '')}`;
      },
    },
    effect(factory) {
      const dispose = factory();
      return typeof dispose === 'function' ? dispose : () => {};
    },
  };

  exports.apply(ctx);
  assert.equal(registered.length, 3, 'expected three file viewers');

  const byKind = new Map(registered.map((descriptor) => [descriptor.id, descriptor]));
  assert.deepEqual([...byKind.keys()].sort(), [
    'dsh-onlyoffice:docx',
    'dsh-onlyoffice:pptx',
    'dsh-onlyoffice:xlsx',
  ]);

  for (const [kind, exts] of Object.entries(EXPECTED_VIEWERS)) {
    const descriptor = byKind.get(`dsh-onlyoffice:${kind}`);
    assert.ok(descriptor, `missing viewer for ${kind}`);
    assert.deepEqual([...descriptor.exts], exts, `wrong exts for ${kind}`);
    assert.equal(descriptor.fetchStrategy, 'mediaUrl');
    assert.equal(descriptor.priority, 10);
    assert.equal(typeof descriptor.component, 'function');
    assert.equal(typeof descriptor.title, 'function');
    assert.equal(typeof descriptor.title(), 'string');
    assert.ok(descriptor.title().length > 0);
    assert.equal(typeof descriptor.icon, 'function');

    const element = descriptor.component({
      ctx,
      scope: { sessionId: 'session-1' },
      path: `docs/report.${kind}`,
      title: `report.${kind}`,
      viewerId: descriptor.id,
      mediaUrl: `/sidebar/file?path=report.${kind}`,
    });
    assert.ok(element && element.$$element, 'the viewer component must render an element');

    // Render-level smoke: a TDZ / hook-order bug must fail HERE instead of
    // blanking the slot in the running app.
    const rendered = element.type(element.props);
    assert.ok(rendered && rendered.$$element, `the ${kind} viewer must render without throwing`);
    assert.equal(typeof rendered.props.className, 'string');
  }

  // Disposers can be called without throwing (unload path).
  for (const dispose of disposers) dispose();
});

test('apply registers the DSH-native document preview implementation and body', async () => {
  const { exports } = await loadClientExports();

  const definitions = [];
  const registrations = [];
  const ctx = {
    documentPreviews: {
      register(definition) {
        definitions.push(definition);
        return () => {};
      },
    },
    slots: {
      inject(key, callback) {
        assert.equal(key, 'sidebar.right.tab.document', 'must register into the document body slot');
        return callback();
      },
      register(options, component) {
        registrations.push({ options, component });
        return () => {};
      },
    },
    effect(factory) {
      const dispose = factory();
      return typeof dispose === 'function' ? dispose : () => {};
    },
  };

  exports.apply(ctx);

  assert.equal(definitions.length, 1, 'expected one document preview implementation');
  const definition = definitions[0];
  assert.equal(definition.id, 'dsh-onlyoffice:document');
  assert.equal(definition.priority, 'extension');
  assert.equal(definition.loading, 'renderer', 'the renderer owns loading in the built-in preview');
  assert.deepEqual([...definition.extensions].sort(), [...EXPECTED_NATIVE_EXTS].sort());
  assert.deepEqual([...definition.binaryExtensions].sort(), [...definition.extensions].sort());
  assert.equal(typeof definition.title, 'function');
  assert.ok(definition.title().length > 0);

  assert.equal(registrations.length, 1, 'expected one document body registration');
  assert.equal(registrations[0].options.name, 'sidebar.right.tab.document');
  assert.equal(registrations[0].options.key, definition.id, 'body key must be the implementation id');
  assert.equal(typeof registrations[0].component, 'function');

  // Renderer mode: the preview hands a revision + completion callbacks and the
  // renderer loads its own bytes (through our Host read route).
  const fileAddress = 'dsh-resource://file/session/session-1/docs/report.docx';
  const calls = [];
  const rendererContent = {
    kind: 'renderer',
    revision: 3,
    loaded: (version) => calls.push(['loaded', version]),
    failed: () => calls.push(['failed']),
    reload: () => calls.push(['reload']),
  };
  const wrapper = registrations[0].component({ resourceAddress: fileAddress, content: rendererContent });
  assert.ok(wrapper && wrapper.$$element, 'the registered body must render an element');
  assert.equal(wrapper.props.ctx, ctx, 'the plugin context must be closed over');
  const loaded = wrapper.type(wrapper.props);
  assert.ok(loaded && loaded.$$element, 'the body must render an element in renderer mode');
  // Render-level smoke for the built-in preview path as well.
  const nativeRendered = loaded.type(loaded.props);
  assert.ok(nativeRendered && nativeRendered.$$element, 'the body must render without throwing');
  assert.equal(loaded.props.kind, 'docx');
  assert.equal(loaded.props.title, 'report.docx');
  assert.equal(loaded.props.sessionId, 'session-1', 'save-back needs the session id');
  assert.equal(loaded.props.sourcePath, 'docs/report.docx', 'save-back needs the file path');
  assert.ok(loaded.props.renderer, 'renderer-mode hooks must be forwarded');
  assert.equal(loaded.props.renderer.revision, 3);
  assert.equal(typeof loaded.props.renderer.load, 'function');
  assert.equal(typeof loaded.props.renderer.loaded, 'function');
  assert.equal(typeof loaded.props.renderer.reload, 'function');

  // Without a renderer payload the body still renders (the preview always
  // provides one, but a stripped-down host must not crash the slot).
  const bare = registrations[0].component({ resourceAddress: fileAddress });
  const bareElement = bare.type(bare.props);
  assert.ok(bareElement && bareElement.$$element);
  assert.equal(bareElement.props.renderer, undefined);

  // `absolute` addresses carry no session: the view still renders, but save-back
  // has no session to validate a write against.
  const absoluteWrapper = registrations[0].component({
    resourceAddress: 'dsh-resource://file/absolute/Users/me/docs/report.docx',
    content: rendererContent,
  });
  const absolute = absoluteWrapper.type(absoluteWrapper.props);
  assert.equal(absolute.props.sessionId, '');
  assert.equal(absolute.props.sourcePath, '/Users/me/docs/report.docx');
});

test('apply picks up a betterSidebar service that appears after activation', async () => {
  const { exports } = await loadClientExports();

  const registered = [];
  const disposers = [];
  const state = { service: undefined };
  const ctx = {
    get(name) {
      return name === 'betterSidebar' ? state.service : undefined;
    },
    effect(factory) {
      const dispose = factory();
      const cleaned = typeof dispose === 'function' ? dispose : () => {};
      disposers.push(cleaned);
      return cleaned;
    },
  };

  exports.apply(ctx);
  assert.equal(registered.length, 0, 'nothing to register while the service is absent');

  state.service = {
    registerFileViewer(descriptor) {
      registered.push(descriptor);
      return () => {};
    },
  };

  const deadline = Date.now() + 3000;
  while (registered.length < 3 && Date.now() < deadline) {
    await new Promise((resolve) => setTimeout(resolve, 100));
  }
  assert.equal(registered.length, 3, 'the waiter must register once the service appears');
  assert.deepEqual(
    registered.map((descriptor) => descriptor.id).sort(),
    ['dsh-onlyoffice:docx', 'dsh-onlyoffice:pptx', 'dsh-onlyoffice:xlsx'],
  );

  for (const dispose of disposers) dispose();
});

test('apply tolerates a cordis context that guards undeclared services', async () => {
  const { exports } = await loadClientExports();

  // Mimic cordis: reading any service that was not injected throws.
  const target = {
    documentPreviews: { register: () => () => {} },
    slots: {
      inject: (_key, callback) => callback(),
      register: () => () => {},
    },
    locale: { subscribe: () => () => {} },
    effect(factory) {
      const dispose = factory();
      return typeof dispose === 'function' ? dispose : () => {};
    },
  };
  const ctx = new Proxy(target, {
    get(object, key) {
      if (key in object) return object[key];
      if (typeof key === 'symbol') return undefined;
      throw new Error(`cannot get property "${key}" without inject`);
    },
  });

  assert.doesNotThrow(() => exports.apply(ctx));
});

test('apply survives a host without betterSidebar', async () => {
  const { exports } = await loadClientExports();
  assert.doesNotThrow(() => exports.apply({}));
  assert.doesNotThrow(() => exports.apply(undefined));
});
