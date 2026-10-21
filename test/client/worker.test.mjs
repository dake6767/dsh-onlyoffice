/**
 * Smoke checks for the inlined x2t worker bundle.
 *
 * The DSH client artifact is one file: `build/build-client.mjs` builds
 * `internal/editor/x2t.worker.ts` as an IIFE and embeds its source as a string
 * in `lib/client.js`. This test evaluates that IIFE in a Worker-shaped VM
 * sandbox and checks that it installs `onmessage`, announces `ready`, and
 * reports a failed conversion as an error message instead of hanging.
 *
 * Run with: node --test test/client
 */

import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { readFile } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import test from 'node:test';
import vm from 'node:vm';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..', '..');
const WORKER_PATH = join(ROOT, 'build', 'generated', 'x2t.worker.iife.js');

function createWorkerSandbox() {
  const posted = [];
  const listeners = [];

  const sandbox = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    location: { origin: 'http://127.0.0.1:53803', href: 'http://127.0.0.1:53803/' },
    setTimeout,
    clearTimeout,
    atob,
    TextDecoder,
    TextEncoder,
    Uint8Array,
    Uint32Array,
    ArrayBuffer,
    Blob,
    Promise,
    fetch: async () => {
      throw new Error('network disabled in the worker smoke test');
    },
    postMessage(message) {
      posted.push(message);
      for (const listener of [...listeners]) listener(message);
    },
    addEventListener(type, listener) {
      if (type === 'message') listeners.push(listener);
    },
    removeEventListener(type, listener) {
      const index = listeners.indexOf(listener);
      if (index >= 0) listeners.splice(index, 1);
    },
  };

  sandbox.self = sandbox;
  sandbox.globalThis = sandbox;

  return { sandbox, posted };
}

function waitFor(posted, predicate, timeoutMs = 2000) {
  const existing = posted.find(predicate);
  if (existing) return Promise.resolve(existing);

  return new Promise((resolvePromise, reject) => {
    const timer = setTimeout(() => {
      reject(new Error('timed out waiting for a worker message'));
    }, timeoutMs);
    // A message callback cannot be registered on the posting side, so poll the
    // array the postMessage stub fills (cheap, bounded, deterministic enough).
    const interval = setInterval(() => {
      const found = posted.find(predicate);
      if (!found) return;
      clearTimeout(timer);
      clearInterval(interval);
      resolvePromise(found);
    }, 5);
  });
}

test('the x2t worker artifact parses (node --check)', () => {
  execFileSync(process.execPath, ['--check', WORKER_PATH], { stdio: 'pipe' });
});

test('the worker installs onmessage and announces ready', async () => {
  const source = await readFile(WORKER_PATH, 'utf8');
  const { sandbox, posted } = createWorkerSandbox();

  vm.runInNewContext(source, sandbox, { filename: WORKER_PATH });

  assert.equal(typeof sandbox.onmessage, 'function', 'the worker must install self.onmessage');
  assert.ok(
    posted.some((message) => message && message.type === 'ready'),
    'the worker must post a ready message',
  );
});

test('a failing conversion is reported back as an error message', async () => {
  const source = await readFile(WORKER_PATH, 'utf8');
  const { sandbox, posted } = createWorkerSandbox();

  vm.runInNewContext(source, sandbox, { filename: WORKER_PATH });

  await sandbox.onmessage({
    data: {
      id: 7,
      type: 'convert',
      payload: {
        data: null,
        fileFrom: 'doc.docx',
        fileTo: 'Editor.bin',
        formatFrom: 65,
        formatTo: 0,
        staticResource: {
          x2t: {
            root: 'http://127.0.0.1:53803/dsh-onlyoffice/sdk/onlyoffice/9.4.0-develop/x2t',
            script:
              'http://127.0.0.1:53803/dsh-onlyoffice/sdk/onlyoffice/9.4.0-develop/x2t/x2t.js',
            wasm: 'http://127.0.0.1:53803/dsh-onlyoffice/sdk/onlyoffice/9.4.0-develop/x2t/x2t.wasm',
          },
        },
      },
    },
  });

  const error = await waitFor(posted, (message) => message && message.id === 7 && message.type === 'error');
  assert.equal(error.id, 7);
  assert.match(String(error.error), /network disabled/);
});
