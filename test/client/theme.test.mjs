/**
 * Unit tests for the host light/dark detection used to pick the OnlyOffice
 * `uiTheme` (the viewer chrome itself only uses `--dsw-alias-*` tokens).
 *
 * Run with: node --test test/client
 */

import assert from 'node:assert/strict';
import test from 'node:test';
import { isHostDark, subscribeHostTheme } from '../../src/client/theme.js';

function makeWindow({
  baseToken = '',
  labelToken = '',
  colorScheme = '',
  attributes = {},
  classes = [],
  mediaDark = false,
} = {}) {
  const observers = [];

  class FakeMutationObserver {
    constructor(callback) {
      this.callback = callback;
      this.disconnected = false;
      observers.push(this);
    }

    observe() {}

    disconnect() {
      this.disconnected = true;
    }
  }

  const documentElement = {
    classList: { contains: (name) => classes.includes(name) },
    getAttribute: (name) => attributes[name] ?? null,
    style: {},
  };
  const body = {
    classList: { contains: () => false },
    getAttribute: () => null,
  };

  const win = {
    document: { documentElement, body },
    getComputedStyle: (element) => ({
      getPropertyValue: (name) => {
        if (element === documentElement && name === '--dsw-alias-bg-base') return baseToken;
        if (element === documentElement && name === '--dsw-alias-label-primary') return labelToken;
        return '';
      },
      colorScheme,
    }),
    matchMedia: () => ({
      matches: mediaDark,
      addEventListener() {},
      removeEventListener() {},
    }),
    MutationObserver: FakeMutationObserver,
    requestAnimationFrame: (callback) => callback(),
    setTimeout,
    observers,
    documentElement,
  };

  return win;
}

test('detects dark/light from the --dsw-alias-bg-base token', () => {
  assert.equal(isHostDark(makeWindow({ baseToken: '#ffffff' })), false);
  assert.equal(isHostDark(makeWindow({ baseToken: '#f5f5f7' })), false);
  assert.equal(isHostDark(makeWindow({ baseToken: '#111827' })), true);
  assert.equal(isHostDark(makeWindow({ baseToken: 'rgb(18, 18, 22)' })), true);
  assert.equal(isHostDark(makeWindow({ baseToken: 'rgba(250, 250, 252, 1)' })), false);
  assert.equal(isHostDark(makeWindow({ baseToken: 'oklch(0.27 0.02 265)' })), true);
  assert.equal(isHostDark(makeWindow({ baseToken: 'oklch(0.97 0.005 265)' })), false);
});

test('falls back to color-scheme, theme attributes and prefers-color-scheme', () => {
  // An unparseable token (e.g. color-mix()) must not mask the other signals.
  assert.equal(isHostDark(makeWindow({ baseToken: 'color-mix(in srgb, black, white)', colorScheme: 'dark' })), true);
  assert.equal(
    isHostDark(makeWindow({ baseToken: 'color-mix(in srgb, black, white)', attributes: { 'data-theme': 'dark' } })),
    true,
  );
  assert.equal(
    isHostDark(makeWindow({ baseToken: 'color-mix(in srgb, black, white)', classes: ['dark'] })),
    true,
  );
  assert.equal(isHostDark(makeWindow({ baseToken: 'color-mix(in srgb, black, white)', mediaDark: true })), true);
  assert.equal(isHostDark(makeWindow({ baseToken: 'color-mix(in srgb, black, white)' })), false);
});

test('uses the label token as a second signal', () => {
  // Light text on a dark page: base token unparseable, label is light → dark mode.
  assert.equal(isHostDark(makeWindow({ labelToken: '#f5f5f7' })), true);
  assert.equal(isHostDark(makeWindow({ labelToken: '#1f2329' })), false);
});

test('subscribeHostTheme fires only on an actual change and stops on dispose', () => {
  const win = makeWindow({ baseToken: '#ffffff' });
  const received = [];

  const unsubscribe = subscribeHostTheme((dark) => received.push(dark), win);
  assert.equal(win.observers.length, 1);

  const observer = win.observers[0];

  // No change → no callback.
  observer.callback();
  assert.deepEqual(received, []);

  // Host switches to dark → one callback.
  win.documentElement.__base = '#111827';
  win.getComputedStyle = (element) => ({
    getPropertyValue: () => (element === win.documentElement ? '#111827' : ''),
    colorScheme: '',
  });
  observer.callback();
  assert.deepEqual(received, [true]);

  // Repeated notifications for the same mode are collapsed.
  observer.callback();
  assert.deepEqual(received, [true]);

  unsubscribe();
  assert.equal(observer.disconnected, true);
  observer.callback();
  assert.deepEqual(received, [true]);
});

test('tolerates a missing window/document', () => {
  assert.equal(isHostDark(undefined), false);
  const dispose = subscribeHostTheme(() => {}, undefined);
  assert.equal(typeof dispose, 'function');
  dispose();
});
