/**
 * Durable guard for the "theme tokens only" rule.
 *
 * The DSH Theme provider lists exactly the tokens below. A `var(--dsw-alias-…)`
 * reference to a token that does not exist is silently dropped by the browser,
 * so the rule is enforced statically: every token the viewer styles or the
 * light/dark detector reads must be in this allow-list.
 *
 * Run with: node --test test/client
 */

import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import test from 'node:test';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..', '..');

/** The tokens the installed Theme.listTokens provider reports. */
const ALLOWED_TOKENS = new Set([
  '--dsw-alias-bg-base',
  '--dsw-alias-bg-layer-1',
  '--dsw-alias-bg-layer-2',
  '--dsw-alias-bg-overlay',
  '--dsw-alias-border-l1',
  '--dsw-alias-border-l2',
  '--dsw-alias-brand-primary',
  '--dsw-alias-label-primary',
  '--dsw-alias-label-secondary',
  '--dsw-alias-state-error-primary',
  '--dsw-alias-state-idle-primary',
  '--dsw-alias-state-success-primary',
  '--dsw-alias-state-warn-primary',
  '--dsw-specific-sidebar-fill',
]);

/** Strip comments so documentation examples cannot mask a real violation. */
function stripComments(source) {
  return source
    .replace(/\/\*[\s\S]*?\*\//g, '')
    .replace(/(^|[^:])\/\/[^\n]*/g, '$1');
}

function tokensIn(source) {
  const matches = source.match(/--dsw-[a-z0-9-]+/g) || [];
  return [...new Set(matches)].sort();
}

async function read(file) {
  return readFile(join(ROOT, 'src', 'client', file), 'utf8');
}

test('styles.js references only tokens the Theme provider reports', async () => {
  const source = stripComments(await read('styles.js'));
  const tokens = tokensIn(source);

  assert.ok(tokens.length > 0, 'styles.js must reference theme tokens');
  const unknown = tokens.filter((token) => !ALLOWED_TOKENS.has(token));
  assert.deepEqual(unknown, [], `unknown theme tokens in styles.js: ${unknown.join(', ')}`);
});

test('theme.js reads only tokens the Theme provider reports', async () => {
  const source = stripComments(await read('theme.js'));
  const tokens = tokensIn(source);

  const unknown = tokens.filter((token) => !ALLOWED_TOKENS.has(token));
  assert.deepEqual(unknown, [], `unknown theme tokens in theme.js: ${unknown.join(', ')}`);
});

test('no literal colors for chrome (every color declaration is token-backed)', async () => {
  const source = stripComments(await read('styles.js'));

  assert.ok(!/#[0-9a-fA-F]{3,8}\b/.test(source), 'styles.js must not contain hex colors');
  assert.ok(
    !/\b(?:rgba?|hsla?|oklch|oklab|color-mix)\s*\(/.test(source),
    'styles.js must not contain color functions',
  );

  // Every declaration that paints must get its value from a theme token (or be
  // an explicit "no paint" value such as `transparent`).
  const declaration =
    /(?:^|\n)\s*(?:color|background|background-color|border|border-color|border-(?:top|right|bottom|left)(?:-color)?|outline|outline-color|fill|stroke)\s*:\s*([^;\n]+);/g;

  const violations = [];
  for (const match of source.matchAll(declaration)) {
    const value = match[1].trim();
    if (value === 'none' || value === 'transparent' || value === '0') continue;
    const literal = /#[0-9a-fA-F]{3,8}\b|\b(?:rgba?|hsla?|oklch|oklab|color-mix)\s*\(/.test(value);
    if (literal) violations.push(match[0].trim());
    else if (!value.includes('var(--dsw-')) violations.push(match[0].trim());
  }

  assert.deepEqual(violations, [], `non-token color declarations in styles.js:\n${violations.join('\n')}`);
});

test('the built bundle carries only allow-listed tokens', async () => {
  const bundle = await readFile(join(ROOT, 'lib', 'client.js'), 'utf8');
  const unknown = tokensIn(bundle).filter((token) => !ALLOWED_TOKENS.has(token));
  // The vendored tree may mention others in comments; the built bundle is what
  // matters. Any unknown token here means the chrome could lose a declaration.
  assert.deepEqual(unknown, [], `unknown theme tokens in lib/client.js: ${unknown.join(', ')}`);
});
