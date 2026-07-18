import { describe, expect, test } from 'bun:test';
import { Buffer } from 'node:buffer';
import { deepRedact, isSecretKeyName, looksLikeSecretValue } from './redact.js';

function fixtureJwt(): string {
  const h = Buffer.from(JSON.stringify({ alg: 'none', typ: 'JWT' })).toString('base64url');
  const p = Buffer.from(JSON.stringify({ upn: 'doris@lassfolk.net' })).toString('base64url');
  return `${h}.${p}.signature-part-here`;
}

describe('isSecretKeyName', () => {
  test('matches common secret field names', () => {
    for (const key of [
      'refreshToken',
      'access_token',
      'password',
      'passwd',
      'client_secret',
      'apiKey',
      'API_KEY',
      'Authorization',
      'auth_code',
      'cookie',
      'privateKey',
      'credential'
    ]) {
      expect(isSecretKeyName(key)).toBe(true);
    }
  });

  test('does not match safe field names', () => {
    for (const key of ['displayName', 'tenantId', 'cacheHealth', 'lastVerifiedAt', 'sizeBytes', 'schemaVersion']) {
      expect(isSecretKeyName(key)).toBe(false);
    }
  });
});

describe('looksLikeSecretValue', () => {
  test('flags JWT-shaped strings', () => {
    expect(looksLikeSecretValue(fixtureJwt())).toBe(true);
  });

  test('flags long opaque refresh-token-like strings', () => {
    expect(looksLikeSecretValue('0.AXoA1234567890abcdefghijklmnopqrstuvwxyzABCDEF')).toBe(true);
  });

  test('does not flag file paths (config dir, cache file paths) — legitimate diagnostic metadata', () => {
    expect(looksLikeSecretValue('/home/user/.config/m365-agent-cli/token-cache-default.json')).toBe(false);
    expect(looksLikeSecretValue('/home/user/.config/m365-agent-cli')).toBe(false);
    expect(looksLikeSecretValue('C:\\Users\\doris\\.config\\m365-agent-cli\\.env')).toBe(false);
  });

  test('does not flag bare lowercase filenames/identifiers even when long', () => {
    expect(looksLikeSecretValue('token-cache-default.json')).toBe(false);
    expect(looksLikeSecretValue('graph-token-cache-default.json')).toBe(false);
  });

  test('does not flag short/safe strings, sentences, dates, or booleans', () => {
    expect(looksLikeSecretValue('healthy')).toBe(false);
    expect(looksLikeSecretValue('2026-07-18T00:00:00.000Z')).toBe(false);
    expect(looksLikeSecretValue('This is a normal sentence with spaces.')).toBe(false);
    expect(looksLikeSecretValue(true)).toBe(false);
    expect(looksLikeSecretValue(42)).toBe(false);
    expect(looksLikeSecretValue(undefined)).toBe(false);
    expect(looksLikeSecretValue('doris@lassfolk.net')).toBe(false);
  });
});

describe('deepRedact', () => {
  test('redacts secret-named fields at any depth', () => {
    const input = {
      cli: { version: '2026.7.7' },
      auth: {
        identity: 'doris',
        refreshToken: 'super-secret-value-that-should-never-appear',
        nested: { password: 'hunter2hunter2', ok: true }
      }
    };
    const out = deepRedact(input);
    expect(out.cli.version).toBe('2026.7.7');
    expect(out.auth.identity).toBe('doris');
    expect(out.auth.refreshToken).toBe('[REDACTED]');
    expect(out.auth.nested.password).toBe('[REDACTED]');
    expect(out.auth.nested.ok).toBe(true);
  });

  test('redacts token/password-shaped VALUES even under an innocuous key name', () => {
    const input = { note: fixtureJwt(), other: 'access is denied' };
    const out = deepRedact(input);
    expect(out.note).toBe('[REDACTED]');
    // "access is denied" contains the substring "access" but has spaces and isn't opaque-looking.
    expect(out.other).toBe('access is denied');
  });

  test('redacts inside arrays', () => {
    const input = { items: [{ apiKey: 'abc' }, { name: 'safe' }] };
    const out = deepRedact(input);
    expect(out.items[0].apiKey).toBe('[REDACTED]');
    expect(out.items[1].name).toBe('safe');
  });

  test('leaves numbers, booleans, null, and non-secret strings untouched', () => {
    const input = { count: 3, ok: false, missing: null, label: 'diagnostic-bundle' };
    expect(deepRedact(input)).toEqual(input);
  });

  test('a boolean/number field whose NAME merely contains a secret-ish word is not itself redacted (regression)', () => {
    // `secretsPrinted: false` is a literal safety marker in the readiness/doctor JSON contracts —
    // it must stay a boolean, not become the string "[REDACTED]" just because its key contains "secret".
    const input = { secretsPrinted: false, tokenCount: 2, hasRefreshToken: true };
    expect(deepRedact(input)).toEqual(input);
  });

  test('caps recursion depth instead of overflowing the stack on pathological input', () => {
    let node: Record<string, unknown> = { leaf: 'safe' };
    for (let i = 0; i < 50; i++) {
      node = { child: node };
    }
    expect(() => deepRedact(node, { maxDepth: 5 })).not.toThrow();
  });

  test('safeKeys exempts a declared display-identifier field from the value-shape heuristic (regression)', () => {
    // A long, mixed-case, digit-containing operator-chosen identifier (e.g. a profile name or
    // cache identity slug) would otherwise match looksLikeSecretValue's high-entropy heuristic and
    // get blanked out of the exact field that exists to say which identity a bundle is about.
    const input = { name: 'ContosoProdMailboxAcct2024', other: 'ContosoProdMailboxAcct2024' };
    const result = deepRedact(input, { safeKeys: ['name'] });
    expect(result.name).toBe('ContosoProdMailboxAcct2024');
    expect(result.other).toBe('[REDACTED]');
  });

  test('safeKeys exempts string elements of a declared array field (e.g. profile names list)', () => {
    const input = { names: ['ContosoProdMailboxAcct2024', 'short'] };
    const result = deepRedact(input, { safeKeys: ['names'] });
    expect(result.names).toEqual(['ContosoProdMailboxAcct2024', 'short']);
  });

  test('safeKeys does NOT exempt a secret-named key even if declared safe — key-name pattern always wins', () => {
    const input = { name: 'not-a-secret-value', apiKey: 'ContosoProdMailboxAcct2024' };
    const result = deepRedact(input, { safeKeys: ['apiKey'] });
    expect(result.apiKey).toBe('[REDACTED]');
  });
});
