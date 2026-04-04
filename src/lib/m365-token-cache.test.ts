import { describe, expect, test } from 'bun:test';
import { assertValidCacheIdentity } from './m365-token-cache.js';

describe('assertValidCacheIdentity', () => {
  test('accepts default and common ids', () => {
    expect(assertValidCacheIdentity('default')).toBe('default');
    expect(assertValidCacheIdentity('beta_user-1')).toBe('beta_user-1');
  });

  test('rejects path injection', () => {
    expect(() => assertValidCacheIdentity('../evil')).toThrow(/Invalid token cache identity/);
    expect(() => assertValidCacheIdentity('a/b')).toThrow(/Invalid token cache identity/);
  });

  test('rejects empty and overlong', () => {
    expect(() => assertValidCacheIdentity('')).toThrow(/Invalid token cache identity/);
    expect(() => assertValidCacheIdentity('   ')).toThrow(/Invalid token cache identity/);
    expect(() => assertValidCacheIdentity('x'.repeat(129))).toThrow(/Invalid token cache identity/);
  });
});
