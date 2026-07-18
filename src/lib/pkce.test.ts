import { describe, expect, test } from 'bun:test';
import { createHash } from 'node:crypto';
import { codeChallengeFromVerifier, generateCodeVerifier, generateOAuthState, generatePkcePair } from './pkce.js';

describe('generateCodeVerifier', () => {
  test('produces a URL-safe string of at least 43 characters (RFC 7636 §4.1)', () => {
    const v = generateCodeVerifier();
    expect(v.length).toBeGreaterThanOrEqual(43);
    expect(v).toMatch(/^[A-Za-z0-9_-]+$/);
  });

  test('produces different verifiers on each call', () => {
    expect(generateCodeVerifier()).not.toBe(generateCodeVerifier());
  });
});

describe('codeChallengeFromVerifier', () => {
  test('computes BASE64URL(SHA256(verifier)) deterministically', () => {
    const verifier = 'dBjftJeZ4CVP-mB92K27uhbUJU1p1r_wW1gFWFOEjXk';
    const expected = createHash('sha256').update(verifier).digest('base64url');
    expect(codeChallengeFromVerifier(verifier)).toBe(expected);
    // Cross-check against the literal value from RFC 7636 Appendix B.
    expect(codeChallengeFromVerifier(verifier)).toBe('E9Melhoa2OwvFrEMTJguCHaoeK1t8URWbuGJSstw-cM');
  });
});

describe('generateOAuthState', () => {
  test('produces a non-empty hex string, different each call', () => {
    const a = generateOAuthState();
    const b = generateOAuthState();
    expect(a).toMatch(/^[0-9a-f]{32}$/);
    expect(a).not.toBe(b);
  });
});

describe('generatePkcePair', () => {
  test('challenge is derived from the paired verifier', () => {
    const pair = generatePkcePair();
    expect(pair.codeChallengeMethod).toBe('S256');
    expect(pair.codeChallenge).toBe(codeChallengeFromVerifier(pair.codeVerifier));
  });
});
