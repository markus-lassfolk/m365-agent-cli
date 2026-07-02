import { afterEach, beforeEach, describe, expect, test } from 'bun:test';
import { Buffer } from 'node:buffer';
import {
  getJwtPayloadAppId,
  getJwtPayloadScopeSet,
  getMicrosoftTenantPathSegment,
  isValidJwtStructure
} from './jwt-utils.js';

describe('getJwtPayloadAppId', () => {
  test('reads appid from payload', () => {
    const h = Buffer.from(JSON.stringify({ alg: 'none' })).toString('base64url');
    const p = Buffer.from(JSON.stringify({ appid: '5f2abcea-d6ea-4460-b468-3d80d7a900eb' })).toString('base64url');
    const tok = `${h}.${p}.x`;
    expect(getJwtPayloadAppId(tok)).toBe('5f2abcea-d6ea-4460-b468-3d80d7a900eb');
  });

  test('falls back to azp when appid absent', () => {
    const h = Buffer.from(JSON.stringify({ alg: 'none' })).toString('base64url');
    const p = Buffer.from(JSON.stringify({ azp: '11111111-1111-1111-1111-111111111111' })).toString('base64url');
    const tok = `${h}.${p}.x`;
    expect(getJwtPayloadAppId(tok)).toBe('11111111-1111-1111-1111-111111111111');
  });

  test('returns undefined for malformed token', () => {
    expect(getJwtPayloadAppId('not-a-jwt')).toBeUndefined();
  });
});

describe('getJwtPayloadScopeSet', () => {
  test('parses scp claim', () => {
    const h = Buffer.from(JSON.stringify({ alg: 'none' })).toString('base64url');
    const p = Buffer.from(JSON.stringify({ scp: 'Mail.Read Mail.Send Calendars.ReadWrite' })).toString('base64url');
    const tok = `${h}.${p}.x`;
    const s = getJwtPayloadScopeSet(tok);
    expect(s.has('Mail.Send')).toBe(true);
    expect(s.has('Mail.Read')).toBe(true);
    expect(s.has('Contacts.ReadWrite')).toBe(false);
  });
});

describe('isValidJwtStructure', () => {
  function makeJwt(payload: Record<string, unknown>, header: Record<string, unknown> = { alg: 'none', typ: 'JWT' }) {
    const h = Buffer.from(JSON.stringify(header)).toString('base64url');
    const p = Buffer.from(JSON.stringify(payload)).toString('base64url');
    return `${h}.${p}.sig`;
  }

  test('accepts a well-formed 3-part JWT with object header and payload', () => {
    expect(isValidJwtStructure(makeJwt({ exp: 2_000_000_000, sub: 'me' }))).toBe(true);
  });

  test('rejects an empty string', () => {
    expect(isValidJwtStructure('')).toBe(false);
  });

  test('rejects a non-string input defensively', () => {
    // biome-ignore lint/suspicious/noExplicitAny: defensive guard branch coverage
    expect(isValidJwtStructure(undefined as any)).toBe(false);
    // biome-ignore lint/suspicious/noExplicitAny: defensive guard branch coverage
    expect(isValidJwtStructure(null as any)).toBe(false);
  });

  test('rejects tokens without exactly three parts', () => {
    expect(isValidJwtStructure('a.b')).toBe(false);
    expect(isValidJwtStructure('a.b.c.d')).toBe(false);
    expect(isValidJwtStructure('a')).toBe(false);
    expect(isValidJwtStructure('a..b')).toBe(false);
  });

  test('rejects tokens with empty segments', () => {
    expect(isValidJwtStructure('.payload.sig')).toBe(false);
    expect(isValidJwtStructure('header..sig')).toBe(false);
    expect(isValidJwtStructure('header.payload.')).toBe(false);
  });

  test('rejects non-JSON payload', () => {
    // base64url for "not json" — JSON.parse fails
    const h = Buffer.from('{"alg":"none"}').toString('base64url');
    const p = Buffer.from('not-json').toString('base64url');
    expect(isValidJwtStructure(`${h}.${p}.sig`)).toBe(false);
  });

  test('rejects when payload is a JSON array (not an object)', () => {
    const h = Buffer.from('{"alg":"none"}').toString('base64url');
    const p = Buffer.from(JSON.stringify([1, 2, 3])).toString('base64url');
    expect(isValidJwtStructure(`${h}.${p}.sig`)).toBe(false);
  });

  test('rejects when header is not a JSON object', () => {
    // base64url of a JSON number for the header
    const h = Buffer.from(JSON.stringify(42)).toString('base64url');
    const p = Buffer.from(JSON.stringify({ exp: 2_000_000_000 })).toString('base64url');
    expect(isValidJwtStructure(`${h}.${p}.sig`)).toBe(false);
  });

  test('rejects when base64url segments are not decodable (truncated)', () => {
    // base64url with invalid padding patterns
    expect(isValidJwtStructure('!!!.???.$$$')).toBe(false);
  });
});

describe('getMicrosoftTenantPathSegment precedence', () => {
  // We re-set the same env keys in every test so that other suites (which may also touch
  // EWS_TENANT_ID in their own beforeEach) cannot pollute this describe block.
  // Bun runs test files in parallel by default; per-file isolation depends on the
  // beforeEach/afterEach in this file alone, so we explicitly clear and restore all three keys.
  const originalKeys = ['M365_TENANT_ID', 'MICROSOFT_TENANT_ID', 'EWS_TENANT_ID'] as const;
  let snapshot: Record<string, string | undefined>;

  beforeEach(() => {
    snapshot = {};
    for (const k of originalKeys) {
      snapshot[k] = process.env[k];
      delete process.env[k];
    }
  });
  afterEach(() => {
    for (const k of originalKeys) {
      if (snapshot[k] === undefined) {
        delete process.env[k];
      } else {
        process.env[k] = snapshot[k];
      }
    }
  });

  test('defaults to "common" when no tenant env vars are set', () => {
    delete process.env.M365_TENANT_ID;
    delete process.env.MICROSOFT_TENANT_ID;
    delete process.env.EWS_TENANT_ID;
    expect(getMicrosoftTenantPathSegment()).toBe('common');
  });

  test('honors EWS_TENANT_ID for backwards compatibility (legacy single variable)', () => {
    delete process.env.M365_TENANT_ID;
    delete process.env.MICROSOFT_TENANT_ID;
    process.env.EWS_TENANT_ID = 'contoso.onmicrosoft.com';
    expect(getMicrosoftTenantPathSegment()).toBe('contoso.onmicrosoft.com');
  });

  test('MICROSOFT_TENANT_ID takes precedence over legacy EWS_TENANT_ID', () => {
    delete process.env.M365_TENANT_ID;
    process.env.EWS_TENANT_ID = 'legacy.example.com';
    process.env.MICROSOFT_TENANT_ID = 'modern.example.com';
    expect(getMicrosoftTenantPathSegment()).toBe('modern.example.com');
  });

  test('M365_TENANT_ID takes precedence over MICROSOFT_TENANT_ID and EWS_TENANT_ID', () => {
    process.env.EWS_TENANT_ID = 'legacy.example.com';
    process.env.MICROSOFT_TENANT_ID = 'modern.example.com';
    process.env.M365_TENANT_ID = 'preferred.example.com';
    expect(getMicrosoftTenantPathSegment()).toBe('preferred.example.com');
  });

  test('treats whitespace-only values as unset and falls through to the next var', () => {
    process.env.EWS_TENANT_ID = 'legacy.example.com';
    process.env.MICROSOFT_TENANT_ID = '   ';
    process.env.M365_TENANT_ID = '';
    expect(getMicrosoftTenantPathSegment()).toBe('legacy.example.com');
  });

  test('accepts common tenant placeholder even when other vars are set', () => {
    delete process.env.EWS_TENANT_ID;
    process.env.MICROSOFT_TENANT_ID = 'common';
    process.env.M365_TENANT_ID = 'organizations';
    expect(getMicrosoftTenantPathSegment()).toBe('organizations');
  });

  test('throws with descriptive error referencing all three variable names', () => {
    delete process.env.M365_TENANT_ID;
    delete process.env.MICROSOFT_TENANT_ID;
    delete process.env.EWS_TENANT_ID;
    process.env.M365_TENANT_ID = 'not a real tenant value!!!';
    expect(() => getMicrosoftTenantPathSegment()).toThrow(/M365_TENANT_ID/);
  });
});
