import { describe, expect, test } from 'bun:test';
import { Buffer } from 'node:buffer';
import {
  getJwtPayloadAppId,
  getJwtPayloadScopeSet,
  getMicrosoftTenantPathSegment,
  isValidJwtStructure,
  resolveTenantPathSegment
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

describe('resolveTenantPathSegment precedence', () => {
  // The pure helper is independent of `process.env`, so these tests do not race against
  // other suites in the same Bun process that mutate shared env state.

  test('defaults to "common" when no tenant env vars are set', () => {
    expect(resolveTenantPathSegment({})).toBe('common');
  });

  test('honors EWS_TENANT_ID for backwards compatibility (legacy single variable)', () => {
    expect(resolveTenantPathSegment({ EWS_TENANT_ID: 'contoso.onmicrosoft.com' })).toBe('contoso.onmicrosoft.com');
  });

  test('MICROSOFT_TENANT_ID takes precedence over legacy EWS_TENANT_ID', () => {
    expect(
      resolveTenantPathSegment({
        EWS_TENANT_ID: 'legacy.example.com',
        MICROSOFT_TENANT_ID: 'modern.example.com'
      })
    ).toBe('modern.example.com');
  });

  test('M365_TENANT_ID takes precedence over MICROSOFT_TENANT_ID and EWS_TENANT_ID', () => {
    expect(
      resolveTenantPathSegment({
        EWS_TENANT_ID: 'legacy.example.com',
        MICROSOFT_TENANT_ID: 'modern.example.com',
        M365_TENANT_ID: 'preferred.example.com'
      })
    ).toBe('preferred.example.com');
  });

  test('treats whitespace-only values as unset and falls through to the next var', () => {
    expect(
      resolveTenantPathSegment({
        EWS_TENANT_ID: 'legacy.example.com',
        MICROSOFT_TENANT_ID: '   ',
        M365_TENANT_ID: ''
      })
    ).toBe('legacy.example.com');
  });

  test('accepts common tenant placeholder even when other vars are set', () => {
    expect(
      resolveTenantPathSegment({
        MICROSOFT_TENANT_ID: 'common',
        M365_TENANT_ID: 'organizations'
      })
    ).toBe('organizations');
  });

  test('throws with descriptive error referencing all three variable names', () => {
    expect(() => resolveTenantPathSegment({ M365_TENANT_ID: 'not a real tenant value!!!' })).toThrow(/M365_TENANT_ID/);
  });
});

describe('getMicrosoftTenantPathSegment (process.env wrapper)', () => {
  // This wrapper reads process.env directly. We do not assert the resolved value because it
  // is host-dependent; we only assert the call shape to lock in the production API.
  test('returns a non-empty string', () => {
    const result = getMicrosoftTenantPathSegment();
    expect(typeof result).toBe('string');
    expect(result.length).toBeGreaterThan(0);
  });
});
