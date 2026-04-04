import { describe, expect, test } from 'bun:test';
import { getJwtPayloadAppId, getJwtPayloadScopeSet } from './jwt-utils.js';

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
