import { afterEach, beforeEach, describe, expect, mock, test } from 'bun:test';
import { GRAPH_DEVICE_CODE_LOGIN_SCOPES, GRAPH_REFRESH_SCOPE_CANDIDATES } from '../lib/graph-oauth-scopes.js';

const mockLoad = mock(() => Promise.resolve(null));
const mockSave = mock(() => Promise.resolve());
const mockFetch = mock(() =>
  Promise.resolve(
    new Response(JSON.stringify({ access_token: 't', refresh_token: 'r', expires_in: 3600 }), { status: 200 })
  )
);

// Re-export real login/refresh lists so parallel test workers do not break graph-oauth-scopes.test.ts.
mock.module('../lib/graph-oauth-scopes.js', () => ({
  GRAPH_DEVICE_CODE_LOGIN_SCOPES,
  GRAPH_REFRESH_SCOPE_CANDIDATES,
  GRAPH_CRITICAL_DELEGATED_SCOPES: [
    'Mail.Send',
    'Contacts.ReadWrite',
    'Notes.ReadWrite.All',
    'OnlineMeetings.ReadWrite'
  ]
}));

mock.module('../lib/jwt-utils.js', () => ({
  getMicrosoftTenantPathSegment: () => 'common',
  isValidJwtStructure: () => true,
  getJwtExpiration: () => Date.now() + 3_600_000,
  getJwtPayloadAppId: (token: string) => {
    try {
      const parts = token.split('.');
      if (parts.length !== 3) return undefined;
      const payload = JSON.parse(Buffer.from(parts[1], 'base64url').toString('utf8')) as {
        appid?: string;
        azp?: string;
      };
      return payload.appid || payload.azp;
    } catch {
      return undefined;
    }
  },
  getJwtPayloadScopeSet: (token: string) => {
    try {
      const parts = token.split('.');
      if (parts.length !== 3) return new Set<string>();
      const payload = JSON.parse(Buffer.from(parts[1], 'base64url').toString('utf8')) as { scp?: string };
      if (typeof payload.scp !== 'string') return new Set<string>();
      return new Set(
        payload.scp
          .split(/\s+/)
          .map((s) => s.trim())
          .filter(Boolean)
      );
    } catch {
      return new Set();
    }
  }
}));

mock.module('../lib/m365-token-cache.js', () => ({
  loadM365TokenCache: mockLoad,
  saveM365TokenCache: mockSave,
  getUnifiedRefreshTokenFromEnv: () =>
    process.env.M365_REFRESH_TOKEN || process.env.GRAPH_REFRESH_TOKEN || process.env.EWS_REFRESH_TOKEN
}));

import { resolveGraphAuth } from '../lib/graph-auth.js';

describe('resolveGraphAuth', () => {
  let originalEnv: NodeJS.ProcessEnv;

  beforeEach(() => {
    originalEnv = { ...process.env };
    global.fetch = mockFetch as unknown as typeof fetch;
    mockLoad.mockReset();
    mockLoad.mockImplementation(() => Promise.resolve(null));
    mockSave.mockClear();
    mockFetch.mockClear();
  });

  afterEach(() => {
    process.env = originalEnv;
  });

  test('returns error when EWS_CLIENT_ID is missing', async () => {
    delete process.env.EWS_CLIENT_ID;
    process.env.M365_REFRESH_TOKEN = 'r';
    const r = await resolveGraphAuth();
    expect(r.success).toBe(false);
    expect(r.error).toContain('EWS_CLIENT_ID');
  });

  test('returns error when no refresh token in env', async () => {
    process.env.EWS_CLIENT_ID = 'client';
    delete process.env.M365_REFRESH_TOKEN;
    delete process.env.GRAPH_REFRESH_TOKEN;
    delete process.env.EWS_REFRESH_TOKEN;
    const r = await resolveGraphAuth();
    expect(r.success).toBe(false);
    expect(r.error).toContain('refresh token');
  });

  test('returns error for invalid identity name', async () => {
    process.env.EWS_CLIENT_ID = 'client';
    process.env.M365_REFRESH_TOKEN = 'r';
    const r = await resolveGraphAuth({ identity: 'bad id' });
    expect(r.success).toBe(false);
    expect(r.error).toContain('Invalid identity');
  });

  function makeAccessTokenJwt(appid: string, scp?: string): string {
    const h = Buffer.from(JSON.stringify({ alg: 'none', typ: 'JWT' })).toString('base64url');
    const payload: Record<string, unknown> = { appid, exp: 2_000_000_000 };
    if (scp !== undefined) payload.scp = scp;
    const p = Buffer.from(JSON.stringify(payload)).toString('base64url');
    return `${h}.${p}.x`;
  }

  test('ignores cached Graph token when app id does not match EWS_CLIENT_ID', async () => {
    process.env.EWS_CLIENT_ID = '5f2abcea-d6ea-4460-b468-3d80d7a900eb';
    process.env.M365_REFRESH_TOKEN = 'env-refresh';

    mockLoad.mockResolvedValue({
      version: 1,
      graph: {
        accessToken: makeAccessTokenJwt(
          'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa',
          'Mail.ReadWrite Calendars.ReadWrite User.Read'
        ),
        expiresAt: Date.now() + 3_600_000
      }
    } as never);

    mockFetch.mockImplementation(() =>
      Promise.resolve(
        new Response(
          JSON.stringify({
            access_token: makeAccessTokenJwt(
              '5f2abcea-d6ea-4460-b468-3d80d7a900eb',
              'Mail.Send Mail.ReadWrite Contacts.ReadWrite Notes.ReadWrite.All OnlineMeetings.ReadWrite User.Read'
            ),
            refresh_token: 'rotated',
            expires_in: 3600
          }),
          { status: 200 }
        )
      )
    );

    const r = await resolveGraphAuth();
    expect(r.success).toBe(true);
    expect(mockFetch).toHaveBeenCalled();
    expect(mockSave).toHaveBeenCalled();
  });

  test('ignores cached Graph token when critical scopes are missing (narrow token)', async () => {
    process.env.EWS_CLIENT_ID = '5f2abcea-d6ea-4460-b468-3d80d7a900eb';
    process.env.M365_REFRESH_TOKEN = 'env-refresh';

    mockLoad.mockResolvedValue({
      version: 1,
      graph: {
        accessToken: makeAccessTokenJwt(
          '5f2abcea-d6ea-4460-b468-3d80d7a900eb',
          'Mail.ReadWrite Calendars.ReadWrite User.Read'
        ),
        expiresAt: Date.now() + 3_600_000
      }
    } as never);

    mockFetch.mockImplementation(() =>
      Promise.resolve(
        new Response(
          JSON.stringify({
            access_token: makeAccessTokenJwt(
              '5f2abcea-d6ea-4460-b468-3d80d7a900eb',
              'Mail.Send Mail.ReadWrite Contacts.ReadWrite Notes.ReadWrite.All OnlineMeetings.ReadWrite User.Read'
            ),
            refresh_token: 'rotated',
            expires_in: 3600
          }),
          { status: 200 }
        )
      )
    );

    const r = await resolveGraphAuth();
    expect(r.success).toBe(true);
    expect(mockFetch).toHaveBeenCalled();
    expect(mockSave).toHaveBeenCalled();
  });

  test('uses explicit token without calling fetch', async () => {
    const r = await resolveGraphAuth({ token: 'explicit-access' });
    expect(r.success).toBe(true);
    expect(r.token).toBe('explicit-access');
    expect(mockFetch).not.toHaveBeenCalled();
  });

  test('falls back to second refresh token when first returns invalid_grant (stale cache)', async () => {
    process.env.EWS_CLIENT_ID = 'client';
    process.env.M365_REFRESH_TOKEN = 'env-refresh-good';

    mockLoad.mockResolvedValue({
      version: 1,
      refreshToken: 'cache-refresh-stale',
      graph: { accessToken: 'expired', expiresAt: Date.now() - 60_000 }
    } as never);

    let calls = 0;
    mockFetch.mockImplementation(() => {
      calls++;
      if (calls === 1) {
        return Promise.resolve(
          new Response(
            JSON.stringify({
              error: 'invalid_grant',
              error_description: 'AADSTS70000: Provided grant is invalid or malformed.'
            }),
            { status: 400 }
          )
        );
      }
      return Promise.resolve(
        new Response(
          JSON.stringify({
            access_token: makeAccessTokenJwt('11111111-1111-1111-1111-111111111111', 'User.Read'),
            refresh_token: 'rotated',
            expires_in: 3600
          }),
          { status: 200 }
        )
      );
    });

    const r = await resolveGraphAuth();
    expect(r.success).toBe(true);
    expect(r.token).toContain('eyJ');
    expect(mockFetch).toHaveBeenCalledTimes(2);
    expect(mockSave).toHaveBeenCalled();
  });

  test('fails when every refresh token candidate fails', async () => {
    process.env.EWS_CLIENT_ID = 'client';
    process.env.M365_REFRESH_TOKEN = 'bad';

    mockLoad.mockResolvedValue({
      version: 1,
      refreshToken: 'also-bad',
      graph: { accessToken: 'x', expiresAt: Date.now() - 60_000 }
    } as never);

    mockFetch.mockImplementation(() =>
      Promise.resolve(
        new Response(JSON.stringify({ error: 'invalid_grant', error_description: 'no' }), { status: 400 })
      )
    );

    const r = await resolveGraphAuth();
    expect(r.success).toBe(false);
    expect(r.error).toContain('Graph token refresh failed');
  });
});
