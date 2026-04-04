import { afterEach, beforeEach, describe, expect, mock, test } from 'bun:test';
import { resolveAuth } from '../lib/auth.js';

const mockRead = mock();
const mockWrite = mock();
const mockFetch = mock();

mock.module('node:fs/promises', () => ({
  readFile: mockRead,
  writeFile: mockWrite,
  mkdir: mock(() => Promise.resolve()),
  rename: mock(() => Promise.resolve()),
  unlink: mock(() => Promise.resolve())
}));

const mockGetJwtExpiration = mock(() => Date.now() + 3600_000);
const mockIsValidJwtStructure = mock(() => true);
const mockGetMicrosoftTenantPathSegment = mock(() => 'common');

mock.module('../lib/jwt-utils.js', () => ({
  getJwtExpiration: mockGetJwtExpiration,
  isValidJwtStructure: mockIsValidJwtStructure,
  getMicrosoftTenantPathSegment: mockGetMicrosoftTenantPathSegment
}));

/** Primary `token-cache-{identity}.json` only — no legacy `graph-token-cache` file. */
function mockPrimaryCacheOnly(json: string) {
  mockRead.mockImplementation((path: string | Buffer | URL) => {
    const p = String(path);
    if (p.includes('graph-token-cache')) {
      return Promise.reject(Object.assign(new Error('ENOENT'), { code: 'ENOENT' }));
    }
    return Promise.resolve(json);
  });
}

describe('auth resolution', () => {
  let originalEnv: NodeJS.ProcessEnv;

  beforeEach(() => {
    originalEnv = { ...process.env };
    global.fetch = mockFetch as any as any;
    mockRead.mockReset();
    mockRead.mockImplementation(() => Promise.resolve('{}'));
    mockWrite.mockClear();
    mockFetch.mockClear();
    mockGetJwtExpiration.mockClear();
    mockIsValidJwtStructure.mockClear();
  });

  afterEach(() => {
    process.env = originalEnv;
  });

  test('uses explicit token when provided', async () => {
    const result = await resolveAuth({ token: 'my-explicit-token' });
    expect(result.success).toBe(true);
    expect(result.token).toBe('my-explicit-token');
  });

  test('returns error if client ID or refresh token missing', async () => {
    delete process.env.EWS_CLIENT_ID;
    delete process.env.EWS_REFRESH_TOKEN;
    delete process.env.GRAPH_REFRESH_TOKEN;
    delete process.env.M365_REFRESH_TOKEN;
    const result = await resolveAuth();
    expect(result.success).toBe(false);
    expect(result.error).toContain('Missing EWS_CLIENT_ID or refresh token');
  });

  test('uses valid cached token', async () => {
    process.env.EWS_CLIENT_ID = 'client';
    process.env.EWS_REFRESH_TOKEN = 'refresh';

    mockPrimaryCacheOnly(
      JSON.stringify({
        version: 1,
        refreshToken: 'cached-refresh-token',
        ews: {
          accessToken: 'cached-access-token',
          expiresAt: Date.now() + 1000_000
        }
      })
    );

    const result = await resolveAuth();
    expect(result.success).toBe(true);
    expect(result.token).toBe('cached-access-token');
    expect(mockFetch).not.toHaveBeenCalled();
  });

  test('accepts legacy flat EWS cache shape', async () => {
    process.env.EWS_CLIENT_ID = 'client';
    process.env.EWS_REFRESH_TOKEN = 'refresh';

    mockPrimaryCacheOnly(
      JSON.stringify({
        accessToken: 'legacy-access',
        refreshToken: 'cached-refresh-token',
        expiresAt: Date.now() + 1000_000
      })
    );

    const result = await resolveAuth();
    expect(result.success).toBe(true);
    expect(result.token).toBe('legacy-access');
    expect(mockFetch).not.toHaveBeenCalled();
  });

  test('M365_REFRESH_TOKEN satisfies auth without EWS_REFRESH_TOKEN', async () => {
    process.env.EWS_CLIENT_ID = 'client';
    delete process.env.EWS_REFRESH_TOKEN;
    process.env.M365_REFRESH_TOKEN = 'unified-refresh';

    mockPrimaryCacheOnly(
      JSON.stringify({
        version: 1,
        ews: {
          accessToken: 'cached-access-token',
          expiresAt: Date.now() + 1000_000
        }
      })
    );

    const result = await resolveAuth();
    expect(result.success).toBe(true);
    expect(result.token).toBe('cached-access-token');
  });

  test('fetches new token if cache expired', async () => {
    process.env.EWS_CLIENT_ID = 'client';
    process.env.EWS_REFRESH_TOKEN = 'refresh';

    mockPrimaryCacheOnly(
      JSON.stringify({
        version: 1,
        refreshToken: 'cached-refresh-token',
        ews: {
          accessToken: 'expired-access-token',
          expiresAt: Date.now() - 1000_000
        }
      })
    );

    mockFetch.mockResolvedValue(
      new Response(
        JSON.stringify({
          access_token: 'new-access-token',
          refresh_token: 'new-refresh-token',
          expires_in: 3600
        }),
        { status: 200 }
      )
    );

    const result = await resolveAuth();
    expect(result.success).toBe(true);
    expect(result.token).toBe('new-access-token');
    expect(mockFetch).toHaveBeenCalled();
    expect(mockWrite).toHaveBeenCalled();
  });
});
