import { afterEach, beforeEach, describe, expect, mock, test } from 'bun:test';
import { resolveAuth } from '../lib/auth.js';

const mockRead = mock();
const mockWrite = mock();
const mockFetch = mock();

mock.module('node:fs/promises', () => ({
  readFile: mockRead,
  writeFile: mockWrite,
  mkdir: mock(() => Promise.resolve())
}));

const mockGetJwtExpiration = mock(() => Date.now() + 3600_000);
const mockIsValidJwtStructure = mock(() => true);
const mockGetMicrosoftTenantPathSegment = mock(() => 'common');

mock.module('../lib/jwt-utils.js', () => ({
  getJwtExpiration: mockGetJwtExpiration,
  isValidJwtStructure: mockIsValidJwtStructure,
  getMicrosoftTenantPathSegment: mockGetMicrosoftTenantPathSegment
}));

describe('auth resolution', () => {
  let originalEnv: NodeJS.ProcessEnv;

  beforeEach(() => {
    originalEnv = { ...process.env };
    global.fetch = mockFetch as any as any;
    mockRead.mockClear();
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
    const result = await resolveAuth();
    expect(result.success).toBe(false);
    expect(result.error).toContain('Missing EWS_CLIENT_ID or EWS_REFRESH_TOKEN');
  });

  test('uses valid cached token', async () => {
    process.env.EWS_CLIENT_ID = 'client';
    process.env.EWS_REFRESH_TOKEN = 'refresh';

    mockRead.mockResolvedValue(
      JSON.stringify({
        accessToken: 'cached-access-token',
        refreshToken: 'cached-refresh-token',
        expiresAt: Date.now() + 1000_000
      })
    );

    const result = await resolveAuth();
    expect(result.success).toBe(true);
    expect(result.token).toBe('cached-access-token');
    expect(mockFetch).not.toHaveBeenCalled();
  });

  test('fetches new token if cache expired', async () => {
    process.env.EWS_CLIENT_ID = 'client';
    process.env.EWS_REFRESH_TOKEN = 'refresh';

    mockRead.mockResolvedValue(
      JSON.stringify({
        accessToken: 'expired-access-token',
        refreshToken: 'cached-refresh-token',
        expiresAt: Date.now() - 1000_000
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
