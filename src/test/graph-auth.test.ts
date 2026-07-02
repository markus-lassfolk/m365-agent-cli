import { afterAll, afterEach, beforeEach, describe, expect, mock, test } from 'bun:test';
import { Buffer } from 'node:buffer';
import { randomBytes } from 'node:crypto';
import { mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import type { AuthResult } from '../lib/auth.js';
import type { GraphAuthResult } from '../lib/graph-auth.js';
import {
  GRAPH_CRITICAL_DELEGATED_SCOPES,
  GRAPH_DEVICE_CODE_LOGIN_SCOPES,
  GRAPH_REFRESH_SCOPE_CANDIDATES
} from '../lib/graph-oauth-scopes.js';
import * as jwtUtilsReal from '../lib/jwt-utils.js';

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
  GRAPH_CRITICAL_DELEGATED_SCOPES
}));

/** Captured once before `mock.module` mutates the live module namespace (spread/`import()` can still see mocked slots). */
const m365TokenCacheOriginal = await import('../lib/m365-token-cache.js');
const m365RealLoad = m365TokenCacheOriginal.loadM365TokenCache;
const m365RealSave = m365TokenCacheOriginal.saveM365TokenCache;
const m365RealGetUnified = m365TokenCacheOriginal.getUnifiedRefreshTokenFromEnv;
const m365RealAssertValid = m365TokenCacheOriginal.assertValidCacheIdentity;

function applyM365TokenCacheMockForGraph() {
  mock.module('../lib/m365-token-cache.js', () => ({
    loadM365TokenCache: mockLoad,
    saveM365TokenCache: mockSave,
    getUnifiedRefreshTokenFromEnv: () =>
      process.env.M365_REFRESH_TOKEN || process.env.GRAPH_REFRESH_TOKEN || process.env.EWS_REFRESH_TOKEN,
    assertValidCacheIdentity: m365RealAssertValid
  }));
}

function restoreRealM365TokenCacheModule() {
  mock.module('../lib/m365-token-cache.js', () => ({
    loadM365TokenCache: m365RealLoad,
    saveM365TokenCache: m365RealSave,
    getUnifiedRefreshTokenFromEnv: m365RealGetUnified,
    assertValidCacheIdentity: m365RealAssertValid
  }));
}

describe('resolveGraphAuth', () => {
  let originalEnv: NodeJS.ProcessEnv;
  let resolveGraphAuth: (options?: {
    token?: string;
    identity?: string;
    forceRefresh?: boolean;
    envPath?: string;
  }) => Promise<GraphAuthResult>;

  function applyJwtUtilsMock() {
    mock.module('../lib/jwt-utils.js', () => ({
      ...jwtUtilsReal,
      getMicrosoftTenantPathSegment: mock(() => 'common')
    }));
  }

  beforeEach(async () => {
    applyJwtUtilsMock();
    applyM365TokenCacheMockForGraph();
    const mod = await import(`../lib/graph-auth.js?graphAuthTest=${Date.now()}`);
    resolveGraphAuth = mod.resolveGraphAuth;
    originalEnv = { ...process.env };
    global.fetch = mockFetch as unknown as typeof fetch;
    mockLoad.mockReset();
    mockLoad.mockImplementation(() => Promise.resolve(null));
    mockSave.mockClear();
    mockFetch.mockClear();
  });

  afterEach(() => {
    process.env = originalEnv;
    // `mock.module` persists across files in Bun; `mock.restore()` then re-bind the real module (see oven-sh/bun#7823).
    mock.restore();
    restoreRealM365TokenCacheModule();
    mock.module('../lib/jwt-utils.js', () => jwtUtilsReal);
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

  function makeAccessTokenJwt(appid: string, scp?: string, tid?: string): string {
    const h = Buffer.from(JSON.stringify({ alg: 'none', typ: 'JWT' })).toString('base64url');
    const payload: Record<string, unknown> = { appid, exp: 2_000_000_000 };
    if (scp !== undefined) payload.scp = scp;
    if (tid !== undefined) payload.tid = tid;
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

  test('ignores cached Graph token when tenant does not match a pinned tenant GUID', async () => {
    // `getMicrosoftTenantPathSegment` is forced to a concrete tenant GUID for this test only, so the
    // new `isPinnedTenantGuid` gate engages (it's a no-op for the default 'common' used elsewhere).
    const pinnedTenant = '11111111-2222-4333-8444-555555555555';
    const staleTenant = '66666666-7777-8888-9999-000000000000';
    mock.module('../lib/jwt-utils.js', () => ({
      ...jwtUtilsReal,
      getMicrosoftTenantPathSegment: mock(() => pinnedTenant)
    }));
    const pinnedMod = await import(`../lib/graph-auth.js?graphAuthTenantTest=${Date.now()}`);
    const resolveGraphAuthPinned: typeof resolveGraphAuth = pinnedMod.resolveGraphAuth;

    process.env.EWS_CLIENT_ID = '5f2abcea-d6ea-4460-b468-3d80d7a900eb';
    process.env.M365_REFRESH_TOKEN = 'env-refresh';

    mockLoad.mockResolvedValue({
      version: 1,
      graph: {
        accessToken: makeAccessTokenJwt(
          '5f2abcea-d6ea-4460-b468-3d80d7a900eb',
          'Mail.ReadWrite Calendars.ReadWrite User.Read',
          staleTenant
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
              'Mail.Send Mail.ReadWrite Contacts.ReadWrite Notes.ReadWrite.All OnlineMeetings.ReadWrite User.Read',
              pinnedTenant
            ),
            refresh_token: 'rotated',
            expires_in: 3600
          }),
          { status: 200 }
        )
      )
    );

    const r = await resolveGraphAuthPinned();
    expect(r.success).toBe(true);
    expect(mockFetch).toHaveBeenCalled();
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

  test('persists rotated refresh token to caller-provided envPath (H-1/H-2)', async () => {
    // H-1/H-2: refreshes triggered by graph-auth should write the rotated refresh token
    // to the active env file the caller passed in (e.g. `--env-file .env.beta`),
    // NOT to the default global .env.
    const testHome = await mkdtemp(join(tmpdir(), 'm365-graph-auth-envpath-'));
    try {
      const envPath = join(testHome, '.env.beta');
      // The CLI is normally configured via `applyEnvFileOverrides(envPath)` first; here we
      // simulate that by not exporting M365_AGENT_ENV_FILE and letting envPath drive the write.
      // We deliberately do NOT set M365_AGENT_SKIP_GLOBAL_ENV or NODE_ENV=test so the persist
      // call actually writes to envPath.
      delete process.env.M365_AGENT_ENV_FILE;
      delete process.env.M365_AGENT_SKIP_GLOBAL_ENV;
      delete process.env.NODE_ENV;
      process.env.EWS_CLIENT_ID = 'beta-client';
      process.env.M365_REFRESH_TOKEN = 'beta-old-refresh';

      mockLoad.mockResolvedValue({
        version: 1,
        graph: { accessToken: 'expired', expiresAt: Date.now() - 60_000 }
      } as never);

      mockFetch.mockResolvedValue(
        new Response(
          JSON.stringify({
            access_token: makeAccessTokenJwt('beta-client', 'User.Read Mail.Read'),
            refresh_token: 'beta-rotated-refresh',
            expires_in: 3600
          }),
          { status: 200 }
        )
      );

      const r = await resolveGraphAuth({ envPath });
      expect(r.success).toBe(true);
      const written = await readFile(envPath, 'utf8');
      expect(written).toContain('M365_REFRESH_TOKEN=beta-rotated-refresh');
      expect(written).toContain('EWS_REFRESH_TOKEN=beta-rotated-refresh');
    } finally {
      await rm(testHome, { recursive: true, force: true }).catch(() => {});
    }
  });

  test('surfaces AADSTS / interaction_required error details on refresh failure (M-1)', async () => {
    process.env.EWS_CLIENT_ID = 'client';
    process.env.M365_REFRESH_TOKEN = 'expired-rt';

    mockLoad.mockResolvedValue({
      version: 1,
      graph: { accessToken: 'expired', expiresAt: Date.now() - 60_000 }
    } as never);

    // Realistic AADSTS500133 / interaction_required payload from Microsoft Entra:
    // user must complete MFA / conditional access before a refresh can succeed.
    // Graph refresh tries multiple scope candidates in a loop, so provide a fresh Response
    // on every call (avoids "Body already used" from a single mocked Response being re-read).
    mockFetch.mockImplementation(() =>
      Promise.resolve(
        new Response(
          JSON.stringify({
            error: 'interaction_required',
            error_description:
              "AADSTS500133: Assertion is not within its valid time range. The current time is 2026-07-02T02:20:00.000Z, the assertion's not-before time is 2026-07-02T02:25:00.000Z.",
            error_codes: [500133],
            timestamp: '2026-07-02T02:20:00Z',
            trace_id: 'abcd1234-5678-90ab-cdef-1234567890ab',
            correlation_id: 'fedcba98-7654-3210-fedc-ba9876543210',
            error_uri: 'https://login.microsoftonline.com/error?code=500133'
          }),
          { status: 400 }
        )
      )
    );

    const r = await resolveGraphAuth();
    expect(r.success).toBe(false);
    expect(r.error).toContain('Graph token refresh failed');
    // Detailed AADSTS / interaction_required context must reach the caller (sanitized).
    expect(r.error).toContain('AADSTS500133');
    expect(r.error).toContain('interaction_required');
    expect(r.error).toContain('re-authentication');
    // lastRefreshError is the structured field for tests / machine-readable output.
    expect(r.lastRefreshError).toBeDefined();
    expect(r.lastRefreshError).toContain('AADSTS500133');
    expect(r.lastRefreshError).toContain('interaction_required');
    // Secrets guard: the new refresh token MUST NOT appear in the surfaced error.
    expect(r.lastRefreshError ?? '').not.toContain('expired-rt');
  });

  afterAll(() => {
    mock.restore();
    restoreRealM365TokenCacheModule();
    mock.module('../lib/jwt-utils.js', () => jwtUtilsReal);
  });
});

/** EWS OAuth disk cache tests live in this file (after `resolveGraphAuth`) so Bun `mock.module` teardown stays scoped. */
function ewsFixtureAccessToken(seed: string): string {
  const h = Buffer.from(JSON.stringify({ alg: 'none', typ: 'JWT' })).toString('base64url');
  const p = Buffer.from(JSON.stringify({ exp: 2_000_000_000, sub: seed })).toString('base64url');
  return `${h}.${p}.x`;
}

function tokenCachePath(home: string, identity: string): string {
  return join(home, '.config', 'm365-agent-cli', `token-cache-${identity}.json`);
}

async function writePrimaryCache(home: string, identity: string, json: Record<string, unknown>): Promise<void> {
  const dir = join(home, '.config', 'm365-agent-cli');
  await mkdir(dir, { recursive: true });
  await writeFile(tokenCachePath(home, identity), JSON.stringify(json), 'utf8');
}

describe('auth resolution', () => {
  let originalEnv: NodeJS.ProcessEnv;
  let resolveAuth: (options?: { token?: string; identity?: string; envPath?: string }) => Promise<AuthResult>;
  let testHome: string;
  let cacheIdentity: string;
  let authFetchMock: ReturnType<typeof mock>;

  beforeEach(async () => {
    originalEnv = { ...process.env };
    testHome = await mkdtemp(join(tmpdir(), 'm365-auth-test-'));
    cacheIdentity = `t${randomBytes(12).toString('hex')}`;
    process.env.M365_AGENT_CLI_CONFIG_DIR = join(testHome, '.config', 'm365-agent-cli');
    process.env.EWS_TENANT_ID = 'common';

    restoreRealM365TokenCacheModule();
    mock.module('../lib/jwt-utils.js', () => jwtUtilsReal);

    authFetchMock = mock();
    const auth = await import(`../lib/auth.js?authDiskTest=${Date.now()}`);
    resolveAuth = auth.resolveAuth;
    global.fetch = authFetchMock as unknown as typeof fetch;
  });

  afterEach(async () => {
    for (const key of Object.keys(process.env)) {
      if (!(key in originalEnv)) {
        delete process.env[key];
      }
    }
    for (const [key, value] of Object.entries(originalEnv)) {
      if (value === undefined) {
        delete process.env[key];
      } else {
        process.env[key] = value;
      }
    }
    await rm(testHome, { recursive: true, force: true }).catch(() => {});
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
    const cachedTok = ewsFixtureAccessToken('cached');

    await writePrimaryCache(testHome, cacheIdentity, {
      version: 1,
      refreshToken: 'cached-refresh-token',
      ews: {
        accessToken: cachedTok,
        expiresAt: Date.now() + 1000_000
      }
    });

    const result = await resolveAuth({ identity: cacheIdentity });
    expect(result.success).toBe(true);
    expect(result.token).toBe(cachedTok);
    expect(authFetchMock).not.toHaveBeenCalled();
  });

  test('accepts legacy flat EWS cache shape', async () => {
    process.env.EWS_CLIENT_ID = 'client';
    process.env.EWS_REFRESH_TOKEN = 'refresh';
    const legacyTok = ewsFixtureAccessToken('legacy');

    await writePrimaryCache(testHome, cacheIdentity, {
      accessToken: legacyTok,
      refreshToken: 'cached-refresh-token',
      expiresAt: Date.now() + 1000_000
    });

    const result = await resolveAuth({ identity: cacheIdentity });
    expect(result.success).toBe(true);
    expect(result.token).toBe(legacyTok);
    expect(authFetchMock).not.toHaveBeenCalled();
  });

  test('M365_REFRESH_TOKEN satisfies auth without EWS_REFRESH_TOKEN', async () => {
    process.env.EWS_CLIENT_ID = 'client';
    delete process.env.EWS_REFRESH_TOKEN;
    process.env.M365_REFRESH_TOKEN = 'unified-refresh';
    const cachedTok = ewsFixtureAccessToken('m365');

    await writePrimaryCache(testHome, cacheIdentity, {
      version: 1,
      ews: {
        accessToken: cachedTok,
        expiresAt: Date.now() + 1000_000
      }
    });

    const result = await resolveAuth({ identity: cacheIdentity });
    expect(result.success).toBe(true);
    expect(result.token).toBe(cachedTok);
  });

  test('fetches new token if cache expired', async () => {
    process.env.EWS_CLIENT_ID = 'client';
    process.env.EWS_REFRESH_TOKEN = 'refresh';
    const expiredTok = ewsFixtureAccessToken('expired');
    const newTok = ewsFixtureAccessToken('new');

    await writePrimaryCache(testHome, cacheIdentity, {
      version: 1,
      refreshToken: 'cached-refresh-token',
      ews: {
        accessToken: expiredTok,
        expiresAt: Date.now() - 1000_000
      }
    });

    authFetchMock.mockResolvedValue(
      new Response(
        JSON.stringify({
          access_token: newTok,
          refresh_token: 'new-refresh-token',
          expires_in: 3600
        }),
        { status: 200 }
      )
    );

    const result = await resolveAuth({ identity: cacheIdentity });
    expect(result.success).toBe(true);
    expect(result.token).toBe(newTok);
    expect(authFetchMock).toHaveBeenCalled();
    const saved = JSON.parse(await readFile(tokenCachePath(testHome, cacheIdentity), 'utf8')) as {
      refreshToken?: string;
    };
    expect(saved.refreshToken).toBe('new-refresh-token');
  });

  test('persists rotated refresh token to .env on EWS refresh', async () => {
    const envPath = join(testHome, 'cli.env');
    process.env.M365_AGENT_ENV_FILE = envPath;
    process.env.EWS_CLIENT_ID = 'client';
    process.env.M365_REFRESH_TOKEN = 'env-refresh';
    await writeFile(envPath, 'EWS_CLIENT_ID=client\nM365_REFRESH_TOKEN=env-refresh\n', 'utf8');

    const expiredTok = ewsFixtureAccessToken('expired-env');
    const newTok = ewsFixtureAccessToken('new-env');

    await writePrimaryCache(testHome, cacheIdentity, {
      version: 1,
      refreshToken: 'cached-refresh-token',
      ews: {
        accessToken: expiredTok,
        expiresAt: Date.now() - 1000_000
      }
    });

    authFetchMock.mockResolvedValue(
      new Response(
        JSON.stringify({
          access_token: newTok,
          refresh_token: 'new-refresh-token',
          expires_in: 3600
        }),
        { status: 200 }
      )
    );

    delete process.env.NODE_ENV;

    const result = await resolveAuth({ identity: cacheIdentity });
    expect(result.success).toBe(true);
    const envAfter = await readFile(envPath, 'utf8');
    expect(envAfter).toContain('M365_REFRESH_TOKEN=new-refresh-token');
    expect(envAfter).toContain('GRAPH_REFRESH_TOKEN=new-refresh-token');
  });
});
