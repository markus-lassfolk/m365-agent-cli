import { afterEach, beforeEach, describe, expect, mock, test } from 'bun:test';
import { Buffer } from 'node:buffer';
import { randomBytes } from 'node:crypto';
import { mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import type { AuthResult } from '../lib/auth.js';

const mockFetch = mock();

/** Minimal three-part JWT so `isValidJwtStructure` passes with real `jwt-utils`. */
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
  /** Per-test identity avoids parallel tests clobbering the same default cache path via shared `HOME`. */
  let cacheIdentity: string;

  beforeEach(async () => {
    // Fresh `auth.js` per test so `loadM365TokenCache` tracks graph-auth `mock.module` teardown in `afterEach`.
    const auth = await import(`../lib/auth.js?authDiskTest=${Date.now()}`);
    resolveAuth = auth.resolveAuth;

    originalEnv = { ...process.env };
    testHome = await mkdtemp(join(tmpdir(), 'm365-auth-test-'));
    cacheIdentity = `t${randomBytes(12).toString('hex')}`;
    process.env.M365_AGENT_CLI_CONFIG_DIR = join(testHome, '.config', 'm365-agent-cli');
    process.env.EWS_TENANT_ID = 'common';
    global.fetch = mockFetch as unknown as typeof fetch;
    mockFetch.mockClear();
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
    expect(mockFetch).not.toHaveBeenCalled();
  });

  test('ignores cached EWS token when app id does not match EWS_CLIENT_ID', async () => {
    // Mirrors graph-auth.test.ts's "ignores cached Graph token when app id does not match
    // EWS_CLIENT_ID": the EWS path lacked this binding check and could silently serve a
    // token minted for a different Entra app registration after EWS_CLIENT_ID changes.
    process.env.EWS_CLIENT_ID = '5f2abcea-d6ea-4460-b468-3d80d7a900eb';
    process.env.M365_REFRESH_TOKEN = 'env-refresh';

    function ewsFixtureAccessTokenWithAppId(appid: string): string {
      const h = Buffer.from(JSON.stringify({ alg: 'none', typ: 'JWT' })).toString('base64url');
      const p = Buffer.from(JSON.stringify({ exp: 2_000_000_000, appid })).toString('base64url');
      return `${h}.${p}.x`;
    }

    await writePrimaryCache(testHome, cacheIdentity, {
      version: 1,
      refreshToken: 'cached-refresh-token',
      ews: {
        accessToken: ewsFixtureAccessTokenWithAppId('aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'),
        expiresAt: Date.now() + 1000_000
      }
    });

    const newTok = ewsFixtureAccessTokenWithAppId('5f2abcea-d6ea-4460-b468-3d80d7a900eb');
    mockFetch.mockResolvedValue(
      new Response(JSON.stringify({ access_token: newTok, refresh_token: 'rotated', expires_in: 3600 }), {
        status: 200
      })
    );

    const result = await resolveAuth({ identity: cacheIdentity });
    expect(result.success).toBe(true);
    expect(result.token).toBe(newTok);
    expect(mockFetch).toHaveBeenCalled();
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
    expect(mockFetch).not.toHaveBeenCalled();
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

    mockFetch.mockResolvedValue(
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
    expect(mockFetch).toHaveBeenCalled();
    const saved = JSON.parse(await readFile(tokenCachePath(testHome, cacheIdentity), 'utf8')) as {
      refreshToken?: string;
    };
    expect(saved.refreshToken).toBe('new-refresh-token');
  });

  test('surfaces AADSTS error details on EWS refresh failure (M-1)', async () => {
    process.env.EWS_CLIENT_ID = 'client';
    process.env.M365_REFRESH_TOKEN = 'expired-rt';

    await writePrimaryCache(testHome, cacheIdentity, {
      version: 1,
      ews: { accessToken: ewsFixtureAccessToken('expired'), expiresAt: Date.now() - 1000_000 }
    });

    // EWS refresh tries two scopes in a loop, so provide a fetch that returns a fresh
    // Response (with its own .json() body) on each call to avoid "Body already used".
    mockFetch.mockImplementation(() =>
      Promise.resolve(
        new Response(
          JSON.stringify({
            error: 'invalid_grant',
            error_description: 'AADSTS70000: Provided grant is invalid or malformed.'
          }),
          { status: 400 }
        )
      )
    );

    const result = await resolveAuth({ identity: cacheIdentity });
    expect(result.success).toBe(false);
    expect(result.error).toContain('Token refresh failed');
    expect(result.error).toContain('AADSTS70000');
    expect(result.lastRefreshError).toBeDefined();
    expect(result.lastRefreshError).toContain('AADSTS70000');
    // Secrets guard
    expect(result.lastRefreshError ?? '').not.toContain('expired-rt');
  });

  test('surfaces interaction_required / AADSTS500133 re-authentication hint on EWS refresh failure (M-1)', async () => {
    process.env.EWS_CLIENT_ID = 'client';
    process.env.M365_REFRESH_TOKEN = 'expired-rt';

    await writePrimaryCache(testHome, cacheIdentity, {
      version: 1,
      ews: { accessToken: ewsFixtureAccessToken('expired'), expiresAt: Date.now() - 1000_000 }
    });

    // EWS refresh tries two scopes in a loop, so provide a fresh Response on each call.
    mockFetch.mockImplementation(() =>
      Promise.resolve(
        new Response(
          JSON.stringify({
            error: 'interaction_required',
            error_description: 'AADSTS500133: Assertion is not within its valid time range.',
            error_codes: [500133]
          }),
          { status: 400 }
        )
      )
    );

    const result = await resolveAuth({ identity: cacheIdentity });
    expect(result.success).toBe(false);
    expect(result.error).toContain('AADSTS500133');
    expect(result.error).toContain('interaction_required');
    expect(result.error).toContain('re-authentication');
    expect(result.lastRefreshError).toContain('AADSTS500133');
    // Secrets guard
    expect(result.lastRefreshError ?? '').not.toContain('expired-rt');
  });

  test('persists rotated refresh token to caller-provided envPath (H-1/H-2 EWS)', async () => {
    const envHome = await mkdtemp(join(tmpdir(), 'm365-auth-envpath-'));
    try {
      const envPath = join(envHome, '.env.beta');
      // Use the explicit envPath to persist; clear M365_AGENT_ENV_FILE so getActiveEnvFilePath
      // returns the explicit one. We deliberately do NOT set M365_AGENT_SKIP_GLOBAL_ENV
      // (the production flag) so the write to envPath happens.
      delete process.env.M365_AGENT_ENV_FILE;
      delete process.env.M365_AGENT_SKIP_GLOBAL_ENV;
      delete process.env.NODE_ENV;
      process.env.EWS_CLIENT_ID = 'beta-client';
      process.env.M365_REFRESH_TOKEN = 'beta-old-refresh';

      await writePrimaryCache(testHome, cacheIdentity, {
        version: 1,
        ews: { accessToken: ewsFixtureAccessToken('expired'), expiresAt: Date.now() - 1000_000 }
      });

      mockFetch.mockResolvedValue(
        new Response(
          JSON.stringify({
            access_token: ewsFixtureAccessToken('new'),
            refresh_token: 'beta-rotated-refresh',
            expires_in: 3600
          }),
          { status: 200 }
        )
      );

      const result = await resolveAuth({ identity: cacheIdentity, envPath });
      expect(result.success).toBe(true);
      const written = await readFile(envPath, 'utf8');
      expect(written).toContain('M365_REFRESH_TOKEN=beta-rotated-refresh');
      expect(written).toContain('EWS_REFRESH_TOKEN=beta-rotated-refresh');
    } finally {
      await rm(envHome, { recursive: true, force: true }).catch(() => {});
    }
  });
});
