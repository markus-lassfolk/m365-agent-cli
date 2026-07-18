import { afterEach, beforeEach, describe, expect, test } from 'bun:test';
import { Buffer } from 'node:buffer';
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { getProfile } from '../lib/identity-profiles.js';
import { LoginAccountMismatchError } from '../lib/login-identity-binding.js';
import { runBrowserLoginFlow, runDeviceCodeLogin } from './login.js';

function fixtureAccessToken(upn: string): string {
  const h = Buffer.from(JSON.stringify({ alg: 'none', typ: 'JWT' })).toString('base64url');
  const p = Buffer.from(JSON.stringify({ exp: 2_000_000_000, upn })).toString('base64url');
  return `${h}.${p}.x`;
}

async function pollUntil(predicate: () => boolean, timeoutMs = 10_000): Promise<void> {
  const start = Date.now();
  while (!predicate()) {
    if (Date.now() - start > timeoutMs) throw new Error('pollUntil timed out');
    await new Promise((r) => setTimeout(r, 5));
  }
}

describe('login command flows', () => {
  let testHome: string;
  let envPath: string;
  let originalEnv: NodeJS.ProcessEnv;
  const originalFetch = global.fetch;
  const originalLog = console.log;
  const originalExit = process.exit;
  let logs: string[];

  beforeEach(async () => {
    testHome = await mkdtemp(join(tmpdir(), 'm365-login-cmd-'));
    envPath = join(testHome, '.env');
    originalEnv = { ...process.env };
    process.env.M365_AGENT_CLI_CONFIG_DIR = join(testHome, '.config', 'm365-agent-cli');
    // Opt back into real env-file persistence for these tests (production behavior; the test
    // runner sets NODE_ENV=test globally, which normally short-circuits persistRefreshTokenToEnv).
    delete process.env.M365_AGENT_ENV_FILE;
    delete process.env.M365_AGENT_SKIP_GLOBAL_ENV;
    delete process.env.NODE_ENV;
    process.env.EWS_CLIENT_ID = 'test-client-id';
    process.env.EWS_TENANT_ID = 'common';

    logs = [];
    console.log = ((...args: unknown[]) => {
      logs.push(args.map(String).join(' '));
    }) as typeof console.log;
    process.exit = ((code?: number) => {
      throw new Error(`process.exit(${code})`);
    }) as never;
  });

  afterEach(async () => {
    global.fetch = originalFetch;
    console.log = originalLog;
    process.exit = originalExit;
    for (const key of Object.keys(process.env)) {
      if (!(key in originalEnv)) delete process.env[key];
    }
    for (const [key, value] of Object.entries(originalEnv)) {
      if (value === undefined) delete process.env[key];
      else process.env[key] = value;
    }
    await rm(testHome, { recursive: true, force: true }).catch(() => {});
  });

  describe('runDeviceCodeLogin', () => {
    function mockDeviceCodeFlow(upn: string): void {
      let step = 0;
      global.fetch = (async (input: string | URL | Request) => {
        const url = typeof input === 'string' ? input : input.toString();
        if (url.includes('/devicecode')) {
          return new Response(
            JSON.stringify({
              device_code: 'dc-1',
              interval: 0.001,
              expires_in: 900,
              message: 'go to microsoft.com/devicelogin'
            }),
            { status: 200, headers: { 'content-type': 'application/json' } }
          );
        }
        if (url.includes('/oauth2/v2.0/token')) {
          step++;
          return new Response(
            JSON.stringify({ access_token: fixtureAccessToken(upn), refresh_token: `rt-${step}`, expires_in: 3600 }),
            { status: 200, headers: { 'content-type': 'application/json' } }
          );
        }
        return originalFetch(input as never);
      }) as unknown as typeof fetch;
    }

    test('persists the refresh token and binds the identity profile on success', async () => {
      mockDeviceCodeFlow('doris@lassfolk.net');
      const result = await runDeviceCodeLogin({ envFile: envPath, identity: 'doris' });
      expect(result.signedInAs).toBe('doris@lassfolk.net');

      const written = await readFile(envPath, 'utf8');
      expect(written).toContain('M365_REFRESH_TOKEN=rt-1');
      expect(written).toContain('EWS_USERNAME=doris@lassfolk.net');

      const profile = await getProfile('doris');
      expect(profile?.signedInAs).toBe('doris@lassfolk.net');
    });

    test('works without --identity (no profile binding, unchanged pre-existing behavior)', async () => {
      mockDeviceCodeFlow('doris@lassfolk.net');
      await runDeviceCodeLogin({ envFile: envPath });
      const written = await readFile(envPath, 'utf8');
      expect(written).toContain('M365_REFRESH_TOKEN=rt-1');
    });

    test('refuses to complete and does not persist anything when the account differs from a previously verified identity', async () => {
      mockDeviceCodeFlow('doris@lassfolk.net');
      await runDeviceCodeLogin({ envFile: envPath, identity: 'doris' });

      mockDeviceCodeFlow('lotta@lassfolk.net');
      await expect(runDeviceCodeLogin({ envFile: envPath, identity: 'doris' })).rejects.toThrow(
        LoginAccountMismatchError
      );

      // Still bound to the original account — not silently switched.
      const profile = await getProfile('doris');
      expect(profile?.signedInAs).toBe('doris@lassfolk.net');
      // The second (rejected) login's refresh token must not have overwritten the env file.
      const written = await readFile(envPath, 'utf8');
      expect(written).toContain('M365_REFRESH_TOKEN=rt-1');
    });

    test('--force-identity-switch allows completing a login into a different account', async () => {
      mockDeviceCodeFlow('doris@lassfolk.net');
      await runDeviceCodeLogin({ envFile: envPath, identity: 'doris' });

      mockDeviceCodeFlow('lotta@lassfolk.net');
      await runDeviceCodeLogin({ envFile: envPath, identity: 'doris', forceIdentitySwitch: true });

      const profile = await getProfile('doris');
      expect(profile?.signedInAs).toBe('lotta@lassfolk.net');
    });
  });

  describe('runBrowserLoginFlow', () => {
    function mockTokenEndpoint(upn: string): void {
      global.fetch = (async (input: string | URL | Request) => {
        const url = typeof input === 'string' ? input : input.toString();
        if (url.includes('login.microsoftonline.com')) {
          return new Response(
            JSON.stringify({ access_token: fixtureAccessToken(upn), refresh_token: 'browser-rt-1', expires_in: 3600 }),
            { status: 200, headers: { 'content-type': 'application/json' } }
          );
        }
        return originalFetch(input as never);
      }) as unknown as typeof fetch;
    }

    function extractPrintedAuthUrl(): URL | undefined {
      for (const line of logs) {
        if (line.startsWith('https://login.microsoftonline.com/')) {
          return new URL(line);
        }
      }
      return undefined;
    }

    test('completes the loopback round trip, persists the refresh token, and binds the identity', async () => {
      mockTokenEndpoint('doris@lassfolk.net');
      const resultPromise = runBrowserLoginFlow({ envFile: envPath, identity: 'doris', open: false });

      let authUrl: URL | undefined;
      await pollUntil(() => {
        authUrl = extractPrintedAuthUrl();
        return authUrl !== undefined;
      });
      const state = authUrl?.searchParams.get('state');
      const redirectUri = authUrl?.searchParams.get('redirect_uri') as string;
      await originalFetch(`${redirectUri}?code=abc123&state=${state}`);

      const result = await resultPromise;
      expect(result.signedInAs).toBe('doris@lassfolk.net');

      const written = await readFile(envPath, 'utf8');
      expect(written).toContain('M365_REFRESH_TOKEN=browser-rt-1');

      const profile = await getProfile('doris');
      expect(profile?.signedInAs).toBe('doris@lassfolk.net');
    }, 15_000);

    test('exits 1 with a clear message when the browser login fails (e.g. timeout)', async () => {
      const resultPromise = runBrowserLoginFlow({
        envFile: envPath,
        open: false,
        callbackTimeout: '30ms'
      });
      await expect(resultPromise).rejects.toThrow(/process\.exit/);
    });
  });
});
