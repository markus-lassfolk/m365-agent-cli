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
    // The real loopback + PKCE round trip is covered exhaustively in `browser-login.test.ts`.
    // Here we inject a fake browser-login runner (via `_runBrowserLogin`) so the command wrapper's
    // env-persist + identity-binding wiring is tested deterministically, without standing up a real
    // HTTP server that can starve under full-suite `--isolate` load (see the module's injection doc).
    function fakeBrowserLogin(result: {
      accessToken?: string;
      refreshToken: string;
      signedInAs?: string;
    }): NonNullable<Parameters<typeof runBrowserLoginFlow>[0]['_runBrowserLogin']> {
      return (async (opts) => {
        opts.onAuthorizationUrl?.('https://login.microsoftonline.com/common/oauth2/v2.0/authorize?state=x');
        return {
          accessToken: result.accessToken ?? fixtureAccessToken(result.signedInAs ?? 'a@b.com'),
          refreshToken: result.refreshToken,
          expiresAt: Date.now() + 3_600_000,
          signedInAs: result.signedInAs
        };
      }) as NonNullable<Parameters<typeof runBrowserLoginFlow>[0]['_runBrowserLogin']>;
    }

    test('persists the refresh token and binds the identity on a successful browser login', async () => {
      const result = await runBrowserLoginFlow({
        envFile: envPath,
        identity: 'doris',
        open: false,
        _runBrowserLogin: fakeBrowserLogin({ refreshToken: 'browser-rt-1', signedInAs: 'doris@lassfolk.net' })
      });
      expect(result.signedInAs).toBe('doris@lassfolk.net');

      const written = await readFile(envPath, 'utf8');
      expect(written).toContain('M365_REFRESH_TOKEN=browser-rt-1');

      const profile = await getProfile('doris');
      expect(profile?.signedInAs).toBe('doris@lassfolk.net');
    });

    test('refuses (does not persist) when the browser login returns a different account than a verified identity', async () => {
      await runBrowserLoginFlow({
        envFile: envPath,
        identity: 'doris',
        open: false,
        _runBrowserLogin: fakeBrowserLogin({ refreshToken: 'browser-rt-1', signedInAs: 'doris@lassfolk.net' })
      });
      await expect(
        runBrowserLoginFlow({
          envFile: envPath,
          identity: 'doris',
          open: false,
          _runBrowserLogin: fakeBrowserLogin({ refreshToken: 'browser-rt-2', signedInAs: 'lotta@lassfolk.net' })
        })
      ).rejects.toThrow(LoginAccountMismatchError);
      // env still bound to the first account's token.
      const written = await readFile(envPath, 'utf8');
      expect(written).toContain('M365_REFRESH_TOKEN=browser-rt-1');
    });

    test('exits 1 with a clear message when the real browser login fails (e.g. timeout)', async () => {
      // Uses the REAL runBrowserLogin with a 30ms callback timeout — exercises the failure path
      // end-to-end without waiting on a redirect that never comes.
      const resultPromise = runBrowserLoginFlow({
        envFile: envPath,
        open: false,
        callbackTimeout: '30ms'
      });
      await expect(resultPromise).rejects.toThrow(/process\.exit/);
    });
  });
});
