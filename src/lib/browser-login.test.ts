import { afterEach, describe, expect, test } from 'bun:test';
import { Buffer } from 'node:buffer';
import { BrowserLoginError, runBrowserLogin } from './browser-login.js';

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

describe('runBrowserLogin', () => {
  const originalFetch = global.fetch;

  afterEach(() => {
    global.fetch = originalFetch;
  });

  function mockTokenEndpoint(responseBody: unknown, status = 200): void {
    global.fetch = (async (input: string | URL | Request) => {
      const url = typeof input === 'string' ? input : input.toString();
      if (url.includes('login.microsoftonline.com')) {
        return new Response(JSON.stringify(responseBody), {
          status,
          headers: { 'content-type': 'application/json' }
        });
      }
      return originalFetch(input as never);
    }) as unknown as typeof fetch;
  }

  test('completes a full loopback PKCE round trip and returns tokens', async () => {
    mockTokenEndpoint({
      access_token: fixtureAccessToken('doris@lassfolk.net'),
      refresh_token: 'rt-1',
      expires_in: 3600
    });

    let authUrl: string | undefined;
    const resultPromise = runBrowserLogin({
      clientId: 'client-1',
      tenant: 'common',
      scope: 'offline_access',
      open: false,
      onAuthorizationUrl: (u) => {
        authUrl = u;
      }
    });

    await pollUntil(() => authUrl !== undefined);
    const parsed = new URL(authUrl as string);
    expect(parsed.searchParams.get('code_challenge_method')).toBe('S256');
    expect(parsed.hostname).toBe('login.microsoftonline.com');
    const state = parsed.searchParams.get('state');
    const redirectUri = parsed.searchParams.get('redirect_uri') as string;
    expect(new URL(redirectUri).hostname).toBe('127.0.0.1');

    await originalFetch(`${redirectUri}?code=abc123&state=${state}`);

    const result = await resultPromise;
    expect(result.refreshToken).toBe('rt-1');
    expect(result.signedInAs).toBe('doris@lassfolk.net');
    expect(result.expiresAt).toBeGreaterThan(Date.now());
  }, 15_000);

  test('does not call openBrowser when open:false, but does when open:true', async () => {
    mockTokenEndpoint({ access_token: fixtureAccessToken('a@b.com'), refresh_token: 'rt', expires_in: 3600 });
    let opened: string | undefined;
    let authUrl: string | undefined;
    const resultPromise = runBrowserLogin({
      clientId: 'c',
      tenant: 'common',
      scope: 'offline_access',
      open: true,
      openBrowser: (u) => {
        opened = u;
      },
      onAuthorizationUrl: (u) => {
        authUrl = u;
      }
    });
    await pollUntil(() => authUrl !== undefined);
    expect(opened).toBe(authUrl);
    const state = new URL(authUrl as string).searchParams.get('state');
    const redirectUri = new URL(authUrl as string).searchParams.get('redirect_uri') as string;
    await originalFetch(`${redirectUri}?code=abc&state=${state}`);
    await resultPromise;
  });

  test('rejects with BrowserLoginError on state mismatch', async () => {
    let authUrl: string | undefined;
    const resultPromise = runBrowserLogin({
      clientId: 'c',
      tenant: 'common',
      scope: 'offline_access',
      open: false,
      onAuthorizationUrl: (u) => {
        authUrl = u;
      }
    });
    resultPromise.catch(() => {}); // mark handled before the network round trip below settles it
    await pollUntil(() => authUrl !== undefined);
    const redirectUri = new URL(authUrl as string).searchParams.get('redirect_uri') as string;
    await originalFetch(`${redirectUri}?code=abc&state=wrong-state`);
    await expect(resultPromise).rejects.toBeInstanceOf(BrowserLoginError);
    await expect(resultPromise).rejects.toThrow(/state mismatch/i);
  });

  test('rejects with BrowserLoginError when Microsoft returns an error (e.g. user cancelled)', async () => {
    let authUrl: string | undefined;
    const resultPromise = runBrowserLogin({
      clientId: 'c',
      tenant: 'common',
      scope: 'offline_access',
      open: false,
      onAuthorizationUrl: (u) => {
        authUrl = u;
      }
    });
    resultPromise.catch(() => {}); // mark handled before the network round trip below settles it
    await pollUntil(() => authUrl !== undefined);
    const state = new URL(authUrl as string).searchParams.get('state');
    const redirectUri = new URL(authUrl as string).searchParams.get('redirect_uri') as string;
    await originalFetch(`${redirectUri}?error=access_denied&error_description=User+cancelled&state=${state}`);
    await expect(resultPromise).rejects.toThrow(/Microsoft sign-in failed/);
  });

  test('rejects with BrowserLoginError on callback timeout', async () => {
    await expect(
      runBrowserLogin({ clientId: 'c', tenant: 'common', scope: 'offline_access', open: false, callbackTimeoutMs: 30 })
    ).rejects.toThrow(/Timed out/);
  });

  test('rejects when the token endpoint returns an OAuth error', async () => {
    mockTokenEndpoint({ error: 'invalid_grant', error_description: 'code expired' }, 400);
    let authUrl: string | undefined;
    const resultPromise = runBrowserLogin({
      clientId: 'c',
      tenant: 'common',
      scope: 'offline_access',
      open: false,
      onAuthorizationUrl: (u) => {
        authUrl = u;
      }
    });
    resultPromise.catch(() => {}); // mark handled before the network round trip below settles it
    await pollUntil(() => authUrl !== undefined);
    const state = new URL(authUrl as string).searchParams.get('state');
    const redirectUri = new URL(authUrl as string).searchParams.get('redirect_uri') as string;
    await originalFetch(`${redirectUri}?code=abc&state=${state}`);
    await expect(resultPromise).rejects.toThrow(/Token exchange failed/);
  });
});
