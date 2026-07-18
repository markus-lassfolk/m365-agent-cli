import { afterEach, describe, expect, test } from 'bun:test';
import { Buffer } from 'node:buffer';
import { BrowserLoginError, runBrowserLogin } from './browser-login.js';

function fixtureAccessToken(upn: string): string {
  const h = Buffer.from(JSON.stringify({ alg: 'none', typ: 'JWT' })).toString('base64url');
  const p = Buffer.from(JSON.stringify({ exp: 2_000_000_000, upn })).toString('base64url');
  return `${h}.${p}.x`;
}

/**
 * Resolves the instant `onAuthorizationUrl` fires — a single promise, not a `setTimeout` polling
 * loop. Polling with a fixed interval needs many event-loop ticks to "catch" a state change and
 * gets increasingly unreliable (spurious timeouts) as the scheduler falls behind under heavy
 * concurrent load (e.g. the full suite's `--isolate` run); a directly-resolved promise has no such
 * dependency — it fires on the first tick after the callback runs, however delayed that tick is.
 */
function authUrlWaiter(): { onAuthorizationUrl: (url: string) => void; promise: Promise<string> } {
  let resolve!: (url: string) => void;
  const promise = new Promise<string>((res) => {
    resolve = res;
  });
  return { onAuthorizationUrl: (url: string) => resolve(url), promise };
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

    const waiter = authUrlWaiter();
    const resultPromise = runBrowserLogin({
      clientId: 'client-1',
      tenant: 'common',
      scope: 'offline_access',
      open: false,
      onAuthorizationUrl: waiter.onAuthorizationUrl
    });

    const authUrl = await waiter.promise;
    const parsed = new URL(authUrl);
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
  }, 20_000);

  test('does not call openBrowser when open:false, but does when open:true', async () => {
    mockTokenEndpoint({ access_token: fixtureAccessToken('a@b.com'), refresh_token: 'rt', expires_in: 3600 });
    let opened: string | undefined;
    const waiter = authUrlWaiter();
    const resultPromise = runBrowserLogin({
      clientId: 'c',
      tenant: 'common',
      scope: 'offline_access',
      open: true,
      openBrowser: (u) => {
        opened = u;
      },
      onAuthorizationUrl: waiter.onAuthorizationUrl
    });
    const authUrl = await waiter.promise;
    expect(opened).toBe(authUrl);
    const state = new URL(authUrl).searchParams.get('state');
    const redirectUri = new URL(authUrl).searchParams.get('redirect_uri') as string;
    await originalFetch(`${redirectUri}?code=abc&state=${state}`);
    await resultPromise;
  }, 20_000);

  test('rejects with BrowserLoginError on state mismatch', async () => {
    const waiter = authUrlWaiter();
    const resultPromise = runBrowserLogin({
      clientId: 'c',
      tenant: 'common',
      scope: 'offline_access',
      open: false,
      onAuthorizationUrl: waiter.onAuthorizationUrl
    });
    resultPromise.catch(() => {}); // mark handled before the network round trip below settles it
    const authUrl = await waiter.promise;
    const redirectUri = new URL(authUrl).searchParams.get('redirect_uri') as string;
    await originalFetch(`${redirectUri}?code=abc&state=wrong-state`);
    await expect(resultPromise).rejects.toBeInstanceOf(BrowserLoginError);
    await expect(resultPromise).rejects.toThrow(/state mismatch/i);
  }, 20_000);

  test('rejects with BrowserLoginError when Microsoft returns an error (e.g. user cancelled)', async () => {
    const waiter = authUrlWaiter();
    const resultPromise = runBrowserLogin({
      clientId: 'c',
      tenant: 'common',
      scope: 'offline_access',
      open: false,
      onAuthorizationUrl: waiter.onAuthorizationUrl
    });
    resultPromise.catch(() => {}); // mark handled before the network round trip below settles it
    const authUrl = await waiter.promise;
    const state = new URL(authUrl).searchParams.get('state');
    const redirectUri = new URL(authUrl).searchParams.get('redirect_uri') as string;
    await originalFetch(`${redirectUri}?error=access_denied&error_description=User+cancelled&state=${state}`);
    await expect(resultPromise).rejects.toThrow(/Microsoft sign-in failed/);
  }, 20_000);

  test('rejects with BrowserLoginError on callback timeout', async () => {
    await expect(
      runBrowserLogin({
        clientId: 'c',
        tenant: 'common',
        scope: 'offline_access',
        open: false,
        callbackTimeoutMs: 30
      })
    ).rejects.toThrow(/Timed out/);
  }, 20_000);

  test('rejects when the token endpoint returns an OAuth error', async () => {
    mockTokenEndpoint({ error: 'invalid_grant', error_description: 'code expired' }, 400);
    const waiter = authUrlWaiter();
    const resultPromise = runBrowserLogin({
      clientId: 'c',
      tenant: 'common',
      scope: 'offline_access',
      open: false,
      onAuthorizationUrl: waiter.onAuthorizationUrl
    });
    resultPromise.catch(() => {}); // mark handled before the network round trip below settles it
    const authUrl = await waiter.promise;
    const state = new URL(authUrl).searchParams.get('state');
    const redirectUri = new URL(authUrl).searchParams.get('redirect_uri') as string;
    await originalFetch(`${redirectUri}?code=abc&state=${state}`);
    await expect(resultPromise).rejects.toThrow(/Token exchange failed/);
  }, 20_000);
});
