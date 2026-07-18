import { afterEach, describe, expect, test } from 'bun:test';
import { Buffer } from 'node:buffer';
import { EventEmitter } from 'node:events';
import type { IncomingMessage, ServerResponse } from 'node:http';
import { BrowserLoginError, type LoopbackServerLike, runBrowserLogin } from './browser-login.js';

function fixtureAccessToken(upn: string): string {
  const h = Buffer.from(JSON.stringify({ alg: 'none', typ: 'JWT' })).toString('base64url');
  const p = Buffer.from(JSON.stringify({ exp: 2_000_000_000, upn })).toString('base64url');
  return `${h}.${p}.x`;
}

/**
 * In-memory stand-in for `node:http`'s `Server`, driven directly by tests instead of a real OS
 * socket. `runBrowserLogin`'s PKCE/state-validation/token-exchange logic only needs the small
 * `LoopbackServerLike` surface (`listen`/`address`/`close`/`on`) — it never inspects the transport
 * itself — so a fake that emits synthetic `request` events exercises exactly the same code paths
 * as a real HTTP round trip, deterministically and without depending on real socket binding
 * succeeding promptly under heavy concurrent load (this repo's full `--isolate` suite spans 97
 * files; real-socket tests were measurably less reliable there than the logic warrants).
 */
class FakeLoopbackServer extends EventEmitter implements LoopbackServerLike {
  closed = false;

  listen(_port: number, _host: string, callback: () => void): void {
    queueMicrotask(callback);
  }

  address(): { port: number } {
    return { port: 44444 };
  }

  close(): void {
    this.closed = true;
  }

  /** Simulate the browser's redirect landing on `/callback?...`. */
  fireCallback(query: string): void {
    const req = { url: `/callback${query}` } as IncomingMessage;
    const res = {
      writeHead(_status: number, _headers?: Record<string, string>) {
        return res;
      },
      end(_body?: string) {
        return res;
      }
    } as unknown as ServerResponse;
    this.emit('request', req, res);
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

  test('completes a full PKCE round trip and returns tokens', async () => {
    mockTokenEndpoint({
      access_token: fixtureAccessToken('doris@lassfolk.net'),
      refresh_token: 'rt-1',
      expires_in: 3600
    });

    const fakeServer = new FakeLoopbackServer();
    let authUrl: string | undefined;
    const resultPromise = runBrowserLogin({
      clientId: 'client-1',
      tenant: 'common',
      scope: 'offline_access',
      open: false,
      _createLoopbackServer: () => fakeServer,
      onAuthorizationUrl: (u) => {
        authUrl = u;
      }
    });

    // onAuthorizationUrl fires synchronously once `listen()`'s queued callback runs.
    await new Promise((r) => queueMicrotask(r as () => void));
    const parsed = new URL(authUrl as string);
    expect(parsed.searchParams.get('code_challenge_method')).toBe('S256');
    expect(parsed.hostname).toBe('login.microsoftonline.com');
    const state = parsed.searchParams.get('state');
    const redirectUri = parsed.searchParams.get('redirect_uri') as string;
    expect(new URL(redirectUri).hostname).toBe('127.0.0.1');

    fakeServer.fireCallback(`?code=abc123&state=${state}`);

    const result = await resultPromise;
    expect(result.refreshToken).toBe('rt-1');
    expect(result.signedInAs).toBe('doris@lassfolk.net');
    expect(result.expiresAt).toBeGreaterThan(Date.now());
    expect(fakeServer.closed).toBe(true);
  });

  test('does not call openBrowser when open:false, but does when open:true', async () => {
    mockTokenEndpoint({ access_token: fixtureAccessToken('a@b.com'), refresh_token: 'rt', expires_in: 3600 });
    const fakeServer = new FakeLoopbackServer();
    let opened: string | undefined;
    let authUrl: string | undefined;
    const resultPromise = runBrowserLogin({
      clientId: 'c',
      tenant: 'common',
      scope: 'offline_access',
      open: true,
      _createLoopbackServer: () => fakeServer,
      openBrowser: (u) => {
        opened = u;
      },
      onAuthorizationUrl: (u) => {
        authUrl = u;
      }
    });
    await new Promise((r) => queueMicrotask(r as () => void));
    expect(opened).toBe(authUrl);
    const state = new URL(authUrl as string).searchParams.get('state');
    fakeServer.fireCallback(`?code=abc&state=${state}`);
    await resultPromise;
  });

  test('rejects with BrowserLoginError on state mismatch', async () => {
    const fakeServer = new FakeLoopbackServer();
    const resultPromise = runBrowserLogin({
      clientId: 'c',
      tenant: 'common',
      scope: 'offline_access',
      open: false,
      _createLoopbackServer: () => fakeServer
    });
    resultPromise.catch(() => {}); // mark handled before the synthetic callback below settles it
    await new Promise((r) => queueMicrotask(r as () => void));
    fakeServer.fireCallback('?code=abc&state=wrong-state');
    await expect(resultPromise).rejects.toBeInstanceOf(BrowserLoginError);
    await expect(resultPromise).rejects.toThrow(/state mismatch/i);
  });

  test('rejects with BrowserLoginError when Microsoft returns an error (e.g. user cancelled)', async () => {
    const fakeServer = new FakeLoopbackServer();
    let authUrl: string | undefined;
    const resultPromise = runBrowserLogin({
      clientId: 'c',
      tenant: 'common',
      scope: 'offline_access',
      open: false,
      _createLoopbackServer: () => fakeServer,
      onAuthorizationUrl: (u) => {
        authUrl = u;
      }
    });
    resultPromise.catch(() => {}); // mark handled before the synthetic callback below settles it
    await new Promise((r) => queueMicrotask(r as () => void));
    const state = new URL(authUrl as string).searchParams.get('state');
    fakeServer.fireCallback(`?error=access_denied&error_description=User+cancelled&state=${state}`);
    await expect(resultPromise).rejects.toThrow(/Microsoft sign-in failed/);
  });

  test('rejects with BrowserLoginError on callback timeout', async () => {
    const fakeServer = new FakeLoopbackServer();
    await expect(
      runBrowserLogin({
        clientId: 'c',
        tenant: 'common',
        scope: 'offline_access',
        open: false,
        callbackTimeoutMs: 30,
        _createLoopbackServer: () => fakeServer
      })
    ).rejects.toThrow(/Timed out/);
  });

  test('rejects when the token endpoint returns an OAuth error', async () => {
    mockTokenEndpoint({ error: 'invalid_grant', error_description: 'code expired' }, 400);
    const fakeServer = new FakeLoopbackServer();
    let authUrl: string | undefined;
    const resultPromise = runBrowserLogin({
      clientId: 'c',
      tenant: 'common',
      scope: 'offline_access',
      open: false,
      _createLoopbackServer: () => fakeServer,
      onAuthorizationUrl: (u) => {
        authUrl = u;
      }
    });
    resultPromise.catch(() => {}); // mark handled before the synthetic callback below settles it
    await new Promise((r) => queueMicrotask(r as () => void));
    const state = new URL(authUrl as string).searchParams.get('state');
    fakeServer.fireCallback(`?code=abc&state=${state}`);
    await expect(resultPromise).rejects.toThrow(/Token exchange failed/);
  });
});
