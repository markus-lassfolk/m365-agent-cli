/**
 * Browser-based OAuth authorization-code + PKCE login (`login --browser`, issue #244).
 *
 * Security properties:
 * - PKCE (S256) — no client secret involved (public client), and the authorization code is
 *   useless without the paired `code_verifier` this process never shares.
 * - The redirect listener binds to `127.0.0.1` only (never `0.0.0.0`), so nothing off-host can
 *   hit the callback endpoint.
 * - The local callback response HTML is a static string — request query params are never
 *   reflected into it, so a malicious redirect can't inject markup/script into the local page.
 * - Authorization codes and tokens are never logged — only returned to the caller in memory.
 * - The listener is time-bounded (`callbackTimeoutMs`) and closes itself on first request,
 *   success or failure, so it can't linger as an open local port.
 */
import { spawn } from 'node:child_process';
import type { IncomingMessage, ServerResponse } from 'node:http';
import { createServer } from 'node:http';
import { platform } from 'node:os';
import { getJwtExpiration, getJwtPayloadUpn, isValidJwtStructure } from './jwt-utils.js';
import { generateOAuthState, generatePkcePair } from './pkce.js';

export class BrowserLoginError extends Error {}

const SUCCESS_HTML =
  '<!doctype html><html><body style="font-family:sans-serif"><h2>Signed in.</h2><p>You can close this tab and return to the terminal.</p></body></html>';
const FAILURE_HTML =
  '<!doctype html><html><body style="font-family:sans-serif"><h2>Sign-in did not complete.</h2><p>Return to the terminal for details.</p></body></html>';

/**
 * Minimal structural subset of `node:http`'s `Server` that {@link runBrowserLogin} actually uses.
 * A real `http.Server` satisfies this automatically; tests can substitute a lightweight in-memory
 * fake (see `browser-login.test.ts`) so the PKCE/state-validation/token-exchange logic is verified
 * without binding a real OS socket — real-socket binding under heavy concurrent load (e.g. this
 * repo's full test suite) is inherently less deterministic than the logic it's testing.
 */
export interface LoopbackServerLike {
  listen(port: number, host: string, callback: () => void): void;
  address(): { port: number } | string | null;
  close(): void;
  on(event: 'request', handler: (req: IncomingMessage, res: ServerResponse) => void): void;
  on(event: 'error', handler: (err: Error) => void): void;
  once(event: 'error', handler: (err: Error) => void): void;
}

export interface BrowserLoginOptions {
  clientId: string;
  tenant: string;
  /** Space-separated OAuth scopes (same convention as device-code login). */
  scope: string;
  /** Loopback port to bind; 0 (default) picks a free ephemeral port. */
  port?: number;
  /** Launch the system browser automatically. Default true — pass false for `--no-open`. */
  open?: boolean;
  /** Max time to wait for the Microsoft redirect. Default 5 minutes. */
  callbackTimeoutMs?: number;
  /** Injectable for tests / alternate platforms; defaults to a real OS "open URL" call. */
  openBrowser?: (url: string) => void;
  /** Called once the authorization URL is known (before opening it) — printing hook for the CLI. */
  onAuthorizationUrl?: (url: string) => void;
  /** Injectable loopback listener (tests only) — defaults to a real `node:http` server. */
  _createLoopbackServer?: () => LoopbackServerLike;
}

export interface BrowserLoginResult {
  accessToken: string;
  refreshToken: string;
  expiresAt: number;
  signedInAs?: string;
}

function defaultOpenBrowser(url: string): void {
  try {
    const plat = platform();
    if (plat === 'darwin') {
      spawn('open', [url], { stdio: 'ignore', detached: true }).unref();
    } else if (plat === 'win32') {
      spawn('cmd', ['/c', 'start', '""', url], { stdio: 'ignore', detached: true, windowsHide: true }).unref();
    } else {
      spawn('xdg-open', [url], { stdio: 'ignore', detached: true }).unref();
    }
  } catch {
    // Best-effort only — the operator can always copy the printed URL manually (--no-open path).
  }
}

function waitForCallback(
  server: LoopbackServerLike,
  expectedState: string,
  timeoutMs: number
): Promise<{ code: string }> {
  return new Promise((resolve, reject) => {
    let settled = false;
    const timer = setTimeout(() => {
      if (settled) return;
      settled = true;
      server.close();
      reject(
        new BrowserLoginError(
          `Timed out after ${Math.round(timeoutMs / 1000)}s waiting for the Microsoft sign-in redirect.`
        )
      );
    }, timeoutMs);

    const finish = (fn: () => void) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      server.close();
      fn();
    };

    server.on('request', (req, res) => {
      let url: URL;
      try {
        url = new URL(req.url ?? '/', 'http://127.0.0.1');
      } catch {
        res.writeHead(400).end();
        return;
      }
      if (url.pathname !== '/callback') {
        res.writeHead(404).end();
        return;
      }

      const error = url.searchParams.get('error');
      const errorDescription = url.searchParams.get('error_description');
      const state = url.searchParams.get('state');
      const code = url.searchParams.get('code');

      if (error) {
        res.writeHead(200, { 'content-type': 'text/html' }).end(FAILURE_HTML);
        finish(() =>
          reject(
            new BrowserLoginError(
              `Microsoft sign-in failed: ${error}${errorDescription ? ` — ${errorDescription}` : ''}`
            )
          )
        );
        return;
      }
      if (!state || state !== expectedState) {
        res.writeHead(400, { 'content-type': 'text/html' }).end(FAILURE_HTML);
        finish(() => reject(new BrowserLoginError('OAuth state mismatch on callback — refusing to continue.')));
        return;
      }
      if (!code) {
        res.writeHead(400, { 'content-type': 'text/html' }).end(FAILURE_HTML);
        finish(() => reject(new BrowserLoginError('Microsoft redirect did not include an authorization code.')));
        return;
      }

      res.writeHead(200, { 'content-type': 'text/html' }).end(SUCCESS_HTML);
      finish(() => resolve({ code }));
    });

    server.on('error', (err) => {
      finish(() => reject(err instanceof Error ? err : new Error(String(err))));
    });
  });
}

/**
 * Run the full browser authorization-code + PKCE flow: bind a loopback listener, print/open the
 * authorization URL, wait for the redirect, then exchange the code for tokens. Throws
 * {@link BrowserLoginError} (never resolves with a partial/invalid result) on any failure —
 * timeout, user cancellation, state mismatch, or a token-endpoint error.
 */
export async function runBrowserLogin(options: BrowserLoginOptions): Promise<BrowserLoginResult> {
  const open = options.open ?? true;
  const timeoutMs = options.callbackTimeoutMs ?? 5 * 60_000;
  const pkce = generatePkcePair();
  const state = generateOAuthState();

  const server: LoopbackServerLike = (options._createLoopbackServer ?? (() => createServer()))();
  await new Promise<void>((resolve, reject) => {
    server.once('error', reject);
    // Loopback only — never bind 0.0.0.0, this listener must not be reachable off-host.
    server.listen(options.port ?? 0, '127.0.0.1', () => resolve());
  });

  const address = server.address();
  if (!address || typeof address === 'string') {
    server.close();
    throw new BrowserLoginError('Failed to bind the local loopback listener.');
  }
  const redirectUri = `http://127.0.0.1:${address.port}/callback`;

  const authUrl = new URL(`https://login.microsoftonline.com/${options.tenant}/oauth2/v2.0/authorize`);
  authUrl.searchParams.set('client_id', options.clientId);
  authUrl.searchParams.set('response_type', 'code');
  authUrl.searchParams.set('redirect_uri', redirectUri);
  authUrl.searchParams.set('response_mode', 'query');
  authUrl.searchParams.set('scope', options.scope);
  authUrl.searchParams.set('state', state);
  authUrl.searchParams.set('code_challenge', pkce.codeChallenge);
  authUrl.searchParams.set('code_challenge_method', pkce.codeChallengeMethod);

  options.onAuthorizationUrl?.(authUrl.toString());
  if (open) {
    (options.openBrowser ?? defaultOpenBrowser)(authUrl.toString());
  }

  const { code } = await waitForCallback(server, state, timeoutMs);

  const tokenRes = await fetch(`https://login.microsoftonline.com/${options.tenant}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: options.clientId,
      grant_type: 'authorization_code',
      code,
      redirect_uri: redirectUri,
      code_verifier: pkce.codeVerifier,
      scope: options.scope
    }).toString()
  });

  const json = (await tokenRes.json().catch(() => ({}))) as {
    access_token?: string;
    refresh_token?: string;
    expires_in?: number;
    error?: string;
    error_description?: string;
  };

  if (!tokenRes.ok || !json.access_token || !json.refresh_token) {
    const detail = [json.error, json.error_description].filter(Boolean).join(': ') || `HTTP ${tokenRes.status}`;
    throw new BrowserLoginError(`Token exchange failed: ${detail}`);
  }
  if (!isValidJwtStructure(json.access_token)) {
    throw new BrowserLoginError('OAuth server returned an invalid token structure — refusing to cache.');
  }

  const expiresAt = getJwtExpiration(json.access_token) ?? Date.now() + (json.expires_in || 3600) * 1000;
  if (expiresAt <= Date.now()) {
    throw new BrowserLoginError('OAuth server returned an already-expired token — refusing to cache.');
  }

  return {
    accessToken: json.access_token,
    refreshToken: json.refresh_token,
    expiresAt,
    signedInAs: getJwtPayloadUpn(json.access_token)
  };
}
