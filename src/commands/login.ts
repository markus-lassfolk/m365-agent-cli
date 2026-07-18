import { chmodSync, existsSync, mkdirSync } from 'node:fs';
import { readFile } from 'node:fs/promises';
import { dirname } from 'node:path';
import { createInterface } from 'node:readline/promises';
import { Command } from 'commander';
import { atomicWriteUtf8File } from '../lib/atomic-write.js';
import { type BrowserLoginResult, runBrowserLogin } from '../lib/browser-login.js';
import { persistRefreshTokenToEnv } from '../lib/env-persist.js';
import { parseCacheTtlMs } from '../lib/graph-cache.js';
import { GRAPH_DEVICE_CODE_LOGIN_SCOPES } from '../lib/graph-oauth-scopes.js';
import { getJwtPayloadUpn, getMicrosoftTenantPathSegment, isValidJwtStructure } from '../lib/jwt-utils.js';
import {
  assertLoginIdentityOrThrow,
  commitLoginIdentity,
  LoginAccountMismatchError
} from '../lib/login-identity-binding.js';
import { applyEnvFileOverrides, getGlobalEnvFilePath, resolveEnvFilePathArgument } from '../lib/utils.js';

/**
 * In `--json` mode the command writes newline-delimited JSON events to stdout so an unattended
 * wrapper can parse them without scraping free-form log lines. No-op when JSON mode is off.
 * See docs/UNATTENDED_LOGIN.md for the device-code event shapes and an end-to-end automation example.
 */
function emitEvent(json: boolean, event: Record<string, unknown>): void {
  if (json) {
    process.stdout.write(`${JSON.stringify(event)}\n`);
  }
}

/**
 * Emit a terminal JSON error event (in `--json` mode) and exit. Unlike emitEvent + process.exit,
 * this awaits the stdout write callback first: on POSIX a piped stdout is asynchronous, so a bare
 * exit could truncate the final event before it flushes. A short fallback timer still exits if the
 * write callback never fires (e.g. the reader closed the pipe → EPIPE).
 */
async function fatalJson(json: boolean, event: Record<string, unknown>, code = 1): Promise<never> {
  if (json) {
    await new Promise<void>((resolve) => {
      let settled = false;
      const finish = (): void => {
        if (settled) return;
        settled = true;
        resolve();
      };
      const timer = setTimeout(finish, 1000);
      if (typeof timer.unref === 'function') timer.unref();
      process.stdout.write(`${JSON.stringify(event)}\n`, finish);
    });
  }
  process.exit(code);
}

/**
 * Route human-readable text to stderr in JSON mode (so stdout stays a clean event stream), else to
 * stdout. Single source of truth for the routing rule used across the command.
 */
function makeHumanLog(json: boolean): (msg: string) => void {
  return (msg: string): void => {
    if (json) console.error(msg);
    else console.log(msg);
  };
}

interface DeviceCodeFlowResult {
  refreshToken: string;
  signedInAs?: string;
}

async function performDeviceCodeFlow(
  clientId: string,
  tenant: string,
  scope: string,
  label: string,
  envPath: string,
  json: boolean
): Promise<DeviceCodeFlowResult> {
  // Human-readable text goes to stderr in JSON mode so stdout stays a clean JSON event stream.
  const humanLog = makeHumanLog(json);

  humanLog(`\nInitiating Device Code flow for ${label}...`);

  const deviceCodeRes = await fetch(`https://login.microsoftonline.com/${tenant}/oauth2/v2.0/devicecode`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: clientId,
      scope: scope
    }).toString()
  });

  const deviceCodeJson = await deviceCodeRes.json();

  if (!deviceCodeRes.ok) {
    console.error(`Failed to initiate ${label} device code flow:`, deviceCodeJson);
    await fatalJson(json, {
      event: 'error',
      error: deviceCodeJson.error ?? 'devicecode_request_failed',
      error_description: deviceCodeJson.error_description
    });
  }

  humanLog('\n=========================================================');
  humanLog(deviceCodeJson.message);
  humanLog('=========================================================\n');

  // Machine-readable device-code details for unattended automation (see docs/UNATTENDED_LOGIN.md).
  emitEvent(json, {
    event: 'device_code',
    user_code: deviceCodeJson.user_code,
    verification_uri: deviceCodeJson.verification_uri,
    verification_uri_complete: deviceCodeJson.verification_uri_complete ?? null,
    expires_in: deviceCodeJson.expires_in,
    interval: deviceCodeJson.interval,
    message: deviceCodeJson.message
  });

  const deviceCode = deviceCodeJson.device_code;
  const interval = (deviceCodeJson.interval || 5) * 1000;
  const expiresAt = Date.now() + (deviceCodeJson.expires_in || 900) * 1000;

  let authenticated = false;
  let refreshToken = '';
  let signedInAs: string | undefined;
  let pollInterval = interval;

  humanLog(`Waiting for ${label} authentication...`);

  while (!authenticated) {
    if (Date.now() > expiresAt) {
      console.error(`\n${label} device code expired. Please run the command again.`);
      await fatalJson(json, {
        event: 'error',
        error: 'expired_token',
        error_description: 'Device code expired before sign-in completed.'
      });
    }

    await new Promise((resolve) => setTimeout(resolve, pollInterval));

    const tokenRes = await fetch(`https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
        client_id: clientId,
        device_code: deviceCode
      }).toString()
    });

    const tokenJson = await tokenRes.json();

    if (tokenRes.ok) {
      authenticated = true;
      refreshToken = tokenJson.refresh_token;
      // Extract username from access token

      try {
        if (isValidJwtStructure(tokenJson.access_token)) {
          // getJwtPayloadUpn covers upn/preferred_username/email (in that precedence) and already
          // strips embedded CR/LF — the same claim-order and sanitization every other
          // identity-guard consumer (auth-diagnostics.ts, identity-guard.ts, browser-login.ts) uses,
          // so `login` and `readiness`/`auth repair` never disagree about the signed-in UPN for the
          // same token.
          const username = getJwtPayloadUpn(tokenJson.access_token);
          signedInAs = username;

          if (username) {
            let envContent = '';

            mkdirSync(dirname(envPath), { recursive: true, mode: 0o700 });
            chmodSync(dirname(envPath), 0o700);

            try {
              envContent = await readFile(envPath, 'utf8');
            } catch (err: any) {
              if (err.code !== 'ENOENT') throw err;
            }

            if (/^EWS_USERNAME=.*$/m.test(envContent)) {
              envContent = envContent.replace(/^EWS_USERNAME=.*$/m, () => `EWS_USERNAME=${username}`);
            } else {
              envContent += `\nEWS_USERNAME=${username}\n`;
            }

            await atomicWriteUtf8File(envPath, `${envContent.trim()}\n`, 0o600);

            humanLog(`Saved EWS_USERNAME (${username}) to ${envPath}`);
          }
        }
      } catch (_e) {
        /* ignore parse errors */
      }
      if (!refreshToken) {
        console.error(`\nFailed to obtain ${label} refresh token. Ensure the offline_access scope is granted.`);
        await fatalJson(json, {
          event: 'error',
          error: 'no_refresh_token',
          error_description: 'No refresh token returned; ensure the offline_access scope is granted.'
        });
      }
    } else if (tokenJson.error === 'authorization_pending') {
      // Continue polling
    } else if (tokenJson.error === 'slow_down') {
      pollInterval += 5000;
    } else {
      console.error(`\n${label} authentication failed:`, tokenJson.error_description || tokenJson.error);
      await fatalJson(json, {
        event: 'error',
        error: tokenJson.error ?? 'authentication_failed',
        error_description: tokenJson.error_description
      });
    }
  }

  humanLog(`\n${label} authentication successful!`);
  emitEvent(json, { event: 'authenticated', username: signedInAs ?? null });

  return { refreshToken, signedInAs };
}

/**
 * Resolve `--env-file`/EWS_CLIENT_ID, prompting for the client id on first run. Shared by both
 * login flows. In `--json` mode, a missing client id is a hard (non-interactive) error rather than
 * a stdin prompt — an unattended wrapper has nothing to answer it with.
 */
async function resolveLoginEnvAndClientId(
  envFile: string | undefined,
  json: boolean
): Promise<{ envPath: string; clientId: string }> {
  const humanLog = makeHumanLog(json);
  let envPath = getGlobalEnvFilePath();
  if (envFile) {
    envPath = resolveEnvFilePathArgument(envFile);
    applyEnvFileOverrides(envPath);
  }

  let clientId = process.env.EWS_CLIENT_ID;
  mkdirSync(dirname(envPath), { recursive: true, mode: 0o700 });
  chmodSync(dirname(envPath), 0o700);
  let envContent = '';
  if (existsSync(envPath)) {
    envContent = await readFile(envPath, 'utf8');
    if (!clientId) {
      const match = envContent.match(/^EWS_CLIENT_ID=(.*)$/m);
      if (match) {
        clientId = match[1].trim();
      }
    }
  }

  if (!clientId) {
    if (json) {
      // Non-interactive mode must not block on stdin waiting for a client id.
      await fatalJson(json, {
        event: 'error',
        error: 'missing_client_id',
        error_description: 'Set EWS_CLIENT_ID in the environment or .env before using --json (non-interactive) mode.'
      });
    }

    const rl = createInterface({ input: process.stdin, output: process.stdout });
    clientId = await rl.question('Enter your EWS_CLIENT_ID: ');
    rl.close();
    clientId = clientId.trim();

    if (!clientId) {
      console.error('EWS_CLIENT_ID is required.');
      process.exit(1);
    }

    envContent += `\nEWS_CLIENT_ID=${clientId}\n`;
    await atomicWriteUtf8File(envPath, `${envContent.trim()}\n`, 0o600);
  }

  humanLog('');
  humanLog(`Configuration file: ${envPath}`);
  if (!process.env.M365_AGENT_ENV_FILE?.trim() && !envFile) {
    humanLog(
      'Tip: For a second app (e.g. beta), use --env-file ~/.config/m365-agent-cli/.env.beta or export M365_AGENT_ENV_FILE to that path before running the CLI.'
    );
  }
  humanLog(`Application (client) ID: ${clientId}`);
  humanLog('');

  return { envPath, clientId };
}

export interface RunLoginOptions {
  envFile?: string;
  identity?: string;
  forceIdentitySwitch?: boolean;
  /** Emit newline-delimited JSON events instead of human text (see docs/UNATTENDED_LOGIN.md). */
  json?: boolean;
}

export interface RunLoginResult {
  envPath: string;
  signedInAs?: string;
}

/** Device-code OAuth login (`m365-agent-cli login`). Exported so `auth repair --start-login` can reuse it. */
export async function runDeviceCodeLogin(options: RunLoginOptions): Promise<RunLoginResult> {
  const json = options.json ?? false;
  const humanLog = makeHumanLog(json);
  const { envPath, clientId } = await resolveLoginEnvAndClientId(options.envFile, json);
  const tenant = getMicrosoftTenantPathSegment();

  const { refreshToken: rawToken, signedInAs } = await performDeviceCodeFlow(
    clientId,
    tenant,
    GRAPH_DEVICE_CODE_LOGIN_SCOPES,
    'Microsoft 365',
    envPath,
    json
  );
  const refreshToken = rawToken.replace(/[\r\n]/g, '');

  // Two-phase: assert (no writes) BEFORE persisting anything, so a refused login (wrong-account
  // mismatch) leaves the env file untouched — "refuse to complete" means no local state changes.
  // commit (the actual profiles.json write) only happens AFTER the refresh token is durably
  // persisted, so a later unrelated failure (disk full, read-only env path) can never leave
  // profiles.json falsely claiming a fresh, verified login with no usable token behind it. Left as
  // a throw (not fatalJson) so callers other than the `login` command itself — e.g. `auth repair
  // --start-login` — can catch LoginAccountMismatchError and present it their own way.
  await assertLoginIdentityOrThrow({
    identity: options.identity,
    signedInAs,
    force: options.forceIdentitySwitch
  });

  await persistRefreshTokenToEnv(refreshToken, { envPath });
  await commitLoginIdentity({ identity: options.identity, signedInAs });
  humanLog(`Saved M365_REFRESH_TOKEN (and legacy GRAPH_REFRESH_TOKEN / EWS_REFRESH_TOKEN) to ${envPath}`);
  if (options.identity) {
    humanLog(`Identity profile "${options.identity}" verified as ${signedInAs ?? '(unknown — no UPN on token)'}.`);
  }
  emitEvent(json, { event: 'complete', env_path: envPath });

  return { envPath, signedInAs };
}

export interface RunBrowserLoginOptions extends RunLoginOptions {
  port?: number;
  open?: boolean;
  callbackTimeout?: string;
  /**
   * Injectable browser-login runner (defaults to the real loopback + PKCE flow). Mirrors the
   * `openBrowser` injection in `browser-login.ts`: lets the command wrapper's own tests verify
   * env-persist + identity-binding wiring without standing up a real HTTP round trip (the loopback
   * mechanism itself is covered exhaustively by `browser-login.test.ts`).
   */
  _runBrowserLogin?: typeof runBrowserLogin;
}

/** Browser authorization-code + PKCE login (`m365-agent-cli login --browser`, issue #244). */
export async function runBrowserLoginFlow(options: RunBrowserLoginOptions): Promise<RunLoginResult> {
  const json = options.json ?? false;
  const humanLog = makeHumanLog(json);
  const { envPath, clientId } = await resolveLoginEnvAndClientId(options.envFile, json);
  const tenant = getMicrosoftTenantPathSegment();
  const timeoutMs = options.callbackTimeout ? (parseCacheTtlMs(options.callbackTimeout) ?? undefined) : undefined;
  const browserLogin = options._runBrowserLogin ?? runBrowserLogin;

  humanLog(`Starting browser login${options.identity ? ` for ${options.identity}` : ''}...`);

  let result: BrowserLoginResult;
  try {
    result = await browserLogin({
      clientId,
      tenant,
      scope: GRAPH_DEVICE_CODE_LOGIN_SCOPES,
      port: options.port,
      open: options.open,
      callbackTimeoutMs: timeoutMs,
      onAuthorizationUrl: (url) => {
        const redirectUri = new URL(url).searchParams.get('redirect_uri');
        humanLog('Open this URL if the browser does not launch automatically:');
        humanLog(url);
        humanLog(`Waiting for Microsoft redirect on ${redirectUri} ...`);
        emitEvent(json, { event: 'authorization_url', url, redirect_uri: redirectUri });
      }
    });
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    console.error(`\nBrowser login failed: ${message}`);
    // fatalJson always calls process.exit — this assignment is never actually reached at runtime,
    // but satisfies TS's definite-assignment analysis for `result` below (a `never` value is
    // assignable to anything).
    result = await fatalJson(json, { event: 'error', error: 'browser_login_failed', error_description: message });
  }

  // See runDeviceCodeLogin's comment: identity assertion stays a throw, not fatalJson, so other
  // callers (auth repair --start-login) can present a mismatch their own way; the actual profile
  // commit happens after the refresh token is durably persisted.
  await assertLoginIdentityOrThrow({
    identity: options.identity,
    signedInAs: result.signedInAs,
    force: options.forceIdentitySwitch
  });

  await persistRefreshTokenToEnv(result.refreshToken, { envPath });
  await commitLoginIdentity({ identity: options.identity, signedInAs: result.signedInAs });
  humanLog('\nLogin complete.');
  humanLog(`Signed in as: ${result.signedInAs ?? '(unknown — no UPN on token)'}`);
  humanLog(`Saved M365_REFRESH_TOKEN (and legacy GRAPH_REFRESH_TOKEN / EWS_REFRESH_TOKEN) to ${envPath}`);
  humanLog('Verification: m365-agent-cli verify-token --capabilities');
  emitEvent(json, { event: 'authenticated', username: result.signedInAs ?? null });
  emitEvent(json, { event: 'complete', env_path: envPath });

  return { envPath, signedInAs: result.signedInAs };
}

export const loginCommand = new Command('login')
  .description(
    'Interactive login to obtain refresh tokens via OAuth2 (device code, or --browser for authorization-code + PKCE)'
  )
  .option(
    '--env-file <path>',
    'Load/save EWS_CLIENT_ID and refresh tokens in this file (e.g. ~/.config/m365-agent-cli/.env.beta). Overrides vars from the default .env loaded at startup.'
  )
  .option(
    '--json',
    'Emit newline-delimited JSON events (device_code/authorization_url, authenticated, complete, error) to stdout for unattended automation; human-readable text goes to stderr. Requires EWS_CLIENT_ID to be preset (no interactive prompt). See docs/UNATTENDED_LOGIN.md.'
  )
  .option(
    '--identity <name>',
    'Bind this login to a named identity profile slug (see `profiles`). Refuses to complete if the resulting account differs from a previously verified account for this slug.'
  )
  .option(
    '--force-identity-switch',
    'With --identity, allow rebinding an identity slug to a different Microsoft account than previously verified.'
  )
  .option('--browser', 'Use a browser authorization-code + PKCE flow instead of device code')
  .option(
    '--localhost-port <port>',
    'Loopback port for --browser (0 = pick a free port automatically)',
    (v) => Number.parseInt(v, 10),
    0
  )
  .option('--no-open', 'With --browser, print the authorization URL instead of launching the system browser')
  .option(
    '--callback-timeout <duration>',
    'With --browser, how long to wait for the Microsoft redirect (e.g. 5m, 10m, 90s). Default 5m.'
  )
  .action(
    async (opts: {
      envFile?: string;
      json?: boolean;
      identity?: string;
      forceIdentitySwitch?: boolean;
      browser?: boolean;
      localhostPort?: number;
      open?: boolean;
      callbackTimeout?: string;
    }) => {
      const json = opts.json ?? false;

      try {
        if (opts.browser) {
          await runBrowserLoginFlow({
            envFile: opts.envFile,
            identity: opts.identity,
            forceIdentitySwitch: opts.forceIdentitySwitch,
            port: opts.localhostPort,
            open: opts.open,
            callbackTimeout: opts.callbackTimeout,
            json
          });
          return;
        }
        await runDeviceCodeLogin({
          envFile: opts.envFile,
          identity: opts.identity,
          forceIdentitySwitch: opts.forceIdentitySwitch,
          json
        });
      } catch (err) {
        // LoginAccountMismatchError (and any other throw not already handled via fatalJson inside
        // the flows above) lands here — present it consistently in both output modes.
        const message = err instanceof Error ? err.message : String(err);
        console.error(`Error: ${message}`);
        const errorCode = err instanceof LoginAccountMismatchError ? 'identity_mismatch' : 'login_failed';
        await fatalJson(json, { event: 'error', error: errorCode, error_description: message });
      }
    }
  );
