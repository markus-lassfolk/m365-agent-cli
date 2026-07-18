import { chmodSync, existsSync, mkdirSync, writeSync } from 'node:fs';
import { readFile } from 'node:fs/promises';
import { dirname } from 'node:path';
import { createInterface } from 'node:readline/promises';
import { Command } from 'commander';
import { atomicWriteUtf8File } from '../lib/atomic-write.js';
import { persistRefreshTokenToEnv } from '../lib/env-persist.js';
import { GRAPH_DEVICE_CODE_LOGIN_SCOPES } from '../lib/graph-oauth-scopes.js';
import { getMicrosoftTenantPathSegment, isValidJwtStructure } from '../lib/jwt-utils.js';
import { applyEnvFileOverrides, getGlobalEnvFilePath, resolveEnvFilePathArgument } from '../lib/utils.js';

/**
 * In `--json` mode the command writes newline-delimited JSON events to stdout so an unattended
 * wrapper can parse them without scraping free-form log lines. No-op when JSON mode is off.
 * See docs/UNATTENDED_LOGIN.md for the event shapes and an end-to-end automation example.
 */
function emitEvent(json: boolean, event: Record<string, unknown>): void {
  if (json) {
    // Synchronous write to fd 1 (not process.stdout.write): on POSIX a piped stdout is async and
    // buffered, so an immediate process.exit(1) after an error event could truncate it before it
    // flushes. writeSync hands the (small) line to the OS before returning, and keeps events ordered.
    writeSync(1, `${JSON.stringify(event)}\n`);
  }
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

async function performDeviceCodeFlow(
  clientId: string,
  tenant: string,
  scope: string,
  label: string,
  envPath: string,
  json: boolean
): Promise<string> {
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
    emitEvent(json, {
      event: 'error',
      error: deviceCodeJson.error ?? 'devicecode_request_failed',
      error_description: deviceCodeJson.error_description
    });
    process.exit(1);
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
  let username: string | undefined;
  let pollInterval = interval;

  humanLog(`Waiting for ${label} authentication...`);

  while (!authenticated) {
    if (Date.now() > expiresAt) {
      console.error(`\n${label} device code expired. Please run the command again.`);
      emitEvent(json, {
        event: 'error',
        error: 'expired_token',
        error_description: 'Device code expired before sign-in completed.'
      });
      process.exit(1);
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
          const parts = tokenJson.access_token.split('.');
          const payload = JSON.parse(Buffer.from(parts[1], 'base64url').toString('utf8'));

          const rawUsername = payload.upn || payload.email;
          username = rawUsername ? rawUsername.replace(/[\r\n]/g, '') : undefined;

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
        emitEvent(json, {
          event: 'error',
          error: 'no_refresh_token',
          error_description: 'No refresh token returned; ensure the offline_access scope is granted.'
        });
        process.exit(1);
      }
    } else if (tokenJson.error === 'authorization_pending') {
      // Continue polling
    } else if (tokenJson.error === 'slow_down') {
      pollInterval += 5000;
    } else {
      console.error(`\n${label} authentication failed:`, tokenJson.error_description || tokenJson.error);
      emitEvent(json, {
        event: 'error',
        error: tokenJson.error ?? 'authentication_failed',
        error_description: tokenJson.error_description
      });
      process.exit(1);
    }
  }

  humanLog(`\n${label} authentication successful!`);
  emitEvent(json, { event: 'authenticated', username: username ?? null });

  return refreshToken;
}

export const loginCommand = new Command('login')
  .description('Interactive login to obtain refresh tokens via OAuth2 Device Code flow')
  .option(
    '--env-file <path>',
    'Load/save EWS_CLIENT_ID and refresh tokens in this file (e.g. ~/.config/m365-agent-cli/.env.beta). Overrides vars from the default .env loaded at startup.'
  )
  .option(
    '--json',
    'Emit newline-delimited JSON events (device_code, authenticated, complete, error) to stdout for unattended automation; human-readable text goes to stderr. Requires EWS_CLIENT_ID to be preset (no interactive prompt). See docs/UNATTENDED_LOGIN.md.'
  )
  .action(async (opts: { envFile?: string; json?: boolean }) => {
    const json = opts.json ?? false;
    const humanLog = makeHumanLog(json);

    let envPath = getGlobalEnvFilePath();
    if (opts.envFile) {
      envPath = resolveEnvFilePathArgument(opts.envFile);
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
        emitEvent(json, {
          event: 'error',
          error: 'missing_client_id',
          error_description: 'Set EWS_CLIENT_ID in the environment or .env before using --json (non-interactive) mode.'
        });
        process.exit(1);
      }

      const rl = createInterface({
        input: process.stdin,
        output: process.stdout
      });
      clientId = await rl.question('Enter your EWS_CLIENT_ID: ');
      rl.close();
      clientId = clientId.trim();

      if (!clientId) {
        console.error('EWS_CLIENT_ID is required.');
        process.exit(1);
      }

      // Save it to .env
      envContent += `\nEWS_CLIENT_ID=${clientId}\n`;
      await atomicWriteUtf8File(envPath, `${envContent.trim()}\n`, 0o600);
    }

    humanLog('');
    humanLog(`Configuration file: ${envPath}`);
    if (!process.env.M365_AGENT_ENV_FILE?.trim() && !opts.envFile) {
      humanLog(
        'Tip: For a second app (e.g. beta), use --env-file ~/.config/m365-agent-cli/.env.beta or export M365_AGENT_ENV_FILE to that path before running the CLI.'
      );
    }
    humanLog(`Application (client) ID: ${clientId}`);
    humanLog('');

    const tenant = getMicrosoftTenantPathSegment();

    // Use a single Graph Device Code flow to obtain a multi-resource refresh token (see src/lib/graph-oauth-scopes.ts)
    const rawToken = await performDeviceCodeFlow(
      clientId,
      tenant,
      GRAPH_DEVICE_CODE_LOGIN_SCOPES,
      'Microsoft 365',
      envPath,
      json
    );
    const refreshToken = rawToken.replace(/[\r\n]/g, '');

    await persistRefreshTokenToEnv(refreshToken, { envPath });

    humanLog(`Saved M365_REFRESH_TOKEN (and legacy GRAPH_REFRESH_TOKEN / EWS_REFRESH_TOKEN) to ${envPath}`);
    emitEvent(json, { event: 'complete', env_path: envPath });
  });
