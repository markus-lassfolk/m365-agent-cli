import { existsSync, mkdirSync } from 'node:fs';
import { readFile } from 'node:fs/promises';
import { homedir } from 'node:os';
import { join } from 'node:path';
import { createInterface } from 'node:readline/promises';
import { Command } from 'commander';
import { atomicWriteUtf8File } from '../lib/atomic-write.js';
import { getMicrosoftTenantPathSegment } from '../lib/jwt-utils.js';

async function performDeviceCodeFlow(clientId: string, tenant: string, scope: string, label: string): Promise<string> {
  console.log(`\nInitiating Device Code flow for ${label}...`);

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
    process.exit(1);
  }

  console.log('\n=========================================================');
  console.log(deviceCodeJson.message);
  console.log('=========================================================\n');

  const deviceCode = deviceCodeJson.device_code;
  const interval = (deviceCodeJson.interval || 5) * 1000;
  const expiresAt = Date.now() + (deviceCodeJson.expires_in || 900) * 1000;

  let authenticated = false;
  let refreshToken = '';
  let pollInterval = interval;

  console.log(`Waiting for ${label} authentication...`);

  while (!authenticated) {
    if (Date.now() > expiresAt) {
      console.error(`\n${label} device code expired. Please run the command again.`);
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
        const parts = tokenJson.access_token.split('.');

        if (parts.length === 3) {
          const payload = JSON.parse(Buffer.from(parts[1], 'base64url').toString('utf8'));

          const rawUsername = payload.upn || payload.email;
          const username = rawUsername ? rawUsername.replace(/[\r\n]/g, '') : undefined;

          if (username) {
            let envContent = '';

            const configDir = join(homedir(), '.config', 'm365-agent-cli');
            mkdirSync(configDir, { recursive: true, mode: 0o700 });
            const envPath = join(configDir, '.env');

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

            console.log(`Saved EWS_USERNAME (${username}) to ${envPath}`);
          }
        }
      } catch (_e) {
        /* ignore parse errors */
      }
      if (!refreshToken) {
        console.error(`\nFailed to obtain ${label} refresh token. Ensure the offline_access scope is granted.`);
        process.exit(1);
      }
    } else if (tokenJson.error === 'authorization_pending') {
      // Continue polling
    } else if (tokenJson.error === 'slow_down') {
      pollInterval += 5000;
    } else {
      console.error(`\n${label} authentication failed:`, tokenJson.error_description || tokenJson.error);
      process.exit(1);
    }
  }

  console.log(`\n${label} authentication successful!`);

  return refreshToken;
}

export const loginCommand = new Command('login')
  .description('Interactive login to obtain refresh tokens via OAuth2 Device Code flow')
  .action(async () => {
    let clientId = process.env.EWS_CLIENT_ID;

    // Read existing .env if present
    const configDir = join(homedir(), '.config', 'm365-agent-cli');
    mkdirSync(configDir, { recursive: true, mode: 0o700 });
    const envPath = join(configDir, '.env');
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

    const tenant = getMicrosoftTenantPathSegment();

    // Use a single Graph Device Code flow to obtain a multi-resource refresh token
    const graphScope =
      'offline_access User.Read Calendars.ReadWrite Mail.ReadWrite MailboxSettings.ReadWrite Files.ReadWrite.All Sites.ReadWrite.All Tasks.ReadWrite Group.ReadWrite.All';
    const rawToken = await performDeviceCodeFlow(clientId, tenant, graphScope, 'Microsoft 365');
    const refreshToken = rawToken.replace(/[\r\n]/g, '');

    // Save tokens immediately
    try {
      envContent = await readFile(envPath, 'utf8');
    } catch (err: any) {
      if (err.code !== 'ENOENT') throw err;
    }

    const upsertEnvLine = (key: string, value: string) => {
      const re = new RegExp(`^${key}=.*$`, 'm');
      if (re.test(envContent)) {
        envContent = envContent.replace(re, () => `${key}=${value}`);
      } else {
        envContent += `\n${key}=${value}\n`;
      }
    };

    upsertEnvLine('M365_REFRESH_TOKEN', refreshToken);
    upsertEnvLine('EWS_REFRESH_TOKEN', refreshToken);
    upsertEnvLine('GRAPH_REFRESH_TOKEN', refreshToken);

    envContent = envContent.replace(/\n{3,}/g, '\n\n');
    await atomicWriteUtf8File(envPath, `${envContent.trim()}\n`, 0o600);

    console.log(`Saved M365_REFRESH_TOKEN (and legacy GRAPH_REFRESH_TOKEN / EWS_REFRESH_TOKEN) to ${envPath}`);
  });
