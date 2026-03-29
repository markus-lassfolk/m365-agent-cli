import { existsSync } from 'node:fs';
import { readFile, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { createInterface } from 'node:readline/promises';
import { Command } from 'commander';
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

          const parts = tokenJson.access_token.split(".");

          if (parts.length === 3) {

            const payload = JSON.parse(Buffer.from(parts[1], "base64url").toString("utf8"));

            const username = payload.upn || payload.email;

            if (username) {

              let envContent = "";

              const envPath = join(process.cwd(), ".env");

              if (existsSync(envPath)) {

                envContent = await readFile(envPath, "utf8");

              }

              if (/^EWS_USERNAME=.*$/m.test(envContent)) {

                envContent = envContent.replace(/^EWS_USERNAME=.*$/m, () => `EWS_USERNAME=${username}`);

              } else {

                envContent += `\nEWS_USERNAME=${username}\n`;

              }

              await writeFile(envPath, `${envContent.trim()}\n`, { encoding: "utf8", mode: 0o600 });

              console.log(`Saved EWS_USERNAME (${username}) to .env file`);

            }

          }

        } catch (_e) { /* ignore parse errors */ }
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
    const envPath = join(process.cwd(), '.env');
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
      await writeFile(envPath, `${envContent.trim()}\n`, { encoding: 'utf8', mode: 0o600 });
    }

    const tenant = getMicrosoftTenantPathSegment();

    // Use a single Graph Device Code flow to obtain a multi-resource refresh token
    const graphScope =
      'offline_access User.Read Calendars.ReadWrite Mail.ReadWrite Files.ReadWrite.All Sites.ReadWrite.All Tasks.ReadWrite Group.ReadWrite.All';
    const refreshToken = await performDeviceCodeFlow(clientId, tenant, graphScope, 'Microsoft 365');

    // Save tokens immediately
    if (existsSync(envPath)) {
      envContent = await readFile(envPath, 'utf8');
    }

    // Update or append EWS_REFRESH_TOKEN
    if (/^EWS_REFRESH_TOKEN=.*$/m.test(envContent)) {
      envContent = envContent.replace(/^EWS_REFRESH_TOKEN=.*$/m, () => `EWS_REFRESH_TOKEN=${refreshToken}`);
    } else {
      envContent += `\nEWS_REFRESH_TOKEN=${refreshToken}\n`;
    }

    // Update or append GRAPH_REFRESH_TOKEN
    if (/^GRAPH_REFRESH_TOKEN=.*$/m.test(envContent)) {
      envContent = envContent.replace(/^GRAPH_REFRESH_TOKEN=.*$/m, () => `GRAPH_REFRESH_TOKEN=${refreshToken}`);
    } else {
      envContent += `\nGRAPH_REFRESH_TOKEN=${refreshToken}\n`;
    }

    envContent = envContent.replace(/\n{3,}/g, '\n\n');
    await writeFile(envPath, `${envContent.trim()}\n`, { encoding: 'utf8', mode: 0o600 });

    console.log('Saved GRAPH_REFRESH_TOKEN and EWS_REFRESH_TOKEN to .env file in the current directory.');
  });
