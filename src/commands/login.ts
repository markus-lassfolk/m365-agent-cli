import { existsSync } from 'node:fs';
import { readFile, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { createInterface } from 'node:readline/promises';
import { Command } from 'commander';
import { getMicrosoftTenantPathSegment } from '../lib/jwt-utils.js';

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
    const scope =
      'offline_access https://outlook.office365.com/EWS.AccessAsUser.All User.Read Calendars.ReadWrite Mail.ReadWrite Files.ReadWrite.All Sites.ReadWrite.All Tasks.ReadWrite Group.ReadWrite.All';

    console.log('\nInitiating Device Code flow...');

    const deviceCodeRes = await fetch(`https://login.microsoftonline.com/${tenant}/oauth2/v2.0/devicecode`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id: clientId,
        scope
      }).toString()
    });

    const deviceCodeJson = await deviceCodeRes.json();

    if (!deviceCodeRes.ok) {
      console.error('Failed to initiate device code flow:', deviceCodeJson);
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

    console.log('Waiting for authentication...');

    while (!authenticated) {
      if (Date.now() > expiresAt) {
        console.error('\nDevice code expired. Please run the command again.');
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
        if (!refreshToken) {
          console.error('\nFailed to obtain refresh token. Ensure the offline_access scope is granted.');
          process.exit(1);
        }
      } else if (tokenJson.error === 'authorization_pending') {
        // Continue polling
      } else if (tokenJson.error === 'slow_down') {
        pollInterval += 5000;
      } else {
        console.error('\nAuthentication failed:', tokenJson.error_description || tokenJson.error);
        process.exit(1);
      }
    }

    console.log('\nAuthentication successful!');

    // Read the current .env content to update/append the refresh tokens
    if (existsSync(envPath)) {
      envContent = await readFile(envPath, 'utf8');
    }

    // Update or append EWS_REFRESH_TOKEN
    if (/^EWS_REFRESH_TOKEN=.*$/m.test(envContent)) {
      envContent = envContent.replace(/^EWS_REFRESH_TOKEN=.*$/m, `EWS_REFRESH_TOKEN=${refreshToken}`);
    } else {
      envContent += `\nEWS_REFRESH_TOKEN=${refreshToken}\n`;
    }

    // Update or append GRAPH_REFRESH_TOKEN
    if (/^GRAPH_REFRESH_TOKEN=.*$/m.test(envContent)) {
      envContent = envContent.replace(/^GRAPH_REFRESH_TOKEN=.*$/m, `GRAPH_REFRESH_TOKEN=${refreshToken}`);
    } else {
      envContent += `\nGRAPH_REFRESH_TOKEN=${refreshToken}\n`;
    }

    // Clean up multiple newlines if any
    envContent = envContent.replace(/\n{3,}/g, '\n\n');

    await writeFile(envPath, `${envContent.trim()}\n`, { encoding: 'utf8', mode: 0o600 });

    console.log('Saved EWS_REFRESH_TOKEN and GRAPH_REFRESH_TOKEN to .env file in the current directory.');
  });
