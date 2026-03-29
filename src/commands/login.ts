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

    // Perform first device code flow for EWS
    const ewsScope = 'offline_access https://outlook.office365.com/EWS.AccessAsUser.All';

    console.log('\nInitiating Device Code flow for EWS...');

    const ewsDeviceCodeRes = await fetch(`https://login.microsoftonline.com/${tenant}/oauth2/v2.0/devicecode`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id: clientId,
        scope: ewsScope
      }).toString()
    });

    const ewsDeviceCodeJson = await ewsDeviceCodeRes.json();

    if (!ewsDeviceCodeRes.ok) {
      console.error('Failed to initiate EWS device code flow:', ewsDeviceCodeJson);
      process.exit(1);
    }

    console.log('\n=========================================================');
    console.log(ewsDeviceCodeJson.message);
    console.log('=========================================================\n');

    let ewsDeviceCode = ewsDeviceCodeJson.device_code;
    let ewsInterval = (ewsDeviceCodeJson.interval || 5) * 1000;
    let ewsExpiresAt = Date.now() + (ewsDeviceCodeJson.expires_in || 900) * 1000;

    let ewsAuthenticated = false;
    let ewsRefreshToken = '';
    let ewsPollInterval = ewsInterval;

    console.log('Waiting for EWS authentication...');

    while (!ewsAuthenticated) {
      if (Date.now() > ewsExpiresAt) {
        console.error('\nEWS device code expired. Please run the command again.');
        process.exit(1);
      }

      await new Promise((resolve) => setTimeout(resolve, ewsPollInterval));

      const tokenRes = await fetch(`https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
          grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
          client_id: clientId,
          device_code: ewsDeviceCode
        }).toString()
      });

      const tokenJson = await tokenRes.json();

      if (tokenRes.ok) {
        ewsAuthenticated = true;
        ewsRefreshToken = tokenJson.refresh_token;
        if (!ewsRefreshToken) {
          console.error('\nFailed to obtain EWS refresh token. Ensure the offline_access scope is granted.');
          process.exit(1);
        }
      } else if (tokenJson.error === 'authorization_pending') {
        // Continue polling
      } else if (tokenJson.error === 'slow_down') {
        ewsPollInterval += 5000;
      } else {
        console.error('\nEWS authentication failed:', tokenJson.error_description || tokenJson.error);
        process.exit(1);
      }
    }

    console.log('\nEWS authentication successful!');

    // Perform second device code flow for Graph
    const graphScope =
      'offline_access User.Read Calendars.ReadWrite Mail.ReadWrite Files.ReadWrite.All Sites.ReadWrite.All Tasks.ReadWrite Group.ReadWrite.All';

    console.log('\nInitiating Device Code flow for Microsoft Graph...');

    const graphDeviceCodeRes = await fetch(`https://login.microsoftonline.com/${tenant}/oauth2/v2.0/devicecode`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id: clientId,
        scope: graphScope
      }).toString()
    });

    const graphDeviceCodeJson = await graphDeviceCodeRes.json();

    if (!graphDeviceCodeRes.ok) {
      console.error('Failed to initiate Graph device code flow:', graphDeviceCodeJson);
      process.exit(1);
    }

    console.log('\n=========================================================');
    console.log(graphDeviceCodeJson.message);
    console.log('=========================================================\n');

    let graphDeviceCode = graphDeviceCodeJson.device_code;
    let graphInterval = (graphDeviceCodeJson.interval || 5) * 1000;
    let graphExpiresAt = Date.now() + (graphDeviceCodeJson.expires_in || 900) * 1000;

    let graphAuthenticated = false;
    let graphRefreshToken = '';
    let graphPollInterval = graphInterval;

    console.log('Waiting for Graph authentication...');

    while (!graphAuthenticated) {
      if (Date.now() > graphExpiresAt) {
        console.error('\nGraph device code expired. Please run the command again.');
        process.exit(1);
      }

      await new Promise((resolve) => setTimeout(resolve, graphPollInterval));

      const tokenRes = await fetch(`https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
          grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
          client_id: clientId,
          device_code: graphDeviceCode
        }).toString()
      });

      const tokenJson = await tokenRes.json();

      if (tokenRes.ok) {
        graphAuthenticated = true;
        graphRefreshToken = tokenJson.refresh_token;
        if (!graphRefreshToken) {
          console.error('\nFailed to obtain Graph refresh token. Ensure the offline_access scope is granted.');
          process.exit(1);
        }
      } else if (tokenJson.error === 'authorization_pending') {
        // Continue polling
      } else if (tokenJson.error === 'slow_down') {
        graphPollInterval += 5000;
      } else {
        console.error('\nGraph authentication failed:', tokenJson.error_description || tokenJson.error);
        process.exit(1);
      }
    }

    console.log('\nGraph authentication successful!');

    // Read the current .env content to update/append the refresh tokens
    if (existsSync(envPath)) {
      envContent = await readFile(envPath, 'utf8');
    }

    // Update or append EWS_REFRESH_TOKEN
    if (/^EWS_REFRESH_TOKEN=.*$/m.test(envContent)) {
      envContent = envContent.replace(/^EWS_REFRESH_TOKEN=.*$/m, `EWS_REFRESH_TOKEN=${ewsRefreshToken}`);
    } else {
      envContent += `\nEWS_REFRESH_TOKEN=${ewsRefreshToken}\n`;
    }

    // Update or append GRAPH_REFRESH_TOKEN
    if (/^GRAPH_REFRESH_TOKEN=.*$/m.test(envContent)) {
      envContent = envContent.replace(/^GRAPH_REFRESH_TOKEN=.*$/m, `GRAPH_REFRESH_TOKEN=${graphRefreshToken}`);
    } else {
      envContent += `\nGRAPH_REFRESH_TOKEN=${graphRefreshToken}\n`;
    }

    // Clean up multiple newlines if any
    envContent = envContent.replace(/\n{3,}/g, '\n\n');

    await writeFile(envPath, `${envContent.trim()}\n`, { encoding: 'utf8', mode: 0o600 });

    console.log('Saved EWS_REFRESH_TOKEN and GRAPH_REFRESH_TOKEN to .env file in the current directory.');
  });
