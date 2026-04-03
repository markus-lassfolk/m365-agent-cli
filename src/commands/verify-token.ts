import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import { getJwtPayloadAppId } from '../lib/jwt-utils.js';
import { applyEnvFileOverrides, getGlobalEnvFilePath, resolveEnvFilePathArgument } from '../lib/utils.js';

export const verifyTokenCommand = new Command('verify-token')
  .description('Verify Graph API token scopes and permissions')
  .option('--token <token>', 'Use a specific Graph token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--json', 'Output as JSON')
  .option(
    '--env-file <path>',
    'Load EWS_CLIENT_ID and refresh token from this file (e.g. ~/.config/m365-agent-cli/.env.beta) before verifying'
  )
  .action(async (options: { token?: string; json?: boolean; identity?: string; envFile?: string }) => {
    if (options.envFile) {
      applyEnvFileOverrides(resolveEnvFilePathArgument(options.envFile));
    }
    const authResult = await resolveGraphAuth({ token: options.token, identity: options.identity });
    if (!authResult.success || !authResult.token) {
      if (options.json) {
        console.log(JSON.stringify({ error: authResult.error || 'Failed to resolve auth token' }, null, 2));
      } else {
        console.error(`Error: ${authResult.error || 'Failed to resolve auth token'}`);
      }
      process.exit(1);
    }

    const token = authResult.token;
    try {
      const parts = token.split('.');
      if (parts.length !== 3) throw new Error('Invalid JWT format');
      const payload = JSON.parse(Buffer.from(parts[1], 'base64url').toString('utf8'));

      if (options.json) {
        console.log(JSON.stringify(payload, null, 2));
        return;
      }

      const envPath = options.envFile ? resolveEnvFilePathArgument(options.envFile) : getGlobalEnvFilePath();
      console.log(`Configuration file: ${envPath}`);
      if (process.env.EWS_CLIENT_ID) {
        console.log(`EWS_CLIENT_ID (Entra app): ${process.env.EWS_CLIENT_ID}`);
      } else {
        console.log('EWS_CLIENT_ID: (not set — check .env or login)');
      }
      if (!options.envFile && !process.env.M365_AGENT_ENV_FILE?.trim()) {
        console.log(
          'Tip: For a second app (.env.beta), set $env:M365_AGENT_ENV_FILE to that path in PowerShell, or use: verify-token --env-file "$env:USERPROFILE\\.config\\m365-agent-cli\\.env.beta"'
        );
      }
      console.log('');

      console.log('\u2713 Token Verified\n');
      const tokenAppId = getJwtPayloadAppId(token);
      console.log(`  App ID (from access token): ${tokenAppId ?? 'N/A'}`);
      console.log(`  Tenant: ${payload.tid || 'N/A'}`);
      console.log(`  User:   ${payload.upn || payload.email || 'N/A'}`);
      console.log(`  Name:   ${payload.name || 'N/A'}`);

      if (payload.scp) {
        console.log('\n  Delegated Scopes (scp):');
        payload.scp.split(' ').forEach((scope: string) => {
          console.log(`    - ${scope}`);
        });
      }

      if (payload.roles && Array.isArray(payload.roles)) {
        console.log('\n  Application Roles (roles):');
        payload.roles.forEach((role: string) => {
          console.log(`    - ${role}`);
        });
      }

      if (!payload.scp && (!payload.roles || payload.roles.length === 0)) {
        console.log('\n  No scopes or roles found in token.');
      }

      const envClient = process.env.EWS_CLIENT_ID?.trim();
      const tokenApp = tokenAppId?.trim();
      if (envClient && tokenApp && envClient.toLowerCase() !== tokenApp.toLowerCase()) {
        console.log(
          '\n  Warning: Access token app id does not match EWS_CLIENT_ID. Use verify-token --env-file for .env.beta, set M365_AGENT_ENV_FILE before starting the CLI, or delete ~/.config/m365-agent-cli/token-cache-*.json and run login again.'
        );
      }
    } catch (err: any) {
      if (options.json) {
        console.log(JSON.stringify({ error: `Failed to parse token: ${err.message}` }, null, 2));
      } else {
        console.error(`Failed to parse token: ${err.message}`);
      }
      process.exit(1);
    }
  });
