import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';

export const verifyTokenCommand = new Command('verify-token')
  .description('Verify Graph API token scopes and permissions')
  .option('--token <token>', 'Use a specific Graph token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--json', 'Output as JSON')
  .action(async (options: { token?: string; json?: boolean; identity?: string }) => {
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

      console.log('\u2713 Token Verified\n');
      console.log(`  App ID: ${payload.appid || 'N/A'}`);
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
    } catch (err: any) {
      if (options.json) {
        console.log(JSON.stringify({ error: `Failed to parse token: ${err.message}` }, null, 2));
      } else {
        console.error(`Failed to parse token: ${err.message}`);
      }
      process.exit(1);
    }
  });
