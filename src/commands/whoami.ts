import { Command } from 'commander';
import { resolveAuth } from '../lib/auth.js';
import { getOwaUserInfo } from '../lib/ews-client.js';

export const whoamiCommand = new Command('whoami')
  .description('Show authenticated user information')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Use a specific authentication identity (default: default)')
  .action(async (options: { json?: boolean; token?: string; identity?: string }) => {
    const authResult = await resolveAuth({
      token: options.token,
      identity: options.identity
    });

    if (!authResult.success) {
      if (options.json) {
        console.log(JSON.stringify({ error: authResult.error }, null, 2));
      } else {
        console.error(`Error: ${authResult.error}`);
        console.error('\nCheck your .env file for EWS_CLIENT_ID and EWS_REFRESH_TOKEN.');
      }
      process.exit(1);
    }

    const userInfo = await getOwaUserInfo(authResult.token!);

    if (!userInfo.ok || !userInfo.data) {
      if (options.json) {
        console.log(
          JSON.stringify(
            {
              error: userInfo.error?.message || 'Failed to fetch user info',
              authenticated: true
            },
            null,
            2
          )
        );
      } else {
        console.log('\u2713 Authenticated');
        console.log('  Could not fetch user details from EWS API');
      }
      process.exit(0);
    }

    const { displayName, email } = userInfo.data;

    if (options.json) {
      const result: { displayName: string; email: string; authenticated: boolean; identity?: string } = {
        displayName,
        email,
        authenticated: true
      };

      // Only include identity if token-based auth was not used
      if (!options.token) {
        result.identity = options.identity || 'default';
      }

      console.log(JSON.stringify(result, null, 2));
    } else {
      console.log('\u2713 Authenticated');

      // Only display identity if token-based auth was not used
      if (!options.token) {
        const identity = options.identity || 'default';
        console.log(`  Identity: ${identity}`);
      }

      console.log(`  Name: ${displayName}`);
      console.log(`  Email: ${email}`);
    }
  });
