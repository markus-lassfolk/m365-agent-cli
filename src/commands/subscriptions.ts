import { Command } from 'commander';
import { listSubscriptions } from '../lib/graph-subscriptions.js';

export const subscriptionsCommand = new Command('subscriptions').description('Manage Microsoft Graph subscriptions');

subscriptionsCommand
  .command('list')
  .description('List all active subscriptions')
  .option('--token <token>', 'Use a specific token')
  .action(async (options: { token?: string }) => {
    try {
      const res = await listSubscriptions(options.token);
      if (!res.ok || !res.data) {
        console.error(`Failed to list subscriptions: ${res.error?.message}`);
        process.exit(1);
      }
      const subs = res.data;
      if (subs.length === 0) {
        console.log('No active subscriptions found.');
        return;
      }
      console.log(`Found ${subs.length} active subscription(s):`);
      console.log(JSON.stringify(subs, null, 2));
    } catch (err) {
      console.error(err instanceof Error ? err.message : err);
      process.exit(1);
    }
  });
