import { Command } from 'commander';
import {
  createSubscription,
  deleteSubscription,
  listSubscriptions,
  renewSubscription
} from '../lib/graph-subscriptions.js';
import { checkReadOnly } from '../lib/utils.js';

export const subscribeCommand = new Command('subscribe')
  .description('Subscribe to Microsoft Graph push notifications (see also: list, renew, cancel)')
  .argument('[resource]', 'Resource to subscribe to (e.g. mail, event, contact, todoTask)')
  .option('--url <url>', 'Webhook notification URL')
  .option('--expiry <datetime>', 'Expiration datetime (ISO 8601, defaults to 3 days from now)')
  .option('--change-type <type>', 'Change type (comma-separated)', 'created,updated')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Subscribe under this user or shared mailbox (users/{id}/...)')
  .action(async (resource, options, cmd) => {
    if (!resource) {
      return cmd.help();
    }
    if (!options.url) {
      console.error('Error: --url is required.');
      process.exit(1);
    }

    checkReadOnly(cmd);

    // Map friendly resource names to graph endpoints
    const mapResource = (res: string, user?: string) => {
      const prefix = user?.trim() ? `users/${encodeURIComponent(user.trim())}` : 'me';
      switch (res.toLowerCase()) {
        case 'mail':
          return `${prefix}/messages`;
        case 'event':
          return `${prefix}/events`;
        case 'contact':
          return `${prefix}/contacts`;
        case 'todotask':
          // Note: Todo subscriptions require a specific list ID.
          // Use the format: me/todo/lists/{listId}/tasks
          // For the default Tasks list, use: me/todo/lists/Tasks/tasks
          return `${prefix}/todo/lists/Tasks/tasks`;
        default:
          return res;
      }
    };

    // Generate clientState for subscription validation (if GRAPH_CLIENT_STATE env is set)
    const clientState = process.env.GRAPH_CLIENT_STATE;

    const graphResource = mapResource(resource, options.user);

    // Default expiration to 3 days (Graph allows up to 3 days for most resources)
    let expiry = options.expiry;
    if (!expiry) {
      const date = new Date();
      date.setDate(date.getDate() + 3);
      // Ensure we don't exceed max limits by shaving off a minute
      date.setMinutes(date.getMinutes() - 1);
      expiry = date.toISOString();
    }

    try {
      console.log(`Creating subscription for ${graphResource}...`);
      const res = await createSubscription(
        graphResource,
        options.changeType,
        options.url,
        expiry,
        clientState,
        options.token,
        options.identity
      );
      if (!res.ok) {
        console.error(`Failed to create subscription: ${res.error?.message}`);
        process.exit(1);
      }
      const sub = res.data;
      console.log('Subscription created successfully!');
      console.log(JSON.stringify(sub, null, 2));
    } catch (err) {
      console.error(err instanceof Error ? err.message : err);
      process.exit(1);
    }
  });

subscribeCommand
  .command('list')
  .description('List active subscriptions (Graph GET /subscriptions)')
  .option('--json', 'Output as JSON array')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (opts: { json?: boolean; token?: string; identity?: string }) => {
    try {
      const res = await listSubscriptions(opts.token, opts.identity);
      if (!res.ok || !res.data) {
        console.error(`Failed to list subscriptions: ${res.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(res.data, null, 2));
      else {
        for (const s of res.data) {
          const exp = s.expirationDateTime ?? '';
          console.log(`${s.id}\t${exp}\t${s.resource}`);
        }
      }
    } catch (err) {
      console.error(err instanceof Error ? err.message : err);
      process.exit(1);
    }
  });

subscribeCommand
  .command('renew <id>')
  .description('Extend subscription expiration (Graph PATCH /subscriptions/{id})')
  .requiredOption('--expiry <datetime>', 'New expiration (ISO 8601)')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (id: string, opts: { expiry: string; token?: string; identity?: string }, cmd) => {
    try {
      checkReadOnly(cmd);
      const res = await renewSubscription(id, opts.expiry, opts.token, opts.identity);
      if (!res.ok) {
        console.error(`Failed to renew subscription: ${res.error?.message}`);
        process.exit(1);
      }
      console.log('Subscription renewed.');
    } catch (err) {
      console.error(err instanceof Error ? err.message : err);
      process.exit(1);
    }
  });

subscribeCommand
  .command('cancel <id>')
  .description('Cancel an existing subscription')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (id, options, cmd) => {
    try {
      checkReadOnly(cmd);
      console.log(`Deleting subscription ${id}...`);
      const res = await deleteSubscription(id, options.token, options.identity);
      if (!res.ok) {
        console.error(`Failed to delete subscription: ${res.error?.message}`);
        process.exit(1);
      }
      console.log('Subscription deleted successfully.');
    } catch (err) {
      console.error(err instanceof Error ? err.message : err);
      process.exit(1);
    }
  });
