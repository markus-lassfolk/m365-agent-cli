import { Command } from 'commander';
import { createSubscription, deleteSubscription } from '../lib/graph-subscriptions.js';

export const subscribeCommand = new Command('subscribe')
  .description('Subscribe to Microsoft Graph push notifications')
  .argument('[resource]', 'Resource to subscribe to (e.g. mail, event, contact, todoTask)')
  .option('--url <url>', 'Webhook notification URL')
  .option('--expiry <datetime>', 'Expiration datetime (ISO 8601, defaults to 3 days from now)')
  .option('--change-type <type>', 'Change type (comma-separated)', 'created,updated')
  .option('--token <token>', 'Use a specific token')
  .action(async (resource, options, cmd) => {
    if (!resource) {
      return cmd.help();
    }
    if (!options.url) {
      console.error('Error: --url is required.');
      process.exit(1);
    }

    // Map friendly resource names to graph endpoints
    const mapResource = (res: string) => {
      switch (res.toLowerCase()) {
        case 'mail':
          return 'me/messages';
        case 'event':
          return 'me/events';
        case 'contact':
          return 'me/contacts';
        case 'todotask':
          // Note: Todo subscriptions require a specific list ID.
          // Use the format: me/todo/lists/{listId}/tasks
          // For the default Tasks list, use: me/todo/lists/Tasks/tasks
          return 'me/todo/lists/Tasks/tasks';
        default:
          return res;
      }
    };

    // Generate clientState for subscription validation (if GRAPH_CLIENT_STATE env is set)
    const clientState = process.env.GRAPH_CLIENT_STATE;

    const graphResource = mapResource(resource);

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
        options.token
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
  .command('cancel <id>')
  .description('Cancel an existing subscription')
  .option('--token <token>', 'Use a specific token')
  .action(async (id, options) => {
    try {
      console.log(`Deleting subscription ${id}...`);
      const res = await deleteSubscription(id, options.token);
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
