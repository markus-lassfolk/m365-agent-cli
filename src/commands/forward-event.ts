import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import { forwardEvent } from '../lib/graph-event.js';
import { checkReadOnly } from '../lib/utils.js';

export const forwardEventCommand = new Command('forward-event')
  .description('Forward a calendar event to additional recipients')
  .alias('forward')
  .argument('<eventId>', 'The ID of the event to forward')
  .argument('<recipients...>', 'Email addresses to forward the event to')
  .option('--comment <text>', 'Optional comment to include in the forwarded invitation')
  .option('--token <token>', 'Use a specific token')
  .action(async (eventId: string, recipients: string[], options: { comment?: string; token?: string }, cmd: any) => {
    checkReadOnly(cmd);
    const authResult = await resolveGraphAuth({ token: options.token });
    if (!authResult.success) {
      console.error(`Error: ${authResult.error}`);
      process.exit(1);
    }

    console.log(`Forwarding event...`);
    console.log(`  Event ID:   ${eventId}`);
    console.log(`  Recipients: ${recipients.join(', ')}`);
    if (options.comment) console.log(`  Comment:    ${options.comment}`);

    const response = await forwardEvent({
      token: authResult.token!,
      eventId,
      toRecipients: recipients,
      comment: options.comment
    });

    if (!response.ok) {
      console.error(`\nError: ${response.error?.message || 'Failed to forward event'}`);
      process.exit(1);
    }

    console.log('\n\u2713 Successfully forwarded the event.');
  });
