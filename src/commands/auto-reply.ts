import { Command } from 'commander';
import { resolveAuth } from '../lib/auth.js';
import { setAutoReplyRule, getAutoReplyRule } from '../lib/ews-client.js';

export const autoReplyCommand = new Command('auto-reply')
  .description('Manage server-side out-of-office (OOF) auto-reply templates via EWS Inbox Rules')
  .option('--message <text>', 'The message text for the auto-reply template')
  .option('--enable', 'Enable the auto-reply rule')
  .option('--disable', 'Disable the auto-reply rule')
  .option('--start <datetime>', 'Start datetime (ISO string)')
  .option('--end <datetime>', 'End datetime (ISO string)')
  .option('--mailbox <email>', 'Target mailbox (if different from authenticated user)')
  .action(async (options) => {
    try {
      const auth = await resolveAuth();
      if (!auth || !auth.token) {
        process.exit(1);
      }

      if (options.message || options.enable || options.disable || options.start || options.end) {
        let enabled = true;
        let messageText = options.message;
        let start = options.start ? new Date(options.start) : undefined;
        let end = options.end ? new Date(options.end) : undefined;

        const currentRuleRes = await getAutoReplyRule(auth.token, options.mailbox);
        
        if (currentRuleRes.ok && currentRuleRes.data) {
          if (options.enable === undefined && options.disable === undefined) {
            enabled = currentRuleRes.data.enabled;
          } else {
            enabled = options.enable ? true : false;
          }
          if (!messageText) messageText = currentRuleRes.data.messageText;
          if (!options.start && currentRuleRes.data.startTime) start = currentRuleRes.data.startTime;
          if (!options.end && currentRuleRes.data.endTime) end = currentRuleRes.data.endTime;
        } else {
          if (options.disable) enabled = false;
          if (!messageText) {
            console.error('Error: --message is required when creating a new auto-reply rule.');
            process.exit(1);
          }
        }

        console.log(`Setting auto-reply rule (enabled: ${enabled})...`);
        const result = await setAutoReplyRule(
          auth.token,
          messageText,
          enabled,
          start,
          end,
          options.mailbox
        );

        if (!result.ok) {
          console.error(`Failed to set auto-reply rule: ${result.status} ${(result.error as any)?.message || result.error}`);
          process.exit(1);
        }

        console.log('Auto-reply rule successfully set.');
      } else {
        const result = await getAutoReplyRule(auth.token, options.mailbox);
        
        if (!result.ok) {
          console.error(`Failed to get auto-reply rule: ${result.status} ${(result.error as any)?.message || result.error}`);
          process.exit(1);
        }

        if (!result.data) {
          console.log('No auto-reply template rule found.');
        } else {
          console.log('Auto-Reply Template Rule:');
          console.log(`  Enabled: ${result.data.enabled}`);
          console.log(`  Start Time: ${result.data.startTime ? result.data.startTime.toISOString() : 'None'}`);
          console.log(`  End Time:   ${result.data.endTime ? result.data.endTime.toISOString() : 'None'}`);
          console.log(`\nMessage:\n${result.data.messageText}`);
        }
      }
    } catch (err) {
      console.error('An unexpected error occurred:', err);
      process.exit(1);
    }
  });
