import { Command } from 'commander';
import { resolveAuth } from '../lib/auth.js';
import { getAutoReplyRule, setAutoReplyRule } from '../lib/ews-client.js';
import { getExchangeBackend } from '../lib/exchange-backend.js';
import { checkReadOnly } from '../lib/utils.js';

export const autoReplyCommand = new Command('auto-reply')
  .description(
    'Manage OOF-style auto-reply via EWS Inbox Rules (legacy). Prefer `oof` for Graph mailboxSettings automatic replies.'
  )
  .option('--message <text>', 'The message text for the auto-reply template')
  .option('--enable', 'Enable the auto-reply rule')
  .option('--disable', 'Disable the auto-reply rule')
  .option('--start <datetime>', 'Start datetime (ISO string)')
  .option('--end <datetime>', 'End datetime (ISO string)')
  .option('--mailbox <email>', 'Target mailbox (if different from authenticated user)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific EWS token')
  .option('--identity <name>', 'Use a specific authentication identity (default: default)')
  .action(async (options, cmd: any) => {
    checkReadOnly(cmd);
    if (getExchangeBackend() === 'graph') {
      const msg =
        'auto-reply uses EWS Inbox Rules only. For Microsoft Graph, use `m365-agent-cli oof` (mailbox automatic replies). To run this command, set M365_EXCHANGE_BACKEND=ews or auto.';
      if (options.json) {
        console.log(JSON.stringify({ error: msg }, null, 2));
      } else {
        console.error(`Error: ${msg}`);
      }
      process.exit(1);
    }
    try {
      const auth = await resolveAuth({ token: options.token, identity: options.identity });
      if (!auth.success || !auth.token) {
        const msg = auth.error || 'Authentication failed';
        if (options.json) {
          console.log(JSON.stringify({ error: msg }, null, 2));
        } else {
          console.error(`Error: ${msg}`);
        }
        process.exit(1);
      }

      // Validate mutually exclusive --enable and --disable
      if (options.enable && options.disable) {
        const msg = 'Error: --enable and --disable cannot be used together.';
        if (options.json) {
          console.log(JSON.stringify({ error: msg }, null, 2));
        } else {
          console.error(msg);
        }
        process.exit(1);
      }

      // Validate date inputs
      let start: Date | undefined;
      let end: Date | undefined;
      if (options.start) {
        start = new Date(options.start);
        if (!Number.isFinite(start.getTime())) {
          const msg = 'Error: --start must be a valid ISO datetime string.';
          if (options.json) {
            console.log(JSON.stringify({ error: msg }, null, 2));
          } else {
            console.error(msg);
          }
          process.exit(1);
        }
      }
      if (options.end) {
        end = new Date(options.end);
        if (!Number.isFinite(end.getTime())) {
          const msg = 'Error: --end must be a valid ISO datetime string.';
          if (options.json) {
            console.log(JSON.stringify({ error: msg }, null, 2));
          } else {
            console.error(msg);
          }
          process.exit(1);
        }
      }

      if (options.message || options.enable || options.disable || options.start || options.end) {
        let enabled = true;
        let messageText = options.message;
        const currentRuleRes = await getAutoReplyRule(auth.token, options.mailbox);

        // Handle error case separately from "not found"
        if (!currentRuleRes.ok) {
          const msg = `Failed to get current auto-reply rule: ${currentRuleRes.status} ${currentRuleRes.error?.message || ''}`;
          if (options.json) {
            console.log(JSON.stringify({ error: msg }, null, 2));
          } else {
            console.error(msg);
          }
          process.exit(1);
        }

        if (currentRuleRes.data) {
          if (options.enable === undefined && options.disable === undefined) {
            enabled = currentRuleRes.data.enabled;
          } else {
            enabled = !!options.enable;
          }
          if (!messageText) messageText = currentRuleRes.data.messageText;
          if (!options.start && currentRuleRes.data.startTime) start = currentRuleRes.data.startTime;
          if (!options.end && currentRuleRes.data.endTime) end = currentRuleRes.data.endTime;
        } else {
          if (options.disable) enabled = false;
          if (!messageText) {
            const msg = 'Error: --message is required when creating a new auto-reply rule.';
            if (options.json) {
              console.log(JSON.stringify({ error: msg }, null, 2));
            } else {
              console.error(msg);
            }
            process.exit(1);
          }
        }

        if (options.json) {
          console.log(JSON.stringify({ status: 'updating', enabled }, null, 2));
        } else {
          console.log(`Setting auto-reply rule (enabled: ${enabled})...`);
        }
        const result = await setAutoReplyRule(auth.token, messageText!, enabled, start, end, options.mailbox);

        if (!result.ok) {
          const msg = `Failed to set auto-reply rule: ${result.status} ${(result.error as any)?.message || result.error}`;
          if (options.json) {
            console.log(JSON.stringify({ error: msg }, null, 2));
          } else {
            console.error(msg);
          }
          process.exit(1);
        }

        if (options.json) {
          console.log(JSON.stringify({ status: 'success' }, null, 2));
        } else {
          console.log('Auto-reply rule successfully set.');
        }
      } else {
        const result = await getAutoReplyRule(auth.token, options.mailbox);

        if (!result.ok) {
          const msg = `Failed to get auto-reply rule: ${result.status} ${(result.error as any)?.message || result.error}`;
          if (options.json) {
            console.log(JSON.stringify({ error: msg }, null, 2));
          } else {
            console.error(msg);
          }
          process.exit(1);
        }

        if (!result.data) {
          if (options.json) {
            console.log(JSON.stringify({ exists: false }, null, 2));
          } else {
            console.log('No auto-reply template rule found.');
          }
        } else {
          if (options.json) {
            console.log(
              JSON.stringify(
                {
                  exists: true,
                  enabled: result.data.enabled,
                  startTime: result.data.startTime ? result.data.startTime.toISOString() : null,
                  endTime: result.data.endTime ? result.data.endTime.toISOString() : null,
                  messageText: result.data.messageText
                },
                null,
                2
              )
            );
          } else {
            console.log('Auto-Reply Template Rule:');
            console.log(`  Enabled: ${result.data.enabled}`);
            console.log(`  Start Time: ${result.data.startTime ? result.data.startTime.toISOString() : 'None'}`);
            console.log(`  End Time:   ${result.data.endTime ? result.data.endTime.toISOString() : 'None'}`);
            console.log(`\nMessage:\n${result.data.messageText}`);
          }
        }
      }
    } catch (err) {
      const msg = err instanceof Error ? err.message : 'An unexpected error occurred';
      if (options.json) {
        console.log(JSON.stringify({ error: msg }, null, 2));
      } else {
        console.error('An unexpected error occurred:', msg);
      }
      process.exit(1);
    }
  });
