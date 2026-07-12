import { Command } from 'commander';
import { parseDay, parseTimeToDate } from '../lib/dates.js';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import { proposeNewTime } from '../lib/graph-event.js';
import { toJsonError } from '../lib/json-error.js';
import { checkReadOnly } from '../lib/utils.js';

export const counterCommand = new Command('counter')
  .description('Propose a new time for a calendar event')
  .alias('propose-new-time')
  .argument('<eventId>', 'The ID of the event')
  .argument('<start>', 'Proposed start time (e.g., 13:00, 1pm)')
  .argument('<end>', 'Proposed end time (e.g., 14:00, 2pm)')
  .option('--day <day>', 'Day for the proposed time (today, tomorrow, YYYY-MM-DD)', 'today')
  .option('--json', 'Output result as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Mailbox whose calendar contains the event (delegation)')
  .action(
    async (
      eventId: string,
      startTime: string,
      endTime: string,
      options: { day: string; token?: string; json?: boolean; identity?: string; user?: string },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const authResult = await resolveGraphAuth({ token: options.token, identity: options.identity });
      if (!authResult.success) {
        console.error(`Error: ${authResult.error}`);
        process.exit(1);
      }

      // Parse dates and times
      let baseDate: Date;
      try {
        baseDate = parseDay(options.day, { throwOnInvalid: true });
      } catch (err) {
        const message = err instanceof Error ? err.message : 'Invalid day value';
        console.error(`Error: ${message}`);
        process.exit(1);
      }

      let start: Date;
      try {
        start = parseTimeToDate(startTime, baseDate, { throwOnInvalid: true });
      } catch (err) {
        const message = err instanceof Error ? err.message : 'Invalid start time';
        console.error(`Error: ${message}`);
        process.exit(1);
      }

      let end: Date;
      try {
        end = parseTimeToDate(endTime, baseDate, { throwOnInvalid: true });
      } catch (err) {
        const message = err instanceof Error ? err.message : 'Invalid end time';
        console.error(`Error: ${message}`);
        process.exit(1);
      }

      // The Graph API expects times with the time zone context. For simplicity, we can pass
      // the ISO 8601 string and UTC as the time zone.
      const startDateTime = start.toISOString();
      const endDateTime = end.toISOString();
      const timeZone = 'UTC';

      if (!options.json) {
        console.log(`Proposing new time for event...`);
        console.log(`  Event ID: ${eventId}`);
        console.log(`  Proposed Start: ${start.toLocaleString()}`);
        console.log(`  Proposed End:   ${end.toLocaleString()}`);
      }

      const response = await proposeNewTime({
        token: authResult.token!,
        eventId,
        startDateTime,
        endDateTime,
        timeZone,
        user: options.user
      });

      if (!response.ok) {
        if (options.json) {
          console.log(
            JSON.stringify(
              { success: false, error: toJsonError(response.error?.message || 'Failed to propose new time') },
              null,
              2
            )
          );
        } else {
          console.error(`\nError: ${response.error?.message || 'Failed to propose new time'}`);
        }
        process.exit(1);
      }

      if (options.json) {
        console.log(
          JSON.stringify(
            {
              success: true,
              eventId,
              proposedStart: start.toISOString(),
              proposedEnd: end.toISOString(),
              timeZone
            },
            null,
            2
          )
        );
      } else {
        console.log('\n\u2713 Successfully proposed a new time for the event.');
      }
    }
  );
