import { Command } from 'commander';
import { parseDay, parseTimeToDate } from '../lib/dates.js';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import { proposeNewTime } from '../lib/graph-event.js';

export const counterCommand = new Command('counter')
  .description('Propose a new time for a calendar event')
  .alias('propose-new-time')
  .argument('<eventId>', 'The ID of the event')
  .argument('<start>', 'Proposed start time (e.g., 13:00, 1pm)')
  .argument('<end>', 'Proposed end time (e.g., 14:00, 2pm)')
  .option('--day <day>', 'Day for the proposed time (today, tomorrow, YYYY-MM-DD)', 'today')
  .option('--token <token>', 'Use a specific token')
  .action(async (eventId: string, startTime: string, endTime: string, options: { day: string; token?: string }) => {
    const authResult = await resolveGraphAuth({ token: options.token });
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

    console.log(`Proposing new time for event...`);
    console.log(`  Event ID: ${eventId}`);
    console.log(`  Proposed Start: ${start.toLocaleString()}`);
    console.log(`  Proposed End:   ${end.toLocaleString()}`);

    const response = await proposeNewTime({
      token: authResult.token!,
      eventId,
      startDateTime,
      endDateTime,
      timeZone
    });

    if (!response.ok) {
      console.error(`\nError: ${response.error?.message || 'Failed to propose new time'}`);
      process.exit(1);
    }

    console.log('\n\u2713 Successfully proposed a new time for the event.');
  });
