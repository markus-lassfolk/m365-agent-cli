import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import { getSchedule } from '../lib/graph-schedule.js';

export const scheduleCommand = new Command('schedule')
  .description('Get merged free/busy schedule for multiple users')
  .argument('<emails...>', 'One or more email addresses to check')
  .requiredOption('--start <date>', 'Start date/time (e.g. 2026-04-01T00:00:00Z or 2026-04-01)')
  .requiredOption('--end <date>', 'End date/time (e.g. 2026-04-07T00:00:00Z or 2026-04-07)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token (bypass interactive auth)')
  .action(async (emails: string[], options: { start: string; end: string; json?: boolean; token?: string }) => {
    const authResult = await resolveGraphAuth({ token: options.token });
    if (!authResult.success || !authResult.token) {
      if (options.json) {
        console.log(JSON.stringify({ error: authResult.error }, null, 2));
      } else {
        console.error(`Error: ${authResult.error}`);
      }
      process.exit(1);
    }

    const startDate = new Date(options.start);
    const endDate = new Date(options.end);

    if (Number.isNaN(startDate.getTime()) || Number.isNaN(endDate.getTime())) {
      const errorMessage =
        'Invalid start or end date. Please provide ISO 8601 date/time values (e.g. 2026-04-01T00:00:00Z or 2026-04-01).';
      if (options.json) {
        console.log(JSON.stringify({ error: errorMessage }, null, 2));
      } else {
        console.error(`Error: ${errorMessage}`);
      }
      process.exit(1);
    }

    // dateTime should not include Z/offset - keep dateTime and timeZone separate
    const startDateTime = startDate.toISOString().replace('Z', '');
    const endDateTime = endDate.toISOString().replace('Z', '');

    const result = await getSchedule(authResult.token, {
      schedules: emails,
      startTime: {
        dateTime: startDateTime,
        timeZone: 'UTC'
      },
      endTime: {
        dateTime: endDateTime,
        timeZone: 'UTC'
      },
      availabilityViewInterval: 60
    });

    if (!result.ok || !result.data) {
      if (options.json) {
        console.log(JSON.stringify({ error: result.error }, null, 2));
      } else {
        console.error('Error fetching schedule:', result.error?.message || 'Unknown error');
      }
      process.exit(1);
    }

    if (options.json) {
      console.log(JSON.stringify(result.data, null, 2));
      return;
    }

    console.log(`\nSchedule for ${emails.join(', ')}`);
    console.log(`From: ${startDateTime}\nTo:   ${endDateTime}\n`);

    for (const schedule of result.data.value) {
      console.log(`User: ${schedule.scheduleId}`);
      if (schedule.error) {
        console.log(`  Error: ${schedule.error.message || schedule.error.responseCode}`);
        continue;
      }

      if (schedule.workingHours) {
        const wh = schedule.workingHours;
        const days = wh.daysOfWeek?.join(', ') || 'N/A';
        console.log(`  Working Hours: ${wh.startTime} - ${wh.endTime} (${wh.timeZone?.name}) on ${days}`);
      }

      if (schedule.scheduleItems && schedule.scheduleItems.length > 0) {
        console.log('  Busy times:');
        for (const item of schedule.scheduleItems) {
          const status = item.status || 'Busy';
          const start = item.start?.dateTime ? new Date(`${item.start.dateTime}Z`).toLocaleString() : 'Unknown';
          const end = item.end?.dateTime ? new Date(`${item.end.dateTime}Z`).toLocaleString() : 'Unknown';
          const subject = item.subject ? ` - ${item.subject}` : '';
          console.log(`    [${status}] ${start} to ${end}${subject}`);
        }
      } else {
        console.log('  No busy times scheduled.');
      }
      console.log();
    }
  });
