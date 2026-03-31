import { Command } from 'commander';
import { resolveAuth } from '../lib/auth.js';
import { parseDay } from '../lib/dates.js';
import { getOwaUserInfo, getScheduleViaOutlook } from '../lib/ews-client.js';

function formatTime(dateStr: string): string {
  const date = new Date(dateStr);
  return date.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit', hour12: false });
}

function formatDate(dateStr: string): string {
  const date = new Date(dateStr);
  return date.toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric' });
}

function _formatDateTime(dateStr: string): string {
  const _date = new Date(dateStr);
  return `${formatDate(dateStr)} ${formatTime(dateStr)}`;
}

function isValidEmail(value: string): boolean {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value);
}

function getDateRange(startDay: string, endDay?: string): { start: Date; end: Date; label: string } {
  const now = new Date();

  switch (startDay.toLowerCase()) {
    case 'week':
    case 'thisweek': {
      const start = new Date(now);
      const dayOfWeek = start.getDay();
      start.setDate(start.getDate() + (dayOfWeek === 0 ? 1 : 8 - dayOfWeek));
      start.setHours(0, 0, 0, 0);
      const end = new Date(start);
      end.setDate(end.getDate() + 4);
      end.setHours(23, 59, 59, 999);
      return { start, end, label: 'This Week (Mon-Fri)' };
    }
    case 'nextweek': {
      const start = new Date(now);
      const dayOfWeek = start.getDay();
      const daysUntilNextMonday = dayOfWeek === 0 ? 1 : 8 - dayOfWeek;
      start.setDate(start.getDate() + daysUntilNextMonday);
      start.setHours(0, 0, 0, 0);
      const end = new Date(start);
      end.setDate(end.getDate() + 4);
      end.setHours(23, 59, 59, 999);
      return { start, end, label: 'Next Week (Mon-Fri)' };
    }
  }

  const startDate = parseDay(startDay, { throwOnInvalid: true });
  startDate.setHours(0, 0, 0, 0);

  if (endDay) {
    const endDate = parseDay(endDay, { baseDate: startDate, weekdayDirection: 'next', throwOnInvalid: true });
    endDate.setHours(23, 59, 59, 999);
    return {
      start: startDate,
      end: endDate,
      label: `${formatDate(startDate.toISOString())} - ${formatDate(endDate.toISOString())}`
    };
  }

  const endDate = new Date(startDate);
  endDate.setHours(23, 59, 59, 999);
  return { start: startDate, end: endDate, label: formatDate(startDate.toISOString()) };
}

export const findtimeCommand = new Command('findtime')
  .description('Find available meeting times with one or more people')
  .argument('[start]', 'Start: today, tomorrow, monday-sunday, week, nextweek, or YYYY-MM-DD', 'nextweek')
  .argument('[endOrEmails...]', 'End day for range AND/OR email addresses')
  .option('--duration <minutes>', 'Meeting duration in minutes', '30')
  .option('--start <hour>', 'Work day start hour (0-23)', '9')
  .option('--end <hour>', 'Work day end hour (0-23)', '17')
  .option('--solo', "Only check specified people, don't include yourself")
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Use a specific authentication identity (default: default)')
  .option('--mailbox <email>', 'EWS anchor mailbox (delegated / shared mailbox context)')
  .action(
    async (
      startDay: string,
      endOrEmails: string[],
      options: {
        duration: string;
        start: string;
        end: string;
        solo?: boolean;
        json?: boolean;
        token?: string;
        identity?: string;
        mailbox?: string;
      },
      _cmd: any
    ) => {
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

      // Parse arguments: figure out which are dates vs emails
      const dateKeywords = [
        'today',
        'tomorrow',
        'monday',
        'tuesday',
        'wednesday',
        'thursday',
        'friday',
        'saturday',
        'sunday',
        'week',
        'thisweek',
        'nextweek'
      ];
      const isDateArg = (arg: string) => {
        if (dateKeywords.includes(arg.toLowerCase())) return true;
        if (/^\d{4}-\d{2}-\d{2}$/.test(arg)) return true;
        return false;
      };

      let endDay: string | undefined;
      const emails: string[] = [];

      for (const arg of endOrEmails) {
        if (isDateArg(arg) && !endDay) {
          endDay = arg;
          continue;
        }

        if (!isValidEmail(arg)) {
          console.error(`Error: Invalid attendee email: ${arg}`);
          console.error('All attendee arguments must be valid email addresses.');
          process.exit(1);
        }

        emails.push(arg);
      }

      // Get current user's email to include in search (unless --solo)
      if (!options.solo) {
        const userInfo = await getOwaUserInfo(authResult.token!);
        if (!userInfo.ok || !userInfo.data?.email) {
          if (options.json) {
            console.log(JSON.stringify({ error: 'Failed to determine user email' }, null, 2));
          } else {
            console.error('Error: Failed to determine user email');
          }
          process.exit(1);
        }
        // Add current user if not already in the list
        if (!emails.includes(userInfo.data.email)) {
          emails.unshift(userInfo.data.email);
        }
      }

      if (emails.length === 0) {
        console.error('Error: Please provide at least one email address.');
        console.error('\nUsage: m365-agent-cli findtime nextweek user@example.com');
        process.exit(1);
      }

      let start: Date;
      let end: Date;
      let label: string;

      try {
        ({ start, end, label } = getDateRange(startDay, endDay));
      } catch (err) {
        const message = err instanceof Error ? err.message : 'Invalid date value';
        if (options.json) {
          console.log(JSON.stringify({ error: message }, null, 2));
        } else {
          console.error(`Error: ${message}`);
        }
        process.exit(1);
      }
      const duration = parseInt(options.duration, 10);

      const result = await getScheduleViaOutlook(
        authResult.token!,
        emails,
        start.toISOString(),
        end.toISOString(),
        duration,
        undefined,
        options.mailbox
      );

      if (!result.ok || !result.data) {
        if (options.json) {
          console.log(JSON.stringify({ error: result.error?.message || 'Failed to find meeting times' }, null, 2));
        } else {
          console.error(`Error: ${result.error?.message || 'Failed to find meeting times'}`);
        }
        process.exit(1);
      }

      // Extract free slots from the response, filtered to working hours
      const workStart = parseInt(options.start, 10);
      const workEnd = parseInt(options.end, 10);

      const freeSlots = (result.data[0]?.scheduleItems || []).filter((item) => {
        if (item.status !== 'Free') return false;
        const hour = new Date(item.start.dateTime).getHours();
        return hour >= workStart && hour < workEnd;
      });

      if (options.json) {
        console.log(
          JSON.stringify(
            {
              attendees: emails,
              duration: duration,
              dateRange: { start: start.toISOString(), end: end.toISOString() },
              availableSlots: freeSlots.map((s) => ({
                start: s.start.dateTime,
                end: s.end.dateTime
              }))
            },
            null,
            2
          )
        );
        return;
      }

      console.log(`\n🗓️  Finding ${duration}-minute meeting times`);
      console.log(`   Attendees: ${emails.join(', ')}`);
      console.log(`   Date range: ${label}`);
      console.log('─'.repeat(50));

      if (freeSlots.length === 0) {
        console.log('\n  ❌ No available times found for all attendees.');
        console.log('     Try a longer date range or shorter meeting duration.');
      } else {
        console.log(`\n  ✅ Found ${freeSlots.length} available slot${freeSlots.length > 1 ? 's' : ''}:\n`);

        // Group by day
        const byDay = new Map<string, typeof freeSlots>();
        for (const slot of freeSlots) {
          const day = slot.start.dateTime.split('T')[0];
          if (!byDay.has(day)) byDay.set(day, []);
          byDay.get(day)?.push(slot);
        }

        for (const [day, slots] of byDay) {
          const dayLabel = formatDate(new Date(day).toISOString());
          console.log(`  ${dayLabel}:`);
          for (const slot of slots) {
            console.log(`    🟢 ${formatTime(slot.start.dateTime)} - ${formatTime(slot.end.dateTime)}`);
          }
        }
      }
      console.log();
    }
  );
