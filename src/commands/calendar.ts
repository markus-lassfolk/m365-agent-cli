import { Command } from 'commander';
import { resolveAuth } from '../lib/auth.js';
import { parseDay, parseLocalDate } from '../lib/dates.js';
import { type CalendarAttendee, type CalendarEvent, getCalendarEvents } from '../lib/ews-client.js';

function formatTime(dateStr: string): string {
  const date = parseLocalDate(dateStr);
  return date.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit', hour12: false });
}

function formatDate(dateStr: string): string {
  const date = parseLocalDate(dateStr);
  return date.toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric' });
}

/**
 * Convert a local-midnight Date to a UTC ISO string for EWS CalendarView.
 *
 * EWS CalendarView StartDate/EndDate are UTC-datetime strings.
 * CalendarView is exclusive on EndDate, so we set end = next day's local midnight.
 *
 * Example for UTC-5 (EST) on March 15:
 *   start = 2024-03-15T00:00 local = 2024-03-15T05:00:00Z
 *   end   = 2024-03-16T00:00 local = 2024-03-16T05:00:00Z
 */
function toEWSRange(localMidnight: Date): { start: string; end: string } {
  const start = new Date(localMidnight);
  start.setHours(0, 0, 0, 0);

  const end = new Date(start);
  end.setDate(end.getDate() + 1);
  end.setHours(0, 0, 0, 0);

  return { start: start.toISOString(), end: end.toISOString() };
}

function getDateRange(startDay: string, endDay?: string): { start: string; end: string; label: string } {
  const now = new Date();

  // Handle special range keywords
  switch (startDay.toLowerCase()) {
    case 'week':
    case 'thisweek': {
      const start = new Date(now);
      const dayOfWeek = start.getDay();
      const diff = dayOfWeek === 0 ? -6 : 1 - dayOfWeek; // Monday
      start.setDate(start.getDate() + diff);
      start.setHours(0, 0, 0, 0);
      const end = new Date(start);
      end.setDate(end.getDate() + 7); // exclusive end = next Monday midnight
      return { start: start.toISOString(), end: end.toISOString(), label: 'This Week' };
    }
    case 'lastweek': {
      const start = new Date(now);
      const dayOfWeek = start.getDay();
      const diff = dayOfWeek === 0 ? -13 : -6 - dayOfWeek; // Last Monday
      start.setDate(start.getDate() + diff);
      start.setHours(0, 0, 0, 0);
      const end = new Date(start);
      end.setDate(end.getDate() + 7); // exclusive end = next Monday midnight
      return { start: start.toISOString(), end: end.toISOString(), label: 'Last Week' };
    }
    case 'nextweek': {
      const start = new Date(now);
      const dayOfWeek = start.getDay();
      const diff = dayOfWeek === 0 ? 1 : 8 - dayOfWeek; // Next Monday
      start.setDate(start.getDate() + diff);
      start.setHours(0, 0, 0, 0);
      const end = new Date(start);
      end.setDate(end.getDate() + 7); // exclusive end = next Monday midnight
      return { start: start.toISOString(), end: end.toISOString(), label: 'Next Week' };
    }
  }

  // Single day or start of range
  const startDate = parseDay(startDay, { weekdayDirection: 'previous' });
  startDate.setHours(0, 0, 0, 0);

  if (endDay) {
    // Date range - use nearestForward for end date
    const endDate = parseDay(endDay, { baseDate: startDate, weekdayDirection: 'nearestForward' });
    endDate.setHours(0, 0, 0, 0);
    // Exclusive end: next day's midnight
    const endExclusive = new Date(endDate);
    endExclusive.setDate(endExclusive.getDate() + 1);

    const label = `${formatDate(startDate.toISOString())} - ${formatDate(endDate.toISOString())}`;
    return { start: startDate.toISOString(), end: endExclusive.toISOString(), label };
  }

  // Single day — use toEWSRange for consistent UTC conversion
  const { end: endISO } = toEWSRange(startDate);

  return {
    start: startDate.toISOString(),
    end: endISO,
    label: formatDate(startDate.toISOString())
  };
}

function getResponseIcon(response: string): string {
  switch (response) {
    case 'Accepted':
      return '✓';
    case 'Declined':
      return '✗';
    case 'TentativelyAccepted':
      return '?';
    case 'NotResponded':
      return '·';
    case 'Organizer':
      return '★';
    default:
      return '·';
  }
}

function displayEvent(event: CalendarEvent, verbose: boolean): void {
  const startTime = formatTime(event.Start.DateTime);
  const endTime = formatTime(event.End.DateTime);
  const location = event.Location?.DisplayName || '';
  const cancelled = event.IsCancelled ? ' [CANCELLED]' : '';

  if (event.IsAllDay) {
    console.log(`  📅 All day: ${event.Subject}${cancelled}`);
  } else {
    console.log(`  ${startTime} - ${endTime}: ${event.Subject}${cancelled}`);
  }

  if (location) {
    console.log(`     📍 ${location}`);
  }

  if (verbose) {
    // Show organizer if not self
    if (!event.IsOrganizer && event.Organizer?.EmailAddress?.Name) {
      console.log(`     👤 Organizer: ${event.Organizer.EmailAddress.Name}`);
    }

    // Show recurrence info
    if (event.RecurrenceDescription) {
      console.log(`     🔁 ${event.RecurrenceDescription}`);
    } else if (event.FirstOccurrence || event.LastOccurrence) {
      // Series bounds without full description
      const first = event.FirstOccurrence ? ` from ${formatDate(event.FirstOccurrence.Start)}` : '';
      const last = event.LastOccurrence ? ` until ${formatDate(event.LastOccurrence.Start)}` : '';
      if (first || last) {
        console.log(`     🔁 Recurring series${first}${last}`);
      }
    }

    // Show attendees
    if (event.Attendees && event.Attendees.length > 0) {
      const attendeeList = event.Attendees.map(
        (a: CalendarAttendee) => `${getResponseIcon(a.Status.Response)} ${a.EmailAddress.Name}`
      ).join(', ');
      console.log(`     👥 ${attendeeList}`);
    }

    // Show categories
    if (event.Categories && event.Categories.length > 0) {
      console.log(`     🏷️  ${event.Categories.join(', ')}`);
    }

    // Show body preview if available
    if (event.BodyPreview) {
      const preview = event.BodyPreview.substring(0, 80).replace(/\n/g, ' ');
      console.log(`     📝 ${preview}${event.BodyPreview.length > 80 ? '...' : ''}`);
    }

    // Show recurrence series info if available
    if (event.FirstOccurrence || event.LastOccurrence) {
      const first = event.FirstOccurrence ? event.FirstOccurrence.Start.substring(0, 10) : 'unknown';
      const last = event.LastOccurrence ? event.LastOccurrence.Start.substring(0, 10) : 'unknown';
      console.log(`     🔄 Series: First: ${first}, Last: ${last}`);
    }
    if (event.ModifiedOccurrences && event.ModifiedOccurrences.length > 0) {
      console.log(
        `     ✏️  Modified exceptions: ${event.ModifiedOccurrences.map((o) => o.OriginalStart.substring(0, 10)).join(', ')}`
      );
    }
    if (event.DeletedOccurrences && event.DeletedOccurrences.length > 0) {
      console.log(
        `     🗑️  Deleted exceptions: ${event.DeletedOccurrences.map((o) => o.Start.substring(0, 10)).join(', ')}`
      );
    }
  }
}

export const calendarCommand = new Command('calendar')
  .description('View calendar events')
  .argument(
    '[start]',
    'Start day: today, yesterday, tomorrow, monday-sunday, week, lastweek, nextweek, or YYYY-MM-DD',
    'today'
  )
  .argument('[end]', 'End day for range (optional)')
  .option('-v, --verbose', 'Show attendees and more details')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--mailbox <email>', 'Delegated or shared mailbox calendar')
  .action(
    async (
      startDay: string,
      endDay: string | undefined,
      options: { json?: boolean; token?: string; verbose?: boolean; mailbox?: string }
    ) => {
      const authResult = await resolveAuth({
        token: options.token
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

      const { start, end, label } = getDateRange(startDay, endDay);
      const result = await getCalendarEvents(authResult.token!, start, end, options.mailbox);

      if (!result.ok || !result.data) {
        if (options.json) {
          console.log(JSON.stringify({ error: result.error?.message || 'Failed to fetch events' }, null, 2));
        } else {
          console.error(`Error: ${result.error?.message || 'Failed to fetch events'}`);
        }
        process.exit(1);
      }

      const events = result.data.filter((e) => !e.IsCancelled);

      if (options.json) {
        console.log(JSON.stringify(events, null, 2));
        return;
      }

      console.log(`\n📆 Calendar for ${label}${options.mailbox ? ` — ${options.mailbox}` : ''}`);
      console.log('─'.repeat(40));

      if (events.length === 0) {
        console.log('  No events scheduled.');
      } else {
        // Group by date for multi-day ranges
        const eventsByDate = new Map<string, CalendarEvent[]>();
        for (const event of events) {
          const localDate = parseLocalDate(event.Start.DateTime);
          const dateKey = `${localDate.getFullYear()}-${String(localDate.getMonth() + 1).padStart(2, '0')}-${String(localDate.getDate()).padStart(2, '0')}`;
          if (!eventsByDate.has(dateKey)) {
            eventsByDate.set(dateKey, []);
          }
          eventsByDate.get(dateKey)?.push(event);
        }

        // Check if multiple days
        if (eventsByDate.size > 1) {
          for (const [dateKey, dayEvents] of eventsByDate) {
            const dayLabel = formatDate(dateKey);
            console.log(`\n  ${dayLabel}`);
            for (const event of dayEvents) {
              displayEvent(event, options.verbose ?? false);
            }
          }
        } else {
          for (const event of events) {
            displayEvent(event, options.verbose ?? false);
          }
        }
      }
      console.log();
    }
  );
