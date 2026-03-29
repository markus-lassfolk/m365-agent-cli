import { Command } from 'commander';
import { resolveAuth } from '../lib/auth.js';
import { parseDay } from '../lib/dates.js';
import { cancelEvent, deleteEvent, getCalendarEvents } from '../lib/ews-client.js';

function formatTime(dateStr: string): string {
  const date = new Date(dateStr);
  return date.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit', hour12: false });
}

function formatDate(dateStr: string): string {
  const date = new Date(dateStr);
  return date.toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric' });
}

export const deleteEventCommand = new Command('delete-event')
  .description('Delete/cancel a calendar event (sends cancellation if there are attendees)')
  .argument('[eventIndex]', 'Event index from the list (deprecated; use --id)')
  .option('--id <eventId>', 'Delete event by stable ID')
  .option('--day <day>', 'Day to show events from (today, tomorrow, YYYY-MM-DD) - note: may miss multi-day events crossing midnight', 'today')
  .option('--search <text>', 'Search for events by title')
  .option('--message <text>', 'Cancellation message to send to attendees')
  .option('--force-delete', 'Delete without sending cancellation (even with attendees)')
  .option('--occurrence <index>', 'Delete only the Nth occurrence of a recurring event')
  .option('--instance <date>', 'Delete only the occurrence on a specific date (YYYY-MM-DD)')
  .option('--scope <scope>', 'Scope: all (default), this (single occurrence), future (this and future)', 'all')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--mailbox <email>', 'Delete event in shared mailbox calendar')
  .action(
    async (
      _eventIndex: string | undefined,
      options: {
        id?: string;
        day: string;
        search?: string;
        message?: string;
        forceDelete?: boolean;
        occurrence?: string;
        instance?: string;
        scope: string;
        json?: boolean;
        token?: string;
        mailbox?: string;
      }
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

      // Get events for the day
      const baseDate = parseDay(options.day);
      const startOfDay = new Date(baseDate);
      startOfDay.setHours(0, 0, 0, 0);
      const endOfDay = new Date(baseDate);
      endOfDay.setHours(23, 59, 59, 999);

      const result = await getCalendarEvents(
        authResult.token!,
        startOfDay.toISOString(),
        endOfDay.toISOString(),
        options.mailbox
      );

      if (!result.ok || !result.data) {
        if (options.json) {
          console.log(JSON.stringify({ error: result.error?.message || 'Failed to fetch events' }, null, 2));
        } else {
          console.error(`Error: ${result.error?.message || 'Failed to fetch events'}`);
        }
        process.exit(1);
      }

      // Filter to events the user owns (IsOrganizer) and optionally by search
      let events = result.data.filter((e) => e.IsOrganizer && !e.IsCancelled);

      if (options.search) {
        const searchLower = options.search.toLowerCase();
        events = events.filter((e) => e.Subject?.toLowerCase().includes(searchLower));
      }

      // If no id provided, list events
      if (!options.id) {
        if (options.json) {
          console.log(
            JSON.stringify(
              {
                events: events.map((e, i) => ({
                  index: i + 1,
                  id: e.Id,
                  subject: e.Subject,
                  start: e.Start.DateTime,
                  end: e.End.DateTime
                }))
              },
              null,
              2
            )
          );
          return;
        }

        console.log(`\nYour events for ${formatDate(baseDate.toISOString())}:\n`);
        console.log('\u2500'.repeat(60));

        if (events.length === 0) {
          console.log('\n  No events found that you can delete.');
          console.log('  (You can only delete events you organized)\n');
          return;
        }

        for (let i = 0; i < events.length; i++) {
          const event = events[i];
          const startTime = formatTime(event.Start.DateTime);
          const endTime = formatTime(event.End.DateTime);
          const attendees = event.Attendees?.filter((a) => a.EmailAddress?.Address && a.Type !== 'Resource') || [];

          console.log(`\n  [${i + 1}] ${event.Subject}`);
          console.log(`      ${startTime} - ${endTime}`);
          console.log(`      ID: ${event.Id}`);
          if (event.Location?.DisplayName) {
            console.log(`      Location: ${event.Location.DisplayName}`);
          }
          if (attendees.length > 0) {
            console.log(`      Attendees: ${attendees.length} (will be notified on cancel)`);
          }
        }

        console.log(`\n${'\u2500'.repeat(60)}`);
        console.log('\nTo delete/cancel an event:');
        console.log('  clippy delete-event <number>                    # Cancel & notify attendees');
        console.log('  clippy delete-event <number> --message "Sorry"  # With cancellation message');
        console.log('  clippy delete-event <number> --force-delete     # Delete without notifying');
        console.log('');
        return;
      }

      // Delete the specified event by ID
      if (!options.id) {
        console.error('Please specify the event id with --id.');
        console.error('Run `clippy delete-event` to list events and IDs.');
        process.exit(1);
      }

      // Determine scope and occurrence ID
      let scope = options.scope as 'all' | 'this' | 'future';
      let occurrenceItemId: string | undefined;
      let targetEvent = events.find((e) => e.Id === options.id);

      // Validate scope: 'future' is not currently supported by EWS
      if (scope === 'future') {
        console.error('Error: --scope future is not supported.');
        console.error('EWS does not provide a native operation to delete "this and future" occurrences.');
        console.error('Use --scope this to delete a single occurrence, or --scope all to delete the entire series.');
        process.exit(1);
      }

      // If occurrence/instance flags are provided without explicit scope, default to 'this'
      if ((options.occurrence || options.instance) && options.scope === 'all') {
        scope = 'this';
      }

      if ((options.occurrence || options.instance) && scope === 'this') {
        // Find the occurrence by index or date, ensuring it matches the provided event ID
        if (options.instance) {
          // Find occurrence matching the specific date and event ID
          const instanceDate = parseDay(options.instance);
          instanceDate.setHours(0, 0, 0, 0);
          const occEvent = events.find((e) => {
            const eventDate = new Date(e.Start.DateTime);
            eventDate.setHours(0, 0, 0, 0);
            return eventDate.getTime() === instanceDate.getTime() && e.Id === options.id;
          });
          if (!occEvent) {
            console.error(
              `No occurrence found on ${options.instance} with ID ${options.id}. Try expanding the date range with --day.`
            );
            process.exit(1);
          }
          // For CalendarView items, the Id we get IS the occurrence ID
          occurrenceItemId = occEvent.Id;
          targetEvent = occEvent;
        } else if (options.occurrence) {
          const idx = parseInt(options.occurrence, 10);
          if (Number.isNaN(idx) || idx < 1) {
            console.error('--occurrence must be a positive integer');
            process.exit(1);
          }
          // Events from CalendarView are already individual occurrences
          if (idx > events.length) {
            console.error(
              `Invalid occurrence index: ${idx}. Only ${events.length} occurrence(s) found in the date range.`
            );
            process.exit(1);
          }
          const occEvent = events[idx - 1];
          if (occEvent.Id !== options.id) {
            console.error(`Occurrence ${idx} does not match the provided event ID ${options.id}.`);
            process.exit(1);
          }
          occurrenceItemId = occEvent.Id;
          targetEvent = occEvent;
        }
        console.log(`\nDeleting single occurrence: ${targetEvent!.Subject}`);
        console.log(
          `  ${formatDate(targetEvent!.Start.DateTime)} ${formatTime(targetEvent!.Start.DateTime)} - ${formatTime(targetEvent!.End.DateTime)}`
        );
      } else if (!targetEvent) {
        console.error(`Invalid event id: ${options.id}`);
        process.exit(1);
      } else {
        // Full series delete
        if (scope !== 'all') {
          // future scope needs the occurrence ID too
          console.log(`\nDeleting: ${targetEvent.Subject} (scope: ${scope})`);
        } else {
          console.log(`\nDeleting: ${targetEvent.Subject}`);
        }
        console.log(
          `  ${formatDate(targetEvent.Start.DateTime)} ${formatTime(targetEvent.Start.DateTime)} - ${formatTime(targetEvent.End.DateTime)}`
        );
      }

      // Check if event has attendees (other than organizer)
      const attendees = targetEvent!.Attendees?.filter((a) => a.EmailAddress?.Address && a.Type !== 'Resource') || [];
      const hasAttendees = attendees.length > 0;

      let deleteResult: Awaited<ReturnType<typeof deleteEvent>>;
      let action: string;

      if (hasAttendees && !options.forceDelete && scope === 'all') {
        // Use cancel to send cancellation notices for full series
        console.log(`  Attendees: ${attendees.map((a) => a.EmailAddress?.Address).join(', ')}`);
        console.log(`  Sending cancellation notices...`);
        deleteResult = await cancelEvent({
          token: authResult.token!,
          eventId: targetEvent!.Id,
          comment: options.message,
          mailbox: options.mailbox
        });
        action = 'cancelled';
      } else {
        // Delete with or without notification based on forceDelete flag
        deleteResult = await deleteEvent({
          token: authResult.token!,
          eventId: targetEvent!.Id,
          occurrenceItemId,
          scope,
          mailbox: options.mailbox,
          forceDelete: options.forceDelete,
          comment: options.message
        });
        action = 'deleted';
      }

      if (!deleteResult.ok) {
        if (options.json) {
          console.log(JSON.stringify({ error: deleteResult.error?.message || `Failed to ${action} event` }, null, 2));
        } else {
          console.error(`\nError: ${deleteResult.error?.message || `Failed to ${action} event`}`);
        }
        process.exit(1);
      }

      if (options.json) {
        console.log(
          JSON.stringify(
            {
              success: true,
              action,
              event: targetEvent!.Subject,
              attendeesNotified: hasAttendees && !options.forceDelete ? attendees.length : 0,
              ...(deleteResult.info ? { info: deleteResult.info } : {})
            },
            null,
            2
          )
        );
      } else {
        if (deleteResult.info) {
          console.warn(`\nNote: ${deleteResult.info}\n`);
        }
        if (hasAttendees && !options.forceDelete) {
          console.log(`\n\u2713 Event cancelled. ${attendees.length} attendee(s) notified.\n`);
        } else {
          console.log('\n\u2713 Event deleted.\n');
        }
      }
    }
  );
