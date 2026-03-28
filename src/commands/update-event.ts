import { Command } from 'commander';
import { resolveAuth } from '../lib/auth.js';
import { getCalendarEvents, updateEvent, searchRooms, getRooms } from '../lib/ews-client.js';
import { parseDay, parseTimeToDate, toLocalISOString } from '../lib/dates.js';

function formatTime(dateStr: string): string {
  const date = new Date(dateStr);
  return date.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit', hour12: false });
}

function formatDate(dateStr: string): string {
  const date = new Date(dateStr);
  return date.toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric' });
}

export const updateEventCommand = new Command('update-event')
  .description('Update a calendar event')
  .argument('[eventIndex]', 'Event index from the list (deprecated; use --id)')
  .option('--id <eventId>', 'Update event by stable ID')
  .option('--day <day>', 'Day to show events from (today, tomorrow, YYYY-MM-DD)', 'today')
  .option('--title <text>', 'New title/subject')
  .option('--description <text>', 'New description/body')
  .option('--start <time>', 'New start time (e.g., 14:00, 2pm)')
  .option('--end <time>', 'New end time (e.g., 15:00, 3pm)')
  .option(
    '--add-attendee <email>',
    'Add an attendee (can be used multiple times)',
    (val, arr: string[]) => [...arr, val],
    []
  )
  .option('--room <room>', 'Set/change meeting room (name or email)')
  .option('--location <text>', 'Set location text')
  .option('--teams', 'Make it a Teams meeting')
  .option('--no-teams', 'Remove Teams meeting')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--mailbox <email>', 'Update event in shared mailbox calendar')
  .action(
    async (
      _eventIndex: string | undefined,
      options: {
        id?: string;
        day: string;
        title?: string;
        description?: string;
        start?: string;
        end?: string;
        addAttendee: string[];
        room?: string;
        location?: string;
        teams?: boolean;
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
      let baseDate: Date;
      try {
        baseDate = parseDay(options.day, { throwOnInvalid: true });
      } catch (err) {
        const message = err instanceof Error ? err.message : 'Invalid day value';
        if (options.json) {
          console.log(JSON.stringify({ error: message }, null, 2));
        } else {
          console.error(`Error: ${message}`);
        }
        process.exit(1);
      }
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

      // Filter to events the user owns
      const events = result.data.filter((e) => e.IsOrganizer && !e.IsCancelled);

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
                  end: e.End.DateTime,
                  attendees: e.Attendees?.map((a) => a.EmailAddress?.Address)
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
          console.log('\n  No events found that you can update.');
          console.log('  (You can only update events you organized)\n');
          return;
        }

        for (let i = 0; i < events.length; i++) {
          const event = events[i];
          const startTime = formatTime(event.Start.DateTime);
          const endTime = formatTime(event.End.DateTime);

          console.log(`\n  [${i + 1}] ${event.Subject}`);
          console.log(`      ${startTime} - ${endTime}`);
          console.log(`      ID: ${event.Id}`);
          if (event.Location?.DisplayName) {
            console.log(`      Location: ${event.Location.DisplayName}`);
          }
          if (event.Attendees && event.Attendees.length > 0) {
            const attendeeList = event.Attendees.filter((a) => a.Type !== 'Resource')
              .map((a) => a.EmailAddress?.Address)
              .filter(Boolean);
            if (attendeeList.length > 0) {
              console.log(`      Attendees: ${attendeeList.join(', ')}`);
            }
          }
        }

        console.log(`\n${'\u2500'.repeat(60)}`);
        console.log('\nTo update an event:');
        console.log('  clippy update-event <number> --title "New Title"');
        console.log('  clippy update-event <number> --add-attendee user@example.com');
        console.log('  clippy update-event <number> --room "Taxi"');
        console.log('  clippy update-event <number> --start 14:00 --end 15:00');
        console.log('');
        return;
      }

      // Get the target event by ID
      const targetEvent = events.find((e) => e.Id === options.id);
      if (!targetEvent) {
        console.error(`Invalid event id: ${options.id}`);
        process.exit(1);
      }

      // Check if any update options were provided
      const hasUpdates =
        options.title ||
        options.description ||
        options.start ||
        options.end ||
        options.addAttendee.length > 0 ||
        options.room ||
        options.location ||
        options.teams !== undefined;

      if (!hasUpdates) {
        // Show current event details
        console.log(`\nEvent: ${targetEvent.Subject}`);
        console.log(
          `  When: ${formatDate(targetEvent.Start.DateTime)} ${formatTime(targetEvent.Start.DateTime)} - ${formatTime(targetEvent.End.DateTime)}`
        );
        if (targetEvent.Location?.DisplayName) {
          console.log(`  Location: ${targetEvent.Location.DisplayName}`);
        }
        if (targetEvent.Attendees && targetEvent.Attendees.length > 0) {
          console.log('  Attendees:');
          for (const a of targetEvent.Attendees) {
            const typeLabel = a.Type === 'Resource' ? ' (Room)' : '';
            console.log(`    - ${a.EmailAddress?.Address}${typeLabel}`);
          }
        }
        console.log('\nUse options like --title, --add-attendee, --room to update.');
        return;
      }

      // Build update payload
      const updateOptions: Parameters<typeof updateEvent>[0] = {
        token: authResult.token!,
        eventId: targetEvent.Id,
        changeKey: targetEvent.ChangeKey,
        mailbox: options.mailbox
      };

      if (options.title) {
        updateOptions.subject = options.title;
      }

      if (options.description) {
        updateOptions.body = options.description;
      }

      // Handle time changes
      if (options.start || options.end) {
        const eventDate = new Date(targetEvent.Start.DateTime);

        if (options.start) {
          const newStart = parseTimeToDate(options.start, eventDate);
          updateOptions.start = toLocalISOString(newStart);
        }

        if (options.end) {
          const newEnd = parseTimeToDate(options.end, eventDate);
          updateOptions.end = toLocalISOString(newEnd);
        }
      }

      // Handle location
      if (options.location) {
        updateOptions.location = options.location;
      }

      // Handle room
      let roomEmail: string | undefined;
      let roomName: string | undefined;

      if (options.room) {
        if (options.room.includes('@')) {
          roomEmail = options.room;
          roomName = options.room;
        } else {
          let roomsResult = await searchRooms(authResult.token!, options.room);
          if (!roomsResult.ok || !roomsResult.data || roomsResult.data.length === 0) {
            roomsResult = await getRooms(authResult.token!);
          }

          if (roomsResult.ok && roomsResult.data) {
            const found = roomsResult.data.find((r) =>
              options.room ? r.Name.toLowerCase().includes(options.room.toLowerCase()) : false
            );
            if (found) {
              roomEmail = found.Address;
              roomName = found.Name;
            } else {
              console.error(`Room not found: ${options.room}`);
              process.exit(1);
            }
          }
        }

        if (roomName) {
          updateOptions.location = roomName;
        }
      }

      // Handle attendees (merge existing with new)
      if (options.addAttendee.length > 0 || roomEmail) {
        const existingAttendees: Array<{ email: string; name?: string; type: 'Required' | 'Optional' | 'Resource' }> = (
          targetEvent.Attendees || []
        ).map((a) => ({
          email: a.EmailAddress?.Address || '',
          name: a.EmailAddress?.Name,
          type: a.Type as 'Required' | 'Optional' | 'Resource'
        }));

        // Add new attendees
        for (const email of options.addAttendee) {
          if (!existingAttendees.find((a) => a.email.toLowerCase() === email.toLowerCase())) {
            existingAttendees.push({ email, type: 'Required' });
          }
        }

        // Add room if specified
        if (roomEmail) {
          // Remove any existing room
          const withoutRooms = existingAttendees.filter((a) => a.type !== 'Resource');
          withoutRooms.push({ email: roomEmail, name: roomName, type: 'Resource' });
          updateOptions.attendees = withoutRooms;
        } else {
          updateOptions.attendees = existingAttendees;
        }
      }

      // Handle Teams
      if (options.teams !== undefined) {
        updateOptions.isOnlineMeeting = options.teams;
      }

      console.log(`\nUpdating: ${targetEvent.Subject}`);

      const updateResult = await updateEvent(updateOptions);

      if (!updateResult.ok) {
        if (options.json) {
          console.log(JSON.stringify({ error: updateResult.error?.message || 'Failed to update event' }, null, 2));
        } else {
          console.error(`\nError: ${updateResult.error?.message || 'Failed to update event'}`);
        }
        process.exit(1);
      }

      if (options.json) {
        console.log(
          JSON.stringify(
            {
              success: true,
              event: {
                id: updateResult.data?.Id,
                subject: updateResult.data?.Subject,
                start: updateResult.data?.Start.DateTime,
                end: updateResult.data?.End.DateTime
              }
            },
            null,
            2
          )
        );
      } else {
        console.log('\n\u2713 Event updated successfully.\n');
        if (updateResult.data) {
          console.log(`  Title: ${updateResult.data.Subject}`);
          console.log(
            `  When:  ${formatDate(updateResult.data.Start.DateTime)} ${formatTime(updateResult.data.Start.DateTime)} - ${formatTime(updateResult.data.End.DateTime)}`
          );
        }
        console.log('');
      }
    }
  );
