import { Command } from 'commander';
import { resolveAuth } from '../lib/auth.js';
import { parseDay, parseTimeToDate, toUTCISOString, toLocalUnzonedISOString } from '../lib/dates.js';
import { getCalendarEvents, getRooms, searchRooms, updateEvent } from '../lib/ews-client.js';

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
  .option(
    '--day <day>',
    'Day to show events from (today, tomorrow, YYYY-MM-DD) - note: may miss multi-day events crossing midnight',
    'today'
  )
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
  .option(
    '--remove-attendee <email>',
    'Remove an attendee by email (can be used multiple times)',
    (val, arr: string[]) => [...arr, val],
    []
  )
  .option('--room <room>', 'Set/change meeting room (name or email)')
  .option('--location <text>', 'Set location text')
  .option('--timezone <timezone>', 'Timezone for the event (e.g., "Pacific Standard Time")')
  .option('--occurrence <index>', 'Update only the Nth occurrence of a recurring event')
  .option('--instance <date>', 'Update only the occurrence on a specific date (YYYY-MM-DD)')
  .option('--teams', 'Make it a Teams meeting')
  .option('--no-teams', 'Remove Teams meeting')
  .option('--all-day', 'Mark as an all-day event')
  .option('--no-all-day', 'Remove all-day flag')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--mailbox <email>', 'Update event in shared mailbox calendar')
  .action(
    async (
      _eventIndex: string | undefined,
      options: {
        id?: string;
        day: string;
        timezone?: string;
        title?: string;
        description?: string;
        start?: string;
        end?: string;
        addAttendee: string[];
        removeAttendee: string[];
        room?: string;
        location?: string;
        occurrence?: string;
        instance?: string;
        teams?: boolean;
        allDay?: boolean;
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
      let occurrenceItemId: string | undefined;
      let displayEvent = targetEvent;

      if (options.occurrence || options.instance) {
        // Find the specific occurrence, ensuring it matches the provided event ID
        if (options.instance) {
          let instanceDate: Date;
          try {
            instanceDate = parseDay(options.instance, { throwOnInvalid: true });
          } catch (err) {
            const message = err instanceof Error ? err.message : 'Invalid instance date';
            if (options.json) {
              console.log(JSON.stringify({ error: message }, null, 2));
            } else {
              console.error(`Error: ${message}`);
            }
            process.exit(1);
          }
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
          occurrenceItemId = occEvent.Id;
          displayEvent = occEvent;
          console.log(`\nUpdating single occurrence of: ${occEvent.Subject}`);
          console.log(
            `  ${formatDate(occEvent.Start.DateTime)} ${formatTime(occEvent.Start.DateTime)} - ${formatTime(occEvent.End.DateTime)}`
          );
        } else if (options.occurrence) {
          const idx = parseInt(options.occurrence, 10);
          if (Number.isNaN(idx) || idx < 1 || idx > events.length) {
            console.error(`Invalid --occurrence index: ${options.occurrence}. Valid range: 1-${events.length}.`);
            process.exit(1);
          }
          const occEvent = events[idx - 1];
          if (occEvent.Id !== options.id) {
            console.error(`Occurrence ${idx} does not match the provided event ID ${options.id}.`);
            process.exit(1);
          }
          occurrenceItemId = occEvent.Id;
          displayEvent = occEvent;
          console.log(`\nUpdating occurrence ${idx} of: ${occEvent.Subject}`);
          console.log(
            `  ${formatDate(occEvent.Start.DateTime)} ${formatTime(occEvent.Start.DateTime)} - ${formatTime(occEvent.End.DateTime)}`
          );
        }
      } else if (!targetEvent) {
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
        options.removeAttendee.length > 0 ||
        options.room ||
        options.location ||
        options.teams !== undefined ||
        options.allDay !== undefined;

      if (!hasUpdates) {
        // Show current event details
        console.log(`\nEvent: ${displayEvent!.Subject}`);
        console.log(
          `  When: ${formatDate(displayEvent!.Start.DateTime)} ${formatTime(displayEvent!.Start.DateTime)} - ${formatTime(displayEvent!.End.DateTime)}`
        );
        if (displayEvent!.Location?.DisplayName) {
          console.log(`  Location: ${displayEvent!.Location.DisplayName}`);
        }
        if (displayEvent!.Attendees && displayEvent!.Attendees.length > 0) {
          console.log('  Attendees:');
          for (const a of displayEvent!.Attendees) {
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
        eventId: targetEvent ? targetEvent.Id : displayEvent!.Id,
        changeKey: displayEvent!.ChangeKey,
        occurrenceItemId,
        mailbox: options.mailbox
      };

      if (options.title) {
        updateOptions.subject = options.title;
      }

      if (options.timezone) {
        updateOptions.timezone = options.timezone;
      }

      if (options.description) {
        updateOptions.body = options.description;
      }

      // Handle time changes
      if (options.start || options.end) {
        const eventDate = new Date(displayEvent!.Start.DateTime);

        if (options.start) {
          try {
            const newStart = parseTimeToDate(options.start, eventDate, { throwOnInvalid: true });
            updateOptions.start = options.timezone ? toLocalUnzonedISOString(newStart) : toUTCISOString(newStart);
          } catch (err) {
            const message = err instanceof Error ? err.message : 'Invalid start time';
            if (options.json) {
              console.log(JSON.stringify({ error: message }, null, 2));
            } else {
              console.error(`Error: ${message}`);
            }
            process.exit(1);
          }
        }

        if (options.end) {
          try {
            const newEnd = parseTimeToDate(options.end, eventDate, { throwOnInvalid: true });
            updateOptions.end = options.timezone ? toLocalUnzonedISOString(newEnd) : toUTCISOString(newEnd);
          } catch (err) {
            const message = err instanceof Error ? err.message : 'Invalid end time';
            if (options.json) {
              console.log(JSON.stringify({ error: message }, null, 2));
            } else {
              console.error(`Error: ${message}`);
            }
            process.exit(1);
          }
        }
      }

      // Handle location
      if (options.location) {
        updateOptions.location = options.location;
      }

      // Handle all-day
      if (options.allDay !== undefined) {
        updateOptions.isAllDay = options.allDay;
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
      // NOTE: updateEvent replaces the entire attendee list via EWS SetItemField.
      // Concurrent edits (e.g., removing an attendee via OWA between fetch and update)
      // can be overwritten. This is a known EWS limitation.
      if (options.addAttendee.length > 0 || options.removeAttendee.length > 0 || roomEmail) {
        const existingAttendees: Array<{ email: string; name?: string; type: 'Required' | 'Optional' | 'Resource' }> = (
          displayEvent!.Attendees || []
        ).map((a) => ({
          email: a.EmailAddress?.Address || '',
          name: a.EmailAddress?.Name,
          type: a.Type as 'Required' | 'Optional' | 'Resource'
        }));

        // Remove attendees specified via --remove-attendee
        for (const email of options.removeAttendee) {
          const idx = existingAttendees.findIndex((a) => a.email.toLowerCase() === email.toLowerCase());
          if (idx !== -1) existingAttendees.splice(idx, 1);
        }

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

      console.log(`\nUpdating: ${displayEvent!.Subject}`);

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
