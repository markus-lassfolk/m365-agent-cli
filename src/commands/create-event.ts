import { Command } from 'commander';
import { resolveAuth } from '../lib/auth.js';
import {
  createEvent,
  getRooms,
  searchRooms,
  isRoomFree,
  type Recurrence,
  type RecurrencePattern,
  type RecurrenceRange
} from '../lib/ews-client.js';

function formatTime(dateStr: string): string {
  const date = new Date(dateStr);
  return date.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit', hour12: false });
}

function formatDate(dateStr: string): string {
  const date = new Date(dateStr);
  return date.toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric' });
}

function parseTimeToDate(timeStr: string, baseDate: Date = new Date()): Date {
  const result = new Date(baseDate);

  // Handle HH:MM format
  const timeMatch = timeStr.match(/^(\d{1,2}):(\d{2})$/);
  if (timeMatch) {
    result.setHours(parseInt(timeMatch[1], 10), parseInt(timeMatch[2], 10), 0, 0);
    return result;
  }

  // Handle "1pm", "13:00", etc.
  const hourMatch = timeStr.match(/^(\d{1,2})(am|pm)?$/i);
  if (hourMatch) {
    let hour = parseInt(hourMatch[1], 10);
    const isPM = hourMatch[2]?.toLowerCase() === 'pm';
    if (isPM && hour < 12) hour += 12;
    if (!isPM && hour === 12) hour = 0;
    result.setHours(hour, 0, 0, 0);
    return result;
  }

  return result;
}

function toLocalISOString(date: Date): string {
  // Format as YYYY-MM-DDTHH:mm:ss without timezone offset
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  const seconds = String(date.getSeconds()).padStart(2, '0');
  return `${year}-${month}-${day}T${hours}:${minutes}:${seconds}`;
}

function parseDay(day: string): Date {
  const now = new Date();

  switch (day.toLowerCase()) {
    case 'today':
      return now;
    case 'tomorrow':
      now.setDate(now.getDate() + 1);
      return now;
    case 'monday':
    case 'tuesday':
    case 'wednesday':
    case 'thursday':
    case 'friday':
    case 'saturday':
    case 'sunday': {
      const days = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];
      const targetDay = days.indexOf(day.toLowerCase());
      const currentDay = now.getDay();
      let diff = targetDay - currentDay;
      if (diff <= 0) diff += 7;
      now.setDate(now.getDate() + diff);
      return now;
    }
    default: {
      const parsed = new Date(day);
      return Number.isNaN(parsed.getTime()) ? now : parsed;
    }
  }
}

export const createEventCommand = new Command('create-event')
  .description('Create a new calendar event')
  .argument('<title>', 'Event title/subject')
  .argument('<start>', 'Start time (e.g., 13:00, 1pm)')
  .argument('<end>', 'End time (e.g., 14:00, 2pm)')
  .option('--day <day>', 'Day for the event (today, tomorrow, monday-sunday, YYYY-MM-DD)', 'today')
  .option('--description <text>', 'Event description/body')
  .option('--attendees <emails>', 'Comma-separated list of attendee emails')
  .option('--room <room>', 'Meeting room (use --list-rooms to see available)')
  .option('--teams', 'Create as Teams meeting')
  .option('--list-rooms', 'List available meeting rooms')
  .option('--find-room', 'Find an available room for the time slot')
  .option('--repeat <type>', 'Recurrence: daily, weekly, monthly, yearly')
  .option('--every <n>', 'Repeat every N days/weeks/months (default: 1)', '1')
  .option('--days <days>', 'Days of week for weekly recurrence (mon,tue,wed,thu,fri,sat,sun)')
  .option('--until <date>', 'End date for recurrence (YYYY-MM-DD)')
  .option('--count <n>', 'Number of occurrences (alternative to --until)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .action(
    async (
      title: string,
      startTime: string,
      endTime: string,
      options: {
        day: string;
        description?: string;
        attendees?: string;
        room?: string;
        teams?: boolean;
        listRooms?: boolean;
        findRoom?: boolean;
        repeat?: string;
        every?: string;
        days?: string;
        until?: string;
        count?: string;
        json?: boolean;
        token?: string;
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

      // Handle --list-rooms
      if (options.listRooms) {
        console.log('\nFetching available meeting rooms...\n');

        // Search with multiple queries to find more rooms
        const allRooms = new Map<string, { Name: string; Address: string }>();
        const queries = ['room', 'meeting', 'vergader', 'nv-', 'conference'];

        for (const q of queries) {
          const result = await searchRooms(authResult.token!, q);
          if (result.ok && result.data) {
            for (const room of result.data) {
              if (!allRooms.has(room.Address)) {
                allRooms.set(room.Address, room);
              }
            }
          }
        }

        if (allRooms.size > 0) {
          console.log('Available rooms:');
          const sortedRooms = [...allRooms.values()].sort((a, b) => a.Name.localeCompare(b.Name));
          for (const room of sortedRooms) {
            console.log(`  - ${room.Name} (${room.Address})`);
          }
          return;
        }

        // Fallback to REST API
        const roomsResult = await getRooms(authResult.token!);
        if (roomsResult.ok && roomsResult.data && roomsResult.data.length > 0) {
          console.log('Available rooms:');
          for (const room of roomsResult.data) {
            console.log(`  - ${room.Name} (${room.Address})`);
          }
        } else {
          console.log("No meeting rooms found or you don't have access to room lists.");
          console.log('You can still specify a room by email address with --room <email>');
        }
        return;
      }

      // Parse date and times
      const baseDate = parseDay(options.day);
      const start = parseTimeToDate(startTime, baseDate);
      const end = parseTimeToDate(endTime, baseDate);

      // Parse attendees
      const attendees: Array<{ email: string; name?: string; type?: 'Required' | 'Optional' | 'Resource' }> =
        options.attendees ? options.attendees.split(',').map((e) => ({ email: e.trim() })) : [];

      // Handle --find-room: find an available room
      let roomEmail: string | undefined;
      let roomName: string | undefined;

      if (options.findRoom) {
        console.log('Searching for available rooms...');

        const roomsResult = await getRooms(authResult.token!);

        if (!roomsResult.ok || !roomsResult.data || roomsResult.data.length === 0) {
          console.error('Could not fetch room list.');
        } else {
          for (const room of roomsResult.data) {
            const free = await isRoomFree(authResult.token!, room.Address, start.toISOString(), end.toISOString());

            if (free) {
              roomEmail = room.Address;
              roomName = room.Name;
              console.log(`Found available room: ${room.Name}`);
              break;
            }
          }

          if (!roomEmail) {
            console.log('No available rooms found for this time slot.');
          }
        }
      } else if (options.room) {
        // User specified a room - could be name or email
        roomName = options.room;
        // If it looks like an email, use it directly
        if (options.room.includes('@')) {
          roomEmail = options.room;
        } else {
          // Try to find the room by name - search for it
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
            }
          }
        }
      }

      // Add room as attendee if found
      if (roomEmail) {
        attendees.push({ email: roomEmail, name: roomName, type: 'Resource' });
      }

      // Build recurrence if specified
      let recurrence: Recurrence | undefined;
      if (options.repeat) {
        const dayMap: Record<string, string> = {
          mon: 'Monday',
          monday: 'Monday',
          tue: 'Tuesday',
          tuesday: 'Tuesday',
          wed: 'Wednesday',
          wednesday: 'Wednesday',
          thu: 'Thursday',
          thursday: 'Thursday',
          fri: 'Friday',
          friday: 'Friday',
          sat: 'Saturday',
          saturday: 'Saturday',
          sun: 'Sunday',
          sunday: 'Sunday'
        };

        const patternTypeMap: Record<string, RecurrencePattern['Type']> = {
          daily: 'Daily',
          weekly: 'Weekly',
          monthly: 'AbsoluteMonthly',
          yearly: 'AbsoluteYearly'
        };

        const patternType = patternTypeMap[options.repeat.toLowerCase()];
        if (!patternType) {
          console.error(`Invalid repeat type: ${options.repeat}`);
          console.error('Valid options: daily, weekly, monthly, yearly');
          process.exit(1);
        }

        const pattern: RecurrencePattern = {
          Type: patternType,
          Interval: parseInt(options.every || '1', 10) || 1
        };

        // For weekly recurrence, add days of week
        if (patternType === 'Weekly') {
          if (options.days) {
            const days = options.days
              .split(',')
              .map((d) => dayMap[d.trim().toLowerCase()])
              .filter(Boolean);
            if (days.length > 0) {
              pattern.DaysOfWeek = days;
            }
          } else {
            // Default to the day of the event
            const dayOfWeek = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'][
              start.getDay()
            ];
            pattern.DaysOfWeek = [dayOfWeek];
          }
        }

        // For monthly, use day of month
        if (patternType === 'AbsoluteMonthly') {
          pattern.DayOfMonth = start.getDate();
        }

        // For yearly, use month and day
        if (patternType === 'AbsoluteYearly') {
          pattern.Month = start.getMonth() + 1;
          pattern.DayOfMonth = start.getDate();
        }

        // Build range
        const range: RecurrenceRange = {
          Type: 'NoEnd',
          StartDate: start.toISOString().split('T')[0]
        };

        if (options.until) {
          range.Type = 'EndDate';
          range.EndDate = options.until;
        } else if (options.count) {
          range.Type = 'Numbered';
          range.NumberOfOccurrences = parseInt(options.count, 10);
        }

        recurrence = { Pattern: pattern, Range: range };
      }

      // Create the event
      const result = await createEvent({
        token: authResult.token!,
        subject: title,
        start: toLocalISOString(start),
        end: toLocalISOString(end),
        body: options.description,
        location: roomName,
        attendees: attendees.length > 0 ? attendees : undefined,
        isOnlineMeeting: options.teams,
        recurrence
      });

      if (!result.ok || !result.data) {
        if (options.json) {
          console.log(JSON.stringify({ error: result.error?.message || 'Failed to create event' }, null, 2));
        } else {
          console.error(`Error: ${result.error?.message || 'Failed to create event'}`);
        }
        process.exit(1);
      }

      if (options.json) {
        console.log(
          JSON.stringify(
            {
              success: true,
              event: {
                id: result.data.Id,
                subject: result.data.Subject,
                start: result.data.Start.DateTime,
                end: result.data.End.DateTime,
                webLink: result.data.WebLink,
                onlineMeetingUrl: result.data.OnlineMeetingUrl,
                recurring: !!recurrence,
                recurrence: recurrence
                  ? {
                      type: recurrence.Pattern.Type,
                      interval: recurrence.Pattern.Interval,
                      daysOfWeek: recurrence.Pattern.DaysOfWeek,
                      endType: recurrence.Range.Type,
                      endDate: recurrence.Range.EndDate,
                      occurrences: recurrence.Range.NumberOfOccurrences
                    }
                  : undefined
              }
            },
            null,
            2
          )
        );
        return;
      }

      console.log('\n\u2713 Event created successfully!\n');
      console.log(`  Title: ${result.data.Subject}`);
      console.log(
        `  When:  ${formatDate(result.data.Start.DateTime)} ${formatTime(result.data.Start.DateTime)} - ${formatTime(result.data.End.DateTime)}`
      );

      if (roomName) {
        console.log(`  Room:  ${roomName}`);
      }

      if (attendees.length > 0) {
        const nonRoomAttendees = attendees.filter((a) => a.type !== 'Resource');
        if (nonRoomAttendees.length > 0) {
          console.log(`  Attendees: ${nonRoomAttendees.map((a) => a.email).join(', ')}`);
        }
      }

      if (result.data.OnlineMeetingUrl) {
        console.log(`  Teams: ${result.data.OnlineMeetingUrl}`);
      }

      if (result.data.WebLink) {
        console.log(`  Link:  ${result.data.WebLink}`);
      }

      if (recurrence) {
        let recurrenceDesc = `Every ${recurrence.Pattern.Interval > 1 ? `${recurrence.Pattern.Interval} ` : ''}`;
        switch (recurrence.Pattern.Type) {
          case 'Daily':
            recurrenceDesc += recurrence.Pattern.Interval > 1 ? 'days' : 'day';
            break;
          case 'Weekly':
            recurrenceDesc += recurrence.Pattern.Interval > 1 ? 'weeks' : 'week';
            if (recurrence.Pattern.DaysOfWeek) {
              recurrenceDesc += ` on ${recurrence.Pattern.DaysOfWeek.join(', ')}`;
            }
            break;
          case 'AbsoluteMonthly':
            recurrenceDesc += recurrence.Pattern.Interval > 1 ? 'months' : 'month';
            break;
          case 'AbsoluteYearly':
            recurrenceDesc += recurrence.Pattern.Interval > 1 ? 'years' : 'year';
            break;
        }
        if (recurrence.Range.Type === 'EndDate' && recurrence.Range.EndDate) {
          recurrenceDesc += ` until ${recurrence.Range.EndDate}`;
        } else if (recurrence.Range.Type === 'Numbered' && recurrence.Range.NumberOfOccurrences) {
          recurrenceDesc += ` (${recurrence.Range.NumberOfOccurrences} occurrences)`;
        }
        console.log(`  Repeat: ${recurrenceDesc}`);
      }

      console.log();
    }
  );
