import { readFile } from 'node:fs/promises';
import { Command } from 'commander';
import { AttachmentLinkSpecError, parseAttachLinkSpec } from '../lib/attach-link-spec.js';
import { AttachmentPathError, validateAttachmentPath } from '../lib/attachments.js';
import { resolveAuth } from '../lib/auth.js';
import { parseDay, parseTimeToDate, toLocalUnzonedISOString, toUTCISOString } from '../lib/dates.js';
import {
  areRoomsFree,
  createEvent,
  type EmailAttachment,
  getRooms,
  type Recurrence,
  type RecurrencePattern,
  type RecurrenceRange,
  type ReferenceAttachmentInput,
  SENSITIVITY_MAP,
  searchRooms
} from '../lib/ews-client.js';
import { getExchangeBackend } from '../lib/exchange-backend.js';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import {
  findFirstAvailableRoomGraph,
  listGraphRooms,
  resolveRoomDisplayNameToPlace
} from '../lib/graph-places-helpers.js';
import { lookupMimeType } from '../lib/mime-type.js';
import { checkReadOnly } from '../lib/utils.js';
import { createEventViaGraph } from './create-event-graph.js';

function formatTime(dateStr: string): string {
  const date = new Date(dateStr);
  return date.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit', hour12: false });
}

function formatDate(dateStr: string): string {
  const date = new Date(dateStr);
  return date.toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric' });
}

export const createEventCommand = new Command('create-event')
  .description('Create a new calendar event')
  .argument('<title>', 'Event title/subject')
  .argument('[start]', 'Start time (e.g., 13:00, 1pm) - not needed for all-day events')
  .argument('[end]', 'End time (e.g., 14:00, 2pm) - not needed for all-day events')
  .option('--day <day>', 'Day for the event (today, tomorrow, monday-sunday, YYYY-MM-DD)', 'today')
  .option('--timezone <timezone>', 'Timezone for the event (e.g., "Pacific Standard Time")')
  .option('--description <text>', 'Event description/body')
  .option('--attendees <emails>', 'Comma-separated list of attendee emails')
  .option('--room <room>', 'Meeting room (use --list-rooms to see available)')
  .option('--teams', 'Create as Teams meeting')
  .option('--category <name>', 'Category label (repeatable)', (v, acc) => [...acc, v], [] as string[])
  .option('--all-day', 'Create as an all-day event (no time slots)')
  .option('--sensitivity <level>', 'Sensitivity: normal, personal, private, confidential')
  .option('--list-rooms', 'List available meeting rooms')
  .option('--find-room', 'Find an available room for the time slot')
  .option('--repeat <type>', 'Recurrence: daily, weekly, monthly, yearly')
  .option('--every <n>', 'Repeat every N days/weeks/months (default: 1)', '1')
  .option('--days <days>', 'Days of week for weekly recurrence (mon,tue,wed,thu,fri,sat,sun)')
  .option('--until <date>', 'End date for recurrence (YYYY-MM-DD)')
  .option('--count <n>', 'Number of occurrences (alternative to --until)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Use a specific authentication identity (default: default)')
  .option('--mailbox <email>', 'Create event in shared mailbox calendar')
  .option('--attach <files>', 'Attach file(s), comma-separated paths (relative to cwd)')
  .option(
    '--attach-link <spec>',
    'Attach link: "Title|https://url" or bare https URL (repeatable)',
    (v: string, prev: string[]) => [...prev, v],
    [] as string[]
  )
  .action(
    async (
      title: string,
      startTime: string | undefined,
      endTime: string | undefined,
      options: {
        day: string;
        timezone?: string;
        description?: string;
        attendees?: string;
        room?: string;
        teams?: boolean;
        allDay?: boolean;
        sensitivity?: string;
        listRooms?: boolean;
        findRoom?: boolean;
        repeat?: string;
        every?: string;
        days?: string;
        until?: string;
        count?: string;
        json?: boolean;
        token?: string;
        identity?: string;
        mailbox?: string;
        category?: string[];
        attach?: string;
        attachLink?: string[];
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const backend = getExchangeBackend();

      // Handle --list-rooms (Graph when graph/auto, else EWS)
      if (options.listRooms) {
        if (backend === 'graph' || backend === 'auto') {
          const ga = await resolveGraphAuth({
            token: options.token,
            identity: options.identity
          });
          if (ga.success && ga.token) {
            const lr = await listGraphRooms(ga.token);
            if (lr.ok && lr.data !== undefined) {
              const sorted = [...lr.data].sort((a, b) =>
                (a.displayName || '').localeCompare(b.displayName || '', undefined, { sensitivity: 'base' })
              );
              if (options.json) {
                console.log(
                  JSON.stringify(
                    {
                      backend: 'graph',
                      rooms: sorted.map((r) => ({
                        name: r.displayName,
                        email: r.emailAddress
                      }))
                    },
                    null,
                    2
                  )
                );
              } else {
                console.log('\nFetching available meeting rooms (Microsoft Graph)...\n');
                if (sorted.length === 0) {
                  console.log('No meeting rooms returned by Places API (empty list).');
                  console.log('You can still specify a room by email address with --room <email>');
                } else {
                  console.log('Available rooms:');
                  for (const room of sorted) {
                    const em = room.emailAddress?.trim();
                    console.log(em ? `  - ${room.displayName} (${em})` : `  - ${room.displayName}`);
                  }
                }
              }
              return;
            }
            if (backend === 'graph') {
              if (options.json) {
                console.log(JSON.stringify({ error: lr.error?.message || 'No rooms returned', rooms: [] }, null, 2));
              } else {
                console.log(lr.error?.message || 'No meeting rooms found or Places API returned no data.');
                console.log('You can still specify a room by email address with --room <email>');
              }
              return;
            }
          } else if (backend === 'graph') {
            if (options.json) {
              console.log(JSON.stringify({ error: ga.error || 'Graph authentication failed' }, null, 2));
            } else {
              console.error(`Error: ${ga.error || 'Graph authentication failed'}`);
            }
            process.exit(1);
          }
        }

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

        console.log('\nFetching available meeting rooms (EWS)...\n');

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

      // Validate start/end times for non-all-day events
      if (!options.allDay && (!startTime || !endTime)) {
        if (options.json) {
          console.log(JSON.stringify({ error: 'Start and end times are required for non-all-day events' }, null, 2));
        } else {
          console.error('Error: Start and end times are required for non-all-day events');
        }
        process.exit(1);
      }

      // Parse date and times
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

      let start: Date;
      let end: Date;

      if (options.allDay) {
        // For all-day events, use midnight boundaries regardless of provided times
        start = new Date(baseDate);
        start.setHours(0, 0, 0, 0);
        end = new Date(baseDate);
        end.setHours(23, 59, 59, 999);
      } else {
        // For regular events, parse the provided times
        try {
          start = parseTimeToDate(startTime!, baseDate, { throwOnInvalid: true });
        } catch (err) {
          const message = err instanceof Error ? err.message : 'Invalid start time';
          if (options.json) {
            console.log(JSON.stringify({ error: message }, null, 2));
          } else {
            console.error(`Error: ${message}`);
          }
          process.exit(1);
        }

        try {
          end = parseTimeToDate(endTime!, baseDate, { throwOnInvalid: true });
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

      // Parse attendees
      const attendees: Array<{ email: string; name?: string; type?: 'Required' | 'Optional' | 'Resource' }> =
        options.attendees ? options.attendees.split(',').map((e) => ({ email: e.trim() })) : [];

      const roomNeedsEwsLookup = Boolean(options.room && !options.room.includes('@'));

      let roomEmail: string | undefined;
      let roomName: string | undefined;
      let ewsToken: string | undefined;

      if (options.findRoom || options.room) {
        if (backend === 'graph' || backend === 'auto') {
          const ga = await resolveGraphAuth({
            token: options.token,
            identity: options.identity
          });
          if (ga.success && ga.token) {
            if (options.findRoom) {
              if (!options.json) {
                console.log('Searching for available rooms...');
              }
              const fr = await findFirstAvailableRoomGraph(ga.token, start, end);
              if (fr) {
                roomEmail = fr.email;
                roomName = fr.name;
                if (!options.json) {
                  console.log(`Found available room: ${fr.name}`);
                }
              }
            } else if (options.room!.includes('@')) {
              roomEmail = options.room;
              roomName = options.room;
            } else {
              const res = await resolveRoomDisplayNameToPlace(ga.token, options.room!);
              if (res.ok) {
                roomEmail = res.place.emailAddress!.trim();
                roomName = res.place.displayName;
              } else if (backend === 'graph') {
                if (options.json) {
                  console.log(JSON.stringify({ error: res.error }, null, 2));
                } else {
                  console.error(`Error: ${res.error}`);
                }
                process.exit(1);
              }
            }
          } else if (backend === 'graph') {
            if (options.json) {
              console.log(JSON.stringify({ error: ga.error || 'Graph authentication failed' }, null, 2));
            } else {
              console.error(`Error: ${ga.error || 'Graph authentication failed'}`);
            }
            process.exit(1);
          }
        }

        if (backend === 'ews' && options.room?.includes('@')) {
          roomEmail = options.room;
          roomName = options.room;
        }

        const needEwsRoom =
          backend === 'ews'
            ? !!(options.findRoom || roomNeedsEwsLookup)
            : backend === 'auto' && !!(options.findRoom || roomNeedsEwsLookup) && !roomEmail;

        if (needEwsRoom) {
          const ar = await resolveAuth({
            token: options.token,
            identity: options.identity
          });
          if (!ar.success) {
            if (options.json) {
              console.log(JSON.stringify({ error: ar.error }, null, 2));
            } else {
              console.error(`Error: ${ar.error}`);
              console.error('\nCheck your .env file for EWS_CLIENT_ID and EWS_REFRESH_TOKEN.');
            }
            process.exit(1);
          }
          const ewsTok = ar.token!;
          ewsToken = ewsTok;

          if (options.findRoom && !roomEmail) {
            if (!options.json) {
              console.log('Searching for available rooms (EWS)...');
            }

            const roomsResult = await getRooms(ewsTok);

            if (!roomsResult.ok || !roomsResult.data || roomsResult.data.length === 0) {
              console.error('Could not fetch room list.');
            } else {
              const roomEmails = roomsResult.data.map((r) => r.Address);
              const freeMap = await areRoomsFree(ewsTok, roomEmails, start.toISOString(), end.toISOString());

              for (const room of roomsResult.data) {
                if (freeMap.get(room.Address)) {
                  roomEmail = room.Address;
                  roomName = room.Name;
                  if (!options.json) {
                    console.log(`Found available room: ${room.Name}`);
                  }
                  break;
                }
              }

              if (!roomEmail && !options.json) {
                console.log('No available rooms found for this time slot.');
              }
            }
          } else if (options.room && !options.room.includes('@') && !roomEmail) {
            const roomQuery = options.room;
            roomName = roomQuery;
            let roomsResult = await searchRooms(ewsToken!, roomQuery);
            if (!roomsResult.ok || !roomsResult.data || roomsResult.data.length === 0) {
              roomsResult = await getRooms(ewsToken!);
            }

            if (roomsResult.ok && roomsResult.data) {
              const found = roomsResult.data.find((r) => r.Name.toLowerCase().includes(roomQuery.toLowerCase()));
              if (found) {
                roomEmail = found.Address;
                roomName = found.Name;
              }
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

        // Build range — use local date to avoid UTC shift for late-evening events
        const year = start.getFullYear();
        const month = String(start.getMonth() + 1).padStart(2, '0');
        const day = String(start.getDate()).padStart(2, '0');
        const localDate = `${year}-${month}-${day}`;
        const range: RecurrenceRange = {
          Type: 'NoEnd',
          StartDate: localDate
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

      const sensitivity = options.sensitivity ? SENSITIVITY_MAP[options.sensitivity.toLowerCase()] : undefined;

      if (options.sensitivity && !sensitivity) {
        console.error(`Invalid sensitivity: ${options.sensitivity}`);
        process.exit(1);
      }

      const workingDirectory = process.cwd();
      let fileAttachments: EmailAttachment[] | undefined;
      if (options.attach?.trim()) {
        fileAttachments = [];
        const filePaths = options.attach
          .split(',')
          .map((f) => f.trim())
          .filter(Boolean);
        for (const filePath of filePaths) {
          try {
            const validated = await validateAttachmentPath(filePath, workingDirectory);
            const content = await readFile(validated.absolutePath);
            const contentType = lookupMimeType(validated.fileName);
            fileAttachments.push({
              name: validated.fileName,
              contentType,
              contentBytes: content.toString('base64')
            });
            if (!options.json) {
              console.log(`  Attaching file: ${validated.fileName} (${Math.round(validated.size / 1024)} KB)`);
            }
          } catch (err) {
            console.error(`Failed to read attachment: ${filePath}`);
            if (err instanceof AttachmentPathError) {
              console.error(err.message);
            } else {
              console.error(err instanceof Error ? err.message : 'Unknown error');
            }
            process.exit(1);
          }
        }
      }

      let referenceAttachments: ReferenceAttachmentInput[] | undefined;
      const linkSpecs = options.attachLink ?? [];
      if (linkSpecs.length > 0) {
        referenceAttachments = [];
        for (const spec of linkSpecs) {
          try {
            const { name, url } = parseAttachLinkSpec(spec);
            referenceAttachments.push({ name, url, contentType: 'text/html' });
            if (!options.json) {
              console.log(`  Attaching link: ${name}`);
            }
          } catch (err) {
            const msg =
              err instanceof AttachmentLinkSpecError ? err.message : err instanceof Error ? err.message : String(err);
            console.error(`Invalid --attach-link: ${msg}`);
            process.exit(1);
          }
        }
      }

      const tryGraphFirst = backend === 'graph' || backend === 'auto';

      if (tryGraphFirst) {
        const ga = await resolveGraphAuth({
          token: options.token,
          identity: options.identity
        });
        if (ga.success && ga.token) {
          const gr = await createEventViaGraph({
            token: ga.token,
            mailbox: options.mailbox,
            subject: title,
            body: options.description,
            start,
            end,
            allDay: options.allDay ?? false,
            timezoneName: options.timezone,
            attendees,
            teams: options.teams ?? false,
            locationDisplay: roomName,
            sensitivity,
            categories: options.category && options.category.length > 0 ? options.category : undefined,
            recurrence,
            fileAttachments,
            referenceAttachments: referenceAttachments?.map((a) => ({ name: a.name, sourceUrl: a.url }))
          });

          if (gr.ok) {
            const ev = gr.event;
            const joinUrl = ev.onlineMeeting?.joinUrl;
            const hasPartialSuccess = 'partialSuccess' in gr && gr.partialSuccess;
            if (options.json) {
              console.log(
                JSON.stringify(
                  {
                    success: true,
                    backend: 'graph',
                    ...(hasPartialSuccess ? { warning: gr.attachmentError } : {}),
                    event: {
                      id: ev.id,
                      changeKey: ev.changeKey,
                      subject: ev.subject,
                      start: ev.start?.dateTime,
                      end: ev.end?.dateTime,
                      webLink: ev.webLink,
                      onlineMeetingUrl: joinUrl,
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
            console.log('  Created via: Microsoft Graph');
            console.log(`  Title: ${ev.subject ?? title}`);
            const st = ev.start?.dateTime ?? '';
            const en = ev.end?.dateTime ?? '';
            if (st && en) {
              console.log(`  When:  ${formatDate(st)} ${formatTime(st)} - ${formatTime(en)}`);
            }
            if (roomName) {
              console.log(`  Room:  ${roomName}`);
            }
            if (attendees.length > 0) {
              const nonRoomAttendees = attendees.filter((a) => a.type !== 'Resource');
              if (nonRoomAttendees.length > 0) {
                console.log(`  Attendees: ${nonRoomAttendees.map((a) => a.email).join(', ')}`);
              }
            }
            if (joinUrl) {
              console.log(`  Teams: ${joinUrl}`);
            }
            if (ev.webLink) {
              console.log(`  Link:  ${ev.webLink}`);
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
            if (hasPartialSuccess) {
              console.log(`\n  Warning: ${gr.attachmentError}`);
            }
            console.log();
            return;
          }
          if (backend === 'graph') {
            if (options.json) {
              console.log(JSON.stringify({ error: gr.error }, null, 2));
            } else {
              console.error(`Error: ${gr.error}`);
            }
            process.exit(1);
          }
          if (!options.json) {
            console.warn(`[create-event] Graph failed (${gr.error}); falling back to EWS.`);
          }
        } else if (backend === 'graph') {
          if (options.json) {
            console.log(JSON.stringify({ error: ga.error || 'Graph authentication failed' }, null, 2));
          } else {
            console.error(`Error: ${ga.error || 'Graph authentication failed'}`);
          }
          process.exit(1);
        }
      }

      if (!ewsToken) {
        const ar = await resolveAuth({
          token: options.token,
          identity: options.identity
        });
        if (!ar.success) {
          if (options.json) {
            console.log(JSON.stringify({ error: ar.error }, null, 2));
          } else {
            console.error(`Error: ${ar.error}`);
            console.error('\nCheck your .env file for EWS_CLIENT_ID and EWS_REFRESH_TOKEN.');
          }
          process.exit(1);
        }
        ewsToken = ar.token;
      }

      // Create the event (EWS)
      const result = await createEvent({
        token: ewsToken!,
        subject: title,
        start: options.timezone ? toLocalUnzonedISOString(start) : toUTCISOString(start),
        end: options.timezone ? toLocalUnzonedISOString(end) : toUTCISOString(end),
        body: options.description,
        location: roomName,
        attendees: attendees.length > 0 ? attendees : undefined,
        isOnlineMeeting: options.teams,
        isAllDay: options.allDay,
        sensitivity,
        recurrence,
        mailbox: options.mailbox,
        timezone: options.timezone,
        categories: options.category && options.category.length > 0 ? options.category : undefined,
        fileAttachments,
        referenceAttachments
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
              backend: 'ews',
              event: {
                id: result.data.Id,
                changeKey: result.data.ChangeKey,
                subject: result.data.Subject,
                start: result.data.Start.DateTime,
                end: result.data.End.DateTime,
                webLink: result.data.WebLink,
                onlineMeetingUrl: result.data.OnlineMeetingUrl,
                fileAttachments: fileAttachments?.length ?? 0,
                referenceAttachments: referenceAttachments?.length ?? 0,
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
