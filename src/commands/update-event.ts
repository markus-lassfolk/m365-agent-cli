import { readFile } from 'node:fs/promises';
import { Command } from 'commander';
import { AttachmentLinkSpecError, parseAttachLinkSpec } from '../lib/attach-link-spec.js';
import { AttachmentPathError, validateAttachmentPath } from '../lib/attachments.js';
import { resolveAuth } from '../lib/auth.js';
import {
  graphDayRangeIso,
  graphEventMatchesOccurrenceFilter,
  graphFilterOrganizerEvents,
  graphGetMailboxOrMeEmail
} from '../lib/calendar-graph-helpers.js';
import { parseDay, parseTimeToDate, toLocalUnzonedISOString, toUTCISOString } from '../lib/dates.js';
import {
  addCalendarEventAttachments,
  type CalendarEvent,
  type EmailAttachment,
  getCalendarEvent,
  getCalendarEvents,
  getRooms,
  type ReferenceAttachmentInput,
  SENSITIVITY_MAP,
  searchRooms,
  updateEvent
} from '../lib/ews-client.js';
import { getExchangeBackend } from '../lib/exchange-backend.js';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import {
  addCalendarEventAttachmentsGraph,
  type GraphCalendarEvent,
  getEvent,
  listCalendarView,
  updateCalendarEvent
} from '../lib/graph-calendar-client.js';
import { resolveRoomDisplayNameToPlace } from '../lib/graph-places-helpers.js';
import { lookupMimeType } from '../lib/mime-type.js';
import { checkReadOnly } from '../lib/utils.js';
import { buildGraphUpdatePatch } from './update-event-graph.js';

/** Shown when Graph cannot resolve an event id (often mixed EWS vs Microsoft Graph ids). */
const GRAPH_EVENT_ID_HINT =
  'With M365_EXCHANGE_BACKEND=graph, use event ids from Graph-backed listing (`calendar`, `respond list`). EWS-format ids will not load.';

function formatTime(dateStr: string): string {
  const date = new Date(dateStr);
  return date.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit', hour12: false });
}

function formatDate(dateStr: string): string {
  const date = new Date(dateStr);
  return date.toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric' });
}

function graphStartDt(e: GraphCalendarEvent): string {
  return e.start?.dateTime ?? '';
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
  .option('--search <text>', 'Search for events by title')
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
  .option('--category <name>', 'Category label (repeatable)', (v, acc) => [...acc, v], [] as string[])
  .option('--clear-categories', 'Clear all categories')
  .option('--sensitivity <level>', 'Set sensitivity: normal, personal, private, confidential')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Use a specific authentication identity (default: default)')
  .option('--mailbox <email>', 'Update event in shared mailbox calendar')
  .option('--attach <files>', 'Add file attachment(s), comma-separated paths (relative to cwd)')
  .option(
    '--attach-link <spec>',
    'Add link attachment: "Title|https://url" or bare https URL (repeatable)',
    (v: string, prev: string[]) => [...prev, v],
    [] as string[]
  )
  .action(
    async (
      _eventIndex: string | undefined,
      options: {
        id?: string;
        day: string;
        search?: string;
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
        sensitivity?: string;
        json?: boolean;
        token?: string;
        identity?: string;
        mailbox?: string;
        category?: string[];
        clearCategories?: boolean;
        attach?: string;
        attachLink?: string[];
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const backend = getExchangeBackend();
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
      const graphRange = graphDayRangeIso(baseDate);

      const tryGraphFirst = backend === 'graph' || backend === 'auto';

      let eventsGraph: GraphCalendarEvent[] | undefined;
      let graphToken: string | undefined;
      let authResult: Awaited<ReturnType<typeof resolveAuth>> | undefined;

      if (tryGraphFirst) {
        const ga = await resolveGraphAuth({
          token: options.token,
          identity: options.identity
        });
        if (ga.success && ga.token) {
          const lv = await listCalendarView(ga.token, graphRange.start, graphRange.end, { user: options.mailbox });
          if (lv.ok && lv.data) {
            const me = await graphGetMailboxOrMeEmail(ga.token, options.mailbox);
            if (me) {
              let evs = graphFilterOrganizerEvents(lv.data, me);
              if (options.search) {
                const searchLower = options.search.toLowerCase();
                evs = evs.filter((e) => e.subject?.toLowerCase().includes(searchLower));
              }
              eventsGraph = evs;
              graphToken = ga.token;
            } else if (backend === 'graph') {
              if (options.json) {
                console.log(JSON.stringify({ error: 'Failed to determine user email' }, null, 2));
              } else {
                console.error('Error: Failed to determine user email');
              }
              process.exit(1);
            } else if (!options.json) {
              console.warn(
                '[update-event] Graph did not return a mailbox identity (/me); falling back to EWS for listing.'
              );
            }
          } else if (backend === 'graph') {
            if (options.json) {
              console.log(JSON.stringify({ error: lv.error?.message || 'Failed to list calendar' }, null, 2));
            } else {
              console.error(`Error: ${lv.error?.message || 'Failed to list calendar'}`);
            }
            process.exit(1);
          } else if (!options.json) {
            console.warn(`[update-event] Graph list failed (${lv.error?.message}); falling back to EWS.`);
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

      let eventsEws: CalendarEvent[] | undefined;
      if (!eventsGraph) {
        authResult = await resolveAuth({
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

        let ev = result.data.filter((e) => e.IsOrganizer && !e.IsCancelled);
        if (options.search) {
          const searchLower = options.search.toLowerCase();
          ev = ev.filter((e) => e.Subject?.toLowerCase().includes(searchLower));
        }
        eventsEws = ev;
      }

      const useGraph = !!eventsGraph && !!graphToken;
      const events = useGraph ? eventsGraph! : eventsEws!;

      // If no id provided, list events
      if (!options.id) {
        if (options.json) {
          if (useGraph) {
            console.log(
              JSON.stringify(
                {
                  backend: 'graph',
                  events: events.map((e, i) => ({
                    index: i + 1,
                    id: (e as GraphCalendarEvent).id,
                    subject: (e as GraphCalendarEvent).subject,
                    start: graphStartDt(e as GraphCalendarEvent),
                    end: (e as GraphCalendarEvent).end?.dateTime
                  }))
                },
                null,
                2
              )
            );
          } else {
            console.log(
              JSON.stringify(
                {
                  backend: 'ews',
                  events: (events as CalendarEvent[]).map((e, i) => ({
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
          }
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
          if (useGraph) {
            const ge = event as GraphCalendarEvent;
            const st = graphStartDt(ge);
            const en = ge.end?.dateTime ?? '';
            console.log(`\n  [${i + 1}] ${ge.subject ?? '(no subject)'}`);
            console.log(`      ${formatTime(st)} - ${formatTime(en)}`);
            console.log(`      ID: ${ge.id}`);
            if (ge.location?.displayName) {
              console.log(`      Location: ${ge.location.displayName}`);
            }
            if (ge.attendees && ge.attendees.length > 0) {
              const attendeeList = ge.attendees
                .filter((a) => (a as { type?: string }).type !== 'resource')
                .map((a) => a.emailAddress?.address)
                .filter(Boolean);
              if (attendeeList.length > 0) {
                console.log(`      Attendees: ${attendeeList.join(', ')}`);
              }
            }
          } else {
            const e = event as CalendarEvent;
            const startTime = formatTime(e.Start.DateTime);
            const endTime = formatTime(e.End.DateTime);

            console.log(`\n  [${i + 1}] ${e.Subject}`);
            console.log(`      ${startTime} - ${endTime}`);
            console.log(`      ID: ${e.Id}`);
            if (e.Location?.DisplayName) {
              console.log(`      Location: ${e.Location.DisplayName}`);
            }
            if (e.Attendees && e.Attendees.length > 0) {
              const attendeeList = e.Attendees.filter((a) => a.Type !== 'Resource')
                .map((a) => a.EmailAddress?.Address)
                .filter(Boolean);
              if (attendeeList.length > 0) {
                console.log(`      Attendees: ${attendeeList.join(', ')}`);
              }
            }
          }
        }

        console.log(`\n${'\u2500'.repeat(60)}`);
        console.log('\nTo update an event:');
        console.log('  m365-agent-cli update-event --id <id> --title "New Title"');
        console.log('  m365-agent-cli update-event --id <id> --add-attendee user@example.com');
        console.log('  m365-agent-cli update-event --id <id> --room "Taxi"');
        console.log('  m365-agent-cli update-event --id <id> --start 14:00 --end 15:00');
        console.log('');
        return;
      }

      let targetGraph: GraphCalendarEvent | undefined;
      let targetEws: CalendarEvent | undefined;
      let occurrenceItemId: string | undefined;
      let displayEws: CalendarEvent | undefined;

      if (useGraph) {
        targetGraph = (events as GraphCalendarEvent[]).find((e) => e.id === options.id);
        if (!targetGraph && options.id) {
          targetGraph = (events as GraphCalendarEvent[]).find((e) => graphEventMatchesOccurrenceFilter(e, options.id!));
        }
        if (!targetGraph && graphToken && options.id) {
          const fetched = await getEvent(graphToken, options.id, options.mailbox);
          if (fetched.ok && fetched.data) {
            targetGraph = fetched.data;
          }
        }
        if (!targetGraph) {
          if (options.json) {
            console.log(
              JSON.stringify({ error: `Invalid event id: ${options.id}`, hint: GRAPH_EVENT_ID_HINT }, null, 2)
            );
          } else {
            console.error(`Invalid event id: ${options.id}`);
            console.error(GRAPH_EVENT_ID_HINT);
          }
          process.exit(1);
        }
        if ((options.occurrence || options.instance) && targetGraph) {
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
            const occEvent = (events as GraphCalendarEvent[]).find((e) => {
              const eventDate = new Date(graphStartDt(e));
              eventDate.setHours(0, 0, 0, 0);
              return (
                eventDate.getTime() === instanceDate.getTime() && graphEventMatchesOccurrenceFilter(e, options.id!)
              );
            });
            if (!occEvent) {
              console.error(
                `No occurrence found on ${options.instance} with ID ${options.id}. Try expanding the date range with --day.`
              );
              process.exit(1);
            }
            targetGraph = occEvent;
            console.log(`\nUpdating single occurrence of: ${occEvent.subject ?? '(no subject)'}`);
            console.log(
              `  ${formatDate(graphStartDt(occEvent))} ${formatTime(graphStartDt(occEvent))} - ${formatTime(occEvent.end?.dateTime ?? '')}`
            );
          } else if (options.occurrence) {
            const idx = parseInt(options.occurrence, 10);
            if (Number.isNaN(idx) || idx < 1 || idx > events.length) {
              console.error(`Invalid --occurrence index: ${options.occurrence}. Valid range: 1-${events.length}.`);
              process.exit(1);
            }
            const occEvent = (events as GraphCalendarEvent[])[idx - 1];
            if (!graphEventMatchesOccurrenceFilter(occEvent, options.id!)) {
              console.error(`Occurrence ${idx} does not match the provided event ID ${options.id}.`);
              process.exit(1);
            }
            targetGraph = occEvent;
            console.log(`\nUpdating occurrence ${idx} of: ${occEvent.subject ?? '(no subject)'}`);
            console.log(
              `  ${formatDate(graphStartDt(occEvent))} ${formatTime(graphStartDt(occEvent))} - ${formatTime(occEvent.end?.dateTime ?? '')}`
            );
          }
        }
      } else {
        targetEws = (events as CalendarEvent[]).find((e) => e.Id === options.id);
        displayEws = targetEws;

        if (options.occurrence || options.instance) {
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
            const occEvent = (events as CalendarEvent[]).find((e) => {
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
            displayEws = occEvent;
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
            const occEvent = (events as CalendarEvent[])[idx - 1];
            if (occEvent.Id !== options.id) {
              console.error(`Occurrence ${idx} does not match the provided event ID ${options.id}.`);
              process.exit(1);
            }
            occurrenceItemId = occEvent.Id;
            displayEws = occEvent;
            console.log(`\nUpdating occurrence ${idx} of: ${occEvent.Subject}`);
            console.log(
              `  ${formatDate(occEvent.Start.DateTime)} ${formatTime(occEvent.Start.DateTime)} - ${formatTime(occEvent.End.DateTime)}`
            );
          }
        } else if (!targetEws && options.id) {
          const fetched = await getCalendarEvent(authResult!.token!, options.id!, options.mailbox);
          if (!fetched.ok || !fetched.data) {
            console.error(`Invalid event id: ${options.id}`);
            process.exit(1);
          }
          displayEws = fetched.data;
          targetEws = fetched.data;
        } else if (!targetEws) {
          console.error(`Invalid event id: ${options.id}`);
          process.exit(1);
        }
      }

      const hasFieldUpdates =
        options.title ||
        options.description ||
        options.start ||
        options.end ||
        options.addAttendee.length > 0 ||
        options.removeAttendee.length > 0 ||
        options.room ||
        options.location ||
        options.timezone ||
        options.teams !== undefined ||
        options.allDay !== undefined ||
        (options.category && options.category.length > 0) ||
        options.clearCategories ||
        !!options.sensitivity;

      const wantsFileAttach = !!options.attach?.trim();
      const linkSpecs = options.attachLink ?? [];
      const wantsLinkAttach = linkSpecs.length > 0;
      const wantsAttachments = wantsFileAttach || wantsLinkAttach;

      if (!hasFieldUpdates && !wantsAttachments) {
        if (useGraph && targetGraph) {
          const tg = targetGraph;
          const st = graphStartDt(tg);
          const en = tg.end?.dateTime ?? '';
          console.log(`\nEvent: ${tg.subject ?? '(no subject)'}`);
          console.log(`  When: ${formatDate(st)} ${formatTime(st)} - ${formatTime(en)}`);
          if (tg.location?.displayName) {
            console.log(`  Location: ${tg.location.displayName}`);
          }
          if (tg.attendees && tg.attendees.length > 0) {
            console.log('  Attendees:');
            for (const a of tg.attendees) {
              const typeLabel = (a as { type?: string }).type === 'resource' ? ' (Room)' : '';
              console.log(`    - ${a.emailAddress?.address ?? ''}${typeLabel}`);
            }
          }
        } else if (displayEws) {
          const de = displayEws;
          console.log(`\nEvent: ${de.Subject}`);
          console.log(
            `  When: ${formatDate(de.Start.DateTime)} ${formatTime(de.Start.DateTime)} - ${formatTime(de.End.DateTime)}`
          );
          if (de.Location?.DisplayName) {
            console.log(`  Location: ${de.Location.DisplayName}`);
          }
          if (de.Attendees && de.Attendees.length > 0) {
            console.log('  Attendees:');
            for (const a of de.Attendees) {
              const typeLabel = a.Type === 'Resource' ? ' (Room)' : '';
              console.log(`    - ${a.EmailAddress?.Address}${typeLabel}`);
            }
          }
        }
        console.log('\nUse options like --title, --add-attendee, --room, --attach, or --attach-link to update.');
        return;
      }

      let fileAttachments: EmailAttachment[] | undefined;
      if (wantsFileAttach) {
        fileAttachments = [];
        const workingDirectory = process.cwd();
        const filePaths = options
          .attach!.split(',')
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
              console.log(`  Adding file attachment: ${validated.fileName}`);
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
      if (wantsLinkAttach) {
        referenceAttachments = [];
        for (const spec of linkSpecs) {
          try {
            const { name, url } = parseAttachLinkSpec(spec);
            referenceAttachments.push({ name, url, contentType: 'text/html' });
            if (!options.json) {
              console.log(`  Adding link attachment: ${name}`);
            }
          } catch (err) {
            const msg =
              err instanceof AttachmentLinkSpecError ? err.message : err instanceof Error ? err.message : String(err);
            console.error(`Invalid --attach-link: ${msg}`);
            process.exit(1);
          }
        }
      }

      let updateResult: Awaited<ReturnType<typeof updateEvent>> | undefined;

      if (
        wantsAttachments &&
        !hasFieldUpdates &&
        (backend === 'graph' || backend === 'auto') &&
        options.id &&
        useGraph
      ) {
        const ga = await resolveGraphAuth({
          token: options.token,
          identity: options.identity
        });
        if (ga.success && ga.token) {
          let gd: GraphCalendarEvent | undefined = targetGraph;
          if (!gd) {
            const ge = await getEvent(ga.token, options.id!, options.mailbox);
            if (ge.ok && ge.data) gd = ge.data;
          }
          if (!gd) {
            if (backend === 'graph') {
              if (options.json) {
                console.log(JSON.stringify({ error: 'Could not load event for attachments' }, null, 2));
              } else {
                console.error(`Could not load event: ${options.id}`);
              }
              process.exit(1);
            }
            if (!options.json) {
              console.warn('[update-event] Could not load event on Graph; falling back to EWS for attachments.');
            }
          } else {
            if (gd.isOrganizer === false) {
              if (options.json) {
                console.log(
                  JSON.stringify(
                    {
                      error: 'Only the organizer can update this event. Use `respond` if you were invited.'
                    },
                    null,
                    2
                  )
                );
              } else {
                console.error('Error: Only the organizer can update this event.');
              }
              process.exit(1);
            }
            const files = fileAttachments ?? [];
            const links = (referenceAttachments ?? []).map((a) => ({ name: a.name, sourceUrl: a.url }));
            const att = await addCalendarEventAttachmentsGraph(
              ga.token,
              gd.id,
              options.mailbox?.trim() || undefined,
              files,
              links
            );
            if (att.ok) {
              if (options.json) {
                console.log(
                  JSON.stringify(
                    {
                      success: true,
                      backend: 'graph',
                      eventId: gd.id,
                      fileAttachmentsAdded: files.length,
                      referenceAttachmentsAdded: links.length
                    },
                    null,
                    2
                  )
                );
              } else {
                console.log(`\n\u2713 Attachment(s) added to calendar event (Graph).`);
              }
              return;
            }
            if (backend === 'graph') {
              if (options.json) {
                console.log(JSON.stringify({ error: att.error?.message || 'Failed to add attachments' }, null, 2));
              } else {
                console.error(`Error: ${att.error?.message || 'Failed to add attachments'}`);
              }
              process.exit(1);
            }
            if (!options.json) {
              console.warn(`[update-event] Graph attachments failed (${att.error?.message}); falling back to EWS.`);
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

      if (hasFieldUpdates) {
        const tryGraphUpdate = (backend === 'graph' || backend === 'auto') && options.id;

        if (tryGraphUpdate) {
          const ga = await resolveGraphAuth({
            token: options.token,
            identity: options.identity
          });
          if (ga.success && ga.token) {
            let gd: GraphCalendarEvent | undefined = useGraph && targetGraph ? targetGraph : undefined;
            let graphGetErr: string | undefined;
            if (!gd) {
              const ge = await getEvent(ga.token, options.id!, options.mailbox);
              if (ge.ok && ge.data) {
                gd = ge.data;
              } else {
                graphGetErr = ge.error?.message;
              }
            }
            if (gd) {
              if (gd.isOrganizer === false) {
                if (options.json) {
                  console.log(
                    JSON.stringify(
                      {
                        error: 'Only the organizer can update this event. Use `respond` if you were invited.'
                      },
                      null,
                      2
                    )
                  );
                } else {
                  console.error('Error: Only the organizer can update this event.');
                }
                process.exit(1);
              }

              let sensitivityEws: 'Normal' | 'Personal' | 'Private' | 'Confidential' | undefined;
              if (options.sensitivity) {
                const s = SENSITIVITY_MAP[options.sensitivity.toLowerCase()];
                if (!s) {
                  console.error(`Invalid sensitivity: ${options.sensitivity}`);
                  process.exit(1);
                }
                sensitivityEws = s;
              }

              const eventDate = new Date(gd.start?.dateTime ?? '');
              let newStart: Date | undefined;
              let newEnd: Date | undefined;
              try {
                if (options.start) {
                  newStart = parseTimeToDate(options.start, eventDate, { throwOnInvalid: true });
                }
                if (options.end) {
                  newEnd = parseTimeToDate(options.end, eventDate, { throwOnInvalid: true });
                }
              } catch (err) {
                const message = err instanceof Error ? err.message : 'Invalid time';
                if (options.json) {
                  console.log(JSON.stringify({ error: message }, null, 2));
                } else {
                  console.error(`Error: ${message}`);
                }
                process.exit(1);
              }

              let locationText = options.location;
              let roomResource: { email: string; name?: string } | undefined;
              if (options.room) {
                if (options.room.includes('@')) {
                  roomResource = { email: options.room, name: options.room };
                  locationText = options.room;
                } else {
                  const rr = await resolveRoomDisplayNameToPlace(ga.token, options.room);
                  if (!rr.ok) {
                    if (options.json) {
                      console.log(JSON.stringify({ error: rr.error }, null, 2));
                    } else {
                      console.error(`Error: ${rr.error}`);
                    }
                    process.exit(1);
                  }
                  const em = rr.place.emailAddress!.trim();
                  roomResource = { email: em, name: rr.place.displayName };
                  locationText = rr.place.displayName?.trim() || em;
                }
              }

              const patch = buildGraphUpdatePatch({
                display: gd,
                title: options.title,
                description: options.description,
                newStart: newStart && newEnd ? newStart : undefined,
                newEnd: newStart && newEnd ? newEnd : undefined,
                timezone: options.timezone,
                location: locationText,
                allDay: options.allDay,
                sensitivity: sensitivityEws,
                categories: options.category && options.category.length > 0 ? options.category : undefined,
                clearCategories: options.clearCategories,
                teams: options.teams === true,
                noTeams: options.teams === false,
                addAttendee: options.addAttendee,
                removeAttendee: options.removeAttendee,
                roomResource
              });

              if (newStart && !newEnd) {
                const endD = new Date(gd.end?.dateTime ?? '');
                patch.end = {
                  dateTime: options.timezone
                    ? toLocalUnzonedISOString(endD)
                    : toUTCISOString(endD).replace(/\.\d{3}Z$/, ''),
                  timeZone: options.timezone?.trim() || gd.end?.timeZone || 'UTC'
                };
                patch.start = {
                  dateTime: options.timezone
                    ? toLocalUnzonedISOString(newStart)
                    : toUTCISOString(newStart).replace(/\.\d{3}Z$/, ''),
                  timeZone: options.timezone?.trim() || gd.start?.timeZone || 'UTC'
                };
              } else if (!newStart && newEnd) {
                const startD = new Date(gd.start?.dateTime ?? '');
                patch.start = {
                  dateTime: options.timezone
                    ? toLocalUnzonedISOString(startD)
                    : toUTCISOString(startD).replace(/\.\d{3}Z$/, ''),
                  timeZone: options.timezone?.trim() || gd.start?.timeZone || 'UTC'
                };
                patch.end = {
                  dateTime: options.timezone
                    ? toLocalUnzonedISOString(newEnd)
                    : toUTCISOString(newEnd).replace(/\.\d{3}Z$/, ''),
                  timeZone: options.timezone?.trim() || gd.end?.timeZone || 'UTC'
                };
              }

              if (Object.keys(patch).length === 0) {
                if (options.json) {
                  console.log(JSON.stringify({ error: 'No field updates to apply' }, null, 2));
                } else {
                  console.error('No field updates to apply.');
                }
                process.exit(1);
              } else {
                console.log(`\nUpdating: ${gd.subject ?? '(no subject)'}`);
                const ur = await updateCalendarEvent(ga.token, gd.id, patch, options.mailbox);
                if (ur.ok && ur.data) {
                  const files = fileAttachments ?? [];
                  const links = (referenceAttachments ?? []).map((a) => ({ name: a.name, sourceUrl: a.url }));
                  if (files.length > 0 || links.length > 0) {
                    const att = await addCalendarEventAttachmentsGraph(
                      ga.token,
                      ur.data.id,
                      options.mailbox?.trim() || undefined,
                      files,
                      links
                    );
                    if (!att.ok) {
                      if (backend === 'graph') {
                        if (options.json) {
                          console.log(
                            JSON.stringify({ error: att.error?.message || 'Failed to add attachments' }, null, 2)
                          );
                        } else {
                          console.error(`Error: ${att.error?.message || 'Failed to add attachments'}`);
                        }
                        process.exit(1);
                      }
                      if (useGraph && backend === 'auto') {
                        if (options.json) {
                          console.log(
                            JSON.stringify({ error: att.error?.message || 'Failed to add attachments' }, null, 2)
                          );
                        } else {
                          console.error(`Error: ${att.error?.message || 'Failed to add attachments'}`);
                        }
                        process.exit(1);
                      }
                    }
                  }
                  if (options.json) {
                    console.log(
                      JSON.stringify(
                        {
                          success: true,
                          backend: 'graph',
                          event: {
                            id: ur.data.id,
                            changeKey: ur.data.changeKey,
                            subject: ur.data.subject,
                            start: ur.data.start?.dateTime,
                            end: ur.data.end?.dateTime
                          },
                          fileAttachmentsAdded: files.length,
                          referenceAttachmentsAdded: links.length
                        },
                        null,
                        2
                      )
                    );
                  } else {
                    console.log('\n\u2713 Event updated successfully.');
                    console.log(`\n  Title: ${ur.data.subject ?? ''}`);
                    const st = ur.data.start?.dateTime ?? '';
                    const en = ur.data.end?.dateTime ?? '';
                    if (st && en) {
                      console.log(`  When:  ${formatDate(st)} ${formatTime(st)} - ${formatTime(en)}`);
                    }
                    if (files.length + links.length > 0) {
                      console.log(`  Attachments: ${files.length} file(s), ${links.length} link(s)`);
                    }
                    console.log('');
                  }
                  return;
                }
                if (backend === 'graph') {
                  if (options.json) {
                    console.log(JSON.stringify({ error: ur.error?.message || 'Failed to update event' }, null, 2));
                  } else {
                    console.error(`Error: ${ur.error?.message || 'Failed to update event'}`);
                  }
                  process.exit(1);
                }
                if (useGraph && backend === 'auto') {
                  if (options.json) {
                    console.log(
                      JSON.stringify(
                        {
                          error:
                            ur.error?.message ||
                            'Graph update failed; cannot fall back to EWS when using Graph calendar data.'
                        },
                        null,
                        2
                      )
                    );
                  } else {
                    console.error(
                      `Error: Graph update failed (${ur.error?.message}). Cannot fall back to EWS when using Graph calendar data; set M365_EXCHANGE_BACKEND=ews or use an EWS event id.`
                    );
                  }
                  process.exit(1);
                }
                if (!options.json) {
                  console.warn(`[update-event] Graph failed (${ur.error?.message}); falling back to EWS.`);
                }
              }
            } else {
              if (backend === 'graph') {
                if (options.json) {
                  console.log(
                    JSON.stringify(
                      {
                        error: graphGetErr || 'Invalid event id',
                        id: options.id,
                        hint: GRAPH_EVENT_ID_HINT
                      },
                      null,
                      2
                    )
                  );
                } else {
                  const detail = graphGetErr ? `: ${graphGetErr}` : '';
                  console.error(`Invalid event id: ${options.id}${detail}`);
                  console.error(GRAPH_EVENT_ID_HINT);
                }
                process.exit(1);
              }
              if (useGraph && backend === 'auto') {
                if (options.json) {
                  console.log(
                    JSON.stringify(
                      {
                        error:
                          graphGetErr ||
                          'Failed to load event from Graph; cannot fall back to EWS when using Graph calendar data.'
                      },
                      null,
                      2
                    )
                  );
                } else {
                  console.error(
                    `Error: ${graphGetErr || 'Failed to load event'}. Cannot fall back to EWS when using Graph calendar data; set M365_EXCHANGE_BACKEND=ews or use an EWS event id.`
                  );
                }
                process.exit(1);
              }
              if (!options.json) {
                console.warn(`[update-event] Graph get event failed (${graphGetErr}); falling back to EWS.`);
              }
            }
          } else if (useGraph && backend === 'auto') {
            if (options.json) {
              console.log(JSON.stringify({ error: ga.error || 'Graph authentication failed' }, null, 2));
            } else {
              console.error(`Error: ${ga.error || 'Graph authentication failed'}`);
            }
            process.exit(1);
          } else if (backend === 'graph') {
            if (options.json) {
              console.log(JSON.stringify({ error: ga.error || 'Graph authentication failed' }, null, 2));
            } else {
              console.error(`Error: ${ga.error || 'Graph authentication failed'}`);
            }
            process.exit(1);
          }
        }

        if (!useGraph) {
          const updateOptions: Parameters<typeof updateEvent>[0] = {
            token: authResult!.token!,
            eventId: (targetEws ?? displayEws!).Id,
            changeKey: displayEws!.ChangeKey,
            occurrenceItemId,
            mailbox: options.mailbox,
            categories: options.clearCategories
              ? []
              : options.category && options.category.length > 0
                ? options.category
                : undefined
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

          if (options.start || options.end) {
            const eventDate = new Date(displayEws!.Start.DateTime);

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

          if (options.location) {
            updateOptions.location = options.location;
          }

          if (options.allDay !== undefined) {
            updateOptions.isAllDay = options.allDay;
          }

          if (options.sensitivity) {
            const sensitivity = SENSITIVITY_MAP[options.sensitivity.toLowerCase()];
            if (!sensitivity) {
              console.error(`Invalid sensitivity: ${options.sensitivity}`);
              process.exit(1);
            }
            updateOptions.sensitivity = sensitivity;
          }

          let roomEmail: string | undefined;
          let roomName: string | undefined;

          if (options.room) {
            if (options.room.includes('@')) {
              roomEmail = options.room;
              roomName = options.room;
            } else {
              let roomsResult = await searchRooms(authResult!.token!, options.room!);
              if (!roomsResult.ok || !roomsResult.data || roomsResult.data.length === 0) {
                roomsResult = await getRooms(authResult!.token!);
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

          if (options.addAttendee.length > 0 || options.removeAttendee.length > 0 || roomEmail) {
            const existingAttendees: Array<{
              email: string;
              name?: string;
              type: 'Required' | 'Optional' | 'Resource';
            }> = (displayEws!.Attendees || []).map((a) => ({
              email: a.EmailAddress?.Address || '',
              name: a.EmailAddress?.Name,
              type: a.Type as 'Required' | 'Optional' | 'Resource'
            }));

            for (const email of options.removeAttendee) {
              const idx = existingAttendees.findIndex((a) => a.email.toLowerCase() === email.toLowerCase());
              if (idx !== -1) existingAttendees.splice(idx, 1);
            }

            for (const email of options.addAttendee) {
              if (!existingAttendees.find((a) => a.email.toLowerCase() === email.toLowerCase())) {
                existingAttendees.push({ email, type: 'Required' });
              }
            }

            if (roomEmail) {
              const withoutRooms = existingAttendees.filter((a) => a.type !== 'Resource');
              withoutRooms.push({ email: roomEmail, name: roomName, type: 'Resource' });
              updateOptions.attendees = withoutRooms;
            } else {
              updateOptions.attendees = existingAttendees;
            }
          }

          if (options.teams !== undefined) {
            updateOptions.isOnlineMeeting = options.teams;
          }

          console.log(`\nUpdating: ${displayEws!.Subject}`);

          updateResult = await updateEvent(updateOptions);

          if (!updateResult.ok) {
            if (options.json) {
              console.log(JSON.stringify({ error: updateResult.error?.message || 'Failed to update event' }, null, 2));
            } else {
              console.error(`\nError: ${updateResult.error?.message || 'Failed to update event'}`);
            }
            process.exit(1);
          }
        }
      }

      const eventIdForAttach = occurrenceItemId || updateResult?.data?.Id || displayEws!.Id;

      if (wantsAttachments) {
        const attachResult = await addCalendarEventAttachments(
          authResult!.token!,
          eventIdForAttach,
          options.mailbox,
          fileAttachments ?? [],
          referenceAttachments ?? []
        );
        if (!attachResult.ok) {
          if (options.json) {
            console.log(JSON.stringify({ error: attachResult.error?.message || 'Failed to add attachments' }, null, 2));
          } else {
            console.error(`\nError: ${attachResult.error?.message || 'Failed to add attachments'}`);
          }
          process.exit(1);
        }
      }

      if (options.json) {
        const dr = updateResult?.data;
        const de = displayEws!;
        console.log(
          JSON.stringify(
            {
              success: true,
              event: {
                id: occurrenceItemId || dr?.Id || de.Id,
                changeKey: dr?.ChangeKey,
                subject: dr?.Subject ?? de.Subject,
                start: dr?.Start.DateTime ?? de.Start.DateTime,
                end: dr?.End.DateTime ?? de.End.DateTime,
                fieldUpdatesApplied: hasFieldUpdates,
                fileAttachmentsAdded: fileAttachments?.length ?? 0,
                referenceAttachmentsAdded: referenceAttachments?.length ?? 0
              }
            },
            null,
            2
          )
        );
      } else {
        if (hasFieldUpdates) {
          console.log('\n\u2713 Event updated successfully.');
        }
        if (wantsAttachments) {
          console.log('\n\u2713 Attachment(s) added to calendar event.');
        }
        const dr = updateResult?.data;
        const de = displayEws!;
        if (dr) {
          console.log(`\n  Title: ${dr.Subject}`);
          console.log(
            `  When:  ${formatDate(dr.Start.DateTime)} ${formatTime(dr.Start.DateTime)} - ${formatTime(dr.End.DateTime)}`
          );
        } else if (wantsAttachments && de) {
          console.log(`\n  Title: ${de.Subject}`);
          console.log(
            `  When:  ${formatDate(de.Start.DateTime)} ${formatTime(de.Start.DateTime)} - ${formatTime(de.End.DateTime)}`
          );
        }
        console.log('');
      }
    }
  );
