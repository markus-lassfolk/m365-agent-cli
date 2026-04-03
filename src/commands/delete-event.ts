import { Command } from 'commander';
import { resolveAuth } from '../lib/auth.js';
import {
  graphDayRangeIso,
  graphEventMatchesOccurrenceFilter,
  graphFilterOrganizerEvents,
  graphGetMailboxOrMeEmail,
  graphNonResourceAttendeeCount
} from '../lib/calendar-graph-helpers.js';
import { parseDay } from '../lib/dates.js';
import { type CalendarEvent, cancelEvent, deleteEvent, getCalendarEvents } from '../lib/ews-client.js';
import { getExchangeBackend } from '../lib/exchange-backend.js';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import {
  cancelCalendarEvent,
  deleteCalendarEvent,
  type GraphCalendarEvent,
  listCalendarView
} from '../lib/graph-calendar-client.js';
import { truncateRecurringSeriesBeforeCut } from '../lib/graph-calendar-recurrence.js';
import { checkReadOnly } from '../lib/utils.js';

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

export const deleteEventCommand = new Command('delete-event')
  .description('Delete/cancel a calendar event (sends cancellation if there are attendees)')
  .argument('[eventIndex]', 'Event index from the list (deprecated; use --id)')
  .option('--id <eventId>', 'Delete event by stable ID')
  .option(
    '--day <day>',
    'Day to show events from (today, tomorrow, YYYY-MM-DD) - note: may miss multi-day events crossing midnight',
    'today'
  )
  .option('--search <text>', 'Search for events by title')
  .option('--message <text>', 'Cancellation message to send to attendees')
  .option('--force-delete', 'Delete without sending cancellation (even with attendees)')
  .option('--occurrence <index>', 'Delete only the Nth occurrence of a recurring event')
  .option('--instance <date>', 'Delete only the occurrence on a specific date (YYYY-MM-DD)')
  .option('--scope <scope>', 'Scope: all (default), this (single occurrence), future (this and future)', 'all')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Use a specific authentication identity (default: default)')
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
        identity?: string;
        mailbox?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const backend = getExchangeBackend();
      const baseDate = parseDay(options.day);
      const startOfDay = new Date(baseDate);
      startOfDay.setHours(0, 0, 0, 0);
      const endOfDay = new Date(baseDate);
      endOfDay.setHours(23, 59, 59, 999);
      const graphRange = graphDayRangeIso(baseDate);

      let eventsGraph: GraphCalendarEvent[] | undefined;
      let graphToken: string | undefined;
      let ewsToken: string | undefined;

      const tryGraphFirst = backend === 'graph' || backend === 'auto';

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
            }
          } else if (backend === 'graph') {
            if (options.json) {
              console.log(JSON.stringify({ error: lv.error?.message || 'Failed to list calendar' }, null, 2));
            } else {
              console.error(`Error: ${lv.error?.message || 'Failed to list calendar'}`);
            }
            process.exit(1);
          } else if (!options.json) {
            console.warn(`[delete-event] Graph list failed (${lv.error?.message}); falling back to EWS.`);
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

        ewsToken = authResult.token!;

        const result = await getCalendarEvents(
          ewsToken,
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
                    end: e.End.DateTime
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
          console.log('\n  No events found that you can delete.');
          console.log('  (You can only delete events you organized)\n');
          return;
        }

        for (let i = 0; i < events.length; i++) {
          const event = events[i];
          if (useGraph) {
            const ge = event as GraphCalendarEvent;
            const st = graphStartDt(ge);
            const en = ge.end?.dateTime ?? '';
            const startTime = formatTime(st);
            const endTime = formatTime(en);
            const ac = graphNonResourceAttendeeCount(ge);
            console.log(`\n  [${i + 1}] ${ge.subject ?? '(no subject)'}`);
            console.log(`      ${startTime} - ${endTime}`);
            console.log(`      ID: ${ge.id}`);
            if (ge.location?.displayName) {
              console.log(`      Location: ${ge.location.displayName}`);
            }
            if (ac > 0) {
              console.log(`      Attendees: ${ac} (will be notified on cancel)`);
            }
          } else {
            const e = event as CalendarEvent;
            const attendees = e.Attendees?.filter((a) => a.EmailAddress?.Address && a.Type !== 'Resource') || [];
            console.log(`\n  [${i + 1}] ${e.Subject}`);
            console.log(`      ${formatTime(e.Start.DateTime)} - ${formatTime(e.End.DateTime)}`);
            console.log(`      ID: ${e.Id}`);
            if (e.Location?.DisplayName) {
              console.log(`      Location: ${e.Location.DisplayName}`);
            }
            if (attendees.length > 0) {
              console.log(`      Attendees: ${attendees.length} (will be notified on cancel)`);
            }
          }
        }

        console.log(`\n${'\u2500'.repeat(60)}`);
        console.log('\nTo delete/cancel an event:');
        console.log('  m365-agent-cli delete-event <number>                    # Cancel & notify attendees');
        console.log('  m365-agent-cli delete-event <number> --message "Sorry"  # With cancellation message');
        console.log('  m365-agent-cli delete-event <number> --force-delete     # Delete without notifying');
        console.log('');
        return;
      }

      // Delete the specified event by ID
      let scope = options.scope as 'all' | 'this' | 'future';
      let occurrenceItemId: string | undefined;
      let targetGraph: GraphCalendarEvent | undefined;
      let targetEws: CalendarEvent | undefined;

      if ((options.occurrence || options.instance) && options.scope === 'all') {
        scope = 'this';
      }

      if (useGraph) {
        targetGraph = events.find((e) => (e as GraphCalendarEvent).id === options.id) as GraphCalendarEvent | undefined;
        if (!targetGraph && options.id) {
          targetGraph = events.find((e) => graphEventMatchesOccurrenceFilter(e as GraphCalendarEvent, options.id!)) as
            | GraphCalendarEvent
            | undefined;
        }
        if ((options.occurrence || options.instance) && (scope === 'this' || scope === 'future') && targetGraph) {
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
              const ge = e as GraphCalendarEvent;
              const eventDate = new Date(graphStartDt(ge));
              eventDate.setHours(0, 0, 0, 0);
              return (
                eventDate.getTime() === instanceDate.getTime() && graphEventMatchesOccurrenceFilter(ge, options.id!)
              );
            }) as GraphCalendarEvent | undefined;
            if (!occEvent) {
              console.error(
                `No occurrence found on ${options.instance} with ID ${options.id}. Try expanding the date range with --day.`
              );
              process.exit(1);
            }
            occurrenceItemId = occEvent.id;
            targetGraph = occEvent;
          } else if (options.occurrence) {
            const idx = parseInt(options.occurrence, 10);
            if (Number.isNaN(idx) || idx < 1) {
              console.error('--occurrence must be a positive integer');
              process.exit(1);
            }
            if (idx > events.length) {
              console.error(
                `Invalid occurrence index: ${idx}. Only ${events.length} occurrence(s) found in the date range.`
              );
              process.exit(1);
            }
            const occEvent = events[idx - 1] as GraphCalendarEvent;
            if (!graphEventMatchesOccurrenceFilter(occEvent, options.id!)) {
              console.error(`Occurrence ${idx} does not match the provided event ID ${options.id}.`);
              process.exit(1);
            }
            occurrenceItemId = occEvent.id;
            targetGraph = occEvent;
          }
          if (!options.json) {
            if (scope === 'future') {
              console.log(`\nTruncating series from occurrence: ${targetGraph.subject ?? '(no subject)'}`);
            } else {
              console.log(`\nDeleting single occurrence: ${targetGraph.subject ?? '(no subject)'}`);
            }
            console.log(
              `  ${formatDate(graphStartDt(targetGraph))} ${formatTime(graphStartDt(targetGraph))} - ${formatTime(targetGraph.end?.dateTime ?? '')}`
            );
          }
        } else if (!targetGraph) {
          console.error(`Invalid event id: ${options.id}`);
          process.exit(1);
        } else if (scope === 'all') {
          if (!options.json) {
            console.log(`\nDeleting: ${targetGraph.subject ?? '(no subject)'}`);
            console.log(
              `  ${formatDate(graphStartDt(targetGraph))} ${formatTime(graphStartDt(targetGraph))} - ${formatTime(targetGraph.end?.dateTime ?? '')}`
            );
          }
        } else {
          if (!options.json) {
            console.log(`\nDeleting: ${targetGraph.subject ?? '(no subject)'} (scope: ${scope})`);
            console.log(
              `  ${formatDate(graphStartDt(targetGraph))} ${formatTime(graphStartDt(targetGraph))} - ${formatTime(targetGraph.end?.dateTime ?? '')}`
            );
          }
        }
      } else {
        targetEws = (events as CalendarEvent[]).find((e) => e.Id === options.id);
        if ((options.occurrence || options.instance) && (scope === 'this' || scope === 'future')) {
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
            targetEws = occEvent;
          } else if (options.occurrence) {
            const idx = parseInt(options.occurrence, 10);
            if (Number.isNaN(idx) || idx < 1) {
              console.error('--occurrence must be a positive integer');
              process.exit(1);
            }
            if (idx > events.length) {
              console.error(
                `Invalid occurrence index: ${idx}. Only ${events.length} occurrence(s) found in the date range.`
              );
              process.exit(1);
            }
            const occEvent = (events as CalendarEvent[])[idx - 1];
            if (occEvent.Id !== options.id) {
              console.error(`Occurrence ${idx} does not match the provided event ID ${options.id}.`);
              process.exit(1);
            }
            occurrenceItemId = occEvent.Id;
            targetEws = occEvent;
          }
          if (!options.json) {
            if (scope === 'future') {
              console.log(`\nTruncating series from occurrence: ${targetEws!.Subject}`);
            } else {
              console.log(`\nDeleting single occurrence: ${targetEws!.Subject}`);
            }
            console.log(
              `  ${formatDate(targetEws!.Start.DateTime)} ${formatTime(targetEws!.Start.DateTime)} - ${formatTime(targetEws!.End.DateTime)}`
            );
          }
        } else if (!targetEws) {
          console.error(`Invalid event id: ${options.id}`);
          process.exit(1);
        } else if (scope !== 'all') {
          if (!options.json) {
            console.log(`\nDeleting: ${targetEws.Subject} (scope: ${scope})`);
            console.log(
              `  ${formatDate(targetEws.Start.DateTime)} ${formatTime(targetEws.Start.DateTime)} - ${formatTime(targetEws.End.DateTime)}`
            );
          }
        } else {
          if (!options.json) {
            console.log(`\nDeleting: ${targetEws.Subject}`);
            console.log(
              `  ${formatDate(targetEws.Start.DateTime)} ${formatTime(targetEws.Start.DateTime)} - ${formatTime(targetEws.End.DateTime)}`
            );
          }
        }
      }

      const deletionId = useGraph ? occurrenceItemId || targetGraph!.id : occurrenceItemId || targetEws!.Id;

      if (useGraph && graphToken && scope === 'future') {
        const tr = await truncateRecurringSeriesBeforeCut(graphToken, options.mailbox, targetGraph!, {
          forceDelete: options.forceDelete
        });
        if (!tr.ok) {
          if (options.json) {
            console.log(JSON.stringify({ error: tr.error?.message || 'Failed to truncate series' }, null, 2));
          } else {
            console.error(`\nError: ${tr.error?.message || 'Failed to truncate series'}`);
          }
          process.exit(1);
        }
        const seriesAction = tr.data!.action;
        const attendeesNotified = tr.data!.attendeesNotified ?? 0;
        if (options.json) {
          console.log(
            JSON.stringify(
              {
                success: true,
                backend: 'graph',
                action: seriesAction,
                event: targetGraph!.subject,
                attendeesNotified
              },
              null,
              2
            )
          );
        } else {
          if (seriesAction === 'truncated') {
            console.log('\n\u2713 Recurring series updated: this and future occurrences were removed.\n');
          } else if (seriesAction === 'cancelled') {
            console.log('\n\u2713 Event cancelled. Attendees were notified.\n');
          } else {
            console.log('\n\u2713 Event deleted.\n');
          }
        }
        return;
      }

      if (useGraph && graphToken) {
        const hasAttendees = graphNonResourceAttendeeCount(targetGraph!) > 0;
        const org = targetGraph!.organizer?.emailAddress?.address?.toLowerCase();
        const attendees =
          targetGraph!.attendees?.filter((a) => {
            const addr = a.emailAddress?.address;
            if (!addr) return false;
            if ((a as { type?: string }).type === 'resource') return false;
            if (addr.toLowerCase() === org) return false;
            return true;
          }) ?? [];

        let graphRes: { ok: boolean; error?: { message?: string } };
        let action: string;

        if (hasAttendees && !options.forceDelete && scope === 'all') {
          console.log(
            `  Attendees: ${attendees
              .map((a) => a.emailAddress?.address)
              .filter(Boolean)
              .join(', ')}`
          );
          console.log(`  Sending cancellation notices...`);
          graphRes = await cancelCalendarEvent(graphToken, deletionId, {
            comment: options.message,
            user: options.mailbox
          });
          action = 'cancelled';
        } else {
          graphRes = await deleteCalendarEvent(graphToken, deletionId, options.mailbox);
          action = 'deleted';
        }

        if (!graphRes.ok) {
          if (options.json) {
            console.log(JSON.stringify({ error: graphRes.error?.message || `Failed to ${action} event` }, null, 2));
          } else {
            console.error(`\nError: ${graphRes.error?.message || `Failed to ${action} event`}`);
          }
          process.exit(1);
        }

        if (options.json) {
          console.log(
            JSON.stringify(
              {
                success: true,
                backend: 'graph',
                action,
                event: targetGraph!.subject,
                attendeesNotified: hasAttendees && !options.forceDelete ? attendees.length : 0
              },
              null,
              2
            )
          );
        } else {
          if (hasAttendees && !options.forceDelete) {
            console.log(`\n\u2713 Event cancelled. ${attendees.length} attendee(s) notified.\n`);
          } else {
            console.log('\n\u2713 Event deleted.\n');
          }
        }
        return;
      }

      const attendeesEws = targetEws!.Attendees?.filter((a) => a.EmailAddress?.Address && a.Type !== 'Resource') || [];
      const hasAttendees = attendeesEws.length > 0;

      let deleteResult: Awaited<ReturnType<typeof deleteEvent>>;
      let action: string;

      if (hasAttendees && !options.forceDelete && scope === 'all') {
        console.log(`  Attendees: ${attendeesEws.map((a) => a.EmailAddress?.Address).join(', ')}`);
        console.log(`  Sending cancellation notices...`);
        deleteResult = await cancelEvent({
          token: ewsToken!,
          eventId: targetEws!.Id,
          comment: options.message,
          mailbox: options.mailbox
        });
        action = 'cancelled';
      } else {
        deleteResult = await deleteEvent({
          token: ewsToken!,
          eventId: targetEws!.Id,
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
              backend: 'ews',
              action,
              event: targetEws!.Subject,
              attendeesNotified: hasAttendees && !options.forceDelete ? attendeesEws.length : 0,
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
          console.log(`\n\u2713 Event cancelled. ${attendeesEws.length} attendee(s) notified.\n`);
        } else {
          console.log('\n\u2713 Event deleted.\n');
        }
      }
    }
  );
