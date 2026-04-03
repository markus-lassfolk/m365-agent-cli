import { Command } from 'commander';
import { resolveAuth } from '../lib/auth.js';
import { graphFilterPendingInvitations, graphGetMailboxOrMeEmail } from '../lib/calendar-graph-helpers.js';
import {
  getCalendarEvent,
  getCalendarEvents,
  getOwaUserInfo,
  type ResponseType,
  respondToEvent
} from '../lib/ews-client.js';
import { getExchangeBackend } from '../lib/exchange-backend.js';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import { type GraphCalendarEvent, getEvent, listCalendarView } from '../lib/graph-calendar-client.js';
import type { GraphResponse } from '../lib/graph-client.js';
import { acceptEventInvitation, declineEventInvitation, tentativelyAcceptEventInvitation } from '../lib/graph-event.js';
import { checkReadOnly } from '../lib/utils.js';

function formatTime(dateStr: string): string {
  const date = new Date(dateStr);
  return date.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit', hour12: false });
}

function formatDate(dateStr: string): string {
  const date = new Date(dateStr);
  return date.toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric' });
}

function getResponseIcon(response: string): string {
  switch (response) {
    case 'Accepted':
      return '\u2713';
    case 'Declined':
      return '\u2717';
    case 'TentativelyAccepted':
      return '?';
    case 'None':
    case 'NotResponded':
      return '\u2022';
    default:
      return ' ';
  }
}

/** Map Graph `responseStatus.response` (or attendee status) to keys expected by {@link getResponseIcon}. */
function graphResponseToIconKey(graphRaw: string | undefined): string {
  const r = (graphRaw ?? 'none').toLowerCase();
  if (r === 'accepted') return 'Accepted';
  if (r === 'declined') return 'Declined';
  if (r === 'tentativelyaccepted') return 'TentativelyAccepted';
  return 'NotResponded';
}

function graphStartStr(e: GraphCalendarEvent): string {
  return e.start?.dateTime ?? '';
}

export const respondCommand = new Command('respond')
  .description('Respond to calendar invitations (accept/decline/tentative)')
  .argument('[action]', 'Action: list, accept, decline, tentative')
  .argument('[eventIndex]', 'Event index from the list (deprecated; use --id)')
  .option('--id <eventId>', 'Respond to a specific event by stable ID')
  .option('--comment <text>', 'Add a comment to your response')
  .option('--no-notify', "Don't send response to organizer")
  .option('--include-optional', 'Include optional invitations (default)', true)
  .option('--only-required', 'Only show required invitations')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Use a specific authentication identity (default: default)')
  .option('--mailbox <email>', 'Respond to event in shared mailbox calendar')
  .action(
    async (
      action: string | undefined,
      _eventIndex: string | undefined,
      options: {
        id?: string;
        comment?: string;
        notify: boolean;
        includeOptional?: boolean;
        onlyRequired?: boolean;
        json?: boolean;
        token?: string;
        identity?: string;
        mailbox?: string;
      },
      cmd: any
    ) => {
      const actionLower = (action || 'list').toLowerCase();
      const backend = getExchangeBackend();
      const tryGraphFirst = backend === 'graph' || backend === 'auto';

      async function runGraphList(): Promise<boolean> {
        const ga = await resolveGraphAuth({
          token: options.token,
          identity: options.identity
        });
        if (!ga.success || !ga.token) {
          if (backend === 'graph') {
            if (options.json) {
              console.log(JSON.stringify({ error: ga.error || 'Graph authentication failed' }, null, 2));
            } else {
              console.error(`Error: ${ga.error || 'Graph authentication failed'}`);
            }
            process.exit(1);
          }
          return false;
        }
        const attendeeEmail = await graphGetMailboxOrMeEmail(ga.token, options.mailbox);
        if (!attendeeEmail) {
          if (backend === 'graph') {
            if (options.json) {
              console.log(JSON.stringify({ error: 'Failed to determine user email' }, null, 2));
            } else {
              console.error('Error: Failed to determine user email');
            }
            process.exit(1);
          }
          return false;
        }

        const now = new Date();
        const futureDate = new Date(now);
        futureDate.setDate(futureDate.getDate() + 31);

        const lv = await listCalendarView(ga.token, now.toISOString(), futureDate.toISOString(), {
          user: options.mailbox
        });
        if (!lv.ok || !lv.data) {
          if (backend === 'graph') {
            if (options.json) {
              console.log(JSON.stringify({ error: lv.error?.message || 'Failed to fetch events' }, null, 2));
            } else {
              console.error(`Error: ${lv.error?.message || 'Failed to fetch events'}`);
            }
            process.exit(1);
          }
          return false;
        }

        let pendingEvents = graphFilterPendingInvitations(lv.data, attendeeEmail);

        if (options.onlyRequired) {
          pendingEvents = pendingEvents.filter((e) => {
            const selfAtt = e.attendees?.find((a) => a.emailAddress?.address?.toLowerCase() === attendeeEmail);
            const t = (selfAtt as { type?: string } | undefined)?.type?.toLowerCase();
            return t !== 'optional';
          });
        }

        if (options.json) {
          console.log(
            JSON.stringify(
              {
                backend: 'graph',
                pendingEvents: pendingEvents.map((e, i) => ({
                  index: i + 1,
                  id: e.id,
                  subject: e.subject,
                  start: graphStartStr(e),
                  end: e.end?.dateTime,
                  organizer: e.organizer?.emailAddress?.name || e.organizer?.emailAddress?.address,
                  location: e.location?.displayName
                }))
              },
              null,
              2
            )
          );
          return true;
        }

        console.log('\nCalendar invitations awaiting your response:\n');
        console.log('\u2500'.repeat(60));

        if (pendingEvents.length === 0) {
          console.log('\n  No pending invitations found.\n');
          return true;
        }

        for (let i = 0; i < pendingEvents.length; i++) {
          const event = pendingEvents[i];
          const dateStr = formatDate(graphStartStr(event));
          const startTime = formatTime(graphStartStr(event));
          const endTime = formatTime(event.end?.dateTime ?? '');
          const selfAtt = event.attendees?.find((a) => a.emailAddress?.address?.toLowerCase() === attendeeEmail);
          const respRaw = selfAtt?.status?.response ?? event.responseStatus?.response;
          const icon = getResponseIcon(graphResponseToIconKey(respRaw));

          console.log(`\n  [${i + 1}] ${icon} ${event.subject ?? '(no subject)'}`);
          console.log(`      ${dateStr} ${startTime} - ${endTime}`);
          console.log(`      ID: ${event.id}`);
          if (event.location?.displayName) {
            console.log(`      Location: ${event.location.displayName}`);
          }
          if (event.organizer?.emailAddress) {
            const org = event.organizer.emailAddress;
            console.log(`      Organizer: ${org.name || org.address}`);
          }
        }

        console.log(`\n${'\u2500'.repeat(60)}`);
        console.log('\nTo respond, use:');
        console.log('  m365-agent-cli respond accept --id <eventId>');
        console.log('  m365-agent-cli respond decline --id <eventId>');
        console.log('  m365-agent-cli respond tentative --id <eventId>');
        console.log('');
        return true;
      }

      async function runEwsList(): Promise<void> {
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

        const userInfo = await getOwaUserInfo(authResult.token!);
        const userEmail = userInfo.ok ? userInfo.data?.email?.toLowerCase() : undefined;
        const attendeeEmail = options.mailbox?.toLowerCase() || userEmail;

        if (!attendeeEmail) {
          if (options.json) {
            console.log(JSON.stringify({ error: 'Failed to determine user email' }, null, 2));
          } else {
            console.error('Error: Failed to determine user email');
          }
          process.exit(1);
        }

        const now = new Date();
        const futureDate = new Date(now);
        futureDate.setDate(futureDate.getDate() + 31);

        const result = await getCalendarEvents(
          authResult.token!,
          now.toISOString(),
          futureDate.toISOString(),
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

        const pendingEvents = result.data.filter((event) => {
          if (event.IsCancelled) return false;
          if (event.IsOrganizer) return false;

          const myAttendance = event.Attendees?.find((a) => a.EmailAddress?.Address?.toLowerCase() === attendeeEmail);
          const eventResponse = (event as { ResponseStatus?: { Response?: string } }).ResponseStatus?.Response as
            | string
            | undefined;
          const response = myAttendance?.Status?.Response || eventResponse || 'None';

          const isPending = response === 'None' || response === 'NotResponded';
          if (!isPending) return false;

          const isOptional = myAttendance?.Type === 'Optional';
          if (options.onlyRequired && isOptional) return false;

          return true;
        });

        if (options.json) {
          console.log(
            JSON.stringify(
              {
                backend: 'ews',
                pendingEvents: pendingEvents.map((e, i) => ({
                  index: i + 1,
                  id: e.Id,
                  subject: e.Subject,
                  start: e.Start.DateTime,
                  end: e.End.DateTime,
                  organizer: e.Organizer?.EmailAddress?.Name || e.Organizer?.EmailAddress?.Address,
                  location: e.Location?.DisplayName
                }))
              },
              null,
              2
            )
          );
          return;
        }

        console.log('\nCalendar invitations awaiting your response:\n');
        console.log('\u2500'.repeat(60));

        if (pendingEvents.length === 0) {
          console.log('\n  No pending invitations found.\n');
          return;
        }

        for (let i = 0; i < pendingEvents.length; i++) {
          const event = pendingEvents[i];
          const dateStr = formatDate(event.Start.DateTime);
          const startTime = formatTime(event.Start.DateTime);
          const endTime = formatTime(event.End.DateTime);

          const myAttendance = event.Attendees?.find((a) => a.EmailAddress?.Address?.toLowerCase() === attendeeEmail);
          const eventResponse = (event as { ResponseStatus?: { Response?: string } }).ResponseStatus?.Response as
            | string
            | undefined;
          const response = myAttendance?.Status?.Response || eventResponse || 'None';
          const icon = getResponseIcon(response);

          console.log(`\n  [${i + 1}] ${icon} ${event.Subject}`);
          console.log(`      ${dateStr} ${startTime} - ${endTime}`);
          console.log(`      ID: ${event.Id}`);
          if (event.Location?.DisplayName) {
            console.log(`      Location: ${event.Location.DisplayName}`);
          }
          if (event.Organizer?.EmailAddress) {
            const org = event.Organizer.EmailAddress;
            console.log(`      Organizer: ${org.Name || org.Address}`);
          }
        }

        console.log(`\n${'\u2500'.repeat(60)}`);
        console.log('\nTo respond, use:');
        console.log('  m365-agent-cli respond accept --id <eventId>');
        console.log('  m365-agent-cli respond decline --id <eventId>');
        console.log('  m365-agent-cli respond tentative --id <eventId>');
        console.log('');
      }

      if (actionLower === 'list') {
        if (tryGraphFirst) {
          const ok = await runGraphList();
          if (ok) {
            return;
          }
          if (!options.json) {
            console.warn('[respond] Graph unavailable; falling back to EWS.');
          }
        }
        await runEwsList();
        return;
      }

      if (!['accept', 'decline', 'tentative'].includes(actionLower)) {
        console.error(`Unknown action: ${action}`);
        console.error('Valid actions: list, accept, decline, tentative');
        process.exit(1);
      }

      checkReadOnly(cmd);

      if (!options.id) {
        console.error('Please specify the event id with --id.');
        console.error('Run `m365-agent-cli respond list` to see pending invitations and IDs.');
        process.exit(1);
      }

      if (tryGraphFirst) {
        const ga = await resolveGraphAuth({
          token: options.token,
          identity: options.identity
        });
        if (ga.success && ga.token) {
          const eventResult = await getEvent(ga.token, options.id, options.mailbox);
          if (eventResult.ok && eventResult.data) {
            const ge = eventResult.data;
            if (ge.isOrganizer === true) {
              if (options.json) {
                console.log(
                  JSON.stringify(
                    {
                      error:
                        "You are the organizer of this meeting. Use 'm365-agent-cli update-event' instead to modify the meeting."
                    },
                    null,
                    2
                  )
                );
              } else {
                console.error(
                  "You are the organizer of this meeting. Use 'm365-agent-cli update-event' instead to modify the meeting."
                );
              }
              process.exit(1);
            }

            console.log(`\nResponding to: ${ge.subject ?? '(no subject)'}`);
            console.log(
              `  ${formatDate(graphStartStr(ge))} ${formatTime(graphStartStr(ge))} - ${formatTime(ge.end?.dateTime ?? '')}`
            );
            console.log(`  Action: ${actionLower}`);
            if (options.comment) {
              console.log(`  Comment: ${options.comment}`);
            }
            console.log('');

            const base = {
              token: ga.token,
              eventId: ge.id,
              comment: options.comment,
              sendResponse: options.notify,
              user: options.mailbox
            };

            let gr: GraphResponse<void>;
            if (actionLower === 'accept') {
              gr = await acceptEventInvitation(base);
            } else if (actionLower === 'decline') {
              gr = await declineEventInvitation(base);
            } else {
              gr = await tentativelyAcceptEventInvitation(base);
            }

            if (!gr.ok) {
              if (options.json) {
                console.log(JSON.stringify({ error: gr.error?.message || 'Failed to respond' }, null, 2));
              } else {
                console.error(`Error: ${gr.error?.message || 'Failed to respond'}`);
              }
              process.exit(1);
            }

            const actionPast = actionLower === 'tentative' ? 'tentatively accepted' : `${actionLower}d`;
            if (options.json) {
              console.log(JSON.stringify({ success: true, backend: 'graph', action: actionLower }, null, 2));
            } else {
              console.log(`\u2713 Successfully ${actionPast} the invitation.`);
            }
            return;
          }
          if (backend === 'graph') {
            if (options.json) {
              console.log(JSON.stringify({ error: eventResult.error?.message || 'Invalid event id' }, null, 2));
            } else {
              console.error(`Invalid event id: ${options.id}`);
            }
            process.exit(1);
          }
          if (!options.json) {
            console.warn(`[respond] Graph get event failed (${eventResult.error?.message}); falling back to EWS.`);
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

      const eventResult = await getCalendarEvent(authResult.token!, options.id, options.mailbox);
      if (!eventResult.ok || !eventResult.data) {
        if (options.json) {
          console.log(JSON.stringify({ error: `Invalid event id: ${options.id}` }, null, 2));
        } else {
          console.error(`Invalid event id: ${options.id}`);
        }
        process.exit(1);
      }

      if (eventResult.data.IsOrganizer) {
        if (options.json) {
          console.log(
            JSON.stringify(
              {
                error:
                  "You are the organizer of this meeting. Use 'm365-agent-cli update-event' instead to modify the meeting."
              },
              null,
              2
            )
          );
        } else {
          console.error(
            "You are the organizer of this meeting. Use 'm365-agent-cli update-event' instead to modify the meeting."
          );
        }
        process.exit(1);
      }

      const targetEvent = eventResult.data;

      console.log(`\nResponding to: ${targetEvent.Subject}`);
      console.log(
        `  ${formatDate(targetEvent.Start.DateTime)} ${formatTime(targetEvent.Start.DateTime)} - ${formatTime(targetEvent.End.DateTime)}`
      );
      console.log(`  Action: ${actionLower}`);
      if (options.comment) {
        console.log(`  Comment: ${options.comment}`);
      }
      console.log('');

      const response = await respondToEvent({
        token: authResult.token!,
        eventId: targetEvent.Id,
        response: actionLower as ResponseType,
        comment: options.comment,
        sendResponse: options.notify,
        mailbox: options.mailbox
      });

      if (!response.ok) {
        if (options.json) {
          console.log(JSON.stringify({ error: response.error?.message || 'Failed to respond' }, null, 2));
        } else {
          console.error(`Error: ${response.error?.message || 'Failed to respond'}`);
        }
        process.exit(1);
      }

      const actionPast = actionLower === 'tentative' ? 'tentatively accepted' : `${actionLower}d`;
      if (options.json) {
        console.log(JSON.stringify({ success: true, backend: 'ews', action: actionLower }, null, 2));
      } else {
        console.log(`\u2713 Successfully ${actionPast} the invitation.`);
      }
    }
  );
