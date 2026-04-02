import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import {
  type GraphCalendarEvent,
  getCalendar,
  getEvent,
  listCalendars,
  listCalendarView
} from '../lib/graph-calendar-client.js';
import { acceptEventInvitation, declineEventInvitation, tentativelyAcceptEventInvitation } from '../lib/graph-event.js';
import { checkReadOnly } from '../lib/utils.js';

export const graphCalendarCommand = new Command('graph-calendar').description(
  'Microsoft Graph calendar REST: calendars, calendarView, events, invitation responses (distinct from EWS `calendar` / `respond`)'
);

function formatEventLine(e: GraphCalendarEvent): string {
  const subj = e.subject?.trim() || '(no subject)';
  const start = e.start?.dateTime;
  const end = e.end?.dateTime;
  const tz = e.start?.timeZone || '';
  const when =
    start && end ? `${start} → ${end}${tz ? ` (${tz})` : ''}` : start ? `${start}${tz ? ` (${tz})` : ''}` : '?';
  const allDay = e.isAllDay ? ' [all-day]' : '';
  return `${when}${allDay}\t${subj}\t${e.id}`;
}

graphCalendarCommand
  .command('list-calendars')
  .description('List calendars (Graph GET /calendars)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target mailbox (delegation)')
  .action(async (opts: { json?: boolean; token?: string; identity?: string; user?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await listCalendars(auth.token, opts.user);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(r.data, null, 2));
      return;
    }
    for (const c of r.data) {
      const label = c.name || '(unnamed)';
      console.log(`${label}\t${c.id}`);
    }
  });

graphCalendarCommand
  .command('get-calendar')
  .description('Get one calendar by id')
  .argument('<calendarId>', 'Calendar id')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target mailbox (delegation)')
  .action(async (calendarId: string, opts: { json?: boolean; token?: string; identity?: string; user?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await getCalendar(auth.token, calendarId, opts.user);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else console.log(JSON.stringify(r.data, null, 2));
  });

graphCalendarCommand
  .command('list-view')
  .description('List events in a time window (Graph GET .../calendarView)')
  .requiredOption('--start <iso>', 'Start (ISO 8601, e.g. 2026-04-01T00:00:00Z)')
  .requiredOption('--end <iso>', 'End (ISO 8601, exclusive upper bound in many cases — see Graph docs)')
  .option('-c, --calendar <calendarId>', 'Calendar id (omit for default calendar)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target mailbox (delegation)')
  .action(
    async (opts: {
      start: string;
      end: string;
      calendar?: string;
      json?: boolean;
      token?: string;
      identity?: string;
      user?: string;
    }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await listCalendarView(auth.token, opts.start, opts.end, {
        calendarId: opts.calendar,
        user: opts.user
      });
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(r.data, null, 2));
        return;
      }
      for (const e of r.data) {
        console.log(formatEventLine(e));
      }
    }
  );

graphCalendarCommand
  .command('get-event')
  .description('Get a single event by id (Graph GET /events/{id})')
  .argument('<eventId>', 'Event id')
  .option('--select <fields>', 'OData $select (comma-separated)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Mailbox that owns the event (delegation)')
  .action(
    async (
      eventId: string,
      opts: {
        select?: string;
        json?: boolean;
        token?: string;
        identity?: string;
        user?: string;
      }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getEvent(auth.token, eventId, opts.user, opts.select);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else console.log(JSON.stringify(r.data, null, 2));
    }
  );

function addRespondCommand(
  name: string,
  description: string,
  fn: (o: {
    token: string;
    eventId: string;
    comment?: string;
    sendResponse: boolean;
    user?: string;
  }) => ReturnType<typeof acceptEventInvitation>
) {
  graphCalendarCommand
    .command(name)
    .description(description)
    .argument('<eventId>', 'Event id')
    .option('--comment <text>', 'Optional comment to organizer')
    .option('--no-notify', "Don't send response to organizer")
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .option('--user <email>', 'Mailbox that owns the invitation (delegation)')
    .action(
      async (
        eventId: string,
        opts: {
          comment?: string;
          notify: boolean;
          token?: string;
          identity?: string;
          user?: string;
        },
        cmd: any
      ) => {
        checkReadOnly(cmd);
        const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
        if (!auth.success || !auth.token) {
          console.error(`Auth error: ${auth.error}`);
          process.exit(1);
        }
        const r = await fn({
          token: auth.token,
          eventId,
          comment: opts.comment,
          sendResponse: opts.notify !== false,
          user: opts.user
        });
        if (!r.ok) {
          console.error(`Error: ${r.error?.message}`);
          process.exit(1);
        }
        console.log('Done.');
      }
    );
}

addRespondCommand('accept', 'Accept a meeting request (Graph POST .../accept)', acceptEventInvitation);
addRespondCommand('decline', 'Decline a meeting request (Graph POST .../decline)', declineEventInvitation);
addRespondCommand(
  'tentative',
  'Tentatively accept without proposing a new time (Graph POST .../tentativelyAccept)',
  tentativelyAcceptEventInvitation
);
