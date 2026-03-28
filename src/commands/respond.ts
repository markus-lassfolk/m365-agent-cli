import { Command } from 'commander';
import { resolveAuth } from '../lib/auth.js';
import {
  getCalendarEvents,
  respondToEvent,
  getOwaUserInfo,
  getCalendarEvent,
  type ResponseType
} from '../lib/ews-client.js';

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

      // Get user's email to identify their response status
      const userInfo = await getOwaUserInfo(authResult.token!);
      const userEmail = userInfo.ok ? userInfo.data?.email?.toLowerCase() : undefined;

      // When using a shared mailbox, the attendee email is the mailbox, not the authenticated user
      const attendeeEmail = options.mailbox?.toLowerCase() || userEmail;

      if (!attendeeEmail) {
        if (options.json) {
          console.log(JSON.stringify({ error: 'Failed to determine user email' }, null, 2));
        } else {
          console.error('Error: Failed to determine user email');
        }
        process.exit(1);
      }

      // Fetch upcoming events
      const now = new Date();
      const futureDate = new Date(now);
      futureDate.setDate(futureDate.getDate() + 30); // Look 30 days ahead

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

      // Filter to events where user is an attendee (and not organizer)
      const pendingEvents = result.data.filter((event) => {
        if (event.IsCancelled) return false;
        if (event.IsOrganizer) return false;

        // Find user's attendance record
        const myAttendance = event.Attendees?.find((a) => a.EmailAddress?.Address?.toLowerCase() === attendeeEmail);

        // Some events don't include attendee records; fall back to event-level ResponseStatus if present
        const eventResponse = (event as any).ResponseStatus?.Response as string | undefined;
        const response = myAttendance?.Status?.Response || eventResponse || 'None';

        // Include events where response is None or NotResponded
        const isPending = response === 'None' || response === 'NotResponded';
        if (!isPending) return false;

        // Optional attendance handling (only if we can detect it)
        const isOptional = myAttendance?.Type === 'Optional';
        if (options.onlyRequired && isOptional) return false;

        return true;
      });

      // Default action is 'list'
      const actionLower = (action || 'list').toLowerCase();

      if (actionLower === 'list') {
        if (options.json) {
          console.log(
            JSON.stringify(
              {
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
          const eventResponse = (event as any).ResponseStatus?.Response as string | undefined;
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
        console.log('  clippy respond accept --id <eventId>');
        console.log('  clippy respond decline --id <eventId>');
        console.log('  clippy respond tentative --id <eventId>');
        console.log('');
        return;
      }

      // Handle accept/decline/tentative
      if (!['accept', 'decline', 'tentative'].includes(actionLower)) {
        console.error(`Unknown action: ${action}`);
        console.error('Valid actions: list, accept, decline, tentative');
        process.exit(1);
      }

      if (!options.id) {
        console.error('Please specify the event id with --id.');
        console.error('Run `clippy respond list` to see pending invitations and IDs.');
        process.exit(1);
      }

      // Look up the event directly to check IsOrganizer (pendingEvents filters out organizer events)
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
                error: "You are the organizer of this meeting. Use 'clippy update-event' instead to modify the meeting."
              },
              null,
              2
            )
          );
        } else {
          console.error(
            "You are the organizer of this meeting. Use 'clippy update-event' instead to modify the meeting."
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
        console.log(JSON.stringify({ success: true, action: actionLower }, null, 2));
      } else {
        console.log(`\u2713 Successfully ${actionPast} the invitation.`);
      }
    }
  );
