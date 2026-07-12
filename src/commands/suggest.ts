import { readFile } from 'node:fs/promises';
import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import { type AttendeeBase, type FindMeetingTimesRequest, findMeetingTimes } from '../lib/graph-schedule.js';
import { buildFindMeetingTimesLocationConstraint } from '../lib/meeting-location-constraint.js';

export const suggestCommand = new Command('suggest')
  .description('AI meeting time suggestions')
  .option('--attendees <emails>', 'Comma-separated email addresses to invite')
  .option('--duration <duration>', 'Duration (e.g., 30m, 1h)', '30m')
  .option('--days <days>', 'Number of days to check from now', '5')
  .option('--suggest-locations', 'Ask findMeetingTimes to suggest locations (locationConstraint.suggestLocation)')
  .option('--require-meeting-location', 'Require a resolved location (locationConstraint.isRequired)')
  .option(
    '--meeting-location <spec>',
    'Location constraint entry: displayName, room@domain, or "Name|room@domain" (repeatable)',
    (v: string, prev: string[]) => [...prev, v],
    [] as string[]
  )
  .option(
    '--resolve-location-availability',
    'Set resolveAvailability on locationConstraint.locations entries (room mailbox resolution)'
  )
  .option(
    '--find-meeting-json <path>',
    'Merge JSON object into the findMeetingTimes request (overrides same top-level keys)'
  )
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token (bypass interactive auth)')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Mailbox whose calendar is used for findMeetingTimes (delegation)')
  .action(
    async (options: {
      attendees?: string;
      duration: string;
      days: string;
      suggestLocations?: boolean;
      requireMeetingLocation?: boolean;
      meetingLocation?: string[];
      resolveLocationAvailability?: boolean;
      findMeetingJson?: string;
      json?: boolean;
      token?: string;
      identity?: string;
      user?: string;
    }) => {
      const authResult = await resolveGraphAuth({ token: options.token, identity: options.identity });
      if (!authResult.success || !authResult.token) {
        if (options.json) {
          console.log(JSON.stringify({ error: authResult.error }, null, 2));
        } else {
          console.error(`Error: ${authResult.error}`);
        }
        process.exit(1);
      }

      const durationMapping: Record<string, string> = {
        '15m': 'PT15M',
        '30m': 'PT30M',
        '45m': 'PT45M',
        '1h': 'PT1H',
        '2h': 'PT2H'
      };

      const durationKey = options.duration.trim().toLowerCase();
      if (!Object.hasOwn(durationMapping, durationKey)) {
        const message = `Invalid duration "${options.duration}". Supported values are: ${Object.keys(durationMapping).join(', ')}.`;
        if (options.json) {
          console.log(JSON.stringify({ error: message }, null, 2));
        } else {
          console.error(`Error: ${message}`);
        }
        process.exit(1);
      }

      const durationStr = durationMapping[durationKey];
      const days = parseInt(options.days, 10) || 5;

      const startDateTime = new Date();
      const endDateTime = new Date(startDateTime);
      endDateTime.setDate(startDateTime.getDate() + days);

      const startDateTimeISO = startDateTime.toISOString().replace('Z', '');
      const endDateTimeISO = endDateTime.toISOString().replace('Z', '');

      const attendeesList: AttendeeBase[] = options.attendees
        ? options.attendees
            .split(',')
            .map((email) => email.trim())
            .filter(Boolean)
            .map((address) => ({
              type: 'required' as const,
              emailAddress: { address }
            }))
        : [];

      let request: FindMeetingTimesRequest = {
        attendees: attendeesList.length > 0 ? attendeesList : undefined,
        meetingDuration: durationStr,
        timeConstraint: {
          activityDomain: 'work',
          timeSlots: [
            {
              start: {
                dateTime: startDateTimeISO,
                timeZone: 'UTC'
              },
              end: {
                dateTime: endDateTimeISO,
                timeZone: 'UTC'
              }
            }
          ]
        },
        isOrganizerOptional: false,
        returnSuggestionReasons: true,
        minimumAttendeePercentage: 100,
        locationConstraint: buildFindMeetingTimesLocationConstraint({
          suggestLocations: options.suggestLocations,
          locationRequired: options.requireMeetingLocation,
          resolveLocationAvailability: options.resolveLocationAvailability,
          meetingLocation: options.meetingLocation ?? []
        })
      };

      if (options.findMeetingJson?.trim()) {
        try {
          const extra = JSON.parse(await readFile(options.findMeetingJson.trim(), 'utf8')) as Record<string, unknown>;
          request = { ...request, ...extra } as FindMeetingTimesRequest;
        } catch (err) {
          const message = err instanceof Error ? err.message : 'Invalid --find-meeting-json file';
          if (options.json) {
            console.log(JSON.stringify({ error: message }, null, 2));
          } else {
            console.error(`Error: ${message}`);
          }
          process.exit(1);
        }
      }

      const result = await findMeetingTimes(authResult.token, request, options.user);

      if (!result.ok || !result.data) {
        if (options.json) {
          console.log(JSON.stringify({ error: result.error }, null, 2));
        } else {
          console.error('Error finding meeting times:', result.error?.message || 'Unknown error');
        }
        process.exit(1);
      }

      if (options.json) {
        console.log(JSON.stringify(result.data, null, 2));
        return;
      }

      const { emptySuggestionsReason, meetingTimeSuggestions } = result.data;

      console.log('\nAI Meeting Time Suggestions:\n');

      if (emptySuggestionsReason) {
        console.log('No suitable meeting times found for the following reason:');
        console.log(`  - ${emptySuggestionsReason}`);
        return;
      }

      if (!meetingTimeSuggestions || meetingTimeSuggestions.length === 0) {
        console.log('No suggestions found.');
        return;
      }

      for (const suggestion of meetingTimeSuggestions) {
        const start = suggestion.meetingTimeSlot?.start?.dateTime
          ? new Date(`${suggestion.meetingTimeSlot.start.dateTime}Z`).toLocaleString()
          : 'Unknown';
        const end = suggestion.meetingTimeSlot?.end?.dateTime
          ? new Date(`${suggestion.meetingTimeSlot.end.dateTime}Z`).toLocaleString()
          : 'Unknown';

        const confidence = suggestion.confidence !== undefined ? `${suggestion.confidence}%` : 'Unknown';

        console.log(`Suggestion: ${start} - ${end}`);
        console.log(`  Confidence: ${confidence}`);

        if (suggestion.suggestionReason) {
          console.log(`  Reason: ${suggestion.suggestionReason}`);
        }

        if (suggestion.attendeeAvailability && suggestion.attendeeAvailability.length > 0) {
          console.log('  Attendee Availability:');
          for (const attendee of suggestion.attendeeAvailability) {
            const email = attendee.attendee?.emailAddress?.address || 'Unknown';
            const status = attendee.availability || 'Unknown';
            console.log(`    - ${email}: ${status}`);
          }
        }
        if (suggestion.locations && suggestion.locations.length > 0) {
          console.log('  Suggested locations:');
          for (const loc of suggestion.locations) {
            const bits = [loc.displayName, loc.locationEmailAddress].filter(Boolean);
            console.log(`    - ${bits.join(' · ') || '(unspecified)'}`);
          }
        }
        console.log();
      }
    }
  );
