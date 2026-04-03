import { access, mkdir, writeFile } from 'node:fs/promises';
import { extname, join } from 'node:path';
import { Command } from 'commander';
import { resolveAuth } from '../lib/auth.js';
import {
  businessDaysBackward,
  businessDaysForward,
  calendarDaysBackward,
  calendarDaysForward,
  isWeekRangeKeyword
} from '../lib/calendar-range.js';
import { parseDay, parseLocalDate } from '../lib/dates.js';
import {
  type CalendarAttendee,
  type CalendarEvent,
  getAttachment,
  getAttachments,
  getCalendarEvent,
  getCalendarEvents
} from '../lib/ews-client.js';
import { getExchangeBackend } from '../lib/exchange-backend.js';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import {
  downloadEventAttachmentBytes,
  type GraphCalendarEvent,
  type GraphEventAttachment,
  getEvent,
  getEventAttachment,
  listCalendarView,
  listEventAttachments
} from '../lib/graph-calendar-client.js';

function sanitizeFileComponent(name: string): string {
  const s = name.replace(/[/\\?%*:|"<>]/g, '_').trim();
  return s.length > 0 ? s : 'attachment';
}

async function pathExists(p: string): Promise<boolean> {
  try {
    await access(p);
    return true;
  } catch {
    return false;
  }
}

function formatTime(dateStr: string): string {
  const date = parseLocalDate(dateStr);
  return date.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit', hour12: false });
}

function formatDate(dateStr: string): string {
  const date = parseLocalDate(dateStr);
  return date.toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric' });
}

/**
 * Convert a local-midnight Date to a UTC ISO string for EWS CalendarView.
 *
 * EWS CalendarView StartDate/EndDate are UTC-datetime strings.
 * CalendarView is exclusive on EndDate, so we set end = next day's local midnight.
 *
 * Example for UTC-5 (EST) on March 15:
 *   start = 2024-03-15T00:00 local = 2024-03-15T05:00:00Z
 *   end   = 2024-03-16T00:00 local = 2024-03-16T05:00:00Z
 */
function toEWSRange(localMidnight: Date): { start: string; end: string } {
  const start = new Date(localMidnight);
  start.setHours(0, 0, 0, 0);

  const end = new Date(start);
  end.setDate(end.getDate() + 1);
  end.setHours(0, 0, 0, 0);

  return { start: start.toISOString(), end: end.toISOString() };
}

function getDateRange(startDay: string, endDay?: string): { start: string; end: string; label: string } {
  const now = new Date();

  // Handle special range keywords
  switch (startDay.toLowerCase()) {
    case 'week':
    case 'thisweek': {
      const start = new Date(now);
      const dayOfWeek = start.getDay();
      const diff = dayOfWeek === 0 ? -6 : 1 - dayOfWeek; // Monday
      start.setDate(start.getDate() + diff);
      start.setHours(0, 0, 0, 0);
      const end = new Date(start);
      end.setDate(end.getDate() + 7); // exclusive end = next Monday midnight
      return { start: start.toISOString(), end: end.toISOString(), label: 'This Week' };
    }
    case 'lastweek': {
      const start = new Date(now);
      const dayOfWeek = start.getDay();
      const diff = dayOfWeek === 0 ? -13 : -6 - dayOfWeek; // Last Monday
      start.setDate(start.getDate() + diff);
      start.setHours(0, 0, 0, 0);
      const end = new Date(start);
      end.setDate(end.getDate() + 7); // exclusive end = next Monday midnight
      return { start: start.toISOString(), end: end.toISOString(), label: 'Last Week' };
    }
    case 'nextweek': {
      const start = new Date(now);
      const dayOfWeek = start.getDay();
      const diff = dayOfWeek === 0 ? 1 : 8 - dayOfWeek; // Next Monday
      start.setDate(start.getDate() + diff);
      start.setHours(0, 0, 0, 0);
      const end = new Date(start);
      end.setDate(end.getDate() + 7); // exclusive end = next Monday midnight
      return { start: start.toISOString(), end: end.toISOString(), label: 'Next Week' };
    }
  }

  // Single day or start of range
  const startDate = parseDay(startDay, { weekdayDirection: 'previous' });
  startDate.setHours(0, 0, 0, 0);

  if (endDay) {
    // Date range - use nearestForward for end date
    const endDate = parseDay(endDay, { baseDate: startDate, weekdayDirection: 'nearestForward' });
    endDate.setHours(0, 0, 0, 0);
    // Exclusive end: next day's midnight
    const endExclusive = new Date(endDate);
    endExclusive.setDate(endExclusive.getDate() + 1);

    const label = `${formatDate(startDate.toISOString())} - ${formatDate(endDate.toISOString())}`;
    return { start: startDate.toISOString(), end: endExclusive.toISOString(), label };
  }

  // Single day — use toEWSRange for consistent UTC conversion
  const { end: endISO } = toEWSRange(startDate);

  return {
    start: startDate.toISOString(),
    end: endISO,
    label: formatDate(startDate.toISOString())
  };
}

function parseCalendarRangeInt(v: string | boolean | undefined, flag: string): number | undefined {
  if (v === undefined || v === '') {
    return undefined;
  }
  if (typeof v === 'boolean') {
    return undefined;
  }
  const n = parseInt(String(v), 10);
  if (!Number.isFinite(n) || n < 1 || n > 366) {
    throw new Error(`${flag} must be an integer between 1 and 366`);
  }
  return n;
}

function parseAnchorDateForDynamicRange(startDay: string): Date {
  const startDate = parseDay(startDay, { weekdayDirection: 'previous' });
  startDate.setHours(0, 0, 0, 0);
  return startDate;
}

/**
 * Resolve EWS window when using --days / --business-days / etc., or delegate to getDateRange.
 */
function resolveCalendarQueryRange(
  startDay: string,
  endDay: string | undefined,
  rangeOpts: {
    days?: number;
    previousDays?: number;
    businessDays?: number;
    previousBusinessDays?: number;
  }
): { start: string; end: string; label: string } {
  const { days, previousDays, businessDays, previousBusinessDays } = rangeOpts;
  const modeCount = [days, previousDays, businessDays, previousBusinessDays].filter((x) => x !== undefined).length;

  if (modeCount === 0) {
    return getDateRange(startDay, endDay);
  }

  if (modeCount > 1) {
    throw new Error(
      'Use only one of: --days, --previous-days, --business-days (--busness-days), --previous-business-days'
    );
  }

  if (endDay !== undefined) {
    throw new Error('Do not pass an end date argument when using --days / --business-days / --previous-days / ...');
  }

  if (isWeekRangeKeyword(startDay)) {
    throw new Error(
      'Week keywords (week, thisweek, lastweek, nextweek) cannot be combined with --days / --business-days / ... — use a single day (e.g. today) as start'
    );
  }

  const anchor = parseAnchorDateForDynamicRange(startDay);
  let result: { start: Date; endExclusive: Date };
  let title: string;

  if (days !== undefined) {
    result = calendarDaysForward(anchor, days);
    title = `Next ${days} calendar day(s)`;
  } else if (previousDays !== undefined) {
    result = calendarDaysBackward(anchor, previousDays);
    title = `Previous ${previousDays} calendar day(s)`;
  } else if (businessDays !== undefined) {
    result = businessDaysForward(anchor, businessDays);
    title = `Next ${businessDays} business day(s)`;
  } else {
    result = businessDaysBackward(anchor, previousBusinessDays!);
    title = `Previous ${previousBusinessDays} business day(s)`;
  }

  const lastInclusive = new Date(result.endExclusive);
  lastInclusive.setDate(lastInclusive.getDate() - 1);

  return {
    start: result.start.toISOString(),
    end: result.endExclusive.toISOString(),
    label: `${title} (${formatDate(result.start.toISOString())} – ${formatDate(lastInclusive.toISOString())})`
  };
}

function getResponseIcon(response: string): string {
  switch (response) {
    case 'Accepted':
      return '✓';
    case 'Declined':
      return '✗';
    case 'TentativelyAccepted':
      return '?';
    case 'NotResponded':
      return '·';
    case 'Organizer':
      return '★';
    default:
      return '·';
  }
}

function graphAttendeeResponseKey(response: string | undefined): string {
  if (!response) return 'NotResponded';
  const u = response.toLowerCase();
  if (u === 'accepted') return 'Accepted';
  if (u === 'declined') return 'Declined';
  if (u === 'tentativelyaccepted') return 'TentativelyAccepted';
  if (u === 'none' || u === 'notresponded') return 'NotResponded';
  return 'NotResponded';
}

function displayGraphCalendarEvent(event: GraphCalendarEvent, verbose: boolean): void {
  const startStr = event.start?.dateTime ?? '';
  const endStr = event.end?.dateTime ?? '';
  const startTime = startStr ? formatTime(startStr) : '?';
  const endTime = endStr ? formatTime(endStr) : '?';
  const location = event.location?.displayName || '';
  const cancelled = event.isCancelled ? ' [CANCELLED]' : '';
  const subject = event.subject || '(no subject)';

  if (event.isAllDay) {
    console.log(`  📅 All day: ${subject}${cancelled}`);
  } else {
    console.log(`  ${startTime} - ${endTime}: ${subject}${cancelled}`);
  }

  if (location) {
    console.log(`     📍 ${location}`);
  }

  if (verbose) {
    const org = event.organizer?.emailAddress?.name || event.organizer?.emailAddress?.address;
    if (org) {
      console.log(`     👤 Organizer: ${org}`);
    }
    if (event.attendees && event.attendees.length > 0) {
      const attendeeList = event.attendees
        .map((a) => {
          const name = a.emailAddress?.name || a.emailAddress?.address || '?';
          const r = graphAttendeeResponseKey(a.status?.response);
          return `${getResponseIcon(r)} ${name}`;
        })
        .join(', ');
      console.log(`     👥 ${attendeeList}`);
    }
    if (event.bodyPreview) {
      const preview = event.bodyPreview.substring(0, 80).replace(/\n/g, ' ');
      console.log(`     📝 ${preview}${event.bodyPreview.length > 80 ? '...' : ''}`);
    }
  }
}

function displayEvent(event: CalendarEvent, verbose: boolean): void {
  const startTime = formatTime(event.Start.DateTime);
  const endTime = formatTime(event.End.DateTime);
  const location = event.Location?.DisplayName || '';
  const cancelled = event.IsCancelled ? ' [CANCELLED]' : '';

  if (event.IsAllDay) {
    console.log(`  📅 All day: ${event.Subject}${cancelled}`);
  } else {
    console.log(`  ${startTime} - ${endTime}: ${event.Subject}${cancelled}`);
  }

  if (location) {
    console.log(`     📍 ${location}`);
  }

  if (event.HasAttachments) {
    console.log(`     📎 Has attachments`);
  }

  if (verbose) {
    // Show organizer if not self
    if (!event.IsOrganizer && event.Organizer?.EmailAddress?.Name) {
      console.log(`     👤 Organizer: ${event.Organizer.EmailAddress.Name}`);
    }

    // Show recurrence info
    if (event.RecurrenceDescription) {
      console.log(`     🔁 ${event.RecurrenceDescription}`);
    } else if (event.FirstOccurrence || event.LastOccurrence) {
      // Series bounds without full description
      const first = event.FirstOccurrence ? ` from ${formatDate(event.FirstOccurrence.Start)}` : '';
      const last = event.LastOccurrence ? ` until ${formatDate(event.LastOccurrence.Start)}` : '';
      if (first || last) {
        console.log(`     🔁 Recurring series${first}${last}`);
      }
    }

    // Show attendees
    if (event.Attendees && event.Attendees.length > 0) {
      const attendeeList = event.Attendees.map(
        (a: CalendarAttendee) => `${getResponseIcon(a.Status.Response)} ${a.EmailAddress.Name}`
      ).join(', ');
      console.log(`     👥 ${attendeeList}`);
    }

    // Show categories
    if (event.Categories && event.Categories.length > 0) {
      console.log(`     🏷️  ${event.Categories.join(', ')}`);
    }

    // Show body preview if available
    if (event.BodyPreview) {
      const preview = event.BodyPreview.substring(0, 80).replace(/\n/g, ' ');
      console.log(`     📝 ${preview}${event.BodyPreview.length > 80 ? '...' : ''}`);
    }

    // Show recurrence series info if available
    if (event.FirstOccurrence || event.LastOccurrence) {
      const first = event.FirstOccurrence ? event.FirstOccurrence.Start.substring(0, 10) : 'unknown';
      const last = event.LastOccurrence ? event.LastOccurrence.Start.substring(0, 10) : 'unknown';
      console.log(`     🔄 Series: First: ${first}, Last: ${last}`);
    }
    if (event.ModifiedOccurrences && event.ModifiedOccurrences.length > 0) {
      console.log(
        `     ✏️  Modified exceptions: ${event.ModifiedOccurrences.map((o) => o.OriginalStart.substring(0, 10)).join(', ')}`
      );
    }
    if (event.DeletedOccurrences && event.DeletedOccurrences.length > 0) {
      console.log(
        `     🗑️  Deleted exceptions: ${event.DeletedOccurrences.map((o) => o.Start.substring(0, 10)).join(', ')}`
      );
    }
  }
}

export const calendarCommand = new Command('calendar')
  .description('View calendar events')
  .argument(
    '[start]',
    'Start day: today, yesterday, tomorrow, monday-sunday, week, lastweek, nextweek, or YYYY-MM-DD',
    'today'
  )
  .argument('[end]', 'End day for range (optional)')
  .option('-v, --verbose', 'Show attendees and more details')
  .option('--list-attachments <eventId>', 'List file and link attachments on an event by id')
  .option('--download-attachments <eventId>', 'Download file attachments; links saved as .url shortcuts')
  .option('-o, --output <dir>', 'Output directory for --download-attachments', '.')
  .option('--force', 'Overwrite existing files when downloading calendar attachments')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Use a specific authentication identity (default: default)')
  .option('--mailbox <email>', 'Delegated or shared mailbox calendar')
  .option(
    '--days <n>',
    'Show N consecutive calendar days forward from start (includes start day; not for use with week keywords)'
  )
  .option('--previous-days <n>', 'Show N consecutive calendar days ending on start day')
  .option(
    '--business-days <n>',
    'Show N weekdays (Mon–Fri) forward from start (skip weekend; use for “next N working days”)'
  )
  .option('--busness-days <n>', 'Same as --business-days (common typo)')
  .option('--previous-business-days <n>', 'Show N weekdays backward ending on the last weekday on or before start')
  .action(
    async (
      startDay: string,
      endDay: string | undefined,
      options: {
        json?: boolean;
        token?: string;
        identity?: string;
        verbose?: boolean;
        mailbox?: string;
        listAttachments?: string;
        downloadAttachments?: string;
        output: string;
        force?: boolean;
        days?: string;
        previousDays?: string;
        businessDays?: string;
        busnessDays?: string;
        previousBusinessDays?: string;
      }
    ) => {
      const backend = getExchangeBackend();
      const mailbox = options.mailbox;

      function graphAttachmentKind(a: GraphEventAttachment): 'file' | 'reference' | 'other' {
        const t = a['@odata.type'] || '';
        if (t.includes('fileAttachment')) return 'file';
        if (t.includes('referenceAttachment')) return 'reference';
        return 'other';
      }

      async function graphListAttachmentsWithToken(graphToken: string): Promise<void> {
        const eventId = options.listAttachments!.trim();
        const eventRes = await getEvent(graphToken, eventId, mailbox);
        if (!eventRes.ok || !eventRes.data) {
          if (options.json) {
            console.log(JSON.stringify({ error: eventRes.error?.message || 'Event not found' }, null, 2));
          } else {
            console.error(`Error: ${eventRes.error?.message || 'Event not found'}`);
          }
          process.exit(1);
        }
        const attsRes = await listEventAttachments(graphToken, eventId, mailbox);
        if (!attsRes.ok || !attsRes.data) {
          if (options.json) {
            console.log(JSON.stringify({ error: attsRes.error?.message || 'Failed to list attachments' }, null, 2));
          } else {
            console.error(`Error: ${attsRes.error?.message || 'Failed to list attachments'}`);
          }
          process.exit(1);
        }
        const atts = attsRes.data.filter((a) => !a.isInline);
        const subject = eventRes.data.subject || '(no subject)';
        if (options.json) {
          console.log(
            JSON.stringify(
              {
                backend: 'graph',
                eventId,
                subject,
                attachments: atts
              },
              null,
              2
            )
          );
          return;
        }
        console.log(`\nAttachments (Graph) — ${subject}`);
        console.log('─'.repeat(40));
        if (atts.length === 0) {
          console.log('  (none)');
        } else {
          for (const a of atts) {
            const k = graphAttachmentKind(a);
            if (k === 'reference' && a.sourceUrl) {
              console.log(`  🔗 ${a.name || a.id}`);
              console.log(`     ${a.sourceUrl}`);
            } else if (k === 'file') {
              const sizeKB = Math.round((a.size || 0) / 1024);
              console.log(`  📄 ${a.name || 'file'} (${sizeKB} KB)`);
            } else {
              console.log(`  📎 ${a.name || a.id} (${a['@odata.type'] || 'attachment'})`);
            }
          }
        }
        console.log();
      }

      if (options.listAttachments) {
        if (backend === 'graph' || backend === 'auto') {
          const ga = await resolveGraphAuth({ token: options.token, identity: options.identity });
          if (ga.success && ga.token) {
            await graphListAttachmentsWithToken(ga.token);
            return;
          }
          if (backend === 'graph') {
            if (options.json) {
              console.log(JSON.stringify({ error: ga.error || 'Graph authentication failed' }, null, 2));
            } else {
              console.error(`Error: ${ga.error || 'Graph authentication failed'}`);
            }
            process.exit(1);
          }
          if (!options.json) {
            console.warn('[calendar] Graph auth failed; falling back to EWS for --list-attachments.');
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

        const token = authResult.token!;

        const eventId = options.listAttachments.trim();
        const eventRes = await getCalendarEvent(token, eventId, mailbox);
        if (!eventRes.ok || !eventRes.data) {
          console.error(`Error: ${eventRes.error?.message || 'Event not found'}`);
          process.exit(1);
        }
        const attsRes = await getAttachments(token, eventId, mailbox);
        if (!attsRes.ok || !attsRes.data) {
          console.error(`Error: ${attsRes.error?.message || 'Failed to list attachments'}`);
          process.exit(1);
        }
        const atts = attsRes.data.value.filter((a) => !a.IsInline);
        if (options.json) {
          console.log(JSON.stringify({ eventId, subject: eventRes.data.Subject, attachments: atts }, null, 2));
          return;
        }
        console.log(`\nAttachments — ${eventRes.data.Subject}`);
        console.log('─'.repeat(40));
        if (atts.length === 0) {
          console.log('  (none)');
        } else {
          for (const a of atts) {
            if (a.Kind === 'reference' && a.AttachLongPathName) {
              console.log(`  🔗 ${a.Name}`);
              console.log(`     ${a.AttachLongPathName}`);
            } else {
              const sizeKB = Math.round(a.Size / 1024);
              console.log(`  📄 ${a.Name} (${sizeKB} KB)`);
            }
          }
        }
        console.log();
        return;
      }

      if (options.downloadAttachments) {
        if (backend === 'graph' || backend === 'auto') {
          const ga = await resolveGraphAuth({ token: options.token, identity: options.identity });
          if (ga.success && ga.token) {
            const eventId = options.downloadAttachments.trim();
            const eventRes = await getEvent(ga.token, eventId, mailbox);
            if (!eventRes.ok || !eventRes.data) {
              console.error(`Error: ${eventRes.error?.message || 'Event not found'}`);
              process.exit(1);
            }
            if (!eventRes.data.hasAttachments) {
              console.log('This event has no attachments.');
              return;
            }
            const attsRes = await listEventAttachments(ga.token, eventId, mailbox);
            if (!attsRes.ok || !attsRes.data) {
              console.error(`Error: ${attsRes.error?.message || 'Failed to fetch attachments'}`);
              process.exit(1);
            }
            const attachments = attsRes.data.filter((a) => !a.isInline);
            if (attachments.length === 0) {
              console.log('No downloadable attachments (inline-only).');
              return;
            }
            await mkdir(options.output, { recursive: true });
            const usedPaths = new Set<string>();
            console.log(`\nDownloading ${attachments.length} attachment(s) to ${options.output}/ (Graph)\n`);
            for (const att of attachments) {
              const kind = graphAttachmentKind(att);
              if (kind === 'reference') {
                let url = att.sourceUrl;
                if (!url && att.id) {
                  const full = await getEventAttachment(ga.token, eventId, att.id, mailbox);
                  if (
                    full.ok &&
                    full.data &&
                    'sourceUrl' in full.data &&
                    (full.data as { sourceUrl?: string }).sourceUrl
                  ) {
                    url = (full.data as { sourceUrl?: string }).sourceUrl;
                  }
                }
                if (!url) {
                  console.error(`  Failed to resolve link: ${att.name || att.id}`);
                  continue;
                }
                const base = sanitizeFileComponent(att.name || 'link');
                let filePath = join(options.output, `${base}.url`);
                let counter = 1;
                while (usedPaths.has(filePath) || (!options.force && (await pathExists(filePath)))) {
                  filePath = join(options.output, `${base} (${counter}).url`);
                  counter++;
                }
                usedPaths.add(filePath);
                const content = `[InternetShortcut]\r\nURL=${url}\r\n`;
                await writeFile(filePath, content, 'utf8');
                console.log(`  ✓ ${filePath.split(/[\\/]/).pop()} (link)`);
                continue;
              }
              if (kind !== 'file') {
                console.warn(`  Skipping non-file attachment: ${att.name || att.id}`);
                continue;
              }
              const dl = await downloadEventAttachmentBytes(ga.token, eventId, att.id, mailbox);
              if (!dl.ok || !dl.data) {
                console.error(`  Failed to download: ${att.name || att.id} (${dl.error?.message})`);
                continue;
              }
              const content = Buffer.from(dl.data);
              const safeName = att.name || `attachment-${att.id}`;
              let filePath = join(options.output, safeName);
              let counter = 1;
              while (true) {
                if (usedPaths.has(filePath)) {
                  const ext = extname(safeName);
                  const base = safeName.slice(0, safeName.length - ext.length);
                  filePath = join(options.output, `${base} (${counter})${ext}`);
                  counter++;
                  continue;
                }
                if (!options.force) {
                  try {
                    await access(filePath);
                    const ext = extname(safeName);
                    const base = safeName.slice(0, safeName.length - ext.length);
                    filePath = join(options.output, `${base} (${counter})${ext}`);
                    counter++;
                    continue;
                  } catch {
                    // missing — ok
                  }
                }
                break;
              }
              usedPaths.add(filePath);
              await writeFile(filePath, content);
              const sizeKB = Math.round(content.length / 1024);
              console.log(`  ✓ ${filePath.split(/[\\/]/).pop()} (${sizeKB} KB)`);
            }
            console.log('\nDone.\n');
            return;
          }
          if (backend === 'graph') {
            if (options.json) {
              console.log(JSON.stringify({ error: ga.error || 'Graph authentication failed' }, null, 2));
            } else {
              console.error(`Error: ${ga.error || 'Graph authentication failed'}`);
            }
            process.exit(1);
          }
          if (!options.json) {
            console.warn('[calendar] Graph auth failed; falling back to EWS for --download-attachments.');
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

        const token = authResult.token!;
        const eventId = options.downloadAttachments.trim();
        const eventRes = await getCalendarEvent(token, eventId, mailbox);
        if (!eventRes.ok || !eventRes.data) {
          console.error(`Error: ${eventRes.error?.message || 'Event not found'}`);
          process.exit(1);
        }
        if (!eventRes.data.HasAttachments) {
          console.log('This event has no attachments.');
          return;
        }
        const attsRes = await getAttachments(token, eventId, mailbox);
        if (!attsRes.ok || !attsRes.data) {
          console.error(`Error: ${attsRes.error?.message || 'Failed to fetch attachments'}`);
          process.exit(1);
        }
        const attachments = attsRes.data.value.filter((a) => !a.IsInline);
        if (attachments.length === 0) {
          console.log('No downloadable attachments (inline-only).');
          return;
        }
        await mkdir(options.output, { recursive: true });
        const usedPaths = new Set<string>();
        console.log(`\nDownloading ${attachments.length} attachment(s) to ${options.output}/\n`);
        for (const att of attachments) {
          if (att.Kind === 'reference' || att.AttachLongPathName) {
            let url = att.AttachLongPathName;
            if (!url) {
              const full = await getAttachment(token, eventId, att.Id, mailbox);
              if (full.ok && full.data?.AttachLongPathName) {
                url = full.data.AttachLongPathName;
              }
            }
            if (!url) {
              console.error(`  Failed to resolve link: ${att.Name}`);
              continue;
            }
            const base = sanitizeFileComponent(att.Name || 'link');
            let filePath = join(options.output, `${base}.url`);
            let counter = 1;
            while (usedPaths.has(filePath) || (!options.force && (await pathExists(filePath)))) {
              filePath = join(options.output, `${base} (${counter}).url`);
              counter++;
            }
            usedPaths.add(filePath);
            const content = `[InternetShortcut]\r\nURL=${url}\r\n`;
            await writeFile(filePath, content, 'utf8');
            console.log(`  ✓ ${filePath.split(/[\\/]/).pop()} (link)`);
            continue;
          }

          const fullAtt = await getAttachment(token, eventId, att.Id, mailbox);
          if (!fullAtt.ok || !fullAtt.data?.ContentBytes) {
            console.error(`  Failed to download: ${att.Name}`);
            continue;
          }
          const content = Buffer.from(fullAtt.data.ContentBytes, 'base64');
          let filePath = join(options.output, att.Name);
          let counter = 1;
          while (true) {
            if (usedPaths.has(filePath)) {
              const ext = extname(att.Name);
              const base = att.Name.slice(0, att.Name.length - ext.length);
              filePath = join(options.output, `${base} (${counter})${ext}`);
              counter++;
              continue;
            }
            if (!options.force) {
              try {
                await access(filePath);
                const ext = extname(att.Name);
                const base = att.Name.slice(0, att.Name.length - ext.length);
                filePath = join(options.output, `${base} (${counter})${ext}`);
                counter++;
                continue;
              } catch {
                // missing — ok
              }
            }
            break;
          }
          usedPaths.add(filePath);
          await writeFile(filePath, content);
          const sizeKB = Math.round(content.length / 1024);
          console.log(`  ✓ ${filePath.split(/[\\/]/).pop()} (${sizeKB} KB)`);
        }
        console.log('\nDone.\n');
        return;
      }

      let start: string;
      let end: string;
      let label: string;
      try {
        const resolved = resolveCalendarQueryRange(startDay, endDay, {
          days: parseCalendarRangeInt(options.days, '--days'),
          previousDays: parseCalendarRangeInt(options.previousDays, '--previous-days'),
          businessDays: parseCalendarRangeInt(options.businessDays ?? options.busnessDays, '--business-days'),
          previousBusinessDays: parseCalendarRangeInt(options.previousBusinessDays, '--previous-business-days')
        });
        start = resolved.start;
        end = resolved.end;
        label = resolved.label;
      } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        if (options.json) {
          console.log(JSON.stringify({ error: message }, null, 2));
        } else {
          console.error(`Error: ${message}`);
        }
        process.exit(1);
      }

      if (backend === 'graph' || backend === 'auto') {
        const ga = await resolveGraphAuth({
          token: options.token,
          identity: options.identity
        });
        if (ga.success && ga.token) {
          const graphResult = await listCalendarView(ga.token, start, end, { user: mailbox });
          if (graphResult.ok && graphResult.data) {
            const graphEvents = graphResult.data.filter((e) => !e.isCancelled);

            if (options.json) {
              console.log(JSON.stringify({ backend: 'graph', label, events: graphEvents }, null, 2));
              return;
            }

            console.log(`\n📆 Calendar for ${label}${mailbox ? ` — ${mailbox}` : ''} (Graph)`);
            console.log('─'.repeat(40));

            if (graphEvents.length === 0) {
              console.log('  No events scheduled.');
            } else {
              const eventsByDate = new Map<string, GraphCalendarEvent[]>();
              for (const event of graphEvents) {
                const startDt = event.start?.dateTime;
                if (!startDt) continue;
                const localDate = parseLocalDate(startDt);
                const dateKey = `${localDate.getFullYear()}-${String(localDate.getMonth() + 1).padStart(2, '0')}-${String(localDate.getDate()).padStart(2, '0')}`;
                if (!eventsByDate.has(dateKey)) {
                  eventsByDate.set(dateKey, []);
                }
                eventsByDate.get(dateKey)?.push(event);
              }

              if (eventsByDate.size > 1) {
                for (const [dateKey, dayEvents] of eventsByDate) {
                  const dayLabel = formatDate(dateKey);
                  console.log(`\n  ${dayLabel}`);
                  for (const event of dayEvents) {
                    displayGraphCalendarEvent(event, options.verbose ?? false);
                  }
                }
              } else {
                for (const event of graphEvents) {
                  displayGraphCalendarEvent(event, options.verbose ?? false);
                }
              }
            }
            console.log();
            return;
          }
          if (backend === 'graph') {
            if (options.json) {
              console.log(
                JSON.stringify({ error: graphResult.error?.message || 'Failed to fetch calendar (Graph)' }, null, 2)
              );
            } else {
              console.error(`Error: ${graphResult.error?.message || 'Failed to fetch calendar (Graph)'}`);
            }
            process.exit(1);
          }
          if (!options.json) {
            console.warn(`[calendar] Graph failed (${graphResult.error?.message || 'unknown'}); falling back to EWS.`);
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

      const result = await getCalendarEvents(authResult.token!, start, end, mailbox);

      if (!result.ok || !result.data) {
        if (options.json) {
          console.log(JSON.stringify({ error: result.error?.message || 'Failed to fetch events' }, null, 2));
        } else {
          console.error(`Error: ${result.error?.message || 'Failed to fetch events'}`);
        }
        process.exit(1);
      }

      const events = result.data.filter((e) => !e.IsCancelled);

      if (options.json) {
        console.log(JSON.stringify(events, null, 2));
        return;
      }

      console.log(`\n📆 Calendar for ${label}${mailbox ? ` — ${mailbox}` : ''}`);
      console.log('─'.repeat(40));

      if (events.length === 0) {
        console.log('  No events scheduled.');
      } else {
        const eventsByDate = new Map<string, CalendarEvent[]>();
        for (const event of events) {
          const localDate = parseLocalDate(event.Start.DateTime);
          const dateKey = `${localDate.getFullYear()}-${String(localDate.getMonth() + 1).padStart(2, '0')}-${String(localDate.getDate()).padStart(2, '0')}`;
          if (!eventsByDate.has(dateKey)) {
            eventsByDate.set(dateKey, []);
          }
          eventsByDate.get(dateKey)?.push(event);
        }

        if (eventsByDate.size > 1) {
          for (const [dateKey, dayEvents] of eventsByDate) {
            const dayLabel = formatDate(dateKey);
            console.log(`\n  ${dayLabel}`);
            for (const event of dayEvents) {
              displayEvent(event, options.verbose ?? false);
            }
          }
        } else {
          for (const event of events) {
            displayEvent(event, options.verbose ?? false);
          }
        }
      }
      console.log();
    }
  );
