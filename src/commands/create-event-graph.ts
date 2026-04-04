/**
 * Microsoft Graph path for `create-event` (POST /me/events).
 */

import { toLocalUnzonedISOString } from '../lib/dates.js';
import type { Recurrence, RecurrencePattern } from '../lib/ews-client.js';
import {
  addCalendarEventAttachmentsGraph,
  createCalendarEvent,
  type GraphCalendarEvent,
  type GraphCreateEventRequest,
  type GraphPatternedRecurrence
} from '../lib/graph-calendar-client.js';

function toGraphUtcDateTime(d: Date): string {
  return d.toISOString().replace(/\.\d{3}Z$/, '');
}

function graphTimeZone(opts: { timezoneName?: string }): string {
  return opts.timezoneName?.trim() || 'UTC';
}

function graphStartEnd(opts: { start: Date; end: Date; allDay: boolean; timezoneName?: string }): {
  start: { dateTime: string; timeZone: string };
  end: { dateTime: string; timeZone: string };
} {
  const tz = graphTimeZone(opts);
  if (opts.allDay) {
    const s = new Date(opts.start);
    s.setHours(0, 0, 0, 0);
    const e = new Date(s);
    e.setDate(e.getDate() + 1);
    if (opts.timezoneName?.trim()) {
      return {
        start: { dateTime: toLocalUnzonedISOString(s), timeZone: tz },
        end: { dateTime: toLocalUnzonedISOString(e), timeZone: tz }
      };
    }
    const sUtc = new Date(opts.start);
    sUtc.setUTCHours(0, 0, 0, 0);
    const eUtc = new Date(sUtc);
    eUtc.setUTCDate(eUtc.getUTCDate() + 1);
    return {
      start: { dateTime: toGraphUtcDateTime(sUtc), timeZone: 'UTC' },
      end: { dateTime: toGraphUtcDateTime(eUtc), timeZone: 'UTC' }
    };
  }
  if (opts.timezoneName?.trim()) {
    return {
      start: { dateTime: toLocalUnzonedISOString(opts.start), timeZone: tz },
      end: { dateTime: toLocalUnzonedISOString(opts.end), timeZone: tz }
    };
  }
  return {
    start: { dateTime: toGraphUtcDateTime(opts.start), timeZone: 'UTC' },
    end: { dateTime: toGraphUtcDateTime(opts.end), timeZone: 'UTC' }
  };
}

function mapEwsDayOfWeekIndex(
  idx: RecurrencePattern['Index'] | undefined
): 'first' | 'second' | 'third' | 'fourth' | 'last' {
  const m: Record<string, 'first' | 'second' | 'third' | 'fourth' | 'last'> = {
    First: 'first',
    Second: 'second',
    Third: 'third',
    Fourth: 'fourth',
    Last: 'last'
  };
  return m[idx || 'First'] ?? 'first';
}

function mapEwsRecurrenceToGraph(r: Recurrence): GraphPatternedRecurrence | undefined {
  const p = r.Pattern;
  const rng = r.Range;

  let patternType: GraphPatternedRecurrence['pattern']['type'];
  switch (p.Type) {
    case 'Daily':
      patternType = 'daily';
      break;
    case 'Weekly':
      patternType = 'weekly';
      break;
    case 'AbsoluteMonthly':
      patternType = 'absoluteMonthly';
      break;
    case 'RelativeMonthly':
      patternType = 'relativeMonthly';
      break;
    case 'AbsoluteYearly':
      patternType = 'absoluteYearly';
      break;
    case 'RelativeYearly':
      patternType = 'relativeYearly';
      break;
    default:
      return undefined;
  }

  const pattern: GraphPatternedRecurrence['pattern'] = {
    type: patternType,
    interval: p.Interval || 1
  };

  if (p.Type === 'Weekly' && p.DaysOfWeek && p.DaysOfWeek.length > 0) {
    pattern.daysOfWeek = p.DaysOfWeek.map((d) => d.toLowerCase());
  }
  if (p.Type === 'AbsoluteMonthly' && p.DayOfMonth !== undefined) {
    pattern.dayOfMonth = p.DayOfMonth;
  }
  if (p.Type === 'AbsoluteYearly') {
    if (p.Month !== undefined) pattern.month = p.Month;
    if (p.DayOfMonth !== undefined) pattern.dayOfMonth = p.DayOfMonth;
  }
  if (p.Type === 'RelativeMonthly' || p.Type === 'RelativeYearly') {
    if (p.DaysOfWeek && p.DaysOfWeek.length > 0) {
      pattern.daysOfWeek = p.DaysOfWeek.map((d) => d.toLowerCase());
    }
    pattern.index = mapEwsDayOfWeekIndex(p.Index);
    if (p.Type === 'RelativeYearly' && p.Month !== undefined) {
      pattern.month = p.Month;
    }
  }

  let rangeType: GraphPatternedRecurrence['range']['type'];
  switch (rng.Type) {
    case 'EndDate':
      rangeType = 'endDate';
      break;
    case 'NoEnd':
      rangeType = 'noEnd';
      break;
    case 'Numbered':
      rangeType = 'numbered';
      break;
    default:
      rangeType = 'noEnd';
  }

  const range: GraphPatternedRecurrence['range'] = {
    type: rangeType,
    startDate: rng.StartDate
  };
  if (rng.Type === 'EndDate' && rng.EndDate) {
    range.endDate = rng.EndDate;
  }
  if (rng.Type === 'Numbered' && rng.NumberOfOccurrences !== undefined) {
    range.numberOfOccurrences = rng.NumberOfOccurrences;
  }

  return { pattern, range };
}

function mapSensitivity(s: 'Normal' | 'Personal' | 'Private' | 'Confidential'): GraphCreateEventRequest['sensitivity'] {
  const m: Record<string, GraphCreateEventRequest['sensitivity']> = {
    Normal: 'normal',
    Personal: 'personal',
    Private: 'private',
    Confidential: 'confidential'
  };
  return m[s] ?? 'normal';
}

function buildGraphCreateEventRequest(opts: {
  subject: string;
  body?: string;
  start: Date;
  end: Date;
  allDay: boolean;
  timezoneName?: string;
  attendees: Array<{ email: string; name?: string; type?: 'Required' | 'Optional' | 'Resource' }>;
  teams: boolean;
  locationDisplay?: string;
  sensitivity?: 'Normal' | 'Personal' | 'Private' | 'Confidential';
  categories?: string[];
  recurrence?: Recurrence;
}): GraphCreateEventRequest {
  const { start, end } = graphStartEnd({
    start: opts.start,
    end: opts.end,
    allDay: opts.allDay,
    timezoneName: opts.timezoneName
  });

  const attendees = opts.attendees.map((a) => {
    const t = a.type || 'Required';
    const graphType: 'required' | 'optional' | 'resource' =
      t === 'Optional' ? 'optional' : t === 'Resource' ? 'resource' : 'required';
    return {
      emailAddress: { address: a.email, ...(a.name ? { name: a.name } : {}) },
      type: graphType
    };
  });

  const body: GraphCreateEventRequest = {
    subject: opts.subject,
    start,
    end,
    ...(opts.body?.trim() ? { body: { contentType: 'text' as const, content: opts.body.trim() } } : {}),
    ...(opts.locationDisplay?.trim() ? { location: { displayName: opts.locationDisplay.trim() } } : {}),
    ...(attendees.length > 0 ? { attendees } : {}),
    ...(opts.allDay ? { isAllDay: true } : {}),
    ...(opts.sensitivity ? { sensitivity: mapSensitivity(opts.sensitivity) } : {}),
    ...(opts.categories && opts.categories.length > 0 ? { categories: opts.categories } : {}),
    ...(opts.teams ? { isOnlineMeeting: true, onlineMeetingProvider: 'teamsForBusiness' as const } : {})
  };

  if (opts.recurrence) {
    const gr = mapEwsRecurrenceToGraph(opts.recurrence);
    if (gr) {
      body.recurrence = gr;
    } else {
      console.warn(
        '[create-event] Unsupported or unknown EWS recurrence pattern for Graph; creating a single occurrence. ' +
          `Pattern type: ${opts.recurrence.Pattern.Type}`
      );
    }
  }

  return body;
}

export async function createEventViaGraph(opts: {
  token: string;
  mailbox?: string;
  subject: string;
  body?: string;
  start: Date;
  end: Date;
  allDay: boolean;
  timezoneName?: string;
  attendees: Array<{ email: string; name?: string; type?: 'Required' | 'Optional' | 'Resource' }>;
  teams: boolean;
  locationDisplay?: string;
  sensitivity?: 'Normal' | 'Personal' | 'Private' | 'Confidential';
  categories?: string[];
  recurrence?: Recurrence;
  fileAttachments?: Array<{ name: string; contentType: string; contentBytes: string }>;
  referenceAttachments?: Array<{ name: string; sourceUrl: string }>;
}): Promise<
  | { ok: true; event: GraphCalendarEvent }
  | { ok: false; error: string }
  | { ok: true; event: GraphCalendarEvent; partialSuccess: true; attachmentError: string }
> {
  const payload = buildGraphCreateEventRequest(opts);
  const result = await createCalendarEvent(opts.token, payload, opts.mailbox?.trim() || undefined);
  if (!result.ok || !result.data) {
    return { ok: false, error: result.error?.message || 'Failed to create event' };
  }
  const event = result.data;
  const files = opts.fileAttachments ?? [];
  const links = opts.referenceAttachments ?? [];
  if (files.length > 0 || links.length > 0) {
    const ar = await addCalendarEventAttachmentsGraph(
      opts.token,
      event.id,
      opts.mailbox?.trim() || undefined,
      files,
      links
    );
    if (!ar.ok) {
      return {
        ok: true,
        event,
        partialSuccess: true,
        attachmentError: ar.error?.message || 'Failed to add attachments to event'
      };
    }
  }
  return { ok: true, event };
}
