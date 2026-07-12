/**
 * Microsoft Graph path for `findtime` via findMeetingTimes (same family as `suggest`),
 * with fallback to calendar/getSchedule + merged availability views.
 */

import { graphEventStartMs } from '../lib/graph-calendar-recurrence.js';
import {
  type AttendeeBase,
  type FindMeetingTimesRequest,
  findMeetingTimes,
  getSchedule
} from '../lib/graph-schedule.js';
import { formatDateInTimeZone, hourInTimeZone } from '../lib/timezone-wallclock.js';

/** Each char in availabilityView: 0=free, 1–5=busy/tentative/OOF/etc. */
function padAvailabilityView(view: string, len: number): string {
  if (view.length >= len) return view.slice(0, len);
  return view + '2'.repeat(len - view.length);
}

/** Merge per-mailbox views: a slot is free only if every mailbox is `0` at that index. */
export function mergeAvailabilityViewsToMergedFree(views: string[]): boolean[] {
  if (views.length === 0) return [];
  const maxLen = Math.max(...views.map((v) => v.length), 0);
  if (maxLen === 0) return [];
  const padded = views.map((v) => padAvailabilityView(v, maxLen));
  const merged: boolean[] = [];
  for (let i = 0; i < maxLen; i++) {
    let allFree = true;
    for (const v of padded) {
      if ((v[i] ?? '2') !== '0') {
        allFree = false;
        break;
      }
    }
    merged.push(allFree);
  }
  return merged;
}

export async function runFindTimeGraph(opts: {
  token: string;
  emails: string[];
  start: Date;
  end: Date;
  durationMinutes: number;
  workStartHour: number;
  workEndHour: number;
  label: string;
  mailbox?: string;
  json?: boolean;
  locationConstraint?: FindMeetingTimesRequest['locationConstraint'];
  findMeetingMerge?: Partial<FindMeetingTimesRequest>;
  optionalEmails?: string[];
  minAttendeePercentage?: number;
  timezone?: string;
}): Promise<{ ok: true } | { ok: false; error: string }> {
  const optionalSet = new Set((opts.optionalEmails ?? []).map((e) => e.toLowerCase()));
  const attendeePayload: AttendeeBase[] = opts.emails.map((address) => ({
    type: optionalSet.has(address.toLowerCase()) ? ('optional' as const) : ('required' as const),
    emailAddress: { address }
  }));

  const tz = opts.timezone || 'UTC';
  const startDateTime = opts.timezone
    ? formatDateInTimeZone(opts.start, opts.timezone)
    : opts.start.toISOString().replace('Z', '');
  const endDateTime = opts.timezone
    ? formatDateInTimeZone(opts.end, opts.timezone)
    : opts.end.toISOString().replace('Z', '');

  const baseRequest: FindMeetingTimesRequest = {
    attendees: attendeePayload,
    meetingDuration: `PT${opts.durationMinutes}M`,
    timeConstraint: {
      activityDomain: 'work',
      timeSlots: [
        {
          start: { dateTime: startDateTime, timeZone: tz },
          end: { dateTime: endDateTime, timeZone: tz }
        }
      ]
    },
    minimumAttendeePercentage: opts.minAttendeePercentage ?? 100,
    isOrganizerOptional: false,
    returnSuggestionReasons: true,
    ...(opts.locationConstraint ? { locationConstraint: opts.locationConstraint } : {})
  };
  const merged: FindMeetingTimesRequest = opts.findMeetingMerge
    ? ({ ...baseRequest, ...opts.findMeetingMerge } as FindMeetingTimesRequest)
    : baseRequest;

  const result = await findMeetingTimes(opts.token, merged, opts.mailbox?.trim() || undefined, tz);

  if (!result.ok || !result.data) {
    return { ok: false, error: result.error?.message || 'Failed to find meeting times' };
  }

  const { emptySuggestionsReason, meetingTimeSuggestions } = result.data;

  if (opts.json) {
    const slots =
      meetingTimeSuggestions?.map((s) => ({
        start: s.meetingTimeSlot?.start?.dateTime,
        end: s.meetingTimeSlot?.end?.dateTime,
        confidence: s.confidence,
        reason: s.suggestionReason,
        locations: s.locations,
        attendeeAvailability: s.attendeeAvailability?.map((a) => ({
          email: a.attendee?.emailAddress?.address,
          availability: a.availability
        }))
      })) ?? [];
    console.log(
      JSON.stringify(
        {
          backend: 'graph',
          attendees: opts.emails,
          duration: opts.durationMinutes,
          dateRange: { start: opts.start.toISOString(), end: opts.end.toISOString() },
          emptySuggestionsReason,
          suggestions: slots
        },
        null,
        2
      )
    );
    return { ok: true };
  }

  console.log(`\n🗓️  Finding ${opts.durationMinutes}-minute meeting times (Graph)`);
  console.log(`   Attendees: ${opts.emails.join(', ')}`);
  console.log(`   Date range: ${opts.label}`);
  console.log('─'.repeat(50));

  if (emptySuggestionsReason) {
    console.log(`\n  No suggestions: ${emptySuggestionsReason}`);
    console.log();
    return { ok: true };
  }

  const suggestions = meetingTimeSuggestions || [];
  const filtered = suggestions.filter((s) => {
    const start = s.meetingTimeSlot?.start;
    if (!start?.dateTime) return false;
    // Graph returns dateTime as wall-clock time already expressed in the zone we sent via the
    // Prefer header (tz — opts.timezone or the UTC default), so when opts.timezone is set the
    // hour digits can be read directly; round-tripping through graphEventStartMs (which only
    // understands UTC/GMT) + hourInTimeZone would re-project an already-zoned time and double-shift it.
    if (opts.timezone) {
      const hour = Number(start.dateTime.slice(11, 13));
      return hour >= opts.workStartHour && hour < opts.workEndHour;
    }
    const ms = graphEventStartMs(start);
    const hour = Number.isFinite(ms) ? new Date(ms).getHours() : new Date(start.dateTime).getHours();
    return hour >= opts.workStartHour && hour < opts.workEndHour;
  });

  if (filtered.length === 0) {
    console.log('\n  No available times in the selected working hours window.');
    console.log('     Try a longer date range, different hours (--start/--end), or shorter duration.');
  } else {
    console.log(`\n  ✅ Found ${filtered.length} suggested slot${filtered.length > 1 ? 's' : ''}:\n`);
    for (const s of filtered) {
      const stSlot = s.meetingTimeSlot?.start;
      const enSlot = s.meetingTimeSlot?.end;
      const stMs = stSlot ? graphEventStartMs(stSlot) : NaN;
      const enMs = enSlot ? graphEventStartMs(enSlot) : NaN;
      const startLabel = stSlot?.dateTime
        ? Number.isFinite(stMs)
          ? new Date(stMs).toLocaleString()
          : new Date(stSlot.dateTime).toLocaleString()
        : '?';
      const endLabel = enSlot?.dateTime
        ? Number.isFinite(enMs)
          ? new Date(enMs).toLocaleString()
          : new Date(enSlot.dateTime).toLocaleString()
        : '?';
      const conf = s.confidence !== undefined ? ` (${s.confidence}% confidence)` : '';
      console.log(`    🟢 ${startLabel} – ${endLabel}${conf}`);
      if (s.suggestionReason) console.log(`       ${s.suggestionReason}`);
      if (s.attendeeAvailability && s.attendeeAvailability.length > 0) {
        for (const a of s.attendeeAvailability) {
          const email = a.attendee?.emailAddress?.address || 'Unknown';
          console.log(`       ${email}: ${a.availability || 'unknown'}`);
        }
      }
      if (s.locations && s.locations.length > 0) {
        const locBits = s.locations
          .map((l) => [l.displayName, l.locationEmailAddress].filter(Boolean).join(' · '))
          .filter(Boolean);
        if (locBits.length > 0) console.log(`       Locations: ${locBits.join('; ')}`);
      }
    }
  }
  console.log();
  return { ok: true };
}

/**
 * Second Graph strategy: `POST /calendar/getSchedule`, merge `availabilityView` strings,
 * then find windows where all attendees are free for `durationMinutes` within work hours.
 */
export async function runFindTimeGraphSchedule(opts: {
  token: string;
  emails: string[];
  start: Date;
  end: Date;
  durationMinutes: number;
  workStartHour: number;
  workEndHour: number;
  label: string;
  mailbox?: string;
  json?: boolean;
  timezone?: string;
}): Promise<{ ok: true } | { ok: false; error: string }> {
  const intervalMinutes = Math.min(30, Math.max(5, opts.durationMinutes));
  const tz = opts.timezone || 'UTC';
  const startIso = opts.timezone
    ? formatDateInTimeZone(opts.start, opts.timezone)
    : opts.start.toISOString().replace(/\.\d{3}Z$/, '');
  const endIso = opts.timezone
    ? formatDateInTimeZone(opts.end, opts.timezone)
    : opts.end.toISOString().replace(/\.\d{3}Z$/, '');

  const sched = await getSchedule(
    opts.token,
    {
      schedules: opts.emails,
      startTime: { dateTime: startIso, timeZone: tz },
      endTime: { dateTime: endIso, timeZone: tz },
      availabilityViewInterval: intervalMinutes
    },
    opts.mailbox?.trim() || undefined,
    tz
  );

  if (!sched.ok || !sched.data?.value) {
    return { ok: false, error: sched.error?.message || 'getSchedule failed' };
  }

  for (const row of sched.data.value) {
    if (row.error?.message) {
      return { ok: false, error: row.error.message };
    }
  }

  const views = sched.data.value.map((s) => s.availabilityView ?? '');
  if (views.some((v) => !v.length)) {
    return { ok: false, error: 'getSchedule returned empty availabilityView for one or more mailboxes' };
  }

  const mergedFree = mergeAvailabilityViewsToMergedFree(views);
  const needSlots = Math.ceil(opts.durationMinutes / intervalMinutes);
  const freeSlots: Array<{ start: string; end: string }> = [];

  for (let i = 0; i <= mergedFree.length - needSlots; i++) {
    let ok = true;
    for (let j = 0; j < needSlots; j++) {
      if (!mergedFree[i + j]) {
        ok = false;
        break;
      }
    }
    if (!ok) continue;
    const t0 = new Date(opts.start.getTime() + i * intervalMinutes * 60 * 1000);
    const t1 = new Date(t0.getTime() + opts.durationMinutes * 60 * 1000);
    if (t1 > opts.end) continue;
    const hour = opts.timezone ? hourInTimeZone(t0, opts.timezone) : t0.getHours();
    if (hour >= opts.workStartHour && hour < opts.workEndHour) {
      freeSlots.push({ start: t0.toISOString(), end: t1.toISOString() });
      i += needSlots - 1;
    }
  }

  if (opts.json) {
    console.log(
      JSON.stringify(
        {
          backend: 'graph',
          strategy: 'getSchedule',
          attendees: opts.emails,
          duration: opts.durationMinutes,
          availabilityViewIntervalMinutes: intervalMinutes,
          dateRange: { start: opts.start.toISOString(), end: opts.end.toISOString() },
          availableSlots: freeSlots.map((s) => ({ start: s.start, end: s.end })),
          // Per-mailbox detail (not just the merged free/busy result): each char in `availabilityView`
          // is one `availabilityViewIntervalMinutes` slot starting at `dateRange.start`
          // (0=free, 1=tentative, 2=busy, 3=out of office, 4=working elsewhere).
          attendeeAvailability: sched.data.value.map((s, i) => ({
            email: opts.emails[i] ?? s.scheduleId,
            availabilityView: s.availabilityView ?? ''
          }))
        },
        null,
        2
      )
    );
    return { ok: true };
  }

  console.log(`\n🗓️  Finding ${opts.durationMinutes}-minute meeting times (Graph getSchedule)`);
  console.log(`   Attendees: ${opts.emails.join(', ')}`);
  console.log(`   Date range: ${opts.label}`);
  console.log('─'.repeat(50));

  if (freeSlots.length === 0) {
    console.log('\n  No available times found for all attendees in this window.');
    console.log('     Try a longer date range, different hours (--start/--end), or shorter duration.');
  } else {
    console.log(`\n  ✅ Found ${freeSlots.length} available slot${freeSlots.length > 1 ? 's' : ''}:\n`);
    const byDay = new Map<string, typeof freeSlots>();
    for (const slot of freeSlots) {
      // Bucket by the requested zone's calendar day when --timezone is set, not the UTC ISO date
      // — otherwise a slot near midnight can print under the wrong day header.
      const day = opts.timezone
        ? formatDateInTimeZone(new Date(slot.start), opts.timezone).slice(0, 10)
        : slot.start.split('T')[0];
      if (!byDay.has(day)) byDay.set(day, []);
      byDay.get(day)?.push(slot);
    }
    for (const [day, slots] of byDay) {
      const dayLabel = new Date(`${day}T12:00:00Z`).toLocaleDateString('en-US', {
        weekday: 'short',
        month: 'short',
        day: 'numeric',
        timeZone: 'UTC'
      });
      console.log(`  ${dayLabel}:`);
      for (const slot of slots) {
        const st = opts.timezone
          ? formatDateInTimeZone(new Date(slot.start), opts.timezone).slice(11, 16)
          : new Date(slot.start).toLocaleTimeString('en-US', {
              hour: '2-digit',
              minute: '2-digit',
              hour12: false
            });
        const en = opts.timezone
          ? formatDateInTimeZone(new Date(slot.end), opts.timezone).slice(11, 16)
          : new Date(slot.end).toLocaleTimeString('en-US', {
              hour: '2-digit',
              minute: '2-digit',
              hour12: false
            });
        console.log(`    🟢 ${st} - ${en}`);
      }
    }
  }
  console.log();
  return { ok: true };
}
