/**
 * Microsoft Graph path for `findtime` via findMeetingTimes (same family as `suggest`),
 * with fallback to calendar/getSchedule + merged availability views.
 */

import { type AttendeeBase, findMeetingTimes, getSchedule } from '../lib/graph-schedule.js';

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
}): Promise<{ ok: true } | { ok: false; error: string }> {
  const attendeePayload: AttendeeBase[] = opts.emails.map((address) => ({
    type: 'required' as const,
    emailAddress: { address }
  }));

  const startDateTime = opts.start.toISOString().replace('Z', '');
  const endDateTime = opts.end.toISOString().replace('Z', '');

  const result = await findMeetingTimes(
    opts.token,
    {
      attendees: attendeePayload,
      meetingDuration: `PT${opts.durationMinutes}M`,
      timeConstraint: {
        activityDomain: 'work',
        timeSlots: [
          {
            start: { dateTime: startDateTime, timeZone: 'UTC' },
            end: { dateTime: endDateTime, timeZone: 'UTC' }
          }
        ]
      },
      minimumAttendeePercentage: 100,
      isOrganizerOptional: false,
      returnSuggestionReasons: true
    },
    opts.mailbox?.trim() || undefined
  );

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
        reason: s.suggestionReason
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
    const slot = s.meetingTimeSlot?.start?.dateTime;
    if (!slot) return false;
    const hour = new Date(slot).getHours();
    return hour >= opts.workStartHour && hour < opts.workEndHour;
  });

  if (filtered.length === 0) {
    console.log('\n  No available times in the selected working hours window.');
    console.log('     Try a longer date range, different hours (--start/--end), or shorter duration.');
  } else {
    console.log(`\n  ✅ Found ${filtered.length} suggested slot${filtered.length > 1 ? 's' : ''}:\n`);
    for (const s of filtered) {
      const st = s.meetingTimeSlot?.start?.dateTime;
      const en = s.meetingTimeSlot?.end?.dateTime;
      const startLabel = st ? new Date(st).toLocaleString() : '?';
      const endLabel = en ? new Date(en).toLocaleString() : '?';
      const conf = s.confidence !== undefined ? ` (${s.confidence}% confidence)` : '';
      console.log(`    🟢 ${startLabel} – ${endLabel}${conf}`);
      if (s.suggestionReason) console.log(`       ${s.suggestionReason}`);
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
}): Promise<{ ok: true } | { ok: false; error: string }> {
  const intervalMinutes = Math.min(30, Math.max(5, opts.durationMinutes));
  const startIso = opts.start.toISOString().replace(/\.\d{3}Z$/, '');
  const endIso = opts.end.toISOString().replace(/\.\d{3}Z$/, '');

  const sched = await getSchedule(
    opts.token,
    {
      schedules: opts.emails,
      startTime: { dateTime: startIso, timeZone: 'UTC' },
      endTime: { dateTime: endIso, timeZone: 'UTC' },
      availabilityViewInterval: intervalMinutes
    },
    opts.mailbox?.trim() || undefined
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
    const hour = t0.getHours();
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
          availableSlots: freeSlots.map((s) => ({ start: s.start, end: s.end }))
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
      const day = slot.start.split('T')[0];
      if (!byDay.has(day)) byDay.set(day, []);
      byDay.get(day)?.push(slot);
    }
    for (const [day, slots] of byDay) {
      const dayLabel = new Date(day).toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric' });
      console.log(`  ${dayLabel}:`);
      for (const slot of slots) {
        const st = new Date(slot.start).toLocaleTimeString('en-US', {
          hour: '2-digit',
          minute: '2-digit',
          hour12: false
        });
        const en = new Date(slot.end).toLocaleTimeString('en-US', {
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
