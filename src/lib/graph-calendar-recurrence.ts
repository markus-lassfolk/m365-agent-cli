/**
 * Truncate a recurring calendar series on Microsoft Graph (EWS “delete this and all future” parity).
 */

import { graphNonResourceAttendeeCount } from './calendar-graph-helpers.js';
import {
  cancelCalendarEvent,
  deleteCalendarEvent,
  type GraphCalendarEvent,
  type GraphPatternedRecurrence,
  getEvent,
  listEventInstances,
  updateCalendarEvent
} from './graph-calendar-client.js';
import type { GraphResponse } from './graph-client.js';
import { graphError, graphResult } from './graph-client.js';

/** Compare two Graph event starts for ordering; UTC-safe when `timeZone` is UTC. */
export function graphEventStartMs(st?: { dateTime?: string; timeZone?: string }): number {
  if (!st?.dateTime) return NaN;
  const raw = st.dateTime.trim();
  if (/[zZ]$|[+-]\d{2}:\d{2}$/.test(raw)) return Date.parse(raw);
  const base = raw.replace(/\.\d+$/, '');
  const tz = (st.timeZone || '').toUpperCase();
  if (tz === 'UTC' || tz === 'GMT' || tz === 'ETC/UTC') return Date.parse(`${base}Z`);
  return Date.parse(base);
}

function isoInstantBeforeCut(cut: GraphCalendarEvent): string {
  const ms = graphEventStartMs(cut.start);
  if (!Number.isFinite(ms)) {
    return new Date(0).toISOString();
  }
  return new Date(ms - 1).toISOString();
}

function recurrenceRangeStartIso(master: GraphCalendarEvent, rec: GraphPatternedRecurrence): string {
  const sd = rec.range?.startDate;
  if (sd?.match(/^\d{4}-\d{2}-\d{2}$/)) {
    return `${sd}T00:00:00.0000000`;
  }
  return master.start?.dateTime ?? sd ?? new Date(0).toISOString();
}

function endDateFromLastKeptOccurrence(last: GraphCalendarEvent): string {
  const dt = last.start?.dateTime;
  if (!dt) return new Date().toISOString().slice(0, 10);
  const day = dt.slice(0, 10);
  if (day.match(/^\d{4}-\d{2}-\d{2}$/)) return day;
  return new Date(graphEventStartMs(last.start)).toISOString().slice(0, 10);
}

export type TruncateSeriesResult = {
  action: 'truncated' | 'deleted' | 'cancelled';
  attendeesNotified?: number;
};

/**
 * Delete “this and all future” for a recurring series: PATCH master recurrence to end before the cut occurrence,
 * or delete/cancel the whole series if the cut is the first occurrence (or the event is not recurring).
 */
export async function truncateRecurringSeriesBeforeCut(
  token: string,
  user: string | undefined,
  target: GraphCalendarEvent,
  options: { forceDelete?: boolean }
): Promise<GraphResponse<TruncateSeriesResult>> {
  const masterId = target.seriesMasterId ?? target.id;
  const masterRes = await getEvent(
    token,
    masterId,
    user,
    'recurrence,start,end,type,subject,isAllDay,organizer,attendees'
  );
  if (!masterRes.ok || !masterRes.data) {
    return { ok: false, error: masterRes.error };
  }
  const master = masterRes.data;
  const rec = master.recurrence;
  const cutMs = graphEventStartMs(target.start);

  if (!Number.isFinite(cutMs)) {
    return graphError('Target event has invalid or missing start date');
  }

  if (!rec?.pattern) {
    const attCount = graphNonResourceAttendeeCount(master);
    if (attCount > 0 && !options.forceDelete) {
      const c = await cancelCalendarEvent(token, masterId, { user, comment: '' });
      if (!c.ok) return { ok: false, error: c.error };
      return graphResult({ action: 'cancelled', attendeesNotified: attCount });
    }
    const d = await deleteCalendarEvent(token, masterId, user);
    if (!d.ok) return { ok: false, error: d.error };
    return graphResult({ action: 'deleted' });
  }

  const startBound = recurrenceRangeStartIso(master, rec);
  const endBound = isoInstantBeforeCut(target);

  const instRes = await listEventInstances(token, masterId, startBound, endBound, {
    user,
    select: 'start,end,type,isCancelled'
  });
  if (!instRes.ok || !instRes.data) {
    return { ok: false, error: instRes.error };
  }

  const beforeCut = instRes.data
    .filter((e) => !e.isCancelled && graphEventStartMs(e.start) < cutMs)
    .sort((a, b) => graphEventStartMs(a.start) - graphEventStartMs(b.start));
  const lastKept = beforeCut[beforeCut.length - 1];

  if (!lastKept) {
    const attCount = graphNonResourceAttendeeCount(master);
    if (attCount > 0 && !options.forceDelete) {
      const c = await cancelCalendarEvent(token, masterId, { user, comment: '' });
      if (!c.ok) return { ok: false, error: c.error };
      return graphResult({ action: 'cancelled', attendeesNotified: attCount });
    }
    const d = await deleteCalendarEvent(token, masterId, user);
    if (!d.ok) return { ok: false, error: d.error };
    return graphResult({ action: 'deleted' });
  }

  const endDateStr = endDateFromLastKeptOccurrence(lastKept);
  const range = rec.range;
  if (!range?.startDate) {
    return graphError('Series master has no recurrence range startDate');
  }

  const newRange: GraphPatternedRecurrence['range'] = {
    type: 'endDate',
    startDate: range.startDate,
    endDate: endDateStr
  };

  const patchRec: GraphPatternedRecurrence = {
    pattern: rec.pattern,
    range: newRange
  };

  const patchRes = await updateCalendarEvent(token, masterId, { recurrence: patchRec }, user);
  if (!patchRes.ok) {
    return { ok: false, error: patchRes.error };
  }
  return graphResult({ action: 'truncated' });
}
