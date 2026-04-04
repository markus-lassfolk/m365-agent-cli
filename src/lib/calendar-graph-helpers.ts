/**
 * Shared helpers for calendar CLI commands using Microsoft Graph (list/filter by mailbox).
 */

import type { GraphCalendarEvent } from './graph-calendar-client.js';
import { callGraph } from './graph-client.js';

export async function graphGetMailboxOrMeEmail(token: string, mailbox?: string): Promise<string | undefined> {
  if (mailbox?.trim()) {
    return mailbox.trim().toLowerCase();
  }
  const r = await callGraph<{ mail?: string; userPrincipalName?: string }>(token, '/me?$select=mail,userPrincipalName');
  const v = r.data?.mail || r.data?.userPrincipalName;
  return v?.toLowerCase();
}

export function graphFilterOrganizerEvents(events: GraphCalendarEvent[], myEmail: string): GraphCalendarEvent[] {
  const me = myEmail.toLowerCase();
  return events.filter((e) => {
    if (e.isCancelled) return false;
    if (e.isOrganizer === true) return true;
    const org = e.organizer?.emailAddress?.address?.toLowerCase();
    return org === me;
  });
}

/** Single calendar day as [start, end) in ISO UTC for `calendarView`. */
export function graphDayRangeIso(baseDate: Date): { start: string; end: string } {
  const start = new Date(baseDate);
  start.setHours(0, 0, 0, 0);
  const end = new Date(start);
  end.setDate(end.getDate() + 1);
  return { start: start.toISOString(), end: end.toISOString() };
}

export function graphNonResourceAttendeeCount(e: GraphCalendarEvent): number {
  const org = e.organizer?.emailAddress?.address?.toLowerCase();
  return (
    e.attendees?.filter((a) => {
      const addr = a.emailAddress?.address?.toLowerCase();
      if (!addr) return false;
      if ((a as { type?: string }).type === 'resource') return false;
      if (addr === org) return false;
      return true;
    }).length ?? 0
  );
}

/**
 * True if `e` is the event with id `idFromUser`, or an occurrence whose `seriesMasterId` matches (series master id).
 */
export function graphEventMatchesOccurrenceFilter(e: GraphCalendarEvent, idFromUser: string): boolean {
  const low = idFromUser.trim().toLowerCase();
  if (e.id.toLowerCase() === low) return true;
  const master = e.seriesMasterId?.trim().toLowerCase();
  return master === low;
}

export function graphFilterPendingInvitations(
  events: GraphCalendarEvent[],
  attendeeEmail: string
): GraphCalendarEvent[] {
  const me = attendeeEmail.toLowerCase();
  return events.filter((e) => {
    if (e.isCancelled) return false;
    if (e.isOrganizer === true) return false;
    const org = e.organizer?.emailAddress?.address?.toLowerCase();
    if (org === me) return false;
    const selfAtt = e.attendees?.find((a) => a.emailAddress?.address?.toLowerCase() === me);
    const resp = (selfAtt?.status?.response || e.responseStatus?.response || 'none').toLowerCase();
    if (resp === 'accepted' || resp === 'declined' || resp === 'tentativelyaccepted' || resp === 'organizer') {
      return false;
    }
    return true;
  });
}
