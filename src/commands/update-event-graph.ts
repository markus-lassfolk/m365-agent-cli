/**
 * Microsoft Graph PATCH path for `update-event` (subset of fields; attachments via POST after create/PATCH).
 */

import { toLocalUnzonedISOString, toUTCISOString } from '../lib/dates.js';
import type { GraphCalendarEvent } from '../lib/graph-calendar-client.js';

function mapSensitivity(
  s: 'Normal' | 'Personal' | 'Private' | 'Confidential'
): 'normal' | 'personal' | 'private' | 'confidential' {
  const m: Record<string, 'normal' | 'personal' | 'private' | 'confidential'> = {
    Normal: 'normal',
    Personal: 'personal',
    Private: 'private',
    Confidential: 'confidential'
  };
  return m[s] ?? 'normal';
}

/**
 * Build full `attendees` array for Graph PATCH from current event + add/remove lists.
 * Preserves `type` (required / optional / resource) for existing attendees.
 * When `roomResource` is set, existing **resource** rows are dropped and the new room is appended.
 */
export function mergeGraphEventAttendees(
  display: GraphCalendarEvent,
  add: string[],
  remove: string[],
  roomResource?: { email: string; name?: string }
): Array<{ emailAddress: { address: string; name?: string }; type: 'required' | 'optional' | 'resource' }> {
  const removeSet = new Set(remove.map((e) => e.trim().toLowerCase()).filter(Boolean));
  const addEmails = add.map((e) => e.trim()).filter(Boolean);

  const result: Array<{
    emailAddress: { address: string; name?: string };
    type: 'required' | 'optional' | 'resource';
  }> = [];
  const seen = new Set<string>();

  for (const a of display.attendees ?? []) {
    const addr = a.emailAddress?.address?.trim();
    if (!addr) continue;
    const low = addr.toLowerCase();
    if (removeSet.has(low)) continue;
    const rawType = (a as { type?: string }).type;
    const graphType: 'required' | 'optional' | 'resource' =
      rawType?.toLowerCase() === 'optional'
        ? 'optional'
        : rawType?.toLowerCase() === 'resource'
          ? 'resource'
          : 'required';
    if (roomResource && graphType === 'resource') {
      continue;
    }
    result.push({
      emailAddress: {
        address: addr,
        ...(a.emailAddress?.name ? { name: a.emailAddress.name } : {})
      },
      type: graphType
    });
    seen.add(low);
  }

  for (const email of addEmails) {
    const low = email.toLowerCase();
    if (removeSet.has(low)) continue;
    if (seen.has(low)) continue;
    result.push({ emailAddress: { address: email }, type: 'required' });
    seen.add(low);
  }

  if (roomResource) {
    const low = roomResource.email.trim().toLowerCase();
    if (!removeSet.has(low) && !seen.has(low)) {
      result.push({
        emailAddress: {
          address: roomResource.email.trim(),
          ...(roomResource.name ? { name: roomResource.name } : {})
        },
        type: 'resource'
      });
    }
  }

  return result;
}

export function buildGraphUpdatePatch(input: {
  display: GraphCalendarEvent;
  title?: string;
  description?: string;
  newStart?: Date;
  newEnd?: Date;
  timezone?: string;
  location?: string;
  allDay?: boolean;
  sensitivity?: 'Normal' | 'Personal' | 'Private' | 'Confidential';
  categories?: string[];
  clearCategories?: boolean;
  teams?: boolean;
  noTeams?: boolean;
  addAttendee?: string[];
  removeAttendee?: string[];
  /** Set/replace conference room (resource attendee + usually location). */
  roomResource?: { email: string; name?: string };
}): Record<string, unknown> {
  const patch: Record<string, unknown> = {};
  const d = input.display;

  if (input.title !== undefined) {
    patch.subject = input.title;
  }
  if (input.description !== undefined) {
    patch.body = { contentType: 'text' as const, content: input.description };
  }
  if (input.location !== undefined) {
    patch.location = { displayName: input.location };
  }

  if (input.newStart !== undefined && input.newEnd !== undefined) {
    const tz = input.timezone?.trim() || d.start?.timeZone || 'UTC';
    if (input.timezone?.trim()) {
      patch.start = { dateTime: toLocalUnzonedISOString(input.newStart), timeZone: tz };
      patch.end = { dateTime: toLocalUnzonedISOString(input.newEnd), timeZone: tz };
    } else {
      patch.start = { dateTime: toUTCISOString(input.newStart).replace(/\.\d{3}Z$/, ''), timeZone: 'UTC' };
      patch.end = { dateTime: toUTCISOString(input.newEnd).replace(/\.\d{3}Z$/, ''), timeZone: 'UTC' };
    }
  }

  if (input.allDay !== undefined) {
    patch.isAllDay = input.allDay;
  }

  if (input.sensitivity !== undefined) {
    patch.sensitivity = mapSensitivity(input.sensitivity);
  }

  if (input.clearCategories) {
    patch.categories = [];
  } else if (input.categories && input.categories.length > 0) {
    patch.categories = input.categories;
  }

  if (input.teams === true) {
    patch.isOnlineMeeting = true;
    patch.onlineMeetingProvider = 'teamsForBusiness';
  } else if (input.noTeams === true) {
    patch.isOnlineMeeting = false;
  }

  const add = input.addAttendee ?? [];
  const rem = input.removeAttendee ?? [];
  const room = input.roomResource;
  if (add.length > 0 || rem.length > 0 || room) {
    patch.attendees = mergeGraphEventAttendees(input.display, add, rem, room);
  }

  return patch;
}
