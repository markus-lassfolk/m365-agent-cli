/**
 * Graph [Places API](https://learn.microsoft.com/en-us/graph/api/resources/place) helpers for room mailboxes.
 */

import { fetchAllPages } from './graph-client.js';
import { isRoomFree, type Place } from './places-client.js';

export type { Place } from './places-client.js';

export function matchRoomByDisplayName(rooms: Place[], query: string): Place | undefined {
  const q = query.trim().toLowerCase();
  if (!q) return undefined;
  return rooms.find((r) => r.displayName?.toLowerCase().includes(q));
}

export async function listGraphRooms(token: string): Promise<{
  ok: boolean;
  data?: Place[];
  error?: { message: string; code?: string; status?: number };
}> {
  return fetchAllPages<Place>(token, '/places/microsoft.graph.room', 'Failed to list rooms');
}

export async function resolveRoomDisplayNameToPlace(
  token: string,
  query: string
): Promise<{ ok: true; place: Place } | { ok: false; error: string }> {
  const r = await listGraphRooms(token);
  if (!r.ok || !r.data) {
    return { ok: false, error: r.error?.message || 'Failed to list rooms' };
  }
  const place = matchRoomByDisplayName(r.data, query);
  const email = place?.emailAddress?.trim();
  if (!place || !email) {
    return { ok: false, error: `Room not found: ${query}` };
  }
  return { ok: true, place };
}

const ROOM_FREE_PROBE_CONCURRENCY = 4;

/** First room with a mailbox that appears free in the given window (calendarView heuristic). */
export async function findFirstAvailableRoomGraph(
  token: string,
  start: Date,
  end: Date
): Promise<{ email: string; name: string } | null> {
  const r = await listGraphRooms(token);
  if (!r.ok || !r.data?.length) {
    return null;
  }
  const startISO = start.toISOString();
  const endISO = end.toISOString();
  const rooms = r.data.filter((room) => room.emailAddress?.trim());
  for (let i = 0; i < rooms.length; i += ROOM_FREE_PROBE_CONCURRENCY) {
    const chunk = rooms.slice(i, i + ROOM_FREE_PROBE_CONCURRENCY);
    const hits = await Promise.all(
      chunk.map(async (room) => {
        const email = room.emailAddress!.trim();
        const free = await isRoomFree(token, email, startISO, endISO);
        if (free === true) {
          return { email, name: room.displayName?.trim() || email };
        }
        return null;
      })
    );
    const found = hits.find((h) => h !== null);
    if (found) return found;
  }
  return null;
}
