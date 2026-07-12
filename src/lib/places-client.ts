import { resolveGraphAuth } from './graph-auth.js';
import {
  callGraph,
  fetchAllPages,
  GraphApiError,
  type GraphResponse,
  graphError,
  graphErrorFromApiError,
  graphResult
} from './graph-client.js';

export interface Place {
  id?: string;
  displayName: string;
  emailAddress?: string;
  address?: {
    street?: string;
    city?: string;
    state?: string;
    countryOrRegion?: string;
    postalCode?: string;
    fullAddress?: string;
  };
  geoCoordinates?: {
    latitude?: number;
    longitude?: number;
  };
  capacity?: number;
  bookingType?: 'standard' | 'reserved';
  tags?: string[];
  building?: string;
  floorNumber?: string;
  isManaged?: boolean;
  isBookable?: boolean;
  phone?: string;
}

export interface RoomList {
  id: string;
  displayName: string;
  emailAddress?: string;
}

async function withAuth<T>(
  fn: (
    token: string
  ) => Promise<{ ok: boolean; data?: T; error?: { message: string; code?: string; status?: number } }>,
  options?: { token?: string; identity?: string }
): Promise<{ ok: boolean; data?: T; error?: { message: string; code?: string; status?: number } }> {
  const auth = await resolveGraphAuth(options);
  if (!auth.success || !auth.token) {
    return graphError(auth.error || 'Authentication failed');
  }
  return fn(auth.token);
}

export async function listPlaceRoomLists(options?: { token?: string; identity?: string }): Promise<{
  ok: boolean;
  data?: RoomList[];
  error?: { message: string; code?: string; status?: number };
}> {
  return withAuth<RoomList[]>(async (token) => {
    return fetchAllPages<RoomList>(token, '/places/microsoft.graph.roomList', 'Failed to list room lists');
  }, options);
}

export async function listRoomsInRoomList(
  roomListEmail: string,
  options?: { token?: string; identity?: string }
): Promise<{
  ok: boolean;
  data?: Place[];
  error?: { message: string; code?: string; status?: number };
}> {
  return withAuth<Place[]>(async (token) => {
    return fetchAllPages<Place>(
      token,
      `/places/${encodeURIComponent(roomListEmail)}/microsoft.graph.roomlist/rooms`,
      'Failed to list rooms'
    );
  }, options);
}

export interface RoomFilters {
  building?: string;
  capacityMin?: number;
  equipment?: string[];
  /** Substring filter on display name, email, building, floor, tags. */
  query?: string;
}

/** Substring match on name, email, building, tags (client-side). */
export function filterPlacesByQuery(places: Place[], query: string): Place[] {
  const q = query.trim().toLowerCase();
  if (!q) return places;
  return places.filter((r) => {
    const name = r.displayName?.toLowerCase() ?? '';
    const email = r.emailAddress?.toLowerCase() ?? '';
    const building = r.building?.toLowerCase() ?? '';
    const floor = String(r.floorNumber ?? '').toLowerCase();
    const tags = (r.tags ?? []).join(' ').toLowerCase();
    return name.includes(q) || email.includes(q) || building.includes(q) || floor.includes(q) || tags.includes(q);
  });
}

export async function findRooms(
  filters?: RoomFilters,
  options?: { token?: string; identity?: string; query?: string }
): Promise<{
  ok: boolean;
  data?: Place[];
  error?: { message: string; code?: string; status?: number };
}> {
  return withAuth<Place[]>(async (token) => {
    const result = await fetchAllPages<Place>(token, '/places/microsoft.graph.room', 'Failed to find rooms');
    if (!result.ok || !result.data) {
      return result;
    }

    let rooms = result.data;

    if (filters?.query?.trim()) {
      rooms = filterPlacesByQuery(rooms, filters.query);
    }

    if (filters?.building) {
      const buildingLower = filters.building.toLowerCase();
      rooms = rooms.filter((r) => r.building?.toLowerCase().includes(buildingLower));
    }

    if (filters?.capacityMin !== undefined) {
      rooms = rooms.filter((r) => r.capacity !== undefined && r.capacity >= filters.capacityMin!);
    }

    if (filters?.equipment && filters.equipment.length > 0) {
      const equipLower = filters.equipment.map((e) => e.toLowerCase());
      rooms = rooms.filter((r) => {
        if (!r.tags || r.tags.length === 0) return false;
        const tagsLower = r.tags.map((t) => t.toLowerCase());
        return equipLower.every((e) => tagsLower.some((t) => t.includes(e)));
      });
    }

    return graphResult(rooms);
  }, options);
}

/** GET /places/{place-id} — rich room/place payload. */
export async function getPlace(token: string, placeId: string): Promise<GraphResponse<Place>> {
  const path = `/places/${encodeURIComponent(placeId)}`;
  try {
    const r = await callGraph<Place>(token, path);
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to get place', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphErrorFromApiError(err);
    return graphError(err instanceof Error ? err.message : 'Failed to get place');
  }
}

export async function isRoomFree(
  token: string,
  roomEmail: string,
  startISO: string,
  endISO: string
): Promise<boolean | null> {
  // Page through the whole window: a busy event on a later calendarView page must not be
  // missed (that would wrongly report the room free). $select=showAs keeps the payload small.
  const path =
    `/users/${encodeURIComponent(roomEmail)}/calendar/calendarView` +
    `?startDateTime=${encodeURIComponent(startISO)}&endDateTime=${encodeURIComponent(endISO)}` +
    `&$select=showAs&$top=100`;
  let result: GraphResponse<Array<{ showAs?: string }>>;
  try {
    result = await fetchAllPages<{ showAs?: string }>(token, path, 'Failed to check room availability');
  } catch (_err) {
    return null;
  }

  if (!result.ok || !result.data) {
    return null;
  }

  const busyEvents = result.data.filter((event) => event.showAs !== 'free');
  return busyEvents.length === 0;
}
