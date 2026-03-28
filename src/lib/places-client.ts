import { resolveGraphAuth } from './graph-auth.js';
import { callGraph, graphResult, graphError, fetchAllPages } from './graph-client.js';

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
  options?: { token?: string }
): Promise<{ ok: boolean; data?: T; error?: { message: string; code?: string; status?: number } }> {
  const auth = await resolveGraphAuth(options);
  if (!auth.success || !auth.token) {
    return graphError(auth.error || 'Authentication failed');
  }
  return fn(auth.token);
}

export async function listPlaceRoomLists(options?: { token?: string }): Promise<{
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
  options?: { token?: string }
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
}

export async function findRooms(
  filters?: RoomFilters,
  options?: { token?: string }
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

export async function isRoomFree(
  token: string,
  roomEmail: string,
  startISO: string,
  endISO: string
): Promise<boolean | null> {
  const result = await callGraph<{ value: Array<{ showAs?: string }> }>(
    token,
    `/users/${encodeURIComponent(roomEmail)}/calendar/calendarView?startDateTime=${encodeURIComponent(
      startISO
    )}&endDateTime=${encodeURIComponent(endISO)}`
  );

  if (!result.ok || !result.data) {
    return null;
  }

  const busyEvents = (result.data.value || []).filter((event) => event.showAs !== 'free');
  return busyEvents.length === 0;
}
