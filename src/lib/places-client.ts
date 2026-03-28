import { resolveGraphAuth } from './graph-auth.js';
import { callGraph, graphResult, graphError } from './graph-client.js';

export interface Place {
  id?: string;
  displayName: string;
  emailAddress?: string;
  capacity?: number;
  bookingType?: string;
  tags?: string[];
  building?: string;
  floorNumber?: string;
  phone?: string;
}

export interface RoomList {
  id: string;
  displayName: string;
  emailAddress?: string;
}

interface PlacesApiResponse<T> {
  value: T[];
  '@odata.nextLink'?: string;
}
export interface RoomFilters {
  building?: string;
  capacityMin?: number;
  equipment?: string[];
}
type PlacesResult<T> = { ok: boolean; data?: T[]; error?: { message: string; code?: string; status?: number } };

async function gf<T>(
  path: string
): Promise<{ ok: boolean; data?: PlacesApiResponse<T>; error?: { message: string; code?: string; status?: number } }> {
  const auth = await resolveGraphAuth();
  if (!auth.success || !auth.token) return graphError(auth.error || 'Auth failed') as ReturnType<typeof graphError>;
  const r = await callGraph<PlacesApiResponse<T>>(auth.token, path);
  if (!r.ok || !r.data)
    return graphError(r.error?.message || 'API failed', r.error?.code, r.error?.status) as ReturnType<
      typeof graphError
    >;
  return graphResult(r.data);
}

export async function listPlaceRoomLists(): Promise<PlacesResult<RoomList>> {
  const r = await gf<RoomList>('/places/microsoft.graph.roomList');
  if (!r.ok || !r.data) return r as PlacesResult<RoomList>;
  return { ok: true, data: r.data.value };
}

export async function listRoomsInRoomList(emailAddress: string): Promise<PlacesResult<Place>> {
  const r = await gf<Place>(`/places/${encodeURIComponent(emailAddress)}/microsoft.graph.roomList/rooms`);
  if (!r.ok || !r.data) return r as PlacesResult<Place>;
  return { ok: true, data: r.data.value };
}

export async function findRooms(f?: RoomFilters): Promise<PlacesResult<Place>> {
  const r = await gf<Place>('/places/microsoft.graph.room');
  if (!r.ok || !r.data) return r as PlacesResult<Place>;
  let rooms = r.data.value;
  if (f?.building) {
    const bl = f.building.toLowerCase();
    rooms = rooms.filter((p) => p.building?.toLowerCase().includes(bl));
  }
  if (f?.capacityMin !== undefined)
    rooms = rooms.filter((p) => p.capacity !== undefined && p.capacity >= f.capacityMin!);
  if (f?.equipment?.length) {
    const el = f.equipment.map((e) => e.toLowerCase());
    rooms = rooms.filter((p) => {
      if (!p.tags?.length) return false;
      const tl = p.tags.map((t) => t.toLowerCase());
      return el.every((e) => tl.some((t) => t.includes(e)));
    });
  }
  return { ok: true, data: rooms };
}

export async function isRoomFree(token: string, email: string, start: string, end: string): Promise<boolean> {
  const r = await callGraph<{ value: unknown[] }>(
    token,
    `/users/${encodeURIComponent(email)}/calendar/calendarView?startDateTime=${encodeURIComponent(start)}&endDateTime=${encodeURIComponent(end)}`
  );
  if (!r.ok || !r.data) return false;
  return r.data.value.length === 0;
}
