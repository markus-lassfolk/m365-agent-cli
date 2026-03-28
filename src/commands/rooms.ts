import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import {
  listPlaceRoomLists,
  listRoomsInRoomList,
  findRooms,
  isRoomFree,
  type Place,
  type RoomList,
  type RoomFilters
} from '../lib/places-client.js';

function parseCapacityFilter(value: string | undefined): number | undefined {
  if (!value) return undefined;
  const cleaned = value.replace('+', '').replace('>', '').trim();
  const parsed = parseInt(cleaned, 10);
  if (isNaN(parsed)) {
    throw new Error(`Invalid capacity value: "${value}". Must be a number.`);
  }
  return parsed;
}

function parseEquipmentFilter(value: string | undefined): string[] | undefined {
  if (!value) return undefined;
  return value
    .split(',')
    .map((s) => s.trim())
    .filter(Boolean);
}

export const roomsCommand = new Command('rooms')
  .description('Discover rooms and room lists via Microsoft Graph Places API')
  .argument('[action]', 'Action: lists, rooms, or find')
  .argument('[roomListId]', 'Room list ID (required for rooms action)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--building <name>', 'Filter by building name (for find action)')
  .option('--capacity <num>', 'Minimum room capacity (for find action)')
  .option('--equipment <items>', 'Required equipment tags, comma-separated (for find action)')
  .option('--start <iso>', 'Start time ISO string (for find action)')
  .option('--end <iso>', 'End time ISO string (for find action)')
  .action(
    async (
      action: string,
      roomListId: string | undefined,
      options: {
        json?: boolean;
        token?: string;
        building?: string;
        capacity?: string;
        equipment?: string;
        start?: string;
        end?: string;
      }
    ) => {
      const authResult = await resolveGraphAuth({ token: options.token });
      if (!authResult.success) {
        console.error(`Error: ${authResult.error}`);
        process.exit(1);
      }

      if (action === 'lists' || action === undefined) {
        console.log('Fetching room lists...');
        const result = await listPlaceRoomLists({ token: authResult.token });
        if (!result.ok || !result.data) {
          console.error(`Error: ${result.error?.message || 'Failed to fetch room lists'}`);
          process.exit(1);
        }
        const lists = result.data;
        if (options.json) {
          console.log(JSON.stringify({ roomLists: lists }, null, 2));
          return;
        }
        if (lists.length === 0) {
          console.log('No room lists found.');
          return;
        }
        console.log(`\nRoom Lists (${lists.length}):\n`);
        console.log('-'.repeat(70));
        for (const list of lists) {
          console.log(`  ${list.displayName || '(no name)'}`);
          if (list.emailAddress) console.log(`    ${list.emailAddress}`);
          console.log(`    ID: ${list.id}`);
          console.log('');
        }
        console.log('-'.repeat(70));
        console.log('\nTip: Use "clippy rooms rooms <listId>" to see rooms.\n');
        return;
      }

      if (action === 'rooms') {
        if (!roomListId) {
          console.error('Error: rooms action requires a room list ID.');
          console.error('Use "clippy rooms lists" to see available room lists.');
          process.exit(1);
        }
        console.log(`Fetching rooms from list: ${roomListId}...`);
        const result = await listRoomsInRoomList(roomListId, { token: authResult.token });
        if (!result.ok || !result.data) {
          console.error(`Error: ${result.error?.message || 'Failed to fetch rooms'}`);
          process.exit(1);
        }
        const rooms = result.data;
        if (options.json) {
          console.log(JSON.stringify({ rooms }, null, 2));
          return;
        }
        if (rooms.length === 0) {
          console.log('No rooms found.');
          return;
        }
        console.log(`\nRooms (${rooms.length}):\n`);
        console.log('-'.repeat(70));
        for (const room of rooms) {
          console.log(`  ${room.displayName || '(no name)'}`);
          if (room.emailAddress) console.log(`    ${room.emailAddress}`);
          if (room.capacity) console.log(`    Capacity: ${room.capacity}`);
          if (room.bookingType) console.log(`    Booking type: ${room.bookingType}`);
          if (room.building) console.log(`    Building: ${room.building}`);
          if (room.floorNumber !== undefined) console.log(`    Floor: ${room.floorNumber}`);
          if (room.tags && room.tags.length > 0) console.log(`    Tags: ${room.tags.join(', ')}`);
          console.log(`    ID: ${room.id}`);
          console.log('');
        }
        console.log('-'.repeat(70));
        return;
      }

      if (action === 'find') {
        let filters: RoomFilters;
        try {
          filters = {
            building: options.building,
            capacityMin: parseCapacityFilter(options.capacity),
            equipment: parseEquipmentFilter(options.equipment)
          };
        } catch (err) {
          console.error(`Error: ${err instanceof Error ? err.message : String(err)}`);
          process.exit(1);
        }
        const hasFilters = !!(filters.building || filters.capacityMin !== undefined || filters.equipment);
        if (!hasFilters) {
          console.error('Error: find action requires at least one filter (--building, --capacity, or --equipment).');
          process.exit(1);
        }
        console.log('Searching for rooms...');
        const result = await findRooms(filters, { token: authResult.token });
        if (!result.ok || !result.data) {
          console.error(`Error: ${result.error?.message || 'Failed to search rooms'}`);
          process.exit(1);
        }
        let availableRooms = result.data;
        if (options.start && options.end) {
          const freeRooms: Place[] = [];
          for (const room of availableRooms) {
            if (room.emailAddress) {
              const free = await isRoomFree(authResult.token!, room.emailAddress, options.start, options.end);
              if (free) freeRooms.push(room);
            }
          }
          availableRooms = freeRooms;
        }
        if (options.json) {
          console.log(JSON.stringify({ rooms: availableRooms }, null, 2));
          return;
        }
        if (availableRooms.length === 0) {
          console.log('No matching rooms found.');
          return;
        }
        console.log(`\nMatching rooms (${availableRooms.length}):\n`);
        for (const room of availableRooms) {
          const tags = room.tags?.length ? ` [${room.tags.join(', ')}]` : '';
          const cap = room.capacity ? ` (cap: ${room.capacity})` : '';
          console.log(`  - ${room.displayName}${cap}${tags}`);
          if (room.emailAddress) console.log(`    ${room.emailAddress}`);
          if (room.building) console.log(`    Building: ${room.building}`);
          console.log('');
        }
        return;
      }

      if (action !== undefined) {
        console.error(`Unknown action: "${action}". Use "lists", "rooms", or "find".`);
        process.exit(1);
      }
    }
  );
