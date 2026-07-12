import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import {
  findRooms,
  getPlace,
  isRoomFree,
  listPlaceRoomLists,
  listRoomsInRoomList,
  type Place,
  type RoomFilters
} from '../lib/places-client.js';

function parseCapacityFilter(value: string | undefined): number | undefined {
  if (!value) return undefined;
  const cleaned = value.replaceAll('+', '').replaceAll('>', '').trim();
  const parsed = parseInt(cleaned, 10);
  if (Number.isNaN(parsed)) {
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
  .argument('[action]', 'Action: lists, rooms, find, or get')
  .argument('[idOrEmail]', 'Room list SMTP (for rooms) or place id (for get)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--building <name>', 'Filter by building name (for find action)')
  .option('--capacity <num>', 'Minimum room capacity (for find action)')
  .option('--equipment <items>', 'Required equipment tags, comma-separated (for find action)')
  .option('--query <text>', 'Substring filter on name, email, building, floor, tags (for find)')
  .option('--start <iso>', 'Start time ISO string (for find action)')
  .option('--end <iso>', 'End time ISO string (for find action)')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      action: string,
      idOrEmail: string | undefined,
      options: {
        json?: boolean;
        token?: string;
        identity?: string;
        building?: string;
        capacity?: string;
        equipment?: string;
        query?: string;
        start?: string;
        end?: string;
      }
    ) => {
      const authResult = await resolveGraphAuth({ token: options.token, identity: options.identity });
      if (!authResult.success) {
        console.error(`Error: ${authResult.error}`);
        process.exit(1);
      }

      if (action === 'lists' || action === undefined) {
        if (!options.json) console.log('Fetching room lists...');
        const result = await listPlaceRoomLists({ token: authResult.token, identity: options.identity });
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
        console.log('\nTip: Use "m365-agent-cli rooms rooms <email>" to see rooms.\n');
        return;
      }

      if (action === 'rooms') {
        if (!idOrEmail) {
          console.error('Error: rooms action requires a room list email address.');
          console.error('Use "m365-agent-cli rooms lists" to see available room lists.');
          process.exit(1);
        }
        if (!options.json) console.log(`Fetching rooms from list: ${idOrEmail}...`);
        const result = await listRoomsInRoomList(idOrEmail, {
          token: authResult.token,
          identity: options.identity
        });
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

      if (action === 'get') {
        if (!idOrEmail) {
          console.error('Error: get action requires a place id (from rooms lists or rooms rooms output).');
          process.exit(1);
        }
        const result = await getPlace(authResult.token!, idOrEmail);
        if (!result.ok || !result.data) {
          console.error(`Error: ${result.error?.message || 'Failed to get place'}`);
          process.exit(1);
        }
        const place = result.data;
        if (options.json) {
          console.log(JSON.stringify(place, null, 2));
          return;
        }
        console.log(`${place.displayName || '(no name)'}`);
        if (place.emailAddress) console.log(`  ${place.emailAddress}`);
        if (place.capacity != null) console.log(`  Capacity: ${place.capacity}`);
        if (place.bookingType) console.log(`  Booking: ${place.bookingType}`);
        if (place.building) console.log(`  Building: ${place.building}`);
        if (place.floorNumber !== undefined) console.log(`  Floor: ${place.floorNumber}`);
        if (place.tags?.length) console.log(`  Tags: ${place.tags.join(', ')}`);
        if (place.id) console.log(`  ID: ${place.id}`);
        return;
      }

      if (action === 'find') {
        let filters: RoomFilters;
        try {
          filters = {
            building: options.building,
            capacityMin: parseCapacityFilter(options.capacity),
            equipment: parseEquipmentFilter(options.equipment),
            query: options.query
          };
        } catch (err) {
          console.error(`Error: ${err instanceof Error ? err.message : String(err)}`);
          process.exit(1);
        }
        if ((options.start && !options.end) || (!options.start && options.end)) {
          console.error('Error: --start and --end must be used together (both or neither).');
          process.exit(1);
        }
        const hasFilters = !!(
          filters.building ||
          filters.capacityMin !== undefined ||
          (filters.equipment && filters.equipment.length > 0) ||
          filters.query?.trim()
        );
        if (!hasFilters && !(options.start && options.end)) {
          console.error(
            'Error: find needs at least one of --query, --building, --capacity, --equipment, or both --start and --end (availability pass).'
          );
          process.exit(1);
        }
        if (!options.json) console.log('Searching for rooms...');
        const result = await findRooms(filters, { token: authResult.token, identity: options.identity });
        if (!result.ok || !result.data) {
          console.error(`Error: ${result.error?.message || 'Failed to search rooms'}`);
          process.exit(1);
        }
        let availableRooms: Array<Place & { availabilityUnknown?: true }> = result.data;
        if (options.start && options.end) {
          const freeRooms: Array<Place & { availabilityUnknown?: true }> = [];
          let availabilityCheckFailed = false;
          for (const room of availableRooms) {
            if (room.emailAddress) {
              const free = await isRoomFree(authResult.token!, room.emailAddress, options.start, options.end);
              if (free === null) {
                // Availability check failed (permission/API error) — include the room but mark it
                // distinctly from a confirmed-free room, so a --json consumer doesn't book it
                // believing it was actually verified as free.
                availabilityCheckFailed = true;
                freeRooms.push({ ...room, availabilityUnknown: true });
              } else if (free) {
                freeRooms.push(room);
              }
            } else {
              freeRooms.push({ ...room, availabilityUnknown: true });
              availabilityCheckFailed = true;
            }
          }
          availableRooms = freeRooms;
          if (availabilityCheckFailed) {
            console.warn(
              'Warning: Could not check availability for some rooms (insufficient permissions or API error).'
            );
          }
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
          const unknown = room.availabilityUnknown ? ' [availability unknown]' : '';
          console.log(`  - ${room.displayName}${cap}${tags}${unknown}`);
          if (room.emailAddress) console.log(`    ${room.emailAddress}`);
          if (room.building) console.log(`    Building: ${room.building}`);
          console.log('');
        }
        return;
      }

      if (action !== undefined) {
        console.error(`Unknown action: "${action}". Use "lists", "rooms", "find", or "get".`);
        process.exit(1);
      }
    }
  );
