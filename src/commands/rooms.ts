import { Command } from 'commander';
import { resolveAuth } from '../lib/auth.js';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import {
  listPlaceRoomLists,
  listRoomsInRoomList,
  findRooms,
  isRoomFree,
  type Place,
  type RoomList,
  type RoomFilters,
} from '../lib/places-client.js';

function parseCapacityFilter(value: string | undefined): number | undefined {
  if (!value) return undefined;
  return parseInt(value.replace('+', '').replace('>', '').trim(), 10);
}

function parseEquipmentFilter(value: string | undefined): string[] | undefined {
  if (!value) return undefined;
  return value.split(',').map(e => e.trim()).filter(Boolean);
}

export const roomsCommand = new Command('rooms')
  .description('Discover meeting rooms and room lists via Microsoft Places API')
  .option('--building <name>', 'Filter rooms by building name (partial, case-insensitive)')
  .option('--capacity <min>', 'Minimum room capacity (e.g. 10, 10+)')
  .option('--equipment <tags>', 'Required equipment tags, comma-separated (e.g. Whiteboard,Video)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .action(async (options: {
    building?: string; capacity?: string; equipment?: string; json?: boolean; token?: string;
  }) => {
    const hasFilters = !!(options.building || options.capacity || options.equipment);
    if (!hasFilters) {
      const result = await listPlaceRoomLists();
      if (!result.ok || !result.data) {
        if (options.json) console.log(JSON.stringify({ error: result.error?.message || 'Failed to list room lists' }, null, 2));
        else console.error(`Error: ${result.error?.message || 'Failed to list room lists'}`);
        process.exit(1);
      }
      const lists: RoomList[] = result.data;
      if (options.json) { console.log(JSON.stringify({ roomLists: lists }, null, 2)); return; }
      console.log(`\n\u{1F4C1}  Room Lists (${lists.length})\n`);
      if (lists.length === 0) console.log('  No room lists found or access denied.');
      else for (const rl of lists) { console.log(`  \u{1F4CB} ${rl.displayName}`); if (rl.emailAddress) console.log(`     ${rl.emailAddress}`); console.log(''); }
      return;
    }
    const filters: RoomFilters = {
      building: options.building,
      capacityMin: parseCapacityFilter(options.capacity),
      equipment: parseEquipmentFilter(options.equipment),
    };
    const result = await findRooms(filters);
    if (!result.ok || !result.data) {
      if (options.json) console.log(JSON.stringify({ error: result.error?.message || 'Failed to list rooms' }, null, 2));
      else console.error(`Error: ${result.error?.message || 'Failed to list rooms'}`);
      process.exit(1);
    }
    const rooms: Place[] = result.data;
    if (options.json) {
      console.log(JSON.stringify({ filters: { building: options.building || null, capacityMin: filters.capacityMin ?? null, equipment: filters.equipment || null }, count: rooms.length, rooms: rooms.map(r => ({ displayName: r.displayName, emailAddress: r.emailAddress, capacity: r.capacity, bookingType: r.bookingType, building: r.building, floorNumber: r.floorNumber, tags: r.tags || [] })) }, null, 2));
      return;
    }
    const fdesc = [options.building ? `building="${options.building}"` : null, filters.capacityMin !== undefined ? `capacity>=${filters.capacityMin}` : null, filters.equipment?.length ? `equipment=${filters.equipment.join(',')}` : null].filter(Boolean).join(' | ');
    console.log(`\n\u{1F3E0}  Meeting Rooms (${rooms.length})${fdesc ? ` -- ${fdesc}` : ''}\n`);
    if (rooms.length === 0) { console.log('  No rooms match your filters, or access is denied.'); }
    else {
      const byBuilding = new Map<string, Place[]>();
      for (const room of rooms) { const key = room.building || 'No Building'; if (!byBuilding.has(key)) byBuilding.set(key, []); byBuilding.get(key)!.push(room); }
      for (const [building, buildingRooms] of byBuilding) {
        console.log(`  ${building !== 'No Building' ? `\u{1F3D7}  ${building}` : '\u{1F3E0}  Unknown Building'} (${buildingRooms.length})`);
        for (const room of buildingRooms) {
          const tags = room.tags?.length ? ` [${room.tags.join(', ')}]` : '';
          const cap = room.capacity ? ` (cap: ${room.capacity})` : '';
          const floor = room.floorNumber ? `, floor ${room.floorNumber}` : '';
          console.log(`    • ${room.displayName}${cap}${floor}${tags}`);
        }
        console.log('');
      }
    }
  });

roomsCommand
  .command('lists')
  .description('List all room lists')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .action(async (options: { json?: boolean; token?: string }) => {
    const result = await listPlaceRoomLists();
    if (!result.ok || !result.data) {
      if (options.json) console.log(JSON.stringify({ error: result.error?.message || 'Failed to list room lists' }, null, 2));
      else console.error(`Error: ${result.error?.message || 'Failed to list room lists'}`);
      process.exit(1);
    }
    const lists: RoomList[] = result.data;
    if (options.json) { console.log(JSON.stringify({ roomLists: lists }, null, 2)); return; }
    console.log(`\n\u{1F4C1}  Room Lists (${lists.length})\n`);
    if (lists.length === 0) console.log('  No room lists found or access denied.');
    else for (const rl of lists) { console.log(`  \u{1F4CB} ${rl.displayName}`); if (rl.emailAddress) console.log(`     ${rl.emailAddress}`); console.log(''); }
  });

roomsCommand
  .command('rooms')
  .description('List all rooms with optional filters')
  .option('--building <name>', 'Filter by building name')
  .option('--capacity <min>', 'Minimum capacity (e.g. 10, 10+)')
  .option('--equipment <tags>', 'Required equipment tags, comma-separated')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .action(async (options: {
    building?: string; capacity?: string; equipment?: string; json?: boolean; token?: string;
  }) => {
    const filters: RoomFilters = {
      building: options.building,
      capacityMin: parseCapacityFilter(options.capacity),
      equipment: parseEquipmentFilter(options.equipment),
    };
    const result = await findRooms(filters);
    if (!result.ok || !result.data) {
      if (options.json) console.log(JSON.stringify({ error: result.error?.message || 'Failed to list rooms' }, null, 2));
      else console.error(`Error: ${result.error?.message || 'Failed to list rooms'}`);
      process.exit(1);
    }
    const rooms: Place[] = result.data;
    if (options.json) {
      console.log(JSON.stringify({ filters: { building: options.building || null, capacityMin: filters.capacityMin ?? null, equipment: filters.equipment || null }, count: rooms.length, rooms: rooms.map(r => ({ displayName: r.displayName, emailAddress: r.emailAddress, capacity: r.capacity, building: r.building, floorNumber: r.floorNumber, tags: r.tags || [] })) }, null, 2));
      return;
    }
    console.log(`\n\u{1F3E0}  Meeting Rooms (${rooms.length})\n`);
    if (rooms.length === 0) { console.log('  No rooms match your filters.'); }
    else {
      const byBuilding = new Map<string, Place[]>();
      for (const room of rooms) { const key = room.building || 'No Building'; if (!byBuilding.has(key)) byBuilding.set(key, []); byBuilding.get(key)!.push(room); }
      for (const [building, buildingRooms] of byBuilding) {
        console.log(`  ${building !== 'No Building' ? `\u{1F3D7}  ${building}` : '\u{1F3E0}  Unknown Building'} (${buildingRooms.length})`);
        for (const room of buildingRooms) {
          const tags = room.tags?.length ? ` [${room.tags.join(', ')}]` : '';
          const cap = room.capacity ? ` (cap: ${room.capacity})` : '';
          const floor = room.floorNumber ? `, floor ${room.floorNumber}` : '';
          console.log(`    • ${room.displayName}${cap}${floor}${tags}`);
        }
        console.log('');
      }
    }
  });

roomsCommand
  .command('find')
  .description('Find an available room for a time slot')
  .argument('<start>', 'Start time (e.g., 13:00, 1pm)')
  .argument('<end>', 'End time (e.g., 14:00, 2pm)')
  .option('--day <day>', 'Day (today, tomorrow, monday-sunday, YYYY-MM-DD)', 'today')
  .option('--building <name>', 'Filter by building name')
  .option('--capacity <min>', 'Minimum capacity (e.g. 10, 10+)')
  .option('--equipment <tags>', 'Required equipment tags, comma-separated')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .action(async (startTime: string, endTime: string, options: {
    day?: string; building?: string; capacity?: string; equipment?: string; json?: boolean; token?: string;
  }) => {
    const authResult = await resolveAuth({ token: options.token });
    if (!authResult.success || !authResult.token) {
      if (options.json) console.log(JSON.stringify({ error: authResult.error }, null, 2));
      else console.error(`Error: ${authResult.error}`);
      process.exit(1);
    }
    const { parseDay, parseTimeToDate } = await import('../lib/dates.js');
    const baseDate = parseDay(options.day || 'today');
    const start = parseTimeToDate(startTime, baseDate);
    const end = parseTimeToDate(endTime, baseDate);
    const filters: RoomFilters = {
      building: options.building,
      capacityMin: parseCapacityFilter(options.capacity),
      equipment: parseEquipmentFilter(options.equipment),
    };
    const result = await findRooms(filters);
    if (!result.ok || !result.data) {
      if (options.json) console.log(JSON.stringify({ error: result.error?.message || 'Failed to find rooms' }, null, 2));
      else console.error(`Error: ${result.error?.message || 'Failed to find rooms'}`);
      process.exit(1);
    }
    const rooms: Place[] = result.data;
    if (options.json) {
      console.log(JSON.stringify({ timeSlot: { start: start.toISOString(), end: end.toISOString() }, filters, count: rooms.length, rooms: rooms.map(r => ({ displayName: r.displayName, emailAddress: r.emailAddress, capacity: r.capacity, building: r.building, floorNumber: r.floorNumber, tags: r.tags || [] })) }, null, 2));
      return;
    }
    const timeStr = (d: Date) => d.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit', hour12: false });
    console.log(`\n\u{1F50D}  Finding available rooms for ${start.toLocaleDateString()} ${timeStr(start)} - ${timeStr(end)}`);
    console.log(`   Checking ${rooms.length} room(s)...\n`);
    const availableRooms: Place[] = [];
    for (const room of rooms) {
      if (!room.emailAddress) continue;
      const free = await isRoomFree(authResult.token, room.emailAddress, start.toISOString(), end.toISOString());
      if (free) availableRooms.push(room);
    }
    if (availableRooms.length === 0) { console.log('  \u{274C}  No available rooms found for this time slot.'); }
    else {
      console.log(`  \u{2705}  ${availableRooms.length} available room(s):\n`);
      for (const room of availableRooms) {
        const tags = room.tags?.length ? ` [${room.tags.join(', ')}]` : '';
        const cap = room.capacity ? ` (cap: ${room.capacity})` : '';
        console.log(`    • ${room.displayName}${cap}${tags}`);
        if (room.emailAddress) console.log(`      ${room.emailAddress}`);
        console.log('');
      }
    }
  });
