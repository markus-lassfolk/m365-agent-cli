import { describe, expect, test } from 'bun:test';
import { matchRoomByDisplayName } from './graph-places-helpers.js';
import type { Place } from './places-client.js';

describe('matchRoomByDisplayName', () => {
  test('matches substring case-insensitive', () => {
    const rooms: Place[] = [
      { displayName: 'Conference A', emailAddress: 'a@x.com' },
      { displayName: 'Lobby', emailAddress: 'b@x.com' }
    ];
    expect(matchRoomByDisplayName(rooms, 'conf')?.emailAddress).toBe('a@x.com');
  });
});
