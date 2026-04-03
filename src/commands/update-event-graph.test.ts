import { describe, expect, test } from 'bun:test';
import { mergeGraphEventAttendees } from './update-event-graph.js';
import type { GraphCalendarEvent } from '../lib/graph-calendar-client.js';

describe('mergeGraphEventAttendees', () => {
  test('removes and adds', () => {
    const display = {
      id: 'e1',
      attendees: [
        { emailAddress: { address: 'a@x.com', name: 'A' }, type: 'required' },
        { emailAddress: { address: 'b@x.com' }, type: 'required' }
      ]
    } as unknown as GraphCalendarEvent;
    const out = mergeGraphEventAttendees(display, ['c@x.com'], ['a@x.com']);
    expect(out).toHaveLength(2);
    expect(out.map((x) => x.emailAddress.address)).toContain('b@x.com');
    expect(out.map((x) => x.emailAddress.address)).toContain('c@x.com');
  });

  test('preserves resource type', () => {
    const display = {
      id: 'e1',
      attendees: [{ emailAddress: { address: 'room@building.com' }, type: 'resource' }]
    } as unknown as GraphCalendarEvent;
    const out = mergeGraphEventAttendees(display, [], []);
    expect(out[0].type).toBe('resource');
  });

  test('roomResource replaces old resource', () => {
    const display = {
      id: 'e1',
      attendees: [
        { emailAddress: { address: 'oldroom@x.com' }, type: 'resource' },
        { emailAddress: { address: 'a@x.com' }, type: 'required' }
      ]
    } as unknown as GraphCalendarEvent;
    const out = mergeGraphEventAttendees(display, [], [], { email: 'newroom@x.com', name: 'New' });
    expect(out).toHaveLength(2);
    expect(out.find((x) => x.type === 'resource')?.emailAddress.address).toBe('newroom@x.com');
    expect(out.find((x) => x.emailAddress.address === 'a@x.com')).toBeDefined();
  });
});
