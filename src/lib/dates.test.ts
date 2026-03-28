import { describe, expect, it } from 'bun:test';
import { parseDay, parseTimeToDate, toLocalISOString } from './dates.js';

describe('dates helpers', () => {
  it('parseTimeToDate handles HH:MM and am/pm inputs', () => {
    const base = new Date('2026-03-27T00:00:00');

    expect(parseTimeToDate('13:45', base).getHours()).toBe(13);
    expect(parseTimeToDate('13:45', base).getMinutes()).toBe(45);
    expect(parseTimeToDate('1pm', base).getHours()).toBe(13);
    expect(parseTimeToDate('12am', base).getHours()).toBe(0);
  });

  it('toLocalISOString formats with local timezone offset', () => {
    const date = new Date(2026, 2, 27, 9, 5, 7);
    // Must include a +HH:MM or -HH:MM offset suffix (not Z, not absent)
    const result = toLocalISOString(date);
    expect(result).toMatch(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}[+-]\d{2}:\d{2}$/);
    expect(result).toMatch(/^2026-03-27T09:05:07[+-]\d{2}:\d{2}$/);
  });

  it('parseDay supports relative values and weekday directions', () => {
    const base = new Date('2026-03-25T10:00:00'); // Wednesday

    expect(parseDay('today', { baseDate: base }).getDate()).toBe(25);
    expect(parseDay('tomorrow', { baseDate: base }).getDate()).toBe(26);
    expect(parseDay('yesterday', { baseDate: base }).getDate()).toBe(24);

    const nextMonday = parseDay('monday', { baseDate: base, weekdayDirection: 'next' });
    expect(nextMonday.toISOString().slice(0, 10)).toBe('2026-03-30');

    const prevMonday = parseDay('monday', { baseDate: base, weekdayDirection: 'previous' });
    expect(prevMonday.toISOString().slice(0, 10)).toBe('2026-03-23');

    const forwardMonday = parseDay('monday', {
      baseDate: new Date('2026-03-30T10:00:00'),
      weekdayDirection: 'nearestForward'
    });
    expect(forwardMonday.toISOString().slice(0, 10)).toBe('2026-03-30');
  });

  it('parseDay throws on invalid input when configured', () => {
    expect(() => parseDay('not-a-date', { throwOnInvalid: true })).toThrow('Invalid day value: not-a-date');
  });
});
