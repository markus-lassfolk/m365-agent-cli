import { describe, expect, it } from 'bun:test';
import { parseDay, parseTimeToDate, toLocalUnzonedISOString, toUTCISOString } from './dates.js';

describe('dates helpers', () => {
  it('parseTimeToDate handles HH:MM and am/pm inputs', () => {
    const base = new Date('2026-03-27T00:00:00');

    expect(parseTimeToDate('13:45', base).getHours()).toBe(13);
    expect(parseTimeToDate('13:45', base).getMinutes()).toBe(45);
    expect(parseTimeToDate('1pm', base).getHours()).toBe(13);
    expect(parseTimeToDate('12am', base).getHours()).toBe(0);
  });

  it('toLocalUnzonedISOString formats as local time without Z', () => {
    const date = new Date(2026, 2, 27, 9, 5, 7); // local time
    const result = toLocalUnzonedISOString(date);
    expect(result).toBe('2026-03-27T09:05:07');
  });

  it('parseTimeToDate throws on invalid input when configured', () => {
    const base = new Date('2026-03-27T00:00:00');
    const opts = { throwOnInvalid: true };

    // Format errors
    expect(() => parseTimeToDate('not-a-time', base, opts)).toThrow(
      'Invalid time format: "not-a-time" — expected HH:MM, H:MM, or H(am|pm)'
    );

    // Value bounds
    expect(() => parseTimeToDate('25:00', base, opts)).toThrow(
      'Invalid time: "25:00" — hours must be 0–23 and minutes 0–59'
    );
    expect(() => parseTimeToDate('9:60', base, opts)).toThrow(
      'Invalid time: "9:60" — hours must be 0–23 and minutes 0–59'
    );

    // AM/PM bounds
    expect(() => parseTimeToDate('13pm', base, opts)).toThrow('Invalid time: "13pm" — 12-hour values must be 1–12');
    expect(() => parseTimeToDate('0am', base, opts)).toThrow('Invalid time: "0am" — 12-hour values must be 1–12');

    // 24-hour hour-only bounds
    expect(() => parseTimeToDate('24', base, opts)).toThrow('Invalid time: "24" — 24-hour values must be 0–23');
  });

  it('toUTCISOString formats as UTC with Z suffix', () => {
    const date = new Date(Date.UTC(2026, 2, 27, 9, 5, 7));
    const result = toUTCISOString(date);
    expect(result).toBe('2026-03-27T09:05:07.000Z');
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
