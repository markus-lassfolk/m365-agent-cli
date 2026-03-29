import { describe, expect, it } from 'bun:test';
import { parseDay, parseTimeToDate, toUTCISOString } from './dates.js';

describe('dates helpers', () => {
  it('parseTimeToDate handles HH:MM and am/pm inputs', () => {
    const base = new Date('2026-03-27T00:00:00');

    expect(parseTimeToDate('13:45', base).getHours()).toBe(13);
    expect(parseTimeToDate('13:45', base).getMinutes()).toBe(45);
    expect(parseTimeToDate('1pm', base).getHours()).toBe(13);
    expect(parseTimeToDate('12am', base).getHours()).toBe(0);
  });

  it('parseTimeToDate throws on invalid hour in HH:MM format', () => {
    const base = new Date('2026-03-27T00:00:00');
    expect(() => parseTimeToDate('25:00', base)).toThrow(
      "Invalid time value: '25:00' — hour must be between 0 and 23."
    );
    expect(() => parseTimeToDate('24:00', base)).toThrow(
      "Invalid time value: '24:00' — hour must be between 0 and 23."
    );
  });

  it('parseTimeToDate throws on invalid minute in HH:MM format', () => {
    const base = new Date('2026-03-27T00:00:00');
    expect(() => parseTimeToDate('13:75', base)).toThrow(
      "Invalid time value: '13:75' — minute must be between 0 and 59."
    );
    expect(() => parseTimeToDate('13:99', base)).toThrow(
      "Invalid time value: '13:99' — minute must be between 0 and 59."
    );
  });

  it('parseTimeToDate throws on invalid hour in H(am/pm) format', () => {
    const base = new Date('2026-03-27T00:00:00');
    expect(() => parseTimeToDate('25pm', base)).toThrow("Invalid time value: '25pm' — hour must be between 1 and 12.");
    expect(() => parseTimeToDate('0am', base)).toThrow("Invalid time value: '0am' — hour must be between 1 and 12.");
  });

  it('parseTimeToDate throws on completely invalid format', () => {
    const base = new Date('2026-03-27T00:00:00');
    expect(() => parseTimeToDate('not-a-time', base)).toThrow(
      "Invalid time value: 'not-a-time' — expected HH:MM, H:MM, or H(am/pm) format."
    );
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
