import { describe, expect, it } from 'bun:test';
import { join } from 'node:path';
import {
  parseDay,
  parseTimeToDate,
  toLocalUnzonedISOString,
  toReinterpretedUTCISOString,
  toUTCISOString
} from './dates.js';

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

  it('toReinterpretedUTCISOString preserves the local calendar date/time regardless of host offset (bug regression)', () => {
    // Constructed via the local Date constructor (mirrors an all-day boundary built via
    // .setHours(0,0,0,0)/.setHours(23,59,59,999)) — must round-trip to the same numbers with a
    // literal "Z", not whatever UTC instant this local time happens to correspond to.
    const midnight = new Date(2026, 3, 15, 0, 0, 0, 0); // Apr 15, local midnight
    expect(toReinterpretedUTCISOString(midnight)).toBe('2026-04-15T00:00:00.000Z');

    const endOfDay = new Date(2026, 3, 15, 23, 59, 59, 999); // Apr 15, local 23:59:59.999
    expect(toReinterpretedUTCISOString(endOfDay)).toBe('2026-04-15T23:59:59.999Z');
  });

  it('does not shift the date on a host with a positive UTC offset, unlike plain toISOString (bug regression)', () => {
    // A mid-process `process.env.TZ` mutation is not reliably honored by every Bun build for
    // Date objects created afterward — observed to work in some environments and be silently
    // ignored in others (e.g. CI) despite an identical reported Bun version. Spawn a fresh child
    // process with TZ set from birth instead: every runtime respects that unambiguously via
    // normal env inheritance, so this can't flake regardless of engine-internal timezone caching.
    const modulePath = join(import.meta.dir, 'dates.js');
    const script = `
      import { toReinterpretedUTCISOString } from ${JSON.stringify(modulePath)};
      const midnight = new Date(2026, 3, 15, 0, 0, 0, 0);
      console.log(JSON.stringify({
        plainIsoDay: midnight.toISOString().slice(0, 10),
        reinterpreted: toReinterpretedUTCISOString(midnight)
      }));
    `;
    const result = Bun.spawnSync({
      cmd: [process.execPath, '-e', script],
      env: { ...process.env, TZ: 'Europe/Berlin' }
    });
    expect(result.exitCode).toBe(0);
    const out = JSON.parse(result.stdout.toString());
    // Europe/Berlin is UTC+1/+2 — local midnight on Apr 15 is still Apr 14 in UTC, which is
    // exactly the scenario that shifted all-day events back one day before this fix. The bug
    // this guards against: a naive date.toISOString() lands on the previous UTC day.
    expect(out.plainIsoDay).toBe('2026-04-14');
    expect(out.reinterpreted).toBe('2026-04-15T00:00:00.000Z');
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

  it('parseDay rejects out-of-range padded dates instead of silently rolling over', () => {
    // Without round-trip validation, `2026-02-30` would become Mar 2 and `2026-13-01` Jan 2027.
    expect(() => parseDay('2026-02-30', { throwOnInvalid: true })).toThrow('Invalid day value: 2026-02-30');
    expect(() => parseDay('2026-13-01', { throwOnInvalid: true })).toThrow('Invalid day value: 2026-13-01');
    // A valid date still parses to the exact day.
    const ok = parseDay('2026-02-28');
    expect(ok.getFullYear()).toBe(2026);
    expect(ok.getMonth()).toBe(1);
    expect(ok.getDate()).toBe(28);
  });
});
