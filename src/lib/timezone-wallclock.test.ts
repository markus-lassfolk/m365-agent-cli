import { describe, expect, it } from 'bun:test';
import { assertValidTimeZone, formatDateInTimeZone, hourInTimeZone } from './timezone-wallclock.js';

describe('assertValidTimeZone', () => {
  it('returns the zone name unchanged when valid', () => {
    expect(assertValidTimeZone('America/New_York')).toBe('America/New_York');
    expect(assertValidTimeZone('UTC')).toBe('UTC');
  });

  it('throws a clean error for an unrecognized zone name', () => {
    expect(() => assertValidTimeZone('Not/AZone')).toThrow(/Invalid IANA time zone/);
  });
});

describe('formatDateInTimeZone', () => {
  it('formats a UTC instant as UTC wall-clock time', () => {
    const d = new Date('2026-06-15T12:30:00.000Z');
    expect(formatDateInTimeZone(d, 'UTC')).toBe('2026-06-15T12:30:00');
  });

  it('formats the same instant differently in a negative-offset zone', () => {
    const d = new Date('2026-06-15T12:30:00.000Z');
    // America/New_York is UTC-4 in June (EDT).
    expect(formatDateInTimeZone(d, 'America/New_York')).toBe('2026-06-15T08:30:00');
  });

  it('formats the same instant differently in a positive-offset zone', () => {
    const d = new Date('2026-06-15T12:30:00.000Z');
    // Asia/Tokyo is UTC+9, no DST.
    expect(formatDateInTimeZone(d, 'Asia/Tokyo')).toBe('2026-06-15T21:30:00');
  });

  it('rolls over to the next day when the zone offset crosses midnight', () => {
    const d = new Date('2026-06-15T23:30:00.000Z');
    expect(formatDateInTimeZone(d, 'Asia/Tokyo')).toBe('2026-06-16T08:30:00');
  });

  it('normalizes a midnight hour to 00 instead of 24', () => {
    const d = new Date('2026-06-15T00:00:00.000Z');
    expect(formatDateInTimeZone(d, 'UTC')).toBe('2026-06-15T00:00:00');
  });
});

describe('hourInTimeZone', () => {
  it('returns the UTC hour for the UTC zone', () => {
    expect(hourInTimeZone(new Date('2026-06-15T12:30:00.000Z'), 'UTC')).toBe(12);
  });

  it('returns the shifted hour for a non-UTC zone', () => {
    expect(hourInTimeZone(new Date('2026-06-15T12:30:00.000Z'), 'America/New_York')).toBe(8);
    expect(hourInTimeZone(new Date('2026-06-15T12:30:00.000Z'), 'Asia/Tokyo')).toBe(21);
  });

  it('normalizes midnight to 0', () => {
    expect(hourInTimeZone(new Date('2026-06-15T00:00:00.000Z'), 'UTC')).toBe(0);
  });
});
