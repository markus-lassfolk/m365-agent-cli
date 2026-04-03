import { describe, expect, test } from 'bun:test';
import { normalizeGraphDateTimeForParsing } from './graph-datetime.js';

describe('normalizeGraphDateTimeForParsing', () => {
  test('appends Z when timeZone is UTC and dateTime has no offset', () => {
    expect(normalizeGraphDateTimeForParsing('2026-04-01T09:00:00.0000000', 'UTC')).toBe('2026-04-01T09:00:00.0000000Z');
  });

  test('leaves values that already have a Z suffix', () => {
    expect(normalizeGraphDateTimeForParsing('2026-04-01T09:00:00Z', 'UTC')).toBe('2026-04-01T09:00:00Z');
  });

  test('returns empty when dateTime missing', () => {
    expect(normalizeGraphDateTimeForParsing(undefined, 'UTC')).toBe('');
  });
});
