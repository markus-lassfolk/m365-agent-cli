import { describe, expect, test } from 'bun:test';
import { graphEventStartMs } from '../lib/graph-calendar-recurrence.js';

describe('graphEventStartMs', () => {
  test('parses Graph UTC-style dateTime without Z', () => {
    expect(graphEventStartMs({ dateTime: '2026-04-15T09:00:00.0000000', timeZone: 'UTC' })).toBe(
      Date.parse('2026-04-15T09:00:00.000Z')
    );
  });

  test('parses ISO with Z', () => {
    expect(graphEventStartMs({ dateTime: '2026-04-15T09:00:00.000Z' })).toBe(Date.parse('2026-04-15T09:00:00.000Z'));
  });
});
