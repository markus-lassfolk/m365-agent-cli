import { describe, expect, test } from 'bun:test';
import { graphEventMatchesOccurrenceFilter } from './calendar-graph-helpers.js';
import type { GraphCalendarEvent } from './graph-calendar-client.js';

describe('graphEventMatchesOccurrenceFilter', () => {
  test('matches instance id', () => {
    const e = { id: 'inst1', seriesMasterId: 'master1' } as GraphCalendarEvent;
    expect(graphEventMatchesOccurrenceFilter(e, 'inst1')).toBe(true);
    expect(graphEventMatchesOccurrenceFilter(e, 'other')).toBe(false);
  });

  test('matches series master id on occurrence row', () => {
    const e = { id: 'inst1', seriesMasterId: 'master1' } as GraphCalendarEvent;
    expect(graphEventMatchesOccurrenceFilter(e, 'master1')).toBe(true);
  });
});
