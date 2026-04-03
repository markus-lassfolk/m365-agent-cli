import { describe, expect, test } from 'bun:test';
import {
  businessDaysBackward,
  businessDaysForward,
  calendarDaysBackward,
  calendarDaysForward,
  clipCalendarRangeStartToNow
} from './calendar-range.js';

describe('calendar-range', () => {
  test('calendarDaysForward: 3 days from Thu', () => {
    const anchor = new Date(2026, 3, 2); // Thu Apr 2 2026
    const { start, endExclusive } = calendarDaysForward(anchor, 3);
    expect(start.getDate()).toBe(2);
    expect(endExclusive.getDate()).toBe(5); // exclusive end = Apr 5 00:00
  });

  test('calendarDaysBackward: 3 days ending Wed', () => {
    const anchor = new Date(2026, 3, 8); // Wed Apr 8
    const { start, endExclusive } = calendarDaysBackward(anchor, 3);
    expect(start.getDate()).toBe(6);
    expect(endExclusive.getDate()).toBe(9);
  });

  test('businessDaysForward: 5 days from Thu skips weekend', () => {
    const anchor = new Date(2026, 3, 2); // Thu
    const { start, endExclusive } = businessDaysForward(anchor, 5);
    expect(start.getDay()).toBe(4); // Thu
    const last = new Date(endExclusive);
    last.setDate(last.getDate() - 1);
    expect(last.getDay()).toBe(3); // Wed next week
  });

  test('businessDaysBackward: 5 days ending Wed', () => {
    const anchor = new Date(2026, 3, 8); // Wed
    const { start, endExclusive } = businessDaysBackward(anchor, 5);
    expect(start.getDate()).toBe(2); // Thu prior week
    const last = new Date(endExclusive);
    last.setDate(last.getDate() - 1);
    expect(last.getDate()).toBe(8);
  });

  test('clipCalendarRangeStartToNow: moves start to at when range began at midnight', () => {
    const at = new Date(2026, 3, 2, 14, 30, 0, 0); // Apr 2 2026 14:30 local
    const range = {
      start: new Date(2026, 3, 2, 0, 0, 0, 0).toISOString(),
      end: new Date(2026, 3, 3, 0, 0, 0, 0).toISOString(),
      label: 'Today'
    };
    const out = clipCalendarRangeStartToNow(range, at);
    expect(out.start).toBe(at.toISOString());
    expect(out.end).toBe(range.end);
    expect(out.label).toContain('(from now)');
  });

  test('clipCalendarRangeStartToNow: no label change when range already starts at or after at', () => {
    const at = new Date(2026, 3, 2, 9, 0, 0, 0);
    const range = {
      start: new Date(2026, 3, 3, 0, 0, 0, 0).toISOString(),
      end: new Date(2026, 3, 4, 0, 0, 0, 0).toISOString(),
      label: 'Tomorrow'
    };
    const out = clipCalendarRangeStartToNow(range, at);
    expect(out.label).toBe('Tomorrow');
  });
});
