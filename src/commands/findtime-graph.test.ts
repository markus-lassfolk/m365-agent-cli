import { describe, expect, test } from 'bun:test';
import { mergeAvailabilityViewsToMergedFree } from './findtime-graph.js';

describe('mergeAvailabilityViewsToMergedFree', () => {
  test('all free when all mailboxes show 0', () => {
    expect(mergeAvailabilityViewsToMergedFree(['00', '00'])).toEqual([true, true]);
  });

  test('busy if any mailbox is non-zero at that slot', () => {
    expect(mergeAvailabilityViewsToMergedFree(['00', '20'])).toEqual([false, true]);
  });

  test('pads shorter views with busy for trailing slots', () => {
    const m = mergeAvailabilityViewsToMergedFree(['0', '00']);
    expect(m).toEqual([true, false]);
  });
});
