import { expect, test } from 'bun:test';
import { describeMailGraphUnhandledCombination } from '../commands/mail-graph.js';

test('describeMailGraphUnhandledCombination: download + read', () => {
  const msg = describeMailGraphUnhandledCombination({
    limit: '10',
    page: '1',
    output: '.',
    read: 'abc',
    download: 'xyz'
  });
  expect(msg).toContain('--read');
  expect(msg).toContain('--download');
});

test('describeMailGraphUnhandledCombination: two mutating groups', () => {
  const msg = describeMailGraphUnhandledCombination({
    limit: '10',
    page: '1',
    output: '.',
    markRead: 'a',
    move: 'b',
    to: 'archive'
  });
  expect(msg).toContain('does not combine');
});
