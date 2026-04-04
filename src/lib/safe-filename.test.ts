import { describe, expect, test } from 'bun:test';
import { safeAttachmentFileName } from './safe-filename.js';

/**
 * Do not assert on real disk I/O here: `src/test/auth.test.ts` globally mocks
 * `node:fs/promises`, so readFile/writeFile behavior differs by test order (CI vs local).
 * `writeInternetShortcutUtf8File` is exercised via mail-graph / calendar download tests.
 */

describe('safeAttachmentFileName', () => {
  test('uses basename and replaces traversal and invalid chars', () => {
    expect(safeAttachmentFileName('a/../evil<x>.pdf', 'fallback.pdf')).toContain('evil');
    expect(safeAttachmentFileName('a/../evil<x>.pdf', 'fallback.pdf')).toMatch(/\.pdf$/);
  });

  test('uses fallback when empty after sanitize', () => {
    expect(safeAttachmentFileName('   ', 'z.pdf')).toBe('z.pdf');
  });

  test('truncates very long names to 255 chars', () => {
    const long = 'a'.repeat(300);
    expect(safeAttachmentFileName(long, 'f').length).toBe(255);
  });
});
