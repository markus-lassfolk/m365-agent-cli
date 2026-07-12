import { describe, expect, test } from 'bun:test';
import { join } from 'node:path';
import { safeAttachmentFileName } from './safe-filename.js';

/**
 * Do not assert on real disk I/O here: `src/test/graph-auth.test.ts` uses `mock.module` for Graph auth.
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

  test('neutralizes deep traversal payloads so join() stays inside the output dir', () => {
    // Exact attack strings a malicious sender could set as the attachment name.
    for (const payload of [
      '../../../../home/user/.bashrc',
      '..\\..\\..\\Windows\\System32\\drivers\\etc\\hosts',
      '/etc/passwd',
      '../.ssh/authorized_keys',
      '../.config/m365-agent-cli/.env'
    ]) {
      const safe = safeAttachmentFileName(payload, 'attachment');
      // Reduced to a single path component: no separators, no parent refs.
      expect(safe).not.toContain('/');
      expect(safe).not.toContain('\\');
      expect(safe).not.toContain('..');
      // And the resolved path never escapes the trusted output directory.
      const outDir = join('/tmp', 'downloads');
      const resolved = join(outDir, safe);
      expect(resolved.startsWith(`${outDir}/`)).toBe(true);
    }
  });
});
