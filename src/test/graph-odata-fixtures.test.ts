import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, test } from 'bun:test';

const __dirname = dirname(fileURLToPath(import.meta.url));
const fixturesDir = join(__dirname, 'fixtures', 'graph');

describe('Graph OData fixtures', () => {
  test('odata-error-v1.example.json matches Graph error envelope', () => {
    const raw = readFileSync(join(fixturesDir, 'odata-error-v1.example.json'), 'utf8');
    const j = JSON.parse(raw) as {
      error?: { code?: string; message?: string; innerError?: Record<string, unknown> };
    };
    expect(j.error?.code).toBeTruthy();
    expect(j.error?.message).toBeTruthy();
    expect(typeof j.error?.message).toBe('string');
  });

  test('sendMail-post-body.example.json has message + saveToSentItems', () => {
    const raw = readFileSync(join(fixturesDir, 'sendMail-post-body.example.json'), 'utf8');
    const j = JSON.parse(raw) as {
      message?: { subject?: string; body?: unknown; toRecipients?: unknown[] };
      saveToSentItems?: boolean;
    };
    expect(j.message?.subject).toBeDefined();
    expect(Array.isArray(j.message?.toRecipients)).toBe(true);
    expect(j.saveToSentItems).toBe(true);
  });
});
