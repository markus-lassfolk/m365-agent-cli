import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

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

  test('oauth-token-error.example.json matches Azure AD token endpoint error shape', () => {
    const raw = readFileSync(join(fixturesDir, 'oauth-token-error.example.json'), 'utf8');
    const j = JSON.parse(raw) as {
      error?: string;
      error_description?: string;
      error_codes?: number[];
    };
    expect(j.error).toBe('invalid_grant');
    expect(j.error_description).toContain('AADSTS');
    expect(Array.isArray(j.error_codes)).toBe(true);
  });
});
