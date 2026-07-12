import { afterEach, beforeEach, describe, expect, test } from 'bun:test';
import { haltForDryRun, isDryRunActive, previewableBody } from './dry-run.js';

describe('isDryRunActive', () => {
  const original = process.env.M365_DRY_RUN;
  afterEach(() => {
    if (original === undefined) delete process.env.M365_DRY_RUN;
    else process.env.M365_DRY_RUN = original;
  });

  test('false when unset', () => {
    delete process.env.M365_DRY_RUN;
    expect(isDryRunActive()).toBe(false);
  });

  test('true for "1" and "true" (case-insensitive)', () => {
    process.env.M365_DRY_RUN = '1';
    expect(isDryRunActive()).toBe(true);
    process.env.M365_DRY_RUN = 'TRUE';
    expect(isDryRunActive()).toBe(true);
  });

  test('false for other values', () => {
    process.env.M365_DRY_RUN = '0';
    expect(isDryRunActive()).toBe(false);
    process.env.M365_DRY_RUN = 'yes';
    expect(isDryRunActive()).toBe(false);
  });
});

describe('previewableBody', () => {
  test('parses a JSON string body', () => {
    expect(previewableBody('{"a":1}')).toEqual({ a: 1 });
  });

  test('returns a non-JSON string as-is', () => {
    expect(previewableBody('plain text')).toBe('plain text');
  });

  test('returns undefined for null/undefined', () => {
    expect(previewableBody(null)).toBeUndefined();
    expect(previewableBody(undefined)).toBeUndefined();
  });

  test('summarizes a binary body by byte length', () => {
    expect(previewableBody(new Uint8Array([1, 2, 3]))).toBe('<binary body, 3 bytes>');
  });

  test('summarizes a stream body without reading it', () => {
    const fakeStream = { getReader: () => ({}) } as unknown as ReadableStream;
    expect(previewableBody(fakeStream)).toBe('<stream body>');
  });
});

describe('haltForDryRun', () => {
  let originalExit: typeof process.exit;
  let originalLog: typeof console.log;
  let logged: string[];
  let exitCode: number | undefined;

  beforeEach(() => {
    logged = [];
    exitCode = undefined;
    originalExit = process.exit;
    originalLog = console.log;
    console.log = ((s: string) => {
      logged.push(s);
    }) as typeof console.log;
    process.exit = ((code?: number) => {
      exitCode = code;
      throw new Error(`process.exit(${code})`);
    }) as never;
  });

  afterEach(() => {
    process.exit = originalExit;
    console.log = originalLog;
  });

  test('prints exactly one JSON object with dryRun:true and exits 0', () => {
    expect(() => haltForDryRun({ backend: 'graph', method: 'POST', url: 'https://x' })).toThrow();
    expect(exitCode).toBe(0);
    expect(logged).toHaveLength(1);
    const parsed = JSON.parse(logged[0]);
    expect(parsed).toEqual({ dryRun: true, backend: 'graph', method: 'POST', url: 'https://x' });
  });
});
