import { afterEach, describe, expect, test } from 'bun:test';
import { exitGraphError } from './copilot.js';

describe('exitGraphError (bug regression: was plain stderr text, not JSON)', () => {
  const originalLog = console.log;
  const originalError = console.error;
  const originalExit = process.exit;
  let logs: string[];
  let errors: string[];
  let exitCode: number | undefined;

  function setup(): void {
    logs = [];
    errors = [];
    exitCode = undefined;
    console.log = ((s: string) => logs.push(s)) as typeof console.log;
    console.error = ((s: string) => errors.push(s)) as typeof console.error;
    process.exit = ((code?: number) => {
      exitCode = code;
      throw new Error(`process.exit(${code})`);
    }) as never;
  }

  afterEach(() => {
    console.log = originalLog;
    console.error = originalError;
    process.exit = originalExit;
  });

  test('prints the structured error envelope on stdout, not plain text on stderr', () => {
    setup();
    expect(() => exitGraphError('Access denied')).toThrow();
    expect(exitCode).toBe(1);
    expect(errors).toHaveLength(0);
    expect(logs).toHaveLength(1);
    expect(JSON.parse(logs[0])).toEqual({ error: { message: 'Access denied' } });
  });

  test('falls back to "Unknown error" when message is undefined', () => {
    setup();
    expect(() => exitGraphError(undefined)).toThrow();
    expect(JSON.parse(logs[0])).toEqual({ error: { message: 'Unknown error' } });
  });
});
