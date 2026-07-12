import { afterEach, describe, expect, test } from 'bun:test';
import { failBookings } from './bookings.js';

describe('failBookings', () => {
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

  test('--json prints the structured error envelope on stdout, not plain text on stderr (bug regression)', () => {
    setup();
    expect(() => failBookings(true, 'Auth error', 'Missing EWS_CLIENT_ID')).toThrow();
    expect(exitCode).toBe(1);
    expect(errors).toHaveLength(0);
    expect(logs).toHaveLength(1);
    expect(JSON.parse(logs[0])).toEqual({ error: { message: 'Missing EWS_CLIENT_ID' } });
  });

  test('--json preserves code/status from a GraphError-shaped object', () => {
    setup();
    expect(() =>
      failBookings(true, 'Error', { message: 'Access denied', code: 'ErrorAccessDenied', status: 403 })
    ).toThrow();
    expect(JSON.parse(logs[0])).toEqual({
      error: { message: 'Access denied', code: 'ErrorAccessDenied', status: 403 }
    });
  });

  test('non-json prints "Auth error: <msg>" to stderr, not stdout', () => {
    setup();
    expect(() => failBookings(false, 'Auth error', 'bad token')).toThrow();
    expect(exitCode).toBe(1);
    expect(logs).toHaveLength(0);
    expect(errors).toEqual(['Auth error: bad token']);
  });

  test('non-json prints "Error: <msg>" for a GraphError-shaped object', () => {
    setup();
    expect(() => failBookings(false, 'Error', { message: 'Not found' })).toThrow();
    expect(errors).toEqual(['Error: Not found']);
  });
});
