import { describe, expect, test } from 'bun:test';
import { GraphApiError } from './graph-client.js';
import { toJsonError } from './json-error.js';

describe('toJsonError', () => {
  test('falls back for null/undefined', () => {
    expect(toJsonError(null)).toEqual({ message: 'Request failed' });
    expect(toJsonError(undefined, 'custom fallback')).toEqual({ message: 'custom fallback' });
  });

  test('wraps a plain string', () => {
    expect(toJsonError('Something broke')).toEqual({ message: 'Something broke' });
  });

  test('falls back for an empty/whitespace string', () => {
    expect(toJsonError('   ', 'fallback')).toEqual({ message: 'fallback' });
    expect(toJsonError('', 'fallback')).toEqual({ message: 'fallback' });
  });

  test('wraps an Error instance', () => {
    expect(toJsonError(new Error('boom'))).toEqual({ message: 'boom' });
  });

  test('falls back for an Error with an empty message', () => {
    expect(toJsonError(new Error(''), 'fallback')).toEqual({ message: 'fallback' });
  });

  test('preserves code/status/requestId from a real GraphApiError instance (bug regression)', () => {
    const err = new GraphApiError('Not found', 'ItemNotFound', 404, 'req-123');
    expect(toJsonError(err)).toEqual({
      message: 'Not found',
      code: 'ItemNotFound',
      status: 404,
      requestId: 'req-123'
    });
  });

  test('preserves code/status/requestId from a GraphError-shaped object', () => {
    const err = { message: 'Not found', code: 'ItemNotFound', status: 404, requestId: 'abc-123' };
    expect(toJsonError(err)).toEqual({ message: 'Not found', code: 'ItemNotFound', status: 404, requestId: 'abc-123' });
  });

  test('preserves code from an OwaError-shaped object (no status field)', () => {
    const err = { message: 'Item not found', code: 'ErrorItemNotFound' };
    expect(toJsonError(err)).toEqual({ message: 'Item not found', code: 'ErrorItemNotFound' });
  });

  test('marks 429/503/504 as retriable', () => {
    expect(toJsonError({ message: 'slow down', status: 429 }).retriable).toBe(true);
    expect(toJsonError({ message: 'unavailable', status: 503 }).retriable).toBe(true);
    expect(toJsonError({ message: 'gateway', status: 504 }).retriable).toBe(true);
    expect(toJsonError({ message: 'not found', status: 404 }).retriable).toBeUndefined();
  });

  test('marks known throttling codes as retriable regardless of status', () => {
    expect(toJsonError({ message: 'throttled', code: 'tooManyRequests' }).retriable).toBe(true);
    expect(toJsonError({ message: 'unavailable', code: 'serviceNotAvailable' }).retriable).toBe(true);
  });

  test('uses fallbackMessage when the object has no usable message', () => {
    expect(toJsonError({ code: 'X' }, 'fallback')).toEqual({ message: 'fallback', code: 'X' });
    expect(toJsonError({}, 'fallback')).toEqual({ message: 'fallback' });
  });

  test('accepts a bare "error" string property (some response shapes use this instead of "message")', () => {
    expect(toJsonError({ error: 'auth failed' })).toEqual({ message: 'auth failed' });
  });

  test('stringifies an unexpected primitive input', () => {
    expect(toJsonError(42)).toEqual({ message: '42' });
  });
});
