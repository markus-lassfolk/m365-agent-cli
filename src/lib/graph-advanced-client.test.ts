import { describe, expect, test } from 'bun:test';
import { assertSafeGraphRelativePath } from './graph-advanced-client.js';
import { GraphApiError } from './graph-client.js';

describe('assertSafeGraphRelativePath', () => {
  test('allows .. inside OData query string', () => {
    const p = assertSafeGraphRelativePath(`/me/messages?$filter=subject eq 'a..b'`);
    expect(p).toContain('..');
  });

  test('rejects .. in path segments', () => {
    expect(() => assertSafeGraphRelativePath('/me/../users')).toThrow(GraphApiError);
  });

  test('rejects . segment', () => {
    expect(() => assertSafeGraphRelativePath('/me/./messages')).toThrow(GraphApiError);
  });
});
