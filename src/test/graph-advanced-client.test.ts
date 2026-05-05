import { describe, expect, test } from 'bun:test';
import {
  type GraphBatchRequestBody,
  graphInvoke,
  graphInvokeText,
  graphPostBatch,
  parseGraphInvokeHeaders
} from '../lib/graph-advanced-client.js';
import { GraphApiError } from '../lib/graph-client.js';

describe('parseGraphInvokeHeaders', () => {
  test('parses valid header lines', () => {
    expect(parseGraphInvokeHeaders(['X-Test: 1', 'Other: value here'])).toEqual({
      'X-Test': '1',
      Other: 'value here'
    });
  });

  test('throws on missing colon', () => {
    expect(() => parseGraphInvokeHeaders(['no-colon'])).toThrow(/Invalid --header format/);
  });

  test('throws on empty name', () => {
    expect(() => parseGraphInvokeHeaders([': value'])).toThrow(/empty name/);
  });
});

describe('graphInvoke path validation', () => {
  test('rejects path without leading slash', async () => {
    const r = await graphInvoke('tok', { method: 'GET', path: 'me', expectJson: false });
    expect(r.ok).toBe(false);
    expect(r.error?.message).toMatch(/start with \//);
  });

  test('rejects path with .. segment', async () => {
    const r = await graphInvoke('tok', { method: 'GET', path: '/me/../admin', expectJson: false });
    expect(r.ok).toBe(false);
    expect(r.error?.message).toMatch(/\.\./);
  });

  test('rejects path with . segment', async () => {
    const r = await graphInvoke('tok', { method: 'GET', path: '/me/./messages', expectJson: false });
    expect(r.ok).toBe(false);
  });

  test('rejects path longer than 8192', async () => {
    const r = await graphInvoke('tok', { method: 'GET', path: `/${'a'.repeat(8200)}`, expectJson: false });
    expect(r.ok).toBe(false);
    expect(r.error?.message).toMatch(/maximum length/i);
  });
});

describe('graphPostBatch validation', () => {
  test('rejects missing requests', async () => {
    const r = await graphPostBatch('tok', {} as GraphBatchRequestBody);
    expect(r.ok).toBe(false);
    expect(r.error?.message).toMatch(/requests/);
  });

  test('rejects more than 20 requests', async () => {
    const requests = Array.from({ length: 21 }, (_, i) => ({ id: String(i) }));
    const r = await graphPostBatch('tok', { requests });
    expect(r.ok).toBe(false);
    expect(r.error?.message).toMatch(/20/);
  });
});

describe('graphInvoke with fetch mock', () => {
  test('returns JSON body on 200', async () => {
    const orig = globalThis.fetch;
    try {
      globalThis.fetch = (async () =>
        new Response(JSON.stringify({ ok: true }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        })) as unknown as typeof fetch;
      const r = await graphInvoke<{ ok: boolean }>('tok', { method: 'GET', path: '/me', pinAccessToken: true });
      expect(r.ok).toBe(true);
      expect(r.data).toEqual({ ok: true });
    } finally {
      globalThis.fetch = orig;
    }
  });

  test('maps GraphApiError to error response', async () => {
    const orig = globalThis.fetch;
    try {
      globalThis.fetch = async () => {
        throw new GraphApiError('nope', 'Custom', 418);
      };
      const r = await graphInvoke('tok', { method: 'GET', path: '/me', pinAccessToken: true });
      expect(r.ok).toBe(false);
      expect(r.error?.message).toContain('nope');
    } finally {
      globalThis.fetch = orig;
    }
  });
});

describe('graphInvokeText with fetch mock', () => {
  test('returns text body', async () => {
    const orig = globalThis.fetch;
    try {
      globalThis.fetch = (async () =>
        new Response('plain', { status: 200, headers: { 'content-type': 'text/plain' } })) as unknown as typeof fetch;
      const r = await graphInvokeText('tok', { method: 'GET', path: '/me', pinAccessToken: true });
      expect(r.ok).toBe(true);
      expect(r.data).toBe('plain');
    } finally {
      globalThis.fetch = orig;
    }
  });
});
