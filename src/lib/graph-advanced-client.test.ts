import { describe, expect, it } from 'bun:test';
import type { GraphBatchRequestBody } from './graph-advanced-client.js';

/** Bypasses `mock.module('./graph-advanced-client.js', …)` used in copilot invoke tests (relative `?query` only; absolute file URLs still resolve to the mock). */
// @ts-expect-error Bun accepts a query on the import specifier; tsc only types the bare `.js` path.
const ga = await import('./graph-advanced-client.js?contractTest=1');
const {
  graphInvoke,
  graphInvokeText,
  graphPostBatch,
  graphBatchAll,
  chunkGraphBatchRequests,
  parseGraphInvokeHeaders
} = ga;

describe('parseGraphInvokeHeaders', () => {
  it('parses Name: value with first colon as separator', () => {
    expect(parseGraphInvokeHeaders(['ConsistencyLevel: eventual', 'Prefer: outlook.timezone="UTC"'])).toEqual({
      ConsistencyLevel: 'eventual',
      Prefer: 'outlook.timezone="UTC"'
    });
  });

  it('trims name and value', () => {
    expect(parseGraphInvokeHeaders(['  X-Test :  hello  '])).toEqual({ 'X-Test': 'hello' });
  });

  it('rejects line without colon', () => {
    expect(() => parseGraphInvokeHeaders(['bad'])).toThrow(/Invalid --header/);
  });

  it('rejects empty header name', () => {
    expect(() => parseGraphInvokeHeaders([': only-value'])).toThrow(/empty name/);
  });
});

describe('graphPostBatch', () => {
  it('rejects more than 20 requests without calling fetch', async () => {
    const originalFetch = globalThis.fetch;
    let fetchCalled = false;
    try {
      globalThis.fetch = (() => {
        fetchCalled = true;
        return Promise.resolve(new Response('{}', { status: 200 }));
      }) as unknown as typeof fetch;

      const requests = Array.from({ length: 21 }, (_, i) => ({ id: String(i), method: 'GET', url: '/me' }));
      const r = await graphPostBatch('t', { requests });
      expect(r.ok).toBe(false);
      expect(r.error?.code).toBe('InvalidBatch');
      expect(fetchCalled).toBe(false);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('rejects batch body without requests array', async () => {
    const r = await graphPostBatch('t', {} as GraphBatchRequestBody);
    expect(r.ok).toBe(false);
    expect(r.error?.message).toMatch(/requests/);
  });

  it('posts $batch and returns JSON', async () => {
    process.env.GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async () =>
        new Response(JSON.stringify({ responses: [] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        })) as unknown as typeof fetch;
      const r = await graphPostBatch('tok', { requests: [{ id: '1', method: 'GET', url: '/me' }] });
      expect(r.ok).toBe(true);
      expect((r.data as { responses: unknown[] }).responses).toEqual([]);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('chunkGraphBatchRequests', () => {
  it('splits into chunks of at most 20, preserving order', () => {
    const requests = Array.from({ length: 45 }, (_, i) => ({ id: String(i) }));
    const chunks = chunkGraphBatchRequests(requests);
    expect(chunks.length).toBe(3);
    expect(chunks[0].length).toBe(20);
    expect(chunks[1].length).toBe(20);
    expect(chunks[2].length).toBe(5);
    expect(chunks.flat()).toEqual(requests);
  });

  it('returns a single chunk when under the limit', () => {
    const requests = Array.from({ length: 5 }, (_, i) => ({ id: String(i) }));
    expect(chunkGraphBatchRequests(requests)).toEqual([requests]);
  });
});

describe('graphBatchAll', () => {
  it('returns empty responses without calling fetch for an empty request list', async () => {
    const originalFetch = globalThis.fetch;
    let fetchCalled = false;
    try {
      globalThis.fetch = (() => {
        fetchCalled = true;
        return Promise.resolve(new Response('{}', { status: 200 }));
      }) as unknown as typeof fetch;
      const r = await graphBatchAll('tok', []);
      expect(r.ok).toBe(true);
      expect(r.data).toEqual({ responses: [] });
      expect(fetchCalled).toBe(false);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('splits 45 requests into 3 sequential /$batch POSTs and merges responses in order', async () => {
    process.env.GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';
    const originalFetch = globalThis.fetch;
    const postedCounts: number[] = [];
    try {
      globalThis.fetch = (async (_input: string | URL | Request, init?: RequestInit) => {
        const parsed = JSON.parse(String(init?.body ?? '{}')) as { requests: Array<{ id: string }> };
        postedCounts.push(parsed.requests.length);
        const responses = parsed.requests.map((req) => ({ id: req.id, status: 200, body: {} }));
        return new Response(JSON.stringify({ responses }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as unknown as typeof fetch;

      const requests = Array.from({ length: 45 }, (_, i) => ({ id: String(i), method: 'GET', url: '/me' }));
      const r = await graphBatchAll('tok', requests);
      expect(r.ok).toBe(true);
      expect(postedCounts).toEqual([20, 20, 5]);
      const data = r.data as { responses: Array<{ id: string }> };
      expect(data.responses.map((x) => x.id)).toEqual(requests.map((x) => x.id));
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('rejects duplicate ids without calling fetch', async () => {
    const originalFetch = globalThis.fetch;
    let fetchCalled = false;
    try {
      globalThis.fetch = (() => {
        fetchCalled = true;
        return Promise.resolve(new Response('{}', { status: 200 }));
      }) as unknown as typeof fetch;
      const r = await graphBatchAll('tok', [
        { id: 'a', method: 'GET', url: '/me' },
        { id: 'a', method: 'GET', url: '/me/messages' }
      ]);
      expect(r.ok).toBe(false);
      expect(r.error?.code).toBe('InvalidBatch');
      expect(r.error?.message).toMatch(/unique/);
      expect(fetchCalled).toBe(false);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('rejects a dependsOn chain that would cross a chunk boundary, without calling fetch', async () => {
    const originalFetch = globalThis.fetch;
    let fetchCalled = false;
    try {
      globalThis.fetch = (() => {
        fetchCalled = true;
        return Promise.resolve(new Response('{}', { status: 200 }));
      }) as unknown as typeof fetch;
      const requests = Array.from({ length: 21 }, (_, i) => ({ id: String(i), method: 'GET', url: '/me' }));
      // request "20" (in the 2nd chunk) depends on "0" (in the 1st chunk).
      (requests[20] as { dependsOn?: string[] }).dependsOn = ['0'];
      const r = await graphBatchAll('tok', requests);
      expect(r.ok).toBe(false);
      expect(r.error?.code).toBe('InvalidBatch');
      expect(r.error?.message).toMatch(/depends on/);
      expect(fetchCalled).toBe(false);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('stops at the first chunk that errors and does not send later chunks', async () => {
    process.env.GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';
    const originalFetch = globalThis.fetch;
    let calls = 0;
    try {
      globalThis.fetch = (async () => {
        calls += 1;
        return new Response('nope', { status: 500 });
      }) as unknown as typeof fetch;
      const requests = Array.from({ length: 25 }, (_, i) => ({ id: String(i), method: 'GET', url: '/me' }));
      const r = await graphBatchAll('tok', requests);
      expect(r.ok).toBe(false);
      expect(calls).toBe(1);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('preserves responses from earlier successful chunks when a later chunk fails outright (bug regression)', async () => {
    process.env.GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';
    const originalFetch = globalThis.fetch;
    let calls = 0;
    try {
      globalThis.fetch = (async (_input: string | URL | Request, init?: RequestInit) => {
        calls += 1;
        if (calls === 1) {
          const parsed = JSON.parse(String(init?.body ?? '{}')) as { requests: Array<{ id: string }> };
          const responses = parsed.requests.map((req) => ({ id: req.id, status: 200, body: {} }));
          return new Response(JSON.stringify({ responses }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        return new Response('server error', { status: 500 });
      }) as unknown as typeof fetch;

      const requests = Array.from({ length: 25 }, (_, i) => ({ id: String(i), method: 'GET', url: '/me' }));
      const r = await graphBatchAll('tok', requests);
      expect(r.ok).toBe(false);
      expect(calls).toBe(2);
      const data = r.data as { responses: Array<{ id: string }> } | undefined;
      const responses = data?.responses ?? [];
      expect(responses.map((x) => x.id)).toEqual(requests.slice(0, 20).map((x) => x.id));
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('graphInvoke / graphInvokeText', () => {
  const baseUrl = 'https://graph.microsoft.com/v1.0';

  it('graphInvoke returns error when path is not relative', async () => {
    const r = await graphInvoke('tok', { path: 'me', method: 'GET', pinAccessToken: true });
    expect(r.ok).toBe(false);
    expect(r.error?.message).toMatch(/start with \//);
  });

  it('graphInvoke GET /me succeeds', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async () =>
        new Response(JSON.stringify({ id: 'u1' }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        })) as unknown as typeof fetch;
      const r = await graphInvoke('tok', { path: '/me', method: 'GET', pinAccessToken: true });
      expect(r.ok).toBe(true);
      expect((r.data as { id: string }).id).toBe('u1');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('graphInvoke POST sends JSON body', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    let body = '';
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async (_input: string | URL | Request, init?: RequestInit) => {
        body = String(init?.body ?? '');
        return new Response(JSON.stringify({ ok: true }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as unknown as typeof fetch;
      const r = await graphInvoke('tok', {
        path: '/me/sendMail',
        method: 'POST',
        body: { x: 1 },
        pinAccessToken: true
      });
      expect(r.ok).toBe(true);
      expect(JSON.parse(body)).toEqual({ x: 1 });
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('graphInvokeText reads non-JSON body', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async () =>
        new Response('plain', {
          status: 200,
          headers: { 'content-type': 'text/plain' }
        })) as unknown as typeof fetch;
      const r = await graphInvokeText('tok', { path: '/me', method: 'GET', pinAccessToken: true });
      expect(r.ok).toBe(true);
      expect(r.data).toBe('plain');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});
