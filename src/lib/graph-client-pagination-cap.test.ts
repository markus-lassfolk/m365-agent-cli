import { describe, expect, it } from 'bun:test';

process.env.GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';

describe('fetchAllPages runaway/cycle guard', () => {
  it('errors with TooManyPages instead of looping forever on a self-referential nextLink', async () => {
    const originalFetch = globalThis.fetch;
    const originalMax = process.env.GRAPH_MAX_PAGES;
    // The cap is read per-call, so setting it here (not at import time) is enough regardless of
    // which test first imported graph-client.js.
    process.env.GRAPH_MAX_PAGES = '3';
    let calls = 0;
    try {
      globalThis.fetch = (async () => {
        calls++;
        return new Response(
          JSON.stringify({
            value: [{ id: 'x' }],
            // Same link every time → would loop forever without the page cap.
            '@odata.nextLink': 'https://graph.microsoft.com/v1.0/things?$skiptoken=loop'
          }),
          { status: 200, headers: { 'content-type': 'application/json' } }
        );
      }) as unknown as typeof fetch;

      const { fetchAllPages } = await import('./graph-client.js');
      const r = await fetchAllPages('tok', '/things', 'fail');
      expect(r.ok).toBe(false);
      expect(r.error?.code).toBe('TooManyPages');
      // Bounded by GRAPH_MAX_PAGES=3 (a handful of calls, not an infinite loop).
      expect(calls).toBeLessThanOrEqual(4);
    } finally {
      globalThis.fetch = originalFetch;
      if (originalMax === undefined) delete process.env.GRAPH_MAX_PAGES;
      else process.env.GRAPH_MAX_PAGES = originalMax;
    }
  });
});
