import { describe, expect, it } from 'bun:test';

const token = 'test-token';
const baseUrl = 'https://graph.microsoft.com/v1.0';

describe('microsoftSearchQuery', () => {
  it('POSTs /search/query with entity types and query string', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const bodies: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        if (init?.body && typeof init.body === 'string') bodies.push(init.body);
        return new Response(
          JSON.stringify({
            value: [
              {
                searchTerms: ['q'],
                hitsContainers: [{ hits: [], total: 0, moreResultsAvailable: false }]
              }
            ]
          }),
          { status: 200, headers: { 'content-type': 'application/json' } }
        );
      }) as typeof fetch;

      const { microsoftSearchQuery } = await import('./graph-microsoft-search.js');
      const r = await microsoftSearchQuery(token, {
        entityTypes: ['message'],
        queryString: 'subject:foo',
        from: 0,
        size: 10
      });

      expect(r.ok).toBe(true);
      expect(urls[0]).toContain('/search/query');
      expect(bodies[0]).toContain('"entityTypes":["message"]');
      expect(bodies[0]).toContain('"queryString":"subject:foo"');
      expect(bodies[0]).toContain('"from":0');
      expect(bodies[0]).toContain('"size":10');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});
