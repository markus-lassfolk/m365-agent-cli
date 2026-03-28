import { describe, expect, it } from 'bun:test';

const token = 'test-token';
const baseUrl = 'https://graph.microsoft.com/v1.0';

describe('searchFiles query encoding', () => {
  it('encodes single quotes in search query before interpolation', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;

    const fetchCalls: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        fetchCalls.push(typeof input === 'string' ? input : input.toString());
        return new Response(JSON.stringify({ value: [] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;

      const { searchFiles } = await import('./graph-client.js');
      await searchFiles(token, "') and name='anything");

      expect(fetchCalls).toHaveLength(1);
      expect(fetchCalls[0]).toContain("/me/drive/root/search(q='%27%29%20and%20name%3D%27anything')");
      expect(fetchCalls[0]).not.toContain("q=') and name='anything'");
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});
