import { describe, expect, it } from 'bun:test';

const token = 'tok';
const betaUrl = 'https://graph.microsoft.com/beta';

describe('graph-excel-comments-client always targets the Graph beta root', () => {
  // workbookComment has no v1.0 equivalent, so every one of these functions must hit the beta
  // root unconditionally (no `--beta` flag exists or is needed at the command layer; only
  // GRAPH_BETA_URL can redirect which beta host is used).
  const originalBetaUrl = process.env.GRAPH_BETA_URL;
  const originalFetch = globalThis.fetch;

  function withMockFetch(fn: (urls: string[]) => Promise<void>) {
    return async () => {
      delete process.env.GRAPH_BETA_URL;
      const urls: string[] = [];
      try {
        globalThis.fetch = (async (input: string | URL | Request) => {
          urls.push(typeof input === 'string' ? input : input.toString());
          return new Response(JSON.stringify({ id: 'c1' }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }) as unknown as typeof fetch;
        await fn(urls);
      } finally {
        globalThis.fetch = originalFetch;
        if (originalBetaUrl === undefined) delete process.env.GRAPH_BETA_URL;
        else process.env.GRAPH_BETA_URL = originalBetaUrl;
      }
    };
  }

  it(
    'listExcelWorkbookComments',
    withMockFetch(async (urls) => {
      const { listExcelWorkbookComments } = await import('./graph-excel-comments-client.js');
      await listExcelWorkbookComments(token, 'item-1');
      expect(urls[0].startsWith(betaUrl)).toBe(true);
    })
  );

  it(
    'getExcelWorkbookComment',
    withMockFetch(async (urls) => {
      const { getExcelWorkbookComment } = await import('./graph-excel-comments-client.js');
      await getExcelWorkbookComment(token, 'item-1', 'c1');
      expect(urls[0].startsWith(betaUrl)).toBe(true);
    })
  );

  it(
    'createExcelWorkbookComment',
    withMockFetch(async (urls) => {
      const { createExcelWorkbookComment } = await import('./graph-excel-comments-client.js');
      await createExcelWorkbookComment(token, 'item-1', { text: 'hi' });
      expect(urls[0].startsWith(betaUrl)).toBe(true);
    })
  );

  it(
    'addExcelWorkbookCommentReply',
    withMockFetch(async (urls) => {
      const { addExcelWorkbookCommentReply } = await import('./graph-excel-comments-client.js');
      await addExcelWorkbookCommentReply(token, 'item-1', 'c1', { text: 'reply' });
      expect(urls[0].startsWith(betaUrl)).toBe(true);
    })
  );

  it(
    'patchExcelWorkbookComment',
    withMockFetch(async (urls) => {
      const { patchExcelWorkbookComment } = await import('./graph-excel-comments-client.js');
      await patchExcelWorkbookComment(token, 'item-1', 'c1', { text: 'edited' });
      expect(urls[0].startsWith(betaUrl)).toBe(true);
    })
  );

  it('respects GRAPH_BETA_URL for the beta host', async () => {
    const customBeta = 'https://graph.example-sovereign.com/beta';
    process.env.GRAPH_BETA_URL = customBeta;
    const urls: string[] = [];
    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        return new Response(JSON.stringify({ id: 'c1' }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as unknown as typeof fetch;
      const { getExcelWorkbookComment } = await import('./graph-excel-comments-client.js');
      await getExcelWorkbookComment(token, 'item-1', 'c1');
      expect(urls[0].startsWith(customBeta)).toBe(true);
    } finally {
      globalThis.fetch = originalFetch;
      if (originalBetaUrl === undefined) delete process.env.GRAPH_BETA_URL;
      else process.env.GRAPH_BETA_URL = originalBetaUrl;
    }
  });
});
