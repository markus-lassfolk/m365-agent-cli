import { afterEach, describe, expect, it } from 'bun:test';
import { tryGraphDraftMutations } from './drafts-graph.js';

const token = 'tok';

describe('tryGraphDraftMutations --edit', () => {
  const originalFetch = globalThis.fetch;
  const originalLog = console.log;
  const originalBaseUrl = process.env.GRAPH_BASE_URL;

  afterEach(() => {
    globalThis.fetch = originalFetch;
    console.log = originalLog;
    if (originalBaseUrl === undefined) delete process.env.GRAPH_BASE_URL;
    else process.env.GRAPH_BASE_URL = originalBaseUrl;
  });

  function mockPatchFetch(): { body: () => Record<string, unknown> } {
    let capturedBody: Record<string, unknown> = {};
    globalThis.fetch = (async (_input: string | URL | Request, init?: RequestInit) => {
      if (init?.method === 'PATCH') {
        capturedBody = JSON.parse(String(init.body ?? '{}'));
      }
      return new Response(JSON.stringify({ id: 'msg-1', subject: 'x' }), {
        status: 200,
        headers: { 'content-type': 'application/json' }
      });
    }) as unknown as typeof fetch;
    return { body: () => capturedBody };
  }

  it('does not touch the body field when only --markdown is set with no --body/--template', async () => {
    process.env.GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';
    const mock = mockPatchFetch();
    console.log = () => {};

    const handled = await tryGraphDraftMutations(
      token,
      undefined,
      { edit: 'msg-1', subject: 'New subject', markdown: true },
      'graph'
    );

    expect(handled).toBe(true);
    expect(mock.body()).not.toHaveProperty('body');
    expect(mock.body().subject).toBe('New subject');
  });

  it('does not touch the body field when only --html is set with no --body/--template', async () => {
    process.env.GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';
    const mock = mockPatchFetch();
    console.log = () => {};

    await tryGraphDraftMutations(token, undefined, { edit: 'msg-1', html: true }, 'graph');

    expect(mock.body()).not.toHaveProperty('body');
  });

  it('does patch the body when --body is explicitly provided (with --markdown)', async () => {
    process.env.GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';
    const mock = mockPatchFetch();
    console.log = () => {};

    await tryGraphDraftMutations(token, undefined, { edit: 'msg-1', body: '**bold**', markdown: true }, 'graph');

    const patchedBody = mock.body().body as { contentType: string; content: string };
    expect(patchedBody.contentType).toBe('HTML');
    expect(patchedBody.content).toContain('<strong>bold</strong>');
  });

  it('patches an explicit empty-string body (--body "") since that is a deliberate clear', async () => {
    process.env.GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';
    const mock = mockPatchFetch();
    console.log = () => {};

    await tryGraphDraftMutations(token, undefined, { edit: 'msg-1', body: '' }, 'graph');

    expect(mock.body().body).toEqual({ contentType: 'Text', content: '' });
  });
});
