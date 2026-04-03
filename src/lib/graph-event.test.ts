import { describe, expect, it } from 'bun:test';

const token = 'test-token';
const baseUrl = 'https://graph.microsoft.com/v1.0';

describe('acceptEventInvitation', () => {
  it('POSTs /events/{id}/accept with sendResponse and optional comment', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const requests: Array<{ url: string; method?: string; body?: string }> = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
        const url = typeof input === 'string' ? input : input.toString();
        requests.push({
          url,
          method: init?.method,
          body: typeof init?.body === 'string' ? init.body : undefined
        });
        return new Response(null, { status: 204 });
      }) as typeof fetch;

      const { acceptEventInvitation } = await import('./graph-event.js');
      const r = await acceptEventInvitation({ token, eventId: 'inv-1', comment: 'On my way' });

      expect(r.ok).toBe(true);
      expect(requests[0].method).toBe('POST');
      expect(requests[0].url).toContain('/me/events/inv-1/accept');
      expect(JSON.parse(requests[0].body ?? '{}')).toEqual({
        sendResponse: true,
        comment: 'On my way'
      });
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('declineEventInvitation', () => {
  it('POSTs /events/{id}/decline', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    let url = '';
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        url = typeof input === 'string' ? input : input.toString();
        return new Response(null, { status: 204 });
      }) as typeof fetch;

      const { declineEventInvitation } = await import('./graph-event.js');
      const r = await declineEventInvitation({ token, eventId: 'inv-2' });
      expect(r.ok).toBe(true);
      expect(url).toContain('/me/events/inv-2/decline');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('tentativelyAcceptEventInvitation', () => {
  it('POSTs /events/{id}/tentativelyAccept', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    let url = '';
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        url = typeof input === 'string' ? input : input.toString();
        return new Response(null, { status: 204 });
      }) as typeof fetch;

      const { tentativelyAcceptEventInvitation } = await import('./graph-event.js');
      const r = await tentativelyAcceptEventInvitation({ token, eventId: 'inv-3' });
      expect(r.ok).toBe(true);
      expect(url).toContain('/me/events/inv-3/tentativelyAccept');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});
