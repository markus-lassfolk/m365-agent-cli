import { describe, expect, it } from 'bun:test';

const token = 'test-token';
const baseUrl = 'https://graph.microsoft.com/v1.0';

describe('listMailFolders', () => {
  it('GETs /mailFolders collection', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        return new Response(
          JSON.stringify({
            value: [{ id: 'inbox-id', displayName: 'Inbox' }]
          }),
          { status: 200, headers: { 'content-type': 'application/json' } }
        );
      }) as typeof fetch;

      const { listMailFolders } = await import('./outlook-graph-client.js');
      const r = await listMailFolders(token);

      expect(r.ok).toBe(true);
      expect(r.data?.[0]?.displayName).toBe('Inbox');
      expect(urls[0]).toContain('/me/mailFolders');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('getMessage', () => {
  it('GETs /messages/{id}', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        return new Response(JSON.stringify({ id: 'msg-1', subject: 'Hi', isRead: false }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;

      const { getMessage } = await import('./outlook-graph-client.js');
      const r = await getMessage(token, 'msg-1', undefined, 'subject,isRead');

      expect(r.ok).toBe(true);
      expect(r.data?.subject).toBe('Hi');
      expect(urls[0]).toContain('/me/messages/msg-1');
      expect(urls[0]).toContain('$select=');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('listMailboxMessages', () => {
  it('GETs /me/messages', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        return new Response(JSON.stringify({ value: [{ id: 'm1', subject: 'A' }] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;

      const { listMailboxMessages } = await import('./outlook-graph-client.js');
      const r = await listMailboxMessages(token, undefined, { top: 10 });

      expect(r.ok).toBe(true);
      expect(urls[0]).toContain('/me/messages');
      expect(decodeURIComponent(urls[0])).toContain('$top=10');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('adds ConsistencyLevel when using search', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    let consistency: string | undefined;
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (_input: string | URL | Request, init?: RequestInit) => {
        const h = init?.headers;
        if (h instanceof Headers) {
          consistency = h.get('ConsistencyLevel') ?? undefined;
        } else if (h && typeof h === 'object') {
          consistency = (h as Record<string, string>).ConsistencyLevel;
        }
        return new Response(JSON.stringify({ value: [] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;

      const { listMailboxMessages } = await import('./outlook-graph-client.js');
      const r = await listMailboxMessages(token, undefined, { top: 5, search: 'budget' });

      expect(r.ok).toBe(true);
      expect(consistency).toBe('eventual');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('sendMail', () => {
  it('POSTs /sendMail', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const bodies: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        if (init?.body && typeof init.body === 'string') bodies.push(init.body);
        return new Response(null, { status: 202 });
      }) as typeof fetch;

      const { sendMail } = await import('./outlook-graph-client.js');
      const r = await sendMail(token, {
        message: { subject: 'Hi', body: { contentType: 'Text', content: 'x' } },
        saveToSentItems: true
      });

      expect(r.ok).toBe(true);
      expect(urls[0]).toContain('/me/sendMail');
      expect(bodies[0]).toContain('saveToSentItems');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('createDraftMessage', () => {
  it('POSTs /me/messages with isDraft', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const bodies: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        if (init?.body && typeof init.body === 'string') bodies.push(init.body);
        return new Response(JSON.stringify({ id: 'draft-1', isDraft: true }), {
          status: 201,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;

      const { createDraftMessage } = await import('./outlook-graph-client.js');
      const r = await createDraftMessage(token, {
        subject: 'S',
        bodyContent: 'hello',
        bodyContentType: 'Text',
        toAddresses: ['a@b.com']
      });

      expect(r.ok).toBe(true);
      expect(r.data?.id).toBe('draft-1');
      expect(urls[0]).toContain('/me/messages');
      expect(bodies[0]).toContain('"isDraft":true');
      expect(bodies[0]).toContain('a@b.com');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('mailMessagesDeltaPage', () => {
  it('GETs /me/messages/delta when no folder', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        return new Response(JSON.stringify({ value: [{ id: 'm1' }] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;

      const { mailMessagesDeltaPage } = await import('./outlook-graph-client.js');
      const r = await mailMessagesDeltaPage(token, {});

      expect(r.ok).toBe(true);
      expect(r.data?.value?.[0]?.id).toBe('m1');
      expect(urls[0]).toContain('/me/messages/delta');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('GETs .../mailFolders/{id}/messages/delta when folder set', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        return new Response(JSON.stringify({ value: [] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;

      const { mailMessagesDeltaPage } = await import('./outlook-graph-client.js');
      const r = await mailMessagesDeltaPage(token, { folderId: 'inbox' });

      expect(r.ok).toBe(true);
      expect(urls[0]).toContain('/me/mailFolders/inbox/messages/delta');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});
