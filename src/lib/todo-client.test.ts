import { describe, expect, it } from 'bun:test';

const token = 'test-token';
const baseUrl = 'https://graph.microsoft.com/v1.0';

describe('linkedResourceToGraphPayload', () => {
  it('prefers displayName over description', async () => {
    const { linkedResourceToGraphPayload } = await import('./todo-client.js');
    expect(linkedResourceToGraphPayload({ displayName: 'A', description: 'B', webUrl: 'https://example.com' })).toEqual(
      {
        displayName: 'A',
        webUrl: 'https://example.com'
      }
    );
  });

  it('uses description when displayName absent', async () => {
    const { linkedResourceToGraphPayload } = await import('./todo-client.js');
    expect(linkedResourceToGraphPayload({ description: 'Only', webUrl: 'https://x' })).toEqual({
      displayName: 'Only',
      webUrl: 'https://x'
    });
  });

  it('includes optional applicationName, externalId, id', async () => {
    const { linkedResourceToGraphPayload } = await import('./todo-client.js');
    expect(
      linkedResourceToGraphPayload({
        displayName: 'T',
        applicationName: 'App',
        externalId: 'ext-1',
        id: 'lr-id'
      })
    ).toEqual({
      displayName: 'T',
      applicationName: 'App',
      externalId: 'ext-1',
      id: 'lr-id'
    });
  });

  it('omits empty displayName string', async () => {
    const { linkedResourceToGraphPayload } = await import('./todo-client.js');
    expect(linkedResourceToGraphPayload({ description: '', webUrl: 'https://u' })).toEqual({
      webUrl: 'https://u'
    });
  });
});

describe('getTasks query options', () => {
  it('requests single page with $top, $skip, $expand, $count', async () => {
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

      const { getTasks } = await import('./todo-client.js');
      const r = await getTasks(token, 'list-1', {
        filter: "status eq 'notStarted'",
        top: 10,
        skip: 5,
        expand: 'attachments',
        count: true
      });

      expect(r.ok).toBe(true);
      expect(r.data).toEqual([]);
      const u = decodeURIComponent(urls[0]);
      expect(u).toContain('$top=10');
      expect(u).toContain('$skip=5');
      expect(u).toContain('$expand=attachments');
      expect(u).toContain('$count=true');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('getTask $select', () => {
  it('appends $select when provided', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        return new Response(JSON.stringify({ id: 't1', title: 'x' }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;

      const { getTask } = await import('./todo-client.js');
      const r = await getTask(token, 'list-1', 'task-1', undefined, { select: 'id,title' });

      expect(r.ok).toBe(true);
      expect(urls[0]).toContain('$select=');
      expect(decodeURIComponent(urls[0])).toContain('id,title');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('getTodoLists / getTodoList', () => {
  it('getTodoLists returns lists from value', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        return new Response(
          JSON.stringify({
            value: [{ id: 'l1', displayName: 'Work', wellknownListName: 'none' }]
          }),
          { status: 200, headers: { 'content-type': 'application/json' } }
        );
      }) as typeof fetch;

      const { getTodoLists } = await import('./todo-client.js');
      const r = await getTodoLists(token);

      expect(r.ok).toBe(true);
      expect(r.data?.[0]?.id).toBe('l1');
      expect(urls[0]).toContain('/me/todo/lists');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('getTodoList returns one list', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (_input: string | URL | Request) => {
        return new Response(JSON.stringify({ id: 'l2', displayName: 'Home', wellknownListName: 'none' }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;

      const { getTodoList } = await import('./todo-client.js');
      const r = await getTodoList(token, 'l2');

      expect(r.ok).toBe(true);
      expect(r.data?.displayName).toBe('Home');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('getTasks string filter (paging)', () => {
  it('uses fetchAllPages for string $filter only', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        return new Response(JSON.stringify({ value: [{ id: 't1', title: 'a', status: 'completed' }] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;

      const { getTasks } = await import('./todo-client.js');
      const r = await getTasks(token, 'list-1', "status eq 'completed'");

      expect(r.ok).toBe(true);
      expect(r.data).toHaveLength(1);
      const q = decodeURIComponent(urls[0].replace(/\+/g, ' '));
      expect(q).toContain("$filter=status eq 'completed'");
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('createTask', () => {
  it('POSTs payload with optional fields', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const posts: { url: string; body: string }[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
        const url = typeof input === 'string' ? input : input.toString();
        if (init?.method === 'POST' && init.body) {
          posts.push({ url, body: String(init.body) });
        }
        return new Response(JSON.stringify({ id: 'new-id', title: 'Created', status: 'notStarted' }), {
          status: 201,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;

      const { createTask } = await import('./todo-client.js');
      const r = await createTask(token, 'list-x', {
        title: 'Created',
        body: '<p>x</p>',
        bodyContentType: 'html',
        dueDateTime: '2026-06-01T12:00:00',
        startDateTime: '2026-06-01T09:00:00',
        timeZone: 'Europe/Stockholm',
        importance: 'high',
        status: 'notStarted',
        isReminderOn: true,
        reminderDateTime: '2026-06-01T08:00:00',
        linkedResources: [{ displayName: 'Link', webUrl: 'https://example.com' }],
        categories: ['Cat1'],
        recurrence: { pattern: { type: 'daily', interval: 1 } }
      });

      expect(r.ok).toBe(true);
      expect(r.data?.id).toBe('new-id');
      expect(posts).toHaveLength(1);
      expect(posts[0].url).toContain('/lists/list-x/tasks');
      const payload = JSON.parse(posts[0].body) as Record<string, unknown>;
      expect(payload.title).toBe('Created');
      expect(payload.body).toEqual({ content: '<p>x</p>', contentType: 'html' });
      expect(payload.importance).toBe('high');
      expect(payload.categories).toEqual(['Cat1']);
      expect(Array.isArray(payload.linkedResources)).toBe(true);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('getChecklistItem', () => {
  it('GETs one checklist item by id', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        return new Response(
          JSON.stringify({
            id: 'ck1',
            displayName: 'Buy milk',
            isChecked: false,
            createdDateTime: '2026-01-01T12:00:00.000Z'
          }),
          { status: 200, headers: { 'content-type': 'application/json' } }
        );
      }) as typeof fetch;

      const { getChecklistItem } = await import('./todo-client.js');
      const r = await getChecklistItem(token, 'list-1', 'task-1', 'ck1');

      expect(r.ok).toBe(true);
      expect(r.data?.displayName).toBe('Buy milk');
      expect(urls[0]).toContain('/checklistItems/ck1');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('getTaskAttachmentContent', () => {
  it('GETs raw bytes from attachments/$value', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        const u = typeof input === 'string' ? input : input.toString();
        expect(u).toContain('$value');
        return new Response(new Uint8Array([7, 8, 9]), { status: 200 });
      }) as typeof fetch;

      const { getTaskAttachmentContent } = await import('./todo-client.js');
      const r = await getTaskAttachmentContent(token, 'list-1', 'task-1', 'att-1');

      expect(r.ok).toBe(true);
      expect(r.data?.length).toBe(3);
      expect(r.data?.[0]).toBe(7);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});
