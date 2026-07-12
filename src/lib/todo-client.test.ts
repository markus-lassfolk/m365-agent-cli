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
    const inits: RequestInit[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        inits.push(init ?? {});
        return new Response(JSON.stringify({ value: [] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as unknown as typeof fetch;

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
      const h = new Headers(inits[0]?.headers);
      expect(h.get('ConsistencyLevel')).toBe('eventual');
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
      }) as unknown as typeof fetch;

      const { getTask } = await import('./todo-client.js');
      const r = await getTask(token, 'list-1', 'task-1', undefined, { select: 'id,title' });

      expect(r.ok).toBe(true);
      expect(decodeURIComponent(urls[0])).toContain('$select=');
      expect(decodeURIComponent(urls[0])).toContain('id,title');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('appends $expand when provided (bug regression: was silently dropped)', async () => {
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
      }) as unknown as typeof fetch;

      const { getTask } = await import('./todo-client.js');
      const r = await getTask(token, 'list-1', 'task-1', undefined, { expand: 'attachments' });

      expect(r.ok).toBe(true);
      expect(decodeURIComponent(urls[0])).toContain('$expand=attachments');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('appends both $select and $expand together', async () => {
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
      }) as unknown as typeof fetch;

      const { getTask } = await import('./todo-client.js');
      await getTask(token, 'list-1', 'task-1', undefined, { select: 'id,title', expand: 'attachments' });

      const decoded = decodeURIComponent(urls[0]);
      expect(decoded).toContain('$select=id,title');
      expect(decoded).toContain('$expand=attachments');
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
      }) as unknown as typeof fetch;

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
      }) as unknown as typeof fetch;

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
      }) as unknown as typeof fetch;

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
      }) as unknown as typeof fetch;

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
      }) as unknown as typeof fetch;

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
      }) as unknown as typeof fetch;

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

describe('createTask', () => {
  it('POSTs task with dueDateTime and linkedResources', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const bodies: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        if (init?.body && typeof init.body === 'string') bodies.push(init.body);
        return new Response(JSON.stringify({ id: 'new-task', title: 'Buy milk' }), {
          status: 201,
          headers: { 'content-type': 'application/json' }
        });
      }) as unknown as typeof fetch;

      const { createTask } = await import('./todo-client.js');
      const r = await createTask(token, 'list-99', {
        title: 'Buy milk',
        dueDateTime: '2026-05-10T12:00:00',
        timeZone: 'UTC',
        linkedResources: [{ displayName: 'Issue', webUrl: 'https://example.com/1' }]
      });

      expect(r.ok).toBe(true);
      expect(r.data?.id).toBe('new-task');
      expect(urls[0]).toContain('/me/todo/lists/list-99/tasks');
      const body = JSON.parse(bodies[0] || '{}');
      expect(body.title).toBe('Buy milk');
      expect(body.dueDateTime).toEqual({ dateTime: '2026-05-10T12:00:00', timeZone: 'UTC' });
      expect(body.linkedResources[0].webUrl).toBe('https://example.com/1');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('getTodoListsDeltaPage', () => {
  it('GETs /me/todo/lists/delta() when no continuation URL', async () => {
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
      }) as unknown as typeof fetch;

      const { getTodoListsDeltaPage } = await import('./todo-client.js');
      const r = await getTodoListsDeltaPage(token, undefined, undefined);
      expect(r.ok).toBe(true);
      expect(urls[0]).toContain('/me/todo/lists/delta()');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('attachment sessions and todo root', () => {
  it('listTaskAttachmentSessions GETs …/tasks/{id}/attachmentSessions', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        return new Response(JSON.stringify({ value: [{ id: 'sess-1' }] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as unknown as typeof fetch;

      const { listTaskAttachmentSessions } = await import('./todo-client.js');
      const r = await listTaskAttachmentSessions(token, 'list-1', 'task-1');
      expect(r.ok).toBe(true);
      expect(r.data?.[0]?.id).toBe('sess-1');
      expect(urls[0]).toContain('/me/todo/lists/list-1/tasks/task-1/attachmentSessions');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('getTodoNavigationResource GETs /me/todo', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        return new Response(JSON.stringify({ id: 'todo-root' }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as unknown as typeof fetch;

      const { getTodoNavigationResource } = await import('./todo-client.js');
      const r = await getTodoNavigationResource(token);
      expect(r.ok).toBe(true);
      expect((r.data as { id?: string }).id).toBe('todo-root');
      expect(urls[0]).toMatch(/\/me\/todo(?:\?|$)/);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('todo lists CRUD and paging', () => {
  it('createTodoList POSTs display name', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async () =>
        new Response(JSON.stringify({ id: 'nl', displayName: 'New', wellknownListName: 'none' }), {
          status: 201,
          headers: { 'content-type': 'application/json' }
        })) as unknown as typeof fetch;
      const { createTodoList } = await import('./todo-client.js');
      const r = await createTodoList(token, 'New');
      expect(r.ok).toBe(true);
      expect(r.data?.id).toBe('nl');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('updateTodoList PATCHes', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async () =>
        new Response(JSON.stringify({ id: 'l1', displayName: 'Ren' }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        })) as unknown as typeof fetch;
      const { updateTodoList } = await import('./todo-client.js');
      const r = await updateTodoList(token, 'l1', 'Ren');
      expect(r.ok).toBe(true);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('deleteTodoList sends DELETE', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async () => new Response(null, { status: 204 })) as unknown as typeof fetch;
      const { deleteTodoList } = await import('./todo-client.js');
      const r = await deleteTodoList(token, 'l-del');
      expect(r.ok).toBe(true);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('getTodoListsPage returns value and nextLink', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async () =>
        new Response(
          JSON.stringify({
            value: [{ id: 'l1', displayName: 'A', wellknownListName: 'none' }],
            '@odata.nextLink': `${baseUrl}/me/todo/lists?$skip=1`
          }),
          { status: 200, headers: { 'content-type': 'application/json' } }
        )) as unknown as typeof fetch;
      const { getTodoListsPage } = await import('./todo-client.js');
      const r = await getTodoListsPage(token, undefined, { top: 5, count: true });
      expect(r.ok).toBe(true);
      expect(r.data?.value?.length).toBe(1);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('todo tasks update/delete and checklist', () => {
  it('updateTask PATCHes title', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async () =>
        new Response(JSON.stringify({ id: 't1', title: 'Up', status: 'notStarted' }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        })) as unknown as typeof fetch;
      const { updateTask } = await import('./todo-client.js');
      const r = await updateTask(token, 'l1', 't1', { title: 'Up' });
      expect(r.ok).toBe(true);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('deleteTask sends DELETE', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async () => new Response(null, { status: 204 })) as unknown as typeof fetch;
      const { deleteTask } = await import('./todo-client.js');
      const r = await deleteTask(token, 'l1', 't1');
      expect(r.ok).toBe(true);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('addChecklistItem POSTs displayName', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async () =>
        new Response(
          JSON.stringify({
            id: 'ck-new',
            displayName: 'Item',
            isChecked: false,
            createdDateTime: '2026-01-01T00:00:00Z'
          }),
          { status: 201, headers: { 'content-type': 'application/json' } }
        )) as unknown as typeof fetch;
      const { addChecklistItem } = await import('./todo-client.js');
      const r = await addChecklistItem(token, 'l1', 't1', 'Item');
      expect(r.ok).toBe(true);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('updateChecklistItem and deleteChecklistItem', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    try {
      let n = 0;
      globalThis.fetch = (async () => {
        n += 1;
        if (n === 1) {
          return new Response(
            JSON.stringify({
              id: 'ck1',
              displayName: 'X',
              isChecked: true,
              createdDateTime: '2026-01-01T00:00:00Z'
            }),
            { status: 200, headers: { 'content-type': 'application/json' } }
          );
        }
        return new Response(null, { status: 204 });
      }) as unknown as typeof fetch;
      const { updateChecklistItem, deleteChecklistItem } = await import('./todo-client.js');
      const u = await updateChecklistItem(token, 'l1', 't1', 'ck1', { isChecked: true });
      expect(u.ok).toBe(true);
      const d = await deleteChecklistItem(token, 'l1', 't1', 'ck1');
      expect(d.ok).toBe(true);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('listTaskChecklistItems pages', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async () =>
        new Response(
          JSON.stringify({ value: [{ id: 'c1', displayName: 'A', isChecked: false, createdDateTime: 't' }] }),
          {
            status: 200,
            headers: { 'content-type': 'application/json' }
          }
        )) as unknown as typeof fetch;
      const { listTaskChecklistItems } = await import('./todo-client.js');
      const r = await listTaskChecklistItems(token, 'l1', 't1');
      expect(r.ok).toBe(true);
      expect(r.data?.length).toBe(1);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('todo attachments and linked resources', () => {
  it('listAttachments, getTaskAttachment, create/delete file and reference attachments', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    try {
      let call = 0;
      globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
        const u = typeof input === 'string' ? input : input.toString();
        const m = (init?.method || 'GET').toUpperCase();
        call += 1;
        if (m === 'DELETE') return new Response(null, { status: 204 });
        if (u.includes('/attachments') && m === 'POST') {
          return new Response(JSON.stringify({ id: `att-${call}`, name: 'f' }), {
            status: 201,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (u.includes('/attachments/att-g') && m === 'GET') {
          return new Response(JSON.stringify({ id: 'att-g', name: 'g' }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        return new Response(JSON.stringify({ value: [{ id: 'a1' }] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as unknown as typeof fetch;
      const t = await import('./todo-client.js');
      const list = await t.listAttachments(token, 'l1', 't1');
      expect(list.ok).toBe(true);
      const g = await t.getTaskAttachment(token, 'l1', 't1', 'att-g');
      expect(g.ok).toBe(true);
      const f = await t.createTaskFileAttachment(token, 'l1', 't1', 'f.bin', 'YQ==', 'application/octet-stream');
      expect(f.ok).toBe(true);
      const r = await t.createTaskReferenceAttachment(token, 'l1', 't1', 'link', 'https://example.com');
      expect(r.ok).toBe(true);
      const d = await t.deleteAttachment(token, 'l1', 't1', 'a1');
      expect(d.ok).toBe(true);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('linkedResources collection CRUD and list', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
        const u = typeof input === 'string' ? input : input.toString();
        const m = (init?.method || 'GET').toUpperCase();
        if (m === 'DELETE') return new Response(null, { status: 204 });
        if (m === 'POST' && u.includes('/linkedResources')) {
          return new Response(JSON.stringify({ id: 'lr1', webUrl: 'https://n' }), {
            status: 201,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (m === 'PATCH' && u.includes('/linkedResources/lr1')) {
          return new Response(JSON.stringify({ id: 'lr1', webUrl: 'https://n2' }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (u.includes('/linkedResources/lr1') && m === 'GET') {
          return new Response(JSON.stringify({ id: 'lr1', webUrl: 'https://n' }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        return new Response(JSON.stringify({ value: [{ id: 'lr1' }] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as unknown as typeof fetch;
      const t = await import('./todo-client.js');
      const li = await t.listTaskLinkedResources(token, 'l1', 't1');
      expect(li.ok).toBe(true);
      const c = await t.createTaskLinkedResource(token, 'l1', 't1', { displayName: 'D', webUrl: 'https://n' });
      expect(c.ok).toBe(true);
      const g = await t.getTaskLinkedResource(token, 'l1', 't1', 'lr1');
      expect(g.ok).toBe(true);
      const u = await t.updateTaskLinkedResource(token, 'l1', 't1', 'lr1', { webUrl: 'https://n2' });
      expect(u.ok).toBe(true);
      const d = await t.deleteTaskLinkedResource(token, 'l1', 't1', 'lr1');
      expect(d.ok).toBe(true);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('addLinkedResource merges into task PATCH', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    try {
      let _n = 0;
      globalThis.fetch = (async (_input: string | URL | Request, init?: RequestInit) => {
        _n += 1;
        const m = (init?.method || 'GET').toUpperCase();
        if (m === 'GET') {
          return new Response(
            JSON.stringify({
              id: 't1',
              title: 'T',
              status: 'notStarted',
              linkedResources: []
            }),
            { status: 200, headers: { 'content-type': 'application/json' } }
          );
        }
        return new Response(JSON.stringify({ id: 't1', title: 'T', status: 'notStarted' }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as unknown as typeof fetch;
      const { addLinkedResource } = await import('./todo-client.js');
      const r = await addLinkedResource(token, 'l1', 't1', { displayName: 'L', webUrl: 'https://x.com' });
      expect(r.ok).toBe(true);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('removeLinkedResourceByWebUrl filters linkedResources', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async (_input: string | URL | Request, init?: RequestInit) => {
        const m = (init?.method || 'GET').toUpperCase();
        if (m === 'GET') {
          return new Response(
            JSON.stringify({
              id: 't1',
              title: 'T',
              status: 'notStarted',
              linkedResources: [{ webUrl: 'https://drop.me' }, { webUrl: 'https://keep.com' }]
            }),
            {
              status: 200,
              headers: { 'content-type': 'application/json' }
            }
          );
        }
        return new Response(JSON.stringify({ id: 't1', title: 'T', status: 'notStarted' }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as unknown as typeof fetch;
      const { removeLinkedResourceByWebUrl } = await import('./todo-client.js');
      const r = await removeLinkedResourceByWebUrl(token, 'l1', 't1', 'https://drop.me');
      expect(r.ok).toBe(true);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('todo attachment sessions and navigation', () => {
  it('getTaskAttachmentSession, patch, delete, content GET/PUT/DELETE', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
        const u = typeof input === 'string' ? input : input.toString();
        const m = (init?.method || 'GET').toUpperCase();
        const isSessionContent =
          u.includes('/attachmentSessions/') && (u.includes('/content') || u.endsWith('/content'));
        if (m === 'DELETE' && isSessionContent) return new Response(null, { status: 204 });
        if (m === 'DELETE') return new Response(null, { status: 204 });
        if (m === 'PUT' && isSessionContent) {
          return new Response(JSON.stringify({ id: 's1', expirationDateTime: 't' }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (isSessionContent && m === 'GET') {
          return new Response(new Uint8Array([1]), { status: 200 });
        }
        if (m === 'PATCH') {
          return new Response(JSON.stringify({ id: 's1', expirationDateTime: 't' }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        return new Response(JSON.stringify({ id: 's1', uploadUrl: 'https://u' }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as unknown as typeof fetch;
      const t = await import('./todo-client.js');
      const g = await t.getTaskAttachmentSession(token, 'l1', 't1', 's1');
      expect(g.ok).toBe(true);
      const p = await t.patchTaskAttachmentSession(token, 'l1', 't1', 's1', {});
      expect(p.ok).toBe(true);
      const gc = await t.getTaskAttachmentSessionContent(token, 'l1', 't1', 's1');
      expect(gc.ok).toBe(true);
      const pc = await t.putTaskAttachmentSessionContent(token, 'l1', 't1', 's1', new Uint8Array([9]));
      expect(pc.ok).toBe(true);
      const dc = await t.deleteTaskAttachmentSessionContent(token, 'l1', 't1', 's1');
      expect(dc.ok).toBe(true);
      const d = await t.deleteTaskAttachmentSession(token, 'l1', 't1', 's1');
      expect(d.ok).toBe(true);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('patchTodoNavigationResource and deleteTodoNavigationResource', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    try {
      let _n = 0;
      globalThis.fetch = (async (_input: string | URL | Request, init?: RequestInit) => {
        _n += 1;
        const m = (init?.method || 'GET').toUpperCase();
        if (m === 'DELETE') return new Response(null, { status: 204 });
        return new Response(JSON.stringify({ id: 'todo' }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as unknown as typeof fetch;
      const t = await import('./todo-client.js');
      const p = await t.patchTodoNavigationResource(token, { isDefaultListVisible: true });
      expect(p.ok).toBe(true);
      const d = await t.deleteTodoNavigationResource(token, undefined, 'W/"1"');
      expect(d.ok).toBe(true);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('todo open extensions and delta URLs', () => {
  it('list/get/set/update/delete todo list open extensions', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
        const u = typeof input === 'string' ? input : input.toString();
        const m = (init?.method || 'GET').toUpperCase();
        if (m === 'DELETE' || (m === 'PATCH' && u.includes('/extensions/ext1'))) {
          return new Response(null, { status: 204 });
        }
        if (m === 'POST' && u.includes('/extensions')) {
          return new Response(JSON.stringify({ extensionName: 'ext1', foo: 1 }), {
            status: 201,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (u.includes('/extensions/ext1') && m === 'GET') {
          return new Response(JSON.stringify({ extensionName: 'ext1' }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        return new Response(JSON.stringify({ value: [] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as unknown as typeof fetch;
      const t = await import('./todo-client.js');
      const l = await t.listTodoListOpenExtensions(token, 'l1');
      expect(l.ok).toBe(true);
      const g = await t.getTodoListOpenExtension(token, 'l1', 'ext1');
      expect(g.ok).toBe(true);
      const s = await t.setTodoListOpenExtension(token, 'l1', 'ext1', { foo: 1 });
      expect(s.ok).toBe(true);
      const u = await t.updateTodoListOpenExtension(token, 'l1', 'ext1', { foo: 2 });
      expect(u.ok).toBe(true);
      const d = await t.deleteTodoListOpenExtension(token, 'l1', 'ext1');
      expect(d.ok).toBe(true);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('task open extensions list/get/set/update/delete', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
        const u = typeof input === 'string' ? input : input.toString();
        const m = (init?.method || 'GET').toUpperCase();
        if (m === 'DELETE' || (m === 'PATCH' && u.includes('/extensions/te1'))) {
          return new Response(null, { status: 204 });
        }
        if (m === 'POST' && u.includes('/tasks/t1/extensions')) {
          return new Response(JSON.stringify({ extensionName: 'te1' }), {
            status: 201,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (u.includes('/extensions/te1') && m === 'GET') {
          return new Response(JSON.stringify({ extensionName: 'te1' }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        return new Response(JSON.stringify({ value: [] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as unknown as typeof fetch;
      const t = await import('./todo-client.js');
      const l = await t.listTaskOpenExtensions(token, 'l1', 't1');
      expect(l.ok).toBe(true);
      const g = await t.getTaskOpenExtension(token, 'l1', 't1', 'te1');
      expect(g.ok).toBe(true);
      const s = await t.setTaskOpenExtension(token, 'l1', 't1', 'te1', { x: 1 });
      expect(s.ok).toBe(true);
      const u = await t.updateTaskOpenExtension(token, 'l1', 't1', 'te1', { x: 2 });
      expect(u.ok).toBe(true);
      const d = await t.deleteTaskOpenExtension(token, 'l1', 't1', 'te1');
      expect(d.ok).toBe(true);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('getTodoTasksDeltaPage with absolute URL', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async () =>
        new Response(JSON.stringify({ value: [] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        })) as unknown as typeof fetch;
      const { getTodoTasksDeltaPage } = await import('./todo-client.js');
      const r = await getTodoTasksDeltaPage(token, 'l1', `${baseUrl}/me/todo/lists/l1/tasks/delta?token=x`);
      expect(r.ok).toBe(true);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('getTodoListsDeltaPage with fullUrl', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async () =>
        new Response(JSON.stringify({ value: [] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        })) as unknown as typeof fetch;
      const { getTodoListsDeltaPage } = await import('./todo-client.js');
      const r = await getTodoListsDeltaPage(token, `${baseUrl}/me/todo/lists/delta()?token=y`);
      expect(r.ok).toBe(true);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});
