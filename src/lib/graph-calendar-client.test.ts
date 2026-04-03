import { describe, expect, it } from 'bun:test';

const token = 'test-token';
const baseUrl = 'https://graph.microsoft.com/v1.0';

describe('listCalendars', () => {
  it('GETs /calendars collection', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        return new Response(
          JSON.stringify({
            value: [{ id: 'cal-1', name: 'Calendar' }]
          }),
          { status: 200, headers: { 'content-type': 'application/json' } }
        );
      }) as typeof fetch;

      const { listCalendars } = await import('./graph-calendar-client.js');
      const r = await listCalendars(token);

      expect(r.ok).toBe(true);
      expect(r.data?.[0]?.name).toBe('Calendar');
      expect(urls[0]).toContain('/me/calendars');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('listCalendarView', () => {
  it('GETs default /calendar/calendarView with start and end', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        return new Response(
          JSON.stringify({
            value: [{ id: 'evt-1', subject: 'Standup' }]
          }),
          { status: 200, headers: { 'content-type': 'application/json' } }
        );
      }) as typeof fetch;

      const { listCalendarView } = await import('./graph-calendar-client.js');
      const r = await listCalendarView(token, '2026-04-01T00:00:00Z', '2026-04-02T00:00:00Z', {});

      expect(r.ok).toBe(true);
      expect(r.data?.[0]?.subject).toBe('Standup');
      expect(urls[0]).toContain('/me/calendar/calendarView');
      expect(urls[0]).toContain('startDateTime=');
      expect(urls[0]).toContain('endDateTime=');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('GETs /calendars/{id}/calendarView when calendarId set', async () => {
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

      const { listCalendarView } = await import('./graph-calendar-client.js');
      const r = await listCalendarView(token, '2026-04-01T00:00:00Z', '2026-04-02T00:00:00Z', {
        calendarId: 'abc/def'
      });

      expect(r.ok).toBe(true);
      expect(urls[0]).toContain('/me/calendars/abc%2Fdef/calendarView');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('getEvent', () => {
  it('GETs /events/{id}', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        return new Response(JSON.stringify({ id: 'evt-1', subject: 'Hi' }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;

      const { getEvent } = await import('./graph-calendar-client.js');
      const r = await getEvent(token, 'evt-1', undefined, 'subject,id');

      expect(r.ok).toBe(true);
      expect(r.data?.subject).toBe('Hi');
      expect(urls[0]).toContain('/me/events/evt-1');
      expect(urls[0]).toContain('$select=');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('updateCalendarEvent', () => {
  it('PATCHes /events/{id} with JSON body and returns updated event', async () => {
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
        return new Response(JSON.stringify({ id: 'evt-1', subject: 'Patched' }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;

      const { updateCalendarEvent } = await import('./graph-calendar-client.js');
      const r = await updateCalendarEvent(token, 'evt-1', { subject: 'Patched' }, undefined);

      expect(r.ok).toBe(true);
      expect(r.data?.subject).toBe('Patched');
      expect(requests[0].method).toBe('PATCH');
      expect(requests[0].url).toContain('/me/events/evt-1');
      expect(requests[0].body).toBe(JSON.stringify({ subject: 'Patched' }));
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('deleteCalendarEvent', () => {
  it('DELETEs /events/{id} and succeeds on 204', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const requests: Array<{ url: string; method?: string }> = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
        const url = typeof input === 'string' ? input : input.toString();
        requests.push({ url, method: init?.method });
        return new Response(null, { status: 204 });
      }) as typeof fetch;

      const { deleteCalendarEvent } = await import('./graph-calendar-client.js');
      const r = await deleteCalendarEvent(token, 'evt-del', undefined);

      expect(r.ok).toBe(true);
      expect(requests[0].method).toBe('DELETE');
      expect(requests[0].url).toContain('/me/events/evt-del');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('cancelCalendarEvent', () => {
  it('POSTs /events/{id}/cancel with lowercase comment field', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const requests: Array<{ url: string; body?: string }> = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
        const url = typeof input === 'string' ? input : input.toString();
        requests.push({
          url,
          body: typeof init?.body === 'string' ? init.body : undefined
        });
        return new Response(null, { status: 204 });
      }) as typeof fetch;

      const { cancelCalendarEvent } = await import('./graph-calendar-client.js');
      const r = await cancelCalendarEvent(token, 'evt-can', { comment: 'Sorry' });

      expect(r.ok).toBe(true);
      expect(requests[0].url).toContain('/me/events/evt-can/cancel');
      expect(requests[0].body).toBe(JSON.stringify({ comment: 'Sorry' }));
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('sends empty comment when omitted', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    let body = '';
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
        body = typeof init?.body === 'string' ? init.body : '';
        return new Response(null, { status: 204 });
      }) as typeof fetch;

      const { cancelCalendarEvent } = await import('./graph-calendar-client.js');
      const r = await cancelCalendarEvent(token, 'evt-x', {});
      expect(r.ok).toBe(true);
      expect(body).toBe(JSON.stringify({ comment: '' }));
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('addFileAttachmentToCalendarEvent', () => {
  it('POSTs to /events/{id}/attachments', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        return new Response(JSON.stringify({ id: 'att-1', name: 'f.txt' }), {
          status: 201,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;

      const { addFileAttachmentToCalendarEvent } = await import('./graph-calendar-client.js');
      const r = await addFileAttachmentToCalendarEvent(token, 'evt-1', {
        name: 'f.txt',
        contentType: 'text/plain',
        contentBytes: 'aGk='
      });

      expect(r.ok).toBe(true);
      expect(urls[0]).toContain('/me/events/evt-1/attachments');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});
