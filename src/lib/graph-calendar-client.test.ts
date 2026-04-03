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

  it('sends Prefer outlook.timezone=UTC when preferOutlookTimezoneUtc', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const headers: Headers[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
        headers.push(new Headers(init?.headers as HeadersInit));
        return new Response(JSON.stringify({ id: 'evt-1' }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;

      const { getEvent } = await import('./graph-calendar-client.js');
      await getEvent(token, 'evt-1', undefined, 'id', { preferOutlookTimezoneUtc: true });

      expect(headers[0].get('Prefer')).toBe('outlook.timezone="UTC"');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('normalizeGraphCalendarRangeInstant', () => {
  it('adds Z to local-style instant without zone', async () => {
    const { normalizeGraphCalendarRangeInstant } = await import('./graph-calendar-client.js');
    expect(normalizeGraphCalendarRangeInstant('2026-01-15T10:00:00')).toBe('2026-01-15T10:00:00Z');
    expect(normalizeGraphCalendarRangeInstant('2026-01-15T10:00:00.0000000')).toBe('2026-01-15T10:00:00Z');
  });

  it('passes through values that already have Z or offset', async () => {
    const { normalizeGraphCalendarRangeInstant } = await import('./graph-calendar-client.js');
    expect(normalizeGraphCalendarRangeInstant('2026-01-15T10:00:00Z')).toBe('2026-01-15T10:00:00Z');
    expect(normalizeGraphCalendarRangeInstant('2026-01-15T11:00:00+01:00')).toBe('2026-01-15T11:00:00+01:00');
  });

  it('expands date-only to UTC midnight', async () => {
    const { normalizeGraphCalendarRangeInstant } = await import('./graph-calendar-client.js');
    expect(normalizeGraphCalendarRangeInstant('2026-03-01')).toBe('2026-03-01T00:00:00.000Z');
  });
});

describe('listEventInstances', () => {
  it('normalizes start/end query params and optional Prefer header', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const captured: Array<{ url: string; prefer?: string | null }> = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
        const url = typeof input === 'string' ? input : input.toString();
        const h = new Headers(init?.headers as HeadersInit);
        captured.push({ url, prefer: h.get('Prefer') });
        return new Response(JSON.stringify({ value: [] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;

      const { listEventInstances } = await import('./graph-calendar-client.js');
      const r = await listEventInstances(token, 'master-1', '2026-01-01', '2026-01-02T12:00:00', {
        preferOutlookTimezoneUtc: true
      });

      expect(r.ok).toBe(true);
      expect(captured[0].url).toContain('startDateTime=2026-01-01T00%3A00%3A00.000Z');
      expect(captured[0].url).toContain('endDateTime=2026-01-02T12%3A00%3A00Z');
      expect(captured[0].prefer).toBe('outlook.timezone="UTC"');
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
      globalThis.fetch = (async (_input: string | URL | Request, init?: RequestInit) => {
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

describe('eventsDeltaPage', () => {
  it('GETs /me/events/delta when no calendar id', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        return new Response(JSON.stringify({ value: [{ id: 'e1', subject: 'X' }] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;

      const { eventsDeltaPage } = await import('./graph-calendar-client.js');
      const r = await eventsDeltaPage(token, {});

      expect(r.ok).toBe(true);
      expect(r.data?.value?.[0]?.id).toBe('e1');
      expect(urls[0]).toContain('/me/events/delta');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('GETs /calendars/{id}/events/delta when calendar set', async () => {
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

      const { eventsDeltaPage } = await import('./graph-calendar-client.js');
      const r = await eventsDeltaPage(token, { calendarId: 'cal-99' });

      expect(r.ok).toBe(true);
      expect(urls[0]).toContain('/me/calendars/cal-99/events/delta');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});
