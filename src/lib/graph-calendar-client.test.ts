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
