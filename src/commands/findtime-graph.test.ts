import { afterEach, describe, expect, test } from 'bun:test';
import { mergeAvailabilityViewsToMergedFree, runFindTimeGraph, runFindTimeGraphSchedule } from './findtime-graph.js';

describe('mergeAvailabilityViewsToMergedFree', () => {
  test('all free when all mailboxes show 0', () => {
    expect(mergeAvailabilityViewsToMergedFree(['00', '00'])).toEqual([true, true]);
  });

  test('busy if any mailbox is non-zero at that slot', () => {
    expect(mergeAvailabilityViewsToMergedFree(['00', '20'])).toEqual([false, true]);
  });

  test('pads shorter views with busy for trailing slots', () => {
    const m = mergeAvailabilityViewsToMergedFree(['0', '00']);
    expect(m).toEqual([true, false]);
  });
});

describe('runFindTimeGraph', () => {
  const originalFetch = globalThis.fetch;
  const originalLog = console.log;
  const originalBaseUrl = process.env.GRAPH_BASE_URL;

  afterEach(() => {
    globalThis.fetch = originalFetch;
    console.log = originalLog;
    if (originalBaseUrl === undefined) delete process.env.GRAPH_BASE_URL;
    else process.env.GRAPH_BASE_URL = originalBaseUrl;
  });

  function mockFindMeetingTimesFetch(): { body: () => Record<string, unknown>; header: () => string | null } {
    let capturedBody: Record<string, unknown> = {};
    let capturedHeader: string | null = null;
    globalThis.fetch = (async (_input: string | URL | Request, init?: RequestInit) => {
      capturedBody = JSON.parse(String(init?.body ?? '{}'));
      capturedHeader = new Headers(init?.headers).get('Prefer');
      return new Response(
        JSON.stringify({
          meetingTimeSuggestions: [
            {
              meetingTimeSlot: { start: { dateTime: '2026-06-15T14:00:00', timeZone: 'UTC' } },
              confidence: 90,
              attendeeAvailability: [
                { attendee: { emailAddress: { address: 'alice@x.com' } }, availability: 'free' },
                { attendee: { emailAddress: { address: 'bob@x.com' } }, availability: 'busy' }
              ]
            }
          ]
        }),
        { status: 200, headers: { 'content-type': 'application/json' } }
      );
    }) as unknown as typeof fetch;
    return { body: () => capturedBody, header: () => capturedHeader };
  }

  test('marks attendees listed in optionalEmails as optional, others stay required', async () => {
    process.env.GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';
    const mock = mockFindMeetingTimesFetch();
    console.log = () => {};

    await runFindTimeGraph({
      token: 'tok',
      emails: ['alice@x.com', 'bob@x.com'],
      start: new Date('2026-06-15T00:00:00Z'),
      end: new Date('2026-06-16T00:00:00Z'),
      durationMinutes: 30,
      workStartHour: 0,
      workEndHour: 24,
      label: 'test',
      json: true,
      optionalEmails: ['bob@x.com']
    });

    const attendees = mock.body().attendees as Array<{ type: string; emailAddress: { address: string } }>;
    expect(attendees).toEqual([
      { type: 'required', emailAddress: { address: 'alice@x.com' } },
      { type: 'optional', emailAddress: { address: 'bob@x.com' } }
    ]);
  });

  test('passes minAttendeePercentage through to the request body', async () => {
    process.env.GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';
    const mock = mockFindMeetingTimesFetch();
    console.log = () => {};

    await runFindTimeGraph({
      token: 'tok',
      emails: ['alice@x.com'],
      start: new Date('2026-06-15T00:00:00Z'),
      end: new Date('2026-06-16T00:00:00Z'),
      durationMinutes: 30,
      workStartHour: 0,
      workEndHour: 24,
      label: 'test',
      json: true,
      minAttendeePercentage: 60
    });

    expect(mock.body().minimumAttendeePercentage).toBe(60);
  });

  test('defaults minAttendeePercentage to 100 when unset', async () => {
    process.env.GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';
    const mock = mockFindMeetingTimesFetch();
    console.log = () => {};

    await runFindTimeGraph({
      token: 'tok',
      emails: ['alice@x.com'],
      start: new Date('2026-06-15T00:00:00Z'),
      end: new Date('2026-06-16T00:00:00Z'),
      durationMinutes: 30,
      workStartHour: 0,
      workEndHour: 24,
      label: 'test',
      json: true
    });

    expect(mock.body().minimumAttendeePercentage).toBe(100);
  });

  test('applies --timezone to the request time zone and Prefer header', async () => {
    process.env.GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';
    const mock = mockFindMeetingTimesFetch();
    console.log = () => {};

    await runFindTimeGraph({
      token: 'tok',
      emails: ['alice@x.com'],
      start: new Date('2026-06-15T12:00:00Z'),
      end: new Date('2026-06-16T00:00:00Z'),
      durationMinutes: 30,
      workStartHour: 0,
      workEndHour: 24,
      label: 'test',
      json: true,
      timezone: 'America/New_York'
    });

    const tc = mock.body().timeConstraint as { timeSlots: Array<{ start: { dateTime: string; timeZone: string } }> };
    expect(tc.timeSlots[0].start.timeZone).toBe('America/New_York');
    expect(tc.timeSlots[0].start.dateTime).toBe('2026-06-15T08:00:00'); // EDT = UTC-4
    expect(mock.header()).toBe('outlook.timezone="America/New_York"');
  });

  test('surfaces attendeeAvailability in the JSON output', async () => {
    process.env.GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';
    mockFindMeetingTimesFetch();
    const logged: string[] = [];
    console.log = ((s: string) => logged.push(s)) as typeof console.log;

    await runFindTimeGraph({
      token: 'tok',
      emails: ['alice@x.com', 'bob@x.com'],
      start: new Date('2026-06-15T00:00:00Z'),
      end: new Date('2026-06-16T00:00:00Z'),
      durationMinutes: 30,
      workStartHour: 0,
      workEndHour: 24,
      label: 'test',
      json: true
    });

    const parsed = JSON.parse(logged[0]);
    expect(parsed.suggestions[0].attendeeAvailability).toEqual([
      { email: 'alice@x.com', availability: 'free' },
      { email: 'bob@x.com', availability: 'busy' }
    ]);
  });
});

describe('runFindTimeGraphSchedule', () => {
  const originalFetch = globalThis.fetch;
  const originalLog = console.log;
  const originalBaseUrl = process.env.GRAPH_BASE_URL;

  afterEach(() => {
    globalThis.fetch = originalFetch;
    console.log = originalLog;
    if (originalBaseUrl === undefined) delete process.env.GRAPH_BASE_URL;
    else process.env.GRAPH_BASE_URL = originalBaseUrl;
  });

  test('includes per-attendee availabilityView detail in the JSON output', async () => {
    process.env.GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';
    globalThis.fetch = (async () =>
      new Response(
        JSON.stringify({
          value: [
            { scheduleId: 'alice@x.com', availabilityView: '000' },
            { scheduleId: 'bob@x.com', availabilityView: '020' }
          ]
        }),
        { status: 200, headers: { 'content-type': 'application/json' } }
      )) as unknown as typeof fetch;
    const logged: string[] = [];
    console.log = ((s: string) => logged.push(s)) as typeof console.log;

    await runFindTimeGraphSchedule({
      token: 'tok',
      emails: ['alice@x.com', 'bob@x.com'],
      start: new Date('2026-06-15T00:00:00Z'),
      end: new Date('2026-06-15T01:30:00Z'),
      durationMinutes: 30,
      workStartHour: 0,
      workEndHour: 24,
      label: 'test',
      json: true
    });

    const parsed = JSON.parse(logged[0]);
    expect(parsed.attendeeAvailability).toEqual([
      { email: 'alice@x.com', availabilityView: '000' },
      { email: 'bob@x.com', availabilityView: '020' }
    ]);
  });

  test('applies --timezone to the request time zone and Prefer header', async () => {
    process.env.GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';
    const captured: { body: Record<string, unknown>; header: string | null } = { body: {}, header: null };
    globalThis.fetch = (async (_input: string | URL | Request, init?: RequestInit) => {
      captured.body = JSON.parse(String(init?.body ?? '{}'));
      captured.header = new Headers(init?.headers).get('Prefer');
      return new Response(JSON.stringify({ value: [{ scheduleId: 'alice@x.com', availabilityView: '00' }] }), {
        status: 200,
        headers: { 'content-type': 'application/json' }
      });
    }) as unknown as typeof fetch;
    console.log = () => {};

    await runFindTimeGraphSchedule({
      token: 'tok',
      emails: ['alice@x.com'],
      start: new Date('2026-06-15T12:00:00Z'),
      end: new Date('2026-06-15T13:00:00Z'),
      durationMinutes: 30,
      workStartHour: 0,
      workEndHour: 24,
      label: 'test',
      json: true,
      timezone: 'Asia/Tokyo'
    });

    expect((captured.body.startTime as { timeZone: string }).timeZone).toBe('Asia/Tokyo');
    expect((captured.body.startTime as { dateTime: string }).dateTime).toBe('2026-06-15T21:00:00');
    expect(captured.header).toBe('outlook.timezone="Asia/Tokyo"');
  });
});
