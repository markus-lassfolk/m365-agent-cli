import { describe, expect, it } from 'bun:test';
import { getMailboxSettings, setMailboxSettings } from './oof-client.js';

describe('oof-client', () => {
  const token = 'test-token';

  it('getMailboxSettings handles GET', async () => {
    const fetchCalls: any[] = [];
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async (input, init) => {
        fetchCalls.push({ input, init });
        return new Response(JSON.stringify({ automaticRepliesSetting: { status: 'disabled' } }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;

      const res = await getMailboxSettings(token);
      expect(res.ok).toBe(true);
      expect(res.data?.automaticRepliesSetting?.status).toBe('disabled');
      expect(fetchCalls).toHaveLength(1);
      expect(fetchCalls[0].input.toString()).toContain('/me/mailboxSettings');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('setMailboxSettings handles PATCH with scheduled time', async () => {
    const fetchCalls: any[] = [];
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async (input, init) => {
        fetchCalls.push({ input, init });
        return new Response('', { status: 204 });
      }) as typeof fetch;

      const res = await setMailboxSettings(token, {
        status: 'scheduled',
        internalReplyMessage: 'Away',
        scheduledStartDateTime: '2025-01-01T00:00:00.000Z'
      });
      
      expect(res.ok).toBe(true);
      expect(fetchCalls).toHaveLength(1);
      expect(fetchCalls[0].init.method).toBe('PATCH');
      
      const body = JSON.parse(fetchCalls[0].init.body);
      expect(body.automaticRepliesSetting.status).toBe('scheduled');
      expect(body.automaticRepliesSetting.internalReplyMessage).toBe('Away');
      expect(body.automaticRepliesSetting.scheduledStartDateTime).toEqual({
        dateTime: '2025-01-01T00:00:00.000Z',
        timeZone: 'UTC'
      });
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});
