import { describe, expect, it } from 'bun:test';
import {
  createCalendarEventFileAttachmentUploadSession,
  createMailMessageFileAttachmentUploadSession,
  GRAPH_OUTLOOK_ATTACHMENT_SESSION_THRESHOLD_BYTES,
  uploadBufferViaGraphUploadUrl
} from './graph-attachment-upload-session.js';

describe('graph-attachment-upload-session', () => {
  it('exports threshold constant (3 MB — Graph upload-session minimum)', () => {
    expect(GRAPH_OUTLOOK_ATTACHMENT_SESSION_THRESHOLD_BYTES).toBe(3 * 1024 * 1024);
  });

  it('uploadBufferViaGraphUploadUrl derives attachment id from the Location header', async () => {
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async () =>
        new Response('', {
          status: 201,
          headers: { location: 'https://outlook.office.com/api/v2.0/me/messages/AAA/attachments/ATTACH-99' }
        })) as unknown as typeof fetch;

      const r = await uploadBufferViaGraphUploadUrl('https://u.example/x', Buffer.from([1, 2, 3]));
      expect(r.ok).toBe(true);
      expect(r.data?.id).toBe('ATTACH-99');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('uploadBufferViaGraphUploadUrl rejects zero bytes', async () => {
    const r = await uploadBufferViaGraphUploadUrl('https://upload.example/u', Buffer.alloc(0));
    expect(r.ok).toBe(false);
    expect(r.error?.message).toMatch(/zero-byte/);
  });

  it('uploadBufferViaGraphUploadUrl completes single chunk and parses JSON', async () => {
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async () =>
        new Response(JSON.stringify({ id: 'att-1' }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        })) as unknown as typeof fetch;

      const r = await uploadBufferViaGraphUploadUrl('https://u.example/x', Buffer.from([1, 2, 3]));
      expect(r.ok).toBe(true);
      expect(r.data?.id).toBe('att-1');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('createMailMessageFileAttachmentUploadSession POSTs AttachmentItem', async () => {
    process.env.GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';
    const bodies: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (_input: string | URL | Request, init?: RequestInit) => {
        if (init?.body && typeof init.body === 'string') bodies.push(init.body);
        return new Response(JSON.stringify({ uploadUrl: 'https://put.example/x' }), {
          status: 201,
          headers: { 'content-type': 'application/json' }
        });
      }) as unknown as typeof fetch;

      const r = await createMailMessageFileAttachmentUploadSession('tok', 'msg-1', 'f.bin', 100, 'application/pdf');
      expect(r.ok).toBe(true);
      expect(r.data?.uploadUrl).toContain('put.example');
      const b = JSON.parse(bodies[0] || '{}');
      expect(b.AttachmentItem.name).toBe('f.bin');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('createCalendarEventFileAttachmentUploadSession targets events path', async () => {
    process.env.GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';
    const urls: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        return new Response(JSON.stringify({ uploadUrl: 'https://put.example/y' }), {
          status: 201,
          headers: { 'content-type': 'application/json' }
        });
      }) as unknown as typeof fetch;

      const r = await createCalendarEventFileAttachmentUploadSession(
        'tok',
        'evt-1',
        'a.png',
        50,
        'image/png',
        'u@x.com'
      );
      expect(r.ok).toBe(true);
      expect(urls[0]).toContain('/users/u%40x.com/events/evt-1/attachments/createUploadSession');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});
