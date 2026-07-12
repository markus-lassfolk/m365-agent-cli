import { describe, expect, it } from 'bun:test';

const token = 'tok';
const baseUrl = 'https://graph.microsoft.com/v1.0';

describe('chat message reactions', () => {
  it('setChatMessageReaction posts to the root message when no replyId is given', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    const urls: string[] = [];
    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        return new Response(null, { status: 204 });
      }) as unknown as typeof fetch;

      const { setChatMessageReaction } = await import('./graph-teams-client.js');
      const r = await setChatMessageReaction(token, 'chat-1', 'msg-1', '👍');
      expect(r.ok).toBe(true);
      expect(urls[0]).toBe(`${baseUrl}/chats/chat-1/messages/msg-1/setReaction`);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('unsetChatMessageReaction posts to the root message when no replyId is given', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    const urls: string[] = [];
    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        return new Response(null, { status: 204 });
      }) as unknown as typeof fetch;

      const { unsetChatMessageReaction } = await import('./graph-teams-client.js');
      const r = await unsetChatMessageReaction(token, 'chat-1', 'msg-1', '👍');
      expect(r.ok).toBe(true);
      expect(urls[0]).toBe(`${baseUrl}/chats/chat-1/messages/msg-1/unsetReaction`);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('URL-encodes chatId/messageId', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    const urls: string[] = [];
    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        return new Response(null, { status: 204 });
      }) as unknown as typeof fetch;

      const { setChatMessageReaction } = await import('./graph-teams-client.js');
      await setChatMessageReaction(token, '19:abc@thread.v2', 'msg id', '👍');
      expect(urls[0]).toBe(`${baseUrl}/chats/19%3Aabc%40thread.v2/messages/msg%20id/setReaction`);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('setChatMessageReaction rejects --reply (replyId) without calling fetch (bug regression: no reply-reaction endpoint exists for chats)', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    let fetchCalled = false;
    try {
      globalThis.fetch = (async () => {
        fetchCalled = true;
        return new Response(null, { status: 204 });
      }) as unknown as typeof fetch;

      const { setChatMessageReaction } = await import('./graph-teams-client.js');
      const r = await setChatMessageReaction(token, 'chat-1', 'msg-1', '👍', 'reply-1');
      expect(r.ok).toBe(false);
      expect(r.error?.message).toMatch(/not supported/i);
      expect(fetchCalled).toBe(false);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('unsetChatMessageReaction rejects --reply (replyId) without calling fetch', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    let fetchCalled = false;
    try {
      globalThis.fetch = (async () => {
        fetchCalled = true;
        return new Response(null, { status: 204 });
      }) as unknown as typeof fetch;

      const { unsetChatMessageReaction } = await import('./graph-teams-client.js');
      const r = await unsetChatMessageReaction(token, 'chat-1', 'msg-1', '👍', 'reply-1');
      expect(r.ok).toBe(false);
      expect(r.error?.message).toMatch(/not supported/i);
      expect(fetchCalled).toBe(false);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});
