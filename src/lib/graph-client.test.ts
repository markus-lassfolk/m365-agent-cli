import { afterEach, describe, expect, it } from 'bun:test';

const token = 'test-token';
const baseUrl = 'https://graph.microsoft.com/v1.0';

describe('searchFiles query encoding', () => {
  it('encodes single quotes in search query before interpolation', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;

    const fetchCalls: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        fetchCalls.push(typeof input === 'string' ? input : input.toString());
        return new Response(JSON.stringify({ value: [] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;

      const { searchFiles } = await import('./graph-client.js');
      await searchFiles(token, "') and name='anything");

      expect(fetchCalls).toHaveLength(1);
      expect(fetchCalls[0]).toContain("/me/drive/root/search(q='%27%29%20and%20name%3D%27anything')");
      expect(fetchCalls[0]).not.toContain("q=') and name='anything'");
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

import { unlink, writeFile } from 'node:fs/promises';
import { uploadLargeFile } from './graph-client.js';

describe('uploadLargeFile chunking', () => {
  const tmpFile = 'test-upload-large.tmp';

  afterEach(async () => {
    try {
      await unlink(tmpFile);
    } catch {}
  });

  it('uploads file in chunks and returns DriveItem', async () => {
    const fileSize = 25 * 1024 * 1024; // 25MB
    const buffer = new Uint8Array(fileSize);
    buffer.fill(42);
    await writeFile(tmpFile, buffer);

    const originalFetch = globalThis.fetch;
    const fetchCalls: any[] = [];

    try {
      globalThis.fetch = (async (input: any, init?: any) => {
        const url = typeof input === 'string' ? input : input.toString();

        // 1. Create session POST
        if (url.includes('createUploadSession')) {
          return new Response(
            JSON.stringify({
              uploadUrl: 'https://upload.example.com/session-123',
              expirationDateTime: '2026-04-01T00:00:00.000Z'
            }),
            { status: 200, headers: { 'content-type': 'application/json' } }
          );
        }

        // 2. Chunk PUTs
        if (init?.method === 'PUT') {
          fetchCalls.push({
            url,
            range: (init.headers as any)?.['Content-Range'],
            bodySize: (init.body as any)?.length
          });
          const range = (init.headers as any)?.['Content-Range'];
          if (range.startsWith('bytes 20971520-26214399')) {
            // Last chunk 10MB*2 to 25MB
            return new Response(JSON.stringify({ id: 'item-123', name: 'test.tmp' }), {
              status: 201,
              headers: { 'content-type': 'application/json' }
            });
          }
          return new Response('{"expirationDateTime": "..."}', { status: 202 });
        }

        return new Response('{}', { status: 200 });
      }) as any;

      const result = await uploadLargeFile('token', tmpFile);

      expect(result.ok).toBe(true);
      expect(result.data?.driveItem?.id).toBe('item-123');
      expect(fetchCalls.length).toBe(3);
      expect(fetchCalls[0].range).toBe('bytes 0-10485759/26214400');
      expect(fetchCalls[1].range).toBe('bytes 10485760-20971519/26214400');
      expect(fetchCalls[2].range).toBe('bytes 20971520-26214399/26214400');
      expect(fetchCalls[2].bodySize).toBe(5 * 1024 * 1024);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});
