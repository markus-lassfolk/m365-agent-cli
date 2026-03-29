import { describe, expect, it } from 'bun:test';

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

import { writeFileSync, unlinkSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { resolve } from 'node:path';
import { uploadLargeFile } from './graph-client.js';

describe('uploadLargeFile chunking', () => {
  it('uploads file in chunks and returns DriveItem', async () => {
    const tmpFile = resolve(tmpdir(), `test-upload-large-${Date.now()}-${Math.random().toString(36).substring(7)}.tmp`);
    const fileSize = 25 * 1024 * 1024; // 25MB
    const buffer = new Uint8Array(fileSize);
    buffer.fill(42);
    writeFileSync(tmpFile, buffer);

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
          if (range?.endsWith('-26214399/26214400')) {
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

      if (!result.ok) throw new Error(JSON.stringify(result));
      expect(result.data?.driveItem?.id).toBe('item-123');
      expect(fetchCalls.length).toBeGreaterThanOrEqual(3);

      const firstCall = fetchCalls[0];
      expect(firstCall.range).toContain('bytes 0-');
      expect(firstCall.range).toContain('/26214400');

      const lastCall = fetchCalls[fetchCalls.length - 1];
      expect(lastCall.range).toContain('-26214399/26214400');
      expect(lastCall.bodySize).toBeGreaterThan(0);
    } finally {
      globalThis.fetch = originalFetch;
      try {
        unlinkSync(tmpFile);
      } catch {}
    }
  });
});
