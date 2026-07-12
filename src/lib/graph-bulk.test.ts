import { afterEach, describe, expect, it } from 'bun:test';
import {
  applyBulkGraphRequests,
  type BulkSubRequestSpec,
  parseBulkIdListOrExit,
  printBulkOutcomeSummary
} from './graph-bulk.js';

const token = 'tok';

describe('applyBulkGraphRequests', () => {
  const originalFetch = globalThis.fetch;
  const originalBaseUrl = process.env.GRAPH_BASE_URL;

  afterEach(() => {
    globalThis.fetch = originalFetch;
    if (originalBaseUrl === undefined) delete process.env.GRAPH_BASE_URL;
    else process.env.GRAPH_BASE_URL = originalBaseUrl;
  });

  it('maps successful and failed sub-responses back to the request id', async () => {
    process.env.GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';
    globalThis.fetch = (async () =>
      new Response(
        JSON.stringify({
          responses: [
            { id: 'a', status: 200, body: { ok: true } },
            { id: 'b', status: 404, body: { error: { message: 'not found' } } }
          ]
        }),
        { status: 200, headers: { 'content-type': 'application/json' } }
      )) as unknown as typeof fetch;

    const requests: BulkSubRequestSpec[] = [
      { id: 'a', method: 'PATCH', url: '/me/todo/lists/l/tasks/a', body: { status: 'completed' } },
      { id: 'b', method: 'PATCH', url: '/me/todo/lists/l/tasks/b', body: { status: 'completed' } }
    ];
    const r = await applyBulkGraphRequests(token, requests);
    expect(r.ok).toBe(true);
    expect(r.data).toEqual([
      { id: 'a', ok: true, status: 200 },
      { id: 'b', ok: false, status: 404, error: 'not found' }
    ]);
  });

  it('falls back to "HTTP {status}" when the failed sub-response has no error.message', async () => {
    process.env.GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';
    globalThis.fetch = (async () =>
      new Response(JSON.stringify({ responses: [{ id: 'a', status: 500, body: {} }] }), {
        status: 200,
        headers: { 'content-type': 'application/json' }
      })) as unknown as typeof fetch;

    const r = await applyBulkGraphRequests(token, [{ id: 'a', method: 'DELETE', url: '/me/todo/lists/l/tasks/a' }]);
    expect(r.ok).toBe(true);
    expect(r.data?.[0]).toEqual({ id: 'a', ok: false, status: 500, error: 'HTTP 500' });
  });

  it('marks a request missing from the batch response as failed rather than crashing', async () => {
    process.env.GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';
    globalThis.fetch = (async () =>
      new Response(JSON.stringify({ responses: [] }), {
        status: 200,
        headers: { 'content-type': 'application/json' }
      })) as unknown as typeof fetch;

    const r = await applyBulkGraphRequests(token, [{ id: 'a', method: 'DELETE', url: '/x' }]);
    expect(r.ok).toBe(true);
    expect(r.data?.[0].ok).toBe(false);
    expect(r.data?.[0].error).toMatch(/No response/);
  });

  it('propagates a top-level batch failure (e.g. auth) as a whole', async () => {
    globalThis.fetch = (async () => new Response('nope', { status: 401 })) as unknown as typeof fetch;
    const r = await applyBulkGraphRequests(token, [{ id: 'a', method: 'DELETE', url: '/x' }]);
    expect(r.ok).toBe(false);
  });
});

describe('parseBulkIdListOrExit', () => {
  const originalExit = process.exit;
  const originalError = console.error;

  afterEach(() => {
    process.exit = originalExit;
    console.error = originalError;
  });

  it('parses a comma-separated --ids list, trimming and dropping blanks', async () => {
    expect(await parseBulkIdListOrExit({ ids: ' a, b ,,c' })).toEqual(['a', 'b', 'c']);
  });

  it('exits with an error when neither --ids nor --json-file is provided', async () => {
    let exitCode: number | undefined;
    process.exit = ((code?: number) => {
      exitCode = code;
      throw new Error('exit');
    }) as never;
    console.error = () => {};
    await expect(parseBulkIdListOrExit({})).rejects.toThrow('exit');
    expect(exitCode).toBe(1);
  });

  it('exits with an error when the id list is empty after trimming', async () => {
    let exitCode: number | undefined;
    process.exit = ((code?: number) => {
      exitCode = code;
      throw new Error('exit');
    }) as never;
    console.error = () => {};
    await expect(parseBulkIdListOrExit({ ids: ' , ,' })).rejects.toThrow('exit');
    expect(exitCode).toBe(1);
  });
});

describe('printBulkOutcomeSummary', () => {
  const originalLog = console.log;
  afterEach(() => {
    console.log = originalLog;
  });

  it('prints a JSON summary with succeeded/failed counts when json is true', () => {
    const logged: string[] = [];
    console.log = ((s: string) => logged.push(s)) as typeof console.log;
    printBulkOutcomeSummary(
      [
        { id: 'a', ok: true, status: 200 },
        { id: 'b', ok: false, status: 404, error: 'nope' }
      ],
      true
    );
    const parsed = JSON.parse(logged[0]);
    expect(parsed).toEqual({
      succeeded: 1,
      failed: 1,
      results: [
        { id: 'a', ok: true, status: 200 },
        { id: 'b', ok: false, status: 404, error: 'nope' }
      ]
    });
  });

  it('prints a per-id human summary line plus a totals line when json is falsy', () => {
    const logged: string[] = [];
    console.log = ((s: string) => logged.push(s)) as typeof console.log;
    printBulkOutcomeSummary(
      [
        { id: 'a', ok: true, status: 200 },
        { id: 'b', ok: false, status: 404, error: 'nope' }
      ],
      false
    );
    expect(logged[0]).toBe('✓ a');
    expect(logged[1]).toBe('✗ b: nope');
    expect(logged[2]).toContain('1 succeeded, 1 failed (2 total)');
  });
});
