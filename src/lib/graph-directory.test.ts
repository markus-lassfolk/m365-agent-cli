import { describe, expect, it } from 'bun:test';

const token = 'tok';
const baseUrl = 'https://graph.microsoft.com/v1.0';

describe('graph-directory', () => {
  it('searchPeople and listPeople with search use ConsistencyLevel', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        const h = init?.headers;
        const ch = h instanceof Headers ? h.get('ConsistencyLevel') : (h as Record<string, string>)?.ConsistencyLevel;
        expect(ch).toBe('eventual');
        return new Response(JSON.stringify({ value: [{ id: 'p1', displayName: 'Pat' }] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as unknown as typeof fetch;
      const d = await import('./graph-directory.js');
      const s = await d.searchPeople(token, 'Pat');
      expect(s.ok).toBe(true);
      const l = await d.listPeople(token, { top: 5, search: 'Pat' });
      expect(l.ok).toBe(true);
      expect(urls.some((u) => u.includes('%24search'))).toBe(true);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('getPerson GETs /me/people/{id}', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async () =>
        new Response(JSON.stringify({ id: 'p1', displayName: 'Pat' }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        })) as unknown as typeof fetch;
      const { getPerson } = await import('./graph-directory.js');
      const r = await getPerson(token, 'p1');
      expect(r.ok).toBe(true);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('searchUsers and searchGroups', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async () =>
        new Response(
          JSON.stringify({
            value: [
              { id: 'u1', displayName: 'Alice' },
              { id: 'g1', displayName: 'Team' }
            ]
          }),
          { status: 200, headers: { 'content-type': 'application/json' } }
        )) as unknown as typeof fetch;
      const d = await import('./graph-directory.js');
      const u = await d.searchUsers(token, 'Ali');
      expect(u.ok).toBe(true);
      const g = await d.searchGroups(token, 'Tea');
      expect(g.ok).toBe(true);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('listPeople without top paginates via fetchAllPages', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    let n = 0;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async () => {
        n += 1;
        if (n === 1) {
          return new Response(
            JSON.stringify({
              value: [{ id: 'p1', displayName: 'A' }],
              '@odata.nextLink': `${baseUrl}/me/people?$skiptoken=x`
            }),
            { status: 200, headers: { 'content-type': 'application/json' } }
          );
        }
        return new Response(JSON.stringify({ value: [{ id: 'p2', displayName: 'B' }] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as unknown as typeof fetch;
      const { listPeople } = await import('./graph-directory.js');
      const r = await listPeople(token, {});
      expect(r.ok).toBe(true);
      expect(r.data?.map((p) => p.id).sort()).toEqual(['p1', 'p2']);
      expect(n).toBe(2);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('expandGroup filters user-like members', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async () =>
        new Response(
          JSON.stringify({
            value: [
              { id: 'u1', displayName: 'U', '@odata.type': '#microsoft.graph.user' },
              { id: 'sp1', displayName: 'Site' }
            ]
          }),
          { status: 200, headers: { 'content-type': 'application/json' } }
        )) as unknown as typeof fetch;
      const { expandGroup } = await import('./graph-directory.js');
      const r = await expandGroup(token, 'g1');
      expect(r.ok).toBe(true);
      expect(r.data?.length).toBe(1);
      expect(r.data?.[0].id).toBe('u1');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('graph-directory expandGroup member classification', () => {
  const token = 'tok';
  const baseUrl = 'https://graph.microsoft.com/v1.0';

  it('includes users but excludes mail-enabled groups / other member types', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async () =>
        new Response(
          JSON.stringify({
            value: [
              { '@odata.type': '#microsoft.graph.user', id: 'u1', mail: 'u1@x.com' },
              // Mail-enabled group: has a `mail` property but must NOT be treated as a user.
              { '@odata.type': '#microsoft.graph.group', id: 'g1', mail: 'dl@x.com', displayName: 'DL' },
              { '@odata.type': '#microsoft.graph.servicePrincipal', id: 'sp1' },
              // Missing discriminator but user-shaped -> included via fallback.
              { id: 'u2', userPrincipalName: 'u2@x.com' }
            ]
          }),
          { status: 200, headers: { 'content-type': 'application/json' } }
        )) as unknown as typeof fetch;
      const d = await import('./graph-directory.js');
      const r = await d.expandGroup(token, 'group-1');
      expect(r.ok).toBe(true);
      expect((r.data ?? []).map((m: any) => m.id).sort()).toEqual(['u1', 'u2']);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});
