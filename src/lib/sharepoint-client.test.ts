import { describe, expect, it } from 'bun:test';

const token = 'tok';
const baseUrl = 'https://graph.microsoft.com/v1.0';

describe('sharepoint-client', () => {
  it('getSiteByGraphPath rejects empty sitePath', async () => {
    const { getSiteByGraphPath } = await import('./sharepoint-client.js');
    const r = await getSiteByGraphPath(token, '   ');
    expect(r.ok).toBe(false);
    expect(r.error?.message).toMatch(/sitePath is required/);
  });

  it('getSiteByGraphPath GETs /sites/{encodedPath}', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        return new Response(JSON.stringify({ id: 's1', displayName: 'Team' }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as unknown as typeof fetch;

      const { getSiteByGraphPath } = await import('./sharepoint-client.js');
      const r = await getSiteByGraphPath(token, 'contoso.sharepoint.com:/sites/T1');
      expect(r.ok).toBe(true);
      expect(r.data?.id).toBe('s1');
      expect(urls[0]).toContain('/sites/contoso.sharepoint.com:/sites/T1');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('getLists returns value array', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async () =>
        new Response(
          JSON.stringify({
            value: [
              {
                id: 'l1',
                name: 'L',
                displayName: 'List',
                createdDateTime: 't',
                lastModifiedDateTime: 't',
                webUrl: 'https://x'
              }
            ]
          }),
          { status: 200, headers: { 'content-type': 'application/json' } }
        )) as unknown as typeof fetch;

      const { getLists } = await import('./sharepoint-client.js');
      const r = await getLists(token, 'site-guid');
      expect(r.ok).toBe(true);
      expect(r.data?.[0]?.name).toBe('L');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('getSiteDefaultDriveId GETs site drive', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        return new Response(JSON.stringify({ id: 'drive-1' }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as unknown as typeof fetch;

      const { getSiteDefaultDriveId } = await import('./sharepoint-client.js');
      const r = await getSiteDefaultDriveId(token, 'site-1');
      expect(r.ok).toBe(true);
      expect(r.data?.id).toBe('drive-1');
      expect(urls[0]).toContain('/sites/site-1/drive');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('getLists treats a response without value as an empty list (paginated)', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async () =>
        new Response(JSON.stringify({}), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        })) as unknown as typeof fetch;
      const { getLists } = await import('./sharepoint-client.js');
      const r = await getLists(token, 'site-guid');
      // Now paginated via fetchAllPages: a missing `value` is an empty page, consistent with siblings.
      expect(r.ok).toBe(true);
      expect(r.data).toEqual([]);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('getLists follows @odata.nextLink across pages', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    try {
      let n = 0;
      globalThis.fetch = (async () => {
        n++;
        if (n === 1) {
          return new Response(
            JSON.stringify({
              value: [{ id: '1', name: 'L1' }],
              '@odata.nextLink': `${baseUrl}/sites/s/lists?$skiptoken=2`
            }),
            { status: 200, headers: { 'content-type': 'application/json' } }
          );
        }
        return new Response(JSON.stringify({ value: [{ id: '2', name: 'L2' }] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as unknown as typeof fetch;
      const { getLists } = await import('./sharepoint-client.js');
      const r = await getLists(token, 'site-guid');
      expect(r.ok).toBe(true);
      expect(r.data?.map((l) => l.id)).toEqual(['1', '2']);
      expect(n).toBe(2);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('getSiteById GETs /sites/{id}', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        return new Response(JSON.stringify({ id: 'sid', displayName: 'S' }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as unknown as typeof fetch;
      const { getSiteById } = await import('./sharepoint-client.js');
      const r = await getSiteById(token, 'contoso.sharepoint.com,x,y');
      expect(r.ok).toBe(true);
      expect(r.data?.id).toBe('sid');
      expect(urls[0]).toContain('/sites/contoso.sharepoint.com%2Cx%2Cy');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('getSiteDrives follows @odata.nextLink', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    let n = 0;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async () => {
        n += 1;
        if (n === 1) {
          return new Response(
            JSON.stringify({
              value: [{ id: 'd1' }],
              '@odata.nextLink': `${baseUrl}/sites/s1/drives?$skip=1`
            }),
            { status: 200, headers: { 'content-type': 'application/json' } }
          );
        }
        return new Response(JSON.stringify({ value: [{ id: 'd2' }] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as unknown as typeof fetch;
      const { getSiteDrives } = await import('./sharepoint-client.js');
      const r = await getSiteDrives(token, 's1');
      expect(r.ok).toBe(true);
      expect(r.data?.map((d) => d.id).sort()).toEqual(['d1', 'd2']);
      expect(n).toBe(2);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('getListMetadata returns list', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async () =>
        new Response(
          JSON.stringify({
            id: 'L1',
            name: 'n',
            displayName: 'D',
            createdDateTime: 'a',
            lastModifiedDateTime: 'b',
            webUrl: 'https://w'
          }),
          { status: 200, headers: { 'content-type': 'application/json' } }
        )) as unknown as typeof fetch;
      const { getListMetadata } = await import('./sharepoint-client.js');
      const r = await getListMetadata(token, 'site', 'list');
      expect(r.ok).toBe(true);
      expect(r.data?.id).toBe('L1');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('getListItemsPage builds query and supports nextLink', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        return new Response(JSON.stringify({ value: [{ id: 'i1' }] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as unknown as typeof fetch;
      const { getListItemsPage } = await import('./sharepoint-client.js');
      const a = await getListItemsPage(token, 's', 'l', {
        filter: "fields/Title eq 'x'",
        orderby: 'createdDateTime desc',
        top: 5
      });
      expect(a.ok).toBe(true);
      expect(urls[0]).toMatch(/expand=fields/);
      expect(urls[0]).toMatch(/filter=/);
      expect(urls[0]).toMatch(/orderby=/);
      expect(urls[0]).toMatch(/top=5/);

      const next = `${baseUrl}/sites/s/lists/l/items?$skip=token`;
      const b = await getListItemsPage(token, 's', 'l', { nextLink: next });
      expect(b.ok).toBe(true);
      expect(urls[1]).toBe(next);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('getAllListItemsPages aggregates pages', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    let n = 0;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async () => {
        n += 1;
        if (n === 1) {
          return new Response(
            JSON.stringify({
              value: [{ id: 'a', createdDateTime: 't', lastModifiedDateTime: 't', webUrl: 'w', fields: {} }],
              '@odata.nextLink': `${baseUrl}/next`
            }),
            { status: 200, headers: { 'content-type': 'application/json' } }
          );
        }
        return new Response(
          JSON.stringify({
            value: [{ id: 'b', createdDateTime: 't', lastModifiedDateTime: 't', webUrl: 'w', fields: {} }]
          }),
          { status: 200, headers: { 'content-type': 'application/json' } }
        );
      }) as unknown as typeof fetch;
      const { getAllListItemsPages } = await import('./sharepoint-client.js');
      const r = await getAllListItemsPages(token, 's', 'l', { top: 10 });
      expect(r.ok).toBe(true);
      expect(r.data?.map((x) => x.id).sort()).toEqual(['a', 'b']);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('createListItem POSTs fields body', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    let body = '';
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async (_input: string | URL | Request, init?: RequestInit) => {
        body = String(init?.body ?? '');
        return new Response(
          JSON.stringify({
            id: 'new',
            createdDateTime: 't',
            lastModifiedDateTime: 't',
            webUrl: 'w',
            fields: { Title: 'x' }
          }),
          { status: 201, headers: { 'content-type': 'application/json' } }
        );
      }) as unknown as typeof fetch;
      const { createListItem } = await import('./sharepoint-client.js');
      const r = await createListItem(token, 's', 'l', { Title: 'x' });
      expect(r.ok).toBe(true);
      expect(JSON.parse(body)).toEqual({ fields: { Title: 'x' } });
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('updateListItem PATCHes fields', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async () =>
        new Response(JSON.stringify({ Title: 'y' }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        })) as unknown as typeof fetch;
      const { updateListItem } = await import('./sharepoint-client.js');
      const r = await updateListItem(token, 's', 'l', 'item-1', { Title: 'y' });
      expect(r.ok).toBe(true);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('getListItem GETs item with expand', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        return new Response(
          JSON.stringify({
            id: 'item-1',
            createdDateTime: 't',
            lastModifiedDateTime: 't',
            webUrl: 'w',
            fields: {}
          }),
          { status: 200, headers: { 'content-type': 'application/json' } }
        );
      }) as unknown as typeof fetch;
      const { getListItem } = await import('./sharepoint-client.js');
      const r = await getListItem(token, 's', 'l', 'item-1');
      expect(r.ok).toBe(true);
      expect(urls[0]).toContain('$expand=fields');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('deleteListItem sends DELETE', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    let method = '';
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async (_input: string | URL | Request, init?: RequestInit) => {
        method = init?.method ?? '';
        return new Response(null, { status: 204 });
      }) as unknown as typeof fetch;
      const { deleteListItem } = await import('./sharepoint-client.js');
      const r = await deleteListItem(token, 's', 'l', 'item-1');
      expect(r.ok).toBe(true);
      expect(method).toBe('DELETE');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('getListItemsDeltaPage uses absolute link when provided', async () => {
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
      }) as unknown as typeof fetch;
      const { getListItemsDeltaPage } = await import('./sharepoint-client.js');
      const link = `${baseUrl}/sites/s/lists/l/items/delta?token=1`;
      const r = await getListItemsDeltaPage(token, 's', 'l', link);
      expect(r.ok).toBe(true);
      expect(urls[0]).toBe(link);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('getSitePermissions lists via fetchAllPages', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async () =>
        new Response(JSON.stringify({ value: [{ id: 'p1' }] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        })) as unknown as typeof fetch;
      const { getSitePermissions } = await import('./sharepoint-client.js');
      const r = await getSitePermissions(token, 's1');
      expect(r.ok).toBe(true);
      expect(r.data?.[0]).toEqual({ id: 'p1' });
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('getListColumns and getListItems use fetchAllPages', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async () =>
        new Response(JSON.stringify({ value: [{ name: 'Title', id: '1' }] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        })) as unknown as typeof fetch;
      const { getListColumns, getListItems } = await import('./sharepoint-client.js');
      const c = await getListColumns(token, 's', 'l');
      expect(c.ok).toBe(true);
      expect(c.data?.[0]?.name).toBe('Title');
      const items = await getListItems(token, 's', 'l');
      expect(items.ok).toBe(true);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('getListItemsDeltaPage starts delta when link omitted', async () => {
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
      }) as unknown as typeof fetch;
      const { getListItemsDeltaPage } = await import('./sharepoint-client.js');
      const r = await getListItemsDeltaPage(token, 's', 'l');
      expect(r.ok).toBe(true);
      expect(urls[0]).toContain('/items/delta');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('updateSitePermission PATCHes body', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async () =>
        new Response(JSON.stringify({ id: 'p1', roles: ['owner'] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        })) as unknown as typeof fetch;
      const { updateSitePermission } = await import('./sharepoint-client.js');
      const r = await updateSitePermission(token, 's1', 'p1', { roles: ['owner'] });
      expect(r.ok).toBe(true);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('getSitePermission GETs /sites/{id}/permissions/{id}', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        expect(init?.method).toBe('GET');
        return new Response(JSON.stringify({ id: 'p1', roles: ['write'] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as unknown as typeof fetch;
      const { getSitePermission } = await import('./sharepoint-client.js');
      const r = await getSitePermission(token, 's1', 'p1');
      expect(r.ok).toBe(true);
      expect(r.data?.id).toBe('p1');
      expect(urls[0]).toContain('/sites/s1/permissions/p1');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('createSitePermission POSTs the permission body', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    let body = '';
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        body = String(init?.body ?? '');
        expect(init?.method).toBe('POST');
        return new Response(JSON.stringify({ id: 'new-p', roles: ['write'] }), {
          status: 201,
          headers: { 'content-type': 'application/json' }
        });
      }) as unknown as typeof fetch;
      const { createSitePermission } = await import('./sharepoint-client.js');
      const r = await createSitePermission(token, 's1', {
        roles: ['write'],
        grantedToIdentities: [{ application: { id: 'app-1', displayName: 'App' } }]
      });
      expect(r.ok).toBe(true);
      expect(r.data?.id).toBe('new-p');
      expect(urls[0]).toContain('/sites/s1/permissions');
      expect(JSON.parse(body)).toEqual({
        roles: ['write'],
        grantedToIdentities: [{ application: { id: 'app-1', displayName: 'App' } }]
      });
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('deleteSitePermission sends DELETE', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    let method = '';
    const urls: string[] = [];
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        method = init?.method ?? '';
        return new Response(null, { status: 204 });
      }) as unknown as typeof fetch;
      const { deleteSitePermission } = await import('./sharepoint-client.js');
      const r = await deleteSitePermission(token, 's1', 'p1');
      expect(r.ok).toBe(true);
      expect(method).toBe('DELETE');
      expect(urls[0]).toContain('/sites/s1/permissions/p1');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});
