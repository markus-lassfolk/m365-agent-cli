import { describe, expect, it } from 'bun:test';
import {
  addExcelTableRows,
  addExcelWorksheet,
  calculateExcelApplication,
  clearExcelRange,
  closeExcelWorkbookSession,
  createExcelTable,
  createExcelWorkbookSession,
  createExcelWorksheetChart,
  createExcelWorksheetPivotTable,
  deleteExcelTable,
  deleteExcelTableRow,
  deleteExcelWorksheet,
  deleteExcelWorksheetChart,
  deleteExcelWorksheetPivotTable,
  getExcelRange,
  getExcelTable,
  getExcelTableColumn,
  getExcelUsedRange,
  getExcelWorkbook,
  getExcelWorkbookNamedItem,
  getExcelWorksheet,
  getExcelWorksheetNamedItem,
  getExcelWorksheetPivotTable,
  listExcelTableColumns,
  listExcelTableRows,
  listExcelTables,
  listExcelWorkbookNames,
  listExcelWorksheetCharts,
  listExcelWorksheetNames,
  listExcelWorksheetPivotTables,
  listExcelWorksheets,
  mergeExcelSessionInit,
  patchExcelRange,
  patchExcelTable,
  patchExcelTableColumn,
  patchExcelTableRow,
  patchExcelWorksheetChart,
  patchExcelWorksheetPivotTable,
  refreshAllExcelWorksheetPivotTables,
  refreshExcelWorkbookSession,
  refreshExcelWorksheetPivotTable,
  updateExcelWorksheet
} from './graph-excel-client.js';

describe('mergeExcelSessionInit', () => {
  it('returns base init unchanged when session id is absent', () => {
    const base = { method: 'PATCH' as const, body: '{}' };
    expect(mergeExcelSessionInit(base, undefined)).toEqual(base);
    expect(mergeExcelSessionInit(base, '')).toEqual(base);
    expect(mergeExcelSessionInit(base, '   ')).toEqual(base);
  });

  it('adds workbook-session-id header', () => {
    const merged = mergeExcelSessionInit({ method: 'POST', body: '{}' }, 'sess-1');
    const h = new Headers(merged.headers as HeadersInit);
    expect(h.get('workbook-session-id')).toBe('sess-1');
  });

  it('preserves existing headers', () => {
    const merged = mergeExcelSessionInit(
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: '{}'
      },
      'abc'
    );
    const h = new Headers(merged.headers as HeadersInit);
    expect(h.get('Content-Type')).toBe('application/json');
    expect(h.get('workbook-session-id')).toBe('abc');
  });

  it('trims session id', () => {
    const merged = mergeExcelSessionInit({ method: 'GET' }, '  trim-me  ');
    expect(new Headers(merged.headers as HeadersInit).get('workbook-session-id')).toBe('trim-me');
  });
});

describe('listExcelWorksheets / getExcelRange', () => {
  const token = 'tok';
  const baseUrl = 'https://graph.microsoft.com/v1.0';

  it('lists worksheets under workbook', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        return new Response(JSON.stringify({ value: [{ id: 'ws1', name: 'Sheet1' }] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as unknown as typeof fetch;

      const r = await listExcelWorksheets(token, 'item-42');
      expect(r.ok).toBe(true);
      expect(r.data?.[0]?.name).toBe('Sheet1');
      expect(urls[0]).toContain('/me/drive/items/item-42/workbook/worksheets');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('reads A1 range with escaped quotes in address', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        return new Response(JSON.stringify({ address: 'A1:A1', values: [[1]] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as unknown as typeof fetch;

      const r = await getExcelRange(token, 'item-42', "Sheet'1", "A'1");
      expect(r.ok).toBe(true);
      expect(r.data?.values).toEqual([[1]]);
      const u = decodeURIComponent(urls[0]);
      expect(u).toContain("range(address='A''1')");
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('returns graphError when list worksheets fails', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async () =>
        new Response(JSON.stringify({ error: { message: 'nope' } }), {
          status: 403,
          headers: { 'content-type': 'application/json' }
        })) as unknown as typeof fetch;

      const r = await listExcelWorksheets(token, 'item-x');
      expect(r.ok).toBe(false);
      expect(r.error?.message).toContain('nope');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('covers used range, tables, worksheet CRUD, and workbook metadata', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async (_input: string | URL | Request, init?: RequestInit) => {
        const method = (init?.method || 'GET').toUpperCase();
        if (method === 'DELETE') {
          return new Response(null, { status: 204 });
        }
        if (method === 'POST' && String(init?.body || '').includes('worksheets/add')) {
          return new Response(JSON.stringify({ id: 'ws-new', name: 'Added' }), {
            status: 201,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (method === 'PATCH') {
          return new Response(JSON.stringify({ id: 'ws1', name: 'Renamed' }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (String(_input).includes('/tables') && !String(_input).includes('/tables/')) {
          return new Response(JSON.stringify({ value: [{ id: 'tbl1', name: 'T' }] }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (String(_input).includes('/workbook')) {
          return new Response(JSON.stringify({ id: 'wb' }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        return new Response(JSON.stringify({ address: 'A1', values: [[]] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as unknown as typeof fetch;

      const item = 'item-excel';
      const u1 = await getExcelUsedRange(token, item, 'Sheet1', true);
      expect(u1.ok).toBe(true);
      const tabs = await listExcelTables(token, item, 'Sheet1');
      expect(tabs.ok).toBe(true);
      expect(tabs.data?.[0]?.id).toBe('tbl1');
      const ws = await getExcelWorksheet(token, item, 'Sheet1');
      expect(ws.ok).toBe(true);
      const added = await addExcelWorksheet(token, item, 'New', 'sess');
      expect(added.ok).toBe(true);
      const upd = await updateExcelWorksheet(token, item, 'Sheet1', { name: 'Renamed' }, 'sess');
      expect(upd.ok).toBe(true);
      const del = await deleteExcelWorksheet(token, item, 'Sheet1', 'sess');
      expect(del.ok).toBe(true);
      const wb = await getExcelWorkbook(token, item);
      expect(wb.ok).toBe(true);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('covers patch/clear range, workbook query, calculate, and named items', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    const item = 'item-meta';
    try {
      globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
        const url = typeof input === 'string' ? input : input.toString();
        const method = (init?.method || 'GET').toUpperCase();
        if (url.includes('/range(address=') && url.endsWith('/clear') && method === 'POST') {
          return new Response(null, { status: 204 });
        }
        if (url.includes('/application/calculate') && method === 'POST') {
          return new Response(null, { status: 204 });
        }
        if (url.includes('/workbook/names') && !url.includes('/worksheets/')) {
          if (url.endsWith('/names') || url.endsWith('/names/')) {
            return new Response(JSON.stringify({ value: [{ id: 'n1', name: 'Foo' }] }), {
              status: 200,
              headers: { 'content-type': 'application/json' }
            });
          }
          return new Response(JSON.stringify({ id: 'n1', name: 'Foo' }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (url.includes('/worksheets/Sheet1/names')) {
          if (method === 'GET' && url.endsWith('/names')) {
            return new Response(JSON.stringify({ value: [{ id: 'wn1', name: 'Bar' }] }), {
              status: 200,
              headers: { 'content-type': 'application/json' }
            });
          }
          return new Response(JSON.stringify({ id: 'wn1', name: 'Bar' }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (method === 'PATCH') {
          return new Response(JSON.stringify({ address: 'A1', values: [[2]] }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (url.includes('workbook') && url.includes('$select=')) {
          return new Response(JSON.stringify({ application: {} }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        return new Response(JSON.stringify({}), { status: 200, headers: { 'content-type': 'application/json' } });
      }) as unknown as typeof fetch;

      const pr = await patchExcelRange(token, item, 'Sheet1', 'A1', { values: [[2]] }, 'sess');
      expect(pr.ok).toBe(true);
      const cl = await clearExcelRange(token, item, 'Sheet1', 'A1', { applyTo: 'All' }, 'sess');
      expect(cl.ok).toBe(true);
      const wbq = await getExcelWorkbook(token, item, '$select=application');
      expect(wbq.ok).toBe(true);
      const calc = await calculateExcelApplication(token, item, { calculationType: 'Recalculate' }, 'sess');
      expect(calc.ok).toBe(true);
      const names = await listExcelWorkbookNames(token, item);
      expect(names.ok).toBe(true);
      expect(names.data?.[0]?.id).toBe('n1');
      const ni = await getExcelWorkbookNamedItem(token, item, 'n1');
      expect(ni.ok).toBe(true);
      const wsn = await listExcelWorksheetNames(token, item, 'Sheet1');
      expect(wsn.ok).toBe(true);
      const wni = await getExcelWorksheetNamedItem(token, item, 'Sheet1', 'wn1');
      expect(wni.ok).toBe(true);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('graph-excel-client tables charts pivot session', () => {
  const token = 'tok';
  const baseUrl = 'https://graph.microsoft.com/v1.0';
  const item = 'item-bulk';

  it('covers table, chart, pivot, and workbook session endpoints', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
        const url = typeof input === 'string' ? input : input instanceof Request ? input.url : String(input);
        const method = (init?.method || 'GET').toUpperCase();

        if (url.includes('/createSession') && method === 'POST') {
          return new Response(JSON.stringify({ id: 'sid-1' }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (url.includes('/refreshSession') && method === 'POST') {
          return new Response(null, { status: 204 });
        }
        if (url.includes('/closeSession') && method === 'POST') {
          return new Response(null, { status: 204 });
        }

        if (url.includes('/pivotTables/refreshAll') && method === 'POST') {
          return new Response(null, { status: 204 });
        }
        if (url.includes('/pivotTables/p1/refresh') && method === 'POST') {
          return new Response(null, { status: 204 });
        }
        if (url.includes('/pivotTables/p1') && method === 'DELETE') {
          return new Response(null, { status: 204 });
        }
        if (url.includes('/pivotTables/p1') && method === 'PATCH') {
          return new Response(JSON.stringify({ id: 'p1' }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (url.includes('/pivotTables/p1') && method === 'GET') {
          return new Response(JSON.stringify({ id: 'p1' }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (url.includes('/worksheets/Sheet1/pivotTables') && method === 'POST') {
          return new Response(JSON.stringify({ id: 'p1' }), {
            status: 201,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (url.includes('/worksheets/Sheet1/pivotTables') && method === 'GET') {
          return new Response(JSON.stringify({ value: [{ id: 'p1' }] }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }

        if (url.includes('/worksheets/Sheet1/charts/ch1') && method === 'DELETE') {
          return new Response(null, { status: 204 });
        }
        if (url.includes('/worksheets/Sheet1/charts/ch1') && method === 'PATCH') {
          return new Response(JSON.stringify({ id: 'ch1' }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (url.includes('/worksheets/Sheet1/charts') && method === 'POST') {
          return new Response(JSON.stringify({ id: 'ch1' }), {
            status: 201,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (url.includes('/worksheets/Sheet1/charts') && method === 'GET') {
          return new Response(JSON.stringify({ value: [{ id: 'ch1' }] }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }

        if (url.includes('/tables/tbl1/rows/add') && method === 'POST') {
          return new Response(JSON.stringify({ index: 0 }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (url.includes('/tables/tbl1/rows/r1') && method === 'DELETE') {
          return new Response(null, { status: 204 });
        }
        if (url.includes('/tables/tbl1/rows/r1') && method === 'PATCH') {
          return new Response(JSON.stringify({ index: 0 }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (url.includes('/tables/tbl1/rows') && method === 'GET') {
          return new Response(JSON.stringify({ value: [{ index: 0, id: 'r1' }] }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }

        if (url.includes('/tables/tbl1/columns/c1') && method === 'PATCH') {
          return new Response(JSON.stringify({ id: 'c1', name: 'Col' }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (url.includes('/tables/tbl1/columns/c1') && method === 'GET') {
          return new Response(JSON.stringify({ id: 'c1', name: 'Col' }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (url.includes('/tables/tbl1/columns') && method === 'GET') {
          return new Response(JSON.stringify({ value: [{ id: 'c1' }] }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }

        if (url.includes('/tables/tbl1') && method === 'DELETE') {
          return new Response(null, { status: 204 });
        }
        if (url.includes('/tables/tbl1') && method === 'PATCH') {
          return new Response(JSON.stringify({ id: 'tbl1', name: 'T' }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (url.includes('/tables/tbl1') && method === 'GET') {
          return new Response(JSON.stringify({ id: 'tbl1', name: 'T' }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }

        if (url.includes('/worksheets/Sheet1/tables') && method === 'POST') {
          return new Response(JSON.stringify({ id: 'tbl-ws' }), {
            status: 201,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (url.includes('/workbook/tables') && method === 'POST' && !url.includes('worksheets')) {
          return new Response(JSON.stringify({ id: 'tbl-root' }), {
            status: 201,
            headers: { 'content-type': 'application/json' }
          });
        }

        return new Response(JSON.stringify({ value: [] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as unknown as typeof fetch;

      const sess = await createExcelWorkbookSession(token, item);
      expect(sess.ok).toBe(true);
      expect(sess.data?.id).toBe('sid-1');
      const refr = await refreshExcelWorkbookSession(token, item, 'sid-1');
      expect(refr.ok).toBe(true);
      const cls = await closeExcelWorkbookSession(token, item, 'sid-1');
      expect(cls.ok).toBe(true);

      expect((await getExcelTable(token, item, 'tbl1')).ok).toBe(true);
      expect((await createExcelTable(token, item, { name: 'T' })).ok).toBe(true);
      expect((await createExcelTable(token, item, { name: 'T' }, 'Sheet1', 's')).ok).toBe(true);
      expect((await patchExcelTable(token, item, 'tbl1', { name: 'X' }, 's')).ok).toBe(true);

      expect((await listExcelTableColumns(token, item, 'tbl1')).ok).toBe(true);
      expect((await getExcelTableColumn(token, item, 'tbl1', 'c1')).ok).toBe(true);
      expect((await patchExcelTableColumn(token, item, 'tbl1', 'c1', { name: 'C' }, 's')).ok).toBe(true);

      expect((await listExcelTableRows(token, item, 'tbl1', 5)).ok).toBe(true);
      expect((await addExcelTableRows(token, item, 'tbl1', { values: [[]] }, 's')).ok).toBe(true);
      expect((await patchExcelTableRow(token, item, 'tbl1', 'r1', {}, 's')).ok).toBe(true);
      expect((await deleteExcelTableRow(token, item, 'tbl1', 'r1', 's')).ok).toBe(true);

      expect((await listExcelWorksheetCharts(token, item, 'Sheet1')).ok).toBe(true);
      expect((await createExcelWorksheetChart(token, item, 'Sheet1', { type: 'ColumnClustered' }, 's')).ok).toBe(true);
      expect((await patchExcelWorksheetChart(token, item, 'Sheet1', 'ch1', {}, 's')).ok).toBe(true);
      expect((await deleteExcelWorksheetChart(token, item, 'Sheet1', 'ch1', 's')).ok).toBe(true);

      expect((await listExcelWorksheetPivotTables(token, item, 'Sheet1')).ok).toBe(true);
      expect((await getExcelWorksheetPivotTable(token, item, 'Sheet1', 'p1')).ok).toBe(true);
      expect((await createExcelWorksheetPivotTable(token, item, 'Sheet1', {}, 's')).ok).toBe(true);
      expect((await patchExcelWorksheetPivotTable(token, item, 'Sheet1', 'p1', {}, 's')).ok).toBe(true);
      expect((await refreshExcelWorksheetPivotTable(token, item, 'Sheet1', 'p1', 's')).ok).toBe(true);
      expect((await refreshAllExcelWorksheetPivotTables(token, item, 'Sheet1', 's')).ok).toBe(true);
      expect((await deleteExcelWorksheetPivotTable(token, item, 'Sheet1', 'p1', 's')).ok).toBe(true);

      expect((await deleteExcelTable(token, item, 'tbl1', 's')).ok).toBe(true);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
  it('listExcelTableRows pages through @odata.nextLink when no top is given', async () => {
    process.env.GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';
    const originalFetch = globalThis.fetch;
    try {
      let n = 0;
      globalThis.fetch = (async () => {
        n++;
        if (n === 1) {
          return new Response(
            JSON.stringify({
              value: [{ index: 0 }],
              '@odata.nextLink': 'https://graph.microsoft.com/v1.0/next?$skiptoken=2'
            }),
            { status: 200, headers: { 'content-type': 'application/json' } }
          );
        }
        return new Response(JSON.stringify({ value: [{ index: 1 }] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as unknown as typeof fetch;
      const { listExcelTableRows } = await import('./graph-excel-client.js');
      const r = await listExcelTableRows(token, 'item1', 'tbl1');
      expect(r.ok).toBe(true);
      expect(r.data?.map((x) => x.index)).toEqual([0, 1]);
      expect(n).toBe(2);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});
