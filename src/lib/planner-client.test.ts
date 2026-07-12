import { describe, expect, it } from 'bun:test';
import { mergePlannerAssignments } from './planner-client.js';

describe('mergePlannerAssignments', () => {
  it('removes an assignee by setting the key to null (open-type merge), not by omitting it', () => {
    const current = {
      a: { '@odata.type': '#microsoft.graph.plannerAssignment', orderHint: 'x' },
      b: { '@odata.type': '#microsoft.graph.plannerAssignment', orderHint: 'y' }
    };
    const out = mergePlannerAssignments(current, ['c'], ['a']);
    // Removed 'a' must be present as null so the server actually drops it.
    expect(Object.hasOwn(out, 'a')).toBe(true);
    expect(out.a).toBeNull();
    // Unchanged 'b' kept; added 'c' is a full assignment object.
    expect(out.b).not.toBeNull();
    expect((out.c as { '@odata.type': string })['@odata.type']).toBe('#microsoft.graph.plannerAssignment');
  });
});

describe('createPlannerPlanForSignedInUser', () => {
  it('resolves /me then POSTs beta /me/planner/plans with user container', async () => {
    process.env.GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';
    process.env.GRAPH_BETA_URL = 'https://graph.microsoft.com/beta';
    const urls: string[] = [];
    const bodies: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        if (init?.body && typeof init.body === 'string') bodies.push(init.body);
        if (urls.length === 1) {
          return new Response(JSON.stringify({ id: 'user-guid-1' }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        return new Response(
          JSON.stringify({
            id: 'plan-1',
            title: 'My plan',
            container: {
              url: 'https://graph.microsoft.com/beta/users/user-guid-1',
              type: 'user'
            }
          }),
          { status: 201, headers: { 'content-type': 'application/json' } }
        );
      }) as unknown as typeof fetch;

      const { createPlannerPlanForSignedInUser } = await import('./planner-client.js');
      const r = await createPlannerPlanForSignedInUser('tok', 'My plan');

      expect(r.ok).toBe(true);
      expect(r.data?.id).toBe('plan-1');
      expect(urls[0]).toContain('/v1.0/me');
      expect(urls[0]).toContain('$select=id');
      expect(urls[1]).toContain('graph.microsoft.com/beta/me/planner/plans');
      const postBody = bodies[0];
      expect(postBody).toContain('user-guid-1');
      expect(postBody).toContain('"type":"user"');
      expect(postBody).toContain('"title":"My plan"');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('listPlannerMyDayTasks and listPlannerRecentPlans', () => {
  it('GETs beta /me/planner/myDayTasks and /me/planner/recentPlans when user is omitted', async () => {
    process.env.GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';
    process.env.GRAPH_BETA_URL = 'https://graph.microsoft.com/beta';
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

      const { listPlannerMyDayTasks, listPlannerRecentPlans } = await import('./planner-client.js');
      await listPlannerMyDayTasks('tok');
      await listPlannerRecentPlans('tok');

      expect(urls.some((u) => u.includes('/beta/me/planner/myDayTasks'))).toBe(true);
      expect(urls.some((u) => u.includes('/beta/me/planner/recentPlans'))).toBe(true);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('uses /users/{id}/planner/... when user is set', async () => {
    process.env.GRAPH_BETA_URL = 'https://graph.microsoft.com/beta';
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

      const { listPlannerMyDayTasks } = await import('./planner-client.js');
      await listPlannerMyDayTasks('tok', 'alice@contoso.com');

      expect(urls[0]).toContain('/beta/users/alice%40contoso.com/planner/myDayTasks');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('listUserTasks, listPlanBuckets, getTask, archivePlannerPlan', () => {
  it('lists /me/planner/tasks and plan buckets', async () => {
    process.env.GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';
    process.env.GRAPH_BETA_URL = 'https://graph.microsoft.com/beta';
    const urls: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        return new Response(JSON.stringify({ value: [{ id: 't1', title: 'Task' }] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as unknown as typeof fetch;

      const { listUserTasks, listPlanBuckets } = await import('./planner-client.js');
      const tasks = await listUserTasks('tok');
      expect(tasks.ok).toBe(true);
      expect(tasks.data?.[0]?.id).toBe('t1');
      expect(urls.some((u) => u.includes('/v1.0/me/planner/tasks'))).toBe(true);

      const buckets = await listPlanBuckets('tok', 'plan-9');
      expect(buckets.ok).toBe(true);
      expect(urls.some((u) => u.includes('/v1.0/planner/plans/plan-9/buckets'))).toBe(true);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('getTask GETs /planner/tasks/{id}', async () => {
    process.env.GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';
    const urls: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        return new Response(JSON.stringify({ id: 'task-1', title: 'One' }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as unknown as typeof fetch;

      const { getTask } = await import('./planner-client.js');
      const r = await getTask('tok', 'task-1');
      expect(r.ok).toBe(true);
      expect(r.data?.title).toBe('One');
      expect(urls[0]).toContain('/planner/tasks/task-1');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('archivePlannerPlan POSTs beta archive with If-Match', async () => {
    process.env.GRAPH_BETA_URL = 'https://graph.microsoft.com/beta';
    const calls: { url: string; init?: RequestInit }[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
        calls.push({ url: typeof input === 'string' ? input : input.toString(), init });
        return new Response(null, { status: 204 });
      }) as unknown as typeof fetch;

      const { archivePlannerPlan } = await import('./planner-client.js');
      const r = await archivePlannerPlan('tok', 'plan-a', 'W/"1"', 'done');
      expect(r.ok).toBe(true);
      expect(calls[0].url).toContain('/beta/planner/plans/plan-a/archive');
      expect(calls[0].init?.method).toBe('POST');
      const h = new Headers(calls[0].init?.headers);
      expect(h.get('If-Match')).toBe('W/"1"');
      expect(JSON.parse((calls[0].init?.body as string) || '{}')).toEqual({ justification: 'done' });
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('updatePlannerUser', () => {
  it('PATCHes beta …/planner with If-Match and merge body', async () => {
    process.env.GRAPH_BETA_URL = 'https://graph.microsoft.com/beta';
    const calls: { url: string; init?: RequestInit }[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
        calls.push({
          url: typeof input === 'string' ? input : input.toString(),
          init
        });
        return new Response(JSON.stringify({ id: 'pu1', '@odata.etag': 'W/"2"', favoritePlanReferences: {} }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as unknown as typeof fetch;

      const { updatePlannerUser } = await import('./planner-client.js');
      const r = await updatePlannerUser('tok', undefined, 'W/"1"', { recentPlanReferences: { p1: null } });

      expect(r.ok).toBe(true);
      expect(calls.length).toBe(1);
      expect(calls[0].url).toContain('/beta/me/planner');
      expect(calls[0].init?.method).toBe('PATCH');
      const h = new Headers(calls[0].init?.headers);
      expect(h.get('If-Match')).toBe('W/"1"');
      expect(JSON.parse((calls[0].init?.body as string) || '{}')).toEqual({ recentPlanReferences: { p1: null } });
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('planner batch coverage', () => {
  const v1 = 'https://graph.microsoft.com/v1.0';
  const beta = 'https://graph.microsoft.com/beta';

  it('plans, buckets, tasks, details, favorites, roster, delta', async () => {
    process.env.GRAPH_BASE_URL = v1;
    process.env.GRAPH_BETA_URL = beta;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
        const u = typeof input === 'string' ? input : input.toString();
        const m = (init?.method || 'GET').toUpperCase();
        if (m === 'DELETE') return new Response(null, { status: 204 });
        if (m === 'PATCH' && u.includes('/planner/tasks/') && !u.includes('taskBoardFormat')) {
          return new Response(null, { status: 204 });
        }
        if (m === 'PATCH' && u.includes('/planner/plans/') && u.includes('/details')) {
          return new Response(null, { status: 204 });
        }
        if (m === 'PATCH' && u.includes('/planner/buckets/')) {
          return new Response(null, { status: 204 });
        }
        if (m === 'PATCH' && u.includes('/planner/plans/') && !u.includes('/details')) {
          return new Response(null, { status: 204 });
        }
        if (u.includes('/planner/plans/') && u.includes('/details') && m === 'GET') {
          return new Response(JSON.stringify({ id: 'pd1', categoryDescriptions: {} }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (u.includes('/planner/tasks/') && u.endsWith('/details') && m === 'GET') {
          return new Response(
            JSON.stringify({
              id: 'td1',
              '@odata.etag': 'W/td',
              checklist: {
                ck: { '@odata.type': '#microsoft.graph.plannerChecklistItem', title: 't', isChecked: false }
              },
              references: {}
            }),
            { status: 200, headers: { 'content-type': 'application/json' } }
          );
        }
        if (u.includes('/planner/taskDetails/') && m === 'PATCH') {
          return new Response(null, { status: 204 });
        }
        if (u.includes('/delta')) {
          return new Response(JSON.stringify({ value: [] }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (m === 'POST' && u.includes('/planner/tasks')) {
          return new Response(JSON.stringify({ id: 'new-task', title: 'T', '@odata.etag': 'W/n' }), {
            status: 201,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (m === 'POST' && u.includes('/planner/buckets')) {
          return new Response(JSON.stringify({ id: 'nb', name: 'B' }), {
            status: 201,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (u.includes('/planner/plans/plan-1') && !u.includes('/details') && m === 'GET') {
          return new Response(JSON.stringify({ id: 'plan-1', title: 'P' }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (u.includes('/planner/buckets/b1') && m === 'GET') {
          return new Response(JSON.stringify({ id: 'b1', name: 'B' }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (u.includes('/planner/tasks/task-1') && !u.includes('/details') && m === 'GET') {
          return new Response(JSON.stringify({ id: 'task-1', title: 'One' }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        return new Response(JSON.stringify({ value: [{ id: 'p1', title: 'P' }] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as unknown as typeof fetch;
      const p = await import('./planner-client.js');
      expect((await p.listUserPlans('tok')).ok).toBe(true);
      expect((await p.listPlannerPlansForUser('tok', 'u1')).ok).toBe(true);
      expect((await p.listGroupPlans('tok', 'g1')).ok).toBe(true);
      expect((await p.listPlanTasks('tok', 'plan-1')).ok).toBe(true);
      expect((await p.getPlanDetails('tok', 'plan-1')).ok).toBe(true);
      expect((await p.getPlannerPlan('tok', 'plan-1')).ok).toBe(true);
      expect((await p.updatePlannerPlan('tok', 'plan-1', 'W/1', 'X')).ok).toBe(true);
      expect((await p.deletePlannerPlan('tok', 'plan-1', 'W/1')).ok).toBe(true);
      expect((await p.unarchivePlannerPlan('tok', 'plan-1', 'W/1', 'ok')).ok).toBe(true);
      expect((await p.getPlannerBucket('tok', 'b1')).ok).toBe(true);
      expect((await p.createPlannerBucket('tok', 'plan-1', 'B')).ok).toBe(true);
      expect((await p.updatePlannerBucket('tok', 'b1', 'W/1', { name: 'N' })).ok).toBe(true);
      expect((await p.deletePlannerBucket('tok', 'b1', 'W/1')).ok).toBe(true);
      expect((await p.createTask('tok', 'plan-1', 'T', 'b1')).ok).toBe(true);
      expect((await p.updateTask('tok', 'task-1', 'W/1', { title: 'U' })).ok).toBe(true);
      expect((await p.deletePlannerTask('tok', 'task-1', 'W/1')).ok).toBe(true);
      expect((await p.getPlannerTaskDetails('tok', 'task-1')).ok).toBe(true);
      expect((await p.updatePlannerTaskDetails('tok', 'td1', 'W/td', { description: 'd' })).ok).toBe(true);
      expect((await p.updatePlannerPlanDetails('tok', 'plan-1', 'W/1', { categoryDescriptions: {} })).ok).toBe(true);
      expect((await p.deletePlannerPlanDetails('tok', 'plan-1', 'W/1')).ok).toBe(true);
      expect((await p.deletePlannerTaskDetails('tok', 'task-1', 'W/1')).ok).toBe(true);
      expect((await p.listFavoritePlans('tok')).ok).toBe(true);
      expect((await p.listRosterPlans('tok')).ok).toBe(true);
      expect((await p.getPlannerDeltaPage('tok', `${beta}/me/planner/all/delta`)).ok).toBe(true);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('checklist items and references on planner task', async () => {
    process.env.GRAPH_BASE_URL = v1;
    const originalFetch = globalThis.fetch;
    const patchBodies: Record<string, unknown>[] = [];
    try {
      globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
        const u = typeof input === 'string' ? input : input.toString();
        const m = (init?.method || 'GET').toUpperCase();
        if (u.includes('/planner/tasks/task-1/details') && m === 'GET') {
          return new Response(
            JSON.stringify({
              id: 'td1',
              '@odata.etag': 'W/td',
              checklist: {
                ck: { '@odata.type': '#microsoft.graph.plannerChecklistItem', title: 'old', isChecked: false }
              },
              references: { 'https://x': { alias: 'a' } }
            }),
            { status: 200, headers: { 'content-type': 'application/json' } }
          );
        }
        if (u.includes('/planner/taskDetails/') && m === 'PATCH') {
          if (init?.body && typeof init.body === 'string') patchBodies.push(JSON.parse(init.body));
          return new Response(null, { status: 204 });
        }
        return new Response('{}', { status: 404 });
      }) as unknown as typeof fetch;
      const p = await import('./planner-client.js');
      expect((await p.addPlannerChecklistItem('tok', 'task-1', 'New item', 'new-ck')).ok).toBe(true);
      expect((await p.removePlannerChecklistItem('tok', 'task-1', 'ck')).ok).toBe(true);
      expect((await p.addPlannerReference('tok', 'task-1', 'https://ref', 'alias')).ok).toBe(true);
      expect((await p.removePlannerReference('tok', 'task-1', 'https://x')).ok).toBe(true);
      expect((await p.updatePlannerChecklistItem('tok', 'task-1', 'ck', { isChecked: true })).ok).toBe(true);

      // Open-type dictionaries are merged per-key: removal MUST send the key as null, not omit it.
      const removeChecklist = patchBodies[1] as { checklist: Record<string, unknown> };
      expect(removeChecklist.checklist.ck).toBeNull();
      const removeRef = patchBodies[3] as { references: Record<string, unknown> };
      // 'https://x' → key encoded as 'https%3A//x'; it must be present with a null value.
      expect(removeRef.references['https%3A//x']).toBeNull();
      expect(Object.hasOwn(removeRef.references, 'https%3A//x')).toBe(true);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('task board formats', async () => {
    process.env.GRAPH_BASE_URL = v1;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
        const u = typeof input === 'string' ? input : input.toString();
        const m = (init?.method || 'GET').toUpperCase();
        if (m === 'PATCH') return new Response(null, { status: 204 });
        if (u.includes('assignedToTaskBoardFormat')) {
          return new Response(JSON.stringify({ id: 'af', orderHintsByAssignee: {} }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (u.includes('bucketTaskBoardFormat')) {
          return new Response(JSON.stringify({ id: 'bf' }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (u.includes('progressTaskBoardFormat')) {
          return new Response(JSON.stringify({ id: 'pf' }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        return new Response('{}', { status: 404 });
      }) as unknown as typeof fetch;
      const p = await import('./planner-client.js');
      expect((await p.getAssignedToTaskBoardFormat('tok', 'task-1')).ok).toBe(true);
      expect((await p.updateAssignedToTaskBoardFormat('tok', 'task-1', 'W/1', { orderHintsByAssignee: {} })).ok).toBe(
        true
      );
      expect((await p.getBucketTaskBoardFormat('tok', 'task-1')).ok).toBe(true);
      expect((await p.updateBucketTaskBoardFormat('tok', 'task-1', 'W/1', ' !')).ok).toBe(true);
      expect((await p.getProgressTaskBoardFormat('tok', 'task-1')).ok).toBe(true);
      expect((await p.updateProgressTaskBoardFormat('tok', 'task-1', 'W/1', ' !')).ok).toBe(true);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('getPlannerUser', async () => {
    process.env.GRAPH_BETA_URL = beta;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async () =>
        new Response(JSON.stringify({ id: 'pu', favoritePlanReferences: {} }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        })) as unknown as typeof fetch;
      const { getPlannerUser } = await import('./planner-client.js');
      expect((await getPlannerUser('tok')).ok).toBe(true);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});
