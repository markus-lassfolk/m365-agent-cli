import { randomUUID } from 'node:crypto';
import {
  callGraph,
  callGraphAbsolute,
  callGraphAt,
  fetchAllPages,
  GraphApiError,
  type GraphResponse,
  graphError,
  graphResult
} from './graph-client.js';
import { GRAPH_BASE_URL, GRAPH_BETA_URL } from './graph-constants.js';

export interface PlannerPlanContainer {
  '@odata.type'?: string;
  url?: string;
  containerId?: string;
  type?: string;
}

export interface PlannerPlan {
  id: string;
  title: string;
  /** @deprecated Prefer `container`; Graph may still return this for group-backed plans. */
  owner?: string;
  container?: PlannerPlanContainer;
  '@odata.etag'?: string;
}

export interface PlannerBucket {
  id: string;
  name: string;
  planId: string;
  orderHint?: string;
  '@odata.etag'?: string;
}

/** Planner label slots (plan defines display names in plan details). */
export type PlannerCategorySlot = 'category1' | 'category2' | 'category3' | 'category4' | 'category5' | 'category6';

export type PlannerAppliedCategories = Partial<Record<PlannerCategorySlot, boolean>>;

export interface PlannerTask {
  id: string;
  planId: string;
  bucketId: string;
  title: string;
  orderHint: string;
  assigneePriority: string;
  percentComplete: number;
  hasDescription: boolean;
  createdDateTime: string;
  /** Planner due/start (ISO date-time strings per Graph). */
  dueDateTime?: string;
  startDateTime?: string;
  /** Teams / Outlook thread id when linked. */
  conversationThreadId?: string;
  /** 0 (highest) .. 10 (lowest). */
  priority?: number;
  /** Card preview on task: automatic, noPreview, checklist, description, reference. */
  previewType?: string;
  assignments?: Record<string, any>;
  /** Label slots (boolean per category1..category6); names come from plan details. */
  appliedCategories?: PlannerAppliedCategories;
  '@odata.etag'?: string;
}

/** Checklist entry on Planner task details (keyed by client-generated id in PATCH body). */
export interface PlannerTaskDetailsChecklistItem {
  '@odata.type'?: string;
  isChecked: boolean;
  title: string;
  orderHint: string;
  lastModifiedDateTime?: string;
  lastModifiedBy?: { user?: { id?: string } };
}

export interface PlannerTaskDetails {
  id: string;
  description?: string;
  checklist?: Record<string, PlannerTaskDetailsChecklistItem>;
  references?: Record<string, unknown>;
  previewType?: string;
  '@odata.etag'?: string;
}

export interface PlannerPlanDetails {
  id: string;
  categoryDescriptions?: Partial<Record<PlannerCategorySlot, string>>;
  sharedWith?: Record<string, boolean>;
  '@odata.etag'?: string;
}

const PLANNER_SLOTS: PlannerCategorySlot[] = [
  'category1',
  'category2',
  'category3',
  'category4',
  'category5',
  'category6'
];

/** Accept `1`..`6` or `category1`..`category6` (case-insensitive). */
export function parsePlannerLabelKey(input: string): PlannerCategorySlot | null {
  const t = input.trim().toLowerCase();
  const m = t.match(/^category([1-6])$/);
  if (m) return `category${m[1]}` as PlannerCategorySlot;
  if (/^[1-6]$/.test(t)) return `category${t}` as PlannerCategorySlot;
  return null;
}

/** Build a full slot map for PATCH (Planner expects explicit booleans per slot). */
export function normalizeAppliedCategories(
  current: PlannerAppliedCategories | undefined,
  patch: { clearAll?: boolean; setTrue?: PlannerCategorySlot[]; setFalse?: PlannerCategorySlot[] }
): PlannerAppliedCategories {
  const out: PlannerAppliedCategories = {};
  for (const s of PLANNER_SLOTS) {
    if (patch.clearAll) {
      out[s] = false;
      continue;
    }
    let v = current?.[s] === true;
    for (const u of patch.setTrue ?? []) if (u === s) v = true;
    for (const u of patch.setFalse ?? []) if (u === s) v = false;
    out[s] = v;
  }
  return out;
}

/** Build `assignments` for create/update from user object IDs. */
export function buildPlannerAssignments(assigneeUserIds: string[]): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  for (const id of assigneeUserIds) {
    out[id] = {
      '@odata.type': '#microsoft.graph.plannerAssignment',
      orderHint: ' !'
    };
  }
  return out;
}

/** Merge assignees: add and remove user IDs from current `assignments` map. */
export function mergePlannerAssignments(
  current: Record<string, unknown> | undefined,
  addUserIds: string[],
  removeUserIds: string[]
): Record<string, unknown> {
  const out: Record<string, unknown> = { ...(current || {}) };
  for (const r of removeUserIds) delete out[r];
  for (const a of addUserIds) {
    out[a] = {
      '@odata.type': '#microsoft.graph.plannerAssignment',
      orderHint: ' !'
    };
  }
  return out;
}

export async function listUserTasks(token: string): Promise<GraphResponse<PlannerTask[]>> {
  return fetchAllPages<PlannerTask>(token, '/me/planner/tasks', 'Failed to list tasks');
}

export async function listUserPlans(token: string): Promise<GraphResponse<PlannerPlan[]>> {
  return fetchAllPages<PlannerPlan>(token, '/me/planner/plans', 'Failed to list plans');
}

/**
 * List Planner tasks for a user (`GET /users/{id}/planner/tasks`).
 * Graph may return 403 for users other than the signed-in user depending on tenant and token.
 */
export async function listPlannerTasksForUser(token: string, userId: string): Promise<GraphResponse<PlannerTask[]>> {
  return fetchAllPages<PlannerTask>(
    token,
    `/users/${encodeURIComponent(userId)}/planner/tasks`,
    'Failed to list tasks for user'
  );
}

/**
 * List Planner plans for a user (`GET /users/{id}/planner/plans`).
 * Same permission caveats as {@link listPlannerTasksForUser}.
 */
export async function listPlannerPlansForUser(token: string, userId: string): Promise<GraphResponse<PlannerPlan[]>> {
  return fetchAllPages<PlannerPlan>(
    token,
    `/users/${encodeURIComponent(userId)}/planner/plans`,
    'Failed to list plans for user'
  );
}

export async function listGroupPlans(token: string, groupId: string): Promise<GraphResponse<PlannerPlan[]>> {
  return fetchAllPages<PlannerPlan>(
    token,
    `/groups/${encodeURIComponent(groupId)}/planner/plans`,
    'Failed to list group plans'
  );
}

export async function listPlanBuckets(token: string, planId: string): Promise<GraphResponse<PlannerBucket[]>> {
  return fetchAllPages<PlannerBucket>(
    token,
    `/planner/plans/${encodeURIComponent(planId)}/buckets`,
    'Failed to list buckets'
  );
}

export async function listPlanTasks(token: string, planId: string): Promise<GraphResponse<PlannerTask[]>> {
  return fetchAllPages<PlannerTask>(
    token,
    `/planner/plans/${encodeURIComponent(planId)}/tasks`,
    'Failed to list plan tasks'
  );
}

export async function getPlanDetails(token: string, planId: string): Promise<GraphResponse<PlannerPlanDetails>> {
  try {
    const result = await callGraph<PlannerPlanDetails>(token, `/planner/plans/${encodeURIComponent(planId)}/details`);
    if (!result.ok || !result.data) {
      return graphError(
        result.error?.message || 'Failed to get plan details',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get plan details');
  }
}

export async function getTask(token: string, taskId: string): Promise<GraphResponse<PlannerTask>> {
  try {
    const result = await callGraph<PlannerTask>(token, `/planner/tasks/${encodeURIComponent(taskId)}`);
    if (!result.ok || !result.data) {
      return graphError(result.error?.message || 'Failed to get task', result.error?.code, result.error?.status);
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get task');
  }
}

export interface CreatePlannerTaskExtras {
  /** ISO date-time after creation (PATCH). */
  dueDateTime?: string | null;
  startDateTime?: string | null;
  conversationThreadId?: string | null;
  orderHint?: string | null;
  assigneePriority?: string | null;
  priority?: number | null;
  previewType?: string | null;
}

export async function createTask(
  token: string,
  planId: string,
  title: string,
  bucketId?: string,
  assignments?: Record<string, any>,
  appliedCategories?: PlannerAppliedCategories,
  extras?: CreatePlannerTaskExtras
): Promise<GraphResponse<PlannerTask>> {
  try {
    const body: Record<string, unknown> = { planId, title };
    if (bucketId) body.bucketId = bucketId;
    if (assignments) body.assignments = assignments;

    const result = await callGraph<PlannerTask>(token, '/planner/tasks', {
      method: 'POST',
      body: JSON.stringify(body)
    });
    if (!result.ok || !result.data) {
      return graphError(result.error?.message || 'Failed to create task', result.error?.code, result.error?.status);
    }
    let task = result.data;
    const etag1 = task['@odata.etag'];
    const patchPayload: Record<string, unknown> = {};
    if (appliedCategories && Object.keys(appliedCategories).length > 0) {
      patchPayload.appliedCategories = appliedCategories;
    }
    if (extras?.dueDateTime !== undefined && extras.dueDateTime !== null) {
      patchPayload.dueDateTime = extras.dueDateTime;
    }
    if (extras?.startDateTime !== undefined && extras.startDateTime !== null) {
      patchPayload.startDateTime = extras.startDateTime;
    }
    if (extras?.conversationThreadId !== undefined && extras.conversationThreadId !== null) {
      patchPayload.conversationThreadId = extras.conversationThreadId;
    }
    if (extras?.orderHint !== undefined && extras.orderHint !== null) {
      patchPayload.orderHint = extras.orderHint;
    }
    if (extras?.assigneePriority !== undefined && extras.assigneePriority !== null) {
      patchPayload.assigneePriority = extras.assigneePriority;
    }
    if (extras?.priority !== undefined && extras.priority !== null) {
      patchPayload.priority = extras.priority;
    }
    if (extras?.previewType !== undefined && extras.previewType !== null) {
      patchPayload.previewType = extras.previewType;
    }
    if (Object.keys(patchPayload).length > 0) {
      if (!etag1) {
        return graphError('Created task missing ETag; cannot set fields', 'MISSING_ETAG', 500);
      }
      const patch = await callGraph<void>(token, `/planner/tasks/${encodeURIComponent(task.id)}`, {
        method: 'PATCH',
        headers: { 'If-Match': etag1 },
        body: JSON.stringify(patchPayload)
      });
      if (!patch.ok) {
        return graphError(patch.error?.message || 'Failed to update new task', patch.error?.code, patch.error?.status);
      }
      const again = await getTask(token, task.id);
      if (again.ok && again.data) task = again.data;
    }
    return graphResult(task);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to create task');
  }
}

export async function updateTask(
  token: string,
  taskId: string,
  etag: string,
  updates: {
    title?: string;
    bucketId?: string;
    assignments?: Record<string, any> | null;
    percentComplete?: number;
    appliedCategories?: PlannerAppliedCategories;
    dueDateTime?: string | null;
    startDateTime?: string | null;
    orderHint?: string | null;
    conversationThreadId?: string | null;
    assigneePriority?: string | null;
    priority?: number | null;
    previewType?: string | null;
  }
): Promise<GraphResponse<void>> {
  try {
    const result = await callGraph<void>(token, `/planner/tasks/${encodeURIComponent(taskId)}`, {
      method: 'PATCH',
      headers: {
        'If-Match': etag
      },
      body: JSON.stringify(updates)
    });
    if (!result.ok) {
      return graphError(result.error?.message || 'Failed to update task', result.error?.code, result.error?.status);
    }
    return graphResult(undefined as undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to update task');
  }
}

export async function getPlannerPlan(token: string, planId: string): Promise<GraphResponse<PlannerPlan>> {
  try {
    const result = await callGraph<PlannerPlan>(token, `/planner/plans/${encodeURIComponent(planId)}`);
    if (!result.ok || !result.data) {
      return graphError(result.error?.message || 'Failed to get plan', result.error?.code, result.error?.status);
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get plan');
  }
}

export async function createPlannerPlan(
  token: string,
  ownerGroupId: string,
  title: string
): Promise<GraphResponse<PlannerPlan>> {
  try {
    const base = GRAPH_BASE_URL.replace(/\/$/, '');
    const result = await callGraph<PlannerPlan>(token, '/planner/plans', {
      method: 'POST',
      body: JSON.stringify({
        title,
        container: {
          url: `${base}/groups/${encodeURIComponent(ownerGroupId)}`
        }
      })
    });
    if (!result.ok || !result.data) {
      return graphError(result.error?.message || 'Failed to create plan', result.error?.code, result.error?.status);
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to create plan');
  }
}

export async function updatePlannerPlan(
  token: string,
  planId: string,
  etag: string,
  title: string
): Promise<GraphResponse<void>> {
  try {
    const result = await callGraph<void>(token, `/planner/plans/${encodeURIComponent(planId)}`, {
      method: 'PATCH',
      headers: { 'If-Match': etag },
      body: JSON.stringify({ title })
    });
    if (!result.ok) {
      return graphError(result.error?.message || 'Failed to update plan', result.error?.code, result.error?.status);
    }
    return graphResult(undefined as undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to update plan');
  }
}

export async function deletePlannerPlan(token: string, planId: string, etag: string): Promise<GraphResponse<void>> {
  try {
    const result = await callGraph<void>(
      token,
      `/planner/plans/${encodeURIComponent(planId)}`,
      {
        method: 'DELETE',
        headers: { 'If-Match': etag }
      },
      false
    );
    if (!result.ok) {
      return graphError(result.error?.message || 'Failed to delete plan', result.error?.code, result.error?.status);
    }
    return graphResult(undefined as undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to delete plan');
  }
}

export async function getPlannerBucket(token: string, bucketId: string): Promise<GraphResponse<PlannerBucket>> {
  try {
    const result = await callGraph<PlannerBucket>(token, `/planner/buckets/${encodeURIComponent(bucketId)}`);
    if (!result.ok || !result.data) {
      return graphError(result.error?.message || 'Failed to get bucket', result.error?.code, result.error?.status);
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get bucket');
  }
}

export async function createPlannerBucket(
  token: string,
  planId: string,
  name: string
): Promise<GraphResponse<PlannerBucket>> {
  try {
    const result = await callGraph<PlannerBucket>(token, '/planner/buckets', {
      method: 'POST',
      body: JSON.stringify({ planId, name, orderHint: ' !' })
    });
    if (!result.ok || !result.data) {
      return graphError(result.error?.message || 'Failed to create bucket', result.error?.code, result.error?.status);
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to create bucket');
  }
}

export async function updatePlannerBucket(
  token: string,
  bucketId: string,
  etag: string,
  updates: { name?: string; orderHint?: string }
): Promise<GraphResponse<void>> {
  const body: Record<string, unknown> = {};
  if (updates.name !== undefined) body.name = updates.name;
  if (updates.orderHint !== undefined) body.orderHint = updates.orderHint;
  if (Object.keys(body).length === 0) {
    return graphError('No bucket updates', 'NO_UPDATES', 400);
  }
  try {
    const result = await callGraph<void>(token, `/planner/buckets/${encodeURIComponent(bucketId)}`, {
      method: 'PATCH',
      headers: { 'If-Match': etag },
      body: JSON.stringify(body)
    });
    if (!result.ok) {
      return graphError(result.error?.message || 'Failed to update bucket', result.error?.code, result.error?.status);
    }
    return graphResult(undefined as undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to update bucket');
  }
}

export async function deletePlannerBucket(token: string, bucketId: string, etag: string): Promise<GraphResponse<void>> {
  try {
    const result = await callGraph<void>(
      token,
      `/planner/buckets/${encodeURIComponent(bucketId)}`,
      {
        method: 'DELETE',
        headers: { 'If-Match': etag }
      },
      false
    );
    if (!result.ok) {
      return graphError(result.error?.message || 'Failed to delete bucket', result.error?.code, result.error?.status);
    }
    return graphResult(undefined as undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to delete bucket');
  }
}

export async function deletePlannerTask(token: string, taskId: string, etag: string): Promise<GraphResponse<void>> {
  try {
    const result = await callGraph<void>(
      token,
      `/planner/tasks/${encodeURIComponent(taskId)}`,
      {
        method: 'DELETE',
        headers: { 'If-Match': etag }
      },
      false
    );
    if (!result.ok) {
      return graphError(result.error?.message || 'Failed to delete task', result.error?.code, result.error?.status);
    }
    return graphResult(undefined as undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to delete task');
  }
}

export async function getPlannerTaskDetails(token: string, taskId: string): Promise<GraphResponse<PlannerTaskDetails>> {
  try {
    const result = await callGraph<PlannerTaskDetails>(token, `/planner/tasks/${encodeURIComponent(taskId)}/details`);
    if (!result.ok || !result.data) {
      return graphError(
        result.error?.message || 'Failed to get task details',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get task details');
  }
}

export interface UpdatePlannerTaskDetailsParams {
  description?: string | null;
  checklist?: Record<string, PlannerTaskDetailsChecklistItem> | null;
  references?: Record<string, unknown> | null;
  previewType?: string;
}

export async function updatePlannerTaskDetails(
  token: string,
  taskDetailsId: string,
  etag: string,
  updates: UpdatePlannerTaskDetailsParams
): Promise<GraphResponse<void>> {
  try {
    const result = await callGraph<void>(token, `/planner/taskDetails/${encodeURIComponent(taskDetailsId)}`, {
      method: 'PATCH',
      headers: { 'If-Match': etag },
      body: JSON.stringify(updates)
    });
    if (!result.ok) {
      return graphError(
        result.error?.message || 'Failed to update task details',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(undefined as undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to update task details');
  }
}

export interface UpdatePlannerPlanDetailsParams {
  categoryDescriptions?: Partial<Record<PlannerCategorySlot, string | null>>;
  sharedWith?: Record<string, boolean>;
}

/** PATCH `/planner/plans/{id}/details` (label names, sharedWith). */
export async function updatePlannerPlanDetails(
  token: string,
  planId: string,
  etag: string,
  updates: UpdatePlannerPlanDetailsParams
): Promise<GraphResponse<void>> {
  try {
    const result = await callGraph<void>(
      token,
      `/planner/plans/${encodeURIComponent(planId)}/details`,
      {
        method: 'PATCH',
        headers: { 'If-Match': etag },
        body: JSON.stringify(updates)
      },
      false
    );
    if (!result.ok) {
      return graphError(
        result.error?.message || 'Failed to update plan details',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(undefined as undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to update plan details');
  }
}

/** Beta: plans marked favorite by the current user. */
export async function listFavoritePlans(token: string): Promise<GraphResponse<PlannerPlan[]>> {
  return fetchAllPages<PlannerPlan>(
    token,
    '/me/planner/favoritePlans',
    'Failed to list favorite plans',
    GRAPH_BETA_URL
  );
}

/** Beta: plans from rosters the user belongs to. */
export async function listRosterPlans(token: string): Promise<GraphResponse<PlannerPlan[]>> {
  return fetchAllPages<PlannerPlan>(token, '/me/planner/rosterPlans', 'Failed to list roster plans', GRAPH_BETA_URL);
}

export interface PlannerDeltaPage {
  value?: unknown[];
  '@odata.nextLink'?: string;
  '@odata.deltaLink'?: string;
}

/** Beta: one page of `/me/planner/all/delta` (pass `nextLink` or `deltaLink` URL from a prior response). */
export async function getPlannerDeltaPage(
  token: string,
  nextOrDeltaUrl?: string
): Promise<GraphResponse<PlannerDeltaPage>> {
  try {
    if (nextOrDeltaUrl) {
      return await callGraphAbsolute<PlannerDeltaPage>(token, nextOrDeltaUrl);
    }
    return await callGraphAt<PlannerDeltaPage>(GRAPH_BETA_URL, token, '/me/planner/all/delta');
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get Planner delta');
  }
}

/** Add a checklist row (client-generated id) and PATCH task details. */
export async function addPlannerChecklistItem(
  token: string,
  taskId: string,
  title: string,
  checklistItemId?: string
): Promise<GraphResponse<void>> {
  const id = checklistItemId || randomUUID();
  const dr = await getPlannerTaskDetails(token, taskId);
  if (!dr.ok || !dr.data) {
    return graphError(dr.error?.message || 'Failed to get task details', dr.error?.code, dr.error?.status);
  }
  const etag = dr.data['@odata.etag'];
  if (!etag) return graphError('Task details missing ETag', 'MISSING_ETAG', 500);
  const checklist = { ...(dr.data.checklist || {}) };
  checklist[id] = {
    '@odata.type': '#microsoft.graph.plannerChecklistItem',
    isChecked: false,
    title,
    orderHint: ' !'
  };
  return updatePlannerTaskDetails(token, dr.data.id, etag, { checklist });
}

/** Remove a checklist item by id. */
export async function removePlannerChecklistItem(
  token: string,
  taskId: string,
  checklistItemId: string
): Promise<GraphResponse<void>> {
  const dr = await getPlannerTaskDetails(token, taskId);
  if (!dr.ok || !dr.data) {
    return graphError(dr.error?.message || 'Failed to get task details', dr.error?.code, dr.error?.status);
  }
  const etag = dr.data['@odata.etag'];
  if (!etag) return graphError('Task details missing ETag', 'MISSING_ETAG', 500);
  const checklist = { ...(dr.data.checklist || {}) };
  delete checklist[checklistItemId];
  return updatePlannerTaskDetails(token, dr.data.id, etag, { checklist });
}

/** Add or replace a reference URL entry on task details. */
export async function addPlannerReference(
  token: string,
  taskId: string,
  resourceUrl: string,
  alias: string,
  type?: string
): Promise<GraphResponse<void>> {
  const dr = await getPlannerTaskDetails(token, taskId);
  if (!dr.ok || !dr.data) {
    return graphError(dr.error?.message || 'Failed to get task details', dr.error?.code, dr.error?.status);
  }
  const etag = dr.data['@odata.etag'];
  if (!etag) return graphError('Task details missing ETag', 'MISSING_ETAG', 500);
  const references = { ...(dr.data.references || {}) };
  references[resourceUrl] = {
    '@odata.type': '#microsoft.graph.plannerExternalReference',
    alias,
    ...(type ? { type } : {}),
    previewPriority: ' !'
  };
  return updatePlannerTaskDetails(token, dr.data.id, etag, { references });
}

/** Remove a reference by key URL. */
export async function removePlannerReference(
  token: string,
  taskId: string,
  resourceUrl: string
): Promise<GraphResponse<void>> {
  const dr = await getPlannerTaskDetails(token, taskId);
  if (!dr.ok || !dr.data) {
    return graphError(dr.error?.message || 'Failed to get task details', dr.error?.code, dr.error?.status);
  }
  const etag = dr.data['@odata.etag'];
  if (!etag) return graphError('Task details missing ETag', 'MISSING_ETAG', 500);
  const references = { ...(dr.data.references || {}) };
  delete references[resourceUrl];
  return updatePlannerTaskDetails(token, dr.data.id, etag, { references });
}

/** Update one checklist row (title, checked state, orderHint). */
export async function updatePlannerChecklistItem(
  token: string,
  taskId: string,
  checklistItemId: string,
  patch: { title?: string; isChecked?: boolean; orderHint?: string }
): Promise<GraphResponse<void>> {
  const dr = await getPlannerTaskDetails(token, taskId);
  if (!dr.ok || !dr.data) {
    return graphError(dr.error?.message || 'Failed to get task details', dr.error?.code, dr.error?.status);
  }
  const etag = dr.data['@odata.etag'];
  if (!etag) return graphError('Task details missing ETag', 'MISSING_ETAG', 500);
  const checklist = { ...(dr.data.checklist || {}) };
  const cur = checklist[checklistItemId];
  if (!cur) {
    return graphError(`Checklist item not found: ${checklistItemId}`, 'NOT_FOUND', 404);
  }
  checklist[checklistItemId] = {
    ...cur,
    '@odata.type': cur['@odata.type'] ?? '#microsoft.graph.plannerChecklistItem',
    ...(patch.title !== undefined ? { title: patch.title } : {}),
    ...(patch.isChecked !== undefined ? { isChecked: patch.isChecked } : {}),
    ...(patch.orderHint !== undefined ? { orderHint: patch.orderHint } : {})
  };
  return updatePlannerTaskDetails(token, dr.data.id, etag, { checklist });
}

export interface PlannerAssignedToTaskBoardFormat {
  id: string;
  unassignedOrderHint?: string;
  orderHintsByAssignee?: Record<string, string>;
  '@odata.etag'?: string;
}

export interface PlannerBucketTaskBoardFormat {
  id: string;
  orderHint?: string;
  '@odata.etag'?: string;
}

export interface PlannerProgressTaskBoardFormat {
  id: string;
  orderHint?: string;
  '@odata.etag'?: string;
}

export async function getAssignedToTaskBoardFormat(
  token: string,
  taskId: string
): Promise<GraphResponse<PlannerAssignedToTaskBoardFormat>> {
  try {
    const result = await callGraph<PlannerAssignedToTaskBoardFormat>(
      token,
      `/planner/tasks/${encodeURIComponent(taskId)}/assignedToTaskBoardFormat`
    );
    if (!result.ok || !result.data) {
      return graphError(
        result.error?.message || 'Failed to get assignedTo task board format',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get assignedTo task board format');
  }
}

export async function updateAssignedToTaskBoardFormat(
  token: string,
  taskId: string,
  etag: string,
  updates: Partial<{ orderHintsByAssignee: Record<string, string> | null; unassignedOrderHint: string | null }>
): Promise<GraphResponse<void>> {
  try {
    const result = await callGraph<void>(
      token,
      `/planner/tasks/${encodeURIComponent(taskId)}/assignedToTaskBoardFormat`,
      {
        method: 'PATCH',
        headers: { 'If-Match': etag },
        body: JSON.stringify(updates)
      },
      false
    );
    if (!result.ok) {
      return graphError(
        result.error?.message || 'Failed to update assignedTo task board format',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(undefined as undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to update assignedTo task board format');
  }
}

export async function getBucketTaskBoardFormat(
  token: string,
  taskId: string
): Promise<GraphResponse<PlannerBucketTaskBoardFormat>> {
  try {
    const result = await callGraph<PlannerBucketTaskBoardFormat>(
      token,
      `/planner/tasks/${encodeURIComponent(taskId)}/bucketTaskBoardFormat`
    );
    if (!result.ok || !result.data) {
      return graphError(
        result.error?.message || 'Failed to get bucket task board format',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get bucket task board format');
  }
}

export async function updateBucketTaskBoardFormat(
  token: string,
  taskId: string,
  etag: string,
  orderHint: string
): Promise<GraphResponse<void>> {
  try {
    const result = await callGraph<void>(
      token,
      `/planner/tasks/${encodeURIComponent(taskId)}/bucketTaskBoardFormat`,
      {
        method: 'PATCH',
        headers: { 'If-Match': etag },
        body: JSON.stringify({ orderHint })
      },
      false
    );
    if (!result.ok) {
      return graphError(
        result.error?.message || 'Failed to update bucket task board format',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(undefined as undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to update bucket task board format');
  }
}

export async function getProgressTaskBoardFormat(
  token: string,
  taskId: string
): Promise<GraphResponse<PlannerProgressTaskBoardFormat>> {
  try {
    const result = await callGraph<PlannerProgressTaskBoardFormat>(
      token,
      `/planner/tasks/${encodeURIComponent(taskId)}/progressTaskBoardFormat`
    );
    if (!result.ok || !result.data) {
      return graphError(
        result.error?.message || 'Failed to get progress task board format',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get progress task board format');
  }
}

export async function updateProgressTaskBoardFormat(
  token: string,
  taskId: string,
  etag: string,
  orderHint: string
): Promise<GraphResponse<void>> {
  try {
    const result = await callGraph<void>(
      token,
      `/planner/tasks/${encodeURIComponent(taskId)}/progressTaskBoardFormat`,
      {
        method: 'PATCH',
        headers: { 'If-Match': etag },
        body: JSON.stringify({ orderHint })
      },
      false
    );
    if (!result.ok) {
      return graphError(
        result.error?.message || 'Failed to update progress task board format',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(undefined as undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to update progress task board format');
  }
}

/** Beta: current user's Planner preferences (favorites, recents). */
export interface PlannerUser {
  id: string;
  '@odata.etag'?: string;
  favoritePlanReferences?: Record<string, unknown>;
  recentPlanReferences?: Record<string, unknown>;
}

export async function getPlannerUser(token: string): Promise<GraphResponse<PlannerUser>> {
  try {
    const result = await callGraphAt<PlannerUser>(GRAPH_BETA_URL, token, '/me/planner');
    if (!result.ok || !result.data) {
      return graphError(
        result.error?.message || 'Failed to get planner user',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get planner user');
  }
}

async function patchPlannerUser(
  token: string,
  etag: string,
  body: {
    favoritePlanReferences?: Record<string, unknown> | null;
    recentPlanReferences?: Record<string, unknown> | null;
  }
): Promise<GraphResponse<PlannerUser | undefined>> {
  try {
    const result = await callGraphAt<PlannerUser>(
      GRAPH_BETA_URL,
      token,
      '/me/planner',
      {
        method: 'PATCH',
        headers: {
          'If-Match': etag,
          Prefer: 'return=representation'
        },
        body: JSON.stringify(body)
      },
      true
    );
    if (!result.ok) {
      return graphError(
        result.error?.message || 'Failed to update planner user',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to update planner user');
  }
}

/** Beta: add or update a favorite plan entry (PATCH /me/planner merge). */
export async function addPlannerFavoritePlan(
  token: string,
  planId: string,
  planTitle: string
): Promise<GraphResponse<PlannerUser | undefined>> {
  const ur = await getPlannerUser(token);
  if (!ur.ok || !ur.data) {
    return graphError(ur.error?.message || 'Failed to get planner user', ur.error?.code, ur.error?.status);
  }
  const etag = ur.data['@odata.etag'];
  if (!etag) return graphError('plannerUser missing ETag', 'MISSING_ETAG', 500);
  return patchPlannerUser(token, etag, {
    favoritePlanReferences: {
      [planId]: {
        '@odata.type': '#microsoft.graph.plannerFavoritePlanReference',
        orderHint: ' !',
        planTitle
      }
    }
  });
}

/** Beta: remove a plan from favorites (set reference to null). */
export async function removePlannerFavoritePlan(
  token: string,
  planId: string
): Promise<GraphResponse<PlannerUser | undefined>> {
  const ur = await getPlannerUser(token);
  if (!ur.ok || !ur.data) {
    return graphError(ur.error?.message || 'Failed to get planner user', ur.error?.code, ur.error?.status);
  }
  const etag = ur.data['@odata.etag'];
  if (!etag) return graphError('plannerUser missing ETag', 'MISSING_ETAG', 500);
  const favoritePlanReferences: Record<string, unknown> = { [planId]: null };
  return patchPlannerUser(token, etag, { favoritePlanReferences });
}

/** Beta: security container for roster-backed plans (see Graph `plannerRoster`). */
export interface PlannerRoster {
  id: string;
  '@odata.type'?: string;
}

/** Beta: one member of a Planner roster. */
export interface PlannerRosterMember {
  id: string;
  userId: string;
  roles?: string[];
  '@odata.type'?: string;
}

/** Beta: `POST /planner/rosters` — create an empty roster (then add members and add a plan). */
export async function createPlannerRoster(token: string): Promise<GraphResponse<PlannerRoster>> {
  try {
    const result = await callGraphAt<PlannerRoster>(GRAPH_BETA_URL, token, '/planner/rosters', {
      method: 'POST',
      body: JSON.stringify({ '@odata.type': '#microsoft.graph.plannerRoster' })
    });
    if (!result.ok || !result.data) {
      return graphError(result.error?.message || 'Failed to create roster', result.error?.code, result.error?.status);
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to create roster');
  }
}

/** Beta: `GET /planner/rosters/{id}`. */
export async function getPlannerRoster(token: string, rosterId: string): Promise<GraphResponse<PlannerRoster>> {
  try {
    const result = await callGraphAt<PlannerRoster>(
      GRAPH_BETA_URL,
      token,
      `/planner/rosters/${encodeURIComponent(rosterId)}`
    );
    if (!result.ok || !result.data) {
      return graphError(result.error?.message || 'Failed to get roster', result.error?.code, result.error?.status);
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get roster');
  }
}

/** Beta: `GET /planner/rosters/{id}/members`. */
export async function listPlannerRosterMembers(
  token: string,
  rosterId: string
): Promise<GraphResponse<PlannerRosterMember[]>> {
  return fetchAllPages<PlannerRosterMember>(
    token,
    `/planner/rosters/${encodeURIComponent(rosterId)}/members`,
    'Failed to list roster members',
    GRAPH_BETA_URL
  );
}

/** Beta: `POST /planner/rosters/{id}/members`. */
export async function addPlannerRosterMember(
  token: string,
  rosterId: string,
  userId: string,
  options?: { tenantId?: string; roles?: string[] }
): Promise<GraphResponse<PlannerRosterMember>> {
  try {
    const body: Record<string, unknown> = {
      '@odata.type': '#microsoft.graph.plannerRosterMember',
      userId
    };
    if (options?.tenantId !== undefined) body.tenantId = options.tenantId;
    if (options?.roles !== undefined && options.roles.length > 0) body.roles = options.roles;
    const result = await callGraphAt<PlannerRosterMember>(
      GRAPH_BETA_URL,
      token,
      `/planner/rosters/${encodeURIComponent(rosterId)}/members`,
      {
        method: 'POST',
        body: JSON.stringify(body)
      }
    );
    if (!result.ok || !result.data) {
      return graphError(
        result.error?.message || 'Failed to add roster member',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to add roster member');
  }
}

/** Beta: `DELETE /planner/rosters/{rosterId}/members/{memberId}` (member id is the roster member resource id). */
export async function removePlannerRosterMember(
  token: string,
  rosterId: string,
  memberId: string
): Promise<GraphResponse<void>> {
  try {
    const result = await callGraphAt<void>(
      GRAPH_BETA_URL,
      token,
      `/planner/rosters/${encodeURIComponent(rosterId)}/members/${encodeURIComponent(memberId)}`,
      { method: 'DELETE' },
      false
    );
    if (!result.ok) {
      return graphError(
        result.error?.message || 'Failed to remove roster member',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(undefined as undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to remove roster member');
  }
}

/**
 * Beta: create a plan contained by a roster (`POST /planner/plans` on beta with `container.type` roster).
 * @see https://learn.microsoft.com/en-us/graph/api/resources/plannerplancontainer
 */
export async function createPlannerPlanInRoster(
  token: string,
  rosterId: string,
  title: string
): Promise<GraphResponse<PlannerPlan>> {
  try {
    const base = GRAPH_BETA_URL.replace(/\/$/, '');
    const result = await callGraphAt<PlannerPlan>(GRAPH_BETA_URL, token, '/planner/plans', {
      method: 'POST',
      headers: { Prefer: 'include-unknown-enum-members' },
      body: JSON.stringify({
        title,
        container: {
          '@odata.type': '#microsoft.graph.plannerPlanContainer',
          url: `${base}/planner/rosters/${encodeURIComponent(rosterId)}`,
          type: 'roster'
        }
      })
    });
    if (!result.ok || !result.data) {
      return graphError(
        result.error?.message || 'Failed to create plan in roster',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to create plan in roster');
  }
}
