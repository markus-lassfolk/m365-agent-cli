import {
  callGraph,
  fetchAllPages,
  GraphApiError,
  type GraphResponse,
  graphError,
  graphResult
} from './graph-client.js';

export interface PlannerPlan {
  id: string;
  title: string;
  owner?: string;
}

export interface PlannerBucket {
  id: string;
  name: string;
  planId: string;
  orderHint?: string;
}

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
  dueDateTime?: string;
  assignments?: Record<string, any>;
  '@odata.etag'?: string;
}

export async function listUserTasks(token: string): Promise<GraphResponse<PlannerTask[]>> {
  return fetchAllPages<PlannerTask>(token, '/me/planner/tasks', 'Failed to list tasks');
}

export async function listUserPlans(token: string): Promise<GraphResponse<PlannerPlan[]>> {
  return fetchAllPages<PlannerPlan>(token, '/me/planner/plans', 'Failed to list plans');
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

export async function createTask(
  token: string,
  planId: string,
  title: string,
  bucketId?: string,
  assignments?: Record<string, any>
): Promise<GraphResponse<PlannerTask>> {
  try {
    const body: any = { planId, title };
    if (bucketId) body.bucketId = bucketId;
    if (assignments) body.assignments = assignments;

    const result = await callGraph<PlannerTask>(token, '/planner/tasks', {
      method: 'POST',
      body: JSON.stringify(body)
    });
    if (!result.ok || !result.data) {
      return graphError(result.error?.message || 'Failed to create task', result.error?.code, result.error?.status);
    }
    return graphResult(result.data);
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
    assignments?: Record<string, any>;
    percentComplete?: number;
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
