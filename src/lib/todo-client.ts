import { readFile, stat } from 'node:fs/promises';
import { basename } from 'node:path';
import {
  callGraph,
  callGraphAbsolute,
  fetchAllPages,
  fetchGraphRaw,
  GraphApiError,
  type GraphResponse,
  graphError,
  graphResult
} from './graph-client.js';
import { graphUserPath } from './graph-user-path.js';

function todoRoot(user?: string): string {
  return graphUserPath(user, 'todo');
}

export type TodoImportance = 'low' | 'normal' | 'high';
export type TodoStatus = 'notStarted' | 'inProgress' | 'completed' | 'waitingOnOthers' | 'deferred';

/** Graph [linkedResource](https://learn.microsoft.com/en-us/graph/api/resources/linkedresource); use `displayName` (alias: `description`). */
export interface TodoLinkedResource {
  id?: string;
  webUrl?: string;
  /** Graph `displayName` (title of the link). */
  displayName?: string;
  /** Legacy alias for `displayName` when creating/updating. */
  description?: string;
  applicationName?: string;
  externalId?: string;
  iconUrl?: string;
}

/** Shape payload for Graph `linkedResources` on todoTask (PATCH/POST). */
export function linkedResourceToGraphPayload(lr: TodoLinkedResource): Record<string, unknown> {
  const displayName = lr.displayName ?? lr.description;
  const out: Record<string, unknown> = {};
  if (displayName !== undefined && displayName !== '') out.displayName = displayName;
  if (lr.webUrl !== undefined) out.webUrl = lr.webUrl;
  if (lr.applicationName !== undefined) out.applicationName = lr.applicationName;
  if (lr.externalId !== undefined) out.externalId = lr.externalId;
  if (lr.id !== undefined) out.id = lr.id;
  return out;
}

export interface TodoChecklistItem {
  id: string;
  displayName: string;
  isChecked: boolean;
  createdDateTime?: string;
  /** Set when `isChecked` is true (Graph). */
  checkedDateTime?: string;
}

export interface TodoTask {
  id: string;
  title: string;
  body?: { content: string; contentType: string };
  isReminderOn?: boolean;
  reminderDateTime?: { dateTime: string; timeZone: string };
  dueDateTime?: { dateTime: string; timeZone: string };
  startDateTime?: { dateTime: string; timeZone: string };
  importance?: TodoImportance;
  status?: TodoStatus;
  /** Outlook-style category labels (strings). */
  categories?: string[];
  linkedResources?: TodoLinkedResource[];
  checklistItems?: TodoChecklistItem[];
  /** Graph `patternedRecurrence` resource (opaque JSON). */
  recurrence?: Record<string, unknown>;
  hasAttachments?: boolean;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  completedDateTime?: { dateTime: string; timeZone: string };
  bodyLastModifiedDateTime?: string;
}

/** Small file attachment on a To Do task (Graph `taskFileAttachment`). */
export interface TodoAttachment {
  id: string;
  name?: string;
  contentType?: string;
  size?: number;
  lastModifiedDateTime?: string;
  '@odata.type'?: string;
}

export interface TodoList {
  id: string;
  displayName: string;
  isOwner?: boolean;
  isShared?: boolean;
  parentSectionId?: string;
  wellknownListName?: string;
}

export async function createTodoList(
  token: string,
  displayName: string,
  user?: string
): Promise<GraphResponse<TodoList>> {
  let result: GraphResponse<TodoList>;
  try {
    result = await callGraph<TodoList>(token, `${todoRoot(user)}/lists`, {
      method: 'POST',
      body: JSON.stringify({ displayName })
    });
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to create list');
  }
  if (!result.ok || !result.data) {
    return graphError(result.error?.message || 'Failed to create list', result.error?.code, result.error?.status);
  }
  return graphResult(result.data);
}

export async function updateTodoList(
  token: string,
  listId: string,
  displayName: string,
  user?: string
): Promise<GraphResponse<TodoList>> {
  let result: GraphResponse<TodoList>;
  try {
    result = await callGraph<TodoList>(token, `${todoRoot(user)}/lists/${encodeURIComponent(listId)}`, {
      method: 'PATCH',
      body: JSON.stringify({ displayName })
    });
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to update list');
  }
  if (!result.ok || !result.data) {
    return graphError(result.error?.message || 'Failed to update list', result.error?.code, result.error?.status);
  }
  return graphResult(result.data);
}

export async function deleteTodoList(token: string, listId: string, user?: string): Promise<GraphResponse<void>> {
  try {
    return await callGraph<void>(
      token,
      `${todoRoot(user)}/lists/${encodeURIComponent(listId)}`,
      { method: 'DELETE' },
      false
    );
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to delete list');
  }
}

export async function getTodoLists(token: string, user?: string): Promise<GraphResponse<TodoList[]>> {
  let result: GraphResponse<{ value: TodoList[] }>;
  try {
    result = await callGraph<{ value: TodoList[] }>(token, `${todoRoot(user)}/lists`);
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to get todo lists');
  }
  if (!result.ok || !result.data) {
    return graphError(result.error?.message || 'Failed to get todo lists', result.error?.code, result.error?.status);
  }
  return graphResult(result.data.value);
}

export async function getTodoList(token: string, listId: string, user?: string): Promise<GraphResponse<TodoList>> {
  try {
    return await callGraph<TodoList>(token, `${todoRoot(user)}/lists/${encodeURIComponent(listId)}`);
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to get todo list');
  }
}

export interface TodoTasksQueryOptions {
  filter?: string;
  orderby?: string;
  select?: string;
  /** When set, only one page is returned (no automatic paging). */
  top?: number;
  skip?: number;
  /** OData `$expand` (e.g. `attachments`). */
  expand?: string;
  /** Set `$count=true` (may require `ConsistencyLevel: eventual` on some tenants). */
  count?: boolean;
}

function tasksListPath(listId: string, user: string | undefined, query?: string | TodoTasksQueryOptions): string {
  const params = new URLSearchParams();
  if (query === undefined) {
    // no query params
  } else if (typeof query === 'string') {
    if (query) params.set('$filter', query);
  } else {
    if (query.filter) params.set('$filter', query.filter);
    if (query.orderby) params.set('$orderby', query.orderby);
    if (query.select) params.set('$select', query.select);
    if (query.top !== undefined) params.set('$top', String(query.top));
    if (query.skip !== undefined) params.set('$skip', String(query.skip));
    if (query.expand) params.set('$expand', query.expand);
    if (query.count) params.set('$count', 'true');
  }
  const qs = params.toString() ? `?${params.toString()}` : '';
  return `${todoRoot(user)}/lists/${encodeURIComponent(listId)}/tasks${qs}`;
}

export async function getTasks(
  token: string,
  listId: string,
  filterOrQuery?: string | TodoTasksQueryOptions,
  user?: string
): Promise<GraphResponse<TodoTask[]>> {
  const path = tasksListPath(listId, user, filterOrQuery);
  const singlePage =
    filterOrQuery !== undefined &&
    typeof filterOrQuery === 'object' &&
    (filterOrQuery.top !== undefined || filterOrQuery.skip !== undefined || filterOrQuery.count === true);

  if (singlePage) {
    let result: GraphResponse<{ value: TodoTask[] }>;
    try {
      result = await callGraph<{ value: TodoTask[] }>(token, path);
    } catch (err) {
      if (err instanceof GraphApiError) {
        return graphError(err.message, err.code, err.status);
      }
      return graphError(err instanceof Error ? err.message : 'Failed to get tasks');
    }
    if (!result.ok || !result.data) {
      return graphError(result.error?.message || 'Failed to get tasks', result.error?.code, result.error?.status);
    }
    return graphResult(result.data.value);
  }

  return fetchAllPages<TodoTask>(token, path, 'Failed to get tasks');
}

export async function getTask(
  token: string,
  listId: string,
  taskId: string,
  user?: string,
  options?: { select?: string }
): Promise<GraphResponse<TodoTask>> {
  try {
    const qs = options?.select ? `?$select=${encodeURIComponent(options.select)}` : '';
    return await callGraph<TodoTask>(
      token,
      `${todoRoot(user)}/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}${qs}`
    );
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to get task');
  }
}

export interface CreateTaskOptions {
  title: string;
  body?: string;
  bodyContentType?: string;
  dueDateTime?: string;
  startDateTime?: string;
  timeZone?: string;
  dueTimeZone?: string;
  startTimeZone?: string;
  reminderTimeZone?: string;
  importance?: TodoImportance;
  status?: TodoStatus;
  isReminderOn?: boolean;
  reminderDateTime?: string;
  linkedResources?: TodoLinkedResource[];
  categories?: string[];
  /** Graph `patternedRecurrence` (see Microsoft Graph docs). */
  recurrence?: Record<string, unknown>;
}

export async function createTask(
  token: string,
  listId: string,
  options: CreateTaskOptions,
  user?: string
): Promise<GraphResponse<TodoTask>> {
  const payload: Record<string, unknown> = { title: options.title };
  if (options.body) payload.body = { content: options.body, contentType: options.bodyContentType || 'text' };
  if (options.dueDateTime) {
    payload.dueDateTime = {
      dateTime: options.dueDateTime,
      timeZone: options.dueTimeZone ?? options.timeZone ?? 'UTC'
    };
  }
  if (options.startDateTime) {
    payload.startDateTime = {
      dateTime: options.startDateTime,
      timeZone: options.startTimeZone ?? options.timeZone ?? 'UTC'
    };
  }
  if (options.importance) payload.importance = options.importance;
  if (options.status) payload.status = options.status;
  if (options.isReminderOn !== undefined) payload.isReminderOn = options.isReminderOn;
  if (options.reminderDateTime) {
    payload.reminderDateTime = {
      dateTime: options.reminderDateTime,
      timeZone: options.reminderTimeZone ?? options.timeZone ?? 'UTC'
    };
  }
  if (options.linkedResources?.length) {
    payload.linkedResources = options.linkedResources.map((lr) => linkedResourceToGraphPayload(lr));
  }
  if (options.categories?.length) payload.categories = options.categories;
  if (options.recurrence !== undefined) payload.recurrence = options.recurrence;
  let result: GraphResponse<TodoTask>;
  try {
    result = await callGraph<TodoTask>(token, `${todoRoot(user)}/lists/${encodeURIComponent(listId)}/tasks`, {
      method: 'POST',
      body: JSON.stringify(payload)
    });
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to create task');
  }
  if (!result.ok || !result.data) {
    return graphError(result.error?.message || 'Failed to create task', result.error?.code, result.error?.status);
  }
  return graphResult(result.data);
}

export interface UpdateTaskOptions {
  title?: string;
  body?: string;
  bodyContentType?: string;
  dueDateTime?: string | null;
  startDateTime?: string | null;
  timeZone?: string;
  dueTimeZone?: string;
  startTimeZone?: string;
  reminderTimeZone?: string;
  importance?: TodoImportance;
  status?: TodoStatus;
  isReminderOn?: boolean;
  reminderDateTime?: string | null;
  completedDateTime?: string | null;
  linkedResources?: TodoLinkedResource[];
  /** Replace categories when set (including empty array). */
  categories?: string[];
  /** When true, PATCH with categories: []. Ignored if categories is set. */
  clearCategories?: boolean;
  /** Set or clear recurrence; `null` removes recurrence from the task. */
  recurrence?: Record<string, unknown> | null;
}

export async function updateTask(
  token: string,
  listId: string,
  taskId: string,
  options: UpdateTaskOptions,
  user?: string
): Promise<GraphResponse<TodoTask>> {
  const payload: Record<string, unknown> = {};
  if (options.title !== undefined) payload.title = options.title;
  if (options.body !== undefined)
    payload.body = { content: options.body, contentType: options.bodyContentType || 'text' };
  if (options.dueDateTime !== undefined) {
    payload.dueDateTime =
      options.dueDateTime === null
        ? null
        : { dateTime: options.dueDateTime, timeZone: options.dueTimeZone ?? options.timeZone ?? 'UTC' };
  }
  if (options.startDateTime !== undefined) {
    payload.startDateTime =
      options.startDateTime === null
        ? null
        : { dateTime: options.startDateTime, timeZone: options.startTimeZone ?? options.timeZone ?? 'UTC' };
  }
  if (options.importance !== undefined) payload.importance = options.importance;
  if (options.status !== undefined) payload.status = options.status;
  if (options.isReminderOn !== undefined) payload.isReminderOn = options.isReminderOn;
  if (options.reminderDateTime !== undefined) {
    payload.reminderDateTime =
      options.reminderDateTime === null
        ? null
        : { dateTime: options.reminderDateTime, timeZone: options.reminderTimeZone ?? options.timeZone ?? 'UTC' };
  }
  if (options.completedDateTime !== undefined) {
    payload.completedDateTime =
      options.completedDateTime === null
        ? null
        : { dateTime: options.completedDateTime, timeZone: options.timeZone || 'UTC' };
  }
  if (options.linkedResources !== undefined) {
    payload.linkedResources = options.linkedResources.map((lr) => linkedResourceToGraphPayload(lr));
  }
  if (options.categories !== undefined) payload.categories = options.categories;
  else if (options.clearCategories) payload.categories = [];
  if (options.recurrence !== undefined) payload.recurrence = options.recurrence;
  let result: GraphResponse<TodoTask>;
  try {
    result = await callGraph<TodoTask>(
      token,
      `${todoRoot(user)}/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}`,
      { method: 'PATCH', body: JSON.stringify(payload) }
    );
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to update task');
  }
  if (!result.ok || !result.data) {
    return graphError(result.error?.message || 'Failed to update task', result.error?.code, result.error?.status);
  }
  return graphResult(result.data);
}

export async function deleteTask(
  token: string,
  listId: string,
  taskId: string,
  user?: string
): Promise<GraphResponse<void>> {
  try {
    return await callGraph<void>(
      token,
      `${todoRoot(user)}/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}`,
      { method: 'DELETE' },
      false
    );
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to delete task');
  }
}

export async function addChecklistItem(
  token: string,
  listId: string,
  taskId: string,
  displayName: string,
  user?: string
): Promise<GraphResponse<TodoChecklistItem>> {
  let result: GraphResponse<TodoChecklistItem>;
  try {
    result = await callGraph<TodoChecklistItem>(
      token,
      `${todoRoot(user)}/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}/checklistItems`,
      { method: 'POST', body: JSON.stringify({ displayName }) }
    );
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to add checklist item');
  }
  if (!result.ok || !result.data) {
    return graphError(
      result.error?.message || 'Failed to add checklist item',
      result.error?.code,
      result.error?.status
    );
  }
  return graphResult(result.data);
}

export async function deleteChecklistItem(
  token: string,
  listId: string,
  taskId: string,
  checklistItemId: string,
  user?: string
): Promise<GraphResponse<void>> {
  try {
    return await callGraph<void>(
      token,
      `${todoRoot(user)}/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}/checklistItems/${encodeURIComponent(checklistItemId)}`,
      { method: 'DELETE' },
      false
    );
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to delete checklist item');
  }
}

export async function updateChecklistItem(
  token: string,
  listId: string,
  taskId: string,
  checklistItemId: string,
  patch: { displayName?: string; isChecked?: boolean },
  user?: string
): Promise<GraphResponse<TodoChecklistItem>> {
  let result: GraphResponse<TodoChecklistItem>;
  try {
    result = await callGraph<TodoChecklistItem>(
      token,
      `${todoRoot(user)}/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}/checklistItems/${encodeURIComponent(checklistItemId)}`,
      { method: 'PATCH', body: JSON.stringify(patch) }
    );
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to update checklist item');
  }
  if (!result.ok || !result.data) {
    return graphError(
      result.error?.message || 'Failed to update checklist item',
      result.error?.code,
      result.error?.status
    );
  }
  return graphResult(result.data);
}

export async function listAttachments(
  token: string,
  listId: string,
  taskId: string,
  user?: string
): Promise<GraphResponse<TodoAttachment[]>> {
  return fetchAllPages<TodoAttachment>(
    token,
    `${todoRoot(user)}/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}/attachments`,
    'Failed to list attachments'
  );
}

export async function createTaskFileAttachment(
  token: string,
  listId: string,
  taskId: string,
  name: string,
  contentBytesBase64: string,
  contentType: string,
  user?: string
): Promise<GraphResponse<TodoAttachment>> {
  const body = {
    '@odata.type': '#microsoft.graph.taskFileAttachment',
    name,
    contentBytes: contentBytesBase64,
    contentType
  };
  let result: GraphResponse<TodoAttachment>;
  try {
    result = await callGraph<TodoAttachment>(
      token,
      `${todoRoot(user)}/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}/attachments`,
      { method: 'POST', body: JSON.stringify(body) }
    );
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to add attachment');
  }
  if (!result.ok || !result.data) {
    return graphError(result.error?.message || 'Failed to add attachment', result.error?.code, result.error?.status);
  }
  return graphResult(result.data);
}

export async function deleteAttachment(
  token: string,
  listId: string,
  taskId: string,
  attachmentId: string,
  user?: string
): Promise<GraphResponse<void>> {
  try {
    return await callGraph<void>(
      token,
      `${todoRoot(user)}/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}/attachments/${encodeURIComponent(attachmentId)}`,
      { method: 'DELETE' },
      false
    );
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to delete attachment');
  }
}

export async function getTaskAttachment(
  token: string,
  listId: string,
  taskId: string,
  attachmentId: string,
  user?: string
): Promise<GraphResponse<TodoAttachment>> {
  try {
    const result = await callGraph<TodoAttachment>(
      token,
      `${todoRoot(user)}/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}/attachments/${encodeURIComponent(attachmentId)}`
    );
    if (!result.ok || !result.data) {
      return graphError(result.error?.message || 'Failed to get attachment', result.error?.code, result.error?.status);
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get attachment');
  }
}

/** Link attachment (URL reference), not file bytes. */
export async function createTaskReferenceAttachment(
  token: string,
  listId: string,
  taskId: string,
  name: string,
  url: string,
  user?: string
): Promise<GraphResponse<TodoAttachment>> {
  const body = {
    '@odata.type': '#microsoft.graph.taskReferenceAttachment',
    name,
    url
  };
  let result: GraphResponse<TodoAttachment>;
  try {
    result = await callGraph<TodoAttachment>(
      token,
      `${todoRoot(user)}/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}/attachments`,
      { method: 'POST', body: JSON.stringify(body) }
    );
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to add reference attachment');
  }
  if (!result.ok || !result.data) {
    return graphError(
      result.error?.message || 'Failed to add reference attachment',
      result.error?.code,
      result.error?.status
    );
  }
  return graphResult(result.data);
}

export async function addLinkedResource(
  token: string,
  listId: string,
  taskId: string,
  resource: TodoLinkedResource,
  user?: string
): Promise<GraphResponse<TodoTask>> {
  const tr = await getTask(token, listId, taskId, user);
  if (!tr.ok || !tr.data) {
    return graphError(tr.error?.message || 'Failed to get task', tr.error?.code, tr.error?.status);
  }
  const existing = (tr.data.linkedResources || []) as TodoLinkedResource[];
  const merged = [...existing, resource];
  return updateTask(token, listId, taskId, { linkedResources: merged }, user);
}

export async function removeLinkedResourceByWebUrl(
  token: string,
  listId: string,
  taskId: string,
  webUrl: string,
  user?: string
): Promise<GraphResponse<TodoTask>> {
  const tr = await getTask(token, listId, taskId, user);
  if (!tr.ok || !tr.data) {
    return graphError(tr.error?.message || 'Failed to get task', tr.error?.code, tr.error?.status);
  }
  const merged = (tr.data.linkedResources || []).filter((r) => r.webUrl !== webUrl);
  return updateTask(token, listId, taskId, { linkedResources: merged as TodoLinkedResource[] }, user);
}

function linkedResourcesCollectionPath(
  listId: string,
  taskId: string,
  user: string | undefined,
  linkedResourceId?: string
): string {
  const b = `${todoRoot(user)}/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}/linkedResources`;
  return linkedResourceId ? `${b}/${encodeURIComponent(linkedResourceId)}` : b;
}

/** List linked resources via the task navigation (same data as `linkedResources` on todoTask; supports paging). */
export async function listTaskLinkedResources(
  token: string,
  listId: string,
  taskId: string,
  user?: string
): Promise<GraphResponse<TodoLinkedResource[]>> {
  return fetchAllPages<TodoLinkedResource>(
    token,
    linkedResourcesCollectionPath(listId, taskId, user),
    'Failed to list linked resources'
  );
}

export async function createTaskLinkedResource(
  token: string,
  listId: string,
  taskId: string,
  resource: TodoLinkedResource,
  user?: string
): Promise<GraphResponse<TodoLinkedResource>> {
  let result: GraphResponse<TodoLinkedResource>;
  try {
    result = await callGraph<TodoLinkedResource>(token, linkedResourcesCollectionPath(listId, taskId, user), {
      method: 'POST',
      body: JSON.stringify(linkedResourceToGraphPayload(resource))
    });
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to create linked resource');
  }
  if (!result.ok || !result.data) {
    return graphError(
      result.error?.message || 'Failed to create linked resource',
      result.error?.code,
      result.error?.status
    );
  }
  return graphResult(result.data);
}

export async function getTaskLinkedResource(
  token: string,
  listId: string,
  taskId: string,
  linkedResourceId: string,
  user?: string
): Promise<GraphResponse<TodoLinkedResource>> {
  try {
    const result = await callGraph<TodoLinkedResource>(
      token,
      linkedResourcesCollectionPath(listId, taskId, user, linkedResourceId)
    );
    if (!result.ok || !result.data) {
      return graphError(
        result.error?.message || 'Failed to get linked resource',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get linked resource');
  }
}

export async function updateTaskLinkedResource(
  token: string,
  listId: string,
  taskId: string,
  linkedResourceId: string,
  patch: Partial<Pick<TodoLinkedResource, 'webUrl' | 'displayName' | 'description' | 'applicationName' | 'externalId'>>,
  user?: string
): Promise<GraphResponse<TodoLinkedResource>> {
  const body: Record<string, unknown> = {};
  if (patch.webUrl !== undefined) body.webUrl = patch.webUrl;
  if (patch.applicationName !== undefined) body.applicationName = patch.applicationName;
  if (patch.externalId !== undefined) body.externalId = patch.externalId;
  const displayName = patch.displayName ?? patch.description;
  if (displayName !== undefined) body.displayName = displayName;
  body['@odata.type'] = '#microsoft.graph.linkedResource';
  let result: GraphResponse<TodoLinkedResource>;
  try {
    result = await callGraph<TodoLinkedResource>(
      token,
      linkedResourcesCollectionPath(listId, taskId, user, linkedResourceId),
      { method: 'PATCH', body: JSON.stringify(body) }
    );
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to update linked resource');
  }
  if (!result.ok || !result.data) {
    return graphError(
      result.error?.message || 'Failed to update linked resource',
      result.error?.code,
      result.error?.status
    );
  }
  return graphResult(result.data);
}

export async function deleteTaskLinkedResource(
  token: string,
  listId: string,
  taskId: string,
  linkedResourceId: string,
  user?: string
): Promise<GraphResponse<void>> {
  try {
    return await callGraph<void>(
      token,
      linkedResourcesCollectionPath(listId, taskId, user, linkedResourceId),
      { method: 'DELETE' },
      false
    );
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to delete linked resource');
  }
}

export async function listTaskChecklistItems(
  token: string,
  listId: string,
  taskId: string,
  user?: string
): Promise<GraphResponse<TodoChecklistItem[]>> {
  return fetchAllPages<TodoChecklistItem>(
    token,
    `${todoRoot(user)}/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}/checklistItems`,
    'Failed to list checklist items'
  );
}

/** `GET .../tasks/{taskId}/checklistItems/{checklistItemId}` (see Graph checklistItem). */
export async function getChecklistItem(
  token: string,
  listId: string,
  taskId: string,
  checklistItemId: string,
  user?: string
): Promise<GraphResponse<TodoChecklistItem>> {
  try {
    const result = await callGraph<TodoChecklistItem>(
      token,
      `${todoRoot(user)}/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}/checklistItems/${encodeURIComponent(checklistItemId)}`
    );
    if (!result.ok || !result.data) {
      return graphError(
        result.error?.message || 'Failed to get checklist item',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get checklist item');
  }
}

/**
 * Raw file bytes for a task file attachment (`GET .../attachments/{id}/$value`).
 * Reference attachments do not support this; use metadata from {@link getTaskAttachment} instead.
 */
export async function getTaskAttachmentContent(
  token: string,
  listId: string,
  taskId: string,
  attachmentId: string,
  user?: string
): Promise<GraphResponse<Uint8Array>> {
  const path = `${todoRoot(user)}/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}/attachments/${encodeURIComponent(attachmentId)}/$value`;
  try {
    const res = await fetchGraphRaw(token, path);
    const buf = new Uint8Array(await res.arrayBuffer());
    if (!res.ok) {
      try {
        const text = new TextDecoder().decode(buf);
        const json = JSON.parse(text) as { error?: { code?: string; message?: string } };
        return graphError(json.error?.message || `HTTP ${res.status}`, json.error?.code, res.status);
      } catch {
        return graphError(`Failed to download attachment: HTTP ${res.status}`, undefined, res.status);
      }
    }
    return graphResult(buf);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to download attachment');
  }
}

function listListExtensionsPath(listId: string, user: string | undefined, extensionName?: string): string {
  const base = `${todoRoot(user)}/lists/${encodeURIComponent(listId)}/extensions`;
  return extensionName ? `${base}/${encodeURIComponent(extensionName)}` : base;
}

export async function listTodoListOpenExtensions(
  token: string,
  listId: string,
  user?: string
): Promise<GraphResponse<Array<Record<string, unknown>>>> {
  try {
    const result = await callGraph<{ value: Array<Record<string, unknown>> }>(
      token,
      listListExtensionsPath(listId, user)
    );
    if (!result.ok || !result.data) {
      return graphError(
        result.error?.message || 'Failed to list list extensions',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(result.data.value || []);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to list list extensions');
  }
}

export async function getTodoListOpenExtension(
  token: string,
  listId: string,
  extensionName: string,
  user?: string
): Promise<GraphResponse<Record<string, unknown>>> {
  try {
    const result = await callGraph<Record<string, unknown>>(token, listListExtensionsPath(listId, user, extensionName));
    if (!result.ok || !result.data) {
      return graphError(
        result.error?.message || 'Failed to get list extension',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get list extension');
  }
}

export async function setTodoListOpenExtension(
  token: string,
  listId: string,
  extensionName: string,
  extensionData: Record<string, unknown>,
  user?: string
): Promise<GraphResponse<Record<string, unknown>>> {
  const body = {
    '@odata.type': 'microsoft.graph.openTypeExtension',
    extensionName,
    ...extensionData
  };
  let result: GraphResponse<Record<string, unknown>>;
  try {
    result = await callGraph<Record<string, unknown>>(token, listListExtensionsPath(listId, user), {
      method: 'POST',
      body: JSON.stringify(body)
    });
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to set list extension');
  }
  if (!result.ok || !result.data) {
    return graphError(
      result.error?.message || 'Failed to set list extension',
      result.error?.code,
      result.error?.status
    );
  }
  return graphResult(result.data);
}

export async function updateTodoListOpenExtension(
  token: string,
  listId: string,
  extensionName: string,
  patch: Record<string, unknown>,
  user?: string
): Promise<GraphResponse<void>> {
  try {
    const result = await callGraph<void>(
      token,
      listListExtensionsPath(listId, user, extensionName),
      {
        method: 'PATCH',
        body: JSON.stringify(patch)
      },
      false
    );
    if (!result.ok) {
      return graphError(
        result.error?.message || 'Failed to update list extension',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(undefined as undefined);
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to update list extension');
  }
}

export async function deleteTodoListOpenExtension(
  token: string,
  listId: string,
  extensionName: string,
  user?: string
): Promise<GraphResponse<void>> {
  try {
    return await callGraph<void>(
      token,
      listListExtensionsPath(listId, user, extensionName),
      { method: 'DELETE' },
      false
    );
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to delete list extension');
  }
}

export interface TodoTaskDeltaPage {
  value?: TodoTask[];
  '@odata.nextLink'?: string;
  '@odata.deltaLink'?: string;
}

export async function getTodoTasksDeltaPage(
  token: string,
  listId: string,
  fullUrl?: string,
  user?: string
): Promise<GraphResponse<TodoTaskDeltaPage>> {
  try {
    if (fullUrl) {
      return await callGraphAbsolute<TodoTaskDeltaPage>(token, fullUrl);
    }
    return await callGraph<TodoTaskDeltaPage>(
      token,
      `${todoRoot(user)}/lists/${encodeURIComponent(listId)}/tasks/delta`
    );
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get todo delta');
  }
}

export interface UploadSessionResult {
  uploadUrl: string;
  expirationDateTime?: string;
  nextExpectedRanges?: string[];
}

async function createTaskAttachmentUploadSession(
  token: string,
  listId: string,
  taskId: string,
  attachmentName: string,
  size: number,
  user?: string
): Promise<GraphResponse<UploadSessionResult>> {
  try {
    const result = await callGraph<UploadSessionResult>(
      token,
      `${todoRoot(user)}/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}/attachments/createUploadSession`,
      {
        method: 'POST',
        body: JSON.stringify({
          attachmentInfo: {
            attachmentType: 'file',
            name: attachmentName,
            size
          }
        })
      }
    );
    if (!result.ok || !result.data) {
      return graphError(
        result.error?.message || 'Failed to create upload session',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to create upload session');
  }
}

/**
 * Upload file bytes via session (no Bearer on PUT; Graph upload URL is pre-authorized).
 * Returns the final attachment object from the last chunk response when JSON.
 */
async function uploadFileViaTodoAttachmentSession(
  uploadUrl: string,
  filePath: string,
  chunkSize = 4 * 1024 * 1024
): Promise<GraphResponse<TodoAttachment>> {
  const buf = await readFile(filePath);
  const total = buf.byteLength;
  let start = 0;
  let lastJson: TodoAttachment | undefined;
  while (start < total) {
    const end = Math.min(start + chunkSize, total);
    const slice = buf.subarray(start, end);
    const contentRange = `bytes ${start}-${end - 1}/${total}`;
    let response: Response;
    try {
      // codeql[js/file-access-to-http]: chunked upload of a user-selected attachment file to Graph (pre-authorized uploadUrl).
      response = await fetch(uploadUrl, {
        method: 'PUT',
        headers: {
          'Content-Length': String(slice.byteLength),
          'Content-Range': contentRange
        },
        body: slice
      });
    } catch (err) {
      return graphError(err instanceof Error ? err.message : 'Upload chunk failed');
    }
    const text = await response.text();
    if (!response.ok) {
      return graphError(text || `Upload failed: HTTP ${response.status}`, undefined, response.status);
    }
    if (text) {
      try {
        const parsed = JSON.parse(text) as TodoAttachment & { value?: unknown };
        if (parsed.id) lastJson = parsed;
      } catch {
        // non-JSON success body
      }
    }
    start = end;
  }
  if (lastJson) return graphResult(lastJson);
  return graphError('Upload completed but attachment body was not returned', 'UPLOAD_PARSE', 500);
}

export async function uploadLargeFileAttachment(
  token: string,
  listId: string,
  taskId: string,
  filePath: string,
  attachmentName?: string,
  user?: string
): Promise<GraphResponse<TodoAttachment>> {
  const name = attachmentName?.trim() || basename(filePath);
  const st = await stat(filePath);
  if (!st.isFile()) return graphError(`Not a file: ${filePath}`, 'NOT_FILE', 400);
  const session = await createTaskAttachmentUploadSession(token, listId, taskId, name, st.size, user);
  if (!session.ok || !session.data?.uploadUrl) {
    return graphError(session.error?.message || 'No upload session', session.error?.code, session.error?.status);
  }
  return uploadFileViaTodoAttachmentSession(session.data.uploadUrl, filePath);
}

function extensionsPath(listId: string, taskId: string, user: string | undefined, extensionName?: string): string {
  const base = `${todoRoot(user)}/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}/extensions`;
  return extensionName ? `${base}/${encodeURIComponent(extensionName)}` : base;
}

export async function listTaskOpenExtensions(
  token: string,
  listId: string,
  taskId: string,
  user?: string
): Promise<GraphResponse<Array<Record<string, unknown>>>> {
  try {
    const result = await callGraph<{ value: Array<Record<string, unknown>> }>(
      token,
      extensionsPath(listId, taskId, user)
    );
    if (!result.ok || !result.data) {
      return graphError(result.error?.message || 'Failed to list extensions', result.error?.code, result.error?.status);
    }
    return graphResult(result.data.value || []);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to list extensions');
  }
}

export async function getTaskOpenExtension(
  token: string,
  listId: string,
  taskId: string,
  extensionName: string,
  user?: string
): Promise<GraphResponse<Record<string, unknown>>> {
  try {
    const result = await callGraph<Record<string, unknown>>(token, extensionsPath(listId, taskId, user, extensionName));
    if (!result.ok || !result.data) {
      return graphError(result.error?.message || 'Failed to get extension', result.error?.code, result.error?.status);
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get extension');
  }
}

export async function setTaskOpenExtension(
  token: string,
  listId: string,
  taskId: string,
  extensionName: string,
  extensionData: Record<string, unknown>,
  user?: string
): Promise<GraphResponse<Record<string, unknown>>> {
  const body = {
    '@odata.type': 'microsoft.graph.openTypeExtension',
    extensionName,
    ...extensionData
  };
  let result: GraphResponse<Record<string, unknown>>;
  try {
    result = await callGraph<Record<string, unknown>>(token, extensionsPath(listId, taskId, user), {
      method: 'POST',
      body: JSON.stringify(body)
    });
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to set extension');
  }
  if (!result.ok || !result.data) {
    return graphError(result.error?.message || 'Failed to set extension', result.error?.code, result.error?.status);
  }
  return graphResult(result.data);
}

export async function updateTaskOpenExtension(
  token: string,
  listId: string,
  taskId: string,
  extensionName: string,
  patch: Record<string, unknown>,
  user?: string
): Promise<GraphResponse<void>> {
  try {
    const result = await callGraph<void>(
      token,
      extensionsPath(listId, taskId, user, extensionName),
      {
        method: 'PATCH',
        body: JSON.stringify(patch)
      },
      false
    );
    if (!result.ok) {
      return graphError(
        result.error?.message || 'Failed to update extension',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(undefined as undefined);
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to update extension');
  }
}

export async function deleteTaskOpenExtension(
  token: string,
  listId: string,
  taskId: string,
  extensionName: string,
  user?: string
): Promise<GraphResponse<void>> {
  try {
    return await callGraph<void>(
      token,
      extensionsPath(listId, taskId, user, extensionName),
      { method: 'DELETE' },
      false
    );
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to delete extension');
  }
}
