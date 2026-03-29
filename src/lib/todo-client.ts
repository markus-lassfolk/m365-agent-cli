import { callGraph, GraphApiError, type GraphResponse, graphError, graphResult } from './graph-client.js';

export type TodoImportance = 'low' | 'normal' | 'high';
export type TodoStatus = 'notStarted' | 'inProgress' | 'completed' | 'waitingOnOthers' | 'deferred';

export interface TodoLinkedResource {
  webUrl: string;
  description: string;
  iconUrl?: string;
}

export interface TodoChecklistItem {
  id: string;
  displayName: string;
  isChecked: boolean;
  createdDateTime?: string;
}

export interface TodoTask {
  id: string;
  title: string;
  body?: { content: string; contentType: string };
  isReminderOn?: boolean;
  reminderDateTime?: { dateTime: string; timeZone: string };
  dueDateTime?: { dateTime: string; timeZone: string };
  importance?: TodoImportance;
  status?: TodoStatus;
  linkedResources?: TodoLinkedResource[];
  checklistItems?: TodoChecklistItem[];
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  completedDateTime?: { dateTime: string; timeZone: string };
}

export interface TodoList {
  id: string;
  displayName: string;
  isOwner?: boolean;
  isShared?: boolean;
  parentSectionId?: string;
  wellknownListName?: string;
}

export async function getTodoLists(token: string): Promise<GraphResponse<TodoList[]>> {
  let result: GraphResponse<{ value: TodoList[] }>;
  try {
    result = await callGraph<{ value: TodoList[] }>(token, '/me/todo/lists');
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

export async function getTodoList(token: string, listId: string): Promise<GraphResponse<TodoList>> {
  try {
    return await callGraph<TodoList>(token, `/me/todo/lists/${encodeURIComponent(listId)}`);
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to get todo list');
  }
}

export async function getTasks(token: string, listId: string, filter?: string): Promise<GraphResponse<TodoTask[]>> {
  const params = new URLSearchParams();
  if (filter) params.set('$filter', filter);
  const query = params.toString() ? `?${params.toString()}` : '';
  let result: GraphResponse<{ value: TodoTask[] }>;
  try {
    result = await callGraph<{ value: TodoTask[] }>(
      token,
      `/me/todo/lists/${encodeURIComponent(listId)}/tasks${query}`
    );
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

export async function getTask(token: string, listId: string, taskId: string): Promise<GraphResponse<TodoTask>> {
  try {
    return await callGraph<TodoTask>(
      token,
      `/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}`
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
  timeZone?: string;
  importance?: TodoImportance;
  status?: TodoStatus;
  isReminderOn?: boolean;
  reminderDateTime?: string;
  linkedResources?: TodoLinkedResource[];
}

export async function createTask(
  token: string,
  listId: string,
  options: CreateTaskOptions
): Promise<GraphResponse<TodoTask>> {
  const payload: Record<string, unknown> = { title: options.title };
  if (options.body) payload.body = { content: options.body, contentType: options.bodyContentType || 'text' };
  if (options.dueDateTime) payload.dueDateTime = { dateTime: options.dueDateTime, timeZone: options.timeZone || 'UTC' };
  if (options.importance) payload.importance = options.importance;
  if (options.status) payload.status = options.status;
  if (options.isReminderOn !== undefined) payload.isReminderOn = options.isReminderOn;
  if (options.reminderDateTime) {
    payload.reminderDateTime = { dateTime: options.reminderDateTime, timeZone: options.timeZone || 'UTC' };
  }
  if (options.linkedResources?.length) payload.linkedResources = options.linkedResources;
  let result: GraphResponse<TodoTask>;
  try {
    result = await callGraph<TodoTask>(token, `/me/todo/lists/${encodeURIComponent(listId)}/tasks`, {
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
  timeZone?: string;
  importance?: TodoImportance;
  status?: TodoStatus;
  isReminderOn?: boolean;
  reminderDateTime?: string | null;
  completedDateTime?: string | null;
  linkedResources?: TodoLinkedResource[];
}

export async function updateTask(
  token: string,
  listId: string,
  taskId: string,
  options: UpdateTaskOptions
): Promise<GraphResponse<TodoTask>> {
  const payload: Record<string, unknown> = {};
  if (options.title !== undefined) payload.title = options.title;
  if (options.body !== undefined)
    payload.body = { content: options.body, contentType: options.bodyContentType || 'text' };
  if (options.dueDateTime !== undefined) {
    payload.dueDateTime =
      options.dueDateTime === null ? null : { dateTime: options.dueDateTime, timeZone: options.timeZone || 'UTC' };
  }
  if (options.importance !== undefined) payload.importance = options.importance;
  if (options.status !== undefined) payload.status = options.status;
  if (options.isReminderOn !== undefined) payload.isReminderOn = options.isReminderOn;
  if (options.reminderDateTime !== undefined) {
    payload.reminderDateTime =
      options.reminderDateTime === null
        ? null
        : { dateTime: options.reminderDateTime, timeZone: options.timeZone || 'UTC' };
  }
  if (options.completedDateTime !== undefined) {
    payload.completedDateTime =
      options.completedDateTime === null
        ? null
        : { dateTime: options.completedDateTime, timeZone: options.timeZone || 'UTC' };
  }
  if (options.linkedResources !== undefined) payload.linkedResources = options.linkedResources;
  let result: GraphResponse<TodoTask>;
  try {
    result = await callGraph<TodoTask>(
      token,
      `/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}`,
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

export async function deleteTask(token: string, listId: string, taskId: string): Promise<GraphResponse<void>> {
  try {
    return await callGraph<void>(
      token,
      `/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}`,
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
  displayName: string
): Promise<GraphResponse<TodoChecklistItem>> {
  let result: GraphResponse<TodoChecklistItem>;
  try {
    result = await callGraph<TodoChecklistItem>(
      token,
      `/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}/checklistItems`,
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
  checklistItemId: string
): Promise<GraphResponse<void>> {
  try {
    return await callGraph<void>(
      token,
      `/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}/checklistItems/${encodeURIComponent(checklistItemId)}`,
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
