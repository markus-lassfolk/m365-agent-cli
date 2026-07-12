import {
  callGraph,
  fetchAllPages,
  GraphApiError,
  type GraphResponse,
  graphError,
  graphErrorFromApiError,
  graphResult
} from './graph-client.js';

/** Microsoft 365 (Outlook) group surface. Subset of `microsoft.graph.group`. */
export interface GraphGroup {
  id: string;
  displayName?: string;
  description?: string;
  mail?: string;
  mailNickname?: string;
  groupTypes?: string[];
  visibility?: string;
}

export interface GroupsListResponse {
  value?: GraphGroup[];
  '@odata.nextLink'?: string;
}

/**
 * `GET /me/memberOf/microsoft.graph.group?$filter=groupTypes/any(c:c eq 'Unified')` —
 * lists Microsoft 365 / Outlook groups the signed-in user belongs to. Requires
 * `ConsistencyLevel: eventual` + `$count=true` for advanced queries on directory objects.
 */
export async function listMyOutlookGroups(
  token: string,
  options: { top?: number } = {}
): Promise<GraphResponse<GroupsListResponse>> {
  const topN = options.top && options.top > 0 ? Math.min(Math.max(1, options.top), 200) : undefined;
  const basePath =
    `/me/memberOf/microsoft.graph.group` +
    `?$count=true&$filter=${encodeURIComponent("groupTypes/any(c:c eq 'Unified')")}` +
    `&$select=id,displayName,description,mail,mailNickname,groupTypes,visibility`;
  // Advanced directory query needs the eventual consistency header.
  const headers = { ConsistencyLevel: 'eventual' };
  try {
    // `--top N` is an explicit bound → single page; otherwise page through all groups so a
    // user in >100 unified groups isn't silently cut off at the first page.
    if (topN !== undefined) {
      return await callGraph<GroupsListResponse>(token, `${basePath}&$top=${topN}`, { headers });
    }
    const all = await fetchAllPages<GraphGroup>(token, basePath, 'Failed to list /me/memberOf groups', undefined, {
      headers
    });
    if (!all.ok || !all.data) return all as unknown as GraphResponse<GroupsListResponse>;
    return graphResult({ value: all.data });
  } catch (err) {
    if (err instanceof GraphApiError) return graphErrorFromApiError(err);
    return graphError(err instanceof Error ? err.message : 'Failed to list /me/memberOf groups');
  }
}

export interface GroupConversation {
  id: string;
  topic?: string;
  hasAttachments?: boolean;
  lastDeliveredDateTime?: string;
  uniqueSenders?: string[];
  preview?: string;
}

export interface ConversationsListResponse {
  value?: GroupConversation[];
  '@odata.nextLink'?: string;
}

/** `GET /groups/{id}/conversations` — list group conversations. */
export async function listGroupConversations(
  token: string,
  groupId: string,
  options: { top?: number } = {}
): Promise<GraphResponse<ConversationsListResponse>> {
  const topN = options.top && options.top > 0 ? Math.min(Math.max(1, options.top), 200) : undefined;
  const basePath = `/groups/${encodeURIComponent(groupId)}/conversations`;
  try {
    if (topN !== undefined) {
      return await callGraph<ConversationsListResponse>(token, `${basePath}?$top=${topN}`);
    }
    const all = await fetchAllPages<GroupConversation>(token, basePath, 'Failed to list group conversations');
    if (!all.ok || !all.data) return all as unknown as GraphResponse<ConversationsListResponse>;
    return graphResult({ value: all.data });
  } catch (err) {
    if (err instanceof GraphApiError) return graphErrorFromApiError(err);
    return graphError(err instanceof Error ? err.message : 'Failed to list group conversations');
  }
}

export interface ConversationThread {
  id: string;
  topic?: string;
  hasAttachments?: boolean;
  lastDeliveredDateTime?: string;
  uniqueSenders?: string[];
  preview?: string;
  isLocked?: boolean;
}

export interface ThreadsListResponse {
  value?: ConversationThread[];
  '@odata.nextLink'?: string;
}

/** `GET /groups/{id}/conversations/{id}/threads` — list threads within a conversation. */
export async function listConversationThreads(
  token: string,
  groupId: string,
  conversationId: string,
  options: { top?: number } = {}
): Promise<GraphResponse<ThreadsListResponse>> {
  const topN = options.top && options.top > 0 ? Math.min(Math.max(1, options.top), 200) : undefined;
  const basePath = `/groups/${encodeURIComponent(groupId)}/conversations/${encodeURIComponent(conversationId)}/threads`;
  try {
    if (topN !== undefined) {
      return await callGraph<ThreadsListResponse>(token, `${basePath}?$top=${topN}`);
    }
    const all = await fetchAllPages<ConversationThread>(token, basePath, 'Failed to list threads');
    if (!all.ok || !all.data) return all as unknown as GraphResponse<ThreadsListResponse>;
    return graphResult({ value: all.data });
  } catch (err) {
    if (err instanceof GraphApiError) return graphErrorFromApiError(err);
    return graphError(err instanceof Error ? err.message : 'Failed to list threads');
  }
}

export interface ConversationPost {
  id: string;
  createdDateTime?: string;
  receivedDateTime?: string;
  hasAttachments?: boolean;
  body?: { contentType?: string; content?: string };
  from?: { emailAddress?: { name?: string; address?: string } };
  sender?: { emailAddress?: { name?: string; address?: string } };
}

export interface PostsListResponse {
  value?: ConversationPost[];
  '@odata.nextLink'?: string;
}

/** `GET /groups/{id}/conversations/{id}/threads/{id}/posts` — list posts in a thread. */
export async function listThreadPosts(
  token: string,
  groupId: string,
  conversationId: string,
  threadId: string,
  options: { top?: number } = {}
): Promise<GraphResponse<PostsListResponse>> {
  const topN = options.top && options.top > 0 ? Math.min(Math.max(1, options.top), 200) : undefined;
  const basePath =
    `/groups/${encodeURIComponent(groupId)}/conversations/${encodeURIComponent(conversationId)}` +
    `/threads/${encodeURIComponent(threadId)}/posts`;
  try {
    if (topN !== undefined) {
      return await callGraph<PostsListResponse>(token, `${basePath}?$top=${topN}`);
    }
    const all = await fetchAllPages<ConversationPost>(token, basePath, 'Failed to list posts');
    if (!all.ok || !all.data) return all as unknown as GraphResponse<PostsListResponse>;
    return graphResult({ value: all.data });
  } catch (err) {
    if (err instanceof GraphApiError) return graphErrorFromApiError(err);
    return graphError(err instanceof Error ? err.message : 'Failed to list posts');
  }
}

/**
 * `POST /groups/{id}/conversations/{id}/threads/{id}/posts/{id}/reply` —
 * reply to a specific post in a group thread.
 */
export async function replyToPost(
  token: string,
  ids: { groupId: string; conversationId: string; threadId: string; postId: string },
  body: { contentType?: 'text' | 'html'; content: string }
): Promise<GraphResponse<void>> {
  const path =
    `/groups/${encodeURIComponent(ids.groupId)}/conversations/${encodeURIComponent(ids.conversationId)}` +
    `/threads/${encodeURIComponent(ids.threadId)}/posts/${encodeURIComponent(ids.postId)}/reply`;
  const payload = {
    post: {
      body: {
        contentType: body.contentType ?? 'text',
        content: body.content
      }
    }
  };
  try {
    return await callGraph<void>(token, path, { method: 'POST', body: JSON.stringify(payload) }, false);
  } catch (err) {
    if (err instanceof GraphApiError) return graphErrorFromApiError(err);
    return graphError(err instanceof Error ? err.message : 'Failed to reply to post');
  }
}
