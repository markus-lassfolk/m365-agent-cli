import {
  callGraph,
  fetchAllPages,
  fetchGraphRaw,
  GraphApiError,
  type GraphResponse,
  graphError,
  graphResult
} from './graph-client.js';
import { graphUserPath } from './graph-user-path.js';

function mailFoldersRoot(user?: string): string {
  return graphUserPath(user, 'mailFolders');
}

function contactsRoot(user?: string): string {
  return graphUserPath(user, 'contacts');
}

/** Graph [mailFolder](https://learn.microsoft.com/en-us/graph/api/resources/mailfolder) (subset). */
export interface OutlookMailFolder {
  id: string;
  displayName: string;
  parentFolderId?: string;
  childFolderCount?: number;
  unreadItemCount?: number;
  totalItemCount?: number;
}

/** Graph [message](https://learn.microsoft.com/en-us/graph/api/resources/message) (subset). */
export interface OutlookMessage {
  id: string;
  subject?: string;
  bodyPreview?: string;
  body?: { contentType?: string; content?: string };
  receivedDateTime?: string;
  sentDateTime?: string;
  lastModifiedDateTime?: string;
  isRead?: boolean;
  importance?: string;
  categories?: string[];
  from?: { emailAddress?: { name?: string; address?: string } };
  toRecipients?: Array<{ emailAddress?: { name?: string; address?: string } }>;
  ccRecipients?: Array<{ emailAddress?: { name?: string; address?: string } }>;
  /** Open in Outlook on the web (when returned by Graph). */
  webLink?: string;
  hasAttachments?: boolean;
  followupFlag?: {
    flagStatus?: 'notFlagged' | 'flagged' | 'complete';
    startDateTime?: { dateTime?: string; timeZone?: string };
    dueDateTime?: { dateTime?: string; timeZone?: string };
  };
}

/** Graph [contact](https://learn.microsoft.com/en-us/graph/api/resources/contact) (subset). */
export interface OutlookContact {
  id: string;
  displayName?: string;
  givenName?: string;
  surname?: string;
  emailAddresses?: Array<{ name?: string; address?: string; type?: string }>;
  mobilePhone?: string;
  businessPhones?: string[];
  companyName?: string;
  jobTitle?: string;
}

export async function listMailFolders(token: string, user?: string): Promise<GraphResponse<OutlookMailFolder[]>> {
  return fetchAllPages<OutlookMailFolder>(token, mailFoldersRoot(user), 'Failed to list mail folders');
}

/** Child folders of a folder (e.g. under Inbox). */
export async function listChildMailFolders(
  token: string,
  parentFolderId: string,
  user?: string
): Promise<GraphResponse<OutlookMailFolder[]>> {
  return fetchAllPages<OutlookMailFolder>(
    token,
    `${mailFoldersRoot(user)}/${encodeURIComponent(parentFolderId)}/childFolders`,
    'Failed to list child mail folders'
  );
}

export async function getMailFolder(
  token: string,
  folderId: string,
  user?: string
): Promise<GraphResponse<OutlookMailFolder>> {
  try {
    const result = await callGraph<OutlookMailFolder>(
      token,
      `${mailFoldersRoot(user)}/${encodeURIComponent(folderId)}`
    );
    if (!result.ok || !result.data) {
      return graphError(result.error?.message || 'Failed to get mail folder', result.error?.code, result.error?.status);
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get mail folder');
  }
}

export async function createMailFolder(
  token: string,
  displayName: string,
  parentFolderId: string | undefined,
  user?: string
): Promise<GraphResponse<OutlookMailFolder>> {
  const body: Record<string, unknown> = { displayName };
  if (parentFolderId) body.parentFolderId = parentFolderId;
  try {
    const result = await callGraph<OutlookMailFolder>(token, mailFoldersRoot(user), {
      method: 'POST',
      body: JSON.stringify(body)
    });
    if (!result.ok || !result.data) {
      return graphError(
        result.error?.message || 'Failed to create mail folder',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to create mail folder');
  }
}

export async function updateMailFolder(
  token: string,
  folderId: string,
  displayName: string,
  user?: string
): Promise<GraphResponse<OutlookMailFolder>> {
  try {
    const result = await callGraph<OutlookMailFolder>(
      token,
      `${mailFoldersRoot(user)}/${encodeURIComponent(folderId)}`,
      {
        method: 'PATCH',
        body: JSON.stringify({ displayName })
      }
    );
    if (!result.ok || !result.data) {
      return graphError(
        result.error?.message || 'Failed to update mail folder',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to update mail folder');
  }
}

export async function deleteMailFolder(token: string, folderId: string, user?: string): Promise<GraphResponse<void>> {
  try {
    return await callGraph<void>(
      token,
      `${mailFoldersRoot(user)}/${encodeURIComponent(folderId)}`,
      { method: 'DELETE' },
      false
    );
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to delete mail folder');
  }
}

/** All mail folders (nested), depth-first under the mailbox root. */
export async function listAllMailFoldersRecursive(
  token: string,
  user?: string
): Promise<GraphResponse<OutlookMailFolder[]>> {
  try {
    const root = await listMailFolders(token, user);
    if (!root.ok || !root.data) return root;
    const all: OutlookMailFolder[] = [];

    async function walk(folder: OutlookMailFolder): Promise<void> {
      all.push(folder);
      const n = folder.childFolderCount ?? 0;
      if (n === 0) return;
      const ch = await listChildMailFolders(token, folder.id, user);
      if (!ch.ok || !ch.data?.length) return;
      for (const c of ch.data) {
        await walk(c);
      }
    }

    for (const f of root.data) {
      await walk(f);
    }
    return graphResult(all);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to list mail folders');
  }
}

export interface MessagesQueryOptions {
  filter?: string;
  orderby?: string;
  select?: string;
  /** When set, only one page (no automatic paging). */
  top?: number;
  skip?: number;
  /**
   * Keyword search (`$search`). Requires `ConsistencyLevel: eventual` (applied automatically).
   * Do not combine with `$filter` on the same request; use client-side filtering if both apply.
   */
  search?: string;
}

/** Query for `GET /me/messages` (mailbox-wide, not folder-scoped). */
export interface RootMailboxMessagesQuery extends MessagesQueryOptions {
  skip?: number;
  /**
   * Sets OData `$search` (keyword search). Graph requires `ConsistencyLevel: eventual` on the request.
   * Do not combine with `$filter` on the same request (Graph limitation).
   */
  search?: string;
}

/** Graph [attachment](https://learn.microsoft.com/en-us/graph/api/resources/attachment) on a message (subset). */
export interface GraphMailMessageAttachment {
  id: string;
  name?: string;
  contentType?: string;
  size?: number;
  isInline?: boolean;
  /** Reference / link attachment target URL. */
  sourceUrl?: string;
  '@odata.type'?: string;
}

function messagesPath(folderId: string, user: string | undefined, query?: MessagesQueryOptions): string {
  const params = new URLSearchParams();
  if (query?.filter) params.set('$filter', query.filter);
  if (query?.orderby) params.set('$orderby', query.orderby);
  if (query?.select) params.set('$select', query.select);
  if (query?.top !== undefined) params.set('$top', String(query.top));
  if (query?.skip !== undefined) params.set('$skip', String(query.skip));
  if (query?.search) {
    const escaped = query.search.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
    params.set('$search', `"${escaped}"`);
  }
  const qs = params.toString() ? `?${params.toString()}` : '';
  return `${mailFoldersRoot(user)}/${encodeURIComponent(folderId)}/messages${qs}`;
}

function folderMessagesSearchRequestInit(query?: MessagesQueryOptions): RequestInit | undefined {
  return query?.search ? { headers: { ConsistencyLevel: 'eventual' } } : undefined;
}

function mailboxMessagesPath(user: string | undefined, query?: RootMailboxMessagesQuery): string {
  const params = new URLSearchParams();
  if (query?.filter) params.set('$filter', query.filter);
  if (query?.orderby) params.set('$orderby', query.orderby);
  if (query?.select) params.set('$select', query.select);
  if (query?.top !== undefined) params.set('$top', String(query.top));
  if (query?.skip !== undefined) params.set('$skip', String(query.skip));
  if (query?.search) {
    const escaped = query.search.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
    params.set('$search', `"${escaped}"`);
  }
  const qs = params.toString() ? `?${params.toString()}` : '';
  return `${graphUserPath(user, 'messages')}${qs}`;
}

function searchRequestInit(query?: RootMailboxMessagesQuery): RequestInit | undefined {
  return query?.search ? { headers: { ConsistencyLevel: 'eventual' } } : undefined;
}

/**
 * List messages in a folder. Well-known folder ids include `inbox`, `sentitems`, `drafts`, `deleteditems`, `archive`, `junkemail`.
 * Use `top` for a single page; omit `top` to page through all results (can be large).
 */
export async function listMessagesInFolder(
  token: string,
  folderId: string,
  user?: string,
  query?: MessagesQueryOptions
): Promise<GraphResponse<OutlookMessage[]>> {
  const path = messagesPath(folderId, user, query);
  const singlePage = query?.top !== undefined;
  const req = folderMessagesSearchRequestInit(query);

  if (singlePage) {
    try {
      const result = await callGraph<{ value: OutlookMessage[] }>(token, path, req ?? {});
      if (!result.ok || !result.data) {
        return graphError(result.error?.message || 'Failed to list messages', result.error?.code, result.error?.status);
      }
      return graphResult(result.data.value || []);
    } catch (err) {
      if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
      return graphError(err instanceof Error ? err.message : 'Failed to list messages');
    }
  }

  return fetchAllPages<OutlookMessage>(token, path, 'Failed to list messages', undefined, req);
}

/**
 * List messages via `GET /me/messages` (entire mailbox, not limited to one folder).
 * When `search` is set, Graph requires `ConsistencyLevel: eventual` (applied automatically).
 */
export async function listMailboxMessages(
  token: string,
  user?: string,
  query?: RootMailboxMessagesQuery
): Promise<GraphResponse<OutlookMessage[]>> {
  const path = mailboxMessagesPath(user, query);
  const singlePage = query?.top !== undefined;
  const req = searchRequestInit(query);

  if (singlePage) {
    try {
      const result = await callGraph<{ value: OutlookMessage[] }>(token, path, req ?? {});
      if (!result.ok || !result.data) {
        return graphError(result.error?.message || 'Failed to list messages', result.error?.code, result.error?.status);
      }
      return graphResult(result.data.value || []);
    } catch (err) {
      if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
      return graphError(err instanceof Error ? err.message : 'Failed to list messages');
    }
  }

  return fetchAllPages<OutlookMessage>(token, path, 'Failed to list messages', undefined, req);
}

/** `GET /me/messages/{id}` — message ids are unique within the mailbox (folder path not required). */
export async function getMessage(
  token: string,
  messageId: string,
  user?: string,
  select?: string
): Promise<GraphResponse<OutlookMessage>> {
  const qs = select ? `?$select=${encodeURIComponent(select)}` : '';
  try {
    const result = await callGraph<OutlookMessage>(
      token,
      `${graphUserPath(user, 'messages')}/${encodeURIComponent(messageId)}${qs}`
    );
    if (!result.ok || !result.data) {
      return graphError(result.error?.message || 'Failed to get message', result.error?.code, result.error?.status);
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get message');
  }
}

/** Payload for `POST /me/messages` with `isDraft: true`. */
export interface GraphCreateDraftMessageInput {
  subject?: string;
  bodyContent: string;
  bodyContentType: 'Text' | 'HTML';
  toAddresses?: string[];
  ccAddresses?: string[];
  categories?: string[];
}

/** Create a draft in the Drafts folder (`POST /me/messages` with `isDraft: true`). */
export async function createDraftMessage(
  token: string,
  input: GraphCreateDraftMessageInput,
  user?: string
): Promise<GraphResponse<OutlookMessage>> {
  const message: Record<string, unknown> = {
    isDraft: true,
    body: {
      contentType: input.bodyContentType,
      content: input.bodyContent
    }
  };
  if (input.subject !== undefined) {
    message.subject = input.subject;
  }
  if (input.toAddresses?.length) {
    message.toRecipients = input.toAddresses.map((address) => ({ emailAddress: { address } }));
  }
  if (input.ccAddresses?.length) {
    message.ccRecipients = input.ccAddresses.map((address) => ({ emailAddress: { address } }));
  }
  if (input.categories?.length) {
    message.categories = input.categories;
  }
  try {
    const result = await callGraph<OutlookMessage>(
      token,
      graphUserPath(user, 'messages'),
      { method: 'POST', body: JSON.stringify(message) },
      true
    );
    if (!result.ok || !result.data) {
      return graphError(result.error?.message || 'Failed to create draft', result.error?.code, result.error?.status);
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to create draft');
  }
}

export async function patchMailMessage(
  token: string,
  messageId: string,
  patch: Record<string, unknown>,
  user?: string
): Promise<GraphResponse<OutlookMessage>> {
  try {
    const result = await callGraph<OutlookMessage>(
      token,
      `${graphUserPath(user, 'messages')}/${encodeURIComponent(messageId)}`,
      { method: 'PATCH', body: JSON.stringify(patch) }
    );
    if (!result.ok || !result.data) {
      return graphError(result.error?.message || 'Failed to update message', result.error?.code, result.error?.status);
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to update message');
  }
}

/** `POST /me/messages/{id}/attachments` — add a file attachment to a draft message. */
export async function addFileAttachmentToMailMessage(
  token: string,
  messageId: string,
  attachment: { name: string; contentType: string; contentBytes: string },
  user?: string
): Promise<GraphResponse<GraphMailMessageAttachment>> {
  const body = {
    '@odata.type': '#microsoft.graph.fileAttachment',
    name: attachment.name,
    contentType: attachment.contentType,
    contentBytes: attachment.contentBytes
  };
  try {
    const result = await callGraph<GraphMailMessageAttachment>(
      token,
      `${graphUserPath(user, 'messages')}/${encodeURIComponent(messageId)}/attachments`,
      { method: 'POST', body: JSON.stringify(body) },
      true
    );
    if (!result.ok || !result.data) {
      return graphError(result.error?.message || 'Failed to add attachment', result.error?.code, result.error?.status);
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to add attachment');
  }
}

/** `POST /me/messages/{id}/attachments` — add a link (`referenceAttachment`) to a draft message. */
export async function addReferenceAttachmentToMailMessage(
  token: string,
  messageId: string,
  attachment: { name: string; sourceUrl: string },
  user?: string
): Promise<GraphResponse<GraphMailMessageAttachment>> {
  const body = {
    '@odata.type': '#microsoft.graph.referenceAttachment',
    name: attachment.name,
    sourceUrl: attachment.sourceUrl
  };
  try {
    const result = await callGraph<GraphMailMessageAttachment>(
      token,
      `${graphUserPath(user, 'messages')}/${encodeURIComponent(messageId)}/attachments`,
      { method: 'POST', body: JSON.stringify(body) },
      true
    );
    if (!result.ok || !result.data) {
      return graphError(
        result.error?.message || 'Failed to add link attachment',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to add link attachment');
  }
}

export async function deleteMailMessage(token: string, messageId: string, user?: string): Promise<GraphResponse<void>> {
  try {
    return await callGraph<void>(
      token,
      `${graphUserPath(user, 'messages')}/${encodeURIComponent(messageId)}`,
      { method: 'DELETE' },
      false
    );
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to delete message');
  }
}

export async function moveMailMessage(
  token: string,
  messageId: string,
  destinationFolderId: string,
  user?: string
): Promise<GraphResponse<OutlookMessage>> {
  try {
    const result = await callGraph<OutlookMessage>(
      token,
      `${graphUserPath(user, 'messages')}/${encodeURIComponent(messageId)}/move`,
      {
        method: 'POST',
        body: JSON.stringify({ destinationId: destinationFolderId })
      }
    );
    if (!result.ok || !result.data) {
      return graphError(result.error?.message || 'Failed to move message', result.error?.code, result.error?.status);
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to move message');
  }
}

export async function copyMailMessage(
  token: string,
  messageId: string,
  destinationFolderId: string,
  user?: string
): Promise<GraphResponse<OutlookMessage>> {
  try {
    const result = await callGraph<OutlookMessage>(
      token,
      `${graphUserPath(user, 'messages')}/${encodeURIComponent(messageId)}/copy`,
      {
        method: 'POST',
        body: JSON.stringify({ destinationId: destinationFolderId })
      }
    );
    if (!result.ok || !result.data) {
      return graphError(result.error?.message || 'Failed to copy message', result.error?.code, result.error?.status);
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to copy message');
  }
}

/** Send a new message in one request (`POST /sendMail`). */
export async function sendMail(
  token: string,
  body: { message: Record<string, unknown>; saveToSentItems?: boolean },
  user?: string
): Promise<GraphResponse<void>> {
  try {
    return await callGraph<void>(
      token,
      `${graphUserPath(user, 'sendMail')}`,
      { method: 'POST', body: JSON.stringify(body) },
      false
    );
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to send mail');
  }
}

/** Send an existing draft (`POST /messages/{id}/send`). */
export async function sendMailMessage(token: string, messageId: string, user?: string): Promise<GraphResponse<void>> {
  try {
    return await callGraph<void>(
      token,
      `${graphUserPath(user, 'messages')}/${encodeURIComponent(messageId)}/send`,
      { method: 'POST' },
      false
    );
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to send message');
  }
}

export async function listMailMessageAttachments(
  token: string,
  messageId: string,
  user?: string
): Promise<GraphResponse<GraphMailMessageAttachment[]>> {
  return fetchAllPages<GraphMailMessageAttachment>(
    token,
    `${graphUserPath(user, 'messages')}/${encodeURIComponent(messageId)}/attachments`,
    'Failed to list attachments'
  );
}

export async function getMailMessageAttachment(
  token: string,
  messageId: string,
  attachmentId: string,
  user?: string
): Promise<GraphResponse<GraphMailMessageAttachment>> {
  try {
    const result = await callGraph<GraphMailMessageAttachment>(
      token,
      `${graphUserPath(user, 'messages')}/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachmentId)}`
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

export async function downloadMailMessageAttachmentBytes(
  token: string,
  messageId: string,
  attachmentId: string,
  user?: string
): Promise<GraphResponse<Uint8Array>> {
  const path = `${graphUserPath(user, 'messages')}/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachmentId)}/$value`;
  try {
    const res = await fetchGraphRaw(token, path);
    if (!res.ok) {
      let message = `Failed to download attachment: HTTP ${res.status}`;
      try {
        const json = (await res.json()) as { error?: { message?: string } };
        message = json.error?.message || message;
      } catch {
        // ignore
      }
      return graphError(message, undefined, res.status);
    }
    const buf = new Uint8Array(await res.arrayBuffer());
    return graphResult(buf);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to download attachment');
  }
}

export async function createMailReplyDraft(
  token: string,
  messageId: string,
  user?: string,
  comment?: string
): Promise<GraphResponse<OutlookMessage>> {
  const body = JSON.stringify(comment ? { comment } : {});
  try {
    const result = await callGraph<OutlookMessage>(
      token,
      `${graphUserPath(user, 'messages')}/${encodeURIComponent(messageId)}/createReply`,
      { method: 'POST', body },
      true
    );
    if (!result.ok || !result.data) {
      return graphError(
        result.error?.message || 'Failed to create reply draft',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to create reply draft');
  }
}

export async function createMailReplyAllDraft(
  token: string,
  messageId: string,
  user?: string,
  comment?: string
): Promise<GraphResponse<OutlookMessage>> {
  const body = JSON.stringify(comment ? { comment } : {});
  try {
    const result = await callGraph<OutlookMessage>(
      token,
      `${graphUserPath(user, 'messages')}/${encodeURIComponent(messageId)}/createReplyAll`,
      { method: 'POST', body },
      true
    );
    if (!result.ok || !result.data) {
      return graphError(
        result.error?.message || 'Failed to create reply-all draft',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to create reply-all draft');
  }
}

export async function createMailForwardDraft(
  token: string,
  messageId: string,
  toRecipients: string[],
  user?: string,
  comment?: string
): Promise<GraphResponse<OutlookMessage>> {
  const recipients = toRecipients.map((address) => ({
    emailAddress: { address }
  }));
  const payload: Record<string, unknown> = { toRecipients: recipients };
  if (comment) payload.comment = comment;
  try {
    const result = await callGraph<OutlookMessage>(
      token,
      `${graphUserPath(user, 'messages')}/${encodeURIComponent(messageId)}/createForward`,
      { method: 'POST', body: JSON.stringify(payload) },
      true
    );
    if (!result.ok || !result.data) {
      return graphError(
        result.error?.message || 'Failed to create forward draft',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to create forward draft');
  }
}

export async function listContacts(token: string, user?: string): Promise<GraphResponse<OutlookContact[]>> {
  return fetchAllPages<OutlookContact>(token, contactsRoot(user), 'Failed to list contacts');
}

export async function getContact(
  token: string,
  contactId: string,
  user?: string,
  select?: string
): Promise<GraphResponse<OutlookContact>> {
  const qs = select ? `?$select=${encodeURIComponent(select)}` : '';
  try {
    const result = await callGraph<OutlookContact>(
      token,
      `${contactsRoot(user)}/${encodeURIComponent(contactId)}${qs}`
    );
    if (!result.ok || !result.data) {
      return graphError(result.error?.message || 'Failed to get contact', result.error?.code, result.error?.status);
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get contact');
  }
}

export async function createContact(
  token: string,
  body: Record<string, unknown>,
  user?: string
): Promise<GraphResponse<OutlookContact>> {
  try {
    const result = await callGraph<OutlookContact>(token, contactsRoot(user), {
      method: 'POST',
      body: JSON.stringify(body)
    });
    if (!result.ok || !result.data) {
      return graphError(result.error?.message || 'Failed to create contact', result.error?.code, result.error?.status);
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to create contact');
  }
}

export async function updateContact(
  token: string,
  contactId: string,
  patch: Record<string, unknown>,
  user?: string
): Promise<GraphResponse<OutlookContact>> {
  try {
    const result = await callGraph<OutlookContact>(token, `${contactsRoot(user)}/${encodeURIComponent(contactId)}`, {
      method: 'PATCH',
      body: JSON.stringify(patch)
    });
    if (!result.ok || !result.data) {
      return graphError(result.error?.message || 'Failed to update contact', result.error?.code, result.error?.status);
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to update contact');
  }
}

export async function deleteContact(token: string, contactId: string, user?: string): Promise<GraphResponse<void>> {
  try {
    return await callGraph<void>(
      token,
      `${contactsRoot(user)}/${encodeURIComponent(contactId)}`,
      { method: 'DELETE' },
      false
    );
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to delete contact');
  }
}
