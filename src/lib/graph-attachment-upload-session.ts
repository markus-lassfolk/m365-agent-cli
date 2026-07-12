/**
 * Large file attachments for Outlook mail and calendar events via Graph upload session
 * (POST …/attachments/createUploadSession + chunked PUT to uploadUrl).
 */

import { callGraph, GraphApiError, type GraphResponse, graphError, graphResult } from './graph-client.js';
import { graphUserPath } from './graph-user-path.js';

/**
 * Raw file size above which we MUST use an upload session instead of an inline base64 POST.
 * Graph rejects `createUploadSession` for files smaller than 3 MB
 * (`ErrorAttachmentSizeShouldNotBeLessThanMinimumSize`) and rejects a single inline POST for
 * files 3 MB or larger, so 3 MB is the exact crossover — not a tunable preference.
 * @see https://learn.microsoft.com/graph/outlook-large-attachments
 */
export const GRAPH_OUTLOOK_ATTACHMENT_SESSION_THRESHOLD_BYTES = 3 * 1024 * 1024;

export interface GraphAttachmentUploadSession {
  uploadUrl: string;
  expirationDateTime?: string;
}

const CHUNK_SIZE = 4 * 1024 * 1024;

/**
 * Upload bytes to a pre-authorized Graph upload URL (no Bearer on PUT).
 * Returns parsed JSON from the final successful response when present.
 */
export async function uploadBufferViaGraphUploadUrl(
  uploadUrl: string,
  data: Buffer
): Promise<GraphResponse<Record<string, unknown> | undefined>> {
  const total = data.byteLength;
  if (total === 0) {
    return graphError('Cannot upload zero-byte attachment via upload session', undefined, 400);
  }
  let start = 0;
  let lastJson: Record<string, unknown> | undefined;
  let locationId: string | undefined;
  while (start < total) {
    const end = Math.min(start + CHUNK_SIZE, total);
    const slice = data.subarray(start, end);
    const contentRange = `bytes ${start}-${end - 1}/${total}`;
    let response: Response;
    try {
      // codeql[js/file-access-to-http]: Chunked PUT to Graph-provided uploadUrl; body is the caller's attachment bytes, not arbitrary file exfiltration.
      response = await fetch(uploadUrl, {
        method: 'PUT',
        headers: {
          'Content-Length': String(slice.byteLength),
          'Content-Range': contentRange
        },
        body: new Blob([Uint8Array.from(slice)])
      });
    } catch (err) {
      return graphError(err instanceof Error ? err.message : 'Upload chunk failed');
    }
    const text = await response.text();
    if (!response.ok) {
      return graphError(text || `Upload failed: HTTP ${response.status}`, undefined, response.status);
    }
    // The final successful PUT carries the attachment ID in the Location header URL
    // (the response body is often empty for large attachments), so capture it here.
    const location = response.headers.get('location');
    if (location) {
      locationId = parseAttachmentIdFromLocation(location) ?? locationId;
    }
    if (text) {
      try {
        const parsed = JSON.parse(text) as Record<string, unknown>;
        if (parsed && typeof parsed === 'object') {
          lastJson = parsed;
        }
      } catch {
        // non-JSON success body
      }
    }
    start = end;
  }
  // Prefer an id from the response body; otherwise fall back to the Location-header id.
  const result: Record<string, unknown> = { ...(lastJson ?? {}) };
  if (typeof result.id !== 'string' && locationId) {
    result.id = locationId;
  }
  return graphResult(Object.keys(result).length > 0 ? result : undefined);
}

/**
 * Extract the attachment id from an upload-session `Location` header URL, e.g.
 * `.../messages/{id}/attachments/{attachmentId}` or an EWS-style `...('{attachmentId}')`.
 */
function parseAttachmentIdFromLocation(location: string): string | undefined {
  let path = location;
  try {
    path = new URL(location).pathname;
  } catch {
    // Not an absolute URL — treat the raw value as the path.
  }
  // Handle OData key syntax: .../Attachments('AAMk...')
  const odataKey = path.match(/\(['"]([^'"]+)['"]\)\/?$/);
  if (odataKey) return decodeURIComponent(odataKey[1]);
  const segments = path.split('/').filter(Boolean);
  const last = segments[segments.length - 1];
  return last ? decodeURIComponent(last) : undefined;
}

export async function createMailMessageFileAttachmentUploadSession(
  token: string,
  messageId: string,
  name: string,
  size: number,
  contentType: string,
  user?: string
): Promise<GraphResponse<GraphAttachmentUploadSession>> {
  const path = `${graphUserPath(user, `messages/${encodeURIComponent(messageId)}/attachments/createUploadSession`)}`;
  const body = {
    AttachmentItem: {
      attachmentType: 'file',
      name,
      size,
      isInline: false,
      contentType: contentType || 'application/octet-stream'
    }
  };
  try {
    const result = await callGraph<GraphAttachmentUploadSession>(token, path, {
      method: 'POST',
      body: JSON.stringify(body)
    });
    if (!result.ok || !result.data?.uploadUrl) {
      return graphError(
        result.error?.message || 'Failed to create message attachment upload session',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to create message attachment upload session');
  }
}

export async function createCalendarEventFileAttachmentUploadSession(
  token: string,
  eventId: string,
  name: string,
  size: number,
  contentType: string,
  user?: string
): Promise<GraphResponse<GraphAttachmentUploadSession>> {
  const path = `${graphUserPath(user, `events/${encodeURIComponent(eventId)}/attachments/createUploadSession`)}`;
  const body = {
    AttachmentItem: {
      attachmentType: 'file',
      name,
      size,
      isInline: false,
      contentType: contentType || 'application/octet-stream'
    }
  };
  try {
    const result = await callGraph<GraphAttachmentUploadSession>(token, path, {
      method: 'POST',
      body: JSON.stringify(body)
    });
    if (!result.ok || !result.data?.uploadUrl) {
      return graphError(
        result.error?.message || 'Failed to create event attachment upload session',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to create event attachment upload session');
  }
}
