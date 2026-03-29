import { randomBytes } from 'node:crypto';
import { createReadStream, createWriteStream } from 'node:fs';
import { open } from 'node:fs/promises';
import { mkdir, rename, stat, unlink } from 'node:fs/promises';
import { homedir } from 'node:os';
import { basename, dirname, resolve } from 'node:path';
import { Readable } from 'node:stream';
import { GRAPH_BASE_URL } from './graph-constants.js';

export { GRAPH_BASE_URL };

export interface GraphError {
  message: string;
  code?: string;
  status?: number;
}

export class GraphApiError extends Error {
  constructor(
    message: string,
    public readonly code?: string,
    public readonly status?: number
  ) {
    super(message);
    this.name = 'GraphApiError';
  }
}

export interface GraphResponse<T> {
  ok: boolean;
  data?: T;
  error?: GraphError;
}

export interface DriveItemReference {
  driveId?: string;
  id?: string;
  path?: string;
}

export interface DriveItem {
  id: string;
  name: string;
  webUrl?: string;
  size?: number;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  file?: { mimeType?: string };
  folder?: { childCount?: number };
  parentReference?: { driveId?: string; id?: string; path?: string };
  '@microsoft.graph.downloadUrl'?: string;
}

export interface DriveItemListResponse {
  value: DriveItem[];
}

export interface SharingLinkResult {
  id?: string;
  webUrl?: string;
  type?: string;
  scope?: string;
}

export interface OfficeCollabLinkResult {
  item: DriveItem;
  link: SharingLinkResult;
  collaborationUrl?: string;
  lockAcquired: boolean;
}

export interface CheckinResult {
  item: DriveItem;
  checkedIn: boolean;
  comment?: string;
}

export interface UploadLargeResult {
  uploadUrl: string;
  expirationDateTime?: string;
  driveItem?: DriveItem;
}

async function streamWebToFile(body: ReadableStream<Uint8Array>, filePath: string): Promise<number> {
  const stream = createWriteStream(filePath, { flags: 'w', mode: 0o600 });
  let bytesWritten = 0;

  try {
    for await (const chunk of body) {
      if (!stream.write(chunk)) {
        await new Promise<void>((resolveDrain, rejectDrain) => {
          const onDrain = () => {
            stream.off('error', onError);
            resolveDrain();
          };
          const onError = (err: Error) => {
            stream.off('drain', onDrain);
            rejectDrain(err);
          };
          stream.once('drain', onDrain);
          stream.once('error', onError);
        });
      }
      bytesWritten += chunk.byteLength;
    }

    await new Promise<void>((resolveClose, rejectClose) => {
      stream.end((err?: Error | null) => {
        if (err) rejectClose(err);
        else resolveClose();
      });
    });

    return bytesWritten;
  } catch (err) {
    stream.destroy();
    try {
      await unlink(filePath);
    } catch {}
    throw err;
  }
}

export function graphResult<T>(data: T): GraphResponse<T> {
  return { ok: true, data };
}

export function graphError(message: string, code?: string, status?: number): GraphResponse<never> {
  return { ok: false, error: { message, code, status } };
}

export async function fetchAllPages<T>(
  token: string,
  initialPath: string,
  errorMessage: string
): Promise<GraphResponse<T[]>> {
  const items: T[] = [];
  let path = initialPath;

  while (path) {
    let result: GraphResponse<{ value: T[]; '@odata.nextLink'?: string }>;
    try {
      result = await callGraph<{ value: T[]; '@odata.nextLink'?: string }>(token, path);
    } catch (err) {
      if (err instanceof GraphApiError) {
        return graphError(err.message, err.code, err.status) as GraphResponse<T[]>;
      }
      return graphError(err instanceof Error ? err.message : errorMessage) as GraphResponse<T[]>;
    }
    if (!result.ok || !result.data) {
      return graphError(
        result.error?.message || errorMessage,
        result.error?.code,
        result.error?.status
      ) as GraphResponse<T[]>;
    }
    items.push(...(result.data.value || []));
    path = result.data['@odata.nextLink']
      ? (() => {
          try {
            const nextLink = result.data['@odata.nextLink']!;
            if (nextLink.startsWith(GRAPH_BASE_URL)) {
              return nextLink.substring(GRAPH_BASE_URL.length);
            }
            const nextUrl = new URL(nextLink);
            const baseUrlUrl = new URL(GRAPH_BASE_URL);
            if (nextUrl.origin === baseUrlUrl.origin) {
              const baseDir = baseUrlUrl.pathname.replace(/\/$/, '');
              if (nextUrl.pathname.startsWith(baseDir)) {
                return nextUrl.pathname.substring(baseDir.length) + nextUrl.search;
              }
            }
            return '';
          } catch {
            return '';
          }
        })()
      : '';
  }
  return graphResult(items);
}

export async function fetchGraphRaw(token: string, path: string, options: RequestInit = {}): Promise<Response> {
  return fetch(`${GRAPH_BASE_URL}${path}`, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      ...(options.headers || {})
    }
  });
}

export async function callGraph<T>(
  token: string,
  path: string,
  options: RequestInit = {},
  expectJson: boolean = true
): Promise<GraphResponse<T>> {
  let response: Response;
  try {
    response = await fetch(`${GRAPH_BASE_URL}${path}`, {
      ...options,
      headers: {
        Authorization: `Bearer ${token}`,
        ...(expectJson ? { Accept: 'application/json' } : {}),
        ...(options.body && !(options.body instanceof Uint8Array) && !(options.body instanceof ArrayBuffer)
          ? { 'Content-Type': 'application/json' }
          : {}),
        ...(options.headers || {})
      }
    });
  } catch (err) {
    // Network-level failure (DNS, connection refused, etc.) — surface it as a thrown error
    throw new GraphApiError(err instanceof Error ? err.message : 'Graph request failed');
  }

  if (!response.ok) {
    let message = `Graph request failed: HTTP ${response.status}`;
    let code: string | undefined;
    try {
      const json = (await response.json()) as { error?: { code?: string; message?: string } };
      message = json.error?.message || message;
      code = json.error?.code;
    } catch {
      // Non-JSON error body — throw with HTTP status instead
      throw new GraphApiError(message, code, response.status);
    }
    throw new GraphApiError(message, code, response.status);
  }

  if (!expectJson || response.status === 204) {
    return graphResult(undefined as T);
  }

  return graphResult((await response.json()) as T);
}

function buildItemPath(reference?: DriveItemReference): string {
  if (!reference?.id) return '/me/drive/root';

  const drivePrefix = reference.driveId ? `/drives/${encodeURIComponent(reference.driveId)}` : '/me/drive';
  return `${drivePrefix}/items/${encodeURIComponent(reference.id)}`;
}

/**
 * Encode a query string for Graph Drive search.
 *
 * encodeURIComponent encodes most characters, but Graph's search(q='...') URL parameter
 * uses single-quoted strings in the URL path. Apostrophes, parentheses, and exclamation marks
 * must therefore also be re-encoded to prevent query syntax injection or truncation.
 *
 * @param query - Raw search query string
 * @returns Percent-encoded query safe for use in Graph search URLs
 */
function encodeGraphSearchQuery(query: string): string {
  return encodeURIComponent(query).replace(/[!'()*]/g, (char) => `%${char.charCodeAt(0).toString(16).toUpperCase()}`);
}

export async function listFiles(token: string, folder?: DriveItemReference): Promise<GraphResponse<DriveItem[]>> {
  const basePath = buildItemPath(folder);
  return fetchAllPages<DriveItem>(token, `${basePath}/children`, 'Failed to list files');
}

export async function searchFiles(token: string, query: string): Promise<GraphResponse<DriveItem[]>> {
  return fetchAllPages<DriveItem>(
    token,
    `/me/drive/root/search(q='${encodeGraphSearchQuery(query)}')`,
    'Failed to search files'
  );
}

export async function getFileMetadata(token: string, itemId: string): Promise<GraphResponse<DriveItem>> {
  try {
    return await callGraph<DriveItem>(token, `/me/drive/items/${encodeURIComponent(itemId)}`);
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to get file metadata');
  }
}

export async function uploadFile(
  token: string,
  localPath: string,
  folder?: DriveItemReference
): Promise<GraphResponse<DriveItem>> {
  try {
    const absolutePath = resolve(localPath);
    const fileStats = await stat(absolutePath);
    if (!fileStats.isFile()) return graphError(`Not a file: ${absolutePath}`);
    if (fileStats.size > 250 * 1024 * 1024) {
      return graphError('File exceeds 250MB simple upload limit. Use upload-large instead.');
    }

    const fileName = basename(absolutePath);
    const folderPath = folder?.id ? `${buildItemPath(folder)}:/` : '/me/drive/root:/';
    const stream = createReadStream(absolutePath);
    try {
      return await callGraph<DriveItem>(token, `${folderPath}${encodeURIComponent(fileName)}:/content`, {
        method: 'PUT',
        body: Readable.toWeb(stream) as unknown as BodyInit,
        headers: {
          'Content-Type': 'application/octet-stream'
        }
      });
    } catch (err) {
      if (err instanceof GraphApiError) {
        return graphError(err.message, err.code, err.status);
      }
      return graphError(err instanceof Error ? err.message : 'Upload failed');
    } finally {
      stream.destroy();
    }
  } catch (err) {
    return graphError(err instanceof Error ? err.message : 'Upload failed');
  }
}

export async function createLargeUploadSession(
  token: string,
  localPath: string,
  folder?: DriveItemReference
): Promise<GraphResponse<UploadLargeResult>> {
  try {
    const absolutePath = resolve(localPath);
    const fileStats = await stat(absolutePath);
    if (!fileStats.isFile()) return graphError(`Not a file: ${absolutePath}`);
    if (fileStats.size > 4 * 1024 * 1024 * 1024) {
      return graphError('File exceeds 4GB large upload limit.');
    }

    const fileName = basename(absolutePath);
    const folderPath = folder?.id ? `${buildItemPath(folder)}:/` : '/me/drive/root:/';

    // Step 1: Create the upload session
    let sessionResult: GraphResponse<UploadLargeResult>;
    try {
      sessionResult = await callGraph<UploadLargeResult>(
        token,
        `${folderPath}${encodeURIComponent(fileName)}:/createUploadSession`,
        {
          method: 'POST',
          body: JSON.stringify({ item: { '@microsoft.graph.conflictBehavior': 'replace', name: fileName } })
        }
      );
    } catch (err) {
      if (err instanceof GraphApiError) {
        return graphError(err.message, err.code, err.status);
      }
      return graphError(err instanceof Error ? err.message : 'Failed to create upload session');
    }

    if (!sessionResult.ok || !sessionResult.data) {
      return sessionResult;
    }

    const { uploadUrl, expirationDateTime } = sessionResult.data;

    // Step 2: Upload the file in chunks
    const fileSize = fileStats.size;
    const CHUNK_SIZE = 10 * 1024 * 1024; // 10MB chunks
    const fileHandle = await open(absolutePath, 'r');

    try {
      let offset = 0;
      let lastSuccessfulResponse: Response | null = null;

      while (offset < fileSize) {
        const chunkLength = Math.min(CHUNK_SIZE, fileSize - offset);
        const chunk = new Uint8Array(chunkLength);
        const { bytesRead } = await fileHandle.read(chunk, { offset: 0, length: chunkLength, position: offset });

        if (bytesRead === 0) break;

        const endOffset = offset + bytesRead - 1;
        const contentRange = `bytes ${offset}-${endOffset}/${fileSize}`;

        lastSuccessfulResponse = await fetch(uploadUrl, {
          method: 'PUT',
          headers: {
            Authorization: `Bearer ${token}`,
            'Content-Length': String(bytesRead),
            'Content-Range': contentRange
          },
          body: chunk.subarray(0, bytesRead)
        });

        if (!lastSuccessfulResponse.ok) {
          const errorBody = await lastSuccessfulResponse.text().catch(() => '');
          return graphError(
            `Chunk upload failed at offset ${offset} (HTTP ${lastSuccessfulResponse.status}): ${errorBody}`,
            String(lastSuccessfulResponse.status),
            lastSuccessfulResponse.status
          );
        }

        offset += bytesRead;
      }

      // Step 3: Parse the final response — Graph returns the complete driveItem on success
      if (lastSuccessfulResponse && lastSuccessfulResponse.ok) {
        const driveItem = (await lastSuccessfulResponse.json()) as DriveItem;
        return {
          ok: true,
          data: { uploadUrl, expirationDateTime, driveItem }
        };
      }

      return graphError('Upload completed but final response was not valid');
    } finally {
      await fileHandle.close();
    }
  } catch (err) {
    return graphError(err instanceof Error ? err.message : 'Failed to create upload session');
  }
}

export async function downloadFile(
  token: string,
  itemId: string,
  outputPath?: string,
  item?: DriveItem
): Promise<GraphResponse<{ path: string; item: DriveItem }>> {
  let resolvedItem = item;
  let targetPath: string | undefined;
  let tmpPath: string | undefined;

  // Step 1: resolve item metadata
  if (!resolvedItem) {
    const metadata = await getFileMetadata(token, itemId);
    if (!metadata.ok || !metadata.data) {
      return graphError(
        metadata.error?.message || 'Failed to fetch file metadata before download',
        metadata.error?.code,
        metadata.error?.status
      );
    }
    resolvedItem = metadata.data;
  }

  const downloadUrl = resolvedItem['@microsoft.graph.downloadUrl'];
  if (!downloadUrl) {
    return graphError('Download URL missing from Graph metadata response.');
  }

  targetPath = resolve(outputPath || defaultDownloadPath(basename(resolvedItem.name || itemId)));
  await mkdir(dirname(targetPath), { recursive: true });

  // Step 2: retry loop for transient network errors
  const MAX_RETRIES = 2;
  let lastError: unknown;

  for (let attempt = 0; attempt <= MAX_RETRIES; attempt++) {
    try {
      const response = await fetch(downloadUrl);

      if (!response.ok) {
        // Non-transient HTTP errors: don't retry
        return graphError(`Download failed: HTTP ${response.status}`);
      }
      if (!response.body) {
        return graphError('Download failed: response body missing');
      }

      const contentLength = response.headers.get('content-length');
      const tmpFileName = `.${resolvedItem.name ?? itemId}.${randomBytes(8).toString('hex')}.tmp`;
      tmpPath = resolve(dirname(targetPath), 'tmp', tmpFileName);
      await mkdir(dirname(tmpPath), { recursive: true });

      const bytesWritten = await streamWebToFile(response.body, tmpPath);

      // Verify integrity when Content-Length is available
      if (contentLength !== null) {
        const expected = Number(contentLength);
        if (!Number.isFinite(expected)) {
          await unlink(tmpPath).catch(() => {});
          tmpPath = undefined;
          return graphError(`Download failed: invalid Content-Length header`);
        }
        if (bytesWritten !== expected) {
          // Clean up corrupted temp file
          await unlink(tmpPath).catch(() => {});
          tmpPath = undefined;
          return graphError(`Download failed: size mismatch (expected ${expected} bytes, got ${bytesWritten})`);
        }
      }

      // Atomic rename: temp → final path
      await rename(tmpPath, targetPath);

      return graphResult({ path: targetPath, item: resolvedItem });
    } catch (err) {
      lastError = err;

      // Clean up temp file on any error
      if (tmpPath) {
        await unlink(tmpPath).catch(() => {});
        tmpPath = undefined;
      }

      // Only retry on network/stream errors, not on business-logic errors (size mismatch etc.)
      const isRetryable =
        err instanceof Error &&
        (err.message.includes('fetch failed') ||
          err.message.includes('network') ||
          err.message.includes('ECONNREFUSED') ||
          err.message.includes('ETIMEDOUT') ||
          err.message.includes('ENOTFOUND'));

      if (isRetryable && attempt < MAX_RETRIES) {
        continue;
      }
      break;
    }
  }

  return graphError(lastError instanceof Error ? lastError.message : 'Download failed');
}

export async function deleteFile(token: string, itemId: string): Promise<GraphResponse<void>> {
  try {
    return await callGraph<void>(token, `/me/drive/items/${encodeURIComponent(itemId)}`, { method: 'DELETE' }, false);
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to delete file');
  }
}

export async function shareFile(
  token: string,
  itemId: string,
  type: 'view' | 'edit' = 'view',
  scope: 'anonymous' | 'organization' = 'organization'
): Promise<GraphResponse<SharingLinkResult>> {
  let result: GraphResponse<{ link?: SharingLinkResult }>;
  try {
    result = await callGraph<{ link?: SharingLinkResult }>(
      token,
      `/me/drive/items/${encodeURIComponent(itemId)}/createLink`,
      {
        method: 'POST',
        body: JSON.stringify({ type, scope })
      }
    );
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to share file');
  }

  if (!result.ok || !result.data) return result as GraphResponse<SharingLinkResult>;
  return graphResult(result.data.link || {});
}

export async function checkoutFile(token: string, itemId: string): Promise<GraphResponse<void>> {
  try {
    return await callGraph<void>(
      token,
      `/me/drive/items/${encodeURIComponent(itemId)}/checkout`,
      { method: 'POST' },
      false
    );
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to checkout file');
  }
}

export async function checkinFile(
  token: string,
  itemId: string,
  comment?: string
): Promise<GraphResponse<CheckinResult>> {
  let result: GraphResponse<void>;
  try {
    result = await callGraph<void>(
      token,
      `/me/drive/items/${encodeURIComponent(itemId)}/checkin`,
      {
        method: 'POST',
        body: JSON.stringify({ comment: comment || '' })
      },
      false
    );
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Check-in failed');
  }

  if (!result.ok) {
    return graphError(result.error?.message || 'Check-in failed', result.error?.code, result.error?.status);
  }

  const item = await getFileMetadata(token, itemId);
  if (!item.ok || !item.data) {
    return graphError(
      item.error?.message || 'Checked in, but failed to refresh file metadata',
      item.error?.code,
      item.error?.status
    );
  }

  return graphResult({ item: item.data, checkedIn: true, comment });
}

export async function createOfficeCollaborationLink(
  token: string,
  itemId: string,
  options: { lock?: boolean } = {}
): Promise<GraphResponse<OfficeCollabLinkResult>> {
  const item = await getFileMetadata(token, itemId);
  if (!item.ok || !item.data) {
    return graphError(item.error?.message || 'Failed to fetch file metadata', item.error?.code, item.error?.status);
  }

  const extension = item.data.name.includes('.') ? item.data.name.split('.').pop()?.toLowerCase() : undefined;
  const supported = new Set(['docx', 'xlsx', 'pptx']);
  if (!extension || !supported.has(extension)) {
    return graphError(
      'Office Online collaboration is only supported for .docx, .xlsx, and .pptx files. Convert legacy Office formats first.'
    );
  }

  if (options.lock) {
    const lock = await checkoutFile(token, itemId);
    if (!lock.ok) {
      return graphError(
        lock.error?.message || 'Failed to checkout file before sharing',
        lock.error?.code,
        lock.error?.status
      );
    }
  }

  const link = await shareFile(token, itemId, 'edit', 'organization');
  if (!link.ok || !link.data) {
    if (options.lock) {
      await checkinFile(token, itemId);
    }
    return graphError(
      link.error?.message || 'Failed to create collaboration link',
      link.error?.code,
      link.error?.status
    );
  }

  return graphResult({
    item: item.data,
    link: link.data,
    collaborationUrl: item.data.webUrl || link.data.webUrl,
    lockAcquired: !!options.lock
  });
}

export function defaultDownloadPath(fileName: string): string {
  return resolve(homedir(), 'Downloads', basename(fileName));
}

export async function cleanupDownloadedFile(path: string): Promise<void> {
  try {
    await unlink(path);
  } catch {
    // Ignore cleanup failures
  }
}
