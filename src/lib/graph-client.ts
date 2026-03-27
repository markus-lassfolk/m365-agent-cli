import { basename, resolve } from 'node:path';
import { mkdir, readFile, stat, unlink, writeFile } from 'node:fs/promises';
import { homedir } from 'node:os';

const GRAPH_BASE_URL = process.env.GRAPH_BASE_URL || 'https://graph.microsoft.com/v1.0';

export interface GraphError {
  message: string;
  code?: string;
  status?: number;
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

export interface UploadLargeResult {
  uploadUrl: string;
  expirationDateTime?: string;
}

function graphResult<T>(data: T): GraphResponse<T> {
  return { ok: true, data };
}

function graphError(message: string, code?: string, status?: number): GraphResponse<never> {
  return { ok: false, error: { message, code, status } };
}

async function callGraph<T>(
  token: string,
  path: string,
  options: RequestInit = {},
  expectJson: boolean = true
): Promise<GraphResponse<T>> {
  try {
    const response = await fetch(`${GRAPH_BASE_URL}${path}`, {
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

    if (!response.ok) {
      let message = `Graph request failed: HTTP ${response.status}`;
      let code: string | undefined;
      try {
        const json = (await response.json()) as { error?: { code?: string; message?: string } };
        message = json.error?.message || message;
        code = json.error?.code;
      } catch {
        // Ignore JSON parse failures for error responses
      }
      return graphError(message, code, response.status);
    }

    if (!expectJson || response.status === 204) {
      return graphResult(undefined as T);
    }

    return graphResult((await response.json()) as T);
  } catch (err) {
    return graphError(err instanceof Error ? err.message : 'Graph request failed');
  }
}

function buildItemPath(reference?: DriveItemReference): string {
  if (!reference?.id) return '/me/drive/root';

  const drivePrefix = reference.driveId ? `/drives/${encodeURIComponent(reference.driveId)}` : '/me/drive';
  return `${drivePrefix}/items/${encodeURIComponent(reference.id)}`;
}

export async function listFiles(token: string, folder?: DriveItemReference): Promise<GraphResponse<DriveItem[]>> {
  const basePath = buildItemPath(folder);
  const result = await callGraph<DriveItemListResponse>(token, `${basePath}/children`);
  if (!result.ok || !result.data) {
    return graphError(result.error?.message || 'Failed to list files', result.error?.code, result.error?.status);
  }
  return graphResult(result.data.value || []);
}

export async function searchFiles(token: string, query: string): Promise<GraphResponse<DriveItem[]>> {
  const encodedQuery = encodeURIComponent(query)
    .replace(/%20/g, ' ')
    .replace(/'/g, '%27');
  const result = await callGraph<DriveItemListResponse>(
    token,
    `/me/drive/root/search(q='${encodedQuery}')`
  );
  if (!result.ok || !result.data) {
    return graphError(result.error?.message || 'Failed to search files', result.error?.code, result.error?.status);
  }
  return graphResult(result.data.value || []);
}

export async function getFileMetadata(token: string, itemId: string): Promise<GraphResponse<DriveItem>> {
  return callGraph<DriveItem>(token, `/me/drive/items/${encodeURIComponent(itemId)}`);
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

    const content = await readFile(absolutePath);
    const fileName = basename(absolutePath);
    const folderPath = folder?.id ? `${buildItemPath(folder)}:/` : '/me/drive/root:/';
    const result = await callGraph<DriveItem>(token, `${folderPath}${encodeURIComponent(fileName)}:/content`, {
      method: 'PUT',
      body: content,
      headers: {
        'Content-Type': 'application/octet-stream'
      }
    });

    return result;
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
    const result = await callGraph<UploadLargeResult>(
      token,
      `${folderPath}${encodeURIComponent(fileName)}:/createUploadSession`,
      {
        method: 'POST',
        body: JSON.stringify({ item: { '@microsoft.graph.conflictBehavior': 'replace', name: fileName } })
      }
    );

    return result;
  } catch (err) {
    return graphError(err instanceof Error ? err.message : 'Failed to create upload session');
  }
}

export async function downloadFile(
  token: string,
  itemId: string,
  outputPath?: string,
  metadata?: DriveItem
): Promise<GraphResponse<{ path: string; item: DriveItem }>> {
  try {
    let itemMetadata: DriveItem;
    if (metadata) {
      itemMetadata = metadata;
    } else {
      const metadataResult = await getFileMetadata(token, itemId);
      if (!metadataResult.ok || !metadataResult.data) {
        return graphError(
          metadataResult.error?.message || 'Failed to fetch file metadata before download',
          metadataResult.error?.code,
          metadataResult.error?.status
        );
      }
      itemMetadata = metadataResult.data;
    }

    const downloadUrl = itemMetadata['@microsoft.graph.downloadUrl'];
    if (!downloadUrl) {
      return graphError('Download URL missing from Graph metadata response.');
    }

    const response = await fetch(downloadUrl);
    if (!response.ok) {
      return graphError(`Download failed: HTTP ${response.status}`);
    }

    const bytes = new Uint8Array(await response.arrayBuffer());
    const targetPath = resolve(outputPath || itemMetadata.name || itemId);
    await mkdir(resolve(targetPath, '..'), { recursive: true }).catch(() => undefined);
    await writeFile(targetPath, bytes);

    return graphResult({ path: targetPath, item: itemMetadata });
  } catch (err) {
    return graphError(err instanceof Error ? err.message : 'Download failed');
  }
}

export async function deleteFile(token: string, itemId: string): Promise<GraphResponse<void>> {
  return callGraph<void>(token, `/me/drive/items/${encodeURIComponent(itemId)}`, { method: 'DELETE' }, false);
}

export async function shareFile(
  token: string,
  itemId: string,
  type: 'view' | 'edit' = 'view',
  scope: 'anonymous' | 'organization' = 'organization'
): Promise<GraphResponse<SharingLinkResult>> {
  const result = await callGraph<{ link?: SharingLinkResult }>(
    token,
    `/me/drive/items/${encodeURIComponent(itemId)}/createLink`,
    {
      method: 'POST',
      body: JSON.stringify({ type, scope })
    }
  );

  if (!result.ok || !result.data) return result as GraphResponse<SharingLinkResult>;
  return graphResult(result.data.link || {});
}

export function defaultDownloadPath(fileName: string): string {
  return resolve(homedir(), 'Downloads', fileName);
}

export async function cleanupDownloadedFile(path: string): Promise<void> {
  try {
    await unlink(path);
  } catch {
    // Ignore cleanup failures
  }
}
