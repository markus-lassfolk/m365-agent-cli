import { randomBytes } from 'node:crypto';
import { createReadStream, createWriteStream } from 'node:fs';
import { mkdir, open, realpath, rename, stat, unlink } from 'node:fs/promises';
import { homedir } from 'node:os';
import { basename, dirname, resolve } from 'node:path';
import { Readable } from 'node:stream';
import type { DriveLocation } from './drive-location.js';
import {
  buildDriveFolderOrRootPath,
  DEFAULT_DRIVE_LOCATION,
  driveDeltaStartPath,
  driveItemPath,
  driveRootSearchPath
} from './drive-location.js';
import { getGraphBaseUrl } from './graph-constants.js';

/** Shown after HTTP 404 on a v1.0-style Graph URL (preview APIs are often beta-only). */
const GRAPH_BETA_404_HINT =
  '\nTip: If this request might be a beta-only Microsoft Graph API, retry with --beta on commands that support it, or set GRAPH_BASE_URL to your Graph beta root (see docs/CLI_REFERENCE.md).';

function graphRequestUrlAppearsToBeV1NotBeta(fullUrl: string): boolean {
  try {
    const { pathname } = new URL(fullUrl);
    if (pathname.includes('/beta/') || pathname.endsWith('/beta')) {
      return false;
    }
    return pathname.includes('/v1.0/') || pathname.endsWith('/v1.0');
  } catch {
    return false;
  }
}

/** Appends a one-line hint for v1.0 404 responses (does not imply 404 always means “use beta”). */
function appendGraphBeta404Hint(fullUrl: string, httpStatus: number, message: string): string {
  if (httpStatus !== 404) return message;
  if (!graphRequestUrlAppearsToBeV1NotBeta(fullUrl)) return message;
  if (message.includes('Tip: If this request might be a beta-only')) return message;
  return message + GRAPH_BETA_404_HINT;
}

function appendGraphBeta404HintForBasePath(
  graphBaseUrl: string,
  relativePath: string,
  httpStatus: number,
  message: string
): string {
  const root = graphBaseUrl.replace(/\/+$/, '');
  const rel = relativePath.startsWith('/') ? relativePath : `/${relativePath}`;
  return appendGraphBeta404Hint(`${root}${rel}`, httpStatus, message);
}

export type { DriveLocation } from './drive-location.js';
export {
  buildDriveFolderOrRootPath,
  DEFAULT_DRIVE_LOCATION,
  driveDeltaStartPath,
  driveItemPath,
  driveLocationFromCliFlags,
  driveRootPrefix,
  driveRootSearchPath
} from './drive-location.js';
export { getGraphBaseUrl, getGraphBetaUrl, graphApiRoot } from './graph-constants.js';

/** Default 60s; override with `GRAPH_TIMEOUT_MS` (milliseconds). */
const GRAPH_TIMEOUT_MS = Number(process.env.GRAPH_TIMEOUT_MS) > 0 ? Number(process.env.GRAPH_TIMEOUT_MS) : 60_000;

/** Optional delay between `@odata.nextLink` pages (milliseconds). Default 0. */
const GRAPH_PAGE_DELAY_MS = Math.max(0, Number(process.env.GRAPH_PAGE_DELAY_MS) || 0);

const GRAPH_RETRY_MAX_ATTEMPTS = Math.min(8, Math.max(1, Number(process.env.GRAPH_MAX_RETRIES) || 4));

const GRAPH_RETRY_MAX_WAIT_MS = Math.max(1000, Number(process.env.GRAPH_RETRY_MAX_WAIT_MS) || 60_000);

export interface GraphInnerError {
  code?: string;
  date?: string;
}

export interface GraphError {
  message: string;
  code?: string;
  status?: number;
  requestId?: string;
  innerError?: GraphInnerError;
}

export class GraphApiError extends Error {
  constructor(
    message: string,
    public readonly code?: string,
    public readonly status?: number,
    public readonly requestId?: string,
    public readonly innerError?: GraphInnerError
  ) {
    super(message);
    this.name = 'GraphApiError';
  }
}

/** Optional extension for `callGraph` / `callGraphAt` options (stripped before `fetch`). */
export type GraphRequestInit = RequestInit & {
  /** If set, on 401 the handler runs once; a returned token triggers a single retry. */
  graphOnUnauthorized?: () => Promise<string | null>;
};

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

export interface DriveItemVersion {
  id: string;
  lastModifiedDateTime?: string;
  size?: number;
  lastModifiedBy?: { user?: { displayName?: string; email?: string } };
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

const MAX_DOWNLOAD_STREAM_BYTES = 5 * 1024 * 1024 * 1024;

/** Streams HTTPS response body to disk with a hard size cap (mitigates unbounded http-to-file writes). */
async function streamWebToFile(
  body: ReadableStream<Uint8Array>,
  filePath: string,
  maxBytes: number = MAX_DOWNLOAD_STREAM_BYTES
): Promise<number> {
  const stream = createWriteStream(filePath, { flags: 'w', mode: 0o600 });
  let bytesWritten = 0;

  try {
    for await (const chunk of body) {
      if (bytesWritten + chunk.byteLength > maxBytes) {
        stream.destroy();
        await unlink(filePath).catch(() => {});
        throw new Error(`Download exceeded maximum size (${maxBytes} bytes)`);
      }
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

export function graphError(
  message: string,
  code?: string,
  status?: number,
  details?: { requestId?: string; innerError?: GraphInnerError }
): GraphResponse<never> {
  return {
    ok: false,
    error: {
      message,
      code,
      status,
      ...(details?.requestId ? { requestId: details.requestId } : {}),
      ...(details?.innerError ? { innerError: details.innerError } : {})
    }
  };
}

export function graphErrorFromApiError(err: GraphApiError): GraphResponse<never> {
  return graphError(err.message, err.code, err.status, {
    requestId: err.requestId,
    innerError: err.innerError
  });
}

function splitGraphRequestInit(options: RequestInit): {
  init: RequestInit;
  onUnauthorized?: () => Promise<string | null>;
} {
  const o = options as GraphRequestInit;
  const { graphOnUnauthorized, ...init } = o;
  return { init, onUnauthorized: graphOnUnauthorized };
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function parseRetryAfterMs(headers: Headers): number | null {
  const ra = headers.get('Retry-After');
  if (!ra) return null;
  const sec = parseInt(ra, 10);
  if (Number.isFinite(sec)) {
    return Math.min(GRAPH_RETRY_MAX_WAIT_MS, Math.max(0, sec) * 1000);
  }
  const when = Date.parse(ra);
  if (Number.isFinite(when)) {
    const delta = when - Date.now();
    return Math.min(GRAPH_RETRY_MAX_WAIT_MS, Math.max(0, delta));
  }
  return null;
}

function throttleBackoffMs(attemptIndex: number): number {
  const base = Math.min(8000, 400 * 2 ** Math.max(0, attemptIndex - 1));
  const jitter = Math.floor(Math.random() * 300);
  return Math.min(GRAPH_RETRY_MAX_WAIT_MS, base + jitter);
}

function isIdempotentMethod(method: string): boolean {
  const m = (method || 'GET').toUpperCase();
  return m === 'GET' || m === 'HEAD';
}

function isTransientNetworkError(err: unknown): boolean {
  if (!(err instanceof Error)) return false;
  return (
    err.name === 'AbortError' ||
    err.message.includes('fetch failed') ||
    err.message.includes('network') ||
    err.message.includes('ECONNRESET') ||
    err.message.includes('ECONNREFUSED') ||
    err.message.includes('ETIMEDOUT') ||
    err.message.includes('ENOTFOUND')
  );
}

interface ParsedGraphFailure {
  message: string;
  code?: string;
  requestId?: string;
  innerError?: GraphInnerError;
}

async function parseGraphFailureResponse(response: Response): Promise<ParsedGraphFailure> {
  const requestId = response.headers.get('request-id') || response.headers.get('client-request-id') || undefined;
  let message = `Graph request failed: HTTP ${response.status}`;
  let code: string | undefined;
  let innerError: GraphInnerError | undefined;
  try {
    const json = (await response.json()) as {
      error?: {
        code?: string;
        message?: string;
        innerError?: { code?: string; date?: string };
      };
    };
    message = json.error?.message || message;
    code = json.error?.code;
    const ie = json.error?.innerError;
    if (ie && (ie.code || ie.date)) {
      innerError = { code: ie.code, date: ie.date };
    }
  } catch {
    // ignore non-JSON body
  }
  return { message, code, requestId, innerError };
}

function shouldRetryThrottle(
  status: number,
  code: string | undefined,
  headers: Headers,
  method: string,
  throttleAttempt: number
): boolean {
  if (throttleAttempt >= GRAPH_RETRY_MAX_ATTEMPTS) return false;
  const idem = isIdempotentMethod(method);
  if (status === 429) {
    if (idem) return true;
    return parseRetryAfterMs(headers) !== null;
  }
  if (status === 503) {
    if (idem) return true;
    return parseRetryAfterMs(headers) !== null;
  }
  if (code === 'tooManyRequests' || code === 'serviceNotAvailable' || code === 'ServiceUnavailable') {
    return idem || parseRetryAfterMs(headers) !== null;
  }
  return false;
}

async function delayBeforeThrottleRetry(headers: Headers, throttleAttempt: number): Promise<void> {
  const ra = parseRetryAfterMs(headers);
  if (ra !== null) {
    await sleep(ra);
    return;
  }
  await sleep(throttleBackoffMs(throttleAttempt));
}

function resolveNextPath(nextLink: string, baseUrl: string): string {
  try {
    const normalizedBase = baseUrl.replace(/\/+$/, '');
    if (nextLink.startsWith(normalizedBase)) {
      return nextLink.substring(normalizedBase.length);
    }
    const nextUrl = new URL(nextLink);
    const baseUrlUrl = new URL(normalizedBase);
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
}

export async function fetchAllPages<T>(
  token: string,
  initialPath: string,
  errorMessage: string,
  baseUrl: string = getGraphBaseUrl(),
  requestInit?: RequestInit
): Promise<GraphResponse<T[]>> {
  const items: T[] = [];
  let path = initialPath;

  while (path) {
    let result: GraphResponse<{ value: T[]; '@odata.nextLink'?: string }>;
    try {
      result = await callGraphAt<{ value: T[]; '@odata.nextLink'?: string }>(baseUrl, token, path, requestInit ?? {});
    } catch (err) {
      if (err instanceof GraphApiError) {
        return graphError(err.message, err.code, err.status, {
          requestId: err.requestId,
          innerError: err.innerError
        }) as GraphResponse<T[]>;
      }
      return graphError(err instanceof Error ? err.message : errorMessage) as GraphResponse<T[]>;
    }
    if (!result.ok || !result.data) {
      return graphError(result.error?.message || errorMessage, result.error?.code, result.error?.status, {
        requestId: result.error?.requestId,
        innerError: result.error?.innerError
      }) as GraphResponse<T[]>;
    }
    items.push(...(result.data.value || []));
    const nextLink = result.data['@odata.nextLink'];
    path = nextLink ? resolveNextPath(nextLink, baseUrl) : '';
    if (path && GRAPH_PAGE_DELAY_MS > 0) {
      await sleep(GRAPH_PAGE_DELAY_MS);
    }
  }
  return graphResult(items);
}

export async function fetchGraphRaw(
  token: string,
  path: string,
  options: RequestInit = {},
  baseUrl: string = getGraphBaseUrl()
): Promise<Response> {
  // codeql[js/file-access-to-http]: Bearer token may come from the local OAuth cache; path is a Graph API path string.
  const { headers: optHeaders, ...rest } = options;
  const headers = new Headers();
  headers.set('Authorization', `Bearer ${token}`);
  if (optHeaders) {
    new Headers(optHeaders).forEach((value, key) => {
      headers.set(key, value);
    });
  }
  const root = baseUrl.replace(/\/+$/, '');
  return fetch(`${root}${path}`, {
    ...rest,
    headers
  });
}

async function callGraphUrlWithRetries<T>(
  fullUrl: string,
  token: string,
  fetchInit: RequestInit,
  expectJson: boolean,
  responseMode: 'json' | 'text',
  onUnauthorized?: () => Promise<string | null>
): Promise<GraphResponse<T>> {
  const method = (fetchInit.method || 'GET').toUpperCase();
  let accessToken = token;
  let throttleAttempt = 0;
  let did401Refresh = false;
  let networkAttempt = 0;

  for (;;) {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), GRAPH_TIMEOUT_MS);
    let response: Response;
    try {
      // codeql[js/file-access-to-http]: Bearer token may come from the local OAuth cache; fullUrl is a Microsoft Graph URL.
      const { headers: existingHeaders, ...fetchRest } = fetchInit;
      const headers = new Headers();
      headers.set('Authorization', `Bearer ${accessToken}`);
      if (responseMode === 'text') {
        headers.set('Accept', '*/*');
      } else if (expectJson) {
        headers.set('Accept', 'application/json');
      }
      const bodyInit = fetchInit.body;
      const isFormData =
        bodyInit !== null && bodyInit !== undefined && typeof FormData !== 'undefined' && bodyInit instanceof FormData;
      if (bodyInit && !(bodyInit instanceof Uint8Array) && !(bodyInit instanceof ArrayBuffer) && !isFormData) {
        headers.set('Content-Type', 'application/json');
      }
      if (existingHeaders) {
        new Headers(existingHeaders).forEach((value, key) => {
          headers.set(key, value);
        });
      }

      // codeql[js/file-access-to-http]: Same Graph client fetch; body is JSON or explicit upload payload from callers, not silent reads of unrelated local files.
      response = await fetch(fullUrl, {
        ...fetchRest,
        method,
        headers,
        signal: controller.signal
      });
    } catch (err) {
      clearTimeout(timeout);
      if (err instanceof Error && err.name === 'AbortError') {
        throw new GraphApiError(`Graph request timed out after ${GRAPH_TIMEOUT_MS / 1000}s`, undefined, 408);
      }
      const idem = isIdempotentMethod(method);
      if (idem && isTransientNetworkError(err) && networkAttempt < GRAPH_RETRY_MAX_ATTEMPTS - 1) {
        networkAttempt++;
        await sleep(throttleBackoffMs(networkAttempt));
        continue;
      }
      throw new GraphApiError(err instanceof Error ? err.message : 'Graph request failed');
    }
    clearTimeout(timeout);

    if (response.status === 401 && onUnauthorized && !did401Refresh) {
      await response.text().catch(() => {});
      const nextTok = await onUnauthorized();
      if (nextTok) {
        accessToken = nextTok;
        did401Refresh = true;
        continue;
      }
      throw new GraphApiError(
        'Microsoft Graph returned 401 and no new access token was obtained.',
        'invalidAuthentication',
        401
      );
    }

    if (!response.ok) {
      const parsed = await parseGraphFailureResponse(response);
      if (shouldRetryThrottle(response.status, parsed.code, response.headers, method, throttleAttempt)) {
        throttleAttempt++;
        await delayBeforeThrottleRetry(response.headers, throttleAttempt);
        continue;
      }
      throw new GraphApiError(
        appendGraphBeta404Hint(fullUrl, response.status, parsed.message),
        parsed.code,
        response.status,
        parsed.requestId,
        parsed.innerError
      );
    }

    if (responseMode === 'text') {
      const text = await response.text();
      return graphResult(text as unknown as T);
    }

    if (!expectJson || response.status === 204) {
      return graphResult(undefined as T);
    }

    const result = await response.json();
    return graphResult(result as T);
  }
}

export async function callGraphAt<T>(
  baseUrl: string,
  token: string,
  path: string,
  options: RequestInit = {},
  expectJson: boolean = true
): Promise<GraphResponse<T>> {
  const normalizedBase = baseUrl.replace(/\/+$/, '');
  const fullUrl = `${normalizedBase}${path}`;
  const { init: fetchInit, onUnauthorized } = splitGraphRequestInit(options);
  return callGraphUrlWithRetries<T>(fullUrl, token, fetchInit, expectJson, 'json', onUnauthorized);
}

/**
 * Like {@link callGraphAt} but always reads a successful response body as UTF-8 text
 * (for `text/event-stream` / SSE and other non-JSON bodies).
 */
export async function callGraphAtText(
  baseUrl: string,
  token: string,
  path: string,
  options: RequestInit = {}
): Promise<GraphResponse<string>> {
  const normalizedBase = baseUrl.replace(/\/+$/, '');
  const fullUrl = `${normalizedBase}${path}`;
  const { init: fetchInit, onUnauthorized } = splitGraphRequestInit(options);
  return callGraphUrlWithRetries<string>(fullUrl, token, fetchInit, true, 'text', onUnauthorized);
}

export async function callGraph<T>(
  token: string,
  path: string,
  options: RequestInit = {},
  expectJson: boolean = true
): Promise<GraphResponse<T>> {
  return callGraphAt(getGraphBaseUrl(), token, path, options, expectJson);
}

function validateGraphUrl(absoluteUrl: string): { valid: boolean; error?: string } {
  let url: URL;
  try {
    url = new URL(absoluteUrl);
  } catch {
    return { valid: false, error: 'Invalid URL format' };
  }

  if (url.protocol !== 'https:') {
    return { valid: false, error: 'Only HTTPS URLs are allowed' };
  }

  const allowedDomains = [
    'graph.microsoft.com',
    'graph.microsoft.us',
    'dod-graph.microsoft.us',
    'microsoftgraph.chinacloudapi.cn',
    'graph.microsoft.de'
  ];

  const isAllowedHost = allowedDomains.some((domain) => url.hostname === domain);

  if (!isAllowedHost) {
    return { valid: false, error: `URL hostname '${url.hostname}' is not a Microsoft Graph endpoint` };
  }

  return { valid: true };
}

/**
 * Drive `copy` async monitor URLs often point at SharePoint / OneDrive hosts, not `graph.microsoft.com`.
 * Same HTTPS + bearer constraints as Graph polling; host allowlist avoids open SSRF from untrusted URLs.
 */
function validateAsyncCopyMonitorUrl(absoluteUrl: string): { valid: boolean; error?: string } {
  let url: URL;
  try {
    url = new URL(absoluteUrl);
  } catch {
    return { valid: false, error: 'Invalid URL format' };
  }
  if (url.protocol !== 'https:') {
    return { valid: false, error: 'Only HTTPS URLs are allowed' };
  }
  const graphCheck = validateGraphUrl(absoluteUrl);
  if (graphCheck.valid) {
    return { valid: true };
  }
  const h = url.hostname.toLowerCase();
  const roots = [
    'sharepoint.com',
    'sharepoint.us',
    'sharepoint.de',
    'sharepoint.cn',
    'sharepoint-mil.us',
    'onedrive.com',
    '1drv.com'
  ];
  for (const root of roots) {
    if (h === root || h.endsWith(`.${root}`)) {
      return { valid: true };
    }
  }
  if (h.endsWith('.1drv.com')) {
    return { valid: true };
  }
  return { valid: false, error: `URL hostname '${url.hostname}' is not allowed for async copy monitor polling` };
}

/** GET/PATCH a full Graph URL (e.g. `@odata.nextLink` / `@odata.deltaLink`). */
export async function callGraphAbsolute<T>(
  token: string,
  absoluteUrl: string,
  options: RequestInit = {},
  expectJson: boolean = true
): Promise<GraphResponse<T>> {
  const validation = validateGraphUrl(absoluteUrl);
  if (!validation.valid) {
    throw new GraphApiError(validation.error || 'Invalid Graph URL', 'InvalidUrl', 400);
  }

  const { init: fetchInit, onUnauthorized } = splitGraphRequestInit(options);
  return callGraphUrlWithRetries<T>(absoluteUrl, token, fetchInit, expectJson, 'json', onUnauthorized);
}

/**
 * Encode a query string for Graph Drive search.
 *
 * encodeURIComponent encodes most characters, but it leaves certain characters
 * like apostrophes and parentheses unescaped. AQS search uses single-quoted
 * strings in the URL path (for example, search(q='...')), so we also percent-
 * encode [!'()*] to prevent query syntax injection and keep the path safe.
 *
 * @param query - Raw search query string
 * @returns Percent-encoded query safe for use in Graph search URLs
 */
function encodeGraphSearchQuery(query: string): string {
  return encodeURIComponent(query).replace(/[!'()*]/g, (char) => `%${char.charCodeAt(0).toString(16).toUpperCase()}`);
}

export async function listFiles(
  token: string,
  folder?: DriveItemReference,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION,
  graphBaseUrl: string = getGraphBaseUrl()
): Promise<GraphResponse<DriveItem[]>> {
  const basePath = buildDriveFolderOrRootPath(location, folder);
  return fetchAllPages<DriveItem>(token, `${basePath}/children`, 'Failed to list files', graphBaseUrl);
}

export async function searchFiles(
  token: string,
  query: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION,
  graphBaseUrl: string = getGraphBaseUrl()
): Promise<GraphResponse<DriveItem[]>> {
  return fetchAllPages<DriveItem>(
    token,
    driveRootSearchPath(location, encodeGraphSearchQuery(query)),
    'Failed to search files',
    graphBaseUrl
  );
}

export async function getFileMetadata(
  token: string,
  itemId: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION,
  graphBaseUrl: string = getGraphBaseUrl()
): Promise<GraphResponse<DriveItem>> {
  try {
    return await callGraphAt<DriveItem>(graphBaseUrl, token, driveItemPath(location, itemId));
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphErrorFromApiError(err);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to get file metadata');
  }
}

export interface UploadLargeResult {
  uploadUrl: string;
  expirationDateTime?: string;
  driveItem?: DriveItem;
}

export async function uploadFile(
  token: string,
  localPath: string,
  folder?: DriveItemReference,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION,
  graphBaseUrl: string = getGraphBaseUrl()
): Promise<GraphResponse<DriveItem>> {
  try {
    const absolutePath = resolve(localPath);
    const st0 = await stat(absolutePath).catch(() => null);
    if (!st0?.isFile()) return graphError(`Not a file or not found: ${absolutePath}`);
    let resolvedPath: string;
    try {
      resolvedPath = await realpath(absolutePath);
    } catch {
      resolvedPath = absolutePath;
    }
    const fileStats = await stat(resolvedPath);
    if (!fileStats.isFile()) return graphError(`Not a file: ${resolvedPath}`);
    if (fileStats.size > 250 * 1024 * 1024) {
      return graphError('File exceeds 250MB simple upload limit. Use upload-large instead.');
    }

    const fileName = basename(resolvedPath);
    const folderPath = `${buildDriveFolderOrRootPath(location, folder)}:/`;
    const stream = createReadStream(resolvedPath);
    try {
      // codeql[js/file-access-to-http]: intentional upload of a user-selected local file to Microsoft Graph after resolve+stat+isFile.
      return await callGraphAt<DriveItem>(
        graphBaseUrl,
        token,
        `${folderPath}${encodeURIComponent(fileName)}:/content`,
        {
          method: 'PUT',
          body: Readable.toWeb(stream) as unknown as BodyInit,
          headers: {
            'Content-Type': 'application/octet-stream'
          }
        }
      );
    } catch (err) {
      if (err instanceof GraphApiError) {
        return graphErrorFromApiError(err);
      }
      return graphError(err instanceof Error ? err.message : 'Upload failed');
    } finally {
      stream.destroy();
    }
  } catch (err) {
    return graphError(err instanceof Error ? err.message : 'Upload failed');
  }
}

export async function uploadLargeFile(
  token: string,
  localPath: string,
  folder?: DriveItemReference,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION,
  graphBaseUrl: string = getGraphBaseUrl()
): Promise<GraphResponse<UploadLargeResult>> {
  try {
    const absolutePath = resolve(localPath);
    const st0 = await stat(absolutePath).catch(() => null);
    if (!st0?.isFile()) return graphError(`Not a file or not found: ${absolutePath}`);
    let resolvedPath: string;
    try {
      resolvedPath = await realpath(absolutePath);
    } catch {
      resolvedPath = absolutePath;
    }
    let fileHandle: any;
    try {
      fileHandle = await open(resolvedPath, 'r');
    } catch (err: any) {
      return graphError(`Failed to open file: ${err.message}`);
    }

    try {
      const fileStats = await fileHandle.stat();
      if (!fileStats.isFile()) return graphError(`Not a file: ${resolvedPath}`);
      if (fileStats.size > 4 * 1024 * 1024 * 1024) {
        return graphError('File exceeds 4GB large upload limit.');
      }

      const fileName = basename(resolvedPath);
      const folderPath = `${buildDriveFolderOrRootPath(location, folder)}:/`;

      // Step 1: Create the upload session
      let sessionResult: GraphResponse<UploadLargeResult>;
      try {
        sessionResult = await callGraphAt<UploadLargeResult>(
          graphBaseUrl,
          token,
          `${folderPath}${encodeURIComponent(fileName)}:/createUploadSession`,
          {
            method: 'POST',
            body: JSON.stringify({ item: { '@microsoft.graph.conflictBehavior': 'replace', name: fileName } })
          }
        );
      } catch (err) {
        if (err instanceof GraphApiError) {
          return graphErrorFromApiError(err);
        }
        return graphError(err instanceof Error ? err.message : 'Failed to create upload session');
      }

      if (!sessionResult.ok || !sessionResult.data) {
        return sessionResult;
      }

      const { uploadUrl, expirationDateTime } = sessionResult.data;

      // Step 2: Upload the file in chunks
      const fileSize = fileStats.size;

      if (fileSize === 0) {
        return graphError('Cannot upload zero-byte files using large upload session. Use simple upload instead.');
      }

      const CHUNK_SIZE = 10 * 1024 * 1024; // 10MB chunks
      const chunkBuffer = new Uint8Array(CHUNK_SIZE);

      let offset = 0;
      let lastSuccessfulResponse: Response | null = null;

      while (offset < fileSize) {
        const chunkLength = Math.min(CHUNK_SIZE, fileSize - offset);
        const { bytesRead } = await fileHandle.read(chunkBuffer, 0, chunkLength, offset);

        if (bytesRead === 0) break;

        const endOffset = offset + bytesRead - 1;
        const contentRange = `bytes ${offset}-${endOffset}/${fileSize}`;
        const chunkData = chunkBuffer.subarray(0, bytesRead);

        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), GRAPH_TIMEOUT_MS);

        try {
          lastSuccessfulResponse = await fetch(uploadUrl, {
            method: 'PUT',
            headers: {
              'Content-Length': String(bytesRead),
              'Content-Range': contentRange
            },
            body: chunkData,
            signal: controller.signal,
            redirect: 'manual'
          });
        } catch (err: any) {
          if (err && err.name === 'AbortError') {
            return graphError(
              `Chunk upload timed out after ${GRAPH_TIMEOUT_MS} ms at offset ${offset}`,
              'RequestTimeout',
              408
            );
          }
          throw err;
        } finally {
          clearTimeout(timeoutId);
        }

        if (!lastSuccessfulResponse.ok) {
          const errorBody = await lastSuccessfulResponse.text().catch(() => '');
          return graphError(
            `Chunk upload failed at offset ${offset} (HTTP ${lastSuccessfulResponse.status}): ${errorBody}`,
            String(lastSuccessfulResponse.status),
            lastSuccessfulResponse.status
          );
        }

        offset += bytesRead;

        if (offset < fileSize) {
          await lastSuccessfulResponse.text().catch(() => {});
        }
      }

      if (offset !== fileSize) {
        return graphError(`Upload stopped early. Expected to upload ${fileSize} bytes but uploaded ${offset}`);
      }

      // Step 3: Parse the final response
      if (lastSuccessfulResponse) {
        const status = lastSuccessfulResponse.status;
        if (status === 200 || status === 201) {
          let body: unknown;
          try {
            body = await lastSuccessfulResponse.json();
          } catch {
            return graphError('Upload completed but failed to parse final response');
          }

          const maybeDriveItem = body as Partial<DriveItem> | null;
          if (
            maybeDriveItem &&
            typeof maybeDriveItem === 'object' &&
            typeof maybeDriveItem.id === 'string' &&
            typeof maybeDriveItem.name === 'string'
          ) {
            const driveItem = maybeDriveItem as DriveItem;
            return {
              ok: true,
              data: { uploadUrl, expirationDateTime, driveItem }
            };
          }

          return graphError('Upload completed but final response did not contain drive item metadata');
        }
      }

      return graphError('Upload completed but final response was not valid');
    } finally {
      await fileHandle.close();
    }
  } catch (err) {
    return graphError(err instanceof Error ? err.message : 'Upload failed');
  }
}

export async function downloadFile(
  token: string,
  itemId: string,
  outputPath?: string,
  item?: DriveItem,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION,
  graphBaseUrl: string = getGraphBaseUrl()
): Promise<GraphResponse<{ path: string; item: DriveItem }>> {
  let resolvedItem = item;
  let targetPath: string | undefined;
  let tmpPath: string | undefined;

  // Step 1: resolve item metadata
  if (!resolvedItem) {
    const metadata = await getFileMetadata(token, itemId, location, graphBaseUrl);
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

  // Security: validate downloadUrl before fetching to prevent SSRF and token exfiltration
  let url: URL;
  try {
    url = new URL(downloadUrl);
  } catch {
    return graphError('Download URL is not a valid URL.');
  }

  if (url.protocol !== 'https:') {
    return graphError('Download URL has unsupported scheme. Only HTTPS is permitted.');
  }

  // Allowed Microsoft domains for download URLs (supports both exact and suffix matching)
  // Includes sovereign cloud domains: .us (GCC High/DoD), .cn (China/21Vianet)
  const allowedDomains = [
    'onedrive.live.com',
    'sharepoint.com',
    'sharepoint.us',
    'sharepoint.cn',
    'graph.microsoft.com',
    'graph.microsoft.us',
    'microsoftgraph.chinacloudapi.cn',
    'files.1drv.com'
  ];

  const isAllowedHost = allowedDomains.some((domain) => url.hostname === domain || url.hostname.endsWith(`.${domain}`));

  if (!isAllowedHost) {
    return graphError(`Download URL hostname '${url.hostname}' is not in the allowlist.`);
  }

  targetPath = resolve(outputPath || defaultDownloadPath(basename(resolvedItem.name || itemId)));
  await mkdir(dirname(targetPath), { recursive: true });

  // Step 2: retry loop for transient network errors
  const MAX_RETRIES = 2;
  let lastError: unknown;

  for (let attempt = 0; attempt <= MAX_RETRIES; attempt++) {
    try {
      const response = await fetch(url.toString(), { redirect: 'manual' });

      // Reject redirects to prevent SSRF bypass
      if (response.status >= 300 && response.status < 400) {
        return graphError('Download failed: redirects are not permitted for security reasons');
      }

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

export async function deleteFile(
  token: string,
  itemId: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION,
  graphBaseUrl: string = getGraphBaseUrl()
): Promise<GraphResponse<void>> {
  try {
    return await callGraphAt<void>(graphBaseUrl, token, driveItemPath(location, itemId), { method: 'DELETE' }, false);
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphErrorFromApiError(err);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to delete file');
  }
}

export async function shareFile(
  token: string,
  itemId: string,
  type: 'view' | 'edit' = 'view',
  scope: 'anonymous' | 'organization' = 'organization',
  location: DriveLocation = DEFAULT_DRIVE_LOCATION,
  graphBaseUrl: string = getGraphBaseUrl()
): Promise<GraphResponse<SharingLinkResult>> {
  let result: GraphResponse<{ link?: SharingLinkResult }>;
  try {
    result = await callGraphAt<{ link?: SharingLinkResult }>(
      graphBaseUrl,
      token,
      `${driveItemPath(location, itemId)}/createLink`,
      {
        method: 'POST',
        body: JSON.stringify({ type, scope })
      }
    );
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphErrorFromApiError(err);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to share file');
  }

  if (!result.ok || !result.data) return result as GraphResponse<SharingLinkResult>;
  return graphResult(result.data.link || {});
}

export async function checkoutFile(
  token: string,
  itemId: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION,
  graphBaseUrl: string = getGraphBaseUrl()
): Promise<GraphResponse<void>> {
  try {
    return await callGraphAt<void>(
      graphBaseUrl,
      token,
      `${driveItemPath(location, itemId)}/checkout`,
      { method: 'POST' },
      false
    );
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphErrorFromApiError(err);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to checkout file');
  }
}

export async function checkinFile(
  token: string,
  itemId: string,
  comment?: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION,
  graphBaseUrl: string = getGraphBaseUrl()
): Promise<GraphResponse<CheckinResult>> {
  let result: GraphResponse<void>;
  try {
    result = await callGraphAt<void>(
      graphBaseUrl,
      token,
      `${driveItemPath(location, itemId)}/checkin`,
      {
        method: 'POST',
        body: JSON.stringify({ comment: comment || '' })
      },
      false
    );
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphErrorFromApiError(err);
    }
    return graphError(err instanceof Error ? err.message : 'Check-in failed');
  }

  if (!result.ok) {
    return graphError(result.error?.message || 'Check-in failed', result.error?.code, result.error?.status);
  }

  const item = await getFileMetadata(token, itemId, location, graphBaseUrl);
  if (!item.ok || !item.data) {
    return graphError(
      item.error?.message || 'Checked in, but failed to refresh file metadata',
      item.error?.code,
      item.error?.status
    );
  }

  return graphResult({ item: item.data, checkedIn: true, comment });
}

/** SharePoint **listItem** for a file in a document library (`GET …/items/{id}/listItem`). Often empty on personal OneDrive. */
export type DriveItemListItem = Record<string, unknown>;

export async function getDriveItemListItem(
  token: string,
  itemId: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION,
  graphBaseUrl: string = getGraphBaseUrl()
): Promise<GraphResponse<DriveItemListItem>> {
  try {
    return await callGraphAt<DriveItemListItem>(graphBaseUrl, token, `${driveItemPath(location, itemId)}/listItem`);
  } catch (err) {
    if (err instanceof GraphApiError) return graphErrorFromApiError(err);
    return graphError(err instanceof Error ? err.message : 'Failed to get listItem');
  }
}

/** OneDrive for Business: follow a file for easy access (`POST …/follow`). */
export async function followDriveItem(
  token: string,
  itemId: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION,
  graphBaseUrl: string = getGraphBaseUrl()
): Promise<GraphResponse<void>> {
  try {
    return await callGraphAt<void>(
      graphBaseUrl,
      token,
      `${driveItemPath(location, itemId)}/follow`,
      { method: 'POST' },
      false
    );
  } catch (err) {
    if (err instanceof GraphApiError) return graphErrorFromApiError(err);
    return graphError(err instanceof Error ? err.message : 'Failed to follow item');
  }
}

export async function unfollowDriveItem(
  token: string,
  itemId: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION,
  graphBaseUrl: string = getGraphBaseUrl()
): Promise<GraphResponse<void>> {
  try {
    return await callGraphAt<void>(
      graphBaseUrl,
      token,
      `${driveItemPath(location, itemId)}/unfollow`,
      { method: 'POST' },
      false
    );
  } catch (err) {
    if (err instanceof GraphApiError) return graphErrorFromApiError(err);
    return graphError(err instanceof Error ? err.message : 'Failed to unfollow item');
  }
}

/** @see https://learn.microsoft.com/en-us/graph/api/driveitem-assignsensitivitylabel */
export async function assignDriveItemSensitivityLabel(
  token: string,
  itemId: string,
  body: Record<string, unknown>,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION,
  graphBaseUrl: string = getGraphBaseUrl()
): Promise<GraphResponse<unknown>> {
  try {
    return await callGraphAt<unknown>(
      graphBaseUrl,
      token,
      `${driveItemPath(location, itemId)}/assignSensitivityLabel`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body)
      }
    );
  } catch (err) {
    if (err instanceof GraphApiError) return graphErrorFromApiError(err);
    return graphError(err instanceof Error ? err.message : 'assignSensitivityLabel failed');
  }
}

/** @see https://learn.microsoft.com/en-us/graph/api/driveitem-extractsensitivitylabels */
export async function extractDriveItemSensitivityLabels(
  token: string,
  itemId: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION,
  graphBaseUrl: string = getGraphBaseUrl()
): Promise<GraphResponse<unknown>> {
  try {
    return await callGraphAt<unknown>(
      graphBaseUrl,
      token,
      `${driveItemPath(location, itemId)}/extractSensitivityLabels`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: '{}'
      }
    );
  } catch (err) {
    if (err instanceof GraphApiError) return graphErrorFromApiError(err);
    return graphError(err instanceof Error ? err.message : 'extractSensitivityLabels failed');
  }
}

/** Permanently delete a drive item (`POST …/permanentDelete`). Irreversible; tenant policies apply. */
export async function permanentDeleteDriveItem(
  token: string,
  itemId: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION,
  graphBaseUrl: string = getGraphBaseUrl()
): Promise<GraphResponse<void>> {
  try {
    return await callGraphAt<void>(
      graphBaseUrl,
      token,
      `${driveItemPath(location, itemId)}/permanentDelete`,
      { method: 'POST' },
      false
    );
  } catch (err) {
    if (err instanceof GraphApiError) return graphErrorFromApiError(err);
    return graphError(err instanceof Error ? err.message : 'permanentDelete failed');
  }
}

/**
 * Restore a deleted drive item from the recycle bin (`POST …/restore`).
 * Optional JSON body (e.g. `parentReference`) per Graph docs; pass `{}` when omitted.
 */
export async function restoreDeletedDriveItem(
  token: string,
  itemId: string,
  body: Record<string, unknown> | undefined,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION,
  graphBaseUrl: string = getGraphBaseUrl()
): Promise<GraphResponse<DriveItem>> {
  try {
    return await callGraphAt<DriveItem>(graphBaseUrl, token, `${driveItemPath(location, itemId)}/restore`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body ?? {})
    });
  } catch (err) {
    if (err instanceof GraphApiError) return graphErrorFromApiError(err);
    return graphError(err instanceof Error ? err.message : 'restore failed');
  }
}

/** @see https://learn.microsoft.com/en-us/graph/api/driveitem-getretentionlabel */
export async function getDriveItemRetentionLabel(
  token: string,
  itemId: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION,
  graphBaseUrl: string = getGraphBaseUrl()
): Promise<GraphResponse<Record<string, unknown>>> {
  try {
    return await callGraphAt<Record<string, unknown>>(
      graphBaseUrl,
      token,
      `${driveItemPath(location, itemId)}/retentionLabel`
    );
  } catch (err) {
    if (err instanceof GraphApiError) return graphErrorFromApiError(err);
    return graphError(err instanceof Error ? err.message : 'Failed to get retentionLabel');
  }
}

/** @see https://learn.microsoft.com/en-us/graph/api/driveitem-removeretentionlabel */
export async function removeDriveItemRetentionLabel(
  token: string,
  itemId: string,
  ifMatch: string | undefined,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION,
  graphBaseUrl: string = getGraphBaseUrl()
): Promise<GraphResponse<void>> {
  try {
    const headers: Record<string, string> = {};
    if (ifMatch?.trim()) headers['If-Match'] = ifMatch.trim();
    return await callGraphAt<void>(
      graphBaseUrl,
      token,
      `${driveItemPath(location, itemId)}/retentionLabel`,
      { method: 'DELETE', headers },
      false
    );
  } catch (err) {
    if (err instanceof GraphApiError) return graphErrorFromApiError(err);
    return graphError(err instanceof Error ? err.message : 'removeRetentionLabel failed');
  }
}

export async function createOfficeCollaborationLink(
  token: string,
  itemId: string,
  options: { lock?: boolean } = {},
  location: DriveLocation = DEFAULT_DRIVE_LOCATION,
  graphBaseUrl: string = getGraphBaseUrl()
): Promise<GraphResponse<OfficeCollabLinkResult>> {
  const item = await getFileMetadata(token, itemId, location, graphBaseUrl);
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
    const lock = await checkoutFile(token, itemId, location, graphBaseUrl);
    if (!lock.ok) {
      return graphError(
        lock.error?.message || 'Failed to checkout file before sharing',
        lock.error?.code,
        lock.error?.status
      );
    }
  }

  const link = await shareFile(token, itemId, 'edit', 'organization', location, graphBaseUrl);
  if (!link.ok || !link.data) {
    if (options.lock) {
      await checkinFile(token, itemId, undefined, location, graphBaseUrl);
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

export async function listFileVersions(
  token: string,
  itemId: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION,
  graphBaseUrl: string = getGraphBaseUrl()
): Promise<GraphResponse<DriveItemVersion[]>> {
  return fetchAllPages<DriveItemVersion>(
    token,
    `${driveItemPath(location, itemId)}/versions`,
    'Failed to list versions',
    graphBaseUrl
  );
}

export async function restoreFileVersion(
  token: string,
  itemId: string,
  versionId: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION,
  graphBaseUrl: string = getGraphBaseUrl()
): Promise<GraphResponse<void>> {
  try {
    return await callGraphAt<void>(
      graphBaseUrl,
      token,
      `${driveItemPath(location, itemId)}/versions/${encodeURIComponent(versionId)}/restoreVersion`,
      { method: 'POST' },
      false
    );
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphErrorFromApiError(err);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to restore version');
  }
}

export async function cleanupDownloadedFile(path: string): Promise<void> {
  try {
    await unlink(path);
  } catch {
    // Ignore cleanup failures
  }
}

export interface FileAnalytics {
  allTime?: {
    access?: { actionCount?: number; actorCount?: number };
  };
  lastSevenDays?: {
    access?: { actionCount?: number; actorCount?: number };
  };
}

export async function getFileAnalytics(
  token: string,
  itemId: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION,
  graphBaseUrl: string = getGraphBaseUrl()
): Promise<GraphResponse<FileAnalytics>> {
  const base = driveItemPath(location, itemId);
  const [allTimeResult, lastSevenDaysResult] = await Promise.allSettled([
    callGraphAt<FileAnalytics['allTime']>(graphBaseUrl, token, `${base}/analytics/allTime`),
    callGraphAt<FileAnalytics['lastSevenDays']>(graphBaseUrl, token, `${base}/analytics/lastSevenDays`)
  ]);

  const analytics: FileAnalytics = {};

  if (allTimeResult.status === 'fulfilled' && allTimeResult.value.ok && allTimeResult.value.data) {
    analytics.allTime = allTimeResult.value.data;
  }

  if (lastSevenDaysResult.status === 'fulfilled' && lastSevenDaysResult.value.ok && lastSevenDaysResult.value.data) {
    analytics.lastSevenDays = lastSevenDaysResult.value.data;
  }

  if (allTimeResult.status === 'rejected' && lastSevenDaysResult.status === 'rejected') {
    const error = allTimeResult.reason;
    if (error instanceof GraphApiError) {
      return graphError(error.message, error.code, error.status);
    }
    return graphError(error instanceof Error ? error.message : 'Failed to get file analytics');
  }

  return graphResult(analytics);
}

export async function downloadConvertedFile(
  token: string,
  itemId: string,
  format: string = 'pdf',
  outputPath?: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION,
  graphBaseUrl: string = getGraphBaseUrl()
): Promise<GraphResponse<{ path: string }>> {
  let tmpPath: string | undefined;
  try {
    const metadata = await getFileMetadata(token, itemId, location, graphBaseUrl);
    if (!metadata.ok || !metadata.data) {
      return graphError(
        metadata.error?.message || 'Failed to fetch file metadata',
        metadata.error?.code,
        metadata.error?.status
      );
    }

    const item = metadata.data;
    const originalName = item.name || itemId;
    const newName = originalName.includes('.')
      ? `${originalName.substring(0, originalName.lastIndexOf('.'))}.${format}`
      : `${originalName}.${format}`;

    const targetPath = resolve(outputPath || defaultDownloadPath(newName));
    await mkdir(dirname(targetPath), { recursive: true });

    const contentPath = `${driveItemPath(location, itemId)}/content?format=${encodeURIComponent(format)}`;

    const redirectResponse = await fetchGraphRaw(token, contentPath, { redirect: 'manual' }, graphBaseUrl);

    if (redirectResponse.status < 300 || redirectResponse.status >= 400) {
      if (!redirectResponse.ok) {
        return graphError(`Failed to convert file: HTTP ${redirectResponse.status}`);
      }
      return graphError('Expected a redirect for file conversion, but got a direct response.');
    }

    const redirectLocation = redirectResponse.headers.get('location');
    if (!redirectLocation) {
      return graphError('Missing redirect location for converted file');
    }

    let url: URL;
    try {
      url = new URL(redirectLocation);
    } catch {
      return graphError('Redirect location is not a valid URL.');
    }

    if (url.protocol !== 'https:') {
      return graphError('Redirect URL has unsupported scheme. Only HTTPS is permitted.');
    }

    const allowedDomains = [
      'onedrive.live.com',
      'sharepoint.com',
      'sharepoint.us',
      'sharepoint.cn',
      'graph.microsoft.com',
      'graph.microsoft.us',
      'microsoftgraph.chinacloudapi.cn',
      'files.1drv.com'
    ];

    const isAllowedHost = allowedDomains.some(
      (domain) => url.hostname === domain || url.hostname.endsWith(`.${domain}`)
    );

    if (!isAllowedHost) {
      return graphError(`Redirect URL hostname '${url.hostname}' is not in the allowlist.`);
    }

    const response = await fetch(url.toString(), { redirect: 'manual' });

    if (response.status >= 300 && response.status < 400) {
      return graphError('Download failed: further redirects are not permitted for security reasons');
    }

    if (!response.ok) {
      return graphError(`Failed to download converted file: HTTP ${response.status}`);
    }
    if (!response.body) {
      return graphError('Response body is empty');
    }

    const tmpFileName = `.${newName}.${randomBytes(8).toString('hex')}.tmp`;
    tmpPath = resolve(dirname(targetPath), 'tmp', tmpFileName);
    await mkdir(dirname(tmpPath), { recursive: true });

    await streamWebToFile(response.body, tmpPath);
    await rename(tmpPath, targetPath);

    return graphResult({ path: targetPath });
  } catch (err) {
    if (tmpPath) {
      await unlink(tmpPath).catch(() => {});
    }
    if (err instanceof GraphApiError) {
      return graphErrorFromApiError(err);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to download converted file');
  }
}

/** One sharing permission entry on a drive item (shape varies by link type / recipient). */
export type DriveItemPermission = Record<string, unknown>;

export async function inviteDriveItem(
  token: string,
  itemId: string,
  body: Record<string, unknown>,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION,
  graphBaseUrl: string = getGraphBaseUrl()
): Promise<GraphResponse<unknown>> {
  try {
    return await callGraphAt<unknown>(graphBaseUrl, token, `${driveItemPath(location, itemId)}/invite`, {
      method: 'POST',
      body: JSON.stringify(body)
    });
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphErrorFromApiError(err);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to send invite');
  }
}

export async function listDriveItemPermissions(
  token: string,
  itemId: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION,
  graphBaseUrl: string = getGraphBaseUrl()
): Promise<GraphResponse<DriveItemPermission[]>> {
  return fetchAllPages<DriveItemPermission>(
    token,
    `${driveItemPath(location, itemId)}/permissions`,
    'Failed to list permissions',
    graphBaseUrl
  );
}

export async function deleteDriveItemPermission(
  token: string,
  itemId: string,
  permissionId: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION,
  graphBaseUrl: string = getGraphBaseUrl()
): Promise<GraphResponse<void>> {
  try {
    return await callGraphAt<void>(
      graphBaseUrl,
      token,
      `${driveItemPath(location, itemId)}/permissions/${encodeURIComponent(permissionId)}`,
      { method: 'DELETE' },
      false
    );
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphErrorFromApiError(err);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to delete permission');
  }
}

/** One size entry in a [thumbnailSet](https://learn.microsoft.com/en-us/graph/api/resources/thumbnailset). */
export interface ThumbnailSizeInfo {
  height?: number;
  width?: number;
  url?: string;
}

export interface DriveItemThumbnailSet {
  id?: string;
  large?: ThumbnailSizeInfo;
  medium?: ThumbnailSizeInfo;
  small?: ThumbnailSizeInfo;
  [key: string]: unknown;
}

/** List generated thumbnails for a drive item (`GET …/items/{id}/thumbnails`). */
export async function listDriveItemThumbnails(
  token: string,
  itemId: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION,
  graphBaseUrl: string = getGraphBaseUrl()
): Promise<GraphResponse<DriveItemThumbnailSet[]>> {
  try {
    const r = await callGraphAt<{ value?: DriveItemThumbnailSet[] }>(
      graphBaseUrl,
      token,
      `${driveItemPath(location, itemId)}/thumbnails`
    );
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to list thumbnails', r.error?.code, r.error?.status);
    }
    return graphResult(r.data.value ?? []);
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphErrorFromApiError(err);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to list thumbnails');
  }
}

/** One page of drive item delta (root or folder); use `@odata.nextLink` / `@odata.deltaLink` for follow-up. */
export interface DriveDeltaPage {
  value?: DriveItem[];
  '@odata.nextLink'?: string;
  '@odata.deltaLink'?: string;
}

export async function getDriveItemDeltaPage(
  token: string,
  options: {
    location: DriveLocation;
    folderItemId?: string;
    nextOrDeltaLink?: string;
    graphBaseUrl?: string;
  }
): Promise<GraphResponse<DriveDeltaPage>> {
  const graphBaseUrl = options.graphBaseUrl ?? getGraphBaseUrl();
  try {
    if (options.nextOrDeltaLink?.trim()) {
      return await callGraphAbsolute<DriveDeltaPage>(token, options.nextOrDeltaLink.trim());
    }
    return await callGraphAt<DriveDeltaPage>(
      graphBaseUrl,
      token,
      driveDeltaStartPath(options.location, options.folderItemId)
    );
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphErrorFromApiError(err);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to fetch drive delta');
  }
}

/** Items shared with the signed-in user (`GET /me/drive/sharedWithMe`). Remote drive locations are not supported by Graph for this collection. */
export interface SharedWithMeDriveItem {
  id?: string;
  name?: string;
  remoteItem?: { id?: string; name?: string; parentReference?: { driveId?: string } };
  webUrl?: string;
  [key: string]: unknown;
}

export async function listDriveSharedWithMe(
  token: string,
  graphBaseUrl: string = getGraphBaseUrl()
): Promise<GraphResponse<{ value?: SharedWithMeDriveItem[] }>> {
  try {
    return await callGraphAt<{ value?: SharedWithMeDriveItem[] }>(graphBaseUrl, token, '/me/drive/sharedWithMe');
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphErrorFromApiError(err);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to list shared-with-me');
  }
}

export interface CopyDriveItemBody {
  parentReference: { id: string; driveId?: string };
  name?: string;
}

/** Starts async copy; returns monitor URL from `Location` when Graph returns 202. */
export async function startCopyDriveItem(
  token: string,
  itemId: string,
  body: CopyDriveItemBody,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION,
  graphBaseUrl: string = getGraphBaseUrl()
): Promise<GraphResponse<{ monitorUrl?: string; status?: number }>> {
  try {
    const res = await fetchGraphRaw(
      token,
      `${driveItemPath(location, itemId)}/copy`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body)
      },
      graphBaseUrl
    );
    if (res.status === 202) {
      const monitorUrl = res.headers.get('location') ?? undefined;
      return graphResult({ monitorUrl, status: 202 });
    }
    if (res.ok) {
      const text = await res.text();
      try {
        return graphResult(JSON.parse(text) as { monitorUrl?: string; status?: number });
      } catch {
        return graphResult({ status: res.status });
      }
    }
    const parsed = await parseGraphFailureResponse(res);
    const relPath = `${driveItemPath(location, itemId)}/copy`;
    throw new GraphApiError(
      appendGraphBeta404HintForBasePath(graphBaseUrl, relPath, res.status, parsed.message),
      parsed.code,
      res.status,
      parsed.requestId,
      parsed.innerError
    );
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphErrorFromApiError(err);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to start copy');
  }
}

export interface AsyncJobStatus {
  status?: string;
  resourceId?: string;
  resourceLocation?: string;
  [key: string]: unknown;
}

/** Poll Graph async job URL (copy monitor) until completed, failed, or timeout. */
export async function pollGraphAsyncJob(
  token: string,
  monitorUrl: string,
  options: { maxAttempts?: number; delayMs?: number } = {}
): Promise<GraphResponse<AsyncJobStatus>> {
  const max = options.maxAttempts ?? 45;
  const delayMs = options.delayMs ?? 2000;
  const validation = validateAsyncCopyMonitorUrl(monitorUrl);
  if (!validation.valid) {
    return graphError(validation.error || 'Invalid monitor URL');
  }
  try {
    for (let attempt = 0; attempt < max; attempt++) {
      if (attempt > 0) await sleep(delayMs);
      const res = await fetch(monitorUrl, { headers: { Authorization: `Bearer ${token}` } });
      const data = (await res.json()) as AsyncJobStatus;
      if (!res.ok) {
        return graphError(`Monitor request failed: HTTP ${res.status}`);
      }
      const st = data.status?.toLowerCase();
      if (st === 'completed' || st === 'succeeded') {
        return graphResult(data);
      }
      if (st === 'failed' || st === 'cancelled') {
        return graphError(data.error ? JSON.stringify(data.error) : `Async job ${st ?? 'failed'}`);
      }
    }
    return graphError('Async copy job timed out while polling');
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphErrorFromApiError(err);
    }
    return graphError(err instanceof Error ? err.message : 'Poll failed');
  }
}

export async function moveDriveItem(
  token: string,
  itemId: string,
  parentReference: { id: string; driveId?: string },
  location: DriveLocation = DEFAULT_DRIVE_LOCATION,
  graphBaseUrl: string = getGraphBaseUrl()
): Promise<GraphResponse<DriveItem>> {
  try {
    return await callGraphAt<DriveItem>(graphBaseUrl, token, driveItemPath(location, itemId), {
      method: 'PATCH',
      body: JSON.stringify({ parentReference })
    });
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphErrorFromApiError(err);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to move item');
  }
}

export async function patchDriveItemPermission(
  token: string,
  itemId: string,
  permissionId: string,
  body: Record<string, unknown>,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION,
  graphBaseUrl: string = getGraphBaseUrl()
): Promise<GraphResponse<DriveItemPermission>> {
  try {
    return await callGraphAt<DriveItemPermission>(
      graphBaseUrl,
      token,
      `${driveItemPath(location, itemId)}/permissions/${encodeURIComponent(permissionId)}`,
      { method: 'PATCH', body: JSON.stringify(body) }
    );
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphErrorFromApiError(err);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to update permission');
  }
}
