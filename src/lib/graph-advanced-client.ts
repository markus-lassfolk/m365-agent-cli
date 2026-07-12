import { resolveGraphAuth } from './graph-auth.js';
import {
  callGraphAt,
  callGraphAtText,
  GraphApiError,
  type GraphRequestInit,
  type GraphResponse,
  graphError,
  graphErrorFromApiError
} from './graph-client.js';
import { getGraphBaseUrl, getGraphBetaUrl } from './graph-constants.js';

/** Parse repeatable CLI `--header "Name: value"` lines (first colon separates name from value). */
export function parseGraphInvokeHeaders(headerLines: string[]): Record<string, string> {
  const out: Record<string, string> = {};
  for (const line of headerLines) {
    const idx = line.indexOf(':');
    if (idx === -1) {
      throw new Error(`Invalid --header format (expected "Name: value"): ${line}`);
    }
    const name = line.slice(0, idx).trim();
    const value = line.slice(idx + 1).trim();
    if (!name) {
      throw new Error(`Invalid --header (empty name): ${line}`);
    }
    out[name] = value;
  }
  return out;
}

/** Reject path traversal and non-relative Graph paths (must be under v1.0/beta root). */
function assertSafeGraphRelativePath(path: string): string {
  const p = path.trim();
  if (!p.startsWith('/')) {
    throw new GraphApiError('Path must start with / (relative to GRAPH_BASE_URL)', 'InvalidPath', 400);
  }
  if (p.length > 8192) {
    throw new GraphApiError('Path exceeds maximum length', 'InvalidPath', 400);
  }
  const q = p.indexOf('?');
  const pathOnly = q === -1 ? p : p.slice(0, q);
  for (const seg of pathOnly.split('/')) {
    if (seg === '.' || seg === '..') {
      throw new GraphApiError('Path must not contain . or .. segments', 'InvalidPath', 400);
    }
  }
  return p;
}

export interface GraphInvokeOptions {
  method: string;
  path: string;
  body?: unknown;
  beta?: boolean;
  expectJson?: boolean;
  extraHeaders?: Record<string, string>;
  /** Used with cached OAuth: on 401, refresh once via `resolveGraphAuth({ identity, forceRefresh: true })`. */
  identity?: string;
  /** When true (e.g. user passed a literal access token), do not attempt refresh on 401. */
  pinAccessToken?: boolean;
}

function graphOnUnauthorizedForCli(
  identity: string | undefined,
  pinAccessToken: boolean | undefined
): (() => Promise<string | null>) | undefined {
  if (pinAccessToken) return undefined;
  return async () => {
    const auth = await resolveGraphAuth({ identity, forceRefresh: true });
    return auth.success && auth.token ? auth.token : null;
  };
}

/**
 * Single request against Microsoft Graph (v1.0 or beta). Path is relative to the API root, e.g. `/me`, `/me/messages?$top=1`.
 */
export async function graphInvoke<T = unknown>(token: string, opts: GraphInvokeOptions): Promise<GraphResponse<T>> {
  try {
    const path = assertSafeGraphRelativePath(opts.path);
    const method = (opts.method || 'GET').toUpperCase();
    const base = opts.beta ? getGraphBetaUrl() : getGraphBaseUrl();
    const expectJson = opts.expectJson !== false;
    const init: GraphRequestInit = { method };
    const hasBody = opts.body !== undefined && method !== 'GET' && method !== 'HEAD';
    if (hasBody) {
      init.body = typeof opts.body === 'string' ? opts.body : JSON.stringify(opts.body);
    }
    if (opts.extraHeaders && Object.keys(opts.extraHeaders).length > 0) {
      init.headers = opts.extraHeaders;
    }
    const onUnauthorized = graphOnUnauthorizedForCli(opts.identity, opts.pinAccessToken === true);
    if (onUnauthorized) {
      init.graphOnUnauthorized = onUnauthorized;
    }
    const r = await callGraphAt<T>(base, token, path, init, expectJson);
    return r;
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphErrorFromApiError(err);
    }
    return graphError(err instanceof Error ? err.message : 'Graph invoke failed');
  }
}

/**
 * Graph invoke where the success body is plain text or SSE (`text/event-stream`), not JSON.
 */
export async function graphInvokeText(token: string, opts: GraphInvokeOptions): Promise<GraphResponse<string>> {
  try {
    const path = assertSafeGraphRelativePath(opts.path);
    const method = (opts.method || 'GET').toUpperCase();
    const base = opts.beta ? getGraphBetaUrl() : getGraphBaseUrl();
    const init: GraphRequestInit = { method };
    const hasBody = opts.body !== undefined && method !== 'GET' && method !== 'HEAD';
    if (hasBody) {
      init.body = typeof opts.body === 'string' ? opts.body : JSON.stringify(opts.body);
    }
    if (opts.extraHeaders && Object.keys(opts.extraHeaders).length > 0) {
      init.headers = opts.extraHeaders;
    }
    const onUnauthorized = graphOnUnauthorizedForCli(opts.identity, opts.pinAccessToken === true);
    if (onUnauthorized) {
      init.graphOnUnauthorized = onUnauthorized;
    }
    return await callGraphAtText(base, token, path, init);
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphErrorFromApiError(err);
    }
    return graphError(err instanceof Error ? err.message : 'Graph invoke failed');
  }
}

export interface GraphBatchRequestBody {
  requests: Array<Record<string, unknown>>;
}

export const GRAPH_BATCH_MAX_REQUESTS = 20;

/** POST `/$batch` ([batch](https://learn.microsoft.com/en-us/graph/json-batching)). Rejects bodies with more than {@link GRAPH_BATCH_MAX_REQUESTS} sub-requests; use `graphBatchAll` to auto-chunk. */
export async function graphPostBatch<T = unknown>(
  token: string,
  body: GraphBatchRequestBody,
  beta?: boolean,
  auth?: { identity?: string; pinAccessToken?: boolean }
): Promise<GraphResponse<T>> {
  try {
    if (!body?.requests || !Array.isArray(body.requests)) {
      return graphError('Body must be a JSON object with a "requests" array', 'InvalidBatch', 400);
    }
    if (body.requests.length > GRAPH_BATCH_MAX_REQUESTS) {
      return graphError(
        `JSON batch supports at most ${GRAPH_BATCH_MAX_REQUESTS} sub-requests per POST. Split into multiple $batch calls, or use graphBatchAll to auto-chunk.`,
        'InvalidBatch',
        400
      );
    }
    const base = beta ? getGraphBetaUrl() : getGraphBaseUrl();
    const init: GraphRequestInit = {
      method: 'POST',
      body: JSON.stringify(body)
    };
    const onUnauthorized = graphOnUnauthorizedForCli(auth?.identity, auth?.pinAccessToken === true);
    if (onUnauthorized) {
      init.graphOnUnauthorized = onUnauthorized;
    }
    return await callGraphAt<T>(base, token, '/$batch', init);
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphErrorFromApiError(err);
    }
    return graphError(err instanceof Error ? err.message : 'Graph batch failed');
  }
}

export interface GraphBatchSubResponse {
  id: string;
  status: number;
  headers?: Record<string, string>;
  body?: unknown;
}

/**
 * Split `requests` into chunks of at most {@link GRAPH_BATCH_MAX_REQUESTS}, in order, without
 * splitting a `dependsOn` chain across chunk boundaries (Graph requires a request and everything
 * it depends on to be in the same `$batch` POST).
 */
export function chunkGraphBatchRequests(
  requests: Array<Record<string, unknown>>,
  size = GRAPH_BATCH_MAX_REQUESTS
): Array<Array<Record<string, unknown>>> {
  const chunks: Array<Array<Record<string, unknown>>> = [];
  for (let i = 0; i < requests.length; i += size) {
    chunks.push(requests.slice(i, i + size));
  }
  return chunks;
}

/**
 * Auto-chunking `$batch`: accepts any number of sub-requests, transparently splits them into
 * `≤ 20`-request POSTs (Graph's per-batch cap), sends them sequentially, and merges the
 * `responses` arrays back into one, in the original request order.
 *
 * Each sub-request needs a unique `id`. If a sub-request's `dependsOn` points at an id that would
 * land in a different chunk, this fails fast with an `InvalidBatch` error before sending anything,
 * since Graph requires dependency chains to stay within a single `$batch` POST — reorder or split
 * such chains into batches of `≤ 20` yourself.
 *
 * If a later chunk's POST fails outright (network error, exhausted retries, ...), the responses
 * already collected from earlier, successfully-sent chunks are NOT discarded: they're returned
 * alongside the error as `data.responses`, so a caller like `applyBulkGraphRequests` can report
 * accurate per-id outcomes instead of leaving the caller to guess whether already-mutated items
 * need to be retried (which could double-apply a non-idempotent mutation).
 */
export async function graphBatchAll(
  token: string,
  requests: Array<Record<string, unknown>>,
  beta?: boolean,
  auth?: { identity?: string; pinAccessToken?: boolean }
): Promise<GraphResponse<{ responses: GraphBatchSubResponse[] }>> {
  if (!Array.isArray(requests)) {
    return graphError('requests must be an array', 'InvalidBatch', 400);
  }
  if (requests.length === 0) {
    return { ok: true, data: { responses: [] } };
  }

  const ids: string[] = [];
  for (const req of requests) {
    const id = req?.id;
    if (typeof id !== 'string' || id.length === 0) {
      return graphError('Every batch sub-request needs a non-empty string "id"', 'InvalidBatch', 400);
    }
    ids.push(id);
  }
  if (new Set(ids).size !== ids.length) {
    return graphError('Batch sub-request "id" values must be unique', 'InvalidBatch', 400);
  }

  const chunks = chunkGraphBatchRequests(requests);
  for (const chunk of chunks) {
    const idsInChunk = new Set(chunk.map((r) => r.id as string));
    for (const req of chunk) {
      const dependsOn = Array.isArray(req.dependsOn) ? (req.dependsOn as unknown[]) : [];
      for (const dep of dependsOn) {
        if (typeof dep !== 'string' || !idsInChunk.has(dep)) {
          return graphError(
            `Batch request "${req.id}" depends on "${String(dep)}", which is not in the same ${GRAPH_BATCH_MAX_REQUESTS}-request chunk. Reorder requests so dependency chains stay together.`,
            'InvalidBatch',
            400
          );
        }
      }
    }
  }

  const responses: GraphBatchSubResponse[] = [];
  for (const chunk of chunks) {
    const r = await graphPostBatch<{ responses: GraphBatchSubResponse[] }>(token, { requests: chunk }, beta, auth);
    if (!r.ok) {
      // Preserve responses already collected from earlier, successfully-sent chunks.
      return { ok: false, error: r.error, data: { responses } };
    }
    responses.push(...(r.data?.responses || []));
  }
  return { ok: true, data: { responses } };
}
