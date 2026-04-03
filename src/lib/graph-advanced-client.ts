import { callGraphAt, GraphApiError, type GraphResponse, graphError } from './graph-client.js';
import { GRAPH_BASE_URL, GRAPH_BETA_URL } from './graph-constants.js';

/** Reject path traversal and non-relative Graph paths (must be under v1.0/beta root). */
export function assertSafeGraphRelativePath(path: string): string {
  const p = path.trim();
  if (!p.startsWith('/')) {
    throw new GraphApiError('Path must start with / (relative to GRAPH_BASE_URL)', 'InvalidPath', 400);
  }
  if (p.includes('..')) {
    throw new GraphApiError('Path must not contain ..', 'InvalidPath', 400);
  }
  if (p.length > 8192) {
    throw new GraphApiError('Path exceeds maximum length', 'InvalidPath', 400);
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
}

/**
 * Single request against Microsoft Graph (v1.0 or beta). Path is relative to the API root, e.g. `/me`, `/me/messages?$top=1`.
 */
export async function graphInvoke<T = unknown>(token: string, opts: GraphInvokeOptions): Promise<GraphResponse<T>> {
  try {
    const path = assertSafeGraphRelativePath(opts.path);
    const method = (opts.method || 'GET').toUpperCase();
    const base = opts.beta ? GRAPH_BETA_URL : GRAPH_BASE_URL;
    const expectJson = opts.expectJson !== false;
    const init: RequestInit = { method };
    const hasBody = opts.body !== undefined && method !== 'GET' && method !== 'HEAD';
    if (hasBody) {
      init.body = typeof opts.body === 'string' ? opts.body : JSON.stringify(opts.body);
    }
    if (opts.extraHeaders && Object.keys(opts.extraHeaders).length > 0) {
      init.headers = opts.extraHeaders;
    }
    const r = await callGraphAt<T>(base, token, path, init, expectJson);
    return r;
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Graph invoke failed');
  }
}

export interface GraphBatchRequestBody {
  requests: Array<Record<string, unknown>>;
}

/** POST `/$batch` ([batch](https://learn.microsoft.com/en-us/graph/json-batching)). */
export async function graphPostBatch<T = unknown>(
  token: string,
  body: GraphBatchRequestBody,
  beta?: boolean
): Promise<GraphResponse<T>> {
  try {
    if (!body?.requests || !Array.isArray(body.requests)) {
      return graphError('Body must be a JSON object with a "requests" array', 'InvalidBatch', 400);
    }
    const base = beta ? GRAPH_BETA_URL : GRAPH_BASE_URL;
    return await callGraphAt<T>(base, token, '/$batch', {
      method: 'POST',
      body: JSON.stringify(body)
    });
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Graph batch failed');
  }
}
