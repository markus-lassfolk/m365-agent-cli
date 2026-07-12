import { callGraph, GraphApiError, type GraphResponse, graphError, graphResult } from './graph-client.js';

/**
 * Microsoft Graph [search query](https://learn.microsoft.com/en-us/graph/api/search-query) (v1.0).
 * Uses entity-specific delegated permissions (e.g. Mail.Read, Files.Read.All, Calendars.Read) per search target — see Microsoft Graph Search API docs.
 */
export interface MicrosoftSearchQueryBody {
  entityTypes: string[];
  queryString: string;
  from?: number;
  size?: number;
  /** Merged into the built `searchRequest` after base fields (deep merge for nested plain objects). */
  requestPatch?: Record<string, unknown>;
}

/** Raw API response shape (subset). */
export interface MicrosoftSearchQueryResponse {
  value?: Array<{
    searchTerms?: string[];
    hitsContainers?: Array<{
      hits?: Array<{
        hitId?: string;
        rank?: number;
        summary?: string;
        resource?: Record<string, unknown>;
      }>;
      total?: number;
      moreResultsAvailable?: boolean;
    }>;
  }>;
}

/** Flattened hit for agents (`graph-search --json-hits`). */
export interface NormalizedSearchHit {
  rank?: number;
  hitId?: string;
  summary?: string;
  entityType?: string;
  id?: string;
  webUrl?: string;
  name?: string;
  title?: string;
  subject?: string;
}

/** Deep-merge plain objects for `searchRequest` overlays (arrays replace; nested objects merge). */
export function deepMergeSearchRequest(
  base: Record<string, unknown>,
  overlay: Record<string, unknown>
): Record<string, unknown> {
  const out: Record<string, unknown> = { ...base };
  for (const [k, v] of Object.entries(overlay)) {
    // Never merge prototype-polluting keys from user-supplied JSON.
    if (k === '__proto__' || k === 'constructor' || k === 'prototype') continue;
    if (
      v !== null &&
      typeof v === 'object' &&
      !Array.isArray(v) &&
      out[k] !== null &&
      typeof out[k] === 'object' &&
      !Array.isArray(out[k])
    ) {
      out[k] = deepMergeSearchRequest(out[k] as Record<string, unknown>, v as Record<string, unknown>);
    } else {
      out[k] = v;
    }
  }
  return out;
}

/** Stable projection of Microsoft Search hits (no OData noise). */
export function flattenMicrosoftSearchHits(response: MicrosoftSearchQueryResponse): NormalizedSearchHit[] {
  const hits: NormalizedSearchHit[] = [];
  for (const block of response.value ?? []) {
    for (const c of block.hitsContainers ?? []) {
      for (const h of c.hits ?? []) {
        const r = h.resource ?? {};
        const odataType = r['@odata.type'];
        hits.push({
          rank: h.rank,
          hitId: h.hitId,
          summary: h.summary,
          entityType: typeof odataType === 'string' ? odataType.replace(/^#microsoft\.graph\./, '') : undefined,
          id: typeof r.id === 'string' ? r.id : undefined,
          webUrl: typeof r.webUrl === 'string' ? r.webUrl : undefined,
          name: typeof r.name === 'string' ? r.name : undefined,
          title: typeof r.title === 'string' ? r.title : undefined,
          subject: typeof r.subject === 'string' ? r.subject : undefined
        });
      }
    }
  }
  return hits;
}

export function buildMicrosoftSearchRequest(body: MicrosoftSearchQueryBody): Record<string, unknown> {
  const from = body.from ?? 0;
  const size = Math.min(Math.max(body.size ?? 25, 1), 1000);
  const base: Record<string, unknown> = {
    entityTypes: body.entityTypes,
    query: { queryString: body.queryString },
    from,
    size
  };
  if (body.requestPatch && Object.keys(body.requestPatch).length > 0) {
    return deepMergeSearchRequest(base, body.requestPatch);
  }
  return base;
}

/** Full `POST /search/query` body: `{ requests: [ searchRequest, … ] }`. */
export async function microsoftSearchQueryRaw(
  token: string,
  payload: { requests: unknown[] }
): Promise<GraphResponse<MicrosoftSearchQueryResponse>> {
  if (!Array.isArray(payload.requests) || payload.requests.length === 0) {
    return graphError('JSON body must include a non-empty "requests" array');
  }
  try {
    const result = await callGraph<MicrosoftSearchQueryResponse>(token, '/search/query', {
      method: 'POST',
      body: JSON.stringify(payload)
    });
    if (!result.ok || !result.data) {
      return graphError(
        result.error?.message || 'Microsoft Search request failed',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Microsoft Search request failed');
  }
}

export async function microsoftSearchQuery(
  token: string,
  body: MicrosoftSearchQueryBody
): Promise<GraphResponse<MicrosoftSearchQueryResponse>> {
  const request = buildMicrosoftSearchRequest(body);
  return microsoftSearchQueryRaw(token, { requests: [request] });
}
