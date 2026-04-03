import {
  callGraph,
  GraphApiError,
  type GraphResponse,
  graphError,
  graphResult
} from './graph-client.js';

/**
 * Microsoft Graph [search query](https://learn.microsoft.com/en-us/graph/api/search-query) (v1.0).
 * Uses entity-specific delegated permissions (e.g. Mail.Read, Files.Read.All, Calendars.Read) per search target — see Microsoft Graph Search API docs.
 */
export interface MicrosoftSearchQueryBody {
  entityTypes: string[];
  queryString: string;
  from?: number;
  size?: number;
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

export async function microsoftSearchQuery(
  token: string,
  body: MicrosoftSearchQueryBody
): Promise<GraphResponse<MicrosoftSearchQueryResponse>> {
  const from = body.from ?? 0;
  const size = Math.min(Math.max(body.size ?? 25, 1), 1000);
  const payload = {
    requests: [
      {
        entityTypes: body.entityTypes,
        query: { queryString: body.queryString },
        from,
        size
      }
    ]
  };
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
