import {
  callGraph,
  fetchAllPages,
  GraphApiError,
  type GraphResponse,
  graphResult,
  graphError
} from './graph-client.js';

export interface SharePointList {
  id: string;
  name: string;
  displayName: string;
  description?: string;
  createdDateTime: string;
  lastModifiedDateTime: string;
  webUrl: string;
}

export interface SharePointListItem {
  id: string;
  createdDateTime: string;
  lastModifiedDateTime: string;
  webUrl: string;
  fields: Record<string, any>;
}

export async function getLists(token: string, siteId: string): Promise<GraphResponse<SharePointList[]>> {
  let res: GraphResponse<{ value: SharePointList[] }>;
  try {
    res = await callGraph<{ value: SharePointList[] }>(token, `/sites/${siteId}/lists`);
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to get lists');
  }
  if (!res.ok) {
    return graphError(res.error?.message ?? 'Failed to get lists', res.error?.code, res.error?.status);
  }
  if (!res.data || !Array.isArray(res.data.value)) {
    return graphError('Failed to get lists');
  }
  return graphResult(res.data.value);
}

export async function getListItems(
  token: string,
  siteId: string,
  listId: string
): Promise<GraphResponse<SharePointListItem[]>> {
  return fetchAllPages<SharePointListItem>(
    token,
    `/sites/${siteId}/lists/${encodeURIComponent(listId)}/items?$expand=fields`,
    'Failed to get list items'
  );
}

export async function createListItem(
  token: string,
  siteId: string,
  listId: string,
  fields: Record<string, any>
): Promise<GraphResponse<SharePointListItem>> {
  try {
    return await callGraph<SharePointListItem>(token, `/sites/${siteId}/lists/${encodeURIComponent(listId)}/items`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ fields })
    });
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to create list item');
  }
}

export async function updateListItem(
  token: string,
  siteId: string,
  listId: string,
  itemId: string,
  fields: Record<string, any>
): Promise<GraphResponse<Record<string, any>>> {
  try {
    return await callGraph<Record<string, any>>(
      token,
      `/sites/${siteId}/lists/${encodeURIComponent(listId)}/items/${encodeURIComponent(itemId)}/fields`,
      {
        method: 'PATCH',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(fields)
      }
    );
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to update list item');
  }
}
