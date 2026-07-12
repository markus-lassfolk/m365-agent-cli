import {
  callGraphAbsolute,
  callGraphAt,
  fetchAllPages,
  GraphApiError,
  type GraphResponse,
  getGraphBaseUrl,
  graphError,
  graphResult
} from './graph-client.js';

export interface SharePointSiteSummary {
  id: string;
  displayName?: string;
  webUrl?: string;
  name?: string;
}

export async function getSiteByGraphPath(
  token: string,
  /** e.g. `contoso.sharepoint.com:/sites/TeamName` (host + `:` + server-relative path) */
  sitePath: string,
  apiBase: string = getGraphBaseUrl()
): Promise<GraphResponse<SharePointSiteSummary>> {
  const raw = sitePath.trim();
  if (!raw) {
    return graphError('sitePath is required');
  }
  // Graph expects `hostname:/server-relative-path` with `:` and `/` preserved — not a single encoded segment.
  const pathSegment = encodeURI(raw);
  return callGraphAt<SharePointSiteSummary>(apiBase, token, `/sites/${pathSegment}`);
}

export async function getSiteDefaultDriveId(
  token: string,
  siteId: string,
  apiBase: string = getGraphBaseUrl()
): Promise<GraphResponse<{ id: string }>> {
  return callGraphAt<{ id: string }>(apiBase, token, `/sites/${encodeURIComponent(siteId)}/drive`);
}

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

export async function getLists(
  token: string,
  siteId: string,
  apiBase: string = getGraphBaseUrl()
): Promise<GraphResponse<SharePointList[]>> {
  // Page through all lists (SharePoint paginates at ~200) — matches the paginating siblings.
  return fetchAllPages<SharePointList>(
    token,
    `/sites/${encodeURIComponent(siteId)}/lists`,
    'Failed to get lists',
    apiBase
  );
}

/** One page of list items (optional `@odata.nextLink` for more pages). */
export interface ListItemsPageResponse {
  value?: SharePointListItem[];
  '@odata.nextLink'?: string;
}

export async function getSiteById(
  token: string,
  siteId: string,
  apiBase: string = getGraphBaseUrl()
): Promise<GraphResponse<SharePointSiteSummary>> {
  return callGraphAt<SharePointSiteSummary>(apiBase, token, `/sites/${encodeURIComponent(siteId)}`);
}

export interface SiteDriveSummary {
  id: string;
  name?: string;
  driveType?: string;
  webUrl?: string;
  description?: string;
}

export async function getSiteDrives(
  token: string,
  siteId: string,
  apiBase: string = getGraphBaseUrl()
): Promise<GraphResponse<SiteDriveSummary[]>> {
  return fetchAllPages<SiteDriveSummary>(
    token,
    `/sites/${encodeURIComponent(siteId)}/drives`,
    'Failed to list site drives',
    apiBase
  );
}

/** Column definition as returned by `GET …/lists/{id}/columns` (shape varies by column type). */
export type SharePointColumnDefinition = Record<string, unknown>;

export async function getListColumns(
  token: string,
  siteId: string,
  listId: string,
  apiBase: string = getGraphBaseUrl()
): Promise<GraphResponse<SharePointColumnDefinition[]>> {
  return fetchAllPages<SharePointColumnDefinition>(
    token,
    `/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/columns`,
    'Failed to list list columns',
    apiBase
  );
}

export async function getListMetadata(
  token: string,
  siteId: string,
  listId: string,
  apiBase: string = getGraphBaseUrl()
): Promise<GraphResponse<SharePointList>> {
  try {
    return await callGraphAt<SharePointList>(
      apiBase,
      token,
      `/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}`
    );
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to get list');
  }
}

export async function getListItems(
  token: string,
  siteId: string,
  listId: string,
  apiBase: string = getGraphBaseUrl()
): Promise<GraphResponse<SharePointListItem[]>> {
  return fetchAllPages<SharePointListItem>(
    token,
    `/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/items?$expand=fields`,
    'Failed to get list items',
    apiBase
  );
}

/**
 * Single GET for list items — use `nextLink` from the response for the next page, or pass `nextLink` to continue.
 * When `nextLink` is set, other options are ignored.
 */
export async function getListItemsPage(
  token: string,
  siteId: string,
  listId: string,
  opts: { nextLink?: string; filter?: string; orderby?: string; top?: number; apiBase?: string }
): Promise<GraphResponse<ListItemsPageResponse>> {
  const apiBase = opts.apiBase ?? getGraphBaseUrl();
  try {
    if (opts.nextLink?.trim()) {
      return await callGraphAbsolute<ListItemsPageResponse>(token, opts.nextLink.trim());
    }
    const params = new URLSearchParams();
    params.set('$expand', 'fields');
    if (opts.filter?.trim()) params.set('$filter', opts.filter.trim());
    if (opts.orderby?.trim()) params.set('$orderby', opts.orderby.trim());
    if (opts.top !== undefined && Number.isFinite(opts.top) && opts.top > 0) {
      params.set('$top', String(Math.min(Math.floor(opts.top), 999)));
    }
    const path = `/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/items?${params.toString()}`;
    return await callGraphAt<ListItemsPageResponse>(apiBase, token, path);
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to get list items page');
  }
}

/** Follow `@odata.nextLink` until exhausted; use after `getListItemsPage` with filter/orderby/top. */
export async function getAllListItemsPages(
  token: string,
  siteId: string,
  listId: string,
  opts: { filter?: string; orderby?: string; top?: number; apiBase?: string }
): Promise<GraphResponse<SharePointListItem[]>> {
  const items: SharePointListItem[] = [];
  let nextLink: string | undefined;
  let first = true;
  while (first || nextLink) {
    const r = await getListItemsPage(token, siteId, listId, {
      nextLink: first ? undefined : nextLink,
      filter: first ? opts.filter : undefined,
      orderby: first ? opts.orderby : undefined,
      top: first ? opts.top : undefined,
      apiBase: opts.apiBase
    });
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to get list items', r.error?.code, r.error?.status);
    }
    items.push(...(r.data.value ?? []));
    nextLink = r.data['@odata.nextLink'];
    first = false;
    if (!nextLink) break;
  }
  return graphResult(items);
}

export async function createListItem(
  token: string,
  siteId: string,
  listId: string,
  fields: Record<string, any>,
  apiBase: string = getGraphBaseUrl()
): Promise<GraphResponse<SharePointListItem>> {
  try {
    return await callGraphAt<SharePointListItem>(
      apiBase,
      token,
      `/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/items`,
      {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ fields })
      }
    );
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
  fields: Record<string, any>,
  apiBase: string = getGraphBaseUrl()
): Promise<GraphResponse<Record<string, any>>> {
  try {
    return await callGraphAt<Record<string, any>>(
      apiBase,
      token,
      `/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/items/${encodeURIComponent(itemId)}/fields`,
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

export async function getListItem(
  token: string,
  siteId: string,
  listId: string,
  itemId: string,
  apiBase: string = getGraphBaseUrl()
): Promise<GraphResponse<SharePointListItem>> {
  try {
    return await callGraphAt<SharePointListItem>(
      apiBase,
      token,
      `/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/items/${encodeURIComponent(itemId)}?$expand=fields`
    );
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to get list item');
  }
}

export async function deleteListItem(
  token: string,
  siteId: string,
  listId: string,
  itemId: string,
  apiBase: string = getGraphBaseUrl()
): Promise<GraphResponse<void>> {
  try {
    return await callGraphAt<void>(
      apiBase,
      token,
      `/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/items/${encodeURIComponent(itemId)}`,
      { method: 'DELETE' },
      false
    );
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to delete list item');
  }
}

export interface ListItemsDeltaPage {
  value?: SharePointListItem[];
  '@odata.nextLink'?: string;
  '@odata.deltaLink'?: string;
}

export async function getListItemsDeltaPage(
  token: string,
  siteId: string,
  listId: string,
  nextOrDeltaLink?: string,
  apiBase: string = getGraphBaseUrl()
): Promise<GraphResponse<ListItemsDeltaPage>> {
  try {
    if (nextOrDeltaLink?.trim()) {
      return await callGraphAbsolute<ListItemsDeltaPage>(token, nextOrDeltaLink.trim());
    }
    const path = `/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/items/delta?$expand=fields`;
    return await callGraphAt<ListItemsDeltaPage>(apiBase, token, path);
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to fetch list items delta');
  }
}

/** Site sharing permissions (owners can list; some updates are beta). */
export type SitePermission = Record<string, unknown>;

export async function getSitePermissions(
  token: string,
  siteId: string,
  apiBase: string = getGraphBaseUrl()
): Promise<GraphResponse<SitePermission[]>> {
  return fetchAllPages<SitePermission>(
    token,
    `/sites/${encodeURIComponent(siteId)}/permissions`,
    'Failed to list site permissions',
    apiBase
  );
}

export async function updateSitePermission(
  token: string,
  siteId: string,
  permissionId: string,
  body: Record<string, unknown>,
  apiBase: string = getGraphBaseUrl()
): Promise<GraphResponse<SitePermission>> {
  try {
    return await callGraphAt<SitePermission>(
      apiBase,
      token,
      `/sites/${encodeURIComponent(siteId)}/permissions/${encodeURIComponent(permissionId)}`,
      {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body)
      }
    );
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to update site permission');
  }
}

/** GET a single site permission by ID. */
export async function getSitePermission(
  token: string,
  siteId: string,
  permissionId: string,
  apiBase: string = getGraphBaseUrl()
): Promise<GraphResponse<SitePermission>> {
  try {
    return await callGraphAt<SitePermission>(
      apiBase,
      token,
      `/sites/${encodeURIComponent(siteId)}/permissions/${encodeURIComponent(permissionId)}`,
      { method: 'GET' }
    );
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to get site permission');
  }
}

/**
 * POST /sites/{id}/permissions — creates a new **application** permission on a site (per Graph
 * docs this cannot be used to grant a new *user* site permission). `body` is a Graph `permission`
 * resource, e.g. `{ roles: ["write"], grantedToIdentities: [{ application: { id, displayName } }] }`.
 */
export async function createSitePermission(
  token: string,
  siteId: string,
  body: Record<string, unknown>,
  apiBase: string = getGraphBaseUrl()
): Promise<GraphResponse<SitePermission>> {
  try {
    return await callGraphAt<SitePermission>(apiBase, token, `/sites/${encodeURIComponent(siteId)}/permissions`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to create site permission');
  }
}

/** DELETE /sites/{id}/permissions/{id} — revoke a site permission. */
export async function deleteSitePermission(
  token: string,
  siteId: string,
  permissionId: string,
  apiBase: string = getGraphBaseUrl()
): Promise<GraphResponse<void>> {
  try {
    return await callGraphAt<void>(
      apiBase,
      token,
      `/sites/${encodeURIComponent(siteId)}/permissions/${encodeURIComponent(permissionId)}`,
      { method: 'DELETE' },
      false
    );
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to delete site permission');
  }
}
