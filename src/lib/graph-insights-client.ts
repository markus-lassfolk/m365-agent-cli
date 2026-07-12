import {
  callGraph,
  callGraphAt,
  type DriveLocation,
  driveItemPath,
  fetchAllPages,
  GraphApiError,
  type GraphResponse,
  getGraphBaseUrl,
  graphError,
  graphErrorFromApiError,
  graphResult
} from './graph-client.js';
import { graphUserPath } from './graph-user-path.js';

/**
 * Office Graph [insights](https://learn.microsoft.com/graph/api/resources/officegraphinsights):
 * `trending` (documents trending around the user), `used` (recently used by the user),
 * `shared` (shared with the user). All three are **delegated** under `/me/insights/...`
 * (or `/users/{id}/insights/...` for another user when the caller has consent).
 */
export type InsightKind = 'trending' | 'used' | 'shared';

/** Minimal shape covering the common surface across the three insight types. */
export interface InsightItem {
  id?: string;
  weight?: number;
  /** trending */
  resourceVisualization?: {
    title?: string;
    type?: string;
    mediaType?: string;
    containerWebUrl?: string;
    containerDisplayName?: string;
    previewImageUrl?: string;
  };
  resourceReference?: {
    webUrl?: string;
    id?: string;
    type?: string;
  };
  /** used */
  lastUsed?: { lastAccessedDateTime?: string; lastModifiedDateTime?: string };
  /** shared */
  lastShared?: {
    sharedDateTime?: string;
    sharingReference?: { webUrl?: string };
    sharingSubject?: string;
    sharingType?: string;
    sharedBy?: { displayName?: string; address?: string };
  };
}

export interface InsightListResponse {
  value?: InsightItem[];
  '@odata.nextLink'?: string;
}

export async function listInsights(
  token: string,
  kind: InsightKind,
  options: { user?: string; top?: number } = {}
): Promise<GraphResponse<InsightListResponse>> {
  const topN = options.top && options.top > 0 ? Math.min(Math.max(1, options.top), 200) : undefined;
  const basePath = graphUserPath(options.user, `insights/${kind}`);
  try {
    if (topN !== undefined) {
      return await callGraph<InsightListResponse>(token, `${basePath}?$top=${topN}`);
    }
    const all = await fetchAllPages<InsightItem>(token, basePath, `Failed to list insights/${kind}`);
    if (!all.ok || !all.data) return all as unknown as GraphResponse<InsightListResponse>;
    return graphResult({ value: all.data });
  } catch (err) {
    if (err instanceof GraphApiError) return graphErrorFromApiError(err);
    return graphError(err instanceof Error ? err.message : `Failed to list insights/${kind}`);
  }
}

export interface ItemActivity {
  id?: string;
  action?: Record<string, unknown>;
  actor?: { user?: { displayName?: string; email?: string }; application?: { displayName?: string } };
  times?: { recordedDateTime?: string; observedDateTime?: string };
}

export interface ItemActivitiesResponse {
  value?: ItemActivity[];
  '@odata.nextLink'?: string;
}

/** `GET /drives/{driveId}/items/{itemId}/activities` — per-item activity feed. */
export async function listDriveItemActivities(
  token: string,
  loc: DriveLocation,
  itemId: string,
  options: { top?: number; graphBaseUrl?: string } = {}
): Promise<GraphResponse<ItemActivitiesResponse>> {
  const base = options.graphBaseUrl ?? getGraphBaseUrl();
  const topN = options.top && options.top > 0 ? Math.min(Math.max(1, options.top), 200) : undefined;
  const basePath = `${driveItemPath(loc, itemId)}/activities`;
  try {
    if (topN !== undefined) {
      return await callGraphAt<ItemActivitiesResponse>(base, token, `${basePath}?$top=${topN}`);
    }
    const all = await fetchAllPages<ItemActivity>(token, basePath, 'Failed to list drive item activities', base);
    if (!all.ok || !all.data) return all as unknown as GraphResponse<ItemActivitiesResponse>;
    return graphResult({ value: all.data });
  } catch (err) {
    if (err instanceof GraphApiError) return graphErrorFromApiError(err);
    return graphError(err instanceof Error ? err.message : 'Failed to list drive item activities');
  }
}

export interface DriveItemPreview {
  getUrl?: string;
  postUrl?: string;
  postParameters?: string;
}

/**
 * `POST /drives/{driveId}/items/{itemId}/preview` — preview session URL for any drive item.
 * Body fields are optional ([driveItem-preview](https://learn.microsoft.com/graph/api/driveitem-preview)).
 */
export async function createDriveItemPreview(
  token: string,
  loc: DriveLocation,
  itemId: string,
  body: { page?: number | string; zoom?: number; allowEdit?: boolean; chromeless?: boolean } = {},
  graphBaseUrl: string = getGraphBaseUrl()
): Promise<GraphResponse<DriveItemPreview>> {
  const path = `${driveItemPath(loc, itemId)}/preview`;
  const cleaned: Record<string, unknown> = {};
  if (body.page !== undefined) cleaned.page = body.page;
  if (body.zoom !== undefined) cleaned.zoom = body.zoom;
  if (body.allowEdit !== undefined) cleaned.allowEdit = body.allowEdit;
  if (body.chromeless !== undefined) cleaned.chromeless = body.chromeless;
  try {
    const r = await callGraphAt<DriveItemPreview>(graphBaseUrl, token, path, {
      method: 'POST',
      body: JSON.stringify(cleaned)
    });
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to create preview', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphErrorFromApiError(err);
    return graphError(err instanceof Error ? err.message : 'Failed to create preview');
  }
}

export interface FollowedSite {
  id?: string;
  webUrl?: string;
  displayName?: string;
  name?: string;
  description?: string;
}

export interface FollowedSitesResponse {
  value?: FollowedSite[];
  '@odata.nextLink'?: string;
}

/** `POST /me/followedSites/remove` may return **207** with per-site errors ([unfollow site](https://learn.microsoft.com/graph/api/site-unfollow)). */
export interface FollowedSitesRemoveMultiStatusBody {
  value?: Array<{ id?: string; error?: { code?: string; message?: string; '@odata.type'?: string } }>;
}

/** `GET /me/followedSites` — sites the user follows. */
export async function listFollowedSites(
  token: string,
  graphBaseUrl: string = getGraphBaseUrl()
): Promise<GraphResponse<FollowedSitesResponse>> {
  try {
    const all = await fetchAllPages<FollowedSite>(
      token,
      '/me/followedSites',
      'Failed to list /me/followedSites',
      graphBaseUrl
    );
    if (!all.ok || !all.data) return all as unknown as GraphResponse<FollowedSitesResponse>;
    return graphResult({ value: all.data });
  } catch (err) {
    if (err instanceof GraphApiError) return graphErrorFromApiError(err);
    return graphError(err instanceof Error ? err.message : 'Failed to list /me/followedSites');
  }
}

/** `POST /me/followedSites/add` — follow one or more sites. */
export async function followSites(
  token: string,
  siteIds: string[],
  graphBaseUrl: string = getGraphBaseUrl()
): Promise<GraphResponse<FollowedSitesResponse>> {
  if (siteIds.length === 0) {
    return graphError('Provide at least one site id');
  }
  const body = { value: siteIds.map((id) => ({ id })) };
  try {
    return await callGraphAt<FollowedSitesResponse>(graphBaseUrl, token, '/me/followedSites/add', {
      method: 'POST',
      body: JSON.stringify(body)
    });
  } catch (err) {
    if (err instanceof GraphApiError) return graphErrorFromApiError(err);
    return graphError(err instanceof Error ? err.message : 'Failed to follow site');
  }
}

/** `POST /me/followedSites/remove` — unfollow one or more sites. */
export async function unfollowSites(
  token: string,
  siteIds: string[],
  graphBaseUrl: string = getGraphBaseUrl()
): Promise<GraphResponse<void>> {
  if (siteIds.length === 0) {
    return graphError('Provide at least one site id');
  }
  const body = { value: siteIds.map((id) => ({ id })) };
  try {
    const r = await callGraphAt<FollowedSitesRemoveMultiStatusBody | undefined>(
      graphBaseUrl,
      token,
      '/me/followedSites/remove',
      { method: 'POST', body: JSON.stringify(body) },
      true
    );
    if (!r.ok) {
      return graphError(r.error?.message || 'Failed to unfollow site', r.error?.code, r.error?.status);
    }
    const entries = r.data?.value;
    if (Array.isArray(entries) && entries.length > 0) {
      const failed = entries.filter((item) => item?.error);
      if (failed.length > 0) {
        const msg = failed
          .map((item) => `${item.id ?? '(unknown)'}: ${item.error?.message ?? item.error?.code ?? 'error'}`)
          .join('; ');
        return graphError(`Unfollow failed for ${failed.length} site(s): ${msg}`, undefined, 207);
      }
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphErrorFromApiError(err);
    return graphError(err instanceof Error ? err.message : 'Failed to unfollow site');
  }
}
