import { callGraph, fetchAllPages, GraphApiError, type GraphResponse, graphError } from './graph-client.js';

export interface SitePage {
  id: string;
  name?: string;
  title?: string;
  pageLayout?: string;
  publishingState?: {
    level: string;
    versionId: string;
  };
  webUrl?: string;
  [key: string]: any;
}

export async function listSitePages(token: string, siteId: string): Promise<GraphResponse<SitePage[]>> {
  return fetchAllPages<SitePage>(token, `/sites/${siteId}/pages/microsoft.graph.sitePage`, 'Failed to list site pages');
}

export async function getSitePage(token: string, siteId: string, pageId: string): Promise<GraphResponse<SitePage>> {
  try {
    return await callGraph<SitePage>(
      token,
      `/sites/${siteId}/pages/${encodeURIComponent(pageId)}/microsoft.graph.sitePage`
    );
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to get site page');
  }
}

export async function updateSitePage(
  token: string,
  siteId: string,
  pageId: string,
  pageData: Partial<SitePage>
): Promise<GraphResponse<SitePage>> {
  try {
    return await callGraph<SitePage>(
      token,
      `/sites/${siteId}/pages/${encodeURIComponent(pageId)}/microsoft.graph.sitePage`,
      {
        method: 'PATCH',
        body: JSON.stringify({ '@odata.type': '#microsoft.graph.sitePage', ...pageData })
      }
    );
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to update site page');
  }
}

export async function publishSitePage(token: string, siteId: string, pageId: string): Promise<GraphResponse<void>> {
  try {
    return await callGraph<void>(
      token,
      `/sites/${siteId}/pages/${encodeURIComponent(pageId)}/microsoft.graph.sitePage/publish`,
      {
        method: 'POST'
      },
      false // might not return JSON, just 204 No Content
    );
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to publish site page');
  }
}
