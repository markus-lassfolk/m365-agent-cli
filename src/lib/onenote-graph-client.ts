import {
  callGraph,
  fetchAllPages,
  fetchGraphRaw,
  GraphApiError,
  type GraphResponse,
  graphError,
  graphResult
} from './graph-client.js';
import { graphUserPath } from './graph-user-path.js';

function notebooksPath(user?: string): string {
  return graphUserPath(user, 'onenote/notebooks');
}

/** Graph [notebook](https://learn.microsoft.com/en-us/graph/api/resources/notebook) (subset). */
export interface OneNoteNotebook {
  id: string;
  displayName?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  links?: { oneNoteWebUrl?: { href?: string } };
}

/** Graph [onenoteSection](https://learn.microsoft.com/en-us/graph/api/resources/onenotesection) (subset). */
export interface OneNoteSection {
  id: string;
  displayName?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  pagesUrl?: string;
}

/** Graph [onenotePage](https://learn.microsoft.com/en-us/graph/api/resources/onenotepage) (subset). */
export interface OneNotePage {
  id: string;
  title?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  links?: { oneNoteWebUrl?: { href?: string } };
  contentUrl?: string;
}

export async function listOneNoteNotebooks(token: string, user?: string): Promise<GraphResponse<OneNoteNotebook[]>> {
  return fetchAllPages<OneNoteNotebook>(token, notebooksPath(user), 'Failed to list OneNote notebooks');
}

export async function listNotebookSections(
  token: string,
  notebookId: string,
  user?: string
): Promise<GraphResponse<OneNoteSection[]>> {
  return fetchAllPages<OneNoteSection>(
    token,
    `${notebooksPath(user)}/${encodeURIComponent(notebookId)}/sections`,
    'Failed to list notebook sections'
  );
}

export async function listSectionPages(
  token: string,
  sectionId: string,
  user?: string
): Promise<GraphResponse<OneNotePage[]>> {
  return fetchAllPages<OneNotePage>(
    token,
    `${graphUserPath(user, 'onenote/sections')}/${encodeURIComponent(sectionId)}/pages`,
    'Failed to list section pages'
  );
}

export async function getOneNotePage(
  token: string,
  pageId: string,
  user?: string
): Promise<GraphResponse<OneNotePage>> {
  try {
    const result = await callGraph<OneNotePage>(
      token,
      `${graphUserPath(user, 'onenote/pages')}/${encodeURIComponent(pageId)}`
    );
    if (!result.ok || !result.data) {
      return graphError(result.error?.message || 'Failed to get page', result.error?.code, result.error?.status);
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get page');
  }
}

export async function getOneNotePageContentHtml(
  token: string,
  pageId: string,
  user?: string
): Promise<GraphResponse<string>> {
  const path = `${graphUserPath(user, 'onenote/pages')}/${encodeURIComponent(pageId)}/content`;
  try {
    const res = await fetchGraphRaw(token, path, { headers: { Accept: 'text/html, application/json' } });
    const text = await res.text();
    if (!res.ok) {
      let msg = text;
      try {
        const j = JSON.parse(text) as { error?: { message?: string } };
        if (j.error?.message) msg = j.error.message;
      } catch {
        /* use raw */
      }
      return graphError(msg || `HTTP ${res.status}`, undefined, res.status);
    }
    return graphResult(text);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get page content');
  }
}

export async function createOneNotePageFromHtml(
  token: string,
  sectionId: string,
  html: string,
  user?: string
): Promise<GraphResponse<OneNotePage>> {
  const path = `${graphUserPath(user, 'onenote/sections')}/${encodeURIComponent(sectionId)}/pages`;
  try {
    const result = await callGraph<OneNotePage>(token, path, {
      method: 'POST',
      body: html,
      headers: { 'Content-Type': 'text/html' }
    });
    if (!result.ok || !result.data) {
      return graphError(
        result.error?.message || 'Failed to create page',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to create page');
  }
}
