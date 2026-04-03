import { readFile } from 'node:fs/promises';
import { basename, resolve } from 'node:path';
import {
  callGraph,
  callGraphAbsolute,
  fetchAllPages,
  fetchGraphRaw,
  GraphApiError,
  type GraphResponse,
  graphError,
  graphResult
} from './graph-client.js';
import { GRAPH_BASE_URL } from './graph-constants.js';
import { graphUserPath } from './graph-user-path.js';
import { lookupMimeType } from './mime-type.js';

/**
 * Optional group or site root for OneNote (`/groups/{id}/onenote`, `/sites/{id}/onenote`).
 * If `siteId` is set it wins over `groupId`. Omit both for `/me/onenote` or `/users/{id}/onenote`.
 */
export type OneNoteGraphScope = { groupId?: string; siteId?: string };

function oneNoteRoot(user: string | undefined, scope?: OneNoteGraphScope): string {
  const site = scope?.siteId?.trim();
  const group = scope?.groupId?.trim();
  if (site) return `/sites/${encodeURIComponent(site)}/onenote`;
  if (group) return `/groups/${encodeURIComponent(group)}/onenote`;
  return graphUserPath(user, 'onenote');
}

function notebooksPath(user?: string, scope?: OneNoteGraphScope): string {
  return `${oneNoteRoot(user, scope)}/notebooks`;
}

function sectionGroupsPath(user?: string, scope?: OneNoteGraphScope): string {
  return `${oneNoteRoot(user, scope)}/sectionGroups`;
}

function sectionsPath(user?: string, scope?: OneNoteGraphScope): string {
  return `${oneNoteRoot(user, scope)}/sections`;
}

function pagesPath(user?: string, scope?: OneNoteGraphScope): string {
  return `${oneNoteRoot(user, scope)}/pages`;
}

function resourcesPath(user?: string, scope?: OneNoteGraphScope): string {
  return `${oneNoteRoot(user, scope)}/resources`;
}

/** Graph [notebook](https://learn.microsoft.com/en-us/graph/api/resources/notebook) (subset). */
export interface OneNoteNotebook {
  id: string;
  displayName?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  links?: { oneNoteWebUrl?: { href?: string } };
  sectionGroupsUrl?: string;
  sectionsUrl?: string;
}

/** Graph [sectionGroup](https://learn.microsoft.com/en-us/graph/api/resources/sectiongroup) (subset). */
export interface OneNoteSectionGroup {
  id: string;
  displayName?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  sectionsUrl?: string;
  sectionGroupsUrl?: string;
}

/** Graph [onenoteSection](https://learn.microsoft.com/en-us/graph/api/resources/onenotesection) (subset). */
export interface OneNoteSection {
  id: string;
  displayName?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  pagesUrl?: string;
  parentNotebook?: { id?: string; displayName?: string };
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

/** [GET …/pages/{id}/preview](https://learn.microsoft.com/en-us/graph/api/page-preview) (subset). */
export interface OneNotePagePreview {
  previewText?: string;
}

/** Async copy / long-running operation ([onenoteOperation](https://learn.microsoft.com/en-us/graph/api/resources/onenoteoperation) subset). */
export interface OneNoteOperation {
  id?: string;
  status?: 'notStarted' | 'running' | 'completed' | 'failed';
  percentComplete?: string;
  resourceLocation?: string;
  resourceId?: string;
  error?: { message?: string; code?: string };
}

/** List notebooks for the signed-in user, another user, a group, or a site. */
export async function listOneNoteNotebooks(
  token: string,
  user?: string,
  scope?: OneNoteGraphScope
): Promise<GraphResponse<OneNoteNotebook[]>> {
  return fetchAllPages<OneNoteNotebook>(token, notebooksPath(user, scope), 'Failed to list OneNote notebooks');
}

/**
 * List pages across all notebooks ([GET …/onenote/pages](https://learn.microsoft.com/en-us/graph/api/onenote-list-pages)).
 * @param odataQuery - Optional query without leading `?` (e.g. `$top=10&$orderby=lastModifiedTime desc`).
 */
export async function listAllOneNotePages(
  token: string,
  user?: string,
  odataQuery?: string,
  scope?: OneNoteGraphScope
): Promise<GraphResponse<OneNotePage[]>> {
  const base = pagesPath(user, scope);
  const path = odataQuery?.trim() ? `${base}?${odataQuery.trim().replace(/^\?/, '')}` : base;
  return fetchAllPages<OneNotePage>(token, path, 'Failed to list OneNote pages');
}

export async function getOneNoteNotebook(
  token: string,
  notebookId: string,
  user?: string,
  scope?: OneNoteGraphScope
): Promise<GraphResponse<OneNoteNotebook>> {
  try {
    const result = await callGraph<OneNoteNotebook>(
      token,
      `${notebooksPath(user, scope)}/${encodeURIComponent(notebookId)}`
    );
    if (!result.ok || !result.data) {
      return graphError(result.error?.message || 'Failed to get notebook', result.error?.code, result.error?.status);
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get notebook');
  }
}

export async function createOneNoteNotebook(
  token: string,
  displayName: string,
  user?: string,
  scope?: OneNoteGraphScope
): Promise<GraphResponse<OneNoteNotebook>> {
  try {
    const result = await callGraph<OneNoteNotebook>(token, notebooksPath(user, scope), {
      method: 'POST',
      body: JSON.stringify({ displayName })
    });
    if (!result.ok || !result.data) {
      return graphError(result.error?.message || 'Failed to create notebook', result.error?.code, result.error?.status);
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to create notebook');
  }
}

/**
 * [getNotebookFromWebUrl](https://learn.microsoft.com/en-us/graph/api/notebook-getnotebookfromweburl) —
 * resolve a notebook by its OneNote web URL.
 */
export async function getOneNoteNotebookFromWebUrl(
  token: string,
  webUrl: string,
  user?: string,
  scope?: OneNoteGraphScope
): Promise<GraphResponse<OneNoteNotebook>> {
  const path = `${notebooksPath(user, scope)}/GetNotebookFromWebUrl`;
  try {
    const result = await callGraph<OneNoteNotebook>(token, path, {
      method: 'POST',
      body: JSON.stringify({ webUrl: webUrl.trim() })
    });
    if (!result.ok || !result.data) {
      return graphError(
        result.error?.message || 'Failed to resolve notebook from URL',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to resolve notebook from URL');
  }
}

export async function updateOneNoteNotebook(
  token: string,
  notebookId: string,
  patch: Record<string, unknown>,
  user?: string,
  scope?: OneNoteGraphScope
): Promise<GraphResponse<OneNoteNotebook | undefined>> {
  try {
    const result = await callGraph<OneNoteNotebook>(
      token,
      `${notebooksPath(user, scope)}/${encodeURIComponent(notebookId)}`,
      { method: 'PATCH', body: JSON.stringify(patch) }
    );
    if (!result.ok) {
      return graphError(result.error?.message || 'Failed to update notebook', result.error?.code, result.error?.status);
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to update notebook');
  }
}

export async function deleteOneNoteNotebook(
  token: string,
  notebookId: string,
  user?: string,
  scope?: OneNoteGraphScope
): Promise<GraphResponse<void>> {
  try {
    return await callGraph<void>(
      token,
      `${notebooksPath(user, scope)}/${encodeURIComponent(notebookId)}`,
      { method: 'DELETE' },
      false
    );
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to delete notebook');
  }
}

export async function listNotebookSectionGroups(
  token: string,
  notebookId: string,
  user?: string,
  scope?: OneNoteGraphScope
): Promise<GraphResponse<OneNoteSectionGroup[]>> {
  return fetchAllPages<OneNoteSectionGroup>(
    token,
    `${notebooksPath(user, scope)}/${encodeURIComponent(notebookId)}/sectionGroups`,
    'Failed to list section groups'
  );
}

export async function getOneNoteSectionGroup(
  token: string,
  sectionGroupId: string,
  user?: string,
  scope?: OneNoteGraphScope
): Promise<GraphResponse<OneNoteSectionGroup>> {
  try {
    const result = await callGraph<OneNoteSectionGroup>(
      token,
      `${sectionGroupsPath(user, scope)}/${encodeURIComponent(sectionGroupId)}`
    );
    if (!result.ok || !result.data) {
      return graphError(
        result.error?.message || 'Failed to get section group',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get section group');
  }
}

export async function createSectionGroupInNotebook(
  token: string,
  notebookId: string,
  displayName: string,
  user?: string,
  scope?: OneNoteGraphScope
): Promise<GraphResponse<OneNoteSectionGroup>> {
  try {
    const result = await callGraph<OneNoteSectionGroup>(
      token,
      `${notebooksPath(user, scope)}/${encodeURIComponent(notebookId)}/sectionGroups`,
      { method: 'POST', body: JSON.stringify({ displayName }) }
    );
    if (!result.ok || !result.data) {
      return graphError(
        result.error?.message || 'Failed to create section group',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to create section group');
  }
}

export async function updateOneNoteSectionGroup(
  token: string,
  sectionGroupId: string,
  patch: Record<string, unknown>,
  user?: string,
  scope?: OneNoteGraphScope
): Promise<GraphResponse<OneNoteSectionGroup | undefined>> {
  try {
    const result = await callGraph<OneNoteSectionGroup>(
      token,
      `${sectionGroupsPath(user, scope)}/${encodeURIComponent(sectionGroupId)}`,
      { method: 'PATCH', body: JSON.stringify(patch) }
    );
    if (!result.ok) {
      return graphError(
        result.error?.message || 'Failed to update section group',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to update section group');
  }
}

export async function deleteOneNoteSectionGroup(
  token: string,
  sectionGroupId: string,
  user?: string,
  scope?: OneNoteGraphScope
): Promise<GraphResponse<void>> {
  try {
    return await callGraph<void>(
      token,
      `${sectionGroupsPath(user, scope)}/${encodeURIComponent(sectionGroupId)}`,
      { method: 'DELETE' },
      false
    );
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to delete section group');
  }
}

export async function listNotebookSections(
  token: string,
  notebookId: string,
  user?: string,
  scope?: OneNoteGraphScope
): Promise<GraphResponse<OneNoteSection[]>> {
  return fetchAllPages<OneNoteSection>(
    token,
    `${notebooksPath(user, scope)}/${encodeURIComponent(notebookId)}/sections`,
    'Failed to list notebook sections'
  );
}

export async function listSectionsInSectionGroup(
  token: string,
  sectionGroupId: string,
  user?: string,
  scope?: OneNoteGraphScope
): Promise<GraphResponse<OneNoteSection[]>> {
  return fetchAllPages<OneNoteSection>(
    token,
    `${sectionGroupsPath(user, scope)}/${encodeURIComponent(sectionGroupId)}/sections`,
    'Failed to list sections in section group'
  );
}

export async function getOneNoteSection(
  token: string,
  sectionId: string,
  user?: string,
  scope?: OneNoteGraphScope
): Promise<GraphResponse<OneNoteSection>> {
  try {
    const result = await callGraph<OneNoteSection>(
      token,
      `${sectionsPath(user, scope)}/${encodeURIComponent(sectionId)}`
    );
    if (!result.ok || !result.data) {
      return graphError(result.error?.message || 'Failed to get section', result.error?.code, result.error?.status);
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get section');
  }
}

export async function createSectionInNotebook(
  token: string,
  notebookId: string,
  displayName: string,
  user?: string,
  scope?: OneNoteGraphScope
): Promise<GraphResponse<OneNoteSection>> {
  try {
    const result = await callGraph<OneNoteSection>(
      token,
      `${notebooksPath(user, scope)}/${encodeURIComponent(notebookId)}/sections`,
      { method: 'POST', body: JSON.stringify({ displayName }) }
    );
    if (!result.ok || !result.data) {
      return graphError(result.error?.message || 'Failed to create section', result.error?.code, result.error?.status);
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to create section');
  }
}

export async function createSectionInSectionGroup(
  token: string,
  sectionGroupId: string,
  displayName: string,
  user?: string,
  scope?: OneNoteGraphScope
): Promise<GraphResponse<OneNoteSection>> {
  try {
    const result = await callGraph<OneNoteSection>(
      token,
      `${sectionGroupsPath(user, scope)}/${encodeURIComponent(sectionGroupId)}/sections`,
      { method: 'POST', body: JSON.stringify({ displayName }) }
    );
    if (!result.ok || !result.data) {
      return graphError(result.error?.message || 'Failed to create section', result.error?.code, result.error?.status);
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to create section');
  }
}

export async function updateOneNoteSection(
  token: string,
  sectionId: string,
  patch: Record<string, unknown>,
  user?: string,
  scope?: OneNoteGraphScope
): Promise<GraphResponse<OneNoteSection | undefined>> {
  try {
    const result = await callGraph<OneNoteSection>(
      token,
      `${sectionsPath(user, scope)}/${encodeURIComponent(sectionId)}`,
      { method: 'PATCH', body: JSON.stringify(patch) }
    );
    if (!result.ok) {
      return graphError(result.error?.message || 'Failed to update section', result.error?.code, result.error?.status);
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to update section');
  }
}

export async function deleteOneNoteSection(
  token: string,
  sectionId: string,
  user?: string,
  scope?: OneNoteGraphScope
): Promise<GraphResponse<void>> {
  try {
    return await callGraph<void>(
      token,
      `${sectionsPath(user, scope)}/${encodeURIComponent(sectionId)}`,
      { method: 'DELETE' },
      false
    );
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to delete section');
  }
}

export async function listSectionPages(
  token: string,
  sectionId: string,
  user?: string,
  scope?: OneNoteGraphScope
): Promise<GraphResponse<OneNotePage[]>> {
  return fetchAllPages<OneNotePage>(
    token,
    `${sectionsPath(user, scope)}/${encodeURIComponent(sectionId)}/pages`,
    'Failed to list section pages'
  );
}

export async function getOneNotePage(
  token: string,
  pageId: string,
  user?: string,
  scope?: OneNoteGraphScope
): Promise<GraphResponse<OneNotePage>> {
  try {
    const result = await callGraph<OneNotePage>(token, `${pagesPath(user, scope)}/${encodeURIComponent(pageId)}`);
    if (!result.ok || !result.data) {
      return graphError(result.error?.message || 'Failed to get page', result.error?.code, result.error?.status);
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get page');
  }
}

export async function getOneNotePagePreview(
  token: string,
  pageId: string,
  user?: string,
  scope?: OneNoteGraphScope
): Promise<GraphResponse<OneNotePagePreview>> {
  try {
    const result = await callGraph<OneNotePagePreview>(
      token,
      `${pagesPath(user, scope)}/${encodeURIComponent(pageId)}/preview`
    );
    if (!result.ok || !result.data) {
      return graphError(
        result.error?.message || 'Failed to get page preview',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get page preview');
  }
}

export async function getOneNotePageContentHtml(
  token: string,
  pageId: string,
  user?: string,
  scope?: OneNoteGraphScope,
  options?: { includeIds?: boolean }
): Promise<GraphResponse<string>> {
  let path = `${pagesPath(user, scope)}/${encodeURIComponent(pageId)}/content`;
  if (options?.includeIds) {
    path += '?includeIDs=true';
  }
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

/**
 * Download binary content for an embedded OneNote resource (image, file)
 * ([GET …/resources/{id}/content](https://learn.microsoft.com/en-us/graph/api/resource-get)).
 * Resource ids appear in page HTML (`data-fullres-src`, `object data`, etc.).
 */
export async function getOneNoteResourceContent(
  token: string,
  resourceId: string,
  user?: string,
  scope?: OneNoteGraphScope
): Promise<GraphResponse<Uint8Array>> {
  const path = `${resourcesPath(user, scope)}/${encodeURIComponent(resourceId)}/content`;
  try {
    const res = await fetchGraphRaw(token, path);
    if (!res.ok) {
      let message = `Failed to download resource: HTTP ${res.status}`;
      try {
        const j = (await res.json()) as { error?: { message?: string } };
        if (j.error?.message) message = j.error.message;
      } catch {
        /* ignore */
      }
      return graphError(message, undefined, res.status);
    }
    const buf = new Uint8Array(await res.arrayBuffer());
    return graphResult(buf);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to download OneNote resource');
  }
}

/** [oneNoteResource](https://learn.microsoft.com/en-us/graph/api/resources/onenoteresource) metadata (GET without `/content`). */
export interface OneNoteResourceInfo {
  id?: string;
  contentUrl?: string;
  self?: string;
  [key: string]: unknown;
}

/**
 * [Get resource](https://learn.microsoft.com/en-us/graph/api/resource-get) — metadata (not binary bytes).
 */
export async function getOneNoteResource(
  token: string,
  resourceId: string,
  user?: string,
  scope?: OneNoteGraphScope
): Promise<GraphResponse<OneNoteResourceInfo>> {
  try {
    const result = await callGraph<OneNoteResourceInfo>(
      token,
      `${resourcesPath(user, scope)}/${encodeURIComponent(resourceId)}`
    );
    if (!result.ok || !result.data) {
      return graphError(result.error?.message || 'Failed to get resource', result.error?.code, result.error?.status);
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get OneNote resource');
  }
}

/** Binary part for multipart create/PATCH (`name:partName` in HTML / `src="name:partName"`). */
export interface OneNoteMultipartBinaryPart {
  partName: string;
  absolutePath: string;
}

/**
 * [Create page with images/files](https://learn.microsoft.com/en-us/graph/api/section-post-pages) —
 * multipart/form-data with a **Presentation** (HTML) part and named binary parts referenced in HTML as `src="name:partName"` or `data="name:partName"`.
 */
export async function createOneNotePageMultipart(
  token: string,
  sectionId: string,
  presentationHtml: string,
  binaryParts: OneNoteMultipartBinaryPart[],
  user?: string,
  scope?: OneNoteGraphScope
): Promise<GraphResponse<OneNotePage>> {
  const path = `${sectionsPath(user, scope)}/${encodeURIComponent(sectionId)}/pages`;
  try {
    const form = new FormData();
    form.append(
      'Presentation',
      new Blob([presentationHtml], { type: 'text/html' }),
      'presentation.html'
    );
    const wd = process.cwd();
    for (const p of binaryParts) {
      const abs = resolve(wd, p.absolutePath);
      const buf = await readFile(abs);
      const fileName = basename(abs);
      const mime = lookupMimeType(fileName) || 'application/octet-stream';
      form.append(p.partName, new Blob([buf], { type: mime }), fileName);
    }

    const res = await fetch(`${GRAPH_BASE_URL}${path}`, {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}` },
      body: form
    });

    const text = await res.text();
    if (!res.ok) {
      let message = text || `HTTP ${res.status}`;
      try {
        const j = JSON.parse(text) as { error?: { message?: string } };
        if (j.error?.message) message = j.error.message;
      } catch {
        /* ignore */
      }
      return graphError(message, undefined, res.status);
    }
    if (!text.trim()) {
      return graphError('Empty response from create page');
    }
    const data = JSON.parse(text) as OneNotePage;
    return graphResult(data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to create page (multipart)');
  }
}

/**
 * [PATCH page content with binary parts](https://learn.microsoft.com/en-us/graph/onenote-update-page) —
 * multipart with **Commands** (JSON array) and image/file parts referenced as `src="name:partName"` in command content.
 */
export async function patchOneNotePageContentMultipart(
  token: string,
  pageId: string,
  commands: unknown[],
  binaryParts: OneNoteMultipartBinaryPart[],
  user?: string,
  scope?: OneNoteGraphScope
): Promise<GraphResponse<void>> {
  const path = `${pagesPath(user, scope)}/${encodeURIComponent(pageId)}/content`;
  try {
    const form = new FormData();
    form.append(
      'Commands',
      new Blob([JSON.stringify(commands)], { type: 'application/json' }),
      'commands.json'
    );
    const wd = process.cwd();
    for (const p of binaryParts) {
      const abs = resolve(wd, p.absolutePath);
      const buf = await readFile(abs);
      const fileName = basename(abs);
      const mime = lookupMimeType(fileName) || 'application/octet-stream';
      form.append(p.partName, new Blob([buf], { type: mime }), fileName);
    }

    const res = await fetch(`${GRAPH_BASE_URL}${path}`, {
      method: 'PATCH',
      headers: { Authorization: `Bearer ${token}` },
      body: form
    });

    if (res.status === 204 || res.ok) {
      return graphResult(undefined as void);
    }
    const text = await res.text();
    let message = text || `HTTP ${res.status}`;
    try {
      const j = JSON.parse(text) as { error?: { message?: string } };
      if (j.error?.message) message = j.error.message;
    } catch {
      /* ignore */
    }
    return graphError(message, undefined, res.status);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to patch page (multipart)');
  }
}

export async function createOneNotePageFromHtml(
  token: string,
  sectionId: string,
  html: string,
  user?: string,
  scope?: OneNoteGraphScope
): Promise<GraphResponse<OneNotePage>> {
  const path = `${sectionsPath(user, scope)}/${encodeURIComponent(sectionId)}/pages`;
  try {
    const result = await callGraph<OneNotePage>(token, path, {
      method: 'POST',
      body: html,
      headers: { 'Content-Type': 'text/html' }
    });
    if (!result.ok || !result.data) {
      return graphError(result.error?.message || 'Failed to create page', result.error?.code, result.error?.status);
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to create page');
  }
}

/** [PATCH page content](https://learn.microsoft.com/en-us/graph/api/page-update) — body is Graph patch command set (JSON). */
export async function updateOneNotePageContent(
  token: string,
  pageId: string,
  patchCommands: unknown,
  user?: string,
  scope?: OneNoteGraphScope
): Promise<GraphResponse<void>> {
  try {
    return await callGraph<void>(
      token,
      `${pagesPath(user, scope)}/${encodeURIComponent(pageId)}/content`,
      {
        method: 'PATCH',
        body: JSON.stringify(patchCommands),
        headers: { 'Content-Type': 'application/json' }
      },
      false
    );
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to update page content');
  }
}

export async function deleteOneNotePage(
  token: string,
  pageId: string,
  user?: string,
  scope?: OneNoteGraphScope
): Promise<GraphResponse<void>> {
  try {
    return await callGraph<void>(
      token,
      `${pagesPath(user, scope)}/${encodeURIComponent(pageId)}`,
      { method: 'DELETE' },
      false
    );
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to delete page');
  }
}

/** Result of `copyToSection` (202 Accepted — poll `operationLocation`). */
export interface OneNoteCopyToSectionResult {
  status: number;
  operationLocation?: string;
}

/**
 * [copyToSection](https://learn.microsoft.com/en-us/graph/api/page-copytosection) — async; poll `operationLocation`.
 * @param copyToSectionGroupId - Optional group id for the **request body** when the destination section is in a group notebook (not the Graph path scope).
 */
export async function copyOneNotePageToSection(
  token: string,
  pageId: string,
  targetSectionId: string,
  user?: string,
  scope?: OneNoteGraphScope,
  copyToSectionGroupId?: string
): Promise<GraphResponse<OneNoteCopyToSectionResult>> {
  const path = `${pagesPath(user, scope)}/${encodeURIComponent(pageId)}/copyToSection`;
  const body: Record<string, string> = { id: targetSectionId };
  if (copyToSectionGroupId?.trim()) body.groupId = copyToSectionGroupId.trim();
  try {
    const res = await fetchGraphRaw(token, path, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });
    const op = res.headers.get('Operation-Location') ?? undefined;
    if (res.status === 202) {
      return graphResult({ status: 202, operationLocation: op });
    }
    if (!res.ok) {
      let msg = `HTTP ${res.status}`;
      try {
        const j = (await res.json()) as { error?: { message?: string } };
        msg = j.error?.message || msg;
      } catch {
        /* ignore */
      }
      return graphError(msg, undefined, res.status);
    }
    return graphResult({ status: res.status, operationLocation: op });
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to copy page');
  }
}

/**
 * [section: copyToNotebook](https://learn.microsoft.com/en-us/graph/api/section-copytonotebook) — async; poll `Operation-Location`.
 * @param opts.copyToNotebookGroupId - Request body `groupId` when the destination notebook is in a Microsoft 365 group.
 */
export async function copyOneNoteSectionToNotebook(
  token: string,
  sectionId: string,
  targetNotebookId: string,
  user?: string,
  scope?: OneNoteGraphScope,
  opts?: { copyToNotebookGroupId?: string; renameAs?: string }
): Promise<GraphResponse<OneNoteCopyToSectionResult>> {
  const path = `${sectionsPath(user, scope)}/${encodeURIComponent(sectionId)}/copyToNotebook`;
  const body: Record<string, string> = { id: targetNotebookId };
  if (opts?.copyToNotebookGroupId?.trim()) body.groupId = opts.copyToNotebookGroupId.trim();
  if (opts?.renameAs?.trim()) body.renameAs = opts.renameAs.trim();
  try {
    const res = await fetchGraphRaw(token, path, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });
    const op = res.headers.get('Operation-Location') ?? undefined;
    if (res.status === 202) {
      return graphResult({ status: 202, operationLocation: op });
    }
    if (!res.ok) {
      let msg = `HTTP ${res.status}`;
      try {
        const j = (await res.json()) as { error?: { message?: string } };
        msg = j.error?.message || msg;
      } catch {
        /* ignore */
      }
      return graphError(msg, undefined, res.status);
    }
    return graphResult({ status: res.status, operationLocation: op });
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to copy section to notebook');
  }
}

/**
 * [copyToSectionGroup](https://learn.microsoft.com/en-us/graph/api/section-copytosectiongroup) —
 * copy a section into a section group (async; poll `Operation-Location`).
 */
export async function copyOneNoteSectionToSectionGroup(
  token: string,
  sectionId: string,
  targetSectionGroupId: string,
  user?: string,
  scope?: OneNoteGraphScope,
  opts?: { copyToGroupId?: string; renameAs?: string }
): Promise<GraphResponse<OneNoteCopyToSectionResult>> {
  const path = `${sectionsPath(user, scope)}/${encodeURIComponent(sectionId)}/copyToSectionGroup`;
  const body: Record<string, string> = { id: targetSectionGroupId };
  if (opts?.copyToGroupId?.trim()) body.groupId = opts.copyToGroupId.trim();
  if (opts?.renameAs?.trim()) body.renameAs = opts.renameAs.trim();
  try {
    const res = await fetchGraphRaw(token, path, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });
    const op = res.headers.get('Operation-Location') ?? undefined;
    if (res.status === 202) {
      return graphResult({ status: 202, operationLocation: op });
    }
    if (!res.ok) {
      let msg = `HTTP ${res.status}`;
      try {
        const j = (await res.json()) as { error?: { message?: string } };
        msg = j.error?.message || msg;
      } catch {
        /* ignore */
      }
      return graphError(msg, undefined, res.status);
    }
    return graphResult({ status: res.status, operationLocation: op });
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to copy section to section group');
  }
}

/** Poll copy/rename status ([get operation](https://learn.microsoft.com/en-us/graph/api/onenoteoperation-get)). */
export async function getOneNoteOperation(
  token: string,
  operationLocationUrl: string
): Promise<GraphResponse<OneNoteOperation>> {
  try {
    const result = await callGraphAbsolute<OneNoteOperation>(token, operationLocationUrl.trim());
    if (!result.ok || !result.data) {
      return graphError(
        result.error?.message || 'Failed to get OneNote operation',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get OneNote operation');
  }
}
