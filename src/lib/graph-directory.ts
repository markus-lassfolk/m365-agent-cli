import {
  callGraph,
  fetchAllPages,
  GraphApiError,
  type GraphResponse,
  getGraphBaseUrl,
  graphError,
  graphResult
} from './graph-client.js';
import { graphUserPath } from './graph-user-path.js';

export interface Person {
  id: string;
  displayName: string;
  givenName?: string;
  surname?: string;
  jobTitle?: string;
  scoredEmailAddresses?: { address: string; name?: string }[];
  userPrincipalName?: string;
  department?: string;
}

export interface User {
  id: string;
  displayName: string;
  givenName?: string;
  surname?: string;
  jobTitle?: string;
  mail?: string;
  userPrincipalName?: string;
  department?: string;
}

export interface Group {
  id: string;
  displayName: string;
  description?: string;
  mail?: string;
  groupTypes?: string[];
}

export async function searchPeople(token: string, query: string): Promise<GraphResponse<Person[]>> {
  const escapedQuery = query.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
  const searchParam = encodeURIComponent(`"${escapedQuery}"`);
  let result: GraphResponse<{ value: Person[] }>;
  try {
    result = await callGraph<{ value: Person[] }>(token, `/me/people?$search=${searchParam}`, {
      headers: { ConsistencyLevel: 'eventual' }
    });
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to search people');
  }
  if (!result.ok || !result.data) {
    return { ok: false, error: result.error };
  }
  return { ok: true, data: result.data.value };
}

function peopleCollectionPath(forUser?: string): string {
  return graphUserPath(forUser, 'people');
}

/**
 * List relevant people for /me or another user (GET …/people). Paginated when no $top.
 * `$search` uses **ConsistencyLevel: eventual** per Graph advanced query rules.
 */
export async function listPeople(
  token: string,
  opts?: { user?: string; top?: number; search?: string }
): Promise<GraphResponse<Person[]>> {
  const base = peopleCollectionPath(opts?.user);
  const params = new URLSearchParams();
  if (opts?.top !== undefined) params.set('$top', String(opts.top));
  if (opts?.search?.trim()) {
    const escaped = opts.search.trim().replace(/\\/g, '\\\\').replace(/"/g, '\\"');
    params.set('$search', `"${escaped}"`);
  }
  const qs = params.toString();
  const path = qs ? `${base}?${qs}` : base;
  const headers: Record<string, string> = opts?.search?.trim() ? { ConsistencyLevel: 'eventual' } : {};

  if (opts?.top !== undefined) {
    let result: GraphResponse<{ value: Person[] }>;
    try {
      result = await callGraph<{ value: Person[] }>(token, path, { headers });
    } catch (err) {
      if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
      return graphError(err instanceof Error ? err.message : 'Failed to list people');
    }
    if (!result.ok || !result.data) return { ok: false, error: result.error };
    return graphResult(result.data.value ?? []);
  }

  return fetchAllPages<Person>(token, path, 'Failed to list people', getGraphBaseUrl(), { headers });
}

/** GET /me/people/{id} or /users/{id}/people/{personId}. */
export async function getPerson(token: string, personId: string, forUser?: string): Promise<GraphResponse<Person>> {
  const base = peopleCollectionPath(forUser);
  const path = `${base}/${encodeURIComponent(personId)}`;
  try {
    const result = await callGraph<Person>(token, path);
    if (!result.ok || !result.data) {
      return { ok: false, error: result.error };
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get person');
  }
}

export async function searchUsers(token: string, query: string): Promise<GraphResponse<User[]>> {
  const escapedQuery = query.replace(/'/g, "''");
  const filter = encodeURIComponent(`startswith(displayName,'${escapedQuery}')`);
  // Page through all matches ($count=true requires ConsistencyLevel: eventual) rather than
  // silently returning only the first page.
  return fetchAllPages<User>(token, `/users?$filter=${filter}&$count=true`, 'Failed to search users', undefined, {
    headers: { ConsistencyLevel: 'eventual' }
  });
}

export async function searchGroups(token: string, query: string): Promise<GraphResponse<Group[]>> {
  const escapedQuery = query.replace(/'/g, "''");
  const filter = encodeURIComponent(`startswith(displayName,'${escapedQuery}')`);
  return fetchAllPages<Group>(token, `/groups?$filter=${filter}&$count=true`, 'Failed to search groups', undefined, {
    headers: { ConsistencyLevel: 'eventual' }
  });
}

export async function expandGroup(token: string, groupId: string): Promise<GraphResponse<User[]>> {
  const result = await fetchAllPages<any>(
    token,
    `/groups/${encodeURIComponent(groupId)}/members?$top=100`,
    'Failed to expand group'
  );

  if (!result.ok || !result.data) {
    return { ok: false, error: result.error };
  }

  const userMembers = result.data.filter((member: any) => {
    // Use the reliable @odata.type discriminator: a mail-enabled group / DL also carries a
    // `mail` property, so the old mail/UPN fallback misclassified nested groups as users.
    const odataType = typeof member['@odata.type'] === 'string' ? member['@odata.type'].toLowerCase() : '';
    if (odataType) return odataType.endsWith('.user');
    // Only when Graph omitted the discriminator do we fall back to a user-shaped heuristic.
    return typeof member.mail === 'string' || typeof member.userPrincipalName === 'string';
  }) as User[];

  return { ok: true, data: userMembers };
}
