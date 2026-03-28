import { callGraph, GRAPH_BASE_URL, type GraphResponse } from './graph-client.js';

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
  const result = await callGraph<{ value: Person[] }>(token, `/me/people?$search=${searchParam}`);
  if (!result.ok || !result.data) {
    return { ok: false, error: result.error };
  }
  return { ok: true, data: result.data.value };
}

export async function searchUsers(token: string, query: string): Promise<GraphResponse<User[]>> {
  const escapedQuery = query.replace(/'/g, "''");
  const filter = encodeURIComponent(`startsWith(displayName,'${escapedQuery}')`);
  const result = await callGraph<{ value: User[] }>(token, `/users?$filter=${filter}&$count=true`, {
    headers: {
      ConsistencyLevel: 'eventual'
    }
  });
  if (!result.ok || !result.data) {
    return { ok: false, error: result.error };
  }
  return { ok: true, data: result.data.value };
}

export async function searchGroups(token: string, query: string): Promise<GraphResponse<Group[]>> {
  const escapedQuery = query.replace(/'/g, "''");
  const filter = encodeURIComponent(`startsWith(displayName,'${escapedQuery}')`);
  const result = await callGraph<{ value: Group[] }>(token, `/groups?$filter=${filter}&$count=true`, {
    headers: {
      ConsistencyLevel: 'eventual'
    }
  });
  if (!result.ok || !result.data) {
    return { ok: false, error: result.error };
  }
  return { ok: true, data: result.data.value };
}

export async function expandGroup(token: string, groupId: string): Promise<GraphResponse<User[]>> {
  const members: User[] = [];
  let path = `/groups/${encodeURIComponent(groupId)}/members?$top=100`;

  while (path) {
    const result = await callGraph<{ value: any[]; '@odata.nextLink'?: string }>(token, path);
    if (!result.ok || !result.data) {
      return { ok: false, error: result.error };
    }

    const userMembers = result.data.value.filter((member: any) => {
      const odataType = member['@odata.type'];
      return (
        (typeof odataType === 'string' && odataType.toLowerCase().endsWith('.user')) ||
        typeof member.mail === 'string' ||
        typeof member.userPrincipalName === 'string'
      );
    }) as User[];

    members.push(...userMembers);
    path = result.data['@odata.nextLink']
      ? result.data['@odata.nextLink'].replace(
          new RegExp(`^${GRAPH_BASE_URL.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}`),
          ''
        )
      : '';
  }

  return { ok: true, data: members };
}
