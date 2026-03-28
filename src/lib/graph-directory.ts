import { callGraph, type GraphResponse, fetchAllPages } from './graph-client.js';

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
  const filter = encodeURIComponent(`startswith(displayName,'${escapedQuery}')`);
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
  const filter = encodeURIComponent(`startswith(displayName,'${escapedQuery}')`);
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
  const result = await fetchAllPages<any>(
    token,
    `/groups/${encodeURIComponent(groupId)}/members?$top=100`,
    'Failed to expand group'
  );

  if (!result.ok || !result.data) {
    return { ok: false, error: result.error };
  }

  const userMembers = result.data.filter((member: any) => {
    const odataType = member['@odata.type'];
    return (
      (typeof odataType === 'string' && odataType.toLowerCase().endsWith('.user')) ||
      typeof member.mail === 'string' ||
      typeof member.userPrincipalName === 'string'
    );
  }) as User[];

  return { ok: true, data: userMembers };
}
