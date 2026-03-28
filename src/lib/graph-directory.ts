import { GraphResponse } from './graph-client.js';

const GRAPH_BASE_URL = process.env.GRAPH_BASE_URL || 'https://graph.microsoft.com/v1.0';

export interface Person {
  id: string;
  displayName: string;
  givenName?: string;
  surname?: string;
  title?: string;
  emailAddresses?: { address: string; name?: string }[];
  userPrincipalName?: string;
}

export interface User {
  id: string;
  displayName: string;
  givenName?: string;
  surname?: string;
  jobTitle?: string;
  mail?: string;
  userPrincipalName?: string;
}

export interface Group {
  id: string;
  displayName: string;
  description?: string;
  mail?: string;
  groupTypes?: string[];
}

export async function searchPeople(token: string, query: string): Promise<GraphResponse<Person[]>> {
  try {
    const url = new URL(`${GRAPH_BASE_URL}/me/people`);
    url.searchParams.set('$search', `"${query}"`);
    const res = await fetch(url.toString(), {
      headers: { Authorization: `Bearer ${token}` }
    });
    if (!res.ok) {
      return { ok: false, error: { status: res.status, message: res.statusText } };
    }
    const data = await res.json();
    return { ok: true, data: data.value as Person[] };
  } catch (err: any) {
    return { ok: false, error: { message: err.message } };
  }
}

export async function searchUsers(token: string, query: string): Promise<GraphResponse<User[]>> {
  try {
    const url = new URL(`${GRAPH_BASE_URL}/users`);
    url.searchParams.set('$filter', `startsWith(displayName,'${query}')`);
    
    const res = await fetch(url.toString(), {
      headers: { 
        Authorization: `Bearer ${token}`,
        ConsistencyLevel: 'eventual'
      }
    });
    if (!res.ok) {
      return { ok: false, error: { status: res.status, message: res.statusText } };
    }
    const data = await res.json();
    return { ok: true, data: data.value as User[] };
  } catch (err: any) {
    return { ok: false, error: { message: err.message } };
  }
}

export async function searchGroups(token: string, query: string): Promise<GraphResponse<Group[]>> {
  try {
    const url = new URL(`${GRAPH_BASE_URL}/groups`);
    url.searchParams.set('$filter', `startsWith(displayName,'${query}')`);
    const res = await fetch(url.toString(), {
      headers: { 
        Authorization: `Bearer ${token}`,
        ConsistencyLevel: 'eventual'
      }
    });
    if (!res.ok) {
      return { ok: false, error: { status: res.status, message: res.statusText } };
    }
    const data = await res.json();
    return { ok: true, data: data.value as Group[] };
  } catch (err: any) {
    return { ok: false, error: { message: err.message } };
  }
}

export async function expandGroup(token: string, groupId: string): Promise<GraphResponse<User[]>> {
  try {
    const url = new URL(`${GRAPH_BASE_URL}/groups/${groupId}/members`);
    const res = await fetch(url.toString(), {
      headers: { Authorization: `Bearer ${token}` }
    });
    if (!res.ok) {
      return { ok: false, error: { status: res.status, message: res.statusText } };
    }
    const data = await res.json();
    return { ok: true, data: data.value as User[] };
  } catch (err: any) {
    return { ok: false, error: { message: err.message } };
  }
}
