import {
  callGraph,
  GraphApiError,
  type GraphResponse,
  graphError,
  graphResult
} from './graph-client.js';

export interface UserPresence {
  id?: string;
  availability?: string;
  activity?: string;
  outOfOfficeSettings?: { message?: { content?: string } };
}

export async function getMyPresence(token: string): Promise<GraphResponse<UserPresence>> {
  try {
    const r = await callGraph<UserPresence>(token, '/me/presence');
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to get presence', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get presence');
  }
}

export async function getUserPresence(
  token: string,
  userIdOrUpn: string
): Promise<GraphResponse<UserPresence>> {
  try {
    const enc = encodeURIComponent(userIdOrUpn.trim());
    const r = await callGraph<UserPresence>(token, `/users/${enc}/presence`);
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to get presence', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get presence');
  }
}
