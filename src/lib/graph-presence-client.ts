import { callGraph, GraphApiError, type GraphResponse, graphError, graphResult } from './graph-client.js';

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

export async function getUserPresence(token: string, userIdOrUpn: string): Promise<GraphResponse<UserPresence>> {
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

/** Up to 650 ids per call. Requires delegated `Presence.Read.All`. */
export async function getPresencesByUserIds(token: string, ids: string[]): Promise<GraphResponse<UserPresence[]>> {
  try {
    const clean = ids.map((x) => x.trim()).filter(Boolean);
    if (clean.length === 0) {
      return graphError('At least one user id is required');
    }
    if (clean.length > 650) {
      return graphError('Maximum 650 user ids per request');
    }
    const r = await callGraph<{ value: UserPresence[] }>(token, '/communications/getPresencesByUserId', {
      method: 'POST',
      body: JSON.stringify({ ids: clean })
    });
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to get presences by user id', r.error?.code, r.error?.status);
    }
    return graphResult(r.data.value ?? []);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get presences by user id');
  }
}

export interface SetPresencePayload {
  sessionId: string;
  availability: string;
  activity: string;
  expirationDuration?: string;
}

/** `userIdOrUpn` omit or empty → `POST /me/presence/setPresence`. Requires `Presence.ReadWrite`. */
export async function setUserPresence(
  token: string,
  payload: SetPresencePayload,
  userIdOrUpn?: string
): Promise<GraphResponse<void>> {
  try {
    const path = userIdOrUpn?.trim()
      ? `/users/${encodeURIComponent(userIdOrUpn.trim())}/presence/setPresence`
      : '/me/presence/setPresence';
    const r = await callGraph<void>(token, path, { method: 'POST', body: JSON.stringify(payload) }, false);
    if (!r.ok) {
      return graphError(r.error?.message || 'Failed to set presence', r.error?.code, r.error?.status);
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to set presence');
  }
}

/** Clear the app’s presence session (same `sessionId` as `setPresence`). `Presence.ReadWrite`. */
export async function clearPresenceSession(
  token: string,
  sessionId: string,
  userIdOrUpn?: string
): Promise<GraphResponse<void>> {
  try {
    const path = userIdOrUpn?.trim()
      ? `/users/${encodeURIComponent(userIdOrUpn.trim())}/presence/clearPresence`
      : '/me/presence/clearPresence';
    const r = await callGraph<void>(
      token,
      path,
      { method: 'POST', body: JSON.stringify({ sessionId: sessionId.trim() }) },
      false
    );
    if (!r.ok) {
      return graphError(r.error?.message || 'Failed to clear presence', r.error?.code, r.error?.status);
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to clear presence');
  }
}
