import { resolveGraphAuth } from './graph-auth.js';
import { callGraph, graphError, type GraphResponse } from './graph-client.js';

export interface Subscription {
  id: string;
  resource: string;
  applicationId?: string;
  changeType: string;
  clientState?: string;
  notificationUrl: string;
  expirationDateTime: string;
  creatorId?: string;
}

async function getAuthToken(token?: string): Promise<string> {
  const auth = await resolveGraphAuth({ token });
  if (!auth.success || !auth.token) {
    throw new Error(auth.error || 'Failed to authenticate to Graph API');
  }
  return auth.token;
}

export async function createSubscription(
  resource: string,
  changeType: string,
  notificationUrl: string,
  expirationDateTime: string,
  clientState?: string,
  token?: string
): Promise<GraphResponse<Subscription>> {
  try {
    const authToken = await getAuthToken(token);
    const body: Record<string, string> = {
      changeType,
      notificationUrl,
      resource,
      expirationDateTime
    };
    if (clientState) {
      body.clientState = clientState;
    }

    return await callGraph<Subscription>(authToken, '/subscriptions', {
      method: 'POST',
      body: JSON.stringify(body)
    });
  } catch (err: any) {
    return graphError(err.message, err.code, err.status);
  }
}

export async function listSubscriptions(token?: string): Promise<GraphResponse<Subscription[]>> {
  try {
    const authToken = await getAuthToken(token);
    const res = await callGraph<{ value: Subscription[] }>(authToken, '/subscriptions', {
      method: 'GET'
    });
    if (!res.ok) return res as unknown as GraphResponse<Subscription[]>;
    return { ok: true, data: res.data!.value };
  } catch (err: any) {
    return graphError(err.message, err.code, err.status);
  }
}

export async function deleteSubscription(id: string, token?: string): Promise<GraphResponse<void>> {
  try {
    const authToken = await getAuthToken(token);
    return await callGraph<void>(authToken, `/subscriptions/${encodeURIComponent(id)}`, {
      method: 'DELETE'
    });
  } catch (err: any) {
    return graphError(err.message, err.code, err.status);
  }
}

export async function renewSubscription(id: string, expirationDateTime: string, token?: string): Promise<GraphResponse<void>> {
  try {
    const authToken = await getAuthToken(token);
    return await callGraph<void>(authToken, `/subscriptions/${encodeURIComponent(id)}`, {
      method: 'PATCH',
      body: JSON.stringify({
        expirationDateTime
      })
    });
  } catch (err: any) {
    return graphError(err.message, err.code, err.status);
  }
}
