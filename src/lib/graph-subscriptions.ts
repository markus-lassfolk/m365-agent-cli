import { resolveGraphAuth } from './graph-auth.js';
import { fetchGraphRaw } from './graph-client.js';

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

async function fetchGraph(endpoint: string, options: RequestInit = {}, token?: string): Promise<Response> {
  const auth = await resolveGraphAuth({ token });
  if (!auth.success || !auth.token) {
    throw new Error(auth.error || 'Failed to authenticate to Graph API');
  }

  return fetchGraphRaw(auth.token, endpoint, {
    ...options,
    headers: {
      ...options.headers,
      'Content-Type': 'application/json'
    }
  });
}

export async function createSubscription(
  resource: string,
  changeType: string,
  notificationUrl: string,
  expirationDateTime: string,
  clientState?: string,
  token?: string
): Promise<Subscription> {
  const body: Record<string, string> = {
    changeType,
    notificationUrl,
    resource,
    expirationDateTime
  };
  if (clientState) {
    body.clientState = clientState;
  }

  const response = await fetchGraph('/subscriptions', {
    method: 'POST',
    body: JSON.stringify(body)
  }, token);

  if (!response.ok) {
    const error = await response.text();
    throw new Error(`Failed to create subscription: ${response.status} ${response.statusText} - ${error}`);
  }

  return response.json() as Promise<Subscription>;
}

export async function listSubscriptions(token?: string): Promise<Subscription[]> {
  const response = await fetchGraph('/subscriptions', {
    method: 'GET'
  }, token);

  if (!response.ok) {
    const error = await response.text();
    throw new Error(`Failed to list subscriptions: ${response.status} ${response.statusText} - ${error}`);
  }

  const data = (await response.json()) as { value: Subscription[] };
  return data.value;
}

export async function deleteSubscription(id: string, token?: string): Promise<void> {
  const response = await fetchGraph(`/subscriptions/${encodeURIComponent(id)}`, {
    method: 'DELETE'
  }, token);

  if (!response.ok) {
    const error = await response.text();
    throw new Error(`Failed to delete subscription: ${response.status} ${response.statusText} - ${error}`);
  }
}

export async function renewSubscription(id: string, expirationDateTime: string, token?: string): Promise<void> {
  const response = await fetchGraph(`/subscriptions/${encodeURIComponent(id)}`, {
    method: 'PATCH',
    body: JSON.stringify({
      expirationDateTime
    })
  }, token);

  if (!response.ok) {
    const error = await response.text();
    throw new Error(`Failed to renew subscription: ${response.status} ${response.statusText} - ${error}`);
  }
}
