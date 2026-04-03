import { callGraph, GraphApiError, type GraphResponse, graphError, graphResult } from './graph-client.js';
import { graphUserPath } from './graph-user-path.js';

function onlineMeetingsRoot(user?: string): string {
  return graphUserPath(user, 'onlineMeetings');
}

/** Graph [onlineMeeting](https://learn.microsoft.com/en-us/graph/api/resources/onlinemeeting) (subset). */
export interface OnlineMeeting {
  id?: string;
  subject?: string;
  startDateTime?: string;
  endDateTime?: string;
  joinWebUrl?: string;
  joinUrl?: string;
}

export async function createOnlineMeeting(
  token: string,
  body: { startDateTime: string; endDateTime: string; subject?: string },
  user?: string
): Promise<GraphResponse<OnlineMeeting>> {
  try {
    const result = await callGraph<OnlineMeeting>(token, onlineMeetingsRoot(user), {
      method: 'POST',
      body: JSON.stringify(body)
    });
    if (!result.ok || !result.data) {
      return graphError(
        result.error?.message || 'Failed to create online meeting',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to create online meeting');
  }
}

export async function getOnlineMeeting(
  token: string,
  meetingId: string,
  user?: string
): Promise<GraphResponse<OnlineMeeting>> {
  try {
    const result = await callGraph<OnlineMeeting>(
      token,
      `${onlineMeetingsRoot(user)}/${encodeURIComponent(meetingId)}`
    );
    if (!result.ok || !result.data) {
      return graphError(
        result.error?.message || 'Failed to get online meeting',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get online meeting');
  }
}
