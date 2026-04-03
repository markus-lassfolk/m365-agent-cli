import { callGraph, GraphApiError, type GraphResponse, graphError, graphResult } from './graph-client.js';
import { graphUserPath } from './graph-user-path.js';

function onlineMeetingsRoot(user?: string): string {
  return graphUserPath(user, 'onlineMeetings');
}

/** Graph [onlineMeeting](https://learn.microsoft.com/en-us/graph/api/resources/onlinemeeting) (subset + common fields). */
export interface OnlineMeeting {
  id?: string;
  subject?: string;
  startDateTime?: string;
  endDateTime?: string;
  joinWebUrl?: string;
  joinUrl?: string;
  /** Conference bridge / dial-in (when returned by Graph). */
  conferenceId?: string;
  participants?: unknown;
  lobbyBypassSettings?: unknown;
  allowedPresenters?: string;
  allowAttendeeToEnableCamera?: boolean;
  allowAttendeeToEnableMic?: boolean;
  recordAutomatically?: boolean;
}

export async function createOnlineMeeting(
  token: string,
  body: { startDateTime: string; endDateTime: string; subject?: string },
  user?: string
): Promise<GraphResponse<OnlineMeeting>> {
  return createOnlineMeetingFromBody(token, body as Record<string, unknown>, user);
}

/** `POST /me/onlineMeetings` with a full Graph body ([user-post-onlineMeetings](https://learn.microsoft.com/en-us/graph/api/user-post-onlineMeetings)). */
export async function createOnlineMeetingFromBody(
  token: string,
  body: Record<string, unknown>,
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

/** `PATCH /me/onlineMeetings/{id}` ([update](https://learn.microsoft.com/en-us/graph/api/onlinemeeting-update)). */
export async function updateOnlineMeeting(
  token: string,
  meetingId: string,
  patch: Record<string, unknown>,
  user?: string
): Promise<GraphResponse<OnlineMeeting>> {
  try {
    const result = await callGraph<OnlineMeeting>(
      token,
      `${onlineMeetingsRoot(user)}/${encodeURIComponent(meetingId)}`,
      { method: 'PATCH', body: JSON.stringify(patch) }
    );
    if (!result.ok || !result.data) {
      return graphError(
        result.error?.message || 'Failed to update online meeting',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to update online meeting');
  }
}

/** `DELETE /me/onlineMeetings/{id}` ([delete](https://learn.microsoft.com/en-us/graph/api/onlinemeeting-delete)). */
export async function deleteOnlineMeeting(
  token: string,
  meetingId: string,
  user?: string
): Promise<GraphResponse<void>> {
  try {
    return await callGraph<void>(
      token,
      `${onlineMeetingsRoot(user)}/${encodeURIComponent(meetingId)}`,
      { method: 'DELETE' },
      false
    );
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to delete online meeting');
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
