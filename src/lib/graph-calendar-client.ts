import {
  callGraph,
  fetchAllPages,
  GraphApiError,
  type GraphResponse,
  graphError,
  graphResult
} from './graph-client.js';
import { graphUserPath } from './graph-user-path.js';

/** Graph [calendar](https://learn.microsoft.com/en-us/graph/api/resources/calendar) (subset). */
export interface GraphCalendarResource {
  id: string;
  name?: string;
  color?: string;
  hexColor?: string;
  owner?: { name?: string; address?: string };
  canEdit?: boolean;
  canShare?: boolean;
  canViewPrivateItems?: boolean;
  defaultOnlineMeetingProvider?: string;
}

/** Graph [event](https://learn.microsoft.com/en-us/graph/api/resources/event) (subset). */
export interface GraphCalendarEvent {
  id: string;
  subject?: string;
  bodyPreview?: string;
  start?: { dateTime: string; timeZone: string };
  end?: { dateTime: string; timeZone: string };
  isAllDay?: boolean;
  isCancelled?: boolean;
  organizer?: { emailAddress?: { name?: string; address?: string } };
  attendees?: Array<{
    emailAddress?: { name?: string; address?: string };
    status?: { response?: string };
  }>;
  location?: { displayName?: string };
  webLink?: string;
  onlineMeeting?: { joinUrl?: string };
}

function calendarsRoot(user?: string): string {
  return graphUserPath(user, 'calendars');
}

export async function listCalendars(token: string, user?: string): Promise<GraphResponse<GraphCalendarResource[]>> {
  return fetchAllPages<GraphCalendarResource>(token, calendarsRoot(user), 'Failed to list calendars');
}

export async function getCalendar(
  token: string,
  calendarId: string,
  user?: string
): Promise<GraphResponse<GraphCalendarResource>> {
  try {
    const result = await callGraph<GraphCalendarResource>(
      token,
      `${calendarsRoot(user)}/${encodeURIComponent(calendarId)}`
    );
    if (!result.ok || !result.data) {
      return graphError(result.error?.message || 'Failed to get calendar', result.error?.code, result.error?.status);
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get calendar');
  }
}

/**
 * List events in a time range (Graph [calendarView](https://learn.microsoft.com/en-us/graph/api/calendar-list-calendarview)).
 * Omit `calendarId` to use the user's default calendar (`/me/calendar/calendarView`).
 */
export async function listCalendarView(
  token: string,
  startDateTime: string,
  endDateTime: string,
  options?: { calendarId?: string; user?: string }
): Promise<GraphResponse<GraphCalendarEvent[]>> {
  const params = new URLSearchParams();
  params.set('startDateTime', startDateTime);
  params.set('endDateTime', endDateTime);
  const qs = `?${params.toString()}`;
  const path = options?.calendarId
    ? `${calendarsRoot(options.user)}/${encodeURIComponent(options.calendarId)}/calendarView${qs}`
    : `${graphUserPath(options?.user, 'calendar/calendarView')}${qs}`;

  return fetchAllPages<GraphCalendarEvent>(token, path, 'Failed to list calendar view');
}

export async function getEvent(
  token: string,
  eventId: string,
  user?: string,
  select?: string
): Promise<GraphResponse<GraphCalendarEvent>> {
  let path = `${graphUserPath(user, `events/${encodeURIComponent(eventId)}`)}`;
  if (select?.trim()) {
    path += `?$select=${encodeURIComponent(select.trim())}`;
  }
  try {
    const result = await callGraph<GraphCalendarEvent>(token, path);
    if (!result.ok || !result.data) {
      return graphError(result.error?.message || 'Failed to get event', result.error?.code, result.error?.status);
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get event');
  }
}
