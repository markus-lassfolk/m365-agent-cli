import {
  callGraph,
  callGraphAbsolute,
  fetchAllPages,
  fetchGraphRaw,
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
  /** Teams / Skype meeting details when `isOnlineMeeting` is true ([onlineMeetingInfo](https://learn.microsoft.com/en-us/graph/api/resources/onlinemeetinginfo)). */
  onlineMeeting?: {
    joinUrl?: string;
    conferenceId?: string;
    quickDial?: string;
    tollNumber?: string;
    tollFreeNumbers?: string[];
    phones?: unknown[];
  };
  changeKey?: string;
  sensitivity?: string;
  hasAttachments?: boolean;
  /** Present on calendarView / get event — user is the organizer */
  isOrganizer?: boolean;
  /** Meeting response for the signed-in user (when not organizer) */
  responseStatus?: { response?: string };
  /** Present on expanded instances — id of the recurring series master */
  seriesMasterId?: string;
  /** From calendarView / get event — `occurrence` | `seriesMaster` | `exception` | `singleInstance` */
  type?: 'singleInstance' | 'occurrence' | 'seriesMaster' | 'exception';
  /** Present on series master from GET event */
  recurrence?: GraphPatternedRecurrence | null;
}

/** Subset of [event resource](https://learn.microsoft.com/en-us/graph/api/resources/event) for POST /me/events. */
export interface GraphCreateEventRequest {
  subject: string;
  body?: { contentType: 'text' | 'html'; content: string };
  start: { dateTime: string; timeZone: string };
  end: { dateTime: string; timeZone: string };
  location?: { displayName: string };
  attendees?: Array<{
    emailAddress: { address: string; name?: string };
    type: 'required' | 'optional' | 'resource';
  }>;
  isOnlineMeeting?: boolean;
  onlineMeetingProvider?: 'teamsForBusiness' | 'unknown' | 'skypeForBusiness' | 'skypeForConsumer';
  isAllDay?: boolean;
  sensitivity?: 'normal' | 'personal' | 'private' | 'confidential';
  categories?: string[];
  recurrence?: GraphPatternedRecurrence;
}

/** [patternedRecurrence](https://learn.microsoft.com/en-us/graph/api/resources/patternedrecurrence) (subset). */
export interface GraphPatternedRecurrence {
  pattern: {
    type: 'daily' | 'weekly' | 'absoluteMonthly' | 'relativeMonthly' | 'absoluteYearly' | 'relativeYearly';
    interval: number;
    month?: number;
    dayOfMonth?: number;
    daysOfWeek?: string[];
    /** Week of month for relative patterns (Graph `index`). */
    index?: 'first' | 'second' | 'third' | 'fourth' | 'last';
    firstDayOfWeek?: 'sunday' | 'monday' | 'tuesday' | 'wednesday' | 'thursday' | 'friday' | 'saturday';
  };
  range: {
    type: 'endDate' | 'noEnd' | 'numbered';
    startDate: string;
    endDate?: string;
    numberOfOccurrences?: number;
  };
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

/**
 * Create an event on the default calendar ([create event](https://learn.microsoft.com/en-us/graph/api/user-post-events)).
 */
export async function createCalendarEvent(
  token: string,
  body: GraphCreateEventRequest,
  user?: string
): Promise<GraphResponse<GraphCalendarEvent>> {
  try {
    const result = await callGraph<GraphCalendarEvent>(token, graphUserPath(user, 'events'), {
      method: 'POST',
      body: JSON.stringify(body)
    });
    if (!result.ok || !result.data) {
      return graphError(result.error?.message || 'Failed to create event', result.error?.code, result.error?.status);
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to create event');
  }
}

/** PATCH [event](https://learn.microsoft.com/en-us/graph/api/event-update) (partial body). */
export async function updateCalendarEvent(
  token: string,
  eventId: string,
  patch: Record<string, unknown>,
  user?: string
): Promise<GraphResponse<GraphCalendarEvent>> {
  try {
    const result = await callGraph<GraphCalendarEvent>(
      token,
      graphUserPath(user, `events/${encodeURIComponent(eventId)}`),
      {
        method: 'PATCH',
        body: JSON.stringify(patch)
      }
    );
    if (!result.ok || !result.data) {
      return graphError(result.error?.message || 'Failed to update event', result.error?.code, result.error?.status);
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to update event');
  }
}

/** DELETE [event](https://learn.microsoft.com/en-us/graph/api/event-delete). */
export async function deleteCalendarEvent(token: string, eventId: string, user?: string): Promise<GraphResponse<void>> {
  try {
    return await callGraph<void>(
      token,
      graphUserPath(user, `events/${encodeURIComponent(eventId)}`),
      { method: 'DELETE' },
      false
    );
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to delete event');
  }
}

/** POST [cancel](https://learn.microsoft.com/en-us/graph/api/event-cancel) — notifies attendees. */
export async function cancelCalendarEvent(
  token: string,
  eventId: string,
  options?: { comment?: string; user?: string }
): Promise<GraphResponse<void>> {
  try {
    return await callGraph<void>(
      token,
      graphUserPath(options?.user, `events/${encodeURIComponent(eventId)}/cancel`),
      {
        method: 'POST',
        body: JSON.stringify({ comment: options?.comment ?? '' })
      },
      false
    );
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to cancel event');
  }
}

/**
 * List [event instances](https://learn.microsoft.com/en-us/graph/api/event-list-instances) for a **series master** in a time range.
 */
export async function listEventInstances(
  token: string,
  seriesMasterId: string,
  startDateTime: string,
  endDateTime: string,
  options?: { user?: string; select?: string }
): Promise<GraphResponse<GraphCalendarEvent[]>> {
  const params = new URLSearchParams();
  params.set('startDateTime', startDateTime);
  params.set('endDateTime', endDateTime);
  if (options?.select?.trim()) {
    params.set('$select', options.select.trim());
  }
  const qs = `?${params.toString()}`;
  const path = `${graphUserPath(options?.user, `events/${encodeURIComponent(seriesMasterId)}/instances`)}${qs}`;
  return fetchAllPages<GraphCalendarEvent>(token, path, 'Failed to list event instances');
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

/** Graph [calendarPermission](https://learn.microsoft.com/en-us/graph/api/resources/calendarpermission) (subset). */
export interface GraphCalendarPermission {
  id?: string;
  emailAddress?: { name?: string; address?: string };
  role?: string;
  allowedRoles?: string[];
  isRemovable?: boolean;
  isInsideOrganization?: boolean;
}

/**
 * List [calendar permissions](https://learn.microsoft.com/en-us/graph/api/calendar-list-calendarpermissions)
 * on the user's default calendar (sharing / delegate-style access).
 */
export async function listCalendarPermissions(
  token: string,
  user?: string
): Promise<GraphResponse<GraphCalendarPermission[]>> {
  return fetchAllPages<GraphCalendarPermission>(
    token,
    graphUserPath(user, 'calendar/calendarPermissions'),
    'Failed to list calendar permissions'
  );
}

/** Body for [POST calendarPermissions](https://learn.microsoft.com/en-us/graph/api/calendar-post-calendarpermissions). */
export interface CreateCalendarPermissionBody {
  emailAddress: { address: string; name?: string };
  role: string;
  isInsideOrganization?: boolean;
  isRemovable?: boolean;
}

/**
 * [Create calendarPermission](https://learn.microsoft.com/en-us/graph/api/calendar-post-calendarpermissions)
 * — share or delegate calendar access (Graph model; not EWS GetDelegates).
 */
export async function createCalendarPermission(
  token: string,
  body: CreateCalendarPermissionBody,
  user?: string
): Promise<GraphResponse<GraphCalendarPermission>> {
  try {
    const result = await callGraph<GraphCalendarPermission>(
      token,
      graphUserPath(user, 'calendar/calendarPermissions'),
      {
        method: 'POST',
        body: JSON.stringify(body)
      }
    );
    if (!result.ok || !result.data) {
      return graphError(
        result.error?.message || 'Failed to create calendar permission',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to create calendar permission');
  }
}

/**
 * [Update calendarPermission](https://learn.microsoft.com/en-us/graph/api/calendarpermission-update)
 * — `permissionId` is the Graph id (see `delegates list` or listCalendarPermissions).
 */
export async function updateCalendarPermission(
  token: string,
  permissionId: string,
  patch: { role: string },
  user?: string
): Promise<GraphResponse<GraphCalendarPermission>> {
  const id = encodeURIComponent(permissionId);
  try {
    const result = await callGraph<GraphCalendarPermission>(
      token,
      graphUserPath(user, `calendar/calendarPermissions/${id}`),
      {
        method: 'PATCH',
        body: JSON.stringify(patch)
      }
    );
    if (!result.ok || !result.data) {
      return graphError(
        result.error?.message || 'Failed to update calendar permission',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to update calendar permission');
  }
}

/**
 * [Delete calendarPermission](https://learn.microsoft.com/en-us/graph/api/calendarpermission-delete).
 */
export async function deleteCalendarPermission(
  token: string,
  permissionId: string,
  user?: string
): Promise<GraphResponse<void>> {
  const id = encodeURIComponent(permissionId);
  try {
    return await callGraph<void>(
      token,
      graphUserPath(user, `calendar/calendarPermissions/${id}`),
      { method: 'DELETE' },
      false
    );
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to delete calendar permission');
  }
}

/** Graph [attachment](https://learn.microsoft.com/en-us/graph/api/resources/attachment) on an event (subset). */
export interface GraphEventAttachment {
  id: string;
  name?: string;
  contentType?: string;
  size?: number;
  isInline?: boolean;
  '@odata.type'?: string;
  /** Present on `#microsoft.graph.referenceAttachment` */
  sourceUrl?: string;
}

/**
 * List attachments on a calendar event ([list attachments](https://learn.microsoft.com/en-us/graph/api/event-list-attachments)).
 */
export async function listEventAttachments(
  token: string,
  eventId: string,
  user?: string
): Promise<GraphResponse<GraphEventAttachment[]>> {
  return fetchAllPages<GraphEventAttachment>(
    token,
    `${graphUserPath(user, `events/${encodeURIComponent(eventId)}/attachments`)}`,
    'Failed to list event attachments'
  );
}

/** Download raw bytes for a **file** attachment (`GET .../attachments/{id}/$value`). */
export async function downloadEventAttachmentBytes(
  token: string,
  eventId: string,
  attachmentId: string,
  user?: string
): Promise<GraphResponse<Uint8Array>> {
  const path = `${graphUserPath(user, `events/${encodeURIComponent(eventId)}/attachments/${encodeURIComponent(attachmentId)}/$value`)}`;
  try {
    const res = await fetchGraphRaw(token, path);
    if (!res.ok) {
      let message = `Failed to download attachment: HTTP ${res.status}`;
      try {
        const json = (await res.json()) as { error?: { message?: string } };
        message = json.error?.message || message;
      } catch {
        // ignore
      }
      return graphError(message, undefined, res.status);
    }
    const buf = new Uint8Array(await res.arrayBuffer());
    return graphResult(buf);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to download attachment');
  }
}

/** Fetch a single attachment metadata (may include `contentBytes` for small file attachments). */
export async function getEventAttachment(
  token: string,
  eventId: string,
  attachmentId: string,
  user?: string
): Promise<GraphResponse<GraphEventAttachment & { contentBytes?: string }>> {
  try {
    const result = await callGraph<GraphEventAttachment & { contentBytes?: string }>(
      token,
      `${graphUserPath(user, `events/${encodeURIComponent(eventId)}/attachments/${encodeURIComponent(attachmentId)}`)}`
    );
    if (!result.ok || !result.data) {
      return graphError(result.error?.message || 'Failed to get attachment', result.error?.code, result.error?.status);
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get attachment');
  }
}

/** `POST /me/events/{id}/attachments` — file attachment on an event (draft or existing). */
export async function addFileAttachmentToCalendarEvent(
  token: string,
  eventId: string,
  attachment: { name: string; contentType: string; contentBytes: string },
  user?: string
): Promise<GraphResponse<GraphEventAttachment>> {
  const body = {
    '@odata.type': '#microsoft.graph.fileAttachment',
    name: attachment.name,
    contentType: attachment.contentType,
    contentBytes: attachment.contentBytes
  };
  try {
    const result = await callGraph<GraphEventAttachment>(
      token,
      `${graphUserPath(user, `events/${encodeURIComponent(eventId)}/attachments`)}`,
      { method: 'POST', body: JSON.stringify(body) },
      true
    );
    if (!result.ok || !result.data) {
      return graphError(
        result.error?.message || 'Failed to add event attachment',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to add event attachment');
  }
}

/** `POST /me/events/{id}/attachments` — link (`referenceAttachment`) on an event. */
async function addReferenceAttachmentToCalendarEvent(
  token: string,
  eventId: string,
  attachment: { name: string; sourceUrl: string },
  user?: string
): Promise<GraphResponse<GraphEventAttachment>> {
  const body = {
    '@odata.type': '#microsoft.graph.referenceAttachment',
    name: attachment.name,
    sourceUrl: attachment.sourceUrl
  };
  try {
    const result = await callGraph<GraphEventAttachment>(
      token,
      `${graphUserPath(user, `events/${encodeURIComponent(eventId)}/attachments`)}`,
      { method: 'POST', body: JSON.stringify(body) },
      true
    );
    if (!result.ok || !result.data) {
      return graphError(
        result.error?.message || 'Failed to add event link attachment',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to add event link attachment');
  }
}

/** Add multiple file + reference attachments to an event (sequential POSTs). */
export async function addCalendarEventAttachmentsGraph(
  token: string,
  eventId: string,
  user: string | undefined,
  files: Array<{ name: string; contentType: string; contentBytes: string }>,
  links: Array<{ name: string; sourceUrl: string }>
): Promise<GraphResponse<void>> {
  for (const f of files) {
    const r = await addFileAttachmentToCalendarEvent(token, eventId, f, user);
    if (!r.ok) {
      return graphError(r.error?.message || 'Failed to add file attachment', r.error?.code, r.error?.status);
    }
  }
  for (const l of links) {
    const r = await addReferenceAttachmentToCalendarEvent(token, eventId, l, user);
    if (!r.ok) {
      return graphError(r.error?.message || 'Failed to add link attachment', r.error?.code, r.error?.status);
    }
  }
  return { ok: true, data: undefined } as GraphResponse<void>;
}

/** One page of event delta sync ([delta](https://learn.microsoft.com/en-us/graph/delta-query-events)). */
export interface EventsDeltaPage {
  value?: GraphCalendarEvent[];
  '@odata.nextLink'?: string;
  '@odata.deltaLink'?: string;
}

export async function eventsDeltaPage(
  token: string,
  options?: { user?: string; calendarId?: string; nextLink?: string }
): Promise<GraphResponse<EventsDeltaPage>> {
  try {
    if (options?.nextLink?.trim()) {
      const result = await callGraphAbsolute<EventsDeltaPage>(token, options.nextLink.trim());
      if (!result.ok || !result.data) {
        return graphError(
          result.error?.message || 'Failed to fetch events delta page',
          result.error?.code,
          result.error?.status
        );
      }
      return graphResult(result.data);
    }
    const calId = options?.calendarId?.trim();
    const path = calId
      ? `${calendarsRoot(options?.user)}/${encodeURIComponent(calId)}/events/delta`
      : `${graphUserPath(options?.user, 'events')}/delta`;
    const result = await callGraph<EventsDeltaPage>(token, path);
    if (!result.ok || !result.data) {
      return graphError(
        result.error?.message || 'Failed to start events delta',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to fetch events delta');
  }
}
