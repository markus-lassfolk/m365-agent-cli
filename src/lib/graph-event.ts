import { callGraph, graphError, type GraphResponse, GraphApiError } from './graph-client.js';

export interface ForwardEventOptions {
  token: string;
  eventId: string;
  toRecipients: string[];
  comment?: string;
}

export async function forwardEvent(options: ForwardEventOptions): Promise<GraphResponse<void>> {
  const { token, eventId, toRecipients, comment } = options;

  const recipientsList = toRecipients.map((email) => ({
    emailAddress: { address: email }
  }));

  const body: any = {
    toRecipients: recipientsList
  };

  if (comment) {
    body.comment = comment;
  }

  try {
    return await callGraph<void>(
      token,
      `/me/events/${encodeURIComponent(eventId)}/forward`,
      {
        method: 'POST',
        body: JSON.stringify(body)
      },
      false
    );
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to forward event');
  }
}

export interface ProposeNewTimeOptions {
  token: string;
  eventId: string;
  startDateTime: string;
  endDateTime: string;
  timeZone?: string;
}

export async function proposeNewTime(options: ProposeNewTimeOptions): Promise<GraphResponse<void>> {
  const { token, eventId, startDateTime, endDateTime, timeZone = 'UTC' } = options;

  const body = {
    proposedNewTime: {
      start: { dateTime: startDateTime, timeZone },
      end: { dateTime: endDateTime, timeZone }
    },
    sendResponse: true
  };

  try {
    return await callGraph<void>(
      token,
      `/me/events/${encodeURIComponent(eventId)}/tentativelyAccept`,
      {
        method: 'POST',
        body: JSON.stringify(body)
      },
      false
    );
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to propose new time');
  }
}
