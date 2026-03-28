import { callGraph, type GraphResponse } from './graph-client.js';

export interface Recipient {
  emailAddress: {
    address: string;
    name?: string;
  };
}

export interface ForwardEventOptions {
  token: string;
  eventId: string;
  toRecipients: string[];
  comment?: string;
}

export async function forwardEvent(options: ForwardEventOptions): Promise<GraphResponse<void>> {
  const { token, eventId, toRecipients, comment } = options;

  const recipientsList = toRecipients.map((email) => ({
    EmailAddress: { Address: email }
  }));

  const body: any = {
    ToRecipients: recipientsList
  };

  if (comment) {
    body.Comment = comment;
  }

  return callGraph<void>(
    token,
    `/me/events/${encodeURIComponent(eventId)}/forward`,
    {
      method: 'POST',
      body: JSON.stringify(body)
    },
    false
  );
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

  return callGraph<void>(
    token,
    `/me/events/${encodeURIComponent(eventId)}/tentativelyAccept`,
    {
      method: 'POST',
      body: JSON.stringify(body)
    },
    false
  );
}
