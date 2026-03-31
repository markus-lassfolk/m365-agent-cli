import { callGraph, GraphApiError, graphError } from './graph-client.js';
import { graphUserPath } from './graph-user-path.js';

export type OofStatus = 'alwaysEnabled' | 'scheduled' | 'disabled';

export interface DateTimeTimeZone {
  dateTime: string;
  timeZone: string;
}

export interface AutomaticRepliesSetting {
  status: OofStatus;
  internalReplyMessage?: string;
  externalReplyMessage?: string;
  scheduledStartDateTime?: DateTimeTimeZone;
  scheduledEndDateTime?: DateTimeTimeZone;
}

export interface MailboxSettings {
  automaticRepliesSetting?: AutomaticRepliesSetting;
  timeZone?: string;
}

export interface GetMailboxSettingsResponse {
  automaticRepliesSetting?: AutomaticRepliesSetting;
  timeZone?: string;
}

export async function getMailboxSettings(token: string, user?: string): Promise<{
  ok: boolean;
  data?: GetMailboxSettingsResponse;
  error?: { message: string; code?: string; status?: number };
}> {
  try {
    return await callGraph<GetMailboxSettingsResponse>(token, graphUserPath(user, 'mailboxSettings'));
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status) as {
        ok: boolean;
        data?: GetMailboxSettingsResponse;
        error?: { message: string; code?: string; status?: number };
      };
    }
    return graphError(err instanceof Error ? err.message : 'Failed to get mailbox settings') as {
      ok: boolean;
      data?: GetMailboxSettingsResponse;
      error?: { message: string; code?: string; status?: number };
    };
  }
}

export async function setMailboxSettings(
  token: string,
  settings: Omit<Partial<AutomaticRepliesSetting>, 'scheduledStartDateTime' | 'scheduledEndDateTime'> & {
    scheduledStartDateTime?: string | DateTimeTimeZone;
    scheduledEndDateTime?: string | DateTimeTimeZone;
  },
  user?: string
): Promise<{
  ok: boolean;
  error?: { message: string; code?: string; status?: number };
}> {
  const payload = {
    automaticRepliesSetting: {
      ...(settings.status !== undefined ? { status: settings.status } : {}),
      ...(settings.internalReplyMessage !== undefined ? { internalReplyMessage: settings.internalReplyMessage } : {}),
      ...(settings.externalReplyMessage !== undefined ? { externalReplyMessage: settings.externalReplyMessage } : {}),
      ...(settings.scheduledStartDateTime !== undefined
        ? {
            scheduledStartDateTime:
              typeof settings.scheduledStartDateTime === 'string'
                ? { dateTime: settings.scheduledStartDateTime, timeZone: 'UTC' }
                : settings.scheduledStartDateTime
          }
        : {}),
      ...(settings.scheduledEndDateTime !== undefined
        ? {
            scheduledEndDateTime:
              typeof settings.scheduledEndDateTime === 'string'
                ? { dateTime: settings.scheduledEndDateTime, timeZone: 'UTC' }
                : settings.scheduledEndDateTime
          }
        : {})
    }
  };

  let result: any;
  try {
    result = await callGraph<Record<string, never>>(
      token,
      graphUserPath(user, 'mailboxSettings'),
      {
        method: 'PATCH',
        body: JSON.stringify(payload)
      },
      false // don't expect JSON on 204
    );
  } catch (err) {
    if (err instanceof GraphApiError) {
      return {
        ok: false,
        error: { message: err.message, code: err.code, status: err.status }
      };
    }
    return {
      ok: false,
      error: { message: err instanceof Error ? err.message : 'Failed to update mailbox settings' }
    };
  }

  if (!result.ok) {
    return {
      ok: false,
      error: result.error || { message: 'Failed to update mailbox settings' }
    };
  }

  return { ok: true };
}
