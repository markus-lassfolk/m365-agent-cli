import { callGraph, type GraphResponse, graphError, graphResult, GraphApiError } from './graph-client.js';

export interface GetScheduleRequest {
  schedules: string[];
  startTime: {
    dateTime: string;
    timeZone: string;
  };
  endTime: {
    dateTime: string;
    timeZone: string;
  };
  availabilityViewInterval?: number;
}

export interface ScheduleInformation {
  scheduleId: string;
  availabilityView?: string;
  scheduleItems?: Array<{
    isPrivate?: boolean;
    status?: string;
    subject?: string;
    location?: string;
    isMeeting?: boolean;
    isRecurring?: boolean;
    isException?: boolean;
    isReminderSet?: boolean;
    start?: {
      dateTime: string;
      timeZone: string;
    };
    end?: {
      dateTime: string;
      timeZone: string;
    };
  }>;
  workingHours?: {
    daysOfWeek?: string[];
    startTime?: string;
    endTime?: string;
    timeZone?: {
      name?: string;
    };
  };
  error?: {
    message?: string;
    responseCode?: string;
  };
}

export interface GetScheduleResponse {
  value: ScheduleInformation[];
}

export async function getSchedule(
  token: string,
  request: GetScheduleRequest
): Promise<GraphResponse<GetScheduleResponse>> {
  let result: GraphResponse<GetScheduleResponse>;
  try {
    result = await callGraph<GetScheduleResponse>(token, '/me/calendar/getSchedule', {
      method: 'POST',
      body: JSON.stringify(request),
      headers: {
        Prefer: 'outlook.timezone="UTC"'
      }
    });
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to get schedule');
  }

  if (!result.ok || !result.data) {
    return graphError(result.error?.message || 'Failed to get schedule', result.error?.code, result.error?.status);
  }
  return graphResult(result.data);
}

export interface AttendeeBase {
  type: 'required' | 'optional' | 'resource';
  emailAddress: {
    address: string;
    name?: string;
  };
}

export interface TimeConstraint {
  activityDomain?: 'work' | 'personal' | 'unrestricted' | 'unknown';
  timeSlots: Array<{
    start: {
      dateTime: string;
      timeZone: string;
    };
    end: {
      dateTime: string;
      timeZone: string;
    };
  }>;
}

export interface FindMeetingTimesRequest {
  attendees?: AttendeeBase[];
  locationConstraint?: {
    isRequired?: boolean;
    suggestLocation?: boolean;
    locations?: Array<{
      resolveAvailability?: boolean;
      displayName?: string;
      locationEmailAddress?: string;
    }>;
  };
  timeConstraint?: TimeConstraint;
  meetingDuration?: string; // e.g. "PT30M"
  returnSuggestionReasons?: boolean;
  minimumAttendeePercentage?: number;
  isOrganizerOptional?: boolean;
  maxCandidates?: number;
}

export interface MeetingTimeSuggestion {
  meetingTimeSlot?: {
    start?: { dateTime: string; timeZone: string };
    end?: { dateTime: string; timeZone: string };
  };
  confidence?: number;
  organizerAvailability?: string;
  attendeeAvailability?: Array<{
    availability?: string;
    attendee?: AttendeeBase;
  }>;
  locations?: Array<{ displayName?: string; locationEmailAddress?: string }>;
  suggestionReason?: string;
}

export interface FindMeetingTimesResponse {
  emptySuggestionsReason?: string;
  meetingTimeSuggestions?: MeetingTimeSuggestion[];
}

export async function findMeetingTimes(
  token: string,
  request: FindMeetingTimesRequest
): Promise<GraphResponse<FindMeetingTimesResponse>> {
  let result: GraphResponse<FindMeetingTimesResponse>;
  try {
    result = await callGraph<FindMeetingTimesResponse>(token, '/me/findMeetingTimes', {
      method: 'POST',
      body: JSON.stringify(request),
      headers: {
        Prefer: 'outlook.timezone="UTC"'
      }
    });
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to find meeting times');
  }

  if (!result.ok || !result.data) {
    return graphError(
      result.error?.message || 'Failed to find meeting times',
      result.error?.code,
      result.error?.status
    );
  }
  return graphResult(result.data);
}
