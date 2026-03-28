/**
 * Mock fetch setup for CLI integration tests.
 * Intercepts all HTTP calls and returns mock responses based on URL and request body.
 */
import {
  mockOAuthTokenResponse,
  mockResolveNamesResponse,
  mockCalendarEventsResponse,
  mockCalendarEventsEmptyResponse,
  mockCalendarEventsWithCancelled,
  mockGetScheduleResponse,
  mockGetScheduleEmptyResponse,
  mockRespondListResponse,
  mockRespondSuccessResponse,
  mockCreateEventResponse,
  mockDeleteEventSuccessResponse,
  mockCancelEventSuccessResponse,
  mockResolveNamesPeopleResponse,
  mockResolveNamesRoomsResponse,
  mockResolveNamesEmptyResponse,
  mockUpdateEventResponse,
  mockGetEmailsResponse,
  mockGetEmailsEmptyResponse,
  mockGetEmailDetailResponse,
  mockGetAttachmentsResponse,
  mockUpdateEmailResponse,
  mockMoveEmailResponse,
  mockSendEmailResponse,
  mockReplyToEmailResponse,
  mockForwardEmailResponse,
  mockGetMailFoldersResponse,
  mockCreateMailFolderResponse,
  mockUpdateMailFolderResponse,
  mockDeleteMailFolderResponse,
  mockGetDraftsResponse,
  mockCreateDraftResponse,
  mockUpdateDraftResponse,
  mockSendDraftResponse,
  mockDeleteDraftResponse,
  mockAddAttachmentResponse,
  mockGetRoomsResponse,
  mockGetRoomsFromListResponse,
  mockSearchRoomsResponse,
  mockIsRoomFreeResponse,
  mockGraphListFilesResponse,
  mockGraphSearchFilesResponse,
  mockGraphGetFileMetadataResponse,
  mockGraphUploadResponse,
  mockGraphDeleteResponse,
  mockGraphShareResponse,
  mockGraphCollabResponse,
  mockGraphCheckinResponse,
  mockGraphCreateUploadSessionResponse
} from './responses.js';

type MockFn = (url: string, request: Request) => { status: number; body: string; contentType: string } | null;

let mockFetchImpl: MockFn | null = null;

export function setMockFetch(impl: MockFn): void {
  mockFetchImpl = impl;
}

export function clearMockFetch(): void {
  mockFetchImpl = null;
}

function makeResponse(body: string, status = 200, contentType = 'text/xml'): Response {
  return new Response(body, {
    status,
    headers: { 'content-type': contentType }
  });
}

function makeJsonResponse(body: object, status = 200): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'content-type': 'application/json' }
  });
}

function extractSoapAction(body: string): string {
  // Extract the first child element of soap:Body
  const match = body.match(/<soap:Body[^>]*>([\s\S]*?)<\/soap:Body>/i);
  if (!match) return '';
  const inner = match[1].trim();
  // Find the first tag name
  const tagMatch = inner.match(/<(?:m:|t:)?(\w+)/);
  return tagMatch ? tagMatch[1] : '';
}

function hasTag(body: string, tag: string): boolean {
  return body.includes(`<${tag}`) || body.includes(`<m:${tag}`) || body.includes(`<t:${tag}`);
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
export function createMockFetch(): any {
  return async (input: string | URL | Request, init?: RequestInit) => {
    const url = typeof input === 'string' ? input : input instanceof URL ? input.toString() : input.url;
    const body = typeof init?.body === 'string' ? init.body : '';

    // Let custom mock take priority
    if (mockFetchImpl) {
      const custom = mockFetchImpl(url, new Request(url, init as RequestInit));
      if (custom) return makeResponse(custom.body, custom.status, custom.contentType);
    }

    // OAuth token endpoint
    if (url.includes('login.microsoftonline.com') && url.includes('/token')) {
      return makeJsonResponse(JSON.parse(mockOAuthTokenResponse));
    }

    // EWS endpoint
    if (url.includes('outlook.office365.com/EWS/Exchange.asmx')) {
      const action = extractSoapAction(body);

      // Auth check / ResolveNames (used by whoami and find)
      if (action === 'ResolveNames' && hasTag(body, 'UnresolvedEntry')) {
        // Check if this is calendar-related (respond list uses DistinguishedFolderId)
        if (hasTag(body, 'DistinguishedFolderId') && hasTag(body, 'CalendarView')) {
          return makeResponse(mockRespondListResponse);
        }

        // Distinguish whoami from find:
        // - whoami: getOwaUserInfo calls ResolveNames with EWS_USERNAME (empty in tests)
        // - find: resolveNames calls ResolveNames with the search query
        // whoami has empty UnresolvedEntry + no RequiredAttendees
        // find has (non-empty UnresolvedEntry) OR (has RequiredAttendees)
        const unresolvedContent = body.match(/<m:UnresolvedEntry>([^<]*)<\/m:UnresolvedEntry>/)?.[1] || '';
        const isPeopleSearch = hasTag(body, 'RequiredAttendees') || unresolvedContent.length > 0;

        if (isPeopleSearch) {
          return makeResponse(mockResolveNamesPeopleResponse);
        }
        // whoami gets here (empty UnresolvedEntry, no RequiredAttendees)
        return makeResponse(mockResolveNamesResponse);
      }

      // Calendar events (used by calendar, respond, delete-event, update-event)
      if (hasTag(body, 'FindItem') && hasTag(body, 'CalendarView')) {
        // Check for specific event IDs (for respond)
        if (hasTag(body, 'invite-')) {
          return makeResponse(mockRespondListResponse);
        }
        // Check if looking for cancelled events
        if (hasTag(body, 'Cancelled')) {
          return makeResponse(mockCalendarEventsWithCancelled);
        }
        // Default calendar events
        return makeResponse(mockCalendarEventsResponse);
      }

      // Create calendar item
      if (hasTag(body, 'CreateItem') && hasTag(body, 'CalendarItem')) {
        return makeResponse(mockCreateEventResponse);
      }

      // Update calendar item
      if (hasTag(body, 'UpdateItem') && hasTag(body, 'CalendarItem')) {
        return makeResponse(mockUpdateEventResponse);
      }

      // Delete calendar item
      if (hasTag(body, 'DeleteItem')) {
        return makeResponse(mockDeleteEventSuccessResponse);
      }

      // Cancel calendar item
      if (hasTag(body, 'CancelItem')) {
        return makeResponse(mockCancelEventSuccessResponse);
      }

      // Respond to item
      if (
        hasTag(body, 'RespondToItem') ||
        hasTag(body, 'AcceptItem') ||
        hasTag(body, 'DeclineItem') ||
        hasTag(body, 'TentativelyAcceptItem')
      ) {
        return makeResponse(mockRespondSuccessResponse);
      }

      // FindItem for mail folders / emails
      if (hasTag(body, 'FindItem') && hasTag(body, 'ItemShape')) {
        // Check if it's a drafts query
        if (hasTag(body, 'drafts') || body.includes('Drafts')) {
          return makeResponse(mockGetDraftsResponse);
        }
        if (hasTag(body, 'sentitems') || body.includes('SentItems')) {
          return makeResponse(mockGetEmailsResponse);
        }
        return makeResponse(mockGetEmailsResponse);
      }

      // GetItem for email detail
      if (hasTag(body, 'GetItem')) {
        return makeResponse(mockGetEmailDetailResponse);
      }

      // GetAttachment
      if (hasTag(body, 'GetAttachment')) {
        return makeResponse(mockGetAttachmentsResponse);
      }

      // UpdateItem for email (mark read/unread/flag)
      if (hasTag(body, 'UpdateItem') && hasTag(body, 'Message')) {
        return makeResponse(mockUpdateEmailResponse);
      }

      // MoveItem
      if (hasTag(body, 'MoveItem')) {
        return makeResponse(mockMoveEmailResponse);
      }

      // SendItem
      if (hasTag(body, 'SendItem')) {
        return makeResponse(mockSendEmailResponse);
      }

      // CreateItem (reply/forward draft)
      if ((hasTag(body, 'CreateItem') && hasTag(body, 'ReplyToItem')) || hasTag(body, 'ForwardItem')) {
        return makeResponse(mockReplyToEmailResponse);
      }

      // GetFolder (mail folders)
      if (hasTag(body, 'GetFolder') || (hasTag(body, 'FindFolder') && hasTag(body, 'DistinguishedFolderId'))) {
        return makeResponse(mockGetMailFoldersResponse);
      }

      // CreateFolder
      if (hasTag(body, 'CreateFolder')) {
        return makeResponse(mockCreateMailFolderResponse);
      }

      // UpdateFolder (rename)
      if (hasTag(body, 'UpdateFolder')) {
        return makeResponse(mockUpdateMailFolderResponse);
      }

      // DeleteFolder
      if (hasTag(body, 'DeleteFolder')) {
        return makeResponse(mockDeleteMailFolderResponse);
      }

      // GetRoomLists
      if (hasTag(body, 'GetRoomLists')) {
        return makeResponse(mockGetRoomsResponse);
      }

      // GetRooms
      if (hasTag(body, 'GetRooms')) {
        return makeResponse(mockGetRoomsFromListResponse);
      }

      // ExpandDL (search rooms)
      if (hasTag(body, 'ExpandDL')) {
        return makeResponse(mockSearchRoomsResponse);
      }

      // GetSchedule (findtime)
      if (hasTag(body, 'GetSchedule')) {
        return makeResponse(mockGetScheduleResponse);
      }

      // CreateAttachment
      if (hasTag(body, 'CreateAttachment')) {
        return makeResponse(mockAddAttachmentResponse);
      }

      // Default: return empty calendar
      return makeResponse(mockCalendarEventsEmptyResponse);
    }

    // Microsoft Graph API (files commands)
    if (url.includes('graph.microsoft.com/v1.0')) {
      // List files
      if (url.includes('/me/drive/root/children') || (url.includes('/me/drive/items') && url.includes('/children'))) {
        return makeJsonResponse(mockGraphListFilesResponse);
      }
      // Search files
      if (url.includes('/me/drive/root/search')) {
        return makeJsonResponse(mockGraphSearchFilesResponse);
      }
      // Upload file
      if (url.includes('/me/drive/items/') && url.includes('/content')) {
        return makeJsonResponse(mockGraphUploadResponse);
      }
      // Get file metadata
      if (url.includes('/me/drive/items/') && !url.includes('/children') && !url.includes('/content') && !url.includes('/createLink') && !url.includes('/checkin') && !url.includes('/checkout') && !url.includes('/createUploadSession')) {
        return makeJsonResponse(mockGraphGetFileMetadataResponse);
      }
      // Create upload session
      if (url.includes('/createUploadSession')) {
        return makeJsonResponse(mockGraphCreateUploadSessionResponse);
      }
      // Delete file
      if ((url.includes('/me/drive/items/') || url.includes('/me/drive/')) && init?.method === 'DELETE') {
        return makeJsonResponse(mockGraphDeleteResponse);
      }
      // Share file / Office collaboration
      if (url.includes('/me/drive/items/') && url.includes('/createLink')) {
        return makeJsonResponse(mockGraphShareResponse);
      }
      // Checkin
      if (url.includes('/me/drive/items/') && url.includes('/checkin')) {
        return makeJsonResponse(mockGraphCheckinResponse);
      }
      // Checkout
      if (url.includes('/me/drive/items/') && url.includes('/checkout')) {
        return makeJsonResponse({});
      }
      return makeJsonResponse({ value: [] });
    }

    // Default: 404
    return new Response('Not found', { status: 404 });
  };
}

// Setup/teardown helpers for use in tests
export function setupMockFetch(): void {
  // @ts-ignore - globalThis.fetch
  // @ts-ignore
  globalThis.fetch = createMockFetch() as typeof fetch;
}

export function teardownMockFetch(): void {
  // @ts-ignore - globalThis.fetch
  globalThis.fetch = undefined;
}
