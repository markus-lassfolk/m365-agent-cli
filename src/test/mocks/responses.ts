/**
 * Mock API responses for CLI integration testing.
 * These mirror the exact XML/JSON structures that the EWS and Graph clients parse.
 */

const SOAP_NS =
  'xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"';

function soapResponse(body: string): string {
  return `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope ${SOAP_NS}>
  <soap:Header><t:RequestServerVersion Version="Exchange2016" /></soap:Header>
  <soap:Body>
    <m:ResponseCode>NoError</m:ResponseCode>
    ${body}
  </soap:Body>
</soap:Envelope>`;
}

// ─── OAuth ────────────────────────────────────────────────────────────────

export const mockOAuthTokenResponse = JSON.stringify({
  access_token: 'mock-access-token-12345',
  refresh_token: 'mock-refresh-token-12345',
  expires_in: 3600,
  token_type: 'Bearer'
});

// ─── whoami ───────────────────────────────────────────────────────────────

export const mockResolveNamesResponse = soapResponse(`
  <m:ResolutionSet>
    <m:Resolution>
      <t:Mailbox>
        <t:Name>Test User</t:Name>
        <t:EmailAddress>test@example.com</t:EmailAddress>
        <t:RoutingType>SMTP</t:RoutingType>
        <t:MailboxType>Mailbox</t:MailboxType>
      </t:Mailbox>
    </m:Resolution>
  </m:ResolutionSet>
`);

// ─── calendar ─────────────────────────────────────────────────────────────

function makeCalendarItem(opts: {
  id: string;
  changeKey?: string;
  subject: string;
  start: string;
  end: string;
  location?: string;
  isAllDay?: boolean;
  isCancelled?: boolean;
  isOrganizer?: boolean;
  organizerName?: string;
  organizerEmail?: string;
  myResponseType?: string;
  attendees?: string;
  categories?: string;
}): string {
  const ck = opts.changeKey ?? 'mockChangeKey1';
  const loc = opts.location ? `<t:Location>${opts.location}</t:Location>` : '';
  const allDay = opts.isAllDay ? '<t:IsAllDayEvent>true</t:IsAllDayEvent>' : '';
  const canc = opts.isCancelled ? '<t:IsCancelled>true</t:IsCancelled>' : '';
  const orgBlock = opts.organizerName
    ? `<t:Organizer><t:Mailbox><t:Name>${opts.organizerName}</t:Name><t:EmailAddress>${opts.organizerEmail ?? 'test@example.com'}</t:EmailAddress></t:Mailbox></t:Organizer>`
    : '';
  const respType = opts.myResponseType ? `<t:MyResponseType>${opts.myResponseType}</t:MyResponseType>` : '';
  const cats = opts.categories ? `<t:Categories>${opts.categories}</t:Categories>` : '';
  return `
  <t:CalendarItem>
    <t:ItemId Id="${opts.id}" ChangeKey="${ck}" />
    <t:Subject>${opts.subject}</t:Subject>
    <t:Start>${opts.start}</t:Start>
    <t:End>${opts.end}</t:End>
    ${loc}
    ${allDay}
    ${canc}
    ${orgBlock}
    ${respType}
    <t:Importance>Normal</t:Importance>
    <t:LegacyFreeBusyStatus>Busy</t:LegacyFreeBusyStatus>
    ${opts.attendees ?? ''}
    ${cats}
    <t:TextBody BodyType="Text">Meeting description</t:TextBody>
  </t:CalendarItem>`;
}

/** GetItem response for calendar event IDs (respond/cancel/delete prefetch). */
export function makeGetCalendarItemDetailResponse(itemId: string): string {
  return soapResponse(`
  <m:GetItemResponse>
    <m:ResponseCode>NoError</m:ResponseCode>
    <m:Items>
      ${makeCalendarItem({
        id: itemId,
        subject: 'Calendar item',
        start: '2026-03-30T10:00:00Z',
        end: '2026-03-30T11:00:00Z',
        isOrganizer: true,
        myResponseType: 'Organizer'
      })}
    </m:Items>
  </m:GetItemResponse>
`);
}

const MOCK_CALENDAR_RESPONSE = soapResponse(`
  <m:FindItemResponse>
    <m:ResponseCode>NoError</m:ResponseCode>
    <m:RootFolder>
      <t:TotalItemsInView>2</t:TotalItemsInView>
      <m:Items>
        ${makeCalendarItem({
          id: 'event-1',
          subject: 'Team Standup',
          start: '2026-03-30T09:00:00Z',
          end: '2026-03-30T09:30:00Z',
          location: 'Conference Room A',
          isOrganizer: true,
          myResponseType: 'Organizer',
          attendees: `
            <t:RequiredAttendees>
              <t:Attendee><t:Mailbox><t:Name>Alice</t:Name><t:EmailAddress>alice@example.com</t:EmailAddress></t:Mailbox><t:ResponseType>Accept</t:ResponseType></t:Attendee>
              <t:Attendee><t:Mailbox><t:Name>Bob</t:Name><t:EmailAddress>bob@example.com</t:EmailAddress></t:Mailbox><t:ResponseType>NoResponseReceived</t:ResponseType></t:Attendee>
            </t:RequiredAttendees>`
        })}
        ${makeCalendarItem({
          id: 'event-2',
          subject: 'Project Review',
          start: '2026-03-30T14:00:00Z',
          end: '2026-03-30T15:00:00Z',
          isOrganizer: true,
          myResponseType: 'Organizer',
          categories: '<t:String>Work</t:String><t:String>Important</t:String>'
        })}
      </m:Items>
    </m:RootFolder>
  </m:FindItemResponse>
`);

export const mockCalendarEventsResponse = MOCK_CALENDAR_RESPONSE;

export const mockCalendarEventsEmptyResponse = soapResponse(`
  <m:FindItemResponse>
    <m:ResponseCode>NoError</m:ResponseCode>
    <m:RootFolder>
      <t:TotalItemsInView>0</t:TotalItemsInView>
      <m:Items />
    </m:RootFolder>
  </m:FindItemResponse>
`);

export const mockCalendarEventsWithCancelled = soapResponse(`
  <m:FindItemResponse>
    <m:ResponseCode>NoError</m:ResponseCode>
    <m:RootFolder>
      <t:TotalItemsInView>2</t:TotalItemsInView>
      <m:Items>
        ${makeCalendarItem({
          id: 'event-1',
          subject: 'Cancelled Meeting',
          start: '2026-03-30T10:00:00Z',
          end: '2026-03-30T11:00:00Z',
          isCancelled: true,
          isOrganizer: true,
          myResponseType: 'Organizer'
        })}
        ${makeCalendarItem({
          id: 'event-2',
          subject: 'Valid Meeting',
          start: '2026-03-30T11:00:00Z',
          end: '2026-03-30T12:00:00Z',
          isOrganizer: true,
          myResponseType: 'Organizer'
        })}
      </m:Items>
    </m:RootFolder>
  </m:FindItemResponse>
`);

// ─── findtime ─────────────────────────────────────────────────────────────

export const mockGetScheduleResponse = soapResponse(`
  <m:GetScheduleResponse>
    <m:ResponseCode>NoError</m:ResponseCode>
    <m:ScheduleInfo>
      <t:ScheduleResourceEmailAddress>test@example.com</t:ScheduleResourceEmailAddress>
      <t:ScheduleItem>
        <t:Start>2026-03-30T09:00:00Z</t:Start>
        <t:End>2026-03-30T10:00:00Z</t:End>
        <t:Status>Free</t:Status>
        <t:IsPrivate>false</t:IsPrivate>
        <t:IsMeeting>true</t:IsMeeting>
        <t:IsRecurring>false</t:IsRecurring>
        <t:IsException>false</t:IsException>
        <t:MeetingTimeZone>UTC</t:MeetingTimeZone>
      </t:ScheduleItem>
      <t:ScheduleItem>
        <t:Start>2026-03-30T11:00:00Z</t:Start>
        <t:End>2026-03-30T12:00:00Z</t:End>
        <t:Status>Busy</t:Status>
        <t:IsPrivate>false</t:IsPrivate>
        <t:IsMeeting>true</t:IsMeeting>
        <t:IsRecurring>false</t:IsRecurring>
        <t:IsException>false</t:IsException>
        <t:MeetingTimeZone>UTC</t:MeetingTimeZone>
      </t:ScheduleItem>
      <t:ScheduleItem>
        <t:Start>2026-03-30T14:00:00Z</t:Start>
        <t:End>2026-03-30T15:00:00Z</t:End>
        <t:Status>Free</t:Status>
        <t:IsPrivate>false</t:IsPrivate>
        <t:IsMeeting>true</t:IsMeeting>
        <t:IsRecurring>false</t:IsRecurring>
        <t:IsException>false</t:IsException>
        <t:MeetingTimeZone>UTC</t:MeetingTimeZone>
      </t:ScheduleItem>
    </t:ScheduleInfo>
  </m:GetScheduleResponse>
`);

export const mockGetScheduleEmptyResponse = soapResponse(`
  <m:GetScheduleResponse>
    <m:ResponseCode>NoError</m:ResponseCode>
    <m:ScheduleInfo>
      <t:ScheduleResourceEmailAddress>test@example.com</t:ScheduleResourceEmailAddress>
      <t:ScheduleItem>
        <t:Start>2026-03-30T09:00:00Z</t:Start>
        <t:End>2026-03-30T17:00:00Z</t:End>
        <t:Status>Busy</t:Status>
        <t:IsPrivate>false</t:IsPrivate>
        <t:IsMeeting>true</t:IsMeeting>
        <t:IsRecurring>false</t:IsRecurring>
        <t:IsException>false</t:IsException>
        <t:MeetingTimeZone>UTC</t:MeetingTimeZone>
      </t:ScheduleItem>
    </t:ScheduleInfo>
  </m:GetScheduleResponse>
`);

// ─── respond (pending invitations) ─────────────────────────────────────────

export const mockRespondListResponse = soapResponse(`
  <m:FindItemResponse>
    <m:ResponseCode>NoError</m:ResponseCode>
    <m:RootFolder>
      <t:TotalItemsInView>2</t:TotalItemsInView>
      <m:Items>
        ${makeCalendarItem({
          id: 'invite-1',
          subject: 'Invited Meeting 1',
          start: '2026-03-31T10:00:00Z',
          end: '2026-03-31T11:00:00Z',
          location: 'Room ABC',
          isOrganizer: false,
          organizerName: 'Organizer Person',
          organizerEmail: 'organizer@example.com',
          attendees: `
            <t:RequiredAttendees>
              <t:Attendee><t:Mailbox><t:Name>Test User</t:Name><t:EmailAddress>test@example.com</t:EmailAddress></t:Mailbox><t:ResponseType>NoResponseReceived</t:ResponseType></t:Attendee>
            </t:RequiredAttendees>`
        })}
        ${makeCalendarItem({
          id: 'invite-2',
          subject: 'Invited Meeting 2',
          start: '2026-04-01T14:00:00Z',
          end: '2026-04-01T15:30:00Z',
          isOrganizer: false,
          organizerName: 'Another Organizer',
          organizerEmail: 'another@example.com'
        })}
      </m:Items>
    </m:RootFolder>
  </m:FindItemResponse>
`);

export const mockRespondSuccessResponse = soapResponse(`
  <m:RespondToItemResponse>
    <m:ResponseCode>NoError</m:ResponseCode>
  </m:RespondToItemResponse>
`);

// ─── create-event ─────────────────────────────────────────────────────────

export const mockCreateEventResponse = soapResponse(`
  <m:CreateItemResponse>
    <m:ResponseCode>NoError</m:ResponseCode>
    <m:Items>
      <t:CalendarItem>
        <t:ItemId Id="new-event-id-123" ChangeKey="newChangeKey456" />
        <t:Subject>New Meeting</t:Subject>
        <t:Start>2026-03-30T10:00:00Z</t:Start>
        <t:End>2026-03-30T11:00:00Z</t:End>
        <t:Location><t:DisplayName>Conference Room A</t:DisplayName></t:Location>
        <t:WebLink>https://outlook.office365.com/owa/item?itemId=abc</t:WebLink>
      </t:CalendarItem>
    </m:Items>
  </m:CreateItemResponse>
`);

// ─── delete-event ──────────────────────────────────────────────────────────

export const mockDeleteEventSuccessResponse = soapResponse(`
  <m:DeleteItemResponse>
    <m:ResponseCode>NoError</m:ResponseCode>
  </m:DeleteItemResponse>
`);

export const mockCancelEventSuccessResponse = soapResponse(`
  <m:CancelItemResponse>
    <m:ResponseCode>NoError</m:ResponseCode>
  </m:CancelItemResponse>
`);

// ─── find (resolveNames) ──────────────────────────────────────────────────

export const mockResolveNamesPeopleResponse = soapResponse(`
  <m:ResolveNamesResponse>
    <m:ResponseCode>NoError</m:ResponseCode>
    <m:ResolutionSet>
      <m:Resolution>
        <t:Mailbox>
          <t:Name>John Doe</t:Name>
          <t:EmailAddress>john.doe@example.com</t:EmailAddress>
          <t:RoutingType>SMTP</t:RoutingType>
          <t:MailboxType>Mailbox</t:MailboxType>
        </t:Mailbox>
        <t:Contact>
          <t:DisplayName>John Doe</t:DisplayName>
          <t:EmailAddresses><t:Entry Key="EmailAddress1">SMTP:john.doe@example.com</t:Entry></t:EmailAddresses>
          <t:JobTitle>Software Engineer</t:JobTitle>
          <t:Department>Engineering</t:Department>
        </t:Contact>
      </m:Resolution>
      <m:Resolution>
        <t:Mailbox>
          <t:Name>Jane Smith</t:Name>
          <t:EmailAddress>jane.smith@example.com</t:EmailAddress>
          <t:RoutingType>SMTP</t:RoutingType>
          <t:MailboxType>Mailbox</t:MailboxType>
        </t:Mailbox>
        <t:Contact>
          <t:DisplayName>Jane Smith</t:DisplayName>
          <t:EmailAddresses><t:Entry Key="EmailAddress1">SMTP:jane.smith@example.com</t:Entry></t:EmailAddresses>
          <t:JobTitle>Product Manager</t:JobTitle>
          <t:Department>Product</t:Department>
        </t:Contact>
      </m:Resolution>
    </m:ResolutionSet>
  </m:ResolveNamesResponse>
`);

export const mockResolveNamesRoomsResponse = soapResponse(`
  <m:ResolveNamesResponse>
    <m:ResponseCode>NoError</m:ResponseCode>
    <m:ResolutionSet>
      <m:Resolution>
        <t:Mailbox>
          <t:Name>Conference Room Alpha</t:Name>
          <t:EmailAddress>conf-alpha@example.com</t:EmailAddress>
          <t:RoutingType>SMTP</t:RoutingType>
          <t:MailboxType>Room</t:MailboxType>
        </t:Mailbox>
      </m:Resolution>
      <m:Resolution>
        <t:Mailbox>
          <t:Name>Conference Room Beta</t:Name>
          <t:EmailAddress>conf-beta@example.com</t:EmailAddress>
          <t:RoutingType>SMTP</t:RoutingType>
          <t:MailboxType>Room</t:MailboxType>
        </t:Mailbox>
      </m:Resolution>
    </m:ResolutionSet>
  </m:ResolveNamesResponse>
`);

export const mockResolveNamesEmptyResponse = soapResponse(`
  <m:ResolveNamesResponse>
    <m:ResponseCode>NoError</m:ResponseCode>
    <m:ResolutionSet TotalItemsInView="0" IncludesLastItemInRange="true" />
  </m:ResolveNamesResponse>
`);

// ─── update-event ─────────────────────────────────────────────────────────

export const mockUpdateEventResponse = soapResponse(`
  <m:UpdateItemResponse>
    <m:ResponseCode>NoError</m:ResponseCode>
    <m:Items>
      <t:CalendarItem>
        <t:ItemId Id="event-update-1" ChangeKey="updatedChangeKey789" />
        <t:Subject>Updated Meeting Title</t:Subject>
        <t:Start>2026-03-30T10:00:00Z</t:Start>
        <t:End>2026-03-30T11:00:00Z</t:End>
      </t:CalendarItem>
    </m:Items>
  </m:UpdateItemResponse>
`);

// ─── mail ─────────────────────────────────────────────────────────────────

function makeEmailItem(opts: {
  id: string;
  subject: string;
  fromName?: string;
  fromEmail?: string;
  to?: string;
  cc?: string;
  body?: string;
  receivedDateTime?: string;
  isRead?: boolean;
  hasAttachments?: boolean;
  importance?: string;
  flagStatus?: string;
}): string {
  const from = opts.fromName
    ? `<t:From><t:Mailbox><t:Name>${opts.fromName}</t:Name><t:EmailAddress>${opts.fromEmail ?? 'sender@example.com'}</t:EmailAddress></t:Mailbox></t:From>`
    : '';
  const toList = opts.to
    ? `<t:ToRecipients>${opts.to
        .split(',')
        .map((e) => `<t:Mailbox><t:EmailAddress>${e.trim()}</t:EmailAddress></t:Mailbox>`)
        .join('')}</t:ToRecipients>`
    : '';
  const ccList = opts.cc
    ? `<t:CcRecipients>${opts.cc
        .split(',')
        .map((e) => `<t:Mailbox><t:EmailAddress>${e.trim()}</t:EmailAddress></t:Mailbox>`)
        .join('')}</t:CcRecipients>`
    : '';
  const body = opts.body ? `<t:Body BodyType="Text">${opts.body}</t:Body>` : '';
  const flag = opts.flagStatus ? `<t:Flag><t:FlagStatus>${opts.flagStatus}</t:FlagStatus></t:Flag>` : '';
  return `
  <t:Message>
    <t:ItemId Id="${opts.id}" ChangeKey="emailChangeKey1" />
    ${from}
    ${toList}
    ${ccList}
    <t:Subject>${opts.subject}</t:Subject>
    ${body}
    <t:DateTimeReceived>${opts.receivedDateTime ?? '2026-03-30T09:00:00Z'}</t:DateTimeReceived>
    <t:IsRead>${(opts.isRead ?? false) ? 'true' : 'false'}</t:IsRead>
    <t:HasAttachments>${(opts.hasAttachments ?? false) ? 'true' : 'false'}</t:HasAttachments>
    <t:Importance>${opts.importance ?? 'Normal'}</t:Importance>
    ${flag}
  </t:Message>`;
}

export const mockGetEmailsResponse = soapResponse(`
  <m:FindItemResponse>
    <m:ResponseCode>NoError</m:ResponseCode>
    <m:RootFolder>
      <t:TotalItemsInView>2</t:TotalItemsInView>
      <m:Items>
        ${makeEmailItem({
          id: 'email-1',
          subject: 'Hello World',
          fromName: 'Alice',
          fromEmail: 'alice@example.com',
          to: 'test@example.com',
          body: 'This is a test email body.',
          receivedDateTime: '2026-03-30T09:00:00Z',
          isRead: false,
          importance: 'Normal'
        })}
        ${makeEmailItem({
          id: 'email-2',
          subject: 'Meeting Tomorrow',
          fromName: 'Bob',
          fromEmail: 'bob@example.com',
          to: 'test@example.com',
          body: "Don't forget our meeting tomorrow.",
          receivedDateTime: '2026-03-29T14:30:00Z',
          isRead: true,
          hasAttachments: true,
          importance: 'High',
          flagStatus: 'Flagged'
        })}
      </m:Items>
    </m:RootFolder>
  </m:FindItemResponse>
`);

export const mockGetEmailsEmptyResponse = soapResponse(`
  <m:FindItemResponse>
    <m:ResponseCode>NoError</m:ResponseCode>
    <m:RootFolder>
      <t:TotalItemsInView>0</t:TotalItemsInView>
      <m:Items />
    </m:RootFolder>
  </m:FindItemResponse>
`);

export const mockGetEmailDetailResponse = soapResponse(`
  <m:GetItemResponse>
    <m:ResponseCode>NoError</m:ResponseCode>
    <m:Items>
      <t:Message>
        <t:ItemId Id="email-detail-1" ChangeKey="emailDetailCK1" />
        <t:From><t:Mailbox><t:Name>Alice</t:Name><t:EmailAddress>alice@example.com</t:EmailAddress></t:Mailbox></t:From>
        <t:ToRecipients><t:Mailbox><t:EmailAddress>test@example.com</t:EmailAddress></t:Mailbox></t:ToRecipients>
        <t:CcRecipients><t:Mailbox><t:EmailAddress>cc@example.com</t:EmailAddress></t:Mailbox></t:CcRecipients>
        <t:Subject>Hello World</t:Subject>
        <t:Body BodyType="Text">This is the full email body.</t:Body>
        <t:DateTimeReceived>2026-03-30T09:00:00Z</t:DateTimeReceived>
        <t:IsRead>false</t:IsRead>
        <t:HasAttachments>true</t:HasAttachments>
        <t:Importance>Normal</t:Importance>
      </t:Message>
    </m:Items>
  </m:GetItemResponse>
`);

export const mockGetAttachmentsResponse = soapResponse(`
  <m:GetAttachmentsResponse>
    <m:ResponseCode>NoError</m:ResponseCode>
    <m:Attachments>
      <t:FileAttachment>
        <t:AttachmentId Id="att-1" />
        <t:Name>document.pdf</t:Name>
        <t:Size>102400</t:Size>
        <t:IsInline>false</t:IsInline>
        <t:ContentId>doc.pdf</t:ContentId>
      </t:FileAttachment>
    </m:Attachments>
  </m:GetAttachmentsResponse>
`);

export const mockUpdateEmailResponse = soapResponse(`
  <m:UpdateItemResponse>
    <m:ResponseCode>NoError</m:ResponseCode>
    <m:Items>
      <t:Message>
        <t:ItemId Id="email-update-1" ChangeKey="updatedCK" />
      </t:Message>
    </m:Items>
  </m:UpdateItemResponse>
`);

export const mockMoveEmailResponse = soapResponse(`
  <m:MoveItemResponse>
    <m:ResponseCode>NoError</m:ResponseCode>
    <m:Items>
      <t:Message>
        <t:ItemId Id="email-moved-1" ChangeKey="movedCK" />
      </t:Message>
    </m:Items>
  </m:MoveItemResponse>
`);

export const mockSendEmailResponse = soapResponse(`
  <m:SendItemResponse>
    <m:ResponseCode>NoError</m:ResponseCode>
  </m:SendItemResponse>
`);

export const mockReplyToEmailResponse = soapResponse(`
  <m:CreateItemResponse>
    <m:ResponseCode>NoError</m:ResponseCode>
    <m:Items>
      <t:Message>
        <t:ItemId Id="reply-draft-1" ChangeKey="replyCK" />
      </t:Message>
    </m:Items>
  </m:CreateItemResponse>
`);

export const mockForwardEmailResponse = soapResponse(`
  <m:ForwardItemResponse>
    <m:ResponseCode>NoError</m:ResponseCode>
  </m:ForwardItemResponse>
`);

export const mockGetMailFoldersResponse = soapResponse(`
  <m:GetFolderResponse>
    <m:ResponseCode>NoError</m:ResponseCode>
    <m:Folders>
      <t:Folder>
        <t:FolderId Id="inbox-id" />
        <t:DisplayName>Inbox</t:DisplayName>
        <t:TotalCount>10</t:TotalCount>
        <t:UnreadCount>3</t:UnreadCount>
        <t:ChildFolderCount>0</t:ChildFolderCount>
      </t:Folder>
      <t:Folder>
        <t:FolderId Id="drafts-id" />
        <t:DisplayName>Drafts</t:DisplayName>
        <t:TotalCount>2</t:TotalCount>
        <t:UnreadCount>0</t:UnreadCount>
        <t:ChildFolderCount>0</t:ChildFolderCount>
      </t:Folder>
      <t:Folder>
        <t:FolderId Id="sentitems-id" />
        <t:DisplayName>Sent Items</t:DisplayName>
        <t:TotalCount>25</t:TotalCount>
        <t:UnreadCount>0</t:UnreadCount>
        <t:ChildFolderCount>0</t:ChildFolderCount>
      </t:Folder>
      <t:Folder>
        <t:FolderId Id="deleteditems-id" />
        <t:DisplayName>Deleted Items</t:DisplayName>
        <t:TotalCount>5</t:TotalCount>
        <t:UnreadCount>0</t:UnreadCount>
        <t:ChildFolderCount>0</t:ChildFolderCount>
      </t:Folder>
      <t:Folder>
        <t:FolderId Id="archive-id" />
        <t:DisplayName>Archive</t:DisplayName>
        <t:TotalCount>8</t:TotalCount>
        <t:UnreadCount>1</t:UnreadCount>
        <t:ChildFolderCount>0</t:ChildFolderCount>
      </t:Folder>
      <t:Folder>
        <t:FolderId Id="custom-id" />
        <t:DisplayName>My Custom Folder</t:DisplayName>
        <t:TotalCount>4</t:TotalCount>
        <t:UnreadCount>0</t:UnreadCount>
        <t:ChildFolderCount>0</t:ChildFolderCount>
      </t:Folder>
    </m:Folders>
  </m:GetFolderResponse>
`);

export const mockCreateMailFolderResponse = soapResponse(`
  <m:CreateFolderResponse>
    <m:ResponseCode>NoError</m:ResponseCode>
    <m:Folders>
      <t:Folder>
        <t:FolderId Id="new-folder-id" />
        <t:DisplayName>New Folder</t:DisplayName>
        <t:TotalCount>0</t:TotalCount>
        <t:UnreadCount>0</t:UnreadCount>
        <t:ChildFolderCount>0</t:ChildFolderCount>
      </t:Folder>
    </m:Folders>
  </m:CreateFolderResponse>
`);

export const mockUpdateMailFolderResponse = soapResponse(`
  <m:UpdateFolderResponse>
    <m:ResponseCode>NoError</m:ResponseCode>
    <m:Folders>
      <t:Folder>
        <t:FolderId Id="folder-updated-id" />
        <t:DisplayName>Renamed Folder</t:DisplayName>
      </t:Folder>
    </m:Folders>
  </m:UpdateFolderResponse>
`);

export const mockDeleteMailFolderResponse = soapResponse(`
  <m:DeleteFolderResponse>
    <m:ResponseCode>NoError</m:ResponseCode>
  </m:DeleteFolderResponse>
`);

// ─── drafts ────────────────────────────────────────────────────────────────

export const mockGetDraftsResponse = soapResponse(`
  <m:FindItemResponse>
    <m:ResponseCode>NoError</m:ResponseCode>
    <m:RootFolder>
      <t:TotalItemsInView>2</t:TotalItemsInView>
      <m:Items>
        ${makeEmailItem({
          id: 'draft-1',
          subject: 'Draft Email One',
          to: 'recipient@example.com',
          body: 'Draft body content.',
          receivedDateTime: '2026-03-30T10:00:00Z'
        })}
        ${makeEmailItem({
          id: 'draft-2',
          subject: 'Draft Email Two',
          to: 'another@example.com',
          body: 'Another draft.',
          receivedDateTime: '2026-03-30T11:00:00Z'
        })}
      </m:Items>
    </m:RootFolder>
  </m:FindItemResponse>
`);

export const mockCreateDraftResponse = soapResponse(`
  <m:CreateItemResponse>
    <m:ResponseCode>NoError</m:ResponseCode>
    <m:Items>
      <t:Message>
        <t:ItemId Id="new-draft-id-123" ChangeKey="draftCK123" />
        <t:Subject>New Draft</t:Subject>
      </t:Message>
    </m:Items>
  </m:CreateItemResponse>
`);

export const mockUpdateDraftResponse = soapResponse(`
  <m:UpdateItemResponse>
    <m:ResponseCode>NoError</m:ResponseCode>
    <m:Items>
      <t:Message>
        <t:ItemId Id="draft-edit-id" ChangeKey="draftEditCK" />
      </t:Message>
    </m:Items>
  </m:UpdateItemResponse>
`);

export const mockSendDraftResponse = soapResponse(`
  <m:SendItemResponse>
    <m:ResponseCode>NoError</m:ResponseCode>
  </m:SendItemResponse>
`);

export const mockDeleteDraftResponse = soapResponse(`
  <m:DeleteItemResponse>
    <m:ResponseCode>NoError</m:ResponseCode>
  </m:DeleteItemResponse>
`);

export const mockAddAttachmentResponse = soapResponse(`
  <m:CreateAttachmentResponse>
    <m:ResponseCode>NoError</m:ResponseCode>
    <m:Attachments>
      <t:FileAttachment>
        <t:AttachmentId Id="new-att-id" />
        <t:Name>attachment.pdf</t:Name>
      </t:FileAttachment>
    </m:Attachments>
    <t:RootItemId Id="new-draft-id-123" ChangeKey="afterAttachCK" />
  </m:CreateAttachmentResponse>
`);

// ─── rooms ─────────────────────────────────────────────────────────────────

export const mockGetRoomsResponse = soapResponse(`
  <m:GetRoomListsResponse>
    <m:ResponseCode>NoError</m:ResponseCode>
    <m:RoomLists>
      <t:Address>
        <t:EmailAddress>rooms@example.com</t:EmailAddress>
      </t:Address>
    </m:RoomLists>
  </m:GetRoomListsResponse>
`);

export const mockGetRoomsFromListResponse = soapResponse(`
  <m:GetRoomsResponse>
    <m:ResponseCode>NoError</m:ResponseCode>
    <m:Rooms>
      <t:Room>
        <t:Id>room-1</t:Id>
        <t:Name>Conference Room Alpha</t:Name>
        <t:EmailAddress>conf-alpha@example.com</t:EmailAddress>
      </t:Room>
      <t:Room>
        <t:Id>room-2</t:Id>
        <t:Name>Conference Room Beta</t:Name>
        <t:EmailAddress>conf-beta@example.com</t:EmailAddress>
      </t:Room>
    </m:Rooms>
  </m:GetRoomsResponse>
`);

export const mockSearchRoomsResponse = soapResponse(`
  <m:ExpandDLResponse>
    <m:ResponseCode>NoError</m:ResponseCode>
    <m:DLbl>Rooms</m:DLbl>
    <m:AddressList>
      <t:Mailbox>
        <t:Name>Conference Room Alpha</t:Name>
        <t:EmailAddress>conf-alpha@example.com</t:EmailAddress>
      </t:Mailbox>
    </m:AddressList>
  </m:ExpandDLResponse>
`);

export const mockIsRoomFreeResponse = soapResponse(`
  <m:GetRoomListsResponse>
    <m:ResponseCode>NoError</m:ResponseCode>
  </m:GetRoomListsResponse>
`);

// ─── Graph API (OneDrive/files) ─────────────────────────────────────────────

export const mockGraphListFilesResponse = {
  value: [
    {
      id: 'drive-item-1',
      name: 'Document.pdf',
      size: 102400,
      lastModifiedDateTime: '2026-03-29T12:00:00Z',
      webUrl: 'https://example.sharepoint.com/doc.pdf',
      file: { mimeType: 'application/pdf' },
      folder: undefined
    },
    {
      id: 'drive-item-2',
      name: 'My Folder',
      size: undefined,
      lastModifiedDateTime: '2026-03-28T09:00:00Z',
      webUrl: 'https://example.sharepoint.com/folder',
      file: undefined,
      folder: { childCount: 5 }
    }
  ]
};

export const mockGraphSearchFilesResponse = {
  value: [
    {
      id: 'drive-item-3',
      name: 'Report.xlsx',
      size: 51200,
      lastModifiedDateTime: '2026-03-27T15:00:00Z',
      webUrl: 'https://example.sharepoint.com/report.xlsx',
      file: { mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' },
      folder: undefined
    }
  ]
};

export const mockGraphGetFileMetadataResponse = {
  id: 'drive-item-1',
  name: 'Report.docx',
  size: 102400,
  createdDateTime: '2026-03-20T10:00:00Z',
  lastModifiedDateTime: '2026-03-29T12:00:00Z',
  webUrl: 'https://example.sharepoint.com/report.docx',
  file: { mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' },
  folder: undefined,
  parentReference: { driveId: 'drive-1', id: 'root', path: '/drive/root' }
};

export const mockGraphUploadResponse = {
  id: 'new-drive-item-id',
  name: 'Uploaded.txt',
  size: 1024,
  lastModifiedDateTime: '2026-03-30T12:00:00Z',
  webUrl: 'https://example.sharepoint.com/uploaded.txt',
  file: { mimeType: 'text/plain' },
  folder: undefined
};

export const mockGraphDeleteResponse = {};
export const mockGraphShareResponse = { id: 'share-link-id', webUrl: 'https://example.sharepoint.com/share/abc123' };
export const mockGraphCollabResponse = {
  item: { id: 'drive-item-1', name: 'Report.docx' },
  link: { webUrl: 'https://example.sharepoint.com/collab' },
  collaborationUrl: 'https://office.com/collab?ItemID=drive-item-1',
  lockAcquired: false
};
export const mockGraphCheckinResponse = {
  item: { id: 'checkin-item-id', name: 'CheckedIn.docx' },
  checkedIn: true,
  comment: 'Checked in'
};
export const mockGraphCreateUploadSessionResponse = {
  uploadUrl: 'https://upload.microsoft.com/upload-session/abc',
  expirationDateTime: '2026-03-31T12:00:00Z'
};
