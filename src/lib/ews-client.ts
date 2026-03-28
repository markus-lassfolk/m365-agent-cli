// ─── Timezone Utilities ───────────────────────────────────────────────────────

/** Returns the IANA timezone name for the current process, e.g. 'Europe/Stockholm'. */
export function getLocalTimezone(): string {
  return Intl.DateTimeFormat().resolvedOptions().timeZone;
}

// ─── XML Utilities ───

export function xmlEscape(value: string): string {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function xmlDecode(value: string): string {
  return String(value || '')
    .replace(/<!\[CDATA\[([\s\S]*?)\]\]>/g, '$1')
    .replace(/&#x([0-9a-f]+);/gi, (_, hex) => {
      const cp = parseInt(hex, 16);
      return Number.isFinite(cp) ? String.fromCodePoint(cp) : _;
    })
    .replace(/&#([0-9]+);/g, (_, digits) => {
      const cp = parseInt(digits, 10);
      return Number.isFinite(cp) ? String.fromCodePoint(cp) : _;
    })
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
    .replace(/\r/g, '');
}

export function extractTag(xml: string, tagName: string): string {
  const regex = new RegExp(
    `<(?:[A-Za-z0-9_]+:)?${tagName}\\b[^>]*>([\\s\\S]*?)<\\/(?:[A-Za-z0-9_]+:)?${tagName}>`,
    'i'
  );
  const match = xml.match(regex);
  return match ? xmlDecode(match[1]) : '';
}

function _extractTagRaw(xml: string, tagName: string): string {
  const regex = new RegExp(
    `<(?:[A-Za-z0-9_]+:)?${tagName}\\b[^>]*>([\\s\\S]*?)<\\/(?:[A-Za-z0-9_]+:)?${tagName}>`,
    'i'
  );
  const match = xml.match(regex);
  return match ? match[1] : '';
}

export function extractAttribute(xml: string, tagName: string, attrName: string): string {
  const regex = new RegExp(`<(?:[A-Za-z0-9_]+:)?${tagName}\\b[^>]*\\b${attrName}="([^"]*)"`, 'i');
  const match = xml.match(regex);
  return match ? xmlDecode(match[1]) : '';
}

export function extractBlocks(xml: string, tagName: string): string[] {
  const regex = new RegExp(`<(?:[A-Za-z0-9_]+:)?${tagName}\\b[\\s\\S]*?<\\/(?:[A-Za-z0-9_]+:)?${tagName}>`, 'g');
  return [...xml.matchAll(regex)].map((m) => m[0]);
}

export function extractSelfClosingOrBlock(xml: string, tagName: string): string {
  // Matches both <Tag ... /> and <Tag ...>...</Tag>
  const regex = new RegExp(
    `<(?:[A-Za-z0-9_]+:)?${tagName}\\b[^>]*(?:\\/>|>[\\s\\S]*?<\\/(?:[A-Za-z0-9_]+:)?${tagName}>)`,
    'i'
  );
  const match = xml.match(regex);
  return match ? match[0] : '';
}

// ─── SOAP Core ───

import { validateUrl } from './url-validation';

export const EWS_ENDPOINT = validateUrl(
  process.env.EWS_ENDPOINT || 'https://outlook.office365.com/EWS/Exchange.asmx',
  'EWS_ENDPOINT'
);
export const EWS_USERNAME = process.env.EWS_USERNAME || '';

export function soapEnvelope(body: string): string {
  return `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2016" />
  </soap:Header>
  <soap:Body>
    ${body}
  </soap:Body>
</soap:Envelope>`;
}

export async function callEws(token: string, envelope: string, mailbox?: string): Promise<string> {
  const anchorMailbox = mailbox || EWS_USERNAME;
  const response = await fetch(EWS_ENDPOINT, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'text/xml; charset=utf-8',
      Accept: 'text/xml',
      'X-AnchorMailbox': anchorMailbox
    },
    body: envelope
  });

  const xml = await response.text();

  if (!response.ok) {
    const soapError = extractTag(xml, 'faultstring') || extractTag(xml, 'MessageText');
    throw new Error(`EWS HTTP ${response.status}${soapError ? `: ${soapError}` : ''}`);
  }

  const responseCode = extractTag(xml, 'ResponseCode');
  if (responseCode && responseCode !== 'NoError') {
    const messageText = extractTag(xml, 'MessageText');
    throw new Error(`EWS ${responseCode}${messageText ? `: ${messageText}` : ''}`);
  }

  return xml;
}

// ─── Types ───

export interface OwaError {
  code: string;
  message: string;
}

export interface OwaResponse<T = unknown> {
  ok: boolean;
  status: number;
  data?: T;
  error?: OwaError;
  /** Informational message (e.g., fallback used, partial success) */
  info?: string;
}

export interface OwaUserInfo {
  displayName: string;
  email: string;
}

export interface CalendarAttendee {
  Type: 'Required' | 'Optional' | 'Resource';
  Status: {
    Response: 'None' | 'Organizer' | 'TentativelyAccepted' | 'Accepted' | 'Declined' | 'NotResponded';
    Time: string;
  };
  EmailAddress: {
    Name: string;
    Address: string;
  };
}

export interface CalendarEvent {
  Id: string;
  ChangeKey?: string;
  Subject: string;
  Start: { DateTime: string; TimeZone: string };
  End: { DateTime: string; TimeZone: string };
  Location?: { DisplayName?: string };
  Organizer?: { EmailAddress?: { Name?: string; Address?: string } };
  Attendees?: CalendarAttendee[];
  IsAllDay?: boolean;
  IsCancelled?: boolean;
  IsOrganizer?: boolean;
  BodyPreview?: string;
  Categories?: string[];
  ShowAs?: string;
  Importance?: string;
  IsOnlineMeeting?: boolean;
  OnlineMeetingUrl?: string;
  WebLink?: string;
}

export interface RecurrencePattern {
  Type: 'Daily' | 'Weekly' | 'AbsoluteMonthly' | 'RelativeMonthly' | 'AbsoluteYearly' | 'RelativeYearly';
  Interval: number;
  DaysOfWeek?: string[];
  DayOfMonth?: number;
  Month?: number;
  Index?: 'First' | 'Second' | 'Third' | 'Fourth' | 'Last';
}

export interface RecurrenceRange {
  Type: 'EndDate' | 'NoEnd' | 'Numbered';
  StartDate: string;
  EndDate?: string;
  NumberOfOccurrences?: number;
}

export interface Recurrence {
  Pattern: RecurrencePattern;
  Range: RecurrenceRange;
}

export interface CreateEventOptions {
  token: string;
  subject: string;
  start: string;
  end: string;
  body?: string;
  location?: string;
  attendees?: Array<{ email: string; name?: string; type?: 'Required' | 'Optional' | 'Resource' }>;
  isOnlineMeeting?: boolean;
  recurrence?: Recurrence;
  mailbox?: string;
}

export interface CreatedEvent {
  Id: string;
  Subject: string;
  Start: { DateTime: string; TimeZone: string };
  End: { DateTime: string; TimeZone: string };
  WebLink?: string;
  OnlineMeetingUrl?: string;
}

export interface UpdateEventOptions {
  token: string;
  eventId: string;
  changeKey?: string;
  subject?: string;
  start?: string;
  end?: string;
  body?: string;
  location?: string;
  attendees?: Array<{ email: string; name?: string; type?: 'Required' | 'Optional' | 'Resource' }>;
  isOnlineMeeting?: boolean;
  mailbox?: string;
}

export interface ScheduleInfo {
  scheduleId: string;
  availabilityView: string;
  scheduleItems: Array<{
    status: string;
    start: { dateTime: string; timeZone: string };
    end: { dateTime: string; timeZone: string };
    subject?: string;
    location?: string;
  }>;
}

export interface FreeBusySlot {
  status: 'Free' | 'Busy' | 'Tentative';
  start: string;
  end: string;
  subject?: string;
}

export interface Room {
  Address: string;
  Name: string;
}

export interface RoomList {
  Address: string;
  Name: string;
}

export interface EmailAddress {
  Name?: string;
  Address?: string;
}

export interface EmailMessage {
  Id: string;
  ChangeKey?: string;
  Subject?: string;
  BodyPreview?: string;
  Body?: { ContentType: string; Content: string };
  From?: { EmailAddress?: EmailAddress };
  ToRecipients?: Array<{ EmailAddress?: EmailAddress }>;
  CcRecipients?: Array<{ EmailAddress?: EmailAddress }>;
  ReceivedDateTime?: string;
  SentDateTime?: string;
  IsRead?: boolean;
  IsDraft?: boolean;
  HasAttachments?: boolean;
  Importance?: 'Low' | 'Normal' | 'High';
  Flag?: { FlagStatus?: 'NotFlagged' | 'Flagged' | 'Complete' };
}

export interface EmailListResponse {
  value: EmailMessage[];
}

export interface GetEmailsOptions {
  token: string;
  folder?: string;
  top?: number;
  skip?: number;
  filter?: string;
  search?: string;
  select?: string[];
  orderBy?: string;
}

export interface Attachment {
  Id: string;
  Name: string;
  ContentType: string;
  Size: number;
  IsInline: boolean;
  ContentId?: string;
  ContentBytes?: string;
}

export interface AttachmentListResponse {
  value: Attachment[];
}

export interface EmailAttachment {
  name: string;
  contentType: string;
  contentBytes: string;
}

export interface MailFolder {
  Id: string;
  DisplayName: string;
  ParentFolderId?: string;
  ChildFolderCount: number;
  UnreadItemCount: number;
  TotalItemCount: number;
}

export interface MailFolderListResponse {
  value: MailFolder[];
}

export type ResponseType = 'accept' | 'decline' | 'tentative';

export interface RespondToEventOptions {
  token: string;
  eventId: string;
  response: ResponseType;
  comment?: string;
  sendResponse?: boolean;
  mailbox?: string;
}

// ─── Parsing Helpers ───

function parseCalendarItem(block: string, mailbox?: string): CalendarEvent {
  const id = extractAttribute(block, 'ItemId', 'Id');
  const changeKey = extractAttribute(block, 'ItemId', 'ChangeKey');
  const subject = extractTag(block, 'Subject');
  const start = extractTag(block, 'Start');
  const end = extractTag(block, 'End');
  const location = extractTag(block, 'Location');
  const isAllDay = extractTag(block, 'IsAllDayEvent').toLowerCase() === 'true';
  const isCancelled = extractTag(block, 'IsCancelled').toLowerCase() === 'true';
  const bodyPreview = extractTag(block, 'TextBody') || extractTag(block, 'Body');
  const importance = extractTag(block, 'Importance') || 'Normal';
  const showAs = extractTag(block, 'LegacyFreeBusyStatus') || 'Busy';

  // Organizer
  const organizerBlock = extractSelfClosingOrBlock(block, 'Organizer');
  const organizerName = extractTag(organizerBlock, 'Name');
  const organizerEmail = extractTag(organizerBlock, 'EmailAddress');
  const myResponseType = extractTag(block, 'MyResponseType');
  const effectiveUser = mailbox || EWS_USERNAME;
  const isOrganizer = myResponseType === 'Organizer' || organizerEmail.toLowerCase() === effectiveUser.toLowerCase();

  // Attendees
  const attendees: CalendarAttendee[] = [];

  for (const type of ['RequiredAttendees', 'OptionalAttendees', 'Resources'] as const) {
    const typeBlock = extractSelfClosingOrBlock(block, type);
    const attendeeBlocks = extractBlocks(typeBlock, 'Attendee');
    const attendeeType =
      type === 'RequiredAttendees' ? 'Required' : type === 'OptionalAttendees' ? 'Optional' : 'Resource';

    for (const ab of attendeeBlocks) {
      const mailboxBlock = extractSelfClosingOrBlock(ab, 'Mailbox');
      const name = extractTag(mailboxBlock, 'Name');
      const email = extractTag(mailboxBlock, 'EmailAddress');
      const responseType = extractTag(ab, 'ResponseType') || 'Unknown';
      const lastResponseTime = extractTag(ab, 'LastResponseTime') || '';

      // Map EWS ResponseType to our format
      const responseMap: Record<string, CalendarAttendee['Status']['Response']> = {
        Accept: 'Accepted',
        Decline: 'Declined',
        Tentative: 'TentativelyAccepted',
        NoResponseReceived: 'NotResponded',
        Organizer: 'Organizer',
        Unknown: 'None'
      };

      attendees.push({
        Type: attendeeType,
        Status: {
          Response: responseMap[responseType] || 'None',
          Time: lastResponseTime
        },
        EmailAddress: { Name: name, Address: email }
      });
    }
  }

  // Categories
  const categoriesBlock = extractSelfClosingOrBlock(block, 'Categories');
  const categories = extractBlocks(categoriesBlock, 'String').map(
    (b) => extractTag(b, 'String') || xmlDecode(b.replace(/<[^>]+>/g, ''))
  );

  return {
    Id: id,
    ChangeKey: changeKey,
    Subject: subject,
    Start: { DateTime: start, TimeZone: 'UTC' },
    End: { DateTime: end, TimeZone: 'UTC' },
    Location: location ? { DisplayName: location } : undefined,
    Organizer: { EmailAddress: { Name: organizerName, Address: organizerEmail } },
    Attendees: attendees.length > 0 ? attendees : undefined,
    IsAllDay: isAllDay,
    IsCancelled: isCancelled,
    IsOrganizer: isOrganizer,
    BodyPreview: bodyPreview ? bodyPreview.substring(0, 200).replace(/\s+/g, ' ').trim() : undefined,
    Categories: categories.length > 0 ? categories : undefined,
    ShowAs: showAs,
    Importance: importance
  };
}

function parseEmailMessage(block: string): EmailMessage {
  const id = extractAttribute(block, 'ItemId', 'Id');
  const changeKey = extractAttribute(block, 'ItemId', 'ChangeKey');
  const subject = extractTag(block, 'Subject');
  const bodyContent = extractTag(block, 'Body') || extractTag(block, 'TextBody');
  const bodyType = extractAttribute(block, 'Body', 'BodyType') || 'Text';
  const preview =
    extractTag(block, 'Preview') || (bodyContent ? bodyContent.substring(0, 200).replace(/\s+/g, ' ').trim() : '');
  const receivedDateTime = extractTag(block, 'DateTimeReceived');
  const sentDateTime = extractTag(block, 'DateTimeSent');
  const isRead = extractTag(block, 'IsRead').toLowerCase() === 'true';
  const isDraft = extractTag(block, 'IsDraft').toLowerCase() === 'true';
  const hasAttachments = extractTag(block, 'HasAttachments').toLowerCase() === 'true';
  const importance = (extractTag(block, 'Importance') || 'Normal') as 'Low' | 'Normal' | 'High';

  // From
  const fromBlock = extractSelfClosingOrBlock(block, 'From');
  const fromMailbox = extractSelfClosingOrBlock(fromBlock, 'Mailbox');
  const fromName = extractTag(fromMailbox, 'Name');
  const fromEmail = extractTag(fromMailbox, 'EmailAddress');

  // To
  const toBlock = extractSelfClosingOrBlock(block, 'ToRecipients');
  const toMailboxes = extractBlocks(toBlock, 'Mailbox');
  const toRecipients = toMailboxes.map((mb) => ({
    EmailAddress: {
      Name: extractTag(mb, 'Name'),
      Address: extractTag(mb, 'EmailAddress')
    }
  }));

  // Cc
  const ccBlock = extractSelfClosingOrBlock(block, 'CcRecipients');
  const ccMailboxes = extractBlocks(ccBlock, 'Mailbox');
  const ccRecipients = ccMailboxes.map((mb) => ({
    EmailAddress: {
      Name: extractTag(mb, 'Name'),
      Address: extractTag(mb, 'EmailAddress')
    }
  }));

  // Flag
  const flagBlock = extractSelfClosingOrBlock(block, 'Flag');
  const flagStatus = extractTag(flagBlock, 'FlagStatus') as 'NotFlagged' | 'Flagged' | 'Complete' | undefined;

  return {
    Id: id,
    ChangeKey: changeKey,
    Subject: subject || undefined,
    BodyPreview: preview || undefined,
    Body: bodyContent ? { ContentType: bodyType, Content: bodyContent } : undefined,
    From: fromEmail ? { EmailAddress: { Name: fromName, Address: fromEmail } } : undefined,
    ToRecipients: toRecipients.length > 0 ? toRecipients : undefined,
    CcRecipients: ccRecipients.length > 0 ? ccRecipients : undefined,
    ReceivedDateTime: receivedDateTime || undefined,
    SentDateTime: sentDateTime || undefined,
    IsRead: isRead,
    IsDraft: isDraft,
    HasAttachments: hasAttachments,
    Importance: importance,
    Flag: flagStatus ? { FlagStatus: flagStatus } : undefined
  };
}

function parseFolder(block: string): MailFolder {
  return {
    Id: extractAttribute(block, 'FolderId', 'Id'),
    DisplayName: extractTag(block, 'DisplayName'),
    ParentFolderId: extractAttribute(block, 'ParentFolderId', 'Id') || undefined,
    ChildFolderCount: parseInt(extractTag(block, 'ChildFolderCount') || '0', 10),
    UnreadItemCount: parseInt(extractTag(block, 'UnreadItemCount') || '0', 10),
    TotalItemCount: parseInt(extractTag(block, 'TotalItemCount') || '0', 10)
  };
}

export function ewsResult<T>(data: T): OwaResponse<T> {
  return { ok: true, status: 200, data };
}

export function ewsError(err: unknown): OwaResponse<never> {
  const message = err instanceof Error ? err.message : 'Unknown error';
  return { ok: false, status: 0, error: { code: 'EWS_ERROR', message } };
}

// Map well-known folder names to EWS DistinguishedFolderId
const FOLDER_MAP: Record<string, string> = {
  inbox: 'inbox',
  drafts: 'drafts',
  sentitems: 'sentitems',
  sent: 'sentitems',
  deleteditems: 'deleteditems',
  deleted: 'deleteditems',
  trash: 'deleteditems',
  junkemail: 'junkemail',
  junk: 'junkemail',
  spam: 'junkemail',
  outbox: 'outbox',
  archive: 'archivemsgfolderoot'
};

function folderIdXml(folder: string): string {
  const distinguished = FOLDER_MAP[folder.toLowerCase()];
  if (distinguished) {
    return `<t:DistinguishedFolderId Id="${distinguished}" />`;
  }
  return `<t:FolderId Id="${xmlEscape(folder)}" />`;
}

// ─── Session Validation ───

export async function validateSession(token: string): Promise<boolean> {
  try {
    const envelope = soapEnvelope(`
    <m:GetFolder>
      <m:FolderShape><t:BaseShape>IdOnly</t:BaseShape></m:FolderShape>
      <m:FolderIds>
        <t:DistinguishedFolderId Id="inbox" />
      </m:FolderIds>
    </m:GetFolder>`);
    await callEws(token, envelope);
    return true;
  } catch {
    return false;
  }
}

// ─── User Info ───

export async function getOwaUserInfo(token: string): Promise<OwaResponse<OwaUserInfo>> {
  try {
    const envelope = soapEnvelope(`
    <m:ResolveNames ReturnFullContactData="true" SearchScope="ActiveDirectory">
      <m:UnresolvedEntry>${xmlEscape(EWS_USERNAME)}</m:UnresolvedEntry>
    </m:ResolveNames>`);
    const xml = await callEws(token, envelope);

    const resolution = extractBlocks(xml, 'Resolution')[0] || '';
    const mailbox = extractSelfClosingOrBlock(resolution, 'Mailbox');
    const name = extractTag(mailbox, 'Name') || EWS_USERNAME;
    const email = extractTag(mailbox, 'EmailAddress') || EWS_USERNAME;

    return ewsResult({ displayName: name, email });
  } catch (err) {
    return ewsError(
      new Error(`Failed to resolve OWA user info: ${err instanceof Error ? err.message : 'Unknown error'}`)
    );
  }
}

// ─── Calendar Operations ───

export async function getCalendarEvents(
  token: string,
  startDateTime: string,
  endDateTime: string,
  mailbox?: string
): Promise<OwaResponse<CalendarEvent[]>> {
  try {
    const calendarFolderXml = mailbox
      ? `<t:DistinguishedFolderId Id="calendar"><t:Mailbox><t:EmailAddress>${xmlEscape(mailbox)}</t:EmailAddress></t:Mailbox></t:DistinguishedFolderId>`
      : `<t:DistinguishedFolderId Id="calendar" />`;

    const envelope = soapEnvelope(`
    <m:FindItem Traversal="Shallow">
      <m:ItemShape>
        <t:BaseShape>Default</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="calendar:Location" />
          <t:FieldURI FieldURI="calendar:Organizer" />
          <t:FieldURI FieldURI="calendar:RequiredAttendees" />
          <t:FieldURI FieldURI="calendar:OptionalAttendees" />
          <t:FieldURI FieldURI="calendar:Resources" />
          <t:FieldURI FieldURI="item:Categories" />
          <t:FieldURI FieldURI="calendar:IsAllDayEvent" />
          <t:FieldURI FieldURI="calendar:IsCancelled" />
          <t:FieldURI FieldURI="calendar:MyResponseType" />
          <t:FieldURI FieldURI="calendar:LegacyFreeBusyStatus" />
          <t:FieldURI FieldURI="item:Importance" />
          <t:FieldURI FieldURI="item:TextBody" />
        </t:AdditionalProperties>
      </m:ItemShape>
      <m:CalendarView StartDate="${xmlEscape(startDateTime)}" EndDate="${xmlEscape(endDateTime)}" />
      <m:ParentFolderIds>
        ${calendarFolderXml}
      </m:ParentFolderIds>
    </m:FindItem>`);

    const xml = await callEws(token, envelope, mailbox);
    const blocks = extractBlocks(xml, 'CalendarItem');
    const events = blocks.map((block) => parseCalendarItem(block, mailbox));

    return ewsResult(events);
  } catch (err) {
    return ewsError(err);
  }
}

export async function getCalendarEvent(
  token: string,
  eventId: string,
  mailbox?: string
): Promise<OwaResponse<CalendarEvent>> {
  try {
    const envelope = soapEnvelope(`
    <m:GetItem>
      <m:ItemShape>
        <t:BaseShape>Default</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="calendar:Location" />
          <t:FieldURI FieldURI="calendar:Organizer" />
          <t:FieldURI FieldURI="calendar:RequiredAttendees" />
          <t:FieldURI FieldURI="calendar:OptionalAttendees" />
          <t:FieldURI FieldURI="calendar:Resources" />
          <t:FieldURI FieldURI="item:Categories" />
          <t:FieldURI FieldURI="calendar:IsAllDayEvent" />
          <t:FieldURI FieldURI="calendar:IsCancelled" />
          <t:FieldURI FieldURI="calendar:MyResponseType" />
          <t:FieldURI FieldURI="calendar:LegacyFreeBusyStatus" />
          <t:FieldURI FieldURI="item:Importance" />
          <t:FieldURI FieldURI="item:TextBody" />
        </t:AdditionalProperties>
      </m:ItemShape>
      <m:ItemIds>
        <t:ItemId Id="${xmlEscape(eventId)}" />
      </m:ItemIds>
    </m:GetItem>`);

    const xml = await callEws(token, envelope, mailbox);
    const block = extractBlocks(xml, 'CalendarItem')[0];
    if (!block) return { ok: false, status: 404, error: { code: 'NOT_FOUND', message: 'Event not found' } };

    return ewsResult(parseCalendarItem(block, mailbox));
  } catch (err) {
    return ewsError(err);
  }
}

/**
 * Validates required fields on a Recurrence input.
 * Throws a descriptive Error for any missing required field.
 */
function validateRecurrenceInput(recurrence: Recurrence): void {
  if (!recurrence) {
    throw new Error('[Recurrence] recurrence object is required');
  }
  if (!recurrence.Pattern) {
    throw new Error('[Recurrence] recurrence.Pattern is required');
  }
  if (!recurrence.Range) {
    throw new Error('[Recurrence] recurrence.Range is required');
  }

  const { Pattern: p, Range: r } = recurrence;

  if (!r.StartDate || r.StartDate.trim() === '') {
    throw new Error('[Recurrence] recurrence.Range.StartDate is required');
  }

  if (r.Type === 'EndDate' && (!r.EndDate || r.EndDate.trim() === '')) {
    throw new Error('[Recurrence] recurrence.Range.EndDate is required when Range.Type is "EndDate"');
  }

  if (p.Interval === undefined || p.Interval <= 0) {
    throw new Error('[Recurrence] recurrence.Pattern.Interval must be a positive integer');
  }
}

function buildRecurrenceXml(recurrence: Recurrence): string {
  validateRecurrenceInput(recurrence);

  let patternXml = '';
  const p = recurrence.Pattern;
  const validTypes = [
    'Daily',
    'Weekly',
    'AbsoluteMonthly',
    'RelativeMonthly',
    'AbsoluteYearly',
    'RelativeYearly'
  ] as const;

  switch (p.Type) {
    case 'Daily':
      patternXml = `<t:DailyRecurrence><t:Interval>${p.Interval}</t:Interval></t:DailyRecurrence>`;
      break;
    case 'Weekly': {
      const days = (p.DaysOfWeek || []).map((d) => `<t:DayOfWeek>${xmlEscape(d)}</t:DayOfWeek>`).join('');
      patternXml = `<t:WeeklyRecurrence><t:Interval>${p.Interval}</t:Interval><t:DaysOfWeek>${days || ''}</t:DaysOfWeek></t:WeeklyRecurrence>`;
      break;
    }
    case 'AbsoluteMonthly':
      patternXml = `<t:AbsoluteMonthlyRecurrence><t:Interval>${p.Interval}</t:Interval><t:DayOfMonth>${p.DayOfMonth || 1}</t:DayOfMonth></t:AbsoluteMonthlyRecurrence>`;
      break;
    case 'AbsoluteYearly':
      patternXml = `<t:AbsoluteYearlyRecurrence><t:DayOfMonth>${p.DayOfMonth || 1}</t:DayOfMonth><t:Month>${['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'][(p.Month || 1) - 1]}</t:Month></t:AbsoluteYearlyRecurrence>`;
      break;
    default:
      if (!validTypes.includes(p.Type)) {
        console.warn(`[Recurrence] Unknown Pattern.Type "${p.Type}", defaulting to Daily`);
      }
      patternXml = `<t:DailyRecurrence><t:Interval>${p.Interval}</t:Interval></t:DailyRecurrence>`;
  }

  let rangeXml = '';
  const r = recurrence.Range;
  switch (r.Type) {
    case 'EndDate': {
      const endDate = r.EndDate || r.StartDate;
      if (!r.EndDate) {
        console.warn(
          '[Recurrence] Range.EndDate is missing; EndDate will equal StartDate (effectively a single-occurrence event)'
        );
      }
      rangeXml = `<t:EndDateRecurrence><t:StartDate>${xmlEscape(r.StartDate)}</t:StartDate><t:EndDate>${xmlEscape(endDate)}</t:EndDate></t:EndDateRecurrence>`;
      break;
    }
    case 'Numbered':
      if (r.NumberOfOccurrences === undefined || r.NumberOfOccurrences <= 0) {
        console.warn('[Recurrence] Range.NumberOfOccurrences is missing or invalid, defaulting to 10');
      }
      rangeXml = `<t:NumberedRecurrence><t:StartDate>${xmlEscape(r.StartDate)}</t:StartDate><t:NumberOfOccurrences>${r.NumberOfOccurrences || 10}</t:NumberOfOccurrences></t:NumberedRecurrence>`;
      break;
    default:
      rangeXml = `<t:NoEndRecurrence><t:StartDate>${xmlEscape(r.StartDate)}</t:StartDate></t:NoEndRecurrence>`;
  }

  return `<t:Recurrence>${patternXml}${rangeXml}</t:Recurrence>`;
}

export async function createEvent(options: CreateEventOptions): Promise<OwaResponse<CreatedEvent>> {
  try {
    const { token, subject, start, end, body, location, attendees, isOnlineMeeting, recurrence, mailbox } = options;

    let attendeesXml = '';
    if (attendees && attendees.length > 0) {
      const required = attendees.filter((a) => (a.type || 'Required') === 'Required');
      const optional = attendees.filter((a) => a.type === 'Optional');
      const resources = attendees.filter((a) => a.type === 'Resource');

      if (required.length > 0) {
        attendeesXml += `<t:RequiredAttendees>${required
          .map(
            (a) =>
              `<t:Attendee><t:Mailbox><t:EmailAddress>${xmlEscape(a.email)}</t:EmailAddress>${a.name ? `<t:Name>${xmlEscape(a.name)}</t:Name>` : ''}</t:Mailbox></t:Attendee>`
          )
          .join('')}</t:RequiredAttendees>`;
      }
      if (optional.length > 0) {
        attendeesXml += `<t:OptionalAttendees>${optional
          .map(
            (a) =>
              `<t:Attendee><t:Mailbox><t:EmailAddress>${xmlEscape(a.email)}</t:EmailAddress>${a.name ? `<t:Name>${xmlEscape(a.name)}</t:Name>` : ''}</t:Mailbox></t:Attendee>`
          )
          .join('')}</t:OptionalAttendees>`;
      }
      if (resources.length > 0) {
        attendeesXml += `<t:Resources>${resources
          .map(
            (a) =>
              `<t:Attendee><t:Mailbox><t:EmailAddress>${xmlEscape(a.email)}</t:EmailAddress>${a.name ? `<t:Name>${xmlEscape(a.name)}</t:Name>` : ''}</t:Mailbox></t:Attendee>`
          )
          .join('')}</t:Resources>`;
      }
    }

    const sendInvitations = attendees && attendees.length > 0 ? 'SendToAllAndSaveCopy' : 'SendToNone';
    const savedItemFolderIdXml = mailbox
      ? `<m:SavedItemFolderId><t:DistinguishedFolderId Id="calendar"><t:Mailbox><t:EmailAddress>${xmlEscape(mailbox)}</t:EmailAddress></t:Mailbox></t:DistinguishedFolderId></m:SavedItemFolderId>`
      : '';

    const envelope = soapEnvelope(`
    <m:CreateItem SendMeetingInvitations="${sendInvitations}">
      ${savedItemFolderIdXml}
      <m:Items>
        <t:CalendarItem>
          <t:Subject>${xmlEscape(subject)}</t:Subject>
          ${body ? `<t:Body BodyType="Text">${xmlEscape(body)}</t:Body>` : ''}
          <t:Start>${xmlEscape(start)}</t:Start>
          <t:End>${xmlEscape(end)}</t:End>
          ${location ? `<t:Location>${xmlEscape(location)}</t:Location>` : ''}
          ${attendeesXml}
          ${isOnlineMeeting ? '<t:IsOnlineMeeting>true</t:IsOnlineMeeting>' : ''}
          ${recurrence ? buildRecurrenceXml(recurrence) : ''}
        </t:CalendarItem>
      </m:Items>
    </m:CreateItem>`);

    const xml = await callEws(token, envelope, mailbox);
    const block = extractBlocks(xml, 'CalendarItem')[0] || '';
    const id = extractAttribute(block, 'ItemId', 'Id');

    return ewsResult({
      Id: id,
      Subject: subject,
      Start: { DateTime: start, TimeZone: undefined },
      End: { DateTime: end, TimeZone: undefined },
      WebLink: undefined,
      OnlineMeetingUrl: undefined
    });
  } catch (err) {
    return ewsError(err);
  }
}

export async function updateEvent(options: UpdateEventOptions): Promise<OwaResponse<CreatedEvent>> {
  try {
    const { token, eventId, changeKey, subject, start, end, body, location, attendees, isOnlineMeeting, mailbox } =
      options;

    const updates: string[] = [];

    if (subject !== undefined) {
      updates.push(
        `<t:SetItemField><t:FieldURI FieldURI="item:Subject" /><t:CalendarItem><t:Subject>${xmlEscape(subject)}</t:Subject></t:CalendarItem></t:SetItemField>`
      );
    }
    if (body !== undefined) {
      updates.push(
        `<t:SetItemField><t:FieldURI FieldURI="item:Body" /><t:CalendarItem><t:Body BodyType="Text">${xmlEscape(body)}</t:Body></t:CalendarItem></t:SetItemField>`
      );
    }
    if (start !== undefined) {
      updates.push(
        `<t:SetItemField><t:FieldURI FieldURI="calendar:Start" /><t:CalendarItem><t:Start>${xmlEscape(start)}</t:Start></t:CalendarItem></t:SetItemField>`
      );
    }
    if (end !== undefined) {
      updates.push(
        `<t:SetItemField><t:FieldURI FieldURI="calendar:End" /><t:CalendarItem><t:End>${xmlEscape(end)}</t:End></t:CalendarItem></t:SetItemField>`
      );
    }
    if (location !== undefined) {
      updates.push(
        `<t:SetItemField><t:FieldURI FieldURI="calendar:Location" /><t:CalendarItem><t:Location>${xmlEscape(location)}</t:Location></t:CalendarItem></t:SetItemField>`
      );
    }
    if (attendees !== undefined) {
      const required = attendees.filter((a) => (a.type || 'Required') !== 'Optional' && a.type !== 'Resource');
      const optional = attendees.filter((a) => a.type === 'Optional');
      const resources = attendees.filter((a) => a.type === 'Resource');

      if (required.length > 0) {
        updates.push(
          `<t:SetItemField><t:FieldURI FieldURI="calendar:RequiredAttendees" /><t:CalendarItem><t:RequiredAttendees>${required
            .map(
              (a) =>
                `<t:Attendee><t:Mailbox><t:EmailAddress>${xmlEscape(a.email)}</t:EmailAddress></t:Mailbox></t:Attendee>`
            )
            .join('')}</t:RequiredAttendees></t:CalendarItem></t:SetItemField>`
        );
      }
      if (optional.length > 0) {
        updates.push(
          `<t:SetItemField><t:FieldURI FieldURI="calendar:OptionalAttendees" /><t:CalendarItem><t:OptionalAttendees>${optional
            .map(
              (a) =>
                `<t:Attendee><t:Mailbox><t:EmailAddress>${xmlEscape(a.email)}</t:EmailAddress></t:Mailbox></t:Attendee>`
            )
            .join('')}</t:OptionalAttendees></t:CalendarItem></t:SetItemField>`
        );
      }
      if (resources.length > 0) {
        updates.push(
          `<t:SetItemField><t:FieldURI FieldURI="calendar:Resources" /><t:CalendarItem><t:Resources>${resources
            .map(
              (a) =>
                `<t:Attendee><t:Mailbox><t:EmailAddress>${xmlEscape(a.email)}</t:EmailAddress></t:Mailbox></t:Attendee>`
            )
            .join('')}</t:Resources></t:CalendarItem></t:SetItemField>`
        );
      }
    }

    if (updates.length === 0) {
      return { ok: false, status: 400, error: { code: 'NO_UPDATES', message: 'No fields to update' } };
    }

    const sendUpdates = attendees && attendees.length > 0 ? 'SendToAllAndSaveCopy' : 'SendToNone';

    const buildEnvelope = (conflictResolution: 'AutoResolve' | 'AlwaysOverwrite', includeChangeKey: boolean): string =>
      soapEnvelope(`
    <m:UpdateItem ConflictResolution="${conflictResolution}" SendMeetingInvitationsOrCancellations="${sendUpdates}">
      <m:ItemChanges>
        <t:ItemChange>
          <t:ItemId Id="${xmlEscape(eventId)}"${includeChangeKey && changeKey ? ` ChangeKey="${xmlEscape(changeKey)}"` : ''} />
          <t:Updates>
            ${updates.join('\n')}
          </t:Updates>
        </t:ItemChange>
      </m:ItemChanges>
    </m:UpdateItem>`);

    let xml: string;
    try {
      xml = await callEws(
        token,
        buildEnvelope(changeKey ? 'AutoResolve' : 'AlwaysOverwrite', Boolean(changeKey)),
        mailbox
      );
    } catch (err) {
      const message = err instanceof Error ? err.message : '';
      const isConflict =
        message.includes('ErrorIrresolvableConflict') ||
        message.includes('ErrorConflictResolutionRequired') ||
        message.includes('ErrorChangeKeyRequiredForWriteOperations');

      if (!changeKey || !isConflict) {
        throw err;
      }

      xml = await callEws(token, buildEnvelope('AlwaysOverwrite', false), mailbox);
    }
    const block = extractBlocks(xml, 'CalendarItem')[0] || '';
    const newId = extractAttribute(block, 'ItemId', 'Id') || eventId;

    return ewsResult({
      Id: newId,
      Subject: subject || '',
      Start: { DateTime: start || '', TimeZone: 'UTC' },
      End: { DateTime: end || '', TimeZone: 'UTC' }
    });
  } catch (err) {
    return ewsError(err);
  }
}

export interface DeleteEventOptions {
  token: string;
  eventId: string;
  mailbox?: string;
}

export async function deleteEvent(options: DeleteEventOptions): Promise<OwaResponse<void>> {
  try {
    const { token, eventId, mailbox } = options;
    const envelope = soapEnvelope(`
    <m:DeleteItem DeleteType="MoveToDeletedItems" SendMeetingCancellations="SendToNone">
      <m:ItemIds>
        <t:ItemId Id="${xmlEscape(eventId)}" />
      </m:ItemIds>
    </m:DeleteItem>`);
    await callEws(token, envelope, mailbox);
    return { ok: true, status: 200 };
  } catch (err) {
    return ewsError(err);
  }
}

export interface CancelEventOptions {
  token: string;
  eventId: string;
  comment?: string;
  mailbox?: string;
}

export async function cancelEvent(options: CancelEventOptions): Promise<OwaResponse<void>> {
  const { token, eventId, comment, mailbox } = options;

  // Primary: CancelCalendarItem
  try {
    const envelope = soapEnvelope(`
    <m:CreateItem MessageDisposition="SendAndSaveCopy">
      <m:Items>
        <t:CancelCalendarItem>
          <t:ReferenceItemId Id="${xmlEscape(eventId)}" />
          ${comment ? `<t:NewBodyContent BodyType="Text">${xmlEscape(comment)}</t:NewBodyContent>` : ''}
        </t:CancelCalendarItem>
      </m:Items>
    </m:CreateItem>`);
    await callEws(token, envelope, mailbox);
    return { ok: true, status: 200 };
  } catch (primaryErr) {
    // Fallback: DeleteItem with SendMeetingCancellations
    try {
      const envelope = soapEnvelope(`
      <m:DeleteItem DeleteType="MoveToDeletedItems" SendMeetingCancellations="SendToAllAndSaveCopy">
        <m:ItemIds>
          <t:ItemId Id="${xmlEscape(eventId)}" />
        </m:ItemIds>
      </m:DeleteItem>`);
      await callEws(token, envelope, mailbox);
      // Fallback succeeded after primary failed — report it so caller knows what happened
      const primaryMsg = primaryErr instanceof Error ? primaryErr.message : String(primaryErr);
      return {
        ok: true,
        status: 200,
        info: `Primary cancellation failed (${primaryMsg}); cancellation sent via fallback DeleteItem instead.`
      };
    } catch (fallbackErr) {
      // Both failed — report both errors clearly
      const primaryMsg = primaryErr instanceof Error ? primaryErr.message : String(primaryErr);
      const fallbackMsg = fallbackErr instanceof Error ? fallbackErr.message : String(fallbackErr);
      return {
        ok: false,
        status: 0,
        error: {
          code: 'EWS_CANCEL_FAILED',
          message: `Primary cancellation failed: ${primaryMsg}. Fallback also failed: ${fallbackMsg}`
        }
      };
    }
  }
}

export async function respondToEvent(options: RespondToEventOptions): Promise<OwaResponse<void>> {
  try {
    const { token, eventId, response, comment, sendResponse = true, mailbox } = options;
    const disposition = sendResponse ? 'SendAndSaveCopy' : 'SaveOnly';

    const responseTagMap: Record<ResponseType, string> = {
      accept: 'AcceptItem',
      decline: 'DeclineItem',
      tentative: 'TentativelyAcceptItem'
    };
    const tag = responseTagMap[response];

    const envelope = soapEnvelope(`
    <m:CreateItem MessageDisposition="${disposition}">
      <m:Items>
        <t:${tag}>
          <t:ReferenceItemId Id="${xmlEscape(eventId)}" />
          ${comment ? `<t:Body BodyType="Text">${xmlEscape(comment)}</t:Body>` : ''}
        </t:${tag}>
      </m:Items>
    </m:CreateItem>`);

    await callEws(token, envelope, mailbox);
    return { ok: true, status: 200 };
  } catch (err) {
    return ewsError(err);
  }
}

// ─── Mail Operations ───

export async function getEmails(options: GetEmailsOptions): Promise<OwaResponse<EmailListResponse>> {
  try {
    const { token, folder = 'inbox', top = 10, skip = 0, filter, search } = options;

    // Build restriction for filters
    let restrictionXml = '';
    if (filter && !search) {
      const restrictions: string[] = [];

      if (filter.includes('IsRead eq false')) {
        restrictions.push(`
        <t:IsEqualTo>
          <t:FieldURI FieldURI="message:IsRead" />
          <t:FieldURIOrConstant><t:Constant Value="false" /></t:FieldURIOrConstant>
        </t:IsEqualTo>`);
      }
      if (filter.includes('FlagStatus') && filter.includes('Flagged')) {
        restrictions.push(`
        <t:IsEqualTo>
          <t:FieldURI FieldURI="item:Flag/FlagStatus" />
          <t:FieldURIOrConstant><t:Constant Value="Flagged" /></t:FieldURIOrConstant>
        </t:IsEqualTo>`);
      }

      if (restrictions.length === 1) {
        restrictionXml = `<m:Restriction>${restrictions[0]}</m:Restriction>`;
      } else if (restrictions.length > 1) {
        restrictionXml = `<m:Restriction><t:And>${restrictions.join('')}</t:And></m:Restriction>`;
      }
    }

    const queryStringXml = search ? `<m:QueryString>${xmlEscape(search)}</m:QueryString>` : '';

    const envelope = soapEnvelope(`
    <m:FindItem Traversal="Shallow">
      <m:ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="item:Subject" />
          <t:FieldURI FieldURI="item:DateTimeReceived" />
          <t:FieldURI FieldURI="item:HasAttachments" />
          <t:FieldURI FieldURI="item:Importance" />
          <t:FieldURI FieldURI="item:Preview" />
          <t:FieldURI FieldURI="message:From" />
          <t:FieldURI FieldURI="message:IsRead" />
          <t:FieldURI FieldURI="item:Flag" />
          <t:FieldURI FieldURI="item:IsDraft" />
        </t:AdditionalProperties>
      </m:ItemShape>
      <m:IndexedPageItemView MaxEntriesReturned="${top}" Offset="${skip}" BasePoint="Beginning" />
      ${restrictionXml}
      ${
        !search
          ? `<m:SortOrder>
        <t:FieldOrder Order="Descending">
          <t:FieldURI FieldURI="item:DateTimeReceived" />
        </t:FieldOrder>
      </m:SortOrder>`
          : ''
      }
      ${queryStringXml}
      <m:ParentFolderIds>
        ${folderIdXml(folder)}
      </m:ParentFolderIds>
    </m:FindItem>`);

    const xml = await callEws(token, envelope);
    const blocks = extractBlocks(xml, 'Message');
    const emails = blocks.map(parseEmailMessage);

    return ewsResult({ value: emails });
  } catch (err) {
    return ewsError(err);
  }
}

export async function getEmail(token: string, messageId: string): Promise<OwaResponse<EmailMessage>> {
  try {
    const envelope = soapEnvelope(`
    <m:GetItem>
      <m:ItemShape>
        <t:BaseShape>Default</t:BaseShape>
        <t:BodyType>Text</t:BodyType>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="item:Body" />
          <t:FieldURI FieldURI="item:DateTimeReceived" />
          <t:FieldURI FieldURI="item:HasAttachments" />
          <t:FieldURI FieldURI="message:From" />
          <t:FieldURI FieldURI="message:ToRecipients" />
          <t:FieldURI FieldURI="message:CcRecipients" />
          <t:FieldURI FieldURI="message:IsRead" />
          <t:FieldURI FieldURI="item:Flag" />
          <t:FieldURI FieldURI="item:Importance" />
        </t:AdditionalProperties>
      </m:ItemShape>
      <m:ItemIds>
        <t:ItemId Id="${xmlEscape(messageId)}" />
      </m:ItemIds>
    </m:GetItem>`);

    const xml = await callEws(token, envelope);
    const block = extractBlocks(xml, 'Message')[0];
    if (!block) return { ok: false, status: 404, error: { code: 'NOT_FOUND', message: 'Message not found' } };

    return ewsResult(parseEmailMessage(block));
  } catch (err) {
    return ewsError(err);
  }
}

export async function sendEmail(
  token: string,
  options: {
    to: string[];
    cc?: string[];
    bcc?: string[];
    subject: string;
    body: string;
    bodyType?: 'Text' | 'HTML';
    attachments?: EmailAttachment[];
    mailbox?: string;
  }
): Promise<OwaResponse<void>> {
  try {
    const { mailbox } = options;

    const toXml = options.to
      .map((e) => `<t:Mailbox><t:EmailAddress>${xmlEscape(e)}</t:EmailAddress></t:Mailbox>`)
      .join('');

    const ccXml =
      options.cc && options.cc.length > 0
        ? `<t:CcRecipients>${options.cc
            .map((e) => `<t:Mailbox><t:EmailAddress>${xmlEscape(e)}</t:EmailAddress></t:Mailbox>`)
            .join('')}</t:CcRecipients>`
        : '';

    const bccXml =
      options.bcc && options.bcc.length > 0
        ? `<t:BccRecipients>${options.bcc
            .map((e) => `<t:Mailbox><t:EmailAddress>${xmlEscape(e)}</t:EmailAddress></t:Mailbox>`)
            .join('')}</t:BccRecipients>`
        : '';

    const bodyType = options.bodyType || 'Text';

    // Build From element for shared mailbox (Send As)
    const fromXml = mailbox
      ? `<t:From><t:Mailbox><t:EmailAddress>${xmlEscape(mailbox)}</t:EmailAddress></t:Mailbox></t:From>`
      : '';

    // Build SavedItemFolderId targeting shared mailbox sentitems
    const savedItemFolderIdXml = mailbox
      ? `<m:SavedItemFolderId><t:DistinguishedFolderId Id="sentitems"><t:Mailbox><t:EmailAddress>${xmlEscape(mailbox)}</t:EmailAddress></t:Mailbox></t:DistinguishedFolderId></m:SavedItemFolderId>`
      : '';

    // If no attachments, send directly
    if (!options.attachments || options.attachments.length === 0) {
      const envelope = soapEnvelope(`
      <m:CreateItem MessageDisposition="SendAndSaveCopy">
        ${savedItemFolderIdXml}
        <m:Items>
          <t:Message>
            ${fromXml}
            <t:Subject>${xmlEscape(options.subject)}</t:Subject>
            <t:Body BodyType="${bodyType}">${xmlEscape(options.body)}</t:Body>
            <t:ToRecipients>${toXml}</t:ToRecipients>
            ${ccXml}
            ${bccXml}
          </t:Message>
        </m:Items>
      </m:CreateItem>`);
      await callEws(token, envelope, mailbox);
      return { ok: true, status: 200 };
    }

    // With attachments: create draft, add attachments, send
    const draftResult = await createDraft(token, {
      to: options.to,
      cc: options.cc,
      subject: options.subject,
      body: options.body,
      bodyType
    });
    if (!draftResult.ok || !draftResult.data) return draftResult as OwaResponse<void>;

    for (const att of options.attachments) {
      await addAttachmentToItem(token, draftResult.data.Id, att);
    }

    await sendItemById(token, draftResult.data.Id);
    return { ok: true, status: 200 };
  } catch (err) {
    return ewsError(err);
  }
}

export async function replyToEmail(
  token: string,
  messageId: string,
  comment: string,
  replyAll: boolean = false,
  isHtml: boolean = false,
  mailbox?: string
): Promise<OwaResponse<void>> {
  try {
    const tag = replyAll ? 'ReplyAllToItem' : 'ReplyToItem';
    const bodyType = isHtml ? 'HTML' : 'Text';

    const envelope = soapEnvelope(`
    <m:CreateItem MessageDisposition="SendAndSaveCopy">
      <m:Items>
        <t:${tag}>
          <t:ReferenceItemId Id="${xmlEscape(messageId)}" />
          <t:NewBodyContent BodyType="${bodyType}">${xmlEscape(comment)}</t:NewBodyContent>
        </t:${tag}>
      </m:Items>
    </m:CreateItem>`);

    await callEws(token, envelope, mailbox);
    return { ok: true, status: 200 };
  } catch (err) {
    return ewsError(err);
  }
}

export async function replyToEmailDraft(
  token: string,
  messageId: string,
  comment: string,
  replyAll: boolean = false,
  isHtml: boolean = false,
  mailbox?: string
): Promise<OwaResponse<{ draftId: string }>> {
  try {
    const tag = replyAll ? 'ReplyAllToItem' : 'ReplyToItem';
    const bodyType = isHtml ? 'HTML' : 'Text';

    const envelope = soapEnvelope(`
    <m:CreateItem MessageDisposition="SaveOnly">
      <m:Items>
        <t:${tag}>
          <t:ReferenceItemId Id="${xmlEscape(messageId)}" />
          <t:NewBodyContent BodyType="${bodyType}">${xmlEscape(comment)}</t:NewBodyContent>
        </t:${tag}>
      </m:Items>
    </m:CreateItem>`);

    const xml = await callEws(token, envelope, mailbox);
    const draftId = extractAttribute(xml, 'ItemId', 'Id');
    return ewsResult({ draftId });
  } catch (err) {
    return ewsError(err);
  }
}

export async function forwardEmail(
  token: string,
  messageId: string,
  toRecipients: string[],
  comment?: string,
  mailbox?: string
): Promise<OwaResponse<void>> {
  try {
    const toXml = toRecipients
      .map((e) => `<t:Mailbox><t:EmailAddress>${xmlEscape(e)}</t:EmailAddress></t:Mailbox>`)
      .join('');

    const envelope = soapEnvelope(`
    <m:CreateItem MessageDisposition="SendAndSaveCopy">
      <m:Items>
        <t:ForwardItem>
          <t:ReferenceItemId Id="${xmlEscape(messageId)}" />
          ${comment ? `<t:NewBodyContent BodyType="Text">${xmlEscape(comment)}</t:NewBodyContent>` : ''}
          <t:ToRecipients>${toXml}</t:ToRecipients>
        </t:ForwardItem>
      </m:Items>
    </m:CreateItem>`);

    await callEws(token, envelope, mailbox);
    return { ok: true, status: 200 };
  } catch (err) {
    return ewsError(err);
  }
}

export async function updateEmail(
  token: string,
  messageId: string,
  updates: {
    IsRead?: boolean;
    Flag?: { FlagStatus: 'NotFlagged' | 'Flagged' | 'Complete' };
  }
): Promise<OwaResponse<EmailMessage>> {
  try {
    const setFields: string[] = [];

    if (updates.IsRead !== undefined) {
      setFields.push(`
      <t:SetItemField>
        <t:FieldURI FieldURI="message:IsRead" />
        <t:Message><t:IsRead>${updates.IsRead}</t:IsRead></t:Message>
      </t:SetItemField>`);
    }

    if (updates.Flag) {
      setFields.push(`
      <t:SetItemField>
        <t:FieldURI FieldURI="item:Flag" />
        <t:Message><t:Flag><t:FlagStatus>${xmlEscape(updates.Flag.FlagStatus)}</t:FlagStatus></t:Flag></t:Message>
      </t:SetItemField>`);
    }

    const envelope = soapEnvelope(`
    <m:UpdateItem ConflictResolution="AlwaysOverwrite" MessageDisposition="SaveOnly" SuppressReadReceipts="true">
      <m:ItemChanges>
        <t:ItemChange>
          <t:ItemId Id="${xmlEscape(messageId)}" />
          <t:Updates>
            ${setFields.join('')}
          </t:Updates>
        </t:ItemChange>
      </m:ItemChanges>
    </m:UpdateItem>`);

    const xml = await callEws(token, envelope);
    const newId = extractAttribute(xml, 'ItemId', 'Id') || messageId;
    return ewsResult({ Id: newId } as EmailMessage);
  } catch (err) {
    return ewsError(err);
  }
}

export async function moveEmail(
  token: string,
  messageId: string,
  destinationFolder: string
): Promise<OwaResponse<EmailMessage>> {
  try {
    const envelope = soapEnvelope(`
    <m:MoveItem>
      <m:ToFolderId>
        ${folderIdXml(destinationFolder)}
      </m:ToFolderId>
      <m:ItemIds>
        <t:ItemId Id="${xmlEscape(messageId)}" />
      </m:ItemIds>
    </m:MoveItem>`);

    const xml = await callEws(token, envelope);
    const newId = extractAttribute(xml, 'ItemId', 'Id') || messageId;
    return ewsResult({ Id: newId } as EmailMessage);
  } catch (err) {
    return ewsError(err);
  }
}

// ─── Draft Operations ───

export async function createDraft(
  token: string,
  options: {
    to?: string[];
    cc?: string[];
    subject?: string;
    body?: string;
    bodyType?: 'Text' | 'HTML';
  }
): Promise<OwaResponse<{ Id: string }>> {
  try {
    const toXml =
      options.to && options.to.length > 0
        ? `<t:ToRecipients>${options.to
            .map((e) => `<t:Mailbox><t:EmailAddress>${xmlEscape(e)}</t:EmailAddress></t:Mailbox>`)
            .join('')}</t:ToRecipients>`
        : '';

    const ccXml =
      options.cc && options.cc.length > 0
        ? `<t:CcRecipients>${options.cc
            .map((e) => `<t:Mailbox><t:EmailAddress>${xmlEscape(e)}</t:EmailAddress></t:Mailbox>`)
            .join('')}</t:CcRecipients>`
        : '';

    const bodyType = options.bodyType || 'Text';

    const envelope = soapEnvelope(`
    <m:CreateItem MessageDisposition="SaveOnly">
      <m:Items>
        <t:Message>
          ${options.subject ? `<t:Subject>${xmlEscape(options.subject)}</t:Subject>` : ''}
          ${options.body ? `<t:Body BodyType="${bodyType}">${xmlEscape(options.body)}</t:Body>` : ''}
          ${toXml}
          ${ccXml}
        </t:Message>
      </m:Items>
    </m:CreateItem>`);

    const xml = await callEws(token, envelope);
    const id = extractAttribute(xml, 'ItemId', 'Id');
    return ewsResult({ Id: id });
  } catch (err) {
    return ewsError(err);
  }
}

export async function updateDraft(
  token: string,
  draftId: string,
  options: {
    to?: string[];
    cc?: string[];
    subject?: string;
    body?: string;
    bodyType?: 'Text' | 'HTML';
  }
): Promise<OwaResponse<void>> {
  try {
    const setFields: string[] = [];

    if (options.subject !== undefined) {
      setFields.push(
        `<t:SetItemField><t:FieldURI FieldURI="item:Subject" /><t:Message><t:Subject>${xmlEscape(options.subject)}</t:Subject></t:Message></t:SetItemField>`
      );
    }
    if (options.body !== undefined) {
      const bodyType = options.bodyType || 'Text';
      setFields.push(
        `<t:SetItemField><t:FieldURI FieldURI="item:Body" /><t:Message><t:Body BodyType="${bodyType}">${xmlEscape(options.body)}</t:Body></t:Message></t:SetItemField>`
      );
    }
    if (options.to !== undefined) {
      setFields.push(
        `<t:SetItemField><t:FieldURI FieldURI="message:ToRecipients" /><t:Message><t:ToRecipients>${options.to
          .map((e) => `<t:Mailbox><t:EmailAddress>${xmlEscape(e)}</t:EmailAddress></t:Mailbox>`)
          .join('')}</t:ToRecipients></t:Message></t:SetItemField>`
      );
    }
    if (options.cc !== undefined) {
      setFields.push(
        `<t:SetItemField><t:FieldURI FieldURI="message:CcRecipients" /><t:Message><t:CcRecipients>${options.cc
          .map((e) => `<t:Mailbox><t:EmailAddress>${xmlEscape(e)}</t:EmailAddress></t:Mailbox>`)
          .join('')}</t:CcRecipients></t:Message></t:SetItemField>`
      );
    }

    const envelope = soapEnvelope(`
    <m:UpdateItem ConflictResolution="AlwaysOverwrite" MessageDisposition="SaveOnly">
      <m:ItemChanges>
        <t:ItemChange>
          <t:ItemId Id="${xmlEscape(draftId)}" />
          <t:Updates>${setFields.join('')}</t:Updates>
        </t:ItemChange>
      </m:ItemChanges>
    </m:UpdateItem>`);

    await callEws(token, envelope);
    return { ok: true, status: 200 };
  } catch (err) {
    return ewsError(err);
  }
}

async function sendItemById(token: string, itemId: string): Promise<void> {
  const envelope = soapEnvelope(`
  <m:SendItem SaveItemToFolder="true">
    <m:ItemIds>
      <t:ItemId Id="${xmlEscape(itemId)}" />
    </m:ItemIds>
    <m:SavedItemFolderId>
      <t:DistinguishedFolderId Id="sentitems" />
    </m:SavedItemFolderId>
  </m:SendItem>`);
  await callEws(token, envelope);
}

export async function sendDraftById(token: string, draftId: string): Promise<OwaResponse<void>> {
  try {
    await sendItemById(token, draftId);
    return { ok: true, status: 200 };
  } catch (err) {
    return ewsError(err);
  }
}

export async function deleteDraftById(token: string, draftId: string): Promise<OwaResponse<void>> {
  try {
    const envelope = soapEnvelope(`
    <m:DeleteItem DeleteType="HardDelete">
      <m:ItemIds>
        <t:ItemId Id="${xmlEscape(draftId)}" />
      </m:ItemIds>
    </m:DeleteItem>`);
    await callEws(token, envelope);
    return { ok: true, status: 200 };
  } catch (err) {
    return ewsError(err);
  }
}

async function addAttachmentToItem(token: string, itemId: string, attachment: EmailAttachment): Promise<void> {
  const envelope = soapEnvelope(`
  <m:CreateAttachment>
    <m:ParentItemId Id="${xmlEscape(itemId)}" />
    <m:Attachments>
      <t:FileAttachment>
        <t:Name>${xmlEscape(attachment.name)}</t:Name>
        <t:ContentType>${xmlEscape(attachment.contentType)}</t:ContentType>
        <t:Content>${attachment.contentBytes}</t:Content>
      </t:FileAttachment>
    </m:Attachments>
  </m:CreateAttachment>`);
  await callEws(token, envelope);
}

export async function addAttachmentToDraft(
  token: string,
  draftId: string,
  attachment: { name: string; contentType: string; contentBytes: string }
): Promise<OwaResponse<void>> {
  try {
    await addAttachmentToItem(token, draftId, attachment);
    return { ok: true, status: 200 };
  } catch (err) {
    return ewsError(err);
  }
}

// ─── Folder Operations ───

export async function getMailFolders(
  token: string,
  parentFolderId?: string
): Promise<OwaResponse<MailFolderListResponse>> {
  try {
    const parentXml = parentFolderId
      ? `<t:FolderId Id="${xmlEscape(parentFolderId)}" />`
      : '<t:DistinguishedFolderId Id="msgfolderroot" />';

    const envelope = soapEnvelope(`
    <m:FindFolder Traversal="Shallow">
      <m:FolderShape>
        <t:BaseShape>Default</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="folder:ChildFolderCount" />
          <t:FieldURI FieldURI="folder:UnreadCount" />
          <t:FieldURI FieldURI="folder:TotalCount" />
        </t:AdditionalProperties>
      </m:FolderShape>
      <m:ParentFolderIds>
        ${parentXml}
      </m:ParentFolderIds>
    </m:FindFolder>`);

    const xml = await callEws(token, envelope);
    const blocks = extractBlocks(xml, 'Folder');
    const folders = blocks.map(parseFolder);

    return ewsResult({ value: folders });
  } catch (err) {
    return ewsError(err);
  }
}

export async function createMailFolder(
  token: string,
  displayName: string,
  parentFolderId?: string
): Promise<OwaResponse<MailFolder>> {
  try {
    const parentXml = parentFolderId
      ? `<t:FolderId Id="${xmlEscape(parentFolderId)}" />`
      : '<t:DistinguishedFolderId Id="msgfolderroot" />';

    const envelope = soapEnvelope(`
    <m:CreateFolder>
      <m:ParentFolderId>
        ${parentXml}
      </m:ParentFolderId>
      <m:Folders>
        <t:Folder>
          <t:DisplayName>${xmlEscape(displayName)}</t:DisplayName>
        </t:Folder>
      </m:Folders>
    </m:CreateFolder>`);

    const xml = await callEws(token, envelope);
    const block = extractBlocks(xml, 'Folder')[0] || '';

    return ewsResult({
      Id: extractAttribute(block, 'FolderId', 'Id'),
      DisplayName: displayName,
      ChildFolderCount: 0,
      UnreadItemCount: 0,
      TotalItemCount: 0
    });
  } catch (err) {
    return ewsError(err);
  }
}

export async function updateMailFolder(
  token: string,
  folderId: string,
  displayName: string
): Promise<OwaResponse<MailFolder>> {
  try {
    const envelope = soapEnvelope(`
    <m:UpdateFolder>
      <m:FolderChanges>
        <t:FolderChange>
          <t:FolderId Id="${xmlEscape(folderId)}" />
          <t:Updates>
            <t:SetFolderField>
              <t:FieldURI FieldURI="folder:DisplayName" />
              <t:Folder>
                <t:DisplayName>${xmlEscape(displayName)}</t:DisplayName>
              </t:Folder>
            </t:SetFolderField>
          </t:Updates>
        </t:FolderChange>
      </m:FolderChanges>
    </m:UpdateFolder>`);

    await callEws(token, envelope);

    return ewsResult({
      Id: folderId,
      DisplayName: displayName,
      ChildFolderCount: 0,
      UnreadItemCount: 0,
      TotalItemCount: 0
    });
  } catch (err) {
    return ewsError(err);
  }
}

export async function deleteMailFolder(token: string, folderId: string): Promise<OwaResponse<void>> {
  try {
    const envelope = soapEnvelope(`
    <m:DeleteFolder DeleteType="MoveToDeletedItems">
      <m:FolderIds>
        <t:FolderId Id="${xmlEscape(folderId)}" />
      </m:FolderIds>
    </m:DeleteFolder>`);

    await callEws(token, envelope);
    return { ok: true, status: 200 };
  } catch (err) {
    return ewsError(err);
  }
}

// ─── Attachment Operations ───

export async function getAttachments(token: string, messageId: string): Promise<OwaResponse<AttachmentListResponse>> {
  try {
    // First get the item to find attachment IDs
    const envelope = soapEnvelope(`
    <m:GetItem>
      <m:ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="item:Attachments" />
        </t:AdditionalProperties>
      </m:ItemShape>
      <m:ItemIds>
        <t:ItemId Id="${xmlEscape(messageId)}" />
      </m:ItemIds>
    </m:GetItem>`);

    const xml = await callEws(token, envelope);
    const attachBlocks = extractBlocks(xml, 'FileAttachment');

    const attachments: Attachment[] = attachBlocks.map((ab) => ({
      Id: extractAttribute(ab, 'AttachmentId', 'Id'),
      Name: extractTag(ab, 'Name'),
      ContentType: extractTag(ab, 'ContentType') || 'application/octet-stream',
      Size: parseInt(extractTag(ab, 'Size') || '0', 10),
      IsInline: extractTag(ab, 'IsInline').toLowerCase() === 'true',
      ContentId: extractTag(ab, 'ContentId') || undefined
    }));

    return ewsResult({ value: attachments });
  } catch (err) {
    return ewsError(err);
  }
}

export async function getAttachment(
  token: string,
  _messageId: string,
  attachmentId: string
): Promise<OwaResponse<Attachment>> {
  try {
    const envelope = soapEnvelope(`
    <m:GetAttachment>
      <m:AttachmentIds>
        <t:AttachmentId Id="${xmlEscape(attachmentId)}" />
      </m:AttachmentIds>
    </m:GetAttachment>`);

    const xml = await callEws(token, envelope);
    const block = extractBlocks(xml, 'FileAttachment')[0] || '';

    return ewsResult({
      Id: extractAttribute(block, 'AttachmentId', 'Id'),
      Name: extractTag(block, 'Name'),
      ContentType: extractTag(block, 'ContentType') || 'application/octet-stream',
      Size: parseInt(extractTag(block, 'Size') || '0', 10),
      IsInline: extractTag(block, 'IsInline').toLowerCase() === 'true',
      ContentId: extractTag(block, 'ContentId') || undefined,
      ContentBytes: extractTag(block, 'Content') || undefined
    });
  } catch (err) {
    return ewsError(err);
  }
}

// ─── People & Rooms ───

export async function resolveNames(
  token: string,
  query: string
): Promise<
  OwaResponse<
    Array<{
      DisplayName?: string;
      EmailAddress?: string;
      JobTitle?: string;
      Department?: string;
      OfficeLocation?: string;
      MailboxType?: string;
    }>
  >
> {
  try {
    const envelope = soapEnvelope(`
    <m:ResolveNames ReturnFullContactData="true" SearchScope="ActiveDirectoryContacts">
      <m:UnresolvedEntry>${xmlEscape(query)}</m:UnresolvedEntry>
    </m:ResolveNames>`);

    const xml = await callEws(token, envelope);
    const resolutions = extractBlocks(xml, 'Resolution');

    const results = resolutions.map((block) => {
      const mailbox = extractSelfClosingOrBlock(block, 'Mailbox');
      const contact = extractSelfClosingOrBlock(block, 'Contact');

      return {
        DisplayName: extractTag(mailbox, 'Name') || extractTag(contact, 'DisplayName'),
        EmailAddress: extractTag(mailbox, 'EmailAddress'),
        JobTitle: extractTag(contact, 'JobTitle') || undefined,
        Department: extractTag(contact, 'Department') || undefined,
        OfficeLocation: extractTag(contact, 'OfficeLocation') || undefined,
        MailboxType: extractTag(mailbox, 'MailboxType') || undefined
      };
    });

    return ewsResult(results);
  } catch (err) {
    return ewsError(err);
  }
}

export async function getRoomLists(token: string): Promise<OwaResponse<RoomList[]>> {
  try {
    const envelope = soapEnvelope('<m:GetRoomLists />');
    const xml = await callEws(token, envelope);
    const addresses = extractBlocks(xml, 'Address');

    const lists: RoomList[] = addresses.map((block) => ({
      Name: extractTag(block, 'Name'),
      Address: extractTag(block, 'EmailAddress')
    }));

    return ewsResult(lists);
  } catch (err) {
    return ewsError(err);
  }
}

export async function getRooms(token: string, roomListAddress?: string): Promise<OwaResponse<Room[]>> {
  try {
    if (roomListAddress) {
      const envelope = soapEnvelope(`
      <m:GetRooms>
        <m:RoomList>
          <t:EmailAddress>${xmlEscape(roomListAddress)}</t:EmailAddress>
        </m:RoomList>
      </m:GetRooms>`);
      const xml = await callEws(token, envelope);
      const rooms = extractBlocks(xml, 'Room').map((block) => {
        const id = extractSelfClosingOrBlock(block, 'Id');
        return {
          Name: extractTag(id, 'Name'),
          Address: extractTag(id, 'EmailAddress')
        };
      });
      return ewsResult(rooms);
    }

    // No room list specified: get all room lists first, then rooms from each
    const listsResult = await getRoomLists(token);
    if (!listsResult.ok || !listsResult.data || listsResult.data.length === 0) {
      return ewsResult([]);
    }

    const allRooms: Room[] = [];
    for (const list of listsResult.data) {
      const roomsResult = await getRooms(token, list.Address);
      if (roomsResult.ok && roomsResult.data) {
        allRooms.push(...roomsResult.data);
      }
    }

    return ewsResult(allRooms);
  } catch (err) {
    return ewsError(err);
  }
}

export async function searchRooms(token: string, query: string = 'room'): Promise<OwaResponse<Room[]>> {
  // Use ResolveNames to find rooms by name
  try {
    const result = await resolveNames(token, query);
    if (!result.ok || !result.data) return ewsResult([]);

    // Try to filter to rooms (MailboxType might indicate this)
    const rooms: Room[] = result.data
      .filter((r) => r.EmailAddress)
      .map((r) => ({
        Name: r.DisplayName || '',
        Address: r.EmailAddress || ''
      }));

    return ewsResult(rooms);
  } catch (err) {
    return ewsError(err);
  }
}

// ─── Availability ───

export async function getScheduleViaOutlook(
  token: string,
  emails: string[],
  startDateTime: string,
  endDateTime: string,
  durationMinutes: number = 30
): Promise<OwaResponse<ScheduleInfo[]>> {
  try {
    // SuggestionsViewOptions requires dates at midnight with no timezone offset
    const suggestStartD = new Date(startDateTime);
    suggestStartD.setHours(0, 0, 0, 0);
    const suggestEndD = new Date(endDateTime);
    suggestEndD.setHours(0, 0, 0, 0);
    suggestEndD.setDate(suggestEndD.getDate() + 1);
    const pad = (n: number) => String(n).padStart(2, '0');
    const toMidnight = (d: Date) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T00:00:00`;
    const suggestStart = toMidnight(suggestStartD);
    const suggestEnd = toMidnight(suggestEndD);

    const mailboxDataXml = emails
      .map(
        (email) => `
    <t:MailboxData>
      <t:Email><t:Address>${xmlEscape(email)}</t:Address></t:Email>
      <t:AttendeeType>Required</t:AttendeeType>
    </t:MailboxData>`
      )
      .join('');

    const envelope = soapEnvelope(`
    <m:GetUserAvailabilityRequest>
      <t:TimeZone>
        <t:Bias>-60</t:Bias>
        <t:StandardTime>
          <t:Bias>0</t:Bias>
          <t:Time>03:00:00</t:Time>
          <t:DayOrder>5</t:DayOrder>
          <t:Month>10</t:Month>
          <t:DayOfWeek>Sunday</t:DayOfWeek>
        </t:StandardTime>
        <t:DaylightTime>
          <t:Bias>-60</t:Bias>
          <t:Time>02:00:00</t:Time>
          <t:DayOrder>5</t:DayOrder>
          <t:Month>3</t:Month>
          <t:DayOfWeek>Sunday</t:DayOfWeek>
        </t:DaylightTime>
      </t:TimeZone>
      <m:MailboxDataArray>
        ${mailboxDataXml}
      </m:MailboxDataArray>
      <t:FreeBusyViewOptions>
        <t:TimeWindow>
          <t:StartTime>${xmlEscape(suggestStart)}</t:StartTime>
          <t:EndTime>${xmlEscape(suggestEnd)}</t:EndTime>
        </t:TimeWindow>
        <t:MergedFreeBusyIntervalInMinutes>${durationMinutes}</t:MergedFreeBusyIntervalInMinutes>
        <t:RequestedView>DetailedMerged</t:RequestedView>
      </t:FreeBusyViewOptions>
      <t:SuggestionsViewOptions>
        <t:GoodThreshold>25</t:GoodThreshold>
        <t:MaximumResultsByDay>10</t:MaximumResultsByDay>
        <t:MaximumNonWorkHourResultsByDay>0</t:MaximumNonWorkHourResultsByDay>
        <t:MeetingDurationInMinutes>${durationMinutes}</t:MeetingDurationInMinutes>
        <t:DetailedSuggestionsWindow>
          <t:StartTime>${xmlEscape(suggestStart)}</t:StartTime>
          <t:EndTime>${xmlEscape(suggestEnd)}</t:EndTime>
        </t:DetailedSuggestionsWindow>
      </t:SuggestionsViewOptions>
    </m:GetUserAvailabilityRequest>`);

    const xml = await callEws(token, envelope);

    // Parse suggestions into free slots
    const schedules: ScheduleInfo[] = emails.map((email) => ({
      scheduleId: email,
      availabilityView: '',
      scheduleItems: []
    }));

    // Extract suggestions
    const suggestions = extractBlocks(xml, 'Suggestion');
    const freeSlots: Array<{ start: string; end: string }> = [];

    for (const suggestion of suggestions) {
      const meetingTime = extractTag(suggestion, 'MeetingTime');
      if (meetingTime) {
        const startTime = new Date(meetingTime);
        const endTime = new Date(startTime.getTime() + durationMinutes * 60 * 1000);
        freeSlots.push({
          start: startTime.toISOString(),
          end: endTime.toISOString()
        });
      }
    }

    // Apply free slots to all schedules
    for (const schedule of schedules) {
      schedule.scheduleItems = freeSlots.map((slot) => ({
        status: 'Free',
        start: { dateTime: slot.start, timeZone: 'W. Europe Standard Time' },
        end: { dateTime: slot.end, timeZone: 'W. Europe Standard Time' }
      }));
    }

    if (freeSlots.length === 0) {
      for (const schedule of schedules) {
        schedule.scheduleItems = [
          {
            status: 'Busy',
            start: { dateTime: startDateTime, timeZone: 'W. Europe Standard Time' },
            end: { dateTime: endDateTime, timeZone: 'W. Europe Standard Time' },
            subject: 'No available times'
          }
        ];
      }
    }

    return ewsResult(schedules);
  } catch (err) {
    return ewsError(err);
  }
}

export async function getFreeBusy(
  token: string,
  startDateTime: string,
  endDateTime: string
): Promise<OwaResponse<FreeBusySlot[]>> {
  const result = await getCalendarEvents(token, startDateTime, endDateTime);
  if (!result.ok || !result.data) return { ok: false, status: result.status, error: result.error };

  const slots: FreeBusySlot[] = result.data
    .filter((event) => !event.IsCancelled)
    .map((event) => ({
      status:
        event.ShowAs === 'Free'
          ? ('Free' as const)
          : event.ShowAs === 'Tentative'
            ? ('Tentative' as const)
            : ('Busy' as const),
      start: event.Start.DateTime,
      end: event.End.DateTime,
      subject: event.Subject
    }));

  return ewsResult(slots);
}

export async function isRoomFree(
  token: string,
  roomEmail: string,
  startDateTime: string,
  endDateTime: string
): Promise<boolean> {
  try {
    const envelope = soapEnvelope(`
    <m:GetUserAvailabilityRequest>
      <t:TimeZone>
        <t:Bias>-60</t:Bias>
        <t:StandardTime>
          <t:Bias>0</t:Bias>
          <t:Time>03:00:00</t:Time>
          <t:DayOrder>5</t:DayOrder>
          <t:Month>10</t:Month>
          <t:DayOfWeek>Sunday</t:DayOfWeek>
        </t:StandardTime>
        <t:DaylightTime>
          <t:Bias>-60</t:Bias>
          <t:Time>02:00:00</t:Time>
          <t:DayOrder>5</t:DayOrder>
          <t:Month>3</t:Month>
          <t:DayOfWeek>Sunday</t:DayOfWeek>
        </t:DaylightTime>
      </t:TimeZone>
      <m:MailboxDataArray>
        <t:MailboxData>
          <t:Email><t:Address>${xmlEscape(roomEmail)}</t:Address></t:Email>
          <t:AttendeeType>Required</t:AttendeeType>
        </t:MailboxData>
      </m:MailboxDataArray>
      <t:FreeBusyViewOptions>
        <t:TimeWindow>
          <t:StartTime>${xmlEscape(startDateTime)}</t:StartTime>
          <t:EndTime>${xmlEscape(endDateTime)}</t:EndTime>
        </t:TimeWindow>
        <t:MergedFreeBusyIntervalInMinutes>15</t:MergedFreeBusyIntervalInMinutes>
        <t:RequestedView>FreeBusy</t:RequestedView>
      </t:FreeBusyViewOptions>
    </m:GetUserAvailabilityRequest>`);

    const xml = await callEws(token, envelope);
    const calendarEvents = extractBlocks(xml, 'CalendarEvent');

    // If no calendar events in the window, room is free
    if (calendarEvents.length === 0) return true;

    // Check if any event overlaps with our requested time
    const reqStart = new Date(startDateTime).getTime();
    const reqEnd = new Date(endDateTime).getTime();

    for (const event of calendarEvents) {
      const busyType = extractTag(event, 'BusyType');
      if (busyType === 'Free') continue;

      const evStart = new Date(extractTag(event, 'StartTime') || '').getTime();
      const evEnd = new Date(extractTag(event, 'EndTime') || '').getTime();

      // Check overlap
      if (evStart < reqEnd && evEnd > reqStart) {
        return false;
      }
    }

    return true;
  } catch {
    return false;
  }
}

export interface AutoReplyRule {
  messageText: string;
  enabled: boolean;
  startTime?: Date;
  endTime?: Date;
}

export async function getAutoReplyRule(token: string, mailbox?: string): Promise<OwaResponse<AutoReplyRule | null>> {
  try {
    const address = mailbox || EWS_USERNAME;
    const envelope = soapEnvelope(`
      <m:GetInboxRules>
        <m:MailboxSmtpAddress>${xmlEscape(address)}</m:MailboxSmtpAddress>
      </m:GetInboxRules>
    `);

    const xml = await callEws(token, envelope, address);

    // Parse the rules
    // Find the rule with DisplayName = "AutoReplyTemplate"
    const rulesRegex = /<t:Rule>(.*?)<\/t:Rule>/gs;
    let match;
    let ruleXml = null;
    while ((match = rulesRegex.exec(xml)) !== null) {
      if (match[1].includes('<t:DisplayName>AutoReplyTemplate</t:DisplayName>')) {
        ruleXml = match[1];
        break;
      }
    }

    if (!ruleXml) {
      return ewsResult(null);
    }

    const enabledStr = extractTag(ruleXml, 'IsEnabled');
    const enabled = enabledStr.toLowerCase() === 'true';

    // Dates
    const startStr = extractTag(ruleXml, 'StartDateTime');
    const endStr = extractTag(ruleXml, 'EndDateTime');

    // To get the message text, we need the template item ID
    const templateId = extractAttribute(ruleXml, 'ItemId', 'Id');
    let messageText = '';

    if (templateId) {
      // Fetch the template draft to read the body
      const getTemplateEnvelope = soapEnvelope(`
        <m:GetItem>
          <m:ItemShape>
            <t:BaseShape>Default</t:BaseShape>
            <t:AdditionalProperties>
              <t:FieldURI FieldURI="item:Body" />
            </t:AdditionalProperties>
          </m:ItemShape>
          <m:ItemIds>
            <t:ItemId Id="${xmlEscape(templateId)}" />
          </m:ItemIds>
        </m:GetItem>
      `);

      const itemXml = await callEws(token, getTemplateEnvelope, address);
      // Extract the GetItemResponseMessage block first to avoid matching the
      // outer <soap:Body> wrapper before the actual <t:Body> item content
      const responseBlocks = extractBlocks(itemXml, 'GetItemResponseMessage');
      const itemBlock = responseBlocks[0] || itemXml;
      messageText = extractTag(itemBlock, 'Body');
    }

    return ewsResult({
      messageText,
      enabled,
      startTime: startStr ? new Date(startStr) : undefined,
      endTime: endStr ? new Date(endStr) : undefined
    });
  } catch (err) {
    return ewsError(err);
  }
}

export async function setAutoReplyRule(
  token: string,
  messageText: string,
  enabled: boolean,
  startTime?: Date,
  endTime?: Date,
  mailbox?: string
): Promise<OwaResponse<void>> {
  try {
    const address = mailbox || EWS_USERNAME;

    // 1. See if the rule exists and extract the old template ID
    const getRulesEnvelope = soapEnvelope(`
      <m:GetInboxRules>
        <m:MailboxSmtpAddress>${xmlEscape(address)}</m:MailboxSmtpAddress>
      </m:GetInboxRules>
    `);
    const rulesXml = await callEws(token, getRulesEnvelope, address);

    let ruleIdStr = '';
    let oldTemplateId = '';
    const rulesRegex = /<t:Rule>(.*?)<\/t:Rule>/gs;
    let match;
    while ((match = rulesRegex.exec(rulesXml)) !== null) {
      if (match[1].includes('<t:DisplayName>AutoReplyTemplate</t:DisplayName>')) {
        ruleIdStr = extractTag(match[1], 'RuleId');
        oldTemplateId = extractAttribute(match[1], 'ItemId', 'Id');
        break;
      }
    }

    // 2. Create a draft message for the template
    const draftEnvelope = soapEnvelope(`
      <m:CreateItem MessageDisposition="SaveOnly">
        <m:Items>
          <t:Message>
            <t:Subject>AutoReplyTemplate</t:Subject>
            <t:Body BodyType="HTML">${xmlEscape(messageText)}</t:Body>
          </t:Message>
        </m:Items>
      </m:CreateItem>
    `);

    const draftXml = await callEws(token, draftEnvelope, address);
    const templateId = extractAttribute(draftXml, 'ItemId', 'Id');
    const templateChangeKey = extractAttribute(draftXml, 'ItemId', 'ChangeKey');

    if (!templateId) {
      throw new Error('Failed to create template message');
    }

    let deleteOp = '';
    if (ruleIdStr) {
      deleteOp = `
        <t:DeleteRuleOperation>
          <t:RuleId>${xmlEscape(ruleIdStr)}</t:RuleId>
        </t:DeleteRuleOperation>
      `;
    }

    // 4. Create the new rule
    let dateRangeXml = '';
    if (startTime || endTime) {
      dateRangeXml = '<t:WithinDateRange>';
      if (startTime) dateRangeXml += `<t:StartDateTime>${startTime.toISOString()}</t:StartDateTime>`;
      if (endTime) dateRangeXml += `<t:EndDateTime>${endTime.toISOString()}</t:EndDateTime>`;
      dateRangeXml += '</t:WithinDateRange>';
    }

    const conditionsXml = dateRangeXml ? `<t:Conditions>${dateRangeXml}</t:Conditions>` : '';
    const templateChangeKeyAttr = templateChangeKey ? ` ChangeKey="${xmlEscape(templateChangeKey)}"` : '';

    const setRulesEnvelope = soapEnvelope(`
      <m:UpdateInboxRules>
        <m:MailboxSmtpAddress>${xmlEscape(address)}</m:MailboxSmtpAddress>
        <m:RemoveOutlookRuleBlob>false</m:RemoveOutlookRuleBlob>
        <m:Operations>
          ${deleteOp}
          <t:CreateRuleOperation>
            <t:Rule>
              <t:DisplayName>AutoReplyTemplate</t:DisplayName>
              <t:Sequence>1</t:Sequence>
              <t:IsEnabled>${enabled ? 'true' : 'false'}</t:IsEnabled>
              ${conditionsXml}
              <t:Actions>
                <t:ServerReplyWithMessage>
                  <t:ItemId Id="${xmlEscape(templateId)}"${templateChangeKeyAttr} />
                </t:ServerReplyWithMessage>
              </t:Actions>
            </t:Rule>
          </t:CreateRuleOperation>
        </m:Operations>
      </m:UpdateInboxRules>
    `);

    try {
      await callEws(token, setRulesEnvelope, address);
    } catch (err) {
      // Clean up the newly created draft template on failure
      try {
        const deleteTemplateEnvelope = soapEnvelope(`
          <m:DeleteItem DeleteType="HardDelete">
            <m:ItemIds>
              <t:ItemId Id="${xmlEscape(templateId)}" />
            </m:ItemIds>
          </m:DeleteItem>
        `);
        await callEws(token, deleteTemplateEnvelope, address);
      } catch {
        // Ignore cleanup errors
      }
      throw err;
    }

    // 5. Delete the old template draft if it exists (after successful rule update)
    if (oldTemplateId) {
      try {
        const deleteTemplateEnvelope = soapEnvelope(`
          <m:DeleteItem DeleteType="HardDelete">
            <m:ItemIds>
              <t:ItemId Id="${xmlEscape(oldTemplateId)}" />
            </m:ItemIds>
          </m:DeleteItem>
        `);
        await callEws(token, deleteTemplateEnvelope, address);
      } catch (err) {
        // Old template might already be deleted, continue
      }
    }

    return ewsResult(undefined);
  } catch (err) {
    return ewsError(err);
  }
}
