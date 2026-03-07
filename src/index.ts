// Library exports for programmatic usage
export { resolveAuth } from './lib/auth.js';
export type { AuthResult } from './lib/auth.js';

export {
  validateSession, getOwaUserInfo,
  getCalendarEvents, getCalendarEvent, createEvent, updateEvent, deleteEvent, cancelEvent, respondToEvent,
  getEmails, getEmail, sendEmail, replyToEmail, replyToEmailDraft, forwardEmail, updateEmail, moveEmail,
  createDraft, updateDraft, sendDraftById, deleteDraftById, addAttachmentToDraft,
  getMailFolders, createMailFolder, updateMailFolder, deleteMailFolder,
  getAttachments, getAttachment,
  resolveNames, getRoomLists, getRooms, searchRooms,
  getScheduleViaOutlook, getFreeBusy,
} from './lib/ews-client.js';

export type {
  OwaResponse, OwaError, OwaUserInfo,
  CalendarEvent, CalendarAttendee, CreatedEvent, CreateEventOptions, UpdateEventOptions,
  EmailMessage, EmailListResponse, GetEmailsOptions, EmailAttachment,
  MailFolder, MailFolderListResponse,
  Attachment, AttachmentListResponse,
  Room, RoomList, ScheduleInfo, FreeBusySlot,
  Recurrence, RecurrencePattern, RecurrenceRange,
  ResponseType, RespondToEventOptions,
} from './lib/ews-client.js';
