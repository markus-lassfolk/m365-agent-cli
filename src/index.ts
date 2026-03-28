// Library exports for programmatic usage
export { resolveAuth } from './lib/auth.js';
export type { AuthResult } from './lib/auth.js';

export {
  validateSession,
  getOwaUserInfo,
  getCalendarEvents,
  getCalendarEvent,
  createEvent,
  updateEvent,
  deleteEvent,
  cancelEvent,
  respondToEvent,
  getEmails,
  getEmail,
  sendEmail,
  replyToEmail,
  replyToEmailDraft,
  forwardEmail,
  updateEmail,
  moveEmail,
  createDraft,
  updateDraft,
  sendDraftById,
  deleteDraftById,
  addAttachmentToDraft,
  getMailFolders,
  createMailFolder,
  updateMailFolder,
  deleteMailFolder,
  getAttachments,
  getAttachment,
  resolveNames,
  getRoomLists,
  getRooms,
  searchRooms,
  getScheduleViaOutlook,
  getFreeBusy,
  setAutoReplyRule,
  getAutoReplyRule
} from './lib/ews-client.js';

export type {
  OwaResponse,
  OwaError,
  OwaUserInfo,
  CalendarEvent,
  CalendarAttendee,
  CreatedEvent,
  CreateEventOptions,
  UpdateEventOptions,
  EmailMessage,
  EmailListResponse,
  GetEmailsOptions,
  EmailAttachment,
  MailFolder,
  MailFolderListResponse,
  Attachment,
  AttachmentListResponse,
  Room,
  RoomList,
  ScheduleInfo,
  FreeBusySlot,
  Recurrence,
  RecurrencePattern,
  RecurrenceRange,
  ResponseType,
  RespondToEventOptions,
  AutoReplyRule
} from './lib/ews-client.js';

export { resolveGraphAuth } from './lib/graph-auth.js';
export type { GraphAuthResult } from './lib/graph-auth.js';

export {
  listFiles,
  searchFiles,
  getFileMetadata,
  uploadFile,
  createLargeUploadSession,
  downloadFile,
  deleteFile,
  shareFile,
  checkoutFile,
  checkinFile,
  createOfficeCollaborationLink,
  defaultDownloadPath,
  cleanupDownloadedFile
} from './lib/graph-client.js';

export type {
  GraphError,
  GraphResponse,
  DriveItemReference,
  DriveItem,
  DriveItemListResponse,
  SharingLinkResult,
  OfficeCollabLinkResult,
  CheckinResult,
  UploadLargeResult
} from './lib/graph-client.js';

export {
  searchPeople,
  searchUsers,
  searchGroups,
  expandGroup
} from './lib/graph-directory.js';

export type {
  Person,
  User,
  Group
} from './lib/graph-directory.js';

export { getSchedule, findMeetingTimes } from './lib/graph-schedule.js';
export type {
  GetScheduleRequest,
  GetScheduleResponse,
  ScheduleInformation,
  FindMeetingTimesRequest,
  FindMeetingTimesResponse,
  MeetingTimeSuggestion,
  TimeConstraint,
  AttendeeBase
} from './lib/graph-schedule.js';
