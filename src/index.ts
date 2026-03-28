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

export {
  createSubscription,
  listSubscriptions,
  deleteSubscription,
  renewSubscription
} from './lib/graph-subscriptions.js';

export type { Subscription } from './lib/graph-subscriptions.js';

// places-client re-exports (aliases from EWS: getRoomLists/getRooms conflict)
export { listPlaceRoomLists as getPlaceRoomLists, listRoomsInRoomList as getPlaceRooms } from './lib/places-client.js';
export type { Place as PlaceRoom, RoomList as PlaceRoomList } from './lib/places-client.js';

export { getMailboxSettings, setMailboxSettings } from './lib/oof-client.js';
export type { OofStatus, AutomaticRepliesSetting, MailboxSettings } from './lib/oof-client.js';

// Delegate management
export {
  getDelegates,
  addDelegate,
  updateDelegate,
  removeDelegate
} from './lib/delegate-client.js';

export type {
  DelegateInfo,
  DelegatePermissions,
  DelegateFolderPermissionLevel,
  DeliverMeetingRequests,
  AddDelegateOptions,
  UpdateDelegateOptions,
  RemoveDelegateOptions
} from './lib/delegate-client.js';

// Inbox rules
export {
  listMessageRules,
  getMessageRule,
  createMessageRule,
  updateMessageRule,
  deleteMessageRule
} from './lib/rules-client.js';

export type {
  MessageRule,
  MessageRuleCondition,
  MessageRuleAction,
  CreateMessageRulePayload,
  UpdateMessageRulePayload
} from './lib/rules-client.js';

// To-Do
export {
  getTodoLists,
  getTodoList,
  getTasks,
  getTask,
  createTask,
  updateTask,
  deleteTask,
  addChecklistItem,
  deleteChecklistItem
} from './lib/todo-client.js';

export type {
  TodoImportance,
  TodoStatus,
  TodoLinkedResource,
  TodoChecklistItem,
  TodoTask,
  TodoList,
  CreateTaskOptions,
  UpdateTaskOptions
} from './lib/todo-client.js';
