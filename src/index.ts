// Library exports for programmatic usage

export type { AuthResult } from './lib/auth.js';
export { resolveAuth } from './lib/auth.js';
export type {
  AddDelegateOptions,
  DelegateFolderPermissionLevel,
  DelegateInfo,
  DelegatePermissions,
  DeliverMeetingRequests,
  RemoveDelegateOptions,
  UpdateDelegateOptions
} from './lib/delegate-client.js';
// Delegate management
export {
  addDelegate,
  getDelegates,
  removeDelegate,
  updateDelegate
} from './lib/delegate-client.js';
export type {
  Attachment,
  AttachmentListResponse,
  AutoReplyRule,
  CalendarAttendee,
  CalendarEvent,
  CreatedEvent,
  CreateEventOptions,
  EmailAttachment,
  EmailListResponse,
  EmailMessage,
  FreeBusySlot,
  GetEmailsOptions,
  MailFolder,
  MailFolderListResponse,
  OwaError,
  OwaResponse,
  OwaUserInfo,
  Recurrence,
  RecurrencePattern,
  RecurrenceRange,
  RespondToEventOptions,
  ResponseType,
  Room,
  RoomList,
  ScheduleInfo,
  UpdateEventOptions
} from './lib/ews-client.js';
export {
  addAttachmentToDraft,
  cancelEvent,
  createDraft,
  createEvent,
  createMailFolder,
  deleteDraftById,
  deleteEvent,
  deleteMailFolder,
  forwardEmail,
  getAttachment,
  getAttachments,
  getAutoReplyRule,
  getCalendarEvent,
  getCalendarEvents,
  getEmail,
  getEmails,
  getFreeBusy,
  getMailFolders,
  getMyFreeBusySlots,
  getOwaUserInfo,
  getRoomLists,
  getRooms,
  getScheduleViaOutlook,
  moveEmail,
  replyToEmail,
  replyToEmailDraft,
  resolveNames,
  respondToEvent,
  searchRooms,
  sendDraftById,
  sendEmail,
  setAutoReplyRule,
  updateDraft,
  updateEmail,
  updateEvent,
  updateMailFolder,
  validateSession
} from './lib/ews-client.js';
export type { GraphAuthResult } from './lib/graph-auth.js';
export { resolveGraphAuth } from './lib/graph-auth.js';
export type {
  CheckinResult,
  DriveItem,
  DriveItemListResponse,
  DriveItemReference,
  GraphError,
  GraphResponse,
  OfficeCollabLinkResult,
  SharingLinkResult,
  UploadLargeResult
} from './lib/graph-client.js';
export {
  checkinFile,
  checkoutFile,
  cleanupDownloadedFile,
  createLargeUploadSession,
  createOfficeCollaborationLink,
  defaultDownloadPath,
  deleteFile,
  downloadFile,
  getFileMetadata,
  listFiles,
  searchFiles,
  shareFile,
  uploadFile
} from './lib/graph-client.js';
export type {
  Group,
  Person,
  User
} from './lib/graph-directory.js';
export {
  expandGroup,
  searchGroups,
  searchPeople,
  searchUsers
} from './lib/graph-directory.js';
export type {
  AttendeeBase,
  FindMeetingTimesRequest,
  FindMeetingTimesResponse,
  GetScheduleRequest,
  GetScheduleResponse,
  MeetingTimeSuggestion,
  ScheduleInformation,
  TimeConstraint
} from './lib/graph-schedule.js';
export { findMeetingTimes, getSchedule } from './lib/graph-schedule.js';
export type { Subscription } from './lib/graph-subscriptions.js';
export {
  createSubscription,
  deleteSubscription,
  listSubscriptions,
  renewSubscription
} from './lib/graph-subscriptions.js';
export type { AutomaticRepliesSetting, MailboxSettings, OofStatus } from './lib/oof-client.js';
export { getMailboxSettings, setMailboxSettings } from './lib/oof-client.js';
export type { Place as PlaceRoom, RoomList as PlaceRoomList } from './lib/places-client.js';
// places-client re-exports (aliases from EWS: getRoomLists/getRooms conflict)
export { listPlaceRoomLists as getPlaceRoomLists, listRoomsInRoomList as getPlaceRooms } from './lib/places-client.js';
export type {
  CreateMessageRulePayload,
  MessageRule,
  MessageRuleAction,
  MessageRuleCondition,
  UpdateMessageRulePayload
} from './lib/rules-client.js';
// Inbox rules
export {
  createMessageRule,
  deleteMessageRule,
  getMessageRule,
  listMessageRules,
  updateMessageRule
} from './lib/rules-client.js';
export type {
  CreateTaskOptions,
  TodoChecklistItem,
  TodoImportance,
  TodoLinkedResource,
  TodoList,
  TodoStatus,
  TodoTask,
  UpdateTaskOptions
} from './lib/todo-client.js';
// To-Do
export {
  addChecklistItem,
  createTask,
  deleteChecklistItem,
  deleteTask,
  getTask,
  getTasks,
  getTodoList,
  getTodoLists,
  updateTask
} from './lib/todo-client.js';
