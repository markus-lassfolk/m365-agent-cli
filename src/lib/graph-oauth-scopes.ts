/**
 * Single source of truth for Microsoft Graph OAuth scopes used by `login` and `graph-auth` refresh.
 * Documentation: docs/GRAPH_SCOPES.md — keep Entra app registration and setup scripts aligned.
 */

const G = (name: string) => `https://graph.microsoft.com/${name}`;

/** Space-separated scopes for OAuth2 v2.0 device code (`m365-agent-cli login`). */
export const GRAPH_DEVICE_CODE_LOGIN_SCOPES = [
  'offline_access',
  'User.Read',
  'Calendars.ReadWrite',
  'Calendars.Read.Shared',
  'Calendars.ReadWrite.Shared',
  'Mail.Send',
  'Mail.ReadWrite',
  'Mail.Read.Shared',
  'Mail.ReadWrite.Shared',
  'MailboxSettings.ReadWrite',
  'Place.Read.All',
  'People.Read',
  'User.Read.All',
  'Files.ReadWrite.All',
  'Sites.ReadWrite.All',
  'Tasks.ReadWrite',
  'Group.ReadWrite.All',
  'Contacts.ReadWrite',
  'OnlineMeetings.ReadWrite',
  'Notes.ReadWrite.All'
].join(' ');

/** Primary delegated resource scopes (URL form) for refresh_token grant, without `offline_access` / `User.Read`. */
const GRAPH_RESOURCE_SCOPES_FULL = [
  G('Mail.Send'),
  G('Mail.ReadWrite'),
  G('Mail.Read.Shared'),
  G('Mail.ReadWrite.Shared'),
  G('Calendars.ReadWrite'),
  G('Calendars.Read.Shared'),
  G('Calendars.ReadWrite.Shared'),
  G('MailboxSettings.ReadWrite'),
  G('Place.Read.All'),
  G('People.Read'),
  G('User.Read.All'),
  G('Files.ReadWrite.All'),
  G('Sites.ReadWrite.All'),
  G('Tasks.ReadWrite'),
  G('Group.ReadWrite.All'),
  G('Contacts.ReadWrite'),
  G('OnlineMeetings.ReadWrite'),
  G('Notes.ReadWrite.All')
].join(' ');

/**
 * Same as full but omits `User.Read.All` (often requires admin consent) so refresh can succeed with user-only consent.
 */
const GRAPH_RESOURCE_SCOPES_WITHOUT_USER_READ_ALL = [
  G('Mail.Send'),
  G('Mail.ReadWrite'),
  G('Mail.Read.Shared'),
  G('Mail.ReadWrite.Shared'),
  G('Calendars.ReadWrite'),
  G('Calendars.Read.Shared'),
  G('Calendars.ReadWrite.Shared'),
  G('MailboxSettings.ReadWrite'),
  G('Place.Read.All'),
  G('People.Read'),
  G('Files.ReadWrite.All'),
  G('Sites.ReadWrite.All'),
  G('Tasks.ReadWrite'),
  G('Group.ReadWrite.All'),
  G('Contacts.ReadWrite'),
  G('OnlineMeetings.ReadWrite'),
  G('Notes.ReadWrite.All')
].join(' ');

/**
 * Ordered candidates for Graph refresh_token exchange. Earlier entries preferred; later entries are fallbacks
 * (e.g. Files-only) so a stale refresh token can still produce some access token.
 */
export const GRAPH_REFRESH_SCOPE_CANDIDATES: readonly string[] = [
  `${G('.default')} offline_access`,
  `${GRAPH_RESOURCE_SCOPES_FULL} offline_access User.Read`,
  `${GRAPH_RESOURCE_SCOPES_WITHOUT_USER_READ_ALL} offline_access User.Read`,
  `${G('Mail.Send')} ${G('Mail.ReadWrite')} ${G('Mail.Read.Shared')} ${G('Mail.ReadWrite.Shared')} ${G('Calendars.ReadWrite')} ${G('Calendars.Read.Shared')} ${G('Calendars.ReadWrite.Shared')} ${G('Files.ReadWrite')} offline_access User.Read`,
  `${G('Files.ReadWrite')} offline_access User.Read`,
  `${G('Files.ReadWrite.All')} offline_access User.Read`,
  `${G('Sites.ReadWrite.All')} offline_access User.Read`,
  `${G('Tasks.ReadWrite')} offline_access User.Read`,
  `${G('Group.ReadWrite.All')} offline_access User.Read`,
  `${G('Contacts.ReadWrite')} ${G('OnlineMeetings.ReadWrite')} ${G('Notes.ReadWrite.All')} offline_access User.Read`,
  `${G('Files.Read')} offline_access User.Read`
];
