import {
  callGraph,
  callGraphAt,
  GraphApiError,
  type GraphResponse,
  graphError,
  graphResult,
  listGraphCollection
} from './graph-client.js';
import { getGraphBaseUrl, getGraphBetaUrl } from './graph-constants.js';
import { graphUserPath } from './graph-user-path.js';

export interface GraphTeam {
  id: string;
  displayName?: string;
  description?: string;
}

export interface GraphChannel {
  id: string;
  displayName?: string;
  description?: string;
  membershipType?: string;
  tenantId?: string;
}

export interface GraphChatMessage {
  id?: string;
  createdDateTime?: string;
  body?: { content?: string; contentType?: string };
  from?: { user?: { displayName?: string; id?: string } };
}

export interface PinnedChatMessageInfo {
  id?: string;
  message?: GraphChatMessage;
}

export interface GraphTeamMember {
  id?: string;
  displayName?: string;
  email?: string;
  roles?: string[];
  userId?: string;
}

export interface GraphChat {
  id: string;
  topic?: string;
  chatType?: string;
  createdDateTime?: string;
}

export interface GraphChatDetail extends GraphChat {
  lastUpdatedDateTime?: string;
  webUrl?: string;
  tenantId?: string;
  members?: GraphTeamMember[];
  lastMessagePreview?: unknown;
}

export interface TeamsChannelTab {
  id?: string;
  displayName?: string;
  webUrl?: string;
  sortOrderIndex?: string;
  teamsApp?: { id?: string; displayName?: string; distributionMethod?: string };
}

export interface TeamsAppInstallation {
  id?: string;
  teamsAppDefinition?: {
    displayName?: string;
    teamsAppId?: string;
    version?: string;
  };
}

/** Entry in `GET /appCatalogs/teamsApps` (catalog id ≠ manifest id unless store distribution). */
export interface TeamsAppCatalogEntry {
  id?: string;
  externalId?: string;
  displayName?: string;
  distributionMethod?: string;
  appDefinitions?: unknown[];
  [key: string]: unknown;
}

/** OData bind URL for `POST …/installedApps` bodies (`teamsApp@odata.bind`). */
export function teamsAppCatalogBindUrl(catalogTeamsAppId: string): string {
  const base = getGraphBaseUrl().replace(/\/+$/, '');
  return `${base}/appCatalogs/teamsApps/${encodeURIComponent(catalogTeamsAppId.trim())}`;
}

export function buildAddTeamsAppInstallationBody(catalogTeamsAppId: string): Record<string, string> {
  return { 'teamsApp@odata.bind': teamsAppCatalogBindUrl(catalogTeamsAppId) };
}

/**
 * List joined teams for `/me` or another user (`GET /users/{id}/joinedTeams` when `forUser` is set).
 * Requires permission to read that user’s team memberships (e.g. Team.ReadBasic.All delegated).
 */
export async function listJoinedTeams(token: string, forUser?: string): Promise<GraphResponse<GraphTeam[]>> {
  return listGraphCollection<GraphTeam>(token, graphUserPath(forUser, 'joinedTeams'), 'Failed to list joined teams');
}

export async function getTeam(token: string, teamId: string): Promise<GraphResponse<GraphTeam>> {
  try {
    const r = await callGraph<GraphTeam>(token, `/teams/${encodeURIComponent(teamId)}`);
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to get team', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get team');
  }
}

/** Default “General” channel. Least privilege: Channel.ReadBasic.All */
export async function getTeamPrimaryChannel(token: string, teamId: string): Promise<GraphResponse<GraphChannel>> {
  try {
    const r = await callGraph<GraphChannel>(token, `/teams/${encodeURIComponent(teamId.trim())}/primaryChannel`);
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to get primary channel', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get primary channel');
  }
}

/** Root drive item for the channel “Files” tab (`GET …/channels/{id}/filesFolder`). Use with `files --drive-id` + folder id. */
export interface ChannelFilesFolderItem {
  id?: string;
  name?: string;
  webUrl?: string;
  parentReference?: { driveId?: string; id?: string; path?: string };
  [key: string]: unknown;
}

export async function getTeamChannelFilesFolder(
  token: string,
  teamId: string,
  channelId: string
): Promise<GraphResponse<ChannelFilesFolderItem>> {
  try {
    const tid = encodeURIComponent(teamId.trim());
    const cid = encodeURIComponent(channelId.trim());
    const r = await callGraph<ChannelFilesFolderItem>(token, `/teams/${tid}/channels/${cid}/filesFolder`);
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to get channel files folder', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get channel files folder');
  }
}

export async function listTeamChannels(token: string, teamId: string): Promise<GraphResponse<GraphChannel[]>> {
  return listGraphCollection<GraphChannel>(
    token,
    `/teams/${encodeURIComponent(teamId)}/channels`,
    'Failed to list channels'
  );
}

/** Includes shared / incoming channels. Least privilege: Channel.ReadBasic.All. Optional OData `$filter` e.g. `membershipType eq 'shared'`. */
export async function listTeamAllChannels(
  token: string,
  teamId: string,
  filter?: string
): Promise<GraphResponse<GraphChannel[]>> {
  const q = filter?.trim() ? `?$filter=${encodeURIComponent(filter.trim())}` : '';
  return listGraphCollection<GraphChannel>(
    token,
    `/teams/${encodeURIComponent(teamId.trim())}/allChannels${q}`,
    'Failed to list all channels'
  );
}

/** Channels shared into this team from other tenants. GET /teams/{id}/incomingChannels (Channel.ReadBasic.All). */
export async function listTeamIncomingChannels(token: string, teamId: string): Promise<GraphResponse<GraphChannel[]>> {
  return listGraphCollection<GraphChannel>(
    token,
    `/teams/${encodeURIComponent(teamId.trim())}/incomingChannels`,
    'Failed to list incoming channels'
  );
}

/** Single channel. Least privilege: Channel.ReadBasic.All */
export async function getTeamChannel(
  token: string,
  teamId: string,
  channelId: string
): Promise<GraphResponse<GraphChannel>> {
  try {
    const t = encodeURIComponent(teamId.trim());
    const c = encodeURIComponent(channelId.trim());
    const r = await callGraph<GraphChannel>(token, `/teams/${t}/channels/${c}`);
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to get channel', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get channel');
  }
}

/**
 * Channel roster (not team-wide members). Least privilege: ChannelMember.Read.All;
 * Group.ReadWrite.All also accepted (Graph compat).
 */
export async function listTeamChannelMembers(
  token: string,
  teamId: string,
  channelId: string,
  top?: number
): Promise<GraphResponse<GraphTeamMember[]>> {
  const t = encodeURIComponent(teamId.trim());
  const c = encodeURIComponent(channelId.trim());
  return listGraphCollection<GraphTeamMember>(
    token,
    `/teams/${t}/channels/${c}/members`,
    'Failed to list channel members',
    {
      top,
      maxTop: 999
    }
  );
}

/** Reply in a channel thread. Same permission as `sendChannelMessage`. */
export async function sendChannelMessageReply(
  token: string,
  teamId: string,
  channelId: string,
  parentMessageId: string,
  body: Record<string, unknown>
): Promise<GraphResponse<GraphChatMessage>> {
  try {
    const t = encodeURIComponent(teamId.trim());
    const c = encodeURIComponent(channelId.trim());
    const m = encodeURIComponent(parentMessageId.trim());
    const r = await callGraph<GraphChatMessage>(token, `/teams/${t}/channels/${c}/messages/${m}/replies`, {
      method: 'POST',
      body: JSON.stringify(body)
    });
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to send reply', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to send reply');
  }
}

/** `body` is a full `chatMessage` resource (or minimal `{ body: { contentType, content } }`). Requires `ChannelMessage.Send`. */
export async function sendChannelMessage(
  token: string,
  teamId: string,
  channelId: string,
  body: Record<string, unknown>
): Promise<GraphResponse<GraphChatMessage>> {
  try {
    const t = encodeURIComponent(teamId.trim());
    const c = encodeURIComponent(channelId.trim());
    const r = await callGraph<GraphChatMessage>(token, `/teams/${t}/channels/${c}/messages`, {
      method: 'POST',
      body: JSON.stringify(body)
    });
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to send channel message', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to send channel message');
  }
}

/** Requires `Chat.ReadWrite` (or narrower send scope where applicable). */
export async function sendChatMessage(
  token: string,
  chatId: string,
  body: Record<string, unknown>
): Promise<GraphResponse<GraphChatMessage>> {
  try {
    const r = await callGraph<GraphChatMessage>(token, `/chats/${encodeURIComponent(chatId.trim())}/messages`, {
      method: 'POST',
      body: JSON.stringify(body)
    });
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to send chat message', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to send chat message');
  }
}

export async function listChannelMessages(
  token: string,
  teamId: string,
  channelId: string,
  top?: number
): Promise<GraphResponse<GraphChatMessage[]>> {
  try {
    const qs = top && top > 0 ? `?$top=${Math.min(top, 50)}` : '';
    const r = await callGraph<{ value: GraphChatMessage[] }>(
      token,
      `/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}/messages${qs}`
    );
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to list channel messages', r.error?.code, r.error?.status);
    }
    return graphResult(r.data.value ?? []);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to list channel messages');
  }
}

export async function listChannelMessageReplies(
  token: string,
  teamId: string,
  channelId: string,
  messageId: string,
  top?: number
): Promise<GraphResponse<GraphChatMessage[]>> {
  const t = encodeURIComponent(teamId.trim());
  const c = encodeURIComponent(channelId.trim());
  const m = encodeURIComponent(messageId.trim());
  return listGraphCollection<GraphChatMessage>(
    token,
    `/teams/${t}/channels/${c}/messages/${m}/replies`,
    'Failed to list message replies',
    { top, maxTop: 50 }
  );
}

/** Single channel message. `ChannelMessage.Read.All` (or compatible). */
export async function getChannelMessage(
  token: string,
  teamId: string,
  channelId: string,
  messageId: string
): Promise<GraphResponse<GraphChatMessage>> {
  try {
    const t = encodeURIComponent(teamId.trim());
    const c = encodeURIComponent(channelId.trim());
    const m = encodeURIComponent(messageId.trim());
    const r = await callGraph<GraphChatMessage>(token, `/teams/${t}/channels/${c}/messages/${m}`);
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to get channel message', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get channel message');
  }
}

/** Single chat message. `Chat.ReadWrite` (or `Chat.Read` where granted). */
export async function getChatMessage(
  token: string,
  chatId: string,
  messageId: string
): Promise<GraphResponse<GraphChatMessage>> {
  try {
    const c = encodeURIComponent(chatId.trim());
    const m = encodeURIComponent(messageId.trim());
    const r = await callGraph<GraphChatMessage>(token, `/chats/${c}/messages/${m}`);
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to get chat message', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get chat message');
  }
}

export async function listChatMessageReplies(
  token: string,
  chatId: string,
  messageId: string,
  top?: number
): Promise<GraphResponse<GraphChatMessage[]>> {
  const c = encodeURIComponent(chatId.trim());
  const m = encodeURIComponent(messageId.trim());
  return listGraphCollection<GraphChatMessage>(
    token,
    `/chats/${c}/messages/${m}/replies`,
    'Failed to list chat replies',
    {
      top,
      maxTop: 50
    }
  );
}

/** Reply in a 1:1 or group chat thread. `Chat.ReadWrite`. */
export async function sendChatMessageReply(
  token: string,
  chatId: string,
  parentMessageId: string,
  body: Record<string, unknown>
): Promise<GraphResponse<GraphChatMessage>> {
  try {
    const c = encodeURIComponent(chatId.trim());
    const m = encodeURIComponent(parentMessageId.trim());
    const r = await callGraph<GraphChatMessage>(token, `/chats/${c}/messages/${m}/replies`, {
      method: 'POST',
      body: JSON.stringify(body)
    });
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to send chat reply', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to send chat reply');
  }
}

export async function listTeamMembers(
  token: string,
  teamId: string,
  top?: number
): Promise<GraphResponse<GraphTeamMember[]>> {
  return listGraphCollection<GraphTeamMember>(
    token,
    `/teams/${encodeURIComponent(teamId)}/members`,
    'Failed to list team members',
    { top, maxTop: 999 }
  );
}

/**
 * `expand` e.g. `members`, `lastMessagePreview`, or `members,lastMessagePreview` (Graph `$expand`).
 * Least privilege: Chat.ReadBasic without expand; Chat.Read for richer expands where required.
 */
export async function getChat(token: string, chatId: string, expand?: string): Promise<GraphResponse<GraphChatDetail>> {
  try {
    const enc = encodeURIComponent(chatId.trim());
    const q = expand?.trim() ? `?$expand=${encodeURIComponent(expand.trim())}` : '';
    const r = await callGraph<GraphChatDetail>(token, `/chats/${enc}${q}`);
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to get chat', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get chat');
  }
}

export async function listMyChats(token: string, top?: number): Promise<GraphResponse<GraphChat[]>> {
  return listGraphCollection<GraphChat>(token, `/me/chats`, 'Failed to list chats', { top, maxTop: 50 });
}

export async function listChatMessages(
  token: string,
  chatId: string,
  top?: number
): Promise<GraphResponse<GraphChatMessage[]>> {
  try {
    const t = top && top > 0 ? Math.min(top, 50) : undefined;
    const qs = t ? `?$top=${t}` : '';
    const r = await callGraph<{ value: GraphChatMessage[] }>(
      token,
      `/chats/${encodeURIComponent(chatId)}/messages${qs}`
    );
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to list chat messages', r.error?.code, r.error?.status);
    }
    return graphResult(r.data.value ?? []);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to list chat messages');
  }
}

/** Optional `$expand=message` for full chatMessage bodies. Requires Chat.Read */
export async function listChatPinnedMessages(
  token: string,
  chatId: string,
  expandMessage?: boolean
): Promise<GraphResponse<PinnedChatMessageInfo[]>> {
  try {
    const enc = encodeURIComponent(chatId.trim());
    const q = expandMessage ? '?$expand=message' : '';
    const r = await callGraph<{ value: PinnedChatMessageInfo[] }>(token, `/chats/${enc}/pinnedMessages${q}`);
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to list pinned messages', r.error?.code, r.error?.status);
    }
    return graphResult(r.data.value ?? []);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to list pinned messages');
  }
}

export async function listChatMembers(token: string, chatId: string): Promise<GraphResponse<GraphTeamMember[]>> {
  return listGraphCollection<GraphTeamMember>(
    token,
    `/chats/${encodeURIComponent(chatId)}/members`,
    'Failed to list chat members'
  );
}

/** Uses `$expand=teamsAppDefinition` for display names. `Group.ReadWrite.All` satisfies Graph (see team-list-installedapps). */
export async function listTeamInstalledApps(
  token: string,
  teamId: string
): Promise<GraphResponse<TeamsAppInstallation[]>> {
  return listGraphCollection<TeamsAppInstallation>(
    token,
    `/teams/${encodeURIComponent(teamId)}/installedApps?$expand=teamsAppDefinition`,
    'Failed to list installed apps'
  );
}

export async function listTeamsAppCatalog(
  token: string,
  filter?: string,
  expand?: string
): Promise<GraphResponse<TeamsAppCatalogEntry[]>> {
  const q: string[] = [];
  if (filter?.trim()) q.push(`$filter=${encodeURIComponent(filter.trim())}`);
  if (expand?.trim()) q.push(`$expand=${encodeURIComponent(expand.trim())}`);
  const suffix = q.length > 0 ? `?${q.join('&')}` : '';
  return listGraphCollection<TeamsAppCatalogEntry>(
    token,
    `/appCatalogs/teamsApps${suffix}`,
    'Failed to list app catalog',
    {}
  );
}

export async function getTeamsAppCatalogEntry(
  token: string,
  catalogTeamsAppId: string,
  expand?: string
): Promise<GraphResponse<TeamsAppCatalogEntry>> {
  try {
    const id = encodeURIComponent(catalogTeamsAppId.trim());
    const q = expand?.trim() ? `?$expand=${encodeURIComponent(expand.trim())}` : '';
    const r = await callGraph<TeamsAppCatalogEntry>(token, `/appCatalogs/teamsApps/${id}${q}`);
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to get catalog app', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get catalog app');
  }
}

export async function getTeamInstalledApp(
  token: string,
  teamId: string,
  installationId: string
): Promise<GraphResponse<TeamsAppInstallation>> {
  try {
    const t = encodeURIComponent(teamId.trim());
    const i = encodeURIComponent(installationId.trim());
    const r = await callGraph<TeamsAppInstallation>(token, `/teams/${t}/installedApps/${i}?$expand=teamsAppDefinition`);
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to get installed app', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get installed app');
  }
}

export async function addTeamInstalledApp(
  token: string,
  teamId: string,
  body: Record<string, unknown>
): Promise<GraphResponse<void>> {
  try {
    const t = encodeURIComponent(teamId.trim());
    const r = await callGraph<void>(
      token,
      `/teams/${t}/installedApps`,
      {
        method: 'POST',
        body: JSON.stringify(body)
      },
      false
    );
    if (!r.ok) {
      return graphError(r.error?.message || 'Failed to install app on team', r.error?.code, r.error?.status);
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to install app on team');
  }
}

export async function patchTeamInstalledApp(
  token: string,
  teamId: string,
  installationId: string,
  body: Record<string, unknown>
): Promise<GraphResponse<TeamsAppInstallation>> {
  try {
    const t = encodeURIComponent(teamId.trim());
    const i = encodeURIComponent(installationId.trim());
    const r = await callGraph<TeamsAppInstallation>(token, `/teams/${t}/installedApps/${i}`, {
      method: 'PATCH',
      body: JSON.stringify(body)
    });
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to patch installed app', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to patch installed app');
  }
}

export async function deleteTeamInstalledApp(
  token: string,
  teamId: string,
  installationId: string,
  ifMatch?: string
): Promise<GraphResponse<void>> {
  try {
    const t = encodeURIComponent(teamId.trim());
    const i = encodeURIComponent(installationId.trim());
    const headers: Record<string, string> = {};
    if (ifMatch?.trim()) headers['If-Match'] = ifMatch.trim();
    const r = await callGraph<void>(token, `/teams/${t}/installedApps/${i}`, { method: 'DELETE', headers }, false);
    if (!r.ok) {
      return graphError(r.error?.message || 'Failed to remove app from team', r.error?.code, r.error?.status);
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to remove app from team');
  }
}

export async function upgradeTeamInstalledApp(
  token: string,
  teamId: string,
  installationId: string
): Promise<GraphResponse<void>> {
  try {
    const t = encodeURIComponent(teamId.trim());
    const i = encodeURIComponent(installationId.trim());
    const r = await callGraph<void>(
      token,
      `/teams/${t}/installedApps/${i}/upgrade`,
      { method: 'POST', body: '{}' },
      false
    );
    if (!r.ok) {
      return graphError(r.error?.message || 'Failed to upgrade installed app', r.error?.code, r.error?.status);
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to upgrade installed app');
  }
}

export async function listChatInstalledApps(
  token: string,
  chatId: string
): Promise<GraphResponse<TeamsAppInstallation[]>> {
  const c = encodeURIComponent(chatId.trim());
  return listGraphCollection<TeamsAppInstallation>(
    token,
    `/chats/${c}/installedApps?$expand=teamsAppDefinition`,
    'Failed to list chat apps'
  );
}

export async function getChatInstalledApp(
  token: string,
  chatId: string,
  installationId: string
): Promise<GraphResponse<TeamsAppInstallation>> {
  try {
    const c = encodeURIComponent(chatId.trim());
    const i = encodeURIComponent(installationId.trim());
    const r = await callGraph<TeamsAppInstallation>(token, `/chats/${c}/installedApps/${i}?$expand=teamsAppDefinition`);
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to get chat app', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get chat app');
  }
}

export async function addChatInstalledApp(
  token: string,
  chatId: string,
  body: Record<string, unknown>
): Promise<GraphResponse<void>> {
  try {
    const c = encodeURIComponent(chatId.trim());
    const r = await callGraph<void>(
      token,
      `/chats/${c}/installedApps`,
      {
        method: 'POST',
        body: JSON.stringify(body)
      },
      false
    );
    if (!r.ok) {
      return graphError(r.error?.message || 'Failed to install app on chat', r.error?.code, r.error?.status);
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to install app on chat');
  }
}

export async function patchChatInstalledApp(
  token: string,
  chatId: string,
  installationId: string,
  body: Record<string, unknown>
): Promise<GraphResponse<TeamsAppInstallation>> {
  try {
    const c = encodeURIComponent(chatId.trim());
    const i = encodeURIComponent(installationId.trim());
    const r = await callGraph<TeamsAppInstallation>(token, `/chats/${c}/installedApps/${i}`, {
      method: 'PATCH',
      body: JSON.stringify(body)
    });
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to patch chat app', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to patch chat app');
  }
}

export async function deleteChatInstalledApp(
  token: string,
  chatId: string,
  installationId: string,
  ifMatch?: string
): Promise<GraphResponse<void>> {
  try {
    const c = encodeURIComponent(chatId.trim());
    const i = encodeURIComponent(installationId.trim());
    const headers: Record<string, string> = {};
    if (ifMatch?.trim()) headers['If-Match'] = ifMatch.trim();
    const r = await callGraph<void>(token, `/chats/${c}/installedApps/${i}`, { method: 'DELETE', headers }, false);
    if (!r.ok) {
      return graphError(r.error?.message || 'Failed to remove app from chat', r.error?.code, r.error?.status);
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to remove app from chat');
  }
}

export async function upgradeChatInstalledApp(
  token: string,
  chatId: string,
  installationId: string
): Promise<GraphResponse<void>> {
  try {
    const c = encodeURIComponent(chatId.trim());
    const i = encodeURIComponent(installationId.trim());
    const r = await callGraph<void>(
      token,
      `/chats/${c}/installedApps/${i}/upgrade`,
      { method: 'POST', body: '{}' },
      false
    );
    if (!r.ok) {
      return graphError(r.error?.message || 'Failed to upgrade chat app', r.error?.code, r.error?.status);
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to upgrade chat app');
  }
}

/** Personal-scope apps: `GET …/teamwork/installedApps` for `/me` or `/users/{id}/…`. */
export async function listUserTeamworkInstalledApps(
  token: string,
  forUser?: string
): Promise<GraphResponse<TeamsAppInstallation[]>> {
  return listGraphCollection<TeamsAppInstallation>(
    token,
    `${graphUserPath(forUser, 'teamwork/installedApps')}?$expand=teamsAppDefinition`,
    'Failed to list user teamwork apps'
  );
}

export async function getUserTeamworkInstalledApp(
  token: string,
  installationId: string,
  forUser?: string
): Promise<GraphResponse<TeamsAppInstallation>> {
  try {
    const i = encodeURIComponent(installationId.trim());
    const path = `${graphUserPath(forUser, `teamwork/installedApps/${i}`)}?$expand=teamsAppDefinition`;
    const r = await callGraph<TeamsAppInstallation>(token, path);
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to get user teamwork app', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get user teamwork app');
  }
}

export async function addUserTeamworkInstalledApp(
  token: string,
  body: Record<string, unknown>,
  forUser?: string
): Promise<GraphResponse<void>> {
  try {
    const path = graphUserPath(forUser, 'teamwork/installedApps');
    const r = await callGraph<void>(token, path, { method: 'POST', body: JSON.stringify(body) }, false);
    if (!r.ok) {
      return graphError(r.error?.message || 'Failed to install user teamwork app', r.error?.code, r.error?.status);
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to install user teamwork app');
  }
}

export async function deleteUserTeamworkInstalledApp(
  token: string,
  installationId: string,
  forUser?: string,
  ifMatch?: string
): Promise<GraphResponse<void>> {
  try {
    const i = encodeURIComponent(installationId.trim());
    const headers: Record<string, string> = {};
    if (ifMatch?.trim()) headers['If-Match'] = ifMatch.trim();
    const path = graphUserPath(forUser, `teamwork/installedApps/${i}`);
    const r = await callGraph<void>(token, path, { method: 'DELETE', headers }, false);
    if (!r.ok) {
      return graphError(r.error?.message || 'Failed to remove user teamwork app', r.error?.code, r.error?.status);
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to remove user teamwork app');
  }
}

/** `$expand=teamsApp` for app names. Least privilege: TeamsTab.Read.All; Group.ReadWrite.All also allowed (compat). */
export async function listChannelTabs(
  token: string,
  teamId: string,
  channelId: string
): Promise<GraphResponse<TeamsChannelTab[]>> {
  try {
    const t = encodeURIComponent(teamId.trim());
    const c = encodeURIComponent(channelId.trim());
    const path = `/teams/${t}/channels/${c}/tabs?$expand=teamsApp`;
    const r = await callGraph<{ value: TeamsChannelTab[] }>(token, path);
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to list channel tabs', r.error?.code, r.error?.status);
    }
    return graphResult(r.data.value ?? []);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to list channel tabs');
  }
}

/** Unicode emoji or Teams-supported reaction string. Returns 204 No Content. */
export async function setChannelMessageReaction(
  token: string,
  teamId: string,
  channelId: string,
  messageId: string,
  reactionType: string,
  replyId?: string
): Promise<GraphResponse<void>> {
  try {
    const t = encodeURIComponent(teamId.trim());
    const c = encodeURIComponent(channelId.trim());
    const m = encodeURIComponent(messageId.trim());
    const path = replyId?.trim()
      ? `/teams/${t}/channels/${c}/messages/${m}/replies/${encodeURIComponent(replyId.trim())}/setReaction`
      : `/teams/${t}/channels/${c}/messages/${m}/setReaction`;
    const r = await callGraph<void>(
      token,
      path,
      { method: 'POST', body: JSON.stringify({ reactionType: reactionType.trim() }) },
      false
    );
    if (!r.ok) {
      return graphError(r.error?.message || 'Failed to set reaction', r.error?.code, r.error?.status);
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to set reaction');
  }
}

export async function unsetChannelMessageReaction(
  token: string,
  teamId: string,
  channelId: string,
  messageId: string,
  reactionType: string,
  replyId?: string
): Promise<GraphResponse<void>> {
  try {
    const t = encodeURIComponent(teamId.trim());
    const c = encodeURIComponent(channelId.trim());
    const m = encodeURIComponent(messageId.trim());
    const path = replyId?.trim()
      ? `/teams/${t}/channels/${c}/messages/${m}/replies/${encodeURIComponent(replyId.trim())}/unsetReaction`
      : `/teams/${t}/channels/${c}/messages/${m}/unsetReaction`;
    const r = await callGraph<void>(
      token,
      path,
      { method: 'POST', body: JSON.stringify({ reactionType: reactionType.trim() }) },
      false
    );
    if (!r.ok) {
      return graphError(r.error?.message || 'Failed to unset reaction', r.error?.code, r.error?.status);
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to unset reaction');
  }
}

export async function setChatMessageReaction(
  token: string,
  chatId: string,
  messageId: string,
  reactionType: string
): Promise<GraphResponse<void>> {
  try {
    const c = encodeURIComponent(chatId.trim());
    const m = encodeURIComponent(messageId.trim());
    const r = await callGraph<void>(
      token,
      `/chats/${c}/messages/${m}/setReaction`,
      { method: 'POST', body: JSON.stringify({ reactionType: reactionType.trim() }) },
      false
    );
    if (!r.ok) {
      return graphError(r.error?.message || 'Failed to set reaction', r.error?.code, r.error?.status);
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to set reaction');
  }
}

export async function unsetChatMessageReaction(
  token: string,
  chatId: string,
  messageId: string,
  reactionType: string
): Promise<GraphResponse<void>> {
  try {
    const c = encodeURIComponent(chatId.trim());
    const m = encodeURIComponent(messageId.trim());
    const r = await callGraph<void>(
      token,
      `/chats/${c}/messages/${m}/unsetReaction`,
      { method: 'POST', body: JSON.stringify({ reactionType: reactionType.trim() }) },
      false
    );
    if (!r.ok) {
      return graphError(r.error?.message || 'Failed to unset reaction', r.error?.code, r.error?.status);
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to unset reaction');
  }
}

/** `POST /me/teamwork/sendActivityNotification` — requires `TeamsActivity.Send`; returns 204. */
export async function sendMeTeamworkActivityNotification(
  token: string,
  body: Record<string, unknown>
): Promise<GraphResponse<void>> {
  try {
    const r = await callGraph<void>(
      token,
      '/me/teamwork/sendActivityNotification',
      {
        method: 'POST',
        body: JSON.stringify(body)
      },
      false
    );
    if (!r.ok) {
      return graphError(r.error?.message || 'Failed to send activity notification', r.error?.code, r.error?.status);
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to send activity notification');
  }
}

/** `POST /users/{id}/teamwork/sendActivityNotification` — typically **application** permission `TeamsActivity.Send`; returns 204. */
export async function sendUserTeamworkActivityNotification(
  token: string,
  userId: string,
  body: Record<string, unknown>
): Promise<GraphResponse<void>> {
  const u = encodeURIComponent(userId.trim());
  try {
    const r = await callGraph<void>(
      token,
      `/users/${u}/teamwork/sendActivityNotification`,
      {
        method: 'POST',
        body: JSON.stringify(body)
      },
      false
    );
    if (!r.ok) {
      return graphError(
        r.error?.message || 'Failed to send user teamwork activity notification',
        r.error?.code,
        r.error?.status
      );
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to send user teamwork activity notification');
  }
}

/** `POST /chats/{id}/sendActivityNotification` — requires `TeamsActivity.Send`; returns 204. */
export async function sendChatActivityNotification(
  token: string,
  chatId: string,
  body: Record<string, unknown>
): Promise<GraphResponse<void>> {
  try {
    const c = encodeURIComponent(chatId.trim());
    const r = await callGraph<void>(
      token,
      `/chats/${c}/sendActivityNotification`,
      {
        method: 'POST',
        body: JSON.stringify(body)
      },
      false
    );
    if (!r.ok) {
      return graphError(
        r.error?.message || 'Failed to send chat activity notification',
        r.error?.code,
        r.error?.status
      );
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to send chat activity notification');
  }
}

/** PATCH root channel message (`ChannelMessage.Send`). */
export async function updateChannelMessage(
  token: string,
  teamId: string,
  channelId: string,
  messageId: string,
  body: Record<string, unknown>,
  useBeta: boolean
): Promise<GraphResponse<GraphChatMessage>> {
  try {
    const t = encodeURIComponent(teamId.trim());
    const c = encodeURIComponent(channelId.trim());
    const m = encodeURIComponent(messageId.trim());
    const path = `/teams/${t}/channels/${c}/messages/${m}`;
    const r = useBeta
      ? await callGraphAt<GraphChatMessage>(getGraphBetaUrl(), token, path, {
          method: 'PATCH',
          body: JSON.stringify(body)
        })
      : await callGraph<GraphChatMessage>(token, path, {
          method: 'PATCH',
          body: JSON.stringify(body)
        });
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to update channel message', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to update channel message');
  }
}

/** PATCH channel message reply. */
export async function updateChannelMessageReply(
  token: string,
  teamId: string,
  channelId: string,
  parentMessageId: string,
  replyId: string,
  body: Record<string, unknown>,
  useBeta: boolean
): Promise<GraphResponse<GraphChatMessage>> {
  try {
    const t = encodeURIComponent(teamId.trim());
    const ch = encodeURIComponent(channelId.trim());
    const p = encodeURIComponent(parentMessageId.trim());
    const r = encodeURIComponent(replyId.trim());
    const path = `/teams/${t}/channels/${ch}/messages/${p}/replies/${r}`;
    const res = useBeta
      ? await callGraphAt<GraphChatMessage>(getGraphBetaUrl(), token, path, {
          method: 'PATCH',
          body: JSON.stringify(body)
        })
      : await callGraph<GraphChatMessage>(token, path, {
          method: 'PATCH',
          body: JSON.stringify(body)
        });
    if (!res.ok || !res.data) {
      return graphError(res.error?.message || 'Failed to update reply', res.error?.code, res.error?.status);
    }
    return graphResult(res.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to update reply');
  }
}

/** Hard-delete channel message (`DELETE`). Optional `If-Match` from message `@odata.etag`. */
export async function deleteChannelMessageHard(
  token: string,
  teamId: string,
  channelId: string,
  messageId: string,
  ifMatch?: string
): Promise<GraphResponse<void>> {
  try {
    const t = encodeURIComponent(teamId.trim());
    const c = encodeURIComponent(channelId.trim());
    const m = encodeURIComponent(messageId.trim());
    const headers: Record<string, string> = {};
    if (ifMatch?.trim()) headers['If-Match'] = ifMatch.trim();
    const r = await callGraph<void>(
      token,
      `/teams/${t}/channels/${c}/messages/${m}`,
      { method: 'DELETE', headers },
      false
    );
    if (!r.ok) {
      return graphError(r.error?.message || 'Failed to delete channel message', r.error?.code, r.error?.status);
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to delete channel message');
  }
}

/** Hard-delete a reply (`DELETE …/messages/{parent}/replies/{reply}`). */
export async function deleteChannelMessageReplyHard(
  token: string,
  teamId: string,
  channelId: string,
  parentMessageId: string,
  replyId: string,
  ifMatch?: string
): Promise<GraphResponse<void>> {
  try {
    const t = encodeURIComponent(teamId.trim());
    const c = encodeURIComponent(channelId.trim());
    const p = encodeURIComponent(parentMessageId.trim());
    const r = encodeURIComponent(replyId.trim());
    const headers: Record<string, string> = {};
    if (ifMatch?.trim()) headers['If-Match'] = ifMatch.trim();
    const res = await callGraph<void>(
      token,
      `/teams/${t}/channels/${c}/messages/${p}/replies/${r}`,
      { method: 'DELETE', headers },
      false
    );
    if (!res.ok) {
      return graphError(res.error?.message || 'Failed to delete reply', res.error?.code, res.error?.status);
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to delete reply');
  }
}

/** POST `…/softDelete` (typically **beta**). */
export async function softDeleteChannelMessage(
  token: string,
  teamId: string,
  channelId: string,
  messageId: string,
  useBeta: boolean
): Promise<GraphResponse<void>> {
  try {
    const t = encodeURIComponent(teamId.trim());
    const c = encodeURIComponent(channelId.trim());
    const m = encodeURIComponent(messageId.trim());
    const path = `/teams/${t}/channels/${c}/messages/${m}/softDelete`;
    const r = useBeta
      ? await callGraphAt<void>(getGraphBetaUrl(), token, path, { method: 'POST' }, false)
      : await callGraph<void>(token, path, { method: 'POST' }, false);
    if (!r.ok) {
      return graphError(r.error?.message || 'Failed to soft-delete channel message', r.error?.code, r.error?.status);
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to soft-delete channel message');
  }
}

export async function undoSoftDeleteChannelMessage(
  token: string,
  teamId: string,
  channelId: string,
  messageId: string,
  useBeta: boolean
): Promise<GraphResponse<void>> {
  try {
    const t = encodeURIComponent(teamId.trim());
    const c = encodeURIComponent(channelId.trim());
    const m = encodeURIComponent(messageId.trim());
    const path = `/teams/${t}/channels/${c}/messages/${m}/undoSoftDelete`;
    const r = useBeta
      ? await callGraphAt<void>(getGraphBetaUrl(), token, path, { method: 'POST' }, false)
      : await callGraph<void>(token, path, { method: 'POST' }, false);
    if (!r.ok) {
      return graphError(r.error?.message || 'Failed to undo soft-delete', r.error?.code, r.error?.status);
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to undo soft-delete');
  }
}

export async function softDeleteChannelMessageReply(
  token: string,
  teamId: string,
  channelId: string,
  parentMessageId: string,
  replyId: string,
  useBeta: boolean
): Promise<GraphResponse<void>> {
  try {
    const t = encodeURIComponent(teamId.trim());
    const c = encodeURIComponent(channelId.trim());
    const p = encodeURIComponent(parentMessageId.trim());
    const r = encodeURIComponent(replyId.trim());
    const path = `/teams/${t}/channels/${c}/messages/${p}/replies/${r}/softDelete`;
    const res = useBeta
      ? await callGraphAt<void>(getGraphBetaUrl(), token, path, { method: 'POST' }, false)
      : await callGraph<void>(token, path, { method: 'POST' }, false);
    if (!res.ok) {
      return graphError(res.error?.message || 'Failed to soft-delete reply', res.error?.code, res.error?.status);
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to soft-delete reply');
  }
}

export async function undoSoftDeleteChannelMessageReply(
  token: string,
  teamId: string,
  channelId: string,
  parentMessageId: string,
  replyId: string,
  useBeta: boolean
): Promise<GraphResponse<void>> {
  try {
    const t = encodeURIComponent(teamId.trim());
    const c = encodeURIComponent(channelId.trim());
    const p = encodeURIComponent(parentMessageId.trim());
    const r = encodeURIComponent(replyId.trim());
    const path = `/teams/${t}/channels/${c}/messages/${p}/replies/${r}/undoSoftDelete`;
    const res = useBeta
      ? await callGraphAt<void>(getGraphBetaUrl(), token, path, { method: 'POST' }, false)
      : await callGraph<void>(token, path, { method: 'POST' }, false);
    if (!res.ok) {
      return graphError(res.error?.message || 'Failed to undo soft-delete reply', res.error?.code, res.error?.status);
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to undo soft-delete reply');
  }
}

export async function updateChatMessage(
  token: string,
  chatId: string,
  messageId: string,
  body: Record<string, unknown>,
  useBeta: boolean
): Promise<GraphResponse<GraphChatMessage>> {
  try {
    const c = encodeURIComponent(chatId.trim());
    const m = encodeURIComponent(messageId.trim());
    const path = `/chats/${c}/messages/${m}`;
    const r = useBeta
      ? await callGraphAt<GraphChatMessage>(getGraphBetaUrl(), token, path, {
          method: 'PATCH',
          body: JSON.stringify(body)
        })
      : await callGraph<GraphChatMessage>(token, path, {
          method: 'PATCH',
          body: JSON.stringify(body)
        });
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to update chat message', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to update chat message');
  }
}

export async function updateChatMessageReply(
  token: string,
  chatId: string,
  parentMessageId: string,
  replyId: string,
  body: Record<string, unknown>,
  useBeta: boolean
): Promise<GraphResponse<GraphChatMessage>> {
  try {
    const c = encodeURIComponent(chatId.trim());
    const p = encodeURIComponent(parentMessageId.trim());
    const r = encodeURIComponent(replyId.trim());
    const path = `/chats/${c}/messages/${p}/replies/${r}`;
    const res = useBeta
      ? await callGraphAt<GraphChatMessage>(getGraphBetaUrl(), token, path, {
          method: 'PATCH',
          body: JSON.stringify(body)
        })
      : await callGraph<GraphChatMessage>(token, path, {
          method: 'PATCH',
          body: JSON.stringify(body)
        });
    if (!res.ok || !res.data) {
      return graphError(res.error?.message || 'Failed to update chat reply', res.error?.code, res.error?.status);
    }
    return graphResult(res.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to update chat reply');
  }
}

export async function deleteChatMessageHard(
  token: string,
  chatId: string,
  messageId: string,
  ifMatch?: string
): Promise<GraphResponse<void>> {
  try {
    const c = encodeURIComponent(chatId.trim());
    const m = encodeURIComponent(messageId.trim());
    const headers: Record<string, string> = {};
    if (ifMatch?.trim()) headers['If-Match'] = ifMatch.trim();
    const r = await callGraph<void>(token, `/chats/${c}/messages/${m}`, { method: 'DELETE', headers }, false);
    if (!r.ok) {
      return graphError(r.error?.message || 'Failed to delete chat message', r.error?.code, r.error?.status);
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to delete chat message');
  }
}

export async function deleteChatMessageReplyHard(
  token: string,
  chatId: string,
  parentMessageId: string,
  replyId: string,
  ifMatch?: string
): Promise<GraphResponse<void>> {
  try {
    const c = encodeURIComponent(chatId.trim());
    const p = encodeURIComponent(parentMessageId.trim());
    const r = encodeURIComponent(replyId.trim());
    const headers: Record<string, string> = {};
    if (ifMatch?.trim()) headers['If-Match'] = ifMatch.trim();
    const res = await callGraph<void>(
      token,
      `/chats/${c}/messages/${p}/replies/${r}`,
      { method: 'DELETE', headers },
      false
    );
    if (!res.ok) {
      return graphError(res.error?.message || 'Failed to delete chat reply', res.error?.code, res.error?.status);
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to delete chat reply');
  }
}

export async function softDeleteChatMessage(
  token: string,
  chatId: string,
  messageId: string,
  useBeta: boolean
): Promise<GraphResponse<void>> {
  try {
    const c = encodeURIComponent(chatId.trim());
    const m = encodeURIComponent(messageId.trim());
    const path = `/chats/${c}/messages/${m}/softDelete`;
    const r = useBeta
      ? await callGraphAt<void>(getGraphBetaUrl(), token, path, { method: 'POST' }, false)
      : await callGraph<void>(token, path, { method: 'POST' }, false);
    if (!r.ok) {
      return graphError(r.error?.message || 'Failed to soft-delete chat message', r.error?.code, r.error?.status);
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to soft-delete chat message');
  }
}

export async function undoSoftDeleteChatMessage(
  token: string,
  chatId: string,
  messageId: string,
  useBeta: boolean
): Promise<GraphResponse<void>> {
  try {
    const c = encodeURIComponent(chatId.trim());
    const m = encodeURIComponent(messageId.trim());
    const path = `/chats/${c}/messages/${m}/undoSoftDelete`;
    const r = useBeta
      ? await callGraphAt<void>(getGraphBetaUrl(), token, path, { method: 'POST' }, false)
      : await callGraph<void>(token, path, { method: 'POST' }, false);
    if (!r.ok) {
      return graphError(r.error?.message || 'Failed to undo soft-delete chat message', r.error?.code, r.error?.status);
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to undo soft-delete chat message');
  }
}

export async function softDeleteChatMessageReply(
  token: string,
  chatId: string,
  parentMessageId: string,
  replyId: string,
  useBeta: boolean
): Promise<GraphResponse<void>> {
  try {
    const c = encodeURIComponent(chatId.trim());
    const p = encodeURIComponent(parentMessageId.trim());
    const r = encodeURIComponent(replyId.trim());
    const path = `/chats/${c}/messages/${p}/replies/${r}/softDelete`;
    const res = useBeta
      ? await callGraphAt<void>(getGraphBetaUrl(), token, path, { method: 'POST' }, false)
      : await callGraph<void>(token, path, { method: 'POST' }, false);
    if (!res.ok) {
      return graphError(res.error?.message || 'Failed to soft-delete chat reply', res.error?.code, res.error?.status);
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to soft-delete chat reply');
  }
}

export async function undoSoftDeleteChatMessageReply(
  token: string,
  chatId: string,
  parentMessageId: string,
  replyId: string,
  useBeta: boolean
): Promise<GraphResponse<void>> {
  try {
    const c = encodeURIComponent(chatId.trim());
    const p = encodeURIComponent(parentMessageId.trim());
    const r = encodeURIComponent(replyId.trim());
    const path = `/chats/${c}/messages/${p}/replies/${r}/undoSoftDelete`;
    const res = useBeta
      ? await callGraphAt<void>(getGraphBetaUrl(), token, path, { method: 'POST' }, false)
      : await callGraph<void>(token, path, { method: 'POST' }, false);
    if (!res.ok) {
      return graphError(
        res.error?.message || 'Failed to undo soft-delete chat reply',
        res.error?.code,
        res.error?.status
      );
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to undo soft-delete chat reply');
  }
}

/** `POST /chats` — create 1:1 or group chat (`Chat.ReadWrite`). */
export async function createChat(
  token: string,
  body: Record<string, unknown>
): Promise<GraphResponse<GraphChatDetail>> {
  try {
    const r = await callGraph<GraphChatDetail>(token, '/chats', {
      method: 'POST',
      body: JSON.stringify(body)
    });
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to create chat', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to create chat');
  }
}

/** `POST /chats/{id}/members` — add member to existing chat. */
export async function addChatMember(
  token: string,
  chatId: string,
  body: Record<string, unknown>
): Promise<GraphResponse<GraphTeamMember>> {
  try {
    const c = encodeURIComponent(chatId.trim());
    const r = await callGraph<GraphTeamMember>(token, `/chats/${c}/members`, {
      method: 'POST',
      body: JSON.stringify(body)
    });
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to add chat member', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to add chat member');
  }
}

/** `POST /teams/{id}/members` — e.g. `conversationMember` payload (`TeamMember.ReadWriteNonGuestRole.All` / `Group.ReadWrite.All`). */
export async function addTeamMember(
  token: string,
  teamId: string,
  body: Record<string, unknown>
): Promise<GraphResponse<GraphTeamMember>> {
  try {
    const t = encodeURIComponent(teamId.trim());
    const r = await callGraph<GraphTeamMember>(token, `/teams/${t}/members`, {
      method: 'POST',
      body: JSON.stringify(body)
    });
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to add team member', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to add team member');
  }
}

/** `POST /teams/{id}/channels/{channelId}/members` */
export async function addChannelMember(
  token: string,
  teamId: string,
  channelId: string,
  body: Record<string, unknown>
): Promise<GraphResponse<GraphTeamMember>> {
  try {
    const t = encodeURIComponent(teamId.trim());
    const c = encodeURIComponent(channelId.trim());
    const r = await callGraph<GraphTeamMember>(token, `/teams/${t}/channels/${c}/members`, {
      method: 'POST',
      body: JSON.stringify(body)
    });
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to add channel member', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to add channel member');
  }
}

/** `GET …/tabs/{tabId}` */
export async function getChannelTab(
  token: string,
  teamId: string,
  channelId: string,
  tabId: string
): Promise<GraphResponse<TeamsChannelTab>> {
  try {
    const t = encodeURIComponent(teamId.trim());
    const c = encodeURIComponent(channelId.trim());
    const id = encodeURIComponent(tabId.trim());
    const r = await callGraph<TeamsChannelTab>(token, `/teams/${t}/channels/${c}/tabs/${id}?$expand=teamsApp`);
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to get tab', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get tab');
  }
}

export async function createChannelTab(
  token: string,
  teamId: string,
  channelId: string,
  body: Record<string, unknown>
): Promise<GraphResponse<TeamsChannelTab>> {
  try {
    const t = encodeURIComponent(teamId.trim());
    const c = encodeURIComponent(channelId.trim());
    const r = await callGraph<TeamsChannelTab>(token, `/teams/${t}/channels/${c}/tabs`, {
      method: 'POST',
      body: JSON.stringify(body)
    });
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to create tab', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to create tab');
  }
}

export async function updateChannelTab(
  token: string,
  teamId: string,
  channelId: string,
  tabId: string,
  body: Record<string, unknown>
): Promise<GraphResponse<TeamsChannelTab>> {
  try {
    const t = encodeURIComponent(teamId.trim());
    const c = encodeURIComponent(channelId.trim());
    const id = encodeURIComponent(tabId.trim());
    const r = await callGraph<TeamsChannelTab>(token, `/teams/${t}/channels/${c}/tabs/${id}`, {
      method: 'PATCH',
      body: JSON.stringify(body)
    });
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to update tab', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to update tab');
  }
}

export async function deleteChannelTab(
  token: string,
  teamId: string,
  channelId: string,
  tabId: string
): Promise<GraphResponse<void>> {
  try {
    const t = encodeURIComponent(teamId.trim());
    const c = encodeURIComponent(channelId.trim());
    const id = encodeURIComponent(tabId.trim());
    const r = await callGraph<void>(token, `/teams/${t}/channels/${c}/tabs/${id}`, { method: 'DELETE' }, false);
    if (!r.ok) {
      return graphError(r.error?.message || 'Failed to delete tab', r.error?.code, r.error?.status);
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to delete tab');
  }
}
