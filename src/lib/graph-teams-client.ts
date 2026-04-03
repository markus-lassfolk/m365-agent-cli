import {
  callGraph,
  GraphApiError,
  type GraphResponse,
  graphError,
  graphResult
} from './graph-client.js';

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

export async function listJoinedTeams(token: string): Promise<GraphResponse<GraphTeam[]>> {
  try {
    const r = await callGraph<{ value: GraphTeam[] }>(token, '/me/joinedTeams');
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to list joined teams', r.error?.code, r.error?.status);
    }
    return graphResult(r.data.value ?? []);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to list joined teams');
  }
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
export async function getTeamPrimaryChannel(
  token: string,
  teamId: string
): Promise<GraphResponse<GraphChannel>> {
  try {
    const r = await callGraph<GraphChannel>(
      token,
      `/teams/${encodeURIComponent(teamId.trim())}/primaryChannel`
    );
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to get primary channel', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get primary channel');
  }
}

export async function listTeamChannels(
  token: string,
  teamId: string
): Promise<GraphResponse<GraphChannel[]>> {
  try {
    const r = await callGraph<{ value: GraphChannel[] }>(
      token,
      `/teams/${encodeURIComponent(teamId)}/channels`
    );
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to list channels', r.error?.code, r.error?.status);
    }
    return graphResult(r.data.value ?? []);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to list channels');
  }
}

/** Includes shared / incoming channels. Least privilege: Channel.ReadBasic.All. Optional OData `$filter` e.g. `membershipType eq 'shared'`. */
export async function listTeamAllChannels(
  token: string,
  teamId: string,
  filter?: string
): Promise<GraphResponse<GraphChannel[]>> {
  try {
    const q =
      filter && filter.trim() ? `?$filter=${encodeURIComponent(filter.trim())}` : '';
    const r = await callGraph<{ value: GraphChannel[] }>(
      token,
      `/teams/${encodeURIComponent(teamId.trim())}/allChannels${q}`
    );
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to list all channels', r.error?.code, r.error?.status);
    }
    return graphResult(r.data.value ?? []);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to list all channels');
  }
}

/** Channels shared into this team from other tenants. GET /teams/{id}/incomingChannels (Channel.ReadBasic.All). */
export async function listTeamIncomingChannels(
  token: string,
  teamId: string
): Promise<GraphResponse<GraphChannel[]>> {
  try {
    const r = await callGraph<{ value: GraphChannel[] }>(
      token,
      `/teams/${encodeURIComponent(teamId.trim())}/incomingChannels`
    );
    if (!r.ok || !r.data) {
      return graphError(
        r.error?.message || 'Failed to list incoming channels',
        r.error?.code,
        r.error?.status
      );
    }
    return graphResult(r.data.value ?? []);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to list incoming channels');
  }
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
  try {
    const t = encodeURIComponent(teamId.trim());
    const c = encodeURIComponent(channelId.trim());
    const n = top && top > 0 ? Math.min(top, 999) : undefined;
    const qs = n ? `?$top=${n}` : '';
    const r = await callGraph<{ value: GraphTeamMember[] }>(
      token,
      `/teams/${t}/channels/${c}/members${qs}`
    );
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to list channel members', r.error?.code, r.error?.status);
    }
    return graphResult(r.data.value ?? []);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to list channel members');
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
  try {
    const t = encodeURIComponent(teamId.trim());
    const c = encodeURIComponent(channelId.trim());
    const m = encodeURIComponent(messageId.trim());
    const n = top && top > 0 ? Math.min(top, 50) : undefined;
    const qs = n ? `?$top=${n}` : '';
    const r = await callGraph<{ value: GraphChatMessage[] }>(
      token,
      `/teams/${t}/channels/${c}/messages/${m}/replies${qs}`
    );
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to list message replies', r.error?.code, r.error?.status);
    }
    return graphResult(r.data.value ?? []);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to list message replies');
  }
}

export async function listTeamMembers(
  token: string,
  teamId: string,
  top?: number
): Promise<GraphResponse<GraphTeamMember[]>> {
  try {
    const t = top && top > 0 ? Math.min(top, 999) : undefined;
    const qs = t ? `?$top=${t}` : '';
    const r = await callGraph<{ value: GraphTeamMember[] }>(
      token,
      `/teams/${encodeURIComponent(teamId)}/members${qs}`
    );
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to list team members', r.error?.code, r.error?.status);
    }
    return graphResult(r.data.value ?? []);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to list team members');
  }
}

/**
 * `expand` e.g. `members`, `lastMessagePreview`, or `members,lastMessagePreview` (Graph `$expand`).
 * Least privilege: Chat.ReadBasic without expand; Chat.Read for richer expands where required.
 */
export async function getChat(
  token: string,
  chatId: string,
  expand?: string
): Promise<GraphResponse<GraphChatDetail>> {
  try {
    const enc = encodeURIComponent(chatId.trim());
    const q =
      expand && expand.trim() ? `?$expand=${encodeURIComponent(expand.trim())}` : '';
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
  try {
    const t = top && top > 0 ? Math.min(top, 50) : undefined;
    const qs = t ? `?$top=${t}` : '';
    const r = await callGraph<{ value: GraphChat[] }>(token, `/me/chats${qs}`);
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to list chats', r.error?.code, r.error?.status);
    }
    return graphResult(r.data.value ?? []);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to list chats');
  }
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
    const r = await callGraph<{ value: PinnedChatMessageInfo[] }>(
      token,
      `/chats/${enc}/pinnedMessages${q}`
    );
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to list pinned messages', r.error?.code, r.error?.status);
    }
    return graphResult(r.data.value ?? []);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to list pinned messages');
  }
}

export async function listChatMembers(
  token: string,
  chatId: string
): Promise<GraphResponse<GraphTeamMember[]>> {
  try {
    const r = await callGraph<{ value: GraphTeamMember[] }>(
      token,
      `/chats/${encodeURIComponent(chatId)}/members`
    );
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to list chat members', r.error?.code, r.error?.status);
    }
    return graphResult(r.data.value ?? []);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to list chat members');
  }
}

/** Uses `$expand=teamsAppDefinition` for display names. `Group.ReadWrite.All` satisfies Graph (see team-list-installedapps). */
export async function listTeamInstalledApps(
  token: string,
  teamId: string
): Promise<GraphResponse<TeamsAppInstallation[]>> {
  try {
    const path = `/teams/${encodeURIComponent(teamId)}/installedApps?$expand=teamsAppDefinition`;
    const r = await callGraph<{ value: TeamsAppInstallation[] }>(token, path);
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to list installed apps', r.error?.code, r.error?.status);
    }
    return graphResult(r.data.value ?? []);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to list installed apps');
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
