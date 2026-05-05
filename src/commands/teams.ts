import { readFile } from 'node:fs/promises';
import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import type { GraphResponse } from '../lib/graph-client.js';
import {
  addChannelMember,
  addChatInstalledApp,
  addChatMember,
  addTeamInstalledApp,
  addTeamMember,
  addUserTeamworkInstalledApp,
  buildAddTeamsAppInstallationBody,
  createChannelTab,
  createChat,
  deleteChannelMessageHard,
  deleteChannelMessageReplyHard,
  deleteChannelTab,
  deleteChatInstalledApp,
  deleteChatMessageHard,
  deleteChatMessageReplyHard,
  deleteTeamInstalledApp,
  deleteUserTeamworkInstalledApp,
  getChannelMessage,
  getChannelTab,
  getChat,
  getChatInstalledApp,
  getChatMessage,
  getTeam,
  getTeamChannel,
  getTeamChannelFilesFolder,
  getTeamInstalledApp,
  getTeamPrimaryChannel,
  getTeamsAppCatalogEntry,
  getUserTeamworkInstalledApp,
  listChannelMessageReplies,
  listChannelMessages,
  listChannelTabs,
  listChatInstalledApps,
  listChatMembers,
  listChatMessageReplies,
  listChatMessages,
  listChatPinnedMessages,
  listJoinedTeams,
  listMyChats,
  listTeamAllChannels,
  listTeamChannelMembers,
  listTeamChannels,
  listTeamIncomingChannels,
  listTeamInstalledApps,
  listTeamMembers,
  listTeamsAppCatalog,
  listUserTeamworkInstalledApps,
  patchChatInstalledApp,
  patchTeamInstalledApp,
  sendChannelMessage,
  sendChannelMessageReply,
  sendChatActivityNotification,
  sendChatMessage,
  sendChatMessageReply,
  sendMeTeamworkActivityNotification,
  sendUserTeamworkActivityNotification,
  setChannelMessageReaction,
  setChatMessageReaction,
  softDeleteChannelMessage,
  softDeleteChannelMessageReply,
  softDeleteChatMessage,
  softDeleteChatMessageReply,
  undoSoftDeleteChannelMessage,
  undoSoftDeleteChannelMessageReply,
  undoSoftDeleteChatMessage,
  undoSoftDeleteChatMessageReply,
  unsetChannelMessageReaction,
  unsetChatMessageReaction,
  updateChannelMessage,
  updateChannelMessageReply,
  updateChannelTab,
  updateChatMessage,
  updateChatMessageReply,
  upgradeChatInstalledApp,
  upgradeTeamInstalledApp
} from '../lib/graph-teams-client.js';
import { buildTeamsHtmlBodyWithMentions, parseAtSpecs } from '../lib/teams-message-compose.js';
import { checkReadOnly } from '../lib/utils.js';

function collectAtSpec(value: string, previous: string[]): string[] {
  return [...previous, value];
}

export const teamsCommand = new Command('teams')
  .description(
    'Microsoft Teams via Graph: teams, channels, chats, messages, tabs, apps, and notifications. See GRAPH_SCOPES.md for permissions.'
  )
  .addHelpText(
    'after',
    `
Examples:
  m365-agent-cli teams list
  m365-agent-cli teams list --user someone@contoso.com
  m365-agent-cli teams channels <teamId>
  m365-agent-cli teams channel-message-send <teamId> <channelId> --text "Hello" --at <userId>:Name

Channel/chat mentions: use --at userId:displayName and put @displayName in --text. Chats list is /me/chats only.
See docs/CLI_REFERENCE.md for full flag lists.
`
  );

teamsCommand
  .command('list')
  .summary('Teams List')
  .description('List joined teams for /me or another user (GET /me/joinedTeams or GET /users/{id}/joinedTeams)')
  .option(
    '--user <upn-or-id>',
    'User object id or UPN — lists that user’s joined teams (requires Team.ReadBasic.All or equivalent)'
  )
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { user?: string; json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await listJoinedTeams(auth.token, opts.user);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(r.data, null, 2));
      return;
    }
    for (const t of r.data) {
      console.log(`${t.displayName ?? '(no name)'}\t${t.id}`);
    }
  });

teamsCommand
  .command('get')
  .summary('Teams Get')
  .description('Get a team by id (GET /teams/{id})')
  .argument('<teamId>', 'Team id (GUID)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (teamId: string, opts: { json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await getTeam(auth.token, teamId);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log(opts.json ? JSON.stringify(r.data, null, 2) : `${r.data.displayName ?? ''}\t${r.data.id}`);
  });

teamsCommand
  .command('primary-channel')
  .summary('Teams Primary Channel')
  .description('Get the team General channel (GET /teams/{id}/primaryChannel; Channel.ReadBasic.All)')
  .argument('<teamId>', 'Team id')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (teamId: string, opts: { json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await getTeamPrimaryChannel(auth.token, teamId);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log(
      opts.json
        ? JSON.stringify(r.data, null, 2)
        : `${r.data.displayName ?? '(channel)'}\t${r.data.id}\t${r.data.membershipType ?? ''}`
    );
  });

teamsCommand
  .command('channel-files-folder')
  .summary('Teams Channel Files Folder')
  .description(
    'Get the channel Files root drive item (GET …/channels/{id}/filesFolder). Use printed driveId with `files list --drive-id … --folder <id>`.'
  )
  .argument('<teamId>', 'Team id')
  .argument('<channelId>', 'Channel id')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (teamId: string, channelId: string, opts: { json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await getTeamChannelFilesFolder(auth.token, teamId, channelId);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(r.data, null, 2));
      return;
    }
    const driveId = r.data.parentReference?.driveId ?? '';
    const folderId = r.data.id ?? '';
    console.log(`driveId:\t${driveId}`);
    console.log(`folderItemId:\t${folderId}`);
    if (r.data.webUrl) console.log(`webUrl:\t${r.data.webUrl}`);
    if (driveId && folderId) {
      console.log(`Example: m365-agent-cli files list --drive-id "${driveId}" --folder "${folderId}"`);
    }
  });

teamsCommand
  .command('channels')
  .summary('Teams Channels')
  .description('List channels in a team (GET /teams/{id}/channels)')
  .argument('<teamId>', 'Team id')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (teamId: string, opts: { json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await listTeamChannels(auth.token, teamId);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(r.data, null, 2));
      return;
    }
    for (const c of r.data) {
      console.log(`${c.displayName ?? '(channel)'}\t${c.id}\t${c.membershipType ?? ''}`);
    }
  });

teamsCommand
  .command('all-channels')
  .summary('Teams All Channels')
  .description('List all channels including shared/incoming (GET /teams/{id}/allChannels; Channel.ReadBasic.All)')
  .argument('<teamId>', 'Team id')
  .option('--filter <odata>', "OData $filter e.g. membershipType eq 'shared' (quoted for your shell)")
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (teamId: string, opts: { filter?: string; json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await listTeamAllChannels(auth.token, teamId, opts.filter);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(r.data, null, 2));
      return;
    }
    for (const c of r.data) {
      console.log(`${c.displayName ?? '(channel)'}\t${c.membershipType ?? ''}\t${c.tenantId ?? ''}\t${c.id}`);
    }
  });

teamsCommand
  .command('incoming-channels')
  .summary('Teams Incoming Channels')
  .description(
    'List channels shared into this team from other tenants (GET /teams/{id}/incomingChannels; Channel.ReadBasic.All)'
  )
  .argument('<teamId>', 'Team id')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (teamId: string, opts: { json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await listTeamIncomingChannels(auth.token, teamId);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(r.data, null, 2));
      return;
    }
    for (const c of r.data) {
      console.log(`${c.displayName ?? '(channel)'}\t${c.membershipType ?? ''}\t${c.tenantId ?? ''}\t${c.id}`);
    }
  });

teamsCommand
  .command('channel-get')
  .summary('Teams Channel Get')
  .description('Get one channel (GET /teams/{id}/channels/{id}; Channel.ReadBasic.All)')
  .argument('<teamId>', 'Team id')
  .argument('<channelId>', 'Channel id')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (teamId: string, channelId: string, opts: { json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await getTeamChannel(auth.token, teamId, channelId);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log(
      opts.json
        ? JSON.stringify(r.data, null, 2)
        : `${r.data.displayName ?? '(channel)'}\t${r.data.id}\t${r.data.membershipType ?? ''}`
    );
  });

teamsCommand
  .command('channel-members')
  .summary('Teams Channel Members')
  .description('List members of a channel (GET …/channels/{id}/members; ChannelMember.Read.All or Group.ReadWrite.All)')
  .argument('<teamId>', 'Team id')
  .argument('<channelId>', 'Channel id')
  .option('-n, --top <n>', 'Page size (max 999)', '')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      teamId: string,
      channelId: string,
      opts: { top?: string; json?: boolean; token?: string; identity?: string }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const top = opts.top ? parseInt(opts.top, 10) : undefined;
      const r = await listTeamChannelMembers(auth.token, teamId, channelId, top && top > 0 ? top : undefined);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(r.data, null, 2));
        return;
      }
      for (const m of r.data) {
        console.log(`${m.displayName ?? '(member)'}\t${m.email ?? ''}\t${m.roles?.join(',') ?? ''}\t${m.userId ?? ''}`);
      }
    }
  );

teamsCommand
  .command('tabs')
  .summary('Teams Tabs')
  .description(
    'List tabs in a channel (GET …/channels/{id}/tabs?$expand=teamsApp; Group.ReadWrite.All or TeamsTab.Read.All)'
  )
  .argument('<teamId>', 'Team id')
  .argument('<channelId>', 'Channel id')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (teamId: string, channelId: string, opts: { json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await listChannelTabs(auth.token, teamId, channelId);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(r.data, null, 2));
      return;
    }
    for (const t of r.data) {
      const app = t.teamsApp?.displayName ?? t.teamsApp?.id ?? '';
      console.log(`${t.displayName ?? '(tab)'}\t${app}\t${t.id ?? ''}`);
    }
  });

teamsCommand
  .command('messages')
  .summary('Teams Messages')
  .description('List recent messages in a channel (GET …/channels/{id}/messages)')
  .argument('<teamId>', 'Team id')
  .argument('<channelId>', 'Channel id')
  .option('-n, --top <n>', 'Page size (max 50)', '10')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      teamId: string,
      channelId: string,
      opts: { top?: string; json?: boolean; token?: string; identity?: string }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const top = Math.min(50, Math.max(1, parseInt(opts.top ?? '10', 10) || 10));
      const r = await listChannelMessages(auth.token, teamId, channelId, top);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(r.data, null, 2));
        return;
      }
      for (const m of r.data) {
        const who = m.from?.user?.displayName ?? m.from?.user?.id ?? '?';
        const preview = (m.body?.content ?? '').replace(/\s+/g, ' ').trim().slice(0, 120);
        console.log(`${m.createdDateTime ?? ''}\t${who}\t${preview}`);
      }
    }
  );

teamsCommand
  .command('channel-message-get')
  .summary('Teams Channel Message Get')
  .description('Get one channel message by id (GET …/channels/{id}/messages/{id})')
  .argument('<teamId>', 'Team id')
  .argument('<channelId>', 'Channel id')
  .argument('<messageId>', 'Message id')
  .option('--json', 'Output full JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      teamId: string,
      channelId: string,
      messageId: string,
      opts: { json?: boolean; token?: string; identity?: string }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getChannelMessage(auth.token, teamId, channelId, messageId);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(r.data, null, 2));
        return;
      }
      const who = r.data.from?.user?.displayName ?? r.data.from?.user?.id ?? '?';
      const preview = (r.data.body?.content ?? '').replace(/\s+/g, ' ').trim().slice(0, 200);
      console.log(`${r.data.createdDateTime ?? ''}\t${who}\t${preview}\t${r.data.id ?? ''}`);
    }
  );

teamsCommand
  .command('channel-message-send')
  .summary('Teams Channel Message Send')
  .description(
    'Post a message to a channel (`ChannelMessage.Send`). Use `--json-file` for full `chatMessage` body, or `--text` / `--html`.'
  )
  .argument('<teamId>', 'Team id')
  .argument('<channelId>', 'Channel id')
  .option('--json-file <path>', 'Full JSON body for POST (overrides --text/--html)')
  .option('--text <s>', 'Plain text body (contentType text)')
  .option('--html <s>', 'HTML body (contentType html)')
  .option(
    '--at <userId:displayName>',
    'User mention (repeatable). Requires --text containing @displayName for each (builds HTML + mentions).',
    collectAtSpec,
    [] as string[]
  )
  .option('--json', 'Print response JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      teamId: string,
      channelId: string,
      opts: {
        jsonFile?: string;
        text?: string;
        html?: string;
        at?: string[];
        json?: boolean;
        token?: string;
        identity?: string;
      },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      let body: Record<string, unknown>;
      const atList = opts.at ?? [];
      if (opts.jsonFile?.trim()) {
        body = JSON.parse(await readFile(opts.jsonFile.trim(), 'utf-8')) as Record<string, unknown>;
      } else if (atList.length > 0) {
        if (!opts.text?.trim()) {
          console.error('Error: --at requires --text (with @displayName tokens matching each mention)');
          process.exit(1);
        }
        if (opts.html?.trim()) {
          console.error('Error: do not combine --html with --at; use --json-file for full control');
          process.exit(1);
        }
        try {
          const built = buildTeamsHtmlBodyWithMentions(opts.text.trim(), parseAtSpecs(atList));
          body = { body: built.body, mentions: built.mentions };
        } catch (e) {
          console.error(e instanceof Error ? e.message : e);
          process.exit(1);
        }
      } else if (opts.html?.trim()) {
        body = { body: { contentType: 'html', content: opts.html } };
      } else if (opts.text?.trim()) {
        body = { body: { contentType: 'text', content: opts.text } };
      } else {
        console.error('Error: provide --json-file, --text, or --html');
        process.exit(1);
      }
      const r = await sendChannelMessage(auth.token, teamId, channelId, body);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(opts.json ? JSON.stringify(r.data, null, 2) : `${r.data.id ?? ''}`);
    }
  );

teamsCommand
  .command('message-replies')
  .summary('Teams Message Replies')
  .description(
    'List replies to a channel message (GET …/messages/{id}/replies; ChannelMessage.Read.All or Group.ReadWrite.All)'
  )
  .argument('<teamId>', 'Team id')
  .argument('<channelId>', 'Channel id')
  .argument('<messageId>', 'Parent message id')
  .option('-n, --top <n>', 'Page size (max 50)', '20')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      teamId: string,
      channelId: string,
      messageId: string,
      opts: { top?: string; json?: boolean; token?: string; identity?: string }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const top = Math.min(50, Math.max(1, parseInt(opts.top ?? '20', 10) || 20));
      const r = await listChannelMessageReplies(auth.token, teamId, channelId, messageId, top);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(r.data, null, 2));
        return;
      }
      for (const m of r.data) {
        const who = m.from?.user?.displayName ?? m.from?.user?.id ?? '?';
        const preview = (m.body?.content ?? '').replace(/\s+/g, ' ').trim().slice(0, 120);
        console.log(`${m.createdDateTime ?? ''}\t${who}\t${preview}\t${m.id ?? ''}`);
      }
    }
  );

teamsCommand
  .command('channel-message-reply')
  .summary('Teams Channel Message Reply')
  .description(
    'Reply to a channel message (`POST …/messages/{id}/replies`; `ChannelMessage.Send`). Same body options as **channel-message-send**.'
  )
  .argument('<teamId>', 'Team id')
  .argument('<channelId>', 'Channel id')
  .argument('<messageId>', 'Parent message id')
  .option('--json-file <path>', 'Full JSON body for POST')
  .option('--text <s>', 'Plain text body')
  .option('--html <s>', 'HTML body')
  .option(
    '--at <userId:displayName>',
    'User mention (repeatable). Requires --text with @displayName per mention.',
    collectAtSpec,
    [] as string[]
  )
  .option('--json', 'Print response JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      teamId: string,
      channelId: string,
      messageId: string,
      opts: {
        jsonFile?: string;
        text?: string;
        html?: string;
        at?: string[];
        json?: boolean;
        token?: string;
        identity?: string;
      },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      let body: Record<string, unknown>;
      const atList = opts.at ?? [];
      if (opts.jsonFile?.trim()) {
        body = JSON.parse(await readFile(opts.jsonFile.trim(), 'utf-8')) as Record<string, unknown>;
      } else if (atList.length > 0) {
        if (!opts.text?.trim()) {
          console.error('Error: --at requires --text (with @displayName tokens matching each mention)');
          process.exit(1);
        }
        if (opts.html?.trim()) {
          console.error('Error: do not combine --html with --at; use --json-file for full control');
          process.exit(1);
        }
        try {
          const built = buildTeamsHtmlBodyWithMentions(opts.text.trim(), parseAtSpecs(atList));
          body = { body: built.body, mentions: built.mentions };
        } catch (e) {
          console.error(e instanceof Error ? e.message : e);
          process.exit(1);
        }
      } else if (opts.html?.trim()) {
        body = { body: { contentType: 'html', content: opts.html } };
      } else if (opts.text?.trim()) {
        body = { body: { contentType: 'text', content: opts.text } };
      } else {
        console.error('Error: provide --json-file, --text, or --html');
        process.exit(1);
      }
      const r = await sendChannelMessageReply(auth.token, teamId, channelId, messageId, body);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(opts.json ? JSON.stringify(r.data, null, 2) : `${r.data.id ?? ''}`);
    }
  );

teamsCommand
  .command('members')
  .summary('Teams Members')
  .description('List members of a team (GET /teams/{id}/members; uses Group.ReadWrite.All or TeamMember.Read.*)')
  .argument('<teamId>', 'Team id')
  .option('-n, --top <n>', 'Page size (default: server default)', '')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (teamId: string, opts: { top?: string; json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const top = opts.top ? parseInt(opts.top, 10) : undefined;
    const r = await listTeamMembers(auth.token, teamId, top && top > 0 ? top : undefined);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(r.data, null, 2));
      return;
    }
    for (const m of r.data) {
      console.log(`${m.displayName ?? '(member)'}\t${m.email ?? ''}\t${m.roles?.join(',') ?? ''}\t${m.id ?? ''}`);
    }
  });

teamsCommand
  .command('chats')
  .summary('Teams Chats')
  .description(
    'List chats for the signed-in user only (GET /me/chats; requires Chat.Read). There is no Graph equivalent to list another user’s chats by id.'
  )
  .option('-n, --top <n>', 'Page size (max 50)', '20')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { top?: string; json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const top = Math.min(50, Math.max(1, parseInt(opts.top ?? '20', 10) || 20));
    const r = await listMyChats(auth.token, top);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(r.data, null, 2));
      return;
    }
    for (const c of r.data) {
      const label = c.topic?.trim() || c.chatType || '(chat)';
      console.log(`${label}\t${c.chatType ?? ''}\t${c.id}`);
    }
  });

teamsCommand
  .command('chat-get')
  .summary('Teams Chat Get')
  .description(
    'Get one chat (GET /chats/{id}); optional --expand members, lastMessagePreview, or both (comma-separated)'
  )
  .argument('<chatId>', 'Chat id')
  .option('--expand <segments>', 'Graph $expand e.g. members or lastMessagePreview or members,lastMessagePreview')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (chatId: string, opts: { expand?: string; json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await getChat(auth.token, chatId, opts.expand);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log(
      opts.json
        ? JSON.stringify(r.data, null, 2)
        : `${r.data.topic ?? r.data.chatType ?? ''}\t${r.data.chatType ?? ''}\t${r.data.id}`
    );
  });

teamsCommand
  .command('chat-messages')
  .summary('Teams Chat Messages')
  .description(
    'List messages in a chat (GET /chats/{id}/messages; requires Chat.ReadWrite for read in same app registration)'
  )
  .argument('<chatId>', 'Chat id')
  .option('-n, --top <n>', 'Page size (max 50)', '10')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (chatId: string, opts: { top?: string; json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const top = Math.min(50, Math.max(1, parseInt(opts.top ?? '10', 10) || 10));
    const r = await listChatMessages(auth.token, chatId, top);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(r.data, null, 2));
      return;
    }
    for (const m of r.data) {
      const who = m.from?.user?.displayName ?? m.from?.user?.id ?? '?';
      const preview = (m.body?.content ?? '').replace(/\s+/g, ' ').trim().slice(0, 120);
      console.log(`${m.createdDateTime ?? ''}\t${who}\t${preview}`);
    }
  });

teamsCommand
  .command('chat-message-get')
  .summary('Teams Chat Message Get')
  .description('Get one chat message by id (GET /chats/{id}/messages/{id})')
  .argument('<chatId>', 'Chat id')
  .argument('<messageId>', 'Message id')
  .option('--json', 'Output full JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (chatId: string, messageId: string, opts: { json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await getChatMessage(auth.token, chatId, messageId);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(r.data, null, 2));
      return;
    }
    const who = r.data.from?.user?.displayName ?? r.data.from?.user?.id ?? '?';
    const preview = (r.data.body?.content ?? '').replace(/\s+/g, ' ').trim().slice(0, 200);
    console.log(`${r.data.createdDateTime ?? ''}\t${who}\t${preview}\t${r.data.id ?? ''}`);
  });

teamsCommand
  .command('chat-message-replies')
  .summary('Teams Chat Message Replies')
  .description('List replies to a chat message (GET …/messages/{id}/replies)')
  .argument('<chatId>', 'Chat id')
  .argument('<messageId>', 'Parent message id')
  .option('-n, --top <n>', 'Page size (max 50)', '20')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      chatId: string,
      messageId: string,
      opts: { top?: string; json?: boolean; token?: string; identity?: string }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const top = Math.min(50, Math.max(1, parseInt(opts.top ?? '20', 10) || 20));
      const r = await listChatMessageReplies(auth.token, chatId, messageId, top);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(r.data, null, 2));
        return;
      }
      for (const m of r.data) {
        const who = m.from?.user?.displayName ?? m.from?.user?.id ?? '?';
        const preview = (m.body?.content ?? '').replace(/\s+/g, ' ').trim().slice(0, 120);
        console.log(`${m.createdDateTime ?? ''}\t${who}\t${preview}\t${m.id ?? ''}`);
      }
    }
  );

teamsCommand
  .command('chat-message-send')
  .summary('Teams Chat Message Send')
  .description('Post a message to a chat (`Chat.ReadWrite`). Use `--json-file` or `--text` / `--html`.')
  .argument('<chatId>', 'Chat id')
  .option('--json-file <path>', 'Full JSON body for POST')
  .option('--text <s>', 'Plain text body')
  .option('--html <s>', 'HTML body')
  .option(
    '--at <userId:displayName>',
    'User mention (repeatable). Requires --text with @displayName per mention.',
    collectAtSpec,
    [] as string[]
  )
  .option('--json', 'Print response JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      chatId: string,
      opts: {
        jsonFile?: string;
        text?: string;
        html?: string;
        at?: string[];
        json?: boolean;
        token?: string;
        identity?: string;
      },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      let body: Record<string, unknown>;
      const atList = opts.at ?? [];
      if (opts.jsonFile?.trim()) {
        body = JSON.parse(await readFile(opts.jsonFile.trim(), 'utf-8')) as Record<string, unknown>;
      } else if (atList.length > 0) {
        if (!opts.text?.trim()) {
          console.error('Error: --at requires --text (with @displayName tokens matching each mention)');
          process.exit(1);
        }
        if (opts.html?.trim()) {
          console.error('Error: do not combine --html with --at; use --json-file for full control');
          process.exit(1);
        }
        try {
          const built = buildTeamsHtmlBodyWithMentions(opts.text.trim(), parseAtSpecs(atList));
          body = { body: built.body, mentions: built.mentions };
        } catch (e) {
          console.error(e instanceof Error ? e.message : e);
          process.exit(1);
        }
      } else if (opts.html?.trim()) {
        body = { body: { contentType: 'html', content: opts.html } };
      } else if (opts.text?.trim()) {
        body = { body: { contentType: 'text', content: opts.text } };
      } else {
        console.error('Error: provide --json-file, --text, or --html');
        process.exit(1);
      }
      const r = await sendChatMessage(auth.token, chatId, body);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(opts.json ? JSON.stringify(r.data, null, 2) : `${r.data.id ?? ''}`);
    }
  );

teamsCommand
  .command('chat-message-reply')
  .summary('Teams Chat Message Reply')
  .description(
    'Reply in a chat thread (`POST …/messages/{id}/replies`; `Chat.ReadWrite`). Same body options as **chat-message-send**.'
  )
  .argument('<chatId>', 'Chat id')
  .argument('<messageId>', 'Parent message id')
  .option('--json-file <path>', 'Full JSON body for POST')
  .option('--text <s>', 'Plain text body')
  .option('--html <s>', 'HTML body')
  .option(
    '--at <userId:displayName>',
    'User mention (repeatable). Requires --text with @displayName per mention.',
    collectAtSpec,
    [] as string[]
  )
  .option('--json', 'Print response JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      chatId: string,
      messageId: string,
      opts: {
        jsonFile?: string;
        text?: string;
        html?: string;
        at?: string[];
        json?: boolean;
        token?: string;
        identity?: string;
      },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      let body: Record<string, unknown>;
      const atList = opts.at ?? [];
      if (opts.jsonFile?.trim()) {
        body = JSON.parse(await readFile(opts.jsonFile.trim(), 'utf-8')) as Record<string, unknown>;
      } else if (atList.length > 0) {
        if (!opts.text?.trim()) {
          console.error('Error: --at requires --text (with @displayName tokens matching each mention)');
          process.exit(1);
        }
        if (opts.html?.trim()) {
          console.error('Error: do not combine --html with --at; use --json-file for full control');
          process.exit(1);
        }
        try {
          const built = buildTeamsHtmlBodyWithMentions(opts.text.trim(), parseAtSpecs(atList));
          body = { body: built.body, mentions: built.mentions };
        } catch (e) {
          console.error(e instanceof Error ? e.message : e);
          process.exit(1);
        }
      } else if (opts.html?.trim()) {
        body = { body: { contentType: 'html', content: opts.html } };
      } else if (opts.text?.trim()) {
        body = { body: { contentType: 'text', content: opts.text } };
      } else {
        console.error('Error: provide --json-file, --text, or --html');
        process.exit(1);
      }
      const r = await sendChatMessageReply(auth.token, chatId, messageId, body);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(opts.json ? JSON.stringify(r.data, null, 2) : `${r.data.id ?? ''}`);
    }
  );

teamsCommand
  .command('chat-pinned')
  .summary('Teams Chat Pinned')
  .description(
    'List pinned messages in a chat (GET /chats/{id}/pinnedMessages; Chat.ReadWrite; --expand-message for bodies)'
  )
  .argument('<chatId>', 'Chat id')
  .option('--expand-message', 'Include full message via $expand=message')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (chatId: string, opts: { expandMessage?: boolean; json?: boolean; token?: string; identity?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await listChatPinnedMessages(auth.token, chatId, opts.expandMessage);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(r.data, null, 2));
        return;
      }
      for (const p of r.data) {
        if (p.message) {
          const who = p.message.from?.user?.displayName ?? p.message.from?.user?.id ?? '?';
          const preview = (p.message.body?.content ?? '').replace(/\s+/g, ' ').trim().slice(0, 120);
          console.log(`${p.id ?? ''}\t${p.message.createdDateTime ?? ''}\t${who}\t${preview}`);
        } else {
          console.log(`${p.id ?? ''}\t\t\t(pinned id only; use --expand-message or --json)`);
        }
      }
    }
  );

teamsCommand
  .command('chat-members')
  .summary('Teams Chat Members')
  .description('List members of a chat (GET /chats/{id}/members; Chat.Read or Chat.ReadBasic)')
  .argument('<chatId>', 'Chat id')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (chatId: string, opts: { json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await listChatMembers(auth.token, chatId);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(r.data, null, 2));
      return;
    }
    for (const m of r.data) {
      console.log(`${m.displayName ?? '(member)'}\t${m.email ?? ''}\t${m.roles?.join(',') ?? ''}\t${m.userId ?? ''}`);
    }
  });

teamsCommand
  .command('apps')
  .summary('Teams Apps')
  .description(
    'List apps installed in a team (GET …/installedApps?$expand=teamsAppDefinition; Group.ReadWrite.All or TeamsAppInstallation.*)'
  )
  .argument('<teamId>', 'Team id')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (teamId: string, opts: { json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await listTeamInstalledApps(auth.token, teamId);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(r.data, null, 2));
      return;
    }
    for (const a of r.data) {
      const name = a.teamsAppDefinition?.displayName ?? '(app)';
      const ver = a.teamsAppDefinition?.version ?? '';
      const appId = a.teamsAppDefinition?.teamsAppId ?? '';
      console.log(`${name}\t${ver}\t${appId}\t${a.id ?? ''}`);
    }
  });

teamsCommand
  .command('app-catalog')
  .summary('Teams App Catalog')
  .description(
    "List Teams apps in the tenant/store catalog (`GET /appCatalogs/teamsApps`; `AppCatalog.Read.All` or related). Use --filter e.g. distributionMethod eq 'organization'"
  )
  .option('--filter <odata>', 'OData $filter')
  .option('--expand <segments>', 'e.g. appDefinitions')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { filter?: string; expand?: string; json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await listTeamsAppCatalog(auth.token, opts.filter, opts.expand);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(r.data, null, 2));
      return;
    }
    for (const a of r.data) {
      console.log(`${a.displayName ?? '(app)'}\t${a.distributionMethod ?? ''}\t${a.externalId ?? ''}\t${a.id ?? ''}`);
    }
  });

teamsCommand
  .command('app-catalog-get')
  .summary('Teams App Catalog Get')
  .description('Get one catalog app by id (`GET /appCatalogs/teamsApps/{id}`)')
  .argument('<catalogTeamsAppId>', 'teamsApp id from app-catalog')
  .option('--expand <segments>', 'e.g. appDefinitions')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (catalogTeamsAppId: string, opts: { expand?: string; json?: boolean; token?: string; identity?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getTeamsAppCatalogEntry(auth.token, catalogTeamsAppId, opts.expand);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(opts.json ? JSON.stringify(r.data, null, 2) : `${r.data.displayName ?? ''}\t${r.data.id ?? ''}`);
    }
  );

teamsCommand
  .command('app-get')
  .summary('Teams App Get')
  .description('Get one app installation on a team (`GET …/installedApps/{id}`)')
  .argument('<teamId>', 'Team id')
  .argument('<installationId>', 'teamsAppInstallation id (from **teams apps**)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (teamId: string, installationId: string, opts: { json?: boolean; token?: string; identity?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getTeamInstalledApp(auth.token, teamId, installationId);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(r.data, null, 2));
        return;
      }
      const name = r.data.teamsAppDefinition?.displayName ?? '(app)';
      console.log(`${name}\t${r.data.id ?? ''}`);
    }
  );

teamsCommand
  .command('app-add')
  .summary('Teams App Add')
  .description(
    'Install an app on a team (`POST …/installedApps`). Use `--teams-app-id` (catalog id) or full `--json-file` (e.g. teamsApp@odata.bind + consentedPermissionSet).'
  )
  .argument('<teamId>', 'Team id')
  .option('--teams-app-id <id>', 'teamsApp id from **teams app-catalog**')
  .option('--json-file <path>', 'Full JSON body (overrides --teams-app-id)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      teamId: string,
      opts: { teamsAppId?: string; jsonFile?: string; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      let body: Record<string, unknown>;
      if (opts.jsonFile?.trim()) {
        body = JSON.parse(await readFile(opts.jsonFile.trim(), 'utf-8')) as Record<string, unknown>;
      } else if (opts.teamsAppId?.trim()) {
        body = buildAddTeamsAppInstallationBody(opts.teamsAppId.trim());
      } else {
        console.error('Error: provide --teams-app-id or --json-file');
        process.exit(1);
      }
      const r = await addTeamInstalledApp(auth.token, teamId, body);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log('App install requested (201).');
    }
  );

teamsCommand
  .command('app-patch')
  .summary('Teams App Patch')
  .description('PATCH a team app installation (`PATCH …/installedApps/{id}`)')
  .argument('<teamId>', 'Team id')
  .argument('<installationId>', 'Installation id')
  .requiredOption('--json-file <path>', 'JSON PATCH body')
  .option('--json', 'Print response JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      teamId: string,
      installationId: string,
      opts: { jsonFile: string; json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const body = JSON.parse(await readFile(opts.jsonFile.trim(), 'utf-8')) as Record<string, unknown>;
      const r = await patchTeamInstalledApp(auth.token, teamId, installationId, body);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(opts.json ? JSON.stringify(r.data, null, 2) : `${r.data.id ?? 'ok'}`);
    }
  );

teamsCommand
  .command('app-upgrade')
  .summary('Teams App Upgrade')
  .description('Upgrade a team app installation to the latest catalog version (`POST …/upgrade`)')
  .argument('<teamId>', 'Team id')
  .argument('<installationId>', 'Installation id')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (teamId: string, installationId: string, opts: { token?: string; identity?: string }, cmd: Command) => {
    checkReadOnly(cmd);
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await upgradeTeamInstalledApp(auth.token, teamId, installationId);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log('Upgrade started (204).');
  });

teamsCommand
  .command('app-delete')
  .summary('Teams App Delete')
  .description('Remove an app from a team (`DELETE …/installedApps/{id}`)')
  .argument('<teamId>', 'Team id')
  .argument('<installationId>', 'Installation id')
  .option('--if-match <etag>', 'If-Match header')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      teamId: string,
      installationId: string,
      opts: { ifMatch?: string; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await deleteTeamInstalledApp(auth.token, teamId, installationId, opts.ifMatch);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log('App removed.');
    }
  );

teamsCommand
  .command('chat-apps')
  .summary('Teams Chat Apps')
  .description('List apps installed in a chat (`GET /chats/{id}/installedApps`)')
  .argument('<chatId>', 'Chat id')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (chatId: string, opts: { json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await listChatInstalledApps(auth.token, chatId);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(r.data, null, 2));
      return;
    }
    for (const a of r.data) {
      const name = a.teamsAppDefinition?.displayName ?? '(app)';
      console.log(`${name}\t${a.teamsAppDefinition?.teamsAppId ?? ''}\t${a.id ?? ''}`);
    }
  });

teamsCommand
  .command('chat-app-get')
  .summary('Teams Chat App Get')
  .description('Get one chat app installation')
  .argument('<chatId>', 'Chat id')
  .argument('<installationId>', 'Installation id')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (chatId: string, installationId: string, opts: { json?: boolean; token?: string; identity?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getChatInstalledApp(auth.token, chatId, installationId);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(
        opts.json
          ? JSON.stringify(r.data, null, 2)
          : `${r.data.teamsAppDefinition?.displayName ?? ''}\t${r.data.id ?? ''}`
      );
    }
  );

teamsCommand
  .command('chat-app-add')
  .summary('Teams Chat App Add')
  .description('Install an app in a chat (`POST /chats/{id}/installedApps`)')
  .argument('<chatId>', 'Chat id')
  .option('--teams-app-id <id>', 'Catalog teamsApp id')
  .option('--json-file <path>', 'Full JSON body')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      chatId: string,
      opts: { teamsAppId?: string; jsonFile?: string; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      let body: Record<string, unknown>;
      if (opts.jsonFile?.trim()) {
        body = JSON.parse(await readFile(opts.jsonFile.trim(), 'utf-8')) as Record<string, unknown>;
      } else if (opts.teamsAppId?.trim()) {
        body = buildAddTeamsAppInstallationBody(opts.teamsAppId.trim());
      } else {
        console.error('Error: provide --teams-app-id or --json-file');
        process.exit(1);
      }
      const r = await addChatInstalledApp(auth.token, chatId, body);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log('App install requested (201).');
    }
  );

teamsCommand
  .command('chat-app-patch')
  .summary('Teams Chat App Patch')
  .description('PATCH a chat app installation')
  .argument('<chatId>', 'Chat id')
  .argument('<installationId>', 'Installation id')
  .requiredOption('--json-file <path>', 'JSON PATCH body')
  .option('--json', 'Print response JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      chatId: string,
      installationId: string,
      opts: { jsonFile: string; json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const body = JSON.parse(await readFile(opts.jsonFile.trim(), 'utf-8')) as Record<string, unknown>;
      const r = await patchChatInstalledApp(auth.token, chatId, installationId, body);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(opts.json ? JSON.stringify(r.data, null, 2) : `${r.data.id ?? 'ok'}`);
    }
  );

teamsCommand
  .command('chat-app-upgrade')
  .summary('Teams Chat App Upgrade')
  .description('Upgrade a chat app installation (`POST …/upgrade`)')
  .argument('<chatId>', 'Chat id')
  .argument('<installationId>', 'Installation id')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (chatId: string, installationId: string, opts: { token?: string; identity?: string }, cmd: Command) => {
    checkReadOnly(cmd);
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await upgradeChatInstalledApp(auth.token, chatId, installationId);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log('Upgrade started (204).');
  });

teamsCommand
  .command('chat-app-delete')
  .summary('Teams Chat App Delete')
  .description('Remove an app from a chat')
  .argument('<chatId>', 'Chat id')
  .argument('<installationId>', 'Installation id')
  .option('--if-match <etag>', 'If-Match header')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      chatId: string,
      installationId: string,
      opts: { ifMatch?: string; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await deleteChatInstalledApp(auth.token, chatId, installationId, opts.ifMatch);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log('App removed.');
    }
  );

teamsCommand
  .command('user-apps')
  .summary('Teams User Apps')
  .description(
    'Apps in the user’s personal Teams scope (`GET …/teamwork/installedApps`). Default: signed-in user; `--user` for `/users/{id}/…` (needs appropriate Teams app permissions).'
  )
  .option('--user <upn-or-id>', 'Another user (admin / delegated scenario)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { user?: string; json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await listUserTeamworkInstalledApps(auth.token, opts.user);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(r.data, null, 2));
      return;
    }
    for (const a of r.data) {
      const name = a.teamsAppDefinition?.displayName ?? '(app)';
      console.log(`${name}\t${a.teamsAppDefinition?.teamsAppId ?? ''}\t${a.id ?? ''}`);
    }
  });

teamsCommand
  .command('user-app-get')
  .summary('Teams User App Get')
  .description('Get one personal-scope app installation')
  .argument('<installationId>', 'Installation id')
  .option('--user <upn-or-id>', 'Target user (omit for /me)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (installationId: string, opts: { user?: string; json?: boolean; token?: string; identity?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getUserTeamworkInstalledApp(auth.token, installationId, opts.user);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(
        opts.json
          ? JSON.stringify(r.data, null, 2)
          : `${r.data.teamsAppDefinition?.displayName ?? ''}\t${r.data.id ?? ''}`
      );
    }
  );

teamsCommand
  .command('user-app-add')
  .summary('Teams User App Add')
  .description('Install an app in the user’s personal scope (`POST …/teamwork/installedApps`)')
  .option('--user <upn-or-id>', 'Target user (omit for /me)')
  .option('--teams-app-id <id>', 'Catalog teamsApp id')
  .option('--json-file <path>', 'Full JSON body (RSC consent, etc.)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      opts: { user?: string; teamsAppId?: string; jsonFile?: string; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      let body: Record<string, unknown>;
      if (opts.jsonFile?.trim()) {
        body = JSON.parse(await readFile(opts.jsonFile.trim(), 'utf-8')) as Record<string, unknown>;
      } else if (opts.teamsAppId?.trim()) {
        body = buildAddTeamsAppInstallationBody(opts.teamsAppId.trim());
      } else {
        console.error('Error: provide --teams-app-id or --json-file');
        process.exit(1);
      }
      const r = await addUserTeamworkInstalledApp(auth.token, body, opts.user);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log('App install requested (201).');
    }
  );

teamsCommand
  .command('user-app-delete')
  .summary('Teams User App Delete')
  .description('Remove an app from the user’s personal scope')
  .argument('<installationId>', 'Installation id')
  .option('--user <upn-or-id>', 'Target user (omit for /me)')
  .option('--if-match <etag>', 'If-Match header')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      installationId: string,
      opts: { user?: string; ifMatch?: string; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await deleteUserTeamworkInstalledApp(auth.token, installationId, opts.user, opts.ifMatch);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log('App removed.');
    }
  );

teamsCommand
  .command('activity-notify')
  .summary('Teams Activity Notify')
  .description(
    'Teams activity feed: POST /me/teamwork/sendActivityNotification (default), POST /chats/{id}/sendActivityNotification with --chat-id, or POST /users/{id}/teamwork/sendActivityNotification with --user-id (usually app-only token + `TeamsActivity.Send`). Body from --json-file (see Microsoft Graph docs).'
  )
  .requiredOption('--json-file <path>', 'JSON body: topic, activityType, previewText, recipient, …')
  .option('--chat-id <id>', 'If set, notify in chat scope (mutually exclusive with --user-id)')
  .option(
    '--user-id <id>',
    'Target user id or UPN for POST /users/{id}/teamwork/sendActivityNotification (typically use --token with an application access token; mutually exclusive with --chat-id)'
  )
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      opts: { jsonFile: string; chatId?: string; userId?: string; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const chat = opts.chatId?.trim();
      const uid = opts.userId?.trim();
      if (chat && uid) {
        console.error('Error: use either --chat-id or --user-id, not both.');
        process.exit(1);
      }
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const body = JSON.parse(await readFile(opts.jsonFile.trim(), 'utf-8')) as Record<string, unknown>;
      const r = chat
        ? await sendChatActivityNotification(auth.token, chat, body)
        : uid
          ? await sendUserTeamworkActivityNotification(auth.token, uid, body)
          : await sendMeTeamworkActivityNotification(auth.token, body);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log('Activity notification sent (204).');
    }
  );

teamsCommand
  .command('channel-message-patch')
  .summary('Teams Channel Message Patch')
  .description('PATCH a channel message or reply (`ChannelMessage.Send`). Use --json-file for the PATCH body.')
  .argument('<teamId>', 'Team id')
  .argument('<channelId>', 'Channel id')
  .argument('<messageId>', 'Message id (root) or reply id when --parent is set')
  .requiredOption('--json-file <path>', 'JSON body (e.g. { "body": { "contentType": "html", "content": "..." } })')
  .option('--parent <parentMessageId>', 'When set, PATCH …/messages/{parent}/replies/{messageId}')
  .option('--beta', 'Call Microsoft Graph beta host for this PATCH', false)
  .option('--json', 'Print response JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      teamId: string,
      channelId: string,
      messageId: string,
      opts: {
        jsonFile: string;
        parent?: string;
        beta?: boolean;
        json?: boolean;
        token?: string;
        identity?: string;
      },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const body = JSON.parse(await readFile(opts.jsonFile.trim(), 'utf-8')) as Record<string, unknown>;
      const useBeta = !!opts.beta;
      const r = opts.parent?.trim()
        ? await updateChannelMessageReply(auth.token, teamId, channelId, opts.parent.trim(), messageId, body, useBeta)
        : await updateChannelMessage(auth.token, teamId, channelId, messageId, body, useBeta);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(opts.json ? JSON.stringify(r.data, null, 2) : (r.data.id ?? 'ok'));
    }
  );

teamsCommand
  .command('channel-message-delete')
  .summary('Teams Channel Message Delete')
  .description(
    'Delete a channel message or reply: hard DELETE (`--hard`), soft-delete POST …/softDelete (default), or `--undo-soft` (often requires `--beta`). Use `--parent` when targeting a reply.'
  )
  .argument('<teamId>', 'Team id')
  .argument('<channelId>', 'Channel id')
  .argument('<messageId>', 'Root message id, or reply id when --parent is set')
  .option('--parent <parentMessageId>', 'Parent message id when deleting a reply')
  .option('--hard', 'Permanent DELETE (optional If-Match via --if-match)', false)
  .option('--undo-soft', 'POST undoSoftDelete instead of softDelete', false)
  .option('--beta', 'Use Graph beta host for soft/undo-soft paths', false)
  .option('--if-match <etag>', 'If-Match header for hard delete')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      teamId: string,
      channelId: string,
      messageId: string,
      opts: {
        parent?: string;
        hard?: boolean;
        undoSoft?: boolean;
        beta?: boolean;
        ifMatch?: string;
        token?: string;
        identity?: string;
      },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const useBeta = !!opts.beta;
      const parent = opts.parent?.trim();
      let r: GraphResponse<void>;
      if (parent) {
        if (opts.hard) {
          r = await deleteChannelMessageReplyHard(auth.token, teamId, channelId, parent, messageId, opts.ifMatch);
        } else if (opts.undoSoft) {
          r = await undoSoftDeleteChannelMessageReply(auth.token, teamId, channelId, parent, messageId, useBeta);
        } else {
          r = await softDeleteChannelMessageReply(auth.token, teamId, channelId, parent, messageId, useBeta);
        }
      } else if (opts.hard) {
        r = await deleteChannelMessageHard(auth.token, teamId, channelId, messageId, opts.ifMatch);
      } else if (opts.undoSoft) {
        r = await undoSoftDeleteChannelMessage(auth.token, teamId, channelId, messageId, useBeta);
      } else {
        r = await softDeleteChannelMessage(auth.token, teamId, channelId, messageId, useBeta);
      }
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log('Done.');
    }
  );

teamsCommand
  .command('chat-message-patch')
  .summary('Teams Chat Message Patch')
  .description('PATCH a chat message (`Chat.ReadWrite`). Use --json-file for the PATCH body.')
  .argument('<chatId>', 'Chat id')
  .argument('<messageId>', 'Message id')
  .requiredOption('--json-file <path>', 'JSON PATCH body')
  .option('--beta', 'Use Graph beta host', false)
  .option('--json', 'Print response JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      chatId: string,
      messageId: string,
      opts: { jsonFile: string; beta?: boolean; json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const body = JSON.parse(await readFile(opts.jsonFile.trim(), 'utf-8')) as Record<string, unknown>;
      const r = await updateChatMessage(auth.token, chatId, messageId, body, !!opts.beta);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(opts.json ? JSON.stringify(r.data, null, 2) : (r.data.id ?? 'ok'));
    }
  );

teamsCommand
  .command('chat-message-reply-patch')
  .summary('Teams Chat Message Reply Patch')
  .description('PATCH a reply in a chat thread (`Chat.ReadWrite`).')
  .argument('<chatId>', 'Chat id')
  .argument('<parentMessageId>', 'Root message id')
  .argument('<replyId>', 'Reply message id')
  .requiredOption('--json-file <path>', 'JSON PATCH body')
  .option('--beta', 'Use Graph beta host', false)
  .option('--json', 'Print response JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      chatId: string,
      parentMessageId: string,
      replyId: string,
      opts: { jsonFile: string; beta?: boolean; json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const body = JSON.parse(await readFile(opts.jsonFile.trim(), 'utf-8')) as Record<string, unknown>;
      const r = await updateChatMessageReply(auth.token, chatId, parentMessageId, replyId, body, !!opts.beta);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(opts.json ? JSON.stringify(r.data, null, 2) : (r.data.id ?? 'ok'));
    }
  );

teamsCommand
  .command('chat-message-delete')
  .summary('Teams Chat Message Delete')
  .description(
    'Delete a chat message or reply: `--hard` (DELETE), default soft-delete POST, or `--undo-soft`. Use `--parent` for replies.'
  )
  .argument('<chatId>', 'Chat id')
  .argument('<messageId>', 'Message id or reply id when --parent is set')
  .option('--parent <parentMessageId>', 'Parent message id when deleting a reply')
  .option('--hard', 'Permanent DELETE', false)
  .option('--undo-soft', 'POST undoSoftDelete', false)
  .option('--beta', 'Use Graph beta host for soft/undo-soft', false)
  .option('--if-match <etag>', 'If-Match for hard delete')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      chatId: string,
      messageId: string,
      opts: {
        parent?: string;
        hard?: boolean;
        undoSoft?: boolean;
        beta?: boolean;
        ifMatch?: string;
        token?: string;
        identity?: string;
      },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const useBeta = !!opts.beta;
      const parent = opts.parent?.trim();
      let r: GraphResponse<void>;
      if (parent) {
        if (opts.hard) {
          r = await deleteChatMessageReplyHard(auth.token, chatId, parent, messageId, opts.ifMatch);
        } else if (opts.undoSoft) {
          r = await undoSoftDeleteChatMessageReply(auth.token, chatId, parent, messageId, useBeta);
        } else {
          r = await softDeleteChatMessageReply(auth.token, chatId, parent, messageId, useBeta);
        }
      } else if (opts.hard) {
        r = await deleteChatMessageHard(auth.token, chatId, messageId, opts.ifMatch);
      } else if (opts.undoSoft) {
        r = await undoSoftDeleteChatMessage(auth.token, chatId, messageId, useBeta);
      } else {
        r = await softDeleteChatMessage(auth.token, chatId, messageId, useBeta);
      }
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log('Done.');
    }
  );

teamsCommand
  .command('chat-create')
  .summary('Teams Chat Create')
  .description('Create a chat (`POST /chats`; `Chat.ReadWrite`). Body from --json-file (members, chatType, …).')
  .requiredOption('--json-file <path>', 'Full JSON body for POST /chats')
  .option('--json', 'Print created chat JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { jsonFile: string; json?: boolean; token?: string; identity?: string }, cmd: Command) => {
    checkReadOnly(cmd);
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const body = JSON.parse(await readFile(opts.jsonFile.trim(), 'utf-8')) as Record<string, unknown>;
    const r = await createChat(auth.token, body);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else console.log(`${r.data.id ?? ''}\t${r.data.topic ?? ''}`);
  });

teamsCommand
  .command('chat-member-add')
  .summary('Teams Chat Member Add')
  .description('Add a member to a chat (`POST /chats/{id}/members`). Body from --json-file (conversationMember).')
  .argument('<chatId>', 'Chat id')
  .requiredOption('--json-file <path>', 'JSON body for new member')
  .option('--json', 'Print response JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      chatId: string,
      opts: { jsonFile: string; json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const body = JSON.parse(await readFile(opts.jsonFile.trim(), 'utf-8')) as Record<string, unknown>;
      const r = await addChatMember(auth.token, chatId, body);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(opts.json ? JSON.stringify(r.data, null, 2) : `${r.data.id ?? ''}\t${r.data.displayName ?? ''}`);
    }
  );

teamsCommand
  .command('team-member-add')
  .summary('Teams Team Member Add')
  .description(
    'Add a member to a team (`POST /teams/{id}/members`). Body from --json-file. Requires elevated team membership permissions.'
  )
  .argument('<teamId>', 'Team id')
  .requiredOption('--json-file <path>', 'JSON body (conversationMember)')
  .option('--json', 'Print response JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      teamId: string,
      opts: { jsonFile: string; json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const body = JSON.parse(await readFile(opts.jsonFile.trim(), 'utf-8')) as Record<string, unknown>;
      const r = await addTeamMember(auth.token, teamId, body);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(opts.json ? JSON.stringify(r.data, null, 2) : `${r.data.id ?? ''}\t${r.data.displayName ?? ''}`);
    }
  );

teamsCommand
  .command('channel-member-add')
  .summary('Teams Channel Member Add')
  .description('Add a member to a channel (`POST …/channels/{id}/members`). Body from --json-file.')
  .argument('<teamId>', 'Team id')
  .argument('<channelId>', 'Channel id')
  .requiredOption('--json-file <path>', 'JSON body (conversationMember)')
  .option('--json', 'Print response JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      teamId: string,
      channelId: string,
      opts: { jsonFile: string; json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const body = JSON.parse(await readFile(opts.jsonFile.trim(), 'utf-8')) as Record<string, unknown>;
      const r = await addChannelMember(auth.token, teamId, channelId, body);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(opts.json ? JSON.stringify(r.data, null, 2) : `${r.data.id ?? ''}\t${r.data.displayName ?? ''}`);
    }
  );

teamsCommand
  .command('tab-get')
  .summary('Teams Tab Get')
  .description('Get one channel tab (`GET …/tabs/{id}` with $expand=teamsApp)')
  .argument('<teamId>', 'Team id')
  .argument('<channelId>', 'Channel id')
  .argument('<tabId>', 'Tab id')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      teamId: string,
      channelId: string,
      tabId: string,
      opts: { json?: boolean; token?: string; identity?: string }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getChannelTab(auth.token, teamId, channelId, tabId);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(opts.json ? JSON.stringify(r.data, null, 2) : `${r.data.displayName ?? ''}\t${r.data.id ?? ''}`);
    }
  );

teamsCommand
  .command('tab-create')
  .summary('Teams Tab Create')
  .description('Create a channel tab (`POST …/tabs`). Body from --json-file.')
  .argument('<teamId>', 'Team id')
  .argument('<channelId>', 'Channel id')
  .requiredOption('--json-file <path>', 'JSON body for teamsTab')
  .option('--json', 'Print response JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      teamId: string,
      channelId: string,
      opts: { jsonFile: string; json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const body = JSON.parse(await readFile(opts.jsonFile.trim(), 'utf-8')) as Record<string, unknown>;
      const r = await createChannelTab(auth.token, teamId, channelId, body);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(opts.json ? JSON.stringify(r.data, null, 2) : `${r.data.id ?? ''}`);
    }
  );

teamsCommand
  .command('tab-update')
  .summary('Teams Tab Update')
  .description('PATCH a channel tab (`PATCH …/tabs/{id}`). Body from --json-file.')
  .argument('<teamId>', 'Team id')
  .argument('<channelId>', 'Channel id')
  .argument('<tabId>', 'Tab id')
  .requiredOption('--json-file <path>', 'JSON PATCH body')
  .option('--json', 'Print response JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      teamId: string,
      channelId: string,
      tabId: string,
      opts: { jsonFile: string; json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const body = JSON.parse(await readFile(opts.jsonFile.trim(), 'utf-8')) as Record<string, unknown>;
      const r = await updateChannelTab(auth.token, teamId, channelId, tabId, body);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(opts.json ? JSON.stringify(r.data, null, 2) : `${r.data.id ?? ''}`);
    }
  );

teamsCommand
  .command('tab-delete')
  .summary('Teams Tab Delete')
  .description('Delete a channel tab (`DELETE …/tabs/{id}`)')
  .argument('<teamId>', 'Team id')
  .argument('<channelId>', 'Channel id')
  .argument('<tabId>', 'Tab id')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      teamId: string,
      channelId: string,
      tabId: string,
      opts: { token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await deleteChannelTab(auth.token, teamId, channelId, tabId);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log('Tab deleted.');
    }
  );

teamsCommand
  .command('channel-message-react')
  .summary('Teams Channel Message React')
  .description('Set or remove a reaction on a channel message (POST setReaction / unsetReaction; ChannelMessage.Send)')
  .argument('<teamId>', 'Team id')
  .argument('<channelId>', 'Channel id')
  .argument('<messageId>', 'Message id')
  .requiredOption('-r, --reaction <unicode>', 'Reaction unicode string (e.g. 👍 or 💙)')
  .option('--reply <replyId>', 'Target a reply in the thread instead of the root message')
  .option('--unset', 'Call unsetReaction instead of setReaction', false)
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      teamId: string,
      channelId: string,
      messageId: string,
      opts: { reaction: string; reply?: string; unset?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = opts.unset
        ? await unsetChannelMessageReaction(auth.token, teamId, channelId, messageId, opts.reaction, opts.reply)
        : await setChannelMessageReaction(auth.token, teamId, channelId, messageId, opts.reaction, opts.reply);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(opts.unset ? 'Reaction removed.' : 'Reaction set.');
    }
  );

teamsCommand
  .command('chat-message-react')
  .summary('Teams Chat Message React')
  .description('Set or remove a reaction on a chat message (POST setReaction / unsetReaction)')
  .argument('<chatId>', 'Chat id')
  .argument('<messageId>', 'Message id')
  .requiredOption('-r, --reaction <unicode>', 'Reaction unicode string')
  .option('--unset', 'Call unsetReaction', false)
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      chatId: string,
      messageId: string,
      opts: { reaction: string; unset?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = opts.unset
        ? await unsetChatMessageReaction(auth.token, chatId, messageId, opts.reaction)
        : await setChatMessageReaction(auth.token, chatId, messageId, opts.reaction);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(opts.unset ? 'Reaction removed.' : 'Reaction set.');
    }
  );
