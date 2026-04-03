import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import {
  getChat,
  getTeam,
  getTeamChannel,
  getTeamPrimaryChannel,
  listChannelMessageReplies,
  listChannelMessages,
  listChannelTabs,
  listChatMembers,
  listChatMessages,
  listChatPinnedMessages,
  listJoinedTeams,
  listMyChats,
  listTeamAllChannels,
  listTeamIncomingChannels,
  listTeamChannels,
  listTeamChannelMembers,
  listTeamInstalledApps,
  listTeamMembers
} from '../lib/graph-teams-client.js';
export const teamsCommand = new Command('teams').description(
  'Microsoft Teams (Graph): teams, channels, tabs, messages, members, chats, apps (delegated; see GRAPH_SCOPES.md)'
);

teamsCommand
  .command('list')
  .description('List teams the signed-in user has joined (GET /me/joinedTeams)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await listJoinedTeams(auth.token);
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
  .command('channels')
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
  .description(
    'List all channels including shared/incoming (GET /teams/{id}/allChannels; Channel.ReadBasic.All)'
  )
  .argument('<teamId>', 'Team id')
  .option(
    '--filter <odata>',
    'OData $filter e.g. membershipType eq \'shared\' (quoted for your shell)'
  )
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (teamId: string, opts: { filter?: string; json?: boolean; token?: string; identity?: string }) => {
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
        console.log(
          `${c.displayName ?? '(channel)'}\t${c.membershipType ?? ''}\t${c.tenantId ?? ''}\t${c.id}`
        );
      }
    }
  );

teamsCommand
  .command('incoming-channels')
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
      console.log(
        `${c.displayName ?? '(channel)'}\t${c.membershipType ?? ''}\t${c.tenantId ?? ''}\t${c.id}`
      );
    }
  });

teamsCommand
  .command('channel-get')
  .description('Get one channel (GET /teams/{id}/channels/{id}; Channel.ReadBasic.All)')
  .argument('<teamId>', 'Team id')
  .argument('<channelId>', 'Channel id')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      teamId: string,
      channelId: string,
      opts: { json?: boolean; token?: string; identity?: string }
    ) => {
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
    }
  );

teamsCommand
  .command('channel-members')
  .description(
    'List members of a channel (GET …/channels/{id}/members; ChannelMember.Read.All or Group.ReadWrite.All)'
  )
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
      const r = await listTeamChannelMembers(
        auth.token,
        teamId,
        channelId,
        top && top > 0 ? top : undefined
      );
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
  .description(
    'List tabs in a channel (GET …/channels/{id}/tabs?$expand=teamsApp; Group.ReadWrite.All or TeamsTab.Read.All)'
  )
  .argument('<teamId>', 'Team id')
  .argument('<channelId>', 'Channel id')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      teamId: string,
      channelId: string,
      opts: { json?: boolean; token?: string; identity?: string }
    ) => {
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
    }
  );

teamsCommand
  .command('messages')
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
  .command('message-replies')
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
  .command('members')
  .description('List members of a team (GET /teams/{id}/members; uses Group.ReadWrite.All or TeamMember.Read.*)')
  .argument('<teamId>', 'Team id')
  .option('-n, --top <n>', 'Page size (default: server default)', '')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (teamId: string, opts: { top?: string; json?: boolean; token?: string; identity?: string }) => {
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
    }
  );

teamsCommand
  .command('chats')
  .description('List chats for the signed-in user (GET /me/chats; requires Chat.Read)')
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
  .description(
    'Get one chat (GET /chats/{id}); optional --expand members, lastMessagePreview, or both (comma-separated)'
  )
  .argument('<chatId>', 'Chat id')
  .option(
    '--expand <segments>',
    'Graph $expand e.g. members or lastMessagePreview or members,lastMessagePreview'
  )
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      chatId: string,
      opts: { expand?: string; json?: boolean; token?: string; identity?: string }
    ) => {
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
    }
  );

teamsCommand
  .command('chat-messages')
  .description('List messages in a chat (GET /chats/{id}/messages; requires Chat.Read)')
  .argument('<chatId>', 'Chat id')
  .option('-n, --top <n>', 'Page size (max 50)', '10')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (chatId: string, opts: { top?: string; json?: boolean; token?: string; identity?: string }) => {
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
    }
  );

teamsCommand
  .command('chat-pinned')
  .description(
    'List pinned messages in a chat (GET /chats/{id}/pinnedMessages; Chat.Read; --expand-message for bodies)'
  )
  .argument('<chatId>', 'Chat id')
  .option('--expand-message', 'Include full message via $expand=message')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      chatId: string,
      opts: { expandMessage?: boolean; json?: boolean; token?: string; identity?: string }
    ) => {
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
          const preview = (p.message.body?.content ?? '')
            .replace(/\s+/g, ' ')
            .trim()
            .slice(0, 120);
          console.log(`${p.id ?? ''}\t${p.message.createdDateTime ?? ''}\t${who}\t${preview}`);
        } else {
          console.log(`${p.id ?? ''}\t\t\t(pinned id only; use --expand-message or --json)`);
        }
      }
    }
  );

teamsCommand
  .command('chat-members')
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
