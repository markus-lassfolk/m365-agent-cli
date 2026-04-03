import { randomUUID } from 'node:crypto';
import { readFile } from 'node:fs/promises';
import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import {
  clearPresenceSession,
  getMyPresence,
  getPresencesByUserIds,
  getUserPresence,
  setUserPresence
} from '../lib/graph-presence-client.js';
import { checkReadOnly } from '../lib/utils.js';

export const presenceCommand = new Command('presence').description(
  'User presence (Graph): read/bulk (`Presence.Read.All`), set/clear session (`Presence.ReadWrite`); see GRAPH_SCOPES.md'
);

presenceCommand
  .command('me')
  .description('Get signed-in user presence (GET /me/presence)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await getMyPresence(auth.token);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log(opts.json ? JSON.stringify(r.data, null, 2) : `${r.data.availability ?? ''}\t${r.data.activity ?? ''}`);
  });

presenceCommand
  .command('user')
  .description('Get presence for a user (GET /users/{id|upn}/presence)')
  .argument('<user>', 'User id (GUID) or UPN')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (user: string, opts: { json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await getUserPresence(auth.token, user);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log(opts.json ? JSON.stringify(r.data, null, 2) : `${r.data.availability ?? ''}\t${r.data.activity ?? ''}`);
  });

presenceCommand
  .command('bulk')
  .description('Batch read presence (POST /communications/getPresencesByUserId; max 650 GUIDs; `Presence.Read.All`)')
  .option('--ids <csv>', 'Comma-separated Azure AD object ids (GUIDs)')
  .option('--json-file <path>', 'JSON array of user ids (overrides --ids)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { ids?: string; jsonFile?: string; json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    let idList: string[];
    if (opts.jsonFile?.trim()) {
      const raw = JSON.parse(await readFile(opts.jsonFile.trim(), 'utf-8')) as unknown;
      idList = Array.isArray(raw) ? (raw as string[]).map(String) : [];
    } else if (opts.ids?.trim()) {
      idList = opts.ids
        .split(',')
        .map((s) => s.trim())
        .filter(Boolean);
    } else {
      console.error('Error: provide --ids or --json-file with a JSON array of GUIDs');
      process.exit(1);
    }
    const r = await getPresencesByUserIds(auth.token, idList);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log(opts.json ? JSON.stringify(r.data, null, 2) : JSON.stringify(r.data, null, 2));
  });

presenceCommand
  .command('set-me')
  .description(
    'Set signed-in user presence (`POST /me/presence/setPresence`; `Presence.ReadWrite`). See Graph for availability/activity enums.'
  )
  .requiredOption('--availability <v>', 'e.g. Available, Busy, Away, DoNotDisturb')
  .requiredOption('--activity <v>', 'e.g. Available, InACall, InAConferenceCall, Away, …')
  .option('--expiration <iso8601duration>', 'Default PT8H', 'PT8H')
  .option('--session-id <guid>', 'Optional; default random UUID per request')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      opts: {
        availability: string;
        activity: string;
        expiration?: string;
        sessionId?: string;
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
      const payload = {
        sessionId: opts.sessionId?.trim() || randomUUID(),
        availability: opts.availability.trim(),
        activity: opts.activity.trim(),
        expirationDuration: opts.expiration?.trim() || 'PT8H'
      };
      const r = await setUserPresence(auth.token, payload);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(`OK\t${payload.sessionId}`);
    }
  );

presenceCommand
  .command('set-user')
  .description("Set another user's presence (`POST /users/{id}/presence/setPresence`; `Presence.ReadWrite`)")
  .argument('<user>', 'User id (GUID) or UPN')
  .requiredOption('--availability <v>', 'Availability value')
  .requiredOption('--activity <v>', 'Activity value')
  .option('--expiration <iso8601duration>', 'Default PT8H', 'PT8H')
  .option('--session-id <guid>', 'Optional; default random UUID')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      user: string,
      opts: {
        availability: string;
        activity: string;
        expiration?: string;
        sessionId?: string;
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
      const payload = {
        sessionId: opts.sessionId?.trim() || randomUUID(),
        availability: opts.availability.trim(),
        activity: opts.activity.trim(),
        expirationDuration: opts.expiration?.trim() || 'PT8H'
      };
      const r = await setUserPresence(auth.token, payload, user);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(`OK\t${payload.sessionId}`);
    }
  );

presenceCommand
  .command('clear-me')
  .description(
    'Clear a presence session (`POST /me/presence/clearPresence`). Use `--session-id` from **set-me** output (second column after OK).'
  )
  .requiredOption('--session-id <guid>', 'Session id from **presence set-me**')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { sessionId: string; token?: string; identity?: string }, cmd: Command) => {
    checkReadOnly(cmd);
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await clearPresenceSession(auth.token, opts.sessionId);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log('OK');
  });

presenceCommand
  .command('clear-user')
  .description('Clear a presence session for another user (`POST /users/{id}/presence/clearPresence`)')
  .argument('<user>', 'User id (GUID) or UPN')
  .requiredOption('--session-id <guid>', 'Session id to clear')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (user: string, opts: { sessionId: string; token?: string; identity?: string }, cmd: Command) => {
    checkReadOnly(cmd);
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await clearPresenceSession(auth.token, opts.sessionId, user);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log('OK');
  });
