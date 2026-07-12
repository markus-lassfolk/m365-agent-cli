import { randomUUID } from 'node:crypto';
import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import {
  clearPreferredPresence,
  clearPresenceLocation,
  clearPresenceSession,
  getMyPresence,
  getPresencesByUserIds,
  getUserPresence,
  setPreferredPresence,
  setPresenceStatusMessage,
  setUserPresence
} from '../lib/graph-presence-client.js';
import { readJsonFileOrExit } from '../lib/read-json-file.js';
import { checkReadOnly } from '../lib/utils.js';

export const presenceCommand = new Command('presence').description(
  'User presence (Graph): read/bulk (`Presence.Read.All`), session set/clear (`Presence.ReadWrite`), **status message**, **preferred presence**, **clear location**. Subscriptions: **`subscribe`** + **`serve`** (see GRAPH_SCOPES.md).'
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
      const raw = await readJsonFileOrExit<unknown>(opts.jsonFile, '--json-file');
      idList = Array.isArray(raw)
        ? raw
            .map(String)
            .map((s) => s.trim())
            .filter(Boolean)
        : [];
    } else if (opts.ids?.trim()) {
      idList = opts.ids
        .split(',')
        .map((s) => s.trim())
        .filter(Boolean);
    } else {
      console.error('Error: provide --ids or --json-file with a JSON array of GUIDs');
      process.exit(1);
    }
    if (idList.length === 0) {
      console.error('Error: no user ids provided (list was empty after removing blanks)');
      process.exit(1);
    }
    const r = await getPresencesByUserIds(auth.token, idList);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(r.data, null, 2));
    } else {
      for (const row of r.data) {
        console.log(`${row.id ?? ''}\t${row.availability ?? ''}\t${row.activity ?? ''}`);
      }
    }
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

presenceCommand
  .command('status-message-set')
  .description(
    'Set Teams status message (`POST …/presence/setStatusMessage`; `Presence.ReadWrite`). Use `--json-file` for full Graph body, or `--text` with optional `--expiry-date-time` + `--timezone`.'
  )
  .option('--user <upn-or-id>', 'Target user (omit for /me)')
  .option(
    '--json-file <path>',
    'Full body e.g. { "statusMessage": { "message": { "contentType":"text", "content":"…" }, … } }'
  )
  .option('--text <s>', 'Plain status text (contentType text)')
  .option('--expiry-date-time <iso>', 'e.g. 2026-05-10T17:00:00 (used with --timezone)')
  .option('--timezone <iana-or-windows>', 'e.g. UTC or Pacific Standard Time')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      opts: {
        user?: string;
        jsonFile?: string;
        text?: string;
        expiryDateTime?: string;
        timezone?: string;
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
      if (opts.jsonFile?.trim()) {
        body = await readJsonFileOrExit(opts.jsonFile, '--json-file');
      } else if (opts.text?.trim()) {
        const statusMessage: Record<string, unknown> = {
          message: { contentType: 'text', content: opts.text.trim() }
        };
        if (opts.expiryDateTime?.trim() && opts.timezone?.trim()) {
          statusMessage.expiryDateTime = {
            dateTime: opts.expiryDateTime.trim(),
            timeZone: opts.timezone.trim()
          };
        } else if (opts.expiryDateTime?.trim() || opts.timezone?.trim()) {
          console.error('Error: use both --expiry-date-time and --timezone, or omit both');
          process.exit(1);
        }
        body = { statusMessage };
      } else {
        console.error('Error: provide --json-file or --text');
        process.exit(1);
      }
      const r = await setPresenceStatusMessage(auth.token, body, opts.user);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log('OK');
    }
  );

presenceCommand
  .command('preferred-set')
  .description(
    'Set preferred availability/activity (`POST …/presence/setUserPreferredPresence`). Effective when the user has a presence session (e.g. after **set-me** or Teams client). See Graph for allowed availability/activity pairs.'
  )
  .requiredOption('--availability <v>', 'e.g. Available, Busy, DoNotDisturb, Away, …')
  .requiredOption('--activity <v>', 'Must match Graph-supported pair for availability')
  .option('--expiration <iso8601duration>', 'e.g. PT8H')
  .option('--user <upn-or-id>', 'Target user (omit for /me)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      opts: {
        availability: string;
        activity: string;
        expiration?: string;
        user?: string;
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
      const payload: { availability: string; activity: string; expirationDuration?: string } = {
        availability: opts.availability.trim(),
        activity: opts.activity.trim()
      };
      if (opts.expiration?.trim()) payload.expirationDuration = opts.expiration.trim();
      const r = await setPreferredPresence(auth.token, payload, opts.user);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log('OK');
    }
  );

presenceCommand
  .command('preferred-clear')
  .description('Clear preferred presence (`POST …/presence/clearUserPreferredPresence`)')
  .option('--user <upn-or-id>', 'Target user (omit for /me)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { user?: string; token?: string; identity?: string }, cmd: Command) => {
    checkReadOnly(cmd);
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await clearPreferredPresence(auth.token, opts.user);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log('OK');
  });

presenceCommand
  .command('clear-location')
  .description(
    'Clear work-location signals for today (`POST …/presence/clearLocation`). See Microsoft Graph presence docs.'
  )
  .option('--user <upn-or-id>', 'Target user (omit for /me)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { user?: string; token?: string; identity?: string }, cmd: Command) => {
    checkReadOnly(cmd);
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await clearPresenceLocation(auth.token, opts.user);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log('OK');
  });
