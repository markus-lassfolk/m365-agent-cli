import { readFile } from 'node:fs/promises';
import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import type { GraphResponse } from '../lib/graph-client.js';
import {
  createOnlineMeeting,
  createOnlineMeetingFromBody,
  deleteOnlineMeeting,
  getOnlineMeeting,
  type OnlineMeeting,
  updateOnlineMeeting
} from '../lib/online-meetings-graph-client.js';
import { checkReadOnly } from '../lib/utils.js';

export const meetingCommand = new Command('meeting').description(
  'Teams online meetings via Microsoft Graph (`OnlineMeetings.ReadWrite`). ' +
    '**Calendar invitations with Teams + attendees:** use `create-event ... --teams` (see `event.teamsMeeting` in `--json` output). ' +
    'This command is for **standalone** `POST /onlineMeetings` (join link without a calendar event, or advanced JSON).'
);

meetingCommand
  .command('create')
  .description(
    'Create an online meeting (`POST /me/onlineMeetings`). Use `--json-file` for full Graph body (participants, lobby, etc.).'
  )
  .option('--json-file <path>', 'Full JSON body (overrides --start/--end/--subject)')
  .option('--start <iso>', 'Start time (ISO 8601, e.g. 2026-04-03T14:00:00-07:00)')
  .option('--end <iso>', 'End time (ISO 8601)')
  .option('-s, --subject <text>', 'Meeting subject')
  .option('--json', 'Output full Graph JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation; same tenant Teams)')
  .action(
    async (
      opts: {
        jsonFile?: string;
        start?: string;
        end?: string;
        subject?: string;
        json?: boolean;
        token?: string;
        identity?: string;
        user?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }

      let r: GraphResponse<OnlineMeeting>;
      if (opts.jsonFile?.trim()) {
        const raw = await readFile(opts.jsonFile.trim(), 'utf-8');
        const body = JSON.parse(raw) as Record<string, unknown>;
        r = await createOnlineMeetingFromBody(auth.token!, body, opts.user);
      } else {
        if (!opts.start?.trim() || !opts.end?.trim()) {
          console.error('Error: provide --start and --end, or use --json-file with a full Graph body.');
          process.exit(1);
        }
        r = await createOnlineMeeting(
          auth.token!,
          {
            startDateTime: opts.start.trim(),
            endDateTime: opts.end.trim(),
            subject: opts.subject
          },
          opts.user
        );
      }

      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      const join = r.data.joinWebUrl ?? r.data.joinUrl;
      if (opts.json) {
        console.log(JSON.stringify(r.data, null, 2));
        return;
      }
      if (r.data.subject) console.log(`Subject: ${r.data.subject}`);
      if (join) console.log(`Join: ${join}`);
      else console.log(JSON.stringify(r.data, null, 2));
      if (r.data.id) console.log(`Meeting id: ${r.data.id}`);
    }
  );

meetingCommand
  .command('get')
  .description('Get an online meeting by id (`GET /me/onlineMeetings/{id}`)')
  .argument('<meetingId>', 'Online meeting id')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
  .action(async (meetingId: string, opts: { json?: boolean; token?: string; identity?: string; user?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await getOnlineMeeting(auth.token!, meetingId, opts.user);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else {
      const join = r.data.joinWebUrl ?? r.data.joinUrl;
      if (r.data.subject) console.log(`Subject: ${r.data.subject}`);
      if (join) console.log(`Join: ${join}`);
      if (r.data.id) console.log(`Meeting id: ${r.data.id}`);
    }
  });

meetingCommand
  .command('update')
  .description('Update an online meeting (`PATCH /me/onlineMeetings/{id}`)')
  .argument('<meetingId>', 'Online meeting id')
  .requiredOption('--json-file <path>', 'JSON patch body per Graph')
  .option('--json', 'Echo updated meeting as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
  .action(
    async (
      meetingId: string,
      opts: { jsonFile: string; json?: boolean; token?: string; identity?: string; user?: string },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const raw = await readFile(opts.jsonFile, 'utf-8');
      const patch = JSON.parse(raw) as Record<string, unknown>;
      const r = await updateOnlineMeeting(auth.token!, meetingId, patch, opts.user);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else {
        const join = r.data.joinWebUrl ?? r.data.joinUrl;
        if (join) console.log(`Join: ${join}`);
        console.log(`Updated meeting: ${r.data.id ?? meetingId}`);
      }
    }
  );

meetingCommand
  .command('delete')
  .description('Delete an online meeting (`DELETE /me/onlineMeetings/{id}`)')
  .argument('<meetingId>', 'Online meeting id')
  .option('--confirm', 'Confirm delete')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
  .action(
    async (
      meetingId: string,
      opts: { confirm?: boolean; token?: string; identity?: string; user?: string },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      if (!opts.confirm) {
        console.error('Refusing to delete without --confirm');
        process.exit(1);
      }
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await deleteOnlineMeeting(auth.token!, meetingId, opts.user);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log('Online meeting deleted.');
    }
  );
