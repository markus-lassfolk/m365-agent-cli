import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import { createOnlineMeeting, getOnlineMeeting } from '../lib/online-meetings-graph-client.js';
import { checkReadOnly } from '../lib/utils.js';

export const meetingCommand = new Command('meeting').description(
  'Teams online meetings via Graph POST /me/onlineMeetings (standalone join link). For calendar events with Teams, prefer create-event --teams (Calendars.ReadWrite + OnlineMeetings behavior differs).'
);

meetingCommand
  .command('create')
  .description('Create an online meeting and print join URL (requires OnlineMeetings.ReadWrite)')
  .requiredOption('--start <iso>', 'Start time (ISO 8601, e.g. 2026-04-03T14:00:00-07:00)')
  .requiredOption('--end <iso>', 'End time (ISO 8601)')
  .option('-s, --subject <text>', 'Meeting subject')
  .option('--json', 'Output full Graph JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
  .action(
    async (
      opts: {
        start: string;
        end: string;
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
      const r = await createOnlineMeeting(
        auth.token!,
        {
          startDateTime: opts.start,
          endDateTime: opts.end,
          subject: opts.subject
        },
        opts.user
      );
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
  .description('Get an online meeting by id')
  .argument('<meetingId>', 'Online meeting id')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
  .action(
    async (
      meetingId: string,
      opts: { json?: boolean; token?: string; identity?: string; user?: string }
    ) => {
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
    }
  );
