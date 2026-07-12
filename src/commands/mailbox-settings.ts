import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import {
  getMailboxSettingsFull,
  normalizeWorkingHourTime,
  parseWorkDaysCsv,
  patchMailboxSettings,
  readJsonPatchFile
} from '../lib/mailbox-settings-client.js';
import { checkReadOnly } from '../lib/utils.js';

function formatWorkingHours(w: {
  daysOfWeek?: string[];
  startTime?: string;
  endTime?: string;
  timeZone?: { name?: string };
}): string {
  const days = w.daysOfWeek?.join(', ') || '(not set)';
  const tz = w.timeZone?.name || '';
  return `${days}  ${w.startTime || '?'}–${w.endTime || '?'}${tz ? `  (${tz})` : ''}`;
}

export const mailboxSettingsCommand = new Command('mailbox-settings')
  .description('Read or update Microsoft Graph mailboxSettings (time zone, working hours, regional formats)')
  .option('--json', 'Output full GET response as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target mailbox (delegation)')
  .action(async (opts: { json?: boolean; token?: string; identity?: string; user?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await getMailboxSettingsFull(auth.token, opts.user);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message || 'Failed to read mailbox settings'}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(r.data, null, 2));
      return;
    }
    const d = r.data;
    console.log('\nMailbox settings (Graph)');
    console.log('─'.repeat(40));
    if (d.timeZone) console.log(`  Time zone:     ${d.timeZone}`);
    if (d.dateFormat) console.log(`  Date format:   ${d.dateFormat}`);
    if (d.timeFormat) console.log(`  Time format:   ${d.timeFormat}`);
    if (d.workingHours) {
      console.log(`  Working hours: ${formatWorkingHours(d.workingHours)}`);
    }
    if (d.archiveFolder) console.log(`  Archive folder id: ${d.archiveFolder}`);
    console.log('\nUse `mailbox-settings set` to patch. Use `--json` for the full payload.\n');
  });

mailboxSettingsCommand
  .command('set')
  .description('PATCH mailboxSettings (time zone, working hours, or custom JSON body)')
  .option('--timezone <name>', 'Mailbox time zone (e.g. "Pacific Standard Time")')
  .option('--work-days <csv>', 'Working days: mon,tue,wed,thu,fri,sat,sun (comma-separated)')
  .option('--work-start <HH:mm>', 'Working hours start (local to working hours time zone)')
  .option('--work-end <HH:mm>', 'Working hours end')
  .option(
    '--work-timezone <name>',
    'IANA or Windows-style name inside workingHours.timeZone.name (defaults to --timezone if set)'
  )
  .option(
    '--json-file <path>',
    'JSON object merged into PATCH body (advanced; overrides single-field flags when keys overlap)'
  )
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target mailbox (delegation)')
  .action(
    async (
      opts: {
        timezone?: string;
        workDays?: string;
        workStart?: string;
        workEnd?: string;
        workTimezone?: string;
        jsonFile?: string;
        token?: string;
        identity?: string;
        user?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }

      const patch: Record<string, unknown> = {};

      if (opts.timezone?.trim()) {
        patch.timeZone = opts.timezone.trim();
      }

      const hasWork =
        opts.workDays?.trim() || opts.workStart?.trim() || opts.workEnd?.trim() || opts.workTimezone?.trim();
      if (hasWork) {
        const cur = await getMailboxSettingsFull(auth.token, opts.user);
        if (!cur.ok) {
          // workingHours is a nested object Graph replaces wholesale on PATCH, not merges — if we
          // can't confirm what's already there, silently treating the fetch failure as "nothing
          // configured" would let a partial --work-* update (e.g. only --work-start) wipe out the
          // existing days/end-time/timezone instead of merging with them.
          console.error(
            `Error: could not fetch existing working hours to merge with --work-* flags: ${cur.error?.message || 'unknown error'}`
          );
          process.exit(1);
        }
        const existing = cur.data?.workingHours ? { ...cur.data.workingHours } : {};
        if (opts.workDays?.trim()) {
          const days = parseWorkDaysCsv(opts.workDays);
          if (days.length === 0) {
            console.error('Error: --work-days had no valid day tokens (use mon,tue,…).');
            process.exit(1);
          }
          existing.daysOfWeek = days;
        }
        if (opts.workStart?.trim()) {
          existing.startTime = normalizeWorkingHourTime(opts.workStart);
        }
        if (opts.workEnd?.trim()) {
          existing.endTime = normalizeWorkingHourTime(opts.workEnd);
        }
        const wtz = opts.workTimezone?.trim() || opts.timezone?.trim();
        if (wtz) {
          existing.timeZone = { name: wtz };
        }
        patch.workingHours = existing;
      }

      if (opts.jsonFile?.trim()) {
        try {
          Object.assign(patch, await readJsonPatchFile(opts.jsonFile.trim()));
        } catch (err) {
          console.error(err instanceof Error ? err.message : String(err));
          process.exit(1);
        }
      }

      if (Object.keys(patch).length === 0) {
        console.error('Error: provide --timezone, --work-*, and/or --json-file');
        process.exit(1);
      }

      const pr = await patchMailboxSettings(auth.token, patch, opts.user);
      if (!pr.ok) {
        console.error(`Error: ${pr.error?.message || 'PATCH failed'}`);
        process.exit(1);
      }
      console.log('Done.');
    }
  );
