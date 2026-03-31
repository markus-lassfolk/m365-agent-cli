import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import { type DateTimeTimeZone, getMailboxSettings, type OofStatus, setMailboxSettings } from '../lib/oof-client.js';
import { checkReadOnly } from '../lib/utils.js';

function formatStatus(status: OofStatus): string {
  switch (status) {
    case 'alwaysEnabled':
      return 'Always On';
    case 'scheduled':
      return 'Scheduled';
    case 'disabled':
      return 'Disabled';
    default:
      return status;
  }
}

function parseIsoDateTime(value: string): string {
  const d = new Date(value);
  if (!Number.isFinite(d.getTime())) {
    throw new Error(`Invalid ISO datetime: ${value}`);
  }
  return d.toISOString();
}

export const oofCommand = new Command('oof')
  .description('Get or set out-of-office (automatic reply) settings via Microsoft Graph')
  .option('--status <status>', 'OOF status: always, scheduled, disabled')
  .option('--internal-message <text>', 'Auto-reply message for internal senders')
  .option('--external-message <text>', 'Auto-reply message for external senders')
  .option('--start <datetime>', 'Scheduled start datetime (ISO 8601, e.g. 2025-12-01T00:00:00)')
  .option('--end <datetime>', 'Scheduled end datetime (ISO 8601, e.g. 2025-12-15T23:59:59)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user or shared mailbox (Graph delegation)')
  .action(async (options: any, cmd: any) => {
    const authResult = await resolveGraphAuth({ token: options.token, identity: options.identity });
    if (!authResult.success || !authResult.token) {
      const msg = authResult.error || 'Graph authentication failed';
      if (options.json) {
        console.log(JSON.stringify({ error: msg }, null, 2));
      } else {
        console.error(`Error: ${msg}`);
      }
      process.exit(1);
    }

    const token = authResult.token;

    // --- READ mode: no write options provided ---
    if (
      !options.status &&
      options.internalMessage === undefined &&
      options.externalMessage === undefined &&
      !options.start &&
      !options.end
    ) {
      const res = await getMailboxSettings(token, options.user);
      if (!res.ok || !res.data) {
        const msg = res.error?.message || 'Failed to retrieve mailbox settings';
        if (options.json) {
          console.log(JSON.stringify({ error: msg }, null, 2));
        } else {
          console.error(`Error: ${msg}`);
        }
        process.exit(1);
      }

      const oof = res.data.automaticRepliesSetting;
      if (!oof) {
        if (options.json) {
          console.log(JSON.stringify({ status: 'disabled', automaticRepliesSetting: null }, null, 2));
        } else {
          console.log('Out-of-office is disabled and no message is configured.');
        }
        return;
      }

      if (options.json) {
        console.log(
          JSON.stringify(
            {
              status: oof.status,
              scheduledStartDateTime: oof.scheduledStartDateTime ?? null,
              scheduledEndDateTime: oof.scheduledEndDateTime ?? null,
              internalReplyMessage: oof.internalReplyMessage ?? null,
              externalReplyMessage: oof.externalReplyMessage ?? null
            },
            null,
            2
          )
        );
      } else {
        console.log('Out-of-Office Settings:');
        console.log(`  Status: ${formatStatus(oof.status)}`);
        if (oof.status === 'scheduled') {
          const startStr = oof.scheduledStartDateTime
            ? typeof oof.scheduledStartDateTime === 'string'
              ? oof.scheduledStartDateTime
              : `${oof.scheduledStartDateTime.dateTime} (${oof.scheduledStartDateTime.timeZone})`
            : '?';
          const endStr = oof.scheduledEndDateTime
            ? typeof oof.scheduledEndDateTime === 'string'
              ? oof.scheduledEndDateTime
              : `${oof.scheduledEndDateTime.dateTime} (${oof.scheduledEndDateTime.timeZone})`
            : '?';
          console.log(`  Scheduled: ${startStr} → ${endStr}`);
        }
        if (oof.internalReplyMessage) {
          console.log(`\n  Internal Reply:\n    ${oof.internalReplyMessage}`);
        }
        if (oof.externalReplyMessage) {
          console.log(`\n  External Reply:\n    ${oof.externalReplyMessage}`);
        }
      }
      return;
    }

    // --- WRITE mode: validate inputs ---
    checkReadOnly(cmd);
    const errors: string[] = [];

    let status: OofStatus | undefined;
    if (options.status) {
      const raw = options.status.toLowerCase();
      if (raw === 'always' || raw === 'alwaysenabled') {
        status = 'alwaysEnabled';
      } else if (raw === 'scheduled') {
        status = 'scheduled';
      } else if (raw === 'disabled') {
        status = 'disabled';
      } else {
        errors.push('--status must be one of: always, scheduled, disabled');
      }
    }

    let scheduledStartDateTime: string | DateTimeTimeZone | undefined;
    let scheduledEndDateTime: string | DateTimeTimeZone | undefined;

    if (options.start) {
      try {
        scheduledStartDateTime = parseIsoDateTime(options.start);
      } catch {
        errors.push('--start must be a valid ISO 8601 datetime (e.g. 2025-12-01T00:00:00)');
      }
    }

    if (options.end) {
      try {
        scheduledEndDateTime = parseIsoDateTime(options.end);
      } catch {
        errors.push('--end must be a valid ISO 8601 datetime (e.g. 2025-12-15T23:59:59)');
      }
    }

    // If start/end are provided without explicit --status, default to scheduled
    if ((scheduledStartDateTime || scheduledEndDateTime) && !status) {
      status = 'scheduled';
    }

    if (errors.length > 0) {
      for (const e of errors) {
        if (options.json) {
          console.log(JSON.stringify({ error: e }, null, 2));
        } else {
          console.error(`Error: ${e}`);
        }
      }
      process.exit(1);
    }

    // Fetch existing settings if we are not explicitly overriding status,
    // because Graph requires status in the PATCH payload or resets it to disabled.
    let statusWasUndefined = false;
    if (!status && (options.internalMessage !== undefined || options.externalMessage !== undefined)) {
      statusWasUndefined = true;
      const currentRes = await getMailboxSettings(token, options.user);
      if (currentRes.ok && currentRes.data?.automaticRepliesSetting) {
        status = currentRes.data.automaticRepliesSetting.status;
        if (status === 'scheduled') {
          scheduledStartDateTime =
            scheduledStartDateTime ?? currentRes.data.automaticRepliesSetting.scheduledStartDateTime;
          scheduledEndDateTime = scheduledEndDateTime ?? currentRes.data.automaticRepliesSetting.scheduledEndDateTime;
        }
      } else {
        status = 'disabled'; // fallback if we couldn't fetch
      }
    }

    // --- Apply updates ---
    const patchResult = await setMailboxSettings(
      token,
      {
        status,
        internalReplyMessage: options.internalMessage,
        externalReplyMessage: options.externalMessage,
        scheduledStartDateTime,
        scheduledEndDateTime
      },
      options.user
    );

    if (!patchResult.ok) {
      const msg = patchResult.error?.message || 'Failed to update mailbox settings';
      if (options.json) {
        console.log(JSON.stringify({ error: msg }, null, 2));
      } else {
        console.error(`Error: ${msg}`);
      }
      process.exit(1);
    }

    if (options.json) {
      const normalizeDateTime = (dt: string | DateTimeTimeZone | undefined): string | null => {
        if (!dt) return null;
        if (typeof dt === 'string') return dt;
        return dt.dateTime;
      };

      const responseBody: any = {
        status: 'success',
        automaticRepliesSetting: {
          scheduledStartDateTime: normalizeDateTime(scheduledStartDateTime),
          scheduledEndDateTime: normalizeDateTime(scheduledEndDateTime),
          internalReplyMessage: options.internalMessage ?? null,
          externalReplyMessage: options.externalMessage ?? null
        }
      };
      if (status !== undefined) {
        responseBody.automaticRepliesSetting.status = status;
      }
      console.log(JSON.stringify(responseBody, null, 2));
    } else {
      console.log('Out-of-office settings updated.');
      if (statusWasUndefined && status !== undefined) {
        console.log(`  Status: ${formatStatus(status)} (unchanged)`);
      } else if (status !== undefined) {
        console.log(`  Status: ${formatStatus(status)}`);
      } else {
        console.log(`  Status: (unchanged)`);
      }
      if (status === 'scheduled' || (status === undefined && (scheduledStartDateTime || scheduledEndDateTime))) {
        if (scheduledStartDateTime) {
          const startStr =
            typeof scheduledStartDateTime === 'string'
              ? scheduledStartDateTime
              : `${scheduledStartDateTime.dateTime} (${scheduledStartDateTime.timeZone})`;
          console.log(`  Start: ${startStr}`);
        }
        if (scheduledEndDateTime) {
          const endStr =
            typeof scheduledEndDateTime === 'string'
              ? scheduledEndDateTime
              : `${scheduledEndDateTime.dateTime} (${scheduledEndDateTime.timeZone})`;
          console.log(`  End:   ${endStr}`);
        }
      }
      if (options.internalMessage) console.log(`  Internal message: ${options.internalMessage}`);
      if (options.externalMessage) console.log(`  External message: ${options.externalMessage}`);
    }
  });
