import { Command } from 'commander';
import { resolveAuth } from '../lib/auth.js';
import {
  addDelegate,
  type DelegateInfo,
  type DelegatePermissions,
  type DeliverMeetingRequests,
  getDelegates,
  removeDelegate,
  updateDelegate
} from '../lib/delegate-client.js';
import { getExchangeBackend } from '../lib/exchange-backend.js';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import { type GraphCalendarPermission, listCalendarPermissions } from '../lib/graph-calendar-client.js';
import { checkReadOnly } from '../lib/utils.js';

const VALID_PERMISSIONS = [
  'None',
  'Owner',
  'PublishingEditor',
  'Editor',
  'PublishingAuthor',
  'Author',
  'Reviewer',
  'NonEditingAuthor',
  'FolderVisible'
] as const;
const VALID_FOLDERS = ['calendar', 'inbox', 'contacts', 'tasks', 'notes'] as const;
const VALID_DELIVER = ['DelegatesAndMe', 'DelegatesOnly', 'DelegatesAndSendInformationToMe', 'NoForward'] as const;

function formatPermissionLevel(level: string | undefined): string {
  return level ?? 'None';
}

function formatDelegate(delegate: DelegateInfo): string {
  const lines: string[] = [];
  const name = delegate.displayName || delegate.primaryEmail || delegate.userId;
  lines.push(`  ${name} <${delegate.userId}>`);
  lines.push(`    View private items: ${delegate.viewPrivateItems}`);
  lines.push(`    Deliver meeting requests: ${delegate.deliverMeetingRequests}`);

  const folderPerms = delegate.permissions;
  if (folderPerms.calendar) lines.push(`    Calendar:     ${formatPermissionLevel(folderPerms.calendar)}`);
  if (folderPerms.inbox) lines.push(`    Inbox:        ${formatPermissionLevel(folderPerms.inbox)}`);
  if (folderPerms.contacts) lines.push(`    Contacts:     ${formatPermissionLevel(folderPerms.contacts)}`);
  if (folderPerms.tasks) lines.push(`    Tasks:        ${formatPermissionLevel(folderPerms.tasks)}`);
  if (folderPerms.notes) lines.push(`    Notes:        ${formatPermissionLevel(folderPerms.notes)}`);

  return lines.join('\n');
}

function formatGraphCalendarPermission(p: GraphCalendarPermission): string {
  const addr = p.emailAddress?.address || p.emailAddress?.name || '(unknown)';
  const lines: string[] = [];
  lines.push(`  ${addr}`);
  if (p.role) lines.push(`    role: ${p.role}`);
  if (p.allowedRoles?.length) lines.push(`    allowedRoles: ${p.allowedRoles.join(', ')}`);
  if (p.isInsideOrganization !== undefined) lines.push(`    insideOrganization: ${p.isInsideOrganization}`);
  if (p.isRemovable !== undefined) lines.push(`    removable: ${p.isRemovable}`);
  return lines.join('\n');
}

function requireEwsForDelegateMutations(): void {
  if (getExchangeBackend() === 'graph') {
    console.error(
      'Error: add/update/remove delegates require EWS (EWS SOAP). Set M365_EXCHANGE_BACKEND=ews, or manage calendar sharing in Outlook.\n' +
        'See: https://learn.microsoft.com/en-us/graph/outlook-share-or-delegate-calendar'
    );
    process.exit(1);
  }
}

// ─── list ───

const listCommand = new Command('list');
listCommand
  .description('List calendar sharing permissions (Graph) or EWS delegates, per M365_EXCHANGE_BACKEND')
  .option('--mailbox <email>', 'mailbox (shared/alternative primary)')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Use a specific authentication identity (default: default)')
  .action(async (opts: { mailbox?: string; token?: string; identity?: string }) => {
    const backend = getExchangeBackend();

    if (backend === 'graph') {
      const ga = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!ga.success || !ga.token) {
        console.error('Auth failed:', ga.error);
        process.exit(1);
      }
      const result = await listCalendarPermissions(ga.token, opts.mailbox);
      if (!result.ok || !result.data) {
        console.error('Failed to list calendar permissions:', result.error?.message);
        process.exit(1);
      }
      const rows = result.data.filter((p) => {
        const role = (p.role || '').toLowerCase();
        return role !== 'freebusyread' && role !== 'none';
      });
      if (rows.length === 0) {
        console.log('No calendar permissions (Graph) — only free/busy or default entries.');
        return;
      }
      console.log(`Calendar permissions — Microsoft Graph (${rows.length}):\n`);
      for (const p of rows) {
        console.log(formatGraphCalendarPermission(p));
        console.log();
      }
      return;
    }

    if (backend === 'auto') {
      const ga = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (ga.success && ga.token) {
        const result = await listCalendarPermissions(ga.token, opts.mailbox);
        if (result.ok && result.data) {
          const rows = result.data.filter((p) => {
            const role = (p.role || '').toLowerCase();
            return role !== 'freebusyread' && role !== 'none';
          });
          if (rows.length > 0) {
            console.log(`Calendar permissions — Microsoft Graph (${rows.length}):\n`);
            for (const p of rows) {
              console.log(formatGraphCalendarPermission(p));
              console.log();
            }
            return;
          }
          // Graph succeeded: empty sharing list — do not substitute EWS delegates (different model).
          console.log('No calendar permissions (Graph) — only free/busy or default entries.');
          return;
        }
        if (!result.ok) {
          console.warn(`[delegates list] Graph failed (${result.error?.message}); falling back to EWS.`);
        }
      }
    }

    const auth = await resolveAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error('Auth failed:', auth.error);
      process.exit(1);
    }

    const result = await getDelegates(auth.token, opts.mailbox);
    if (!result.ok) {
      console.error('GetDelegates failed:', result.error?.message);
      process.exit(1);
    }

    const delegates = result.data ?? [];
    if (delegates.length === 0) {
      console.log('No delegates configured (EWS).');
      return;
    }

    console.log(`Delegates — EWS (${delegates.length}):\n`);
    for (const d of delegates) {
      console.log(formatDelegate(d));
      console.log();
    }
  });

// ─── add ───

const addCommand = new Command('add');
addCommand
  .description('Add a delegate with per-folder permissions')
  .requiredOption('--email <email>', 'delegate email address')
  .option('--name <name>', 'display name for the delegate')
  .option('--calendar <level>', `Calendar permission level (${VALID_PERMISSIONS.join('|')})`)
  .option('--inbox <level>', `Inbox permission level (${VALID_PERMISSIONS.join('|')})`)
  .option('--contacts <level>', `Contacts permission level (${VALID_PERMISSIONS.join('|')})`)
  .option('--tasks <level>', `Tasks permission level (${VALID_PERMISSIONS.join('|')})`)
  .option('--notes <level>', `Notes permission level (${VALID_PERMISSIONS.join('|')})`)
  .option('--view-private', 'allow delegate to view private items', false)
  .option('--deliver <mode>', `deliver meeting requests (${VALID_DELIVER.join('|')})`, 'DelegatesAndMe')
  .option('--mailbox <email>', 'mailbox to add delegate to (shared/alternative primary)')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Use a specific authentication identity (default: default)')
  .action(
    async (
      opts: {
        email: string;
        name?: string;
        calendar?: string;
        inbox?: string;
        contacts?: string;
        tasks?: string;
        notes?: string;
        viewPrivate?: boolean;
        deliver: string;
        mailbox?: string;
        token?: string;
        identity?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      requireEwsForDelegateMutations();
      // Validate permission levels
      const perms: DelegatePermissions = {};
      for (const folder of VALID_FOLDERS) {
        const key = folder as (typeof VALID_FOLDERS)[number];
        const level = opts[key] as string | undefined;
        if (level) {
          if (!VALID_PERMISSIONS.includes(level as (typeof VALID_PERMISSIONS)[number])) {
            console.error(`Invalid permission level "${level}" for ${folder}. Valid: ${VALID_PERMISSIONS.join(', ')}`);
            process.exit(1);
          }
          (perms as Record<string, string>)[key] = level;
        }
      }

      const deliver = opts.deliver as DeliverMeetingRequests;
      if (!VALID_DELIVER.includes(deliver)) {
        console.error(`Invalid deliver mode "${deliver}". Valid: ${VALID_DELIVER.join(', ')}`);
        process.exit(1);
      }

      const auth = await resolveAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error('Auth failed:', auth.error);
        process.exit(1);
      }

      const result = await addDelegate({
        token: auth.token,
        delegateEmail: opts.email,
        delegateName: opts.name,
        permissions: perms,
        viewPrivateItems: opts.viewPrivate,
        deliverMeetingRequests: deliver,
        mailbox: opts.mailbox
      });

      if (!result.ok) {
        console.error('AddDelegate failed:', result.error?.message);
        process.exit(1);
      }

      console.log('Delegate added:');
      console.log(formatDelegate(result.data!));
    }
  );

// ─── update ───

const updateCommand = new Command('update');
updateCommand
  .description("Update an existing delegate's permissions")
  .requiredOption('--email <email>', 'delegate email address')
  .option('--calendar <level>', `Calendar permission level (${VALID_PERMISSIONS.join('|')})`)
  .option('--inbox <level>', `Inbox permission level (${VALID_PERMISSIONS.join('|')})`)
  .option('--contacts <level>', `Contacts permission level (${VALID_PERMISSIONS.join('|')})`)
  .option('--tasks <level>', `Tasks permission level (${VALID_PERMISSIONS.join('|')})`)
  .option('--notes <level>', `Notes permission level (${VALID_PERMISSIONS.join('|')})`)
  .option('--view-private <boolean>', 'allow delegate to view private items (true/false)')
  .option('--deliver <mode>', `deliver meeting requests (${VALID_DELIVER.join('|')})`)
  .option('--mailbox <email>', 'mailbox (shared/alternative primary)')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Use a specific authentication identity (default: default)')
  .action(
    async (
      opts: {
        email: string;
        calendar?: string;
        inbox?: string;
        contacts?: string;
        tasks?: string;
        notes?: string;
        viewPrivate?: string | boolean;
        deliver?: string;
        mailbox?: string;
        token?: string;
        identity?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      requireEwsForDelegateMutations();
      const permsOut: DelegatePermissions = {};
      let hasPerms = false;

      for (const folder of VALID_FOLDERS) {
        const key = folder as (typeof VALID_FOLDERS)[number];
        const level = opts[key] as string | undefined;
        if (level !== undefined) {
          if (!VALID_PERMISSIONS.includes(level as (typeof VALID_PERMISSIONS)[number])) {
            console.error(`Invalid permission level "${level}" for ${folder}. Valid: ${VALID_PERMISSIONS.join(', ')}`);
            process.exit(1);
          }
          (permsOut as Record<string, string>)[key] = level;
          hasPerms = true;
        }
      }

      const deliver = opts.deliver as DeliverMeetingRequests | undefined;
      if (deliver && !VALID_DELIVER.includes(deliver)) {
        console.error(`Invalid deliver mode "${deliver}". Valid: ${VALID_DELIVER.join(', ')}`);
        process.exit(1);
      }

      const auth = await resolveAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error('Auth failed:', auth.error);
        process.exit(1);
      }

      const result = await updateDelegate({
        token: auth.token,
        delegateEmail: opts.email,
        permissions: hasPerms ? permsOut : undefined,
        viewPrivateItems:
          opts.viewPrivate === undefined ? undefined : opts.viewPrivate === 'true' || opts.viewPrivate === true,
        deliverMeetingRequests: deliver,
        mailbox: opts.mailbox
      });

      if (!result.ok) {
        console.error('UpdateDelegate failed:', result.error?.message);
        process.exit(1);
      }

      console.log('Delegate updated:');
      console.log(formatDelegate(result.data!));
    }
  );

// ─── remove ───

const removeCommand = new Command('remove');
removeCommand
  .description('Remove a delegate from the mailbox')
  .requiredOption('--email <email>', 'delegate email address')
  .option('--mailbox <email>', 'mailbox (shared/alternative primary)')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Use a specific authentication identity (default: default)')
  .action(async (opts: { email: string; mailbox?: string; token?: string; identity?: string }, cmd: any) => {
    checkReadOnly(cmd);
    requireEwsForDelegateMutations();
    const auth = await resolveAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error('Auth failed:', auth.error);
      process.exit(1);
    }

    const result = await removeDelegate({
      token: auth.token,
      delegateEmail: opts.email,
      mailbox: opts.mailbox
    });

    if (!result.ok) {
      console.error('RemoveDelegate failed:', result.error?.message);
      process.exit(1);
    }

    console.log(`Delegate ${opts.email} removed.`);
  });

// ─── Root ───

export const delegatesCommand = new Command('delegates');
delegatesCommand
  .description(
    'Delegates: list (Graph calendar permissions or EWS); add/update/remove (EWS — set M365_EXCHANGE_BACKEND=ews or auto)'
  )
  .addCommand(listCommand)
  .addCommand(addCommand)
  .addCommand(updateCommand)
  .addCommand(removeCommand);
