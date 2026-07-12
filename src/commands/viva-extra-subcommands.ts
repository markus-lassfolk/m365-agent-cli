import { readFile } from 'node:fs/promises';
import type { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import {
  buildVivaListQuery,
  createOnlineMeetingEngagementConversation,
  deleteAdminPeopleItemInsights,
  deleteOnlineMeetingEngagementConversation,
  deleteOrganizationItemInsights,
  deleteWorkHoursOccurrence,
  deleteWorkHoursRecurrence,
  getAdminPeopleItemInsights,
  getOnlineMeetingEngagementConversation,
  getOnlineMeetingEngagementConversationMessage,
  getOrganizationItemInsights,
  getWorkHoursAndLocations,
  getWorkHoursOccurrence,
  getWorkHoursOccurrencesView,
  getWorkHoursRecurrence,
  listOnlineMeetingEngagementConversationMessages,
  listOnlineMeetingEngagementConversations,
  listWorkHoursOccurrences,
  listWorkHoursRecurrences,
  patchAdminPeopleItemInsights,
  patchOnlineMeetingEngagementConversation,
  patchOrganizationItemInsights,
  patchWorkHoursAndLocations,
  patchWorkHoursOccurrence,
  patchWorkHoursRecurrence,
  postWorkHoursSetCurrentLocation
} from '../lib/graph-viva-client.js';
import {
  createOnlineMeetingConversationMessage,
  createOnlineMeetingConversationMessageReaction,
  createOnlineMeetingConversationMessageReply,
  createOnlineMeetingConversationMessageReplyReaction,
  deleteOnlineMeetingConversationMessage,
  deleteOnlineMeetingConversationMessageReaction,
  deleteOnlineMeetingConversationMessageReply,
  deleteOnlineMeetingConversationMessageReplyReaction,
  getOnlineMeetingConversationMessageConversation,
  getOnlineMeetingConversationMessageReaction,
  getOnlineMeetingConversationMessageReply,
  getOnlineMeetingConversationMessageReplyConversation,
  getOnlineMeetingConversationMessageReplyReaction,
  getOnlineMeetingConversationMessageReplyReplyTo,
  getOnlineMeetingConversationMessageReplyTo,
  getOnlineMeetingConversationOnlineMeeting,
  listOnlineMeetingConversationMessageReactions,
  listOnlineMeetingConversationMessageReplies,
  listOnlineMeetingConversationMessageReplyReactions,
  patchOnlineMeetingConversationMessage,
  patchOnlineMeetingConversationMessageReaction,
  patchOnlineMeetingConversationMessageReply,
  patchOnlineMeetingConversationMessageReplyReaction
} from '../lib/graph-viva-meeting-engage-deep.js';
import { toJsonError } from '../lib/json-error.js';
import { checkReadOnly } from '../lib/utils.js';

async function readJsonBody(path: string): Promise<unknown> {
  const raw = await readFile(path.trim(), 'utf8');
  return JSON.parse(raw) as unknown;
}

/**
 * Prints the `--json` structured error envelope (or the matching plain-text "Auth error: ..." /
 * "Error: ...") for the two failure shapes the ~7 `--json`-supporting list subcommands in this
 * file hit — an auth failure from `resolveGraphAuth` (a plain string) or a Graph API failure
 * (`GraphResponse.error`, a GraphError-shaped object) — then exits 1. Without this, a `--json`
 * list call that fails printed plain text on stderr instead of `{ error: {...} } ` on stdout like
 * every other --json-supporting command, leaving an agent piping the output nothing valid to parse.
 */
function failVivaExtra(
  json: boolean | undefined,
  prefix: 'Auth error' | 'Error',
  error: unknown,
  fallbackMessage?: string
): never {
  if (json) {
    console.log(JSON.stringify({ error: toJsonError(error, fallbackMessage) }, null, 2));
  } else {
    const message =
      (typeof error === 'string' ? error : (error as { message?: string } | undefined)?.message) ?? fallbackMessage;
    console.error(`${prefix}: ${message}`);
  }
  process.exit(1);
}

type ODataOpts = {
  filter?: string;
  select?: string;
  top?: number;
  skip?: number;
  count?: boolean;
};

function odataQ(o: ODataOpts): string {
  return buildVivaListQuery({
    filter: o.filter,
    select: o.select,
    top: o.top,
    skip: o.skip,
    count: o.count
  });
}

function addODataListOpts(cmd: Command): Command {
  return cmd
    .option('--filter <odata>', 'OData $filter')
    .option('--select <fields>', 'OData $select (comma-separated)')
    .option('--top <n>', 'OData $top', (v) => parseInt(v, 10))
    .option('--skip <n>', 'OData $skip', (v) => parseInt(v, 10))
    .option('--count', 'Include $count=true');
}

/** Admin/org insights, work hours & locations, Viva Engage meeting conversations (beta). */
export function registerVivaExtraSubcommands(viva: Command): void {
  viva
    .command('admin-item-insights-get')
    .description('GET /admin/people/itemInsights (beta; org admin)')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(async (opts: { token?: string; identity?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getAdminPeopleItemInsights(auth.token);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  viva
    .command('admin-item-insights-patch')
    .description('PATCH /admin/people/itemInsights (beta)')
    .requiredOption('--body-file <path>', 'JSON object for PATCH body')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(async (opts: { bodyFile: string; token?: string; identity?: string }, cmd: Command) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      let body: unknown;
      try {
        body = await readJsonBody(opts.bodyFile);
      } catch (e) {
        console.error(e instanceof Error ? e.message : 'Invalid --body-file');
        process.exit(1);
      }
      const r = await patchAdminPeopleItemInsights(auth.token, body);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  viva
    .command('admin-item-insights-delete')
    .description('DELETE /admin/people/itemInsights (beta)')
    .option('--if-match <etag>', 'If-Match header value')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(async (opts: { ifMatch?: string; token?: string; identity?: string }, cmd: Command) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await deleteAdminPeopleItemInsights(auth.token, opts.ifMatch);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
    });

  viva
    .command('org-item-insights-get')
    .description('GET /organization/{id}/settings/itemInsights (beta)')
    .requiredOption('--organization-id <id>', 'Organization id (tenant id)')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(async (opts: { organizationId: string; token?: string; identity?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getOrganizationItemInsights(auth.token, opts.organizationId);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  viva
    .command('org-item-insights-patch')
    .description('PATCH /organization/{id}/settings/itemInsights (beta)')
    .requiredOption('--organization-id <id>', 'Organization id')
    .requiredOption('--body-file <path>', 'JSON object for PATCH body')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(
      async (opts: { organizationId: string; bodyFile: string; token?: string; identity?: string }, cmd: Command) => {
        checkReadOnly(cmd);
        const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
        if (!auth.success || !auth.token) {
          console.error(`Auth error: ${auth.error}`);
          process.exit(1);
        }
        let body: unknown;
        try {
          body = await readJsonBody(opts.bodyFile);
        } catch (e) {
          console.error(e instanceof Error ? e.message : 'Invalid --body-file');
          process.exit(1);
        }
        const r = await patchOrganizationItemInsights(auth.token, opts.organizationId, body);
        if (!r.ok) {
          console.error(`Error: ${r.error?.message}`);
          process.exit(1);
        }
        console.log(JSON.stringify(r.data ?? null, null, 2));
      }
    );

  viva
    .command('org-item-insights-delete')
    .description('DELETE /organization/{id}/settings/itemInsights (beta)')
    .requiredOption('--organization-id <id>', 'Organization id')
    .option('--if-match <etag>', 'If-Match header value')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(
      async (opts: { organizationId: string; ifMatch?: string; token?: string; identity?: string }, cmd: Command) => {
        checkReadOnly(cmd);
        const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
        if (!auth.success || !auth.token) {
          console.error(`Auth error: ${auth.error}`);
          process.exit(1);
        }
        const r = await deleteOrganizationItemInsights(auth.token, opts.organizationId, opts.ifMatch);
        if (!r.ok) {
          console.error(`Error: ${r.error?.message}`);
          process.exit(1);
        }
      }
    );

  viva
    .command('work-hours-get')
    .description('GET /me|/users/{id}/settings/workHoursAndLocations (beta)')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .option('--user <id>', 'Target user id or UPN (omit for /me)')
    .action(async (opts: { token?: string; identity?: string; user?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getWorkHoursAndLocations(auth.token, opts.user);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  viva
    .command('work-hours-patch')
    .description('PATCH workHoursAndLocations (beta)')
    .requiredOption('--body-file <path>', 'JSON object for PATCH body')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .option('--user <id>', 'Target user id or UPN (omit for /me)')
    .action(async (opts: { bodyFile: string; token?: string; identity?: string; user?: string }, cmd: Command) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      let body: unknown;
      try {
        body = await readJsonBody(opts.bodyFile);
      } catch (e) {
        console.error(e instanceof Error ? e.message : 'Invalid --body-file');
        process.exit(1);
      }
      const r = await patchWorkHoursAndLocations(auth.token, body, opts.user);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  addODataListOpts(
    viva
      .command('work-hours-occurrences-list')
      .description('List workHoursAndLocations/occurrences (beta, paged)')
      .option('--json', 'Output as JSON array')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity (default: default)')
      .option('--user <id>', 'Target user id or UPN (omit for /me)')
  ).action(async (opts: ODataOpts & { json?: boolean; token?: string; identity?: string; user?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      failVivaExtra(opts.json, 'Auth error', auth.error);
    }
    const r = await listWorkHoursOccurrences(auth.token, opts.user, odataQ(opts));
    if (!r.ok || !r.data) {
      failVivaExtra(opts.json, 'Error', r.error, 'List failed');
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else for (const row of r.data) console.log(JSON.stringify(row));
  });

  viva
    .command('work-hours-occurrence-get')
    .description('GET workPlanOccurrence by id (beta)')
    .requiredOption('--occurrence-id <id>', 'workPlanOccurrence id')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .option('--user <id>', 'Target user id or UPN (omit for /me)')
    .action(async (opts: { occurrenceId: string; token?: string; identity?: string; user?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getWorkHoursOccurrence(auth.token, opts.occurrenceId, opts.user);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  viva
    .command('work-hours-occurrence-patch')
    .description('PATCH workPlanOccurrence (beta)')
    .requiredOption('--occurrence-id <id>', 'workPlanOccurrence id')
    .requiredOption('--body-file <path>', 'JSON object for PATCH body')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .option('--user <id>', 'Target user id or UPN (omit for /me)')
    .action(
      async (
        opts: { occurrenceId: string; bodyFile: string; token?: string; identity?: string; user?: string },
        cmd: Command
      ) => {
        checkReadOnly(cmd);
        const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
        if (!auth.success || !auth.token) {
          console.error(`Auth error: ${auth.error}`);
          process.exit(1);
        }
        let body: unknown;
        try {
          body = await readJsonBody(opts.bodyFile);
        } catch (e) {
          console.error(e instanceof Error ? e.message : 'Invalid --body-file');
          process.exit(1);
        }
        const r = await patchWorkHoursOccurrence(auth.token, opts.occurrenceId, body, opts.user);
        if (!r.ok) {
          console.error(`Error: ${r.error?.message}`);
          process.exit(1);
        }
        console.log(JSON.stringify(r.data ?? null, null, 2));
      }
    );

  viva
    .command('work-hours-occurrence-delete')
    .description('DELETE workPlanOccurrence (beta)')
    .requiredOption('--occurrence-id <id>', 'workPlanOccurrence id')
    .option('--if-match <etag>', 'If-Match header value')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .option('--user <id>', 'Target user id or UPN (omit for /me)')
    .action(
      async (
        opts: { occurrenceId: string; ifMatch?: string; token?: string; identity?: string; user?: string },
        cmd: Command
      ) => {
        checkReadOnly(cmd);
        const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
        if (!auth.success || !auth.token) {
          console.error(`Auth error: ${auth.error}`);
          process.exit(1);
        }
        const r = await deleteWorkHoursOccurrence(auth.token, opts.occurrenceId, opts.user, opts.ifMatch);
        if (!r.ok) {
          console.error(`Error: ${r.error?.message}`);
          process.exit(1);
        }
      }
    );

  viva
    .command('work-hours-set-current-location')
    .description('POST .../occurrences/setCurrentLocation (beta; optional --body-file, default {})')
    .option('--body-file <path>', 'JSON POST body (optional)')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .option('--user <id>', 'Target user id or UPN (omit for /me)')
    .action(async (opts: { bodyFile?: string; token?: string; identity?: string; user?: string }, cmd: Command) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      let body: unknown = {};
      if (opts.bodyFile?.trim()) {
        try {
          body = await readJsonBody(opts.bodyFile);
        } catch (e) {
          console.error(e instanceof Error ? e.message : 'Invalid --body-file');
          process.exit(1);
        }
      }
      const r = await postWorkHoursSetCurrentLocation(auth.token, opts.user, body);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  viva
    .command('work-hours-occurrences-view')
    .description('GET occurrencesView(startDateTime,endDateTime) (beta; ISO 8601 strings, OData-escaped)')
    .requiredOption('--start <iso>', 'startDateTime e.g. 2025-01-01T00:00:00Z')
    .requiredOption('--end <iso>', 'endDateTime e.g. 2025-01-07T23:59:59Z')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .option('--user <id>', 'Target user id or UPN (omit for /me)')
    .action(async (opts: { start: string; end: string; token?: string; identity?: string; user?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getWorkHoursOccurrencesView(auth.token, opts.start, opts.end, opts.user);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  addODataListOpts(
    viva
      .command('work-hours-recurrences-list')
      .description('List workHoursAndLocations/recurrences (beta, paged)')
      .option('--json', 'Output as JSON array')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity (default: default)')
      .option('--user <id>', 'Target user id or UPN (omit for /me)')
  ).action(async (opts: ODataOpts & { json?: boolean; token?: string; identity?: string; user?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      failVivaExtra(opts.json, 'Auth error', auth.error);
    }
    const r = await listWorkHoursRecurrences(auth.token, opts.user, odataQ(opts));
    if (!r.ok || !r.data) {
      failVivaExtra(opts.json, 'Error', r.error, 'List failed');
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else for (const row of r.data) console.log(JSON.stringify(row));
  });

  viva
    .command('work-hours-recurrence-get')
    .description('GET workPlanRecurrence by id (beta)')
    .requiredOption('--recurrence-id <id>', 'workPlanRecurrence id')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .option('--user <id>', 'Target user id or UPN (omit for /me)')
    .action(async (opts: { recurrenceId: string; token?: string; identity?: string; user?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getWorkHoursRecurrence(auth.token, opts.recurrenceId, opts.user);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  viva
    .command('work-hours-recurrence-patch')
    .description('PATCH workPlanRecurrence (beta)')
    .requiredOption('--recurrence-id <id>', 'workPlanRecurrence id')
    .requiredOption('--body-file <path>', 'JSON object for PATCH body')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .option('--user <id>', 'Target user id or UPN (omit for /me)')
    .action(
      async (
        opts: { recurrenceId: string; bodyFile: string; token?: string; identity?: string; user?: string },
        cmd: Command
      ) => {
        checkReadOnly(cmd);
        const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
        if (!auth.success || !auth.token) {
          console.error(`Auth error: ${auth.error}`);
          process.exit(1);
        }
        let body: unknown;
        try {
          body = await readJsonBody(opts.bodyFile);
        } catch (e) {
          console.error(e instanceof Error ? e.message : 'Invalid --body-file');
          process.exit(1);
        }
        const r = await patchWorkHoursRecurrence(auth.token, opts.recurrenceId, body, opts.user);
        if (!r.ok) {
          console.error(`Error: ${r.error?.message}`);
          process.exit(1);
        }
        console.log(JSON.stringify(r.data ?? null, null, 2));
      }
    );

  viva
    .command('work-hours-recurrence-delete')
    .description('DELETE workPlanRecurrence (beta)')
    .requiredOption('--recurrence-id <id>', 'workPlanRecurrence id')
    .option('--if-match <etag>', 'If-Match header value')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .option('--user <id>', 'Target user id or UPN (omit for /me)')
    .action(
      async (
        opts: { recurrenceId: string; ifMatch?: string; token?: string; identity?: string; user?: string },
        cmd: Command
      ) => {
        checkReadOnly(cmd);
        const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
        if (!auth.success || !auth.token) {
          console.error(`Auth error: ${auth.error}`);
          process.exit(1);
        }
        const r = await deleteWorkHoursRecurrence(auth.token, opts.recurrenceId, opts.user, opts.ifMatch);
        if (!r.ok) {
          console.error(`Error: ${r.error?.message}`);
          process.exit(1);
        }
      }
    );

  addODataListOpts(
    viva
      .command('meeting-engage-conversations-list')
      .description('List /communications/onlineMeetingConversations (beta, paged)')
      .option('--json', 'Output as JSON array')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity (default: default)')
  ).action(async (opts: ODataOpts & { json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      failVivaExtra(opts.json, 'Auth error', auth.error);
    }
    const r = await listOnlineMeetingEngagementConversations(auth.token, odataQ(opts));
    if (!r.ok || !r.data) {
      failVivaExtra(opts.json, 'Error', r.error, 'List failed');
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else for (const row of r.data) console.log(JSON.stringify(row));
  });

  viva
    .command('meeting-engage-conversation-create')
    .description('POST onlineMeetingConversation (beta)')
    .requiredOption('--body-file <path>', 'JSON object for POST body')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(async (opts: { bodyFile: string; token?: string; identity?: string }, cmd: Command) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      let body: unknown;
      try {
        body = await readJsonBody(opts.bodyFile);
      } catch (e) {
        console.error(e instanceof Error ? e.message : 'Invalid --body-file');
        process.exit(1);
      }
      const r = await createOnlineMeetingEngagementConversation(auth.token, body);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  viva
    .command('meeting-engage-conversation-get')
    .description('GET onlineMeetingEngagementConversation by id (beta)')
    .requiredOption('--conversation-id <id>', 'Conversation id')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(async (opts: { conversationId: string; token?: string; identity?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getOnlineMeetingEngagementConversation(auth.token, opts.conversationId);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  viva
    .command('meeting-engage-conversation-patch')
    .description('PATCH onlineMeetingEngagementConversation (beta)')
    .requiredOption('--conversation-id <id>', 'Conversation id')
    .requiredOption('--body-file <path>', 'JSON object for PATCH body')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(
      async (opts: { conversationId: string; bodyFile: string; token?: string; identity?: string }, cmd: Command) => {
        checkReadOnly(cmd);
        const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
        if (!auth.success || !auth.token) {
          console.error(`Auth error: ${auth.error}`);
          process.exit(1);
        }
        let body: unknown;
        try {
          body = await readJsonBody(opts.bodyFile);
        } catch (e) {
          console.error(e instanceof Error ? e.message : 'Invalid --body-file');
          process.exit(1);
        }
        const r = await patchOnlineMeetingEngagementConversation(auth.token, opts.conversationId, body);
        if (!r.ok) {
          console.error(`Error: ${r.error?.message}`);
          process.exit(1);
        }
        console.log(JSON.stringify(r.data ?? null, null, 2));
      }
    );

  viva
    .command('meeting-engage-conversation-delete')
    .description('DELETE onlineMeetingEngagementConversation (beta)')
    .requiredOption('--conversation-id <id>', 'Conversation id')
    .option('--if-match <etag>', 'If-Match header value')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(
      async (opts: { conversationId: string; ifMatch?: string; token?: string; identity?: string }, cmd: Command) => {
        checkReadOnly(cmd);
        const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
        if (!auth.success || !auth.token) {
          console.error(`Auth error: ${auth.error}`);
          process.exit(1);
        }
        const r = await deleteOnlineMeetingEngagementConversation(auth.token, opts.conversationId, opts.ifMatch);
        if (!r.ok) {
          console.error(`Error: ${r.error?.message}`);
          process.exit(1);
        }
      }
    );

  addODataListOpts(
    viva
      .command('meeting-engage-conversation-messages-list')
      .description('List messages on a Viva Engage meeting conversation (beta, paged)')
      .requiredOption('--conversation-id <id>', 'Conversation id')
      .option('--json', 'Output as JSON array')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity (default: default)')
  ).action(async (opts: ODataOpts & { conversationId: string; json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      failVivaExtra(opts.json, 'Auth error', auth.error);
    }
    const r = await listOnlineMeetingEngagementConversationMessages(auth.token, opts.conversationId, odataQ(opts));
    if (!r.ok || !r.data) {
      failVivaExtra(opts.json, 'Error', r.error, 'List failed');
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else for (const row of r.data) console.log(JSON.stringify(row));
  });

  viva
    .command('meeting-engage-conversation-message-get')
    .description('GET single message on meeting conversation (beta)')
    .requiredOption('--conversation-id <id>', 'Conversation id')
    .requiredOption('--message-id <id>', 'Message id')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(async (opts: { conversationId: string; messageId: string; token?: string; identity?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getOnlineMeetingEngagementConversationMessage(auth.token, opts.conversationId, opts.messageId);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  viva
    .command('meeting-engage-conversation-message-create')
    .description('POST message on meeting conversation (beta)')
    .requiredOption('--conversation-id <id>', 'Conversation id')
    .requiredOption('--body-file <path>', 'JSON object for POST body')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(
      async (opts: { conversationId: string; bodyFile: string; token?: string; identity?: string }, cmd: Command) => {
        checkReadOnly(cmd);
        const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
        if (!auth.success || !auth.token) {
          console.error(`Auth error: ${auth.error}`);
          process.exit(1);
        }
        let body: unknown;
        try {
          body = await readJsonBody(opts.bodyFile);
        } catch (e) {
          console.error(e instanceof Error ? e.message : 'Invalid --body-file');
          process.exit(1);
        }
        const r = await createOnlineMeetingConversationMessage(auth.token, opts.conversationId, body);
        if (!r.ok) {
          console.error(`Error: ${r.error?.message}`);
          process.exit(1);
        }
        console.log(JSON.stringify(r.data ?? null, null, 2));
      }
    );

  viva
    .command('meeting-engage-conversation-message-patch')
    .description('PATCH message on meeting conversation (beta)')
    .requiredOption('--conversation-id <id>', 'Conversation id')
    .requiredOption('--message-id <id>', 'Message id')
    .requiredOption('--body-file <path>', 'JSON object for PATCH body')
    .option('--if-match <etag>', 'If-Match header value')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(
      async (
        opts: {
          conversationId: string;
          messageId: string;
          bodyFile: string;
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
        let body: unknown;
        try {
          body = await readJsonBody(opts.bodyFile);
        } catch (e) {
          console.error(e instanceof Error ? e.message : 'Invalid --body-file');
          process.exit(1);
        }
        const r = await patchOnlineMeetingConversationMessage(
          auth.token,
          opts.conversationId,
          opts.messageId,
          body,
          opts.ifMatch
        );
        if (!r.ok) {
          console.error(`Error: ${r.error?.message}`);
          process.exit(1);
        }
        console.log(JSON.stringify(r.data ?? null, null, 2));
      }
    );

  viva
    .command('meeting-engage-conversation-message-delete')
    .description('DELETE message on meeting conversation (beta)')
    .requiredOption('--conversation-id <id>', 'Conversation id')
    .requiredOption('--message-id <id>', 'Message id')
    .option('--if-match <etag>', 'If-Match header value')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(
      async (
        opts: { conversationId: string; messageId: string; ifMatch?: string; token?: string; identity?: string },
        cmd: Command
      ) => {
        checkReadOnly(cmd);
        const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
        if (!auth.success || !auth.token) {
          console.error(`Auth error: ${auth.error}`);
          process.exit(1);
        }
        const r = await deleteOnlineMeetingConversationMessage(
          auth.token,
          opts.conversationId,
          opts.messageId,
          opts.ifMatch
        );
        if (!r.ok) {
          console.error(`Error: ${r.error?.message}`);
          process.exit(1);
        }
      }
    );

  viva
    .command('meeting-engage-conversation-message-conversation-get')
    .description('GET conversation navigation from a message (beta)')
    .requiredOption('--conversation-id <id>', 'Conversation id')
    .requiredOption('--message-id <id>', 'Message id')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(async (opts: { conversationId: string; messageId: string; token?: string; identity?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getOnlineMeetingConversationMessageConversation(auth.token, opts.conversationId, opts.messageId);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  viva
    .command('meeting-engage-conversation-message-reply-to-get')
    .description('GET replyTo navigation on a message (beta)')
    .requiredOption('--conversation-id <id>', 'Conversation id')
    .requiredOption('--message-id <id>', 'Message id')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(async (opts: { conversationId: string; messageId: string; token?: string; identity?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getOnlineMeetingConversationMessageReplyTo(auth.token, opts.conversationId, opts.messageId);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  addODataListOpts(
    viva
      .command('meeting-engage-conversation-message-reactions-list')
      .description('List reactions on a meeting conversation message (beta, paged)')
      .requiredOption('--conversation-id <id>', 'Conversation id')
      .requiredOption('--message-id <id>', 'Message id')
      .option('--json', 'Output as JSON array')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity (default: default)')
  ).action(
    async (
      opts: ODataOpts & { conversationId: string; messageId: string; json?: boolean; token?: string; identity?: string }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        failVivaExtra(opts.json, 'Auth error', auth.error);
      }
      const r = await listOnlineMeetingConversationMessageReactions(
        auth.token,
        opts.conversationId,
        opts.messageId,
        odataQ(opts)
      );
      if (!r.ok || !r.data) {
        failVivaExtra(opts.json, 'Error', r.error, 'List failed');
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else for (const row of r.data) console.log(JSON.stringify(row));
    }
  );

  viva
    .command('meeting-engage-conversation-message-reaction-create')
    .description('POST reaction on a meeting conversation message (beta)')
    .requiredOption('--conversation-id <id>', 'Conversation id')
    .requiredOption('--message-id <id>', 'Message id')
    .requiredOption('--body-file <path>', 'JSON object for POST body')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(
      async (
        opts: { conversationId: string; messageId: string; bodyFile: string; token?: string; identity?: string },
        cmd: Command
      ) => {
        checkReadOnly(cmd);
        const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
        if (!auth.success || !auth.token) {
          console.error(`Auth error: ${auth.error}`);
          process.exit(1);
        }
        let body: unknown;
        try {
          body = await readJsonBody(opts.bodyFile);
        } catch (e) {
          console.error(e instanceof Error ? e.message : 'Invalid --body-file');
          process.exit(1);
        }
        const r = await createOnlineMeetingConversationMessageReaction(
          auth.token,
          opts.conversationId,
          opts.messageId,
          body
        );
        if (!r.ok) {
          console.error(`Error: ${r.error?.message}`);
          process.exit(1);
        }
        console.log(JSON.stringify(r.data ?? null, null, 2));
      }
    );

  viva
    .command('meeting-engage-conversation-message-reaction-get')
    .description('GET reaction on a meeting conversation message (beta)')
    .requiredOption('--conversation-id <id>', 'Conversation id')
    .requiredOption('--message-id <id>', 'Message id')
    .requiredOption('--reaction-id <id>', 'Reaction id')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(
      async (opts: {
        conversationId: string;
        messageId: string;
        reactionId: string;
        token?: string;
        identity?: string;
      }) => {
        const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
        if (!auth.success || !auth.token) {
          console.error(`Auth error: ${auth.error}`);
          process.exit(1);
        }
        const r = await getOnlineMeetingConversationMessageReaction(
          auth.token,
          opts.conversationId,
          opts.messageId,
          opts.reactionId
        );
        if (!r.ok) {
          console.error(`Error: ${r.error?.message}`);
          process.exit(1);
        }
        console.log(JSON.stringify(r.data ?? null, null, 2));
      }
    );

  viva
    .command('meeting-engage-conversation-message-reaction-patch')
    .description('PATCH reaction on a meeting conversation message (beta)')
    .requiredOption('--conversation-id <id>', 'Conversation id')
    .requiredOption('--message-id <id>', 'Message id')
    .requiredOption('--reaction-id <id>', 'Reaction id')
    .requiredOption('--body-file <path>', 'JSON object for PATCH body')
    .option('--if-match <etag>', 'If-Match header value')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(
      async (
        opts: {
          conversationId: string;
          messageId: string;
          reactionId: string;
          bodyFile: string;
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
        let body: unknown;
        try {
          body = await readJsonBody(opts.bodyFile);
        } catch (e) {
          console.error(e instanceof Error ? e.message : 'Invalid --body-file');
          process.exit(1);
        }
        const r = await patchOnlineMeetingConversationMessageReaction(
          auth.token,
          opts.conversationId,
          opts.messageId,
          opts.reactionId,
          body,
          opts.ifMatch
        );
        if (!r.ok) {
          console.error(`Error: ${r.error?.message}`);
          process.exit(1);
        }
        console.log(JSON.stringify(r.data ?? null, null, 2));
      }
    );

  viva
    .command('meeting-engage-conversation-message-reaction-delete')
    .description('DELETE reaction on a meeting conversation message (beta)')
    .requiredOption('--conversation-id <id>', 'Conversation id')
    .requiredOption('--message-id <id>', 'Message id')
    .requiredOption('--reaction-id <id>', 'Reaction id')
    .option('--if-match <etag>', 'If-Match header value')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(
      async (
        opts: {
          conversationId: string;
          messageId: string;
          reactionId: string;
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
        const r = await deleteOnlineMeetingConversationMessageReaction(
          auth.token,
          opts.conversationId,
          opts.messageId,
          opts.reactionId,
          opts.ifMatch
        );
        if (!r.ok) {
          console.error(`Error: ${r.error?.message}`);
          process.exit(1);
        }
      }
    );

  addODataListOpts(
    viva
      .command('meeting-engage-conversation-message-replies-list')
      .description('List replies on a meeting conversation message (beta, paged)')
      .requiredOption('--conversation-id <id>', 'Conversation id')
      .requiredOption('--message-id <id>', 'Message id')
      .option('--json', 'Output as JSON array')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity (default: default)')
  ).action(
    async (
      opts: ODataOpts & { conversationId: string; messageId: string; json?: boolean; token?: string; identity?: string }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        failVivaExtra(opts.json, 'Auth error', auth.error);
      }
      const r = await listOnlineMeetingConversationMessageReplies(
        auth.token,
        opts.conversationId,
        opts.messageId,
        odataQ(opts)
      );
      if (!r.ok || !r.data) {
        failVivaExtra(opts.json, 'Error', r.error, 'List failed');
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else for (const row of r.data) console.log(JSON.stringify(row));
    }
  );

  viva
    .command('meeting-engage-conversation-message-reply-create')
    .description('POST reply on a meeting conversation message (beta)')
    .requiredOption('--conversation-id <id>', 'Conversation id')
    .requiredOption('--message-id <id>', 'Parent message id')
    .requiredOption('--body-file <path>', 'JSON object for POST body')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(
      async (
        opts: { conversationId: string; messageId: string; bodyFile: string; token?: string; identity?: string },
        cmd: Command
      ) => {
        checkReadOnly(cmd);
        const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
        if (!auth.success || !auth.token) {
          console.error(`Auth error: ${auth.error}`);
          process.exit(1);
        }
        let body: unknown;
        try {
          body = await readJsonBody(opts.bodyFile);
        } catch (e) {
          console.error(e instanceof Error ? e.message : 'Invalid --body-file');
          process.exit(1);
        }
        const r = await createOnlineMeetingConversationMessageReply(
          auth.token,
          opts.conversationId,
          opts.messageId,
          body
        );
        if (!r.ok) {
          console.error(`Error: ${r.error?.message}`);
          process.exit(1);
        }
        console.log(JSON.stringify(r.data ?? null, null, 2));
      }
    );

  viva
    .command('meeting-engage-conversation-message-reply-get')
    .description('GET reply on a meeting conversation message (beta)')
    .requiredOption('--conversation-id <id>', 'Conversation id')
    .requiredOption('--message-id <id>', 'Parent message id')
    .requiredOption('--reply-id <id>', 'Reply id')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(
      async (opts: {
        conversationId: string;
        messageId: string;
        replyId: string;
        token?: string;
        identity?: string;
      }) => {
        const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
        if (!auth.success || !auth.token) {
          console.error(`Auth error: ${auth.error}`);
          process.exit(1);
        }
        const r = await getOnlineMeetingConversationMessageReply(
          auth.token,
          opts.conversationId,
          opts.messageId,
          opts.replyId
        );
        if (!r.ok) {
          console.error(`Error: ${r.error?.message}`);
          process.exit(1);
        }
        console.log(JSON.stringify(r.data ?? null, null, 2));
      }
    );

  viva
    .command('meeting-engage-conversation-message-reply-patch')
    .description('PATCH reply on a meeting conversation message (beta)')
    .requiredOption('--conversation-id <id>', 'Conversation id')
    .requiredOption('--message-id <id>', 'Parent message id')
    .requiredOption('--reply-id <id>', 'Reply id')
    .requiredOption('--body-file <path>', 'JSON object for PATCH body')
    .option('--if-match <etag>', 'If-Match header value')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(
      async (
        opts: {
          conversationId: string;
          messageId: string;
          replyId: string;
          bodyFile: string;
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
        let body: unknown;
        try {
          body = await readJsonBody(opts.bodyFile);
        } catch (e) {
          console.error(e instanceof Error ? e.message : 'Invalid --body-file');
          process.exit(1);
        }
        const r = await patchOnlineMeetingConversationMessageReply(
          auth.token,
          opts.conversationId,
          opts.messageId,
          opts.replyId,
          body,
          opts.ifMatch
        );
        if (!r.ok) {
          console.error(`Error: ${r.error?.message}`);
          process.exit(1);
        }
        console.log(JSON.stringify(r.data ?? null, null, 2));
      }
    );

  viva
    .command('meeting-engage-conversation-message-reply-delete')
    .description('DELETE reply on a meeting conversation message (beta)')
    .requiredOption('--conversation-id <id>', 'Conversation id')
    .requiredOption('--message-id <id>', 'Parent message id')
    .requiredOption('--reply-id <id>', 'Reply id')
    .option('--if-match <etag>', 'If-Match header value')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(
      async (
        opts: {
          conversationId: string;
          messageId: string;
          replyId: string;
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
        const r = await deleteOnlineMeetingConversationMessageReply(
          auth.token,
          opts.conversationId,
          opts.messageId,
          opts.replyId,
          opts.ifMatch
        );
        if (!r.ok) {
          console.error(`Error: ${r.error?.message}`);
          process.exit(1);
        }
      }
    );

  viva
    .command('meeting-engage-conversation-message-reply-conversation-get')
    .description('GET conversation navigation from a reply (beta)')
    .requiredOption('--conversation-id <id>', 'Conversation id')
    .requiredOption('--message-id <id>', 'Parent message id')
    .requiredOption('--reply-id <id>', 'Reply id')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(
      async (opts: {
        conversationId: string;
        messageId: string;
        replyId: string;
        token?: string;
        identity?: string;
      }) => {
        const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
        if (!auth.success || !auth.token) {
          console.error(`Auth error: ${auth.error}`);
          process.exit(1);
        }
        const r = await getOnlineMeetingConversationMessageReplyConversation(
          auth.token,
          opts.conversationId,
          opts.messageId,
          opts.replyId
        );
        if (!r.ok) {
          console.error(`Error: ${r.error?.message}`);
          process.exit(1);
        }
        console.log(JSON.stringify(r.data ?? null, null, 2));
      }
    );

  viva
    .command('meeting-engage-conversation-message-reply-reply-to-get')
    .description('GET replyTo navigation on a reply (beta)')
    .requiredOption('--conversation-id <id>', 'Conversation id')
    .requiredOption('--message-id <id>', 'Parent message id')
    .requiredOption('--reply-id <id>', 'Reply id')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(
      async (opts: {
        conversationId: string;
        messageId: string;
        replyId: string;
        token?: string;
        identity?: string;
      }) => {
        const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
        if (!auth.success || !auth.token) {
          console.error(`Auth error: ${auth.error}`);
          process.exit(1);
        }
        const r = await getOnlineMeetingConversationMessageReplyReplyTo(
          auth.token,
          opts.conversationId,
          opts.messageId,
          opts.replyId
        );
        if (!r.ok) {
          console.error(`Error: ${r.error?.message}`);
          process.exit(1);
        }
        console.log(JSON.stringify(r.data ?? null, null, 2));
      }
    );

  addODataListOpts(
    viva
      .command('meeting-engage-conversation-message-reply-reactions-list')
      .description('List reactions on a meeting conversation reply (beta, paged)')
      .requiredOption('--conversation-id <id>', 'Conversation id')
      .requiredOption('--message-id <id>', 'Parent message id')
      .requiredOption('--reply-id <id>', 'Reply id')
      .option('--json', 'Output as JSON array')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity (default: default)')
  ).action(
    async (
      opts: ODataOpts & {
        conversationId: string;
        messageId: string;
        replyId: string;
        json?: boolean;
        token?: string;
        identity?: string;
      }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        failVivaExtra(opts.json, 'Auth error', auth.error);
      }
      const r = await listOnlineMeetingConversationMessageReplyReactions(
        auth.token,
        opts.conversationId,
        opts.messageId,
        opts.replyId,
        odataQ(opts)
      );
      if (!r.ok || !r.data) {
        failVivaExtra(opts.json, 'Error', r.error, 'List failed');
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else for (const row of r.data) console.log(JSON.stringify(row));
    }
  );

  viva
    .command('meeting-engage-conversation-message-reply-reaction-create')
    .description('POST reaction on a meeting conversation reply (beta)')
    .requiredOption('--conversation-id <id>', 'Conversation id')
    .requiredOption('--message-id <id>', 'Parent message id')
    .requiredOption('--reply-id <id>', 'Reply id')
    .requiredOption('--body-file <path>', 'JSON object for POST body')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(
      async (
        opts: {
          conversationId: string;
          messageId: string;
          replyId: string;
          bodyFile: string;
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
        let body: unknown;
        try {
          body = await readJsonBody(opts.bodyFile);
        } catch (e) {
          console.error(e instanceof Error ? e.message : 'Invalid --body-file');
          process.exit(1);
        }
        const r = await createOnlineMeetingConversationMessageReplyReaction(
          auth.token,
          opts.conversationId,
          opts.messageId,
          opts.replyId,
          body
        );
        if (!r.ok) {
          console.error(`Error: ${r.error?.message}`);
          process.exit(1);
        }
        console.log(JSON.stringify(r.data ?? null, null, 2));
      }
    );

  viva
    .command('meeting-engage-conversation-message-reply-reaction-get')
    .description('GET reaction on a meeting conversation reply (beta)')
    .requiredOption('--conversation-id <id>', 'Conversation id')
    .requiredOption('--message-id <id>', 'Parent message id')
    .requiredOption('--reply-id <id>', 'Reply id')
    .requiredOption('--reaction-id <id>', 'Reaction id')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(
      async (opts: {
        conversationId: string;
        messageId: string;
        replyId: string;
        reactionId: string;
        token?: string;
        identity?: string;
      }) => {
        const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
        if (!auth.success || !auth.token) {
          console.error(`Auth error: ${auth.error}`);
          process.exit(1);
        }
        const r = await getOnlineMeetingConversationMessageReplyReaction(
          auth.token,
          opts.conversationId,
          opts.messageId,
          opts.replyId,
          opts.reactionId
        );
        if (!r.ok) {
          console.error(`Error: ${r.error?.message}`);
          process.exit(1);
        }
        console.log(JSON.stringify(r.data ?? null, null, 2));
      }
    );

  viva
    .command('meeting-engage-conversation-message-reply-reaction-patch')
    .description('PATCH reaction on a meeting conversation reply (beta)')
    .requiredOption('--conversation-id <id>', 'Conversation id')
    .requiredOption('--message-id <id>', 'Parent message id')
    .requiredOption('--reply-id <id>', 'Reply id')
    .requiredOption('--reaction-id <id>', 'Reaction id')
    .requiredOption('--body-file <path>', 'JSON object for PATCH body')
    .option('--if-match <etag>', 'If-Match header value')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(
      async (
        opts: {
          conversationId: string;
          messageId: string;
          replyId: string;
          reactionId: string;
          bodyFile: string;
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
        let body: unknown;
        try {
          body = await readJsonBody(opts.bodyFile);
        } catch (e) {
          console.error(e instanceof Error ? e.message : 'Invalid --body-file');
          process.exit(1);
        }
        const r = await patchOnlineMeetingConversationMessageReplyReaction(
          auth.token,
          opts.conversationId,
          opts.messageId,
          opts.replyId,
          opts.reactionId,
          body,
          opts.ifMatch
        );
        if (!r.ok) {
          console.error(`Error: ${r.error?.message}`);
          process.exit(1);
        }
        console.log(JSON.stringify(r.data ?? null, null, 2));
      }
    );

  viva
    .command('meeting-engage-conversation-message-reply-reaction-delete')
    .description('DELETE reaction on a meeting conversation reply (beta)')
    .requiredOption('--conversation-id <id>', 'Conversation id')
    .requiredOption('--message-id <id>', 'Parent message id')
    .requiredOption('--reply-id <id>', 'Reply id')
    .requiredOption('--reaction-id <id>', 'Reaction id')
    .option('--if-match <etag>', 'If-Match header value')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(
      async (
        opts: {
          conversationId: string;
          messageId: string;
          replyId: string;
          reactionId: string;
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
        const r = await deleteOnlineMeetingConversationMessageReplyReaction(
          auth.token,
          opts.conversationId,
          opts.messageId,
          opts.replyId,
          opts.reactionId,
          opts.ifMatch
        );
        if (!r.ok) {
          console.error(`Error: ${r.error?.message}`);
          process.exit(1);
        }
      }
    );

  viva
    .command('meeting-engage-conversation-online-meeting-get')
    .description('GET onlineMeeting linked to conversation (beta)')
    .requiredOption('--conversation-id <id>', 'Conversation id')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(async (opts: { conversationId: string; token?: string; identity?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getOnlineMeetingConversationOnlineMeeting(auth.token, opts.conversationId);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });
}
