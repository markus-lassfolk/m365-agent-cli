import { readFile } from 'node:fs/promises';
import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import {
  buildVivaListQuery,
  createEmployeeExperienceAssignedRole,
  createEmployeeExperienceAssignedRoleMember,
  deleteEmployeeExperience,
  deleteEmployeeExperienceAssignedRole,
  deleteEmployeeExperienceAssignedRoleMember,
  deleteUserItemInsightsSettings,
  deleteWorkingTimeSchedule,
  endWorkingTime,
  getEmployeeExperience,
  getEmployeeExperienceAssignedRole,
  getEmployeeExperienceAssignedRoleMember,
  getEmployeeExperienceAssignedRoleMemberUser,
  getEmployeeExperienceAssignedRoleMemberUserMailboxSettings,
  getLearningCourseActivity,
  getLearningCourseActivityByExternalId,
  getUserItemInsightsSettings,
  getWorkingTimeSchedule,
  listEmployeeExperienceAssignedRoleMembers,
  listEmployeeExperienceAssignedRoleMemberUserServiceProvisioningErrors,
  listEmployeeExperienceAssignedRoles,
  listLearningCourseActivities,
  patchEmployeeExperience,
  patchEmployeeExperienceAssignedRole,
  patchEmployeeExperienceAssignedRoleMember,
  patchEmployeeExperienceAssignedRoleMemberUserMailboxSettings,
  patchUserItemInsightsSettings,
  patchWorkingTimeSchedule,
  startWorkingTime
} from '../lib/graph-viva-client.js';
import { toJsonError } from '../lib/json-error.js';
import { checkReadOnly } from '../lib/utils.js';
import { registerVivaExtraSubcommands } from './viva-extra-subcommands.js';
import { registerVivaTenantSubcommands } from './viva-tenant-subcommands.js';

export const vivaCommand = new Command('viva').description(
  'Microsoft Graph **beta** Viva / employee experience: user + tenant `/employeeExperience`, insights, work hours, Engage roles/learning, meeting Q&A (see docs/GRAPH_SCOPES.md)'
);

/**
 * Prints the `--json` structured error envelope (or the matching plain-text "Auth error: .../
 * Error: ...") for the two failure shapes viva.ts's `--json`-capable list subcommands hit — an
 * auth failure from `resolveGraphAuth` (a plain string) or a Graph API failure
 * (`GraphResponse.error`, a GraphError-shaped object) — then exits 1. Mirrors
 * failBookings/failGroups so a `--json` viva call that fails gets `{ error: {...} } ` on stdout
 * instead of plain text on stderr.
 */
function failViva(
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

async function readJsonBody(path: string): Promise<unknown> {
  const raw = await readFile(path.trim(), 'utf8');
  return JSON.parse(raw) as unknown;
}

vivaCommand
  .command('employee-experience-get')
  .description('GET /me|/users/{id}/employeeExperience (beta)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <id>', 'Target user id or UPN (omit for /me)')
  .action(async (opts: { token?: string; identity?: string; user?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await getEmployeeExperience(auth.token, opts.user);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log(JSON.stringify(r.data ?? null, null, 2));
  });

vivaCommand
  .command('employee-experience-patch')
  .description('PATCH /me|/users/{id}/employeeExperience (beta); body from --body-file')
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
    const r = await patchEmployeeExperience(auth.token, body, opts.user);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log(JSON.stringify(r.data ?? null, null, 2));
  });

vivaCommand
  .command('employee-experience-delete')
  .description('DELETE /me|/users/{id}/employeeExperience navigation (beta); optional If-Match')
  .option('--if-match <etag>', 'If-Match header value')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <id>', 'Target user id or UPN (omit for /me)')
  .action(async (opts: { ifMatch?: string; token?: string; identity?: string; user?: string }, cmd: Command) => {
    checkReadOnly(cmd);
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await deleteEmployeeExperience(auth.token, opts.user, opts.ifMatch);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
  });

vivaCommand
  .command('working-time-schedule-get')
  .description('GET /me|/users/{id}/solutions/workingTimeSchedule (beta)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <id>', 'Target user id or UPN (omit for /me)')
  .action(async (opts: { token?: string; identity?: string; user?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await getWorkingTimeSchedule(auth.token, opts.user);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log(JSON.stringify(r.data ?? null, null, 2));
  });

vivaCommand
  .command('working-time-schedule-patch')
  .description('PATCH workingTimeSchedule (beta); body from --body-file JSON')
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
    const r = await patchWorkingTimeSchedule(auth.token, body, opts.user);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log(JSON.stringify(r.data ?? null, null, 2));
  });

vivaCommand
  .command('working-time-schedule-delete')
  .description('DELETE workingTimeSchedule (beta); optional If-Match')
  .option('--if-match <etag>', 'If-Match header value')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <id>', 'Target user id or UPN (omit for /me)')
  .action(async (opts: { ifMatch?: string; token?: string; identity?: string; user?: string }, cmd: Command) => {
    checkReadOnly(cmd);
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await deleteWorkingTimeSchedule(auth.token, opts.user, opts.ifMatch);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
  });

vivaCommand
  .command('start-working-time')
  .description('POST .../workingTimeSchedule/startWorkingTime (beta)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <id>', 'Target user id or UPN (omit for /me)')
  .action(async (opts: { token?: string; identity?: string; user?: string }, cmd: Command) => {
    checkReadOnly(cmd);
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await startWorkingTime(auth.token, opts.user);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
  });

vivaCommand
  .command('end-working-time')
  .description('POST .../workingTimeSchedule/endWorkingTime (beta)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <id>', 'Target user id or UPN (omit for /me)')
  .action(async (opts: { token?: string; identity?: string; user?: string }, cmd: Command) => {
    checkReadOnly(cmd);
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await endWorkingTime(auth.token, opts.user);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
  });

vivaCommand
  .command('insights-settings-get')
  .description('GET /me|/users/{id}/settings/itemInsights (beta)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <id>', 'Target user id or UPN (omit for /me)')
  .action(async (opts: { token?: string; identity?: string; user?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await getUserItemInsightsSettings(auth.token, opts.user);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log(JSON.stringify(r.data ?? null, null, 2));
  });

vivaCommand
  .command('insights-settings-patch')
  .description('PATCH user itemInsights / userInsightsSettings (beta)')
  .requiredOption('--body-file <path>', 'JSON object for PATCH body')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <id>', 'Target user id or UPN')
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
    const r = await patchUserItemInsightsSettings(auth.token, body, opts.user);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log(JSON.stringify(r.data ?? null, null, 2));
  });

vivaCommand
  .command('insights-settings-delete')
  .description('DELETE /me|/users/{id}/settings/itemInsights (beta)')
  .option('--if-match <etag>', 'If-Match header value')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <id>', 'Target user id or UPN (omit for /me)')
  .action(async (opts: { ifMatch?: string; token?: string; identity?: string; user?: string }, cmd: Command) => {
    checkReadOnly(cmd);
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await deleteUserItemInsightsSettings(auth.token, opts.user, opts.ifMatch);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
  });

vivaCommand
  .command('engage-assigned-roles-list')
  .description('List GET /me|/users/{id}/employeeExperience/assignedRoles (beta, paged)')
  .option('--json', 'Output as JSON array')
  .option('--filter <odata>', 'OData $filter')
  .option('--select <fields>', 'OData $select (comma-separated)')
  .option('--top <n>', 'OData $top', (v) => parseInt(v, 10))
  .option('--skip <n>', 'OData $skip', (v) => parseInt(v, 10))
  .option('--count', 'Include $count=true')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <id>', 'Target user id or UPN (omit for /me)')
  .action(
    async (opts: {
      json?: boolean;
      filter?: string;
      select?: string;
      top?: number;
      skip?: number;
      count?: boolean;
      token?: string;
      identity?: string;
      user?: string;
    }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        failViva(opts.json, 'Auth error', auth.error);
      }
      const q = buildVivaListQuery({
        filter: opts.filter,
        select: opts.select,
        top: opts.top,
        skip: opts.skip,
        count: opts.count
      });
      const r = await listEmployeeExperienceAssignedRoles(auth.token, opts.user, q);
      if (!r.ok || !r.data) {
        failViva(opts.json, 'Error', r.error, 'List failed');
      }
      if (opts.json) {
        console.log(JSON.stringify(r.data, null, 2));
      } else {
        for (const row of r.data) {
          console.log(JSON.stringify(row));
        }
      }
    }
  );

vivaCommand
  .command('engage-assigned-role-create')
  .description('POST /me|/users/{id}/employeeExperience/assignedRoles (beta)')
  .requiredOption('--body-file <path>', 'JSON object for POST body')
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
    const r = await createEmployeeExperienceAssignedRole(auth.token, body, opts.user);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log(JSON.stringify(r.data ?? null, null, 2));
  });

vivaCommand
  .command('engage-assigned-role-get')
  .description('GET assigned role by id (beta)')
  .requiredOption('--role-id <id>', 'engagementRole id')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <id>', 'Target user id or UPN (omit for /me)')
  .action(async (opts: { roleId: string; token?: string; identity?: string; user?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await getEmployeeExperienceAssignedRole(auth.token, opts.roleId, opts.user);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log(JSON.stringify(r.data ?? null, null, 2));
  });

vivaCommand
  .command('engage-assigned-role-patch')
  .description('PATCH assigned role by id (beta)')
  .requiredOption('--role-id <id>', 'engagementRole id')
  .requiredOption('--body-file <path>', 'JSON object for PATCH body')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <id>', 'Target user id or UPN (omit for /me)')
  .action(
    async (
      opts: { roleId: string; bodyFile: string; token?: string; identity?: string; user?: string },
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
      const r = await patchEmployeeExperienceAssignedRole(auth.token, opts.roleId, body, opts.user);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    }
  );

vivaCommand
  .command('engage-assigned-role-delete')
  .description('DELETE assigned role by id (beta)')
  .requiredOption('--role-id <id>', 'engagementRole id')
  .option('--if-match <etag>', 'If-Match header value')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <id>', 'Target user id or UPN (omit for /me)')
  .action(
    async (
      opts: { roleId: string; ifMatch?: string; token?: string; identity?: string; user?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await deleteEmployeeExperienceAssignedRole(auth.token, opts.roleId, opts.user, opts.ifMatch);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
    }
  );

vivaCommand
  .command('engage-assigned-role-members-list')
  .description('List members for an assigned role (beta, paged)')
  .requiredOption('--role-id <id>', 'engagementRole id')
  .option('--json', 'Output as JSON array')
  .option('--filter <odata>', 'OData $filter')
  .option('--select <fields>', 'OData $select (comma-separated)')
  .option('--top <n>', 'OData $top', (v) => parseInt(v, 10))
  .option('--skip <n>', 'OData $skip', (v) => parseInt(v, 10))
  .option('--count', 'Include $count=true')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <id>', 'Target user id or UPN (omit for /me)')
  .action(
    async (opts: {
      roleId: string;
      json?: boolean;
      filter?: string;
      select?: string;
      top?: number;
      skip?: number;
      count?: boolean;
      token?: string;
      identity?: string;
      user?: string;
    }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        failViva(opts.json, 'Auth error', auth.error);
      }
      const q = buildVivaListQuery({
        filter: opts.filter,
        select: opts.select,
        top: opts.top,
        skip: opts.skip,
        count: opts.count
      });
      const r = await listEmployeeExperienceAssignedRoleMembers(auth.token, opts.roleId, opts.user, q);
      if (!r.ok || !r.data) {
        failViva(opts.json, 'Error', r.error, 'List failed');
      }
      if (opts.json) {
        console.log(JSON.stringify(r.data, null, 2));
      } else {
        for (const row of r.data) {
          console.log(JSON.stringify(row));
        }
      }
    }
  );

vivaCommand
  .command('engage-assigned-role-member-create')
  .description('POST member on an assigned role (beta)')
  .requiredOption('--role-id <id>', 'engagementRole id')
  .requiredOption('--body-file <path>', 'JSON object for POST body')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <id>', 'Target user id or UPN (omit for /me)')
  .action(
    async (
      opts: { roleId: string; bodyFile: string; token?: string; identity?: string; user?: string },
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
      const r = await createEmployeeExperienceAssignedRoleMember(auth.token, opts.roleId, body, opts.user);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    }
  );

vivaCommand
  .command('engage-assigned-role-member-get')
  .description('GET assigned role member by id (beta)')
  .requiredOption('--role-id <id>', 'engagementRole id')
  .requiredOption('--member-id <id>', 'engagementRoleMember id')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <id>', 'Target user id or UPN (omit for /me)')
  .action(async (opts: { roleId: string; memberId: string; token?: string; identity?: string; user?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await getEmployeeExperienceAssignedRoleMember(auth.token, opts.roleId, opts.memberId, opts.user);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log(JSON.stringify(r.data ?? null, null, 2));
  });

vivaCommand
  .command('engage-assigned-role-member-patch')
  .description('PATCH assigned role member (beta)')
  .requiredOption('--role-id <id>', 'engagementRole id')
  .requiredOption('--member-id <id>', 'engagementRoleMember id')
  .requiredOption('--body-file <path>', 'JSON object for PATCH body')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <id>', 'Target user id or UPN (omit for /me)')
  .action(
    async (
      opts: {
        roleId: string;
        memberId: string;
        bodyFile: string;
        token?: string;
        identity?: string;
        user?: string;
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
      const r = await patchEmployeeExperienceAssignedRoleMember(
        auth.token,
        opts.roleId,
        opts.memberId,
        body,
        opts.user
      );
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    }
  );

vivaCommand
  .command('engage-assigned-role-member-delete')
  .description('DELETE assigned role member (beta)')
  .requiredOption('--role-id <id>', 'engagementRole id')
  .requiredOption('--member-id <id>', 'engagementRoleMember id')
  .option('--if-match <etag>', 'If-Match header value')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <id>', 'Target user id or UPN (omit for /me)')
  .action(
    async (
      opts: { roleId: string; memberId: string; ifMatch?: string; token?: string; identity?: string; user?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await deleteEmployeeExperienceAssignedRoleMember(
        auth.token,
        opts.roleId,
        opts.memberId,
        opts.user,
        opts.ifMatch
      );
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
    }
  );

vivaCommand
  .command('engage-assigned-role-member-user-get')
  .description('GET nested user on assigned role member (beta)')
  .requiredOption('--role-id <id>', 'engagementRole id')
  .requiredOption('--member-id <id>', 'engagementRoleMember id')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <id>', 'Target user id or UPN (omit for /me)')
  .action(async (opts: { roleId: string; memberId: string; token?: string; identity?: string; user?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await getEmployeeExperienceAssignedRoleMemberUser(auth.token, opts.roleId, opts.memberId, opts.user);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log(JSON.stringify(r.data ?? null, null, 2));
  });

vivaCommand
  .command('engage-assigned-role-member-user-mailbox-settings-get')
  .description('GET mailboxSettings on assigned role member user (beta)')
  .requiredOption('--role-id <id>', 'engagementRole id')
  .requiredOption('--member-id <id>', 'engagementRoleMember id')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <id>', 'Target user id or UPN (omit for /me)')
  .action(async (opts: { roleId: string; memberId: string; token?: string; identity?: string; user?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await getEmployeeExperienceAssignedRoleMemberUserMailboxSettings(
      auth.token,
      opts.roleId,
      opts.memberId,
      opts.user
    );
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log(JSON.stringify(r.data ?? null, null, 2));
  });

vivaCommand
  .command('engage-assigned-role-member-user-mailbox-settings-patch')
  .description('PATCH mailboxSettings on assigned role member user (beta)')
  .requiredOption('--role-id <id>', 'engagementRole id')
  .requiredOption('--member-id <id>', 'engagementRoleMember id')
  .requiredOption('--body-file <path>', 'JSON object for PATCH body')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <id>', 'Target user id or UPN (omit for /me)')
  .action(
    async (
      opts: { roleId: string; memberId: string; bodyFile: string; token?: string; identity?: string; user?: string },
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
      const r = await patchEmployeeExperienceAssignedRoleMemberUserMailboxSettings(
        auth.token,
        opts.roleId,
        opts.memberId,
        body,
        opts.user
      );
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    }
  );

vivaCommand
  .command('engage-assigned-role-member-user-service-provisioning-errors-list')
  .description('List serviceProvisioningErrors on assigned role member user (beta, paged)')
  .requiredOption('--role-id <id>', 'engagementRole id')
  .requiredOption('--member-id <id>', 'engagementRoleMember id')
  .option('--json', 'Output as JSON array')
  .option('--filter <odata>', 'OData $filter')
  .option('--select <fields>', 'OData $select (comma-separated)')
  .option('--top <n>', 'OData $top', (v) => parseInt(v, 10))
  .option('--skip <n>', 'OData $skip', (v) => parseInt(v, 10))
  .option('--count', 'Include $count=true')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <id>', 'Target user id or UPN (omit for /me)')
  .action(
    async (opts: {
      roleId: string;
      memberId: string;
      json?: boolean;
      filter?: string;
      select?: string;
      top?: number;
      skip?: number;
      count?: boolean;
      token?: string;
      identity?: string;
      user?: string;
    }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        failViva(opts.json, 'Auth error', auth.error);
      }
      const q = buildVivaListQuery({
        filter: opts.filter,
        select: opts.select,
        top: opts.top,
        skip: opts.skip,
        count: opts.count
      });
      const r = await listEmployeeExperienceAssignedRoleMemberUserServiceProvisioningErrors(
        auth.token,
        opts.roleId,
        opts.memberId,
        opts.user,
        q
      );
      if (!r.ok || !r.data) {
        failViva(opts.json, 'Error', r.error, 'List failed');
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else for (const row of r.data) console.log(JSON.stringify(row));
    }
  );

vivaCommand
  .command('learning-activities-list')
  .description('List GET /me|/users/{id}/employeeExperience/learningCourseActivities (beta, paged)')
  .option('--json', 'Output as JSON array')
  .option('--filter <odata>', 'OData $filter')
  .option('--select <fields>', 'OData $select (comma-separated)')
  .option('--top <n>', 'OData $top', (v) => parseInt(v, 10))
  .option('--skip <n>', 'OData $skip', (v) => parseInt(v, 10))
  .option('--count', 'Include $count=true')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <id>', 'Target user id or UPN (omit for /me)')
  .action(
    async (opts: {
      json?: boolean;
      filter?: string;
      select?: string;
      top?: number;
      skip?: number;
      count?: boolean;
      token?: string;
      identity?: string;
      user?: string;
    }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        failViva(opts.json, 'Auth error', auth.error);
      }
      const q = buildVivaListQuery({
        filter: opts.filter,
        select: opts.select,
        top: opts.top,
        skip: opts.skip,
        count: opts.count
      });
      const r = await listLearningCourseActivities(auth.token, opts.user, q);
      if (!r.ok || !r.data) {
        failViva(opts.json, 'Error', r.error, 'List failed');
      }
      if (opts.json) {
        console.log(JSON.stringify(r.data, null, 2));
      } else {
        for (const row of r.data) {
          console.log(JSON.stringify(row));
        }
      }
    }
  );

vivaCommand
  .command('learning-activity-get')
  .description('GET learning course activity by Graph id (beta)')
  .requiredOption('--activity-id <id>', 'learningCourseActivity id')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <id>', 'Target user id or UPN (omit for /me)')
  .action(async (opts: { activityId: string; token?: string; identity?: string; user?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await getLearningCourseActivity(auth.token, opts.activityId, opts.user);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log(JSON.stringify(r.data ?? null, null, 2));
  });

vivaCommand
  .command('learning-activity-get-external')
  .description('GET activity by alternate key externalcourseActivityId (beta; OData key, quotes escaped)')
  .requiredOption('--external-activity-id <id>', 'External course activity id')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <id>', 'Target user id or UPN (omit for /me)')
  .action(async (opts: { externalActivityId: string; token?: string; identity?: string; user?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await getLearningCourseActivityByExternalId(auth.token, opts.externalActivityId, opts.user);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log(JSON.stringify(r.data ?? null, null, 2));
  });

registerVivaTenantSubcommands(vivaCommand);
registerVivaExtraSubcommands(vivaCommand);
