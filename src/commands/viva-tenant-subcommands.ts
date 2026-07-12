import { readFile } from 'node:fs/promises';
import type { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import { buildVivaListQuery } from '../lib/graph-viva-client.js';
import {
  createTenantCommunity,
  createTenantEngagementAsyncOperation,
  createTenantEngagementRole,
  createTenantEngagementRoleMember,
  createTenantGoalsExportJob,
  createTenantLearningContent,
  createTenantLearningProvider,
  createTenantProviderLearningCourseActivity,
  createTenantRootLearningCourseActivity,
  deleteTenantCommunity,
  deleteTenantEmployeeExperience,
  deleteTenantEngagementAsyncOperation,
  deleteTenantEngagementRole,
  deleteTenantEngagementRoleMember,
  deleteTenantGoals,
  deleteTenantGoalsExportJob,
  deleteTenantLearningContent,
  deleteTenantLearningProvider,
  deleteTenantProviderLearningCourseActivity,
  deleteTenantRootLearningCourseActivity,
  getTenantCommunity,
  getTenantCommunityGroup,
  getTenantCommunityOwner,
  getTenantCommunityOwnerByUserPrincipalName,
  getTenantCommunityOwnerMailboxSettings,
  getTenantEmployeeExperience,
  getTenantEngagementAsyncOperation,
  getTenantEngagementRole,
  getTenantEngagementRoleMember,
  getTenantEngagementRoleMemberUser,
  getTenantEngagementRoleMemberUserMailboxSettings,
  getTenantGoals,
  getTenantGoalsExportJob,
  getTenantGoalsExportJobContent,
  getTenantLearningContent,
  getTenantLearningContentByExternal,
  getTenantLearningProvider,
  getTenantProviderLearningCourseActivity,
  getTenantProviderLearningCourseActivityByExternal,
  getTenantRootLearningCourseActivity,
  getTenantRootLearningCourseActivityByExternal,
  listTenantCommunities,
  listTenantCommunityOwnerServiceProvisioningErrors,
  listTenantCommunityOwners,
  listTenantEngagementAsyncOperations,
  listTenantEngagementRoleMembers,
  listTenantEngagementRoleMemberUserServiceProvisioningErrors,
  listTenantEngagementRoles,
  listTenantGoalsExportJobs,
  listTenantLearningContents,
  listTenantLearningProviders,
  listTenantProviderLearningCourseActivities,
  listTenantRootLearningCourseActivities,
  patchTenantCommunity,
  patchTenantCommunityOwnerMailboxSettings,
  patchTenantEmployeeExperience,
  patchTenantEngagementAsyncOperation,
  patchTenantEngagementRole,
  patchTenantEngagementRoleMember,
  patchTenantEngagementRoleMemberUserMailboxSettings,
  patchTenantGoals,
  patchTenantGoalsExportJob,
  patchTenantLearningContent,
  patchTenantLearningProvider,
  patchTenantProviderLearningCourseActivity,
  patchTenantRootLearningCourseActivity
} from '../lib/graph-viva-tenant-client.js';
import { toJsonError } from '../lib/json-error.js';
import { checkReadOnly } from '../lib/utils.js';

async function readJsonBody(path: string): Promise<unknown> {
  const raw = await readFile(path.trim(), 'utf8');
  return JSON.parse(raw) as unknown;
}

/** Prints the --json structured error envelope (or the matching plain-text "Auth error: .../
 *  Error: ...") for the two failure shapes the tenant `-list` subcommands hit, then exits 1. */
function failVivaTenant(
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

/** Register **`/employeeExperience`** tenant (org-wide) beta subcommands on **`viva`**. */
export function registerVivaTenantSubcommands(viva: Command): void {
  viva
    .command('tenant-employee-experience-get')
    .description('GET /employeeExperience tenant singleton (beta)')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(async (opts: { token?: string; identity?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getTenantEmployeeExperience(auth.token);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  viva
    .command('tenant-employee-experience-patch')
    .description('PATCH /employeeExperience (beta); --body-file JSON')
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
      const r = await patchTenantEmployeeExperience(auth.token, body);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  viva
    .command('tenant-employee-experience-delete')
    .description('DELETE /employeeExperience navigation (beta); optional If-Match')
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
      const r = await deleteTenantEmployeeExperience(auth.token, opts.ifMatch);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
    });

  addODataListOpts(
    viva
      .command('tenant-communities-list')
      .description('List GET /employeeExperience/communities (beta, paged)')
      .option('--json', 'Output as JSON array')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity (default: default)')
  ).action(async (opts: ODataOpts & { json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      failVivaTenant(opts.json, 'Auth error', auth.error);
    }
    const r = await listTenantCommunities(auth.token, odataQ(opts));
    if (!r.ok || !r.data) {
      failVivaTenant(opts.json, 'Error', r.error, 'List failed');
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else for (const row of r.data) console.log(JSON.stringify(row));
  });

  viva
    .command('tenant-community-create')
    .description('POST /employeeExperience/communities (beta)')
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
      const r = await createTenantCommunity(auth.token, body);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  viva
    .command('tenant-community-get')
    .description('GET community by id (beta)')
    .requiredOption('--community-id <id>', 'Community id')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(async (opts: { communityId: string; token?: string; identity?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getTenantCommunity(auth.token, opts.communityId);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  viva
    .command('tenant-community-patch')
    .description('PATCH community (beta)')
    .requiredOption('--community-id <id>', 'Community id')
    .requiredOption('--body-file <path>', 'JSON object for PATCH body')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(
      async (opts: { communityId: string; bodyFile: string; token?: string; identity?: string }, cmd: Command) => {
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
        const r = await patchTenantCommunity(auth.token, opts.communityId, body);
        if (!r.ok) {
          console.error(`Error: ${r.error?.message}`);
          process.exit(1);
        }
        console.log(JSON.stringify(r.data ?? null, null, 2));
      }
    );

  viva
    .command('tenant-community-delete')
    .description('DELETE community (beta)')
    .requiredOption('--community-id <id>', 'Community id')
    .option('--if-match <etag>', 'If-Match header value')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(
      async (opts: { communityId: string; ifMatch?: string; token?: string; identity?: string }, cmd: Command) => {
        checkReadOnly(cmd);
        const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
        if (!auth.success || !auth.token) {
          console.error(`Auth error: ${auth.error}`);
          process.exit(1);
        }
        const r = await deleteTenantCommunity(auth.token, opts.communityId, opts.ifMatch);
        if (!r.ok) {
          console.error(`Error: ${r.error?.message}`);
          process.exit(1);
        }
      }
    );

  viva
    .command('tenant-community-group-get')
    .description('GET .../communities/{id}/group (beta)')
    .requiredOption('--community-id <id>', 'Community id')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(async (opts: { communityId: string; token?: string; identity?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getTenantCommunityGroup(auth.token, opts.communityId);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  addODataListOpts(
    viva
      .command('tenant-community-owners-list')
      .description('List community owners (beta, paged)')
      .requiredOption('--community-id <id>', 'Community id')
      .option('--json', 'Output as JSON array')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity (default: default)')
  ).action(async (opts: ODataOpts & { communityId: string; json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      failVivaTenant(opts.json, 'Auth error', auth.error);
    }
    const r = await listTenantCommunityOwners(auth.token, opts.communityId, odataQ(opts));
    if (!r.ok || !r.data) {
      failVivaTenant(opts.json, 'Error', r.error, 'List failed');
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else for (const row of r.data) console.log(JSON.stringify(row));
  });

  viva
    .command('tenant-community-owner-get')
    .description('GET community owner user (beta)')
    .requiredOption('--community-id <id>', 'Community id')
    .requiredOption('--owner-user-id <id>', 'Owner user id')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(async (opts: { communityId: string; ownerUserId: string; token?: string; identity?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getTenantCommunityOwner(auth.token, opts.communityId, opts.ownerUserId);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  viva
    .command('tenant-community-owner-get-by-upn')
    .description('GET community owner by userPrincipalName alternate key (beta)')
    .requiredOption('--community-id <id>', 'Community id')
    .requiredOption('--user-principal-name <upn>', 'Owner userPrincipalName')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(async (opts: { communityId: string; userPrincipalName: string; token?: string; identity?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getTenantCommunityOwnerByUserPrincipalName(auth.token, opts.communityId, opts.userPrincipalName);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  viva
    .command('tenant-community-owner-mailbox-settings-get')
    .description('GET mailboxSettings on community owner (beta)')
    .requiredOption('--community-id <id>', 'Community id')
    .requiredOption('--owner-user-id <id>', 'Owner user id')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(async (opts: { communityId: string; ownerUserId: string; token?: string; identity?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getTenantCommunityOwnerMailboxSettings(auth.token, opts.communityId, opts.ownerUserId);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  viva
    .command('tenant-community-owner-mailbox-settings-patch')
    .description('PATCH mailboxSettings on community owner (beta)')
    .requiredOption('--community-id <id>', 'Community id')
    .requiredOption('--owner-user-id <id>', 'Owner user id')
    .requiredOption('--body-file <path>', 'JSON object for PATCH body')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(
      async (
        opts: { communityId: string; ownerUserId: string; bodyFile: string; token?: string; identity?: string },
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
        const r = await patchTenantCommunityOwnerMailboxSettings(auth.token, opts.communityId, opts.ownerUserId, body);
        if (!r.ok) {
          console.error(`Error: ${r.error?.message}`);
          process.exit(1);
        }
        console.log(JSON.stringify(r.data ?? null, null, 2));
      }
    );

  addODataListOpts(
    viva
      .command('tenant-community-owner-service-provisioning-errors-list')
      .description('List serviceProvisioningErrors on community owner (beta, paged)')
      .requiredOption('--community-id <id>', 'Community id')
      .requiredOption('--owner-user-id <id>', 'Owner user id')
      .option('--json', 'Output as JSON array')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity (default: default)')
  ).action(
    async (
      opts: ODataOpts & { communityId: string; ownerUserId: string; json?: boolean; token?: string; identity?: string }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        failVivaTenant(opts.json, 'Auth error', auth.error);
      }
      const r = await listTenantCommunityOwnerServiceProvisioningErrors(
        auth.token,
        opts.communityId,
        opts.ownerUserId,
        odataQ(opts)
      );
      if (!r.ok || !r.data) {
        failVivaTenant(opts.json, 'Error', r.error, 'List failed');
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else for (const row of r.data) console.log(JSON.stringify(row));
    }
  );

  addODataListOpts(
    viva
      .command('tenant-async-ops-list')
      .description('List /employeeExperience/engagementAsyncOperations (beta, paged)')
      .option('--json', 'Output as JSON array')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity (default: default)')
  ).action(async (opts: ODataOpts & { json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      failVivaTenant(opts.json, 'Auth error', auth.error);
    }
    const r = await listTenantEngagementAsyncOperations(auth.token, odataQ(opts));
    if (!r.ok || !r.data) {
      failVivaTenant(opts.json, 'Error', r.error, 'List failed');
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else for (const row of r.data) console.log(JSON.stringify(row));
  });

  viva
    .command('tenant-async-op-create')
    .description('POST engagementAsyncOperations (beta)')
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
      const r = await createTenantEngagementAsyncOperation(auth.token, body);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  viva
    .command('tenant-async-op-get')
    .description('GET engagementAsyncOperation by id (beta)')
    .requiredOption('--operation-id <id>', 'engagementAsyncOperation id')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(async (opts: { operationId: string; token?: string; identity?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getTenantEngagementAsyncOperation(auth.token, opts.operationId);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  viva
    .command('tenant-async-op-patch')
    .description('PATCH engagementAsyncOperation (beta)')
    .requiredOption('--operation-id <id>', 'engagementAsyncOperation id')
    .requiredOption('--body-file <path>', 'JSON object for PATCH body')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(
      async (opts: { operationId: string; bodyFile: string; token?: string; identity?: string }, cmd: Command) => {
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
        const r = await patchTenantEngagementAsyncOperation(auth.token, opts.operationId, body);
        if (!r.ok) {
          console.error(`Error: ${r.error?.message}`);
          process.exit(1);
        }
        console.log(JSON.stringify(r.data ?? null, null, 2));
      }
    );

  viva
    .command('tenant-async-op-delete')
    .description('DELETE engagementAsyncOperation (beta)')
    .requiredOption('--operation-id <id>', 'engagementAsyncOperation id')
    .option('--if-match <etag>', 'If-Match header value')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(
      async (opts: { operationId: string; ifMatch?: string; token?: string; identity?: string }, cmd: Command) => {
        checkReadOnly(cmd);
        const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
        if (!auth.success || !auth.token) {
          console.error(`Auth error: ${auth.error}`);
          process.exit(1);
        }
        const r = await deleteTenantEngagementAsyncOperation(auth.token, opts.operationId, opts.ifMatch);
        if (!r.ok) {
          console.error(`Error: ${r.error?.message}`);
          process.exit(1);
        }
      }
    );

  viva
    .command('tenant-goals-get')
    .description('GET /employeeExperience/goals (beta)')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(async (opts: { token?: string; identity?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getTenantGoals(auth.token);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  viva
    .command('tenant-goals-patch')
    .description('PATCH /employeeExperience/goals (beta)')
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
      const r = await patchTenantGoals(auth.token, body);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  viva
    .command('tenant-goals-delete')
    .description('DELETE /employeeExperience/goals navigation (beta)')
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
      const r = await deleteTenantGoals(auth.token, opts.ifMatch);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
    });

  addODataListOpts(
    viva
      .command('tenant-goals-export-jobs-list')
      .description('List goals export jobs (beta, paged)')
      .option('--json', 'Output as JSON array')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity (default: default)')
  ).action(async (opts: ODataOpts & { json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      failVivaTenant(opts.json, 'Auth error', auth.error);
    }
    const r = await listTenantGoalsExportJobs(auth.token, odataQ(opts));
    if (!r.ok || !r.data) {
      failVivaTenant(opts.json, 'Error', r.error, 'List failed');
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else for (const row of r.data) console.log(JSON.stringify(row));
  });

  viva
    .command('tenant-goals-export-job-create')
    .description('POST goals export job (beta)')
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
      const r = await createTenantGoalsExportJob(auth.token, body);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  viva
    .command('tenant-goals-export-job-get')
    .description('GET goals export job (beta)')
    .requiredOption('--job-id <id>', 'goalsExportJob id')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(async (opts: { jobId: string; token?: string; identity?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getTenantGoalsExportJob(auth.token, opts.jobId);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  viva
    .command('tenant-goals-export-job-patch')
    .description('PATCH goals export job (beta)')
    .requiredOption('--job-id <id>', 'goalsExportJob id')
    .requiredOption('--body-file <path>', 'JSON object for PATCH body')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(async (opts: { jobId: string; bodyFile: string; token?: string; identity?: string }, cmd: Command) => {
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
      const r = await patchTenantGoalsExportJob(auth.token, opts.jobId, body);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  viva
    .command('tenant-goals-export-job-delete')
    .description('DELETE goals export job (beta)')
    .requiredOption('--job-id <id>', 'goalsExportJob id')
    .option('--if-match <etag>', 'If-Match header value')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(async (opts: { jobId: string; ifMatch?: string; token?: string; identity?: string }, cmd: Command) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await deleteTenantGoalsExportJob(auth.token, opts.jobId, opts.ifMatch);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
    });

  viva
    .command('tenant-goals-export-job-content')
    .description('GET export job content (beta; response body printed as text)')
    .requiredOption('--job-id <id>', 'goalsExportJob id')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(async (opts: { jobId: string; token?: string; identity?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getTenantGoalsExportJobContent(auth.token, opts.jobId);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (r.data !== undefined) process.stdout.write(r.data);
    });

  addODataListOpts(
    viva
      .command('tenant-learning-activities-list')
      .description('List /employeeExperience/learningCourseActivities (tenant root, beta, paged)')
      .option('--json', 'Output as JSON array')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity (default: default)')
  ).action(async (opts: ODataOpts & { json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      failVivaTenant(opts.json, 'Auth error', auth.error);
    }
    const r = await listTenantRootLearningCourseActivities(auth.token, odataQ(opts));
    if (!r.ok || !r.data) {
      failVivaTenant(opts.json, 'Error', r.error, 'List failed');
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else for (const row of r.data) console.log(JSON.stringify(row));
  });

  viva
    .command('tenant-learning-activity-create')
    .description('POST tenant learningCourseActivities (beta)')
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
      const r = await createTenantRootLearningCourseActivity(auth.token, body);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  viva
    .command('tenant-learning-activity-get')
    .description('GET tenant learning course activity by id (beta)')
    .requiredOption('--activity-id <id>', 'learningCourseActivity id')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(async (opts: { activityId: string; token?: string; identity?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getTenantRootLearningCourseActivity(auth.token, opts.activityId);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  viva
    .command('tenant-learning-activity-get-external')
    .description('GET tenant activity by externalcourseActivityId (beta)')
    .requiredOption('--external-activity-id <id>', 'External course activity id')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(async (opts: { externalActivityId: string; token?: string; identity?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getTenantRootLearningCourseActivityByExternal(auth.token, opts.externalActivityId);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  viva
    .command('tenant-learning-activity-patch')
    .description('PATCH tenant learning course activity (beta)')
    .requiredOption('--activity-id <id>', 'learningCourseActivity id')
    .requiredOption('--body-file <path>', 'JSON object for PATCH body')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(async (opts: { activityId: string; bodyFile: string; token?: string; identity?: string }, cmd: Command) => {
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
      const r = await patchTenantRootLearningCourseActivity(auth.token, opts.activityId, body);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  viva
    .command('tenant-learning-activity-delete')
    .description('DELETE tenant learning course activity (beta)')
    .requiredOption('--activity-id <id>', 'learningCourseActivity id')
    .option('--if-match <etag>', 'If-Match header value')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(async (opts: { activityId: string; ifMatch?: string; token?: string; identity?: string }, cmd: Command) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await deleteTenantRootLearningCourseActivity(auth.token, opts.activityId, opts.ifMatch);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
    });

  addODataListOpts(
    viva
      .command('tenant-learning-providers-list')
      .description('List /employeeExperience/learningProviders (beta, paged)')
      .option('--json', 'Output as JSON array')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity (default: default)')
  ).action(async (opts: ODataOpts & { json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      failVivaTenant(opts.json, 'Auth error', auth.error);
    }
    const r = await listTenantLearningProviders(auth.token, odataQ(opts));
    if (!r.ok || !r.data) {
      failVivaTenant(opts.json, 'Error', r.error, 'List failed');
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else for (const row of r.data) console.log(JSON.stringify(row));
  });

  viva
    .command('tenant-learning-provider-create')
    .description('POST learningProvider (beta)')
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
      const r = await createTenantLearningProvider(auth.token, body);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  viva
    .command('tenant-learning-provider-get')
    .description('GET learningProvider (beta)')
    .requiredOption('--provider-id <id>', 'learningProvider id')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(async (opts: { providerId: string; token?: string; identity?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getTenantLearningProvider(auth.token, opts.providerId);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  viva
    .command('tenant-learning-provider-patch')
    .description('PATCH learningProvider (beta)')
    .requiredOption('--provider-id <id>', 'learningProvider id')
    .requiredOption('--body-file <path>', 'JSON object for PATCH body')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(async (opts: { providerId: string; bodyFile: string; token?: string; identity?: string }, cmd: Command) => {
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
      const r = await patchTenantLearningProvider(auth.token, opts.providerId, body);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  viva
    .command('tenant-learning-provider-delete')
    .description('DELETE learningProvider (beta)')
    .requiredOption('--provider-id <id>', 'learningProvider id')
    .option('--if-match <etag>', 'If-Match header value')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(async (opts: { providerId: string; ifMatch?: string; token?: string; identity?: string }, cmd: Command) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await deleteTenantLearningProvider(auth.token, opts.providerId, opts.ifMatch);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
    });

  addODataListOpts(
    viva
      .command('tenant-learning-contents-list')
      .description('List learningContents for a provider (beta, paged)')
      .requiredOption('--provider-id <id>', 'learningProvider id')
      .option('--json', 'Output as JSON array')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity (default: default)')
  ).action(async (opts: ODataOpts & { providerId: string; json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      failVivaTenant(opts.json, 'Auth error', auth.error);
    }
    const r = await listTenantLearningContents(auth.token, opts.providerId, odataQ(opts));
    if (!r.ok || !r.data) {
      failVivaTenant(opts.json, 'Error', r.error, 'List failed');
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else for (const row of r.data) console.log(JSON.stringify(row));
  });

  viva
    .command('tenant-learning-content-create')
    .description('POST learningContent under provider (beta)')
    .requiredOption('--provider-id <id>', 'learningProvider id')
    .requiredOption('--body-file <path>', 'JSON object for POST body')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(async (opts: { providerId: string; bodyFile: string; token?: string; identity?: string }, cmd: Command) => {
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
      const r = await createTenantLearningContent(auth.token, opts.providerId, body);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  viva
    .command('tenant-learning-content-get')
    .description('GET learningContent (beta)')
    .requiredOption('--provider-id <id>', 'learningProvider id')
    .requiredOption('--content-id <id>', 'learningContent id')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(async (opts: { providerId: string; contentId: string; token?: string; identity?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getTenantLearningContent(auth.token, opts.providerId, opts.contentId);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  viva
    .command('tenant-learning-content-get-external')
    .description('GET learningContent by externalId alternate key (beta)')
    .requiredOption('--provider-id <id>', 'learningProvider id')
    .requiredOption('--external-id <id>', 'External id')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(async (opts: { providerId: string; externalId: string; token?: string; identity?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getTenantLearningContentByExternal(auth.token, opts.providerId, opts.externalId);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  viva
    .command('tenant-learning-content-patch')
    .description('PATCH learningContent (beta)')
    .requiredOption('--provider-id <id>', 'learningProvider id')
    .requiredOption('--content-id <id>', 'learningContent id')
    .requiredOption('--body-file <path>', 'JSON object for PATCH body')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(
      async (
        opts: { providerId: string; contentId: string; bodyFile: string; token?: string; identity?: string },
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
        const r = await patchTenantLearningContent(auth.token, opts.providerId, opts.contentId, body);
        if (!r.ok) {
          console.error(`Error: ${r.error?.message}`);
          process.exit(1);
        }
        console.log(JSON.stringify(r.data ?? null, null, 2));
      }
    );

  viva
    .command('tenant-learning-content-delete')
    .description('DELETE learningContent (beta)')
    .requiredOption('--provider-id <id>', 'learningProvider id')
    .requiredOption('--content-id <id>', 'learningContent id')
    .option('--if-match <etag>', 'If-Match header value')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(
      async (
        opts: { providerId: string; contentId: string; ifMatch?: string; token?: string; identity?: string },
        cmd: Command
      ) => {
        checkReadOnly(cmd);
        const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
        if (!auth.success || !auth.token) {
          console.error(`Auth error: ${auth.error}`);
          process.exit(1);
        }
        const r = await deleteTenantLearningContent(auth.token, opts.providerId, opts.contentId, opts.ifMatch);
        if (!r.ok) {
          console.error(`Error: ${r.error?.message}`);
          process.exit(1);
        }
      }
    );

  addODataListOpts(
    viva
      .command('tenant-provider-learning-activities-list')
      .description('List learningCourseActivities under a provider (beta, paged)')
      .requiredOption('--provider-id <id>', 'learningProvider id')
      .option('--json', 'Output as JSON array')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity (default: default)')
  ).action(async (opts: ODataOpts & { providerId: string; json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      failVivaTenant(opts.json, 'Auth error', auth.error);
    }
    const r = await listTenantProviderLearningCourseActivities(auth.token, opts.providerId, odataQ(opts));
    if (!r.ok || !r.data) {
      failVivaTenant(opts.json, 'Error', r.error, 'List failed');
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else for (const row of r.data) console.log(JSON.stringify(row));
  });

  viva
    .command('tenant-provider-learning-activity-create')
    .description('POST learningCourseActivity under provider (beta)')
    .requiredOption('--provider-id <id>', 'learningProvider id')
    .requiredOption('--body-file <path>', 'JSON object for POST body')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(async (opts: { providerId: string; bodyFile: string; token?: string; identity?: string }, cmd: Command) => {
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
      const r = await createTenantProviderLearningCourseActivity(auth.token, opts.providerId, body);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  viva
    .command('tenant-provider-learning-activity-get')
    .description('GET provider-scoped learning course activity (beta)')
    .requiredOption('--provider-id <id>', 'learningProvider id')
    .requiredOption('--activity-id <id>', 'learningCourseActivity id')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(async (opts: { providerId: string; activityId: string; token?: string; identity?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getTenantProviderLearningCourseActivity(auth.token, opts.providerId, opts.activityId);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  viva
    .command('tenant-provider-learning-activity-get-external')
    .description('GET provider activity by externalcourseActivityId (beta)')
    .requiredOption('--provider-id <id>', 'learningProvider id')
    .requiredOption('--external-activity-id <id>', 'External course activity id')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(async (opts: { providerId: string; externalActivityId: string; token?: string; identity?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getTenantProviderLearningCourseActivityByExternal(
        auth.token,
        opts.providerId,
        opts.externalActivityId
      );
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  viva
    .command('tenant-provider-learning-activity-patch')
    .description('PATCH provider-scoped learning course activity (beta)')
    .requiredOption('--provider-id <id>', 'learningProvider id')
    .requiredOption('--activity-id <id>', 'learningCourseActivity id')
    .requiredOption('--body-file <path>', 'JSON object for PATCH body')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(
      async (
        opts: { providerId: string; activityId: string; bodyFile: string; token?: string; identity?: string },
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
        const r = await patchTenantProviderLearningCourseActivity(auth.token, opts.providerId, opts.activityId, body);
        if (!r.ok) {
          console.error(`Error: ${r.error?.message}`);
          process.exit(1);
        }
        console.log(JSON.stringify(r.data ?? null, null, 2));
      }
    );

  viva
    .command('tenant-provider-learning-activity-delete')
    .description('DELETE provider-scoped learning course activity (beta)')
    .requiredOption('--provider-id <id>', 'learningProvider id')
    .requiredOption('--activity-id <id>', 'learningCourseActivity id')
    .option('--if-match <etag>', 'If-Match header value')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(
      async (
        opts: { providerId: string; activityId: string; ifMatch?: string; token?: string; identity?: string },
        cmd: Command
      ) => {
        checkReadOnly(cmd);
        const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
        if (!auth.success || !auth.token) {
          console.error(`Auth error: ${auth.error}`);
          process.exit(1);
        }
        const r = await deleteTenantProviderLearningCourseActivity(
          auth.token,
          opts.providerId,
          opts.activityId,
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
      .command('tenant-engagement-roles-list')
      .description('List tenant /employeeExperience/roles catalog (beta, paged)')
      .option('--json', 'Output as JSON array')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity (default: default)')
  ).action(async (opts: ODataOpts & { json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      failVivaTenant(opts.json, 'Auth error', auth.error);
    }
    const r = await listTenantEngagementRoles(auth.token, odataQ(opts));
    if (!r.ok || !r.data) {
      failVivaTenant(opts.json, 'Error', r.error, 'List failed');
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else for (const row of r.data) console.log(JSON.stringify(row));
  });

  viva
    .command('tenant-engagement-role-create')
    .description('POST tenant engagement role (beta)')
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
      const r = await createTenantEngagementRole(auth.token, body);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  viva
    .command('tenant-engagement-role-get')
    .description('GET tenant engagement role by id (beta)')
    .requiredOption('--role-id <id>', 'engagementRole id')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(async (opts: { roleId: string; token?: string; identity?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getTenantEngagementRole(auth.token, opts.roleId);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  viva
    .command('tenant-engagement-role-patch')
    .description('PATCH tenant engagement role (beta)')
    .requiredOption('--role-id <id>', 'engagementRole id')
    .requiredOption('--body-file <path>', 'JSON object for PATCH body')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(async (opts: { roleId: string; bodyFile: string; token?: string; identity?: string }, cmd: Command) => {
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
      const r = await patchTenantEngagementRole(auth.token, opts.roleId, body);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  viva
    .command('tenant-engagement-role-delete')
    .description('DELETE tenant engagement role (beta)')
    .requiredOption('--role-id <id>', 'engagementRole id')
    .option('--if-match <etag>', 'If-Match header value')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(async (opts: { roleId: string; ifMatch?: string; token?: string; identity?: string }, cmd: Command) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await deleteTenantEngagementRole(auth.token, opts.roleId, opts.ifMatch);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
    });

  addODataListOpts(
    viva
      .command('tenant-engagement-role-members-list')
      .description('List members on tenant engagement role (beta, paged)')
      .requiredOption('--role-id <id>', 'engagementRole id')
      .option('--json', 'Output as JSON array')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity (default: default)')
  ).action(async (opts: ODataOpts & { roleId: string; json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      failVivaTenant(opts.json, 'Auth error', auth.error);
    }
    const r = await listTenantEngagementRoleMembers(auth.token, opts.roleId, odataQ(opts));
    if (!r.ok || !r.data) {
      failVivaTenant(opts.json, 'Error', r.error, 'List failed');
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else for (const row of r.data) console.log(JSON.stringify(row));
  });

  viva
    .command('tenant-engagement-role-member-create')
    .description('POST member on tenant engagement role (beta)')
    .requiredOption('--role-id <id>', 'engagementRole id')
    .requiredOption('--body-file <path>', 'JSON object for POST body')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(async (opts: { roleId: string; bodyFile: string; token?: string; identity?: string }, cmd: Command) => {
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
      const r = await createTenantEngagementRoleMember(auth.token, opts.roleId, body);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  viva
    .command('tenant-engagement-role-member-get')
    .description('GET tenant engagement role member (beta)')
    .requiredOption('--role-id <id>', 'engagementRole id')
    .requiredOption('--member-id <id>', 'engagementRoleMember id')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(async (opts: { roleId: string; memberId: string; token?: string; identity?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getTenantEngagementRoleMember(auth.token, opts.roleId, opts.memberId);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  viva
    .command('tenant-engagement-role-member-patch')
    .description('PATCH tenant engagement role member (beta)')
    .requiredOption('--role-id <id>', 'engagementRole id')
    .requiredOption('--member-id <id>', 'engagementRoleMember id')
    .requiredOption('--body-file <path>', 'JSON object for PATCH body')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(
      async (
        opts: { roleId: string; memberId: string; bodyFile: string; token?: string; identity?: string },
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
        const r = await patchTenantEngagementRoleMember(auth.token, opts.roleId, opts.memberId, body);
        if (!r.ok) {
          console.error(`Error: ${r.error?.message}`);
          process.exit(1);
        }
        console.log(JSON.stringify(r.data ?? null, null, 2));
      }
    );

  viva
    .command('tenant-engagement-role-member-delete')
    .description('DELETE tenant engagement role member (beta)')
    .requiredOption('--role-id <id>', 'engagementRole id')
    .requiredOption('--member-id <id>', 'engagementRoleMember id')
    .option('--if-match <etag>', 'If-Match header value')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(
      async (
        opts: { roleId: string; memberId: string; ifMatch?: string; token?: string; identity?: string },
        cmd: Command
      ) => {
        checkReadOnly(cmd);
        const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
        if (!auth.success || !auth.token) {
          console.error(`Auth error: ${auth.error}`);
          process.exit(1);
        }
        const r = await deleteTenantEngagementRoleMember(auth.token, opts.roleId, opts.memberId, opts.ifMatch);
        if (!r.ok) {
          console.error(`Error: ${r.error?.message}`);
          process.exit(1);
        }
      }
    );

  viva
    .command('tenant-engagement-role-member-user-get')
    .description('GET nested user on tenant engagement role member (beta)')
    .requiredOption('--role-id <id>', 'engagementRole id')
    .requiredOption('--member-id <id>', 'engagementRoleMember id')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(async (opts: { roleId: string; memberId: string; token?: string; identity?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getTenantEngagementRoleMemberUser(auth.token, opts.roleId, opts.memberId);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  viva
    .command('tenant-engagement-role-member-user-mailbox-settings-get')
    .description('GET mailboxSettings on tenant engagement role member user (beta)')
    .requiredOption('--role-id <id>', 'engagementRole id')
    .requiredOption('--member-id <id>', 'engagementRoleMember id')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(async (opts: { roleId: string; memberId: string; token?: string; identity?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getTenantEngagementRoleMemberUserMailboxSettings(auth.token, opts.roleId, opts.memberId);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    });

  viva
    .command('tenant-engagement-role-member-user-mailbox-settings-patch')
    .description('PATCH mailboxSettings on tenant engagement role member user (beta)')
    .requiredOption('--role-id <id>', 'engagementRole id')
    .requiredOption('--member-id <id>', 'engagementRoleMember id')
    .requiredOption('--body-file <path>', 'JSON object for PATCH body')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .action(
      async (
        opts: { roleId: string; memberId: string; bodyFile: string; token?: string; identity?: string },
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
        const r = await patchTenantEngagementRoleMemberUserMailboxSettings(
          auth.token,
          opts.roleId,
          opts.memberId,
          body
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
      .command('tenant-engagement-role-member-user-service-provisioning-errors-list')
      .description('List serviceProvisioningErrors on tenant engagement role member user (beta, paged)')
      .requiredOption('--role-id <id>', 'engagementRole id')
      .requiredOption('--member-id <id>', 'engagementRoleMember id')
      .option('--json', 'Output as JSON array')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity (default: default)')
  ).action(
    async (
      opts: ODataOpts & { roleId: string; memberId: string; json?: boolean; token?: string; identity?: string }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        failVivaTenant(opts.json, 'Auth error', auth.error);
      }
      const r = await listTenantEngagementRoleMemberUserServiceProvisioningErrors(
        auth.token,
        opts.roleId,
        opts.memberId,
        odataQ(opts)
      );
      if (!r.ok || !r.data) {
        failVivaTenant(opts.json, 'Error', r.error, 'List failed');
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else for (const row of r.data) console.log(JSON.stringify(row));
    }
  );
}
