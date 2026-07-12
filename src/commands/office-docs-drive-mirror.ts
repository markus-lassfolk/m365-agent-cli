/**
 * Registers drive-item subcommands on `word` / `powerpoint` that mirror `files` — same Graph calls
 * and flags, so agents can round-trip Office docs without switching command roots.
 */
import { readFile } from 'node:fs/promises';
import type { Command } from 'commander';
import type { DriveLocationCliFlags } from '../lib/drive-location.js';
import { registerDriveLocationCliOptions, resolveDriveLocationForCli } from '../lib/drive-location-cli.js';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import {
  assignDriveItemSensitivityLabel,
  checkinFile,
  checkoutFile,
  createOfficeCollaborationLink,
  type DriveItemReference,
  deleteDriveItemPermission,
  deleteFile,
  downloadConvertedFile,
  extractDriveItemSensitivityLabels,
  followDriveItem,
  getDriveItemListItem,
  getDriveItemPermission,
  getDriveItemRetentionLabel,
  getFileAnalytics,
  graphApiRoot,
  inviteDriveItem,
  listDriveItemPermissions,
  listFileVersions,
  moveDriveItem,
  patchDriveItemPermission,
  permanentDeleteDriveItem,
  pollGraphAsyncJob,
  removeDriveItemRetentionLabel,
  restoreFileVersion,
  shareFile,
  startCopyDriveItem,
  unfollowDriveItem,
  uploadFile,
  uploadLargeFile
} from '../lib/graph-client.js';
import { type ItemActivity, listDriveItemActivities } from '../lib/graph-insights-client.js';
import { checkReadOnly } from '../lib/utils.js';

type DriveLocOpts = DriveLocationCliFlags;

function graphRoot(flags: { beta?: boolean }): string {
  return graphApiRoot(!!flags.beta);
}

function withDrive<T extends Command>(cmd: T): T {
  return registerDriveLocationCliOptions(cmd) as T;
}

function parseLoc(flags: DriveLocOpts) {
  return resolveDriveLocationForCli(flags);
}

function parseFolderRef(folder?: string): DriveItemReference | undefined {
  if (!folder) return undefined;
  const trimmed = folder.trim();
  if (!trimmed) return undefined;
  return { id: trimmed };
}

function formatBytes(bytes?: number): string {
  if (bytes === undefined) return '-';
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  if (bytes < 1024 * 1024 * 1024) return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  return `${(bytes / (1024 * 1024 * 1024)).toFixed(1)} GB`;
}

function renderActivityLine(a: ItemActivity): void {
  const when = a.times?.recordedDateTime ?? a.times?.observedDateTime ?? '';
  const who = a.actor?.user?.displayName ?? a.actor?.user?.email ?? a.actor?.application?.displayName ?? '';
  const action = a.action ? Object.keys(a.action).join(',') : '';
  console.log(`${when}\t${action}\t${who}${a.id ? `\tid:${a.id}` : ''}`);
}

export function registerOfficeDriveMirroredCommands(parent: Command): void {
  withDrive(
    parent
      .command('upload <path>')
      .description('Upload a file (same as `files upload`; ≤250MB by default Graph limit)')
      .option('--folder <id>', 'Target folder item ID')
      .option('--json', 'Output as JSON')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity')
  ).action(
    async (
      path: string,
      opts: DriveLocOpts & { folder?: string; json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const loc = parseLoc(opts);
      const result = await uploadFile(auth.token, path, parseFolderRef(opts.folder), loc, graphRoot(opts));
      if (!result.ok || !result.data) {
        console.error(`Error: ${result.error?.message || 'Upload failed'}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(result.data, null, 2));
        return;
      }
      console.log(`✓ Uploaded: ${result.data.name}`);
      console.log(`  ID: ${result.data.id}`);
      if (result.data.webUrl) console.log(`  URL: ${result.data.webUrl}`);
    }
  );

  withDrive(
    parent
      .command('upload-large <path>')
      .description('Chunked upload up to ~4GB (same as `files upload-large`)')
      .option('--folder <id>', 'Target folder item ID')
      .option('--json', 'Output as JSON')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity')
  ).action(
    async (
      path: string,
      opts: DriveLocOpts & { folder?: string; json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const loc = parseLoc(opts);
      const result = await uploadLargeFile(auth.token, path, parseFolderRef(opts.folder), loc, graphRoot(opts));
      if (!result.ok || !result.data) {
        console.error(`Error: ${result.error?.message || 'Failed to upload file'}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(result.data, null, 2));
        return;
      }
      if (result.data.driveItem) {
        console.log(`✓ Uploaded: ${result.data.driveItem.name}`);
        console.log(`  ID: ${result.data.driveItem.id}`);
        if (result.data.driveItem.webUrl) console.log(`  URL: ${result.data.driveItem.webUrl}`);
      } else {
        console.log('✓ Large upload session created');
        console.log(`  Upload URL: ${result.data.uploadUrl}`);
        if (result.data.expirationDateTime) console.log(`  Expires: ${result.data.expirationDateTime}`);
      }
    }
  );

  withDrive(
    parent
      .command('delete <fileId>')
      .description('Delete a drive item (same as `files delete`)')
      .option('--json', 'Output as JSON')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity')
  ).action(
    async (
      fileId: string,
      opts: DriveLocOpts & { json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const loc = parseLoc(opts);
      const result = await deleteFile(auth.token, fileId, loc, graphRoot(opts));
      if (!result.ok) {
        console.error(`Error: ${result.error?.message || 'Delete failed'}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify({ success: true, id: fileId }, null, 2));
        return;
      }
      console.log(`✓ Deleted file: ${fileId}`);
    }
  );

  withDrive(
    parent
      .command('share <fileId>')
      .description('Create sharing link or Office Online collaboration handoff (same as `files share`)')
      .option('--type <type>', 'Link type: view or edit', 'view')
      .option('--scope <scope>', 'Link scope: org or anonymous', 'org')
      .option('--collab', 'Office Online collaboration handoff (edit/org + webUrl)')
      .option('--lock', 'Checkout before collaboration link (with --collab)')
      .option('--expiration <iso>', 'Link expiration, ISO 8601 (e.g. 2026-01-01T00:00:00Z)')
      .option('--password <password>', 'Sharing-link password (OneDrive Personal only, per Graph)')
      .option(
        '--no-retain-inherited-permissions',
        'On first share of this item, remove existing inherited permissions instead of keeping them (Graph default: retain)'
      )
      .option('--json', 'Output as JSON')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity')
  ).action(
    async (
      fileId: string,
      opts: DriveLocOpts & {
        type?: 'view' | 'edit';
        scope?: 'org' | 'anonymous';
        collab?: boolean;
        lock?: boolean;
        expiration?: string;
        password?: string;
        retainInheritedPermissions?: boolean;
        json?: boolean;
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
      if (opts.lock && !opts.collab) {
        console.error('Error: --lock is only supported together with --collab.');
        process.exit(1);
      }
      if (opts.collab && (opts.expiration || opts.password || opts.retainInheritedPermissions === false)) {
        console.error(
          'Error: --expiration/--password/--no-retain-inherited-permissions are not supported together with --collab (collaboration links are always edit/organization).'
        );
        process.exit(1);
      }
      const loc = parseLoc(opts);
      if (opts.collab) {
        const result = await createOfficeCollaborationLink(
          auth.token,
          fileId,
          { lock: opts.lock },
          loc,
          graphRoot(opts)
        );
        if (!result.ok || !result.data) {
          console.error(`Error: ${result.error?.message || 'Collaboration link creation failed'}`);
          process.exit(1);
        }
        if (opts.json) {
          console.log(JSON.stringify(result.data, null, 2));
          return;
        }
        console.log('✓ Office Online collaboration handoff created');
        console.log(`  File: ${result.data.item.name}`);
        console.log(`  File ID: ${result.data.item.id}`);
        if (result.data.link.webUrl) console.log(`  Share URL: ${result.data.link.webUrl}`);
        if (result.data.collaborationUrl) console.log(`  Open in Office Online: ${result.data.collaborationUrl}`);
        console.log('  Mode: edit / organization');
        console.log(`  Lock acquired: ${result.data.lockAcquired ? 'yes' : 'no'}`);
        return;
      }
      const type = opts.type === 'edit' ? 'edit' : 'view';
      const scope = opts.scope === 'anonymous' ? 'anonymous' : 'organization';
      const result = await shareFile(auth.token, fileId, type, scope, loc, graphRoot(opts), {
        expirationDateTime: opts.expiration,
        password: opts.password,
        retainInheritedPermissions: opts.retainInheritedPermissions
      });
      if (!result.ok || !result.data) {
        console.error(`Error: ${result.error?.message || 'Share failed'}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(result.data, null, 2));
        return;
      }
      console.log(`✓ Sharing link created (${type}/${scope})`);
      if (result.data.webUrl) console.log(`  URL: ${result.data.webUrl}`);
      if (result.data.id) console.log(`  Link ID: ${result.data.id}`);
    }
  );

  withDrive(
    parent
      .command('invite <fileId>')
      .description('POST driveItem invite (same as `files invite`)')
      .requiredOption('--body <path>', 'Path to JSON body file')
      .option('--json', 'Output as JSON')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity')
  ).action(
    async (
      fileId: string,
      opts: DriveLocOpts & { body: string; json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      let raw: string;
      try {
        raw = await readFile(opts.body, 'utf-8');
      } catch (e) {
        console.error(`Error: could not read --body file: ${e instanceof Error ? e.message : String(e)}`);
        process.exit(1);
      }
      let body: Record<string, unknown>;
      try {
        body = JSON.parse(raw) as Record<string, unknown>;
      } catch {
        console.error('Error: --body must contain valid JSON');
        process.exit(1);
      }
      const loc = parseLoc(opts);
      const result = await inviteDriveItem(auth.token, fileId, body, loc, graphRoot(opts));
      if (!result.ok) {
        console.error(`Error: ${result.error?.message || 'Invite failed'}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(result.data ?? null, null, 2));
        return;
      }
      console.log('✓ Invite request completed');
      if (result.data !== undefined) console.log(JSON.stringify(result.data, null, 2));
    }
  );

  withDrive(
    parent
      .command('permissions <fileId>')
      .description('List permissions (same as `files permissions`)')
      .option('--json', 'Output as JSON')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity')
  ).action(async (fileId: string, opts: DriveLocOpts & { json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const loc = parseLoc(opts);
    const result = await listDriveItemPermissions(auth.token, fileId, loc, graphRoot(opts));
    if (!result.ok || !result.data) {
      console.error(`Error: ${result.error?.message || 'Failed to list permissions'}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify({ permissions: result.data }, null, 2));
      return;
    }
    if (result.data.length === 0) {
      console.log('No permissions returned.');
      return;
    }
    for (const p of result.data) {
      const id = typeof p.id === 'string' ? p.id : '';
      const roles = Array.isArray(p.roles) ? (p.roles as string[]).join(',') : '';
      console.log(`${id}\t${roles}`);
    }
  });

  withDrive(
    parent
      .command('permission-get <fileId> <permissionId>')
      .description('GET a single permission (same as `files permission-get`)')
      .option('--json', 'Output as JSON')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity')
  ).action(
    async (
      fileId: string,
      permissionId: string,
      opts: DriveLocOpts & { json?: boolean; token?: string; identity?: string }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const loc = parseLoc(opts);
      const result = await getDriveItemPermission(auth.token, fileId, permissionId, loc, graphRoot(opts));
      if (!result.ok || !result.data) {
        console.error(`Error: ${result.error?.message || 'Failed to get permission'}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(result.data, null, 2));
        return;
      }
      const roles = Array.isArray(result.data.roles) ? (result.data.roles as string[]).join(',') : '';
      console.log(`${result.data.id}\t${roles}`);
    }
  );

  withDrive(
    parent
      .command('permission-remove <fileId> <permissionId>')
      .description('DELETE a permission (same as `files permission-remove`)')
      .option('--json', 'Output as JSON')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity')
  ).action(
    async (
      fileId: string,
      permissionId: string,
      opts: DriveLocOpts & { json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const loc = parseLoc(opts);
      const result = await deleteDriveItemPermission(auth.token, fileId, permissionId, loc, graphRoot(opts));
      if (!result.ok) {
        console.error(`Error: ${result.error?.message || 'Failed to remove permission'}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify({ success: true, fileId, permissionId }, null, 2));
        return;
      }
      console.log(`✓ Removed permission ${permissionId} from item ${fileId}`);
    }
  );

  withDrive(
    parent
      .command('permission-update <fileId> <permissionId>')
      .description('PATCH a permission (same as `files permission-update`)')
      .requiredOption('--json-file <path>', 'JSON body for PATCH')
      .option('--json', 'Output as JSON')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity')
  ).action(
    async (
      fileId: string,
      permissionId: string,
      opts: DriveLocOpts & { jsonFile: string; json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      let raw: string;
      try {
        raw = await readFile(opts.jsonFile, 'utf-8');
      } catch (e) {
        console.error(`Error: could not read --json-file: ${e instanceof Error ? e.message : String(e)}`);
        process.exit(1);
      }
      let body: Record<string, unknown>;
      try {
        body = JSON.parse(raw) as Record<string, unknown>;
      } catch {
        console.error('Error: --json-file must contain valid JSON');
        process.exit(1);
      }
      const loc = parseLoc(opts);
      const result = await patchDriveItemPermission(auth.token, fileId, permissionId, body, loc, graphRoot(opts));
      if (!result.ok) {
        console.error(`Error: ${result.error?.message || 'Permission update failed'}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(result.data ?? null, null, 2));
        return;
      }
      console.log('✓ Permission updated');
    }
  );

  withDrive(
    parent
      .command('copy <fileId>')
      .description('Copy item (same as `files copy`)')
      .requiredOption('--parent-id <id>', 'Destination folder item id')
      .option('--parent-drive-id <id>', 'Destination drive id when copying across drives')
      .option('--name <name>', 'Optional new file name')
      .option('--wait', 'Poll async job until completion')
      .option('--json', 'Output as JSON')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity')
  ).action(
    async (
      fileId: string,
      opts: DriveLocOpts & {
        parentId: string;
        parentDriveId?: string;
        name?: string;
        wait?: boolean;
        json?: boolean;
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
      const loc = parseLoc(opts);
      const body = {
        parentReference: {
          id: opts.parentId.trim(),
          ...(opts.parentDriveId?.trim() ? { driveId: opts.parentDriveId.trim() } : {})
        },
        ...(opts.name?.trim() ? { name: opts.name.trim() } : {})
      };
      const started = await startCopyDriveItem(auth.token, fileId, body, loc, graphRoot(opts));
      if (!started.ok || !started.data) {
        console.error(`Error: ${started.error?.message || 'Copy failed'}`);
        process.exit(1);
      }
      const monitorUrl = started.data.monitorUrl;
      if (opts.wait && monitorUrl) {
        const done = await pollGraphAsyncJob(auth.token, monitorUrl);
        if (!done.ok || !done.data) {
          console.error(`Error: ${done.error?.message || 'Copy job failed'}`);
          process.exit(1);
        }
        if (opts.json) {
          console.log(JSON.stringify(done.data, null, 2));
          return;
        }
        console.log('✓ Copy completed');
        console.log(JSON.stringify(done.data, null, 2));
        return;
      }
      if (opts.json) {
        console.log(JSON.stringify(started.data, null, 2));
        return;
      }
      if (monitorUrl) {
        console.log('Copy job started (HTTP 202). Poll the monitor URL:');
        console.log(monitorUrl);
        console.log('Re-run with --wait to poll until completion.');
      } else {
        console.log('Copy request finished.');
        console.log(JSON.stringify(started.data, null, 2));
      }
    }
  );

  withDrive(
    parent
      .command('move <fileId>')
      .description('Move item (same as `files move`)')
      .requiredOption('--parent-id <id>', 'Destination folder item id')
      .option('--parent-drive-id <id>', 'Destination drive id when moving across drives')
      .option('--json', 'Output as JSON')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity')
  ).action(
    async (
      fileId: string,
      opts: DriveLocOpts & {
        parentId: string;
        parentDriveId?: string;
        json?: boolean;
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
      const loc = parseLoc(opts);
      const parentReference: { id: string; driveId?: string } = { id: opts.parentId.trim() };
      if (opts.parentDriveId?.trim()) parentReference.driveId = opts.parentDriveId.trim();
      const result = await moveDriveItem(auth.token, fileId, parentReference, loc, graphRoot(opts));
      if (!result.ok || !result.data) {
        console.error(`Error: ${result.error?.message || 'Move failed'}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(result.data, null, 2));
        return;
      }
      console.log(`✓ Moved item ${fileId} → parent ${opts.parentId}`);
      if (result.data.name) console.log(`  Name: ${result.data.name}`);
      if (result.data.id) console.log(`  New id: ${result.data.id}`);
    }
  );

  withDrive(
    parent
      .command('versions <fileId>')
      .description('List versions (same as `files versions`)')
      .option('--json', 'Output as JSON')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity')
  ).action(async (fileId: string, opts: DriveLocOpts & { json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const loc = parseLoc(opts);
    const result = await listFileVersions(auth.token, fileId, loc, graphRoot(opts));
    if (!result.ok || !result.data) {
      console.error(`Error: ${result.error?.message || 'Failed to list versions'}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify({ versions: result.data }, null, 2));
      return;
    }
    if (result.data.length === 0) {
      console.log('No versions found.');
      return;
    }
    for (const v of result.data) {
      console.log(`VERSION  ${v.id}`);
      if (v.lastModifiedDateTime) console.log(`      Modified: ${v.lastModifiedDateTime}`);
      if (v.size !== undefined) console.log(`      Size:     ${formatBytes(v.size)}`);
      if (v.lastModifiedBy?.user?.displayName) console.log(`      By:       ${v.lastModifiedBy.user.displayName}`);
    }
  });

  withDrive(
    parent
      .command('restore <fileId> <versionId>')
      .description('Restore a version (same as `files restore`)')
      .option('--json', 'Output as JSON')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity')
  ).action(
    async (
      fileId: string,
      versionId: string,
      opts: DriveLocOpts & { json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const loc = parseLoc(opts);
      const result = await restoreFileVersion(auth.token, fileId, versionId, loc, graphRoot(opts));
      if (!result.ok) {
        console.error(`Error: ${result.error?.message || 'Restore failed'}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify({ success: true, fileId, versionId }, null, 2));
        return;
      }
      console.log(`✓ Restored version ${versionId} of file ${fileId}`);
    }
  );

  withDrive(
    parent
      .command('checkout <fileId>')
      .description('Check out an Office document (POST …/checkout; same as `files checkout`)')
      .option('--json', 'Output as JSON')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity')
  ).action(
    async (
      fileId: string,
      opts: DriveLocOpts & { json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const loc = parseLoc(opts);
      const result = await checkoutFile(auth.token, fileId, loc, graphRoot(opts));
      if (!result.ok) {
        console.error(`Error: ${result.error?.message || 'Checkout failed'}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify({ success: true, fileId }, null, 2));
        return;
      }
      console.log(`✓ Checked out: ${fileId}`);
    }
  );

  withDrive(
    parent
      .command('checkin <fileId>')
      .description('Check in (same as `files checkin`)')
      .option('--comment <comment>', 'Optional check-in comment')
      .option('--json', 'Output as JSON')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity')
  ).action(
    async (
      fileId: string,
      opts: DriveLocOpts & { comment?: string; json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const loc = parseLoc(opts);
      const result = await checkinFile(auth.token, fileId, opts.comment, loc, graphRoot(opts));
      if (!result.ok || !result.data) {
        console.error(`Error: ${result.error?.message || 'Check-in failed'}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(result.data, null, 2));
        return;
      }
      console.log('✓ File checked in');
      console.log(`  File: ${result.data.item.name}`);
      console.log(`  File ID: ${result.data.item.id}`);
      if (result.data.comment) console.log(`  Comment: ${result.data.comment}`);
    }
  );

  withDrive(
    parent
      .command('convert <fileId>')
      .description('Download converted format (same as `files convert`)')
      .option('--format <format>', 'Target format (e.g. pdf)', 'pdf')
      .option('--out <path>', 'Output path')
      .option('--json', 'Output as JSON')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity')
  ).action(
    async (
      fileId: string,
      opts: DriveLocOpts & { format: string; out?: string; json?: boolean; token?: string; identity?: string }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const loc = parseLoc(opts);
      const result = await downloadConvertedFile(auth.token, fileId, opts.format, opts.out, loc, graphRoot(opts));
      if (!result.ok || !result.data) {
        console.error(`Error: ${result.error?.message || 'Conversion download failed'}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(result.data, null, 2));
        return;
      }
      console.log('✓ Converted file downloaded');
      console.log(`  Saved to: ${result.data.path}`);
    }
  );

  withDrive(
    parent
      .command('analytics <fileId>')
      .description('File analytics (same as `files analytics`)')
      .option('--json', 'Output as JSON')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity')
  ).action(async (fileId: string, opts: DriveLocOpts & { json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const loc = parseLoc(opts);
    const result = await getFileAnalytics(auth.token, fileId, loc, graphRoot(opts));
    if (!result.ok || !result.data) {
      console.error(`Error: ${result.error?.message || 'Failed to get file analytics'}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(result.data, null, 2));
      return;
    }
    console.log(`✓ Analytics for ${fileId}:`);
    if (result.data.allTime?.access) {
      console.log('  All Time:');
      console.log(`    Actions: ${result.data.allTime.access.actionCount || 0}`);
      console.log(`    Actors:  ${result.data.allTime.access.actorCount || 0}`);
    }
    if (result.data.lastSevenDays?.access) {
      console.log('  Last 7 Days:');
      console.log(`    Actions: ${result.data.lastSevenDays.access.actionCount || 0}`);
      console.log(`    Actors:  ${result.data.lastSevenDays.access.actorCount || 0}`);
    }
    if (!result.data.allTime?.access && !result.data.lastSevenDays?.access) {
      console.log('  No analytics data available.');
    }
  });

  withDrive(
    parent
      .command('activities <fileId>')
      .description('Per-item activity feed (same as `files activities`)')
      .option('--top <n>', 'Limit ($top, max 200)')
      .option('--json', 'Output as JSON')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity')
  ).action(
    async (
      fileId: string,
      opts: DriveLocOpts & { top?: string; json?: boolean; token?: string; identity?: string }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const top = opts.top ? Number.parseInt(opts.top, 10) : undefined;
      if (opts.top && (!Number.isFinite(top) || (top as number) <= 0)) {
        console.error('Error: --top must be a positive integer');
        process.exit(1);
      }
      const loc = parseLoc(opts);
      const r = await listDriveItemActivities(auth.token, loc, fileId, { top, graphBaseUrl: graphRoot(opts) });
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message ?? 'Activities request failed'}`);
        process.exit(1);
      }
      const items = r.data.value ?? [];
      if (opts.json) {
        console.log(JSON.stringify(r.data, null, 2));
        return;
      }
      if (items.length === 0) {
        console.log('No activities recorded.');
        return;
      }
      for (const a of items) renderActivityLine(a);
    }
  );

  withDrive(
    parent
      .command('list-item <fileId>')
      .description('GET SharePoint listItem (`…/listItem`); same as `files list-item`')
      .option('--json', 'Output as JSON')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity')
  ).action(async (fileId: string, opts: DriveLocOpts & { json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const loc = parseLoc(opts);
    const r = await getDriveItemListItem(auth.token, fileId, loc, graphRoot(opts));
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message || 'listItem failed'}`);
      process.exit(1);
    }
    console.log(opts.json ? JSON.stringify(r.data, null, 2) : JSON.stringify(r.data, null, 2));
  });

  withDrive(
    parent
      .command('follow <fileId>')
      .description('POST …/follow (same as `files follow`)')
      .option('--json', 'Output as JSON')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity')
  ).action(
    async (
      fileId: string,
      opts: DriveLocOpts & { json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const loc = parseLoc(opts);
      const r = await followDriveItem(auth.token, fileId, loc, graphRoot(opts));
      if (!r.ok) {
        console.error(`Error: ${r.error?.message || 'follow failed'}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify({ success: true, fileId }, null, 2));
        return;
      }
      console.log(`✓ Following: ${fileId}`);
    }
  );

  withDrive(
    parent
      .command('unfollow <fileId>')
      .description('POST …/unfollow (same as `files unfollow`)')
      .option('--json', 'Output as JSON')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity')
  ).action(
    async (
      fileId: string,
      opts: DriveLocOpts & { json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const loc = parseLoc(opts);
      const r = await unfollowDriveItem(auth.token, fileId, loc, graphRoot(opts));
      if (!r.ok) {
        console.error(`Error: ${r.error?.message || 'unfollow failed'}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify({ success: true, fileId }, null, 2));
        return;
      }
      console.log(`✓ Unfollowed: ${fileId}`);
    }
  );

  withDrive(
    parent
      .command('sensitivity-assign <fileId>')
      .description('POST assignSensitivityLabel (same as `files sensitivity-assign`)')
      .requiredOption('--json-file <path>', 'JSON parameters (see Graph driveItem assignSensitivityLabel)')
      .option('--json', 'Output as JSON')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity')
  ).action(
    async (
      fileId: string,
      opts: DriveLocOpts & { jsonFile: string; json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      let raw: string;
      try {
        raw = await readFile(opts.jsonFile, 'utf-8');
      } catch (e) {
        console.error(`Error: could not read --json-file: ${e instanceof Error ? e.message : String(e)}`);
        process.exit(1);
      }
      let body: Record<string, unknown>;
      try {
        body = JSON.parse(raw) as Record<string, unknown>;
      } catch {
        console.error('Error: --json-file must contain valid JSON');
        process.exit(1);
      }
      const loc = parseLoc(opts);
      const r = await assignDriveItemSensitivityLabel(auth.token, fileId, body, loc, graphRoot(opts));
      if (!r.ok) {
        console.error(`Error: ${r.error?.message || 'assignSensitivityLabel failed'}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(r.data ?? null, null, 2));
        return;
      }
      console.log('✓ Sensitivity label assignment request completed');
      if (r.data !== undefined) console.log(JSON.stringify(r.data, null, 2));
    }
  );

  withDrive(
    parent
      .command('sensitivity-extract <fileId>')
      .description('POST extractSensitivityLabels (same as `files sensitivity-extract`)')
      .option('--json', 'Output as JSON')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity')
  ).action(async (fileId: string, opts: DriveLocOpts & { json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const loc = parseLoc(opts);
    const r = await extractDriveItemSensitivityLabels(auth.token, fileId, loc, graphRoot(opts));
    if (!r.ok) {
      console.error(`Error: ${r.error?.message || 'extractSensitivityLabels failed'}`);
      process.exit(1);
    }
    console.log(opts.json ? JSON.stringify(r.data ?? null, null, 2) : JSON.stringify(r.data ?? null, null, 2));
  });

  withDrive(
    parent
      .command('permanent-delete <fileId>')
      .description('POST permanentDelete (same as `files permanent-delete`)')
      .option('--json', 'Output as JSON')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity')
  ).action(
    async (
      fileId: string,
      opts: DriveLocOpts & { json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const loc = parseLoc(opts);
      const r = await permanentDeleteDriveItem(auth.token, fileId, loc, graphRoot(opts));
      if (!r.ok) {
        console.error(`Error: ${r.error?.message || 'permanentDelete failed'}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify({ success: true, fileId }, null, 2));
        return;
      }
      console.log(`✓ Permanently deleted: ${fileId}`);
    }
  );

  withDrive(
    parent
      .command('retention-label <fileId>')
      .description('GET …/retentionLabel (same as `files retention-label`)')
      .option('--json', 'Output as JSON')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity')
  ).action(async (fileId: string, opts: DriveLocOpts & { json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const loc = parseLoc(opts);
    const r = await getDriveItemRetentionLabel(auth.token, fileId, loc, graphRoot(opts));
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message || 'getRetentionLabel failed'}`);
      process.exit(1);
    }
    console.log(opts.json ? JSON.stringify(r.data, null, 2) : JSON.stringify(r.data, null, 2));
  });

  withDrive(
    parent
      .command('retention-label-remove <fileId>')
      .description('DELETE …/retentionLabel (same as `files retention-label-remove`)')
      .option('--if-match <etag>', 'Optional If-Match header')
      .option('--json', 'Output as JSON')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity')
  ).action(
    async (
      fileId: string,
      opts: DriveLocOpts & { ifMatch?: string; json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const loc = parseLoc(opts);
      const r = await removeDriveItemRetentionLabel(auth.token, fileId, opts.ifMatch, loc, graphRoot(opts));
      if (!r.ok) {
        console.error(`Error: ${r.error?.message || 'removeRetentionLabel failed'}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify({ success: true, fileId }, null, 2));
        return;
      }
      console.log(`✓ Removed retention label from: ${fileId}`);
    }
  );
}
