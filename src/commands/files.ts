import { readFile } from 'node:fs/promises';
import { Command } from 'commander';
import type { DriveLocation, DriveLocationCliFlags } from '../lib/drive-location.js';
import { registerDriveLocationCliOptions, resolveDriveLocationForCli } from '../lib/drive-location-cli.js';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import {
  assignDriveItemSensitivityLabel,
  callGraphAt,
  checkinFile,
  checkoutFile,
  createOfficeCollaborationLink,
  type DriveItem,
  type DriveItemReference,
  defaultDownloadPath,
  deleteDriveItemPermission,
  deleteFile,
  downloadConvertedFile,
  downloadFile,
  extractDriveItemSensitivityLabels,
  followDriveItem,
  getDriveItemDeltaPage,
  getDriveItemListItem,
  getDriveItemPermission,
  getDriveItemRetentionLabel,
  getFileAnalytics,
  getFileMetadata,
  graphApiRoot,
  inviteDriveItem,
  listDriveItemPermissions,
  listDriveItemThumbnails,
  listDriveSharedWithMe,
  listFiles,
  listFileVersions,
  moveDriveItem,
  patchDriveItemPermission,
  permanentDeleteDriveItem,
  pollGraphAsyncJob,
  removeDriveItemRetentionLabel,
  restoreDeletedDriveItem,
  restoreFileVersion,
  searchFiles,
  shareFile,
  startCopyDriveItem,
  unfollowDriveItem,
  uploadFile,
  uploadLargeFile
} from '../lib/graph-client.js';
import {
  applyDeltaPageToState,
  assertDeltaScopeMatchesState,
  driveDeltaScopeFromLocation,
  readDeltaStateFile,
  resolveDeltaContinuationUrl,
  writeDeltaStateFile
} from '../lib/graph-delta-state-file.js';
import { createDriveItemPreview, type ItemActivity, listDriveItemActivities } from '../lib/graph-insights-client.js';
import { checkReadOnly } from '../lib/utils.js';

/** CLI flags shared by all `files` subcommands that hit a drive root (includes `--beta`). */
export type FilesDriveCliFlags = DriveLocationCliFlags;

function graphRoot(flags: { beta?: boolean }): string {
  return graphApiRoot(!!flags.beta);
}

function withDriveOptions<T extends Command>(cmd: T): T {
  return registerDriveLocationCliOptions(cmd) as T;
}

function parseDriveLocation(flags: FilesDriveCliFlags): DriveLocation {
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

function renderItems(items: DriveItem[]): void {
  if (items.length === 0) {
    console.log('No files found.');
    return;
  }

  for (const item of items) {
    const kind = item.folder ? 'DIR ' : 'FILE';
    console.log(`${kind}  ${item.id}`);
    console.log(`      Name: ${item.name}`);
    console.log(`      Size: ${formatBytes(item.size)}`);
    if (item.webUrl) console.log(`      URL:  ${item.webUrl}`);
    if (item.lastModifiedDateTime) console.log(`      Modified: ${item.lastModifiedDateTime}`);
  }
}

export const filesCommand = new Command('files')
  .description(
    'OneDrive and SharePoint drives via Microsoft Graph: browse, upload, share, permissions, versions, labels.'
  )
  .addHelpText(
    'after',
    `
Examples:
  m365-agent-cli files list
  m365-agent-cli files search "report"
  m365-agent-cli files upload ./deck.pptx --folder <folderId>
  m365-agent-cli files download <fileId> --out ./out.bin

Drive targets use the same flags as other file commands (default /me/drive). See docs/CLI_REFERENCE.md and GRAPH_SCOPES.md.
`
  );

withDriveOptions(filesCommand.command('list'))
  .summary('List folder or drive root children')
  .description('List files in drive root or a folder')
  .option('--folder <id>', 'Folder item ID')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (options: FilesDriveCliFlags & { folder?: string; json?: boolean; token?: string; identity?: string }) => {
      const auth = await resolveGraphAuth({ token: options.token, identity: options.identity });
      if (!auth.success) {
        console.error(`Error: ${auth.error}`);
        process.exit(1);
      }

      const loc = parseDriveLocation(options);
      const result = await listFiles(auth.token!, parseFolderRef(options.folder), loc, graphRoot(options));
      if (!result.ok || !result.data) {
        console.error(`Error: ${result.error?.message || 'Failed to list files'}`);
        process.exit(1);
      }

      if (options.json) {
        console.log(JSON.stringify({ items: result.data }, null, 2));
        return;
      }

      renderItems(result.data);
    }
  );

withDriveOptions(filesCommand.command('delta'))
  .summary('Incremental sync with deltaLink state')
  .description(
    'One page of drive item delta sync (root or --folder). Use --url for @odata.nextLink/@odata.deltaLink; optional --state-file persists cursor (kind: driveDelta).'
  )
  .option('--folder <id>', 'Folder item id (delta under that folder instead of drive root)')
  .option('--url <url>', 'Full nextLink or deltaLink URL (overrides state-file continuation)')
  .option('--state-file <path>', 'Read/write JSON delta cursor')
  .option('--json', 'Output raw page JSON (value, nextLink, deltaLink)')
  .option('--token <token>', 'Use a specific Graph token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      opts: FilesDriveCliFlags & {
        folder?: string;
        url?: string;
        stateFile?: string;
        json?: boolean;
        token?: string;
        identity?: string;
      }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Error: ${auth.error}`);
        process.exit(1);
      }

      const loc = parseDriveLocation(opts);
      const driveScope = driveDeltaScopeFromLocation(loc, opts.folder);
      const existingState = opts.stateFile ? await readDeltaStateFile(opts.stateFile) : null;
      if (existingState && existingState.kind !== 'driveDelta') {
        console.error('Error: state file is not for files drive delta (kind must be driveDelta).');
        process.exit(1);
      }
      try {
        if (existingState) {
          assertDeltaScopeMatchesState(existingState, driveScope);
        }
      } catch (err) {
        console.error(err instanceof Error ? err.message : err);
        process.exit(1);
      }

      const continueUrl = resolveDeltaContinuationUrl({ explicitNext: opts.url, state: existingState });
      const r = await getDriveItemDeltaPage(auth.token, {
        location: loc,
        folderItemId: opts.folder?.trim(),
        nextOrDeltaLink: continueUrl,
        graphBaseUrl: graphRoot(opts)
      });
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message || 'Drive delta failed'}`);
        process.exit(1);
      }

      if (opts.stateFile) {
        const merged = applyDeltaPageToState(existingState, 'driveDelta', r.data, driveScope);
        await writeDeltaStateFile(opts.stateFile, merged);
      }

      if (opts.json) {
        console.log(JSON.stringify(r.data, null, 2));
        return;
      }
      console.log(`Changes: ${r.data.value?.length ?? 0} item(s)`);
      if (r.data['@odata.nextLink']) console.log(`nextLink: ${r.data['@odata.nextLink']}`);
      if (r.data['@odata.deltaLink']) console.log(`deltaLink: ${r.data['@odata.deltaLink']}`);
      if (opts.stateFile) console.log(`state-file: ${opts.stateFile} (updated)`);
    }
  );

withDriveOptions(filesCommand.command('search <query>'))
  .summary('Search drive by query')
  .description(
    'Search under the drive root only (Microsoft Graph GET …/root/search(q=…)); there is no --folder flag. Folder-scoped search uses …/items/{folderId}/search in Graph — use graph invoke if you need that.'
  )
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (query: string, options: FilesDriveCliFlags & { json?: boolean; token?: string; identity?: string }) => {
      const auth = await resolveGraphAuth({ token: options.token, identity: options.identity });
      if (!auth.success) {
        console.error(`Error: ${auth.error}`);
        process.exit(1);
      }

      const loc = parseDriveLocation(options);
      const result = await searchFiles(auth.token!, query, loc, graphRoot(options));
      if (!result.ok || !result.data) {
        console.error(`Error: ${result.error?.message || 'Failed to search files'}`);
        process.exit(1);
      }

      if (options.json) {
        console.log(JSON.stringify({ items: result.data }, null, 2));
        return;
      }

      renderItems(result.data);
    }
  );

withDriveOptions(filesCommand.command('meta <fileId>'))
  .summary('Show item metadata')
  .description('Get file metadata')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (fileId: string, options: FilesDriveCliFlags & { json?: boolean; token?: string; identity?: string }) => {
      const auth = await resolveGraphAuth({ token: options.token, identity: options.identity });
      if (!auth.success) {
        console.error(`Error: ${auth.error}`);
        process.exit(1);
      }

      const loc = parseDriveLocation(options);
      const result = await getFileMetadata(auth.token!, fileId, loc, graphRoot(options));
      if (!result.ok || !result.data) {
        console.error(`Error: ${result.error?.message || 'Failed to fetch metadata'}`);
        process.exit(1);
      }

      if (options.json) {
        console.log(JSON.stringify(result.data, null, 2));
        return;
      }

      renderItems([result.data]);
    }
  );

withDriveOptions(filesCommand.command('thumbnails <fileId>'))
  .summary('List image thumbnail URLs')
  .description('List thumbnail sets for a drive item (GET …/thumbnails — small/medium/large URLs per Graph)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (fileId: string, options: FilesDriveCliFlags & { json?: boolean; token?: string; identity?: string }) => {
      const auth = await resolveGraphAuth({ token: options.token, identity: options.identity });
      if (!auth.success || !auth.token) {
        console.error(`Error: ${auth.error}`);
        process.exit(1);
      }
      const loc = parseDriveLocation(options);
      const result = await listDriveItemThumbnails(auth.token, fileId, loc, graphRoot(options));
      if (!result.ok || !result.data) {
        console.error(`Error: ${result.error?.message || 'Failed to list thumbnails'}`);
        process.exit(1);
      }
      if (options.json) {
        console.log(JSON.stringify({ thumbnails: result.data }, null, 2));
        return;
      }
      if (result.data.length === 0) {
        console.log('No thumbnails returned (Graph may not have generated any yet).');
        return;
      }
      for (const set of result.data) {
        const id = set.id ?? '(set)';
        console.log(`thumbnailSet ${id}`);
        for (const size of ['small', 'medium', 'large'] as const) {
          const info = set[size];
          if (info?.url) {
            console.log(`  ${size}: ${info.width ?? '?'}x${info.height ?? '?'}  ${info.url}`);
          }
        }
      }
    }
  );

withDriveOptions(filesCommand.command('upload <path>'))
  .summary('Upload a small file')
  .description('Upload a file up to 250MB')
  .option('--folder <id>', 'Target folder item ID')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      path: string,
      options: FilesDriveCliFlags & { folder?: string; json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: options.token, identity: options.identity });
      if (!auth.success) {
        console.error(`Error: ${auth.error}`);
        process.exit(1);
      }

      const loc = parseDriveLocation(options);
      const result = await uploadFile(auth.token!, path, parseFolderRef(options.folder), loc, graphRoot(options));
      if (!result.ok || !result.data) {
        console.error(`Error: ${result.error?.message || 'Upload failed'}`);
        process.exit(1);
      }

      if (options.json) {
        console.log(JSON.stringify(result.data, null, 2));
        return;
      }

      console.log(`✓ Uploaded: ${result.data.name}`);
      console.log(`  ID: ${result.data.id}`);
      if (result.data.webUrl) console.log(`  URL: ${result.data.webUrl}`);
    }
  );

withDriveOptions(filesCommand.command('upload-large <path>'))
  .summary('Resumable large upload')
  .description('Upload a file up to 4GB using a chunked upload session')
  .option('--folder <id>', 'Target folder item ID')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      path: string,
      options: FilesDriveCliFlags & { folder?: string; json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: options.token, identity: options.identity });
      if (!auth.success) {
        console.error(`Error: ${auth.error}`);
        process.exit(1);
      }

      const loc = parseDriveLocation(options);
      const result = await uploadLargeFile(auth.token!, path, parseFolderRef(options.folder), loc, graphRoot(options));
      if (!result.ok || !result.data) {
        console.error(`Error: ${result.error?.message || 'Failed to upload file'}`);
        process.exit(1);
      }

      if (options.json) {
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

withDriveOptions(filesCommand.command('download <fileId>'))
  .summary('Download file bytes')
  .description('Download a file by ID')
  .option('--out <path>', 'Output path (defaults to ~/Downloads/<name>)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      fileId: string,
      options: FilesDriveCliFlags & { out?: string; json?: boolean; token?: string; identity?: string }
    ) => {
      const auth = await resolveGraphAuth({ token: options.token, identity: options.identity });
      if (!auth.success) {
        console.error(`Error: ${auth.error}`);
        process.exit(1);
      }

      const loc = parseDriveLocation(options);
      const meta = options.out ? undefined : await getFileMetadata(auth.token!, fileId, loc, graphRoot(options));
      const defaultOut = meta?.ok && meta.data ? defaultDownloadPath(meta.data.name || fileId) : undefined;
      const result = await downloadFile(
        auth.token!,
        fileId,
        options.out || defaultOut,
        meta?.ok ? meta.data : undefined,
        loc,
        graphRoot(options)
      );
      if (!result.ok || !result.data) {
        console.error(`Error: ${result.error?.message || 'Download failed'}`);
        process.exit(1);
      }

      if (options.json) {
        console.log(JSON.stringify(result.data, null, 2));
        return;
      }

      console.log(`✓ Downloaded: ${result.data.item.name}`);
      console.log(`  Saved to: ${result.data.path}`);
    }
  );

withDriveOptions(filesCommand.command('delete <fileId>'))
  .summary('Delete a drive item')
  .description('Delete a file by ID')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      fileId: string,
      options: FilesDriveCliFlags & { json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: options.token, identity: options.identity });
      if (!auth.success) {
        console.error(`Error: ${auth.error}`);
        process.exit(1);
      }

      const loc = parseDriveLocation(options);
      const result = await deleteFile(auth.token!, fileId, loc, graphRoot(options));
      if (!result.ok) {
        console.error(`Error: ${result.error?.message || 'Delete failed'}`);
        process.exit(1);
      }

      if (options.json) {
        console.log(JSON.stringify({ success: true, id: fileId }, null, 2));
        return;
      }

      console.log(`✓ Deleted file: ${fileId}`);
    }
  );

withDriveOptions(filesCommand.command('restore-deleted <fileId>'))
  .summary('Restore item from recycle bin')
  .description(
    'Restore a deleted drive item from the recycle bin (`POST …/items/{id}/restore`). Optional `--json-file` body (e.g. parentReference).'
  )
  .option('--json-file <path>', 'JSON body for restore (omit or `{}` for default)')
  .option('--json', 'Output restored item as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      fileId: string,
      options: FilesDriveCliFlags & { jsonFile?: string; json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: options.token, identity: options.identity });
      if (!auth.success || !auth.token) {
        console.error(`Error: ${auth.error}`);
        process.exit(1);
      }
      let body: Record<string, unknown> | undefined;
      if (options.jsonFile?.trim()) {
        let raw: string;
        try {
          raw = await readFile(options.jsonFile.trim(), 'utf-8');
        } catch (e) {
          console.error(`Error: could not read --json-file: ${e instanceof Error ? e.message : String(e)}`);
          process.exit(1);
        }
        try {
          body = JSON.parse(raw) as Record<string, unknown>;
        } catch {
          console.error('Error: --json-file must contain valid JSON');
          process.exit(1);
        }
      }
      const loc = parseDriveLocation(options);
      const result = await restoreDeletedDriveItem(auth.token, fileId, body, loc, graphRoot(options));
      if (!result.ok || !result.data) {
        console.error(`Error: ${result.error?.message || 'Restore failed'}`);
        process.exit(1);
      }
      if (options.json) {
        console.log(JSON.stringify(result.data, null, 2));
        return;
      }
      console.log(`✓ Restored item: ${result.data.id ?? fileId}`);
      if (result.data.webUrl) console.log(`  URL: ${result.data.webUrl}`);
    }
  );

withDriveOptions(filesCommand.command('share <fileId>'))
  .summary('Create sharing link')
  .description('Create a sharing link or Office Online collaboration handoff')
  .option('--type <type>', 'Link type: view or edit', 'view')
  .option('--scope <scope>', 'Link scope: org or anonymous', 'org')
  .option('--collab', 'Create an Office Online collaboration handoff (edit/org + webUrl)')
  .option('--lock', 'Checkout the file before creating a collaboration link (use with --collab)')
  .option('--expiration <iso>', 'Link expiration, ISO 8601 (e.g. 2026-01-01T00:00:00Z)')
  .option('--password <password>', 'Sharing-link password (OneDrive Personal only, per Graph)')
  .option(
    '--no-retain-inherited-permissions',
    'On first share of this item, remove existing inherited permissions instead of keeping them (Graph default: retain)'
  )
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      fileId: string,
      options: FilesDriveCliFlags & {
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
      const auth = await resolveGraphAuth({ token: options.token, identity: options.identity });
      if (!auth.success) {
        console.error(`Error: ${auth.error}`);
        process.exit(1);
      }

      if (options.lock && !options.collab) {
        console.error('Error: --lock is only supported together with --collab.');
        process.exit(1);
      }

      const loc = parseDriveLocation(options);

      if (options.collab) {
        const result = await createOfficeCollaborationLink(
          auth.token!,
          fileId,
          { lock: options.lock },
          loc,
          graphRoot(options)
        );
        if (!result.ok || !result.data) {
          console.error(`Error: ${result.error?.message || 'Collaboration link creation failed'}`);
          process.exit(1);
        }

        if (options.json) {
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
        console.log('  Note: real-time co-authoring happens in Office Online after the user opens the document URL.');
        return;
      }

      const type = options.type === 'edit' ? 'edit' : 'view';
      const scope = options.scope === 'anonymous' ? 'anonymous' : 'organization';
      const result = await shareFile(auth.token!, fileId, type, scope, loc, graphRoot(options), {
        expirationDateTime: options.expiration,
        password: options.password,
        retainInheritedPermissions: options.retainInheritedPermissions
      });
      if (!result.ok || !result.data) {
        console.error(`Error: ${result.error?.message || 'Share failed'}`);
        process.exit(1);
      }

      if (options.json) {
        console.log(JSON.stringify(result.data, null, 2));
        return;
      }

      console.log(`✓ Sharing link created (${type}/${scope})`);
      if (result.data.webUrl) console.log(`  URL: ${result.data.webUrl}`);
      if (result.data.id) console.log(`  Link ID: ${result.data.id}`);
    }
  );

withDriveOptions(filesCommand.command('invite <fileId>'))
  .summary('Send invite to people')
  .description('POST driveItem invite — share with specific recipients (JSON body; see Graph driveItem: invite)')
  .requiredOption('--body <path>', 'Path to JSON file for the invite request body')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      fileId: string,
      options: FilesDriveCliFlags & { body: string; json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: options.token, identity: options.identity });
      if (!auth.success) {
        console.error(`Error: ${auth.error}`);
        process.exit(1);
      }

      let raw: string;
      try {
        raw = await readFile(options.body, 'utf-8');
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

      const loc = parseDriveLocation(options);
      const result = await inviteDriveItem(auth.token!, fileId, body, loc, graphRoot(options));
      if (!result.ok) {
        console.error(`Error: ${result.error?.message || 'Invite failed'}`);
        process.exit(1);
      }

      if (options.json) {
        console.log(JSON.stringify(result.data ?? null, null, 2));
        return;
      }

      console.log('✓ Invite request completed');
      if (result.data !== undefined) {
        console.log(JSON.stringify(result.data, null, 2));
      }
    }
  );

withDriveOptions(filesCommand.command('permissions <fileId>'))
  .summary('List sharing permissions')
  .description('List permissions on a drive item (sharing links and grants)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (fileId: string, options: FilesDriveCliFlags & { json?: boolean; token?: string; identity?: string }) => {
      const auth = await resolveGraphAuth({ token: options.token, identity: options.identity });
      if (!auth.success) {
        console.error(`Error: ${auth.error}`);
        process.exit(1);
      }

      const loc = parseDriveLocation(options);
      const result = await listDriveItemPermissions(auth.token!, fileId, loc, graphRoot(options));
      if (!result.ok || !result.data) {
        console.error(`Error: ${result.error?.message || 'Failed to list permissions'}`);
        process.exit(1);
      }

      if (options.json) {
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
    }
  );

withDriveOptions(filesCommand.command('permission-get <fileId> <permissionId>'))
  .summary('Get one permission')
  .description('GET a single permission on a drive item by ID (sharing link or invitation grant)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      fileId: string,
      permissionId: string,
      options: FilesDriveCliFlags & { json?: boolean; token?: string; identity?: string }
    ) => {
      const auth = await resolveGraphAuth({ token: options.token, identity: options.identity });
      if (!auth.success) {
        console.error(`Error: ${auth.error}`);
        process.exit(1);
      }

      const loc = parseDriveLocation(options);
      const result = await getDriveItemPermission(auth.token!, fileId, permissionId, loc, graphRoot(options));
      if (!result.ok || !result.data) {
        console.error(`Error: ${result.error?.message || 'Failed to get permission'}`);
        process.exit(1);
      }

      if (options.json) {
        console.log(JSON.stringify(result.data, null, 2));
        return;
      }

      const roles = Array.isArray(result.data.roles) ? (result.data.roles as string[]).join(',') : '';
      console.log(`${result.data.id}\t${roles}`);
    }
  );

withDriveOptions(filesCommand.command('permission-remove <fileId> <permissionId>'))
  .summary('Remove a permission')
  .description('DELETE a permission from a drive item (revoke access granted by that permission)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      fileId: string,
      permissionId: string,
      options: FilesDriveCliFlags & { json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: options.token, identity: options.identity });
      if (!auth.success) {
        console.error(`Error: ${auth.error}`);
        process.exit(1);
      }

      const loc = parseDriveLocation(options);
      const result = await deleteDriveItemPermission(auth.token!, fileId, permissionId, loc, graphRoot(options));
      if (!result.ok) {
        console.error(`Error: ${result.error?.message || 'Failed to remove permission'}`);
        process.exit(1);
      }

      if (options.json) {
        console.log(JSON.stringify({ success: true, fileId, permissionId }, null, 2));
        return;
      }

      console.log(`✓ Removed permission ${permissionId} from item ${fileId}`);
    }
  );

withDriveOptions(filesCommand.command('permission-update <fileId> <permissionId>'))
  .summary('Update permission role')
  .description('PATCH a permission on a drive item (e.g. change roles — JSON body per Graph driveItem permission)')
  .requiredOption('--json-file <path>', 'Path to JSON body for PATCH')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      fileId: string,
      permissionId: string,
      options: FilesDriveCliFlags & { jsonFile: string; json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: options.token, identity: options.identity });
      if (!auth.success || !auth.token) {
        console.error(`Error: ${auth.error}`);
        process.exit(1);
      }
      let raw: string;
      try {
        raw = await readFile(options.jsonFile, 'utf-8');
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
      const loc = parseDriveLocation(options);
      const result = await patchDriveItemPermission(auth.token, fileId, permissionId, body, loc, graphRoot(options));
      if (!result.ok) {
        console.error(`Error: ${result.error?.message || 'Permission update failed'}`);
        process.exit(1);
      }
      if (options.json) {
        console.log(JSON.stringify(result.data ?? null, null, 2));
        return;
      }
      console.log('✓ Permission updated');
    }
  );

filesCommand
  .command('shared-with-me')
  .summary('List files shared with me')
  .description(
    'List items shared with the signed-in user (GET /me/drive/sharedWithMe only — not available for --user/--drive-id/--site-id)'
  )
  .option('--json', 'Output as JSON')
  .option('--beta', 'Use Microsoft Graph beta API host for this call', false)
  .option('--token <token>', 'Use a specific Graph token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (options: { json?: boolean; beta?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: options.token, identity: options.identity });
    if (!auth.success || !auth.token) {
      console.error(`Error: ${auth.error}`);
      process.exit(1);
    }
    const result = await listDriveSharedWithMe(auth.token, graphRoot(options));
    if (!result.ok || !result.data) {
      console.error(`Error: ${result.error?.message || 'Request failed'}`);
      process.exit(1);
    }
    const items = result.data.value ?? [];
    if (options.json) {
      console.log(JSON.stringify({ value: items }, null, 2));
      return;
    }
    if (items.length === 0) {
      console.log('No shared items returned.');
      return;
    }
    for (const it of items) {
      const ri = it.remoteItem as { id?: string; name?: string; parentReference?: { driveId?: string } } | undefined;
      const name = (ri?.name ?? it.name ?? '(no name)') as string;
      const id = (ri?.id ?? it.id ?? '') as string;
      const driveId = ri?.parentReference?.driveId ?? '';
      console.log(`${id}\t${name}${driveId ? `\tdrive:${driveId}` : ''}`);
      if (it.webUrl && typeof it.webUrl === 'string') console.log(`  ${it.webUrl}`);
    }
  });

withDriveOptions(filesCommand.command('copy <fileId>'))
  .summary('Copy item to another folder')
  .description('Copy a drive item into another folder (async; use --wait to poll the Graph job to completion)')
  .requiredOption('--parent-id <id>', 'Destination folder item id (parentReference.id)')
  .option('--parent-drive-id <id>', 'Destination drive id when copying across drives (parentReference.driveId)')
  .option('--name <name>', 'Optional new file name')
  .option('--wait', 'Poll async job until completed or failed')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      fileId: string,
      options: FilesDriveCliFlags & {
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
      const auth = await resolveGraphAuth({ token: options.token, identity: options.identity });
      if (!auth.success || !auth.token) {
        console.error(`Error: ${auth.error}`);
        process.exit(1);
      }
      const loc = parseDriveLocation(options);
      const body = {
        parentReference: {
          id: options.parentId.trim(),
          ...(options.parentDriveId?.trim() ? { driveId: options.parentDriveId.trim() } : {})
        },
        ...(options.name?.trim() ? { name: options.name.trim() } : {})
      };
      const started = await startCopyDriveItem(auth.token, fileId, body, loc, graphRoot(options));
      if (!started.ok || !started.data) {
        console.error(`Error: ${started.error?.message || 'Copy failed'}`);
        process.exit(1);
      }
      const monitorUrl = started.data.monitorUrl;
      if (options.wait && monitorUrl) {
        const done = await pollGraphAsyncJob(auth.token, monitorUrl);
        if (!done.ok || !done.data) {
          console.error(`Error: ${done.error?.message || 'Copy job failed'}`);
          process.exit(1);
        }
        if (options.json) {
          console.log(JSON.stringify(done.data, null, 2));
          return;
        }
        console.log('✓ Copy completed');
        console.log(JSON.stringify(done.data, null, 2));
        return;
      }
      if (options.json) {
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

withDriveOptions(filesCommand.command('move <fileId>'))
  .summary('Move item to another folder')
  .description('Move a drive item to another folder (PATCH parentReference)')
  .requiredOption('--parent-id <id>', 'Destination folder item id')
  .option('--parent-drive-id <id>', 'Destination drive id when moving across drives')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      fileId: string,
      options: FilesDriveCliFlags & {
        parentId: string;
        parentDriveId?: string;
        json?: boolean;
        token?: string;
        identity?: string;
      },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: options.token, identity: options.identity });
      if (!auth.success || !auth.token) {
        console.error(`Error: ${auth.error}`);
        process.exit(1);
      }
      const loc = parseDriveLocation(options);
      const parentReference: { id: string; driveId?: string } = { id: options.parentId.trim() };
      if (options.parentDriveId?.trim()) {
        parentReference.driveId = options.parentDriveId.trim();
      }
      const result = await moveDriveItem(auth.token, fileId, parentReference, loc, graphRoot(options));
      if (!result.ok || !result.data) {
        console.error(`Error: ${result.error?.message || 'Move failed'}`);
        process.exit(1);
      }
      if (options.json) {
        console.log(JSON.stringify(result.data, null, 2));
        return;
      }
      console.log(`✓ Moved item ${fileId} → parent ${options.parentId}`);
      if (result.data.name) console.log(`  Name: ${result.data.name}`);
      if (result.data.id) console.log(`  New id: ${result.data.id}`);
    }
  );

withDriveOptions(filesCommand.command('versions <fileId>'))
  .summary('List file versions')
  .description('List versions of a file')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (fileId: string, options: FilesDriveCliFlags & { json?: boolean; token?: string; identity?: string }) => {
      const auth = await resolveGraphAuth({ token: options.token, identity: options.identity });
      if (!auth.success) {
        console.error(`Error: ${auth.error}`);
        process.exit(1);
      }

      const loc = parseDriveLocation(options);
      const result = await listFileVersions(auth.token!, fileId, loc, graphRoot(options));
      if (!result.ok || !result.data) {
        console.error(`Error: ${result.error?.message || 'Failed to list versions'}`);
        process.exit(1);
      }

      if (options.json) {
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
    }
  );

withDriveOptions(filesCommand.command('restore <fileId> <versionId>'))
  .summary('Restore prior version')
  .description('Restore a specific version of a file')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      fileId: string,
      versionId: string,
      options: FilesDriveCliFlags & { json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: options.token, identity: options.identity });
      if (!auth.success) {
        console.error(`Error: ${auth.error}`);
        process.exit(1);
      }

      const loc = parseDriveLocation(options);
      const result = await restoreFileVersion(auth.token!, fileId, versionId, loc, graphRoot(options));
      if (!result.ok) {
        console.error(`Error: ${result.error?.message || 'Restore failed'}`);
        process.exit(1);
      }

      if (options.json) {
        console.log(JSON.stringify({ success: true, fileId, versionId }, null, 2));
        return;
      }

      console.log(`✓ Restored version ${versionId} of file ${fileId}`);
    }
  );

withDriveOptions(filesCommand.command('checkout <fileId>'))
  .summary('Check out for exclusive edit')
  .description('Check out a drive item for exclusive edit (`POST …/checkout`) — pair with `files checkin` when done')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      fileId: string,
      options: FilesDriveCliFlags & { json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: options.token, identity: options.identity });
      if (!auth.success || !auth.token) {
        console.error(`Error: ${auth.error}`);
        process.exit(1);
      }
      const loc = parseDriveLocation(options);
      const result = await checkoutFile(auth.token, fileId, loc, graphRoot(options));
      if (!result.ok) {
        console.error(`Error: ${result.error?.message || 'Checkout failed'}`);
        process.exit(1);
      }
      if (options.json) {
        console.log(JSON.stringify({ success: true, fileId }, null, 2));
        return;
      }
      console.log(`✓ Checked out: ${fileId}`);
    }
  );

withDriveOptions(filesCommand.command('checkin <fileId>'))
  .summary('Check in after checkout')
  .description('Check in a previously checked-out Office document')
  .option('--comment <comment>', 'Optional check-in comment')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      fileId: string,
      options: FilesDriveCliFlags & { comment?: string; json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: options.token, identity: options.identity });
      if (!auth.success) {
        console.error(`Error: ${auth.error}`);
        process.exit(1);
      }

      const loc = parseDriveLocation(options);
      const result = await checkinFile(auth.token!, fileId, options.comment, loc, graphRoot(options));
      if (!result.ok || !result.data) {
        console.error(`Error: ${result.error?.message || 'Check-in failed'}`);
        process.exit(1);
      }

      if (options.json) {
        console.log(JSON.stringify(result.data, null, 2));
        return;
      }

      console.log('✓ File checked in');
      console.log(`  File: ${result.data.item.name}`);
      console.log(`  File ID: ${result.data.item.id}`);
      if (result.data.comment) console.log(`  Comment: ${result.data.comment}`);
    }
  );

withDriveOptions(filesCommand.command('list-item <fileId>'))
  .summary('Get SharePoint listItem fields')
  .description(
    'GET SharePoint listItem for a file (`…/items/{id}/listItem`). Returns library columns/metadata; often 404 on personal OneDrive.'
  )
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (fileId: string, options: FilesDriveCliFlags & { json?: boolean; token?: string; identity?: string }) => {
      const auth = await resolveGraphAuth({ token: options.token, identity: options.identity });
      if (!auth.success || !auth.token) {
        console.error(`Error: ${auth.error}`);
        process.exit(1);
      }
      const loc = parseDriveLocation(options);
      const r = await getDriveItemListItem(auth.token, fileId, loc, graphRoot(options));
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message || 'listItem failed'}`);
        process.exit(1);
      }
      if (options.json) {
        console.log(JSON.stringify(r.data, null, 2));
        return;
      }
      console.log(JSON.stringify(r.data, null, 2));
    }
  );

withDriveOptions(filesCommand.command('follow <fileId>'))
  .summary('Follow item in Office hub')
  .description('Follow a drive item (OneDrive for Business — POST …/follow)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      fileId: string,
      options: FilesDriveCliFlags & { json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: options.token, identity: options.identity });
      if (!auth.success || !auth.token) {
        console.error(`Error: ${auth.error}`);
        process.exit(1);
      }
      const loc = parseDriveLocation(options);
      const r = await followDriveItem(auth.token, fileId, loc, graphRoot(options));
      if (!r.ok) {
        console.error(`Error: ${r.error?.message || 'follow failed'}`);
        process.exit(1);
      }
      if (options.json) {
        console.log(JSON.stringify({ success: true, fileId }, null, 2));
        return;
      }
      console.log(`✓ Following: ${fileId}`);
    }
  );

withDriveOptions(filesCommand.command('unfollow <fileId>'))
  .summary('Stop following item')
  .description('Unfollow a drive item (POST …/unfollow)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      fileId: string,
      options: FilesDriveCliFlags & { json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: options.token, identity: options.identity });
      if (!auth.success || !auth.token) {
        console.error(`Error: ${auth.error}`);
        process.exit(1);
      }
      const loc = parseDriveLocation(options);
      const r = await unfollowDriveItem(auth.token, fileId, loc, graphRoot(options));
      if (!r.ok) {
        console.error(`Error: ${r.error?.message || 'unfollow failed'}`);
        process.exit(1);
      }
      if (options.json) {
        console.log(JSON.stringify({ success: true, fileId }, null, 2));
        return;
      }
      console.log(`✓ Unfollowed: ${fileId}`);
    }
  );

withDriveOptions(filesCommand.command('sensitivity-assign <fileId>'))
  .summary('Apply MIP sensitivity label')
  .description('POST assignSensitivityLabel (Microsoft Information Protection — JSON body per Graph docs)')
  .requiredOption('--json-file <path>', 'Path to JSON action parameters')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      fileId: string,
      options: FilesDriveCliFlags & { jsonFile: string; json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: options.token, identity: options.identity });
      if (!auth.success || !auth.token) {
        console.error(`Error: ${auth.error}`);
        process.exit(1);
      }
      let raw: string;
      try {
        raw = await readFile(options.jsonFile, 'utf-8');
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
      const loc = parseDriveLocation(options);
      const r = await assignDriveItemSensitivityLabel(auth.token, fileId, body, loc, graphRoot(options));
      if (!r.ok) {
        console.error(`Error: ${r.error?.message || 'assignSensitivityLabel failed'}`);
        process.exit(1);
      }
      if (options.json) {
        console.log(JSON.stringify(r.data ?? null, null, 2));
        return;
      }
      console.log('✓ Sensitivity label assignment request completed');
      if (r.data !== undefined) console.log(JSON.stringify(r.data, null, 2));
    }
  );

withDriveOptions(filesCommand.command('sensitivity-extract <fileId>'))
  .summary('Read sensitivity labels')
  .description('POST extractSensitivityLabels — scan content for applicable labels (tenant/licensing dependent)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (fileId: string, options: FilesDriveCliFlags & { json?: boolean; token?: string; identity?: string }) => {
      const auth = await resolveGraphAuth({ token: options.token, identity: options.identity });
      if (!auth.success || !auth.token) {
        console.error(`Error: ${auth.error}`);
        process.exit(1);
      }
      const loc = parseDriveLocation(options);
      const r = await extractDriveItemSensitivityLabels(auth.token, fileId, loc, graphRoot(options));
      if (!r.ok) {
        console.error(`Error: ${r.error?.message || 'extractSensitivityLabels failed'}`);
        process.exit(1);
      }
      if (options.json) {
        console.log(JSON.stringify(r.data ?? null, null, 2));
        return;
      }
      console.log(JSON.stringify(r.data ?? null, null, 2));
    }
  );

withDriveOptions(filesCommand.command('permanent-delete <fileId>'))
  .summary('Permanently delete from recycle bin')
  .description('POST permanentDelete — irreversible; bypasses recycle bin where permitted')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      fileId: string,
      options: FilesDriveCliFlags & { json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: options.token, identity: options.identity });
      if (!auth.success || !auth.token) {
        console.error(`Error: ${auth.error}`);
        process.exit(1);
      }
      const loc = parseDriveLocation(options);
      const r = await permanentDeleteDriveItem(auth.token, fileId, loc, graphRoot(options));
      if (!r.ok) {
        console.error(`Error: ${r.error?.message || 'permanentDelete failed'}`);
        process.exit(1);
      }
      if (options.json) {
        console.log(JSON.stringify({ success: true, fileId }, null, 2));
        return;
      }
      console.log(`✓ Permanently deleted: ${fileId}`);
    }
  );

withDriveOptions(filesCommand.command('retention-label <fileId>'))
  .summary('Assign retention label')
  .description('GET retention label metadata (`…/retentionLabel`)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (fileId: string, options: FilesDriveCliFlags & { json?: boolean; token?: string; identity?: string }) => {
      const auth = await resolveGraphAuth({ token: options.token, identity: options.identity });
      if (!auth.success || !auth.token) {
        console.error(`Error: ${auth.error}`);
        process.exit(1);
      }
      const loc = parseDriveLocation(options);
      const r = await getDriveItemRetentionLabel(auth.token, fileId, loc, graphRoot(options));
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message || 'getRetentionLabel failed'}`);
        process.exit(1);
      }
      console.log(options.json ? JSON.stringify(r.data, null, 2) : JSON.stringify(r.data, null, 2));
    }
  );

withDriveOptions(filesCommand.command('retention-label-remove <fileId>'))
  .summary('Clear retention label')
  .description('DELETE retention label from item (`…/retentionLabel`)')
  .option('--if-match <etag>', 'Optional If-Match (ETag) header')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      fileId: string,
      options: FilesDriveCliFlags & { ifMatch?: string; json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: options.token, identity: options.identity });
      if (!auth.success || !auth.token) {
        console.error(`Error: ${auth.error}`);
        process.exit(1);
      }
      const loc = parseDriveLocation(options);
      const r = await removeDriveItemRetentionLabel(auth.token, fileId, options.ifMatch, loc, graphRoot(options));
      if (!r.ok) {
        console.error(`Error: ${r.error?.message || 'removeRetentionLabel failed'}`);
        process.exit(1);
      }
      if (options.json) {
        console.log(JSON.stringify({ success: true, fileId }, null, 2));
        return;
      }
      console.log(`✓ Removed retention label from: ${fileId}`);
    }
  );

withDriveOptions(filesCommand.command('convert <fileId>'))
  .summary('Download converted format (e.g. PDF)')
  .description('Download a file converted to a specific format (default: pdf)')
  .option('--format <format>', 'Target format (e.g. pdf, html, glb)', 'pdf')
  .option('--out <path>', 'Output path')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      fileId: string,
      options: FilesDriveCliFlags & { format: string; out?: string; json?: boolean; token?: string; identity?: string }
    ) => {
      const auth = await resolveGraphAuth({ token: options.token, identity: options.identity });
      if (!auth.success) {
        console.error(`Error: ${auth.error}`);
        process.exit(1);
      }

      const loc = parseDriveLocation(options);
      const result = await downloadConvertedFile(
        auth.token!,
        fileId,
        options.format,
        options.out,
        loc,
        graphRoot(options)
      );
      if (!result.ok || !result.data) {
        console.error(`Error: ${result.error?.message || 'Conversion download failed'}`);
        process.exit(1);
      }

      if (options.json) {
        console.log(JSON.stringify(result.data, null, 2));
        return;
      }

      console.log(`✓ Converted file downloaded`);
      console.log(`  Saved to: ${result.data.path}`);
    }
  );

withDriveOptions(filesCommand.command('analytics <fileId>'))
  .summary('Item access analytics')
  .description('Get file analytics (access/action counts)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (fileId: string, options: FilesDriveCliFlags & { json?: boolean; token?: string; identity?: string }) => {
      const auth = await resolveGraphAuth({ token: options.token, identity: options.identity });
      if (!auth.success) {
        console.error(`Error: ${auth.error}`);
        process.exit(1);
      }

      const loc = parseDriveLocation(options);
      const result = await getFileAnalytics(auth.token!, fileId, loc, graphRoot(options));
      if (!result.ok || !result.data) {
        console.error(`Error: ${result.error?.message || 'Failed to get file analytics'}`);
        process.exit(1);
      }

      if (options.json) {
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
    }
  );

filesCommand
  .command('recent')
  .summary('Recently used items for user')
  .description(
    'List items recently used by the signed-in user (`GET /me/drive/recent` or `/users/{id}/drive/recent` with `--user`).'
  )
  .option('--user <upn-or-id>', 'Target user delegate (`/users/{id}/drive/recent`)')
  .option('--top <n>', 'Limit results (Graph $top, max 200)')
  .option('--json', 'Output as JSON')
  .option('--beta', 'Use Microsoft Graph beta API host for this call', false)
  .option('--token <token>', 'Use a specific Graph token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (opts: {
      user?: string;
      top?: string;
      json?: boolean;
      beta?: boolean;
      token?: string;
      identity?: string;
    }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Error: ${auth.error}`);
        process.exit(1);
      }
      const top = opts.top ? Number.parseInt(opts.top, 10) : undefined;
      if (opts.top && (!Number.isFinite(top) || (top as number) <= 0)) {
        console.error('Error: --top must be a positive integer');
        process.exit(1);
      }
      const base = opts.user?.trim() ? `/users/${encodeURIComponent(opts.user.trim())}/drive` : '/me/drive';
      const query = top ? `?$top=${Math.min(top, 200)}` : '';
      const r = await callGraphAt<{ value?: DriveItem[] }>(graphRoot(opts), auth.token, `${base}/recent${query}`);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message ?? '/drive/recent failed'}`);
        process.exit(1);
      }
      const items = r.data.value ?? [];
      if (opts.json) {
        console.log(JSON.stringify({ value: items }, null, 2));
        return;
      }
      if (items.length === 0) {
        console.log('No recent items.');
        return;
      }
      renderItems(items);
    }
  );

withDriveOptions(filesCommand.command('activities <fileId>'))
  .summary('Recent activities on an item')
  .description(
    'List recent activities on a drive item (`GET /drives/{id}/items/{id}/activities`). Drive location flags select the item source (default `/me/drive`).'
  )
  .option('--top <n>', 'Limit results (Graph $top, max 200)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      fileId: string,
      options: FilesDriveCliFlags & { top?: string; json?: boolean; token?: string; identity?: string }
    ) => {
      const auth = await resolveGraphAuth({ token: options.token, identity: options.identity });
      if (!auth.success || !auth.token) {
        console.error(`Error: ${auth.error}`);
        process.exit(1);
      }
      const top = options.top ? Number.parseInt(options.top, 10) : undefined;
      if (options.top && (!Number.isFinite(top) || (top as number) <= 0)) {
        console.error('Error: --top must be a positive integer');
        process.exit(1);
      }
      const loc = parseDriveLocation(options);
      const r = await listDriveItemActivities(auth.token, loc, fileId, { top, graphBaseUrl: graphRoot(options) });
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message ?? 'Activities request failed'}`);
        process.exit(1);
      }
      const items = r.data.value ?? [];
      if (options.json) {
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

withDriveOptions(filesCommand.command('preview <fileId>'))
  .summary('Short-lived preview/embed URL')
  .description(
    'Create a short-lived preview session URL for any drive item (`POST /drives/{id}/items/{id}/preview`). Complements `word preview` / `powerpoint preview` for non-Office items.'
  )
  .option('--page <pageOrName>', 'Initial page number or name (per Graph API)')
  .option('--zoom <factor>', 'Zoom factor (per Graph API)')
  .option('--allow-edit', 'Request an embed URL with edit enabled (where supported)')
  .option('--chromeless', 'Request a chromeless embed URL (default true on Graph)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      fileId: string,
      options: FilesDriveCliFlags & {
        page?: string;
        zoom?: string;
        allowEdit?: boolean;
        chromeless?: boolean;
        json?: boolean;
        token?: string;
        identity?: string;
      }
    ) => {
      const auth = await resolveGraphAuth({ token: options.token, identity: options.identity });
      if (!auth.success || !auth.token) {
        console.error(`Error: ${auth.error}`);
        process.exit(1);
      }
      const loc = parseDriveLocation(options);
      let page: number | string | undefined;
      if (options.page !== undefined) {
        const asNum = Number(options.page);
        page = Number.isFinite(asNum) && options.page.trim() !== '' ? asNum : options.page;
      }
      let zoom: number | undefined;
      if (options.zoom !== undefined) {
        zoom = Number(options.zoom);
        if (!Number.isFinite(zoom)) {
          console.error('Error: --zoom must be numeric');
          process.exit(1);
        }
      }
      const r = await createDriveItemPreview(
        auth.token,
        loc,
        fileId,
        {
          page,
          zoom,
          allowEdit: options.allowEdit,
          chromeless: options.chromeless
        },
        graphRoot(options)
      );
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message ?? 'Preview request failed'}`);
        process.exit(1);
      }
      if (options.json) {
        console.log(JSON.stringify(r.data, null, 2));
        return;
      }
      if (r.data.getUrl) console.log(`getUrl:  ${r.data.getUrl}`);
      if (r.data.postUrl) console.log(`postUrl: ${r.data.postUrl}`);
      if (r.data.postParameters) console.log(`postParameters: ${r.data.postParameters}`);
    }
  );

function renderActivityLine(a: ItemActivity): void {
  const when = a.times?.recordedDateTime ?? a.times?.observedDateTime ?? '';
  const who = a.actor?.user?.displayName ?? a.actor?.user?.email ?? a.actor?.application?.displayName ?? '';
  const action = a.action ? Object.keys(a.action).join(',') : '';
  console.log(`${when}\t${action}\t${who}${a.id ? `\tid:${a.id}` : ''}`);
}
