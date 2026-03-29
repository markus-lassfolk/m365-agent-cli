import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import {
  checkinFile,
  createOfficeCollaborationLink,
  type DriveItem,
  type DriveItemReference,
  defaultDownloadPath,
  deleteFile,
  downloadFile,
  getFileMetadata,
  listFiles,
  searchFiles,
  shareFile,
  uploadFile,
  uploadLargeFile,
  listFileVersions,
  restoreFileVersion,
  type DriveItemVersion
} from '../lib/graph-client.js';

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

export const filesCommand = new Command('files').description('Manage OneDrive files via Microsoft Graph');

filesCommand
  .command('list')
  .description('List files in OneDrive root or a folder')
  .option('--folder <id>', 'Folder item ID')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .action(async (options: { folder?: string; json?: boolean; token?: string }) => {
    const auth = await resolveGraphAuth({ token: options.token });
    if (!auth.success) {
      console.error(`Error: ${auth.error}`);
      process.exit(1);
    }

    const result = await listFiles(auth.token!, parseFolderRef(options.folder));
    if (!result.ok || !result.data) {
      console.error(`Error: ${result.error?.message || 'Failed to list files'}`);
      process.exit(1);
    }

    if (options.json) {
      console.log(JSON.stringify({ items: result.data }, null, 2));
      return;
    }

    renderItems(result.data);
  });

filesCommand
  .command('search <query>')
  .description('Search OneDrive files')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .action(async (query: string, options: { json?: boolean; token?: string }) => {
    const auth = await resolveGraphAuth({ token: options.token });
    if (!auth.success) {
      console.error(`Error: ${auth.error}`);
      process.exit(1);
    }

    const result = await searchFiles(auth.token!, query);
    if (!result.ok || !result.data) {
      console.error(`Error: ${result.error?.message || 'Failed to search files'}`);
      process.exit(1);
    }

    if (options.json) {
      console.log(JSON.stringify({ items: result.data }, null, 2));
      return;
    }

    renderItems(result.data);
  });

filesCommand
  .command('meta <fileId>')
  .description('Get file metadata')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .action(async (fileId: string, options: { json?: boolean; token?: string }) => {
    const auth = await resolveGraphAuth({ token: options.token });
    if (!auth.success) {
      console.error(`Error: ${auth.error}`);
      process.exit(1);
    }

    const result = await getFileMetadata(auth.token!, fileId);
    if (!result.ok || !result.data) {
      console.error(`Error: ${result.error?.message || 'Failed to fetch metadata'}`);
      process.exit(1);
    }

    if (options.json) {
      console.log(JSON.stringify(result.data, null, 2));
      return;
    }

    renderItems([result.data]);
  });

filesCommand
  .command('upload <path>')
  .description('Upload a file up to 250MB')
  .option('--folder <id>', 'Target folder item ID')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .action(async (path: string, options: { folder?: string; json?: boolean; token?: string }) => {
    const auth = await resolveGraphAuth({ token: options.token });
    if (!auth.success) {
      console.error(`Error: ${auth.error}`);
      process.exit(1);
    }

    const result = await uploadFile(auth.token!, path, parseFolderRef(options.folder));
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
  });

filesCommand
  .command('upload-large <path>')
  .description('Upload a file up to 4GB using a chunked upload session')
  .option('--folder <id>', 'Target folder item ID')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .action(async (path: string, options: { folder?: string; json?: boolean; token?: string }) => {
    const auth = await resolveGraphAuth({ token: options.token });
    if (!auth.success) {
      console.error(`Error: ${auth.error}`);
      process.exit(1);
    }

    const result = await uploadLargeFile(auth.token!, path, parseFolderRef(options.folder));
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
  });

filesCommand
  .command('download <fileId>')
  .description('Download a file by ID')
  .option('--out <path>', 'Output path (defaults to ~/Downloads/<name>)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .action(async (fileId: string, options: { out?: string; json?: boolean; token?: string }) => {
    const auth = await resolveGraphAuth({ token: options.token });
    if (!auth.success) {
      console.error(`Error: ${auth.error}`);
      process.exit(1);
    }

    const meta = options.out ? undefined : await getFileMetadata(auth.token!, fileId);
    const defaultOut = meta?.ok && meta.data ? defaultDownloadPath(meta.data.name || fileId) : undefined;
    const result = await downloadFile(auth.token!, fileId, options.out || defaultOut, meta?.ok ? meta.data : undefined);
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
  });

filesCommand
  .command('delete <fileId>')
  .description('Delete a file by ID')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .action(async (fileId: string, options: { json?: boolean; token?: string }) => {
    const auth = await resolveGraphAuth({ token: options.token });
    if (!auth.success) {
      console.error(`Error: ${auth.error}`);
      process.exit(1);
    }

    const result = await deleteFile(auth.token!, fileId);
    if (!result.ok) {
      console.error(`Error: ${result.error?.message || 'Delete failed'}`);
      process.exit(1);
    }

    if (options.json) {
      console.log(JSON.stringify({ success: true, id: fileId }, null, 2));
      return;
    }

    console.log(`✓ Deleted file: ${fileId}`);
  });

filesCommand
  .command('share <fileId>')
  .description('Create a OneDrive sharing link or Office Online collaboration handoff')
  .option('--type <type>', 'Link type: view or edit', 'view')
  .option('--scope <scope>', 'Link scope: org or anonymous', 'org')
  .option('--collab', 'Create an Office Online collaboration handoff (edit/org + webUrl)')
  .option('--lock', 'Checkout the file before creating a collaboration link (use with --collab)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .action(
    async (
      fileId: string,
      options: {
        type?: 'view' | 'edit';
        scope?: 'org' | 'anonymous';
        collab?: boolean;
        lock?: boolean;
        json?: boolean;
        token?: string;
      }
    ) => {
      const auth = await resolveGraphAuth({ token: options.token });
      if (!auth.success) {
        console.error(`Error: ${auth.error}`);
        process.exit(1);
      }

      if (options.lock && !options.collab) {
        console.error('Error: --lock is only supported together with --collab.');
        process.exit(1);
      }

      if (options.collab) {
        const result = await createOfficeCollaborationLink(auth.token!, fileId, { lock: options.lock });
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
      const result = await shareFile(auth.token!, fileId, type, scope);
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


filesCommand
  .command('versions <fileId>')
  .description('List versions of a file')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .action(async (fileId: string, options: { json?: boolean; token?: string }) => {
    const auth = await resolveGraphAuth({ token: options.token });
    if (!auth.success) {
      console.error(`Error: ${auth.error}`);
      process.exit(1);
    }

    const result = await listFileVersions(auth.token!, fileId);
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
  });

filesCommand
  .command('restore <fileId> <versionId>')
  .description('Restore a specific version of a file')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .action(async (fileId: string, versionId: string, options: { json?: boolean; token?: string }) => {
    const auth = await resolveGraphAuth({ token: options.token });
    if (!auth.success) {
      console.error(`Error: ${auth.error}`);
      process.exit(1);
    }

    const result = await restoreFileVersion(auth.token!, fileId, versionId);
    if (!result.ok) {
      console.error(`Error: ${result.error?.message || 'Restore failed'}`);
      process.exit(1);
    }

    if (options.json) {
      console.log(JSON.stringify({ success: true, fileId, versionId }, null, 2));
      return;
    }

    console.log(`✓ Restored version ${versionId} of file ${fileId}`);
  });

filesCommand
  .command('checkin <fileId>')
  .description('Check in a previously checked-out Office document')
  .option('--comment <comment>', 'Optional check-in comment')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .action(async (fileId: string, options: { comment?: string; json?: boolean; token?: string }) => {
    const auth = await resolveGraphAuth({ token: options.token });
    if (!auth.success) {
      console.error(`Error: ${auth.error}`);
      process.exit(1);
    }

    const result = await checkinFile(auth.token!, fileId, options.comment);
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
  });
