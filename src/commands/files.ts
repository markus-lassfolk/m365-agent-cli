import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import {
  defaultDownloadPath,
  deleteFile,
  downloadFile,
  getFileMetadata,
  listFiles,
  searchFiles,
  shareFile,
  uploadFile,
  createLargeUploadSession,
  type DriveItem,
  type DriveItemReference
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

export const filesCommand = new Command('files')
  .description('Manage OneDrive files via Microsoft Graph');

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
  .description('Create a large-upload session for files up to 4GB')
  .option('--folder <id>', 'Target folder item ID')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .action(async (path: string, options: { folder?: string; json?: boolean; token?: string }) => {
    const auth = await resolveGraphAuth({ token: options.token });
    if (!auth.success) {
      console.error(`Error: ${auth.error}`);
      process.exit(1);
    }

    const result = await createLargeUploadSession(auth.token!, path, parseFolderRef(options.folder));
    if (!result.ok || !result.data) {
      console.error(`Error: ${result.error?.message || 'Failed to create upload session'}`);
      process.exit(1);
    }

    if (options.json) {
      console.log(JSON.stringify(result.data, null, 2));
      return;
    }

    console.log('✓ Large upload session created');
    console.log(`  Upload URL: ${result.data.uploadUrl}`);
    if (result.data.expirationDateTime) console.log(`  Expires: ${result.data.expirationDateTime}`);
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

    const meta = await getFileMetadata(auth.token!, fileId);
    const defaultOut = meta.ok && meta.data ? defaultDownloadPath(meta.data.name || fileId) : undefined;
    const result = await downloadFile(auth.token!, fileId, options.out || defaultOut, meta.ok ? meta.data : undefined);
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
  .description('Create a OneDrive sharing link')
  .option('--type <type>', 'Link type: view or edit', 'view')
  .option('--scope <scope>', 'Link scope: org or anonymous', 'org')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .action(
    async (
      fileId: string,
      options: { type?: 'view' | 'edit'; scope?: 'org' | 'anonymous'; json?: boolean; token?: string }
    ) => {
      const auth = await resolveGraphAuth({ token: options.token });
      if (!auth.success) {
        console.error(`Error: ${auth.error}`);
        process.exit(1);
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
