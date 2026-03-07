import { Command } from 'commander';
import { resolveAuth } from '../lib/auth.js';
import { getMailFolders, createMailFolder, updateMailFolder, deleteMailFolder } from '../lib/ews-client.js';

export const foldersCommand = new Command('folders')
  .description('Manage mail folders')
  .option('--create <name>', 'Create a new folder')
  .option('--rename <name>', 'Rename a folder (use with --to)')
  .option('--delete <name>', 'Delete a folder')
  .option('--to <newname>', 'New name for rename operation')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .action(async (options: {
    create?: string;
    rename?: string;
    delete?: string;
    to?: string;
    json?: boolean;
    token?: string;
  }) => {
    const authResult = await resolveAuth({
      token: options.token,
    });

    if (!authResult.success) {
      if (options.json) {
        console.log(JSON.stringify({ error: authResult.error }, null, 2));
      } else {
        console.error(`Error: ${authResult.error}`);
        console.error('\nCheck your .env file for EWS_CLIENT_ID and EWS_REFRESH_TOKEN.');
      }
      process.exit(1);
    }

    // Get all folders first (needed for most operations)
    const foldersResult = await getMailFolders(authResult.token!);
    if (!foldersResult.ok || !foldersResult.data) {
      console.error(`Error: ${foldersResult.error?.message || 'Failed to fetch folders'}`);
      process.exit(1);
    }

    const folders = foldersResult.data.value;

    // Helper to find folder by name (case-insensitive)
    const findFolder = (name: string) => {
      return folders.find(f => f.DisplayName.toLowerCase() === name.toLowerCase());
    };

    // Handle create
    if (options.create) {
      const existing = findFolder(options.create);
      if (existing) {
        console.error(`Folder "${options.create}" already exists.`);
        process.exit(1);
      }

      const result = await createMailFolder(authResult.token!, options.create);
      if (!result.ok || !result.data) {
        console.error(`Error: ${result.error?.message || 'Failed to create folder'}`);
        process.exit(1);
      }

      if (options.json) {
        console.log(JSON.stringify({ success: true, folder: result.data }, null, 2));
      } else {
        console.log(`\u2713 Created folder: ${result.data.DisplayName}`);
      }
      return;
    }

    // Handle rename
    if (options.rename) {
      if (!options.to) {
        console.error('Please specify new name with --to');
        console.error('Example: clippy folders --rename "Old Name" --to "New Name"');
        process.exit(1);
      }

      const folder = findFolder(options.rename);
      if (!folder) {
        console.error(`Folder "${options.rename}" not found.`);
        process.exit(1);
      }

      const result = await updateMailFolder(authResult.token!, folder.Id, options.to);
      if (!result.ok || !result.data) {
        console.error(`Error: ${result.error?.message || 'Failed to rename folder'}`);
        process.exit(1);
      }

      if (options.json) {
        console.log(JSON.stringify({ success: true, folder: result.data }, null, 2));
      } else {
        console.log(`\u2713 Renamed "${options.rename}" to "${result.data.DisplayName}"`);
      }
      return;
    }

    // Handle delete
    if (options.delete) {
      const folder = findFolder(options.delete);
      if (!folder) {
        console.error(`Folder "${options.delete}" not found.`);
        process.exit(1);
      }

      // Prevent deleting system folders
      const systemFolders = ['inbox', 'drafts', 'sent items', 'deleted items', 'junk email', 'archive', 'outbox'];
      if (systemFolders.includes(folder.DisplayName.toLowerCase())) {
        console.error(`Cannot delete system folder "${folder.DisplayName}".`);
        process.exit(1);
      }

      const result = await deleteMailFolder(authResult.token!, folder.Id);
      if (!result.ok) {
        console.error(`Error: ${result.error?.message || 'Failed to delete folder'}`);
        process.exit(1);
      }

      if (options.json) {
        console.log(JSON.stringify({ success: true, deleted: options.delete }, null, 2));
      } else {
        console.log(`\u2713 Deleted folder: ${options.delete}`);
      }
      return;
    }

    // List folders (default action)
    if (options.json) {
      console.log(JSON.stringify({
        folders: folders.map(f => ({
          id: f.Id,
          name: f.DisplayName,
          unread: f.UnreadItemCount,
          total: f.TotalItemCount,
          childFolders: f.ChildFolderCount,
        })),
      }, null, 2));
      return;
    }

    console.log('\n\ud83d\udcc1 Mail Folders:\n');
    console.log('\u2500'.repeat(50));

    for (const folder of folders) {
      const unreadBadge = folder.UnreadItemCount > 0 ? ` (${folder.UnreadItemCount} unread)` : '';
      console.log(`  ${folder.DisplayName}${unreadBadge}`);
      console.log(`    ${folder.TotalItemCount} items`);
    }

    console.log('\n' + '\u2500'.repeat(50));
    console.log('\nCommands:');
    console.log('  clippy folders --create "Folder Name"');
    console.log('  clippy folders --rename "Old" --to "New"');
    console.log('  clippy folders --delete "Folder Name"');
    console.log('');
  });
