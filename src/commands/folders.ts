import { Command } from 'commander';
import { resolveAuth } from '../lib/auth.js';
import {
  createMailFolder as ewsCreateMailFolder,
  deleteMailFolder as ewsDeleteMailFolder,
  updateMailFolder as ewsUpdateMailFolder,
  getMailFolders
} from '../lib/ews-client.js';
import { getExchangeBackend } from '../lib/exchange-backend.js';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import {
  createMailFolder as graphCreateMailFolder,
  deleteMailFolder as graphDeleteMailFolder,
  updateMailFolder as graphUpdateMailFolder,
  listAllMailFoldersRecursive,
  type OutlookMailFolder
} from '../lib/outlook-graph-client.js';
import { checkReadOnly } from '../lib/utils.js';

const SYSTEM_FOLDER_NAMES = ['inbox', 'drafts', 'sent items', 'deleted items', 'junk email', 'archive', 'outbox'];

function isSystemFolderName(displayName: string): boolean {
  return SYSTEM_FOLDER_NAMES.includes(displayName.toLowerCase());
}

export const foldersCommand = new Command('folders')
  .description('Manage mail folders (EWS or Microsoft Graph per M365_EXCHANGE_BACKEND)')
  .option('--create <name>', 'Create a new folder')
  .option('--rename <name>', 'Rename a folder (use with --to)')
  .option('--delete <name>', 'Delete a folder')
  .option('--to <newname>', 'New name for rename operation')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Access token (EWS or Graph depending on backend)')
  .option('--identity <name>', 'Authentication identity (default: default)')
  .option('--mailbox <email>', 'Delegated or shared mailbox')
  .action(
    async (
      options: {
        create?: string;
        rename?: string;
        delete?: string;
        to?: string;
        json?: boolean;
        token?: string;
        identity?: string;
        mailbox?: string;
      },
      cmd: any
    ) => {
      const backend = getExchangeBackend();
      const user = options.mailbox?.trim() || undefined;

      async function runGraph(token: string): Promise<void> {
        const foldersResult = await listAllMailFoldersRecursive(token, user);
        if (!foldersResult.ok || !foldersResult.data) {
          console.error(`Error: ${foldersResult.error?.message || 'Failed to fetch folders'}`);
          process.exit(1);
        }

        const folders = foldersResult.data;
        const findFolder = (name: string) => folders.find((f) => f.displayName.toLowerCase() === name.toLowerCase());

        if (options.create) {
          checkReadOnly(cmd);
          if (findFolder(options.create)) {
            console.error(`Folder "${options.create}" already exists.`);
            process.exit(1);
          }
          const result = await graphCreateMailFolder(token, options.create, undefined, user);
          if (!result.ok || !result.data) {
            console.error(`Error: ${result.error?.message || 'Failed to create folder'}`);
            process.exit(1);
          }
          if (options.json) {
            console.log(JSON.stringify({ success: true, folder: result.data, backend: 'graph' }, null, 2));
          } else {
            console.log(`\u2713 Created folder: ${result.data.displayName}`);
          }
          return;
        }

        if (options.rename) {
          checkReadOnly(cmd);
          if (!options.to) {
            console.error('Please specify new name with --to');
            console.error('Example: m365-agent-cli folders --rename "Old Name" --to "New Name"');
            process.exit(1);
          }
          const folder = findFolder(options.rename);
          if (!folder) {
            console.error(`Folder "${options.rename}" not found.`);
            process.exit(1);
          }
          const result = await graphUpdateMailFolder(token, folder.id, options.to, user);
          if (!result.ok || !result.data) {
            console.error(`Error: ${result.error?.message || 'Failed to rename folder'}`);
            process.exit(1);
          }
          if (options.json) {
            console.log(JSON.stringify({ success: true, folder: result.data, backend: 'graph' }, null, 2));
          } else {
            console.log(`\u2713 Renamed "${options.rename}" to "${result.data.displayName}"`);
          }
          return;
        }

        if (options.delete) {
          checkReadOnly(cmd);
          const folder = findFolder(options.delete);
          if (!folder) {
            console.error(`Folder "${options.delete}" not found.`);
            process.exit(1);
          }
          if (isSystemFolderName(folder.displayName)) {
            console.error(`Cannot delete system folder "${folder.displayName}".`);
            process.exit(1);
          }
          const result = await graphDeleteMailFolder(token, folder.id, user);
          if (!result.ok) {
            console.error(`Error: ${result.error?.message || 'Failed to delete folder'}`);
            process.exit(1);
          }
          if (options.json) {
            console.log(JSON.stringify({ success: true, deleted: options.delete, backend: 'graph' }, null, 2));
          } else {
            console.log(`\u2713 Deleted folder: ${options.delete}`);
          }
          return;
        }

        if (options.json) {
          console.log(
            JSON.stringify(
              {
                backend: 'graph',
                folders: folders.map((f: OutlookMailFolder) => ({
                  id: f.id,
                  name: f.displayName,
                  unread: f.unreadItemCount,
                  total: f.totalItemCount,
                  childFolders: f.childFolderCount
                }))
              },
              null,
              2
            )
          );
          return;
        }

        console.log(`\n\ud83d\udcc1 Mail Folders (Graph)${user ? ` — ${user}` : ''}:\n`);
        console.log('\u2500'.repeat(50));
        for (const folder of folders) {
          const unreadBadge = (folder.unreadItemCount ?? 0) > 0 ? ` (${folder.unreadItemCount} unread)` : '';
          console.log(`  ${folder.displayName}${unreadBadge}`);
          console.log(`    ${folder.totalItemCount ?? 0} items`);
        }
        console.log(`\n${'\u2500'.repeat(50)}`);
        console.log('\nCommands:');
        console.log('  m365-agent-cli folders --create "Folder Name"');
        console.log('  m365-agent-cli folders --rename "Old" --to "New"');
        console.log('  m365-agent-cli folders --delete "Folder Name"');
        console.log('');
      }

      async function runEws(): Promise<void> {
        const authResult = await resolveAuth({
          token: options.token,
          identity: options.identity
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

        const foldersResult = await getMailFolders(authResult.token!, undefined, options.mailbox);
        if (!foldersResult.ok || !foldersResult.data) {
          console.error(`Error: ${foldersResult.error?.message || 'Failed to fetch folders'}`);
          process.exit(1);
        }

        const folders = foldersResult.data.value;

        const findFolder = (name: string) => {
          return folders.find((f) => f.DisplayName.toLowerCase() === name.toLowerCase());
        };

        if (options.create) {
          checkReadOnly(cmd);
          const existing = findFolder(options.create);
          if (existing) {
            console.error(`Folder "${options.create}" already exists.`);
            process.exit(1);
          }

          const result = await ewsCreateMailFolder(authResult.token!, options.create, undefined, options.mailbox);
          if (!result.ok || !result.data) {
            console.error(`Error: ${result.error?.message || 'Failed to create folder'}`);
            process.exit(1);
          }

          if (options.json) {
            console.log(JSON.stringify({ success: true, folder: result.data, backend: 'ews' }, null, 2));
          } else {
            console.log(`\u2713 Created folder: ${result.data.DisplayName}`);
          }
          return;
        }

        if (options.rename) {
          checkReadOnly(cmd);
          if (!options.to) {
            console.error('Please specify new name with --to');
            console.error('Example: m365-agent-cli folders --rename "Old Name" --to "New Name"');
            process.exit(1);
          }

          const folder = findFolder(options.rename);
          if (!folder) {
            console.error(`Folder "${options.rename}" not found.`);
            process.exit(1);
          }

          const result = await ewsUpdateMailFolder(authResult.token!, folder.Id, options.to, options.mailbox);
          if (!result.ok || !result.data) {
            console.error(`Error: ${result.error?.message || 'Failed to rename folder'}`);
            process.exit(1);
          }

          if (options.json) {
            console.log(JSON.stringify({ success: true, folder: result.data, backend: 'ews' }, null, 2));
          } else {
            console.log(`\u2713 Renamed "${options.rename}" to "${result.data.DisplayName}"`);
          }
          return;
        }

        if (options.delete) {
          checkReadOnly(cmd);
          const folder = findFolder(options.delete);
          if (!folder) {
            console.error(`Folder "${options.delete}" not found.`);
            process.exit(1);
          }

          if (isSystemFolderName(folder.DisplayName)) {
            console.error(`Cannot delete system folder "${folder.DisplayName}".`);
            process.exit(1);
          }

          const result = await ewsDeleteMailFolder(authResult.token!, folder.Id, options.mailbox);
          if (!result.ok) {
            console.error(`Error: ${result.error?.message || 'Failed to delete folder'}`);
            process.exit(1);
          }

          if (options.json) {
            console.log(JSON.stringify({ success: true, deleted: options.delete, backend: 'ews' }, null, 2));
          } else {
            console.log(`\u2713 Deleted folder: ${options.delete}`);
          }
          return;
        }

        if (options.json) {
          console.log(
            JSON.stringify(
              {
                backend: 'ews',
                folders: folders.map((f) => ({
                  id: f.Id,
                  name: f.DisplayName,
                  unread: f.UnreadItemCount,
                  total: f.TotalItemCount,
                  childFolders: f.ChildFolderCount
                }))
              },
              null,
              2
            )
          );
          return;
        }

        console.log(`\n\ud83d\udcc1 Mail Folders${options.mailbox ? ` — ${options.mailbox}` : ''}:\n`);
        console.log('\u2500'.repeat(50));

        for (const folder of folders) {
          const unreadBadge = folder.UnreadItemCount > 0 ? ` (${folder.UnreadItemCount} unread)` : '';
          console.log(`  ${folder.DisplayName}${unreadBadge}`);
          console.log(`    ${folder.TotalItemCount} items`);
        }

        console.log(`\n${'\u2500'.repeat(50)}`);
        console.log('\nCommands:');
        console.log('  m365-agent-cli folders --create "Folder Name"');
        console.log('  m365-agent-cli folders --rename "Old" --to "New"');
        console.log('  m365-agent-cli folders --delete "Folder Name"');
        console.log('');
      }

      if (backend === 'ews') {
        await runEws();
        return;
      }

      const graphAuth = await resolveGraphAuth({
        token: options.token,
        identity: options.identity
      });

      if (backend === 'graph') {
        if (!graphAuth.success || !graphAuth.token) {
          if (options.json) {
            console.log(JSON.stringify({ error: graphAuth.error || 'Graph auth failed' }, null, 2));
          } else {
            console.error(`Error: ${graphAuth.error || 'Graph authentication failed'}`);
            console.error('\nSet EWS_CLIENT_ID and M365_REFRESH_TOKEN for Graph, or run `m365-agent-cli login`.');
          }
          process.exit(1);
        }
        await runGraph(graphAuth.token);
        return;
      }

      // auto
      if (graphAuth.success && graphAuth.token) {
        await runGraph(graphAuth.token);
      } else {
        await runEws();
      }
    }
  );
