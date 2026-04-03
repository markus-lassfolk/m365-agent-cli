import { readFile } from 'node:fs/promises';
import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import {
  createContact,
  deleteContact,
  getContact,
  listContactFolders,
  listContacts,
  listContactsInFolder,
  updateContact
} from '../lib/outlook-graph-client.js';
import { checkReadOnly } from '../lib/utils.js';

export const contactsCommand = new Command('contacts').description(
  'Outlook contacts via Microsoft Graph (delegated Contacts.ReadWrite; run login after Entra permission changes)'
);

contactsCommand
  .command('folders')
  .description('List contact folders (Graph GET /contactFolders)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
  .action(async (opts: { json?: boolean; token?: string; identity?: string; user?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await listContactFolders(auth.token!, opts.user);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else {
      for (const f of r.data) {
        console.log(`${f.displayName ?? '(folder)'}\t${f.id}`);
      }
    }
  });

contactsCommand
  .command('list')
  .description('List contacts (default folder or --folder)')
  .option('-f, --folder <folderId>', 'Contact folder id (omit for default contacts)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
  .action(async (opts: { folder?: string; json?: boolean; token?: string; identity?: string; user?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = opts.folder
      ? await listContactsInFolder(auth.token!, opts.folder, opts.user)
      : await listContacts(auth.token!, opts.user);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else {
      for (const c of r.data) {
        const em = c.emailAddresses?.[0]?.address ?? '';
        console.log(`${c.displayName ?? '(no name)'}\t${em}\t${c.id}`);
      }
    }
  });

contactsCommand
  .command('show')
  .description('Get one contact by id')
  .argument('<contactId>', 'Contact id')
  .option('--select <fields>', 'OData $select')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
  .action(
    async (
      contactId: string,
      opts: { select?: string; json?: boolean; token?: string; identity?: string; user?: string }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getContact(auth.token!, contactId, opts.user, opts.select);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else console.log(JSON.stringify(r.data, null, 2));
    }
  );

contactsCommand
  .command('create')
  .description('Create a contact (JSON body per Graph contact resource)')
  .requiredOption('--json-file <path>', 'Path to JSON file')
  .option('--json', 'Echo created contact as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
  .action(
    async (opts: { jsonFile: string; json?: boolean; token?: string; identity?: string; user?: string }, cmd: any) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const raw = await readFile(opts.jsonFile, 'utf-8');
      const body = JSON.parse(raw) as Record<string, unknown>;
      const r = await createContact(auth.token!, body, opts.user);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else console.log(`Created contact: ${r.data.displayName ?? r.data.id} (${r.data.id})`);
    }
  );

contactsCommand
  .command('update')
  .description('PATCH a contact (merge JSON from file)')
  .argument('<contactId>', 'Contact id')
  .requiredOption('--json-file <path>', 'Path to JSON patch body')
  .option('--json', 'Echo updated contact as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
  .action(
    async (
      contactId: string,
      opts: { jsonFile: string; json?: boolean; token?: string; identity?: string; user?: string },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const raw = await readFile(opts.jsonFile, 'utf-8');
      const patch = JSON.parse(raw) as Record<string, unknown>;
      const r = await updateContact(auth.token!, contactId, patch, opts.user);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else console.log(`Updated contact: ${r.data.id}`);
    }
  );

contactsCommand
  .command('delete')
  .description('Delete a contact')
  .argument('<contactId>', 'Contact id')
  .option('--confirm', 'Confirm delete')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
  .action(
    async (
      contactId: string,
      opts: { confirm?: boolean; token?: string; identity?: string; user?: string },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      if (!opts.confirm) {
        console.error('Refusing to delete without --confirm');
        process.exit(1);
      }
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await deleteContact(auth.token!, contactId, opts.user);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log('Deleted contact.');
    }
  );
