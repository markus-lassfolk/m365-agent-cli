import { readFile, writeFile } from 'node:fs/promises';
import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import {
  deleteContactMergeSuggestions,
  getContactMergeSuggestions,
  patchContactMergeSuggestions
} from '../lib/graph-contact-merge-suggestions-client.js';
import {
  applyDeltaPageToState,
  assertDeltaScopeMatchesState,
  readDeltaStateFile,
  resolveDeltaContinuationUrl,
  writeDeltaStateFile
} from '../lib/graph-delta-state-file.js';
import { toJsonError } from '../lib/json-error.js';
import {
  addFileAttachmentToContact,
  addReferenceAttachmentToContact,
  type ContactExtensionLocation,
  type ContactListQuery,
  contactsDeltaPage,
  createContact,
  createContactFolder,
  deleteContact,
  deleteContactAttachment,
  deleteContactFolder,
  deleteContactOpenExtension,
  deleteContactPhoto,
  downloadContactAttachmentBytes,
  getContact,
  getContactAttachment,
  getContactFolder,
  getContactOpenExtension,
  getContactPhotoBytes,
  listChildContactFolders,
  listContactAttachments,
  listContactFolders,
  listContactOpenExtensions,
  listContacts,
  listContactsInFolder,
  listContactsRawPage,
  searchContacts,
  setContactOpenExtension,
  setContactPhoto,
  updateContact,
  updateContactFolder,
  updateContactOpenExtension
} from '../lib/outlook-graph-client.js';
import { checkReadOnly } from '../lib/utils.js';

const contactsDesc =
  'Outlook contacts via Microsoft Graph (Contacts.ReadWrite; shared mailboxes: Contacts.Read.Shared / Contacts.ReadWrite.Shared — see docs/GRAPH_SCOPES.md)';

function failContacts(
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

function resolveContactExtensionLocation(opts: {
  folder?: string;
  childFolder?: string;
}): ContactExtensionLocation | undefined | 'invalid' {
  const folder = opts.folder?.trim();
  const child = opts.childFolder?.trim();
  if (child && !folder) return 'invalid';
  if (!folder) return undefined;
  return child ? { folderId: folder, childFolderId: child } : { folderId: folder };
}

export const contactsCommand = new Command('contacts').description(contactsDesc);

// ─── folders (list — backward compatible) ───────────────────────────────────

async function listContactFoldersAction(opts: { json?: boolean; token?: string; identity?: string; user?: string }) {
  const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
  if (!auth.success || !auth.token) {
    failContacts(opts.json, 'Auth error', auth.error);
  }
  const token = auth.token;
  const r = await listContactFolders(token, opts.user);
  if (!r.ok || !r.data) {
    failContacts(opts.json, 'Error', r.error);
  }
  if (opts.json) console.log(JSON.stringify(r.data, null, 2));
  else {
    for (const f of r.data) {
      console.log(`${f.displayName ?? '(folder)'}\t${f.id}`);
    }
  }
}

contactsCommand
  .command('folders')
  .description('List contact folders (Graph GET /contactFolders); see also `contacts folder list`')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (delegated / shared mailbox)')
  .action(listContactFoldersAction);

// ─── folder (CRUD + children) ───────────────────────────────────────────────

const folderCmd = new Command('folder').description('Contact folder create, read, update, delete, list');

folderCmd
  .command('list')
  .description('List contact folders')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user')
  .action(listContactFoldersAction);

folderCmd
  .command('get')
  .description('Get one contact folder by id')
  .argument('<folderId>', 'Folder id')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user')
  .action(async (folderId: string, opts: { json?: boolean; token?: string; identity?: string; user?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      failContacts(opts.json, 'Auth error', auth.error);
    }
    const token = auth.token;
    const r = await getContactFolder(token, folderId, opts.user);
    if (!r.ok || !r.data) {
      failContacts(opts.json, 'Error', r.error);
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else console.log(`${r.data.displayName ?? '(folder)'}\t${r.data.id}`);
  });

folderCmd
  .command('create')
  .description('Create a contact folder (optional parent)')
  .requiredOption('--name <displayName>', 'Display name')
  .option('--parent <folderId>', 'Parent folder id (omit for top level)')
  .option('--json', 'Echo created folder as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user')
  .action(
    async (
      opts: {
        name: string;
        parent?: string;
        json?: boolean;
        token?: string;
        identity?: string;
        user?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        failContacts(opts.json, 'Auth error', auth.error);
      }
      const token = auth.token;
      const r = await createContactFolder(token, opts.name, opts.user, opts.parent);
      if (!r.ok || !r.data) {
        failContacts(opts.json, 'Error', r.error);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else console.log(`Created folder: ${r.data.displayName ?? r.data.id} (${r.data.id})`);
    }
  );

folderCmd
  .command('update')
  .description('PATCH a contact folder (JSON merge from file)')
  .argument('<folderId>', 'Folder id')
  .requiredOption('--json-file <path>', 'Path to JSON patch body')
  .option('--json', 'Echo updated folder as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user')
  .action(
    async (
      folderId: string,
      opts: { jsonFile: string; json?: boolean; token?: string; identity?: string; user?: string },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        failContacts(opts.json, 'Auth error', auth.error);
      }
      const token = auth.token;
      const raw = await readFile(opts.jsonFile, 'utf-8');
      const patch = JSON.parse(raw) as Record<string, unknown>;
      const r = await updateContactFolder(token, folderId, patch, opts.user);
      if (!r.ok || !r.data) {
        failContacts(opts.json, 'Error', r.error);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else console.log(`Updated folder: ${r.data.id}`);
    }
  );

folderCmd
  .command('delete')
  .description('Delete a contact folder')
  .argument('<folderId>', 'Folder id')
  .option('--confirm', 'Confirm delete')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user')
  .action(
    async (
      folderId: string,
      opts: { confirm?: boolean; token?: string; identity?: string; user?: string },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      if (!opts.confirm) {
        console.error('Refusing to delete without --confirm');
        process.exit(1);
      }
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        failContacts(undefined, 'Auth error', auth.error);
      }
      const token = auth.token;
      const r = await deleteContactFolder(token, folderId, opts.user);
      if (!r.ok) {
        failContacts(undefined, 'Error', r.error);
      }
      console.log('Deleted folder.');
    }
  );

folderCmd
  .command('children')
  .description('List child folders under a contact folder')
  .argument('<parentFolderId>', 'Parent folder id')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user')
  .action(
    async (parentFolderId: string, opts: { json?: boolean; token?: string; identity?: string; user?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        failContacts(opts.json, 'Auth error', auth.error);
      }
      const token = auth.token;
      const r = await listChildContactFolders(token, parentFolderId, opts.user);
      if (!r.ok || !r.data) {
        failContacts(opts.json, 'Error', r.error);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else {
        for (const f of r.data) {
          console.log(`${f.displayName ?? '(folder)'}\t${f.id}`);
        }
      }
    }
  );

contactsCommand.addCommand(folderCmd);

// ─── list / show / create / update / delete (contacts) ──────────────────────

contactsCommand
  .command('list')
  .description('List contacts (default folder or --folder); follows OData paging unless --top / --skip / --count')
  .option('-f, --folder <folderId>', 'Contact folder id (omit for default contacts)')
  .option('--filter <odata>', "OData $filter expression (e.g. startswith(displayName,'A'))")
  .option('--orderby <expr>', 'OData $orderby')
  .option('--select <fields>', 'OData $select (comma-separated)')
  .option('--top <n>', 'OData $top (single page)')
  .option('--skip <n>', 'OData $skip (single page)')
  .option('--count', 'Include total count ($count=true; ConsistencyLevel: eventual)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user')
  .action(
    async (opts: {
      folder?: string;
      filter?: string;
      orderby?: string;
      select?: string;
      top?: string;
      skip?: string;
      count?: boolean;
      json?: boolean;
      token?: string;
      identity?: string;
      user?: string;
    }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        failContacts(opts.json, 'Auth error', auth.error);
      }
      const token = auth.token;
      const useStructured = !!(
        opts.filter?.trim() ||
        opts.orderby?.trim() ||
        opts.select?.trim() ||
        opts.top !== undefined ||
        opts.skip !== undefined ||
        opts.count === true
      );

      const query: ContactListQuery | undefined = useStructured
        ? {
            filter: opts.filter?.trim(),
            orderby: opts.orderby?.trim(),
            select: opts.select?.trim(),
            top: opts.top !== undefined ? Number(opts.top) : undefined,
            skip: opts.skip !== undefined ? Number(opts.skip) : undefined,
            count: opts.count === true
          }
        : undefined;

      const singlePage = opts.top !== undefined || opts.skip !== undefined || opts.count === true;

      if (singlePage) {
        const r = await listContactsRawPage(token, {
          user: opts.user,
          folderId: opts.folder,
          query: query ?? {}
        });
        if (!r.ok || !r.data) {
          failContacts(opts.json, 'Error', r.error);
        }
        if (opts.json) {
          console.log(JSON.stringify(r.data, null, 2));
          return;
        }
        for (const c of r.data.value || []) {
          const em = c.emailAddresses?.[0]?.address ?? '';
          console.log(`${c.displayName ?? '(no name)'}\t${em}\t${c.id}`);
        }
        if (r.data['@odata.count'] !== undefined) {
          console.log(`Total count (@odata.count): ${r.data['@odata.count']}`);
        }
        return;
      }

      const r = opts.folder
        ? await listContactsInFolder(token, opts.folder, opts.user, query)
        : await listContacts(token, opts.user, query);
      if (!r.ok || !r.data) {
        failContacts(opts.json, 'Error', r.error);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else {
        for (const c of r.data) {
          const em = c.emailAddresses?.[0]?.address ?? '';
          console.log(`${c.displayName ?? '(no name)'}\t${em}\t${c.id}`);
        }
      }
    }
  );

contactsCommand
  .command('show')
  .description('Get one contact by id')
  .argument('<contactId>', 'Contact id')
  .option('--select <fields>', 'OData $select')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user')
  .action(
    async (
      contactId: string,
      opts: { select?: string; json?: boolean; token?: string; identity?: string; user?: string }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        failContacts(opts.json, 'Auth error', auth.error);
      }
      const token = auth.token;
      const r = await getContact(token, contactId, opts.user, opts.select);
      if (!r.ok || !r.data) {
        failContacts(opts.json, 'Error', r.error);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else {
        const em = r.data.emailAddresses?.[0]?.address ?? '';
        console.log(`${r.data.displayName ?? '(no name)'}\t${em}\t${r.data.id}`);
      }
    }
  );

contactsCommand
  .command('create')
  .description('Create a contact (JSON body per Graph contact resource)')
  .requiredOption('--json-file <path>', 'Path to JSON file')
  .option('-f, --folder <folderId>', 'Create under this contact folder (POST .../contactFolders/{id}/contacts)')
  .option('--json', 'Echo created contact as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user')
  .action(
    async (
      opts: {
        jsonFile: string;
        folder?: string;
        json?: boolean;
        token?: string;
        identity?: string;
        user?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        failContacts(opts.json, 'Auth error', auth.error);
      }
      const token = auth.token;
      const raw = await readFile(opts.jsonFile, 'utf-8');
      const body = JSON.parse(raw) as Record<string, unknown>;
      const r = await createContact(token, body, opts.user, opts.folder);
      if (!r.ok || !r.data) {
        failContacts(opts.json, 'Error', r.error);
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
  .option('--user <email>', 'Target user')
  .action(
    async (
      contactId: string,
      opts: { jsonFile: string; json?: boolean; token?: string; identity?: string; user?: string },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        failContacts(opts.json, 'Auth error', auth.error);
      }
      const token = auth.token;
      const raw = await readFile(opts.jsonFile, 'utf-8');
      const patch = JSON.parse(raw) as Record<string, unknown>;
      const r = await updateContact(token, contactId, patch, opts.user);
      if (!r.ok || !r.data) {
        failContacts(opts.json, 'Error', r.error);
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
  .option('--user <email>', 'Target user')
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
      if (!auth.success || !auth.token) {
        failContacts(undefined, 'Auth error', auth.error);
      }
      const token = auth.token;
      const r = await deleteContact(token, contactId, opts.user);
      if (!r.ok) {
        failContacts(undefined, 'Error', r.error);
      }
      console.log('Deleted contact.');
    }
  );

// ─── search ─────────────────────────────────────────────────────────────────

contactsCommand
  .command('search')
  .description('Search contacts ($search; Graph requires ConsistencyLevel: eventual)')
  .argument('<query>', 'Search string (see Graph $search syntax)')
  .option('-f, --folder <folderId>', 'Limit to a contact folder')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user')
  .action(
    async (
      query: string,
      opts: { folder?: string; json?: boolean; token?: string; identity?: string; user?: string }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        failContacts(opts.json, 'Auth error', auth.error);
      }
      const token = auth.token;
      const r = await searchContacts(token, query, opts.user, opts.folder);
      if (!r.ok || !r.data) {
        failContacts(opts.json, 'Error', r.error);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else {
        for (const c of r.data) {
          const em = c.emailAddresses?.[0]?.address ?? '';
          console.log(`${c.displayName ?? '(no name)'}\t${em}\t${c.id}`);
        }
      }
    }
  );

// ─── delta ──────────────────────────────────────────────────────────────────

contactsCommand
  .command('delta')
  .description(
    'Delta sync page (use @odata.nextLink from JSON as --next for more pages; @odata.deltaLink for baseline). Optional --state-file persists cursor.'
  )
  .option('-f, --folder <folderId>', 'Delta for contacts in this folder')
  .option('--next <url>', 'Full @odata.nextLink URL from a previous response (overrides --state-file continuation)')
  .option('--state-file <path>', 'Read/write JSON delta cursor')
  .option('--json', 'Output raw page JSON (value, nextLink, deltaLink)')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user')
  .action(
    async (opts: {
      folder?: string;
      next?: string;
      stateFile?: string;
      json?: boolean;
      token?: string;
      identity?: string;
      user?: string;
    }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        failContacts(opts.json, 'Auth error', auth.error);
      }
      const token = auth.token;
      const existingState = opts.stateFile ? await readDeltaStateFile(opts.stateFile) : null;
      if (existingState && existingState.kind !== 'contacts') {
        console.error('Error: state file is not for contacts delta (kind must be contacts).');
        process.exit(1);
      }
      try {
        if (existingState) {
          assertDeltaScopeMatchesState(existingState, { folderId: opts.folder, user: opts.user });
        }
      } catch (err) {
        console.error(err instanceof Error ? err.message : err);
        process.exit(1);
      }
      const continueUrl = resolveDeltaContinuationUrl({ explicitNext: opts.next, state: existingState });
      const r = await contactsDeltaPage(token, {
        user: opts.user,
        folderId: opts.folder,
        nextLink: continueUrl
      });
      if (!r.ok || !r.data) {
        failContacts(opts.json, 'Error', r.error);
      }
      if (opts.stateFile && r.data) {
        const merged = applyDeltaPageToState(existingState, 'contacts', r.data, {
          folderId: opts.folder,
          user: opts.user
        });
        await writeDeltaStateFile(opts.stateFile, merged);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else {
        console.log(`Changes: ${r.data.value?.length ?? 0} item(s)`);
        if (r.data['@odata.nextLink']) console.log(`nextLink: ${r.data['@odata.nextLink']}`);
        if (r.data['@odata.deltaLink']) console.log(`deltaLink: ${r.data['@odata.deltaLink']}`);
        if (opts.stateFile) console.log(`state-file: ${opts.stateFile} (updated)`);
      }
    }
  );

// ─── extension (open extensions) ────────────────────────────────────────────

const contactExtensionCmd = new Command('extension').description('Open type extensions on a contact (Graph)');

contactExtensionCmd
  .command('list')
  .description(
    'List open extensions on a contact (default …/contacts/{id}/extensions; use --folder for contactFolders path)'
  )
  .argument('<contactId>', 'Contact id')
  .option('-f, --folder <folderId>', 'Contact folder id (Graph …/contactFolders/{id}/contacts/{contactId}/extensions)')
  .option('--child-folder <folderId>', 'Child folder under --folder (…/childFolders/{id}/contacts/…)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user')
  .action(
    async (
      contactId: string,
      opts: {
        folder?: string;
        childFolder?: string;
        json?: boolean;
        token?: string;
        identity?: string;
        user?: string;
      }
    ) => {
      const loc = resolveContactExtensionLocation(opts);
      if (loc === 'invalid') {
        console.error('Error: --child-folder requires -f/--folder');
        process.exit(1);
      }
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        failContacts(opts.json, 'Auth error', auth.error);
      }
      const token = auth.token;
      const r = await listContactOpenExtensions(token, contactId, opts.user, loc);
      if (!r.ok || !r.data) {
        failContacts(opts.json, 'Error', r.error);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else {
        for (const ext of r.data) {
          const name = (ext.extensionName as string) || JSON.stringify(ext);
          console.log(`- ${name}`);
        }
      }
    }
  );

contactExtensionCmd
  .command('get')
  .description('Get one open extension by name')
  .argument('<contactId>', 'Contact id')
  .requiredOption('-n, --name <id>', 'extensionName')
  .option('-f, --folder <folderId>', 'Contact folder id')
  .option('--child-folder <folderId>', 'Child folder under --folder')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user')
  .action(
    async (
      contactId: string,
      opts: {
        name: string;
        folder?: string;
        childFolder?: string;
        json?: boolean;
        token?: string;
        identity?: string;
        user?: string;
      }
    ) => {
      const loc = resolveContactExtensionLocation(opts);
      if (loc === 'invalid') {
        console.error('Error: --child-folder requires -f/--folder');
        process.exit(1);
      }
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        failContacts(opts.json, 'Auth error', auth.error);
      }
      const token = auth.token;
      const r = await getContactOpenExtension(token, contactId, opts.name, opts.user, loc);
      if (!r.ok || !r.data) {
        failContacts(opts.json, 'Error', r.error);
      }
      if (opts.json) {
        console.log(JSON.stringify(r.data, null, 2));
      } else {
        const { extensionName, ...props } = r.data;
        console.log(`- ${(extensionName as string) || opts.name}`);
        for (const [key, value] of Object.entries(props)) {
          if (key.startsWith('@') || key === 'id') continue;
          console.log(`    ${key}: ${JSON.stringify(value)}`);
        }
      }
    }
  );

contactExtensionCmd
  .command('set')
  .description('Create an open extension (POST); JSON file is merged with extensionName')
  .argument('<contactId>', 'Contact id')
  .requiredOption('-n, --name <id>', 'extensionName')
  .requiredOption('--json-file <path>', 'JSON object: custom properties (extensionName added automatically)')
  .option('-f, --folder <folderId>', 'Contact folder id')
  .option('--child-folder <folderId>', 'Child folder under --folder')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user')
  .action(
    async (
      contactId: string,
      opts: {
        name: string;
        jsonFile: string;
        folder?: string;
        childFolder?: string;
        json?: boolean;
        token?: string;
        identity?: string;
        user?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const loc = resolveContactExtensionLocation(opts);
      if (loc === 'invalid') {
        console.error('Error: --child-folder requires -f/--folder');
        process.exit(1);
      }
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        failContacts(opts.json, 'Auth error', auth.error);
      }
      const token = auth.token;
      const raw = await readFile(opts.jsonFile, 'utf-8');
      const data = JSON.parse(raw) as Record<string, unknown>;
      const r = await setContactOpenExtension(token, contactId, opts.name, data, opts.user, loc);
      if (!r.ok || !r.data) {
        failContacts(opts.json, 'Error', r.error);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else console.log(`\n\u2705 Extension set: ${opts.name}\n`);
    }
  );

contactExtensionCmd
  .command('update')
  .description('PATCH an open extension (partial update)')
  .argument('<contactId>', 'Contact id')
  .requiredOption('-n, --name <id>', 'extensionName')
  .requiredOption('--json-file <path>', 'JSON object: properties to patch')
  .option('-f, --folder <folderId>', 'Contact folder id')
  .option('--child-folder <folderId>', 'Child folder under --folder')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user')
  .action(
    async (
      contactId: string,
      opts: {
        name: string;
        jsonFile: string;
        folder?: string;
        childFolder?: string;
        token?: string;
        identity?: string;
        user?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const loc = resolveContactExtensionLocation(opts);
      if (loc === 'invalid') {
        console.error('Error: --child-folder requires -f/--folder');
        process.exit(1);
      }
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        failContacts(undefined, 'Auth error', auth.error);
      }
      const token = auth.token;
      const raw = await readFile(opts.jsonFile, 'utf-8');
      const patch = JSON.parse(raw) as Record<string, unknown>;
      const r = await updateContactOpenExtension(token, contactId, opts.name, patch, opts.user, loc);
      if (!r.ok) {
        failContacts(undefined, 'Error', r.error);
      }
      console.log('\n\u2705 Extension updated.\n');
    }
  );

contactExtensionCmd
  .command('delete')
  .description('Delete an open extension')
  .argument('<contactId>', 'Contact id')
  .requiredOption('-n, --name <id>', 'extensionName')
  .option('-f, --folder <folderId>', 'Contact folder id')
  .option('--child-folder <folderId>', 'Child folder under --folder')
  .option('--confirm', 'Confirm without prompt')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user')
  .action(
    async (
      contactId: string,
      opts: {
        name: string;
        folder?: string;
        childFolder?: string;
        confirm?: boolean;
        token?: string;
        identity?: string;
        user?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      if (!opts.confirm) {
        console.log(`Delete extension "${opts.name}"? Run with --confirm.`);
        process.exit(1);
      }
      const loc = resolveContactExtensionLocation(opts);
      if (loc === 'invalid') {
        console.error('Error: --child-folder requires -f/--folder');
        process.exit(1);
      }
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        failContacts(undefined, 'Auth error', auth.error);
      }
      const token = auth.token;
      const r = await deleteContactOpenExtension(token, contactId, opts.name, opts.user, loc);
      if (!r.ok) {
        failContacts(undefined, 'Error', r.error);
      }
      console.log(`\n\u2705 Deleted extension: ${opts.name}\n`);
    }
  );

contactsCommand.addCommand(contactExtensionCmd);

// ─── photo ──────────────────────────────────────────────────────────────────

const photoCmd = new Command('photo').description('Contact profile photo (GET/PUT/DELETE .../photo)');

photoCmd
  .command('get')
  .description('Download contact photo bytes to a file')
  .argument('<contactId>', 'Contact id')
  .requiredOption('--out <path>', 'Output file path')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user')
  .action(async (contactId: string, opts: { out: string; token?: string; identity?: string; user?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      failContacts(undefined, 'Auth error', auth.error);
    }
    const token = auth.token;
    const r = await getContactPhotoBytes(token, contactId, opts.user);
    if (!r.ok || !r.data) {
      failContacts(undefined, 'Error', r.error);
    }
    await writeFile(opts.out, r.data);
    console.log(`Wrote ${r.data.byteLength} bytes to ${opts.out}`);
  });

photoCmd
  .command('set')
  .description('Upload a photo (JPEG/PNG recommended)')
  .argument('<contactId>', 'Contact id')
  .requiredOption('--file <path>', 'Image file path')
  .option('--content-type <mime>', 'Content-Type (default: image/jpeg)')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user')
  .action(
    async (
      contactId: string,
      opts: { file: string; contentType?: string; token?: string; identity?: string; user?: string },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        failContacts(undefined, 'Auth error', auth.error);
      }
      const token = auth.token;
      const bytes = await readFile(opts.file);
      const ct = opts.contentType ?? 'image/jpeg';
      const r = await setContactPhoto(token, contactId, new Uint8Array(bytes), ct, opts.user);
      if (!r.ok) {
        failContacts(undefined, 'Error', r.error);
      }
      console.log('Photo updated.');
    }
  );

photoCmd
  .command('delete')
  .description('Remove contact photo')
  .argument('<contactId>', 'Contact id')
  .option('--confirm', 'Confirm')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user')
  .action(
    async (
      contactId: string,
      opts: { confirm?: boolean; token?: string; identity?: string; user?: string },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      if (!opts.confirm) {
        console.error('Refusing to delete photo without --confirm');
        process.exit(1);
      }
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        failContacts(undefined, 'Auth error', auth.error);
      }
      const token = auth.token;
      const r = await deleteContactPhoto(token, contactId, opts.user);
      if (!r.ok) {
        failContacts(undefined, 'Error', r.error);
      }
      console.log('Photo deleted.');
    }
  );

contactsCommand.addCommand(photoCmd);

// ─── attachments ────────────────────────────────────────────────────────────

const attachCmd = new Command('attachments').description('File attachments on a contact');

attachCmd
  .command('list')
  .description('List attachments')
  .argument('<contactId>', 'Contact id')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user')
  .action(async (contactId: string, opts: { json?: boolean; token?: string; identity?: string; user?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      failContacts(opts.json, 'Auth error', auth.error);
    }
    const token = auth.token;
    const r = await listContactAttachments(token, contactId, opts.user);
    if (!r.ok || !r.data) {
      failContacts(opts.json, 'Error', r.error);
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else {
      for (const a of r.data) {
        console.log(`${a.name ?? a.id}\t${a.id}\t${a.contentType ?? ''}`);
      }
    }
  });

attachCmd
  .command('add')
  .description('Add a file attachment (base64 upload)')
  .argument('<contactId>', 'Contact id')
  .requiredOption('--file <path>', 'File to attach')
  .option('--name <filename>', 'Attachment name (default: basename of file)')
  .option('--content-type <mime>', 'MIME type (default: application/octet-stream)')
  .option('--json', 'Echo attachment metadata as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user')
  .action(
    async (
      contactId: string,
      opts: {
        file: string;
        name?: string;
        contentType?: string;
        json?: boolean;
        token?: string;
        identity?: string;
        user?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        failContacts(opts.json, 'Auth error', auth.error);
      }
      const token = auth.token;
      const buf = await readFile(opts.file);
      const base64 = Buffer.from(buf).toString('base64');
      const pathParts = opts.file.replace(/\\/g, '/').split('/');
      const baseName = pathParts[pathParts.length - 1] ?? 'attachment';
      const name = opts.name ?? baseName;
      const contentType = opts.contentType ?? 'application/octet-stream';
      const r = await addFileAttachmentToContact(
        token,
        contactId,
        { name, contentType, contentBytes: base64 },
        opts.user
      );
      if (!r.ok || !r.data) {
        failContacts(opts.json, 'Error', r.error);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else console.log(`Added attachment: ${r.data.name ?? r.data.id} (${r.data.id})`);
    }
  );

attachCmd
  .command('add-link')
  .description('Add a link attachment (Graph `referenceAttachment`)')
  .argument('<contactId>', 'Contact id')
  .requiredOption('--name <title>', 'Link title / attachment name')
  .requiredOption('--url <httpsUrl>', 'Target URL')
  .option('--json', 'Echo attachment metadata as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user')
  .action(
    async (
      contactId: string,
      opts: {
        name: string;
        url: string;
        json?: boolean;
        token?: string;
        identity?: string;
        user?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        failContacts(opts.json, 'Auth error', auth.error);
      }
      const token = auth.token;
      const r = await addReferenceAttachmentToContact(
        token,
        contactId,
        {
          name: opts.name,
          sourceUrl: opts.url
        },
        opts.user
      );
      if (!r.ok || !r.data) {
        failContacts(opts.json, 'Error', r.error);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else console.log(`Added link attachment: ${r.data.name ?? r.data.id} (${r.data.id})`);
    }
  );

attachCmd
  .command('show')
  .description('Get attachment metadata')
  .argument('<contactId>', 'Contact id')
  .argument('<attachmentId>', 'Attachment id')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user')
  .action(
    async (
      contactId: string,
      attachmentId: string,
      opts: { json?: boolean; token?: string; identity?: string; user?: string }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        failContacts(opts.json, 'Auth error', auth.error);
      }
      const token = auth.token;
      const r = await getContactAttachment(token, contactId, attachmentId, opts.user);
      if (!r.ok || !r.data) {
        failContacts(opts.json, 'Error', r.error);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else console.log(`${r.data.name ?? r.data.id}\t${r.data.id}\t${r.data.contentType ?? ''}`);
    }
  );

attachCmd
  .command('download')
  .description('Download file attachment raw bytes')
  .argument('<contactId>', 'Contact id')
  .argument('<attachmentId>', 'Attachment id')
  .requiredOption('--out <path>', 'Output file path')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user')
  .action(
    async (
      contactId: string,
      attachmentId: string,
      opts: { out: string; token?: string; identity?: string; user?: string }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        failContacts(undefined, 'Auth error', auth.error);
      }
      const token = auth.token;
      const r = await downloadContactAttachmentBytes(token, contactId, attachmentId, opts.user);
      if (!r.ok || !r.data) {
        failContacts(undefined, 'Error', r.error);
      }
      await writeFile(opts.out, r.data);
      console.log(`Wrote ${r.data.byteLength} bytes to ${opts.out}`);
    }
  );

attachCmd
  .command('delete')
  .description('Delete an attachment')
  .argument('<contactId>', 'Contact id')
  .argument('<attachmentId>', 'Attachment id')
  .option('--confirm', 'Confirm delete')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user')
  .action(
    async (
      contactId: string,
      attachmentId: string,
      opts: { confirm?: boolean; token?: string; identity?: string; user?: string },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      if (!opts.confirm) {
        console.error('Refusing to delete without --confirm');
        process.exit(1);
      }
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        failContacts(undefined, 'Auth error', auth.error);
      }
      const token = auth.token;
      const r = await deleteContactAttachment(token, contactId, attachmentId, opts.user);
      if (!r.ok) {
        failContacts(undefined, 'Error', r.error);
      }
      console.log('Attachment deleted.');
    }
  );

const mergeSuggestionsCmd = new Command('merge-suggestions').description(
  'Duplicate contact merge suggestions visibility (Graph **beta**: `/me/settings/contactMergeSuggestions` or `/users/{id}/…`). Uses `User.Read` / `User.ReadWrite` for self; delegated other users typically need `User.Read.All` / `User.ReadWrite.All`.'
);

mergeSuggestionsCmd
  .command('get')
  .description('Read contactMergeSuggestions settings (JSON to stdout)')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (delegated)')
  .action(async (opts: { json?: boolean; token?: string; identity?: string; user?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      failContacts(opts.json, 'Auth error', auth.error);
    }
    const token = auth.token;
    const r = await getContactMergeSuggestions(token, opts.user);
    if (!r.ok || !r.data) {
      failContacts(opts.json, 'Error', r.error);
    }
    console.log(JSON.stringify(r.data, null, 2));
  });

mergeSuggestionsCmd
  .command('set')
  .description('PATCH contactMergeSuggestions (`--json-file` body per Graph schema)')
  .requiredOption('--json-file <path>', 'JSON body for PATCH')
  .option('--json', 'Echo updated resource as JSON after PATCH')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (delegated)')
  .action(
    async (
      opts: { jsonFile: string; json?: boolean; token?: string; identity?: string; user?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        failContacts(opts.json, 'Auth error', auth.error);
      }
      const token = auth.token;
      const raw = await readFile(opts.jsonFile.trim(), 'utf8');
      const body = JSON.parse(raw) as Record<string, unknown>;
      const r = await patchContactMergeSuggestions(token, body, opts.user);
      if (!r.ok || !r.data) {
        failContacts(opts.json, 'Error', r.error);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else console.log('Updated contact merge suggestions settings.');
    }
  );

mergeSuggestionsCmd
  .command('delete')
  .description('DELETE contactMergeSuggestions navigation property (requires If-Match; omit to fetch ETag first)')
  .option('--if-match <etag>', 'If-Match header from `merge-suggestions get --json`')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (delegated)')
  .action(async (opts: { ifMatch?: string; token?: string; identity?: string; user?: string }, cmd: Command) => {
    checkReadOnly(cmd);
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      failContacts(undefined, 'Auth error', auth.error);
    }
    const token = auth.token;
    let ifMatch = opts.ifMatch?.trim();
    if (!ifMatch) {
      const gr = await getContactMergeSuggestions(token, opts.user);
      if (!gr.ok || !gr.data) {
        failContacts(undefined, 'Error', gr.error);
      }
      ifMatch = (gr.data as { '@odata.etag'?: string })['@odata.etag']?.trim();
      if (!ifMatch) {
        console.error('Error: missing @odata.etag; pass --if-match');
        process.exit(1);
      }
    }
    const r = await deleteContactMergeSuggestions(token, ifMatch, opts.user);
    if (!r.ok) {
      failContacts(undefined, 'Error', r.error);
    }
    console.log('Deleted contact merge suggestions settings resource.');
  });

contactsCommand.addCommand(mergeSuggestionsCmd);
contactsCommand.addCommand(attachCmd);
