import { readFile, writeFile } from 'node:fs/promises';
import { Command } from 'commander';
import { requireGraphAuth } from '../lib/graph-auth.js';
import {
  addFileAttachmentToContact,
  contactsDeltaPage,
  createContact,
  createContactFolder,
  deleteContact,
  deleteContactAttachment,
  deleteContactFolder,
  deleteContactPhoto,
  downloadContactAttachmentBytes,
  getContact,
  getContactAttachment,
  getContactFolder,
  getContactPhotoBytes,
  listChildContactFolders,
  listContactAttachments,
  listContactFolders,
  listContacts,
  listContactsInFolder,
  searchContacts,
  setContactPhoto,
  updateContact,
  updateContactFolder
} from '../lib/outlook-graph-client.js';
import { checkReadOnly } from '../lib/utils.js';

const contactsDesc =
  'Outlook contacts via Microsoft Graph (Contacts.ReadWrite; shared mailboxes: Contacts.Read.Shared / Contacts.ReadWrite.Shared — see docs/GRAPH_SCOPES.md)';

export const contactsCommand = new Command('contacts').description(contactsDesc);

// ─── folders (list — backward compatible) ───────────────────────────────────

contactsCommand
  .command('folders')
  .description('List contact folders (Graph GET /contactFolders); see also `contacts folder list`')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (delegated / shared mailbox)')
  .action(async (opts: { json?: boolean; token?: string; identity?: string; user?: string }) => {
    const token = await requireGraphAuth(opts);
    const r = await listContactFolders(token, opts.user);
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

// ─── folder (CRUD + children) ───────────────────────────────────────────────

const folderCmd = new Command('folder').description('Contact folder create, read, update, delete, list');

folderCmd
  .command('list')
  .description('List contact folders')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user')
  .action(async (opts: { json?: boolean; token?: string; identity?: string; user?: string }) => {
    const token = await requireGraphAuth(opts);
    const r = await listContactFolders(token, opts.user);
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

folderCmd
  .command('get')
  .description('Get one contact folder by id')
  .argument('<folderId>', 'Folder id')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user')
  .action(async (folderId: string, opts: { json?: boolean; token?: string; identity?: string; user?: string }) => {
    const token = await requireGraphAuth(opts);
    const r = await getContactFolder(token, folderId, opts.user);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
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
      const token = await requireGraphAuth(opts);
      const r = await createContactFolder(token, opts.name, opts.user, opts.parent);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
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
      const token = await requireGraphAuth(opts);
      const raw = await readFile(opts.jsonFile, 'utf-8');
      const patch = JSON.parse(raw) as Record<string, unknown>;
      const r = await updateContactFolder(token, folderId, patch, opts.user);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
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
      const token = await requireGraphAuth(opts);
      const r = await deleteContactFolder(token, folderId, opts.user);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
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
      const token = await requireGraphAuth(opts);
      const r = await listChildContactFolders(token, parentFolderId, opts.user);
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
    }
  );

contactsCommand.addCommand(folderCmd);

// ─── list / show / create / update / delete (contacts) ──────────────────────

contactsCommand
  .command('list')
  .description('List contacts (default folder or --folder)')
  .option('-f, --folder <folderId>', 'Contact folder id (omit for default contacts)')
  .option('--filter <odata>', "OData fragment for $filter=… (e.g. `startswith(displayName,\\'A\\')`)")
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user')
  .action(
    async (opts: {
      folder?: string;
      filter?: string;
      json?: boolean;
      token?: string;
      identity?: string;
      user?: string;
    }) => {
      const token = await requireGraphAuth(opts);
      let odata: string | undefined;
      if (opts.filter?.trim()) {
        odata = `$filter=${encodeURIComponent(opts.filter.trim())}`;
      }
      const r = opts.folder
        ? await listContactsInFolder(token, opts.folder, opts.user, odata)
        : await listContacts(token, opts.user, odata);
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
      const token = await requireGraphAuth(opts);
      const r = await getContact(token, contactId, opts.user, opts.select);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
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
      const token = await requireGraphAuth(opts);
      const raw = await readFile(opts.jsonFile, 'utf-8');
      const body = JSON.parse(raw) as Record<string, unknown>;
      const r = await createContact(token, body, opts.user, opts.folder);
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
  .option('--user <email>', 'Target user')
  .action(
    async (
      contactId: string,
      opts: { jsonFile: string; json?: boolean; token?: string; identity?: string; user?: string },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const token = await requireGraphAuth(opts);
      const raw = await readFile(opts.jsonFile, 'utf-8');
      const patch = JSON.parse(raw) as Record<string, unknown>;
      const r = await updateContact(token, contactId, patch, opts.user);
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
      const token = await requireGraphAuth(opts);
      const r = await deleteContact(token, contactId, opts.user);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
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
      const token = await requireGraphAuth(opts);
      const r = await searchContacts(token, query, opts.user, opts.folder);
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
    }
  );

// ─── delta ──────────────────────────────────────────────────────────────────

contactsCommand
  .command('delta')
  .description(
    'Delta sync page (use @odata.nextLink from JSON as --next for more pages; @odata.deltaLink for baseline)'
  )
  .option('-f, --folder <folderId>', 'Delta for contacts in this folder')
  .option('--next <url>', 'Full @odata.nextLink URL from a previous response')
  .option('--json', 'Output raw page JSON (value, nextLink, deltaLink)')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user')
  .action(
    async (opts: {
      folder?: string;
      next?: string;
      json?: boolean;
      token?: string;
      identity?: string;
      user?: string;
    }) => {
      const token = await requireGraphAuth(opts);
      const r = await contactsDeltaPage(token, {
        user: opts.user,
        folderId: opts.folder,
        nextLink: opts.next
      });
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else {
        console.log(`Changes: ${r.data.value?.length ?? 0} item(s)`);
        if (r.data['@odata.nextLink']) console.log(`nextLink: ${r.data['@odata.nextLink']}`);
        if (r.data['@odata.deltaLink']) console.log(`deltaLink: ${r.data['@odata.deltaLink']}`);
      }
    }
  );

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
    const token = await requireGraphAuth(opts);
    const r = await getContactPhotoBytes(token, contactId, opts.user);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
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
      const token = await requireGraphAuth(opts);
      const bytes = await readFile(opts.file);
      const ct = opts.contentType ?? 'image/jpeg';
      const r = await setContactPhoto(token, contactId, new Uint8Array(bytes), ct, opts.user);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
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
      const token = await requireGraphAuth(opts);
      const r = await deleteContactPhoto(token, contactId, opts.user);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
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
    const token = await requireGraphAuth(opts);
    const r = await listContactAttachments(token, contactId, opts.user);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
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
      const token = await requireGraphAuth(opts);
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
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else console.log(`Added attachment: ${r.data.name ?? r.data.id} (${r.data.id})`);
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
      const token = await requireGraphAuth(opts);
      const r = await getContactAttachment(token, contactId, attachmentId, opts.user);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
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
      const token = await requireGraphAuth(opts);
      const r = await downloadContactAttachmentBytes(token, contactId, attachmentId, opts.user);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
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
      const token = await requireGraphAuth(opts);
      const r = await deleteContactAttachment(token, contactId, attachmentId, opts.user);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log('Attachment deleted.');
    }
  );

contactsCommand.addCommand(attachCmd);
