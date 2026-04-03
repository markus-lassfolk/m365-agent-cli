import { readFile, writeFile } from 'node:fs/promises';
import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import {
  copyMailMessage,
  createContact,
  createMailFolder,
  createMailForwardDraft,
  createMailReplyAllDraft,
  createMailReplyDraft,
  deleteContact,
  deleteMailFolder,
  deleteMailMessage,
  downloadMailMessageAttachmentBytes,
  getContact,
  getMailFolder,
  getMailMessageAttachment,
  getMessage,
  listChildMailFolders,
  listContacts,
  listMailboxMessages,
  listMailFolders,
  listMailMessageAttachments,
  listMessagesInFolder,
  type MessagesQueryOptions,
  mailMessagesDeltaPage,
  moveMailMessage,
  patchMailMessage,
  type RootMailboxMessagesQuery,
  sendMail,
  sendMailMessage,
  updateContact,
  updateMailFolder
} from '../lib/outlook-graph-client.js';
import { checkReadOnly } from '../lib/utils.js';

export const outlookGraphCommand = new Command('outlook-graph').description(
  'Microsoft Graph Outlook REST: mail folders, messages (list/send/patch/move/copy/attachments/reply), contacts (distinct from EWS mail/folders)'
);

outlookGraphCommand
  .command('list-folders')
  .description('List mail folders (Graph GET /mailFolders)')
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
    const r = await listMailFolders(auth.token!, opts.user);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else {
      for (const f of r.data) {
        console.log(`${f.displayName}\t${f.id}`);
      }
    }
  });

outlookGraphCommand
  .command('child-folders')
  .description('List child folders under a parent folder id')
  .requiredOption('-p, --parent <folderId>', 'Parent folder id (e.g. Inbox id from list-folders)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
  .action(async (opts: { parent: string; json?: boolean; token?: string; identity?: string; user?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await listChildMailFolders(auth.token!, opts.parent, opts.user);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else {
      for (const f of r.data) {
        console.log(`${f.displayName}\t${f.id}`);
      }
    }
  });

outlookGraphCommand
  .command('get-folder')
  .description('Get one mail folder by id (well-known: inbox, sentitems, drafts, deleteditems, archive, junkemail)')
  .argument('<folderId>', 'Folder id')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
  .action(async (folderId: string, opts: { json?: boolean; token?: string; identity?: string; user?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await getMailFolder(auth.token!, folderId, opts.user);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else console.log(JSON.stringify(r.data, null, 2));
  });

outlookGraphCommand
  .command('create-folder')
  .description('Create a mail folder (optionally under --parent)')
  .requiredOption('-n, --name <displayName>', 'Folder display name')
  .option('-p, --parent <folderId>', 'Parent folder id (omit for root-level where allowed)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
  .action(
    async (
      opts: { name: string; parent?: string; json?: boolean; token?: string; identity?: string; user?: string },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await createMailFolder(auth.token!, opts.name, opts.parent, opts.user);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else console.log(`Created folder: ${r.data.displayName} (${r.data.id})`);
    }
  );

outlookGraphCommand
  .command('update-folder')
  .description('Rename a mail folder (PATCH displayName)')
  .argument('<folderId>', 'Folder id')
  .requiredOption('-n, --name <displayName>', 'New display name')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
  .action(
    async (
      folderId: string,
      opts: { name: string; json?: boolean; token?: string; identity?: string; user?: string },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await updateMailFolder(auth.token!, folderId, opts.name, opts.user);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else console.log(`Updated folder: ${r.data.displayName}`);
    }
  );

outlookGraphCommand
  .command('delete-folder')
  .description('Delete a mail folder (not allowed for default folders)')
  .argument('<folderId>', 'Folder id')
  .option('--confirm', 'Confirm delete')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
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
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await deleteMailFolder(auth.token!, folderId, opts.user);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log('Deleted folder.');
    }
  );

outlookGraphCommand
  .command('list-messages')
  .description('List messages in a folder (Graph; use well-known id e.g. inbox)')
  .requiredOption('-f, --folder <folderId>', 'Folder id (e.g. inbox)')
  .option('--top <n>', 'Page size (default 25). Omit with --all to page entire folder', '25')
  .option('--all', 'Follow all pages (may be slow/large)')
  .option('--filter <odata>', 'OData $filter')
  .option('--orderby <odata>', 'OData $orderby')
  .option('--select <fields>', 'OData $select (comma-separated)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
  .action(
    async (opts: {
      folder: string;
      top?: string;
      all?: boolean;
      filter?: string;
      orderby?: string;
      select?: string;
      json?: boolean;
      token?: string;
      identity?: string;
      user?: string;
    }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const q: MessagesQueryOptions = {};
      if (opts.filter) q.filter = opts.filter;
      if (opts.orderby) q.orderby = opts.orderby;
      if (opts.select) q.select = opts.select;
      if (opts.all) {
        // fetch all pages — no $top
      } else {
        q.top = Math.max(1, parseInt(opts.top ?? '25', 10) || 25);
      }
      const r = await listMessagesInFolder(auth.token!, opts.folder, opts.user, q);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else {
        for (const m of r.data) {
          const from = m.from?.emailAddress?.address ?? '';
          const sub = m.subject ?? '(no subject)';
          console.log(`${m.receivedDateTime ?? ''}\t${from}\t${sub}\t${m.id}`);
        }
      }
    }
  );

outlookGraphCommand
  .command('messages-delta')
  .description(
    'One page of messages delta sync (use @odata.nextLink as --next for more pages; @odata.deltaLink for baseline)'
  )
  .option('-f, --folder <folderId>', 'Delta for messages in this mail folder only (omit for all messages)')
  .option('--next <url>', 'Full @odata.nextLink URL from a previous response')
  .option('--json', 'Output raw page JSON (value, nextLink, deltaLink)')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
  .action(
    async (opts: {
      folder?: string;
      next?: string;
      json?: boolean;
      token?: string;
      identity?: string;
      user?: string;
    }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await mailMessagesDeltaPage(auth.token!, {
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

outlookGraphCommand
  .command('list-mail')
  .description('List messages mailbox-wide (Graph GET /messages; use --folder on list-messages for one folder)')
  .option('--top <n>', 'Page size (default 25). Omit with --all to page entire result set', '25')
  .option('--all', 'Follow all pages (may be slow/large)')
  .option('--filter <odata>', 'OData $filter (do not combine with --search)')
  .option('--orderby <odata>', 'OData $orderby')
  .option('--select <fields>', 'OData $select (comma-separated)')
  .option('--skip <n>', 'OData $skip')
  .option('--search <text>', 'Keyword search ($search; adds ConsistencyLevel; do not combine with --filter)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
  .action(
    async (opts: {
      top?: string;
      all?: boolean;
      filter?: string;
      orderby?: string;
      select?: string;
      skip?: string;
      search?: string;
      json?: boolean;
      token?: string;
      identity?: string;
      user?: string;
    }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const q: RootMailboxMessagesQuery = {};
      if (opts.filter) q.filter = opts.filter;
      if (opts.orderby) q.orderby = opts.orderby;
      if (opts.select) q.select = opts.select;
      if (opts.search) q.search = opts.search;
      if (opts.skip !== undefined) q.skip = Math.max(0, parseInt(opts.skip, 10) || 0);
      if (opts.all) {
        // full pagination
      } else {
        q.top = Math.max(1, parseInt(opts.top ?? '25', 10) || 25);
      }
      const r = await listMailboxMessages(auth.token!, opts.user, q);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else {
        for (const m of r.data) {
          const from = m.from?.emailAddress?.address ?? '';
          const sub = m.subject ?? '(no subject)';
          console.log(`${m.receivedDateTime ?? ''}\t${from}\t${sub}\t${m.id}`);
        }
      }
    }
  );

outlookGraphCommand
  .command('get-message')
  .description('Get one message by id (Graph GET /messages/{id})')
  .requiredOption('-i, --id <messageId>', 'Message id')
  .option('--select <fields>', 'OData $select')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
  .action(
    async (opts: { id: string; select?: string; json?: boolean; token?: string; identity?: string; user?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getMessage(auth.token!, opts.id, opts.user, opts.select);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else console.log(JSON.stringify(r.data, null, 2));
    }
  );

outlookGraphCommand
  .command('send-mail')
  .description('Send email in one request (Graph POST /sendMail; JSON body with message + optional saveToSentItems)')
  .requiredOption('--json-file <path>', 'Path to JSON: { "message": { ... }, "saveToSentItems": true }')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
  .action(async (opts: { jsonFile: string; token?: string; identity?: string; user?: string }, cmd: any) => {
    checkReadOnly(cmd);
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const raw = await readFile(opts.jsonFile, 'utf-8');
    const body = JSON.parse(raw) as { message: Record<string, unknown>; saveToSentItems?: boolean };
    if (!body.message) {
      console.error('JSON must include a "message" object');
      process.exit(1);
    }
    const r = await sendMail(auth.token!, body, opts.user);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log('Sent.');
  });

outlookGraphCommand
  .command('patch-message')
  .description('PATCH a message (e.g. isRead, flag, categories — Graph PATCH /messages/{id})')
  .argument('<messageId>', 'Message id')
  .requiredOption('--json-file <path>', 'Path to JSON patch body')
  .option('--json', 'Echo updated message as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
  .action(
    async (
      messageId: string,
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
      const r = await patchMailMessage(auth.token!, messageId, patch, opts.user);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else console.log(`Updated message ${r.data.id}`);
    }
  );

outlookGraphCommand
  .command('delete-message')
  .description('Delete a message (Graph DELETE /messages/{id})')
  .argument('<messageId>', 'Message id')
  .option('--confirm', 'Confirm delete')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
  .action(
    async (
      messageId: string,
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
      const r = await deleteMailMessage(auth.token!, messageId, opts.user);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log('Deleted message.');
    }
  );

outlookGraphCommand
  .command('move-message')
  .description('Move a message to another folder (Graph POST .../move)')
  .argument('<messageId>', 'Message id')
  .requiredOption('-d, --destination <folderId>', 'Destination folder id (e.g. deleteditems)')
  .option('--json', 'Echo moved message as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
  .action(
    async (
      messageId: string,
      opts: { destination: string; json?: boolean; token?: string; identity?: string; user?: string },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await moveMailMessage(auth.token!, messageId, opts.destination, opts.user);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else console.log(`Moved to folder; message id: ${r.data.id}`);
    }
  );

outlookGraphCommand
  .command('copy-message')
  .description('Copy a message to another folder (Graph POST .../copy)')
  .argument('<messageId>', 'Message id')
  .requiredOption('-d, --destination <folderId>', 'Destination folder id')
  .option('--json', 'Echo copy as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
  .action(
    async (
      messageId: string,
      opts: { destination: string; json?: boolean; token?: string; identity?: string; user?: string },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await copyMailMessage(auth.token!, messageId, opts.destination, opts.user);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else console.log(`Copy id: ${r.data.id}`);
    }
  );

outlookGraphCommand
  .command('list-message-attachments')
  .description('List attachments on a message (Graph GET .../messages/{id}/attachments)')
  .requiredOption('-i, --id <messageId>', 'Message id')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
  .action(async (opts: { id: string; json?: boolean; token?: string; identity?: string; user?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await listMailMessageAttachments(auth.token!, opts.id, opts.user);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else {
      for (const a of r.data) {
        console.log(`${a.name ?? a.id}\t${a.contentType ?? ''}\t${a.id}`);
      }
    }
  });

outlookGraphCommand
  .command('get-message-attachment')
  .description('Get attachment metadata (Graph GET .../attachments/{id})')
  .requiredOption('-i, --id <messageId>', 'Message id')
  .requiredOption('-a, --attachment <attachmentId>', 'Attachment id')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
  .action(
    async (opts: {
      id: string;
      attachment: string;
      json?: boolean;
      token?: string;
      identity?: string;
      user?: string;
    }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getMailMessageAttachment(auth.token!, opts.id, opts.attachment, opts.user);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else console.log(JSON.stringify(r.data, null, 2));
    }
  );

outlookGraphCommand
  .command('download-message-attachment')
  .description('Download file attachment bytes (Graph GET .../attachments/{id}/$value)')
  .requiredOption('-i, --id <messageId>', 'Message id')
  .requiredOption('-a, --attachment <attachmentId>', 'Attachment id')
  .requiredOption('-o, --output <path>', 'Output file path')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
  .action(
    async (opts: {
      id: string;
      attachment: string;
      output: string;
      token?: string;
      identity?: string;
      user?: string;
    }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await downloadMailMessageAttachmentBytes(auth.token!, opts.id, opts.attachment, opts.user);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      await writeFile(opts.output, r.data);
      console.log(`Wrote ${r.data.byteLength} bytes to ${opts.output}`);
    }
  );

outlookGraphCommand
  .command('create-reply')
  .description('Create a reply draft (Graph POST .../createReply); then patch body and outlook-graph send-message')
  .argument('<messageId>', 'Message id to reply to')
  .option('--comment <text>', 'Optional comment included in the draft')
  .option('--json', 'Output draft message as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
  .action(
    async (
      messageId: string,
      opts: { comment?: string; json?: boolean; token?: string; identity?: string; user?: string },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await createMailReplyDraft(auth.token!, messageId, opts.user, opts.comment);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else console.log(`Draft message id: ${r.data.id}`);
    }
  );

outlookGraphCommand
  .command('create-reply-all')
  .description('Create a reply-all draft (Graph POST .../createReplyAll)')
  .argument('<messageId>', 'Message id')
  .option('--comment <text>', 'Optional comment')
  .option('--json', 'Output draft message as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
  .action(
    async (
      messageId: string,
      opts: { comment?: string; json?: boolean; token?: string; identity?: string; user?: string },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await createMailReplyAllDraft(auth.token!, messageId, opts.user, opts.comment);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else console.log(`Draft message id: ${r.data.id}`);
    }
  );

outlookGraphCommand
  .command('create-forward')
  .description('Create a forward draft (Graph POST .../createForward)')
  .argument('<messageId>', 'Message id')
  .requiredOption('--to <emails>', 'Comma-separated recipient addresses')
  .option('--comment <text>', 'Optional comment')
  .option('--json', 'Output draft message as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
  .action(
    async (
      messageId: string,
      opts: {
        to: string;
        comment?: string;
        json?: boolean;
        token?: string;
        identity?: string;
        user?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const recipients = opts.to
        .split(',')
        .map((s) => s.trim())
        .filter(Boolean);
      if (recipients.length === 0) {
        console.error('Provide at least one address in --to');
        process.exit(1);
      }
      const r = await createMailForwardDraft(auth.token!, messageId, recipients, opts.user, opts.comment);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else console.log(`Draft message id: ${r.data.id}`);
    }
  );

outlookGraphCommand
  .command('send-message')
  .description('Send a draft message (Graph POST /messages/{id}/send)')
  .argument('<messageId>', 'Draft message id')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
  .action(async (messageId: string, opts: { token?: string; identity?: string; user?: string }, cmd: any) => {
    checkReadOnly(cmd);
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await sendMailMessage(auth.token!, messageId, opts.user);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log('Sent.');
  });

outlookGraphCommand
  .command('list-contacts')
  .description('List personal contacts (Graph GET /contacts)')
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
    const r = await listContacts(auth.token!, opts.user);
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

outlookGraphCommand
  .command('get-contact')
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

outlookGraphCommand
  .command('create-contact')
  .description('Create a contact (JSON body per Graph contact resource)')
  .requiredOption('--json-file <path>', 'Path to JSON file')
  .option('-f, --folder <folderId>', 'Create under this contact folder')
  .option('--json', 'Echo created contact as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
  .action(
    async (
      opts: { jsonFile: string; folder?: string; json?: boolean; token?: string; identity?: string; user?: string },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const raw = await readFile(opts.jsonFile, 'utf-8');
      const body = JSON.parse(raw) as Record<string, unknown>;
      const r = await createContact(auth.token!, body, opts.user, opts.folder);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else console.log(`Created contact: ${r.data.displayName ?? r.data.id} (${r.data.id})`);
    }
  );

outlookGraphCommand
  .command('update-contact')
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

outlookGraphCommand
  .command('delete-contact')
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
