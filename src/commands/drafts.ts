import { open, readFile } from 'node:fs/promises';
import { Command } from 'commander';
import { AttachmentLinkSpecError, parseAttachLinkSpec } from '../lib/attach-link-spec.js';
import { AttachmentPathError, validateAttachmentPath } from '../lib/attachments.js';
import { resolveAuth } from '../lib/auth.js';
import {
  addAttachmentToDraft,
  addReferenceAttachmentToDraft,
  createDraft,
  deleteDraftById,
  getEmail,
  getEmails,
  sendDraftById,
  updateDraft
} from '../lib/ews-client.js';
import { getExchangeBackend } from '../lib/exchange-backend.js';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import { markdownToHtml } from '../lib/markdown.js';
import { lookupMimeType } from '../lib/mime-type.js';
import { getMessage, listMessagesInFolder } from '../lib/outlook-graph-client.js';
import { checkReadOnly } from '../lib/utils.js';
import { tryGraphDraftMutations } from './drafts-graph.js';

function formatDate(dateStr: string): string {
  const date = new Date(dateStr);
  const now = new Date();
  const isToday = date.toDateString() === now.toDateString();

  if (isToday) {
    return date.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit', hour12: false });
  } else {
    return date.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
  }
}

function truncate(str: string, maxLen: number): string {
  if (!str) return '';
  str = str.replace(/\s+/g, ' ').trim();
  if (str.length <= maxLen) return str;
  return `${str.substring(0, maxLen - 1)}\u2026`;
}

export const draftsCommand = new Command('drafts')
  .description('Manage email drafts')
  .option('-n, --limit <number>', 'Number of drafts to show', '10')
  .option('-r, --read <id>', 'Read draft by ID')
  .option('--create', 'Create a new draft')
  .option('--edit <id>', 'Edit draft by ID')
  .option('--send <id>', 'Send draft by ID')
  .option('--delete <id>', 'Delete draft by ID')
  .option('--to <emails>', 'Recipient(s) for create/edit, comma-separated')
  .option('--cc <emails>', 'CC recipient(s), comma-separated')
  .option('--subject <text>', 'Subject for create/edit')
  .option('--body <text>', 'Body for create/edit')
  .option('--attach <files>', 'Attach file(s), comma-separated paths')
  .option(
    '--attach-link <spec>',
    'Attach link: "Title|https://url" or bare https URL (repeatable)',
    (v: string, prev: string[]) => [...prev, v],
    [] as string[]
  )
  .option('--markdown', 'Parse body as markdown')
  .option('--html', 'Treat body as HTML')
  .option(
    '--category <name>',
    'Outlook category (repeatable; colors follow mailbox master list)',
    (v: string, prev: string[]) => [...prev, v],
    [] as string[]
  )
  .option('--clear-categories', 'On --edit, remove all categories from the draft')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Use a specific authentication identity (default: default)')
  .option('--mailbox <email>', 'Delegated or shared mailbox drafts folder')
  .action(
    async (
      options: {
        limit: string;
        read?: string;
        create?: boolean;
        edit?: string;
        send?: string;
        delete?: string;
        to?: string;
        cc?: string;
        subject?: string;
        body?: string;
        attach?: string;
        attachLink?: string[];
        markdown?: boolean;
        html?: boolean;
        json?: boolean;
        token?: string;
        identity?: string;
        mailbox?: string;
        category?: string[];
        clearCategories?: boolean;
      },
      cmd: any
    ) => {
      if (options.send || options.delete || options.create || options.edit) {
        checkReadOnly(cmd);
      }

      const backend = getExchangeBackend();
      const needsEwsMutation = !!(options.create || options.edit || options.send || options.delete);
      const graphDraftsListOrRead = !needsEwsMutation;

      if (graphDraftsListOrRead && (backend === 'graph' || backend === 'auto')) {
        const ga = await resolveGraphAuth({ token: options.token, identity: options.identity });
        if (ga.success && ga.token) {
          const user = options.mailbox?.trim() || undefined;

          if (options.read) {
            const id = options.read.trim();
            const select =
              'subject,body,bodyPreview,toRecipients,ccRecipients,categories,lastModifiedDateTime,receivedDateTime';
            const full = await getMessage(ga.token, id, user, select);
            if (!full.ok || !full.data) {
              if (options.json) {
                console.log(JSON.stringify({ error: full.error?.message || 'Failed to fetch draft' }, null, 2));
              } else {
                console.error(`Error: ${full.error?.message || 'Failed to fetch draft'}`);
              }
              process.exit(1);
            }
            const d = full.data;

            if (options.json) {
              console.log(JSON.stringify({ backend: 'graph', draft: d }, null, 2));
              return;
            }

            const toLine =
              d.toRecipients?.map((x) => x.emailAddress?.address).filter(Boolean).join(', ') || '(none)';
            console.log(`\n${'\u2500'.repeat(60)}`);
            console.log(`To: ${toLine}`);
            if (d.ccRecipients?.length) {
              console.log(
                `Cc: ${d.ccRecipients.map((x) => x.emailAddress?.address).filter(Boolean).join(', ') || '(none)'}`
              );
            }
            console.log(`Subject: ${d.subject || '(no subject)'}`);
            if (d.categories?.length) console.log(`Categories: ${d.categories.join(', ')}`);
            console.log(`${'\u2500'.repeat(60)}\n`);
            const content = d.body?.content ?? d.bodyPreview ?? '(no content)';
            console.log(content);
            console.log(`\n${'\u2500'.repeat(60)}\n`);
            return;
          }

          const limit = parseInt(options.limit, 10) || 10;
          const r = await listMessagesInFolder(ga.token, 'drafts', user, {
            top: limit,
            orderby: 'lastModifiedDateTime desc'
          });
          if (!r.ok || !r.data) {
            if (options.json) {
              console.log(JSON.stringify({ error: r.error?.message || 'Failed to fetch drafts' }, null, 2));
            } else {
              console.error(`Error: ${r.error?.message || 'Failed to fetch drafts'}`);
            }
            process.exit(1);
          }
          const graphDrafts = r.data;
          if (options.json) {
            console.log(
              JSON.stringify(
                {
                  backend: 'graph',
                  drafts: graphDrafts.map((d, i) => ({
                    index: i + 1,
                    id: d.id,
                    to: d.toRecipients?.map((x) => x.emailAddress?.address),
                    subject: d.subject,
                    preview: d.bodyPreview,
                    lastModified: d.lastModifiedDateTime || d.receivedDateTime,
                    categories: d.categories
                  }))
                },
                null,
                2
              )
            );
            return;
          }

          console.log(`\n\ud83d\udcdd Drafts (Graph)${options.mailbox ? ` — ${options.mailbox}` : ''}:\n`);
          console.log('\u2500'.repeat(70));
          if (graphDrafts.length === 0) {
            console.log('\n  No drafts found.\n');
            return;
          }
          for (let i = 0; i < graphDrafts.length; i++) {
            const draft = graphDrafts[i];
            const to = draft.toRecipients?.map((x) => x.emailAddress?.address).join(', ') || '(no recipient)';
            const subject = draft.subject || '(no subject)';
            const when = draft.lastModifiedDateTime || draft.receivedDateTime;
            const date = when ? formatDate(when) : '';
            console.log(
              `  [${(i + 1).toString().padStart(2)}] ${truncate(to, 25).padEnd(25)} ${truncate(subject, 32).padEnd(32)} ${date}`
            );
            console.log(`       ID: ${draft.id}`);
            if (draft.categories?.length) console.log(`       Categories: ${draft.categories.join(', ')}`);
          }
          console.log(`\n${'\u2500'.repeat(70)}`);
          console.log('\nCommands:');
          console.log('  m365-agent-cli drafts -r <id>                  # Read draft by id');
          console.log('  m365-agent-cli drafts --create --to "..." ...   # Graph when backend=graph|auto');
          console.log('');
          return;
        }
        if (backend === 'graph') {
          console.error('Error: Graph authentication failed. Set EWS_CLIENT_ID and GRAPH_REFRESH_TOKEN.');
          process.exit(1);
        }
      }

      if (needsEwsMutation && (backend === 'graph' || backend === 'auto')) {
        const ga = await resolveGraphAuth({ token: options.token, identity: options.identity });
        if (ga.success && ga.token) {
          const user = options.mailbox?.trim() || undefined;
          const graphDone = await tryGraphDraftMutations(ga.token, user, options, backend);
          if (graphDone) return;
        }
        if (backend === 'graph') {
          console.error('Error: Graph authentication failed. Set EWS_CLIENT_ID and GRAPH_REFRESH_TOKEN.');
          process.exit(1);
        }
      }

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

      const limit = parseInt(options.limit, 10) || 10;

      // Get drafts for listing
      const draftsResult = await getEmails({
        token: authResult.token!,
        folder: 'drafts',
        mailbox: options.mailbox,
        top: limit
      });

      if (!draftsResult.ok || !draftsResult.data) {
        if (options.json) {
          console.log(JSON.stringify({ error: draftsResult.error?.message || 'Failed to fetch drafts' }, null, 2));
        } else {
          console.error(`Error: ${draftsResult.error?.message || 'Failed to fetch drafts'}`);
        }
        process.exit(1);
      }

      const drafts = draftsResult.data.value;

      // Handle create
      if (options.create) {
        const toList = options.to
          ? options.to
              .split(',')
              .map((e) => e.trim())
              .filter(Boolean)
          : undefined;
        const ccList = options.cc
          ? options.cc
              .split(',')
              .map((e) => e.trim())
              .filter(Boolean)
          : undefined;

        let body = options.body;
        if (body) body = body.replace(/\\n/g, '\n');
        let bodyType: 'Text' | 'HTML' = 'Text';
        if (options.html && body) {
          const escaped = body
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/\n/g, '<br>');
          body = body.match(/<\w+[^>]*>/) ? body : escaped;
          bodyType = 'HTML';
        } else if (options.markdown && body) {
          body = markdownToHtml(body);
          bodyType = 'HTML';
        }

        const cats = (options.category ?? []).map((c) => c.trim()).filter(Boolean);
        const result = await createDraft(authResult.token!, {
          to: toList,
          cc: ccList,
          subject: options.subject,
          body,
          bodyType,
          mailbox: options.mailbox,
          categories: cats.length ? cats : undefined
        });

        if (!result.ok || !result.data) {
          console.error(`Error: ${result.error?.message || 'Failed to create draft'}`);
          process.exit(1);
        }

        // Add attachments if specified
        const workingDirectory = process.cwd();
        if (options.attach) {
          const filePaths = options.attach
            .split(',')
            .map((f) => f.trim())
            .filter(Boolean);
          for (const filePath of filePaths) {
            try {
              const validated = await validateAttachmentPath(filePath, workingDirectory);
              const fh = await open(validated.absolutePath, 'r');
              let content: Buffer;
              try {
                const st = await fh.stat();
                if (!st.isFile()) {
                  console.error(`Not a file: ${validated.absolutePath}`);
                  process.exit(1);
                }
                if (st.size > 25 * 1024 * 1024) {
                  console.error(`File too large (>25MB): ${validated.absolutePath}`);
                  process.exit(1);
                }
                content = await fh.readFile();
              } finally {
                await fh.close();
              }
              const contentType = lookupMimeType(validated.fileName) || 'application/octet-stream';

              const attachResult = await addAttachmentToDraft(
                authResult.token!,
                result.data.Id,
                {
                  name: validated.fileName,
                  contentType,
                  contentBytes: content.toString('base64')
                },
                options.mailbox
              );

              if (!attachResult.ok) {
                console.error(`Failed to attach ${validated.fileName}: ${attachResult.error?.message}`);
              } else if (!options.json) {
                console.log(`  Attached: ${validated.fileName}`);
              }
            } catch (err) {
              if (err instanceof AttachmentPathError) {
                console.error(`Invalid attachment path: ${filePath}: ${err.message}`);
              } else {
                console.error(`Failed to attach: ${filePath}`);
              }
              process.exit(1);
            }
          }
        }

        const linkSpecsCreate = options.attachLink ?? [];
        for (const spec of linkSpecsCreate) {
          try {
            const { name, url } = parseAttachLinkSpec(spec);
            const linkRes = await addReferenceAttachmentToDraft(
              authResult.token!,
              result.data.Id,
              { name, url, contentType: 'text/html' },
              options.mailbox
            );
            if (!linkRes.ok) {
              console.error(`Failed to attach link ${name}: ${linkRes.error?.message}`);
              process.exit(1);
            }
            if (!options.json) {
              console.log(`  Attached link: ${name}`);
            }
          } catch (err) {
            const msg =
              err instanceof AttachmentLinkSpecError ? err.message : err instanceof Error ? err.message : String(err);
            console.error(`Invalid --attach-link: ${msg}`);
            process.exit(1);
          }
        }

        if (options.json) {
          console.log(JSON.stringify({ success: true, draftId: result.data.Id }, null, 2));
        } else {
          console.log(`\n\u2713 Draft created`);
          if (options.subject) console.log(`  Subject: ${options.subject}`);
          if (toList) console.log(`  To: ${toList.join(', ')}`);
          console.log();
        }

        return;
      }

      // Handle read
      if (options.read) {
        const id = options.read.trim();
        const fullDraft = await getEmail(authResult.token!, id, options.mailbox);

        if (!fullDraft.ok || !fullDraft.data) {
          console.error(`Error: ${fullDraft.error?.message || 'Failed to fetch draft'}`);
          process.exit(1);
        }

        const d = fullDraft.data;

        if (options.json) {
          console.log(JSON.stringify(d, null, 2));
          return;
        }

        console.log(`\n${'\u2500'.repeat(60)}`);
        console.log(`To: ${d.ToRecipients?.map((r) => r.EmailAddress?.Address).join(', ') || '(none)'}`);
        console.log(`Subject: ${d.Subject || '(no subject)'}`);
        if (d.Categories?.length) console.log(`Categories: ${d.Categories.join(', ')}`);
        console.log(`${'\u2500'.repeat(60)}\n`);
        console.log(d.Body?.Content || d.BodyPreview || '(no content)');
        console.log(`\n${'\u2500'.repeat(60)}\n`);
        return;
      }

      // Handle edit
      if (options.edit) {
        const id = options.edit.trim();
        const toList = options.to
          ? options.to
              .split(',')
              .map((e) => e.trim())
              .filter(Boolean)
          : undefined;
        const ccList = options.cc
          ? options.cc
              .split(',')
              .map((e) => e.trim())
              .filter(Boolean)
          : undefined;

        let body = options.body;
        if (body) body = body.replace(/\\n/g, '\n');
        let bodyType: 'Text' | 'HTML' = 'Text';
        if (options.html && body) {
          const escaped = body
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/\n/g, '<br>');
          body = body.match(/<\w+[^>]*>/) ? body : escaped;
          bodyType = 'HTML';
        } else if (options.markdown && body) {
          body = markdownToHtml(body);
          bodyType = 'HTML';
        }

        const cats = (options.category ?? []).map((c) => c.trim()).filter(Boolean);
        const result = await updateDraft(authResult.token!, id, {
          to: toList,
          cc: ccList,
          subject: options.subject,
          body,
          bodyType,
          mailbox: options.mailbox,
          categories: cats.length ? cats : undefined,
          clearCategories: options.clearCategories
        });

        if (!result.ok) {
          console.error(`Error: ${result.error?.message || 'Failed to update draft'}`);
          process.exit(1);
        }

        // Add attachments if specified
        const workingDirectory = process.cwd();
        if (options.attach) {
          const filePaths = options.attach
            .split(',')
            .map((f) => f.trim())
            .filter(Boolean);
          for (const filePath of filePaths) {
            try {
              const validated = await validateAttachmentPath(filePath, workingDirectory);
              const content = await readFile(validated.absolutePath);
              const contentType = lookupMimeType(validated.fileName) || 'application/octet-stream';

              await addAttachmentToDraft(
                authResult.token!,
                id,
                {
                  name: validated.fileName,
                  contentType,
                  contentBytes: content.toString('base64')
                },
                options.mailbox
              );

              if (!options.json) {
                console.log(`  Attached: ${validated.fileName}`);
              }
            } catch (err) {
              if (err instanceof AttachmentPathError) {
                console.error(`Invalid attachment path: ${filePath}: ${err.message}`);
              } else {
                console.error(`Failed to attach: ${filePath}`);
              }
              process.exit(1);
            }
          }
        }

        const linkSpecsEdit = options.attachLink ?? [];
        for (const spec of linkSpecsEdit) {
          try {
            const { name, url } = parseAttachLinkSpec(spec);
            const linkRes = await addReferenceAttachmentToDraft(
              authResult.token!,
              id,
              { name, url, contentType: 'text/html' },
              options.mailbox
            );
            if (!linkRes.ok) {
              console.error(`Failed to attach link ${name}: ${linkRes.error?.message}`);
              process.exit(1);
            }
            if (!options.json) {
              console.log(`  Attached link: ${name}`);
            }
          } catch (err) {
            const msg =
              err instanceof AttachmentLinkSpecError ? err.message : err instanceof Error ? err.message : String(err);
            console.error(`Invalid --attach-link: ${msg}`);
            process.exit(1);
          }
        }

        console.log(`\u2713 Draft updated: ${id}`);
        return;
      }

      // Handle send
      if (options.send) {
        const id = options.send.trim();
        const result = await sendDraftById(authResult.token!, id, options.mailbox);

        if (!result.ok) {
          console.error(`Error: ${result.error?.message || 'Failed to send draft'}`);
          process.exit(1);
        }

        console.log(`\u2713 Draft sent: ${id}`);
        return;
      }

      // Handle delete
      if (options.delete) {
        const id = options.delete.trim();
        if (!id) {
          console.error('Error: --delete requires a draft ID');
          process.exit(1);
        }
        const result = await deleteDraftById(authResult.token!, id, options.mailbox);

        if (!result.ok) {
          console.error(`Error: ${result.error?.message || 'Failed to delete draft'}`);
          process.exit(1);
        }

        console.log(`\u2713 Draft deleted: ${id}`);
        return;
      }

      // List drafts
      if (options.json) {
        console.log(
          JSON.stringify(
            {
              drafts: drafts.map((d, i) => ({
                index: i + 1,
                id: d.Id,
                to: d.ToRecipients?.map((r) => r.EmailAddress?.Address),
                subject: d.Subject,
                preview: d.BodyPreview,
                lastModified: d.ReceivedDateTime,
                categories: d.Categories
              }))
            },
            null,
            2
          )
        );
        return;
      }

      console.log(`\n\ud83d\udcdd Drafts${options.mailbox ? ` — ${options.mailbox}` : ''}:\n`);
      console.log('\u2500'.repeat(70));

      if (drafts.length === 0) {
        console.log('\n  No drafts found.\n');
        return;
      }

      for (let i = 0; i < drafts.length; i++) {
        const draft = drafts[i];
        const to = draft.ToRecipients?.map((r) => r.EmailAddress?.Address).join(', ') || '(no recipient)';
        const subject = draft.Subject || '(no subject)';
        const date = draft.ReceivedDateTime ? formatDate(draft.ReceivedDateTime) : '';

        console.log(
          `  [${(i + 1).toString().padStart(2)}] ${truncate(to, 25).padEnd(25)} ${truncate(subject, 32).padEnd(32)} ${date}`
        );
        console.log(`       ID: ${draft.Id}`);
        if (draft.Categories?.length) console.log(`       Categories: ${draft.Categories.join(', ')}`);
      }

      console.log(`\n${'\u2500'.repeat(70)}`);
      console.log('\nCommands:');
      console.log('  m365-agent-cli drafts -r <id>                  # Read draft');
      console.log('  m365-agent-cli drafts --create --to "..." --subject "..." --body "..."');
      console.log('  m365-agent-cli drafts --edit <id> --body "new text"');
      console.log('  m365-agent-cli drafts --send <id>              # Send draft');
      console.log('  m365-agent-cli drafts --delete <id>            # Delete draft');
      console.log();
    }
  );
