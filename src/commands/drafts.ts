import { readFile, stat } from 'node:fs/promises';
import { Command } from 'commander';
import { AttachmentPathError, validateAttachmentPath } from '../lib/attachments.js';
import { resolveAuth } from '../lib/auth.js';
import {
  addAttachmentToDraft,
  createDraft,
  deleteDraftById,
  getEmail,
  getEmails,
  sendDraftById,
  updateDraft
} from '../lib/ews-client.js';
import { markdownToHtml } from '../lib/markdown.js';
import { lookupMimeType } from '../lib/mime-type.js';

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
  .option('--markdown', 'Parse body as markdown')
  .option('--html', 'Treat body as HTML')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Use a specific authentication identity (default: default)')
  .action(
    async (options: {
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
      markdown?: boolean;
      html?: boolean;
      json?: boolean;
      token?: string;
      identity?: string;
    }) => {
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

        const result = await createDraft(authResult.token!, {
          to: toList,
          cc: ccList,
          subject: options.subject,
          body,
          bodyType
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
              const fileStat = await stat(validated.absolutePath);
              if (fileStat.size > 25 * 1024 * 1024) {
                console.error(`File too large (>25MB): ${validated.absolutePath}`);
                process.exit(1);
              }
              const content = await readFile(validated.absolutePath);
              const contentType = lookupMimeType(validated.fileName) || 'application/octet-stream';

              const attachResult = await addAttachmentToDraft(authResult.token!, result.data.Id, {
                name: validated.fileName,
                contentType,
                contentBytes: content.toString('base64')
              });

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
        const fullDraft = await getEmail(authResult.token!, id);

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

        const result = await updateDraft(authResult.token!, id, {
          to: toList,
          cc: ccList,
          subject: options.subject,
          body,
          bodyType
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

              await addAttachmentToDraft(authResult.token!, id, {
                name: validated.fileName,
                contentType,
                contentBytes: content.toString('base64')
              });

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

        console.log(`\u2713 Draft updated: ${id}`);
        return;
      }

      // Handle send
      if (options.send) {
        const id = options.send.trim();
        const result = await sendDraftById(authResult.token!, id);

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
        const result = await deleteDraftById(authResult.token!, id);

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
                lastModified: d.ReceivedDateTime
              }))
            },
            null,
            2
          )
        );
        return;
      }

      console.log('\n\ud83d\udcdd Drafts:\n');
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
      }

      console.log(`\n${'\u2500'.repeat(70)}`);
      console.log('\nCommands:');
      console.log('  clippy drafts -r <id>                  # Read draft');
      console.log('  clippy drafts --create --to "..." --subject "..." --body "..."');
      console.log('  clippy drafts --edit <id> --body "new text"');
      console.log('  clippy drafts --send <id>              # Send draft');
      console.log('  clippy drafts --delete <id>            # Delete draft');
      console.log();
    }
  );
