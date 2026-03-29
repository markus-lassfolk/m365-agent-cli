import { access, mkdir, writeFile } from 'node:fs/promises';
import { extname, join } from 'node:path';
import { Command } from 'commander';
import { resolveAuth } from '../lib/auth.js';
import {
  forwardEmail,
  getAttachment,
  getAttachments,
  getEmail,
  getEmails,
  getMailFolders,
  moveEmail,
  replyToEmail,
  replyToEmailDraft,
  SENSITIVITY_MAP,
  updateEmail
} from '../lib/ews-client.js';
import { markdownToHtml } from '../lib/markdown.js';

function formatDate(dateStr: string): string {
  const date = new Date(dateStr);
  const now = new Date();
  const isToday = date.toDateString() === now.toDateString();
  const yesterday = new Date(now);
  yesterday.setDate(yesterday.getDate() - 1);
  const isYesterday = date.toDateString() === yesterday.toDateString();

  if (isToday) {
    return date.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit', hour12: false });
  } else if (isYesterday) {
    return 'Yesterday';
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

export const mailCommand = new Command('mail')
  .description('List and read emails')
  .argument('[folder]', 'Folder: inbox, sent, drafts, deleted, archive, junk', 'inbox')
  .option('-n, --limit <number>', 'Number of emails to show', '10')
  .option('-p, --page <number>', 'Page number (1-based)', '1')
  .option('--unread', 'Show only unread emails')
  .option('--flagged', 'Show only flagged emails')
  .option('-s, --search <query>', 'Search emails (subject, body, sender)')
  .option('-r, --read <id>', 'Read email by ID')
  .option('-d, --download <id>', 'Download attachments from email by ID')
  .option('-o, --output <dir>', 'Output directory for attachments', '.')
  .option('--mark-read <id>', 'Mark email as read (by ID)')
  .option('--mark-unread <id>', 'Mark email as unread (by ID)')
  .option('--flag <id>', 'Flag email (by ID)')
  .option('--start-date <date>', 'Start date for flag (YYYY-MM-DD)')
  .option('--due <date>', 'Due date for flag (YYYY-MM-DD)')
  .option('--unflag <id>', 'Remove flag (by ID)')
  .option('--complete <id>', 'Mark flagged email as complete (by ID)')
  .option('--sensitivity <id>', 'Set sensitivity on email by ID (use with --level)')
  .option('--level <level>', 'Sensitivity level: normal, personal, private, confidential')
  .option('--move <id>', 'Move email to folder (use with --to)')
  .option('--to <folder>', 'Destination folder for move (inbox, archive, deleted, junk)')
  .option('--reply <id>', 'Reply to email by ID')
  .option('--reply-all <id>', 'Reply all to email by ID')
  .option('--draft', 'Create a reply draft (do not send)')
  .option('--forward <id>', 'Forward email by ID (use with --to-addr)')
  .option('--to-addr <emails>', 'Forward recipients (comma-separated)')
  .option('--message <text>', 'Reply/forward message text')
  .option('--markdown', 'Parse message as markdown (bold, links, lists)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--mailbox <email>', 'Shared mailbox for reply/forward (routes via X-AnchorMailbox)')
  .option('--identity <name>', 'Use a specific authentication identity (default: default)')
  .action(
    async (
      folder: string,
      options: {
        limit: string;
        page: string;
        unread?: boolean;
        flagged?: boolean;
        search?: string;
        read?: string;
        download?: string;
        output: string;
        force?: boolean;
        markRead?: string;
        markUnread?: string;
        flag?: string;
        startDate?: string;
        due?: string;
        unflag?: string;
        complete?: string;
        sensitivity?: string;
        level?: string;
        move?: string;
        to?: string;
        reply?: string;
        replyAll?: string;
        forward?: string;
        toAddr?: string;
        message?: string;
        markdown?: boolean;
        json?: boolean;
        token?: string;
        draft?: boolean;
        mailbox?: string;
        identity?: string;
      }
    ) => {
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

      // Map folder names to API folder IDs
      const folderMap: Record<string, string> = {
        inbox: 'inbox',
        sent: 'sentitems',
        sentitems: 'sentitems',
        drafts: 'drafts',
        deleted: 'deleteditems',
        deleteditems: 'deleteditems',
        trash: 'deleteditems',
        archive: 'archive',
        junk: 'junkemail',
        junkemail: 'junkemail',
        spam: 'junkemail'
      };

      let apiFolder = folderMap[folder.toLowerCase()];

      // If not a well-known folder, look up by name
      if (!apiFolder) {
        const foldersResult = await getMailFolders(authResult.token!);
        if (foldersResult.ok && foldersResult.data) {
          const found = foldersResult.data.value.find((f) => f.DisplayName.toLowerCase() === folder.toLowerCase());
          if (found) {
            apiFolder = found.Id;
          } else {
            console.error(`Folder "${folder}" not found.`);
            console.error('Use "clippy folders" to see available folders.');
            process.exit(1);
          }
        } else {
          apiFolder = folder; // Fallback to using the name directly
        }
      }

      const limit = parseInt(options.limit, 10) || 10;
      const page = parseInt(options.page, 10) || 1;
      const skip = (page - 1) * limit;

      const result = await getEmails({
        token: authResult.token!,
        folder: apiFolder,
        top: limit,
        skip,
        search: options.search,
        isRead: options.unread ? false : undefined,
        flagStatus: options.flagged ? 'Flagged' : undefined
      });

      if (!result.ok || !result.data) {
        if (options.json) {
          console.log(JSON.stringify({ error: result.error?.message || 'Failed to fetch emails' }, null, 2));
        } else {
          console.error(`Error: ${result.error?.message || 'Failed to fetch emails'}`);
        }
        process.exit(1);
      }

      const emails = result.data.value;

      // Handle reading a specific email
      if (options.read) {
        const id = options.read.trim();
        const fullEmail = await getEmail(authResult.token!, id);

        if (!fullEmail.ok || !fullEmail.data) {
          console.error(`Error: ${fullEmail.error?.message || 'Failed to fetch email'}`);
          process.exit(1);
        }

        const email = fullEmail.data;

        if (options.json) {
          console.log(JSON.stringify(email, null, 2));
          return;
        }

        console.log(`\n${'\u2500'.repeat(60)}`);
        console.log(`From: ${email.From?.EmailAddress?.Name || email.From?.EmailAddress?.Address || 'Unknown'}`);
        if (email.From?.EmailAddress?.Address) {
          console.log(`      <${email.From.EmailAddress.Address}>`);
        }
        console.log(`Subject: ${email.Subject || '(no subject)'}`);
        console.log(`Date: ${email.ReceivedDateTime ? new Date(email.ReceivedDateTime).toLocaleString() : 'Unknown'}`);

        if (email.ToRecipients && email.ToRecipients.length > 0) {
          const to = email.ToRecipients.map((r) => r.EmailAddress?.Address)
            .filter(Boolean)
            .join(', ');
          console.log(`To: ${to}`);
        }

        if (email.CcRecipients && email.CcRecipients.length > 0) {
          const cc = email.CcRecipients.map((r) => r.EmailAddress?.Address)
            .filter(Boolean)
            .join(', ');
          console.log(`Cc: ${cc}`);
        }

        if (email.HasAttachments) {
          const attachmentsResult = await getAttachments(authResult.token!, email.Id);
          if (attachmentsResult.ok && attachmentsResult.data) {
            const atts = attachmentsResult.data.value.filter((a) => !a.IsInline);
            if (atts.length > 0) {
              console.log('Attachments:');
              for (const att of atts) {
                const sizeKB = Math.round(att.Size / 1024);
                console.log(`  - ${att.Name} (${sizeKB} KB)`);
              }
            }
          }
        }

        console.log(`${'\u2500'.repeat(60)}\n`);
        console.log(email.Body?.Content || email.BodyPreview || '(no content)');
        console.log(`\n${'\u2500'.repeat(60)}\n`);
        return;
      }

      // Handle downloading attachments
      if (options.download) {
        const id = options.download.trim();
        const emailSummary = await getEmail(authResult.token!, id);
        if (!emailSummary.ok || !emailSummary.data) {
          console.error(`Error: ${emailSummary.error?.message || 'Failed to fetch email'}`);
          process.exit(1);
        }

        if (!emailSummary.data.HasAttachments) {
          console.log('This email has no attachments.');
          return;
        }

        const attachmentsResult = await getAttachments(authResult.token!, emailSummary.data.Id);
        if (!attachmentsResult.ok || !attachmentsResult.data) {
          console.error(`Error: ${attachmentsResult.error?.message || 'Failed to fetch attachments'}`);
          process.exit(1);
        }

        const attachments = attachmentsResult.data.value.filter((a) => !a.IsInline);

        if (attachments.length === 0) {
          console.log('This email has no downloadable attachments.');
          return;
        }

        // Ensure output directory exists
        await mkdir(options.output, { recursive: true });

        console.log(`\nDownloading ${attachments.length} attachment(s) to ${options.output}/\n`);

        const usedPaths = new Set<string>();

        for (const att of attachments) {
          // Get full attachment with content
          const fullAtt = await getAttachment(authResult.token!, emailSummary.data.Id, att.Id);
          if (!fullAtt.ok || !fullAtt.data?.ContentBytes) {
            console.error(`  Failed to download: ${att.Name}`);
            continue;
          }

          const content = Buffer.from(fullAtt.data.ContentBytes, 'base64');

          // Resolve the actual file path, avoiding collisions and existing files
          let filePath = join(options.output, att.Name);
          let counter = 1;
          while (true) {
            // Always check for intra-download collisions
            if (usedPaths.has(filePath)) {
              const ext = extname(att.Name);
              const base = att.Name.slice(0, att.Name.length - ext.length);
              filePath = join(options.output, `${base} (${counter})${ext}`);
              counter++;
              continue;
            }

            // Check for pre-existing files only if --force is not set
            if (!options.force) {
              try {
                await access(filePath);
                // File exists — resolve collision with a numeric suffix
                const ext = extname(att.Name);
                const base = att.Name.slice(0, att.Name.length - ext.length);
                filePath = join(options.output, `${base} (${counter})${ext}`);
                counter++;
                continue;
              } catch {
                // File doesn't exist — safe to use
              }
            }

            // Path is safe to use
            break;
          }

          usedPaths.add(filePath);
          await writeFile(filePath, content);

          const sizeKB = Math.round(content.length / 1024);
          const written = filePath === join(options.output, att.Name) ? att.Name : filePath.split(/[\\/]/).pop();
          console.log(`  \u2713 ${written} (${sizeKB} KB)`);
        }

        console.log('\nDone.\n');
        return;
      }

      // Handle mark as read/unread
      if (options.markRead || options.markUnread) {
        const id = (options.markRead || options.markUnread)?.trim();
        if (!id) {
          console.error('Error: --mark-read/--mark-unread requires a message ID');
          process.exit(1);
        }
        const isRead = !!options.markRead;

        const result = await updateEmail(authResult.token!, id, { IsRead: isRead });

        if (!result.ok) {
          console.error(`Error: ${result.error?.message || 'Failed to update email'}`);
          process.exit(1);
        }

        console.log(`\u2713 Marked as ${isRead ? 'read' : 'unread'}: ${id}`);
        return;
      }

      // Handle flag/unflag/complete
      if (options.flag || options.unflag || options.complete) {
        const id = (options.flag || options.unflag || options.complete)?.trim();
        let flagStatus: 'NotFlagged' | 'Flagged' | 'Complete';
        let actionLabel: string;
        let startDate: { DateTime: string; TimeZone: string } | undefined;
        let dueDate: { DateTime: string; TimeZone: string } | undefined;

        if (options.flag) {
          flagStatus = 'Flagged';
          actionLabel = 'Flagged';

          if (options.startDate) {
            const parsedStartDate = new Date(options.startDate);
            if (Number.isNaN(parsedStartDate.getTime())) {
              console.error('Error: Invalid start date. Please provide a valid ISO 8601 date/time value.');
              process.exit(1);
            }
            startDate = { DateTime: parsedStartDate.toISOString(), TimeZone: 'UTC' };
          }
          if (options.due) {
            const parsedDueDate = new Date(options.due);
            if (Number.isNaN(parsedDueDate.getTime())) {
              console.error('Error: Invalid due date. Please provide a valid ISO 8601 date/time value.');
              process.exit(1);
            }
            dueDate = { DateTime: parsedDueDate.toISOString(), TimeZone: 'UTC' };
          }
        } else if (options.complete) {
          flagStatus = 'Complete';
          actionLabel = 'Marked complete';
        } else {
          flagStatus = 'NotFlagged';
          actionLabel = 'Unflagged';
        }

        if (!id) {
          console.error('Error: --flag/--unflag/--complete requires a message ID');
          process.exit(1);
        }
        const result = await updateEmail(authResult.token!, id, {
          Flag: { FlagStatus: flagStatus, StartDate: startDate, DueDate: dueDate }
        });

        if (!result.ok) {
          console.error(`Error: ${result.error?.message || 'Failed to update email'}`);
          process.exit(1);
        }

        console.log(`\u2713 ${actionLabel}: ${id}`);
        return;
      }

      // Handle sensitivity
      if (options.sensitivity) {
        const id = options.sensitivity.trim();

        if (!options.level) {
          console.error('Error: --sensitivity requires --level to be specified');
          console.error('Example: clippy mail --sensitivity <id> --level personal');
          console.error('Levels: normal, personal, private, confidential');
          process.exit(1);
        }

        const sensitivity = SENSITIVITY_MAP[options.level.toLowerCase()];

        if (!sensitivity) {
          console.error(`Invalid sensitivity level: ${options.level}`);
          console.error('Valid levels: normal, personal, private, confidential');
          process.exit(1);
        }

        const result = await updateEmail(authResult.token!, id, {
          Sensitivity: sensitivity
        });

        if (!result.ok) {
          console.error(`Error: ${result.error?.message || 'Failed to update email sensitivity'}`);
          process.exit(1);
        }

        console.log(`\u2713 Sensitivity set to ${sensitivity}: ${id}`);
        return;
      }

      // Handle move
      if (options.move) {
        if (!options.to) {
          console.error('Please specify destination folder with --to');
          console.error('Example: clippy mail --move <id> --to archive');
          console.error('Folders: inbox, archive, deleted, junk, drafts, sent');
          process.exit(1);
        }

        const id = options.move.trim();

        // Map folder names to API folder IDs
        const destFolderMap: Record<string, string> = {
          inbox: 'inbox',
          archive: 'archive',
          deleted: 'deleteditems',
          deleteditems: 'deleteditems',
          trash: 'deleteditems',
          junk: 'junkemail',
          junkemail: 'junkemail',
          spam: 'junkemail',
          drafts: 'drafts',
          sent: 'sentitems',
          sentitems: 'sentitems'
        };

        let destFolder = destFolderMap[options.to.toLowerCase()];

        // If not a well-known folder, look up by name
        if (!destFolder) {
          const foldersResult = await getMailFolders(authResult.token!);
          if (foldersResult.ok && foldersResult.data) {
            const found = foldersResult.data.value.find(
              (f) => f.DisplayName.toLowerCase() === options.to?.toLowerCase()
            );
            if (found) {
              destFolder = found.Id;
            } else {
              console.error(`Folder "${options.to}" not found.`);
              console.error('Use "clippy folders" to see available folders.');
              process.exit(1);
            }
          } else {
            console.error('Failed to look up folder.');
            process.exit(1);
          }
        }

        const result = await moveEmail(authResult.token!, id, destFolder);

        if (!result.ok) {
          console.error(`Error: ${result.error?.message || 'Failed to move email'}`);
          process.exit(1);
        }

        const folderDisplay = options.to.charAt(0).toUpperCase() + options.to.slice(1);
        console.log(`\u2713 Moved to ${folderDisplay}: ${id}`);
        return;
      }

      // Handle reply
      if (options.reply || options.replyAll) {
        const id = (options.reply || options.replyAll)?.trim();

        if (!id) {
          console.error('Error: --reply/--reply-all requires a message ID');
          process.exit(1);
        }

        if (!options.message) {
          console.error('Please provide reply text with --message');
          console.error('Example: clippy mail --reply <id> --message "Thanks for your email!"');
          process.exit(1);
        }

        const isReplyAll = !!options.replyAll;

        let message = options.message;
        let isHtml = false;

        if (options.markdown) {
          message = markdownToHtml(options.message);
          isHtml = true;
        }

        if (options.draft) {
          const result = await replyToEmailDraft(authResult.token!, id, message, isReplyAll, isHtml, options.mailbox);

          if (!result.ok || !result.data) {
            console.error(`Error: ${result.error?.message || 'Failed to create reply draft'}`);
            process.exit(1);
          }

          const replyType = isReplyAll ? 'Reply all' : 'Reply';
          console.log(`\u2713 ${replyType} draft created: ${result.data.draftId}`);
          return;
        }

        const result = await replyToEmail(authResult.token!, id, message, isReplyAll, isHtml, options.mailbox);

        if (!result.ok) {
          console.error(`Error: ${result.error?.message || 'Failed to send reply'}`);
          process.exit(1);
        }

        const replyType = isReplyAll ? 'Reply all' : 'Reply';
        console.log(`\u2713 ${replyType} sent to: ${id}`);
        return;
      }

      // Handle forward
      if (options.forward) {
        const id = options.forward.trim();

        if (!options.toAddr) {
          console.error('Please provide forward recipients with --to-addr');
          console.error('Example: clippy mail --forward <id> --to-addr "user@example.com"');
          process.exit(1);
        }

        const recipients = options.toAddr
          .split(',')
          .map((e) => e.trim())
          .filter(Boolean);

        if (!id) {
          console.error('Error: --forward requires a message ID');
          process.exit(1);
        }
        const result = await forwardEmail(authResult.token!, id, recipients, options.message, options.mailbox);

        if (!result.ok) {
          console.error(`Error: ${result.error?.message || 'Failed to forward email'}`);
          process.exit(1);
        }

        console.log(`\u2713 Forwarded to ${recipients.join(', ')}: ${id}`);
        return;
      }

      // List emails
      if (options.json) {
        console.log(
          JSON.stringify(
            {
              folder: apiFolder,
              page,
              limit,
              emails: emails.map((e, i) => ({
                index: skip + i + 1,
                id: e.Id,
                from: e.From?.EmailAddress?.Address,
                fromName: e.From?.EmailAddress?.Name,
                subject: e.Subject,
                preview: e.BodyPreview,
                receivedAt: e.ReceivedDateTime,
                isRead: e.IsRead,
                hasAttachments: e.HasAttachments,
                importance: e.Importance,
                flagged: e.Flag?.FlagStatus === 'Flagged'
              }))
            },
            null,
            2
          )
        );
        return;
      }

      const folderDisplay = folder.charAt(0).toUpperCase() + folder.slice(1);
      const searchInfo = options.search ? ` - search: "${options.search}"` : '';
      const pageInfo = page > 1 ? ` (page ${page})` : '';
      console.log(`\n\ud83d\udcec ${folderDisplay}${searchInfo}${pageInfo}:\n`);
      console.log('\u2500'.repeat(70));

      if (emails.length === 0) {
        console.log('\n  No emails found.\n');
        return;
      }

      for (let i = 0; i < emails.length; i++) {
        const email = emails[i];
        const idx = skip + i + 1;
        const unreadMark = email.IsRead ? ' ' : '\u2022';
        const flagMark = email.Flag?.FlagStatus === 'Flagged' ? '\u2691' : ' ';
        const attachMark = email.HasAttachments ? '\ud83d\udcce' : ' ';
        const importanceMark = email.Importance === 'High' ? '!' : ' ';

        const from = email.From?.EmailAddress?.Name || email.From?.EmailAddress?.Address || 'Unknown';
        const subject = email.Subject || '(no subject)';
        const date = email.ReceivedDateTime ? formatDate(email.ReceivedDateTime) : '';

        // Format: [idx] marks | from | subject | date
        const marks = `${unreadMark}${flagMark}${attachMark}${importanceMark}`;
        const fromTrunc = truncate(from, 20);
        const subjectTrunc = truncate(subject, 35);

        console.log(
          `  [${idx.toString().padStart(2)}] ${marks} ${fromTrunc.padEnd(20)} ${subjectTrunc.padEnd(35)} ${date}`
        );
        console.log(`       ID: ${email.Id}`);
      }

      console.log(`\n${'\u2500'.repeat(70)}`);
      console.log('\nCommands:');
      console.log(`  clippy mail -r <id>               # Read email`);
      console.log(`  clippy mail -p ${page + 1}                   # Next page`);
      console.log(`  clippy mail --unread              # Only unread`);
      console.log(`  clippy mail -s "keyword"          # Search emails`);
      console.log(`  clippy mail sent                  # Sent folder`);
      console.log('');
    }
  );
