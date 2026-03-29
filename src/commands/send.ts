import { readFile } from 'node:fs/promises';
import { Command } from 'commander';
import { AttachmentPathError, validateAttachmentPath } from '../lib/attachments.js';
import { resolveAuth } from '../lib/auth.js';
import { type EmailAttachment, sendEmail } from '../lib/ews-client.js';
import { markdownToHtml } from '../lib/markdown.js';
import { lookupMimeType } from '../lib/mime-type.js';

export const sendCommand = new Command('send')
  .description('Send an email')
  .requiredOption('--to <emails>', 'Recipient email(s), comma-separated')
  .requiredOption('--subject <text>', 'Email subject')
  .option('--body <text>', 'Email body', '')
  .option('--cc <emails>', 'CC recipient(s), comma-separated')
  .option('--bcc <emails>', 'BCC recipient(s), comma-separated')
  .option('--attach <files>', 'Attach file(s), comma-separated paths')
  .option('--html', 'Send body as HTML')
  .option('--markdown', 'Parse body as markdown (bold, links, lists)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--mailbox <email>', 'Send from shared mailbox (Send As)')
  .action(
    async (options: {
      to: string;
      subject: string;
      body?: string;
      cc?: string;
      bcc?: string;
      attach?: string;
      html?: boolean;
      markdown?: boolean;
      json?: boolean;
      token?: string;
      mailbox?: string;
    }) => {
      const authResult = await resolveAuth({
        token: options.token
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

      const toList = options.to
        .split(',')
        .map((e) => e.trim())
        .filter(Boolean);
      const ccList = options.cc
        ? options.cc
            .split(',')
            .map((e) => e.trim())
            .filter(Boolean)
        : undefined;
      const bccList = options.bcc
        ? options.bcc
            .split(',')
            .map((e) => e.trim())
            .filter(Boolean)
        : undefined;

      if (toList.length === 0) {
        console.error('At least one recipient is required.');
        process.exit(1);
      }

      let body = options.body ?? '';
      let bodyType: 'Text' | 'HTML' = 'Text';

      if (options.markdown) {
        body = markdownToHtml(body);
        bodyType = 'HTML';
      } else if (options.html) {
        bodyType = 'HTML';
      }

      // Process attachments
      let attachments: EmailAttachment[] | undefined;
      const workingDirectory = process.cwd();
      if (options.attach) {
        const filePaths = options.attach
          .split(',')
          .map((f) => f.trim())
          .filter(Boolean);
        attachments = [];

        for (const filePath of filePaths) {
          try {
            const validated = await validateAttachmentPath(filePath, workingDirectory);
            const content = await readFile(validated.absolutePath);
            const contentType = lookupMimeType(validated.fileName);

            attachments.push({
              name: validated.fileName,
              contentType,
              contentBytes: content.toString('base64')
            });

            if (!options.json) {
              console.log(`  Attaching: ${validated.fileName} (${Math.round(validated.size / 1024)} KB)`);
            }
          } catch (err) {
            console.error(`Failed to read attachment: ${filePath}`);
            if (err instanceof AttachmentPathError) {
              console.error(err.message);
            } else {
              console.error(err instanceof Error ? err.message : 'Unknown error');
            }
            process.exit(1);
          }
        }
      }

      const result = await sendEmail(authResult.token!, {
        to: toList,
        cc: ccList,
        bcc: bccList,
        subject: options.subject,
        body,
        bodyType,
        attachments,
        mailbox: options.mailbox
      });

      if (!result.ok) {
        if (options.json) {
          console.log(JSON.stringify({ error: result.error?.message || 'Failed to send email' }, null, 2));
        } else {
          console.error(`Error: ${result.error?.message || 'Failed to send email'}`);
        }
        process.exit(1);
      }

      if (options.json) {
        console.log(
          JSON.stringify(
            {
              success: true,
              to: toList,
              subject: options.subject,
              attachments: attachments?.map((a) => a.name)
            },
            null,
            2
          )
        );
      } else {
        console.log(`\n\u2713 Email sent to ${toList.join(', ')}`);
        console.log(`  Subject: ${options.subject}`);
        if (attachments && attachments.length > 0) {
          console.log(`  Attachments: ${attachments.length}`);
        }
        console.log();
      }
    }
  );
