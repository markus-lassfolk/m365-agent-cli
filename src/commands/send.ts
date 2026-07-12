import { readFile } from 'node:fs/promises';
import { Command } from 'commander';
import { AttachmentLinkSpecError, parseAttachLinkSpec } from '../lib/attach-link-spec.js';
import { AttachmentPathError, validateAttachmentPath } from '../lib/attachments.js';
import { resolveAuth } from '../lib/auth.js';
import { type EmailAttachment, type ReferenceAttachmentInput, sendEmail } from '../lib/ews-client.js';
import { getExchangeBackend } from '../lib/exchange-backend.js';
import { GRAPH_OUTLOOK_ATTACHMENT_SESSION_THRESHOLD_BYTES } from '../lib/graph-attachment-upload-session.js';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import { buildGraphSendMailPayload } from '../lib/graph-send-mail.js';
import { toJsonError } from '../lib/json-error.js';
import { MailTemplateError, parseTemplateVars, renderMailTemplate } from '../lib/mail-template.js';
import { markdownToHtml } from '../lib/markdown.js';
import { lookupMimeType } from '../lib/mime-type.js';
import {
  addFileAttachmentToMailMessage,
  addReferenceAttachmentToMailMessage,
  createDraftMessage,
  sendMail as graphSendMail,
  sendMailMessage
} from '../lib/outlook-graph-client.js';
import { checkReadOnly } from '../lib/utils.js';

/** Graph sendMail often returns access denied without Mail.Send on the app registration + refreshed token. */
function graphSendDeniedHint(message: string, code?: string): string | undefined {
  const lower = message.toLowerCase();
  const codeMatch = (code || '').toLowerCase().includes('erroraccess') || lower.includes('erroraccessdenied');
  if (codeMatch || lower.includes('access is denied')) {
    return 'Hint: add delegated Mail.Send (and Mail.ReadWrite) on the Entra app, admin-consent if needed, then `m365-agent-cli login` again. See docs/GRAPH_SCOPES.md. Or set M365_EXCHANGE_BACKEND=ews to use EWS for this send.';
  }
  return undefined;
}

export const sendCommand = new Command('send')
  .description('Send an email (EWS or Microsoft Graph per M365_EXCHANGE_BACKEND)')
  .requiredOption('--to <emails>', 'Recipient email(s), comma-separated')
  .requiredOption('--subject <text>', 'Email subject')
  .option('--body <text>', 'Email body', '')
  .option(
    '--template <path>',
    'Read the body from a template file with {{variable}} / {{variable|default}} placeholders (mutually exclusive with --body)'
  )
  .option(
    '--var <nameValue>',
    'Template variable "name=value" (repeatable; use with --template)',
    (v: string, prev: string[]) => [...prev, v],
    [] as string[]
  )
  .option('--category <name>', 'Category label (repeatable)', (v, acc) => [...acc, v], [] as string[])
  .option('--cc <emails>', 'CC recipient(s), comma-separated')
  .option('--bcc <emails>', 'BCC recipient(s), comma-separated')
  .option('--attach <files>', 'Attach file(s), comma-separated paths')
  .option(
    '--attach-link <spec>',
    'Attach link: "Title|https://url" or bare https URL (repeatable; Graph `referenceAttachment` or EWS)',
    (v: string, prev: string[]) => [...prev, v],
    [] as string[]
  )
  .option('--html', 'Send body as HTML')
  .option('--markdown', 'Parse body as markdown (bold, links, lists)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific access token')
  .option('--mailbox <email>', 'Send from shared mailbox (Send As / Graph user)')
  .option('--identity <name>', 'Use a specific authentication identity (default: default)')
  .action(
    async (
      options: {
        to: string;
        subject: string;
        body?: string;
        template?: string;
        var?: string[];
        cc?: string;
        bcc?: string;
        attach?: string;
        attachLink?: string[];
        html?: boolean;
        markdown?: boolean;
        json?: boolean;
        token?: string;
        mailbox?: string;
        identity?: string;
        category?: string[];
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const backend = getExchangeBackend();
      const linkSpecs = options.attachLink ?? [];

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

      if (options.template && options.body) {
        console.error('Error: --template and --body are mutually exclusive.');
        process.exit(1);
      }

      let body = options.body ?? '';
      if (options.template) {
        try {
          const source = await readFile(options.template.trim(), 'utf-8');
          const vars = parseTemplateVars(options.var ?? []);
          body = renderMailTemplate(source, vars);
        } catch (err) {
          const message =
            err instanceof MailTemplateError
              ? err.message
              : `Could not read/render --template: ${err instanceof Error ? err.message : String(err)}`;
          if (options.json) {
            console.log(JSON.stringify({ error: toJsonError(message) }, null, 2));
          } else {
            console.error(`Error: ${message}`);
          }
          process.exit(1);
        }
      }
      let html = Boolean(options.html);
      if (options.markdown) {
        body = markdownToHtml(body);
        html = true;
      } else if (options.html) {
        html = true;
      }

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

      let referenceAttachments: ReferenceAttachmentInput[] | undefined;
      if (linkSpecs.length > 0) {
        referenceAttachments = [];
        for (const spec of linkSpecs) {
          try {
            const { name, url } = parseAttachLinkSpec(spec);
            referenceAttachments.push({ name, url, contentType: 'text/html' });
            if (!options.json) {
              console.log(`  Attaching link: ${name}`);
            }
          } catch (err) {
            const msg =
              err instanceof AttachmentLinkSpecError ? err.message : err instanceof Error ? err.message : String(err);
            console.error(`Invalid --attach-link: ${msg}`);
            process.exit(1);
          }
        }
      }

      const categories = options.category && options.category.length > 0 ? options.category : undefined;
      const user = options.mailbox?.trim() || undefined;

      async function sendEws(): Promise<void> {
        const authResult = await resolveAuth({
          token: options.token,
          identity: options.identity
        });

        if (!authResult.success) {
          if (options.json) {
            console.log(JSON.stringify({ error: toJsonError(authResult.error) }, null, 2));
          } else {
            console.error(`Error: ${authResult.error}`);
            console.error('\nCheck your .env file for EWS_CLIENT_ID and EWS_REFRESH_TOKEN.');
          }
          process.exit(1);
        }

        const bodyType: 'Text' | 'HTML' = html ? 'HTML' : 'Text';
        const result = await sendEmail(authResult.token!, {
          to: toList,
          cc: ccList,
          bcc: bccList,
          subject: options.subject,
          body,
          bodyType,
          attachments,
          referenceAttachments,
          mailbox: options.mailbox,
          categories
        });

        if (!result.ok) {
          if (options.json) {
            console.log(
              JSON.stringify({ error: toJsonError(result.error?.message || 'Failed to send email') }, null, 2)
            );
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
                backend: 'ews',
                to: toList,
                subject: options.subject,
                attachments: attachments?.map((a) => a.name),
                attachLinks: referenceAttachments?.map((a) => a.name)
              },
              null,
              2
            )
          );
        } else {
          console.log(`\n\u2713 Email sent to ${toList.join(', ')}`);
          console.log(`  Subject: ${options.subject}`);
          const nFile = attachments?.length ?? 0;
          const nLink = referenceAttachments?.length ?? 0;
          if (nFile + nLink > 0) {
            console.log(`  Attachments: ${nFile} file(s), ${nLink} link(s)`);
          }
          console.log();
        }
      }

      if (backend === 'ews') {
        await sendEws();
        return;
      }

      const graphAuth = await resolveGraphAuth({
        token: options.token,
        identity: options.identity
      });

      if (!graphAuth.success || !graphAuth.token) {
        if (backend === 'graph') {
          if (options.json) {
            console.log(JSON.stringify({ error: toJsonError(graphAuth.error || 'Graph auth failed') }, null, 2));
          } else {
            console.error(`Error: ${graphAuth.error || 'Graph authentication failed'}`);
            console.error('\nSet EWS_CLIENT_ID and M365_REFRESH_TOKEN for Graph, or run `m365-agent-cli login`.');
          }
          process.exit(1);
        }
        await sendEws();
        return;
      }

      const anyLargeFile =
        attachments?.some((a) => {
          try {
            return Buffer.from(a.contentBytes, 'base64').byteLength >= GRAPH_OUTLOOK_ATTACHMENT_SESSION_THRESHOLD_BYTES;
          } catch {
            return false;
          }
        }) ?? false;

      if (anyLargeFile) {
        const dr = await createDraftMessage(
          graphAuth.token,
          {
            subject: options.subject,
            bodyContent: body,
            bodyContentType: html ? 'HTML' : 'Text',
            toAddresses: toList,
            ccAddresses: ccList,
            bccAddresses: bccList,
            categories
          },
          user
        );
        if (!dr.ok || !dr.data?.id) {
          if (backend === 'auto') {
            await sendEws();
            return;
          }
          if (options.json) {
            console.log(
              JSON.stringify(
                { error: toJsonError(dr.error?.message || 'Failed to create draft for large attachments') },
                null,
                2
              )
            );
          } else {
            console.error(`Error: ${dr.error?.message || 'Failed to create draft for large attachments'}`);
          }
          process.exit(1);
        }
        const draftId = dr.data.id;
        for (const a of attachments ?? []) {
          const ar = await addFileAttachmentToMailMessage(
            graphAuth.token,
            draftId,
            { name: a.name, contentType: a.contentType, contentBytes: a.contentBytes },
            user
          );
          if (!ar.ok) {
            if (backend === 'auto') {
              await sendEws();
              return;
            }
            if (options.json) {
              console.log(
                JSON.stringify({ error: toJsonError(ar.error?.message || 'Failed to attach file') }, null, 2)
              );
            } else {
              console.error(`Error: ${ar.error?.message || 'Failed to attach file'}`);
            }
            process.exit(1);
          }
        }
        for (const ref of referenceAttachments ?? []) {
          const lr = await addReferenceAttachmentToMailMessage(
            graphAuth.token,
            draftId,
            { name: ref.name, sourceUrl: ref.url },
            user
          );
          if (!lr.ok) {
            if (backend === 'auto') {
              await sendEws();
              return;
            }
            if (options.json) {
              console.log(
                JSON.stringify({ error: toJsonError(lr.error?.message || 'Failed to attach link') }, null, 2)
              );
            } else {
              console.error(`Error: ${lr.error?.message || 'Failed to attach link'}`);
            }
            process.exit(1);
          }
        }
        const sr = await sendMailMessage(graphAuth.token, draftId, user);
        if (!sr.ok) {
          if (backend === 'auto') {
            await sendEws();
            return;
          }
          if (options.json) {
            const payloadErr: Record<string, unknown> = {
              error: sr.error?.message || 'Failed to send draft',
              backend: 'graph'
            };
            const hint = graphSendDeniedHint(sr.error?.message || '', sr.error?.code);
            if (hint) payloadErr.hint = hint;
            console.log(JSON.stringify(payloadErr, null, 2));
          } else {
            console.error(`Error: ${sr.error?.message || 'Failed to send draft'}`);
            const hint = graphSendDeniedHint(sr.error?.message || '', sr.error?.code);
            if (hint) console.error(hint);
          }
          process.exit(1);
        }
      } else {
        const payload = buildGraphSendMailPayload({
          to: toList,
          cc: ccList,
          bcc: bccList,
          subject: options.subject,
          body,
          html,
          categories,
          fileAttachments: attachments,
          referenceAttachments: referenceAttachments?.map((a) => ({ name: a.name, sourceUrl: a.url }))
        });

        const result = await graphSendMail(graphAuth.token, payload, user);
        if (!result.ok) {
          if (backend === 'auto') {
            await sendEws();
            return;
          }
          if (options.json) {
            const payload: Record<string, unknown> = {
              error: result.error?.message || 'Failed to send email',
              backend: 'graph'
            };
            const hint = graphSendDeniedHint(result.error?.message || '', result.error?.code);
            if (hint) payload.hint = hint;
            console.log(JSON.stringify(payload, null, 2));
          } else {
            console.error(`Error: ${result.error?.message || 'Failed to send email'}`);
            const hint = graphSendDeniedHint(result.error?.message || '', result.error?.code);
            if (hint) console.error(hint);
          }
          process.exit(1);
        }
      }

      if (options.json) {
        console.log(
          JSON.stringify(
            {
              success: true,
              backend: 'graph',
              to: toList,
              subject: options.subject,
              attachments: attachments?.map((a) => a.name),
              attachLinks: referenceAttachments?.map((a) => a.name)
            },
            null,
            2
          )
        );
      } else {
        console.log(`\n\u2713 Email sent (Graph) to ${toList.join(', ')}`);
        console.log(`  Subject: ${options.subject}`);
        const nFile = attachments?.length ?? 0;
        const nLink = referenceAttachments?.length ?? 0;
        if (nFile + nLink > 0) {
          console.log(`  Attachments: ${nFile} file(s), ${nLink} link(s)`);
        }
        console.log();
      }
    }
  );
