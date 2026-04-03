/**
 * Microsoft Graph path for `mail` when M365_EXCHANGE_BACKEND is graph or auto.
 * Covers list/read/search, mark read, categories, move, download attachments, flags,
 * sensitivity, reply/reply-all/forward (with optional attachments and categories).
 */

import { access, mkdir, readFile, writeFile } from 'node:fs/promises';
import { extname, join } from 'node:path';
import { AttachmentLinkSpecError, parseAttachLinkSpec } from '../lib/attach-link-spec.js';
import { AttachmentPathError, validateAttachmentPath } from '../lib/attachments.js';
import { parseDay, toLocalUnzonedISOString } from '../lib/dates.js';
import { markdownToHtml } from '../lib/markdown.js';
import { lookupMimeType } from '../lib/mime-type.js';
import {
  addFileAttachmentToMailMessage,
  addReferenceAttachmentToMailMessage,
  createMailForwardDraft,
  createMailReplyAllDraft,
  createMailReplyDraft,
  downloadMailMessageAttachmentBytes,
  type GraphMailMessageAttachment,
  getMailMessageAttachment,
  getMessage,
  listAllMailFoldersRecursive,
  listMailMessageAttachments,
  listMessagesInFolder,
  moveMailMessage,
  type OutlookMessage,
  patchMailMessage,
  sendMailMessage
} from '../lib/outlook-graph-client.js';
import { safeAttachmentFileName, writeInternetShortcutUtf8File } from '../lib/safe-filename.js';

export interface MailGraphCommandOptions {
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
  reply?: string;
  replyAll?: string;
  forward?: string;
  to?: string;
  toAddr?: string;
  message?: string;
  markdown?: boolean;
  draft?: boolean;
  attach?: string;
  attachLink?: string[];
  withCategory?: string[];
  setCategories?: string;
  clearCategories?: string;
  category?: string[];
  json?: boolean;
  token?: string;
  mailbox?: string;
  identity?: string;
}

const GRAPH_SENSITIVITY: Record<string, 'normal' | 'personal' | 'private' | 'confidential'> = {
  normal: 'normal',
  personal: 'personal',
  private: 'private',
  confidential: 'confidential'
};

function formatDateShort(dateStr: string): string {
  const date = new Date(dateStr);
  const now = new Date();
  const isToday = date.toDateString() === now.toDateString();
  if (isToday) {
    return date.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit', hour12: false });
  }
  return date.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
}

function truncate(str: string, maxLen: number): string {
  if (!str) return '';
  str = str.replace(/\s+/g, ' ').trim();
  if (str.length <= maxLen) return str;
  return `${str.substring(0, maxLen - 1)}\u2026`;
}

const FOLDER_MAP: Record<string, string> = {
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

const DEST_FOLDER_MAP: Record<string, string> = {
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

/** Whether `opts` only uses the mail operations allowed by the mask (others must be absent). */
function mailOpsMatch(
  opts: MailGraphCommandOptions,
  allow: {
    markRead?: boolean;
    categoryEdit?: boolean;
    move?: boolean;
    download?: boolean;
    readId?: boolean;
    search?: boolean;
    reply?: boolean;
    forward?: boolean;
    flag?: boolean;
    sensitivity?: boolean;
  }
): boolean {
  if (!allow.markRead && (opts.markRead || opts.markUnread)) return false;
  if (!allow.categoryEdit && (opts.setCategories || opts.clearCategories)) return false;
  if (!allow.move && opts.move) return false;
  if (!allow.download && opts.download) return false;
  if (!allow.readId && opts.read) return false;
  if (!allow.search && opts.search) return false;
  if (!allow.reply && (opts.reply || opts.replyAll)) return false;
  if (!allow.forward && opts.forward) return false;
  if (!allow.flag && (opts.flag || opts.unflag || opts.complete)) return false;
  if (!allow.sensitivity && opts.sensitivity) return false;
  return true;
}

/** True when only --mark-read/--mark-unread (by id) — handled via Graph PATCH. */
function isGraphMailMarkReadOnlyOptions(opts: MailGraphCommandOptions): boolean {
  return Boolean(
    (opts.markRead || opts.markUnread) &&
      mailOpsMatch(opts, { markRead: true }) &&
      !opts.startDate &&
      !opts.due &&
      !opts.flagged
  );
}

/** True when only --set-categories / --clear-categories (by message id). */
function isGraphMailCategoriesOnlyOptions(opts: MailGraphCommandOptions): boolean {
  return Boolean(
    (opts.setCategories || opts.clearCategories) &&
      mailOpsMatch(opts, { categoryEdit: true }) &&
      !opts.startDate &&
      !opts.due &&
      !opts.flagged
  );
}

function isGraphMailMoveOnlyOptions(opts: MailGraphCommandOptions): boolean {
  return Boolean(
    opts.move && opts.to && mailOpsMatch(opts, { move: true }) && !opts.startDate && !opts.due && !opts.flagged
  );
}

function isGraphMailDownloadOnlyOptions(opts: MailGraphCommandOptions): boolean {
  return Boolean(opts.download && mailOpsMatch(opts, { download: true }) && !opts.flagged);
}

function isGraphMailFlagOnlyOptions(opts: MailGraphCommandOptions): boolean {
  return Boolean((opts.flag || opts.unflag || opts.complete) && mailOpsMatch(opts, { flag: true }) && !opts.flagged);
}

function isGraphMailSensitivityOnlyOptions(opts: MailGraphCommandOptions): boolean {
  return Boolean(opts.sensitivity && opts.level && mailOpsMatch(opts, { sensitivity: true }) && !opts.flagged);
}

/** Reply, reply-all, or forward — Graph create draft + optional extras + send. */
function isGraphMailReplyForwardGraphOptions(opts: MailGraphCommandOptions): boolean {
  if (opts.reply || opts.replyAll) {
    if (!opts.message?.trim()) return false;
    return mailOpsMatch(opts, { reply: true }) && !opts.flagged;
  }
  if (opts.forward) {
    if (!opts.toAddr?.trim()) return false;
    return mailOpsMatch(opts, { forward: true }) && !opts.flagged;
  }
  return false;
}

export function isGraphMailPortionEligible(opts: MailGraphCommandOptions): boolean {
  return (
    isGraphMailMarkReadOnlyOptions(opts) ||
    isGraphMailCategoriesOnlyOptions(opts) ||
    isGraphMailMoveOnlyOptions(opts) ||
    isGraphMailDownloadOnlyOptions(opts) ||
    isGraphMailFlagOnlyOptions(opts) ||
    isGraphMailSensitivityOnlyOptions(opts) ||
    isGraphMailReplyForwardGraphOptions(opts)
  );
}

function graphUnsupportedForList(opts: MailGraphCommandOptions): boolean {
  return Boolean(
    opts.download ||
      opts.markRead ||
      opts.markUnread ||
      opts.flag ||
      opts.unflag ||
      opts.complete ||
      opts.sensitivity ||
      opts.move ||
      opts.reply ||
      opts.replyAll ||
      opts.forward ||
      opts.setCategories ||
      opts.clearCategories ||
      opts.startDate ||
      opts.due
  );
}

async function resolveDestinationFolderId(
  token: string,
  user: string | undefined,
  toArg: string
): Promise<string | undefined> {
  const key = toArg.toLowerCase();
  const wellKnown = DEST_FOLDER_MAP[key];
  if (wellKnown) return wellKnown;
  const all = await listAllMailFoldersRecursive(token, user);
  if (!all.ok || !all.data) return undefined;
  const found = all.data.find((f) => f.displayName.toLowerCase() === toArg.toLowerCase());
  return found?.id;
}

/**
 * @returns handled true if the Graph path completed (list or read).
 */
export async function tryMailGraphPortion(
  token: string,
  folderArg: string,
  options: MailGraphCommandOptions,
  _cmd: unknown
): Promise<{ handled: boolean }> {
  const user = options.mailbox?.trim() || undefined;

  if (isGraphMailMarkReadOnlyOptions(options)) {
    const id = (options.markRead || options.markUnread)!.trim();
    if (!id) {
      console.error('Error: --mark-read/--mark-unread requires a message ID');
      process.exit(1);
    }
    const isRead = Boolean(options.markRead);
    const r = await patchMailMessage(token, id, { isRead }, user);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message || 'Failed to update message'}`);
      process.exit(1);
    }
    if (options.json) {
      console.log(JSON.stringify({ success: true, id, isRead }, null, 2));
    } else {
      console.log(`\u2713 Marked as ${isRead ? 'read' : 'unread'}: ${id}`);
    }
    return { handled: true };
  }

  if (isGraphMailCategoriesOnlyOptions(options)) {
    if (options.setCategories && options.clearCategories) {
      console.error('Error: use either --set-categories or --clear-categories, not both');
      process.exit(1);
    }
    const id = (options.setCategories || options.clearCategories)!.trim();
    if (!id) {
      console.error('Error: --set-categories/--clear-categories requires a message ID');
      process.exit(1);
    }
    if (options.setCategories) {
      const cats = (options.category ?? []).map((c) => c.trim()).filter(Boolean);
      if (cats.length === 0) {
        console.error('Error: --set-categories requires at least one --category <name>');
        process.exit(1);
      }
      const r = await patchMailMessage(token, id, { categories: cats }, user);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message || 'Failed to set categories'}`);
        process.exit(1);
      }
      if (options.json) {
        console.log(JSON.stringify({ success: true, id, categories: cats }, null, 2));
      } else {
        console.log(`\u2713 Categories set (${cats.join(', ')}): ${id}`);
      }
    } else {
      const r = await patchMailMessage(token, id, { categories: [] }, user);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message || 'Failed to clear categories'}`);
        process.exit(1);
      }
      if (options.json) {
        console.log(JSON.stringify({ success: true, id, categories: [] }, null, 2));
      } else {
        console.log(`\u2713 Categories cleared: ${id}`);
      }
    }
    return { handled: true };
  }

  if (isGraphMailMoveOnlyOptions(options)) {
    const id = options.move!.trim();
    const destId = await resolveDestinationFolderId(token, user, options.to!);
    if (!destId) {
      console.error(`Folder "${options.to}" not found.`);
      console.error('Use "m365-agent-cli folders" to see available folders.');
      process.exit(1);
    }
    const r = await moveMailMessage(token, id, destId, user);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message || 'Failed to move email'}`);
      process.exit(1);
    }
    const folderDisplay = options.to!.charAt(0).toUpperCase() + options.to!.slice(1);
    if (options.json) {
      console.log(JSON.stringify({ success: true, id, destination: options.to }, null, 2));
    } else {
      console.log(`\u2713 Moved to ${folderDisplay}: ${id}`);
    }
    return { handled: true };
  }

  if (isGraphMailDownloadOnlyOptions(options)) {
    const id = options.download!.trim();
    const outDir = options.output || '.';
    const force = Boolean(options.force);

    const summary = await getMessage(token, id, user, 'hasAttachments');
    if (!summary.ok || !summary.data) {
      console.error(`Error: ${summary.error?.message || 'Failed to fetch email'}`);
      process.exit(1);
    }
    if (!summary.data.hasAttachments) {
      console.log('This email has no attachments.');
      return { handled: true };
    }

    const attachmentsResult = await listMailMessageAttachments(token, id, user);
    if (!attachmentsResult.ok || !attachmentsResult.data) {
      console.error(`Error: ${attachmentsResult.error?.message || 'Failed to fetch attachments'}`);
      process.exit(1);
    }

    const attachments = attachmentsResult.data.filter((a) => !a.isInline);
    if (attachments.length === 0) {
      console.log('This email has no downloadable attachments.');
      return { handled: true };
    }

    await mkdir(outDir, { recursive: true });
    console.log(`\nDownloading ${attachments.length} attachment(s) to ${outDir}/\n`);

    const usedPaths = new Set<string>();
    const _wd = process.cwd();

    for (const att of attachments) {
      const odataType = att['@odata.type'] || '';
      const isRef = odataType.includes('referenceAttachment');

      if (isRef) {
        let url = att.sourceUrl;
        if (!url) {
          const full = await getMailMessageAttachment(token, id, att.id, user);
          if (full.ok && full.data && 'sourceUrl' in full.data) {
            url = (full.data as GraphMailMessageAttachment).sourceUrl;
          }
        }
        if (!url) {
          console.error(`  Failed to resolve link: ${att.name || att.id}`);
          continue;
        }
        const safeBase = safeAttachmentFileName(att.name || 'link', 'link');
        let filePath = join(outDir, `${safeBase}.url`);
        let counter = 1;
        while (usedPaths.has(filePath)) {
          filePath = join(outDir, `${safeBase} (${counter}).url`);
          counter++;
        }
        if (!force) {
          while (true) {
            try {
              await access(filePath);
              filePath = join(outDir, `${safeBase} (${counter}).url`);
              counter++;
            } catch {
              break;
            }
          }
        }
        usedPaths.add(filePath);
        const wrote = await writeInternetShortcutUtf8File(filePath, url);
        if (!wrote) {
          usedPaths.delete(filePath);
          console.error(`  Refusing unsafe or invalid link URL: ${att.name || att.id}`);
          continue;
        }
        console.log(`  \u2713 ${filePath.split(/[\\/]/).pop()} (link)`);
        continue;
      }

      const dl = await downloadMailMessageAttachmentBytes(token, id, att.id, user);
      if (!dl.ok || !dl.data) {
        console.error(`  Failed to download: ${att.name || att.id}`);
        continue;
      }
      const content = Buffer.from(dl.data);

      const safeFileName = safeAttachmentFileName(att.name || 'attachment', 'attachment');
      let filePath = join(outDir, safeFileName);
      let counter = 1;
      while (true) {
        if (usedPaths.has(filePath)) {
          const ext = extname(safeFileName);
          const base = safeFileName.slice(0, safeFileName.length - ext.length);
          filePath = join(outDir, `${base} (${counter})${ext}`);
          counter++;
          continue;
        }
        if (!force) {
          try {
            await access(filePath);
            const ext = extname(safeFileName);
            const base = safeFileName.slice(0, safeFileName.length - ext.length);
            filePath = join(outDir, `${base} (${counter})${ext}`);
            counter++;
            continue;
          } catch {
            // free
          }
        }
        break;
      }

      usedPaths.add(filePath);
      await writeFile(filePath, content);
      const sizeKB = Math.round(content.length / 1024);
      const written = filePath.endsWith(safeFileName) ? safeFileName : filePath.split(/[\\/]/).pop();
      console.log(`  \u2713 ${written} (${sizeKB} KB)`);
    }

    console.log('\nDone.\n');
    return { handled: true };
  }

  if (isGraphMailFlagOnlyOptions(options)) {
    const id = (options.flag || options.unflag || options.complete)!.trim();
    let followupFlag: Record<string, unknown>;
    let actionLabel: string;

    if (options.flag) {
      actionLabel = 'Flagged';
      followupFlag = { flagStatus: 'flagged' };
      if (options.startDate) {
        let parsedStartDate: Date;
        try {
          parsedStartDate = parseDay(options.startDate, { throwOnInvalid: true });
        } catch (err) {
          console.error(`Error: Invalid start date: ${err instanceof Error ? err.message : String(err)}`);
          process.exit(1);
        }
        followupFlag.startDateTime = {
          dateTime: toLocalUnzonedISOString(parsedStartDate),
          timeZone: 'UTC'
        };
      }
      if (options.due) {
        let parsedDueDate: Date;
        try {
          parsedDueDate = parseDay(options.due, { throwOnInvalid: true });
        } catch (err) {
          console.error(`Error: Invalid due date: ${err instanceof Error ? err.message : String(err)}`);
          process.exit(1);
        }
        followupFlag.dueDateTime = {
          dateTime: toLocalUnzonedISOString(parsedDueDate),
          timeZone: 'UTC'
        };
      }
    } else if (options.complete) {
      actionLabel = 'Marked complete';
      followupFlag = { flagStatus: 'complete' };
    } else {
      actionLabel = 'Unflagged';
      followupFlag = { flagStatus: 'notFlagged' };
    }

    if (!id) {
      console.error('Error: --flag/--unflag/--complete requires a message ID');
      process.exit(1);
    }

    const r = await patchMailMessage(token, id, { followupFlag }, user);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message || 'Failed to update email'}`);
      process.exit(1);
    }
    if (options.json) {
      console.log(JSON.stringify({ success: true, id, action: actionLabel }, null, 2));
    } else {
      console.log(`\u2713 ${actionLabel}: ${id}`);
    }
    return { handled: true };
  }

  if (isGraphMailSensitivityOnlyOptions(options)) {
    const id = options.sensitivity!.trim();
    const levelKey = options.level!.toLowerCase();
    const sens = GRAPH_SENSITIVITY[levelKey];
    if (!sens) {
      console.error(`Invalid sensitivity level: ${options.level}`);
      console.error('Valid levels: normal, personal, private, confidential');
      process.exit(1);
    }
    const r = await patchMailMessage(token, id, { sensitivity: sens }, user);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message || 'Failed to update email sensitivity'}`);
      process.exit(1);
    }
    if (options.json) {
      console.log(JSON.stringify({ success: true, id, sensitivity: sens }, null, 2));
    } else {
      console.log(`\u2713 Sensitivity set to ${sens}: ${id}`);
    }
    return { handled: true };
  }

  if (isGraphMailReplyForwardGraphOptions(options)) {
    const withCat = (options.withCategory ?? []).map((c) => c.trim()).filter(Boolean);
    const hasAttach = !!options.attach?.trim();
    const hasLinks = (options.attachLink?.length ?? 0) > 0;

    let draftR: Awaited<ReturnType<typeof createMailReplyDraft>>;
    let messageIdForSend: string;

    if (options.reply || options.replyAll) {
      const srcId = (options.reply || options.replyAll)!.trim();
      let bodyText = options.message!;
      if (options.markdown) {
        bodyText = markdownToHtml(bodyText);
      }
      const isHtml = Boolean(options.markdown);
      const comment = isHtml ? undefined : bodyText;
      if (options.replyAll) {
        draftR = await createMailReplyAllDraft(token, srcId, user, comment);
      } else {
        draftR = await createMailReplyDraft(token, srcId, user, comment);
      }
      if (!draftR.ok || !draftR.data) {
        console.error(`Error: ${draftR.error?.message || 'Failed to create reply draft'}`);
        process.exit(1);
      }
      messageIdForSend = draftR.data.id;
      if (isHtml) {
        const pr = await patchMailMessage(
          token,
          messageIdForSend,
          { body: { contentType: 'HTML', content: bodyText } },
          user
        );
        if (!pr.ok) {
          console.error(`Error: ${pr.error?.message || 'Failed to set reply body'}`);
          process.exit(1);
        }
      }
    } else {
      const srcId = options.forward!.trim();
      const recipients = options
        .toAddr!.split(',')
        .map((e) => e.trim())
        .filter(Boolean);
      let bodyText = options.message ?? '';
      if (options.markdown) {
        bodyText = markdownToHtml(bodyText);
      }
      const isHtml = Boolean(options.markdown);
      draftR = await createMailForwardDraft(token, srcId, recipients, user, isHtml ? undefined : bodyText || undefined);
      if (!draftR.ok || !draftR.data) {
        console.error(`Error: ${draftR.error?.message || 'Failed to create forward draft'}`);
        process.exit(1);
      }
      messageIdForSend = draftR.data.id;
      if (isHtml && bodyText) {
        const pr = await patchMailMessage(
          token,
          messageIdForSend,
          { body: { contentType: 'HTML', content: bodyText } },
          user
        );
        if (!pr.ok) {
          console.error(`Error: ${pr.error?.message || 'Failed to set forward body'}`);
          process.exit(1);
        }
      }
    }

    if (withCat.length) {
      const cr = await patchMailMessage(token, messageIdForSend, { categories: withCat }, user);
      if (!cr.ok) {
        console.error(`Error: ${cr.error?.message || 'Failed to set categories on draft'}`);
        process.exit(1);
      }
    }

    if (hasAttach) {
      const paths = options
        .attach!.split(',')
        .map((f) => f.trim())
        .filter(Boolean);
      for (const filePath of paths) {
        try {
          const validated = await validateAttachmentPath(filePath, process.cwd());
          const content = await readFile(validated.absolutePath);
          const contentType = lookupMimeType(validated.fileName) || 'application/octet-stream';
          const ar = await addFileAttachmentToMailMessage(
            token,
            messageIdForSend,
            {
              name: validated.fileName,
              contentType,
              contentBytes: content.toString('base64')
            },
            user
          );
          if (!ar.ok) {
            console.error(`Failed to attach ${validated.fileName}: ${ar.error?.message}`);
            process.exit(1);
          }
          if (!options.json) {
            console.log(`  Attached: ${validated.fileName}`);
          }
        } catch (err) {
          if (err instanceof AttachmentPathError) {
            console.error(err.message);
          } else {
            console.error(`Failed to read attachment: ${filePath}`);
          }
          process.exit(1);
        }
      }
    }

    for (const spec of options.attachLink ?? []) {
      try {
        const { name, url } = parseAttachLinkSpec(spec);
        const lr = await addReferenceAttachmentToMailMessage(token, messageIdForSend, { name, sourceUrl: url }, user);
        if (!lr.ok) {
          console.error(`Failed to attach link ${name}: ${lr.error?.message}`);
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

    if (options.draft) {
      const replyType = options.replyAll ? 'Reply all' : options.forward ? 'Forward' : 'Reply';
      if (options.json) {
        console.log(JSON.stringify({ success: true, draftId: messageIdForSend, kind: replyType }, null, 2));
      } else {
        console.log(`\u2713 ${replyType} draft created: ${messageIdForSend}`);
      }
      return { handled: true };
    }

    const sendR = await sendMailMessage(token, messageIdForSend, user);
    if (!sendR.ok) {
      console.error(`Error: ${sendR.error?.message || 'Failed to send message'}`);
      process.exit(1);
    }

    const replyType = options.replyAll ? 'Reply all' : options.forward ? 'Forward' : 'Reply';
    if (options.json) {
      console.log(
        JSON.stringify(
          { success: true, sent: true, sourceMessageId: options.reply || options.replyAll || options.forward },
          null,
          2
        )
      );
    } else if (hasAttach || hasLinks || withCat.length) {
      console.log(`\u2713 ${replyType} sent (with extras)`);
    } else {
      console.log(`\u2713 ${replyType} sent`);
    }
    return { handled: true };
  }

  if (graphUnsupportedForList(options)) {
    return { handled: false };
  }

  const folderKey = folderArg.toLowerCase();
  let folderId = FOLDER_MAP[folderKey];

  if (!folderId) {
    const all = await listAllMailFoldersRecursive(token, user);
    if (!all.ok || !all.data) {
      console.error(`Error: ${all.error?.message || 'Failed to list folders'}`);
      process.exit(1);
    }
    const found = all.data.find((f) => f.displayName.toLowerCase() === folderArg.toLowerCase());
    if (!found) {
      console.error(`Folder "${folderArg}" not found.`);
      console.error('Use "m365-agent-cli folders" to see available folders.');
      process.exit(1);
    }
    folderId = found.id;
  }

  const limit = Math.max(1, parseInt(options.limit, 10) || 10);
  const page = Math.max(1, parseInt(options.page, 10) || 1);
  const skip = (page - 1) * limit;

  const searchQ = options.search?.trim();
  const filters: string[] = [];
  if (options.unread && !searchQ) {
    filters.push('isRead eq false');
  }
  if (options.flagged && !searchQ) {
    filters.push("followupFlag/flagStatus eq 'flagged'");
  }
  const filter = filters.length ? filters.join(' and ') : undefined;

  if (options.read) {
    const id = options.read.trim();
    const select =
      'subject,body,bodyPreview,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,categories,isRead';
    const full = await getMessage(token, id, user, select);
    if (!full.ok || !full.data) {
      console.error(`Error: ${full.error?.message || 'Failed to fetch email'}`);
      process.exit(1);
    }
    const email = full.data;

    if (options.json) {
      console.log(JSON.stringify(email, null, 2));
      return { handled: true };
    }

    const fromAddr = email.from?.emailAddress?.address ?? '';
    const fromName = email.from?.emailAddress?.name ?? '';
    console.log(`\n${'\u2500'.repeat(60)}`);
    console.log(`From: ${fromName || fromAddr || 'Unknown'}`);
    if (fromAddr) {
      console.log(`      <${fromAddr}>`);
    }
    console.log(`Subject: ${email.subject || '(no subject)'}`);
    const when = email.receivedDateTime || email.sentDateTime;
    console.log(`Date: ${when ? new Date(when).toLocaleString() : 'Unknown'}`);
    if (email.categories?.length) {
      console.log(`Categories: ${email.categories.join(', ')}`);
    }
    console.log(`${'\u2500'.repeat(60)}\n`);
    const content = email.body?.content ?? email.bodyPreview ?? '(no content)';
    console.log(content);
    console.log(`\n${'\u2500'.repeat(60)}\n`);
    return { handled: true };
  }

  const listResult = await listMessagesInFolder(token, folderId, user, {
    top: limit,
    skip,
    orderby: 'receivedDateTime desc',
    filter,
    search: searchQ || undefined
  });

  if (!listResult.ok || !listResult.data) {
    console.error(`Error: ${listResult.error?.message || 'Failed to fetch emails'}`);
    process.exit(1);
  }

  let emails = listResult.data;
  if (options.unread && searchQ) {
    emails = emails.filter((m) => m.isRead === false);
  }
  if (options.flagged && searchQ) {
    emails = emails.filter((m) => m.followupFlag?.flagStatus === 'flagged');
  }

  if (options.json) {
    console.log(JSON.stringify({ value: emails }, null, 2));
    return { handled: true };
  }

  console.log(`\n${'\u2500'.repeat(60)}`);
  console.log(
    `Folder: ${folderArg} (${emails.length} message${emails.length === 1 ? '' : 's'} shown)${user ? ` — ${user}` : ''}`
  );
  console.log(`${'\u2500'.repeat(60)}\n`);

  if (emails.length === 0) {
    console.log('No messages found.\n');
    return { handled: true };
  }

  for (const m of emails as OutlookMessage[]) {
    const from = m.from?.emailAddress?.address ?? '';
    const subj = truncate(m.subject || '(no subject)', 50);
    const when = m.receivedDateTime ? formatDateShort(m.receivedDateTime) : '';
    const read = m.isRead === false ? ' *' : '';
    console.log(`${when}\t${from}\t${subj}\t${m.id}${read}`);
  }

  console.log(`\n${'\u2500'.repeat(60)}`);
  console.log('\nTip: m365-agent-cli mail -r <id> --read <id>');
  console.log('');
  return { handled: true };
}
