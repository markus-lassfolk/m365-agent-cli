/**
 * Microsoft Graph path for `drafts` mutations when `M365_EXCHANGE_BACKEND` is `graph` or `auto`.
 */

import { readFile } from 'node:fs/promises';
import { AttachmentLinkSpecError, parseAttachLinkSpec } from '../lib/attach-link-spec.js';
import { AttachmentPathError, validateAttachmentPath } from '../lib/attachments.js';
import { markdownToHtml } from '../lib/markdown.js';
import { lookupMimeType } from '../lib/mime-type.js';
import {
  addFileAttachmentToMailMessage,
  addReferenceAttachmentToMailMessage,
  createDraftMessage,
  deleteMailMessage,
  patchMailMessage,
  sendMailMessage
} from '../lib/outlook-graph-client.js';

export interface DraftsGraphOptions {
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
  category?: string[];
  clearCategories?: boolean;
}

function normalizeDraftBody(options: DraftsGraphOptions): { body: string; bodyType: 'Text' | 'HTML' } {
  let body = options.body ?? '';
  body = body.replace(/\\n/g, '\n');
  let bodyType: 'Text' | 'HTML' = 'Text';
  if (options.html && body) {
    const escaped = body.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\n/g, '<br>');
    body = body.match(/<\w+[^>]*>/) ? body : escaped;
    bodyType = 'HTML';
  } else if (options.markdown && body) {
    body = markdownToHtml(body);
    bodyType = 'HTML';
  }
  return { body, bodyType };
}

async function addDraftAttachments(
  token: string,
  draftId: string,
  user: string | undefined,
  options: DraftsGraphOptions,
  backend: 'graph' | 'auto'
): Promise<boolean> {
  const wd = process.cwd();
  if (options.attach) {
    const filePaths = options.attach
      .split(',')
      .map((f) => f.trim())
      .filter(Boolean);
    for (const filePath of filePaths) {
      try {
        const validated = await validateAttachmentPath(filePath, wd);
        const content = await readFile(validated.absolutePath);
        if (content.length > 25 * 1024 * 1024) {
          console.error(`File too large (>25MB): ${validated.absolutePath}`);
          process.exit(1);
        }
        const contentType = lookupMimeType(validated.fileName) || 'application/octet-stream';
        const ar = await addFileAttachmentToMailMessage(
          token,
          draftId,
          { name: validated.fileName, contentType, contentBytes: content.toString('base64') },
          user
        );
        if (!ar.ok) {
          if (backend === 'auto') return false;
          console.error(`Failed to attach ${validated.fileName}: ${ar.error?.message}`);
          process.exit(1);
        }
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

  for (const spec of options.attachLink ?? []) {
    try {
      const { name, url } = parseAttachLinkSpec(spec);
      const linkRes = await addReferenceAttachmentToMailMessage(token, draftId, { name, sourceUrl: url }, user);
      if (!linkRes.ok) {
        if (backend === 'auto') return false;
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

  return true;
}

/**
 * @returns `true` if Graph handled the mutation (success or fatal user error); `false` if caller should try EWS (`auto` only).
 */
export async function tryGraphDraftMutations(
  token: string,
  user: string | undefined,
  options: DraftsGraphOptions,
  backend: 'graph' | 'auto'
): Promise<boolean> {
  if (!options.create && !options.edit && !options.send && !options.delete) {
    return false;
  }

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
    const { body, bodyType } = normalizeDraftBody(options);
    const cats = (options.category ?? []).map((c) => c.trim()).filter(Boolean);

    const cr = await createDraftMessage(
      token,
      {
        subject: options.subject,
        bodyContent: body,
        bodyContentType: bodyType,
        toAddresses: toList,
        ccAddresses: ccList,
        categories: cats.length ? cats : undefined
      },
      user
    );
    if (!cr.ok || !cr.data) {
      if (backend === 'auto') return false;
      console.error(`Error: ${cr.error?.message || 'Failed to create draft'}`);
      process.exit(1);
    }
    const draftId = cr.data.id;

    const okAttach = await addDraftAttachments(token, draftId, user, options, backend);
    if (!okAttach) return false;

    if (options.json) {
      console.log(JSON.stringify({ success: true, backend: 'graph', draftId }, null, 2));
    } else {
      console.log(`\n\u2713 Draft created (Graph)`);
      if (options.subject) console.log(`  Subject: ${options.subject}`);
      if (toList) console.log(`  To: ${toList.join(', ')}`);
      console.log();
    }
    return true;
  }

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

    const patch: Record<string, unknown> = {};
    if (options.subject !== undefined) patch.subject = options.subject;
    if (options.body !== undefined || options.markdown || options.html) {
      const { body, bodyType } = normalizeDraftBody(options);
      patch.body = { contentType: bodyType, content: body };
    }
    if (toList !== undefined) {
      patch.toRecipients = toList.map((address) => ({ emailAddress: { address } }));
    }
    if (ccList !== undefined) {
      patch.ccRecipients = ccList.map((address) => ({ emailAddress: { address } }));
    }
    const cats = (options.category ?? []).map((c) => c.trim()).filter(Boolean);
    if (cats.length) {
      patch.categories = cats;
    } else if (options.clearCategories) {
      patch.categories = [];
    }

    if (Object.keys(patch).length > 0) {
      const ur = await patchMailMessage(token, id, patch, user);
      if (!ur.ok) {
        if (backend === 'auto') return false;
        console.error(`Error: ${ur.error?.message || 'Failed to update draft'}`);
        process.exit(1);
      }
    }

    const okAttach = await addDraftAttachments(token, id, user, options, backend);
    if (!okAttach) return false;

    if (options.json) {
      console.log(JSON.stringify({ success: true, backend: 'graph', draftId: id }, null, 2));
    } else {
      console.log(`\u2713 Draft updated (Graph): ${id}`);
    }
    return true;
  }

  if (options.send) {
    const id = options.send.trim();
    const sr = await sendMailMessage(token, id, user);
    if (!sr.ok) {
      if (backend === 'auto') return false;
      console.error(`Error: ${sr.error?.message || 'Failed to send draft'}`);
      process.exit(1);
    }
    if (options.json) {
      console.log(JSON.stringify({ success: true, backend: 'graph', sent: id }, null, 2));
    } else {
      console.log(`\u2713 Draft sent (Graph): ${id}`);
    }
    return true;
  }

  if (options.delete) {
    const id = options.delete.trim();
    if (!id) {
      console.error('Error: --delete requires a draft ID');
      process.exit(1);
    }
    const dr = await deleteMailMessage(token, id, user);
    if (!dr.ok) {
      if (backend === 'auto') return false;
      console.error(`Error: ${dr.error?.message || 'Failed to delete draft'}`);
      process.exit(1);
    }
    if (options.json) {
      console.log(JSON.stringify({ success: true, backend: 'graph', deleted: id }, null, 2));
    } else {
      console.log(`\u2713 Draft deleted (Graph): ${id}`);
    }
    return true;
  }

  return false;
}
