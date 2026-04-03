/**
 * Build Microsoft Graph `sendMail` payload (POST /me/sendMail).
 */

export interface GraphSendFileAttachment {
  name: string;
  contentType: string;
  /** Base64-encoded file content */
  contentBytes: string;
}

/** Link attachment (Graph `referenceAttachment` with `sourceUrl`). */
export interface GraphSendReferenceAttachment {
  name: string;
  /** HTTPS URL shown as a linked attachment in Outlook. */
  sourceUrl: string;
}

export function buildGraphSendMailPayload(opts: {
  to: string[];
  cc?: string[];
  bcc?: string[];
  subject: string;
  body: string;
  html: boolean;
  categories?: string[];
  fileAttachments?: GraphSendFileAttachment[];
  referenceAttachments?: GraphSendReferenceAttachment[];
}): { message: Record<string, unknown>; saveToSentItems: boolean } {
  const toRecipients = opts.to.map((address) => ({ emailAddress: { address } }));
  const ccRecipients = opts.cc?.filter(Boolean).map((address) => ({ emailAddress: { address } }));
  const bccRecipients = opts.bcc?.filter(Boolean).map((address) => ({ emailAddress: { address } }));

  const body = {
    contentType: opts.html ? 'HTML' : 'Text',
    content: opts.body
  };

  const message: Record<string, unknown> = {
    subject: opts.subject,
    body,
    toRecipients
  };
  if (ccRecipients?.length) message.ccRecipients = ccRecipients;
  if (bccRecipients?.length) message.bccRecipients = bccRecipients;
  if (opts.categories?.length) message.categories = opts.categories;

  const fileParts =
    opts.fileAttachments?.map((a) => ({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: a.name,
      contentType: a.contentType,
      contentBytes: a.contentBytes
    })) ?? [];
  const refParts =
    opts.referenceAttachments?.map((a) => ({
      '@odata.type': '#microsoft.graph.referenceAttachment',
      name: a.name,
      sourceUrl: a.sourceUrl
    })) ?? [];
  const combined = [...fileParts, ...refParts];
  if (combined.length) {
    message.attachments = combined;
  }

  return { message, saveToSentItems: true };
}
