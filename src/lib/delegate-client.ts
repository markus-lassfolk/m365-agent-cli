// Delegate management via EWS SOAP (AddDelegate / GetDelegate / UpdateDelegate / RemoveDelegate)
// No Microsoft Graph equivalent — only available via EWS.

import {
  callEws,
  soapEnvelope,
  extractBlocks,
  extractTag,
  xmlEscape,
  ewsResult,
  ewsError,
  EWS_USERNAME
} from './ews-client.js';

// ─── Types ───

export type DelegateFolderPermissionLevel =
  | 'None'
  | 'Owner'
  | 'PublishingEditor'
  | 'Editor'
  | 'PublishingAuthor'
  | 'Author'
  | 'Reviewer'
  | 'NonEditingAuthor'
  | 'FolderVisible';

export type DeliverMeetingRequests =
  | 'DelegatesAndMe'
  | 'DelegatesOnly'
  | 'DelegatesAndSendInformationToMe'
  | 'NoForward';

export interface DelegatePermissions {
  calendar?: DelegateFolderPermissionLevel;
  inbox?: DelegateFolderPermissionLevel;
  contacts?: DelegateFolderPermissionLevel;
  tasks?: DelegateFolderPermissionLevel;
  notes?: DelegateFolderPermissionLevel;
}

export interface DelegateInfo {
  userId: string;
  displayName?: string;
  primaryEmail?: string;
  permissions: DelegatePermissions;
  viewPrivateItems: boolean;
  deliverMeetingRequests: DeliverMeetingRequests;
}

export interface AddDelegateOptions {
  token: string;
  delegateEmail: string;
  delegateName?: string;
  permissions: DelegatePermissions;
  viewPrivateItems?: boolean;
  deliverMeetingRequests?: DeliverMeetingRequests;
  mailbox?: string;
}

export interface UpdateDelegateOptions {
  token: string;
  delegateEmail: string;
  permissions?: DelegatePermissions;
  viewPrivateItems?: boolean;
  deliverMeetingRequests?: DeliverMeetingRequests;
  mailbox?: string;
}

export interface RemoveDelegateOptions {
  token: string;
  delegateEmail: string;
  mailbox?: string;
}

// ─── Helpers ───

const FOLDER_PERMISSION_ELEMENT_MAP: Record<string, string> = {
  calendar: 'CalendarFolderPermissionLevel',
  inbox: 'InboxFolderPermissionLevel',
  contacts: 'ContactsFolderPermissionLevel',
  tasks: 'TasksFolderPermissionLevel',
  notes: 'NotesFolderPermissionLevel'
};

function buildDelegatePermissionsXml(permissions: DelegatePermissions): string {
  const entries = Object.entries(permissions).filter(([, level]) => level !== undefined);

  if (entries.length === 0) return '';

  return entries
    .map(([folder, level]) => {
      const elementName = FOLDER_PERMISSION_ELEMENT_MAP[folder];
      if (!elementName) return '';
      return `<t:${elementName}>${level}</t:${elementName}>`;
    })
    .join('\n          ');
}

function parseDelegateInfo(block: string, globalDeliver?: DeliverMeetingRequests): DelegateInfo {
  const primaryEmail = extractTag(block, 'PrimarySmtpAddress') || undefined;
  const smtpAddress = extractTag(block, 'SmtpAddress') || undefined;
  const userId = primaryEmail || smtpAddress || '';
  const displayName = extractTag(block, 'DisplayName') || undefined;

  const viewPrivateStr = extractTag(block, 'ViewPrivateItems').toLowerCase();
  const viewPrivateItems = viewPrivateStr === 'true';

  const deliverStr = globalDeliver || extractTag(block, 'DeliverMeetingRequests');
  const deliverMeetingRequests = (deliverStr || 'DelegatesAndMe') as DeliverMeetingRequests;

  // Parse per-folder permissions
  const permissions: DelegatePermissions = {};
  const permissionsBlock = extractBlocks(block, 'DelegatePermissions')[0] || block;

  for (const [folder, elementName] of Object.entries(FOLDER_PERMISSION_ELEMENT_MAP)) {
    const level = extractTag(permissionsBlock, elementName) as DelegateFolderPermissionLevel | '';
    if (level) {
      (permissions as Record<string, string>)[folder] = level;
    }
  }

  return {
    userId,
    displayName,
    primaryEmail,
    permissions,
    viewPrivateItems,
    deliverMeetingRequests
  };
}

// ─── Operations ───

/**
 * Get all delegates (and their permissions) on a mailbox.
 * https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/getdelegate-operation
 */
export async function getDelegates(
  token: string,
  mailbox?: string
): Promise<{ ok: boolean; status: number; data?: DelegateInfo[]; error?: { code: string; message: string } }> {
  try {
    const address = mailbox || EWS_USERNAME;
    const envelope = soapEnvelope(`
    <m:GetDelegate IncludePermissions="true">
      <m:Mailbox>
        <t:EmailAddress>${xmlEscape(address)}</t:EmailAddress>
      </m:Mailbox>
      <m:UserIds />
    </m:GetDelegate>`);

    const xml = await callEws(token, envelope, address);

    const delegateBlocks = extractBlocks(xml, 'DelegateUser');
    const globalDeliverStr = extractTag(xml, 'DeliverMeetingRequests');
    const globalDeliver = globalDeliverStr ? (globalDeliverStr as DeliverMeetingRequests) : undefined;
    const delegates = delegateBlocks.map((block) => parseDelegateInfo(block, globalDeliver));

    return ewsResult(delegates);
  } catch (err) {
    return ewsError(err);
  }
}

/**
 * Add a delegate with per-folder permissions.
 * https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/adddelegate-operation
 */
export async function addDelegate(
  options: AddDelegateOptions
): Promise<{ ok: boolean; status: number; data?: DelegateInfo; error?: { code: string; message: string } }> {
  try {
    const {
      token,
      delegateEmail,
      delegateName,
      permissions,
      viewPrivateItems = false,
      deliverMeetingRequests = 'DelegatesAndMe',
      mailbox
    } = options;

    const address = mailbox || EWS_USERNAME;

    const delegateUserXml = `
      <t:DelegateUser>
        <t:UserId>
          <t:PrimarySmtpAddress>${xmlEscape(delegateEmail)}</t:PrimarySmtpAddress>
          ${delegateName ? `<t:DisplayName>${xmlEscape(delegateName)}</t:DisplayName>` : ''}
        </t:UserId>
        <t:DelegatePermissions>
          ${buildDelegatePermissionsXml(permissions)}
        </t:DelegatePermissions>
        <t:ViewPrivateItems>${viewPrivateItems}</t:ViewPrivateItems>
      </t:DelegateUser>`.trim();

    const envelope = soapEnvelope(`
    <m:AddDelegate>
      <m:Mailbox>
        <t:EmailAddress>${xmlEscape(address)}</t:EmailAddress>
      </m:Mailbox>
      <m:DelegateUsers>
        ${delegateUserXml}
      </m:DelegateUsers>
      <m:DeliverMeetingRequests>${deliverMeetingRequests}</m:DeliverMeetingRequests>
    </m:AddDelegate>`);

    const xml = await callEws(token, envelope, address);
    const delegateBlock = extractBlocks(xml, 'DelegateUser')[0] || '';
    const globalDeliverStr = extractTag(xml, 'DeliverMeetingRequests');
    const globalDeliver = globalDeliverStr ? (globalDeliverStr as DeliverMeetingRequests) : undefined;

    return ewsResult(parseDelegateInfo(delegateBlock, globalDeliver));
  } catch (err) {
    return ewsError(err);
  }
}

/**
 * Update an existing delegate's permissions.
 * https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/updatedelegate-operation
 */
export async function updateDelegate(
  options: UpdateDelegateOptions
): Promise<{ ok: boolean; status: number; data?: DelegateInfo; error?: { code: string; message: string } }> {
  try {
    const { token, delegateEmail, permissions, viewPrivateItems, deliverMeetingRequests, mailbox } = options;

    const address = mailbox || EWS_USERNAME;

    // Build update elements — only include fields that are defined
    const permissionsParts: string[] = [];
    const delegateUserParts: string[] = [];
    let deliverMeetingRequestsXml = '';

    if (permissions !== undefined) {
      const permXml = buildDelegatePermissionsXml(permissions);
      if (permXml) {
        permissionsParts.push(permXml);
      }
    }

    if (viewPrivateItems !== undefined) {
      delegateUserParts.push(`<t:ViewPrivateItems>${viewPrivateItems}</t:ViewPrivateItems>`);
    }

    if (deliverMeetingRequests !== undefined) {
      deliverMeetingRequestsXml = `<m:DeliverMeetingRequests>${deliverMeetingRequests}</m:DeliverMeetingRequests>`;
    }

    if (permissionsParts.length === 0 && delegateUserParts.length === 0 && !deliverMeetingRequestsXml) {
      return { ok: false, status: 400, error: { code: 'NO_UPDATES', message: 'No fields to update' } };
    }

    const delegatePermissionsXml =
      permissionsParts.length > 0
        ? `<t:DelegatePermissions>
          ${permissionsParts.join('\n')}
        </t:DelegatePermissions>`
        : '';

    const delegateUserXml = `
      <t:DelegateUser>
        <t:UserId>
          <t:PrimarySmtpAddress>${xmlEscape(delegateEmail)}</t:PrimarySmtpAddress>
        </t:UserId>
        ${delegatePermissionsXml}
        ${delegateUserParts.join('\n        ')}
      </t:DelegateUser>`.trim();

    const envelope = soapEnvelope(`
    <m:UpdateDelegate>
      <m:Mailbox>
        <t:EmailAddress>${xmlEscape(address)}</t:EmailAddress>
      </m:Mailbox>
      <m:DelegateUsers>
        ${delegateUserXml}
      </m:DelegateUsers>
      ${deliverMeetingRequestsXml}
    </m:UpdateDelegate>`);

    const xml = await callEws(token, envelope, address);
    const delegateBlock = extractBlocks(xml, 'DelegateUser')[0] || '';
    const globalDeliverStr = extractTag(xml, 'DeliverMeetingRequests');
    const globalDeliver = globalDeliverStr ? (globalDeliverStr as DeliverMeetingRequests) : undefined;

    return ewsResult(parseDelegateInfo(delegateBlock, globalDeliver));
  } catch (err) {
    return ewsError(err);
  }
}

/**
 * Remove a delegate from a mailbox.
 * https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/removedelegate-operation
 */
export async function removeDelegate(
  options: RemoveDelegateOptions
): Promise<{ ok: boolean; status: number; error?: { code: string; message: string } }> {
  try {
    const { token, delegateEmail, mailbox } = options;
    const address = mailbox || EWS_USERNAME;

    const envelope = soapEnvelope(`
    <m:RemoveDelegate>
      <m:Mailbox>
        <t:EmailAddress>${xmlEscape(address)}</t:EmailAddress>
      </m:Mailbox>
      <m:UserIds>
        <t:UserId>
          <t:PrimarySmtpAddress>${xmlEscape(delegateEmail)}</t:PrimarySmtpAddress>
        </t:UserId>
      </m:UserIds>
    </m:RemoveDelegate>`);

    await callEws(token, envelope, address);
    return { ok: true, status: 200 };
  } catch (err) {
    return ewsError(err);
  }
}
