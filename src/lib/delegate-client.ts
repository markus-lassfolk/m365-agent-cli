// Delegate management via EWS SOAP (AddDelegate / GetDelegate / UpdateDelegate / RemoveDelegate)
// No Microsoft Graph equivalent — only available via EWS.

import {
  callEws,
  soapEnvelope,
  extractBlocks,
  extractSelfClosingOrBlock,
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

const FOLDER_MAP: Record<string, string> = {
  calendar: 'Calendar',
  inbox: 'Inbox',
  contacts: 'Contacts',
  tasks: 'Tasks',
  notes: 'Notes'
};

function buildDelegatePermissionsXml(permissions: DelegatePermissions): string {
  const entries = Object.entries(permissions).filter(([, level]) => level !== undefined);

  if (entries.length === 0) return '';

  return entries
    .map(([folder, level]) => {
      const folderName = FOLDER_MAP[folder];
      if (!folderName) return '';
      return `<t:${folderName}FolderPermissionLevel>${level}</t:${folderName}FolderPermissionLevel>`;
    })
    .join('\n          ');
}

function parseDelegateInfo(block: string, deliverMeetingRequests: DeliverMeetingRequests): DelegateInfo {
  const userIdBlock = extractSelfClosingOrBlock(block, 'UserId');
  const userId = extractTag(userIdBlock, 'PrimarySmtpAddress') || extractTag(userIdBlock, 'SmtpAddress') || '';
  const displayName = extractTag(userIdBlock, 'DisplayName') || undefined;
  const primaryEmail = extractTag(userIdBlock, 'PrimarySmtpAddress') || undefined;

  const viewPrivateStr = extractTag(block, 'ViewPrivateItems').toLowerCase();
  const viewPrivateItems = viewPrivateStr === 'true';

  const permissions: DelegatePermissions = {};
  for (const [key, folderName] of Object.entries(FOLDER_MAP)) {
    const level = extractTag(block, `${folderName}FolderPermissionLevel`) as DelegateFolderPermissionLevel | '';
    if (level) {
      (permissions as Record<string, string>)[key] = level;
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

    const deliverStr = extractTag(xml, 'DeliverMeetingRequests');
    const deliverMeetingRequests = (deliverStr || 'DelegatesAndMe') as DeliverMeetingRequests;

    const delegateBlocks = extractBlocks(xml, 'DelegateUser');
    const delegates = delegateBlocks.map((block) => parseDelegateInfo(block, deliverMeetingRequests));

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

    const deliverStr = extractTag(xml, 'DeliverMeetingRequests');
    const actualDeliverMeetingRequests = (deliverStr || deliverMeetingRequests) as DeliverMeetingRequests;

    const delegateBlock = extractBlocks(xml, 'DelegateUser')[0] || '';

    return ewsResult(parseDelegateInfo(delegateBlock, actualDeliverMeetingRequests));
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

    const delegateUserParts: string[] = [];

    if (permissions !== undefined) {
      const permXml = buildDelegatePermissionsXml(permissions);
      if (permXml) {
        delegateUserParts.push(`<t:DelegatePermissions>
          ${permXml}
        </t:DelegatePermissions>`);
      }
    }

    if (viewPrivateItems !== undefined) {
      delegateUserParts.push(`<t:ViewPrivateItems>${viewPrivateItems}</t:ViewPrivateItems>`);
    }

    if (delegateUserParts.length === 0 && deliverMeetingRequests === undefined) {
      return { ok: false, status: 400, error: { code: 'NO_UPDATES', message: 'No fields to update' } };
    }

    const delegateUserXml = `
      <t:DelegateUser>
        <t:UserId>
          <t:PrimarySmtpAddress>${xmlEscape(delegateEmail)}</t:PrimarySmtpAddress>
        </t:UserId>
        ${delegateUserParts.join('\n        ')}
      </t:DelegateUser>`.trim();

    const deliverMeetingRequestsXml =
      deliverMeetingRequests !== undefined
        ? `<m:DeliverMeetingRequests>${deliverMeetingRequests}</m:DeliverMeetingRequests>`
        : '';

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

    const deliverStr = extractTag(xml, 'DeliverMeetingRequests');
    const actualDeliverMeetingRequests = (deliverStr || 'DelegatesAndMe') as DeliverMeetingRequests;

    const delegateBlock = extractBlocks(xml, 'DelegateUser')[0] || '';

    return ewsResult(parseDelegateInfo(delegateBlock, actualDeliverMeetingRequests));
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
