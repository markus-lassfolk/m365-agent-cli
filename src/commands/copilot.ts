import { access, readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { Command } from 'commander';
import {
  assertCopilotReportPeriod,
  buildCopilotRetrievalBody,
  buildCopilotSearchBody,
  COPILOT_RETRIEVAL_DATA_SOURCES,
  copilotAdminCatalogDelete,
  copilotAdminCatalogGet,
  copilotAdminCatalogPatch,
  copilotAdminLimitedModeDelete,
  copilotAdminLimitedModeGet,
  copilotAdminLimitedModePatch,
  copilotAdminNavDelete,
  copilotAdminNavGet,
  copilotAdminNavPatch,
  copilotAdminSettingsDelete,
  copilotAdminSettingsGet,
  copilotAdminSettingsPatch,
  copilotAgentGet,
  copilotAgentsCount,
  copilotAgentsList,
  copilotAiUserCreate,
  copilotAiUserDelete,
  copilotAiUserGet,
  copilotAiUserInteractionHistoryDelete,
  copilotAiUserInteractionHistoryGet,
  copilotAiUserInteractionHistoryPatch,
  copilotAiUserOnlineMeetingCreate,
  copilotAiUserOnlineMeetingDelete,
  copilotAiUserOnlineMeetingGet,
  copilotAiUserOnlineMeetingPatch,
  copilotAiUserOnlineMeetingsCount,
  copilotAiUserOnlineMeetingsList,
  copilotAiUserPatch,
  copilotAiUsersCount,
  copilotAiUsersList,
  copilotCommunicationsDelete,
  copilotCommunicationsGet,
  copilotCommunicationsPatch,
  copilotConversationChat,
  copilotConversationChatOverStream,
  copilotConversationCreate,
  copilotConversationDelete,
  copilotConversationDeleteByThreadId,
  copilotConversationGet,
  copilotConversationMessageCreate,
  copilotConversationMessageDelete,
  copilotConversationMessageGet,
  copilotConversationMessagePatch,
  copilotConversationMessagesCount,
  copilotConversationMessagesList,
  copilotConversationPatch,
  copilotConversationsCount,
  copilotConversationsList,
  copilotInteractionHistoryNavDelete,
  copilotInteractionHistoryNavGet,
  copilotInteractionHistoryNavPatch,
  copilotInteractionsExportList,
  copilotInteractionsTenantExportList,
  copilotMeetingAiInsightDelete,
  copilotMeetingAiInsightPatch,
  copilotMeetingAiInsightsCount,
  copilotMeetingAiInsightsCreate,
  copilotMeetingInsightGet,
  copilotMeetingInsightsList,
  copilotPackagesBlock,
  copilotPackagesCount,
  copilotPackagesCreate,
  copilotPackagesDelete,
  copilotPackagesGet,
  copilotPackagesList,
  copilotPackagesReassign,
  copilotPackagesUnblock,
  copilotPackagesUpdate,
  copilotPackageZipDelete,
  copilotPackageZipDownload,
  copilotPackageZipUpload,
  copilotRealtimeActivityFeedDelete,
  copilotRealtimeActivityFeedGet,
  copilotRealtimeActivityFeedPatch,
  copilotRealtimeMeetingCreate,
  copilotRealtimeMeetingDelete,
  copilotRealtimeMeetingGet,
  copilotRealtimeMeetingPatch,
  copilotRealtimeMeetingsCount,
  copilotRealtimeMeetingsList,
  copilotRealtimeSubscriptionCreate,
  copilotRealtimeSubscriptionDelete,
  copilotRealtimeSubscriptionGet,
  copilotRealtimeSubscriptionGetArtifacts,
  copilotRealtimeSubscriptionPatch,
  copilotRealtimeSubscriptionsCount,
  copilotRealtimeSubscriptionsList,
  copilotRealtimeTranscriptCreate,
  copilotRealtimeTranscriptDelete,
  copilotRealtimeTranscriptGet,
  copilotRealtimeTranscriptPatch,
  copilotRealtimeTranscriptsCount,
  copilotRealtimeTranscriptsList,
  copilotReportGet,
  copilotReportsNavDelete,
  copilotReportsNavGet,
  copilotReportsNavPatch,
  copilotRetrieval,
  copilotRootGet,
  copilotRootPatch,
  copilotSearch,
  copilotSearchNextPage,
  copilotSettingsDelete,
  copilotSettingsEnhancedPersonalizationDelete,
  copilotSettingsEnhancedPersonalizationGet,
  copilotSettingsEnhancedPersonalizationPatch,
  copilotSettingsGet,
  copilotSettingsPatch,
  copilotSettingsPeopleDelete,
  copilotSettingsPeopleGet,
  copilotSettingsPeoplePatch
} from '../lib/copilot-graph-client.js';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import { readJsonFileOrExit } from '../lib/read-json-file.js';
import { checkReadOnly } from '../lib/utils.js';

export const copilotCommand = new Command('copilot').description(
  'Microsoft 365 Copilot APIs on Microsoft Graph (/copilot/...). Licensing, roles, and preview terms apply; see https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/copilot-apis-overview'
);

type AuthOpts = { token?: string; identity?: string };

async function resolveTokenOrExit(opts: AuthOpts): Promise<string> {
  const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
  if (!auth.success || !auth.token) {
    console.error(`Auth error: ${auth.error}`);
    process.exit(1);
  }
  return auth.token;
}

function printJson(data: unknown): void {
  console.log(JSON.stringify(data, null, 2));
}

function exitGraphError(prefix: string, message: string | undefined): never {
  console.error(`${prefix}${message || 'Unknown error'}`);
  process.exit(1);
}

function ifMatchHeader(ifMatch: string | undefined): Record<string, string> | undefined {
  const v = ifMatch?.trim();
  return v ? { 'If-Match': v } : undefined;
}

/** POST /copilot/retrieval */
copilotCommand
  .command('retrieval')
  .description('POST /copilot/retrieval — grounding extracts (SharePoint, OneDrive, connectors)')
  .option(
    '-q, --query <text>',
    'Natural language query (max 1500 chars; required with --data-source unless --json-file)'
  )
  .option('-s, --data-source <source>', `With --query: ${COPILOT_RETRIEVAL_DATA_SOURCES.join(' | ')}`)
  .option('--filter-expression <kql>', 'Optional KQL filterExpression')
  .option('--max <n>', 'maximumNumberOfResults (1–25)', (v) => parseInt(String(v), 10))
  .option('-m, --metadata <fields>', 'Comma-separated resourceMetadata names')
  .option('-f, --json-file <path>', 'Full JSON body (overrides query flags)')
  .option('--beta', 'Use Graph beta')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (opts: {
      query?: string;
      dataSource?: string;
      filterExpression?: string;
      max?: number;
      metadata?: string;
      jsonFile?: string;
      beta?: boolean;
      token?: string;
      identity?: string;
    }) => {
      const token = await resolveTokenOrExit(opts);
      let body: Record<string, unknown>;
      if (opts.jsonFile) {
        body = await readJsonFileOrExit(opts.jsonFile, '--json-file');
      } else {
        try {
          body = buildCopilotRetrievalBody({
            queryString: opts.query ?? '',
            dataSource: opts.dataSource ?? '',
            filterExpression: opts.filterExpression,
            maximumNumberOfResults: opts.max,
            resourceMetadata: opts.metadata
              ?.split(',')
              .map((s) => s.trim())
              .filter(Boolean)
          });
        } catch (e) {
          console.error(e instanceof Error ? e.message : String(e));
          process.exit(1);
        }
      }
      const r = await copilotRetrieval(token, body, Boolean(opts.beta));
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      printJson(r.data);
    }
  );

/** POST /copilot/search (preview; defaults to beta) */
copilotCommand
  .command('search')
  .description('POST /copilot/search — hybrid search over OneDrive for work or school (preview; beta by default)')
  .option('-q, --query <text>', 'Natural language query (required unless --json-file)')
  .option('--page-size <n>', 'Results per page (1–100)', (v) => parseInt(String(v), 10))
  .option('--one-drive-filter <kql>', 'dataSources.oneDrive.filterExpression (path KQL)')
  .option('-m, --metadata <fields>', 'Comma-separated resourceMetadataNames for OneDrive')
  .option('-f, --json-file <path>', 'Full JSON body')
  .option('--v1', 'Use Graph v1.0 (search is generally beta; v1 may 404 until GA)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (opts: {
      query?: string;
      pageSize?: number;
      oneDriveFilter?: string;
      metadata?: string;
      jsonFile?: string;
      v1?: boolean;
      token?: string;
      identity?: string;
    }) => {
      const token = await resolveTokenOrExit(opts);
      let body: Record<string, unknown>;
      if (opts.jsonFile) {
        body = await readJsonFileOrExit(opts.jsonFile, '--json-file');
      } else {
        try {
          body = buildCopilotSearchBody({
            query: opts.query ?? '',
            pageSize: opts.pageSize,
            oneDriveFilterExpression: opts.oneDriveFilter,
            resourceMetadataNames: opts.metadata
              ?.split(',')
              .map((s) => s.trim())
              .filter(Boolean)
          });
        } catch (e) {
          console.error(e instanceof Error ? e.message : String(e));
          process.exit(1);
        }
      }
      const r = await copilotSearch(token, body, !opts.v1);
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      printJson(r.data);
    }
  );

/** GET @odata.nextLink from search */
copilotCommand
  .command('search-next')
  .description('GET Copilot search next page — pass full @odata.nextLink URL from a prior search response')
  .requiredOption('--next-link <url>', 'Full HTTPS nextLink from copilot search response')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { nextLink: string; token?: string; identity?: string }) => {
    const token = await resolveTokenOrExit(opts);
    const r = await copilotSearchNextPage(token, opts.nextLink);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    printJson(r.data);
  });

/** GET /copilot/conversations */
copilotCommand
  .command('conversations-list')
  .description('GET /copilot/conversations — list Copilot chat conversations for the signed-in user')
  .option('--odata <query>', 'OData without leading ?')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { odata?: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }) => {
    const token = await resolveTokenOrExit(opts);
    const r = await copilotConversationsList(token, opts.odata, !opts.v1);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    printJson(r.data);
  });

copilotCommand
  .command('conversation-get')
  .description('GET /copilot/conversations/{id}')
  .argument('<conversationId>', 'Conversation id')
  .option('--odata <query>', 'OData without leading ?')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      conversationId: string,
      opts: { odata?: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }
    ) => {
      const token = await resolveTokenOrExit(opts);
      const r = await copilotConversationGet(token, conversationId, opts.odata, !opts.v1);
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      printJson(r.data);
    }
  );

copilotCommand
  .command('conversation-patch')
  .description('PATCH /copilot/conversations/{id}')
  .argument('<conversationId>', 'Conversation id')
  .requiredOption('-f, --json-file <path>', 'JSON body')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--if-match <etag>', 'Optional If-Match header')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      conversationId: string,
      opts: { jsonFile: string; beta?: boolean; v1?: boolean; ifMatch?: string; token?: string; identity?: string },
      cmd
    ) => {
      checkReadOnly(cmd);
      const token = await resolveTokenOrExit(opts);
      const body = await readJsonFileOrExit(opts.jsonFile, '--json-file');
      const r = await copilotConversationPatch(token, conversationId, body, !opts.v1, ifMatchHeader(opts.ifMatch));
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      if (r.data !== undefined) printJson(r.data);
      else console.log('OK (204 No Content)');
    }
  );

copilotCommand
  .command('conversation-delete')
  .description('DELETE /copilot/conversations/{id}')
  .argument('<conversationId>', 'Conversation id')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--if-match <etag>', 'Optional If-Match header')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      conversationId: string,
      opts: { beta?: boolean; v1?: boolean; ifMatch?: string; token?: string; identity?: string },
      cmd
    ) => {
      checkReadOnly(cmd);
      const token = await resolveTokenOrExit(opts);
      const r = await copilotConversationDelete(token, conversationId, !opts.v1, ifMatchHeader(opts.ifMatch));
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      if (r.data !== undefined) printJson(r.data);
      else console.log('OK (204 No Content)');
    }
  );

copilotCommand
  .command('conversation-delete-by-thread')
  .description('POST /copilot/conversations/microsoft.graph.copilot.deleteByThreadId')
  .requiredOption('-f, --json-file <path>', 'Action parameters JSON (see Graph docs)')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { jsonFile: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }, cmd) => {
    checkReadOnly(cmd);
    const token = await resolveTokenOrExit(opts);
    const body = await readJsonFileOrExit(opts.jsonFile, '--json-file');
    const r = await copilotConversationDeleteByThreadId(token, body, !opts.v1);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    if (r.data !== undefined) printJson(r.data);
    else console.log('OK (204 No Content)');
  });

copilotCommand
  .command('messages-list')
  .description('GET /copilot/conversations/{id}/messages')
  .argument('<conversationId>', 'Conversation id')
  .option('--odata <query>', 'OData without leading ?')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      conversationId: string,
      opts: { odata?: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }
    ) => {
      const token = await resolveTokenOrExit(opts);
      const r = await copilotConversationMessagesList(token, conversationId, opts.odata, !opts.v1);
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      printJson(r.data);
    }
  );

copilotCommand
  .command('message-get')
  .description('GET /copilot/conversations/{id}/messages/{messageId}')
  .argument('<conversationId>', 'Conversation id')
  .argument('<messageId>', 'Message id')
  .option('--odata <query>', 'OData without leading ?')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      conversationId: string,
      messageId: string,
      opts: { odata?: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }
    ) => {
      const token = await resolveTokenOrExit(opts);
      const r = await copilotConversationMessageGet(token, conversationId, messageId, opts.odata, !opts.v1);
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      printJson(r.data);
    }
  );

copilotCommand
  .command('message-create')
  .description('POST /copilot/conversations/{id}/messages')
  .argument('<conversationId>', 'Conversation id')
  .requiredOption('-f, --json-file <path>', 'JSON body (microsoft.graph.copilotConversationMessage)')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      conversationId: string,
      opts: { jsonFile: string; beta?: boolean; v1?: boolean; token?: string; identity?: string },
      cmd
    ) => {
      checkReadOnly(cmd);
      const token = await resolveTokenOrExit(opts);
      const body = await readJsonFileOrExit(opts.jsonFile, '--json-file');
      const r = await copilotConversationMessageCreate(token, conversationId, body, !opts.v1);
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      printJson(r.data);
    }
  );

copilotCommand
  .command('message-patch')
  .description('PATCH /copilot/conversations/{id}/messages/{messageId}')
  .argument('<conversationId>', 'Conversation id')
  .argument('<messageId>', 'Message id')
  .requiredOption('-f, --json-file <path>', 'JSON body')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--if-match <etag>', 'Optional If-Match header')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      conversationId: string,
      messageId: string,
      opts: { jsonFile: string; beta?: boolean; v1?: boolean; ifMatch?: string; token?: string; identity?: string },
      cmd
    ) => {
      checkReadOnly(cmd);
      const token = await resolveTokenOrExit(opts);
      const body = await readJsonFileOrExit(opts.jsonFile, '--json-file');
      const r = await copilotConversationMessagePatch(
        token,
        conversationId,
        messageId,
        body,
        !opts.v1,
        ifMatchHeader(opts.ifMatch)
      );
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      if (r.data !== undefined) printJson(r.data);
      else console.log('OK (204 No Content)');
    }
  );

copilotCommand
  .command('message-delete')
  .description('DELETE /copilot/conversations/{id}/messages/{messageId}')
  .argument('<conversationId>', 'Conversation id')
  .argument('<messageId>', 'Message id')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--if-match <etag>', 'Optional If-Match header')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      conversationId: string,
      messageId: string,
      opts: { beta?: boolean; v1?: boolean; ifMatch?: string; token?: string; identity?: string },
      cmd
    ) => {
      checkReadOnly(cmd);
      const token = await resolveTokenOrExit(opts);
      const r = await copilotConversationMessageDelete(
        token,
        conversationId,
        messageId,
        !opts.v1,
        ifMatchHeader(opts.ifMatch)
      );
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      if (r.data !== undefined) printJson(r.data);
      else console.log('OK (204 No Content)');
    }
  );

copilotCommand
  .command('agents-list')
  .description('GET /copilot/agents')
  .option('--odata <query>', 'OData without leading ?')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { odata?: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }) => {
    const token = await resolveTokenOrExit(opts);
    const r = await copilotAgentsList(token, opts.odata, !opts.v1);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    printJson(r.data);
  });

copilotCommand
  .command('agent-get')
  .description('GET /copilot/agents/{id}')
  .argument('<agentId>', 'Agent id')
  .option('--odata <query>', 'OData without leading ?')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      agentId: string,
      opts: { odata?: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }
    ) => {
      const token = await resolveTokenOrExit(opts);
      const r = await copilotAgentGet(token, agentId, opts.odata, !opts.v1);
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      printJson(r.data);
    }
  );

copilotCommand
  .command('settings-get')
  .description('GET /copilot/settings')
  .option('--odata <query>', 'OData without leading ?')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { odata?: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }) => {
    const token = await resolveTokenOrExit(opts);
    const r = await copilotSettingsGet(token, opts.odata, !opts.v1);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    printJson(r.data);
  });

copilotCommand
  .command('settings-patch')
  .description('PATCH /copilot/settings')
  .requiredOption('-f, --json-file <path>', 'JSON body')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { jsonFile: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }, cmd) => {
    checkReadOnly(cmd);
    const token = await resolveTokenOrExit(opts);
    const body = await readJsonFileOrExit(opts.jsonFile, '--json-file');
    const r = await copilotSettingsPatch(token, body, !opts.v1);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    if (r.data !== undefined) printJson(r.data);
    else console.log('OK (204 No Content)');
  });

copilotCommand
  .command('settings-people-get')
  .description('GET /copilot/settings/people')
  .option('--odata <query>', 'OData without leading ?')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { odata?: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }) => {
    const token = await resolveTokenOrExit(opts);
    const r = await copilotSettingsPeopleGet(token, opts.odata, !opts.v1);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    printJson(r.data);
  });

copilotCommand
  .command('settings-people-patch')
  .description('PATCH /copilot/settings/people')
  .requiredOption('-f, --json-file <path>', 'JSON body')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { jsonFile: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }, cmd) => {
    checkReadOnly(cmd);
    const token = await resolveTokenOrExit(opts);
    const body = await readJsonFileOrExit(opts.jsonFile, '--json-file');
    const r = await copilotSettingsPeoplePatch(token, body, !opts.v1);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    if (r.data !== undefined) printJson(r.data);
    else console.log('OK (204 No Content)');
  });

copilotCommand
  .command('settings-enhanced-personalization-get')
  .description('GET /copilot/settings/people/enhancedPersonalization')
  .option('--odata <query>', 'OData without leading ?')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { odata?: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }) => {
    const token = await resolveTokenOrExit(opts);
    const r = await copilotSettingsEnhancedPersonalizationGet(token, opts.odata, !opts.v1);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    printJson(r.data);
  });

copilotCommand
  .command('settings-enhanced-personalization-patch')
  .description('PATCH /copilot/settings/people/enhancedPersonalization')
  .requiredOption('-f, --json-file <path>', 'JSON body')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { jsonFile: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }, cmd) => {
    checkReadOnly(cmd);
    const token = await resolveTokenOrExit(opts);
    const body = await readJsonFileOrExit(opts.jsonFile, '--json-file');
    const r = await copilotSettingsEnhancedPersonalizationPatch(token, body, !opts.v1);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    if (r.data !== undefined) printJson(r.data);
    else console.log('OK (204 No Content)');
  });

copilotCommand
  .command('settings-delete')
  .description('DELETE /copilot/settings; destructive')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--if-match <etag>', 'Optional If-Match header')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { beta?: boolean; v1?: boolean; ifMatch?: string; token?: string; identity?: string }, cmd) => {
    checkReadOnly(cmd);
    const token = await resolveTokenOrExit(opts);
    const r = await copilotSettingsDelete(token, !opts.v1, ifMatchHeader(opts.ifMatch));
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    if (r.data !== undefined) printJson(r.data);
    else console.log('OK (204 No Content)');
  });

copilotCommand
  .command('settings-people-delete')
  .description('DELETE /copilot/settings/people; destructive')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--if-match <etag>', 'Optional If-Match header')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { beta?: boolean; v1?: boolean; ifMatch?: string; token?: string; identity?: string }, cmd) => {
    checkReadOnly(cmd);
    const token = await resolveTokenOrExit(opts);
    const r = await copilotSettingsPeopleDelete(token, !opts.v1, ifMatchHeader(opts.ifMatch));
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    if (r.data !== undefined) printJson(r.data);
    else console.log('OK (204 No Content)');
  });

copilotCommand
  .command('settings-enhanced-personalization-delete')
  .description('DELETE /copilot/settings/people/enhancedPersonalization; destructive')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--if-match <etag>', 'Optional If-Match header')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { beta?: boolean; v1?: boolean; ifMatch?: string; token?: string; identity?: string }, cmd) => {
    checkReadOnly(cmd);
    const token = await resolveTokenOrExit(opts);
    const r = await copilotSettingsEnhancedPersonalizationDelete(token, !opts.v1, ifMatchHeader(opts.ifMatch));
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    if (r.data !== undefined) printJson(r.data);
    else console.log('OK (204 No Content)');
  });

copilotCommand
  .command('admin-settings-get')
  .description('GET /copilot/admin/settings (beta)')
  .option('--odata <query>', 'OData without leading ?')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { odata?: string; token?: string; identity?: string }) => {
    const token = await resolveTokenOrExit(opts);
    const r = await copilotAdminSettingsGet(token, opts.odata);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    printJson(r.data);
  });

copilotCommand
  .command('admin-settings-patch')
  .description('PATCH /copilot/admin/settings (beta)')
  .requiredOption('-f, --json-file <path>', 'JSON body')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { jsonFile: string; token?: string; identity?: string }, cmd) => {
    checkReadOnly(cmd);
    const token = await resolveTokenOrExit(opts);
    const body = await readJsonFileOrExit(opts.jsonFile, '--json-file');
    const r = await copilotAdminSettingsPatch(token, body);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    if (r.data !== undefined) printJson(r.data);
    else console.log('OK (204 No Content)');
  });

copilotCommand
  .command('admin-limited-mode-get')
  .description('GET /copilot/admin/settings/limitedMode (beta)')
  .option('--odata <query>', 'OData without leading ?')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { odata?: string; token?: string; identity?: string }) => {
    const token = await resolveTokenOrExit(opts);
    const r = await copilotAdminLimitedModeGet(token, opts.odata);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    printJson(r.data);
  });

copilotCommand
  .command('admin-limited-mode-patch')
  .description('PATCH /copilot/admin/settings/limitedMode (beta)')
  .requiredOption('-f, --json-file <path>', 'JSON body')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { jsonFile: string; token?: string; identity?: string }, cmd) => {
    checkReadOnly(cmd);
    const token = await resolveTokenOrExit(opts);
    const body = await readJsonFileOrExit(opts.jsonFile, '--json-file');
    const r = await copilotAdminLimitedModePatch(token, body);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    if (r.data !== undefined) printJson(r.data);
    else console.log('OK (204 No Content)');
  });

copilotCommand
  .command('admin-settings-delete')
  .description('DELETE /copilot/admin/settings (beta); destructive')
  .option('--if-match <etag>', 'Optional If-Match header')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { ifMatch?: string; token?: string; identity?: string }, cmd) => {
    checkReadOnly(cmd);
    const token = await resolveTokenOrExit(opts);
    const r = await copilotAdminSettingsDelete(token, ifMatchHeader(opts.ifMatch));
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    if (r.data !== undefined) printJson(r.data);
    else console.log('OK (204 No Content)');
  });

copilotCommand
  .command('admin-limited-mode-delete')
  .description('DELETE /copilot/admin/settings/limitedMode (beta); destructive')
  .option('--if-match <etag>', 'Optional If-Match header')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { ifMatch?: string; token?: string; identity?: string }, cmd) => {
    checkReadOnly(cmd);
    const token = await resolveTokenOrExit(opts);
    const r = await copilotAdminLimitedModeDelete(token, ifMatchHeader(opts.ifMatch));
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    if (r.data !== undefined) printJson(r.data);
    else console.log('OK (204 No Content)');
  });

copilotCommand
  .command('root-get')
  .description('GET /copilot — root Copilot resource')
  .option('--odata <query>', 'OData without leading ?')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { odata?: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }) => {
    const token = await resolveTokenOrExit(opts);
    const r = await copilotRootGet(token, opts.odata, !opts.v1);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    printJson(r.data);
  });

copilotCommand
  .command('root-patch')
  .description('PATCH /copilot')
  .requiredOption('-f, --json-file <path>', 'JSON body')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { jsonFile: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }, cmd) => {
    checkReadOnly(cmd);
    const token = await resolveTokenOrExit(opts);
    const body = await readJsonFileOrExit(opts.jsonFile, '--json-file');
    const r = await copilotRootPatch(token, body, !opts.v1);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    if (r.data !== undefined) printJson(r.data);
    else console.log('OK (204 No Content)');
  });

copilotCommand
  .command('admin-nav-get')
  .description('GET /copilot/admin (beta navigation)')
  .option('--odata <query>', 'OData without leading ?')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { odata?: string; token?: string; identity?: string }) => {
    const token = await resolveTokenOrExit(opts);
    const r = await copilotAdminNavGet(token, opts.odata);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    printJson(r.data);
  });

copilotCommand
  .command('admin-nav-patch')
  .description('PATCH /copilot/admin (beta)')
  .requiredOption('-f, --json-file <path>', 'JSON body')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { jsonFile: string; token?: string; identity?: string }, cmd) => {
    checkReadOnly(cmd);
    const token = await resolveTokenOrExit(opts);
    const body = await readJsonFileOrExit(opts.jsonFile, '--json-file');
    const r = await copilotAdminNavPatch(token, body);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    if (r.data !== undefined) printJson(r.data);
    else console.log('OK (204 No Content)');
  });

copilotCommand
  .command('admin-nav-delete')
  .description('DELETE /copilot/admin (beta); destructive')
  .option('--if-match <etag>', 'Optional If-Match header')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { ifMatch?: string; token?: string; identity?: string }, cmd) => {
    checkReadOnly(cmd);
    const token = await resolveTokenOrExit(opts);
    const r = await copilotAdminNavDelete(token, ifMatchHeader(opts.ifMatch));
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    if (r.data !== undefined) printJson(r.data);
    else console.log('OK (204 No Content)');
  });

copilotCommand
  .command('admin-catalog-get')
  .description('GET /copilot/admin/catalog (beta)')
  .option('--odata <query>', 'OData without leading ?')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { odata?: string; token?: string; identity?: string }) => {
    const token = await resolveTokenOrExit(opts);
    const r = await copilotAdminCatalogGet(token, opts.odata);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    printJson(r.data);
  });

copilotCommand
  .command('admin-catalog-patch')
  .description('PATCH /copilot/admin/catalog (beta)')
  .requiredOption('-f, --json-file <path>', 'JSON body')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { jsonFile: string; token?: string; identity?: string }, cmd) => {
    checkReadOnly(cmd);
    const token = await resolveTokenOrExit(opts);
    const body = await readJsonFileOrExit(opts.jsonFile, '--json-file');
    const r = await copilotAdminCatalogPatch(token, body);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    if (r.data !== undefined) printJson(r.data);
    else console.log('OK (204 No Content)');
  });

copilotCommand
  .command('admin-catalog-delete')
  .description('DELETE /copilot/admin/catalog (beta); destructive')
  .option('--if-match <etag>', 'Optional If-Match header')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { ifMatch?: string; token?: string; identity?: string }, cmd) => {
    checkReadOnly(cmd);
    const token = await resolveTokenOrExit(opts);
    const r = await copilotAdminCatalogDelete(token, ifMatchHeader(opts.ifMatch));
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    if (r.data !== undefined) printJson(r.data);
    else console.log('OK (204 No Content)');
  });

copilotCommand
  .command('communications-get')
  .description('GET /copilot/communications')
  .option('--odata <query>', 'OData without leading ?')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { odata?: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }) => {
    const token = await resolveTokenOrExit(opts);
    const r = await copilotCommunicationsGet(token, opts.odata, !opts.v1);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    printJson(r.data);
  });

copilotCommand
  .command('communications-patch')
  .description('PATCH /copilot/communications')
  .requiredOption('-f, --json-file <path>', 'JSON body')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { jsonFile: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }, cmd) => {
    checkReadOnly(cmd);
    const token = await resolveTokenOrExit(opts);
    const body = await readJsonFileOrExit(opts.jsonFile, '--json-file');
    const r = await copilotCommunicationsPatch(token, body, !opts.v1);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    if (r.data !== undefined) printJson(r.data);
    else console.log('OK (204 No Content)');
  });

copilotCommand
  .command('communications-delete')
  .description('DELETE /copilot/communications; destructive')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--if-match <etag>', 'Optional If-Match header')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { beta?: boolean; v1?: boolean; ifMatch?: string; token?: string; identity?: string }, cmd) => {
    checkReadOnly(cmd);
    const token = await resolveTokenOrExit(opts);
    const r = await copilotCommunicationsDelete(token, !opts.v1, ifMatchHeader(opts.ifMatch));
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    if (r.data !== undefined) printJson(r.data);
    else console.log('OK (204 No Content)');
  });

copilotCommand
  .command('interaction-history-nav-get')
  .description('GET /copilot/interactionHistory (tenant navigation; not the export function)')
  .option('--odata <query>', 'OData without leading ?')
  .option('--beta', 'Use Graph beta')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { odata?: string; beta?: boolean; token?: string; identity?: string }) => {
    const token = await resolveTokenOrExit(opts);
    const r = await copilotInteractionHistoryNavGet(token, opts.odata, Boolean(opts.beta));
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    printJson(r.data);
  });

copilotCommand
  .command('interaction-history-nav-patch')
  .description('PATCH /copilot/interactionHistory (tenant)')
  .requiredOption('-f, --json-file <path>', 'JSON body')
  .option('--beta', 'Use Graph beta')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { jsonFile: string; beta?: boolean; token?: string; identity?: string }, cmd) => {
    checkReadOnly(cmd);
    const token = await resolveTokenOrExit(opts);
    const body = await readJsonFileOrExit(opts.jsonFile, '--json-file');
    const r = await copilotInteractionHistoryNavPatch(token, body, Boolean(opts.beta));
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    if (r.data !== undefined) printJson(r.data);
    else console.log('OK (204 No Content)');
  });

copilotCommand
  .command('interaction-history-nav-delete')
  .description('DELETE /copilot/interactionHistory (tenant); destructive')
  .option('--beta', 'Use Graph beta')
  .option('--if-match <etag>', 'Optional If-Match header')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { beta?: boolean; ifMatch?: string; token?: string; identity?: string }, cmd) => {
    checkReadOnly(cmd);
    const token = await resolveTokenOrExit(opts);
    const r = await copilotInteractionHistoryNavDelete(token, Boolean(opts.beta), ifMatchHeader(opts.ifMatch));
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    if (r.data !== undefined) printJson(r.data);
    else console.log('OK (204 No Content)');
  });

copilotCommand
  .command('conversations-count')
  .description('GET /copilot/conversations/$count')
  .option('--odata <query>', 'OData without leading ? ($filter may need ConsistencyLevel)')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { odata?: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }) => {
    const token = await resolveTokenOrExit(opts);
    const r = await copilotConversationsCount(token, opts.odata, !opts.v1);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    printJson(r.data);
  });

copilotCommand
  .command('messages-count')
  .description('GET /copilot/conversations/{id}/messages/$count')
  .argument('<conversationId>', 'Conversation id')
  .option('--odata <query>', 'OData without leading ?')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      conversationId: string,
      opts: { odata?: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }
    ) => {
      const token = await resolveTokenOrExit(opts);
      const r = await copilotConversationMessagesCount(token, conversationId, opts.odata, !opts.v1);
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      printJson(r.data);
    }
  );

copilotCommand
  .command('agents-count')
  .description('GET /copilot/agents/$count')
  .option('--odata <query>', 'OData without leading ?')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { odata?: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }) => {
    const token = await resolveTokenOrExit(opts);
    const r = await copilotAgentsCount(token, opts.odata, !opts.v1);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    printJson(r.data);
  });

copilotCommand
  .command('packages-count')
  .description('GET /copilot/admin/catalog/packages/$count (beta)')
  .option('--odata <query>', 'OData without leading ?')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { odata?: string; token?: string; identity?: string }) => {
    const token = await resolveTokenOrExit(opts);
    const r = await copilotPackagesCount(token, opts.odata);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    printJson(r.data);
  });

copilotCommand
  .command('package-zip-delete')
  .description('DELETE /copilot/admin/catalog/packages/{id}/zipFile (beta)')
  .argument('<packageId>', 'Package id')
  .option('--if-match <etag>', 'Optional If-Match header')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (packageId: string, opts: AuthOpts & { ifMatch?: string }, cmd) => {
    checkReadOnly(cmd);
    const token = await resolveTokenOrExit(opts);
    const r = await copilotPackageZipDelete(token, packageId, ifMatchHeader(opts.ifMatch));
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    if (r.data !== undefined) printJson(r.data);
    else console.log('OK (204 No Content)');
  });

copilotCommand
  .command('interactions-export-tenant')
  .description(
    'GET /copilot/interactionHistory/getAllEnterpriseInteractions() — tenant-wide export (app-only AiEnterpriseInteraction.Read.All typical)'
  )
  .option('--odata <query>', 'OData without leading ?')
  .option('--beta', 'Use Graph beta')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { odata?: string; beta?: boolean; token?: string; identity?: string }) => {
    const token = await resolveTokenOrExit(opts);
    const r = await copilotInteractionsTenantExportList(token, opts.odata, Boolean(opts.beta));
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    printJson(r.data);
  });

/** POST /copilot/conversations */
copilotCommand
  .command('conversation-create')
  .description('POST /copilot/conversations — create an empty Copilot chat conversation (returns id)')
  .option('--v1', 'Use Graph v1.0 (default is beta, as in Microsoft docs)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { v1?: boolean; token?: string; identity?: string }, cmd) => {
    checkReadOnly(cmd);
    const token = await resolveTokenOrExit(opts);
    const r = await copilotConversationCreate(token, !opts.v1);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    printJson(r.data);
  });

/** POST .../microsoft.graph.copilot.chat */
copilotCommand
  .command('chat')
  .description(
    'POST /copilot/conversations/{id}/microsoft.graph.copilot.chat — synchronous Copilot message (requires locationHint)'
  )
  .argument('<conversationId>', 'Conversation id from conversation-create')
  .option('-m, --message <text>', 'User message text (required unless --json-file)')
  .option(
    '-z, --timezone <iana>',
    'locationHint.timeZone (e.g. America/New_York; required with --message unless --json-file)'
  )
  .option('-f, --json-file <path>', 'Full JSON body (message, locationHint, contextualResources, …)')
  .option('--v1', 'Use Graph v1.0 (default is beta)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      conversationId: string,
      opts: {
        message?: string;
        timezone?: string;
        jsonFile?: string;
        v1?: boolean;
        token?: string;
        identity?: string;
      },
      cmd
    ) => {
      checkReadOnly(cmd);
      const token = await resolveTokenOrExit(opts);
      let body: Record<string, unknown>;
      if (opts.jsonFile) {
        body = await readJsonFileOrExit(opts.jsonFile, '--json-file');
      } else {
        const text = (opts.message ?? '').trim();
        const tz = (opts.timezone ?? '').trim();
        if (!text || !tz) {
          console.error('Error: --message and --timezone are required unless --json-file is set');
          process.exit(1);
        }
        body = {
          message: { text },
          locationHint: { timeZone: tz }
        };
      }
      const r = await copilotConversationChat(token, conversationId, body, !opts.v1);
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      printJson(r.data);
    }
  );

/** POST .../microsoft.graph.copilot.chatOverStream */
copilotCommand
  .command('chat-stream')
  .description(
    'POST /copilot/conversations/{id}/microsoft.graph.copilot.chatOverStream — streamed SSE response (printed as raw text)'
  )
  .argument('<conversationId>', 'Conversation id')
  .option('-m, --message <text>', 'User message (required unless --json-file)')
  .option('-z, --timezone <iana>', 'locationHint.timeZone (required with --message unless --json-file)')
  .option('-f, --json-file <path>', 'Full JSON body')
  .option('--v1', 'Use Graph v1.0 (default is beta)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      conversationId: string,
      opts: {
        message?: string;
        timezone?: string;
        jsonFile?: string;
        v1?: boolean;
        token?: string;
        identity?: string;
      },
      cmd
    ) => {
      checkReadOnly(cmd);
      const token = await resolveTokenOrExit(opts);
      let body: Record<string, unknown>;
      if (opts.jsonFile) {
        body = await readJsonFileOrExit(opts.jsonFile, '--json-file');
      } else {
        const text = (opts.message ?? '').trim();
        const tz = (opts.timezone ?? '').trim();
        if (!text || !tz) {
          console.error('Error: --message and --timezone are required unless --json-file is set');
          process.exit(1);
        }
        body = {
          message: { text },
          locationHint: { timeZone: tz }
        };
      }
      const r = await copilotConversationChatOverStream(token, conversationId, body, !opts.v1);
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      console.log(r.data ?? '');
    }
  );

/** GET interaction export (application permission typical) */
copilotCommand
  .command('interactions-export')
  .description(
    'GET .../interactionHistory/getAllEnterpriseInteractions — export Copilot interactions for a user (app-only AiEnterpriseInteraction.Read.All typical; see Graph docs)'
  )
  .requiredOption('--user <id>', 'User id (GUID or UPN) in the path')
  .option('--odata <query>', 'OData query without leading ? (e.g. $top=100&$filter=... )')
  .option('--beta', 'Use Graph beta')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { user: string; odata?: string; beta?: boolean; token?: string; identity?: string }) => {
    const token = await resolveTokenOrExit(opts);
    const r = await copilotInteractionsExportList(token, opts.user, opts.odata, Boolean(opts.beta));
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    printJson(r.data);
  });

/** GET meeting AI insights list */
copilotCommand
  .command('meeting-insights-list')
  .description('GET /copilot/users/{user}/onlineMeetings/{meetingId}/aiInsights')
  .requiredOption('--user <id>', 'User id')
  .requiredOption('--meeting <id>', 'Online meeting id')
  .option('--odata <query>', 'OData without leading ? (e.g. $select=id)')
  .option('--beta', 'Use Graph beta')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (opts: {
      user: string;
      meeting: string;
      odata?: string;
      beta?: boolean;
      token?: string;
      identity?: string;
    }) => {
      const token = await resolveTokenOrExit(opts);
      const r = await copilotMeetingInsightsList(token, opts.user, opts.meeting, opts.odata, Boolean(opts.beta));
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      printJson(r.data);
    }
  );

/** GET single meeting insight */
copilotCommand
  .command('meeting-insight-get')
  .description('GET /copilot/users/{user}/onlineMeetings/{meetingId}/aiInsights/{insightId}')
  .requiredOption('--user <id>', 'User id')
  .requiredOption('--meeting <id>', 'Online meeting id')
  .requiredOption('--insight <id>', 'AI insight id')
  .option('--odata <query>', 'OData without leading ? (e.g. $select=meetingNotes)')
  .option('--beta', 'Use Graph beta')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (opts: {
      user: string;
      meeting: string;
      insight: string;
      odata?: string;
      beta?: boolean;
      token?: string;
      identity?: string;
    }) => {
      const token = await resolveTokenOrExit(opts);
      const r = await copilotMeetingInsightGet(
        token,
        opts.user,
        opts.meeting,
        opts.insight,
        opts.odata,
        Boolean(opts.beta)
      );
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      printJson(r.data);
    }
  );

copilotCommand
  .command('meeting-insights-count')
  .description('GET .../onlineMeetings/{meetingId}/aiInsights/$count')
  .requiredOption('--user <id>', 'User id')
  .requiredOption('--meeting <id>', 'Online meeting id')
  .option('--odata <query>', 'OData without leading ?')
  .option('--beta', 'Use Graph beta')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (opts: {
      user: string;
      meeting: string;
      odata?: string;
      beta?: boolean;
      token?: string;
      identity?: string;
    }) => {
      const token = await resolveTokenOrExit(opts);
      const r = await copilotMeetingAiInsightsCount(token, opts.user, opts.meeting, opts.odata, Boolean(opts.beta));
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      printJson(r.data);
    }
  );

copilotCommand
  .command('meeting-insight-create')
  .description('POST .../onlineMeetings/{meetingId}/aiInsights')
  .requiredOption('--user <id>', 'User id')
  .requiredOption('--meeting <id>', 'Online meeting id')
  .requiredOption('-f, --json-file <path>', 'JSON body')
  .option('--beta', 'Use Graph beta')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      opts: { user: string; meeting: string; jsonFile: string; beta?: boolean; token?: string; identity?: string },
      cmd
    ) => {
      checkReadOnly(cmd);
      const token = await resolveTokenOrExit(opts);
      const body = await readJsonFileOrExit(opts.jsonFile, '--json-file');
      const r = await copilotMeetingAiInsightsCreate(token, opts.user, opts.meeting, body, Boolean(opts.beta));
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      printJson(r.data);
    }
  );

copilotCommand
  .command('meeting-insight-patch')
  .description('PATCH .../aiInsights/{insightId}')
  .requiredOption('--user <id>', 'User id')
  .requiredOption('--meeting <id>', 'Online meeting id')
  .requiredOption('--insight <id>', 'AI insight id')
  .requiredOption('-f, --json-file <path>', 'JSON body')
  .option('--beta', 'Use Graph beta')
  .option('--if-match <etag>', 'Optional If-Match header')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      opts: {
        user: string;
        meeting: string;
        insight: string;
        jsonFile: string;
        beta?: boolean;
        ifMatch?: string;
        token?: string;
        identity?: string;
      },
      cmd
    ) => {
      checkReadOnly(cmd);
      const token = await resolveTokenOrExit(opts);
      const body = await readJsonFileOrExit(opts.jsonFile, '--json-file');
      const r = await copilotMeetingAiInsightPatch(
        token,
        opts.user,
        opts.meeting,
        opts.insight,
        body,
        Boolean(opts.beta),
        ifMatchHeader(opts.ifMatch)
      );
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      if (r.data !== undefined) printJson(r.data);
      else console.log('OK (204 No Content)');
    }
  );

copilotCommand
  .command('meeting-insight-delete')
  .description('DELETE .../aiInsights/{insightId}')
  .requiredOption('--user <id>', 'User id')
  .requiredOption('--meeting <id>', 'Online meeting id')
  .requiredOption('--insight <id>', 'AI insight id')
  .option('--beta', 'Use Graph beta')
  .option('--if-match <etag>', 'Optional If-Match header')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      opts: {
        user: string;
        meeting: string;
        insight: string;
        beta?: boolean;
        ifMatch?: string;
        token?: string;
        identity?: string;
      },
      cmd
    ) => {
      checkReadOnly(cmd);
      const token = await resolveTokenOrExit(opts);
      const r = await copilotMeetingAiInsightDelete(
        token,
        opts.user,
        opts.meeting,
        opts.insight,
        Boolean(opts.beta),
        ifMatchHeader(opts.ifMatch)
      );
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      if (r.data !== undefined) printJson(r.data);
      else console.log('OK (204 No Content)');
    }
  );

const reportsCmd = new Command('reports').description(
  'Copilot usage reports (GET /copilot/reports/...); requires Reports.Read.All and admin report reader roles where applicable'
);

reportsCmd.addCommand(
  new Command('user-count-summary')
    .description('getMicrosoft365CopilotUserCountSummary(period=...)')
    .requiredOption('-p, --period <code>', 'D7 | D30 | D90 | D180 | ALL')
    .option('--beta', 'Use Graph beta')
    .option('--v1', 'Use Graph v1.0')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity')
    .action(async (opts: { period: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }) => {
      const token = await resolveTokenOrExit(opts);
      let period: string;
      try {
        period = assertCopilotReportPeriod(opts.period);
      } catch (e) {
        console.error(e instanceof Error ? e.message : String(e));
        process.exit(1);
      }
      const beta = Boolean(opts.beta) && !opts.v1;
      const r = await copilotReportGet(token, 'getMicrosoft365CopilotUserCountSummary', period, beta);
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      printJson(r.data);
    })
);

reportsCmd.addCommand(
  new Command('user-count-trend')
    .description('getMicrosoft365CopilotUserCountTrend(period=...)')
    .requiredOption('-p, --period <code>', 'D7 | D30 | D90 | D180 | ALL')
    .option('--beta', 'Use Graph beta')
    .option('--v1', 'Use Graph v1.0')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity')
    .action(async (opts: { period: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }) => {
      const token = await resolveTokenOrExit(opts);
      let period: string;
      try {
        period = assertCopilotReportPeriod(opts.period);
      } catch (e) {
        console.error(e instanceof Error ? e.message : String(e));
        process.exit(1);
      }
      const beta = Boolean(opts.beta) && !opts.v1;
      const r = await copilotReportGet(token, 'getMicrosoft365CopilotUserCountTrend', period, beta);
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      printJson(r.data);
    })
);

reportsCmd.addCommand(
  new Command('usage-user-detail')
    .description('getMicrosoft365CopilotUsageUserDetail(period=...)')
    .requiredOption('-p, --period <code>', 'D7 | D30 | D90 | D180 | ALL')
    .option('--beta', 'Use Graph beta')
    .option('--v1', 'Use Graph v1.0')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity')
    .action(async (opts: { period: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }) => {
      const token = await resolveTokenOrExit(opts);
      let period: string;
      try {
        period = assertCopilotReportPeriod(opts.period);
      } catch (e) {
        console.error(e instanceof Error ? e.message : String(e));
        process.exit(1);
      }
      const beta = Boolean(opts.beta) && !opts.v1;
      const r = await copilotReportGet(token, 'getMicrosoft365CopilotUsageUserDetail', period, beta);
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      printJson(r.data);
    })
);

reportsCmd.addCommand(
  new Command('nav-get')
    .description('GET /copilot/reports (navigation; distinct from usage functions)')
    .option('--odata <query>', 'OData without leading ?')
    .option('--beta', 'Use Graph beta')
    .option('--v1', 'Use Graph v1.0')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity')
    .action(async (opts: { odata?: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }) => {
      const token = await resolveTokenOrExit(opts);
      const beta = Boolean(opts.beta) && !opts.v1;
      const r = await copilotReportsNavGet(token, opts.odata, beta);
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      printJson(r.data);
    })
);

reportsCmd.addCommand(
  new Command('nav-patch')
    .description('PATCH /copilot/reports (navigation)')
    .requiredOption('-f, --json-file <path>', 'JSON body')
    .option('--beta', 'Use Graph beta')
    .option('--v1', 'Use Graph v1.0')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity')
    .action(
      async (opts: { jsonFile: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }, cmd) => {
        checkReadOnly(cmd);
        const token = await resolveTokenOrExit(opts);
        const body = await readJsonFileOrExit(opts.jsonFile, '--json-file');
        const beta = Boolean(opts.beta) && !opts.v1;
        const r = await copilotReportsNavPatch(token, body, beta);
        if (!r.ok) exitGraphError('Error: ', r.error?.message);
        if (r.data !== undefined) printJson(r.data);
        else console.log('OK (204 No Content)');
      }
    )
);

reportsCmd.addCommand(
  new Command('nav-delete')
    .description('DELETE /copilot/reports; destructive')
    .option('--beta', 'Use Graph beta')
    .option('--v1', 'Use Graph v1.0')
    .option('--if-match <etag>', 'Optional If-Match header')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity')
    .action(
      async (opts: { beta?: boolean; v1?: boolean; ifMatch?: string; token?: string; identity?: string }, cmd) => {
        checkReadOnly(cmd);
        const token = await resolveTokenOrExit(opts);
        const beta = Boolean(opts.beta) && !opts.v1;
        const r = await copilotReportsNavDelete(token, beta, ifMatchHeader(opts.ifMatch));
        if (!r.ok) exitGraphError('Error: ', r.error?.message);
        if (r.data !== undefined) printJson(r.data);
        else console.log('OK (204 No Content)');
      }
    )
);

copilotCommand.addCommand(reportsCmd);

const packagesCmd = new Command('packages').description(
  'Copilot package catalog (beta /copilot/admin/catalog/packages); CopilotPackages.Read*. See Microsoft Agent 365 licensing.'
);

packagesCmd
  .command('list')
  .description('GET /copilot/admin/catalog/packages')
  .option('--odata <query>', 'OData without leading ?')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { odata?: string; token?: string; identity?: string }) => {
    const token = await resolveTokenOrExit(opts);
    const r = await copilotPackagesList(token, opts.odata);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    printJson(r.data);
  });

packagesCmd
  .command('create')
  .description('POST /copilot/admin/catalog/packages (beta)')
  .requiredOption('-f, --json-file <path>', 'JSON body for new package')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { jsonFile: string; token?: string; identity?: string }, cmd) => {
    checkReadOnly(cmd);
    const token = await resolveTokenOrExit(opts);
    const body = await readJsonFileOrExit(opts.jsonFile, '--json-file');
    const r = await copilotPackagesCreate(token, body);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    printJson(r.data);
  });

packagesCmd
  .command('get')
  .description('GET /copilot/admin/catalog/packages/{id}')
  .argument('<packageId>', 'Package id (e.g. P_...)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (packageId: string, opts: AuthOpts) => {
    const token = await resolveTokenOrExit(opts);
    const r = await copilotPackagesGet(token, packageId);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    printJson(r.data);
  });

packagesCmd
  .command('update')
  .description('PATCH /copilot/admin/catalog/packages/{id}')
  .argument('<packageId>', 'Package id')
  .requiredOption('-f, --json-file <path>', 'JSON body (allowedUsersAndGroups, acquireUsersAndGroups, …)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (packageId: string, opts: AuthOpts & { jsonFile: string }, cmd) => {
    checkReadOnly(cmd);
    const token = await resolveTokenOrExit(opts);
    const body = await readJsonFileOrExit(opts.jsonFile, '--json-file');
    const r = await copilotPackagesUpdate(token, packageId, body);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    if (r.data !== undefined) printJson(r.data);
    else console.log('OK (204 No Content)');
  });

packagesCmd
  .command('block')
  .description('POST /copilot/admin/catalog/packages/{id}/block')
  .argument('<packageId>', 'Package id')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (packageId: string, opts: AuthOpts, cmd) => {
    checkReadOnly(cmd);
    const token = await resolveTokenOrExit(opts);
    const r = await copilotPackagesBlock(token, packageId);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    if (r.data !== undefined) printJson(r.data);
    else console.log('OK (204 No Content)');
  });

packagesCmd
  .command('unblock')
  .description('POST /copilot/admin/catalog/packages/{id}/unblock')
  .argument('<packageId>', 'Package id')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (packageId: string, opts: AuthOpts, cmd) => {
    checkReadOnly(cmd);
    const token = await resolveTokenOrExit(opts);
    const r = await copilotPackagesUnblock(token, packageId);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    if (r.data !== undefined) printJson(r.data);
    else console.log('OK (204 No Content)');
  });

packagesCmd
  .command('reassign')
  .description('POST /copilot/admin/catalog/packages/{id}/reassign')
  .argument('<packageId>', 'Package id')
  .requiredOption('--new-owner-user-id <guid>', 'New owner user id')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (packageId: string, opts: AuthOpts & { newOwnerUserId: string }, cmd) => {
    checkReadOnly(cmd);
    const token = await resolveTokenOrExit(opts);
    const r = await copilotPackagesReassign(token, packageId, opts.newOwnerUserId);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    if (r.data !== undefined) printJson(r.data);
    else console.log('OK (204 No Content)');
  });

packagesCmd
  .command('delete')
  .description('DELETE /copilot/admin/catalog/packages/{id} (beta)')
  .argument('<packageId>', 'Package id')
  .option('--if-match <etag>', 'Optional If-Match header')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (packageId: string, opts: AuthOpts & { ifMatch?: string }, cmd) => {
    checkReadOnly(cmd);
    const token = await resolveTokenOrExit(opts);
    const r = await copilotPackagesDelete(token, packageId, ifMatchHeader(opts.ifMatch));
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    if (r.data !== undefined) printJson(r.data);
    else console.log('OK (204 No Content)');
  });

packagesCmd
  .command('zip-download')
  .description('GET /copilot/admin/catalog/packages/{id}/zipFile — write binary to --output (beta)')
  .argument('<packageId>', 'Package id')
  .requiredOption('-o, --output <path>', 'Local file path for .zip bytes')
  .option('--force', 'Overwrite the output file if it already exists')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (packageId: string, opts: AuthOpts & { output: string; force?: boolean }) => {
    const token = await resolveTokenOrExit(opts);
    const outPath = resolve(process.cwd(), opts.output.trim());
    if (!opts.force) {
      const exists = await access(outPath).then(
        () => true,
        () => false
      );
      if (exists) {
        console.error(`Error: ${opts.output} already exists. Pass --force to overwrite.`);
        process.exit(1);
      }
    }
    const r = await copilotPackageZipDownload(token, packageId);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    if (!r.data) exitGraphError('Error: ', 'Empty response');
    await writeFile(outPath, r.data);
    console.log(`Wrote ${r.data.byteLength} bytes to ${opts.output}`);
  });

packagesCmd
  .command('zip-upload')
  .description('PUT /copilot/admin/catalog/packages/{id}/zipFile (beta)')
  .argument('<packageId>', 'Package id')
  .requiredOption('--file <path>', 'Local .zip file to upload')
  .option('--content-type <mime>', 'Content-Type header (default: application/zip)', 'application/zip')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (packageId: string, opts: AuthOpts & { file: string; contentType?: string }, cmd) => {
    checkReadOnly(cmd);
    const token = await resolveTokenOrExit(opts);
    const bytes = new Uint8Array(await readFile(resolve(process.cwd(), opts.file.trim())));
    const r = await copilotPackageZipUpload(token, packageId, bytes, opts.contentType ?? 'application/zip');
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    if (r.data !== undefined) printJson(r.data);
    else console.log('OK (204 No Content)');
  });

copilotCommand.addCommand(packagesCmd);

const activityFeedCmd = new Command('activity-feed').description(
  'Copilot realtime activity feed (beta /copilot/communications/realtimeActivityFeed/...) — meetings, subscriptions, transcripts'
);

activityFeedCmd
  .command('get')
  .description('GET /copilot/communications/realtimeActivityFeed')
  .option('--odata <query>', 'OData without leading ?')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { odata?: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }) => {
    const token = await resolveTokenOrExit(opts);
    const r = await copilotRealtimeActivityFeedGet(token, opts.odata, !opts.v1);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    printJson(r.data);
  });

activityFeedCmd
  .command('patch-root')
  .description('PATCH /copilot/communications/realtimeActivityFeed')
  .requiredOption('-f, --json-file <path>', 'JSON body')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--if-match <etag>', 'Optional If-Match header')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      opts: {
        jsonFile: string;
        beta?: boolean;
        v1?: boolean;
        ifMatch?: string;
        token?: string;
        identity?: string;
      },
      cmd
    ) => {
      checkReadOnly(cmd);
      const token = await resolveTokenOrExit(opts);
      const body = await readJsonFileOrExit(opts.jsonFile, '--json-file');
      const r = await copilotRealtimeActivityFeedPatch(token, body, !opts.v1, ifMatchHeader(opts.ifMatch));
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      if (r.data !== undefined) printJson(r.data);
      else console.log('OK (204 No Content)');
    }
  );

activityFeedCmd
  .command('delete-root')
  .description('DELETE /copilot/communications/realtimeActivityFeed; destructive')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--if-match <etag>', 'Optional If-Match header')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { beta?: boolean; v1?: boolean; ifMatch?: string; token?: string; identity?: string }, cmd) => {
    checkReadOnly(cmd);
    const token = await resolveTokenOrExit(opts);
    const r = await copilotRealtimeActivityFeedDelete(token, !opts.v1, ifMatchHeader(opts.ifMatch));
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    if (r.data !== undefined) printJson(r.data);
    else console.log('OK (204 No Content)');
  });

activityFeedCmd
  .command('meetings-list')
  .description('GET .../realtimeActivityFeed/meetings')
  .option('--odata <query>', 'OData without leading ?')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { odata?: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }) => {
    const token = await resolveTokenOrExit(opts);
    const r = await copilotRealtimeMeetingsList(token, opts.odata, !opts.v1);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    printJson(r.data);
  });

activityFeedCmd
  .command('meetings-count')
  .description('GET .../realtimeActivityFeed/meetings/$count')
  .option('--odata <query>', 'OData without leading ?')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { odata?: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }) => {
    const token = await resolveTokenOrExit(opts);
    const r = await copilotRealtimeMeetingsCount(token, opts.odata, !opts.v1);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    printJson(r.data);
  });

activityFeedCmd
  .command('meeting-create')
  .description('POST .../realtimeActivityFeed/meetings')
  .requiredOption('-f, --json-file <path>', 'JSON body')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { jsonFile: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }, cmd) => {
    checkReadOnly(cmd);
    const token = await resolveTokenOrExit(opts);
    const body = await readJsonFileOrExit(opts.jsonFile, '--json-file');
    const r = await copilotRealtimeMeetingCreate(token, body, !opts.v1);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    printJson(r.data);
  });

activityFeedCmd
  .command('meeting-get')
  .description('GET .../realtimeActivityFeed/meetings/{id}')
  .argument('<meetingId>', 'Realtime activity meeting id')
  .option('--odata <query>', 'OData without leading ?')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      meetingId: string,
      opts: { odata?: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }
    ) => {
      const token = await resolveTokenOrExit(opts);
      const r = await copilotRealtimeMeetingGet(token, meetingId, opts.odata, !opts.v1);
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      printJson(r.data);
    }
  );

activityFeedCmd
  .command('meeting-patch')
  .description('PATCH .../realtimeActivityFeed/meetings/{id}')
  .argument('<meetingId>', 'Realtime activity meeting id')
  .requiredOption('-f, --json-file <path>', 'JSON body')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--if-match <etag>', 'Optional If-Match header')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      meetingId: string,
      opts: {
        jsonFile: string;
        beta?: boolean;
        v1?: boolean;
        ifMatch?: string;
        token?: string;
        identity?: string;
      },
      cmd
    ) => {
      checkReadOnly(cmd);
      const token = await resolveTokenOrExit(opts);
      const body = await readJsonFileOrExit(opts.jsonFile, '--json-file');
      const r = await copilotRealtimeMeetingPatch(token, meetingId, body, !opts.v1, ifMatchHeader(opts.ifMatch));
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      if (r.data !== undefined) printJson(r.data);
      else console.log('OK (204 No Content)');
    }
  );

activityFeedCmd
  .command('meeting-delete')
  .description('DELETE .../realtimeActivityFeed/meetings/{id}')
  .argument('<meetingId>', 'Realtime activity meeting id')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--if-match <etag>', 'Optional If-Match header')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      meetingId: string,
      opts: { beta?: boolean; v1?: boolean; ifMatch?: string; token?: string; identity?: string },
      cmd
    ) => {
      checkReadOnly(cmd);
      const token = await resolveTokenOrExit(opts);
      const r = await copilotRealtimeMeetingDelete(token, meetingId, !opts.v1, ifMatchHeader(opts.ifMatch));
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      if (r.data !== undefined) printJson(r.data);
      else console.log('OK (204 No Content)');
    }
  );

activityFeedCmd
  .command('subscriptions-list')
  .description('GET .../realtimeActivityFeed/multiActivitySubscriptions')
  .option('--odata <query>', 'OData without leading ?')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { odata?: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }) => {
    const token = await resolveTokenOrExit(opts);
    const r = await copilotRealtimeSubscriptionsList(token, opts.odata, !opts.v1);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    printJson(r.data);
  });

activityFeedCmd
  .command('subscriptions-count')
  .description('GET .../multiActivitySubscriptions/$count')
  .option('--odata <query>', 'OData without leading ?')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { odata?: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }) => {
    const token = await resolveTokenOrExit(opts);
    const r = await copilotRealtimeSubscriptionsCount(token, opts.odata, !opts.v1);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    printJson(r.data);
  });

activityFeedCmd
  .command('subscription-create')
  .description('POST .../realtimeActivityFeed/multiActivitySubscriptions')
  .requiredOption('-f, --json-file <path>', 'JSON body')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { jsonFile: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }, cmd) => {
    checkReadOnly(cmd);
    const token = await resolveTokenOrExit(opts);
    const body = await readJsonFileOrExit(opts.jsonFile, '--json-file');
    const r = await copilotRealtimeSubscriptionCreate(token, body, !opts.v1);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    printJson(r.data);
  });

activityFeedCmd
  .command('subscription-get')
  .description('GET .../multiActivitySubscriptions/{id}')
  .argument('<subscriptionId>', 'Subscription id')
  .option('--odata <query>', 'OData without leading ?')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      subscriptionId: string,
      opts: { odata?: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }
    ) => {
      const token = await resolveTokenOrExit(opts);
      const r = await copilotRealtimeSubscriptionGet(token, subscriptionId, opts.odata, !opts.v1);
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      printJson(r.data);
    }
  );

activityFeedCmd
  .command('subscription-patch')
  .description('PATCH .../multiActivitySubscriptions/{id}')
  .argument('<subscriptionId>', 'Subscription id')
  .requiredOption('-f, --json-file <path>', 'JSON body')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--if-match <etag>', 'Optional If-Match header')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      subscriptionId: string,
      opts: {
        jsonFile: string;
        beta?: boolean;
        v1?: boolean;
        ifMatch?: string;
        token?: string;
        identity?: string;
      },
      cmd
    ) => {
      checkReadOnly(cmd);
      const token = await resolveTokenOrExit(opts);
      const body = await readJsonFileOrExit(opts.jsonFile, '--json-file');
      const r = await copilotRealtimeSubscriptionPatch(
        token,
        subscriptionId,
        body,
        !opts.v1,
        ifMatchHeader(opts.ifMatch)
      );
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      if (r.data !== undefined) printJson(r.data);
      else console.log('OK (204 No Content)');
    }
  );

activityFeedCmd
  .command('subscription-delete')
  .description('DELETE .../multiActivitySubscriptions/{id}')
  .argument('<subscriptionId>', 'Subscription id')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--if-match <etag>', 'Optional If-Match header')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      subscriptionId: string,
      opts: { beta?: boolean; v1?: boolean; ifMatch?: string; token?: string; identity?: string },
      cmd
    ) => {
      checkReadOnly(cmd);
      const token = await resolveTokenOrExit(opts);
      const r = await copilotRealtimeSubscriptionDelete(token, subscriptionId, !opts.v1, ifMatchHeader(opts.ifMatch));
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      if (r.data !== undefined) printJson(r.data);
      else console.log('OK (204 No Content)');
    }
  );

activityFeedCmd
  .command('subscription-get-artifacts')
  .description('POST .../multiActivitySubscriptions/{id}/getArtifacts')
  .argument('<subscriptionId>', 'Subscription id')
  .option('-f, --json-file <path>', 'Optional JSON body (default {})')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      subscriptionId: string,
      opts: { jsonFile?: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }
    ) => {
      const token = await resolveTokenOrExit(opts);
      let body: Record<string, unknown> | undefined;
      if (opts.jsonFile?.trim()) {
        body = await readJsonFileOrExit(opts.jsonFile, '--json-file');
      }
      const r = await copilotRealtimeSubscriptionGetArtifacts(token, subscriptionId, body, !opts.v1);
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      printJson(r.data);
    }
  );

activityFeedCmd
  .command('transcripts-list')
  .description('GET .../meetings/{meetingId}/transcripts')
  .argument('<meetingId>', 'Realtime activity meeting id')
  .option('--odata <query>', 'OData without leading ?')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      meetingId: string,
      opts: { odata?: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }
    ) => {
      const token = await resolveTokenOrExit(opts);
      const r = await copilotRealtimeTranscriptsList(token, meetingId, opts.odata, !opts.v1);
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      printJson(r.data);
    }
  );

activityFeedCmd
  .command('transcripts-count')
  .description('GET .../meetings/{meetingId}/transcripts/$count')
  .argument('<meetingId>', 'Realtime activity meeting id')
  .option('--odata <query>', 'OData without leading ?')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      meetingId: string,
      opts: { odata?: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }
    ) => {
      const token = await resolveTokenOrExit(opts);
      const r = await copilotRealtimeTranscriptsCount(token, meetingId, opts.odata, !opts.v1);
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      printJson(r.data);
    }
  );

activityFeedCmd
  .command('transcript-create')
  .description('POST .../meetings/{meetingId}/transcripts')
  .argument('<meetingId>', 'Realtime activity meeting id')
  .requiredOption('-f, --json-file <path>', 'JSON body')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      meetingId: string,
      opts: { jsonFile: string; beta?: boolean; v1?: boolean; token?: string; identity?: string },
      cmd
    ) => {
      checkReadOnly(cmd);
      const token = await resolveTokenOrExit(opts);
      const body = await readJsonFileOrExit(opts.jsonFile, '--json-file');
      const r = await copilotRealtimeTranscriptCreate(token, meetingId, body, !opts.v1);
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      printJson(r.data);
    }
  );

activityFeedCmd
  .command('transcript-get')
  .description('GET .../meetings/{meetingId}/transcripts/{transcriptId}')
  .argument('<meetingId>', 'Realtime activity meeting id')
  .argument('<transcriptId>', 'Transcript id')
  .option('--odata <query>', 'OData without leading ?')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      meetingId: string,
      transcriptId: string,
      opts: { odata?: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }
    ) => {
      const token = await resolveTokenOrExit(opts);
      const r = await copilotRealtimeTranscriptGet(token, meetingId, transcriptId, opts.odata, !opts.v1);
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      printJson(r.data);
    }
  );

activityFeedCmd
  .command('transcript-patch')
  .description('PATCH .../meetings/{meetingId}/transcripts/{transcriptId}')
  .argument('<meetingId>', 'Realtime activity meeting id')
  .argument('<transcriptId>', 'Transcript id')
  .requiredOption('-f, --json-file <path>', 'JSON body')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--if-match <etag>', 'Optional If-Match header')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      meetingId: string,
      transcriptId: string,
      opts: {
        jsonFile: string;
        beta?: boolean;
        v1?: boolean;
        ifMatch?: string;
        token?: string;
        identity?: string;
      },
      cmd
    ) => {
      checkReadOnly(cmd);
      const token = await resolveTokenOrExit(opts);
      const body = await readJsonFileOrExit(opts.jsonFile, '--json-file');
      const r = await copilotRealtimeTranscriptPatch(
        token,
        meetingId,
        transcriptId,
        body,
        !opts.v1,
        ifMatchHeader(opts.ifMatch)
      );
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      if (r.data !== undefined) printJson(r.data);
      else console.log('OK (204 No Content)');
    }
  );

activityFeedCmd
  .command('transcript-delete')
  .description('DELETE .../meetings/{meetingId}/transcripts/{transcriptId}')
  .argument('<meetingId>', 'Realtime activity meeting id')
  .argument('<transcriptId>', 'Transcript id')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--if-match <etag>', 'Optional If-Match header')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      meetingId: string,
      transcriptId: string,
      opts: { beta?: boolean; v1?: boolean; ifMatch?: string; token?: string; identity?: string },
      cmd
    ) => {
      checkReadOnly(cmd);
      const token = await resolveTokenOrExit(opts);
      const r = await copilotRealtimeTranscriptDelete(
        token,
        meetingId,
        transcriptId,
        !opts.v1,
        ifMatchHeader(opts.ifMatch)
      );
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      if (r.data !== undefined) printJson(r.data);
      else console.log('OK (204 No Content)');
    }
  );

const aiUserCmd = new Command('ai-user').description(
  'Graph /copilot/users (Copilot aiUser resources): CRUD, per-user interactionHistory, onlineMeetings'
);

aiUserCmd
  .command('list')
  .description('GET /copilot/users')
  .option('--odata <query>', 'OData without leading ?')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { odata?: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }) => {
    const token = await resolveTokenOrExit(opts);
    const r = await copilotAiUsersList(token, opts.odata, !opts.v1);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    printJson(r.data);
  });

aiUserCmd
  .command('count')
  .description('GET /copilot/users/$count')
  .option('--odata <query>', 'OData without leading ?')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { odata?: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }) => {
    const token = await resolveTokenOrExit(opts);
    const r = await copilotAiUsersCount(token, opts.odata, !opts.v1);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    printJson(r.data);
  });

aiUserCmd
  .command('create')
  .description('POST /copilot/users')
  .requiredOption('-f, --json-file <path>', 'JSON body')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { jsonFile: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }, cmd) => {
    checkReadOnly(cmd);
    const token = await resolveTokenOrExit(opts);
    const body = await readJsonFileOrExit(opts.jsonFile, '--json-file');
    const r = await copilotAiUserCreate(token, body, !opts.v1);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    printJson(r.data);
  });

aiUserCmd
  .command('get')
  .description('GET /copilot/users/{aiUserId}')
  .argument('<aiUserId>', 'aiUser id')
  .option('--odata <query>', 'OData without leading ?')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      aiUserId: string,
      opts: { odata?: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }
    ) => {
      const token = await resolveTokenOrExit(opts);
      const r = await copilotAiUserGet(token, aiUserId, opts.odata, !opts.v1);
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      printJson(r.data);
    }
  );

aiUserCmd
  .command('patch')
  .description('PATCH /copilot/users/{aiUserId}')
  .argument('<aiUserId>', 'aiUser id')
  .requiredOption('-f, --json-file <path>', 'JSON body')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--if-match <etag>', 'Optional If-Match header')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      aiUserId: string,
      opts: {
        jsonFile: string;
        beta?: boolean;
        v1?: boolean;
        ifMatch?: string;
        token?: string;
        identity?: string;
      },
      cmd
    ) => {
      checkReadOnly(cmd);
      const token = await resolveTokenOrExit(opts);
      const body = await readJsonFileOrExit(opts.jsonFile, '--json-file');
      const r = await copilotAiUserPatch(token, aiUserId, body, !opts.v1, ifMatchHeader(opts.ifMatch));
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      if (r.data !== undefined) printJson(r.data);
      else console.log('OK (204 No Content)');
    }
  );

aiUserCmd
  .command('delete')
  .description('DELETE /copilot/users/{aiUserId}')
  .argument('<aiUserId>', 'aiUser id')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--if-match <etag>', 'Optional If-Match header')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      aiUserId: string,
      opts: { beta?: boolean; v1?: boolean; ifMatch?: string; token?: string; identity?: string },
      cmd
    ) => {
      checkReadOnly(cmd);
      const token = await resolveTokenOrExit(opts);
      const r = await copilotAiUserDelete(token, aiUserId, !opts.v1, ifMatchHeader(opts.ifMatch));
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      if (r.data !== undefined) printJson(r.data);
      else console.log('OK (204 No Content)');
    }
  );

aiUserCmd
  .command('interaction-history-get')
  .description('GET /copilot/users/{user}/interactionHistory (navigation)')
  .requiredOption('--user <id>', 'aiUser id in path')
  .option('--odata <query>', 'OData without leading ?')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (opts: { user: string; odata?: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }) => {
      const token = await resolveTokenOrExit(opts);
      const r = await copilotAiUserInteractionHistoryGet(token, opts.user, opts.odata, !opts.v1);
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      printJson(r.data);
    }
  );

aiUserCmd
  .command('interaction-history-patch')
  .description('PATCH /copilot/users/{user}/interactionHistory')
  .requiredOption('--user <id>', 'aiUser id in path')
  .requiredOption('-f, --json-file <path>', 'JSON body')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      opts: { user: string; jsonFile: string; beta?: boolean; v1?: boolean; token?: string; identity?: string },
      cmd
    ) => {
      checkReadOnly(cmd);
      const token = await resolveTokenOrExit(opts);
      const body = await readJsonFileOrExit(opts.jsonFile, '--json-file');
      const r = await copilotAiUserInteractionHistoryPatch(token, opts.user, body, !opts.v1);
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      if (r.data !== undefined) printJson(r.data);
      else console.log('OK (204 No Content)');
    }
  );

aiUserCmd
  .command('interaction-history-delete')
  .description('DELETE /copilot/users/{user}/interactionHistory; destructive')
  .requiredOption('--user <id>', 'aiUser id in path')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--if-match <etag>', 'Optional If-Match header')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      opts: { user: string; beta?: boolean; v1?: boolean; ifMatch?: string; token?: string; identity?: string },
      cmd
    ) => {
      checkReadOnly(cmd);
      const token = await resolveTokenOrExit(opts);
      const r = await copilotAiUserInteractionHistoryDelete(token, opts.user, !opts.v1, ifMatchHeader(opts.ifMatch));
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      if (r.data !== undefined) printJson(r.data);
      else console.log('OK (204 No Content)');
    }
  );

aiUserCmd
  .command('meetings-list')
  .description('GET /copilot/users/{user}/onlineMeetings')
  .requiredOption('--user <id>', 'aiUser id in path')
  .option('--odata <query>', 'OData without leading ?')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (opts: { user: string; odata?: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }) => {
      const token = await resolveTokenOrExit(opts);
      const r = await copilotAiUserOnlineMeetingsList(token, opts.user, opts.odata, !opts.v1);
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      printJson(r.data);
    }
  );

aiUserCmd
  .command('meetings-count')
  .description('GET /copilot/users/{user}/onlineMeetings/$count')
  .requiredOption('--user <id>', 'aiUser id in path')
  .option('--odata <query>', 'OData without leading ?')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (opts: { user: string; odata?: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }) => {
      const token = await resolveTokenOrExit(opts);
      const r = await copilotAiUserOnlineMeetingsCount(token, opts.user, opts.odata, !opts.v1);
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      printJson(r.data);
    }
  );

aiUserCmd
  .command('meeting-create')
  .description('POST /copilot/users/{user}/onlineMeetings')
  .requiredOption('--user <id>', 'aiUser id in path')
  .requiredOption('-f, --json-file <path>', 'JSON body')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      opts: { user: string; jsonFile: string; beta?: boolean; v1?: boolean; token?: string; identity?: string },
      cmd
    ) => {
      checkReadOnly(cmd);
      const token = await resolveTokenOrExit(opts);
      const body = await readJsonFileOrExit(opts.jsonFile, '--json-file');
      const r = await copilotAiUserOnlineMeetingCreate(token, opts.user, body, !opts.v1);
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      printJson(r.data);
    }
  );

aiUserCmd
  .command('meeting-get')
  .description('GET /copilot/users/{user}/onlineMeetings/{id}')
  .requiredOption('--user <id>', 'aiUser id in path')
  .argument('<onlineMeetingId>', 'Online meeting id')
  .option('--odata <query>', 'OData without leading ?')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      onlineMeetingId: string,
      opts: { user: string; odata?: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }
    ) => {
      const token = await resolveTokenOrExit(opts);
      const r = await copilotAiUserOnlineMeetingGet(token, opts.user, onlineMeetingId, opts.odata, !opts.v1);
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      printJson(r.data);
    }
  );

aiUserCmd
  .command('meeting-patch')
  .description('PATCH /copilot/users/{user}/onlineMeetings/{id}')
  .requiredOption('--user <id>', 'aiUser id in path')
  .argument('<onlineMeetingId>', 'Online meeting id')
  .requiredOption('-f, --json-file <path>', 'JSON body')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--if-match <etag>', 'Optional If-Match header')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      onlineMeetingId: string,
      opts: {
        user: string;
        jsonFile: string;
        beta?: boolean;
        v1?: boolean;
        ifMatch?: string;
        token?: string;
        identity?: string;
      },
      cmd
    ) => {
      checkReadOnly(cmd);
      const token = await resolveTokenOrExit(opts);
      const body = await readJsonFileOrExit(opts.jsonFile, '--json-file');
      const r = await copilotAiUserOnlineMeetingPatch(
        token,
        opts.user,
        onlineMeetingId,
        body,
        !opts.v1,
        ifMatchHeader(opts.ifMatch)
      );
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      if (r.data !== undefined) printJson(r.data);
      else console.log('OK (204 No Content)');
    }
  );

aiUserCmd
  .command('meeting-delete')
  .description('DELETE /copilot/users/{user}/onlineMeetings/{id}')
  .requiredOption('--user <id>', 'aiUser id in path')
  .argument('<onlineMeetingId>', 'Online meeting id')
  .option('--beta', 'Use Graph beta (default)')
  .option('--v1', 'Use Graph v1.0')
  .option('--if-match <etag>', 'Optional If-Match header')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      onlineMeetingId: string,
      opts: { user: string; beta?: boolean; v1?: boolean; ifMatch?: string; token?: string; identity?: string },
      cmd
    ) => {
      checkReadOnly(cmd);
      const token = await resolveTokenOrExit(opts);
      const r = await copilotAiUserOnlineMeetingDelete(
        token,
        opts.user,
        onlineMeetingId,
        !opts.v1,
        ifMatchHeader(opts.ifMatch)
      );
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      if (r.data !== undefined) printJson(r.data);
      else console.log('OK (204 No Content)');
    }
  );

copilotCommand.addCommand(aiUserCmd);
copilotCommand.addCommand(activityFeedCmd);

copilotCommand
  .command('notify-help')
  .description(
    'Print Copilot change-notification resource paths (use `subscribe` with --url and a JSON file for encrypted payloads)'
  )
  .action(() => {
    console.log(`Copilot AI interactions (per-user, delegated AiEnterpriseInteraction.Read — include resource data needs encryption):
  /copilot/users/{user-id}/interactionHistory/getAllEnterpriseInteractions()
  Optional OData on resource string, e.g. ?$filter=appClass eq 'IPM.SkypeTeams.Message.Copilot.Teams'

Tenant-wide (application AiEnterpriseInteraction.Read.All):
  /copilot/interactionHistory/getAllEnterpriseInteractions()

Meeting AI insights are listed per online meeting (GET; OnlineMeetingAiInsight.Read.All):
  /copilot/users/{user-id}/onlineMeetings/{online-meeting-id}/aiInsights
  See: https://learn.microsoft.com/graph/api/onlinemeeting-list-aiinsights

Also: m365-agent-cli subscribe copilot-interactions --user <id> --url <webhook> (shortcut; see subscribe --help)
`);
  });
