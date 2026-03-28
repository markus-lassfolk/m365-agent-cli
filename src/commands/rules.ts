import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import {
  listMessageRules,
  getMessageRule,
  createMessageRule,
  updateMessageRule,
  deleteMessageRule,
  type MessageRule,
  type CreateMessageRulePayload,
  type UpdateMessageRulePayload,
  type MessageRuleCondition,
  type MessageRuleAction
} from '../lib/rules-client.js';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function parseEmailAddresses(raw: string): { emailAddress: { name?: string; address: string } }[] {
  try {
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed)) {
      return parsed.map((item) => (typeof item === 'string' ? { emailAddress: { address: item } } : item));
    }
    return parsed;
  } catch {
    return raw.split(',').map((s) => ({ emailAddress: { address: s.trim() } }));
  }
}

function parseCondition(key: string, raw: string): unknown {
  if (key === 'hasAttachments' || key === 'isAutomaticForward') {
    return raw.toLowerCase() === 'true';
  }
  // Addresses fields expect JSON array; plain strings are split by comma
  if (key === 'fromAddresses' || key === 'sentToAddresses') {
    return parseEmailAddresses(raw);
  }
  // Contains fields expect string arrays
  if (key === 'bodyContains' || key === 'subjectContains' || key === 'senderContains' || key === 'recipientContains') {
    return [raw];
  }
  return raw;
}

function parseAction(key: string, raw: any): MessageRuleAction[keyof MessageRuleAction] {
  if (key === 'delete' || key === 'permanentDelete' || key === 'markAsRead' || key === 'stopProcessingRules') {
    if (typeof raw === 'boolean') return raw;
    return String(raw).toLowerCase() === 'true';
  }
  if (key === 'forwardToRecipients' || key === 'forwardAsAttachmentToRecipients') {
    return parseEmailAddresses(String(raw));
  }
  if (key === 'assignCategories') {
    return String(raw)
      .split(',')
      .map((s: string) => s.trim());
  }
  return raw;
}

function conditionsFromOpts(options: Record<string, unknown>): MessageRuleCondition | undefined {
  const conditions: MessageRuleCondition = {};
  let used = false;

  const entries: [keyof MessageRuleCondition, unknown][] = [
    ['bodyContains', options.bodyContains],
    ['subjectContains', options.subjectContains],
    ['senderContains', options.senderContains],
    ['recipientContains', options.recipientContains],
    ['fromAddresses', options.fromAddresses],
    ['sentToAddresses', options.sentToAddresses],
    ['hasAttachments', options.hasAttachments],
    ['importance', options.importance],
    ['isAutomaticForward', options.isAutomaticForward]
  ];

  for (const [key, val] of entries) {
    if (val !== undefined) {
      (conditions as Record<string, unknown>)[key] = parseCondition(key as string, String(val));
      used = true;
    }
  }
  return used ? conditions : undefined;
}

function actionsFromOpts(options: Record<string, unknown>): MessageRuleAction {
  const actions: MessageRuleAction = {};

  if (options.delete === true || options.delete === 'true') actions.delete = true;
  if (options.permanentDelete === true || options.permanentDelete === 'true') actions.permanentDelete = true;
  if (options.markAsRead === true || options.markAsRead === 'true') actions.markAsRead = true;
  if (options.stopProcessingRules === true || options.stopProcessingRules === 'true')
    actions.stopProcessingRules = true;

  if (options.moveToFolder !== undefined) actions.moveToFolder = String(options.moveToFolder);
  if (options.copyToFolder !== undefined) actions.copyToFolder = String(options.copyToFolder);
  if (options.markImportance !== undefined) actions.markImportance = String(options.markImportance) as any;

  if (options.forwardTo !== undefined)
    actions.forwardToRecipients = parseAction('forwardToRecipients', options.forwardTo as string) as any;
  if (options.forwardAsAttachmentTo !== undefined)
    actions.forwardAsAttachmentToRecipients = parseAction(
      'forwardAsAttachmentToRecipients',
      options.forwardAsAttachmentTo as string
    ) as any;
  if (options.assignCategories !== undefined)
    actions.assignCategories = parseAction('assignCategories', options.assignCategories as string) as any;

  return actions;
}

function printRule(rule: MessageRule, json: boolean): void {
  if (json) {
    console.log(JSON.stringify(rule, null, 2));
    return;
  }
  console.log(`\nRule: ${rule.displayName}`);
  console.log(`  ID:        ${rule.id}`);
  console.log(`  Priority:  ${rule.priority}`);
  console.log(`  Enabled:   ${rule.isEnabled}`);

  if (rule.conditions) {
    const c = rule.conditions;
    const parts: string[] = [];
    if (c.bodyContains?.length) parts.push(`body contains: ${c.bodyContains.join(', ')}`);
    if (c.subjectContains?.length) parts.push(`subject contains: ${c.subjectContains.join(', ')}`);
    if (c.senderContains?.length) parts.push(`sender contains: ${c.senderContains.join(', ')}`);
    if (c.recipientContains?.length) parts.push(`recipient contains: ${c.recipientContains.join(', ')}`);
    if (c.fromAddresses?.length) parts.push(`from: ${c.fromAddresses.map((a) => a.emailAddress.address).join(', ')}`);
    if (c.sentToAddresses?.length)
      parts.push(`sent to: ${c.sentToAddresses.map((a) => a.emailAddress.address).join(', ')}`);
    if (c.hasAttachments !== undefined) parts.push(`has attachments: ${c.hasAttachments}`);
    if (c.importance) parts.push(`importance: ${c.importance}`);
    if (c.isAutomaticForward !== undefined) parts.push(`auto forward: ${c.isAutomaticForward}`);
    if (parts.length) console.log(`  Conditions: ${parts.join(' | ')}`);
  }

  if (rule.actions) {
    const a = rule.actions;
    const parts: string[] = [];
    if (a.moveToFolder) parts.push(`move to: ${a.moveToFolder}`);
    if (a.copyToFolder) parts.push(`copy to: ${a.copyToFolder}`);
    if (a.delete) parts.push('delete');
    if (a.permanentDelete) parts.push('permanent delete');
    if (a.markAsRead) parts.push('mark as read');
    if (a.markImportance) parts.push(`mark importance: ${a.markImportance}`);
    if (a.forwardToRecipients?.length)
      parts.push(`forward to: ${a.forwardToRecipients.map((r) => r.emailAddress.address).join(', ')}`);
    if (a.forwardAsAttachmentToRecipients?.length) {
      parts.push(
        `forward as attachment to: ${a.forwardAsAttachmentToRecipients.map((r) => r.emailAddress.address).join(', ')}`
      );
    }
    if (a.assignCategories?.length) parts.push(`categories: ${a.assignCategories.join(', ')}`);
    if (a.stopProcessingRules) parts.push('stop processing rules');
    if (parts.length) console.log(`  Actions:    ${parts.join(' | ')}`);
  }

  if (rule.exceptionConditions) {
    console.log(`  Exceptions: (see --json for details)`);
  }
  console.log('');
}

// ---------------------------------------------------------------------------
// List subcommand
// ---------------------------------------------------------------------------
const listCmd = new Command('list')
  .description('List all inbox message rules')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .addHelpText(
    'after',
    `
Conditions you can filter by:
  --bodyContains <text>         Match message body contains text
  --subjectContains <text>      Match subject contains text
  --senderContains <text>       Match sender display name/address contains text
  --recipientContains <text>    Match any recipient contains text
  --fromAddresses <emails>      Match from addresses (comma-separated or JSON)
  --sentToAddresses <emails>    Match sent to addresses (comma-separated or JSON)
  --hasAttachments <true|false> Match attachment presence
  --importance <Low|Normal|High> Match importance
  --isAutomaticForward <true|false> Match auto-forward messages
`
  )
  .option('--bodyContains <text>', 'Match body contains text')
  .option('--subjectContains <text>', 'Match subject contains text')
  .option('--senderContains <text>', 'Match sender contains text')
  .option('--recipientContains <text>', 'Match recipient contains text')
  .option('--fromAddresses <emails>', 'Match from addresses (comma-separated or JSON)')
  .option('--sentToAddresses <emails>', 'Match sent to addresses (comma-separated or JSON)')
  .option('--hasAttachments <value>', 'Match has attachments (true|false)')
  .option('--importance <level>', 'Match importance (Low|Normal|High)')
  .option('--isAutomaticForward <value>', 'Match is auto-forward (true|false)')
  .action(async (opts) => {
    const auth = await resolveGraphAuth({ token: opts.token });
    if (!auth.success) {
      console.error(`Error: ${auth.error}`);
      process.exit(1);
    }

    const result = await listMessageRules(auth.token!);
    if (!result.ok) {
      console.error(`Error: ${result.error?.message}`);
      process.exit(1);
    }

    let rules = result.data || [];

    // Apply client-side filtering based on condition options
    const filterConditions = conditionsFromOpts(opts as Record<string, unknown>);
    if (filterConditions) {
      rules = rules.filter((rule) => {
        if (!rule.conditions) return false;
        const c = rule.conditions;

        // Check each filter condition
        if (
          filterConditions.bodyContains &&
          (!c.bodyContains || !filterConditions.bodyContains.some((term) => c.bodyContains?.includes(term)))
        ) {
          return false;
        }
        if (
          filterConditions.subjectContains &&
          (!c.subjectContains || !filterConditions.subjectContains.some((term) => c.subjectContains?.includes(term)))
        ) {
          return false;
        }
        if (
          filterConditions.senderContains &&
          (!c.senderContains || !filterConditions.senderContains.some((term) => c.senderContains?.includes(term)))
        ) {
          return false;
        }
        if (
          filterConditions.recipientContains &&
          (!c.recipientContains ||
            !filterConditions.recipientContains.some((term) => c.recipientContains?.includes(term)))
        ) {
          return false;
        }
        if (
          filterConditions.fromAddresses &&
          (!c.fromAddresses ||
            !filterConditions.fromAddresses.some((addr) =>
              c.fromAddresses?.some((ruleAddr) => ruleAddr.emailAddress.address === addr.emailAddress.address)
            ))
        ) {
          return false;
        }
        if (
          filterConditions.sentToAddresses &&
          (!c.sentToAddresses ||
            !filterConditions.sentToAddresses.some((addr) =>
              c.sentToAddresses?.some((ruleAddr) => ruleAddr.emailAddress.address === addr.emailAddress.address)
            ))
        ) {
          return false;
        }
        if (filterConditions.hasAttachments !== undefined && c.hasAttachments !== filterConditions.hasAttachments) {
          return false;
        }
        if (filterConditions.importance && c.importance !== filterConditions.importance) {
          return false;
        }
        if (
          filterConditions.isAutomaticForward !== undefined &&
          c.isAutomaticForward !== filterConditions.isAutomaticForward
        ) {
          return false;
        }

        return true;
      });
    }

    if (rules.length === 0) {
      console.log(opts.json ? '[]' : 'No inbox rules found.');
      return;
    }

    if (opts.json) {
      console.log(JSON.stringify(rules, null, 2));
      return;
    }

    for (const rule of rules) {
      printRule(rule, false);
    }
  });

// ---------------------------------------------------------------------------
// Get subcommand
// ---------------------------------------------------------------------------
const getCmd = new Command('get')
  .description('Get a single inbox rule by ID')
  .argument('<ruleId>', 'The rule ID')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .action(async (ruleId, opts) => {
    const auth = await resolveGraphAuth({ token: opts.token });
    if (!auth.success) {
      console.error(`Error: ${auth.error}`);
      process.exit(1);
    }

    const result = await getMessageRule(auth.token!, ruleId);
    if (!result.ok) {
      console.error(`Error: ${result.error?.message}`);
      process.exit(1);
    }

    printRule(result.data!, !!opts.json);
  });

// ---------------------------------------------------------------------------
// Create subcommand
// ---------------------------------------------------------------------------
const createCmd = new Command('create')
  .description('Create a new inbox message rule')
  .requiredOption('--name <name>', 'Rule display name')
  .option('--priority <number>', 'Rule priority (lower = runs first)', parseInt)
  .option('--disable', 'Create rule in disabled state')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  // Conditions
  .option('--bodyContains <text>', 'Condition: body contains text')
  .option('--subjectContains <text>', 'Condition: subject contains text')
  .option('--senderContains <text>', 'Condition: sender contains text')
  .option('--recipientContains <text>', 'Condition: recipient contains text')
  .option('--fromAddresses <emails>', 'Condition: from addresses (comma-separated or JSON)')
  .option('--sentToAddresses <emails>', 'Condition: sent to addresses (comma-separated or JSON)')
  .option('--hasAttachments <value>', 'Condition: has attachments (true|false)')
  .option('--importance <level>', 'Condition: importance (Low|Normal|High)')
  .option('--isAutomaticForward <value>', 'Condition: is auto-forward (true|false)')
  // Actions
  .option('--moveToFolder <folder>', 'Action: move to folder')
  .option('--copyToFolder <folder>', 'Action: copy to folder')
  .option('--delete', 'Action: soft-delete')
  .option('--permanentDelete', 'Action: permanent delete')
  .option('--markAsRead', 'Action: mark as read')
  .option('--markImportance <level>', 'Action: mark importance (Low|Normal|High)')
  .option('--forwardTo <emails>', 'Action: forward to recipients (comma-separated or JSON)')
  .option('--forwardAsAttachmentTo <emails>', 'Action: forward as attachment to recipients (comma-separated or JSON)')
  .option('--assignCategories <cats>', 'Action: assign categories (comma-separated)')
  .option('--stopProcessingRules', 'Action: stop processing more rules')
  .action(async (opts) => {
    const auth = await resolveGraphAuth({ token: opts.token });
    if (!auth.success) {
      console.error(`Error: ${auth.error}`);
      process.exit(1);
    }

    const conditions = conditionsFromOpts(opts as Record<string, unknown>);
    const actions = actionsFromOpts(opts as Record<string, unknown>);

    if (!conditions) {
      console.error('Error: at least one condition is required.');
      process.exit(1);
    }
    if (Object.keys(actions).length === 0) {
      console.error('Error: at least one action is required.');
      process.exit(1);
    }

    const payload: CreateMessageRulePayload = {
      displayName: opts.name,
      isEnabled: !opts.disable,
      priority: opts.priority,
      conditions,
      actions
    };

    console.log('Creating rule…');
    if (!opts.json) {
      console.log(`  Name:      ${payload.displayName}`);
      console.log(`  Priority:  ${payload.priority ?? 'default'}`);
      console.log(`  Enabled:   ${payload.isEnabled}`);
    }

    const result = await createMessageRule(auth.token!, payload);
    if (!result.ok) {
      console.error(`Error: ${result.error?.message}`);
      process.exit(1);
    }

    console.log(
      opts.json
        ? JSON.stringify(result.data, null, 2)
        : `\u2713 Rule created: ${result.data!.displayName} (${result.data!.id})`
    );
  });

// ---------------------------------------------------------------------------
// Update subcommand
// ---------------------------------------------------------------------------
const updateCmd = new Command('update')
  .description('Update an existing inbox message rule')
  .requiredOption('--id <ruleId>', 'The rule ID to update')
  .option('--name <name>', 'New rule display name')
  .option('--priority <number>', 'New rule priority', parseInt)
  .option('--enable', 'Enable the rule')
  .option('--disable', 'Disable the rule')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  // Conditions (replace all)
  .option('--bodyContains <text>', 'Condition: body contains text')
  .option('--subjectContains <text>', 'Condition: subject contains text')
  .option('--senderContains <text>', 'Condition: sender contains text')
  .option('--recipientContains <text>', 'Condition: recipient contains text')
  .option('--fromAddresses <emails>', 'Condition: from addresses (comma-separated or JSON)')
  .option('--sentToAddresses <emails>', 'Condition: sent to addresses (comma-separated or JSON)')
  .option('--hasAttachments <value>', 'Condition: has attachments (true|false)')
  .option('--importance <level>', 'Condition: importance (Low|Normal|High)')
  .option('--isAutomaticForward <value>', 'Condition: is auto-forward (true|false)')
  // Actions (replace all)
  .option('--moveToFolder <folder>', 'Action: move to folder')
  .option('--copyToFolder <folder>', 'Action: copy to folder')
  .option('--delete', 'Action: soft-delete')
  .option('--permanentDelete', 'Action: permanent delete')
  .option('--markAsRead', 'Action: mark as read')
  .option('--markImportance <level>', 'Action: mark importance (Low|Normal|High)')
  .option('--forwardTo <emails>', 'Action: forward to recipients (comma-separated or JSON)')
  .option('--forwardAsAttachmentTo <emails>', 'Action: forward as attachment to recipients (comma-separated or JSON)')
  .option('--assignCategories <cats>', 'Action: assign categories (comma-separated)')
  .option('--stopProcessingRules', 'Action: stop processing more rules')
  .action(async (opts) => {
    const auth = await resolveGraphAuth({ token: opts.token });
    if (!auth.success) {
      console.error(`Error: ${auth.error}`);
      process.exit(1);
    }

    if (opts.enable && opts.disable) {
      console.error('Error: --enable and --disable cannot be used together.');
      process.exit(1);
    }

    const conditions = conditionsFromOpts(opts as Record<string, unknown>);
    const actions = actionsFromOpts(opts as Record<string, unknown>);

    const payload: UpdateMessageRulePayload = {};
    if (opts.name) payload.displayName = opts.name;
    if (opts.priority !== undefined) payload.priority = opts.priority;
    if (opts.enable) payload.isEnabled = true;
    if (opts.disable) payload.isEnabled = false;
    if (conditions) payload.conditions = conditions;
    if (Object.keys(actions).length > 0) payload.actions = actions;

    console.log(`Updating rule ${opts.id}…`);
    const result = await updateMessageRule(auth.token!, opts.id, payload);
    if (!result.ok) {
      console.error(`Error: ${result.error?.message}`);
      process.exit(1);
    }

    console.log(opts.json ? JSON.stringify(result.data, null, 2) : `\u2713 Rule updated: ${result.data!.displayName}`);
  });

// ---------------------------------------------------------------------------
// Delete subcommand
// ---------------------------------------------------------------------------
const deleteCmd = new Command('delete')
  .description('Delete an inbox message rule')
  .argument('<ruleId>', 'The rule ID to delete')
  .option('--token <token>', 'Use a specific token')
  .option('--json', 'Output as JSON')
  .action(async (ruleId, opts) => {
    const auth = await resolveGraphAuth({ token: opts.token });
    if (!auth.success) {
      console.error(`Error: ${auth.error}`);
      process.exit(1);
    }

    const result = await deleteMessageRule(auth.token!, ruleId);
    if (!result.ok) {
      console.error(`Error: ${result.error?.message}`);
      process.exit(1);
    }

    console.log(opts.json ? JSON.stringify({ deleted: ruleId }) : `\u2713 Rule deleted: ${ruleId}`);
  });

// ---------------------------------------------------------------------------
// Main command
// ---------------------------------------------------------------------------
export const rulesCommand = new Command('rules')
  .description('Manage server-side inbox message rules (Graph messageRules API)')
  .addCommand(listCmd)
  .addCommand(getCmd)
  .addCommand(createCmd)
  .addCommand(updateCmd)
  .addCommand(deleteCmd);
