import { Command } from 'commander';
import { checkMailboxAccess, diagnoseAuth, type MailboxAccessResult } from '../lib/auth-diagnostics.js';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import { evaluateGraphCapabilities, GRAPH_CAPABILITY_MATRIX } from '../lib/graph-capability-matrix.js';
import { upnsMatch } from '../lib/identity-guard.js';
import { getProfileByIdentity, resolveIdentitySlug } from '../lib/identity-profiles.js';
import { loadM365TokenCache } from '../lib/m365-token-cache.js';
import { deepRedact, IDENTITY_LABEL_SAFE_KEYS } from '../lib/redact.js';
import { applyEnvFileOverrides, resolveEnvFilePathArgument } from '../lib/utils.js';

/** Bump when the JSON shape changes in a way automation should branch on. */
export const READINESS_SCHEMA_VERSION = 1;

/** Friendly capability tokens (issue #247's examples) mapped onto `graph-capability-matrix.ts` rows. */
const CAPABILITY_ALIASES: Record<string, { rowId: string; dimension: 'read' | 'write' }> = {
  'mail.read': { rowId: 'mail.own', dimension: 'read' },
  'mail.write': { rowId: 'mail.own', dimension: 'write' },
  'mail.send': { rowId: 'mail.send', dimension: 'write' },
  'calendar.read': { rowId: 'calendar.own', dimension: 'read' },
  'calendar.write': { rowId: 'calendar.own', dimension: 'write' },
  'mail.shared.read': { rowId: 'mail.shared', dimension: 'read' },
  'mail.shared.write': { rowId: 'mail.shared', dimension: 'write' },
  'calendar.shared.read': { rowId: 'calendar.shared', dimension: 'read' },
  'calendar.shared.write': { rowId: 'calendar.shared', dimension: 'write' }
};

/** `--require <token>`: a known alias, or `<matrixRowId>.read`/`.write` for any other row. */
function resolveCapabilityRequirement(token: string): { rowId: string; dimension: 'read' | 'write' } | undefined {
  const alias = CAPABILITY_ALIASES[token];
  if (alias) return alias;
  const m = /^(.+)\.(read|write)$/.exec(token);
  if (!m) return undefined;
  const [, rowId, dim] = m;
  if (!GRAPH_CAPABILITY_MATRIX.some((r) => r.id === rowId)) return undefined;
  return { rowId, dimension: dim as 'read' | 'write' };
}

export interface ReadinessResult {
  schemaVersion: number;
  ready: boolean;
  signedInAs: string | null;
  expectedIdentity: string | null;
  identityMismatch: boolean;
  mailbox: string | null;
  tenantId: string | null;
  identity: string;
  profile: string | null;
  authHealth: string;
  cacheHealth: string;
  tokenExpiresAt: string | null;
  refreshTokenPresent: boolean;
  capabilities: string[];
  missingCapabilities: string[];
  mailboxAccess: MailboxAccessResult | null;
  recommendedAction: string | null;
  safeCommand: string | null;
  secretsPrinted: false;
}

export interface ComputeReadinessOptions {
  identity?: string;
  mailbox?: string;
  requireTokens?: string[];
  expectIdentity?: string;
  envPath?: string;
}

async function computeMailboxAccess(
  diag: Awaited<ReturnType<typeof diagnoseAuth>>,
  options: ComputeReadinessOptions
): Promise<MailboxAccessResult | null> {
  if (diag.status !== 'healthy' || !options.mailbox) return null;
  if (diag.authBackend !== 'graph') {
    // Mailbox delegation is only checkable via Graph today (`checkMailboxAccess` calls the Graph
    // API) — an EWS-only-healthy identity is genuinely healthy, but this specific check can't run
    // for it. Report unchecked rather than a misleading "no token" failure that would flip
    // `ready` to false for an identity that can actually operate on mail/calendar via EWS.
    return { checked: false, mailbox: options.mailbox };
  }
  const graphAuth = await resolveGraphAuth({ identity: diag.identity, envPath: options.envPath });
  if (graphAuth.success && graphAuth.token) {
    return checkMailboxAccess(graphAuth.token, options.mailbox);
  }
  return {
    checked: true,
    mailbox: options.mailbox,
    ok: false,
    error: 'Could not obtain a token to check mailbox access.'
  };
}

export async function computeReadiness(options: ComputeReadinessOptions): Promise<ReadinessResult> {
  const identity = await resolveIdentitySlug(options.identity);
  const diag = await diagnoseAuth({ identity, envPath: options.envPath });
  const mailboxAccess = await computeMailboxAccess(diag, options);

  const requireTokens = options.requireTokens ?? [];
  const permSet = new Set(diag.capabilities);
  const evaluatedRows = evaluateGraphCapabilities(permSet);
  const missingCapabilities: string[] = [];
  for (const token of requireTokens) {
    const req = resolveCapabilityRequirement(token);
    if (!req) {
      missingCapabilities.push(token);
      continue;
    }
    const row = evaluatedRows.find((r) => r.id === req.rowId);
    const ok = req.dimension === 'read' ? row?.readOk : row?.writeOk;
    if (!ok) missingCapabilities.push(token);
  }

  const identityMismatch = Boolean(options.expectIdentity && !upnsMatch(diag.signedInAs, options.expectIdentity));

  const ready =
    diag.status === 'healthy' && missingCapabilities.length === 0 && mailboxAccess?.ok !== false && !identityMismatch;

  let recommendedAction: string | null = null;
  let safeCommand: string | null = null;
  if (diag.status === 'healthy') {
    if (identityMismatch) {
      recommendedAction = 'interactive_login';
      safeCommand = `m365-agent-cli login --identity ${identity}`;
    } else if (missingCapabilities.length > 0) {
      recommendedAction = 'interactive_login';
      safeCommand = 'm365-agent-cli login';
    } else if (mailboxAccess?.ok === false) {
      recommendedAction = 'check_config';
      safeCommand = `m365-agent-cli delegates list --mailbox ${options.mailbox}`;
    }
  } else {
    recommendedAction = diag.recommendedAction;
    safeCommand = diag.safeCommand ?? null;
  }

  const [cache, profile] = await Promise.all([
    loadM365TokenCache(identity).catch(() => null),
    getProfileByIdentity(identity)
  ]);
  const tokenExpiresAt = cache?.graph?.expiresAt ?? cache?.ews?.expiresAt;

  return {
    schemaVersion: READINESS_SCHEMA_VERSION,
    ready,
    signedInAs: diag.signedInAs ?? null,
    expectedIdentity: options.expectIdentity ?? null,
    identityMismatch,
    mailbox: options.mailbox ?? null,
    tenantId: diag.tenantId ?? null,
    identity,
    profile: profile?.name ?? null,
    authHealth: diag.failureClass,
    cacheHealth: diag.cacheHealth,
    tokenExpiresAt: tokenExpiresAt ? new Date(tokenExpiresAt).toISOString() : null,
    refreshTokenPresent: diag.failureClass !== 'missing_credentials',
    capabilities: diag.capabilities,
    missingCapabilities,
    mailboxAccess,
    recommendedAction,
    safeCommand,
    secretsPrinted: false
  };
}

function collectRequire(value: string, previous: string[]): string[] {
  return [...previous, value];
}

export const readinessCommand = new Command('readiness')
  .description('Machine-readable readiness contract: can mail/calendar operations run right now?')
  .option('--identity <name>', 'Identity/cache slot to check (default: the default profile, else "default")')
  .option('--mailbox <email>', 'Also verify delegated/shared access to this mailbox')
  .option(
    '--require <capability>',
    'Require a capability (repeatable), e.g. mail.read, mail.send, calendar.read, or <matrixRowId>.read/.write',
    collectRequire,
    []
  )
  .option('--expect-identity <upn>', 'Report identityMismatch/ready:false unless the signed-in UPN matches exactly')
  .option('--env-file <path>', 'Load EWS_CLIENT_ID / refresh token from this file before checking')
  .option('--json', 'Output as JSON (this is the primary, documented output — see docs/AUTHENTICATION.md)')
  .action(
    async (opts: {
      identity?: string;
      mailbox?: string;
      require: string[];
      expectIdentity?: string;
      envFile?: string;
      json?: boolean;
    }) => {
      if (opts.envFile) {
        applyEnvFileOverrides(resolveEnvFilePathArgument(opts.envFile));
      }
      const resolvedEnvPath = opts.envFile ? resolveEnvFilePathArgument(opts.envFile) : undefined;

      const result = await computeReadiness({
        identity: opts.identity,
        mailbox: opts.mailbox,
        requireTokens: opts.require,
        expectIdentity: opts.expectIdentity,
        envPath: resolvedEnvPath
      });

      if (opts.json) {
        console.log(JSON.stringify(deepRedact(result, { safeKeys: IDENTITY_LABEL_SAFE_KEYS }), null, 2));
      } else {
        console.log(`Ready: ${result.ready ? 'yes' : 'no'}`);
        console.log(`Signed in as: ${result.signedInAs ?? '(not signed in)'}`);
        if (result.mailbox) console.log(`Mailbox: ${result.mailbox}`);
        console.log(`Auth health: ${result.authHealth}`);
        if (result.missingCapabilities.length > 0) {
          console.log(`Missing capabilities: ${result.missingCapabilities.join(', ')}`);
        }
        if (!result.ready) {
          if (result.recommendedAction) console.log(`Recommended action: ${result.recommendedAction}`);
          if (result.safeCommand) console.log(`Command: ${result.safeCommand}`);
        }
        console.log('Tip: run with --json for the full machine-readable readiness contract.');
      }

      // Exit 0 whenever the CLI itself ran successfully — readiness (true/false) lives in the
      // `ready` JSON field, not the exit code (issue #247's exit semantics).
    }
  );
