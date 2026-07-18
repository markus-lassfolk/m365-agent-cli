/**
 * Shared auth failure classification + diagnostic orchestration used by `auth repair` (#243),
 * `readiness` (#247), and `doctor` (#246). One place to keep AADSTS-code → failure-class mapping
 * and the "what should the operator do next" recommendation consistent across all three surfaces.
 *
 * Never logs or returns raw access/refresh tokens — only sanitized error text (already stripped of
 * control characters and truncated by `auth.ts`/`graph-auth.ts`) and derived metadata.
 */
import { resolveAuth } from './auth.js';
import { getExchangeBackend } from './exchange-backend.js';
import { permissionSetFromGraphPayload } from './graph-capability-matrix.js';
import { graphUserPath } from './graph-user-path.js';
import { resolveGraphAuth } from './graph-auth.js';
import { callGraph, GraphApiError } from './graph-client.js';
import { type CacheHealth, probeCacheHealth } from './identity-profiles.js';
import { getJwtPayloadTenantId, getJwtPayloadUpn } from './jwt-utils.js';
import { getUnifiedRefreshTokenFromEnv } from './m365-token-cache.js';

export type AuthFailureClass =
  | 'healthy'
  | 'missing_credentials'
  | 'missing_cache'
  | 'malformed_cache'
  | 'refresh_grant_revoked'
  | 'refresh_grant_expired'
  | 'interaction_required'
  | 'mfa_required'
  | 'conditional_access_blocked'
  | 'consent_required'
  | 'tenant_client_mismatch'
  | 'unknown_error';

export type RecommendedAction =
  | 'none'
  | 'run_login'
  | 'interactive_login'
  | 'interactive_login_browser'
  | 'contact_admin_or_interactive_login'
  | 'check_config';

export interface ClassifiedFailure {
  failureClass: AuthFailureClass;
  evidence: string[];
  recommendedAction: RecommendedAction;
}

const AADSTS_CODE_RE = /AADSTS\d+/gi;

/**
 * Classify a sanitized OAuth/Graph error string (already stripped of secrets by
 * `auth.ts`/`graph-auth.ts`'s `sanitizeRefreshError`) into a stable failure class.
 *
 * AADSTS50173 ("The provided grant has expired due to it being revoked... TokensValidFrom")
 * is deliberately classified as tenant-side grant invalidation (`refresh_grant_revoked`), not
 * generic local cache corruption — see issue #243's acceptance criteria.
 */
export function classifyAuthFailure(errorText: string | undefined | null): ClassifiedFailure {
  const text = (errorText ?? '').toLowerCase();
  const evidence = [...new Set((errorText?.match(AADSTS_CODE_RE) ?? []).map((c) => c.toUpperCase()))];

  if (text.includes('aadsts50173')) {
    if (text.includes('tokensvalidfrom')) evidence.push('tokens_valid_from_after_grant');
    return { failureClass: 'refresh_grant_revoked', evidence, recommendedAction: 'interactive_login' };
  }
  if (text.includes('interaction_required') || text.includes('aadsts500133')) {
    evidence.push('interaction_required');
    return { failureClass: 'interaction_required', evidence, recommendedAction: 'interactive_login' };
  }
  if (text.includes('aadsts700082') || text.includes('aadsts70008')) {
    return { failureClass: 'refresh_grant_expired', evidence, recommendedAction: 'interactive_login' };
  }
  if (text.includes('aadsts65001')) {
    return { failureClass: 'consent_required', evidence, recommendedAction: 'interactive_login' };
  }
  if (text.includes('aadsts53003')) {
    return {
      failureClass: 'conditional_access_blocked',
      evidence,
      recommendedAction: 'contact_admin_or_interactive_login'
    };
  }
  if (text.includes('aadsts50076') || text.includes('aadsts50079')) {
    evidence.push('mfa_required');
    return { failureClass: 'mfa_required', evidence, recommendedAction: 'interactive_login_browser' };
  }
  if (
    text.includes('aadsts700016') ||
    text.includes('aadsts7000215') ||
    text.includes('unauthorized_client') ||
    text.includes('invalid_client')
  ) {
    return { failureClass: 'tenant_client_mismatch', evidence, recommendedAction: 'check_config' };
  }
  if (!text.trim()) {
    return { failureClass: 'unknown_error', evidence, recommendedAction: 'interactive_login' };
  }
  return { failureClass: 'unknown_error', evidence, recommendedAction: 'interactive_login' };
}

export interface AuthDiagnosis {
  status: 'healthy' | 'repair_required';
  identity: string;
  signedInAs?: string;
  tenantId?: string;
  failureClass: AuthFailureClass;
  evidence: string[];
  recommendedAction: RecommendedAction;
  safeCommand?: string;
  cacheHealth: CacheHealth;
  capabilities: string[];
  /** Always false — this module never surfaces raw token/secret material. */
  secretsPrinted: false;
}

function safeCommandFor(action: RecommendedAction, identity: string): string | undefined {
  const idFlag = identity !== 'default' ? ` --identity ${identity}` : '';
  switch (action) {
    case 'run_login':
      return `m365-agent-cli login${idFlag}`;
    case 'interactive_login':
      return `m365-agent-cli login${idFlag}`;
    case 'interactive_login_browser':
      return `m365-agent-cli login --browser${idFlag}`;
    case 'contact_admin_or_interactive_login':
      return `m365-agent-cli login${idFlag}`;
    case 'check_config':
      return 'm365-agent-cli verify-token --capabilities';
    default:
      return undefined;
  }
}

/** Offline (no network) capability list decoded from whatever Graph access token was just obtained. */
function capabilitiesFromToken(token: string | undefined): string[] {
  if (!token) return [];
  try {
    const parts = token.split('.');
    if (parts.length !== 3) return [];
    const payload = JSON.parse(Buffer.from(parts[1], 'base64url').toString('utf8'));
    return [...permissionSetFromGraphPayload(payload)].sort();
  } catch {
    return [];
  }
}

/**
 * Run a full, read-only auth diagnosis for one identity: offline cache-health probe, then a
 * (cache-hit-when-possible) auth resolution attempt, classified into a stable failure taxonomy.
 * Never mutates auth state beyond whatever refresh the resolution attempt itself performs — the
 * same refresh any other command would trigger.
 */
export async function diagnoseAuth(options: { identity: string; envPath?: string }): Promise<AuthDiagnosis> {
  const { identity, envPath } = options;
  const hasClientId = Boolean(process.env.EWS_CLIENT_ID?.trim());
  const hasRefreshTokenEnv = Boolean(getUnifiedRefreshTokenFromEnv());
  const cacheHealth = await probeCacheHealth(identity);

  if (!hasClientId || !hasRefreshTokenEnv) {
    const evidence = [
      ...(!hasClientId ? ['missing_client_id'] : []),
      ...(!hasRefreshTokenEnv ? ['missing_refresh_token'] : [])
    ];
    return {
      status: 'repair_required',
      identity,
      failureClass: 'missing_credentials',
      evidence,
      recommendedAction: 'run_login',
      safeCommand: safeCommandFor('run_login', identity),
      cacheHealth,
      capabilities: [],
      secretsPrinted: false
    };
  }

  // Prefer Graph (richer claim/scope set); fall back to EWS so `M365_EXCHANGE_BACKEND=ews` setups
  // still get a real diagnosis instead of a Graph-only false negative.
  const graph = await resolveGraphAuth({ identity, envPath });
  if (graph.success && graph.token) {
    return {
      status: 'healthy',
      identity,
      signedInAs: getJwtPayloadUpn(graph.token),
      tenantId: getJwtPayloadTenantId(graph.token),
      failureClass: 'healthy',
      evidence: [],
      recommendedAction: 'none',
      cacheHealth: await probeCacheHealth(identity),
      capabilities: capabilitiesFromToken(graph.token),
      secretsPrinted: false
    };
  }

  const ews = getExchangeBackend() === 'graph' ? null : await resolveAuth({ identity, envPath });
  if (ews?.success && ews.token) {
    return {
      status: 'healthy',
      identity,
      signedInAs: getJwtPayloadUpn(ews.token),
      tenantId: getJwtPayloadTenantId(ews.token),
      failureClass: 'healthy',
      evidence: [],
      recommendedAction: 'none',
      cacheHealth: await probeCacheHealth(identity),
      capabilities: [],
      secretsPrinted: false
    };
  }

  const errorText = graph.lastRefreshError || graph.error || ews?.lastRefreshError || ews?.error;
  let classified = classifyAuthFailure(errorText);
  if (classified.failureClass === 'unknown_error' && cacheHealth === 'missing') {
    classified = {
      failureClass: 'missing_cache',
      evidence: classified.evidence,
      recommendedAction: 'interactive_login'
    };
  }

  return {
    status: 'repair_required',
    identity,
    failureClass: classified.failureClass,
    evidence: classified.evidence,
    recommendedAction: classified.recommendedAction,
    safeCommand: safeCommandFor(classified.recommendedAction, identity),
    cacheHealth,
    capabilities: [],
    secretsPrinted: false
  };
}

export interface MailboxAccessResult {
  checked: boolean;
  mailbox?: string;
  ok?: boolean;
  error?: string;
}

/**
 * Minimal, read-only mailbox delegation probe: a single `GET .../mailFolders/inbox` against the
 * target mailbox. Succeeds only if the signed-in identity actually has delegated/shared access —
 * a cheap, safe stand-in for "can this identity operate on this mailbox" without reading any mail.
 */
export async function checkMailboxAccess(token: string, mailbox: string | undefined): Promise<MailboxAccessResult> {
  if (!mailbox) return { checked: false };
  try {
    const res = await callGraph<{ id?: string }>(token, graphUserPath(mailbox, 'mailFolders/inbox'));
    return { checked: true, mailbox, ok: res.ok, error: res.ok ? undefined : res.error?.message };
  } catch (err) {
    const message = err instanceof GraphApiError ? err.message : err instanceof Error ? err.message : String(err);
    return { checked: true, mailbox, ok: false, error: message };
  }
}
