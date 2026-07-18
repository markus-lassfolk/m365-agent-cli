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
import { resolveGraphAuth } from './graph-auth.js';
import { permissionSetFromGraphPayload } from './graph-capability-matrix.js';
import { callGraph, GraphApiError } from './graph-client.js';
import { graphUserPath } from './graph-user-path.js';
import { type CacheHealth, probeCacheHealth } from './identity-profiles.js';
import { type DecodedJwtPayload, decodeJwtPayload, tenantIdFromJwtPayload, upnFromJwtPayload } from './jwt-utils.js';
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
  // MFA (AADSTS50076/50079) is checked BEFORE the generic `interaction_required`/AADSTS500133
  // branch below: a real conditional-MFA token-endpoint response commonly has the shape
  // `{"error":"interaction_required","error_description":"AADSTS50076: ... multi-factor
  // authentication ..."}` — both signals present in the same text. Checking the generic case
  // first would always shadow the more specific (and more actionable — `interactive_login_browser`
  // vs `interactive_login`) MFA classification.
  if (text.includes('aadsts50076') || text.includes('aadsts50079')) {
    evidence.push('mfa_required');
    return { failureClass: 'mfa_required', evidence, recommendedAction: 'interactive_login_browser' };
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
  if (
    text.includes('aadsts700016') ||
    text.includes('aadsts7000215') ||
    text.includes('unauthorized_client') ||
    text.includes('invalid_client')
  ) {
    return { failureClass: 'tenant_client_mismatch', evidence, recommendedAction: 'check_config' };
  }
  return { failureClass: 'unknown_error', evidence, recommendedAction: 'interactive_login' };
}

export interface AuthDiagnosis {
  status: 'healthy' | 'repair_required';
  identity: string;
  signedInAs?: string;
  tenantId?: string;
  /** Which backend the healthy diagnosis actually resolved through. Only set when `status` is
   *  `'healthy'` — a mailbox/capability check that requires Graph should gate on this rather than
   *  assume `resolveGraphAuth` will succeed just because the identity is generally healthy. */
  authBackend?: 'graph' | 'ews';
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
    case 'interactive_login':
    case 'contact_admin_or_interactive_login':
      return `m365-agent-cli login${idFlag}`;
    case 'interactive_login_browser':
      return `m365-agent-cli login --browser${idFlag}`;
    case 'check_config':
      return 'm365-agent-cli verify-token --capabilities';
    default:
      return undefined;
  }
}

/**
 * `EWS.AccessAsUser.All` is an all-or-nothing delegated grant covering full mail + calendar
 * read/write for the signed-in mailbox — it isn't represented in a Graph token's `scp`/`roles`
 * claims at all (see `graph-capability-matrix.ts`'s `ews` row), so a healthy EWS-only diagnosis
 * maps onto the Graph-scope-shaped capability ids that `readiness --require`/`doctor` check
 * against, instead of leaving `capabilities` empty (which would make `readiness --require
 * mail.read` report `missingCapabilities` for an identity that can, in fact, read mail via EWS).
 */
const EWS_FULL_ACCESS_CAPABILITIES = ['Mail.ReadWrite', 'Mail.Send', 'Calendars.ReadWrite'];

/** Offline (no network) capability list decoded from an already-decoded token payload. */
function capabilitiesFromTokenPayload(payload: DecodedJwtPayload | undefined): string[] {
  if (!payload) return [];
  try {
    return [...permissionSetFromGraphPayload(payload)].sort();
  } catch {
    return [];
  }
}

/** Offline (no network) capability list decoded from a Graph access token. Exported so other
 *  commands (e.g. `profiles show`) that need the same decode don't each reimplement it. */
export function capabilitiesFromToken(token: string | undefined): string[] {
  if (!token) return [];
  return capabilitiesFromTokenPayload(decodeJwtPayload(token));
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
    const payload = decodeJwtPayload(graph.token);
    return {
      status: 'healthy',
      identity,
      signedInAs: upnFromJwtPayload(payload),
      tenantId: tenantIdFromJwtPayload(payload),
      authBackend: 'graph',
      failureClass: 'healthy',
      evidence: [],
      recommendedAction: 'none',
      cacheHealth: await probeCacheHealth(identity),
      capabilities: capabilitiesFromTokenPayload(payload),
      secretsPrinted: false
    };
  }

  const ews = getExchangeBackend() === 'graph' ? null : await resolveAuth({ identity, envPath });
  if (ews?.success && ews.token) {
    const payload = decodeJwtPayload(ews.token);
    return {
      status: 'healthy',
      identity,
      signedInAs: upnFromJwtPayload(payload),
      tenantId: tenantIdFromJwtPayload(payload),
      authBackend: 'ews',
      failureClass: 'healthy',
      evidence: [],
      recommendedAction: 'none',
      cacheHealth: await probeCacheHealth(identity),
      capabilities: EWS_FULL_ACCESS_CAPABILITIES,
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
