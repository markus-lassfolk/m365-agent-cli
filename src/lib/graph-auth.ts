import { getActiveEnvFilePath } from './active-env.js';
import { persistRefreshTokenToEnv } from './env-persist.js';
import { GRAPH_CRITICAL_DELEGATED_SCOPES, GRAPH_REFRESH_SCOPE_CANDIDATES } from './graph-oauth-scopes.js';
import {
  getJwtExpiration,
  getJwtPayloadAppId,
  getJwtPayloadScopeSet,
  getJwtPayloadTenantId,
  getMicrosoftTenantPathSegment,
  isPinnedTenantGuid,
  isValidJwtStructure
} from './jwt-utils.js';
import {
  getUnifiedRefreshTokenFromEnv,
  loadM365TokenCache,
  type M365TokenCacheV1,
  saveM365TokenCache
} from './m365-token-cache.js';
import { withRefreshTokenLock } from './refresh-token-lock.js';

export interface GraphAuthResult {
  success: boolean;
  token?: string;
  error?: string;
  /**
   * Sanitized last refresh-token-exchange error (AADSTS code + description, HTTP status, etc.).
   * Surfaces detailed OAuth failure context to callers and tests without leaking secrets.
   */
  lastRefreshError?: string;
}

/**
 * Sanitize an OAuth token-endpoint error for surfacing in CLI output / tests.
 * Strips control characters and truncates long error_descriptions to keep logs bounded.
 * Does NOT touch refresh tokens or access tokens — callers must not pass those here.
 */
function sanitizeRefreshError(raw: string | undefined | null): string {
  if (!raw) return '';
  return raw
    .replace(/\p{Cc}/gu, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 500);
}

async function refreshGraphAccessToken(
  clientId: string,
  refreshToken: string,
  tenant: string
): Promise<{ accessToken: string; refreshToken: string; expiresAt: number }> {
  let lastError = '';
  let lastInteractionRequired = false;

  for (const scope of GRAPH_REFRESH_SCOPE_CANDIDATES) {
    const response = await fetch(`https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id: clientId,
        grant_type: 'refresh_token',
        refresh_token: refreshToken,
        scope
      }).toString()
    });

    const json = (await response.json()) as {
      access_token?: string;
      refresh_token?: string;
      expires_in?: number;
      error?: string;
      error_description?: string;
      error_codes?: number[];
      suberror?: string;
    };

    if (response.ok && json.access_token) {
      const accessToken = json.access_token;

      if (!isValidJwtStructure(accessToken)) {
        throw new Error('OAuth server returned an invalid token structure — refusing to cache');
      }

      const expiresAt = getJwtExpiration(accessToken) ?? Date.now() + (json.expires_in || 3600) * 1000;
      if (expiresAt <= Date.now()) {
        throw new Error('OAuth server returned an already-expired token — refusing to cache');
      }

      return {
        accessToken,
        refreshToken: json.refresh_token || refreshToken,
        expiresAt
      };
    }

    lastError = [json.error, json.error_description].filter(Boolean).join(': ') || `HTTP ${response.status}`;
    // AADSTS error code 500133 / sub error "interaction_required" means the user must re-authenticate
    // (e.g. MFA, conditional access, or revoked grant). Track this so the caller can prompt re-login.
    if (json.error === 'interaction_required' || (json.error_codes ?? []).includes(500133)) {
      lastInteractionRequired = true;
    }
  }

  const err = new Error(`Graph token refresh failed: ${lastError}`) as Error & {
    lastRefreshError?: string;
    interactionRequired?: boolean;
  };
  err.lastRefreshError = sanitizeRefreshError(lastError);
  err.interactionRequired = lastInteractionRequired;
  throw err;
}

export async function resolveGraphAuth(options?: {
  token?: string;
  identity?: string;
  /** When true, skip the "cached access token still valid" shortcut and refresh from the refresh token. */
  forceRefresh?: boolean;
  /** Resolved env file path used to persist rotated refresh tokens (e.g. from `--env-file`). */
  envPath?: string;
}): Promise<GraphAuthResult> {
  if (options?.token) {
    return { success: true, token: options.token };
  }

  try {
    const clientId = process.env.EWS_CLIENT_ID;
    const envRefreshToken = getUnifiedRefreshTokenFromEnv();

    if (!clientId) {
      return {
        success: false,
        error: 'Missing EWS_CLIENT_ID in environment. Check your .env file or Azure app registration.'
      };
    }

    if (!envRefreshToken) {
      return {
        success: false,
        error:
          'Missing refresh token for Microsoft Graph. Set M365_REFRESH_TOKEN (preferred) or GRAPH_REFRESH_TOKEN or EWS_REFRESH_TOKEN, or run `m365-agent-cli login`.'
      };
    }

    const identity = options?.identity || 'default';
    if (!/^[a-zA-Z0-9_-]+$/.test(identity)) {
      return {
        success: false,
        error: 'Invalid identity name. Only alphanumeric characters, hyphens, and underscores are allowed.'
      };
    }

    const tenant = getMicrosoftTenantPathSegment();

    const tryCachedGraph = (cached: M365TokenCacheV1 | null | undefined): GraphAuthResult | null => {
      if (options?.forceRefresh) return null;
      if (!(cached?.graph && cached.graph.expiresAt > Date.now() + 60_000)) return null;
      if (!isValidJwtStructure(cached.graph.accessToken)) {
        console.warn(
          '[graph-auth] Cached Graph token has an invalid JWT structure — falling back to token refresh. ' +
            'You may need to re-authenticate if this persists.'
        );
        return null;
      }
      const tokenAppId = getJwtPayloadAppId(cached.graph.accessToken);
      const expectId = clientId.trim();
      const appIdMismatch = tokenAppId && expectId && tokenAppId.toLowerCase() !== expectId.toLowerCase();
      const tokenTenantId = getJwtPayloadTenantId(cached.graph.accessToken);
      // Only enforce tenant equality when the operator pinned a concrete tenant GUID — `common` /
      // `organizations` / `consumers` / domain-name tenants resolve to a `tid` that legitimately
      // varies per user, so comparing those would produce false-positive refreshes.
      const tenantMismatch =
        isPinnedTenantGuid(tenant) && tokenTenantId && tokenTenantId.toLowerCase() !== tenant.toLowerCase();
      if (appIdMismatch) {
        console.warn(
          `[graph-auth] Ignoring cached Graph access token: token app id (${tokenAppId}) does not match EWS_CLIENT_ID (${expectId}). Refreshing.`
        );
        return null;
      }
      if (tenantMismatch) {
        console.warn(
          `[graph-auth] Ignoring cached Graph access token: token tenant (${tokenTenantId}) does not match configured tenant (${tenant}). Refreshing.`
        );
        return null;
      }
      const scopeSet = getJwtPayloadScopeSet(cached.graph.accessToken);
      const missingCritical = GRAPH_CRITICAL_DELEGATED_SCOPES.filter((s) => !scopeSet.has(s));
      const narrowOk = cached.graphNarrowScopeAccepted === true;
      if (missingCritical.length > 0 && !narrowOk) {
        console.warn(
          `[graph-auth] Ignoring cached Graph access token: missing delegated scopes ${missingCritical.join(', ')} (narrow token — e.g. stale cache from another host). Refreshing.`
        );
        return null;
      }
      return { success: true, token: cached.graph.accessToken };
    };

    const cachedBeforeLock = await loadM365TokenCache(identity);
    const hit = tryCachedGraph(cachedBeforeLock);
    if (hit) return hit;

    // Serialize refresh against EWS auth for the same identity (shared refresh token).
    return await withRefreshTokenLock(identity, async () => {
      const cached = await loadM365TokenCache(identity);
      const afterWait = tryCachedGraph(cached);
      if (afterWait) return afterWait;

      const refreshTokens = [
        ...new Set([cached?.refreshToken, getUnifiedRefreshTokenFromEnv() ?? envRefreshToken].filter((t): t is string => !!t))
      ];

      // Resolve the active env file once so all refresh attempts persist to the same file the
      // CLI loaded from (M365_AGENT_ENV_FILE / --env-file / default). Without this, refreshes
      // can silently write rotated tokens to the default global .env.
      const activeEnvPath = getActiveEnvFilePath(options?.envPath);
      const lastRefreshErrors: string[] = [];
      let interactionRequired = false;

      for (let i = 0; i < refreshTokens.length; i++) {
        try {
          const result = await refreshGraphAccessToken(clientId, refreshTokens[i], tenant);
          const newScopeSet = getJwtPayloadScopeSet(result.accessToken);
          const missingAfterRefresh = GRAPH_CRITICAL_DELEGATED_SCOPES.filter((s) => !newScopeSet.has(s));
          const next: M365TokenCacheV1 = {
            version: 1,
            refreshToken: result.refreshToken,
            ews: cached?.ews,
            graph: { accessToken: result.accessToken, expiresAt: result.expiresAt },
            graphNarrowScopeAccepted: missingAfterRefresh.length > 0
          };
          await saveM365TokenCache(identity, next);
          await persistRefreshTokenToEnv(result.refreshToken, {
            envPath: activeEnvPath,
            previousRefreshToken: cached?.refreshToken ?? envRefreshToken
          });
          return { success: true, token: result.accessToken };
        } catch (err) {
          const msg = err instanceof Error ? err.message : String(err);
          const isLast = i === refreshTokens.length - 1;
          // Capture detailed OAuth error context (AADSTS codes, sub-error, descriptions) for
          // surfacing in the returned `error` / `lastRefreshError` fields. Falls back to the
          // formatted Error.message if the refresh helper did not attach a structured detail.
          const detailed =
            err && typeof err === 'object' && 'lastRefreshError' in err
              ? (err as { lastRefreshError?: string }).lastRefreshError
              : undefined;
          if (err && typeof err === 'object' && (err as { interactionRequired?: boolean }).interactionRequired) {
            interactionRequired = true;
          }
          lastRefreshErrors.push(sanitizeRefreshError(detailed || msg));
          console.warn(
            `[graph-auth] Token refresh attempt failed: ${msg}${isLast ? '.' : ' — trying next token candidate.'}`
          );
        }
      }

      const lastErrorDetail = lastRefreshErrors[lastRefreshErrors.length - 1] ?? '';
      const interactiveHint = interactionRequired
        ? ' Azure requires re-authentication (interaction_required / AADSTS500133). Run `m365-agent-cli login` again.'
        : '';
      return {
        success: false,
        error:
          'Graph token refresh failed. Update M365_REFRESH_TOKEN (or GRAPH_REFRESH_TOKEN) in .env or run `login`.' +
          (lastErrorDetail ? ` Last error: ${lastErrorDetail}` : '') +
          interactiveHint,
        lastRefreshError: lastErrorDetail || undefined
      };
    });
  } catch (err) {
    return {
      success: false,
      error: err instanceof Error ? err.message : 'Graph authentication failed'
    };
  }
}

export async function requireGraphAuth(opts: { token?: string; identity?: string }): Promise<string> {
  const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
  if (!auth.success || !auth.token) {
    console.error(`Auth error: ${auth.error}`);
    process.exit(1);
  }
  return auth.token;
}
