import { GRAPH_CRITICAL_DELEGATED_SCOPES, GRAPH_REFRESH_SCOPE_CANDIDATES } from './graph-oauth-scopes.js';
import {
  getJwtExpiration,
  getJwtPayloadAppId,
  getJwtPayloadScopeSet,
  getMicrosoftTenantPathSegment,
  isValidJwtStructure
} from './jwt-utils.js';
import {
  getUnifiedRefreshTokenFromEnv,
  loadM365TokenCache,
  type M365TokenCacheV1,
  saveM365TokenCache
} from './m365-token-cache.js';

export interface GraphAuthResult {
  success: boolean;
  token?: string;
  error?: string;
}

async function refreshGraphAccessToken(
  clientId: string,
  refreshToken: string,
  tenant: string
): Promise<{ accessToken: string; refreshToken: string; expiresAt: number }> {
  let lastError = '';

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
  }

  throw new Error(`Graph token refresh failed: ${lastError}`);
}

export async function resolveGraphAuth(options?: { token?: string; identity?: string }): Promise<GraphAuthResult> {
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

    const cached = await loadM365TokenCache(identity);
    if (cached?.graph && cached.graph.expiresAt > Date.now() + 60_000) {
      if (isValidJwtStructure(cached.graph.accessToken)) {
        const tokenAppId = getJwtPayloadAppId(cached.graph.accessToken);
        const expectId = clientId.trim();
        const mismatch = tokenAppId && expectId && tokenAppId.toLowerCase() !== expectId.toLowerCase();
        if (mismatch) {
          console.warn(
            `[graph-auth] Ignoring cached Graph access token: token app id (${tokenAppId}) does not match EWS_CLIENT_ID (${expectId}). Refreshing.`
          );
        } else {
          const scopeSet = getJwtPayloadScopeSet(cached.graph.accessToken);
          const missingCritical = GRAPH_CRITICAL_DELEGATED_SCOPES.filter((s) => !scopeSet.has(s));
          const narrowOk = cached.graphNarrowScopeAccepted === true;
          if (missingCritical.length > 0 && !narrowOk) {
            console.warn(
              `[graph-auth] Ignoring cached Graph access token: missing delegated scopes ${missingCritical.join(', ')} (narrow token — e.g. stale cache from another host). Refreshing.`
            );
          } else {
            return { success: true, token: cached.graph.accessToken };
          }
        }
      } else {
        console.warn(
          '[graph-auth] Cached Graph token has an invalid JWT structure — falling back to token refresh. ' +
            'You may need to re-authenticate if this persists.'
        );
      }
    }

    const refreshTokens = [...new Set([cached?.refreshToken, envRefreshToken].filter((t): t is string => !!t))];

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
        return { success: true, token: result.accessToken };
      } catch (err) {
        const msg = err instanceof Error ? err.message : String(err);
        const isLast = i === refreshTokens.length - 1;
        console.warn(
          `[graph-auth] Token refresh attempt failed: ${msg}${isLast ? '.' : ' — trying next token candidate.'}`
        );
      }
    }

    return {
      success: false,
      error: 'Graph token refresh failed. Update M365_REFRESH_TOKEN (or GRAPH_REFRESH_TOKEN) in .env or run `login`.'
    };
  } catch (err) {
    return {
      success: false,
      error: err instanceof Error ? err.message : 'Graph authentication failed'
    };
  }
}
