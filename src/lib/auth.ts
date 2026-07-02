import { getActiveEnvFilePath } from './active-env.js';
import { persistRefreshTokenToEnv } from './env-persist.js';
import { getJwtExpiration, getMicrosoftTenantPathSegment, isValidJwtStructure } from './jwt-utils.js';
import {
  getUnifiedRefreshTokenFromEnv,
  loadM365TokenCache,
  type M365TokenCacheV1,
  saveM365TokenCache
} from './m365-token-cache.js';

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

export interface AuthResult {
  success: boolean;
  token?: string;
  error?: string;
  /**
   * Sanitized last refresh-token-exchange error (AADSTS code + description, HTTP status, etc.).
   * Surfaces detailed OAuth failure context to callers and tests without leaking secrets.
   */
  lastRefreshError?: string;
}

async function refreshAccessToken(clientId: string, refreshToken: string, tenant: string) {
  const scopes = [
    'https://outlook.office365.com/EWS.AccessAsUser.All offline_access',
    'https://outlook.office365.com/.default offline_access'
  ];

  let lastError = '';
  let lastInteractionRequired = false;

  for (const scope of scopes) {
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

  const error = new Error(`Token refresh failed: ${lastError}`) as Error & {
    lastRefreshError?: string;
    interactionRequired?: boolean;
  };
  // Preserve the most recent (last) OAuth error so callers can surface AADSTS / interaction_required details.
  error.lastRefreshError = sanitizeRefreshError(lastError);
  error.interactionRequired = lastInteractionRequired;
  throw error;
}

export async function resolveAuth(options?: {
  token?: string;
  identity?: string;
  envPath?: string;
}): Promise<AuthResult> {
  if (options?.token) {
    return { success: true, token: options.token };
  }

  try {
    const clientId = process.env.EWS_CLIENT_ID;
    const envRefreshToken = getUnifiedRefreshTokenFromEnv();

    if (!clientId || !envRefreshToken) {
      return {
        success: false,
        error:
          'Missing EWS_CLIENT_ID or refresh token. Set M365_REFRESH_TOKEN (preferred) or GRAPH_REFRESH_TOKEN or EWS_REFRESH_TOKEN in environment or run `m365-agent-cli login`.'
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
    if (cached?.ews && cached.ews.expiresAt > Date.now() + 60_000) {
      if (isValidJwtStructure(cached.ews.accessToken)) {
        return { success: true, token: cached.ews.accessToken };
      }
    }

    const refreshTokens = [...new Set([cached?.refreshToken, envRefreshToken].filter((t): t is string => !!t))];

    // Resolve the active env file once so all refresh attempts persist to the same file the
    // CLI loaded from (M365_AGENT_ENV_FILE / --env-file / default). Without this, refreshes
    // can silently write rotated tokens to the default global .env.
    const activeEnvPath = getActiveEnvFilePath(options?.envPath);
    const lastRefreshErrors: string[] = [];
    let interactionRequired = false;

    for (const rt of refreshTokens) {
      try {
        const result = await refreshAccessToken(clientId, rt, tenant);
        const next: M365TokenCacheV1 = {
          version: 1,
          refreshToken: result.refreshToken,
          ews: { accessToken: result.accessToken, expiresAt: result.expiresAt },
          graph: cached?.graph
        };
        await saveM365TokenCache(identity, next);
        await persistRefreshTokenToEnv(result.refreshToken, {
          envPath: activeEnvPath,
          previousRefreshToken: cached?.refreshToken ?? envRefreshToken
        });
        return { success: true, token: result.accessToken };
      } catch (err) {
        // Capture the most recent OAuth error (AADSTS code, description, HTTP status) per attempt
        // so the caller can surface it on failure instead of a generic message.
        const msg = err instanceof Error ? err.message : String(err);
        const detailed =
          err && typeof err === 'object' && 'lastRefreshError' in err
            ? (err as { lastRefreshError?: string }).lastRefreshError
            : undefined;
        if (err && typeof err === 'object' && (err as { interactionRequired?: boolean }).interactionRequired) {
          interactionRequired = true;
        }
        lastRefreshErrors.push(sanitizeRefreshError(detailed || msg));
      }
    }

    // Aggregate the last error from each candidate so the caller can see AADSTS / interaction_required
    // details instead of a generic "Token refresh failed" message. The env-persist call above
    // re-resolves via getActiveEnvFilePath(options?.envPath) when envPath is omitted, so the
    // default path remains correct.
    const lastErrorDetail = lastRefreshErrors[lastRefreshErrors.length - 1] ?? '';
    const interactiveHint = interactionRequired
      ? ' Azure requires re-authentication (interaction_required / AADSTS500133). Run `m365-agent-cli login` again.'
      : '';
    return {
      success: false,
      error:
        'Token refresh failed. Update M365_REFRESH_TOKEN (or GRAPH_REFRESH_TOKEN / EWS_REFRESH_TOKEN) in .env or run `login`.' +
        (lastErrorDetail ? ` Last error: ${lastErrorDetail}` : '') +
        interactiveHint,
      lastRefreshError: lastErrorDetail || undefined
    };
  } catch (err) {
    return {
      success: false,
      error: err instanceof Error ? err.message : 'Authentication failed'
    };
  }
}
