import { getJwtExpiration, getMicrosoftTenantPathSegment, isValidJwtStructure } from './jwt-utils.js';
import {
  getUnifiedRefreshTokenFromEnv,
  loadM365TokenCache,
  type M365TokenCacheV1,
  saveM365TokenCache
} from './m365-token-cache.js';

export interface AuthResult {
  success: boolean;
  token?: string;
  error?: string;
}

async function refreshAccessToken(clientId: string, refreshToken: string, tenant: string) {
  const scopes = [
    'https://outlook.office365.com/EWS.AccessAsUser.All offline_access',
    'https://outlook.office365.com/.default offline_access'
  ];

  let lastError = '';

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

  throw new Error(`Token refresh failed: ${lastError}`);
}

export async function resolveAuth(options?: { token?: string; identity?: string }): Promise<AuthResult> {
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
        return { success: true, token: result.accessToken };
      } catch {
        // Try next
      }
    }

    return {
      success: false,
      error:
        'Token refresh failed. Update M365_REFRESH_TOKEN (or GRAPH_REFRESH_TOKEN / EWS_REFRESH_TOKEN) in .env or run `login`.'
    };
  } catch (err) {
    return {
      success: false,
      error: err instanceof Error ? err.message : 'Authentication failed'
    };
  }
}
