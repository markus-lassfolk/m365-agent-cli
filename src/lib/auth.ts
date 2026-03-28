import { homedir } from 'node:os';
import { join } from 'node:path';
import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { getJwtExpiration, getMicrosoftTenantPathSegment, isValidJwtStructure } from './jwt-utils.js';

export interface AuthResult {
  success: boolean;
  token?: string;
  error?: string;
}

interface CachedToken {
  accessToken: string;
  refreshToken: string;
  expiresAt: number;
}

// Security model: cache file stores bearer/refresh tokens and must be owner-only.
// Directory is created as 0700 and file writes enforce 0600 to satisfy least-privilege.
// The cache path is anchored to a fixed, local per-user directory under homedir();
// network values (token contents) are written only as file data, never used to select
// an arbitrary write location.
const TOKEN_CACHE_FILE_TEMPLATE = join(homedir(), '.config', 'clippy', 'token-cache-${identity}.json');

async function loadCachedToken(identity: string): Promise<CachedToken | null> {
  try {
    const TOKEN_CACHE_FILE = TOKEN_CACHE_FILE_TEMPLATE.replace('${identity}', identity);
    const data = await readFile(TOKEN_CACHE_FILE, 'utf-8');
    return JSON.parse(data) as CachedToken;
  } catch {
    return null;
  }
}

async function saveCachedToken(identity: string, token: CachedToken): Promise<void> {
  try {
    const dir = join(homedir(), '.config', 'clippy');
    await mkdir(dir, { recursive: true, mode: 0o700 });
    const TOKEN_CACHE_FILE = TOKEN_CACHE_FILE_TEMPLATE.replace('${identity}', identity);
    await writeFile(TOKEN_CACHE_FILE, JSON.stringify(token, null, 2), {
      encoding: 'utf-8',
      mode: 0o600
    });
  } catch {
    // Ignore write errors
  }
}

async function refreshAccessToken(clientId: string, refreshToken: string, tenant: string): Promise<CachedToken> {
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

      // Refuse to cache tokens that are not well-formed JWTs
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
    const envRefreshToken = process.env.EWS_REFRESH_TOKEN;

    if (!clientId || !envRefreshToken) {
      return {
        success: false,
        error: 'Missing EWS_CLIENT_ID or EWS_REFRESH_TOKEN in environment. Check your .env file.'
      };
    }

    const identity = options?.identity || 'default';

    const tenant = getMicrosoftTenantPathSegment();

    // Check cached token
    const cached = await loadCachedToken(identity);
    if (cached && cached.expiresAt > Date.now() + 60_000) {
      // Guard against corrupted cache: validate JWT structure before returning
      if (!isValidJwtStructure(cached.accessToken)) {
        // Treat a malformed cached token as if there were no cache
      } else {
        return { success: true, token: cached.accessToken };
      }
    }

    // Refresh - try cached refresh token first (may have been rotated), then .env
    const refreshTokens = [...new Set([cached?.refreshToken, envRefreshToken].filter((t): t is string => !!t))];

    for (const rt of refreshTokens) {
      try {
        const result = await refreshAccessToken(clientId, rt, tenant);
        await saveCachedToken(identity, result);
        return { success: true, token: result.accessToken };
      } catch {
        // Try next
      }
    }

    return {
      success: false,
      error: 'Token refresh failed. You may need to update EWS_REFRESH_TOKEN in .env.'
    };
  } catch (err) {
    return {
      success: false,
      error: err instanceof Error ? err.message : 'Authentication failed'
    };
  }
}
