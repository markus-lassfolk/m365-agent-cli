import { homedir } from 'node:os';
import { join } from 'node:path';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { getJwtExpiration, getMicrosoftTenantPathSegment, isValidJwtStructure } from './jwt-utils.js';

export interface GraphAuthResult {
  success: boolean;
  token?: string;
  error?: string;
}

interface CachedGraphToken {
  accessToken: string;
  refreshToken: string;
  expiresAt: number;
}

const GRAPH_TOKEN_CACHE_FILE = join(homedir(), '.config', 'clippy', 'graph-token-cache.json');
const GRAPH_SCOPES = [
  'https://graph.microsoft.com/Files.ReadWrite offline_access User.Read',
  'https://graph.microsoft.com/Files.ReadWrite.All offline_access User.Read',
  'https://graph.microsoft.com/Sites.ReadWrite.All offline_access User.Read',
  'https://graph.microsoft.com/.default offline_access',
  'https://graph.microsoft.com/Files.Read offline_access User.Read'
];

async function loadCachedGraphToken(): Promise<CachedGraphToken | null> {
  try {
    const data = await readFile(GRAPH_TOKEN_CACHE_FILE, 'utf-8');
    return JSON.parse(data) as CachedGraphToken;
  } catch {
    return null;
  }
}

async function saveCachedGraphToken(token: CachedGraphToken): Promise<void> {
  try {
    const dir = join(homedir(), '.config', 'clippy');
    await mkdir(dir, { recursive: true, mode: 0o700 });
    await writeFile(GRAPH_TOKEN_CACHE_FILE, JSON.stringify(token, null, 2), {
      encoding: 'utf-8',
      mode: 0o600
    });
  } catch {
    // Ignore cache write failures
  }
}

async function refreshGraphAccessToken(
  clientId: string,
  refreshToken: string,
  tenant: string
): Promise<CachedGraphToken> {
  let lastError = '';

  for (const scope of GRAPH_SCOPES) {
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

  throw new Error(`Graph token refresh failed: ${lastError}`);
}

export async function resolveGraphAuth(options?: { token?: string }): Promise<GraphAuthResult> {
  if (options?.token) {
    return { success: true, token: options.token };
  }

  try {
    const clientId = process.env.EWS_CLIENT_ID;
    const envRefreshToken = process.env.GRAPH_REFRESH_TOKEN || process.env.EWS_REFRESH_TOKEN;

    if (!clientId || !envRefreshToken) {
      return {
        success: false,
        error: 'Missing EWS_CLIENT_ID and/or GRAPH_REFRESH_TOKEN (or EWS_REFRESH_TOKEN) in environment.'
      };
    }

    const tenant = getMicrosoftTenantPathSegment();

    const cached = await loadCachedGraphToken();
    if (cached && cached.expiresAt > Date.now() + 60_000) {
      // Guard against corrupted cache: validate JWT structure before returning
      if (!isValidJwtStructure(cached.accessToken)) {
        // Treat a malformed cached token as if there were no cache
      } else {
        return { success: true, token: cached.accessToken };
      }
    }

    const refreshTokens = [...new Set([cached?.refreshToken, envRefreshToken].filter((t): t is string => !!t))];

    for (const refreshToken of refreshTokens) {
      try {
        const result = await refreshGraphAccessToken(clientId, refreshToken, tenant);
        await saveCachedGraphToken(result);
        return { success: true, token: result.accessToken };
      } catch {
        // Try next token candidate
      }
    }

    return {
      success: false,
      error: 'Graph token refresh failed. You may need to update GRAPH_REFRESH_TOKEN (or EWS_REFRESH_TOKEN) in .env.'
    };
  } catch (err) {
    return {
      success: false,
      error: err instanceof Error ? err.message : 'Graph authentication failed'
    };
  }
}
