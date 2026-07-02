const VALID_TENANT_ID =
  /^(?:common|organizations|consumers|[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[1-5][0-9a-fA-F]{3}-[89abAB][0-9a-fA-F]{3}-[0-9a-fA-F]{12}|[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)+)$/;

/**
 * Pure precedence resolver. Exported separately so tests can call it with a stub env
 * (independent of any `mock.module('./jwt-utils.js', ...)` that parallel suites may
 * install on the wrapper).
 */
export function resolveTenantPathSegment(envSource: NodeJS.ProcessEnv): string {
  const rawTenant =
    envSource.M365_TENANT_ID?.trim() ||
    envSource.MICROSOFT_TENANT_ID?.trim() ||
    envSource.EWS_TENANT_ID?.trim() ||
    'common';
  if (!VALID_TENANT_ID.test(rawTenant)) {
    throw new Error(
      'Invalid tenant id (M365_TENANT_ID / MICROSOFT_TENANT_ID / EWS_TENANT_ID). Use common/organizations/consumers, a valid tenant UUID, or a domain name.'
    );
  }
  return rawTenant;
}

/**
 * Resolve the Microsoft OAuth tenant path segment.
 *
 * Precedence: `M365_TENANT_ID` > `MICROSOFT_TENANT_ID` > `EWS_TENANT_ID` (legacy) > `'common'`.
 * The legacy `EWS_TENANT_ID` name predates the CLI's Graph scope and remains supported.
 */
export function getMicrosoftTenantPathSegment(): string {
  return resolveTenantPathSegment(process.env);
}

export function getJwtExpiration(token: string): number | null {
  try {
    const parts = token.split('.');
    if (parts.length !== 3) return null;
    const payload = JSON.parse(Buffer.from(parts[1], 'base64url').toString());
    return payload.exp ? payload.exp * 1000 : null;
  } catch {
    return null;
  }
}

/**
 * Strict structural JWT validation.
 *
 * A valid bearer JWT has exactly three non-empty base64url segments (header.payload.signature)
 * and a JSON-decodable payload that parses as an object. The previous implementation only called
 * `Buffer.from(parts[1], 'base64url').toString()` (which never throws and yields an empty string
 * for malformed input), so tokens like `"a.."` or `"...."` were silently accepted.
 *
 * The header and signature are not cryptographically verified here — that requires the JWK set.
 * This check is intentionally fast and offline-safe so cached / refreshed access tokens can be
 * rejected before they hit the Graph SDK or get written to disk.
 */
export function isValidJwtStructure(token: string): boolean {
  if (typeof token !== 'string' || token.length === 0) return false;
  const parts = token.split('.');
  if (parts.length !== 3) return false;
  // Reject empty segments — RFC 7519 §7.2 forbids them.
  if (!parts[0] || !parts[1] || !parts[2]) return false;
  try {
    const headerRaw = Buffer.from(parts[0], 'base64url').toString('utf8');
    const payloadRaw = Buffer.from(parts[1], 'base64url').toString('utf8');
    if (!headerRaw || !payloadRaw) return false;
    const payload = JSON.parse(payloadRaw) as unknown;
    if (!payload || typeof payload !== 'object' || Array.isArray(payload)) return false;
    // Header is required to be a JSON object too (some issuers omit `typ` but always emit `alg`).
    const header = JSON.parse(headerRaw) as unknown;
    if (!header || typeof header !== 'object' || Array.isArray(header)) return false;
    return true;
  } catch {
    return false;
  }
}

/**
 * Application (client) id embedded in an access token (`appid`, or `azp` when present).
 * Used to avoid reusing a cached Graph access token issued for a different Entra app than `EWS_CLIENT_ID`.
 */
export function getJwtPayloadAppId(token: string): string | undefined {
  try {
    const parts = token.split('.');
    if (parts.length !== 3) return undefined;
    const payload = JSON.parse(Buffer.from(parts[1], 'base64url').toString('utf8')) as {
      appid?: string;
      azp?: string;
    };
    if (typeof payload.appid === 'string' && payload.appid.length > 0) return payload.appid;
    if (typeof payload.azp === 'string' && payload.azp.length > 0) return payload.azp;
    return undefined;
  } catch {
    return undefined;
  }
}

/** Space-separated delegated scopes on a Graph access token (`scp` claim). */
export function getJwtPayloadScopeSet(token: string): Set<string> {
  try {
    const parts = token.split('.');
    if (parts.length !== 3) return new Set();
    const payload = JSON.parse(Buffer.from(parts[1], 'base64url').toString('utf8')) as { scp?: string };
    if (typeof payload.scp !== 'string' || !payload.scp.trim()) return new Set();
    return new Set(
      payload.scp
        .split(/\s+/)
        .map((s) => s.trim())
        .filter(Boolean)
    );
  } catch {
    return new Set();
  }
}
