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

/** Tenant id embedded in an access token (`tid` claim). See {@link tenantIdFromJwtPayload}. */
export function getJwtPayloadTenantId(token: string): string | undefined {
  return tenantIdFromJwtPayload(decodeJwtPayload(token));
}

const TENANT_GUID_RE = /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[1-5][0-9a-fA-F]{3}-[89abAB][0-9a-fA-F]{3}-[0-9a-fA-F]{12}$/;

/**
 * True when `tenant` is a concrete tenant GUID (as opposed to the `common`/`organizations`/`consumers`
 * placeholders or a domain name, none of which ever appear as a token's `tid` claim). Used to gate
 * cached-token tenant-mismatch checks: only enforce `tid` equality when the operator pinned a specific
 * tenant, since placeholder/domain tenants resolve to a `tid` that legitimately varies per user.
 */
export function isPinnedTenantGuid(tenant: string): boolean {
  return TENANT_GUID_RE.test(tenant);
}

export interface DecodedJwtPayload {
  upn?: string;
  preferred_username?: string;
  email?: string;
  tid?: string;
  scp?: string;
  roles?: string[];
  [key: string]: unknown;
}

/**
 * Decode a JWT's payload segment once. Callers that need more than one claim (e.g. `diagnoseAuth`
 * deriving UPN + tenant id + capabilities from the same token) should decode once via this and
 * pass the result to {@link upnFromJwtPayload} / {@link tenantIdFromJwtPayload} instead of calling
 * the token-accepting per-claim getters below multiple times on the same token.
 */
export function decodeJwtPayload(token: string): DecodedJwtPayload | undefined {
  try {
    const parts = token.split('.');
    if (parts.length !== 3) return undefined;
    const payload = JSON.parse(Buffer.from(parts[1], 'base64url').toString('utf8')) as unknown;
    if (!payload || typeof payload !== 'object' || Array.isArray(payload)) return undefined;
    return payload as DecodedJwtPayload;
  } catch {
    return undefined;
  }
}

/**
 * Best-effort signed-in identity (UPN/email) from an already-decoded token payload.
 * Prefers `upn`, then `preferred_username`, then `email` — the same precedence `whoami`/
 * `verify-token` already use ad hoc when reading `payload.upn || payload.email`. Centralized here
 * so identity-guardrail checks (`--require-identity`, `--as-delegate-of`, `auth repair`, `readiness`)
 * compare against one consistent claim order. Strips embedded CR/LF (not just leading/trailing
 * whitespace) before returning — a claim value is attacker-influenceable IdP output, and downstream
 * callers echo it into terminal/log lines and persist it into profiles.json, so an embedded newline
 * could otherwise forge extra output lines.
 */
export function upnFromJwtPayload(payload: DecodedJwtPayload | undefined): string | undefined {
  const candidate = payload?.upn || payload?.preferred_username || payload?.email;
  if (typeof candidate !== 'string') return undefined;
  const cleaned = candidate.replace(/[\r\n]/g, '').trim();
  return cleaned ? cleaned : undefined;
}

/** Tenant id from an already-decoded token payload (`tid` claim). */
export function tenantIdFromJwtPayload(payload: DecodedJwtPayload | undefined): string | undefined {
  return typeof payload?.tid === 'string' && payload.tid.length > 0 ? payload.tid : undefined;
}

/** Best-effort signed-in identity (UPN/email) embedded in an access token. See {@link upnFromJwtPayload}. */
export function getJwtPayloadUpn(token: string): string | undefined {
  return upnFromJwtPayload(decodeJwtPayload(token));
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
