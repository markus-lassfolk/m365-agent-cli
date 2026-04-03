const VALID_TENANT_ID =
  /^(?:common|organizations|consumers|[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[1-5][0-9a-fA-F]{3}-[89abAB][0-9a-fA-F]{3}-[0-9a-fA-F]{12}|[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)+)$/;

export function getMicrosoftTenantPathSegment(): string {
  const rawTenant = process.env.EWS_TENANT_ID?.trim() || 'common';
  if (!VALID_TENANT_ID.test(rawTenant)) {
    throw new Error(
      'Invalid EWS_TENANT_ID. Use common/organizations/consumers, a valid tenant UUID, or a domain name.'
    );
  }
  return rawTenant;
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

export function isValidJwtStructure(token: string): boolean {
  const parts = token.split('.');
  if (parts.length !== 3) return false;
  try {
    Buffer.from(parts[1], 'base64url').toString();
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
