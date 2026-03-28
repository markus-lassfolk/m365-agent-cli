const VALID_TENANT_ID =
  /^(?:common|organizations|consumers|[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[1-5][0-9a-fA-F]{3}-[89abAB][0-9a-fA-F]{3}-[0-9a-fA-F]{12})$/;

export function getMicrosoftTenantPathSegment(): string {
  const rawTenant = process.env.EWS_TENANT_ID?.trim() || 'common';
  if (!VALID_TENANT_ID.test(rawTenant)) {
    throw new Error('Invalid EWS_TENANT_ID. Use common/organizations/consumers or a valid tenant UUID.');
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
