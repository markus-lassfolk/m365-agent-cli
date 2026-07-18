import { afterEach, beforeEach, describe, expect, test } from 'bun:test';
import { Buffer } from 'node:buffer';
import { mkdir, mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { checkMailboxAccess, classifyAuthFailure, diagnoseAuth } from './auth-diagnostics.js';
import { GRAPH_CRITICAL_DELEGATED_SCOPES } from './graph-oauth-scopes.js';

const CLIENT_ID = '5f2abcea-d6ea-4460-b468-3d80d7a900eb';

// Real-world AADSTS50173 payload shape (redacted timestamps/user), as returned by the v2.0 token
// endpoint on a revoked refresh grant (e.g. password reset). Per issue #243 acceptance criteria,
// this must classify as tenant-side grant invalidation, not generic cache corruption.
const AADSTS50173_FIXTURE =
  "invalid_grant: AADSTS50173: The provided grant has expired due to it being revoked, a fresh auth token is needed. The user might have changed or reset their password. The grant was issued on '2026-01-01T00:00:00.0000000Z' and the TokensValidFrom date (before which tokens are not valid) for this user is '2026-02-01T00:00:00.0000000Z'.";

describe('classifyAuthFailure', () => {
  test('AADSTS50173 classifies as refresh_grant_revoked, not cache corruption', () => {
    const result = classifyAuthFailure(AADSTS50173_FIXTURE);
    expect(result.failureClass).toBe('refresh_grant_revoked');
    expect(result.evidence).toContain('AADSTS50173');
    expect(result.evidence).toContain('tokens_valid_from_after_grant');
    expect(result.recommendedAction).toBe('interactive_login');
  });

  test('interaction_required / AADSTS500133 classifies as interaction_required', () => {
    const result = classifyAuthFailure(
      'interaction_required: AADSTS500133: Assertion is not within its valid time range.'
    );
    expect(result.failureClass).toBe('interaction_required');
    expect(result.evidence).toContain('AADSTS500133');
  });

  test('AADSTS70008 / AADSTS700082 classify as refresh_grant_expired', () => {
    expect(classifyAuthFailure('AADSTS70008: expired token').failureClass).toBe('refresh_grant_expired');
    expect(classifyAuthFailure('AADSTS700082: expired due to inactivity').failureClass).toBe('refresh_grant_expired');
  });

  test('AADSTS65001 classifies as consent_required', () => {
    expect(classifyAuthFailure('AADSTS65001: consent required').failureClass).toBe('consent_required');
  });

  test('AADSTS53003 classifies as conditional_access_blocked', () => {
    const result = classifyAuthFailure('AADSTS53003: blocked by Conditional Access policy');
    expect(result.failureClass).toBe('conditional_access_blocked');
    expect(result.recommendedAction).toBe('contact_admin_or_interactive_login');
  });

  test('AADSTS50076/50079 classify as mfa_required and recommend browser login', () => {
    expect(classifyAuthFailure('AADSTS50076: MFA required').failureClass).toBe('mfa_required');
    expect(classifyAuthFailure('AADSTS50076: MFA required').recommendedAction).toBe('interactive_login_browser');
  });

  test(
    'MFA is classified as mfa_required (not the generic interaction_required) even when the ' +
      'response also carries the generic "interaction_required" error code — a real conditional-MFA ' +
      'token-endpoint response commonly has both signals in the same text',
    () => {
      const result = classifyAuthFailure(
        'interaction_required: AADSTS50076: Due to a configuration change made by your administrator, ' +
          'or because you moved to a new location, you must use multi-factor authentication.'
      );
      expect(result.failureClass).toBe('mfa_required');
      expect(result.recommendedAction).toBe('interactive_login_browser');
    }
  );

  test('empty/undefined error text classifies as unknown_error', () => {
    expect(classifyAuthFailure(undefined).failureClass).toBe('unknown_error');
    expect(classifyAuthFailure('').failureClass).toBe('unknown_error');
  });

  test('never echoes raw secret material — only takes already-sanitized strings', () => {
    // classifyAuthFailure has no access to tokens at all; it only inspects the string it is given.
    const result = classifyAuthFailure(AADSTS50173_FIXTURE);
    expect(JSON.stringify(result)).not.toMatch(/ey[A-Za-z0-9_-]{10,}/); // no JWT-looking substrings
  });
});

describe('diagnoseAuth', () => {
  let testHome: string;
  let originalEnv: NodeJS.ProcessEnv;

  beforeEach(async () => {
    testHome = await mkdtemp(join(tmpdir(), 'm365-auth-diagnostics-'));
    originalEnv = { ...process.env };
    process.env.M365_AGENT_CLI_CONFIG_DIR = join(testHome, '.config', 'm365-agent-cli');
    process.env.EWS_TENANT_ID = 'common';
    process.env.NODE_ENV = 'test';
  });

  afterEach(async () => {
    for (const key of Object.keys(process.env)) {
      if (!(key in originalEnv)) delete process.env[key];
    }
    for (const [key, value] of Object.entries(originalEnv)) {
      if (value === undefined) delete process.env[key];
      else process.env[key] = value;
    }
    await rm(testHome, { recursive: true, force: true }).catch(() => {});
  });

  test('missing_credentials when EWS_CLIENT_ID / refresh token are both absent', async () => {
    delete process.env.EWS_CLIENT_ID;
    delete process.env.M365_REFRESH_TOKEN;
    delete process.env.GRAPH_REFRESH_TOKEN;
    delete process.env.EWS_REFRESH_TOKEN;

    const diag = await diagnoseAuth({ identity: 'default' });
    expect(diag.status).toBe('repair_required');
    expect(diag.failureClass).toBe('missing_credentials');
    expect(diag.safeCommand).toBe('m365-agent-cli login');
    expect(diag.secretsPrinted).toBe(false);
  });

  test('healthy when a valid, unexpired cached Graph token exists', async () => {
    process.env.EWS_CLIENT_ID = CLIENT_ID;
    process.env.M365_REFRESH_TOKEN = 'fake-refresh-token';
    const dir = process.env.M365_AGENT_CLI_CONFIG_DIR as string;
    await mkdir(dir, { recursive: true });
    const h = Buffer.from(JSON.stringify({ alg: 'none', typ: 'JWT' })).toString('base64url');
    const p = Buffer.from(
      JSON.stringify({
        exp: 2_000_000_000,
        appid: CLIENT_ID,
        upn: 'doris@lassfolk.net',
        tid: '11111111-2222-4333-8444-555555555555',
        scp: GRAPH_CRITICAL_DELEGATED_SCOPES.join(' ')
      })
    ).toString('base64url');
    await writeFile(
      join(dir, 'token-cache-default.json'),
      JSON.stringify({
        version: 1,
        refreshToken: 'fake-refresh-token',
        graph: { accessToken: `${h}.${p}.x`, expiresAt: Date.now() + 3_600_000 }
      }),
      'utf8'
    );

    const diag = await diagnoseAuth({ identity: 'default' });
    expect(diag.status).toBe('healthy');
    expect(diag.signedInAs).toBe('doris@lassfolk.net');
    expect(diag.tenantId).toBe('11111111-2222-4333-8444-555555555555');
    expect(diag.cacheHealth).toBe('healthy');
    expect(diag.capabilities.length).toBeGreaterThan(0);
    expect(diag.authBackend).toBe('graph');
  });

  test('healthy via EWS fallback maps to a non-empty synthetic capability set and authBackend "ews"', async () => {
    // EWS.AccessAsUser.All is an all-or-nothing grant that isn't representable in a Graph token's
    // scp/roles claims — a healthy EWS-only diagnosis used to leave `capabilities: []`, making
    // `readiness --require mail.read` falsely report the capability missing for an identity that
    // can, in fact, read mail via EWS (found in QA review).
    process.env.EWS_CLIENT_ID = CLIENT_ID;
    process.env.M365_REFRESH_TOKEN = 'fake-refresh-token';
    // No `graph` cache entry and every fetch fails — resolveGraphAuth can't succeed, so
    // diagnoseAuth falls through to the EWS branch, which hits the seeded `ews` cache entry below
    // without ever needing a real network call.
    global.fetch = (async () => new Response('', { status: 500 })) as unknown as typeof fetch;

    const dir = process.env.M365_AGENT_CLI_CONFIG_DIR as string;
    await mkdir(dir, { recursive: true });
    const h = Buffer.from(JSON.stringify({ alg: 'none', typ: 'JWT' })).toString('base64url');
    const p = Buffer.from(JSON.stringify({ exp: 2_000_000_000, appid: CLIENT_ID, upn: 'doris@lassfolk.net' })).toString(
      'base64url'
    );
    await writeFile(
      join(dir, 'token-cache-default.json'),
      JSON.stringify({
        version: 1,
        refreshToken: 'fake-refresh-token',
        ews: { accessToken: `${h}.${p}.x`, expiresAt: Date.now() + 3_600_000 }
      }),
      'utf8'
    );

    const diag = await diagnoseAuth({ identity: 'default' });
    expect(diag.status).toBe('healthy');
    expect(diag.authBackend).toBe('ews');
    expect(diag.signedInAs).toBe('doris@lassfolk.net');
    expect(diag.capabilities.length).toBeGreaterThan(0);
    expect(diag.capabilities).toContain('Mail.ReadWrite');
  });

  test('classifies a revoked refresh grant (AADSTS50173) from a failed refresh attempt', async () => {
    process.env.EWS_CLIENT_ID = CLIENT_ID;
    process.env.M365_REFRESH_TOKEN = 'fake-refresh-token';
    global.fetch = (async () =>
      new Response(JSON.stringify({ error: 'invalid_grant', error_description: AADSTS50173_FIXTURE }), {
        status: 400
      })) as unknown as typeof fetch;

    const diag = await diagnoseAuth({ identity: 'default' });
    expect(diag.status).toBe('repair_required');
    expect(diag.failureClass).toBe('refresh_grant_revoked');
    expect(diag.evidence).toContain('AADSTS50173');
    expect(diag.safeCommand).toBe('m365-agent-cli login');
    expect(diag.secretsPrinted).toBe(false);
  });

  test('missing_cache when credentials exist but there is no cache and refresh fails generically', async () => {
    process.env.EWS_CLIENT_ID = CLIENT_ID;
    process.env.M365_REFRESH_TOKEN = 'fake-refresh-token';
    global.fetch = (async () => new Response('', { status: 500 })) as unknown as typeof fetch;

    const diag = await diagnoseAuth({ identity: 'default' });
    expect(diag.status).toBe('repair_required');
    expect(diag.failureClass).toBe('missing_cache');
    expect(diag.cacheHealth).toBe('missing');
  });
});

describe('checkMailboxAccess', () => {
  test('returns checked:false when no mailbox is given', async () => {
    expect(await checkMailboxAccess('tok', undefined)).toEqual({ checked: false });
  });

  test('reports ok:true on a successful mailbox read', async () => {
    const originalFetch = global.fetch;
    global.fetch = (async () =>
      new Response(JSON.stringify({ id: 'inbox-id' }), {
        status: 200,
        headers: { 'content-type': 'application/json' }
      })) as unknown as typeof fetch;
    try {
      const result = await checkMailboxAccess('tok', 'lotta@lassfolk.net');
      expect(result).toEqual({ checked: true, mailbox: 'lotta@lassfolk.net', ok: true, error: undefined });
    } finally {
      global.fetch = originalFetch;
    }
  });

  test('reports ok:false with an error message on access denial', async () => {
    const originalFetch = global.fetch;
    global.fetch = (async () =>
      new Response(JSON.stringify({ error: { code: 'ErrorAccessDenied', message: 'Access is denied.' } }), {
        status: 403,
        headers: { 'content-type': 'application/json' }
      })) as unknown as typeof fetch;
    try {
      const result = await checkMailboxAccess('tok', 'lotta@lassfolk.net');
      expect(result.checked).toBe(true);
      expect(result.ok).toBe(false);
      expect(result.error).toContain('Access is denied');
    } finally {
      global.fetch = originalFetch;
    }
  });
});
