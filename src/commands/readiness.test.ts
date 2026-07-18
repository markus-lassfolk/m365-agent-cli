import { afterEach, beforeEach, describe, expect, test } from 'bun:test';
import { Buffer } from 'node:buffer';
import { mkdir, mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { GRAPH_CRITICAL_DELEGATED_SCOPES } from '../lib/graph-oauth-scopes.js';
import { upsertProfile } from '../lib/identity-profiles.js';
import { computeReadiness, READINESS_SCHEMA_VERSION } from './readiness.js';

const CLIENT_ID = '5f2abcea-d6ea-4460-b468-3d80d7a900eb';

function fixtureAccessToken(opts: { upn?: string; scp?: string[] }): string {
  const h = Buffer.from(JSON.stringify({ alg: 'none', typ: 'JWT' })).toString('base64url');
  const p = Buffer.from(
    JSON.stringify({
      exp: 2_000_000_000,
      appid: CLIENT_ID,
      upn: opts.upn,
      scp: (opts.scp ?? GRAPH_CRITICAL_DELEGATED_SCOPES).join(' ')
    })
  ).toString('base64url');
  return `${h}.${p}.x`;
}

describe('computeReadiness', () => {
  let testHome: string;
  let originalEnv: NodeJS.ProcessEnv;
  const originalFetch = global.fetch;

  beforeEach(async () => {
    testHome = await mkdtemp(join(tmpdir(), 'm365-readiness-'));
    originalEnv = { ...process.env };
    process.env.M365_AGENT_CLI_CONFIG_DIR = join(testHome, '.config', 'm365-agent-cli');
    process.env.EWS_TENANT_ID = 'common';
    process.env.NODE_ENV = 'test';
  });

  afterEach(async () => {
    global.fetch = originalFetch;
    for (const key of Object.keys(process.env)) {
      if (!(key in originalEnv)) delete process.env[key];
    }
    for (const [key, value] of Object.entries(originalEnv)) {
      if (value === undefined) delete process.env[key];
      else process.env[key] = value;
    }
    await rm(testHome, { recursive: true, force: true }).catch(() => {});
  });

  async function seedHealthyCache(identity: string, opts: { upn: string; scp?: string[] }): Promise<void> {
    const dir = process.env.M365_AGENT_CLI_CONFIG_DIR as string;
    await mkdir(dir, { recursive: true });
    await writeFile(
      join(dir, `token-cache-${identity}.json`),
      JSON.stringify({
        version: 1,
        refreshToken: 'fake-refresh-token',
        graph: { accessToken: fixtureAccessToken(opts), expiresAt: Date.now() + 3_600_000 },
        // Avoid the "narrow token" refresh path when a test deliberately omits a critical scope.
        graphNarrowScopeAccepted: true
      }),
      'utf8'
    );
  }

  test('scenario: healthy auth — ready with no missing capabilities', async () => {
    process.env.EWS_CLIENT_ID = CLIENT_ID;
    process.env.M365_REFRESH_TOKEN = 'fake-refresh-token';
    await seedHealthyCache('default', {
      upn: 'doris@lassfolk.net',
      scp: [...GRAPH_CRITICAL_DELEGATED_SCOPES, 'Mail.ReadWrite', 'Calendars.ReadWrite']
    });

    const result = await computeReadiness({ requireTokens: ['mail.read', 'mail.send', 'calendar.read'] });
    expect(result.ready).toBe(true);
    expect(result.authHealth).toBe('healthy');
    expect(result.signedInAs).toBe('doris@lassfolk.net');
    expect(result.missingCapabilities).toEqual([]);
    expect(result.recommendedAction).toBeNull();
    expect(result.secretsPrinted).toBe(false);
    expect(result.schemaVersion).toBe(READINESS_SCHEMA_VERSION);
  });

  test('scenario: missing auth — not ready, recommends login', async () => {
    delete process.env.EWS_CLIENT_ID;
    delete process.env.M365_REFRESH_TOKEN;
    delete process.env.GRAPH_REFRESH_TOKEN;
    delete process.env.EWS_REFRESH_TOKEN;

    const result = await computeReadiness({});
    expect(result.ready).toBe(false);
    expect(result.authHealth).toBe('missing_credentials');
    expect(result.recommendedAction).toBe('run_login');
    expect(result.safeCommand).toBe('m365-agent-cli login');
    expect(result.signedInAs).toBeNull();
  });

  test('scenario: revoked refresh grant — not ready, classified correctly with a safe command', async () => {
    process.env.EWS_CLIENT_ID = CLIENT_ID;
    process.env.M365_REFRESH_TOKEN = 'fake-refresh-token';
    global.fetch = (async () =>
      new Response(
        JSON.stringify({
          error: 'invalid_grant',
          error_description: 'AADSTS50173: revoked grant, TokensValidFrom is later.'
        }),
        { status: 400 }
      )) as unknown as typeof fetch;

    const result = await computeReadiness({});
    expect(result.ready).toBe(false);
    expect(result.authHealth).toBe('refresh_grant_revoked');
    expect(result.recommendedAction).toBe('interactive_login');
    expect(result.safeCommand).toContain('m365-agent-cli login');
  });

  test('scenario: wrong identity — not ready via --expect-identity even though auth itself is healthy', async () => {
    process.env.EWS_CLIENT_ID = CLIENT_ID;
    process.env.M365_REFRESH_TOKEN = 'fake-refresh-token';
    await seedHealthyCache('default', { upn: 'doris@lassfolk.net' });

    const result = await computeReadiness({ expectIdentity: 'lotta@lassfolk.net' });
    expect(result.ready).toBe(false);
    expect(result.authHealth).toBe('healthy');
    expect(result.identityMismatch).toBe(true);
    expect(result.signedInAs).toBe('doris@lassfolk.net');
    expect(result.expectedIdentity).toBe('lotta@lassfolk.net');
    expect(result.recommendedAction).toBe('interactive_login');
  });

  test('scenario: wrong identity — ready when --expect-identity matches (case-insensitive)', async () => {
    process.env.EWS_CLIENT_ID = CLIENT_ID;
    process.env.M365_REFRESH_TOKEN = 'fake-refresh-token';
    await seedHealthyCache('default', { upn: 'doris@lassfolk.net' });

    const result = await computeReadiness({ expectIdentity: 'DORIS@lassfolk.net' });
    expect(result.identityMismatch).toBe(false);
    expect(result.ready).toBe(true);
  });

  test('scenario: missing mailbox delegation — not ready, mailboxAccess reflects the denial', async () => {
    process.env.EWS_CLIENT_ID = CLIENT_ID;
    process.env.M365_REFRESH_TOKEN = 'fake-refresh-token';
    await seedHealthyCache('default', { upn: 'doris@lassfolk.net' });
    global.fetch = (async () =>
      new Response(JSON.stringify({ error: { code: 'ErrorAccessDenied', message: 'Access is denied.' } }), {
        status: 403,
        headers: { 'content-type': 'application/json' }
      })) as unknown as typeof fetch;

    const result = await computeReadiness({ mailbox: 'lotta@lassfolk.net' });
    expect(result.ready).toBe(false);
    expect(result.mailboxAccess).toEqual({
      checked: true,
      mailbox: 'lotta@lassfolk.net',
      ok: false,
      error: expect.stringContaining('Access is denied')
    });
    expect(result.recommendedAction).toBe('check_config');
    expect(result.safeCommand).toContain('delegates list --mailbox lotta@lassfolk.net');
  });

  test('scenario: EWS-only healthy identity + --mailbox — reports unchecked, not a false failure', async () => {
    // Mailbox delegation is only checkable via Graph today; an EWS-only-healthy identity used to
    // get a misleading "could not obtain a token" mailboxAccess failure (and ready:false) purely
    // because it isn't Graph-authenticated, even though it's genuinely healthy (found in QA review).
    process.env.EWS_CLIENT_ID = CLIENT_ID;
    process.env.M365_REFRESH_TOKEN = 'fake-refresh-token';
    // No graph cache entry and every fetch fails — resolveGraphAuth can't succeed, so diagnoseAuth
    // (and the mailbox check) fall through to / gate on the EWS branch below.
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

    const result = await computeReadiness({ mailbox: 'lotta@lassfolk.net' });
    expect(result.authHealth).toBe('healthy');
    expect(result.mailboxAccess).toEqual({ checked: false, mailbox: 'lotta@lassfolk.net' });
    expect(result.ready).toBe(true);
  });

  test('scenario: profile field resolves by bound identity, not just by matching name', async () => {
    // A profile can be registered under a name that differs from the cache-slot identity it's
    // bound to (`upsertProfile('work', {identity: 'acct2'})`); `getProfile(identity)` used to look
    // up by that mismatched key and always return undefined for such profiles (found in QA review).
    process.env.EWS_CLIENT_ID = CLIENT_ID;
    process.env.M365_REFRESH_TOKEN = 'fake-refresh-token';
    await seedHealthyCache('acct2', { upn: 'doris@lassfolk.net' });
    await upsertProfile('work', { identity: 'acct2' });

    const result = await computeReadiness({ identity: 'acct2' });
    expect(result.identity).toBe('acct2');
    expect(result.profile).toBe('work');
  });

  test('missing capabilities are reported and block readiness even when auth is healthy', async () => {
    process.env.EWS_CLIENT_ID = CLIENT_ID;
    process.env.M365_REFRESH_TOKEN = 'fake-refresh-token';
    // Cache has scopes that satisfy the "narrow token" check but not Mail.Send specifically.
    await seedHealthyCache('default', {
      upn: 'doris@lassfolk.net',
      scp: GRAPH_CRITICAL_DELEGATED_SCOPES.filter((s) => s !== 'Mail.Send')
    });

    const result = await computeReadiness({ requireTokens: ['mail.send'] });
    expect(result.ready).toBe(false);
    expect(result.missingCapabilities).toEqual(['mail.send']);
  });

  test('unknown --require token is treated as unmet (fail closed)', async () => {
    process.env.EWS_CLIENT_ID = CLIENT_ID;
    process.env.M365_REFRESH_TOKEN = 'fake-refresh-token';
    await seedHealthyCache('default', { upn: 'doris@lassfolk.net' });

    const result = await computeReadiness({ requireTokens: ['not.a.real.capability'] });
    expect(result.ready).toBe(false);
    expect(result.missingCapabilities).toEqual(['not.a.real.capability']);
  });

  test('never includes raw token material in the result', async () => {
    process.env.EWS_CLIENT_ID = CLIENT_ID;
    process.env.M365_REFRESH_TOKEN = 'super-secret-refresh-token';
    await seedHealthyCache('default', { upn: 'doris@lassfolk.net' });

    const result = await computeReadiness({});
    expect(JSON.stringify(result)).not.toContain('super-secret-refresh-token');
    expect(result.secretsPrinted).toBe(false);
  });
});
