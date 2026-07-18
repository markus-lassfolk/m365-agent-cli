import { afterEach, beforeEach, describe, expect, test } from 'bun:test';
import { Buffer } from 'node:buffer';
import { mkdir, mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { GRAPH_CRITICAL_DELEGATED_SCOPES } from './graph-oauth-scopes.js';
import { checkIdentityGuards, resolveSignedInUpn } from './identity-guard.js';
import { setDefaultProfile } from './identity-profiles.js';

const CLIENT_ID = '5f2abcea-d6ea-4460-b468-3d80d7a900eb';

function graphFixtureAccessToken(upn: string): string {
  const h = Buffer.from(JSON.stringify({ alg: 'none', typ: 'JWT' })).toString('base64url');
  const p = Buffer.from(
    JSON.stringify({
      exp: 2_000_000_000,
      appid: CLIENT_ID,
      upn,
      scp: GRAPH_CRITICAL_DELEGATED_SCOPES.join(' ')
    })
  ).toString('base64url');
  return `${h}.${p}.x`;
}

describe('identity-guard', () => {
  let testHome: string;
  let originalEnv: NodeJS.ProcessEnv;

  beforeEach(async () => {
    testHome = await mkdtemp(join(tmpdir(), 'm365-identity-guard-'));
    originalEnv = { ...process.env };
    process.env.M365_AGENT_CLI_CONFIG_DIR = join(testHome, '.config', 'm365-agent-cli');
    process.env.EWS_TENANT_ID = 'common';
    process.env.EWS_CLIENT_ID = CLIENT_ID;
    process.env.M365_REFRESH_TOKEN = 'fake-refresh-token';
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

  async function seedCache(identity: string, upn: string): Promise<void> {
    const dir = join(process.env.M365_AGENT_CLI_CONFIG_DIR as string);
    await mkdir(dir, { recursive: true });
    await writeFile(
      join(dir, `token-cache-${identity}.json`),
      JSON.stringify({
        version: 1,
        refreshToken: 'fake-refresh-token',
        graph: { accessToken: graphFixtureAccessToken(upn), expiresAt: Date.now() + 3_600_000 }
      }),
      'utf8'
    );
  }

  test('no-op when neither --require-identity nor --as-delegate-of is set', async () => {
    const result = await checkIdentityGuards({});
    expect(result).toEqual({ ok: true });
  });

  test('resolveSignedInUpn reads the upn claim from a cached Graph token', async () => {
    await seedCache('doris', 'doris@lassfolk.net');
    expect(await resolveSignedInUpn('doris')).toBe('doris@lassfolk.net');
  });

  test('--require-identity passes when signed-in identity matches (case-insensitive)', async () => {
    await seedCache('doris', 'doris@lassfolk.net');
    const result = await checkIdentityGuards({ identity: 'doris', requireIdentity: 'DORIS@lassfolk.net' });
    expect(result).toEqual({ ok: true, signedInAs: 'doris@lassfolk.net' });
  });

  test('--require-identity fails closed on mismatch', async () => {
    await seedCache('doris', 'doris@lassfolk.net');
    const result = await checkIdentityGuards({ identity: 'doris', requireIdentity: 'lotta@lassfolk.net' });
    expect(result.ok).toBe(false);
    expect(result.message).toContain('signed in as "doris@lassfolk.net"');
    expect(result.message).toContain('lotta@lassfolk.net');
  });

  test('--require-identity fails closed when identity cannot be verified at all', async () => {
    // No cache seeded, no valid refresh token exchange possible (fetch not mocked) — must fail closed.
    global.fetch = (async () =>
      new Response(JSON.stringify({ error: 'invalid_grant', error_description: 'no' }), {
        status: 400
      })) as unknown as typeof fetch;
    const result = await checkIdentityGuards({ identity: 'ghost', requireIdentity: 'doris@lassfolk.net' });
    expect(result.ok).toBe(false);
    expect(result.message).toContain('Could not verify the signed-in identity');
  });

  test('--as-delegate-of requires --mailbox to also be set', async () => {
    const result = await checkIdentityGuards({ asDelegateOf: 'doris@lassfolk.net' });
    expect(result.ok).toBe(false);
    expect(result.message).toContain('--as-delegate-of requires --mailbox');
  });

  test('--as-delegate-of passes when identity matches and --mailbox is present', async () => {
    await seedCache('doris', 'doris@lassfolk.net');
    const result = await checkIdentityGuards({
      identity: 'doris',
      asDelegateOf: 'doris@lassfolk.net',
      mailbox: 'lotta@lassfolk.net'
    });
    expect(result).toEqual({ ok: true, signedInAs: 'doris@lassfolk.net' });
  });

  test('falls back to the default profile identity when --identity is omitted', async () => {
    await seedCache('doris-slot', 'doris@lassfolk.net');
    await setDefaultProfile('doris-slot');
    const result = await checkIdentityGuards({ requireIdentity: 'doris@lassfolk.net' });
    expect(result).toEqual({ ok: true, signedInAs: 'doris@lassfolk.net' });
  });
});
