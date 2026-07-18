import { afterEach, beforeEach, describe, expect, test } from 'bun:test';
import { Buffer } from 'node:buffer';
import { mkdir, mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { buildDoctorBundle, DOCTOR_BUNDLE_SCHEMA_VERSION } from './doctor-bundle.js';
import { GRAPH_CRITICAL_DELEGATED_SCOPES } from './graph-oauth-scopes.js';
import { setDefaultProfile } from './identity-profiles.js';

const CLIENT_ID = '5f2abcea-d6ea-4460-b468-3d80d7a900eb';

function fixtureAccessToken(upn: string): string {
  const h = Buffer.from(JSON.stringify({ alg: 'none', typ: 'JWT' })).toString('base64url');
  const p = Buffer.from(
    JSON.stringify({ exp: 2_000_000_000, appid: CLIENT_ID, upn, scp: GRAPH_CRITICAL_DELEGATED_SCOPES.join(' ') })
  ).toString('base64url');
  return `${h}.${p}.x`;
}

describe('buildDoctorBundle', () => {
  let testHome: string;
  let originalEnv: NodeJS.ProcessEnv;

  beforeEach(async () => {
    testHome = await mkdtemp(join(tmpdir(), 'm365-doctor-bundle-'));
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

  test('reports missing cache/env file presence without any network call, credentials absent', async () => {
    delete process.env.EWS_CLIENT_ID;
    delete process.env.M365_REFRESH_TOKEN;
    delete process.env.GRAPH_REFRESH_TOKEN;
    delete process.env.EWS_REFRESH_TOKEN;

    const bundle = await buildDoctorBundle({ identity: 'default' });
    expect(bundle.identity.cacheFile.exists).toBe(false);
    expect(bundle.config.envFile.exists).toBe(false);
    expect(bundle.authDiagnosis.status).toBe('repair_required');
    expect(bundle.authDiagnosis.failureClass).toBe('missing_credentials');
    expect(bundle.clientId).toBeNull();
    expect(bundle.secretsPrinted).toBe(false);
    expect(bundle.unsafeFieldsIncluded).toBe(false);
  });

  test('reports cache file presence/size/mtime and healthy auth for a valid cached token', async () => {
    process.env.EWS_CLIENT_ID = CLIENT_ID;
    process.env.M365_REFRESH_TOKEN = 'fake-refresh-token';
    const dir = process.env.M365_AGENT_CLI_CONFIG_DIR as string;
    await mkdir(dir, { recursive: true });
    const cachePath = join(dir, 'token-cache-default.json');
    await writeFile(
      cachePath,
      JSON.stringify({
        version: 1,
        refreshToken: 'fake-refresh-token',
        graph: { accessToken: fixtureAccessToken('doris@lassfolk.net'), expiresAt: Date.now() + 3_600_000 }
      }),
      'utf8'
    );

    const bundle = await buildDoctorBundle({ identity: 'default' });
    expect(bundle.identity.cacheFile.exists).toBe(true);
    expect(bundle.identity.cacheFile.path).toBe(cachePath);
    expect(bundle.identity.cacheFile.sizeBytes).toBeGreaterThan(0);
    expect(bundle.identity.cacheFile.mtime).toBeTruthy();
    expect(bundle.authDiagnosis.status).toBe('healthy');
    expect(bundle.clientId).toBe(CLIENT_ID);
  });

  test('reflects the default profile and registered profile names', async () => {
    delete process.env.EWS_CLIENT_ID;
    await setDefaultProfile('doris');
    const bundle = await buildDoctorBundle({ identity: 'doris' });
    expect(bundle.profiles.defaultProfile).toBe('doris');
    expect(bundle.profiles.names).toEqual(['doris']);
  });

  test('never includes the raw refresh token, access token, or client secret material', async () => {
    process.env.EWS_CLIENT_ID = CLIENT_ID;
    process.env.M365_REFRESH_TOKEN = 'super-secret-refresh-token-value-that-must-never-leak';
    const dir = process.env.M365_AGENT_CLI_CONFIG_DIR as string;
    await mkdir(dir, { recursive: true });
    await writeFile(
      join(dir, 'token-cache-default.json'),
      JSON.stringify({
        version: 1,
        refreshToken: 'super-secret-refresh-token-value-that-must-never-leak',
        graph: { accessToken: fixtureAccessToken('doris@lassfolk.net'), expiresAt: Date.now() + 3_600_000 }
      }),
      'utf8'
    );

    const bundle = await buildDoctorBundle({ identity: 'default' });
    const serialized = JSON.stringify(bundle);
    expect(serialized).not.toContain('super-secret-refresh-token-value-that-must-never-leak');
    expect(serialized).not.toContain(fixtureAccessToken('doris@lassfolk.net'));
  });

  test('mailboxCheck is null when --mailbox was not requested', async () => {
    delete process.env.EWS_CLIENT_ID;
    const bundle = await buildDoctorBundle({ identity: 'default' });
    expect(bundle.mailboxCheck).toBeNull();
  });

  test('mailboxCheck reports checked:false without a network call when auth itself is not healthy', async () => {
    delete process.env.EWS_CLIENT_ID;
    const bundle = await buildDoctorBundle({ identity: 'default', mailbox: 'lotta@lassfolk.net' });
    expect(bundle.mailboxCheck).toEqual({ checked: false, mailbox: 'lotta@lassfolk.net', ok: null });
  });

  test('includes CLI/runtime/platform metadata', async () => {
    delete process.env.EWS_CLIENT_ID;
    const bundle = await buildDoctorBundle({ identity: 'default' });
    expect(bundle.cli.nodeVersion).toBe(process.version);
    expect(typeof bundle.cli.platform).toBe('string');
    expect(typeof bundle.cli.version).toBe('string');
    expect(bundle.schemaVersion).toBe(DOCTOR_BUNDLE_SCHEMA_VERSION);
  });
});
