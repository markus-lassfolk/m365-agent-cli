import { afterEach, beforeEach, describe, expect, test } from 'bun:test';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import {
  assertValidProfileName,
  deleteProfile,
  getDefaultProfileIdentity,
  getDefaultProfileName,
  getProfile,
  listProfiles,
  probeCacheHealth,
  setDefaultProfile,
  upsertProfile
} from './identity-profiles.js';
import { saveM365TokenCache } from './m365-token-cache.js';

describe('identity-profiles', () => {
  let testHome: string;
  let originalConfigDir: string | undefined;

  beforeEach(async () => {
    testHome = await mkdtemp(join(tmpdir(), 'm365-profiles-'));
    originalConfigDir = process.env.M365_AGENT_CLI_CONFIG_DIR;
    process.env.M365_AGENT_CLI_CONFIG_DIR = join(testHome, '.config', 'm365-agent-cli');
  });

  afterEach(async () => {
    if (originalConfigDir === undefined) {
      delete process.env.M365_AGENT_CLI_CONFIG_DIR;
    } else {
      process.env.M365_AGENT_CLI_CONFIG_DIR = originalConfigDir;
    }
    await rm(testHome, { recursive: true, force: true }).catch(() => {});
  });

  test('assertValidProfileName rejects path injection and empty names', () => {
    expect(() => assertValidProfileName('../evil')).toThrow(/Invalid profile name/);
    expect(() => assertValidProfileName('')).toThrow(/Invalid profile name/);
    expect(assertValidProfileName('doris')).toBe('doris');
  });

  test('listProfiles is empty before any profile is created', async () => {
    expect(await listProfiles()).toEqual([]);
    expect(await getDefaultProfileName()).toBeUndefined();
    expect(await getDefaultProfileIdentity()).toBeUndefined();
  });

  test('upsertProfile creates then updates a profile record', async () => {
    const created = await upsertProfile('doris', { tenantId: 'tenant-1' });
    expect(created.name).toBe('doris');
    expect(created.identity).toBe('doris');
    expect(created.tenantId).toBe('tenant-1');
    expect(created.createdAt).toBeTruthy();

    const updated = await upsertProfile('doris', { signedInAs: 'doris@lassfolk.net' });
    expect(updated.tenantId).toBe('tenant-1'); // preserved
    expect(updated.signedInAs).toBe('doris@lassfolk.net');
    expect(updated.createdAt).toBe(created.createdAt); // not reset on update

    const fetched = await getProfile('doris');
    expect(fetched?.signedInAs).toBe('doris@lassfolk.net');
  });

  test('setDefaultProfile auto-registers an unknown profile name and default profile selection resolves its identity', async () => {
    const entry = await setDefaultProfile('doris');
    expect(entry.identity).toBe('doris');
    expect(await getDefaultProfileName()).toBe('doris');
    expect(await getDefaultProfileIdentity()).toBe('doris');
  });

  test('setDefaultProfile switches default across multiple profiles', async () => {
    await upsertProfile('doris');
    await upsertProfile('lotta', { identity: 'lotta-cache' });
    await setDefaultProfile('doris');
    expect(await getDefaultProfileIdentity()).toBe('doris');
    await setDefaultProfile('lotta');
    expect(await getDefaultProfileIdentity()).toBe('lotta-cache');
  });

  test('deleteProfile removes metadata and clears default when the default is deleted', async () => {
    await setDefaultProfile('doris');
    expect(await deleteProfile('doris')).toBe(true);
    expect(await getProfile('doris')).toBeUndefined();
    expect(await getDefaultProfileName()).toBeUndefined();
    // deleting an unknown profile is a no-op, not an error
    expect(await deleteProfile('doris')).toBe(false);
  });

  test('probeCacheHealth reflects offline cache state without any network call', async () => {
    expect(await probeCacheHealth('doris')).toBe('missing');

    await saveM365TokenCache('doris', {
      version: 1,
      refreshToken: 'rt',
      graph: { accessToken: 'a.b.c', expiresAt: Date.now() + 60_000 }
    });
    expect(await probeCacheHealth('doris')).toBe('healthy');

    await saveM365TokenCache('doris', {
      version: 1,
      refreshToken: 'rt',
      graph: { accessToken: 'a.b.c', expiresAt: Date.now() - 60_000 }
    });
    expect(await probeCacheHealth('doris')).toBe('expired');
  });
});
