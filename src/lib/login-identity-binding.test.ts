import { afterEach, beforeEach, describe, expect, test } from 'bun:test';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { getProfile } from './identity-profiles.js';
import { bindLoginIdentityOrThrow, LoginAccountMismatchError } from './login-identity-binding.js';

describe('bindLoginIdentityOrThrow', () => {
  let testHome: string;
  let originalConfigDir: string | undefined;

  beforeEach(async () => {
    testHome = await mkdtemp(join(tmpdir(), 'm365-login-binding-'));
    originalConfigDir = process.env.M365_AGENT_CLI_CONFIG_DIR;
    process.env.M365_AGENT_CLI_CONFIG_DIR = join(testHome, '.config', 'm365-agent-cli');
  });

  afterEach(async () => {
    if (originalConfigDir === undefined) delete process.env.M365_AGENT_CLI_CONFIG_DIR;
    else process.env.M365_AGENT_CLI_CONFIG_DIR = originalConfigDir;
    await rm(testHome, { recursive: true, force: true }).catch(() => {});
  });

  test('no-op when no --identity is given', async () => {
    await bindLoginIdentityOrThrow({ signedInAs: 'doris@lassfolk.net' });
    expect(await getProfile('default')).toBeUndefined();
  });

  test('first login into a slug registers it without any mismatch check', async () => {
    await bindLoginIdentityOrThrow({ identity: 'doris', signedInAs: 'doris@lassfolk.net' });
    const profile = await getProfile('doris');
    expect(profile?.signedInAs).toBe('doris@lassfolk.net');
    expect(profile?.lastVerifiedAt).toBeTruthy();
  });

  test('re-login with the same account updates lastVerifiedAt without error', async () => {
    await bindLoginIdentityOrThrow({ identity: 'doris', signedInAs: 'doris@lassfolk.net' });
    await bindLoginIdentityOrThrow({ identity: 'doris', signedInAs: 'DORIS@lassfolk.net' });
    const profile = await getProfile('doris');
    expect(profile?.signedInAs).toBe('DORIS@lassfolk.net');
  });

  test('refuses to complete when a different account lands on an already-verified slug', async () => {
    await bindLoginIdentityOrThrow({ identity: 'doris', signedInAs: 'doris@lassfolk.net' });
    await expect(bindLoginIdentityOrThrow({ identity: 'doris', signedInAs: 'lotta@lassfolk.net' })).rejects.toThrow(
      LoginAccountMismatchError
    );
    // The profile must NOT have been silently rebound.
    const profile = await getProfile('doris');
    expect(profile?.signedInAs).toBe('doris@lassfolk.net');
  });

  test('--force-identity-switch allows intentionally rebinding the slug', async () => {
    await bindLoginIdentityOrThrow({ identity: 'doris', signedInAs: 'doris@lassfolk.net' });
    await bindLoginIdentityOrThrow({ identity: 'doris', signedInAs: 'lotta@lassfolk.net', force: true });
    const profile = await getProfile('doris');
    expect(profile?.signedInAs).toBe('lotta@lassfolk.net');
  });

  test('registers the slug even when no UPN could be decoded', async () => {
    await bindLoginIdentityOrThrow({ identity: 'doris' });
    const profile = await getProfile('doris');
    expect(profile?.identity).toBe('doris');
    expect(profile?.signedInAs).toBeUndefined();
  });
});
