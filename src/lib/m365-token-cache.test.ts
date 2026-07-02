import { afterEach, beforeEach, describe, expect, test } from 'bun:test';
import { mkdir, mkdtemp, rm, stat, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { assertValidCacheIdentity, loadM365TokenCache } from './m365-token-cache.js';

describe('assertValidCacheIdentity', () => {
  test('accepts default and common ids', () => {
    expect(assertValidCacheIdentity('default')).toBe('default');
    expect(assertValidCacheIdentity('beta_user-1')).toBe('beta_user-1');
  });

  test('rejects path injection', () => {
    expect(() => assertValidCacheIdentity('../evil')).toThrow(/Invalid token cache identity/);
    expect(() => assertValidCacheIdentity('a/b')).toThrow(/Invalid token cache identity/);
  });

  test('rejects empty and overlong', () => {
    expect(() => assertValidCacheIdentity('')).toThrow(/Invalid token cache identity/);
    expect(() => assertValidCacheIdentity('   ')).toThrow(/Invalid token cache identity/);
    expect(() => assertValidCacheIdentity('x'.repeat(129))).toThrow(/Invalid token cache identity/);
  });
});

describe('legacy cache migration re-tightens permissions', () => {
  let testHome: string;
  let originalConfigDir: string | undefined;

  beforeEach(async () => {
    testHome = await mkdtemp(join(tmpdir(), 'm365-token-cache-migrate-'));
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

  test('root graph-token-cache.json migrated from a loosely-permissioned file ends up 0600', async () => {
    const dir = process.env.M365_AGENT_CLI_CONFIG_DIR as string;
    await mkdir(dir, { recursive: true });
    const legacyPath = join(dir, 'graph-token-cache.json');
    // Simulate a pre-hardening file created under a permissive umask (world/group readable).
    await writeFile(
      legacyPath,
      JSON.stringify({ version: 1, refreshToken: 'rt', graph: { accessToken: 'a.b.c', expiresAt: Date.now() } }),
      { mode: 0o644 }
    );

    await loadM365TokenCache('default');

    const migratedPath = join(dir, 'graph-token-cache-default.json');
    const s = await stat(migratedPath);
    expect(s.mode & 0o777).toBe(0o600);
  });
});
