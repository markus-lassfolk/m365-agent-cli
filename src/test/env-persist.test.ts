import { afterEach, beforeEach, describe, expect, test } from 'bun:test';
import { mkdtemp, readFile, rm, stat, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

describe('persistRefreshTokenToEnv', () => {
  let originalEnv: NodeJS.ProcessEnv;
  let testHome: string;
  let envPath: string;

  beforeEach(async () => {
    originalEnv = { ...process.env };
    testHome = await mkdtemp(join(tmpdir(), 'm365-env-persist-'));
    envPath = join(testHome, '.env');
    delete process.env.M365_AGENT_SKIP_GLOBAL_ENV;
    delete process.env.NODE_ENV;
  });

  afterEach(async () => {
    process.env = originalEnv;
    await rm(testHome, { recursive: true, force: true }).catch(() => {});
  });

  test('writes M365_REFRESH_TOKEN and legacy aliases', async () => {
    const { persistRefreshTokenToEnv } = await import('../lib/env-persist.js');
    const wrote = await persistRefreshTokenToEnv('rotated-token', { envPath });
    expect(wrote).toBe(true);
    const content = await readFile(envPath, 'utf8');
    expect(content).toContain('M365_REFRESH_TOKEN=rotated-token');
    expect(content).toContain('GRAPH_REFRESH_TOKEN=rotated-token');
    expect(content).toContain('EWS_REFRESH_TOKEN=rotated-token');
    const mode = (await stat(envPath)).mode & 0o777;
    expect(mode).toBe(0o600);
    expect(process.env.M365_REFRESH_TOKEN).toBe('rotated-token');
  });

  test('upserts existing keys without duplicating', async () => {
    await writeFile(envPath, 'EWS_CLIENT_ID=abc\nM365_REFRESH_TOKEN=old\n', 'utf8');
    const { persistRefreshTokenToEnv } = await import('../lib/env-persist.js');
    await persistRefreshTokenToEnv('new-token', { envPath });
    const content = await readFile(envPath, 'utf8');
    expect(content.match(/M365_REFRESH_TOKEN=/g)?.length).toBe(1);
    expect(content).toContain('M365_REFRESH_TOKEN=new-token');
    expect(content).toContain('EWS_CLIENT_ID=abc');
  });

  test('skips write when refresh token unchanged', async () => {
    await writeFile(envPath, 'M365_REFRESH_TOKEN=same\n', 'utf8');
    const before = await stat(envPath);
    const { persistRefreshTokenToEnv } = await import('../lib/env-persist.js');
    const wrote = await persistRefreshTokenToEnv('same', {
      envPath,
      previousRefreshToken: 'same'
    });
    expect(wrote).toBe(false);
    const after = await stat(envPath);
    expect(after.mtimeMs).toBe(before.mtimeMs);
  });

  test('skips when M365_AGENT_SKIP_GLOBAL_ENV=1', async () => {
    process.env.M365_AGENT_SKIP_GLOBAL_ENV = '1';
    const { persistRefreshTokenToEnv } = await import('../lib/env-persist.js');
    const wrote = await persistRefreshTokenToEnv('token', { envPath });
    expect(wrote).toBe(false);
  });
});
