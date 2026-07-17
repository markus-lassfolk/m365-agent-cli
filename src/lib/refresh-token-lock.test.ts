import { afterEach, beforeEach, describe, expect, test } from 'bun:test';
import { mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { refreshTokenLockPath, withRefreshTokenLock } from './refresh-token-lock.js';

describe('withRefreshTokenLock', () => {
  let originalEnv: NodeJS.ProcessEnv;
  let configDir: string;

  beforeEach(async () => {
    originalEnv = { ...process.env };
    const root = await mkdtemp(join(tmpdir(), 'm365-refresh-lock-'));
    configDir = join(root, 'cfg');
    await mkdir(configDir, { recursive: true });
    process.env.M365_AGENT_CLI_CONFIG_DIR = configDir;
  });

  afterEach(async () => {
    for (const key of Object.keys(process.env)) {
      if (!(key in originalEnv)) delete process.env[key];
    }
    for (const [key, value] of Object.entries(originalEnv)) {
      if (value === undefined) delete process.env[key];
      else process.env[key] = value;
    }
    await rm(configDir, { recursive: true, force: true }).catch(() => {});
  });

  test('serializes concurrent critical sections for the same identity', async () => {
    const order: string[] = [];
    const releaseA = Promise.withResolvers<void>();

    const a = withRefreshTokenLock('default', async () => {
      order.push('a-enter');
      await releaseA.promise;
      order.push('a-exit');
      return 'a';
    });

    // Wait until A holds the lock.
    for (let i = 0; i < 40 && order[0] !== 'a-enter'; i++) {
      await Bun.sleep(5);
    }
    expect(order[0]).toBe('a-enter');

    const bStarted = Promise.withResolvers<void>();
    const b = withRefreshTokenLock('default', async () => {
      bStarted.resolve();
      order.push('b-run');
      return 'b';
    });

    // B must not enter while A holds the lock.
    await Bun.sleep(30);
    expect(order).toEqual(['a-enter']);

    releaseA.resolve();
    await expect(a).resolves.toBe('a');
    await expect(b).resolves.toBe('b');
    await bStarted.promise;
    expect(order).toEqual(['a-enter', 'a-exit', 'b-run']);
  });

  test('removes a stale lock from a dead holder and proceeds', async () => {
    const lockPath = refreshTokenLockPath('default');
    // PID 1 is usually init and alive; use an extremely unlikely dead PID with old timestamp.
    await writeFile(lockPath, `999999999\n${Date.now() - 200_000}\n`, 'utf8');

    const result = await withRefreshTokenLock('default', async () => 'ok', { staleMs: 1000, maxWaitMs: 5000 });
    expect(result).toBe('ok');
    await expect(readFile(lockPath, 'utf8')).rejects.toMatchObject({ code: 'ENOENT' });
  });

  test('rejects invalid identity names', async () => {
    await expect(withRefreshTokenLock('bad id', async () => 1)).rejects.toThrow(/Invalid identity/);
  });
});
