/**
 * Exclusive lock around refresh-token exchange for a given identity.
 *
 * Graph and EWS share one refresh token in `token-cache-{identity}.json`. Concurrent CLI
 * processes can otherwise both redeem the same RT, Entra rotates/invalidates one grant, and
 * the loser persists a stale RT (or leaves `.env`/cache divergent).
 *
 * Lock file: `{configDir}/.refresh-{identity}.lock`
 */
import { mkdir, open, readFile, unlink } from 'node:fs/promises';
import { join } from 'node:path';
import { getM365AgentCliConfigDir } from './m365-token-cache.js';

const DEFAULT_STALE_MS = 120_000;
const DEFAULT_MAX_WAIT_MS = 60_000;
const DEFAULT_POLL_MS = 50;

export function refreshTokenLockPath(identity: string): string {
  return join(getM365AgentCliConfigDir(), `.refresh-${identity}.lock`);
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function processAlive(pid: number): boolean {
  if (!Number.isFinite(pid) || pid <= 0) return false;
  try {
    process.kill(pid, 0);
    return true;
  } catch {
    return false;
  }
}

async function lockLooksStale(lockPath: string, staleMs: number): Promise<boolean> {
  try {
    const raw = await readFile(lockPath, 'utf8');
    const lines = raw.trim().split(/\r?\n/);
    const pid = Number(lines[0]);
    const ts = Number(lines[1]);
    if (!Number.isFinite(ts)) return true;
    if (Date.now() - ts > staleMs) return true;
    if (Number.isFinite(pid) && pid > 0 && !processAlive(pid)) return true;
    return false;
  } catch (err) {
    if (err && typeof err === 'object' && 'code' in err && (err as NodeJS.ErrnoException).code === 'ENOENT') {
      return true;
    }
    // Unreadable lock — treat as stale so we do not wedge forever.
    return true;
  }
}

async function tryAcquireExclusive(lockPath: string): Promise<boolean> {
  try {
    const fh = await open(lockPath, 'wx');
    try {
      await fh.writeFile(`${process.pid}\n${Date.now()}\n`, 'utf8');
    } finally {
      await fh.close();
    }
    return true;
  } catch (err) {
    if (err && typeof err === 'object' && 'code' in err && (err as NodeJS.ErrnoException).code === 'EEXIST') {
      return false;
    }
    throw err;
  }
}

export type RefreshTokenLockOptions = {
  /** Consider a lock stale after this age (or sooner if holder PID is gone). Default 120s. */
  staleMs?: number;
  /** Give up waiting after this long. Default 60s. */
  maxWaitMs?: number;
  /** Poll interval while waiting. Default 50ms. */
  pollMs?: number;
};

/**
 * Run `fn` while holding an exclusive per-identity refresh lock.
 * Callers should re-load cache (and short-circuit on a fresh access token) inside `fn`.
 */
export async function withRefreshTokenLock<T>(
  identity: string,
  fn: () => Promise<T>,
  options?: RefreshTokenLockOptions
): Promise<T> {
  if (!/^[a-zA-Z0-9_-]+$/.test(identity)) {
    throw new Error('Invalid identity name for refresh lock.');
  }

  const staleMs = options?.staleMs ?? DEFAULT_STALE_MS;
  const maxWaitMs = options?.maxWaitMs ?? DEFAULT_MAX_WAIT_MS;
  const pollMs = options?.pollMs ?? DEFAULT_POLL_MS;
  const lockPath = refreshTokenLockPath(identity);
  const dir = getM365AgentCliConfigDir();
  await mkdir(dir, { recursive: true, mode: 0o700 });

  const started = Date.now();
  let acquired = false;

  while (!acquired) {
    acquired = await tryAcquireExclusive(lockPath);
    if (acquired) break;

    if (await lockLooksStale(lockPath, staleMs)) {
      await unlink(lockPath).catch(() => {});
      continue;
    }

    if (Date.now() - started > maxWaitMs) {
      throw new Error(
        `Timed out after ${maxWaitMs}ms waiting for M365 refresh lock (${lockPath}). ` +
          'Another m365-agent-cli process may be stuck refreshing tokens.'
      );
    }
    await sleep(pollMs);
  }

  try {
    return await fn();
  } finally {
    await unlink(lockPath).catch(() => {});
  }
}
