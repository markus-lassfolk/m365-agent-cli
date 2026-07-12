/**
 * Opt-in on-disk cache for idempotent (`GET`) Microsoft Graph responses, keyed by a hash of the
 * bearer token + method + URL so different signed-in identities never share cache entries.
 * Enabled by setting `M365_CACHE_TTL` (synced from the root `--cache <duration>` flag); a miss or
 * disabled cache is a normal pass-through — cache errors never fail the underlying Graph call.
 */
import { createHash } from 'node:crypto';
import { mkdir, readdir, readFile, unlink } from 'node:fs/promises';
import { homedir } from 'node:os';
import { join, resolve } from 'node:path';
import { atomicWriteUtf8File } from './atomic-write.js';

/** Resolve at call time (not module load) so tests can redirect the cache dir per-case. */
function configDir(): string {
  const dirOverride = process.env.M365_AGENT_CLI_CONFIG_DIR?.trim();
  if (dirOverride) {
    return resolve(dirOverride);
  }
  const xdg = process.env.XDG_CONFIG_HOME?.trim();
  if (xdg) {
    return join(xdg, 'm365-agent-cli');
  }
  return join(homedir(), '.config', 'm365-agent-cli');
}

function cacheDir(): string {
  return join(configDir(), 'graph-cache');
}

const DURATION_RE = /^(\d+)(ms|s|m|h|d)?$/i;
const DURATION_UNIT_MS: Record<string, number> = { ms: 1, s: 1000, m: 60_000, h: 3_600_000, d: 86_400_000 };

/** Parses `"30"` (bare number = seconds), `"30s"`, `"5m"`, `"2h"`, `"1d"` into milliseconds; `null` if unset/invalid. */
export function parseCacheTtlMs(raw: string | undefined | null): number | null {
  const v = raw?.trim();
  if (!v) return null;
  const m = DURATION_RE.exec(v);
  if (!m) return null;
  const n = Number(m[1]);
  if (!Number.isFinite(n) || n <= 0) return null;
  const unit = (m[2] || 's').toLowerCase();
  return n * DURATION_UNIT_MS[unit];
}

/** TTL in ms from `M365_CACHE_TTL`, or `null` when unset/invalid (cache disabled). */
export function activeCacheTtlMs(): number | null {
  return parseCacheTtlMs(process.env.M365_CACHE_TTL);
}

export function isCacheEnabled(): boolean {
  return activeCacheTtlMs() !== null;
}

export interface CachedGraphEntry {
  status: number;
  body: unknown;
  cachedAt: number;
  expiresAt: number;
}

/** Headers (besides Authorization, already covered by `token`) that can change a GET's response
 *  body for the same method+URL — e.g. `Prefer: outlook.timezone=...`, `ConsistencyLevel:
 *  eventual`, `Accept-Language`. Sorted/lowercased so header order/case never affects the key. */
function normalizeHeaders(headers?: HeadersInit): string {
  if (!headers) return '';
  const pairs: Array<[string, string]> = [];
  new Headers(headers).forEach((value, key) => {
    if (key.toLowerCase() === 'authorization') return;
    pairs.push([key.toLowerCase(), value]);
  });
  // JSON-encode the sorted [key, value] tuples rather than joining with "key=value&..." — a raw
  // delimiter join lets an unrelated header set collide onto the same string (and thus the same
  // cache key) whenever a header's own value contains "&" or "=".
  pairs.sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0));
  return JSON.stringify(pairs);
}

function cacheKey(token: string, method: string, url: string, headers?: HeadersInit): string {
  return createHash('sha256')
    .update(`${token} ${method.toUpperCase()} ${url} ${normalizeHeaders(headers)}`)
    .digest('hex');
}

function cacheFilePath(key: string): string {
  return join(cacheDir(), `${key}.json`);
}

function isExpired(entry: Pick<CachedGraphEntry, 'expiresAt'>, now: number): boolean {
  return typeof entry.expiresAt !== 'number' || now >= entry.expiresAt;
}

/** Returns the cached entry for `GET token+method+url+headers`, or `null` on a miss, expiry, or any read/parse error. */
export async function readGraphCache(
  token: string,
  method: string,
  url: string,
  headers?: HeadersInit
): Promise<CachedGraphEntry | null> {
  try {
    const raw = await readFile(cacheFilePath(cacheKey(token, method, url, headers)), 'utf8');
    const parsed = JSON.parse(raw) as CachedGraphEntry;
    if (isExpired(parsed, Date.now())) return null;
    return parsed;
  } catch {
    return null;
  }
}

/** Best-effort write; failures (e.g. read-only filesystem) are swallowed so caching never breaks the underlying call. */
export async function writeGraphCache(
  token: string,
  method: string,
  url: string,
  status: number,
  body: unknown,
  ttlMs: number,
  headers?: HeadersInit
): Promise<void> {
  try {
    const now = Date.now();
    const entry: CachedGraphEntry = { status, body, cachedAt: now, expiresAt: now + ttlMs };
    await mkdir(cacheDir(), { recursive: true, mode: 0o700 });
    await atomicWriteUtf8File(cacheFilePath(cacheKey(token, method, url, headers)), JSON.stringify(entry), 0o600);
    await pruneExpiredCacheEntries(now);
  } catch {
    // Cache writes are best-effort.
  }
}

/** Opportunistically deletes expired entries so the cache directory doesn't grow without bound. */
async function pruneExpiredCacheEntries(now: number): Promise<void> {
  try {
    const dir = cacheDir();
    const names = await readdir(dir);
    for (const name of names) {
      if (!name.endsWith('.json')) continue;
      const path = join(dir, name);
      try {
        const raw = await readFile(path, 'utf8');
        const parsed = JSON.parse(raw) as CachedGraphEntry;
        if (isExpired(parsed, now)) {
          await unlink(path).catch(() => {});
        }
      } catch {
        await unlink(path).catch(() => {});
      }
    }
  } catch {
    // Directory may not exist yet, or listing may fail transiently — non-fatal.
  }
}
