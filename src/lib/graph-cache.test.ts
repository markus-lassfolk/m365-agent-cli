import { afterEach, beforeEach, describe, expect, it } from 'bun:test';
import { mkdtemp, readdir, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { activeCacheTtlMs, isCacheEnabled, parseCacheTtlMs, readGraphCache, writeGraphCache } from './graph-cache.js';

describe('parseCacheTtlMs', () => {
  it('treats a bare number as seconds', () => {
    expect(parseCacheTtlMs('30')).toBe(30_000);
  });

  it('parses s/m/h/d/ms suffixes', () => {
    expect(parseCacheTtlMs('45s')).toBe(45_000);
    expect(parseCacheTtlMs('5m')).toBe(300_000);
    expect(parseCacheTtlMs('2h')).toBe(7_200_000);
    expect(parseCacheTtlMs('1d')).toBe(86_400_000);
    expect(parseCacheTtlMs('500ms')).toBe(500);
  });

  it('is case-insensitive on the unit', () => {
    expect(parseCacheTtlMs('5M')).toBe(300_000);
  });

  it('returns null for unset, empty, zero, negative, or malformed input', () => {
    expect(parseCacheTtlMs(undefined)).toBeNull();
    expect(parseCacheTtlMs(null)).toBeNull();
    expect(parseCacheTtlMs('')).toBeNull();
    expect(parseCacheTtlMs('0')).toBeNull();
    expect(parseCacheTtlMs('-5s')).toBeNull();
    expect(parseCacheTtlMs('soon')).toBeNull();
    expect(parseCacheTtlMs('5x')).toBeNull();
  });
});

describe('activeCacheTtlMs / isCacheEnabled', () => {
  const original = process.env.M365_CACHE_TTL;

  afterEach(() => {
    if (original === undefined) delete process.env.M365_CACHE_TTL;
    else process.env.M365_CACHE_TTL = original;
  });

  it('is disabled when M365_CACHE_TTL is unset', () => {
    delete process.env.M365_CACHE_TTL;
    expect(isCacheEnabled()).toBe(false);
    expect(activeCacheTtlMs()).toBeNull();
  });

  it('is enabled with a valid duration', () => {
    process.env.M365_CACHE_TTL = '1m';
    expect(isCacheEnabled()).toBe(true);
    expect(activeCacheTtlMs()).toBe(60_000);
  });
});

describe('readGraphCache / writeGraphCache', () => {
  let testHome: string;
  let originalConfigDir: string | undefined;

  beforeEach(async () => {
    testHome = await mkdtemp(join(tmpdir(), 'm365-graph-cache-'));
    originalConfigDir = process.env.M365_AGENT_CLI_CONFIG_DIR;
    process.env.M365_AGENT_CLI_CONFIG_DIR = testHome;
  });

  afterEach(async () => {
    if (originalConfigDir === undefined) delete process.env.M365_AGENT_CLI_CONFIG_DIR;
    else process.env.M365_AGENT_CLI_CONFIG_DIR = originalConfigDir;
    await rm(testHome, { recursive: true, force: true }).catch(() => {});
  });

  it('misses when nothing was written', async () => {
    expect(await readGraphCache('tok', 'GET', 'https://graph.microsoft.com/v1.0/me')).toBeNull();
  });

  it('round-trips a written entry', async () => {
    await writeGraphCache('tok', 'GET', 'https://graph.microsoft.com/v1.0/me', 200, { id: 'u1' }, 60_000);
    const entry = await readGraphCache('tok', 'GET', 'https://graph.microsoft.com/v1.0/me');
    expect(entry).not.toBeNull();
    expect(entry?.status).toBe(200);
    expect(entry?.body).toEqual({ id: 'u1' });
  });

  it('is a miss once the TTL has elapsed', async () => {
    await writeGraphCache('tok', 'GET', 'https://graph.microsoft.com/v1.0/me', 200, { id: 'u1' }, -1);
    expect(await readGraphCache('tok', 'GET', 'https://graph.microsoft.com/v1.0/me')).toBeNull();
  });

  it('isolates entries by token (different identity => different cache)', async () => {
    await writeGraphCache('tok-a', 'GET', 'https://graph.microsoft.com/v1.0/me', 200, { id: 'a' }, 60_000);
    expect(await readGraphCache('tok-b', 'GET', 'https://graph.microsoft.com/v1.0/me')).toBeNull();
  });

  it('isolates entries by URL', async () => {
    await writeGraphCache('tok', 'GET', 'https://graph.microsoft.com/v1.0/me', 200, { id: 'a' }, 60_000);
    expect(await readGraphCache('tok', 'GET', 'https://graph.microsoft.com/v1.0/me/messages')).toBeNull();
  });

  it('isolates entries by headers (bug regression: Prefer/ConsistencyLevel must not collide)', async () => {
    const url = 'https://graph.microsoft.com/v1.0/me/events';
    await writeGraphCache('tok', 'GET', url, 200, { tz: 'UTC' }, 60_000, {
      Prefer: 'outlook.timezone="UTC"'
    });
    await writeGraphCache('tok', 'GET', url, 200, { tz: 'America/New_York' }, 60_000, {
      Prefer: 'outlook.timezone="America/New_York"'
    });

    const utcEntry = await readGraphCache('tok', 'GET', url, { Prefer: 'outlook.timezone="UTC"' });
    const nyEntry = await readGraphCache('tok', 'GET', url, { Prefer: 'outlook.timezone="America/New_York"' });
    expect(utcEntry?.body).toEqual({ tz: 'UTC' });
    expect(nyEntry?.body).toEqual({ tz: 'America/New_York' });

    // No headers at all is a third, distinct cache slot.
    expect(await readGraphCache('tok', 'GET', url)).toBeNull();
  });

  it('is unaffected by header key case or declaration order', async () => {
    const url = 'https://graph.microsoft.com/v1.0/me/events';
    await writeGraphCache('tok', 'GET', url, 200, { ok: true }, 60_000, {
      Prefer: 'a',
      ConsistencyLevel: 'eventual'
    });
    const entry = await readGraphCache('tok', 'GET', url, {
      consistencylevel: 'eventual',
      prefer: 'a'
    });
    expect(entry?.body).toEqual({ ok: true });
  });

  it('does not collide when a header value itself contains "&" or "=" (bug regression)', async () => {
    const url = 'https://graph.microsoft.com/v1.0/me/events';
    await writeGraphCache('tok', 'GET', url, 200, { which: 'combined-header' }, 60_000, {
      'X-Test': 'a&y=1'
    });
    await writeGraphCache('tok', 'GET', url, 200, { which: 'split-headers' }, 60_000, {
      'X-Test': 'a',
      y: '1'
    });

    const combined = await readGraphCache('tok', 'GET', url, { 'X-Test': 'a&y=1' });
    const split = await readGraphCache('tok', 'GET', url, { 'X-Test': 'a', y: '1' });
    expect(combined?.body).toEqual({ which: 'combined-header' });
    expect(split?.body).toEqual({ which: 'split-headers' });
  });

  it('prunes expired entries opportunistically on every write, keeping only live entries', async () => {
    const dir = join(testHome, 'graph-cache');

    // Already-expired at write time: the opportunistic prune inside this same write removes it.
    await writeGraphCache('tok', 'GET', 'https://graph.microsoft.com/v1.0/stale', 200, { x: 1 }, -1);
    expect((await readdir(dir)).length).toBe(0);

    // A later write for a live entry must not be pruned by that same opportunistic pass.
    await writeGraphCache('tok', 'GET', 'https://graph.microsoft.com/v1.0/fresh', 200, { x: 2 }, 60_000);
    const names = await readdir(dir);
    expect(names.length).toBe(1);
    expect(await readGraphCache('tok', 'GET', 'https://graph.microsoft.com/v1.0/fresh')).not.toBeNull();
  });
});
