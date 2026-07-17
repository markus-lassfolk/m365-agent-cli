/**
 * Unified OAuth token cache: one file per identity with separate EWS and Graph access tokens
 * (different audiences) and a single refresh token. See docs/GOALS.md.
 */
import { chmod, mkdir, readFile, rename, stat, unlink } from 'node:fs/promises';
import { homedir } from 'node:os';
import { join, resolve } from 'node:path';
import { atomicWriteUtf8File } from './atomic-write.js';

/**
 * Resolve at call time so tests and shells can redirect the cache without relying on `os.homedir()`
 * (some runtimes cache homedir early).
 * Optional absolute path to the `m365-agent-cli` config directory (same folder as `token-cache-*.json`).
 * Tests use this to avoid mutating global HOME/XDG across parallel cases.
 */
/**
 * Config directory for `.env`, token caches, and refresh locks.
 * Honors `M365_AGENT_CLI_CONFIG_DIR`, then `XDG_CONFIG_HOME/m365-agent-cli`, else `~/.config/m365-agent-cli`.
 */
export function getM365AgentCliConfigDir(): string {
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

function configDir(): string {
  return getM365AgentCliConfigDir();
}

function tokenCachePath(identity: string): string {
  return join(configDir(), `token-cache-${identity}.json`);
}

function graphTokenCachePath(identity: string): string {
  return join(configDir(), `graph-token-cache-${identity}.json`);
}

function legacyGraphRootCacheFile(): string {
  return join(configDir(), 'graph-token-cache.json');
}

function legacyClippyGraphCacheFile(): string {
  return join(homedir(), '.config', 'clippy', 'graph-token-cache.json');
}

/** Legacy EWS-only cache from the `clippy` package name (`auth.ts` previously wrote here). */
function legacyClippyEwsTokenCachePath(identity: string): string {
  return join(homedir(), '.config', 'clippy', `token-cache-${identity}.json`);
}

export interface TokenSlot {
  accessToken: string;
  expiresAt: number;
}

/** v1 on-disk shape for `token-cache-{identity}.json` */
export interface M365TokenCacheV1 {
  version: 1;
  /** Last refresh token returned by the token endpoint (may rotate). */
  refreshToken?: string;
  ews?: TokenSlot;
  graph?: TokenSlot;
  /**
   * True when we accepted a Graph access token that still lacks critical delegated scopes
   * after refresh (tenant may not grant them). Avoids infinite refresh loops.
   */
  graphNarrowScopeAccepted?: boolean;
}

const CACHE_IDENTITY_RE = /^[a-zA-Z0-9_-]{1,128}$/;

/** Rejects path-separator injection in cache filenames (`token-cache-{identity}.json`). */
export function assertValidCacheIdentity(identity: string): string {
  const id = identity.trim();
  if (!CACHE_IDENTITY_RE.test(id)) {
    throw new Error('Invalid token cache identity: use only letters, digits, underscore, hyphen (max 128 chars).');
  }
  return id;
}

function assertTokenSlot(o: unknown, label: string): TokenSlot {
  if (!o || typeof o !== 'object') throw new Error(`invalid ${label} slot`);
  const s = o as Record<string, unknown>;
  if (typeof s.accessToken !== 'string' || s.accessToken.length > 100_000) throw new Error(`invalid ${label} slot`);
  if (typeof s.expiresAt !== 'number' || !Number.isFinite(s.expiresAt)) throw new Error(`invalid ${label} slot`);
  return { accessToken: s.accessToken, expiresAt: s.expiresAt };
}

/** Malformed slots are treated as absent so refresh-from-env can recover (Codex / robustness). */
function tryParseTokenSlot(o: unknown, label: string): TokenSlot | undefined {
  try {
    return assertTokenSlot(o, label);
  } catch {
    return undefined;
  }
}

function isLegacyFlat(data: Record<string, unknown>): boolean {
  return (
    data.version === undefined &&
    typeof data.accessToken === 'string' &&
    typeof data.expiresAt === 'number' &&
    typeof data.refreshToken === 'string'
  );
}

/** Single refresh token: prefer M365_*, then legacy env names (same value after `login`). */
export function getUnifiedRefreshTokenFromEnv(): string | undefined {
  const m365 = process.env.M365_REFRESH_TOKEN?.trim();
  if (m365) return m365;
  const graph = process.env.GRAPH_REFRESH_TOKEN?.trim();
  if (graph) return graph;
  const ews = process.env.EWS_REFRESH_TOKEN?.trim();
  if (ews) return ews;
  return undefined;
}

async function readJsonFile(path: string): Promise<unknown | null> {
  try {
    const raw = await readFile(path, 'utf-8');
    return JSON.parse(raw) as unknown;
  } catch {
    return null;
  }
}

/**
 * Load unified cache; merges legacy `graph-token-cache-{identity}.json` into `token-cache-{identity}.json`
 * on first read (and may delete the legacy graph file after a successful merged save — caller triggers save).
 */
export async function loadM365TokenCache(identity: string): Promise<M365TokenCacheV1 | null> {
  const id = assertValidCacheIdentity(identity);
  await migrateLegacyGraphRootFiles();
  await migrateLegacyEwsClippyCache(id);

  const primaryPath = tokenCachePath(id);
  const graphPath = graphTokenCachePath(id);

  let merged: M365TokenCacheV1 = { version: 1 };
  let hadPrimary = false;

  const primary = await readJsonFile(primaryPath);
  if (primary && typeof primary === 'object') {
    const p = primary as Record<string, unknown>;
    if (p.version === 1) {
      merged = {
        version: 1,
        refreshToken: typeof p.refreshToken === 'string' ? p.refreshToken : undefined,
        ews: tryParseTokenSlot(p.ews, 'ews'),
        graph: tryParseTokenSlot(p.graph, 'graph'),
        graphNarrowScopeAccepted:
          typeof p.graphNarrowScopeAccepted === 'boolean' ? p.graphNarrowScopeAccepted : undefined
      };
      hadPrimary = true;
    } else if (isLegacyFlat(p)) {
      merged.ews = tryParseTokenSlot({ accessToken: p.accessToken, expiresAt: p.expiresAt }, 'ews');
      merged.refreshToken = p.refreshToken as string;
      hadPrimary = true;
    }
  }

  const graphOnly = await readJsonFile(graphPath);
  if (!merged.graph && graphOnly && typeof graphOnly === 'object') {
    const g = graphOnly as Record<string, unknown>;
    if (g.version === 1 && g.graph) {
      merged.graph = tryParseTokenSlot(g.graph, 'graph');
      if (!merged.refreshToken && typeof g.refreshToken === 'string') merged.refreshToken = g.refreshToken;
    } else if (isLegacyFlat(g)) {
      merged.graph = tryParseTokenSlot({ accessToken: g.accessToken, expiresAt: g.expiresAt }, 'graph');
      if (!merged.refreshToken && typeof g.refreshToken === 'string') merged.refreshToken = g.refreshToken;
    }
  }

  if (!hadPrimary && !merged.graph && !merged.ews && !merged.refreshToken) {
    return null;
  }

  return merged;
}

/** Persist unified cache; pass full object (merge in caller). */
export async function saveM365TokenCache(identity: string, cache: M365TokenCacheV1): Promise<void> {
  const id = assertValidCacheIdentity(identity);
  const safe: M365TokenCacheV1 = {
    version: 1,
    refreshToken: cache.refreshToken,
    ews: cache.ews,
    graph: cache.graph,
    graphNarrowScopeAccepted: cache.graphNarrowScopeAccepted
  };
  await mkdir(configDir(), { recursive: true, mode: 0o700 });
  await atomicWriteUtf8File(tokenCachePath(id), JSON.stringify(safe, null, 2), 0o600);

  try {
    await unlink(graphTokenCachePath(id));
  } catch {
    // no legacy file
  }
}

async function migrateLegacyGraphRootFiles(): Promise<void> {
  try {
    const defaultPath = graphTokenCachePath('default');
    const defaultStats = await stat(defaultPath).catch(() => null);
    if (!defaultStats) {
      const rootLegacy = legacyGraphRootCacheFile();
      const legacyStats = await stat(rootLegacy).catch(() => null);
      if (legacyStats) {
        await mkdir(configDir(), { recursive: true, mode: 0o700 });
        await chmod(configDir(), 0o700).catch(() => {});
        await rename(rootLegacy, defaultPath);
        // `rename` preserves the source file's mode, which may predate 0600 hardening — re-tighten it.
        await chmod(defaultPath, 0o600).catch(() => {});
        return;
      }
      const clippyGraph = legacyClippyGraphCacheFile();
      const oldClippyStats = await stat(clippyGraph).catch(() => null);
      if (oldClippyStats) {
        await mkdir(configDir(), { recursive: true, mode: 0o700 });
        await chmod(configDir(), 0o700).catch(() => {});
        await rename(clippyGraph, defaultPath);
        await chmod(defaultPath, 0o600).catch(() => {});
      }
    }
  } catch {
    // ignore
  }
}

/** One-time: move `~/.config/clippy/token-cache-{id}.json` into the unified cache path when the latter is absent. */
async function migrateLegacyEwsClippyCache(identity: string): Promise<void> {
  try {
    const dest = tokenCachePath(identity);
    const destStats = await stat(dest).catch(() => null);
    if (destStats) return;

    const legacy = legacyClippyEwsTokenCachePath(identity);
    const legacyStats = await stat(legacy).catch(() => null);
    if (!legacyStats) return;

    await mkdir(configDir(), { recursive: true, mode: 0o700 });
    await chmod(configDir(), 0o700).catch(() => {});
    await rename(legacy, dest);
    // `rename` preserves the source file's mode, which may predate 0600 hardening — re-tighten it.
    await chmod(dest, 0o600).catch(() => {});
  } catch {
    // ignore
  }
}
