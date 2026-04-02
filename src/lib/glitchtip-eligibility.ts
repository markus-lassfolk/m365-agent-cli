/**
 * Gate GlitchTip: only when this install matches the latest npm release and the embedded
 * commit matches the GitHub tag v{version} (see docs/GLITCHTIP.md).
 */
import { readFileSync } from 'node:fs';
import { readFile } from 'node:fs/promises';
import { homedir } from 'node:os';
import { join } from 'node:path';
import { atomicWriteUtf8File } from './atomic-write.js';
import { COMMIT_SHA } from './git-commit.js';
import { getPackageJsonPath, getPackageVersion } from './package-info.js';

const NPM_PKG = 'm365-agent-cli';
/** How long to cache npm + GitHub tag lookups (avoid hammering registries on every invocation). */
const CACHE_TTL_MS = 60 * 60 * 1000;
const FETCH_TIMEOUT_MS = 12_000;

export interface GlitchTipEligibility {
  ok: boolean;
  reason?: string;
}

interface EligibilityCache {
  fetchedAt: number;
  npmLatest: string;
  /** Resolved commit SHA for tag v{npmLatest} */
  tagCommitSha: string | null;
}

function parseGithubRepo(repoUrl: string | undefined): { owner: string; repo: string } | null {
  if (!repoUrl) return null;
  const m = repoUrl.match(/github\.com[/:]([^/]+)\/([^/.]+)/i);
  if (!m) return null;
  return { owner: m[1], repo: m[2] };
}

async function readGithubCoords(): Promise<{ owner: string; repo: string } | null> {
  const raw = readFileSync(getPackageJsonPath(), 'utf8');
  const j = JSON.parse(raw) as { repository?: { url?: string } };
  return parseGithubRepo(j.repository?.url);
}

function cacheFile(): string {
  return join(homedir(), '.config', 'm365-agent-cli', 'cache', 'glitchtip-eligibility.json');
}

function assertEligibilityCache(data: unknown): EligibilityCache {
  if (!data || typeof data !== 'object') throw new Error('invalid eligibility cache');
  const o = data as Record<string, unknown>;
  if (typeof o.fetchedAt !== 'number' || !Number.isFinite(o.fetchedAt)) throw new Error('invalid eligibility cache');
  if (typeof o.npmLatest !== 'string' || o.npmLatest.trim().length < 3 || o.npmLatest.length > 80) {
    throw new Error('invalid eligibility cache');
  }
  const sha = o.tagCommitSha;
  if (sha !== null && (typeof sha !== 'string' || !/^[0-9a-f]{40}$/i.test(sha))) {
    throw new Error('invalid eligibility cache');
  }
  return {
    fetchedAt: o.fetchedAt,
    npmLatest: o.npmLatest.trim(),
    tagCommitSha: sha
  };
}

async function loadCache(): Promise<EligibilityCache | null> {
  try {
    const raw = await readFile(cacheFile(), 'utf8');
    return assertEligibilityCache(JSON.parse(raw));
  } catch {
    return null;
  }
}

async function saveCache(c: EligibilityCache): Promise<void> {
  const safe = assertEligibilityCache(c);
  await atomicWriteUtf8File(cacheFile(), JSON.stringify(safe, null, 2), 0o600);
}

function assertEligibilityFetchUrl(url: string): URL {
  const u = new URL(url);
  if (u.protocol !== 'https:') throw new Error('unsupported URL');
  const host = u.hostname;
  if (host !== 'registry.npmjs.org' && host !== 'api.github.com') {
    throw new Error('unsupported URL');
  }
  return u;
}

async function fetchJson<T>(url: string): Promise<{ ok: boolean; data?: T; status: number }> {
  let href: string;
  try {
    href = assertEligibilityFetchUrl(url).toString();
  } catch {
    return { ok: false, status: 0 };
  }
  const ac = new AbortController();
  const t = setTimeout(() => ac.abort(), FETCH_TIMEOUT_MS);
  try {
    const r = await fetch(href, {
      signal: ac.signal,
      headers: {
        Accept: 'application/vnd.github+json',
        'User-Agent': 'm365-agent-cli-glitchtip-eligibility'
      }
    });
    if (!r.ok) return { ok: false, status: r.status };
    const data = (await r.json()) as T;
    return { ok: true, data, status: r.status };
  } catch {
    return { ok: false, status: 0 };
  } finally {
    clearTimeout(t);
  }
}

async function fetchNpmLatestVersion(): Promise<string | null> {
  const url = `https://registry.npmjs.org/${NPM_PKG}/latest`;
  const r = await fetchJson<{ version?: string }>(url);
  if (!r.ok || !r.data?.version) return null;
  return r.data.version.trim();
}

/** Resolve annotated or lightweight tag to commit SHA. */
async function resolveTagCommitSha(owner: string, repo: string, version: string): Promise<string | null> {
  const refUrl = `https://api.github.com/repos/${owner}/${repo}/git/ref/tags/v${encodeURIComponent(version)}`;
  const ref = await fetchJson<{ object?: { sha?: string; type?: string } }>(refUrl);
  if (!ref.ok || !ref.data?.object?.sha) return null;
  const { sha, type } = ref.data.object;
  if (type === 'commit') return sha;
  if (type === 'tag') {
    const tagUrl = `https://api.github.com/repos/${owner}/${repo}/git/tags/${sha}`;
    const tag = await fetchJson<{ object?: { sha?: string; type?: string } }>(tagUrl);
    if (!tag.ok || !tag.data?.object?.sha) return null;
    return tag.data.object.sha;
  }
  return null;
}

function normalizeSha(s: string): string {
  return s.trim().toLowerCase();
}

/**
 * Returns whether GlitchTip should initialize. Uses cached npm + GitHub tag resolution (~1h).
 */
export async function checkGlitchTipEligibility(): Promise<GlitchTipEligibility> {
  if (process.env.GLITCHTIP_SKIP_VERSION_CHECK === '1' || process.env.GLITCHTIP_SKIP_VERSION_CHECK === 'true') {
    return { ok: true, reason: 'GLITCHTIP_SKIP_VERSION_CHECK' };
  }

  const currentVersion = await getPackageVersion();
  const coords = await readGithubCoords();
  if (!coords) {
    return { ok: false, reason: 'package.json repository URL could not be parsed for GitHub' };
  }

  let cache = await loadCache();
  const now = Date.now();
  let npmLatest: string;
  let tagCommitSha: string | null;

  if (cache && now - cache.fetchedAt < CACHE_TTL_MS && cache.npmLatest) {
    npmLatest = cache.npmLatest;
    tagCommitSha = cache.tagCommitSha;
  } else {
    const nv = await fetchNpmLatestVersion();
    if (!nv) {
      return { ok: false, reason: 'could not fetch latest version from npm registry' };
    }
    npmLatest = nv;
    tagCommitSha = await resolveTagCommitSha(coords.owner, coords.repo, npmLatest);
    cache = { fetchedAt: now, npmLatest, tagCommitSha };
    try {
      await saveCache(cache);
    } catch {
      // cache is optional
    }
  }

  if (currentVersion !== npmLatest) {
    return {
      ok: false,
      reason: `not on latest npm release (running ${currentVersion}, latest is ${npmLatest})`
    };
  }

  const localSha = normalizeSha(COMMIT_SHA);
  if (localSha === 'unknown') {
    if (
      process.env.GLITCHTIP_ALLOW_UNVERIFIED_COMMIT === '1' ||
      process.env.GLITCHTIP_ALLOW_UNVERIFIED_COMMIT === 'true'
    ) {
      return {
        ok: true,
        reason: 'COMMIT_SHA unknown; allowed by GLITCHTIP_ALLOW_UNVERIFIED_COMMIT (git match skipped)'
      };
    }
    return {
      ok: false,
      reason:
        'COMMIT_SHA is unknown — run `npm run embed-sha` from a git checkout before release, or set GLITCHTIP_ALLOW_UNVERIFIED_COMMIT=1'
    };
  }

  if (!tagCommitSha) {
    return {
      ok: false,
      reason: `could not resolve GitHub tag v${npmLatest} (create tag on the release commit)`
    };
  }

  if (normalizeSha(tagCommitSha) !== localSha) {
    return {
      ok: false,
      reason: `embedded commit does not match GitHub tag v${npmLatest} (local ${localSha.slice(0, 7)}… vs tag ${normalizeSha(tagCommitSha).slice(0, 7)}…)`
    };
  }

  return { ok: true, reason: `npm ${npmLatest} and commit match tag v${npmLatest}` };
}
