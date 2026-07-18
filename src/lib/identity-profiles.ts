/**
 * Named identity profiles: a small registry on top of the existing per-identity token cache
 * (`m365-token-cache.ts`) that lets an operator/agent bind a friendly name to a cache identity
 * slot, mark one profile as the default (used whenever a command's `--identity` is omitted), and
 * record last-verified metadata (signed-in UPN, tenant, capabilities) without an extra network
 * call on every read.
 *
 * Storage: `{configDir}/profiles.json`. Never stores tokens — only references the existing
 * `token-cache-{identity}.json` slot by name.
 */
import { mkdir, readFile, unlink } from 'node:fs/promises';
import { join } from 'node:path';
import { atomicWriteUtf8File } from './atomic-write.js';
import { assertValidCacheIdentity, getM365AgentCliConfigDir, loadM365TokenCache } from './m365-token-cache.js';

function profilesFilePath(): string {
  return join(getM365AgentCliConfigDir(), 'profiles.json');
}

export interface ProfileEntry {
  /** Profile name — also used as the token-cache identity slot unless `identity` overrides it. */
  name: string;
  /** Token-cache identity slot this profile binds to (defaults to `name`). */
  identity: string;
  /** Optional `--env-file` path this profile's credentials live in (e.g. a second app/tenant). */
  envFile?: string;
  createdAt: string;
  /** Last time a live check (login, `profiles show --verify`, `readiness`, `auth repair`) confirmed this identity. */
  lastVerifiedAt?: string;
  /** UPN/email last observed signed in for this profile — used to detect account mismatches on re-login. */
  signedInAs?: string;
  tenantId?: string;
}

export interface ProfilesFileV1 {
  version: 1;
  defaultProfile?: string;
  profiles: Record<string, ProfileEntry>;
}

const PROFILE_NAME_RE = /^[a-zA-Z0-9_-]{1,128}$/;

export function assertValidProfileName(name: string): string {
  const n = name.trim();
  if (!PROFILE_NAME_RE.test(n)) {
    throw new Error('Invalid profile name: use only letters, digits, underscore, hyphen (max 128 chars).');
  }
  return n;
}

async function readProfilesFile(): Promise<ProfilesFileV1> {
  try {
    const raw = await readFile(profilesFilePath(), 'utf8');
    const parsed = JSON.parse(raw) as unknown;
    if (parsed && typeof parsed === 'object' && (parsed as { version?: unknown }).version === 1) {
      const p = parsed as ProfilesFileV1;
      return {
        version: 1,
        defaultProfile: typeof p.defaultProfile === 'string' ? p.defaultProfile : undefined,
        profiles: p.profiles && typeof p.profiles === 'object' ? p.profiles : {}
      };
    }
  } catch {
    // missing or malformed — treat as empty
  }
  return { version: 1, profiles: {} };
}

async function writeProfilesFile(data: ProfilesFileV1): Promise<void> {
  await mkdir(getM365AgentCliConfigDir(), { recursive: true, mode: 0o700 });
  await atomicWriteUtf8File(profilesFilePath(), JSON.stringify(data, null, 2), 0o600);
}

/** List all registered profiles (empty array when none registered yet). */
export async function listProfiles(): Promise<ProfileEntry[]> {
  const data = await readProfilesFile();
  return Object.values(data.profiles);
}

export async function getProfile(name: string): Promise<ProfileEntry | undefined> {
  const n = assertValidProfileName(name);
  const data = await readProfilesFile();
  return data.profiles[n];
}

/** Name of the default profile, or undefined when none is set. */
export async function getDefaultProfileName(): Promise<string | undefined> {
  const data = await readProfilesFile();
  return data.defaultProfile;
}

/**
 * Resolve the token-cache identity slot that should be used when a command's `--identity` flag
 * is omitted: the default profile's bound identity, or `undefined` when no default profile is
 * set (callers fall back to the literal `'default'` cache slot, unchanged from prior behavior).
 */
export async function getDefaultProfileIdentity(): Promise<string | undefined> {
  const data = await readProfilesFile();
  if (!data.defaultProfile) return undefined;
  return data.profiles[data.defaultProfile]?.identity;
}

/**
 * Create or update a profile record. Creates the profile if it does not exist yet (profiles are
 * lightweight references to a cache identity, so `set-default`/login can register on first use).
 */
export async function upsertProfile(
  name: string,
  fields: Partial<Omit<ProfileEntry, 'name' | 'createdAt'>> = {}
): Promise<ProfileEntry> {
  const n = assertValidProfileName(name);
  const identity = fields.identity ? assertValidCacheIdentity(fields.identity) : n;
  const data = await readProfilesFile();
  const existing = data.profiles[n];
  const entry: ProfileEntry = {
    name: n,
    identity,
    envFile: fields.envFile ?? existing?.envFile,
    createdAt: existing?.createdAt ?? new Date().toISOString(),
    lastVerifiedAt: fields.lastVerifiedAt ?? existing?.lastVerifiedAt,
    signedInAs: fields.signedInAs ?? existing?.signedInAs,
    tenantId: fields.tenantId ?? existing?.tenantId
  };
  data.profiles[n] = entry;
  await writeProfilesFile(data);
  return entry;
}

/** Set the default profile. Auto-registers the profile (identity = name) if not already known. */
export async function setDefaultProfile(name: string): Promise<ProfileEntry> {
  const n = assertValidProfileName(name);
  const data = await readProfilesFile();
  if (!data.profiles[n]) {
    data.profiles[n] = { name: n, identity: n, createdAt: new Date().toISOString() };
  }
  data.defaultProfile = n;
  await writeProfilesFile(data);
  return data.profiles[n];
}

/**
 * Delete a profile's metadata record. Does not delete the underlying `token-cache-{identity}.json`
 * unless `purgeCache` is true (kept opt-in — the cache slot may still be in use directly via
 * `--identity`, and deleting cached auth material is a one-way, hard-to-reverse action).
 */
export async function deleteProfile(name: string, options?: { purgeCache?: boolean }): Promise<boolean> {
  const n = assertValidProfileName(name);
  const data = await readProfilesFile();
  const entry = data.profiles[n];
  if (!entry) return false;
  delete data.profiles[n];
  if (data.defaultProfile === n) {
    data.defaultProfile = undefined;
  }
  await writeProfilesFile(data);
  if (options?.purgeCache) {
    const identity = assertValidCacheIdentity(entry.identity);
    await unlink(join(getM365AgentCliConfigDir(), `token-cache-${identity}.json`)).catch(() => {});
  }
  return true;
}

export type CacheHealth = 'healthy' | 'expired' | 'missing' | 'malformed';

/** Offline (no network) cache-health probe for a profile's bound identity — reads the local token cache only. */
export async function probeCacheHealth(identity: string): Promise<CacheHealth> {
  try {
    const cache = await loadM365TokenCache(identity);
    if (!cache) return 'missing';
    const slot = cache.graph ?? cache.ews;
    if (!slot) return 'missing';
    return slot.expiresAt > Date.now() ? 'healthy' : 'expired';
  } catch {
    return 'malformed';
  }
}
