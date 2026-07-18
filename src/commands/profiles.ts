import { Command } from 'commander';
import {
  type CacheHealth,
  deleteProfile,
  getDefaultProfileName,
  getProfile,
  listProfiles,
  type ProfileEntry,
  probeCacheHealth,
  setDefaultProfile
} from '../lib/identity-profiles.js';
import { permissionSetFromGraphPayload } from '../lib/graph-capability-matrix.js';
import { toJsonError } from '../lib/json-error.js';
import { loadM365TokenCache } from '../lib/m365-token-cache.js';

interface ProfileSummary {
  name: string;
  identity: string;
  envFile?: string;
  tenantId?: string;
  signedInAs?: string;
  lastVerifiedAt?: string;
  cacheHealth: CacheHealth;
  capabilities: string[];
  isDefault: boolean;
}

/** Offline (no network) capability list decoded from whatever Graph access token is already cached. */
async function offlineCapabilities(identity: string): Promise<string[]> {
  const cache = await loadM365TokenCache(identity).catch(() => null);
  if (!cache?.graph?.accessToken) return [];
  try {
    const parts = cache.graph.accessToken.split('.');
    if (parts.length !== 3) return [];
    const payload = JSON.parse(Buffer.from(parts[1], 'base64url').toString('utf8'));
    return [...permissionSetFromGraphPayload(payload)].sort();
  } catch {
    return [];
  }
}

async function summarizeProfile(entry: ProfileEntry, defaultProfile: string | undefined): Promise<ProfileSummary> {
  const [cacheHealth, capabilities] = await Promise.all([
    probeCacheHealth(entry.identity),
    offlineCapabilities(entry.identity)
  ]);
  return {
    name: entry.name,
    identity: entry.identity,
    envFile: entry.envFile,
    tenantId: entry.tenantId,
    signedInAs: entry.signedInAs,
    lastVerifiedAt: entry.lastVerifiedAt,
    cacheHealth,
    capabilities,
    isDefault: entry.name === defaultProfile
  };
}

function printProfileText(p: ProfileSummary): void {
  console.log(`  ${p.name}${p.isDefault ? '  (default)' : ''}`);
  console.log(`    identity (cache slot): ${p.identity}`);
  if (p.envFile) console.log(`    env file:               ${p.envFile}`);
  if (p.tenantId) console.log(`    tenant:                 ${p.tenantId}`);
  console.log(`    signed in as:           ${p.signedInAs ?? '(unknown — never verified)'}`);
  console.log(`    cache health:           ${p.cacheHealth}`);
  console.log(`    last verified:          ${p.lastVerifiedAt ?? '(never)'}`);
  console.log(`    capabilities:           ${p.capabilities.length ? p.capabilities.join(', ') : '(none cached)'}`);
}

// ─── list ───

const listCmd = new Command('list')
  .description('List registered identity profiles')
  .option('--json', 'Output as JSON')
  .action(async (opts: { json?: boolean }) => {
    const [entries, defaultProfile] = await Promise.all([listProfiles(), getDefaultProfileName()]);
    const summaries = await Promise.all(entries.map((e) => summarizeProfile(e, defaultProfile)));

    if (opts.json) {
      console.log(JSON.stringify({ defaultProfile: defaultProfile ?? null, profiles: summaries }, null, 2));
      return;
    }

    if (summaries.length === 0) {
      console.log('No identity profiles registered yet.');
      console.log('Tip: `m365-agent-cli profiles set-default <name>` registers and selects a profile.');
      return;
    }

    console.log(`Identity profiles (${summaries.length}):\n`);
    for (const p of summaries) {
      printProfileText(p);
      console.log();
    }
  });

// ─── show ───

const showCmd = new Command('show')
  .description('Show one identity profile (default profile when no name is given)')
  .argument('[name]', 'Profile name (defaults to the current default profile)')
  .option('--json', 'Output as JSON')
  .action(async (name: string | undefined, opts: { json?: boolean }) => {
    const defaultProfile = await getDefaultProfileName();
    const resolvedName = name ?? defaultProfile;

    if (!resolvedName) {
      const message = 'No profile name given and no default profile is set.';
      if (opts.json) {
        console.log(JSON.stringify({ error: toJsonError(message) }, null, 2));
      } else {
        console.error(`Error: ${message}`);
        console.error('Tip: `m365-agent-cli profiles set-default <name>` first, or pass a profile name.');
      }
      process.exit(1);
    }

    const entry = await getProfile(resolvedName);
    if (!entry) {
      const message = `No such profile: ${resolvedName}`;
      if (opts.json) {
        console.log(JSON.stringify({ error: toJsonError(message) }, null, 2));
      } else {
        console.error(`Error: ${message}`);
      }
      process.exit(1);
    }

    const summary = await summarizeProfile(entry, defaultProfile);
    if (opts.json) {
      console.log(JSON.stringify(summary, null, 2));
      return;
    }
    printProfileText(summary);
  });

// ─── set-default ───

const setDefaultCmd = new Command('set-default')
  .description('Set the default identity profile (used whenever --identity is omitted)')
  .argument('<name>', 'Profile name (auto-registered if not already known)')
  .option('--json', 'Output as JSON')
  .action(async (name: string, opts: { json?: boolean }) => {
    try {
      const entry = await setDefaultProfile(name);
      if (opts.json) {
        console.log(JSON.stringify({ defaultProfile: entry.name, identity: entry.identity }, null, 2));
      } else {
        console.log(`✓ Default profile set to "${entry.name}" (cache identity: ${entry.identity}).`);
      }
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      if (opts.json) {
        console.log(JSON.stringify({ error: toJsonError(message) }, null, 2));
      } else {
        console.error(`Error: ${message}`);
      }
      process.exit(1);
    }
  });

// ─── delete ───

const deleteCmd = new Command('delete')
  .description('Delete a profile record (metadata only; does not delete the underlying token cache)')
  .argument('<name>', 'Profile name')
  .option('--purge-cache', 'Also delete the underlying token-cache file for this profile’s identity')
  .option('--json', 'Output as JSON')
  .action(async (name: string, opts: { purgeCache?: boolean; json?: boolean }) => {
    const removed = await deleteProfile(name, { purgeCache: opts.purgeCache });
    if (!removed) {
      const message = `No such profile: ${name}`;
      if (opts.json) {
        console.log(JSON.stringify({ error: toJsonError(message) }, null, 2));
      } else {
        console.error(`Error: ${message}`);
      }
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify({ deleted: name, purgedCache: Boolean(opts.purgeCache) }, null, 2));
    } else {
      console.log(`✓ Deleted profile "${name}"${opts.purgeCache ? ' (and its token cache)' : ''}.`);
    }
  });

export const profilesCommand = new Command('profiles')
  .description('Manage named identity profiles (default selection, wrong-account guardrails)')
  .addCommand(listCmd)
  .addCommand(showCmd)
  .addCommand(setDefaultCmd)
  .addCommand(deleteCmd);
