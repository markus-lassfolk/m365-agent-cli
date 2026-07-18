import { Command } from 'commander';
import { capabilitiesFromToken } from '../lib/auth-diagnostics.js';
import {
  assertValidProfileName,
  type CacheHealth,
  deleteProfile,
  getProfilesSnapshot,
  type ProfileEntry,
  probeCacheHealth,
  setDefaultProfile
} from '../lib/identity-profiles.js';
import { toJsonError } from '../lib/json-error.js';
import { loadM365TokenCache } from '../lib/m365-token-cache.js';
import { deepRedact, IDENTITY_LABEL_SAFE_KEYS } from '../lib/redact.js';

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
  return capabilitiesFromToken(cache?.graph?.accessToken);
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

/** Print a `{error}` JSON envelope (or plain-text error, plus any extra tip lines) and exit 1 —
 *  single source of truth for every error path in this file so every `--json` caller gets the
 *  same shape and every text-mode caller gets the same "Error: " prefix. */
function failWith(json: boolean | undefined, message: string, extraTextLines: string[] = []): never {
  if (json) {
    console.log(JSON.stringify({ error: toJsonError(message) }, null, 2));
  } else {
    console.error(`Error: ${message}`);
    for (const line of extraTextLines) console.error(line);
  }
  process.exit(1);
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
    const { profiles: entries, defaultProfile } = await getProfilesSnapshot();
    const summaries = await Promise.all(entries.map((e) => summarizeProfile(e, defaultProfile)));

    if (opts.json) {
      console.log(
        JSON.stringify(
          deepRedact(
            { defaultProfile: defaultProfile ?? null, profiles: summaries },
            { safeKeys: IDENTITY_LABEL_SAFE_KEYS }
          ),
          null,
          2
        )
      );
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
    const { profiles: entries, defaultProfile } = await getProfilesSnapshot();

    // Narrow try/catch around just the validating call — NOT around the failWith calls below,
    // which themselves call `process.exit` (mocked to throw in tests): wrapping them too would
    // catch that throw here and re-report it as a generic "process.exit(1)" error instead of the
    // real message.
    let resolvedName: string | undefined;
    try {
      resolvedName = name ? assertValidProfileName(name) : defaultProfile;
    } catch (err) {
      failWith(opts.json, err instanceof Error ? err.message : String(err));
    }

    if (!resolvedName) {
      failWith(opts.json, 'No profile name given and no default profile is set.', [
        'Tip: `m365-agent-cli profiles set-default <name>` first, or pass a profile name.'
      ]);
    }

    const entry = entries.find((e) => e.name === resolvedName);
    if (!entry) {
      failWith(opts.json, `No such profile: ${resolvedName}`);
    }

    const summary = await summarizeProfile(entry, defaultProfile);
    if (opts.json) {
      console.log(JSON.stringify(deepRedact(summary, { safeKeys: IDENTITY_LABEL_SAFE_KEYS }), null, 2));
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
        console.log(
          JSON.stringify(
            deepRedact(
              { defaultProfile: entry.name, identity: entry.identity },
              { safeKeys: IDENTITY_LABEL_SAFE_KEYS }
            ),
            null,
            2
          )
        );
      } else {
        console.log(`✓ Default profile set to "${entry.name}" (cache identity: ${entry.identity}).`);
      }
    } catch (err) {
      failWith(opts.json, err instanceof Error ? err.message : String(err));
    }
  });

// ─── delete ───

const deleteCmd = new Command('delete')
  .description('Delete a profile record (metadata only; does not delete the underlying token cache)')
  .argument('<name>', 'Profile name')
  .option('--purge-cache', 'Also delete the underlying token-cache file for this profile’s identity')
  .option('--json', 'Output as JSON')
  .action(async (name: string, opts: { purgeCache?: boolean; json?: boolean }) => {
    // `removed` is resolved inside try/catch (deleteProfile validates the name and can throw);
    // the not-found failWith below stays OUTSIDE the try so its `process.exit` (mocked to throw in
    // tests) isn't re-caught here and re-reported as a generic "process.exit(1)" error.
    let removed: boolean;
    try {
      removed = await deleteProfile(name, { purgeCache: opts.purgeCache });
    } catch (err) {
      failWith(opts.json, err instanceof Error ? err.message : String(err));
    }

    if (!removed) {
      failWith(opts.json, `No such profile: ${name}`);
    }

    if (opts.json) {
      console.log(
        JSON.stringify(
          deepRedact({ deleted: name, purgedCache: Boolean(opts.purgeCache) }, { safeKeys: IDENTITY_LABEL_SAFE_KEYS }),
          null,
          2
        )
      );
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
