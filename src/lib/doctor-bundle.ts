/**
 * Builds the non-secret diagnostic bundle for `doctor` / `doctor --redacted-bundle` (issue #246).
 *
 * Every field here is metadata *about* auth state (presence, size, mtime, classification) — never
 * the state itself. The result still passes through {@link deepRedact} before being returned as a
 * final defense-in-depth pass (see `redact.ts`'s module doc).
 */
import { stat } from 'node:fs/promises';
import { arch, platform, release } from 'node:os';
import { checkMailboxAccess, diagnoseAuth } from './auth-diagnostics.js';
import { getExchangeBackend } from './exchange-backend.js';
import { resolveGraphAuth } from './graph-auth.js';
import { getProfilesSnapshot } from './identity-profiles.js';
import { getM365AgentCliConfigDir, tokenCachePath } from './m365-token-cache.js';
import { getPackageVersionSync } from './package-info.js';
import { deepRedact, IDENTITY_LABEL_SAFE_KEYS } from './redact.js';
import { getGlobalEnvFilePath } from './utils.js';

/** Bump when the bundle shape changes in a way tooling should branch on. */
export const DOCTOR_BUNDLE_SCHEMA_VERSION = 1;

export interface FilePresenceInfo {
  path: string;
  exists: boolean;
  sizeBytes: number | null;
  mtime: string | null;
}

export interface DoctorBundle {
  schemaVersion: number;
  generatedAt: string;
  cli: {
    version: string;
    nodeVersion: string;
    platform: string;
    arch: string;
    osRelease: string;
  };
  config: {
    configDir: string;
    envFile: FilePresenceInfo;
  };
  exchangeBackend: string;
  /** Entra application (client) id — a public OAuth client identifier, not a secret; same value
   *  `login`/`verify-token` already print unredacted. */
  clientId: string | null;
  profiles: {
    defaultProfile: string | null;
    names: string[];
  };
  identity: {
    name: string;
    cacheFile: FilePresenceInfo;
  };
  authDiagnosis: {
    status: string;
    failureClass: string;
    evidence: string[];
    recommendedAction: string | null;
    safeCommand: string | null;
    cacheHealth: string;
    tenantId: string | null;
    tokenExpiryKnown: boolean;
    capabilitiesCount: number;
  };
  mailboxCheck: {
    checked: boolean;
    mailbox: string | null;
    ok: boolean | null;
  } | null;
  secretsPrinted: false;
  unsafeFieldsIncluded: false;
}

async function inspectFile(path: string): Promise<FilePresenceInfo> {
  try {
    const s = await stat(path);
    return { path, exists: true, sizeBytes: s.size, mtime: s.mtime.toISOString() };
  } catch {
    return { path, exists: false, sizeBytes: null, mtime: null };
  }
}

export interface BuildDoctorBundleOptions {
  identity: string;
  mailbox?: string;
  envPath?: string;
  /** Explicit opt-in required for anything beyond the always-safe default fields. Currently a
   *  no-op placeholder — no unsafe field exists yet — but keeps the "refuses unless explicit
   *  unsafe flag" contract enforceable if one is ever added. */
  allowUnsafeFields?: boolean;
}

async function computeMailboxCheck(
  diag: Awaited<ReturnType<typeof diagnoseAuth>>,
  options: BuildDoctorBundleOptions
): Promise<DoctorBundle['mailboxCheck']> {
  if (!options.mailbox) return null;
  if (diag.status !== 'healthy') {
    return { checked: false, mailbox: options.mailbox, ok: null };
  }
  if (diag.authBackend !== 'graph') {
    // Mailbox delegation is only checkable via Graph today (`checkMailboxAccess` calls the Graph
    // API) — an EWS-only-healthy identity is genuinely healthy, but this specific check can't run
    // for it. Report unchecked rather than a misleading "no token" failure.
    return { checked: false, mailbox: options.mailbox, ok: null };
  }
  const graphAuth = await resolveGraphAuth({ identity: diag.identity, envPath: options.envPath });
  const access =
    graphAuth.success && graphAuth.token
      ? await checkMailboxAccess(graphAuth.token, options.mailbox)
      : { checked: true, ok: false };
  return { checked: true, mailbox: options.mailbox, ok: access.ok ?? false };
}

export async function buildDoctorBundle(options: BuildDoctorBundleOptions): Promise<DoctorBundle> {
  const { identity } = options;
  const diag = await diagnoseAuth({ identity, envPath: options.envPath });

  const [mailboxCheck, envFileInfo, cacheFileInfo, profilesSnapshot] = await Promise.all([
    computeMailboxCheck(diag, options),
    inspectFile(options.envPath ?? getGlobalEnvFilePath()),
    inspectFile(tokenCachePath(identity)),
    getProfilesSnapshot()
  ]);
  const { profiles, defaultProfile } = profilesSnapshot;

  const bundle: DoctorBundle = {
    schemaVersion: DOCTOR_BUNDLE_SCHEMA_VERSION,
    generatedAt: new Date().toISOString(),
    cli: {
      version: getPackageVersionSync(),
      nodeVersion: process.version,
      platform: platform(),
      arch: arch(),
      osRelease: release()
    },
    config: {
      configDir: getM365AgentCliConfigDir(),
      envFile: envFileInfo
    },
    exchangeBackend: getExchangeBackend(),
    clientId: process.env.EWS_CLIENT_ID?.trim() || null,
    profiles: {
      defaultProfile: defaultProfile ?? null,
      names: profiles.map((p) => p.name)
    },
    identity: {
      name: identity,
      cacheFile: cacheFileInfo
    },
    authDiagnosis: {
      status: diag.status,
      failureClass: diag.failureClass,
      evidence: diag.evidence,
      recommendedAction: diag.status === 'healthy' ? null : diag.recommendedAction,
      safeCommand: diag.status === 'healthy' ? null : (diag.safeCommand ?? null),
      cacheHealth: diag.cacheHealth,
      tenantId: diag.tenantId ?? null,
      tokenExpiryKnown: cacheFileInfo.exists,
      capabilitiesCount: diag.capabilities.length
    },
    mailboxCheck,
    secretsPrinted: false,
    unsafeFieldsIncluded: false
  };

  return deepRedact(bundle, { safeKeys: IDENTITY_LABEL_SAFE_KEYS });
}
