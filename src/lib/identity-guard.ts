/**
 * Global `--require-identity` / `--as-delegate-of` wrong-account guardrails (see docs/AUTHENTICATION.md).
 *
 * Wired into `m365-program.ts`'s root `preAction` hook so it applies uniformly to every command
 * without per-command changes — the same pattern `--dry-run`/`--cache` already use.
 */
import { resolveAuth } from './auth.js';
import { getExchangeBackend } from './exchange-backend.js';
import { resolveGraphAuth } from './graph-auth.js';
import { resolveIdentitySlug } from './identity-profiles.js';
import { getJwtPayloadUpn } from './jwt-utils.js';

export interface IdentityGuardOptions {
  identity?: string;
  requireIdentity?: string;
  asDelegateOf?: string;
  mailbox?: string;
}

export interface IdentityGuardResult {
  ok: boolean;
  message?: string;
  signedInAs?: string;
}

/** Case-insensitive UPN equality — the one comparison every wrong-account guardrail in this
 *  module (and `readiness --expect-identity`) should use, so a future hardening (e.g. normalizing
 *  UPN casing/domain-alias differences) only needs to change in one place. */
export function upnsMatch(a: string | undefined, b: string | undefined): boolean {
  return Boolean(a && b && a.toLowerCase() === b.toLowerCase());
}

/**
 * Resolve the signed-in UPN for a cache identity slot without assuming Graph or EWS backend —
 * tries Graph first (richer claim set on most tenants), falls back to EWS (same underlying login,
 * different resource audience) unless `M365_EXCHANGE_BACKEND=graph` pins Graph-only (matching
 * `auth-diagnostics.ts`'s `diagnoseAuth`, so this guard and `readiness`/`auth repair` never
 * disagree about which backend a "healthy" identity resolved through). Read-only: never mutates
 * auth state beyond the refresh either path would already perform.
 */
export async function resolveSignedInUpn(identity: string): Promise<string | undefined> {
  const graph = await resolveGraphAuth({ identity });
  if (graph.success && graph.token) {
    const upn = getJwtPayloadUpn(graph.token);
    if (upn) return upn;
  }
  if (getExchangeBackend() === 'graph') return undefined;
  const ews = await resolveAuth({ identity });
  if (ews.success && ews.token) {
    return getJwtPayloadUpn(ews.token);
  }
  return undefined;
}

/**
 * Enforce `--require-identity` / `--as-delegate-of` before a command's action runs.
 *
 * Fails closed whenever the signed-in identity cannot be verified at all — an unauthenticatable
 * or undecodable identity is treated as a mismatch, not a pass-through. This mirrors issue #245's
 * premise: for assistant-driven workflows, an unverifiable identity is more dangerous than no
 * identity, since a silent pass-through could let a command run against the wrong account.
 */
export async function checkIdentityGuards(opts: IdentityGuardOptions): Promise<IdentityGuardResult> {
  if (!opts.requireIdentity && !opts.asDelegateOf) {
    return { ok: true };
  }

  if (opts.asDelegateOf && !opts.mailbox) {
    return {
      ok: false,
      message:
        '--as-delegate-of requires --mailbox on this command to name which mailbox you are operating on (distinguishes "signed in as X" from "operating on mailbox Y via delegation").'
    };
  }

  const identity = await resolveIdentitySlug(opts.identity);
  const signedInAs = await resolveSignedInUpn(identity);

  if (!signedInAs) {
    return {
      ok: false,
      message:
        'Could not verify the signed-in identity (no valid cached/refreshable token for this identity) — refusing to run with --require-identity/--as-delegate-of set. Run `m365-agent-cli login` or `m365-agent-cli whoami` first.'
    };
  }

  // Checked independently (not `requireIdentity ?? asDelegateOf`, which would only ever verify
  // whichever flag came first) so passing both flags together can't let one silently short-circuit
  // verification of the other.
  const checks: Array<{ flag: '--require-identity' | '--as-delegate-of'; wanted: string }> = [];
  if (opts.requireIdentity) checks.push({ flag: '--require-identity', wanted: opts.requireIdentity });
  if (opts.asDelegateOf) checks.push({ flag: '--as-delegate-of', wanted: opts.asDelegateOf });

  for (const check of checks) {
    if (!upnsMatch(signedInAs, check.wanted)) {
      return {
        ok: false,
        message: `Identity guard failed: signed in as "${signedInAs}", but ${check.flag} requires "${check.wanted}".`,
        signedInAs
      };
    }
  }

  return { ok: true, signedInAs };
}
