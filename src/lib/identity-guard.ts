/**
 * Global `--require-identity` / `--as-delegate-of` wrong-account guardrails (see docs/AUTHENTICATION.md).
 *
 * Wired into `m365-program.ts`'s root `preAction` hook so it applies uniformly to every command
 * without per-command changes — the same pattern `--dry-run`/`--cache` already use.
 */
import { resolveAuth } from './auth.js';
import { resolveGraphAuth } from './graph-auth.js';
import { getDefaultProfileIdentity } from './identity-profiles.js';
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

/**
 * Resolve the signed-in UPN for a cache identity slot without assuming Graph or EWS backend —
 * tries Graph first (richer claim set on most tenants), falls back to EWS (same underlying login,
 * different resource audience). Read-only: never mutates auth state beyond the refresh either
 * path would already perform.
 */
export async function resolveSignedInUpn(identity: string): Promise<string | undefined> {
  const graph = await resolveGraphAuth({ identity });
  if (graph.success && graph.token) {
    const upn = getJwtPayloadUpn(graph.token);
    if (upn) return upn;
  }
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

  const identity = opts.identity || (await getDefaultProfileIdentity()) || 'default';
  const signedInAs = await resolveSignedInUpn(identity);

  if (!signedInAs) {
    return {
      ok: false,
      message:
        'Could not verify the signed-in identity (no valid cached/refreshable token for this identity) — refusing to run with --require-identity/--as-delegate-of set. Run `m365-agent-cli login` or `m365-agent-cli whoami` first.'
    };
  }

  const wanted = opts.requireIdentity ?? opts.asDelegateOf;
  if (wanted && signedInAs.toLowerCase() !== wanted.toLowerCase()) {
    const flag = opts.requireIdentity ? '--require-identity' : '--as-delegate-of';
    return {
      ok: false,
      message: `Identity guard failed: signed in as "${signedInAs}", but ${flag} requires "${wanted}".`,
      signedInAs
    };
  }

  return { ok: true, signedInAs };
}
