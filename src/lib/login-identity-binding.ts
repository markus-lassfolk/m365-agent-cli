/**
 * Shared "bind a completed login to a named identity profile" step for both the device-code and
 * browser (`--browser`) login flows (issues #244/#245).
 *
 * This codebase's `--identity` flag is a local cache-slot slug (letters/digits/underscore/hyphen —
 * see `assertValidCacheIdentity`), not necessarily the Microsoft UPN itself. Once a slug has been
 * verified against a real signed-in account, re-running `login --identity <same slug>` and landing
 * on a *different* account is exactly the "wrong-account" failure mode issue #245 is about — so a
 * mismatch refuses to complete instead of silently overwriting which account that slug means.
 */
import { getProfile, upsertProfile } from './identity-profiles.js';
import { assertValidCacheIdentity } from './m365-token-cache.js';

export class LoginAccountMismatchError extends Error {}

/**
 * Validate a completed login against a previously-verified identity slug WITHOUT writing
 * anything — throws {@link LoginAccountMismatchError} on a mismatch. Callers should call this
 * BEFORE persisting any other login state (env file, refresh token), then call
 * {@link commitLoginIdentity} only AFTER that state is safely persisted. That ordering keeps two
 * separate guarantees: a mismatch refuses to complete with NO local state changed at all, and an
 * unrelated failure while persisting the refresh token (disk full, read-only path) never leaves
 * `profiles.json` falsely claiming a fresh, verified login when no usable token was actually saved.
 */
export async function assertLoginIdentityOrThrow(options: {
  identity?: string;
  signedInAs?: string;
  force?: boolean;
}): Promise<void> {
  if (!options.identity || !options.signedInAs) return;
  const slug = assertValidCacheIdentity(options.identity);
  const existing = await getProfile(slug);
  if (
    existing?.signedInAs &&
    existing.signedInAs.toLowerCase() !== options.signedInAs.toLowerCase() &&
    !options.force
  ) {
    throw new LoginAccountMismatchError(
      `Refusing to complete: identity "${slug}" was previously verified as "${existing.signedInAs}", but this login resolved to "${options.signedInAs}". Re-run with --force-identity-switch to intentionally rebind this identity slug, or choose a different --identity.`
    );
  }
}

/**
 * Commit a completed, already-{@link assertLoginIdentityOrThrow}-validated login to the named
 * identity profile. Call only after all other login state (env file / refresh token) has been
 * durably persisted — see {@link assertLoginIdentityOrThrow}'s doc for why the ordering matters.
 */
export async function commitLoginIdentity(options: { identity?: string; signedInAs?: string }): Promise<void> {
  if (!options.identity) return;
  const slug = assertValidCacheIdentity(options.identity);
  // signedInAs may be undefined (couldn't decode a UPN from the token) — still register the slug
  // so `profiles list` reflects it; upsertProfile keeps the prior `signedInAs` when a new one
  // isn't given, so this never erases a previously-verified account.
  await upsertProfile(slug, { signedInAs: options.signedInAs, lastVerifiedAt: new Date().toISOString() });
}

/** @deprecated Use {@link assertLoginIdentityOrThrow} + {@link commitLoginIdentity} instead, with
 *  the commit happening after other login state is durably persisted. Kept for any external/test
 *  caller that wants the old atomic-looking (check-then-write) behavior in one call. */
export async function bindLoginIdentityOrThrow(options: {
  identity?: string;
  signedInAs?: string;
  force?: boolean;
}): Promise<void> {
  await assertLoginIdentityOrThrow(options);
  await commitLoginIdentity(options);
}
