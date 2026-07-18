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

export async function bindLoginIdentityOrThrow(options: {
  identity?: string;
  signedInAs?: string;
  force?: boolean;
}): Promise<void> {
  if (!options.identity) return;
  const slug = assertValidCacheIdentity(options.identity);

  if (!options.signedInAs) {
    // Couldn't decode a UPN from the token — still register the slug so `profiles list` reflects
    // it, but there is nothing to compare, so there is no mismatch to guard against.
    await upsertProfile(slug, { lastVerifiedAt: new Date().toISOString() });
    return;
  }

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

  await upsertProfile(slug, { signedInAs: options.signedInAs, lastVerifiedAt: new Date().toISOString() });
}
