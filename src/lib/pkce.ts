/**
 * RFC 7636 PKCE (Proof Key for Code Exchange) helpers for the browser authorization-code login
 * flow (`login --browser`, issue #244). `code_verifier`/`code_challenge` only — never persisted,
 * used once per login attempt and discarded.
 */
import { createHash, randomBytes } from 'node:crypto';

/** 43-128 char unreserved-charset string per RFC 7636 §4.1. 32 random bytes → 43 base64url chars. */
export function generateCodeVerifier(): string {
  return randomBytes(32).toString('base64url');
}

/** S256 code_challenge = BASE64URL(SHA256(code_verifier)). */
export function codeChallengeFromVerifier(codeVerifier: string): string {
  return createHash('sha256').update(codeVerifier).digest('base64url');
}

/** Opaque anti-CSRF `state` value for the authorization request / redirect round trip. */
export function generateOAuthState(): string {
  return randomBytes(16).toString('hex');
}

export interface PkcePair {
  codeVerifier: string;
  codeChallenge: string;
  codeChallengeMethod: 'S256';
}

export function generatePkcePair(): PkcePair {
  const codeVerifier = generateCodeVerifier();
  return { codeVerifier, codeChallenge: codeChallengeFromVerifier(codeVerifier), codeChallengeMethod: 'S256' };
}
