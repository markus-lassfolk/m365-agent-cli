import { getGlobalEnvFilePath, resolveEnvFilePathArgument } from './utils.js';

/**
 * Resolve the active env file path used for refresh-token persistence.
 *
 * Precedence:
 *   1. Caller-provided `envPath` (e.g. `--env-file` from `login` / `verify-token`).
 *   2. `M365_AGENT_ENV_FILE` shell override (e.g. `.env.beta`).
 *   3. Default global path (`~/.config/m365-agent-cli/.env`).
 *
 * Empty/whitespace strings fall through to the next candidate so a stale env var
 * does not silently route tokens to a non-existent file.
 *
 * Callers should pass the resolved path to `persistRefreshTokenToEnv({ envPath })`
 * to keep rotated refresh tokens landing in the same file the CLI loaded from.
 */
export function getActiveEnvFilePath(explicitEnvPath?: string): string {
  if (explicitEnvPath?.trim()) {
    return resolveEnvFilePathArgument(explicitEnvPath);
  }
  return getGlobalEnvFilePath();
}
