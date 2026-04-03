import { existsSync, readFileSync } from 'node:fs';
import { homedir } from 'node:os';
import { join, resolve } from 'node:path';

/**
 * Resolve a path with tilde expansion support.
 * @param raw - The raw path string to resolve
 * @param fallback - The fallback path if raw is empty
 * @returns The resolved absolute path
 */
function resolveTildePath(raw: string, fallback: string): string {
  const s = raw.trim();
  if (!s) {
    return fallback;
  }
  if (s === '~') {
    return homedir();
  }
  if (s.startsWith('~/') || s.startsWith('~\\')) {
    return join(homedir(), s.slice(2));
  }
  return resolve(s);
}

/**
 * Path to the CLI env file. Override with `M365_AGENT_ENV_FILE` (e.g. `~/.config/m365-agent-cli/.env.beta`).
 * Must be set in the shell before starting the process (not from inside `.env`).
 */
export function getGlobalEnvFilePath(): string {
  const raw = process.env.M365_AGENT_ENV_FILE?.trim() ?? '';
  return resolveTildePath(raw, join(homedir(), '.config', 'm365-agent-cli', '.env'));
}

/**
 * Resolve a user-supplied path (e.g. `~/.config/m365-agent-cli/.env.beta`) for `--env-file`.
 */
export function resolveEnvFilePathArgument(raw: string): string {
  return resolveTildePath(raw, join(homedir(), '.config', 'm365-agent-cli', '.env'));
}

/**
 * Parse a `.env` file and set `process.env`.
 * @param envPath Path to the .env file
 * @param overwrite If true, overwrites existing keys; if false, only sets undefined keys
 */
function parseEnvFile(envPath: string, overwrite: boolean): void {
  if (!existsSync(envPath)) {
    return;
  }
  const content = readFileSync(envPath, 'utf8');
  for (const line of content.split(/\r?\n/)) {
    const match = line.match(/^\s*([^#\s=]+)\s*=\s*(.*)$/);
    if (match) {
      const key = match[1];
      let val = match[2].trim();
      if ((val.startsWith('"') && val.endsWith('"')) || (val.startsWith("'") && val.endsWith("'"))) {
        val = val.slice(1, -1);
      }
      if (overwrite) {
        process.env[key] = val;
      } else if (process.env[key] === undefined) {
        process.env[key] = val;
      }
    }
  }
}

/**
 * Parse a `.env` file and set `process.env` (overwrites existing keys).
 * Use with `login --env-file` / `verify-token --env-file` so the beta app id and tokens apply
 * even when `M365_AGENT_ENV_FILE` was not exported before starting the process.
 */
export function applyEnvFileOverrides(envPath: string): void {
  parseEnvFile(envPath, true);
}

export function loadGlobalEnv() {
  const globalEnvPath = getGlobalEnvFilePath();
  parseEnvFile(globalEnvPath, false);
}

export function checkReadOnly(cmdOrOptions?: any) {
  let isReadOnly = process.env.READ_ONLY_MODE === 'true';

  if (cmdOrOptions) {
    // If it's a Commander Command instance
    if (typeof cmdOrOptions.optsWithGlobals === 'function') {
      if (cmdOrOptions.optsWithGlobals().readOnly) {
        isReadOnly = true;
      }
    }
    // If it's just an options object
    else if (cmdOrOptions.readOnly) {
      isReadOnly = true;
    }
  }

  if (isReadOnly) {
    console.error('Error: Command blocked. The CLI is running in read-only mode.');
    process.exit(1);
  }
}
