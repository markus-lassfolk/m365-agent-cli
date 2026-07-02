import { readFile } from 'node:fs/promises';
import { atomicWriteUtf8File } from './atomic-write.js';
import { getGlobalEnvFilePath } from './utils.js';

const REFRESH_TOKEN_KEYS = ['M365_REFRESH_TOKEN', 'EWS_REFRESH_TOKEN', 'GRAPH_REFRESH_TOKEN'] as const;

function shouldSkipEnvPersist(): boolean {
  return process.env.NODE_ENV === 'test' || process.env.M365_AGENT_SKIP_GLOBAL_ENV === '1';
}

function upsertEnvLine(content: string, key: string, value: string): string {
  const escaped = key.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const re = new RegExp(`^${escaped}=.*$`, 'm');
  if (re.test(content)) {
    return content.replace(re, () => `${key}=${value}`);
  }
  return `${content.trimEnd()}\n${key}=${value}\n`;
}

/**
 * Upsert unified refresh token fields in the global `.env` (same keys as `login`).
 * Skips when `NODE_ENV=test` or `M365_AGENT_SKIP_GLOBAL_ENV=1`.
 * Returns true when the file was written.
 */
export async function persistRefreshTokenToEnv(
  refreshToken: string,
  options?: { envPath?: string; previousRefreshToken?: string }
): Promise<boolean> {
  if (shouldSkipEnvPersist()) {
    return false;
  }

  const sanitized = refreshToken.replace(/[\r\n]/g, '');
  if (!sanitized) {
    return false;
  }

  const previous = options?.previousRefreshToken?.replace(/[\r\n]/g, '');
  if (previous && previous === sanitized) {
    return false;
  }

  const envPath = options?.envPath?.trim() || getGlobalEnvFilePath();
  if (!envPath) {
    // Defensive: getGlobalEnvFilePath() always returns a path, so this branch is unreachable
    // in normal flows, but guard against future regressions that pass an explicitly empty path.
    return false;
  }

  let envContent = '';
  try {
    envContent = await readFile(envPath, 'utf8');
  } catch (err: unknown) {
    if (err && typeof err === 'object' && 'code' in err && (err as NodeJS.ErrnoException).code !== 'ENOENT') {
      throw err;
    }
  }

  for (const key of REFRESH_TOKEN_KEYS) {
    envContent = upsertEnvLine(envContent, key, sanitized);
  }
  envContent = envContent.replace(/\n{3,}/g, '\n\n');
  await atomicWriteUtf8File(envPath, `${envContent.trim()}\n`, 0o600);

  process.env.M365_REFRESH_TOKEN = sanitized;
  process.env.EWS_REFRESH_TOKEN = sanitized;
  process.env.GRAPH_REFRESH_TOKEN = sanitized;

  return true;
}
