import { existsSync, readFileSync } from 'node:fs';
import { homedir } from 'node:os';
import { join, resolve } from 'node:path';

/**
 * Path to the CLI env file. Override with `M365_AGENT_ENV_FILE` (e.g. `~/.config/m365-agent-cli/.env.beta`).
 * Must be set in the shell before starting the process (not from inside `.env`).
 */
export function getGlobalEnvFilePath(): string {
  const raw = process.env.M365_AGENT_ENV_FILE?.trim();
  if (!raw) {
    return join(homedir(), '.config', 'm365-agent-cli', '.env');
  }
  if (raw === '~') {
    return homedir();
  }
  if (raw.startsWith('~/') || raw.startsWith('~\\')) {
    return join(homedir(), raw.slice(2));
  }
  return resolve(raw);
}

export function loadGlobalEnv() {
  const globalEnvPath = getGlobalEnvFilePath();
  if (existsSync(globalEnvPath)) {
    const content = readFileSync(globalEnvPath, 'utf8');
    for (const line of content.split(/\r?\n/)) {
      const match = line.match(/^\s*([^#\s=]+)\s*=\s*(.*)$/);
      if (match) {
        const key = match[1];
        let val = match[2].trim();
        if ((val.startsWith('"') && val.endsWith('"')) || (val.startsWith("'") && val.endsWith("'"))) {
          val = val.slice(1, -1);
        }
        if (process.env[key] === undefined) {
          process.env[key] = val;
        }
      }
    }
  }
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
