import { existsSync, readFileSync } from 'node:fs';
import { homedir } from 'node:os';
import { join } from 'node:path';

export function loadGlobalEnv() {
  const globalEnvPath = join(homedir(), '.config', 'm365-agent-cli', '.env');
  if (existsSync(globalEnvPath)) {
    const content = readFileSync(globalEnvPath, 'utf8');
    for (const line of content.split('\n')) {
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
