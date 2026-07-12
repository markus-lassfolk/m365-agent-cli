import { spawn } from 'node:child_process';
import { realpathSync } from 'node:fs';
import { sep } from 'node:path';
import { Command } from 'commander';
import semver from 'semver';
import { getPackageVersion } from '../lib/package-info.js';
import { checkReadOnly } from '../lib/utils.js';

const NPM_PKG = 'm365-agent-cli';
const NPM_LATEST_URL = `https://registry.npmjs.org/${NPM_PKG}/latest`;

const NODE_MODULES_PKG = `${sep}node_modules${sep}m365-agent-cli${sep}`;

/**
 * When the CLI is executed by Bun but lives under npm/pnpm/yarn `node_modules/m365-agent-cli`
 * (not under Bun's `~/.bun` install), `bun install -g` only updates Bun's global bin — the copy
 * your shell usually runs stays old. Use npm for the global install in that case.
 */
function isInstalledFromNpmStyleGlobalLayout(): boolean {
  try {
    const exe = realpathSync(process.argv[1] || '');
    return exe.includes(NODE_MODULES_PKG) && !exe.includes(`${sep}.bun${sep}`);
  } catch {
    return false;
  }
}

function compareVersions(a: string, b: string): number {
  const va = semver.valid(semver.coerce(a));
  const vb = semver.valid(semver.coerce(b));
  if (va && vb) return semver.compare(va, vb);
  return a.localeCompare(b);
}

async function fetchLatestVersion(): Promise<string> {
  const ac = new AbortController();
  const t = setTimeout(() => ac.abort(), 15_000);
  try {
    const r = await fetch(NPM_LATEST_URL, { signal: ac.signal });
    if (!r.ok) {
      throw new Error(`npm registry returned ${r.status}`);
    }
    const j = (await r.json()) as { version?: string };
    if (!j.version?.trim()) {
      throw new Error('missing version in registry response');
    }
    return j.version.trim();
  } finally {
    clearTimeout(t);
  }
}

function runGlobalInstall(): Promise<{ code: number; packageManager: 'npm' | 'bun' }> {
  const pkg = `${NPM_PKG}@latest`;
  const bunRuntime = process.versions.bun !== undefined;
  const useBun = bunRuntime && !isInstalledFromNpmStyleGlobalLayout();
  const packageManager: 'npm' | 'bun' = useBun ? 'bun' : 'npm';
  const args = ['install', '-g', pkg];

  return new Promise((resolve, reject) => {
    const child = spawn(packageManager, args, {
      stdio: 'inherit',
      shell: process.platform === 'win32',
      env: process.env
    });
    child.on('error', reject);
    child.on('close', (code) => resolve({ code: code ?? 1, packageManager }));
  });
}

export const updateCommand = new Command('update')
  .description('Check for and install the latest npm release of this CLI')
  .option('-c, --check', 'Only check if a newer version exists on npm (exit 1 if newer is available)')
  .action(async (opts: { check?: boolean }, cmd) => {
    const current = await getPackageVersion();
    let latest: string;
    try {
      latest = await fetchLatestVersion();
    } catch (e) {
      console.error(`Error: ${e instanceof Error ? e.message : String(e)}`);
      process.exit(1);
    }

    const cmp = compareVersions(current, latest);
    if (cmp === 0) {
      console.log(`m365-agent-cli is up to date (${current}).`);
      process.exit(0);
    }
    if (cmp > 0) {
      console.log(`m365-agent-cli local version ${current} is newer than npm latest (${latest}).`);
      process.exit(0);
    }

    if (opts.check) {
      console.log(`Update available: ${current} → ${latest}`);
      process.exit(1);
    }

    checkReadOnly(cmd);

    console.log(`Updating ${current} → ${latest}…`);
    const { code, packageManager } = await runGlobalInstall();
    if (code !== 0) {
      console.error(`Error: ${packageManager} install -g exited with code ${code}`);
      console.error('Try manually: npm install -g m365-agent-cli@latest   or   bun install -g m365-agent-cli@latest');
      process.exit(1);
    }
    console.log(`Updated to ${latest} (${packageManager} install -g).`);
    if (packageManager === 'bun') {
      console.log(
        'If m365-agent-cli --version is still old, another copy is earlier on PATH. Check: type -a m365-agent-cli  (or put Bun’s global bin before npm).'
      );
    } else {
      console.log('Run m365-agent-cli --version to confirm (same shell is fine for npm global).');
    }
    process.exit(0);
  });
