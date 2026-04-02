import { spawn } from 'node:child_process';
import { Command } from 'commander';
import semver from 'semver';
import { getPackageVersion } from '../lib/package-info.js';

const NPM_PKG = 'm365-agent-cli';
const NPM_LATEST_URL = `https://registry.npmjs.org/${NPM_PKG}/latest`;

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

function runGlobalInstall(): Promise<number> {
  const pkg = `${NPM_PKG}@latest`;
  const useBun = process.versions.bun !== undefined;
  const cmd = useBun ? 'bun' : 'npm';
  const args = ['install', '-g', pkg];

  return new Promise((resolve, reject) => {
    const child = spawn(cmd, args, {
      stdio: 'inherit',
      shell: process.platform === 'win32',
      env: process.env
    });
    child.on('error', reject);
    child.on('close', (code) => resolve(code ?? 1));
  });
}

export const updateCommand = new Command('update')
  .description('Check for and install the latest npm release of this CLI')
  .option('-c, --check', 'Only check if a newer version exists on npm (exit 1 if newer is available)')
  .action(async (opts: { check?: boolean }) => {
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

    console.log(`Updating ${current} → ${latest}…`);
    const code = await runGlobalInstall();
    if (code !== 0) {
      console.error(`Error: ${process.versions.bun ? 'bun' : 'npm'} install exited with code ${code}`);
      console.error('Try manually: npm install -g m365-agent-cli@latest   or   bun install -g m365-agent-cli@latest');
      process.exit(1);
    }
    console.log(`Updated to ${latest}. Run again to use the new version.`);
    process.exit(0);
  });
