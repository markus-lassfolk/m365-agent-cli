import { existsSync, readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const PKG_NAME = 'm365-agent-cli';

function isOurPackageJson(path: string): boolean {
  try {
    const raw = readFileSync(path, 'utf8');
    const j = JSON.parse(raw) as { name?: string };
    return j.name === PKG_NAME;
  } catch {
    return false;
  }
}

function walkUpForOurPackageJson(startDir: string): string | null {
  let dir = startDir;
  for (let i = 0; i < 14; i++) {
    const candidate = join(dir, 'package.json');
    if (existsSync(candidate) && isOurPackageJson(candidate)) {
      return candidate;
    }
    const parent = dirname(dir);
    if (parent === dir) break;
    dir = parent;
  }
  return null;
}

/**
 * Resolves package.json for this package (installed layout, repo checkout, or tests).
 */
export function getPackageJsonPath(): string {
  const libDir = dirname(fileURLToPath(import.meta.url));
  const fromAdjacent = join(libDir, '../../package.json');
  if (existsSync(fromAdjacent) && isOurPackageJson(fromAdjacent)) {
    return fromAdjacent;
  }
  const fromModuleWalk = walkUpForOurPackageJson(libDir);
  if (fromModuleWalk) {
    return fromModuleWalk;
  }
  const fromCwdWalk = walkUpForOurPackageJson(process.cwd());
  if (fromCwdWalk) {
    return fromCwdWalk;
  }
  return fromAdjacent;
}

/** Uses sync fs so tests that mock `node:fs/promises` (e.g. auth) do not break version reads. */
export async function getPackageVersion(): Promise<string> {
  return getPackageVersionSync();
}

export function getPackageVersionSync(): string {
  const raw = readFileSync(getPackageJsonPath(), 'utf8');
  return (JSON.parse(raw) as { version?: string }).version?.trim() ?? '0.0.0';
}
