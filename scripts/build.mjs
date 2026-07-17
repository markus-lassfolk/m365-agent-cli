#!/usr/bin/env node
/**
 * Compiles src/ to dist/ (tsconfig.build.json) for the published npm package: a Node-runnable
 * `dist/cli.js` with a `#!/usr/bin/env node` shebang, so `npm install -g m365-agent-cli` works
 * without a separately installed Bun runtime (see issue #239). Source keeps its `#!/usr/bin/env
 * bun` shebang for the bun-based dev workflow (`bun run src/cli.ts`); only the compiled output
 * is rewritten here.
 */
import { execFileSync } from 'node:child_process';
import { chmodSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const root = join(__dirname, '..');
const distDir = join(root, 'dist');
const cliEntry = join(distDir, 'cli.js');

rmSync(distDir, { recursive: true, force: true });

execFileSync('tsc', ['-p', 'tsconfig.build.json'], { cwd: root, stdio: 'inherit' });

const cli = readFileSync(cliEntry, 'utf8');
const lines = cli.split('\n');
const nodeShebang = '#!/usr/bin/env node';
if (lines[0].startsWith('#!')) {
  lines[0] = nodeShebang;
} else {
  lines.unshift(nodeShebang);
}
writeFileSync(cliEntry, lines.join('\n'), 'utf8');
chmodSync(cliEntry, 0o755);

console.log('build: wrote', cliEntry, 'with', nodeShebang);
