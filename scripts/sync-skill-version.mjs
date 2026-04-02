#!/usr/bin/env node
/**
 * Sets `version:` in skills/m365-agent-cli/SKILL.md frontmatter from package.json.
 * Run when bumping the package version so the published skill metadata matches the CLI release.
 *
 * OpenClaw/ClawHub: `requires.bins` is only binary names (no semver). The top-level `version`
 * field is the skill bundle / aligned CLI release version.
 */
import { readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const root = join(__dirname, '..');
const pkg = JSON.parse(readFileSync(join(root, 'package.json'), 'utf8'));
const ver = pkg.version;
if (!ver || typeof ver !== 'string') {
  console.error('sync-skill-version: package.json missing version');
  process.exit(1);
}

const skillPath = join(root, 'skills', 'm365-agent-cli', 'SKILL.md');
let s = readFileSync(skillPath, 'utf8');

if (!s.startsWith('---')) {
  console.error('sync-skill-version: SKILL.md must start with YAML frontmatter ---');
  process.exit(1);
}
const m = s.match(/^---\r?\n([\s\S]*?)\r?\n---\r?\n/);
if (!m) {
  console.error('sync-skill-version: could not parse YAML frontmatter in SKILL.md');
  process.exit(1);
}

const front = m[0];
const rest = s.slice(front.length);

let newFront;
if (/^version:\s/m.test(front)) {
  newFront = front.replace(/^version:\s*[^\n]+/m, `version: ${ver}`);
} else {
  newFront = front.replace(/^(name:\s[^\n]+\n)/m, `$1version: ${ver}\n`);
}

writeFileSync(skillPath, newFront + rest, 'utf8');
console.log('sync-skill-version: set skills/m365-agent-cli/SKILL.md version to', ver);
