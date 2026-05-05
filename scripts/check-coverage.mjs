#!/usr/bin/env node
/**
 * Enforce a minimum line coverage from Bun's lcov output (coverage/lcov.info).
 * Set COVERAGE_MIN_LINES (default 31) to adjust the bar.
 *
 * Optional COVERAGE_EXCLUDE_PREFIXES: comma-separated path prefixes (forward slashes, relative to repo root).
 * Example: `src/commands/` omits command entrypoints so the gate tracks **library + test** code (large command files otherwise dominate the denominator).
 */
import { readFileSync } from 'node:fs';

const min = Number(process.env.COVERAGE_MIN_LINES ?? '31');
const lcovPath = process.argv[2] ?? 'coverage/lcov.info';

const excludePrefixes = (process.env.COVERAGE_EXCLUDE_PREFIXES ?? '')
  .split(',')
  .map((s) => s.trim().replace(/\\/g, '/'))
  .filter(Boolean);

function normalizeSf(sf) {
  return sf.trim().replace(/\\/g, '/');
}

function isExcluded(sfNorm) {
  for (const p of excludePrefixes) {
    const dir = p.endsWith('/') ? p : `${p}/`;
    if (sfNorm === p || sfNorm.startsWith(dir)) {
      return true;
    }
  }
  return false;
}

let raw;
try {
  raw = readFileSync(lcovPath, 'utf8');
} catch {
  console.error(`check-coverage: missing or unreadable ${lcovPath} (run bun test --coverage first)`);
  process.exit(1);
}

let lf = 0;
let lh = 0;

for (const block of raw.split('end_of_record')) {
  let sf = '';
  let blockLf = 0;
  let blockLh = 0;
  for (const line of block.split(/\r?\n/)) {
    const s = line.trim();
    if (s.startsWith('SF:')) sf = s.slice(3).trim();
    if (s.startsWith('LF:')) blockLf = Number(s.slice(3).trim()) || 0;
    if (s.startsWith('LH:')) blockLh = Number(s.slice(3).trim()) || 0;
  }
  if (!sf) continue;
  const n = normalizeSf(sf);
  if (isExcluded(n)) continue;
  lf += blockLf;
  lh += blockLh;
}

if (excludePrefixes.length > 0) {
  console.log(`check-coverage: excluding prefixes ${excludePrefixes.join(', ')}`);
}

const pct = lf === 0 ? 100 : (lh / lf) * 100;
console.log(`Line coverage: ${pct.toFixed(1)}% (${lh}/${lf} lines), minimum ${min}%`);

if (pct + 1e-9 < min) {
  console.error(`check-coverage: FAILED — raise coverage or set COVERAGE_MIN_LINES to lower the bar`);
  process.exit(1);
}
