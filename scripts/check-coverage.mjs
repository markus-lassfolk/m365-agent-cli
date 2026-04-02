#!/usr/bin/env node
/**
 * Enforce a minimum line coverage from Bun's lcov output (coverage/lcov.info).
 * Set COVERAGE_MIN_LINES (default 40) to adjust the bar.
 */
import { readFileSync } from 'node:fs';

const min = Number(process.env.COVERAGE_MIN_LINES ?? '40');
const path = process.argv[2] ?? 'coverage/lcov.info';

let raw;
try {
  raw = readFileSync(path, 'utf8');
} catch {
  console.error(`check-coverage: missing or unreadable ${path} (run bun test --coverage first)`);
  process.exit(1);
}

let lf = 0;
let lh = 0;
for (const line of raw.split('\n')) {
  if (line.startsWith('LF:')) lf += Number(line.slice(3).trim()) || 0;
  if (line.startsWith('LH:')) lh += Number(line.slice(3).trim()) || 0;
}

const pct = lf === 0 ? 100 : (lh / lf) * 100;
console.log(`Line coverage: ${pct.toFixed(1)}% (${lh}/${lf} lines), minimum ${min}%`);

if (pct + 1e-9 < min) {
  console.error(`check-coverage: FAILED — raise coverage or set COVERAGE_MIN_LINES to lower the bar`);
  process.exit(1);
}
