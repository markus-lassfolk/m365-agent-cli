#!/usr/bin/env node
/**
 * Enforce a minimum line coverage from Bun's lcov output (coverage/lcov.info).
 * Set COVERAGE_MIN_LINES (default 32) to adjust the bar.
 * Bun/OS versions can shift totals slightly; keep CI threshold below Bun's "All files" % if needed.
 */
import { readFileSync } from 'node:fs';

const min = Number(process.env.COVERAGE_MIN_LINES ?? '32');
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
for (const line of raw.split(/\r?\n/)) {
  const s = line.trim();
  if (s.startsWith('LF:')) lf += Number(s.slice(3).trim()) || 0;
  if (s.startsWith('LH:')) lh += Number(s.slice(3).trim()) || 0;
}

const pct = lf === 0 ? 100 : (lh / lf) * 100;
console.log(`Line coverage: ${pct.toFixed(1)}% (${lh}/${lf} lines), minimum ${min}%`);

if (pct + 1e-9 < min) {
  console.error(`check-coverage: FAILED — raise coverage or set COVERAGE_MIN_LINES to lower the bar`);
  process.exit(1);
}
