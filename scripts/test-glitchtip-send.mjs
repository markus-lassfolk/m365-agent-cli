#!/usr/bin/env node
/**
 * Sends a single test message to GlitchTip (Sentry-compatible ingest).
 * Usage: GLITCHTIP_DSN=... node scripts/test-glitchtip-send.mjs
 * Or:    node scripts/test-glitchtip-send.mjs   (uses DSN from docs/env.glitchtip.example if env unset)
 */
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { captureMessage, flush, init, isInitialized } from '@sentry/node';

const __dirname = dirname(fileURLToPath(import.meta.url));
const root = join(__dirname, '..');

function dsnFromExampleFile() {
  try {
    const raw = readFileSync(join(root, 'docs', 'env.glitchtip.example'), 'utf8');
    const m = raw.match(/^GLITCHTIP_DSN=(.+)$/m);
    return m?.[1]?.trim();
  } catch {
    return undefined;
  }
}

const dsn = process.env.GLITCHTIP_DSN?.trim() || process.env.SENTRY_DSN?.trim() || dsnFromExampleFile();

if (!dsn) {
  console.error('Set GLITCHTIP_DSN or add docs/env.glitchtip.example with GLITCHTIP_DSN=');
  process.exit(1);
}

function defaultReleaseFromPackageJson() {
  try {
    const pkg = JSON.parse(readFileSync(join(root, 'package.json'), 'utf8'));
    const v = typeof pkg.version === 'string' ? pkg.version.trim() : '';
    return v ? `m365-agent-cli@${v}` : 'm365-agent-cli-connectivity-test';
  } catch {
    return 'm365-agent-cli-connectivity-test';
  }
}

const release = process.env.GLITCHTIP_RELEASE?.trim() || defaultReleaseFromPackageJson();
console.log('release:', release);

init({
  dsn,
  environment: process.env.GLITCHTIP_ENVIRONMENT || 'production',
  release,
  tracesSampleRate: 0,
  sendDefaultPii: false
});

if (!isInitialized()) {
  console.error('Sentry failed to initialize');
  process.exit(1);
}

const id = captureMessage('m365-agent-cli manual GlitchTip connectivity test', 'info');
console.log('captureMessage submitted, event id:', id);
const ok = await flush(8000);
console.log(ok ? 'flush: ok (check GlitchTip project for the event)' : 'flush: timed out');
process.exit(ok ? 0 : 1);
