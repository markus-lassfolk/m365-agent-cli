/**
 * Optional error reporting to [GlitchTip](https://glitchtip.com/) (Sentry-compatible ingest).
 * Set `GLITCHTIP_DSN` or `SENTRY_DSN` to enable. No reporting when unset.
 *
 * Events are scrubbed to avoid PII: no argv content, no user/request/breadcrumbs, paths and
 * messages redacted (see stripSensitiveEventData).
 */

import type { ErrorEvent, EventHint, StackFrame } from '@sentry/core';
import { captureException, flush, init, isInitialized } from '@sentry/node';
import { checkGlitchTipEligibility } from './glitchtip-eligibility.js';
import { getPackageVersionSync } from './package-info.js';

/** Keys we never attach to reports (callers may pass `extra`; these are dropped). */
const EXTRA_KEY_DENYLIST =
  /^(body|bodyHtml|html|text|message|subject|snippet|preview|email|mail|address|token|password|secret|authorization|refresh|credential|content|attachment)/i;

const EMAIL_RE = /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b/g;
/** Looks like a JWT or opaque bearer token (redact). */
const JWT_LIKE_RE = /\beyJ[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}\b/gi;

function redactString(s: string): string {
  return s
    .replace(EMAIL_RE, '[email]')
    .replace(JWT_LIKE_RE, '[token]')
    .replace(/Bearer\s+[A-Za-z0-9._~-]+/gi, 'Bearer [token]')
    .replace(/\\Users\\[^\\]+\\/gi, '\\Users\\<user>\\')
    .replace(/\/Users\/[^/]+/g, '/Users/<user>')
    .replace(/\/home\/[^/]+/g, '/home/<user>');
}

function deepRedactValue(value: unknown, depth = 0): unknown {
  if (depth > 6) return '[truncated]';
  if (typeof value === 'string') return redactString(value);
  if (value === null || typeof value !== 'object') return value;
  if (Array.isArray(value)) return value.slice(0, 50).map((v) => deepRedactValue(v, depth + 1));
  const out: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(value)) {
    if (EXTRA_KEY_DENYLIST.test(k)) continue;
    out[k] = deepRedactValue(v, depth + 1);
  }
  return out;
}

function sanitizeExtraForReport(extra: Record<string, unknown>): Record<string, unknown> {
  const cleaned: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(extra)) {
    if (EXTRA_KEY_DENYLIST.test(k)) continue;
    cleaned[k] = deepRedactValue(v);
  }
  return cleaned;
}

/** Remove PII from the outbound event; keeps stack structure for debugging. */
function stripSensitiveEventData(event: ErrorEvent): void {
  delete event.user;
  delete event.request;
  delete event.server_name;
  event.breadcrumbs = [];

  if (event.sdkProcessingMetadata) {
    delete event.sdkProcessingMetadata;
  }

  if (event.message) {
    event.message = redactString(event.message);
  }
  if (event.logentry?.message) {
    event.logentry.message = redactString(event.logentry.message);
  }

  if (event.extra) {
    const next: Record<string, unknown> = {};
    for (const [k, v] of Object.entries(event.extra)) {
      if (EXTRA_KEY_DENYLIST.test(k)) continue;
      next[k] = deepRedactValue(v);
    }
    event.extra = next;
  }

  if (event.tags) {
    const next: Record<string, string | number | boolean> = {};
    for (const [k, v] of Object.entries(event.tags)) {
      if (typeof v === 'string') next[k] = redactString(v);
      else if (typeof v === 'number' || typeof v === 'boolean') next[k] = v;
    }
    event.tags = next;
  }

  const exValues = event.exception?.values;
  if (exValues) {
    for (const ex of exValues) {
      if (ex.value) ex.value = redactString(ex.value);
      scrubFrames(ex.stacktrace?.frames);
    }
  }

  if (event.modules) {
    const next: Record<string, string> = {};
    for (const [k, v] of Object.entries(event.modules)) {
      next[k] = typeof v === 'string' ? redactString(v) : String(v);
    }
    event.modules = next;
  }

  if (event.contexts) {
    const { os, runtime, app } = event.contexts;
    event.contexts = {};
    if (os && typeof os === 'object') {
      const o = os as Record<string, unknown>;
      event.contexts.os = {
        name: typeof o.name === 'string' ? o.name : undefined,
        version: typeof o.version === 'string' ? o.version : undefined,
        kernel_version: typeof o.kernel_version === 'string' ? o.kernel_version : undefined
      };
    }
    if (runtime && typeof runtime === 'object') {
      event.contexts.runtime = runtime;
    }
    if (app && typeof app === 'object') {
      const a = app as Record<string, unknown>;
      event.contexts.app = {
        app_start_time: typeof a.app_start_time === 'string' ? a.app_start_time : undefined,
        app_memory: typeof a.app_memory === 'number' ? a.app_memory : undefined
      };
    }
  }

  if (event.threads?.values) {
    for (const thread of event.threads.values) {
      delete thread.name;
      scrubFrames(thread.stacktrace?.frames);
    }
  }
}

function scrubFrames(frames: StackFrame[] | undefined): void {
  if (!frames) return;
  for (const frame of frames) {
    if (frame.filename) frame.filename = redactString(frame.filename);
    if (frame.abs_path) frame.abs_path = redactString(frame.abs_path);
    if (frame.context_line) frame.context_line = redactString(frame.context_line);
    if (frame.pre_context) frame.pre_context = frame.pre_context.map(redactString);
    if (frame.post_context) frame.post_context = frame.post_context.map(redactString);
    delete frame.vars;
  }
}

function getDsn(): string | undefined {
  const raw = process.env.GLITCHTIP_DSN?.trim() || process.env.SENTRY_DSN?.trim();
  if (!raw) return undefined;
  const disabled = process.env.GLITCHTIP_ENABLED === '0' || process.env.GLITCHTIP_ENABLED === 'false';
  if (disabled) return undefined;
  return raw;
}

/** Errno codes that usually indicate environment/network, not application bugs. */
const DEFAULT_IGNORE_ERRNO = new Set([
  'ECONNREFUSED',
  'ECONNRESET',
  'ETIMEDOUT',
  'ENOTFOUND',
  'EAI_AGAIN',
  'ENETUNREACH'
]);

function beforeSend(event: ErrorEvent, hint: EventHint): ErrorEvent | null {
  const reportAll = process.env.GLITCHTIP_REPORT_ALL === '1';
  if (!reportAll) {
    const ex = hint.originalException;
    if (ex && typeof ex === 'object' && 'code' in ex) {
      const code = String((ex as NodeJS.ErrnoException).code);
      if (code && DEFAULT_IGNORE_ERRNO.has(code)) {
        return null;
      }
    }

    if (ex instanceof Error) {
      const msg = ex.message;
      if (/invalid[_ ]?grant|refresh[_ ]?token.*invalid|AADSTS\d+/i.test(msg)) {
        return null;
      }
    }
  }

  stripSensitiveEventData(event);
  enrichSafeEvent(event);
  return event;
}

/** Safe CLI context only: no argv text (may contain addresses, paths, mail snippets). */
function enrichSafeEvent(event: ErrorEvent): void {
  const argv = process.argv.slice(2);
  const first = argv[0];
  event.tags = {
    ...(event.tags ?? {}),
    'cli.argc': String(argv.length)
  };
  if (first && /^[a-z][a-z0-9-]*$/i.test(first) && first.length <= 48) {
    event.tags['cli.command'] = first;
  }
}

let didInit = false;

/** Call once after `loadGlobalEnv()` so `.env` can define `GLITCHTIP_DSN`. */
export async function initGlitchTip(): Promise<void> {
  if (didInit) return;
  didInit = true;

  const dsn = getDsn();
  if (!dsn) return;

  const eligibility = await checkGlitchTipEligibility();
  if (!eligibility.ok) {
    if (process.env.GLITCHTIP_DEBUG_ELIGIBILITY === '1' || process.env.GLITCHTIP_DEBUG_ELIGIBILITY === 'true') {
      console.error(`[GlitchTip] Disabled: ${eligibility.reason ?? 'not eligible'}`);
    }
    return;
  }

  const releaseFromEnv = process.env.GLITCHTIP_RELEASE?.trim();
  const release =
    releaseFromEnv && releaseFromEnv.length > 0
      ? releaseFromEnv
      : `m365-agent-cli@${getPackageVersionSync()}`;

  init({
    dsn,
    sendDefaultPii: false,
    environment: process.env.GLITCHTIP_ENVIRONMENT || process.env.NODE_ENV || 'production',
    release,
    tracesSampleRate: 0,
    profilesSampleRate: 0,
    maxBreadcrumbs: 0,
    beforeBreadcrumb: () => null,
    beforeSend: beforeSend as (event: ErrorEvent, hint: EventHint) => ErrorEvent | null
  });
}

/** Wait for the transport to finish (call before `process.exit` after a capture). */
export async function flushGlitchTip(timeoutMs = 2000): Promise<boolean> {
  if (!isInitialized()) return true;
  return flush(timeoutMs);
}

/** Report an exception (e.g. Commander parse failure). No-op if GlitchTip is not configured. */
export function captureCliException(error: unknown, extra?: Record<string, unknown>): void {
  if (!isInitialized()) return;
  if (extra && Object.keys(extra).length > 0) {
    captureException(error, { extra: sanitizeExtraForReport(extra) });
  } else {
    captureException(error);
  }
}
