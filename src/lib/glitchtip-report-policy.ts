import type { EventHint } from '@sentry/core';

/**
 * Errno codes that usually indicate environment/network, not application bugs.
 * Used by GlitchTip `beforeSend` unless `GLITCHTIP_REPORT_ALL=1`.
 */
const GLITCHTIP_DEFAULT_IGNORE_ERRNO = new Set([
  'ECONNREFUSED',
  'ECONNRESET',
  'ETIMEDOUT',
  'ENOTFOUND',
  'EAI_AGAIN',
  'ENETUNREACH'
]);

/**
 * Whether this event hint should be dropped before sending to GlitchTip.
 * @param reportAll — When true (`GLITCHTIP_REPORT_ALL=1`), never suppress.
 */
export function glitchTipShouldSuppressHint(hint: EventHint | undefined, reportAll: boolean): boolean {
  if (reportAll) return false;

  const ex = hint?.originalException;
  if (ex && typeof ex === 'object' && 'code' in ex) {
    const code = String((ex as NodeJS.ErrnoException).code);
    if (code && GLITCHTIP_DEFAULT_IGNORE_ERRNO.has(code)) {
      return true;
    }
  }

  if (ex instanceof Error) {
    const msg = ex.message;
    if (/invalid[_ ]?grant|refresh[_ ]?token.*invalid|AADSTS\d+/i.test(msg)) {
      return true;
    }
  }

  return false;
}
