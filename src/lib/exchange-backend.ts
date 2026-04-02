/**
 * Exchange (mail/calendar) API backend selection for EWS → Graph migration.
 * On branch `dev_v2`, default is `graph` (pure Graph-first); set `M365_EXCHANGE_BACKEND=ews|auto` for legacy.
 *
 * @see docs/GRAPH_V2_STATUS.md
 * @see docs/EWS_TO_GRAPH_MIGRATION_EPIC.md
 */

export type ExchangeBackend = 'graph' | 'ews' | 'auto';

const VALID: ReadonlySet<string> = new Set(['graph', 'ews', 'auto']);

/** Default on dev_v2: Microsoft Graph only (no EWS for commands that honor this). */
export const DEFAULT_EXCHANGE_BACKEND: ExchangeBackend = 'graph';

/**
 * Reads `M365_EXCHANGE_BACKEND` (`graph` | `ews` | `auto`).
 * Invalid or empty values fall back to {@link DEFAULT_EXCHANGE_BACKEND}.
 */
export function getExchangeBackend(): ExchangeBackend {
  const raw = process.env.M365_EXCHANGE_BACKEND?.trim().toLowerCase();
  if (!raw || !VALID.has(raw)) {
    return DEFAULT_EXCHANGE_BACKEND;
  }
  return raw as ExchangeBackend;
}

/** True when Graph should be tried first for the current backend mode. */
export function shouldTryGraphFirst(): boolean {
  const b = getExchangeBackend();
  return b === 'graph' || b === 'auto';
}

/** True when EWS may be used (either exclusively or as fallback). */
export function mayUseEws(): boolean {
  const b = getExchangeBackend();
  return b === 'ews' || b === 'auto';
}
