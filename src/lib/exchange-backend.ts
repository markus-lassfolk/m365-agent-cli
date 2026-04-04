/**
 * Exchange (mail/calendar) API backend selection for EWS → Graph migration.
 *
 * **Modes**
 * - **`auto` (default)** — **Graph first**: try Graph for every operation that has a Graph implementation; **EWS only as fallback**
 *   when Graph auth fails, the Graph API call fails, or the feature has **no Graph equivalent** (see `docs/MIGRATION_TRACKING.md`).
 *   A **successful** Graph response (including “empty list”) is **not** replaced by EWS.
 * - **`graph`** — Microsoft Graph only; never fall back to EWS (strict Graph errors).
 * - **`ews`** — Legacy EWS-only path for troubleshooting or parity testing.
 *
 * @see docs/MIGRATION_TRACKING.md
 * @see docs/GRAPH_V2_STATUS.md
 * @see docs/EWS_TO_GRAPH_MIGRATION_EPIC.md
 */

export type ExchangeBackend = 'graph' | 'ews' | 'auto';

const VALID: ReadonlySet<string> = new Set(['graph', 'ews', 'auto']);

/** Default: Graph first with EWS fallback — smooth upgrade for existing EWS-centric setups. */
export const DEFAULT_EXCHANGE_BACKEND: ExchangeBackend = 'auto';

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

/** True when Graph should be tried first for the current backend mode (`graph` or `auto`). */
export function shouldTryGraphFirst(): boolean {
  const b = getExchangeBackend();
  return b === 'graph' || b === 'auto';
}

/** `M365_EXCHANGE_BACKEND=auto` — Graph first; EWS only when Graph cannot satisfy the request. */
export function isAutoMode(): boolean {
  return getExchangeBackend() === 'auto';
}

/** `M365_EXCHANGE_BACKEND=graph` — never use EWS fallback. */
export function isGraphOnlyMode(): boolean {
  return getExchangeBackend() === 'graph';
}

/** `M365_EXCHANGE_BACKEND=ews` — EWS only (no Graph). */
export function isEwsExclusiveMode(): boolean {
  return getExchangeBackend() === 'ews';
}

/** True when EWS may be used (`ews` always; `auto` only as fallback after Graph failure or no Graph path). */
export function mayUseEws(): boolean {
  const b = getExchangeBackend();
  return b === 'ews' || b === 'auto';
}
