import { validateUrl } from './url-validation.js';

const DEFAULT_GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
const DEFAULT_GRAPH_BETA = 'https://graph.microsoft.com/beta';

/** Microsoft Graph v1.0 root — resolved from `process.env.GRAPH_BASE_URL` on each call. */
export function getGraphBaseUrl(): string {
  return validateUrl(process.env.GRAPH_BASE_URL || DEFAULT_GRAPH_BASE, 'GRAPH_BASE_URL');
}

/** Microsoft Graph beta root — resolved from `process.env.GRAPH_BETA_URL` on each call. */
export function getGraphBetaUrl(): string {
  return validateUrl(process.env.GRAPH_BETA_URL || DEFAULT_GRAPH_BETA, 'GRAPH_BETA_URL');
}

/** Resolve v1.0 vs beta root for CLI `--beta` (or `GRAPH_BETA_URL` / `GRAPH_BASE_URL` env). */
export function graphApiRoot(useBeta?: boolean): string {
  return useBeta ? getGraphBetaUrl() : getGraphBaseUrl();
}
