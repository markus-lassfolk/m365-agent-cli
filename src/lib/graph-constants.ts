import { validateUrl } from './url-validation';

export const GRAPH_BASE_URL = validateUrl(
  process.env.GRAPH_BASE_URL || 'https://graph.microsoft.com/v1.0',
  'GRAPH_BASE_URL'
);

/** Microsoft Graph beta root (Planner favorites, roster, delta — subject to change). */
export const GRAPH_BETA_URL = validateUrl(
  process.env.GRAPH_BETA_URL || 'https://graph.microsoft.com/beta',
  'GRAPH_BETA_URL'
);
