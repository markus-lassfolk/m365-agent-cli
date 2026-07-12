import {
  callGraph,
  fetchAllPages,
  GraphApiError,
  type GraphResponse,
  graphError,
  graphResult
} from './graph-client.js';
import { graphUserPath } from './graph-user-path.js';

/** Preset color name from Graph (`preset0`..`preset24`). */
export interface OutlookMasterCategory {
  id: string;
  displayName: string;
  color: string;
}

/** Valid Microsoft Graph outlookCategory color constants (25 presets). */
export const OUTLOOK_CATEGORY_COLOR_PRESETS = [
  'preset0',
  'preset1',
  'preset2',
  'preset3',
  'preset4',
  'preset5',
  'preset6',
  'preset7',
  'preset8',
  'preset9',
  'preset10',
  'preset11',
  'preset12',
  'preset13',
  'preset14',
  'preset15',
  'preset16',
  'preset17',
  'preset18',
  'preset19',
  'preset20',
  'preset21',
  'preset22',
  'preset23',
  'preset24'
] as const;

export type OutlookCategoryColorPreset = (typeof OUTLOOK_CATEGORY_COLOR_PRESETS)[number];

export function isValidOutlookCategoryColor(s: string): s is OutlookCategoryColorPreset {
  return OUTLOOK_CATEGORY_COLOR_PRESETS.includes(s as OutlookCategoryColorPreset);
}

export async function listOutlookMasterCategories(
  token: string,
  user?: string
): Promise<GraphResponse<OutlookMasterCategory[]>> {
  return fetchAllPages<OutlookMasterCategory>(
    token,
    graphUserPath(user, 'outlook/masterCategories'),
    'Failed to list master categories'
  );
}

export async function createOutlookMasterCategory(
  token: string,
  displayName: string,
  color: string,
  user?: string
): Promise<GraphResponse<OutlookMasterCategory>> {
  const path = graphUserPath(user, 'outlook/masterCategories');
  try {
    const result = await callGraph<OutlookMasterCategory>(token, path, {
      method: 'POST',
      body: JSON.stringify({ displayName: displayName.trim(), color })
    });
    if (!result.ok || !result.data) {
      return graphError(
        result.error?.message || 'Failed to create master category',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to create master category');
  }
}

export async function updateOutlookMasterCategory(
  token: string,
  categoryId: string,
  updates: { displayName?: string; color?: string },
  user?: string
): Promise<GraphResponse<OutlookMasterCategory>> {
  const path = `${graphUserPath(user, 'outlook/masterCategories')}/${encodeURIComponent(categoryId)}`;
  try {
    const body: Record<string, string> = {};
    if (updates.displayName !== undefined) body.displayName = updates.displayName.trim();
    if (updates.color !== undefined) body.color = updates.color;
    const result = await callGraph<OutlookMasterCategory>(token, path, {
      method: 'PATCH',
      body: JSON.stringify(body)
    });
    if (!result.ok || !result.data) {
      return graphError(
        result.error?.message || 'Failed to update master category',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to update master category');
  }
}

export async function deleteOutlookMasterCategory(
  token: string,
  categoryId: string,
  user?: string
): Promise<GraphResponse<void>> {
  const path = `${graphUserPath(user, 'outlook/masterCategories')}/${encodeURIComponent(categoryId)}`;
  try {
    const result = await callGraph<void>(token, path, { method: 'DELETE' });
    if (!result.ok) {
      return graphError(
        result.error?.message || 'Failed to delete master category',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(undefined as undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to delete master category');
  }
}
