import { callGraph, GraphApiError, type GraphResponse, graphError, graphResult } from './graph-client.js';
import { graphUserPath } from './graph-user-path.js';

export interface ExcelWorksheet {
  id?: string;
  name?: string;
  position?: number;
  visibility?: string;
}

export interface ExcelRange {
  address?: string;
  values?: unknown[][];
}

export interface ExcelTable {
  id?: string;
  name?: string;
  style?: string;
}

export interface ExcelNamedItem {
  name?: string;
  type?: string;
  value?: unknown;
}

export interface ExcelChart {
  id?: string;
  name?: string;
  height?: number;
  width?: number;
  left?: number;
  top?: number;
}

export interface ExcelTableRow {
  index?: number;
  values?: unknown[][];
}

function driveItemWorkbookPrefix(user: string | undefined, itemId: string): string {
  const id = encodeURIComponent(itemId.trim());
  return `${graphUserPath(user, `drive/items/${id}/workbook`)}`;
}

function worksheetSegment(sheet: string): string {
  return `/worksheets/${encodeURIComponent(sheet.trim())}`;
}

function rangePathForAddress(user: string | undefined, itemId: string, sheet: string, address: string): string {
  const base = driveItemWorkbookPrefix(user, itemId);
  const addrLiteral = address.trim().replace(/'/g, "''");
  return `${base}${worksheetSegment(sheet)}/range(address='${addrLiteral}')`;
}

export async function listExcelWorksheets(
  token: string,
  itemId: string,
  user?: string
): Promise<GraphResponse<ExcelWorksheet[]>> {
  try {
    const path = `${driveItemWorkbookPrefix(user, itemId)}/worksheets`;
    const r = await callGraph<{ value: ExcelWorksheet[] }>(token, path);
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to list worksheets', r.error?.code, r.error?.status);
    }
    return graphResult(r.data.value ?? []);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to list worksheets');
  }
}

/**
 * Read a range (A1 notation). `sheet` is worksheet name or id; `address` e.g. A1:D10.
 */
export async function getExcelRange(
  token: string,
  itemId: string,
  sheet: string,
  address: string,
  user?: string
): Promise<GraphResponse<ExcelRange>> {
  try {
    const path = rangePathForAddress(user, itemId, sheet, address);
    const r = await callGraph<ExcelRange>(token, path);
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to read range', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to read range');
  }
}

/** Smallest used rectangle on the sheet. `valuesOnly` ignores formatted-empty cells per Graph. */
export async function getExcelUsedRange(
  token: string,
  itemId: string,
  sheet: string,
  user?: string,
  valuesOnly?: boolean
): Promise<GraphResponse<ExcelRange>> {
  try {
    const sheetEnc = encodeURIComponent(sheet.trim());
    const q = valuesOnly ? '?valuesOnly=true' : '';
    const path = `${driveItemWorkbookPrefix(user, itemId)}/worksheets/${sheetEnc}/usedRange${q}`;
    const r = await callGraph<ExcelRange>(token, path);
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to read used range', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to read used range');
  }
}

/** List tables in the workbook, or only on one worksheet if `worksheet` is set (name or id). */
export async function listExcelTables(
  token: string,
  itemId: string,
  user?: string,
  worksheet?: string
): Promise<GraphResponse<ExcelTable[]>> {
  try {
    const base = driveItemWorkbookPrefix(user, itemId);
    const path = worksheet ? `${base}/worksheets/${encodeURIComponent(worksheet.trim())}/tables` : `${base}/tables`;
    const r = await callGraph<{ value: ExcelTable[] }>(token, path);
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to list tables', r.error?.code, r.error?.status);
    }
    return graphResult(r.data.value ?? []);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to list tables');
  }
}

export async function getExcelWorksheet(
  token: string,
  itemId: string,
  sheet: string,
  user?: string
): Promise<GraphResponse<ExcelWorksheet>> {
  try {
    const path = `${driveItemWorkbookPrefix(user, itemId)}${worksheetSegment(sheet)}`;
    const r = await callGraph<ExcelWorksheet>(token, path);
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to get worksheet', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get worksheet');
  }
}

/** POST …/workbook/worksheets/add */
export async function addExcelWorksheet(
  token: string,
  itemId: string,
  name: string,
  user?: string
): Promise<GraphResponse<ExcelWorksheet>> {
  try {
    const path = `${driveItemWorkbookPrefix(user, itemId)}/worksheets/add`;
    const r = await callGraph<ExcelWorksheet>(token, path, {
      method: 'POST',
      body: JSON.stringify({ name: name.trim() })
    });
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to add worksheet', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to add worksheet');
  }
}

export async function updateExcelWorksheet(
  token: string,
  itemId: string,
  sheet: string,
  patch: Record<string, unknown>,
  user?: string
): Promise<GraphResponse<ExcelWorksheet>> {
  try {
    const path = `${driveItemWorkbookPrefix(user, itemId)}${worksheetSegment(sheet)}`;
    const r = await callGraph<ExcelWorksheet>(token, path, {
      method: 'PATCH',
      body: JSON.stringify(patch)
    });
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to update worksheet', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to update worksheet');
  }
}

export async function deleteExcelWorksheet(
  token: string,
  itemId: string,
  sheet: string,
  user?: string
): Promise<GraphResponse<void>> {
  try {
    const path = `${driveItemWorkbookPrefix(user, itemId)}${worksheetSegment(sheet)}`;
    const r = await callGraph<void>(token, path, { method: 'DELETE' }, false);
    if (!r.ok) {
      return graphError(r.error?.message || 'Failed to delete worksheet', r.error?.code, r.error?.status);
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to delete worksheet');
  }
}

/** PATCH range (e.g. `{ values: [[1,2]] }` or formats). */
export async function patchExcelRange(
  token: string,
  itemId: string,
  sheet: string,
  address: string,
  body: Record<string, unknown>,
  user?: string
): Promise<GraphResponse<ExcelRange>> {
  try {
    const path = rangePathForAddress(user, itemId, sheet, address);
    const r = await callGraph<ExcelRange>(token, path, {
      method: 'PATCH',
      body: JSON.stringify(body)
    });
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to patch range', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to patch range');
  }
}

export async function listExcelWorkbookNames(
  token: string,
  itemId: string,
  user?: string
): Promise<GraphResponse<ExcelNamedItem[]>> {
  try {
    const path = `${driveItemWorkbookPrefix(user, itemId)}/names`;
    const r = await callGraph<{ value: ExcelNamedItem[] }>(token, path);
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to list names', r.error?.code, r.error?.status);
    }
    return graphResult(r.data.value ?? []);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to list names');
  }
}

export async function getExcelTable(
  token: string,
  itemId: string,
  tableId: string,
  user?: string
): Promise<GraphResponse<ExcelTable>> {
  try {
    const path = `${driveItemWorkbookPrefix(user, itemId)}/tables/${encodeURIComponent(tableId.trim())}`;
    const r = await callGraph<ExcelTable>(token, path);
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to get table', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get table');
  }
}

export async function listExcelTableRows(
  token: string,
  itemId: string,
  tableId: string,
  user?: string,
  top?: number
): Promise<GraphResponse<ExcelTableRow[]>> {
  try {
    const t = top && top > 0 ? Math.min(top, 9999) : undefined;
    const qs = t ? `?$top=${t}` : '';
    const path = `${driveItemWorkbookPrefix(user, itemId)}/tables/${encodeURIComponent(tableId.trim())}/rows${qs}`;
    const r = await callGraph<{ value: ExcelTableRow[] }>(token, path);
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to list table rows', r.error?.code, r.error?.status);
    }
    return graphResult(r.data.value ?? []);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to list table rows');
  }
}

/** POST …/tables/{id}/rows/add — body e.g. `{ index: null, values: [["a","b"]] }`. */
export async function addExcelTableRows(
  token: string,
  itemId: string,
  tableId: string,
  body: Record<string, unknown>,
  user?: string
): Promise<GraphResponse<unknown>> {
  try {
    const path = `${driveItemWorkbookPrefix(user, itemId)}/tables/${encodeURIComponent(tableId.trim())}/rows/add`;
    const r = await callGraph<unknown>(token, path, {
      method: 'POST',
      body: JSON.stringify(body)
    });
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to add table rows', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to add table rows');
  }
}

export async function listExcelWorksheetCharts(
  token: string,
  itemId: string,
  sheet: string,
  user?: string
): Promise<GraphResponse<ExcelChart[]>> {
  try {
    const path = `${driveItemWorkbookPrefix(user, itemId)}${worksheetSegment(sheet)}/charts`;
    const r = await callGraph<{ value: ExcelChart[] }>(token, path);
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to list charts', r.error?.code, r.error?.status);
    }
    return graphResult(r.data.value ?? []);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to list charts');
  }
}
