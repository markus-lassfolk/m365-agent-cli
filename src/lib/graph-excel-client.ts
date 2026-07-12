import type { DriveLocation } from './drive-location.js';
import { DEFAULT_DRIVE_LOCATION, driveRootPrefix } from './drive-location.js';
import {
  callGraph,
  GraphApiError,
  type GraphResponse,
  graphError,
  graphResult,
  listGraphCollection
} from './graph-client.js';

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
  id?: string;
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

export type ExcelPivotTable = Record<string, unknown>;

export interface ExcelTableColumn {
  id?: string;
  name?: string;
  index?: number;
}

/** Merges `workbook-session-id` into Graph request init when `sessionId` is set (exported for unit tests). */
export function mergeExcelSessionInit(base: RequestInit, workbookSessionId?: string): RequestInit {
  const sid = workbookSessionId?.trim();
  if (!sid) return base;
  const merged = new Headers(base.headers as HeadersInit);
  merged.set('workbook-session-id', sid);
  return { ...base, headers: merged };
}

function driveItemWorkbookPrefix(location: DriveLocation, itemId: string): string {
  const id = encodeURIComponent(itemId.trim());
  return `${driveRootPrefix(location)}/items/${id}/workbook`;
}

function worksheetSegment(sheet: string): string {
  return `/worksheets/${encodeURIComponent(sheet.trim())}`;
}

function rangePathForAddress(location: DriveLocation, itemId: string, sheet: string, address: string): string {
  const base = driveItemWorkbookPrefix(location, itemId);
  const addrLiteral = address.trim().replace(/'/g, "''");
  return `${base}${worksheetSegment(sheet)}/range(address='${addrLiteral}')`;
}

export async function listExcelWorksheets(
  token: string,
  itemId: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION
): Promise<GraphResponse<ExcelWorksheet[]>> {
  return listGraphCollection<ExcelWorksheet>(
    token,
    `${driveItemWorkbookPrefix(location, itemId)}/worksheets`,
    'Failed to list worksheets'
  );
}

/**
 * Read a range (A1 notation). `sheet` is worksheet name or id; `address` e.g. A1:D10.
 */
export async function getExcelRange(
  token: string,
  itemId: string,
  sheet: string,
  address: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION
): Promise<GraphResponse<ExcelRange>> {
  try {
    const path = rangePathForAddress(location, itemId, sheet, address);
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
  valuesOnly?: boolean,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION
): Promise<GraphResponse<ExcelRange>> {
  try {
    const sheetEnc = encodeURIComponent(sheet.trim());
    // `valuesOnly` is an OData function parameter (usedRange(valuesOnly=true)), NOT a query option;
    // sent as `?valuesOnly=true` Graph ignores it and returns the formatting-inclusive range.
    const q = valuesOnly ? '(valuesOnly=true)' : '';
    const path = `${driveItemWorkbookPrefix(location, itemId)}/worksheets/${sheetEnc}/usedRange${q}`;
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
  worksheet?: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION
): Promise<GraphResponse<ExcelTable[]>> {
  const base = driveItemWorkbookPrefix(location, itemId);
  const path = worksheet ? `${base}/worksheets/${encodeURIComponent(worksheet.trim())}/tables` : `${base}/tables`;
  return listGraphCollection<ExcelTable>(token, path, 'Failed to list tables');
}

export async function getExcelWorksheet(
  token: string,
  itemId: string,
  sheet: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION
): Promise<GraphResponse<ExcelWorksheet>> {
  try {
    const path = `${driveItemWorkbookPrefix(location, itemId)}${worksheetSegment(sheet)}`;
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
  workbookSessionId?: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION
): Promise<GraphResponse<ExcelWorksheet>> {
  try {
    const path = `${driveItemWorkbookPrefix(location, itemId)}/worksheets/add`;
    const r = await callGraph<ExcelWorksheet>(
      token,
      path,
      mergeExcelSessionInit(
        {
          method: 'POST',
          body: JSON.stringify({ name: name.trim() })
        },
        workbookSessionId
      )
    );
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
  workbookSessionId?: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION
): Promise<GraphResponse<ExcelWorksheet>> {
  try {
    const path = `${driveItemWorkbookPrefix(location, itemId)}${worksheetSegment(sheet)}`;
    const r = await callGraph<ExcelWorksheet>(
      token,
      path,
      mergeExcelSessionInit(
        {
          method: 'PATCH',
          body: JSON.stringify(patch)
        },
        workbookSessionId
      )
    );
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
  workbookSessionId?: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION
): Promise<GraphResponse<void>> {
  try {
    const path = `${driveItemWorkbookPrefix(location, itemId)}${worksheetSegment(sheet)}`;
    const r = await callGraph<void>(token, path, mergeExcelSessionInit({ method: 'DELETE' }, workbookSessionId), false);
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
  workbookSessionId?: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION
): Promise<GraphResponse<ExcelRange>> {
  try {
    const path = rangePathForAddress(location, itemId, sheet, address);
    const r = await callGraph<ExcelRange>(
      token,
      path,
      mergeExcelSessionInit(
        {
          method: 'PATCH',
          body: JSON.stringify(body)
        },
        workbookSessionId
      )
    );
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to patch range', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to patch range');
  }
}

/** POST …/range(address='…')/clear — body per Graph `range.clear` (e.g. applyTo). */
export async function clearExcelRange(
  token: string,
  itemId: string,
  sheet: string,
  address: string,
  body: Record<string, unknown>,
  workbookSessionId?: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION
): Promise<GraphResponse<void>> {
  try {
    const path = `${rangePathForAddress(location, itemId, sheet, address)}/clear`;
    const r = await callGraph<void>(
      token,
      path,
      mergeExcelSessionInit(
        {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(body)
        },
        workbookSessionId
      ),
      false
    );
    if (!r.ok) {
      return graphError(r.error?.message || 'Failed to clear range', r.error?.code, r.error?.status);
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to clear range');
  }
}

/** GET …/workbook — optional OData query e.g. `$select=application`. */
export async function getExcelWorkbook(
  token: string,
  itemId: string,
  query?: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION
): Promise<GraphResponse<Record<string, unknown>>> {
  try {
    const q = query?.trim() ? (query.trim().startsWith('?') ? query.trim() : `?${query.trim()}`) : '';
    const path = `${driveItemWorkbookPrefix(location, itemId)}${q}`;
    const r = await callGraph<Record<string, unknown>>(token, path);
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to get workbook', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get workbook');
  }
}

/** POST …/workbook/application/calculate */
export async function calculateExcelApplication(
  token: string,
  itemId: string,
  body: Record<string, unknown>,
  workbookSessionId?: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION
): Promise<GraphResponse<void>> {
  try {
    const path = `${driveItemWorkbookPrefix(location, itemId)}/application/calculate`;
    const r = await callGraph<void>(
      token,
      path,
      mergeExcelSessionInit(
        {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(body)
        },
        workbookSessionId
      ),
      false
    );
    if (!r.ok) {
      return graphError(r.error?.message || 'Failed to calculate', r.error?.code, r.error?.status);
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to calculate');
  }
}

export async function listExcelWorkbookNames(
  token: string,
  itemId: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION
): Promise<GraphResponse<ExcelNamedItem[]>> {
  return listGraphCollection<ExcelNamedItem>(
    token,
    `${driveItemWorkbookPrefix(location, itemId)}/names`,
    'Failed to list names'
  );
}

/** GET …/workbook/names/{nameId} — `nameId` is the named item id from `names` list. */
export async function getExcelWorkbookNamedItem(
  token: string,
  itemId: string,
  nameId: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION
): Promise<GraphResponse<ExcelNamedItem>> {
  try {
    const path = `${driveItemWorkbookPrefix(location, itemId)}/names/${encodeURIComponent(nameId.trim())}`;
    const r = await callGraph<ExcelNamedItem>(token, path);
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to get named item', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get named item');
  }
}

/** List worksheet-scoped names (GET …/worksheets/{sheet}/names). */
export async function listExcelWorksheetNames(
  token: string,
  itemId: string,
  sheet: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION
): Promise<GraphResponse<ExcelNamedItem[]>> {
  return listGraphCollection<ExcelNamedItem>(
    token,
    `${driveItemWorkbookPrefix(location, itemId)}${worksheetSegment(sheet)}/names`,
    'Failed to list worksheet names'
  );
}

/** GET …/worksheets/{sheet}/names/{nameId} */
export async function getExcelWorksheetNamedItem(
  token: string,
  itemId: string,
  sheet: string,
  nameId: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION
): Promise<GraphResponse<ExcelNamedItem>> {
  try {
    const path = `${driveItemWorkbookPrefix(location, itemId)}${worksheetSegment(sheet)}/names/${encodeURIComponent(nameId.trim())}`;
    const r = await callGraph<ExcelNamedItem>(token, path);
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to get worksheet named item', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get worksheet named item');
  }
}

export async function getExcelTable(
  token: string,
  itemId: string,
  tableId: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION
): Promise<GraphResponse<ExcelTable>> {
  try {
    const path = `${driveItemWorkbookPrefix(location, itemId)}/tables/${encodeURIComponent(tableId.trim())}`;
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

/** POST …/workbook/tables or …/worksheets/{sheet}/tables when `worksheet` is set. */
export async function createExcelTable(
  token: string,
  itemId: string,
  body: Record<string, unknown>,
  worksheet?: string,
  workbookSessionId?: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION
): Promise<GraphResponse<ExcelTable>> {
  try {
    const base = driveItemWorkbookPrefix(location, itemId);
    const path = worksheet ? `${base}${worksheetSegment(worksheet)}/tables` : `${base}/tables`;
    const r = await callGraph<ExcelTable>(
      token,
      path,
      mergeExcelSessionInit(
        {
          method: 'POST',
          body: JSON.stringify(body)
        },
        workbookSessionId
      )
    );
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to create table', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to create table');
  }
}

export async function patchExcelTable(
  token: string,
  itemId: string,
  tableId: string,
  patch: Record<string, unknown>,
  workbookSessionId?: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION
): Promise<GraphResponse<ExcelTable>> {
  try {
    const path = `${driveItemWorkbookPrefix(location, itemId)}/tables/${encodeURIComponent(tableId.trim())}`;
    const r = await callGraph<ExcelTable>(
      token,
      path,
      mergeExcelSessionInit(
        {
          method: 'PATCH',
          body: JSON.stringify(patch)
        },
        workbookSessionId
      )
    );
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to patch table', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to patch table');
  }
}

export async function deleteExcelTable(
  token: string,
  itemId: string,
  tableId: string,
  workbookSessionId?: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION
): Promise<GraphResponse<void>> {
  try {
    const path = `${driveItemWorkbookPrefix(location, itemId)}/tables/${encodeURIComponent(tableId.trim())}`;
    const r = await callGraph<void>(token, path, mergeExcelSessionInit({ method: 'DELETE' }, workbookSessionId), false);
    if (!r.ok) {
      return graphError(r.error?.message || 'Failed to delete table', r.error?.code, r.error?.status);
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to delete table');
  }
}

export async function listExcelTableColumns(
  token: string,
  itemId: string,
  tableId: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION
): Promise<GraphResponse<ExcelTableColumn[]>> {
  return listGraphCollection<ExcelTableColumn>(
    token,
    `${driveItemWorkbookPrefix(location, itemId)}/tables/${encodeURIComponent(tableId.trim())}/columns`,
    'Failed to list table columns'
  );
}

export async function getExcelTableColumn(
  token: string,
  itemId: string,
  tableId: string,
  columnId: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION
): Promise<GraphResponse<ExcelTableColumn>> {
  try {
    const path = `${driveItemWorkbookPrefix(location, itemId)}/tables/${encodeURIComponent(tableId.trim())}/columns/${encodeURIComponent(columnId.trim())}`;
    const r = await callGraph<ExcelTableColumn>(token, path);
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to get table column', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get table column');
  }
}

export async function patchExcelTableColumn(
  token: string,
  itemId: string,
  tableId: string,
  columnId: string,
  patch: Record<string, unknown>,
  workbookSessionId?: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION
): Promise<GraphResponse<ExcelTableColumn>> {
  try {
    const path = `${driveItemWorkbookPrefix(location, itemId)}/tables/${encodeURIComponent(tableId.trim())}/columns/${encodeURIComponent(columnId.trim())}`;
    const r = await callGraph<ExcelTableColumn>(
      token,
      path,
      mergeExcelSessionInit(
        {
          method: 'PATCH',
          body: JSON.stringify(patch)
        },
        workbookSessionId
      )
    );
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to patch table column', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to patch table column');
  }
}

export async function listExcelTableRows(
  token: string,
  itemId: string,
  tableId: string,
  top?: number,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION
): Promise<GraphResponse<ExcelTableRow[]>> {
  const basePath = `${driveItemWorkbookPrefix(location, itemId)}/tables/${encodeURIComponent(tableId.trim())}/rows`;
  return listGraphCollection<ExcelTableRow>(token, basePath, 'Failed to list table rows', { top, maxTop: 9999 });
}

/** POST …/tables/{id}/rows/add — body e.g. `{ index: null, values: [["a","b"]] }`. */
export async function addExcelTableRows(
  token: string,
  itemId: string,
  tableId: string,
  body: Record<string, unknown>,
  workbookSessionId?: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION
): Promise<GraphResponse<unknown>> {
  try {
    const path = `${driveItemWorkbookPrefix(location, itemId)}/tables/${encodeURIComponent(tableId.trim())}/rows/add`;
    const r = await callGraph<unknown>(
      token,
      path,
      mergeExcelSessionInit(
        {
          method: 'POST',
          body: JSON.stringify(body)
        },
        workbookSessionId
      )
    );
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to add table rows', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to add table rows');
  }
}

/** PATCH …/tables/{tableId}/rows/{rowId} — `rowId` is typically the row index from list rows. */
export async function patchExcelTableRow(
  token: string,
  itemId: string,
  tableId: string,
  rowId: string,
  patch: Record<string, unknown>,
  workbookSessionId?: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION
): Promise<GraphResponse<ExcelTableRow>> {
  try {
    const path = `${driveItemWorkbookPrefix(location, itemId)}/tables/${encodeURIComponent(tableId.trim())}/rows/${encodeURIComponent(rowId.trim())}`;
    const r = await callGraph<ExcelTableRow>(
      token,
      path,
      mergeExcelSessionInit(
        {
          method: 'PATCH',
          body: JSON.stringify(patch)
        },
        workbookSessionId
      )
    );
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to patch table row', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to patch table row');
  }
}

export async function deleteExcelTableRow(
  token: string,
  itemId: string,
  tableId: string,
  rowId: string,
  workbookSessionId?: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION
): Promise<GraphResponse<void>> {
  try {
    const path = `${driveItemWorkbookPrefix(location, itemId)}/tables/${encodeURIComponent(tableId.trim())}/rows/${encodeURIComponent(rowId.trim())}`;
    const r = await callGraph<void>(token, path, mergeExcelSessionInit({ method: 'DELETE' }, workbookSessionId), false);
    if (!r.ok) {
      return graphError(r.error?.message || 'Failed to delete table row', r.error?.code, r.error?.status);
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to delete table row');
  }
}

export async function listExcelWorksheetCharts(
  token: string,
  itemId: string,
  sheet: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION
): Promise<GraphResponse<ExcelChart[]>> {
  return listGraphCollection<ExcelChart>(
    token,
    `${driveItemWorkbookPrefix(location, itemId)}${worksheetSegment(sheet)}/charts`,
    'Failed to list charts'
  );
}

/** POST …/worksheets/{sheet}/charts — body is a [workbookChart](https://learn.microsoft.com/en-us/graph/api/resources/workbookchart) JSON object. */
export async function createExcelWorksheetChart(
  token: string,
  itemId: string,
  sheet: string,
  chartBody: Record<string, unknown>,
  workbookSessionId?: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION
): Promise<GraphResponse<ExcelChart>> {
  try {
    const path = `${driveItemWorkbookPrefix(location, itemId)}${worksheetSegment(sheet)}/charts`;
    const r = await callGraph<ExcelChart>(
      token,
      path,
      mergeExcelSessionInit(
        {
          method: 'POST',
          body: JSON.stringify(chartBody)
        },
        workbookSessionId
      )
    );
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to create chart', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to create chart');
  }
}

function worksheetChartsSegment(sheet: string, chartName: string): string {
  const chartEnc = encodeURIComponent(chartName.trim());
  return `${worksheetSegment(sheet)}/charts/${chartEnc}`;
}

/** PATCH …/worksheets/{sheet}/charts/{name} — partial [workbookChart](https://learn.microsoft.com/en-us/graph/api/resources/workbookchart) JSON body. */
export async function patchExcelWorksheetChart(
  token: string,
  itemId: string,
  sheet: string,
  chartName: string,
  chartPatch: Record<string, unknown>,
  workbookSessionId?: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION
): Promise<GraphResponse<ExcelChart>> {
  try {
    const path = `${driveItemWorkbookPrefix(location, itemId)}${worksheetChartsSegment(sheet, chartName)}`;
    const r = await callGraph<ExcelChart>(
      token,
      path,
      mergeExcelSessionInit(
        {
          method: 'PATCH',
          body: JSON.stringify(chartPatch)
        },
        workbookSessionId
      )
    );
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to patch chart', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to patch chart');
  }
}

/** DELETE …/worksheets/{sheet}/charts/{name} */
export async function deleteExcelWorksheetChart(
  token: string,
  itemId: string,
  sheet: string,
  chartName: string,
  workbookSessionId?: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION
): Promise<GraphResponse<void>> {
  try {
    const path = `${driveItemWorkbookPrefix(location, itemId)}${worksheetChartsSegment(sheet, chartName)}`;
    const r = await callGraph<void>(token, path, mergeExcelSessionInit({ method: 'DELETE' }, workbookSessionId), false);
    if (!r.ok) {
      return graphError(r.error?.message || 'Failed to delete chart', r.error?.code, r.error?.status);
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to delete chart');
  }
}

function pivotTablesBase(location: DriveLocation, itemId: string, sheet: string): string {
  return `${driveItemWorkbookPrefix(location, itemId)}${worksheetSegment(sheet)}/pivotTables`;
}

export async function listExcelWorksheetPivotTables(
  token: string,
  itemId: string,
  sheet: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION
): Promise<GraphResponse<ExcelPivotTable[]>> {
  return listGraphCollection<ExcelPivotTable>(
    token,
    pivotTablesBase(location, itemId, sheet),
    'Failed to list pivot tables'
  );
}

export async function getExcelWorksheetPivotTable(
  token: string,
  itemId: string,
  sheet: string,
  pivotId: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION
): Promise<GraphResponse<ExcelPivotTable>> {
  try {
    const path = `${pivotTablesBase(location, itemId, sheet)}/${encodeURIComponent(pivotId.trim())}`;
    const r = await callGraph<ExcelPivotTable>(token, path);
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to get pivot table', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get pivot table');
  }
}

export async function createExcelWorksheetPivotTable(
  token: string,
  itemId: string,
  sheet: string,
  body: Record<string, unknown>,
  workbookSessionId?: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION
): Promise<GraphResponse<ExcelPivotTable>> {
  try {
    const path = pivotTablesBase(location, itemId, sheet);
    const r = await callGraph<ExcelPivotTable>(
      token,
      path,
      mergeExcelSessionInit(
        {
          method: 'POST',
          body: JSON.stringify(body)
        },
        workbookSessionId
      )
    );
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to create pivot table', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to create pivot table');
  }
}

export async function patchExcelWorksheetPivotTable(
  token: string,
  itemId: string,
  sheet: string,
  pivotId: string,
  patch: Record<string, unknown>,
  workbookSessionId?: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION
): Promise<GraphResponse<ExcelPivotTable>> {
  try {
    const path = `${pivotTablesBase(location, itemId, sheet)}/${encodeURIComponent(pivotId.trim())}`;
    const r = await callGraph<ExcelPivotTable>(
      token,
      path,
      mergeExcelSessionInit(
        {
          method: 'PATCH',
          body: JSON.stringify(patch)
        },
        workbookSessionId
      )
    );
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to patch pivot table', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to patch pivot table');
  }
}

export async function deleteExcelWorksheetPivotTable(
  token: string,
  itemId: string,
  sheet: string,
  pivotId: string,
  workbookSessionId?: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION
): Promise<GraphResponse<void>> {
  try {
    const path = `${pivotTablesBase(location, itemId, sheet)}/${encodeURIComponent(pivotId.trim())}`;
    const r = await callGraph<void>(token, path, mergeExcelSessionInit({ method: 'DELETE' }, workbookSessionId), false);
    if (!r.ok) {
      return graphError(r.error?.message || 'Failed to delete pivot table', r.error?.code, r.error?.status);
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to delete pivot table');
  }
}

export async function refreshExcelWorksheetPivotTable(
  token: string,
  itemId: string,
  sheet: string,
  pivotId: string,
  workbookSessionId?: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION
): Promise<GraphResponse<void>> {
  try {
    const path = `${pivotTablesBase(location, itemId, sheet)}/${encodeURIComponent(pivotId.trim())}/refresh`;
    const r = await callGraph<void>(
      token,
      path,
      mergeExcelSessionInit(
        {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: '{}'
        },
        workbookSessionId
      ),
      false
    );
    if (!r.ok) {
      return graphError(r.error?.message || 'Failed to refresh pivot table', r.error?.code, r.error?.status);
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to refresh pivot table');
  }
}

export async function refreshAllExcelWorksheetPivotTables(
  token: string,
  itemId: string,
  sheet: string,
  workbookSessionId?: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION
): Promise<GraphResponse<void>> {
  try {
    const path = `${pivotTablesBase(location, itemId, sheet)}/refreshAll`;
    const r = await callGraph<void>(
      token,
      path,
      mergeExcelSessionInit(
        {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: '{}'
        },
        workbookSessionId
      ),
      false
    );
    if (!r.ok) {
      return graphError(r.error?.message || 'Failed to refresh all pivot tables', r.error?.code, r.error?.status);
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to refresh all pivot tables');
  }
}

/** POST …/workbook/createSession — returns `{ id }` for concurrent workbook edits ([docs](https://learn.microsoft.com/en-us/graph/api/workbook-createsession)). */
export async function createExcelWorkbookSession(
  token: string,
  itemId: string,
  persistChanges = true,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION
): Promise<GraphResponse<{ id: string }>> {
  try {
    const path = `${driveItemWorkbookPrefix(location, itemId)}/createSession`;
    const r = await callGraph<{ id: string }>(token, path, {
      method: 'POST',
      body: JSON.stringify({ persistChanges })
    });
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to create workbook session', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to create workbook session');
  }
}

/** POST …/workbook/refreshSession — extend session lifetime ([docs](https://learn.microsoft.com/en-us/graph/api/workbook-refreshsession)). */
export async function refreshExcelWorkbookSession(
  token: string,
  itemId: string,
  sessionId: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION
): Promise<GraphResponse<void>> {
  try {
    const path = `${driveItemWorkbookPrefix(location, itemId)}/refreshSession`;
    const r = await callGraph<void>(
      token,
      path,
      {
        method: 'POST',
        headers: {
          'workbook-session-id': sessionId.trim(),
          'Content-Type': 'application/json'
        },
        body: '{}'
      },
      false
    );
    if (!r.ok) {
      return graphError(r.error?.message || 'Failed to refresh workbook session', r.error?.code, r.error?.status);
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to refresh workbook session');
  }
}

/** POST …/workbook/closeSession — pass session id from createSession. */
export async function closeExcelWorkbookSession(
  token: string,
  itemId: string,
  sessionId: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION
): Promise<GraphResponse<void>> {
  try {
    const path = `${driveItemWorkbookPrefix(location, itemId)}/closeSession`;
    const r = await callGraph<void>(
      token,
      path,
      {
        method: 'POST',
        headers: {
          'workbook-session-id': sessionId.trim(),
          'Content-Type': 'application/json'
        },
        body: '{}'
      },
      false
    );
    if (!r.ok) {
      return graphError(r.error?.message || 'Failed to close workbook session', r.error?.code, r.error?.status);
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to close workbook session');
  }
}
