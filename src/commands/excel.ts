import { readFile } from 'node:fs/promises';
import { Command } from 'commander';
import { resolveDriveLocationForCli } from '../lib/drive-location-cli.js';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import {
  addExcelTableRows,
  addExcelWorksheet,
  calculateExcelApplication,
  clearExcelRange,
  closeExcelWorkbookSession,
  createExcelTable,
  createExcelWorkbookSession,
  createExcelWorksheetChart,
  createExcelWorksheetPivotTable,
  deleteExcelTable,
  deleteExcelTableRow,
  deleteExcelWorksheet,
  deleteExcelWorksheetChart,
  deleteExcelWorksheetPivotTable,
  getExcelRange,
  getExcelTable,
  getExcelTableColumn,
  getExcelUsedRange,
  getExcelWorkbook,
  getExcelWorkbookNamedItem,
  getExcelWorksheet,
  getExcelWorksheetNamedItem,
  getExcelWorksheetPivotTable,
  listExcelTableColumns,
  listExcelTableRows,
  listExcelTables,
  listExcelWorkbookNames,
  listExcelWorksheetCharts,
  listExcelWorksheetNames,
  listExcelWorksheetPivotTables,
  listExcelWorksheets,
  patchExcelRange,
  patchExcelTable,
  patchExcelTableColumn,
  patchExcelTableRow,
  patchExcelWorksheetChart,
  patchExcelWorksheetPivotTable,
  refreshAllExcelWorksheetPivotTables,
  refreshExcelWorkbookSession,
  refreshExcelWorksheetPivotTable,
  updateExcelWorksheet
} from '../lib/graph-excel-client.js';
import {
  addExcelWorkbookCommentReply,
  createExcelWorkbookComment,
  getExcelWorkbookComment,
  listExcelWorkbookComments,
  patchExcelWorkbookComment
} from '../lib/graph-excel-comments-client.js';
import { toJsonError } from '../lib/json-error.js';
import { checkReadOnly } from '../lib/utils.js';

export const excelCommand = new Command('excel').description(
  'Excel workbook on OneDrive / SharePoint (Graph): worksheets, ranges, tables, pivots, charts, names, application, sessions, workbook comments beta (`Files.ReadWrite.All`; see GRAPH_SCOPES.md)'
);

excelCommand
  .command('worksheets')
  .description('List worksheets in a drive item workbook (drive item id from files list)')
  .argument('<itemId>', 'Drive item id of the Excel file')
  .option('--user <upn>', "Target user's default OneDrive (not with --drive-id / --site-id)")
  .option('--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)')
  .option('--site-id <id>', 'SharePoint site id (default document library)')
  .option('--library-drive-id <id>', 'Library drive id (only with --site-id)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (itemId: string, opts: { user?: string; json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await listExcelWorksheets(auth.token, itemId, resolveDriveLocationForCli(opts));
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(r.data, null, 2));
      return;
    }
    for (const w of r.data) {
      console.log(`${w.name ?? '(sheet)'}\t${w.id ?? ''}`);
    }
  });

excelCommand
  .command('worksheet-get')
  .description('Get one worksheet by name or id')
  .argument('<itemId>', 'Drive item id of the Excel file')
  .argument('<sheet>', 'Worksheet name or id')
  .option('--user <upn>', "Target user's default OneDrive (not with --drive-id / --site-id)")
  .option('--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)')
  .option('--site-id <id>', 'SharePoint site id (default document library)')
  .option('--library-drive-id <id>', 'Library drive id (only with --site-id)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      sheet: string,
      opts: { user?: string; json?: boolean; token?: string; identity?: string }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getExcelWorksheet(auth.token, itemId, sheet, resolveDriveLocationForCli(opts));
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(opts.json ? JSON.stringify(r.data, null, 2) : `${r.data.name ?? ''}\t${r.data.id ?? ''}`);
    }
  );

excelCommand
  .command('worksheet-add')
  .description('Add a worksheet (POST …/workbook/worksheets/add)')
  .argument('<itemId>', 'Drive item id of the Excel file')
  .requiredOption('--name <name>', 'Name for the new sheet')
  .option('--user <upn>', "Target user's default OneDrive (not with --drive-id / --site-id)")
  .option('--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)')
  .option('--site-id <id>', 'SharePoint site id (default document library)')
  .option('--library-drive-id <id>', 'Library drive id (only with --site-id)')
  .option('--session-id <id>', 'Optional workbook-session-id header')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      opts: { name: string; sessionId?: string; user?: string; json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await addExcelWorksheet(
        auth.token,
        itemId,
        opts.name,
        opts.sessionId,
        resolveDriveLocationForCli(opts)
      );
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(opts.json ? JSON.stringify(r.data, null, 2) : `${r.data.name ?? ''}\t${r.data.id ?? ''}`);
    }
  );

excelCommand
  .command('worksheet-update')
  .description('PATCH a worksheet (name, visibility, position — use --json-file for full body)')
  .argument('<itemId>', 'Drive item id')
  .argument('<sheet>', 'Worksheet name or id')
  .option('--json-file <path>', 'Full JSON patch body (overrides single-field flags)')
  .option('--name <name>', 'New display name')
  .option('--visibility <v>', 'visible | hidden (if supported)')
  .option('--position <n>', 'Sheet position (integer)', (v) => parseInt(v, 10))
  .option('--user <upn>', "Target user's default OneDrive (not with --drive-id / --site-id)")
  .option('--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)')
  .option('--site-id <id>', 'SharePoint site id (default document library)')
  .option('--library-drive-id <id>', 'Library drive id (only with --site-id)')
  .option('--session-id <id>', 'Optional workbook-session-id header')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      sheet: string,
      opts: {
        jsonFile?: string;
        name?: string;
        visibility?: string;
        position?: number;
        sessionId?: string;
        user?: string;
        json?: boolean;
        token?: string;
        identity?: string;
      },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      let patch: Record<string, unknown>;
      if (opts.jsonFile?.trim()) {
        patch = JSON.parse(await readFile(opts.jsonFile.trim(), 'utf-8')) as Record<string, unknown>;
      } else {
        patch = {};
        if (opts.name?.trim()) patch.name = opts.name.trim();
        if (opts.visibility?.trim()) patch.visibility = opts.visibility.trim();
        if (opts.position !== undefined && !Number.isNaN(opts.position)) patch.position = opts.position;
        if (Object.keys(patch).length === 0) {
          console.error('Error: provide --json-file or at least one of --name, --visibility, --position');
          process.exit(1);
        }
      }
      const r = await updateExcelWorksheet(
        auth.token,
        itemId,
        sheet,
        patch,
        opts.sessionId,
        resolveDriveLocationForCli(opts)
      );
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(opts.json ? JSON.stringify(r.data, null, 2) : `${r.data.name ?? ''}\t${r.data.id ?? ''}`);
    }
  );

excelCommand
  .command('worksheet-delete')
  .description('Delete a worksheet')
  .argument('<itemId>', 'Drive item id')
  .argument('<sheet>', 'Worksheet name or id')
  .option('--user <upn>', "Target user's default OneDrive (not with --drive-id / --site-id)")
  .option('--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)')
  .option('--site-id <id>', 'SharePoint site id (default document library)')
  .option('--library-drive-id <id>', 'Library drive id (only with --site-id)')
  .option('--session-id <id>', 'Optional workbook-session-id header')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      sheet: string,
      opts: { sessionId?: string; user?: string; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await deleteExcelWorksheet(auth.token, itemId, sheet, opts.sessionId, resolveDriveLocationForCli(opts));
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log('Deleted.');
    }
  );

excelCommand
  .command('range')
  .description('Read a range (A1 notation) from a worksheet')
  .argument('<itemId>', 'Drive item id of the Excel file')
  .argument('<sheet>', 'Worksheet name or id')
  .argument('<address>', 'Range address e.g. A1:D10')
  .option('--user <upn>', "Target user's default OneDrive (not with --drive-id / --site-id)")
  .option('--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)')
  .option('--site-id <id>', 'SharePoint site id (default document library)')
  .option('--library-drive-id <id>', 'Library drive id (only with --site-id)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      sheet: string,
      address: string,
      opts: { user?: string; json?: boolean; token?: string; identity?: string }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getExcelRange(auth.token, itemId, sheet, address, resolveDriveLocationForCli(opts));
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(opts.json ? JSON.stringify(r.data, null, 2) : JSON.stringify(r.data.values ?? r.data, null, 2));
    }
  );

excelCommand
  .command('range-patch')
  .description('PATCH a range (e.g. values or format); body via --json-file')
  .argument('<itemId>', 'Drive item id')
  .argument('<sheet>', 'Worksheet name or id')
  .argument('<address>', 'Range in A1 notation')
  .requiredOption('--json-file <path>', 'JSON body for PATCH (e.g. { "values": [[1,2]] })')
  .option('--session-id <id>', 'Optional workbook-session-id header')
  .option('--user <upn>', "Target user's default OneDrive (not with --drive-id / --site-id)")
  .option('--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)')
  .option('--site-id <id>', 'SharePoint site id (default document library)')
  .option('--library-drive-id <id>', 'Library drive id (only with --site-id)')
  .option('--json', 'Output full response JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      sheet: string,
      address: string,
      opts: { jsonFile: string; sessionId?: string; user?: string; json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const body = JSON.parse(await readFile(opts.jsonFile.trim(), 'utf-8')) as Record<string, unknown>;
      const r = await patchExcelRange(
        auth.token,
        itemId,
        sheet,
        address,
        body,
        opts.sessionId,
        resolveDriveLocationForCli(opts)
      );
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(opts.json ? JSON.stringify(r.data, null, 2) : JSON.stringify(r.data.values ?? r.data, null, 2));
    }
  );

excelCommand
  .command('used-range')
  .description('Read the worksheet used range (GET …/worksheets/{sheet}/usedRange)')
  .argument('<itemId>', 'Drive item id of the Excel file')
  .argument('<sheet>', 'Worksheet name or id')
  .option('--values-only', 'valuesOnly=true (ignore format-only cells)')
  .option('--user <upn>', "Target user's default OneDrive (not with --drive-id / --site-id)")
  .option('--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)')
  .option('--site-id <id>', 'SharePoint site id (default document library)')
  .option('--library-drive-id <id>', 'Library drive id (only with --site-id)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      sheet: string,
      opts: {
        valuesOnly?: boolean;
        user?: string;
        json?: boolean;
        token?: string;
        identity?: string;
      }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getExcelUsedRange(auth.token, itemId, sheet, opts.valuesOnly, resolveDriveLocationForCli(opts));
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(opts.json ? JSON.stringify(r.data, null, 2) : JSON.stringify(r.data.values ?? r.data, null, 2));
    }
  );

excelCommand
  .command('range-clear')
  .description('POST …/range(address=…)/clear (Graph range.clear); body via --json-file (use {} if none)')
  .argument('<itemId>', 'Drive item id')
  .argument('<sheet>', 'Worksheet name or id')
  .argument('<address>', 'Range in A1 notation')
  .requiredOption('--json-file <path>', 'JSON body for clear (e.g. {})')
  .option('--session-id <id>', 'Optional workbook-session-id header')
  .option('--user <upn>', "Target user's default OneDrive (not with --drive-id / --site-id)")
  .option('--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)')
  .option('--site-id <id>', 'SharePoint site id (default document library)')
  .option('--library-drive-id <id>', 'Library drive id (only with --site-id)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      sheet: string,
      address: string,
      opts: { jsonFile: string; sessionId?: string; user?: string; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      let body: Record<string, unknown>;
      try {
        body = JSON.parse(await readFile(opts.jsonFile.trim(), 'utf-8')) as Record<string, unknown>;
      } catch (e) {
        console.error(e instanceof Error ? e.message : 'Invalid --json-file');
        process.exit(1);
      }
      const r = await clearExcelRange(
        auth.token,
        itemId,
        sheet,
        address,
        body,
        opts.sessionId,
        resolveDriveLocationForCli(opts)
      );
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify({ ok: true }, null, 2));
    }
  );

excelCommand
  .command('workbook-get')
  .description('GET …/workbook (optional OData query e.g. --query "$select=application")')
  .argument('<itemId>', 'Drive item id')
  .option('--query <q>', 'OData query string without leading ?')
  .option('--user <upn>', "Target user's default OneDrive (not with --drive-id / --site-id)")
  .option('--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)')
  .option('--site-id <id>', 'SharePoint site id (default document library)')
  .option('--library-drive-id <id>', 'Library drive id (only with --site-id)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      opts: { query?: string; user?: string; json?: boolean; token?: string; identity?: string }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getExcelWorkbook(auth.token, itemId, opts.query, resolveDriveLocationForCli(opts));
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(opts.json ? JSON.stringify(r.data, null, 2) : JSON.stringify(r.data, null, 2));
    }
  );

excelCommand
  .command('application-calculate')
  .description('POST …/workbook/application/calculate (body e.g. { "calculationType": "Recalculate" })')
  .argument('<itemId>', 'Drive item id')
  .requiredOption('--json-file <path>', 'JSON body')
  .option('--session-id <id>', 'Optional workbook-session-id header')
  .option('--user <upn>', "Target user's default OneDrive (not with --drive-id / --site-id)")
  .option('--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)')
  .option('--site-id <id>', 'SharePoint site id (default document library)')
  .option('--library-drive-id <id>', 'Library drive id (only with --site-id)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      opts: { jsonFile: string; sessionId?: string; user?: string; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const body = JSON.parse(await readFile(opts.jsonFile.trim(), 'utf-8')) as Record<string, unknown>;
      const r = await calculateExcelApplication(
        auth.token,
        itemId,
        body,
        opts.sessionId,
        resolveDriveLocationForCli(opts)
      );
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify({ ok: true }, null, 2));
    }
  );

excelCommand
  .command('tables')
  .description('List Excel tables (whole workbook, or one worksheet with --sheet)')
  .argument('<itemId>', 'Drive item id of the Excel file')
  .option('--sheet <name>', 'Worksheet name or id (omit to list all workbook tables)')
  .option('--user <upn>', "Target user's default OneDrive (not with --drive-id / --site-id)")
  .option('--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)')
  .option('--site-id <id>', 'SharePoint site id (default document library)')
  .option('--library-drive-id <id>', 'Library drive id (only with --site-id)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      opts: { sheet?: string; user?: string; json?: boolean; token?: string; identity?: string }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await listExcelTables(auth.token, itemId, opts.sheet, resolveDriveLocationForCli(opts));
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(r.data, null, 2));
        return;
      }
      for (const t of r.data) {
        console.log(`${t.name ?? '(table)'}\t${t.style ?? ''}\t${t.id ?? ''}`);
      }
    }
  );

excelCommand
  .command('table-get')
  .description('Get one table by id')
  .argument('<itemId>', 'Drive item id')
  .argument('<tableId>', 'Table id')
  .option('--user <upn>', "Target user's default OneDrive (not with --drive-id / --site-id)")
  .option('--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)')
  .option('--site-id <id>', 'SharePoint site id (default document library)')
  .option('--library-drive-id <id>', 'Library drive id (only with --site-id)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      tableId: string,
      opts: { user?: string; json?: boolean; token?: string; identity?: string }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getExcelTable(auth.token, itemId, tableId, resolveDriveLocationForCli(opts));
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(opts.json ? JSON.stringify(r.data, null, 2) : `${r.data.name ?? ''}\t${r.data.id ?? ''}`);
    }
  );

excelCommand
  .command('table-rows')
  .description('List rows for a table')
  .argument('<itemId>', 'Drive item id')
  .argument('<tableId>', 'Table id')
  .option('--top <n>', 'Max rows', '500')
  .option('--user <upn>', "Target user's default OneDrive (not with --drive-id / --site-id)")
  .option('--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)')
  .option('--site-id <id>', 'SharePoint site id (default document library)')
  .option('--library-drive-id <id>', 'Library drive id (only with --site-id)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      tableId: string,
      opts: { top?: string; user?: string; json?: boolean; token?: string; identity?: string }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      let top: number | undefined;
      if (opts.top !== undefined) {
        top = parseInt(opts.top, 10);
        if (Number.isNaN(top) || top < 0) {
          console.error('Error: --top must be a non-negative integer');
          process.exit(1);
        }
      }
      const r = await listExcelTableRows(auth.token, itemId, tableId, top, resolveDriveLocationForCli(opts));
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(opts.json ? JSON.stringify(r.data, null, 2) : JSON.stringify(r.data, null, 2));
    }
  );

excelCommand
  .command('table-rows-add')
  .description('Add rows to a table (POST …/rows/add; see Graph `workbookTableRow` body)')
  .argument('<itemId>', 'Drive item id')
  .argument('<tableId>', 'Table id')
  .requiredOption('--json-file <path>', 'JSON e.g. { "index": null, "values": [["a","b"]] }')
  .option('--session-id <id>', 'Optional workbook-session-id header')
  .option('--user <upn>', "Target user's default OneDrive (not with --drive-id / --site-id)")
  .option('--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)')
  .option('--site-id <id>', 'SharePoint site id (default document library)')
  .option('--library-drive-id <id>', 'Library drive id (only with --site-id)')
  .option('--json', 'Output response JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      tableId: string,
      opts: { jsonFile: string; sessionId?: string; user?: string; json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const body = JSON.parse(await readFile(opts.jsonFile.trim(), 'utf-8')) as Record<string, unknown>;
      const r = await addExcelTableRows(
        auth.token,
        itemId,
        tableId,
        body,
        opts.sessionId,
        resolveDriveLocationForCli(opts)
      );
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(r.data ?? { ok: true }, null, 2));
      } else {
        console.log('OK');
      }
    }
  );

excelCommand
  .command('table-add')
  .description('Create a table (POST …/workbook/tables or …/worksheets/{sheet}/tables with --sheet)')
  .argument('<itemId>', 'Drive item id')
  .requiredOption(
    '--json-file <path>',
    'JSON [workbookTable](https://learn.microsoft.com/en-us/graph/api/resources/workbooktable) body'
  )
  .option('--sheet <name>', 'Worksheet name or id (omit for workbook-level /tables)')
  .option('--session-id <id>', 'Optional workbook-session-id header')
  .option('--user <upn>', "Target user's default OneDrive (not with --drive-id / --site-id)")
  .option('--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)')
  .option('--site-id <id>', 'SharePoint site id (default document library)')
  .option('--library-drive-id <id>', 'Library drive id (only with --site-id)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      opts: {
        jsonFile: string;
        sheet?: string;
        sessionId?: string;
        user?: string;
        token?: string;
        identity?: string;
      },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const body = JSON.parse(await readFile(opts.jsonFile.trim(), 'utf-8')) as Record<string, unknown>;
      const r = await createExcelTable(
        auth.token,
        itemId,
        body,
        opts.sheet,
        opts.sessionId,
        resolveDriveLocationForCli(opts)
      );
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data, null, 2));
    }
  );

excelCommand
  .command('table-patch')
  .description('PATCH a table by id')
  .argument('<itemId>', 'Drive item id')
  .argument('<tableId>', 'Table id')
  .requiredOption('--json-file <path>', 'JSON patch body')
  .option('--session-id <id>', 'Optional workbook-session-id header')
  .option('--user <upn>', "Target user's default OneDrive (not with --drive-id / --site-id)")
  .option('--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)')
  .option('--site-id <id>', 'SharePoint site id (default document library)')
  .option('--library-drive-id <id>', 'Library drive id (only with --site-id)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      tableId: string,
      opts: { jsonFile: string; sessionId?: string; user?: string; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const patch = JSON.parse(await readFile(opts.jsonFile.trim(), 'utf-8')) as Record<string, unknown>;
      const r = await patchExcelTable(
        auth.token,
        itemId,
        tableId,
        patch,
        opts.sessionId,
        resolveDriveLocationForCli(opts)
      );
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data, null, 2));
    }
  );

excelCommand
  .command('table-delete')
  .description('Delete a table by id')
  .argument('<itemId>', 'Drive item id')
  .argument('<tableId>', 'Table id')
  .option('--session-id <id>', 'Optional workbook-session-id header')
  .option('--user <upn>', "Target user's default OneDrive (not with --drive-id / --site-id)")
  .option('--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)')
  .option('--site-id <id>', 'SharePoint site id (default document library)')
  .option('--library-drive-id <id>', 'Library drive id (only with --site-id)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      tableId: string,
      opts: { sessionId?: string; user?: string; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await deleteExcelTable(auth.token, itemId, tableId, opts.sessionId, resolveDriveLocationForCli(opts));
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify({ ok: true }, null, 2));
    }
  );

excelCommand
  .command('table-columns')
  .description('List columns for a table')
  .argument('<itemId>', 'Drive item id')
  .argument('<tableId>', 'Table id')
  .option('--user <upn>', "Target user's default OneDrive (not with --drive-id / --site-id)")
  .option('--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)')
  .option('--site-id <id>', 'SharePoint site id (default document library)')
  .option('--library-drive-id <id>', 'Library drive id (only with --site-id)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      tableId: string,
      opts: { user?: string; json?: boolean; token?: string; identity?: string }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await listExcelTableColumns(auth.token, itemId, tableId, resolveDriveLocationForCli(opts));
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(opts.json ? JSON.stringify(r.data, null, 2) : JSON.stringify(r.data, null, 2));
    }
  );

excelCommand
  .command('table-column-get')
  .description('Get one table column by id')
  .argument('<itemId>', 'Drive item id')
  .argument('<tableId>', 'Table id')
  .argument('<columnId>', 'Column id')
  .option('--user <upn>', "Target user's default OneDrive (not with --drive-id / --site-id)")
  .option('--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)')
  .option('--site-id <id>', 'SharePoint site id (default document library)')
  .option('--library-drive-id <id>', 'Library drive id (only with --site-id)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      tableId: string,
      columnId: string,
      opts: { user?: string; json?: boolean; token?: string; identity?: string }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getExcelTableColumn(auth.token, itemId, tableId, columnId, resolveDriveLocationForCli(opts));
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(opts.json ? JSON.stringify(r.data, null, 2) : JSON.stringify(r.data, null, 2));
    }
  );

excelCommand
  .command('table-column-patch')
  .description('PATCH a table column')
  .argument('<itemId>', 'Drive item id')
  .argument('<tableId>', 'Table id')
  .argument('<columnId>', 'Column id')
  .requiredOption('--json-file <path>', 'JSON patch body')
  .option('--session-id <id>', 'Optional workbook-session-id header')
  .option('--user <upn>', "Target user's default OneDrive (not with --drive-id / --site-id)")
  .option('--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)')
  .option('--site-id <id>', 'SharePoint site id (default document library)')
  .option('--library-drive-id <id>', 'Library drive id (only with --site-id)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      tableId: string,
      columnId: string,
      opts: { jsonFile: string; sessionId?: string; user?: string; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const patch = JSON.parse(await readFile(opts.jsonFile.trim(), 'utf-8')) as Record<string, unknown>;
      const r = await patchExcelTableColumn(
        auth.token,
        itemId,
        tableId,
        columnId,
        patch,
        opts.sessionId,
        resolveDriveLocationForCli(opts)
      );
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data, null, 2));
    }
  );

excelCommand
  .command('table-row-patch')
  .description('PATCH a table row (row id is usually index from table-rows)')
  .argument('<itemId>', 'Drive item id')
  .argument('<tableId>', 'Table id')
  .argument('<rowId>', 'Row id')
  .requiredOption('--json-file <path>', 'JSON patch body (e.g. values)')
  .option('--session-id <id>', 'Optional workbook-session-id header')
  .option('--user <upn>', "Target user's default OneDrive (not with --drive-id / --site-id)")
  .option('--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)')
  .option('--site-id <id>', 'SharePoint site id (default document library)')
  .option('--library-drive-id <id>', 'Library drive id (only with --site-id)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      tableId: string,
      rowId: string,
      opts: { jsonFile: string; sessionId?: string; user?: string; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const patch = JSON.parse(await readFile(opts.jsonFile.trim(), 'utf-8')) as Record<string, unknown>;
      const r = await patchExcelTableRow(
        auth.token,
        itemId,
        tableId,
        rowId,
        patch,
        opts.sessionId,
        resolveDriveLocationForCli(opts)
      );
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data, null, 2));
    }
  );

excelCommand
  .command('table-row-delete')
  .description('DELETE a table row by id')
  .argument('<itemId>', 'Drive item id')
  .argument('<tableId>', 'Table id')
  .argument('<rowId>', 'Row id')
  .option('--session-id <id>', 'Optional workbook-session-id header')
  .option('--user <upn>', "Target user's default OneDrive (not with --drive-id / --site-id)")
  .option('--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)')
  .option('--site-id <id>', 'SharePoint site id (default document library)')
  .option('--library-drive-id <id>', 'Library drive id (only with --site-id)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      tableId: string,
      rowId: string,
      opts: { sessionId?: string; user?: string; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await deleteExcelTableRow(
        auth.token,
        itemId,
        tableId,
        rowId,
        opts.sessionId,
        resolveDriveLocationForCli(opts)
      );
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify({ ok: true }, null, 2));
    }
  );

excelCommand
  .command('names')
  .description('List defined names in the workbook')
  .argument('<itemId>', 'Drive item id')
  .option('--user <upn>', "Target user's default OneDrive (not with --drive-id / --site-id)")
  .option('--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)')
  .option('--site-id <id>', 'SharePoint site id (default document library)')
  .option('--library-drive-id <id>', 'Library drive id (only with --site-id)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (itemId: string, opts: { user?: string; json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await listExcelWorkbookNames(auth.token, itemId, resolveDriveLocationForCli(opts));
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log(opts.json ? JSON.stringify(r.data, null, 2) : JSON.stringify(r.data, null, 2));
  });

excelCommand
  .command('name-get')
  .description('Get one workbook-scoped named item by id (from excel names)')
  .argument('<itemId>', 'Drive item id')
  .argument('<nameId>', 'Named item id')
  .option('--user <upn>', "Target user's default OneDrive (not with --drive-id / --site-id)")
  .option('--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)')
  .option('--site-id <id>', 'SharePoint site id (default document library)')
  .option('--library-drive-id <id>', 'Library drive id (only with --site-id)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      nameId: string,
      opts: { user?: string; json?: boolean; token?: string; identity?: string }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getExcelWorkbookNamedItem(auth.token, itemId, nameId, resolveDriveLocationForCli(opts));
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(opts.json ? JSON.stringify(r.data, null, 2) : JSON.stringify(r.data, null, 2));
    }
  );

excelCommand
  .command('worksheet-names')
  .description('List defined names scoped to a worksheet')
  .argument('<itemId>', 'Drive item id')
  .argument('<sheet>', 'Worksheet name or id')
  .option('--user <upn>', "Target user's default OneDrive (not with --drive-id / --site-id)")
  .option('--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)')
  .option('--site-id <id>', 'SharePoint site id (default document library)')
  .option('--library-drive-id <id>', 'Library drive id (only with --site-id)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      sheet: string,
      opts: { user?: string; json?: boolean; token?: string; identity?: string }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await listExcelWorksheetNames(auth.token, itemId, sheet, resolveDriveLocationForCli(opts));
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(opts.json ? JSON.stringify(r.data, null, 2) : JSON.stringify(r.data, null, 2));
    }
  );

excelCommand
  .command('worksheet-name-get')
  .description('Get one worksheet-scoped named item')
  .argument('<itemId>', 'Drive item id')
  .argument('<sheet>', 'Worksheet name or id')
  .argument('<nameId>', 'Named item id')
  .option('--user <upn>', "Target user's default OneDrive (not with --drive-id / --site-id)")
  .option('--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)')
  .option('--site-id <id>', 'SharePoint site id (default document library)')
  .option('--library-drive-id <id>', 'Library drive id (only with --site-id)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      sheet: string,
      nameId: string,
      opts: { user?: string; json?: boolean; token?: string; identity?: string }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getExcelWorksheetNamedItem(auth.token, itemId, sheet, nameId, resolveDriveLocationForCli(opts));
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(opts.json ? JSON.stringify(r.data, null, 2) : JSON.stringify(r.data, null, 2));
    }
  );

excelCommand
  .command('charts')
  .description('List charts on a worksheet')
  .argument('<itemId>', 'Drive item id')
  .argument('<sheet>', 'Worksheet name or id')
  .option('--user <upn>', "Target user's default OneDrive (not with --drive-id / --site-id)")
  .option('--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)')
  .option('--site-id <id>', 'SharePoint site id (default document library)')
  .option('--library-drive-id <id>', 'Library drive id (only with --site-id)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      sheet: string,
      opts: { user?: string; json?: boolean; token?: string; identity?: string }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await listExcelWorksheetCharts(auth.token, itemId, sheet, resolveDriveLocationForCli(opts));
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(r.data, null, 2));
        return;
      }
      for (const c of r.data) {
        console.log(`${c.name ?? '(chart)'}\t${c.id ?? ''}`);
      }
    }
  );

excelCommand
  .command('chart-create')
  .description(
    'Create a chart on a worksheet (POST …/charts). Pass [workbookChart](https://learn.microsoft.com/en-us/graph/api/resources/workbookchart) JSON via --json-file'
  )
  .argument('<itemId>', 'Drive item id of the Excel file')
  .argument('<sheet>', 'Worksheet name or id')
  .requiredOption('--json-file <path>', 'JSON body for the new chart resource')
  .option('--session-id <id>', 'Optional workbook-session-id header')
  .option('--user <upn>', "Target user's default OneDrive (not with --drive-id / --site-id)")
  .option('--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)')
  .option('--site-id <id>', 'SharePoint site id (default document library)')
  .option('--library-drive-id <id>', 'Library drive id (only with --site-id)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      sheet: string,
      opts: {
        jsonFile: string;
        sessionId?: string;
        user?: string;
        driveId?: string;
        siteId?: string;
        libraryDriveId?: string;
        token?: string;
        identity?: string;
      },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      let chartBody: Record<string, unknown>;
      try {
        const raw = await readFile(opts.jsonFile.trim(), 'utf8');
        chartBody = JSON.parse(raw) as Record<string, unknown>;
      } catch (e) {
        console.error(e instanceof Error ? e.message : 'Invalid --json-file');
        process.exit(1);
      }
      const r = await createExcelWorksheetChart(
        auth.token,
        itemId,
        sheet,
        chartBody,
        opts.sessionId,
        resolveDriveLocationForCli(opts)
      );
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data, null, 2));
    }
  );

excelCommand
  .command('chart-patch')
  .description(
    'Patch an existing chart (PATCH …/charts/{name}). Pass partial [workbookChart](https://learn.microsoft.com/en-us/graph/api/resources/workbookchart) JSON via --json-file; chart name is usually the chart id from excel charts.'
  )
  .argument('<itemId>', 'Drive item id of the Excel file')
  .argument('<sheet>', 'Worksheet name or id')
  .argument('<chartName>', 'Chart name (id) from excel charts list')
  .requiredOption('--json-file <path>', 'JSON body for PATCH')
  .option('--session-id <id>', 'Optional workbook-session-id header')
  .option('--user <upn>', "Target user's default OneDrive (not with --drive-id / --site-id)")
  .option('--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)')
  .option('--site-id <id>', 'SharePoint site id (default document library)')
  .option('--library-drive-id <id>', 'Library drive id (only with --site-id)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      sheet: string,
      chartName: string,
      opts: {
        jsonFile: string;
        sessionId?: string;
        user?: string;
        driveId?: string;
        siteId?: string;
        libraryDriveId?: string;
        token?: string;
        identity?: string;
      },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      let body: Record<string, unknown>;
      try {
        const raw = await readFile(opts.jsonFile.trim(), 'utf8');
        body = JSON.parse(raw) as Record<string, unknown>;
      } catch (e) {
        console.error(e instanceof Error ? e.message : 'Invalid --json-file');
        process.exit(1);
      }
      const r = await patchExcelWorksheetChart(
        auth.token,
        itemId,
        sheet,
        chartName,
        body,
        opts.sessionId,
        resolveDriveLocationForCli(opts)
      );
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data, null, 2));
    }
  );

excelCommand
  .command('chart-delete')
  .description('Delete a chart from a worksheet (DELETE …/charts/{name})')
  .argument('<itemId>', 'Drive item id of the Excel file')
  .argument('<sheet>', 'Worksheet name or id')
  .argument('<chartName>', 'Chart name (id) from excel charts list')
  .option('--session-id <id>', 'Optional workbook-session-id header')
  .option('--user <upn>', "Target user's default OneDrive (not with --drive-id / --site-id)")
  .option('--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)')
  .option('--site-id <id>', 'SharePoint site id (default document library)')
  .option('--library-drive-id <id>', 'Library drive id (only with --site-id)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      sheet: string,
      chartName: string,
      opts: {
        sessionId?: string;
        user?: string;
        driveId?: string;
        siteId?: string;
        libraryDriveId?: string;
        token?: string;
        identity?: string;
      },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await deleteExcelWorksheetChart(
        auth.token,
        itemId,
        sheet,
        chartName,
        opts.sessionId,
        resolveDriveLocationForCli(opts)
      );
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify({ ok: true }, null, 2));
    }
  );

excelCommand
  .command('pivot-tables')
  .description('List pivot tables on a worksheet')
  .argument('<itemId>', 'Drive item id')
  .argument('<sheet>', 'Worksheet name or id')
  .option('--user <upn>', "Target user's default OneDrive (not with --drive-id / --site-id)")
  .option('--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)')
  .option('--site-id <id>', 'SharePoint site id (default document library)')
  .option('--library-drive-id <id>', 'Library drive id (only with --site-id)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      sheet: string,
      opts: { user?: string; json?: boolean; token?: string; identity?: string }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await listExcelWorksheetPivotTables(auth.token, itemId, sheet, resolveDriveLocationForCli(opts));
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(opts.json ? JSON.stringify(r.data, null, 2) : JSON.stringify(r.data, null, 2));
    }
  );

excelCommand
  .command('pivot-table-get')
  .description('Get one pivot table by id')
  .argument('<itemId>', 'Drive item id')
  .argument('<sheet>', 'Worksheet name or id')
  .argument('<pivotId>', 'Pivot table id')
  .option('--user <upn>', "Target user's default OneDrive (not with --drive-id / --site-id)")
  .option('--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)')
  .option('--site-id <id>', 'SharePoint site id (default document library)')
  .option('--library-drive-id <id>', 'Library drive id (only with --site-id)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      sheet: string,
      pivotId: string,
      opts: { user?: string; json?: boolean; token?: string; identity?: string }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getExcelWorksheetPivotTable(auth.token, itemId, sheet, pivotId, resolveDriveLocationForCli(opts));
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(opts.json ? JSON.stringify(r.data, null, 2) : JSON.stringify(r.data, null, 2));
    }
  );

excelCommand
  .command('pivot-table-create')
  .description('Create a pivot table (POST …/worksheets/{sheet}/pivotTables)')
  .argument('<itemId>', 'Drive item id')
  .argument('<sheet>', 'Worksheet name or id')
  .requiredOption(
    '--json-file <path>',
    'JSON [workbookPivotTable](https://learn.microsoft.com/en-us/graph/api/resources/workbookpivottable) body'
  )
  .option('--session-id <id>', 'Optional workbook-session-id header')
  .option('--user <upn>', "Target user's default OneDrive (not with --drive-id / --site-id)")
  .option('--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)')
  .option('--site-id <id>', 'SharePoint site id (default document library)')
  .option('--library-drive-id <id>', 'Library drive id (only with --site-id)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      sheet: string,
      opts: { jsonFile: string; sessionId?: string; user?: string; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const body = JSON.parse(await readFile(opts.jsonFile.trim(), 'utf-8')) as Record<string, unknown>;
      const r = await createExcelWorksheetPivotTable(
        auth.token,
        itemId,
        sheet,
        body,
        opts.sessionId,
        resolveDriveLocationForCli(opts)
      );
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data, null, 2));
    }
  );

excelCommand
  .command('pivot-table-patch')
  .description('PATCH a pivot table')
  .argument('<itemId>', 'Drive item id')
  .argument('<sheet>', 'Worksheet name or id')
  .argument('<pivotId>', 'Pivot table id')
  .requiredOption('--json-file <path>', 'JSON patch body')
  .option('--session-id <id>', 'Optional workbook-session-id header')
  .option('--user <upn>', "Target user's default OneDrive (not with --drive-id / --site-id)")
  .option('--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)')
  .option('--site-id <id>', 'SharePoint site id (default document library)')
  .option('--library-drive-id <id>', 'Library drive id (only with --site-id)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      sheet: string,
      pivotId: string,
      opts: { jsonFile: string; sessionId?: string; user?: string; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const patch = JSON.parse(await readFile(opts.jsonFile.trim(), 'utf-8')) as Record<string, unknown>;
      const r = await patchExcelWorksheetPivotTable(
        auth.token,
        itemId,
        sheet,
        pivotId,
        patch,
        opts.sessionId,
        resolveDriveLocationForCli(opts)
      );
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data, null, 2));
    }
  );

excelCommand
  .command('pivot-table-delete')
  .description('Delete a pivot table')
  .argument('<itemId>', 'Drive item id')
  .argument('<sheet>', 'Worksheet name or id')
  .argument('<pivotId>', 'Pivot table id')
  .option('--session-id <id>', 'Optional workbook-session-id header')
  .option('--user <upn>', "Target user's default OneDrive (not with --drive-id / --site-id)")
  .option('--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)')
  .option('--site-id <id>', 'SharePoint site id (default document library)')
  .option('--library-drive-id <id>', 'Library drive id (only with --site-id)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      sheet: string,
      pivotId: string,
      opts: { sessionId?: string; user?: string; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await deleteExcelWorksheetPivotTable(
        auth.token,
        itemId,
        sheet,
        pivotId,
        opts.sessionId,
        resolveDriveLocationForCli(opts)
      );
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify({ ok: true }, null, 2));
    }
  );

excelCommand
  .command('pivot-table-refresh')
  .description('POST …/pivotTables/{id}/refresh')
  .argument('<itemId>', 'Drive item id')
  .argument('<sheet>', 'Worksheet name or id')
  .argument('<pivotId>', 'Pivot table id')
  .option('--session-id <id>', 'Optional workbook-session-id header')
  .option('--user <upn>', "Target user's default OneDrive (not with --drive-id / --site-id)")
  .option('--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)')
  .option('--site-id <id>', 'SharePoint site id (default document library)')
  .option('--library-drive-id <id>', 'Library drive id (only with --site-id)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      sheet: string,
      pivotId: string,
      opts: { sessionId?: string; user?: string; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await refreshExcelWorksheetPivotTable(
        auth.token,
        itemId,
        sheet,
        pivotId,
        opts.sessionId,
        resolveDriveLocationForCli(opts)
      );
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify({ ok: true }, null, 2));
    }
  );

excelCommand
  .command('pivot-tables-refresh-all')
  .description('POST …/worksheets/{sheet}/pivotTables/refreshAll')
  .argument('<itemId>', 'Drive item id')
  .argument('<sheet>', 'Worksheet name or id')
  .option('--session-id <id>', 'Optional workbook-session-id header')
  .option('--user <upn>', "Target user's default OneDrive (not with --drive-id / --site-id)")
  .option('--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)')
  .option('--site-id <id>', 'SharePoint site id (default document library)')
  .option('--library-drive-id <id>', 'Library drive id (only with --site-id)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      sheet: string,
      opts: { sessionId?: string; user?: string; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await refreshAllExcelWorksheetPivotTables(
        auth.token,
        itemId,
        sheet,
        opts.sessionId,
        resolveDriveLocationForCli(opts)
      );
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify({ ok: true }, null, 2));
    }
  );

excelCommand
  .command('session-create')
  .description(
    'Create an Excel workbook session (POST …/workbook/createSession). Use session id as workbook-session-id on subsequent requests.'
  )
  .argument('<itemId>', 'Drive item id of the Excel file')
  .option('--volatile', 'persistChanges: false (discard edits when the session ends)', false)
  .option('--user <upn>', "Target user's default OneDrive (not with --drive-id / --site-id)")
  .option('--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)')
  .option('--site-id <id>', 'SharePoint site id (default document library)')
  .option('--library-drive-id <id>', 'Library drive id (only with --site-id)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      opts: { volatile?: boolean; user?: string; json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const persistChanges = opts.volatile !== true;
      const r = await createExcelWorkbookSession(auth.token, itemId, persistChanges, resolveDriveLocationForCli(opts));
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(r.data, null, 2));
        return;
      }
      console.log(`sessionId: ${r.data.id}`);
    }
  );

excelCommand
  .command('session-close')
  .description('Close an Excel workbook session (POST …/workbook/closeSession)')
  .argument('<itemId>', 'Drive item id of the Excel file')
  .requiredOption('--session-id <id>', 'Session id from session-create')
  .option('--user <upn>', "Target user's default OneDrive (not with --drive-id / --site-id)")
  .option('--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)')
  .option('--site-id <id>', 'SharePoint site id (default document library)')
  .option('--library-drive-id <id>', 'Library drive id (only with --site-id)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      opts: { sessionId: string; user?: string; json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await closeExcelWorkbookSession(auth.token, itemId, opts.sessionId, resolveDriveLocationForCli(opts));
      if (!r.ok) {
        if (opts.json) {
          console.log(JSON.stringify({ error: toJsonError(r.error?.message || 'Failed') }, null, 2));
        } else {
          console.error(`Error: ${r.error?.message}`);
        }
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify({ success: true }, null, 2));
      } else {
        console.log('Session closed.');
      }
    }
  );

excelCommand
  .command('session-refresh')
  .description('Refresh an Excel workbook session (POST …/workbook/refreshSession; extend lifetime)')
  .argument('<itemId>', 'Drive item id of the Excel file')
  .requiredOption('--session-id <id>', 'Session id from session-create')
  .option('--user <upn>', "Target user's default OneDrive (not with --drive-id / --site-id)")
  .option('--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)')
  .option('--site-id <id>', 'SharePoint site id (default document library)')
  .option('--library-drive-id <id>', 'Library drive id (only with --site-id)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      opts: { sessionId: string; user?: string; json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await refreshExcelWorkbookSession(auth.token, itemId, opts.sessionId, resolveDriveLocationForCli(opts));
      if (!r.ok) {
        if (opts.json) {
          console.log(JSON.stringify({ error: toJsonError(r.error?.message || 'Failed') }, null, 2));
        } else {
          console.error(`Error: ${r.error?.message}`);
        }
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify({ success: true }, null, 2));
      } else {
        console.log('Session refreshed.');
      }
    }
  );

excelCommand
  .command('comments-list <itemId>')
  .description(
    'List threaded workbook comments (always uses the Microsoft Graph beta root — workbookComment has no v1.0 equivalent; set GRAPH_BETA_URL to target a different beta host)'
  )
  .option('--user <upn>', "Target user's default OneDrive (not with --drive-id / --site-id)")
  .option('--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)')
  .option('--site-id <id>', 'SharePoint site id (default document library)')
  .option('--library-drive-id <id>', 'Library drive id (only with --site-id)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      opts: {
        user?: string;
        driveId?: string;
        siteId?: string;
        libraryDriveId?: string;
        json?: boolean;
        token?: string;
        identity?: string;
      }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const loc = resolveDriveLocationForCli(opts);
      const r = await listExcelWorkbookComments(auth.token, itemId, loc);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(opts.json ? JSON.stringify({ comments: r.data }, null, 2) : JSON.stringify(r.data, null, 2));
    }
  );

excelCommand
  .command('comments-get <itemId> <commentId>')
  .description('Get one workbook comment by id (always uses the Microsoft Graph beta root; see GRAPH_BETA_URL)')
  .option('--user <upn>', "Target user's default OneDrive (not with --drive-id / --site-id)")
  .option('--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)')
  .option('--site-id <id>', 'SharePoint site id (default document library)')
  .option('--library-drive-id <id>', 'Library drive id (only with --site-id)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      commentId: string,
      opts: {
        user?: string;
        driveId?: string;
        siteId?: string;
        libraryDriveId?: string;
        json?: boolean;
        token?: string;
        identity?: string;
      }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const loc = resolveDriveLocationForCli(opts);
      const r = await getExcelWorkbookComment(auth.token, itemId, commentId, loc);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(opts.json ? JSON.stringify(r.data, null, 2) : JSON.stringify(r.data, null, 2));
    }
  );

excelCommand
  .command('comments-create <itemId>')
  .description(
    'POST a new workbook comment (always uses the Microsoft Graph beta root; body shape see workbookComment; see GRAPH_BETA_URL)'
  )
  .requiredOption('--body <path>', 'JSON file for the new comment resource')
  .option('--user <upn>', "Target user's default OneDrive (not with --drive-id / --site-id)")
  .option('--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)')
  .option('--site-id <id>', 'SharePoint site id (default document library)')
  .option('--library-drive-id <id>', 'Library drive id (only with --site-id)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      opts: {
        body: string;
        user?: string;
        driveId?: string;
        siteId?: string;
        libraryDriveId?: string;
        token?: string;
        identity?: string;
      },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      let body: Record<string, unknown>;
      try {
        body = JSON.parse(await readFile(opts.body.trim(), 'utf-8')) as Record<string, unknown>;
      } catch (e) {
        console.error(e instanceof Error ? e.message : 'Invalid --body');
        process.exit(1);
      }
      const loc = resolveDriveLocationForCli(opts);
      const r = await createExcelWorkbookComment(auth.token, itemId, body, loc);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data, null, 2));
    }
  );

excelCommand
  .command('comments-reply <itemId> <commentId>')
  .description('POST a reply on a workbook comment (always uses the Microsoft Graph beta root; see GRAPH_BETA_URL)')
  .requiredOption('--body <path>', 'JSON file for the reply resource')
  .option('--user <upn>', "Target user's default OneDrive (not with --drive-id / --site-id)")
  .option('--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)')
  .option('--site-id <id>', 'SharePoint site id (default document library)')
  .option('--library-drive-id <id>', 'Library drive id (only with --site-id)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      commentId: string,
      opts: {
        body: string;
        user?: string;
        driveId?: string;
        siteId?: string;
        libraryDriveId?: string;
        token?: string;
        identity?: string;
      },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      let body: Record<string, unknown>;
      try {
        body = JSON.parse(await readFile(opts.body.trim(), 'utf-8')) as Record<string, unknown>;
      } catch (e) {
        console.error(e instanceof Error ? e.message : 'Invalid --body');
        process.exit(1);
      }
      const loc = resolveDriveLocationForCli(opts);
      const r = await addExcelWorkbookCommentReply(auth.token, itemId, commentId, body, loc);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data, null, 2));
    }
  );

excelCommand
  .command('comments-patch <itemId> <commentId>')
  .description(
    'PATCH a workbook comment (always uses the Microsoft Graph beta root; e.g. update content or associated task; see GRAPH_BETA_URL)'
  )
  .requiredOption('--body <path>', 'JSON file for the PATCH body')
  .option('--user <upn>', "Target user's default OneDrive (not with --drive-id / --site-id)")
  .option('--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)')
  .option('--site-id <id>', 'SharePoint site id (default document library)')
  .option('--library-drive-id <id>', 'Library drive id (only with --site-id)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      commentId: string,
      opts: {
        body: string;
        user?: string;
        driveId?: string;
        siteId?: string;
        libraryDriveId?: string;
        token?: string;
        identity?: string;
      },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      let body: Record<string, unknown>;
      try {
        body = JSON.parse(await readFile(opts.body.trim(), 'utf-8')) as Record<string, unknown>;
      } catch (e) {
        console.error(e instanceof Error ? e.message : 'Invalid --body');
        process.exit(1);
      }
      const loc = resolveDriveLocationForCli(opts);
      const r = await patchExcelWorkbookComment(auth.token, itemId, commentId, body, loc);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data, null, 2));
    }
  );
