import { readFile } from 'node:fs/promises';
import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import {
  addExcelTableRows,
  addExcelWorksheet,
  deleteExcelWorksheet,
  getExcelRange,
  getExcelTable,
  getExcelUsedRange,
  getExcelWorksheet,
  listExcelTableRows,
  listExcelTables,
  listExcelWorkbookNames,
  listExcelWorksheetCharts,
  listExcelWorksheets,
  patchExcelRange,
  updateExcelWorksheet
} from '../lib/graph-excel-client.js';
import { checkReadOnly } from '../lib/utils.js';

export const excelCommand = new Command('excel').description(
  'Excel workbook on OneDrive (Graph): worksheets CRUD, range read/patch, tables, rows, charts, names (`Files.ReadWrite.All`; see GRAPH_SCOPES.md)'
);

excelCommand
  .command('worksheets')
  .description('List worksheets in a drive item workbook (drive item id from files list)')
  .argument('<itemId>', 'Drive item id of the Excel file')
  .option('--user <upn>', "Mailbox UPN when reading another user's drive (optional)")
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (itemId: string, opts: { user?: string; json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await listExcelWorksheets(auth.token, itemId, opts.user);
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
  .option('--user <upn>', 'Delegated drive owner UPN (optional)')
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
      const r = await getExcelWorksheet(auth.token, itemId, sheet, opts.user);
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
  .option('--user <upn>', 'Delegated drive owner UPN (optional)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      opts: { name: string; user?: string; json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await addExcelWorksheet(auth.token, itemId, opts.name, opts.user);
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
  .option('--user <upn>', 'Delegated drive owner UPN (optional)')
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
      const r = await updateExcelWorksheet(auth.token, itemId, sheet, patch, opts.user);
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
  .option('--user <upn>', 'Delegated drive owner UPN (optional)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (itemId: string, sheet: string, opts: { user?: string; token?: string; identity?: string }, cmd: Command) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await deleteExcelWorksheet(auth.token, itemId, sheet, opts.user);
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
  .option('--user <upn>', "Mailbox UPN when reading another user's drive (optional)")
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
      const r = await getExcelRange(auth.token, itemId, sheet, address, opts.user);
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
  .option('--user <upn>', 'Delegated drive owner UPN (optional)')
  .option('--json', 'Output full response JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      sheet: string,
      address: string,
      opts: { jsonFile: string; user?: string; json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const body = JSON.parse(await readFile(opts.jsonFile.trim(), 'utf-8')) as Record<string, unknown>;
      const r = await patchExcelRange(auth.token, itemId, sheet, address, body, opts.user);
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
  .option('--user <upn>', "Mailbox UPN when reading another user's drive (optional)")
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
      const r = await getExcelUsedRange(auth.token, itemId, sheet, opts.user, opts.valuesOnly);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(opts.json ? JSON.stringify(r.data, null, 2) : JSON.stringify(r.data.values ?? r.data, null, 2));
    }
  );

excelCommand
  .command('tables')
  .description('List Excel tables (whole workbook, or one worksheet with --sheet)')
  .argument('<itemId>', 'Drive item id of the Excel file')
  .option('--sheet <name>', 'Worksheet name or id (omit to list all workbook tables)')
  .option('--user <upn>', "Mailbox UPN when reading another user's drive (optional)")
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
      const r = await listExcelTables(auth.token, itemId, opts.user, opts.sheet);
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
  .option('--user <upn>', 'Delegated drive owner UPN (optional)')
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
      const r = await getExcelTable(auth.token, itemId, tableId, opts.user);
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
  .option('--user <upn>', 'Delegated drive owner UPN (optional)')
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
      const top = opts.top ? parseInt(opts.top, 10) : undefined;
      const r = await listExcelTableRows(auth.token, itemId, tableId, opts.user, top);
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
  .option('--user <upn>', 'Delegated drive owner UPN (optional)')
  .option('--json', 'Output response JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      itemId: string,
      tableId: string,
      opts: { jsonFile: string; user?: string; json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const body = JSON.parse(await readFile(opts.jsonFile.trim(), 'utf-8')) as Record<string, unknown>;
      const r = await addExcelTableRows(auth.token, itemId, tableId, body, opts.user);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(opts.json && r.data !== undefined ? JSON.stringify(r.data, null, 2) : 'OK');
    }
  );

excelCommand
  .command('names')
  .description('List defined names in the workbook')
  .argument('<itemId>', 'Drive item id')
  .option('--user <upn>', 'Delegated drive owner UPN (optional)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (itemId: string, opts: { user?: string; json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await listExcelWorkbookNames(auth.token, itemId, opts.user);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log(opts.json ? JSON.stringify(r.data, null, 2) : JSON.stringify(r.data, null, 2));
  });

excelCommand
  .command('charts')
  .description('List charts on a worksheet')
  .argument('<itemId>', 'Drive item id')
  .argument('<sheet>', 'Worksheet name or id')
  .option('--user <upn>', 'Delegated drive owner UPN (optional)')
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
      const r = await listExcelWorksheetCharts(auth.token, itemId, sheet, opts.user);
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
