import { readFile } from 'node:fs/promises';
import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import { graphApiRoot } from '../lib/graph-client.js';
import {
  applyDeltaPageToState,
  assertDeltaScopeMatchesState,
  readDeltaStateFile,
  resolveDeltaContinuationUrl,
  writeDeltaStateFile
} from '../lib/graph-delta-state-file.js';
import { followSites, listFollowedSites, unfollowSites } from '../lib/graph-insights-client.js';
import {
  createListItem,
  createSitePermission,
  deleteListItem,
  deleteSitePermission,
  getAllListItemsPages,
  getListColumns,
  getListItem,
  getListItems,
  getListItemsDeltaPage,
  getListItemsPage,
  getListMetadata,
  getLists,
  getSiteByGraphPath,
  getSiteById,
  getSiteDefaultDriveId,
  getSiteDrives,
  getSitePermission,
  getSitePermissions,
  type SharePointListItem,
  updateListItem,
  updateSitePermission
} from '../lib/sharepoint-client.js';
import { checkReadOnly } from '../lib/utils.js';

function spApi(o: { beta?: boolean }): string {
  return graphApiRoot(!!o.beta);
}

async function parseFieldsJsonOrFile(opts: {
  fields?: string;
  jsonFile?: string;
  mode: 'create' | 'update';
}): Promise<Record<string, any>> {
  let raw: string;
  if (opts.jsonFile?.trim()) {
    try {
      raw = await readFile(opts.jsonFile.trim(), 'utf-8');
    } catch (e) {
      throw new Error(`Could not read --json-file: ${e instanceof Error ? e.message : String(e)}`);
    }
  } else if (opts.fields?.trim()) {
    raw = opts.fields.trim();
  } else {
    throw new Error(opts.mode === 'create' ? 'Provide --fields or --json-file' : 'Provide --fields or --json-file');
  }
  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch (e) {
    throw new Error(`Invalid JSON: ${e instanceof Error ? e.message : String(e)}`);
  }
  if (typeof parsed !== 'object' || parsed === null || Array.isArray(parsed)) {
    throw new Error('JSON must be an object');
  }
  const o = parsed as Record<string, unknown>;
  if (opts.mode === 'create') {
    if ('fields' in o && typeof o.fields === 'object' && o.fields !== null && !Array.isArray(o.fields)) {
      return o.fields as Record<string, any>;
    }
    return o as Record<string, any>;
  }
  if ('fields' in o && typeof o.fields === 'object' && o.fields !== null && !Array.isArray(o.fields)) {
    return o.fields as Record<string, any>;
  }
  return o as Record<string, any>;
}

function printListItemHuman(item: SharePointListItem): void {
  console.log(`Item ID: ${item.id}`);
  if (item.fields) {
    for (const [key, val] of Object.entries(item.fields)) {
      if (!key.startsWith('@odata')) {
        console.log(`  ${key}: ${val}`);
      }
    }
  }
  console.log('---');
}

function printListItemsPageHuman(items: SharePointListItem[], nextLink?: string): void {
  if (items.length === 0) {
    console.log('No items in this page.');
  } else {
    for (const item of items) {
      printListItemHuman(item);
    }
  }
  if (nextLink) {
    console.log(`nextLink:\t${nextLink}`);
    console.log('Re-run: sharepoint items --site-id … --list-id … --url "<nextLink>"');
  }
}

export const sharepointCommand = new Command('sharepoint')
  .description(
    'Microsoft SharePoint: sites, libraries, lists, items, site sharing permissions (owners), followed sites. Use --beta on subcommands for Graph beta host.'
  )
  .alias('sp');

sharepointCommand
  .command('resolve-site <siteResource>')
  .description(
    'Resolve a site by Graph path (GET /sites/{resource}) — e.g. `contoso.sharepoint.com:/sites/TeamName`. Prints site id and default document library drive id for `files --site-id`.'
  )
  .option('--json', 'Output as JSON')
  .option('--beta', 'Use Microsoft Graph beta API host for this call', false)
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (siteResource: string, opts: { json?: boolean; beta?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error || 'Unknown error'}`);
      process.exit(1);
    }
    const site = await getSiteByGraphPath(auth.token, siteResource, spApi(opts));
    if (!site.ok || !site.data) {
      console.error(`Error: ${site.error?.message || 'Unknown error'}`);
      process.exit(1);
    }
    const drive = await getSiteDefaultDriveId(auth.token, site.data.id, spApi(opts));
    const driveId = drive.ok && drive.data?.id ? drive.data.id : '';
    if (opts.json) {
      console.log(
        JSON.stringify(
          { site: site.data, defaultDriveId: driveId || null, driveError: drive.ok ? null : drive.error },
          null,
          2
        )
      );
      return;
    }
    console.log(`siteId:\t${site.data.id}`);
    if (site.data.displayName) console.log(`name:\t${site.data.displayName}`);
    if (site.data.webUrl) console.log(`webUrl:\t${site.data.webUrl}`);
    if (driveId) {
      console.log(`defaultDriveId:\t${driveId}`);
      console.log(`Example: m365-agent-cli files list --site-id "${site.data.id}"`);
      console.log(`Other libraries: m365-agent-cli sharepoint drives --site-id "${site.data.id}"`);
    } else if (!drive.ok) {
      console.error(`Warning: could not load default drive: ${drive.error?.message ?? 'unknown'}`);
    }
  });

sharepointCommand
  .command('get-site <siteId>')
  .description('Get a site by Graph site id (`GET /sites/{id}`) — displayName, webUrl, ids')
  .option('--json', 'Output as JSON')
  .option('--beta', 'Use Microsoft Graph beta API host for this call', false)
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (siteId: string, opts: { json?: boolean; beta?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error || 'Unknown error'}`);
      process.exit(1);
    }
    const site = await getSiteById(auth.token, siteId.trim(), spApi(opts));
    if (!site.ok || !site.data) {
      console.error(`Error: ${site.error?.message || 'Unknown error'}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(site.data, null, 2));
      return;
    }
    console.log(`id:\t${site.data.id}`);
    if (site.data.displayName) console.log(`displayName:\t${site.data.displayName}`);
    if (site.data.name) console.log(`name:\t${site.data.name}`);
    if (site.data.webUrl) console.log(`webUrl:\t${site.data.webUrl}`);
  });

sharepointCommand
  .command('drives')
  .description(
    'List document libraries / drives under a site (`GET /sites/{id}/drives`) — use drive id with `files --library-drive-id`'
  )
  .requiredOption('--site-id <id>', 'SharePoint Site ID')
  .option('--json', 'Output as JSON')
  .option('--beta', 'Use Microsoft Graph beta API host for this call', false)
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (opts: { siteId: string; json?: boolean; beta?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error || 'Unknown error'}`);
      process.exit(1);
    }
    const res = await getSiteDrives(auth.token, opts.siteId, spApi(opts));
    if (!res.ok) {
      console.error(`Error: ${res.error?.message || 'Unknown error'}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify({ value: res.data ?? [] }, null, 2));
      return;
    }
    const drives = res.data ?? [];
    if (drives.length === 0) {
      console.log('No drives returned.');
      return;
    }
    for (const d of drives) {
      console.log(`${d.id}\t${d.name ?? '(no name)'}${d.driveType ? `\t${d.driveType}` : ''}`);
      if (d.webUrl) console.log(`  ${d.webUrl}`);
    }
    console.log('');
    console.log('Use with files: m365-agent-cli files list --site-id "<siteId>" --library-drive-id "<driveId>"');
  });

sharepointCommand
  .command('get-list')
  .description('Get list metadata (`GET /sites/{id}/lists/{listId}`)')
  .requiredOption('--site-id <id>', 'SharePoint Site ID')
  .requiredOption('--list-id <id>', 'SharePoint List ID')
  .option('--json', 'Output as JSON')
  .option('--beta', 'Use Microsoft Graph beta API host for this call', false)
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (opts: {
      siteId: string;
      listId: string;
      json?: boolean;
      beta?: boolean;
      token?: string;
      identity?: string;
    }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error || 'Unknown error'}`);
        process.exit(1);
      }
      const res = await getListMetadata(auth.token, opts.siteId, opts.listId, spApi(opts));
      if (!res.ok || !res.data) {
        console.error(`Error: ${res.error?.message || 'Unknown error'}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(res.data, null, 2));
        return;
      }
      console.log(`${res.data.name} (${res.data.id})`);
      if (res.data.description) console.log(`  ${res.data.description}`);
      if (res.data.webUrl) console.log(`  ${res.data.webUrl}`);
    }
  );

sharepointCommand
  .command('columns')
  .description('List columns (schema) for a SharePoint list (`GET …/lists/{id}/columns`)')
  .requiredOption('--site-id <id>', 'SharePoint Site ID')
  .requiredOption('--list-id <id>', 'SharePoint List ID')
  .option('--json', 'Output as JSON')
  .option('--beta', 'Use Microsoft Graph beta API host for this call', false)
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (opts: {
      siteId: string;
      listId: string;
      json?: boolean;
      beta?: boolean;
      token?: string;
      identity?: string;
    }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error || 'Unknown error'}`);
        process.exit(1);
      }
      const res = await getListColumns(auth.token, opts.siteId, opts.listId, spApi(opts));
      if (!res.ok) {
        console.error(`Error: ${res.error?.message || 'Unknown error'}`);
        process.exit(1);
      }
      const cols = res.data ?? [];
      if (opts.json) {
        console.log(JSON.stringify({ value: cols }, null, 2));
        return;
      }
      if (cols.length === 0) {
        console.log('No columns returned.');
        return;
      }
      for (const c of cols) {
        const name = typeof c.name === 'string' ? c.name : typeof c.id === 'string' ? c.id : '?';
        const colType = typeof c.type === 'string' ? c.type : '';
        console.log(`${name}${colType ? `\t${colType}` : ''}`);
      }
    }
  );

sharepointCommand
  .command('lists')
  .description('List all SharePoint lists in a site')
  .requiredOption('--site-id <id>', 'SharePoint Site ID')
  .option('--json', 'Output as JSON')
  .option('--beta', 'Use Microsoft Graph beta API host for this call', false)
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (opts: { siteId: string; json?: boolean; beta?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error || 'Unknown error'}`);
      process.exit(1);
    }
    const res = await getLists(auth.token, opts.siteId, spApi(opts));
    if (!res.ok) {
      console.error(`Error listing lists: ${res.error?.message || 'Unknown error'}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(res.data, null, 2));
      return;
    }
    if (!res.data || res.data.length === 0) {
      console.log('No lists found in this site.');
      return;
    }
    for (const list of res.data) {
      console.log(`${list.name} (${list.id})`);
      if (list.description) console.log(`  ${list.description}`);
    }
  });

sharepointCommand
  .command('items')
  .description(
    'Get items from a SharePoint list. Default: all pages. With --url / --filter / --orderby / --top: one page unless --all-pages.'
  )
  .requiredOption('--site-id <id>', 'SharePoint Site ID')
  .requiredOption('--list-id <id>', 'SharePoint List ID')
  .option('--url <url>', 'Full @odata.nextLink URL (single page; overrides other query flags)')
  .option('--filter <odata>', 'OData $filter (implies paged mode unless combined with --all-pages)')
  .option('--orderby <odata>', 'OData $orderby')
  .option('--top <n>', 'Page size ($top, max 999)')
  .option('--all-pages', 'With --filter/--orderby/--top: follow @odata.nextLink until exhausted')
  .option('--json', 'Output as JSON')
  .option('--beta', 'Use Microsoft Graph beta API host for this call', false)
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (opts: {
      siteId: string;
      listId: string;
      url?: string;
      filter?: string;
      orderby?: string;
      top?: string;
      allPages?: boolean;
      json?: boolean;
      beta?: boolean;
      token?: string;
      identity?: string;
    }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error || 'Unknown error'}`);
        process.exit(1);
      }
      const api = spApi(opts);
      const topNum = opts.top ? Number.parseInt(opts.top, 10) : undefined;
      if (opts.top && (!Number.isFinite(topNum) || (topNum as number) <= 0)) {
        console.error('Error: --top must be a positive integer');
        process.exit(1);
      }
      const hasPagedQuery =
        Boolean(opts.url?.trim()) ||
        Boolean(opts.filter?.trim()) ||
        Boolean(opts.orderby?.trim()) ||
        opts.top !== undefined;

      if (opts.url?.trim()) {
        const page = await getListItemsPage(auth.token, opts.siteId, opts.listId, {
          nextLink: opts.url.trim(),
          apiBase: api
        });
        if (!page.ok || !page.data) {
          console.error(`Error getting list items: ${page.error?.message || 'Unknown error'}`);
          process.exit(1);
        }
        if (opts.json) {
          console.log(JSON.stringify(page.data, null, 2));
          return;
        }
        printListItemsPageHuman(page.data.value ?? [], page.data['@odata.nextLink']);
        return;
      }

      if (hasPagedQuery) {
        if (opts.allPages) {
          const res = await getAllListItemsPages(auth.token, opts.siteId, opts.listId, {
            filter: opts.filter,
            orderby: opts.orderby,
            top: topNum,
            apiBase: api
          });
          if (!res.ok) {
            console.error(`Error getting list items: ${res.error?.message || 'Unknown error'}`);
            process.exit(1);
          }
          if (opts.json) {
            console.log(JSON.stringify(res.data ?? [], null, 2));
            return;
          }
          if (!res.data || res.data.length === 0) {
            console.log('No items found in this list.');
            return;
          }
          for (const item of res.data) {
            printListItemHuman(item);
          }
          return;
        }
        const page = await getListItemsPage(auth.token, opts.siteId, opts.listId, {
          filter: opts.filter,
          orderby: opts.orderby,
          top: topNum,
          apiBase: api
        });
        if (!page.ok || !page.data) {
          console.error(`Error getting list items: ${page.error?.message || 'Unknown error'}`);
          process.exit(1);
        }
        if (opts.json) {
          console.log(JSON.stringify(page.data, null, 2));
          return;
        }
        printListItemsPageHuman(page.data.value ?? [], page.data['@odata.nextLink']);
        return;
      }

      const res = await getListItems(auth.token, opts.siteId, opts.listId, api);
      if (!res.ok) {
        console.error(`Error getting list items: ${res.error?.message || 'Unknown error'}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(res.data, null, 2));
        return;
      }
      if (!res.data || res.data.length === 0) {
        console.log('No items found in this list.');
        return;
      }
      for (const item of res.data) {
        printListItemHuman(item);
      }
    }
  );

sharepointCommand
  .command('site-permissions')
  .description(
    'List sharing permissions on a site (`GET /sites/{id}/permissions`). Site owners use this to audit sharing; often documented under beta.'
  )
  .requiredOption('--site-id <id>', 'SharePoint Site ID')
  .option('--json', 'Output as JSON')
  .option('--beta', 'Use Microsoft Graph beta API host for this call', false)
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (opts: { siteId: string; json?: boolean; beta?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error || 'Unknown error'}`);
      process.exit(1);
    }
    const res = await getSitePermissions(auth.token, opts.siteId, spApi(opts));
    if (!res.ok) {
      console.error(`Error: ${res.error?.message || 'Unknown error'}`);
      process.exit(1);
    }
    const rows = res.data ?? [];
    if (opts.json) {
      console.log(JSON.stringify({ value: rows }, null, 2));
      return;
    }
    if (rows.length === 0) {
      console.log('No permissions returned.');
      return;
    }
    for (const p of rows) {
      const id = typeof p.id === 'string' ? p.id : '';
      const roles = Array.isArray(p.roles) ? (p.roles as string[]).join(',') : '';
      console.log(`${id}\t${roles}`);
    }
  });

sharepointCommand
  .command('site-permission-update')
  .description(
    'PATCH a site permission (`PATCH /sites/{id}/permissions/{permissionId}`). Body per Graph; site admin/owner scenarios.'
  )
  .requiredOption('--site-id <id>', 'SharePoint Site ID')
  .requiredOption('--permission-id <id>', 'Permission id from site-permissions')
  .requiredOption('--json-file <path>', 'JSON body for PATCH')
  .option('--json', 'Output as JSON')
  .option('--beta', 'Use Microsoft Graph beta API host for this call', false)
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      opts: {
        siteId: string;
        permissionId: string;
        jsonFile: string;
        json?: boolean;
        beta?: boolean;
        token?: string;
        identity?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error || 'Unknown error'}`);
        process.exit(1);
      }
      let raw: string;
      try {
        raw = await readFile(opts.jsonFile.trim(), 'utf-8');
      } catch (e) {
        console.error(`Error: could not read --json-file: ${e instanceof Error ? e.message : String(e)}`);
        process.exit(1);
      }
      let body: Record<string, unknown>;
      try {
        body = JSON.parse(raw) as Record<string, unknown>;
      } catch {
        console.error('Error: --json-file must contain valid JSON');
        process.exit(1);
      }
      const res = await updateSitePermission(auth.token, opts.siteId, opts.permissionId, body, spApi(opts));
      if (!res.ok || !res.data) {
        console.error(`Error: ${res.error?.message || 'Unknown error'}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(res.data, null, 2));
        return;
      }
      console.log('✓ Site permission updated');
    }
  );

sharepointCommand
  .command('site-permission-get')
  .description('GET a single site permission by ID (`GET /sites/{id}/permissions/{permissionId}`).')
  .requiredOption('--site-id <id>', 'SharePoint Site ID')
  .requiredOption('--permission-id <id>', 'Permission id from site-permissions')
  .option('--json', 'Output as JSON')
  .option('--beta', 'Use Microsoft Graph beta API host for this call', false)
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (opts: {
      siteId: string;
      permissionId: string;
      json?: boolean;
      beta?: boolean;
      token?: string;
      identity?: string;
    }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error || 'Unknown error'}`);
        process.exit(1);
      }
      const res = await getSitePermission(auth.token, opts.siteId, opts.permissionId, spApi(opts));
      if (!res.ok || !res.data) {
        console.error(`Error: ${res.error?.message || 'Unknown error'}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(res.data, null, 2));
        return;
      }
      const roles = Array.isArray(res.data.roles) ? (res.data.roles as string[]).join(',') : '';
      console.log(`${res.data.id}\t${roles}`);
    }
  );

sharepointCommand
  .command('site-permission-create')
  .description(
    'POST a new site permission (`POST /sites/{id}/permissions`). Per Graph, this creates an *application* permission only — it cannot grant a new user site permission. Body is a Graph `permission` resource, e.g. { "roles": ["write"], "grantedToIdentities": [{ "application": { "id": "...", "displayName": "..." } }] }.'
  )
  .requiredOption('--site-id <id>', 'SharePoint Site ID')
  .requiredOption('--json-file <path>', 'JSON body for POST (a Graph `permission` resource)')
  .option('--json', 'Output as JSON')
  .option('--beta', 'Use Microsoft Graph beta API host for this call', false)
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      opts: {
        siteId: string;
        jsonFile: string;
        json?: boolean;
        beta?: boolean;
        token?: string;
        identity?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error || 'Unknown error'}`);
        process.exit(1);
      }
      let raw: string;
      try {
        raw = await readFile(opts.jsonFile.trim(), 'utf-8');
      } catch (e) {
        console.error(`Error: could not read --json-file: ${e instanceof Error ? e.message : String(e)}`);
        process.exit(1);
      }
      let body: Record<string, unknown>;
      try {
        body = JSON.parse(raw) as Record<string, unknown>;
      } catch {
        console.error('Error: --json-file must contain valid JSON');
        process.exit(1);
      }
      const res = await createSitePermission(auth.token, opts.siteId, body, spApi(opts));
      if (!res.ok || !res.data) {
        console.error(`Error: ${res.error?.message || 'Unknown error'}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(res.data, null, 2));
        return;
      }
      console.log(`✓ Site permission created${typeof res.data.id === 'string' ? `: ${res.data.id}` : ''}`);
    }
  );

sharepointCommand
  .command('site-permission-delete')
  .description('DELETE a site permission (`DELETE /sites/{id}/permissions/{permissionId}`) — revoke access.')
  .requiredOption('--site-id <id>', 'SharePoint Site ID')
  .requiredOption('--permission-id <id>', 'Permission id from site-permissions')
  .option('--json', 'Output as JSON')
  .option('--beta', 'Use Microsoft Graph beta API host for this call', false)
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      opts: {
        siteId: string;
        permissionId: string;
        json?: boolean;
        beta?: boolean;
        token?: string;
        identity?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error || 'Unknown error'}`);
        process.exit(1);
      }
      const res = await deleteSitePermission(auth.token, opts.siteId, opts.permissionId, spApi(opts));
      if (!res.ok) {
        console.error(`Error: ${res.error?.message || 'Unknown error'}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify({ success: true, siteId: opts.siteId, permissionId: opts.permissionId }, null, 2));
        return;
      }
      console.log(`✓ Removed permission ${opts.permissionId} from site ${opts.siteId}`);
    }
  );

sharepointCommand
  .command('create-item')
  .description(
    'Create an item in a SharePoint list (--fields JSON string or --json-file with field object or { "fields": { … } })'
  )
  .requiredOption('--site-id <id>', 'SharePoint Site ID')
  .requiredOption('--list-id <id>', 'SharePoint List ID')
  .option('--fields <json>', 'JSON string of fields (e.g. \'{"Title": "My Item"}\')')
  .option('--json-file <path>', 'JSON file: either { "fields": { … } } or a flat fields object')
  .option('--json', 'Output as JSON')
  .option('--beta', 'Use Microsoft Graph beta API host for this call', false)
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      opts: {
        siteId: string;
        listId: string;
        fields?: string;
        jsonFile?: string;
        json?: boolean;
        beta?: boolean;
        token?: string;
        identity?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error || 'Unknown error'}`);
        process.exit(1);
      }
      let parsedFields: Record<string, any>;
      try {
        parsedFields = await parseFieldsJsonOrFile({ fields: opts.fields, jsonFile: opts.jsonFile, mode: 'create' });
      } catch (err: any) {
        console.error(`Error: ${err.message}`);
        process.exit(1);
      }
      const res = await createListItem(auth.token, opts.siteId, opts.listId, parsedFields, spApi(opts));
      if (!res.ok) {
        console.error(`Error creating list item: ${res.error?.message || 'Unknown error'}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(res.data, null, 2));
        return;
      }
      console.log(`Successfully created item ${res.data?.id}`);
    }
  );

sharepointCommand
  .command('update-item')
  .description(
    'Update an item in a SharePoint list (--fields or --json-file; file may be { "fields": { … } } or flat fields)'
  )
  .requiredOption('--site-id <id>', 'SharePoint Site ID')
  .requiredOption('--list-id <id>', 'SharePoint List ID')
  .requiredOption('--item-id <id>', 'SharePoint List Item ID')
  .option('--fields <json>', 'JSON string of fields to patch (e.g. \'{"Title": "New Title"}\')')
  .option('--json-file <path>', 'JSON file with fields to PATCH')
  .option('--json', 'Output as JSON')
  .option('--beta', 'Use Microsoft Graph beta API host for this call', false)
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      opts: {
        siteId: string;
        listId: string;
        itemId: string;
        fields?: string;
        jsonFile?: string;
        json?: boolean;
        beta?: boolean;
        token?: string;
        identity?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error || 'Unknown error'}`);
        process.exit(1);
      }
      let parsedFields: Record<string, any>;
      try {
        parsedFields = await parseFieldsJsonOrFile({ fields: opts.fields, jsonFile: opts.jsonFile, mode: 'update' });
      } catch (err: any) {
        console.error(`Error: ${err.message}`);
        process.exit(1);
      }
      const res = await updateListItem(auth.token, opts.siteId, opts.listId, opts.itemId, parsedFields, spApi(opts));
      if (!res.ok) {
        console.error(`Error updating list item: ${res.error?.message || 'Unknown error'}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(res.data, null, 2));
        return;
      }
      console.log(`Successfully updated item ${opts.itemId}`);
    }
  );

sharepointCommand
  .command('get-item')
  .description('Get one SharePoint list item by id ($expand=fields)')
  .requiredOption('--site-id <id>', 'SharePoint Site ID')
  .requiredOption('--list-id <id>', 'SharePoint List ID')
  .requiredOption('--item-id <id>', 'List item ID')
  .option('--json', 'Output as JSON')
  .option('--beta', 'Use Microsoft Graph beta API host for this call', false)
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (opts: {
      siteId: string;
      listId: string;
      itemId: string;
      json?: boolean;
      beta?: boolean;
      token?: string;
      identity?: string;
    }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error || 'Unknown error'}`);
        process.exit(1);
      }
      const res = await getListItem(auth.token, opts.siteId, opts.listId, opts.itemId, spApi(opts));
      if (!res.ok || !res.data) {
        console.error(`Error: ${res.error?.message || 'Unknown error'}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(res.data, null, 2));
        return;
      }
      console.log(`Item ID: ${res.data.id}`);
      if (res.data.fields) {
        for (const [key, val] of Object.entries(res.data.fields)) {
          if (!key.startsWith('@odata')) {
            console.log(`  ${key}: ${val}`);
          }
        }
      }
    }
  );

sharepointCommand
  .command('delete-item')
  .description('Delete a SharePoint list item')
  .requiredOption('--site-id <id>', 'SharePoint Site ID')
  .requiredOption('--list-id <id>', 'SharePoint List ID')
  .requiredOption('--item-id <id>', 'List item ID')
  .option('--json', 'Output as JSON')
  .option('--beta', 'Use Microsoft Graph beta API host for this call', false)
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      opts: {
        siteId: string;
        listId: string;
        itemId: string;
        json?: boolean;
        beta?: boolean;
        token?: string;
        identity?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error || 'Unknown error'}`);
        process.exit(1);
      }
      const res = await deleteListItem(auth.token, opts.siteId, opts.listId, opts.itemId, spApi(opts));
      if (!res.ok) {
        console.error(`Error: ${res.error?.message || 'Unknown error'}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify({ success: true, itemId: opts.itemId }, null, 2));
        return;
      }
      console.log(`Deleted item ${opts.itemId}`);
    }
  );

sharepointCommand
  .command('items-delta')
  .description(
    'One page of SharePoint list items delta (GET …/items/delta?$expand=fields). Use --url for nextLink/deltaLink; optional --state-file (kind: sharePointListItems).'
  )
  .requiredOption('--site-id <id>', 'SharePoint Site ID')
  .requiredOption('--list-id <id>', 'SharePoint List ID')
  .option('--url <url>', 'Full nextLink or deltaLink URL')
  .option('--state-file <path>', 'Read/write JSON delta cursor')
  .option('--json', 'Output raw page JSON')
  .option('--beta', 'Use Microsoft Graph beta API host for this call', false)
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (opts: {
      siteId: string;
      listId: string;
      url?: string;
      stateFile?: string;
      json?: boolean;
      beta?: boolean;
      token?: string;
      identity?: string;
    }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error || 'Unknown error'}`);
        process.exit(1);
      }
      const scope = { sharePointSiteId: opts.siteId.trim(), sharePointListId: opts.listId.trim() };
      const existingState = opts.stateFile ? await readDeltaStateFile(opts.stateFile) : null;
      if (existingState && existingState.kind !== 'sharePointListItems') {
        console.error('Error: state file is not for sharepoint items-delta (kind must be sharePointListItems).');
        process.exit(1);
      }
      try {
        if (existingState) {
          assertDeltaScopeMatchesState(existingState, scope);
        }
      } catch (err) {
        console.error(err instanceof Error ? err.message : err);
        process.exit(1);
      }
      const continueUrl = resolveDeltaContinuationUrl({ explicitNext: opts.url, state: existingState });
      const r = await getListItemsDeltaPage(auth.token, opts.siteId, opts.listId, continueUrl, spApi(opts));
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message || 'Delta failed'}`);
        process.exit(1);
      }
      if (opts.stateFile) {
        const merged = applyDeltaPageToState(existingState, 'sharePointListItems', r.data, scope);
        await writeDeltaStateFile(opts.stateFile, merged);
      }
      if (opts.json) {
        console.log(JSON.stringify(r.data, null, 2));
        return;
      }
      console.log(`Changes: ${r.data.value?.length ?? 0} item(s)`);
      if (r.data['@odata.nextLink']) console.log(`nextLink: ${r.data['@odata.nextLink']}`);
      if (r.data['@odata.deltaLink']) console.log(`deltaLink: ${r.data['@odata.deltaLink']}`);
      if (opts.stateFile) console.log(`state-file: ${opts.stateFile} (updated)`);
    }
  );

sharepointCommand
  .command('followed-sites')
  .description('List sites the signed-in user follows (`GET /me/followedSites`).')
  .option('--json', 'Output as JSON')
  .option('--beta', 'Use Microsoft Graph beta API host for this call', false)
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (opts: { json?: boolean; beta?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error || 'Unknown error'}`);
      process.exit(1);
    }
    const r = await listFollowedSites(auth.token, spApi(opts));
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message || 'followedSites failed'}`);
      process.exit(1);
    }
    const items = r.data.value ?? [];
    if (opts.json) {
      console.log(JSON.stringify({ value: items }, null, 2));
      return;
    }
    if (items.length === 0) {
      console.log('No followed sites.');
      return;
    }
    for (const s of items) {
      const name = s.displayName ?? s.name ?? '(no name)';
      console.log(`${s.id ?? ''}\t${name}`);
      if (s.webUrl) console.log(`  ${s.webUrl}`);
    }
  });

sharepointCommand
  .command('follow <siteId...>')
  .description(
    'Follow one or more SharePoint sites (`POST /me/followedSites/add`). Pass multiple ids to follow many in one call.'
  )
  .option('--json', 'Output as JSON')
  .option('--beta', 'Use Microsoft Graph beta API host for this call', false)
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      siteIds: string[],
      opts: { json?: boolean; beta?: boolean; token?: string; identity?: string },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error || 'Unknown error'}`);
        process.exit(1);
      }
      const ids = siteIds.map((s) => s.trim()).filter(Boolean);
      if (ids.length === 0) {
        console.error('Error: provide at least one site id');
        process.exit(1);
      }
      const r = await followSites(auth.token, ids, spApi(opts));
      if (!r.ok) {
        console.error(`Error: ${r.error?.message || 'follow failed'}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(r.data ?? { value: [] }, null, 2));
        return;
      }
      const items = r.data?.value ?? [];
      console.log(`✓ Following ${items.length} site(s)`);
      for (const s of items) {
        const name = s.displayName ?? s.name ?? '(no name)';
        console.log(`  ${s.id ?? ''}\t${name}`);
      }
    }
  );

sharepointCommand
  .command('unfollow <siteId...>')
  .description('Unfollow one or more SharePoint sites (`POST /me/followedSites/remove`).')
  .option('--beta', 'Use Microsoft Graph beta API host for this call', false)
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (siteIds: string[], opts: { beta?: boolean; token?: string; identity?: string }, cmd: any) => {
    checkReadOnly(cmd);
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error || 'Unknown error'}`);
      process.exit(1);
    }
    const ids = siteIds.map((s) => s.trim()).filter(Boolean);
    if (ids.length === 0) {
      console.error('Error: provide at least one site id');
      process.exit(1);
    }
    const r = await unfollowSites(auth.token, ids, spApi(opts));
    if (!r.ok) {
      console.error(`Error: ${r.error?.message || 'unfollow failed'}`);
      process.exit(1);
    }
    console.log(`✓ Unfollowed ${ids.length} site(s)`);
  });
