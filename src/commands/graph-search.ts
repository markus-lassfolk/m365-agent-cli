import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import { microsoftSearchQuery } from '../lib/graph-microsoft-search.js';

const DEFAULT_ENTITY_TYPES = ['message', 'event', 'driveItem', 'listItem', 'person'];

function summarizeResource(r: Record<string, unknown> | undefined): string {
  if (!r) return '(no resource)';
  const type = (r['@odata.type'] as string | undefined)?.replace(/^#microsoft\.graph\./, '') ?? 'item';
  const subject = r.subject as string | undefined;
  const name = r.name as string | undefined;
  const title = r.title as string | undefined;
  const displayName = r.displayName as string | undefined;
  const line = subject || name || title || displayName || (r.id as string) || '';
  return line ? `[${type}] ${line}` : `[${type}]`;
}

export const graphSearchCommand = new Command('graph-search').description(
  'Microsoft Graph Search API (POST /search/query); entity-specific delegated scopes (Mail, Files, Calendars, etc.) — see Graph docs and docs/GRAPH_SCOPES.md'
);

graphSearchCommand
  .argument('<query>', 'Search query string (KQL-style per Graph docs)')
  .option(
    '-t, --types <list>',
    `Comma-separated entity types (default: ${DEFAULT_ENTITY_TYPES.join(',')})`
  )
  .option('--from <n>', 'Result offset', '0')
  .option('--size <n>', 'Page size (1–1000)', '25')
  .option('--json', 'Output raw JSON response')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      query: string,
      opts: {
        types?: string;
        from?: string;
        size?: string;
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
      const parsedTypes = opts.types
        ?.split(',')
        .map((s) => s.trim())
        .filter(Boolean);
      const entityTypes =
        parsedTypes && parsedTypes.length > 0 ? parsedTypes : [...DEFAULT_ENTITY_TYPES];
      const from = Math.max(0, parseInt(opts.from ?? '0', 10) || 0);
      const size = Math.min(1000, Math.max(1, parseInt(opts.size ?? '25', 10) || 25));

      const r = await microsoftSearchQuery(auth.token, {
        entityTypes,
        queryString: query,
        from,
        size
      });
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(r.data, null, 2));
        return;
      }

      const blocks = r.data.value ?? [];
      if (blocks.length === 0) {
        console.log('No result blocks (empty value array).');
        return;
      }
      for (const block of blocks) {
        const terms = block.searchTerms?.join(', ') ?? query;
        console.log(`Search terms: ${terms}`);
        const containers = block.hitsContainers ?? [];
        if (containers.length === 0) {
          console.log('  (no hitsContainers)');
          continue;
        }
        for (const c of containers) {
          const hits = c.hits ?? [];
          const total = c.total ?? hits.length;
          console.log(`  Hits (showing ${hits.length}, total reported: ${total})`);
          for (const h of hits) {
            const line = summarizeResource(h.resource);
            const rank = h.rank != null ? `#${h.rank} ` : '';
            console.log(`    ${rank}${line}`);
            if (h.summary?.trim()) {
              const oneLine = h.summary.replace(/\s+/g, ' ').trim().slice(0, 200);
              if (oneLine) console.log(`      ${oneLine}`);
            }
          }
          if (c.moreResultsAvailable) console.log('    … more results available (increase --size or paginate with --from)');
        }
      }
    }
  );
