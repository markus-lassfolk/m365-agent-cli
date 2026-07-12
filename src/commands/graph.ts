import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { Command } from 'commander';
import {
  GRAPH_BATCH_MAX_REQUESTS,
  type GraphBatchRequestBody,
  graphBatchAll,
  graphInvoke,
  parseGraphInvokeHeaders
} from '../lib/graph-advanced-client.js';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import { checkReadOnly } from '../lib/utils.js';

/** Read + parse a JSON file, exiting with a clean message (not a raw stack trace) on failure. */
async function readJsonFileOrExit(path: string, label: string): Promise<unknown> {
  let raw: string;
  try {
    raw = await readFile(resolve(process.cwd(), path.trim()), 'utf8');
  } catch (err) {
    console.error(`Error: could not read ${label}: ${err instanceof Error ? err.message : String(err)}`);
    process.exit(1);
  }
  try {
    return JSON.parse(raw);
  } catch (err) {
    console.error(`Error: ${label} must contain valid JSON: ${err instanceof Error ? err.message : String(err)}`);
    process.exit(1);
  }
}

function batchHasMutations(body: GraphBatchRequestBody): boolean {
  if (!Array.isArray(body?.requests)) return false;
  for (const req of body.requests) {
    const m = String((req as { method?: string }).method || 'GET').toUpperCase();
    if (m !== 'GET' && m !== 'HEAD') return true;
  }
  return false;
}

export const graphCommand = new Command('graph').description(
  'Advanced Microsoft Graph: raw REST invoke and JSON batch ($batch). Paths are relative to GRAPH_BASE_URL (v1.0) or beta.'
);

graphCommand
  .command('invoke')
  .description(
    'Call Graph with a relative path (e.g. /me/messages?$top=5); JSON response only. Advanced OData ($search, some $filter/$count) may need headers such as ConsistencyLevel: eventual — use --header "ConsistencyLevel: eventual" or a higher-level CLI command.'
  )
  .argument('<path>', 'Path starting with / (under v1.0 or beta root)')
  .option('-X, --method <method>', 'HTTP method', 'GET')
  .option('-d, --data <json>', 'JSON request body (for POST/PATCH/PUT)')
  .option('--body-file <path>', 'Read JSON body from file (overrides --data)')
  .option('--beta', 'Use GRAPH_BETA_URL instead of v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .option(
    '-H, --header <nameValue>',
    'Extra HTTP header ("Name: value", first colon separates name from value). Repeatable, e.g. -H "ConsistencyLevel: eventual"',
    (val: string, prev: string[]) => {
      const acc = prev ?? [];
      acc.push(val);
      return acc;
    },
    [] as string[]
  )
  .action(
    async (
      pathArg: string,
      opts: {
        method: string;
        data?: string;
        bodyFile?: string;
        beta?: boolean;
        token?: string;
        identity?: string;
        header?: string[];
      },
      cmd
    ) => {
      const method = (opts.method || 'GET').toUpperCase();
      if (method !== 'GET' && method !== 'HEAD') {
        checkReadOnly(cmd);
      }

      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }

      let body: unknown | undefined;
      if (opts.bodyFile) {
        body = await readJsonFileOrExit(opts.bodyFile, '--body-file');
      } else if (opts.data) {
        try {
          body = JSON.parse(opts.data) as unknown;
        } catch (err) {
          console.error(`Error: --data must be valid JSON: ${err instanceof Error ? err.message : String(err)}`);
          process.exit(1);
        }
      }

      let extraHeaders: Record<string, string> | undefined;
      try {
        const lines = opts.header && opts.header.length > 0 ? opts.header : [];
        extraHeaders = lines.length > 0 ? parseGraphInvokeHeaders(lines) : undefined;
      } catch (e) {
        console.error(e instanceof Error ? e.message : String(e));
        process.exit(1);
      }

      const r = await graphInvoke(auth.token, {
        method,
        path: pathArg,
        body,
        beta: opts.beta,
        expectJson: true,
        extraHeaders,
        identity: opts.identity,
        pinAccessToken: !!opts.token
      });

      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        if (r.error?.requestId) {
          console.error(`request-id: ${r.error.requestId}`);
        }
        process.exit(1);
      }
      console.log(JSON.stringify(r.data, null, 2));
    }
  );

graphCommand
  .command('batch')
  .description(
    `POST JSON batch body to /$batch. Any number of sub-requests is accepted: requests beyond the API's ${GRAPH_BATCH_MAX_REQUESTS}-per-call cap are transparently split into multiple /$batch POSTs (sent sequentially) and the "responses" arrays are merged back into one, in request order. Sub-requests with a "dependsOn" chain must land within the same ${GRAPH_BATCH_MAX_REQUESTS}-request chunk. See https://learn.microsoft.com/en-us/graph/json-batching`
  )
  .requiredOption('-f, --file <path>', 'JSON file: { "requests": [ { "id", "method", "url", ... }, ... ] }')
  .option('--beta', 'Use GRAPH_BETA_URL')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { file: string; beta?: boolean; token?: string; identity?: string }, cmd) => {
    const body = (await readJsonFileOrExit(opts.file, '--file')) as GraphBatchRequestBody;
    if (!Array.isArray(body?.requests)) {
      console.error('Error: --file must be a JSON object with a "requests" array (see --help).');
      process.exit(1);
    }
    if (batchHasMutations(body)) {
      checkReadOnly(cmd);
    }

    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }

    const r = await graphBatchAll(auth.token, body.requests, opts.beta, {
      identity: opts.identity,
      pinAccessToken: !!opts.token
    });
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      if (r.error?.requestId) {
        console.error(`request-id: ${r.error.requestId}`);
      }
      process.exit(1);
    }
    console.log(JSON.stringify(r.data, null, 2));
  });
