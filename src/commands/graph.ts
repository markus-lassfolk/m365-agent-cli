import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import {
  graphInvoke,
  graphPostBatch,
  type GraphBatchRequestBody
} from '../lib/graph-advanced-client.js';
import { checkReadOnly } from '../lib/utils.js';

function batchHasMutations(body: GraphBatchRequestBody): boolean {
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
  .description('Call Graph with a relative path (e.g. /me/messages?$top=5); JSON response only')
  .argument('<path>', 'Path starting with / (under v1.0 or beta root)')
  .option('-X, --method <method>', 'HTTP method', 'GET')
  .option('-d, --data <json>', 'JSON request body (for POST/PATCH/PUT)')
  .option('--body-file <path>', 'Read JSON body from file (overrides --data)')
  .option('--beta', 'Use GRAPH_BETA_URL instead of v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
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
        const raw = await readFile(resolve(process.cwd(), opts.bodyFile.trim()), 'utf8');
        body = JSON.parse(raw) as unknown;
      } else if (opts.data) {
        body = JSON.parse(opts.data) as unknown;
      }

      const r = await graphInvoke(auth.token, {
        method,
        path: pathArg,
        body,
        beta: opts.beta,
        expectJson: true
      });

      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data, null, 2));
    }
  );

graphCommand
  .command('batch')
  .description('POST JSON batch body to /$batch (see Graph JSON batching docs)')
  .requiredOption('-f, --file <path>', 'JSON file: { "requests": [ { "id", "method", "url", ... }, ... ] }')
  .option('--beta', 'Use GRAPH_BETA_URL')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      opts: { file: string; beta?: boolean; token?: string; identity?: string },
      cmd
    ) => {
      const raw = await readFile(resolve(process.cwd(), opts.file.trim()), 'utf8');
      const body = JSON.parse(raw) as GraphBatchRequestBody;
      if (batchHasMutations(body)) {
        checkReadOnly(cmd);
      }

      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }

      const r = await graphPostBatch(auth.token, body, opts.beta);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data, null, 2));
    }
  );
