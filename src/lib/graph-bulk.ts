/**
 * Bulk Graph mutations: build one JSON-batch sub-request per target id, fire them through
 * `graphBatchAll` (auto-chunked into `<=20`-request `/$batch` POSTs — see `graph-advanced-client.ts`),
 * and normalize the per-id outcome. Avoids N sequential single-item Graph calls for CLI commands
 * that operate on an ID list (`--ids <csv>` / `--json-file <path>`).
 */

import { graphBatchAll } from './graph-advanced-client.js';
import type { GraphResponse } from './graph-client.js';
import { readJsonFileOrExit } from './read-json-file.js';

export interface BulkSubRequestSpec {
  id: string;
  method: string;
  url: string;
  headers?: Record<string, string>;
  body?: unknown;
}

export interface BulkMutationOutcome {
  id: string;
  ok: boolean;
  status?: number;
  error?: string;
}

function extractBatchErrorMessage(body: unknown): string | undefined {
  if (!body || typeof body !== 'object') return undefined;
  return (body as { error?: { message?: string } }).error?.message;
}

/** Runs one batched Graph mutation per target id, matching each `/$batch` sub-response back to its request by `id`. */
export async function applyBulkGraphRequests(
  token: string,
  requests: BulkSubRequestSpec[],
  beta?: boolean,
  auth?: { identity?: string; pinAccessToken?: boolean }
): Promise<GraphResponse<BulkMutationOutcome[]>> {
  const r = await graphBatchAll(token, requests as unknown as Array<Record<string, unknown>>, beta, auth);
  if (!r.ok) {
    return { ok: false, error: r.error };
  }
  const byId = new Map((r.data?.responses ?? []).map((resp) => [resp.id, resp]));
  const outcomes: BulkMutationOutcome[] = requests.map((req) => {
    const resp = byId.get(req.id);
    if (!resp) {
      return { id: req.id, ok: false, error: 'No response returned for this request (unexpected)' };
    }
    const ok = resp.status >= 200 && resp.status < 300;
    return {
      id: req.id,
      ok,
      status: resp.status,
      ...(ok ? {} : { error: extractBatchErrorMessage(resp.body) || `HTTP ${resp.status}` })
    };
  });
  return { ok: true, data: outcomes };
}

/**
 * Parses a `--ids <csv>` / `--json-file <path>` pair (the convention used by `presence bulk`) into
 * a trimmed, blank-filtered id list. Exits the process with a clean error message — never throws —
 * on missing input, an unreadable/invalid `--json-file`, or an empty resulting list, so every
 * caller gets the same user-facing behavior.
 */
export async function parseBulkIdListOrExit(opts: { ids?: string; jsonFile?: string }): Promise<string[]> {
  let idList: string[];
  if (opts.jsonFile?.trim()) {
    const raw = await readJsonFileOrExit<unknown>(opts.jsonFile, '--json-file');
    idList = Array.isArray(raw)
      ? raw
          .map(String)
          .map((s) => s.trim())
          .filter(Boolean)
      : [];
  } else if (opts.ids?.trim()) {
    idList = opts.ids
      .split(',')
      .map((s) => s.trim())
      .filter(Boolean);
  } else {
    console.error('Error: provide --ids or --json-file with a JSON array of ids');
    process.exit(1);
  }
  if (idList.length === 0) {
    console.error('Error: no ids provided (list was empty after removing blanks)');
    process.exit(1);
  }
  return idList;
}

/** Prints a per-id bulk outcome summary consistently across bulk commands. */
export function printBulkOutcomeSummary(outcomes: BulkMutationOutcome[], json: boolean | undefined): void {
  const succeeded = outcomes.filter((o) => o.ok).length;
  const failed = outcomes.length - succeeded;
  if (json) {
    console.log(JSON.stringify({ succeeded, failed, results: outcomes }, null, 2));
    return;
  }
  for (const o of outcomes) {
    console.log(o.ok ? `✓ ${o.id}` : `✗ ${o.id}: ${o.error}`);
  }
  console.log(`\n${succeeded} succeeded, ${failed} failed (${outcomes.length} total)`);
}
