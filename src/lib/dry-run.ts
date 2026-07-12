/**
 * `--dry-run` support: preview the exact resolved request a mutating command would send,
 * without sending it. Implemented at the transport layer (graph-client.ts / ews-client.ts)
 * so it works uniformly for every command — including multi-step flows — without each of the
 * ~60 command files needing to hand-assemble a preview of its own request.
 *
 * Activation is a process-wide signal (`M365_DRY_RUN=1`), set from the root `--dry-run` flag by
 * a Commander `preAction` hook (see m365-program.ts) before the command action runs. The
 * transport layer has no access to the parsed Command instance, so an env var is the only way
 * to carry the flag down to `callGraphAt` / `callEws`.
 *
 * For a multi-step flow (e.g. createDraft → addAttachment → send), only the FIRST mutating
 * request is shown: halting immediately (via `process.exit`) at the first write is required to
 * make the guarantee "dry-run never causes a real side effect" — printing all N would-be
 * requests would mean pretending step 1 succeeded, which we don't know without truly sending it.
 */

const TRUE_VALUES = new Set(['1', 'true']);

/** True when `--dry-run` (synced to this env var by the CLI) is active for this process. */
export function isDryRunActive(): boolean {
  return TRUE_VALUES.has((process.env.M365_DRY_RUN || '').toLowerCase());
}

export interface DryRunPreview {
  backend: 'graph' | 'ews';
  [key: string]: unknown;
}

/**
 * Prints the dry-run preview as the sole line of stdout and exits 0 — nothing was sent.
 * Uses `process.exit` (not a thrown error) so the preview can never be caught, re-wrapped into a
 * normal error response by a client wrapper's `catch`, and printed a second time as JSON on the
 * same stdout stream (which would break `--json` parsers).
 */
export function haltForDryRun(preview: DryRunPreview): never {
  console.log(JSON.stringify({ dryRun: true, ...preview }, null, 2));
  return process.exit(0);
}

/** Best-effort human/JSON-friendly rendering of a fetch body for the dry-run preview. */
export function previewableBody(body: unknown): unknown {
  if (body === null || body === undefined) return undefined;
  if (typeof body === 'string') {
    try {
      return JSON.parse(body);
    } catch {
      return body;
    }
  }
  if (body instanceof Uint8Array || body instanceof ArrayBuffer) {
    const len = body instanceof Uint8Array ? body.byteLength : body.byteLength;
    return `<binary body, ${len} bytes>`;
  }
  if (typeof FormData !== 'undefined' && body instanceof FormData) {
    return '<FormData body>';
  }
  return '<stream body>';
}
