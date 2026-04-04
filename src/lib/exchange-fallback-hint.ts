/**
 * User-visible hints when `M365_EXCHANGE_BACKEND=auto` tries Microsoft Graph first
 * and continues on Exchange Web Services (EWS).
 */

const SCOPES_DOC = 'docs/GRAPH_SCOPES.md';

/** Heuristic: Graph or auth layer reported something that often means scopes / consent / RBAC. */
function graphErrorLooksPermissionRelated(message: string | undefined): boolean {
  if (!message) return false;
  const m = message.toLowerCase();
  return (
    /\b403\b/.test(m) ||
    m.includes('access denied') ||
    m.includes('access is denied') ||
    m.includes('forbidden') ||
    m.includes('insufficient privileges') ||
    m.includes('not allowed') ||
    m.includes('authorization_requestdenied') ||
    m.includes('required permissions')
  );
}

export type AutoFallbackReason = 'auth' | 'api';

/**
 * Emits stderr guidance when auto mode falls back from Graph to EWS.
 * Skipped for `--json` so machine-readable stdout stays clean (warnings still go to real stderr in production).
 */
export function warnAutoGraphToEwsFallback(
  commandLabel: string,
  opts: {
    json?: boolean;
    verbose?: boolean;
    graphError?: string;
    /** Auth failures often warrant the same scope/consent hint as permission-like API errors. */
    reason?: AutoFallbackReason;
  }
): void {
  if (opts.json) return;

  const err = opts.graphError?.trim();
  const permissionish =
    opts.reason === 'auth' || (err !== undefined && err.length > 0 && graphErrorLooksPermissionRelated(err));

  console.warn(
    `[${commandLabel}] Microsoft Graph did not satisfy this request; continuing with Exchange Web Services (EWS). (M365_EXCHANGE_BACKEND=auto)`
  );

  if (permissionish) {
    console.warn(
      `[${commandLabel}] Hint: missing or narrow Microsoft Graph delegated permissions is a common cause. See ${SCOPES_DOC} and run \`m365-agent-cli login\` after your admin adds scopes. Use \`M365_EXCHANGE_BACKEND=graph\` to surface Graph-only errors.`
    );
  } else if (err) {
    console.warn(`[${commandLabel}] Graph message: ${err}`);
  }

  if (opts.verbose && err) {
    console.warn(`[${commandLabel}] Verbose — Graph detail: ${err}`);
  }
}
