/**
 * Structured `--json` error envelope: `{ error: { message, code, status, retriable, requestId } }`
 * instead of a bare `{ error: "some string" }`. Mirrors the shape Microsoft Graph itself uses for
 * error responses (`{ "error": { "code": ..., "message": ... } }`), so agents parsing `--json`
 * output can branch on `error.code`/`error.retriable` instead of regex-matching prose.
 *
 * `GraphError` (graph-client.ts) and `OwaError` (ews-client.ts) already carry this data — the gap
 * was command code extracting only `.message` before printing `--json` output, discarding
 * `code`/`status`/`requestId`. `toJsonError` accepts whatever a call site already has (the full
 * error object, a bare string, an `Error`, or `undefined`) and normalizes it into one shape, so it
 * is a safe drop-in wrap at any existing `{ error: EXPR }` site regardless of what EXPR is.
 */

export interface JsonErrorEnvelope {
  message: string;
  code?: string;
  status?: number;
  retriable?: boolean;
  requestId?: string;
}

const RETRIABLE_CODES = new Set([
  'tooManyRequests',
  'TooManyRequests',
  'serviceNotAvailable',
  'ServiceUnavailable',
  'activityLimitReached'
]);
const RETRIABLE_STATUS = new Set([429, 502, 503, 504]);

function isRetriable(code: string | undefined, status: number | undefined): boolean | undefined {
  const retriable =
    (code !== undefined && RETRIABLE_CODES.has(code)) || (status !== undefined && RETRIABLE_STATUS.has(status));
  return retriable || undefined;
}

/**
 * Normalizes any error-ish value into the structured envelope. Handles:
 * - `GraphError` / `OwaError`-shaped objects (`{ message, code?, status?, requestId? }`)
 * - a plain string (already-reduced `err.message` style)
 * - an `Error` instance
 * - `null`/`undefined` (falls back to `fallbackMessage`)
 */
export function toJsonError(input: unknown, fallbackMessage = 'Request failed'): JsonErrorEnvelope {
  if (input === null || input === undefined) {
    return { message: fallbackMessage };
  }

  if (typeof input === 'string') {
    return { message: input.trim() || fallbackMessage };
  }

  if (input instanceof Error) {
    return { message: input.message || fallbackMessage };
  }

  if (typeof input === 'object') {
    const obj = input as Record<string, unknown>;
    const message =
      typeof obj.message === 'string' && obj.message.trim()
        ? obj.message
        : typeof obj.error === 'string' && obj.error.trim()
          ? obj.error
          : fallbackMessage;
    const code = typeof obj.code === 'string' ? obj.code : undefined;
    const status = typeof obj.status === 'number' ? obj.status : undefined;
    const requestId = typeof obj.requestId === 'string' ? obj.requestId : undefined;
    const retriable = isRetriable(code, status);
    return {
      message,
      ...(code !== undefined ? { code } : {}),
      ...(status !== undefined ? { status } : {}),
      ...(retriable !== undefined ? { retriable } : {}),
      ...(requestId !== undefined ? { requestId } : {})
    };
  }

  return { message: String(input) || fallbackMessage };
}
