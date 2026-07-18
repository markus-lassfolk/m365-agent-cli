/**
 * Deep redaction for anything that might end up in a shareable diagnostic bundle (`doctor
 * --redacted-bundle`, issue #246). Defense in depth: `doctor-bundle.ts` only ever *builds* safe,
 * non-secret metadata by construction (presence/size/mtime, not raw file contents), but every
 * object still passes through {@link deepRedact} before being written or printed, so a future
 * field added by mistake gets caught instead of silently leaking.
 */

const REDACTED = '[REDACTED]';

/** Field-name patterns that always redact their value, whatever it looks like. */
const SECRET_KEY_PATTERN =
  /(token|secret|password|passwd|pwd|refresh|access[-_]?key|api[-_]?key|apikey|authorization|auth[-_]?code|cookie|client[-_]?secret|private[-_]?key|credential)/i;

/** Three dot-separated base64url segments — a JWT access/id token shape. */
const JWT_LIKE_RE = /^[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]*$/;

/**
 * Long, whitespace-free, opaque-looking strings (refresh tokens, auth codes, API keys, …).
 * Deliberately excludes `/` and `+` (so POSIX file paths and URLs never match — those are common,
 * legitimate values in diagnostic metadata) and requires base64url's charset only.
 */
const LONG_OPAQUE_RE = /^[A-Za-z0-9_\-.=]{24,}$/;

export function isSecretKeyName(key: string): boolean {
  return SECRET_KEY_PATTERN.test(key);
}

/** High-entropy heuristic: real opaque tokens mix case and digits; filenames/identifiers usually don't. */
function looksHighEntropy(v: string): boolean {
  return /[a-z]/.test(v) && /[A-Z]/.test(v) && /[0-9]/.test(v);
}

/** True for values that look like a token/secret regardless of the field name holding them. */
export function looksLikeSecretValue(value: unknown): boolean {
  if (typeof value !== 'string') return false;
  const v = value.trim();
  if (!v) return false;
  if (JWT_LIKE_RE.test(v)) return true;
  if (v.length >= 24 && !/\s/.test(v) && LONG_OPAQUE_RE.test(v) && looksHighEntropy(v)) return true;
  return false;
}

/**
 * Recursively redact any object/array: a key matching a known secret-field pattern, or any string
 * value that looks token/secret-shaped, becomes `"[REDACTED]"`. Everything else (numbers,
 * booleans, dates, short/safe strings) passes through unchanged.
 */
export function deepRedact<T>(value: T, options?: { maxDepth?: number }): T {
  const maxDepth = options?.maxDepth ?? 10;

  function walk(v: unknown, depth: number): unknown {
    if (depth > maxDepth) return '[TRUNCATED]';
    if (Array.isArray(v)) return v.map((item) => walk(item, depth + 1));
    if (v && typeof v === 'object') {
      const out: Record<string, unknown> = {};
      for (const [k, val] of Object.entries(v as Record<string, unknown>)) {
        // Key-name matching only redacts STRING leaves — a boolean/number field whose name merely
        // contains "secret"/"token" (e.g. a `secretsPrinted: false` marker) is not itself secret
        // material. Objects/arrays under such a key are still walked so any real nested secret
        // leaf gets caught by its own key or value shape.
        if (typeof val === 'string' && (isSecretKeyName(k) || looksLikeSecretValue(val))) {
          out[k] = REDACTED;
        } else {
          out[k] = walk(val, depth + 1);
        }
      }
      return out;
    }
    if (looksLikeSecretValue(v)) return REDACTED;
    return v;
  }

  return walk(value, 0) as T;
}
