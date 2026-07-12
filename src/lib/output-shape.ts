/**
 * Client-side output shaping for `--json` list output: `--fields` projects each row down to a
 * comma-separated set of dot-paths, and `--ndjson` prints one JSON object per line instead of a
 * single pretty-printed array, so an agent can stream-process large lists without buffering.
 * Both are opt-in per command (there is no single choke point every list command funnels
 * through, unlike the Graph transport layer), so each command wires this in explicitly.
 */

/** Parses `--fields "subject,from.emailAddress.address"` into `["subject", "from.emailAddress.address"]`, or `undefined` if unset/empty. */
export function parseFieldsOption(raw: string | undefined): string[] | undefined {
  if (!raw) return undefined;
  const fields = raw
    .split(',')
    .map((s) => s.trim())
    .filter(Boolean);
  return fields.length > 0 ? fields : undefined;
}

function getByPath(obj: unknown, segments: string[]): { found: boolean; value?: unknown } {
  let cur: unknown = obj;
  for (const seg of segments) {
    if (cur === null || typeof cur !== 'object' || !(seg in (cur as Record<string, unknown>))) {
      return { found: false };
    }
    cur = (cur as Record<string, unknown>)[seg];
  }
  return { found: true, value: cur };
}

function setByPath(target: Record<string, unknown>, segments: string[], value: unknown): void {
  let cur = target;
  for (let i = 0; i < segments.length - 1; i++) {
    const seg = segments[i];
    const existing = cur[seg];
    if (typeof existing !== 'object' || existing === null || Array.isArray(existing)) {
      cur[seg] = {};
    }
    cur = cur[seg] as Record<string, unknown>;
  }
  cur[segments[segments.length - 1]] = value;
}

/**
 * Projects `value` down to the given dot-paths (e.g. `"from.emailAddress.address"`), preserving
 * nesting shape. Paths that don't exist on `value` are silently omitted (not an error) — Graph
 * and EWS payloads are not uniform, so a field present on some rows and absent on others is
 * normal. Non-object `value` (e.g. `null`, a primitive) is returned unchanged.
 */
export function projectFields(value: unknown, fields: string[]): unknown {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) return value;
  const out: Record<string, unknown> = {};
  for (const path of fields) {
    const segments = path
      .split('.')
      .map((s) => s.trim())
      .filter(Boolean);
    if (segments.length === 0) continue;
    const got = getByPath(value, segments);
    if (got.found) setByPath(out, segments, got.value);
  }
  return out;
}

/** Applies {@link projectFields} to every row when `fields` is set; otherwise returns `rows` unchanged. */
export function shapeRows(rows: unknown[], fields?: string[]): unknown[] {
  if (!fields || fields.length === 0) return rows;
  return rows.map((row) => projectFields(row, fields));
}

/** One compact JSON object per line (no pretty-printing — NDJSON is meant to be streamed/parsed line-by-line). */
export function formatNdjson(rows: unknown[]): string {
  return rows.map((row) => JSON.stringify(row)).join('\n');
}
