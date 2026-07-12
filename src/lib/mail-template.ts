/**
 * Reusable mail/draft templates: a text/HTML file with `{{variable}}` (optionally
 * `{{variable|default text}}`) placeholders, filled in from repeatable `--var name=value` CLI
 * flags. Used by `send` and `drafts --create`/`--edit` as an alternative to `--body`.
 */

const PLACEHOLDER_RE = /\{\{\s*([a-zA-Z_][a-zA-Z0-9_]*)\s*(?:\|([^}]*))?\}\}/g;
/** Catches `{{...}}`-shaped tokens PLACEHOLDER_RE's strict identifier grammar doesn't match
 *  (hyphenated names, leading digits, empty braces, ...) so they don't pass through unsubstituted. */
const MALFORMED_PLACEHOLDER_RE = /\{\{[^}]*\}\}/g;

export class MailTemplateError extends Error {}

/** Parses repeatable `--var name=value` entries into a lookup object; first `=` splits name from value. */
export function parseTemplateVars(pairs: string[]): Record<string, string> {
  const vars: Record<string, string> = {};
  for (const pair of pairs) {
    const idx = pair.indexOf('=');
    if (idx === -1) {
      throw new MailTemplateError(`Invalid --var (expected "name=value"): ${pair}`);
    }
    const name = pair.slice(0, idx).trim();
    const value = pair.slice(idx + 1);
    if (!name) {
      throw new MailTemplateError(`Invalid --var (empty name): ${pair}`);
    }
    vars[name] = value;
  }
  return vars;
}

/**
 * Replaces every `{{name}}` / `{{name|default}}` placeholder in `source` with `vars[name]`, or the
 * placeholder's own default text when `vars[name]` is unset. Throws {@link MailTemplateError}
 * listing any placeholder left with neither a supplied value nor a default — better a clear error
 * than silently mailing a recipient a literal `{{name}}`.
 */
export function renderMailTemplate(source: string, vars: Record<string, string>): string {
  const unresolved = new Set<string>();
  const rendered = source.replace(PLACEHOLDER_RE, (_match, name: string, fallback: string | undefined) => {
    if (Object.hasOwn(vars, name)) return vars[name];
    if (fallback !== undefined) return fallback;
    unresolved.add(name);
    return _match;
  });
  if (unresolved.size > 0) {
    throw new MailTemplateError(
      `Template has unresolved placeholder(s) with no --var and no default: ${[...unresolved].join(', ')}`
    );
  }
  const malformed = [...new Set([...rendered.matchAll(MALFORMED_PLACEHOLDER_RE)].map((m) => m[0]))];
  if (malformed.length > 0) {
    throw new MailTemplateError(
      `Template has malformed placeholder(s) — expected {{name}} or {{name|default}}: ${malformed.join(', ')}`
    );
  }
  return rendered;
}
