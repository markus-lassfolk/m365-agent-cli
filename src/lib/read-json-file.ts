import { readFile } from 'node:fs/promises';

/**
 * Read + parse a JSON file, exiting with a clean message (not a raw stack trace or a silent
 * exit) on failure. Keeps routine user input errors (missing/malformed `--json-file`) out of
 * exception tracking, which reports uncaught throws from the CLI entrypoint.
 */
export async function readJsonFileOrExit<T = Record<string, unknown>>(path: string, label: string): Promise<T> {
  let raw: string;
  try {
    raw = await readFile(path.trim(), 'utf-8');
  } catch (err) {
    console.error(`Error: could not read ${label}: ${err instanceof Error ? err.message : String(err)}`);
    process.exit(1);
  }
  try {
    return JSON.parse(raw) as T;
  } catch (err) {
    console.error(`Error: ${label} must contain valid JSON: ${err instanceof Error ? err.message : String(err)}`);
    process.exit(1);
  }
}
