import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import type { Command } from 'commander';

const __dirname = dirname(fileURLToPath(import.meta.url));

/** Long or short flag tokens we never treat as errors (Commander built-ins + root tip). */
const IMPLICIT_FLAGS = new Set(['--help', '-h', '-V', '--version']);

function forEachCommandDepthFirst(cmd: Command, fn: (c: Command) => void): void {
  fn(cmd);
  for (const sub of cmd.commands) {
    forEachCommandDepthFirst(sub, fn);
  }
}

function resetCommanderHooks(cmd: Command): void {
  (cmd as unknown as { _exitCallback: null })._exitCallback = null;
  cmd.configureOutput({
    writeOut: (str: string) => process.stdout.write(str),
    writeErr: (str: string) => process.stderr.write(str)
  });
  for (const sub of cmd.commands) {
    resetCommanderHooks(sub);
  }
}

function captureOutputHelp(cmd: Command): string {
  const chunks: string[] = [];
  cmd.configureOutput({
    writeOut: (s: string) => {
      chunks.push(s);
    },
    writeErr: (s: string) => {
      chunks.push(s);
    }
  });
  cmd.exitOverride();
  cmd.outputHelp({
    write: (s: string) => {
      chunks.push(s);
    }
  } as unknown as Parameters<Command['outputHelp']>[0]);
  return chunks.join('');
}

function mergedOptionFlags(cmd: Command): Set<string> {
  const s = new Set<string>(IMPLICIT_FLAGS);
  let c: Command | null = cmd;
  while (c) {
    for (const opt of c.options) {
      if (opt.hidden) continue;
      if (opt.long) s.add(opt.long);
      if (opt.short) s.add(opt.short);
    }
    c = c.parent ?? null;
  }
  return s;
}

function findNamedSubcommand(parent: Command, token: string): Command | undefined {
  return parent.commands.find((c) => {
    const aliases = c.aliases();
    return c.name() === token || (aliases.length > 0 && aliases.includes(token));
  });
}

function isPlaceholderToken(t: string): boolean {
  return t.startsWith('<') && t.endsWith('>') && t.length > 2;
}

/**
 * Walk from program: subcommand names until a flag or positional; consume positionals
 * for the resolved leaf (Graph paths, ids, etc.).
 */
function getDefaultSubcommand(parent: Command): Command | undefined {
  const defName = (parent as unknown as { _defaultCommandName?: string | null })._defaultCommandName;
  if (!defName) return undefined;
  return parent.commands.find((c) => c.name() === defName);
}

/**
 * Walk from program through explicit subcommands, then positionals; if the walk ends on a parent
 * that has a default subcommand and the next token is a flag, treat flags as belonging to the default
 * (same as Commander: `calendar --mailbox` → `calendar list --mailbox`).
 */
function resolveExampleCommand(program: Command, tokens: string[]): Command | string {
  if (tokens.length === 0 || tokens[0] !== 'm365-agent-cli') {
    return 'line does not start with m365-agent-cli';
  }
  let cur: Command = program;
  let i = 1;
  while (i < tokens.length) {
    const t = tokens[i];
    if (t.startsWith('-')) break;
    if (isPlaceholderToken(t)) {
      i++;
      continue;
    }
    const child = findNamedSubcommand(cur, t);
    if (child) {
      cur = child;
      i++;
      continue;
    }
    const defChild = getDefaultSubcommand(cur);
    if (defChild) {
      cur = defChild;
      i++;
      continue;
    }
    if (cur === program) {
      return `unknown root subcommand: ${t}`;
    }
    i++;
  }
  if (i < tokens.length && tokens[i].startsWith('-')) {
    const defChild = getDefaultSubcommand(cur);
    if (defChild) cur = defChild;
  }
  return cur;
}

function shellSplit(line: string): string[] {
  const trimmed = line.trim();
  if (!trimmed) return [];
  const tokens: string[] = [];
  let cur = '';
  let quote: "'" | '"' | null = null;
  for (let idx = 0; idx < trimmed.length; idx++) {
    const ch = trimmed[idx];
    if (quote) {
      if (ch === '\\' && quote === '"' && idx + 1 < trimmed.length) {
        cur += trimmed[++idx];
        continue;
      }
      if (ch === quote) {
        quote = null;
        continue;
      }
      cur += ch;
      continue;
    }
    if (ch === '"' || ch === "'") {
      quote = ch as "'" | '"';
      continue;
    }
    if (/\s/.test(ch)) {
      if (cur) {
        tokens.push(cur);
        cur = '';
      }
      continue;
    }
    if (ch === '#') break;
    cur += ch;
  }
  if (cur) tokens.push(cur);
  return tokens;
}

const LONG_FLAG = /--[a-zA-Z][-a-zA-Z0-9]*/g;

function extractLongFlagsFromSegment(segment: string): string[] {
  const out: string[] = [];
  const re = new RegExp(LONG_FLAG.source, 'g');
  for (;;) {
    const m = re.exec(segment);
    if (m === null) break;
    let flag = m[0];
    const eq = flag.indexOf('=');
    if (eq !== -1) flag = flag.slice(0, eq);
    out.push(flag);
  }
  return out;
}

/** Short flags as standalone tokens: -n, -X (not -1, not --foo). */
function extractShortFlagTokens(tokens: string[], startIdx: number): string[] {
  const out: string[] = [];
  for (let i = startIdx; i < tokens.length; i++) {
    const t = tokens[i];
    if (!t.startsWith('-') || t.startsWith('--')) continue;
    if (/^-\d+(\.\d+)?$/.test(t)) continue;
    if (t.length >= 2) out.push(t);
  }
  return out;
}

function validateTokensAgainstCommand(program: Command, tokens: string[]): string | null {
  const resolved = resolveExampleCommand(program, tokens);
  if (typeof resolved === 'string') return resolved;

  const allowed = mergedOptionFlags(resolved);
  let i = 1;
  while (i < tokens.length && !tokens[i].startsWith('-')) i++;

  for (const f of extractLongFlagsFromSegment(tokens.join(' '))) {
    if (!allowed.has(f)) return `unknown flag ${f} for command path ending at "${resolved.name()}"`;
  }
  for (const sf of extractShortFlagTokens(tokens, i)) {
    if (!allowed.has(sf)) return `unknown short flag ${sf} for command path ending at "${resolved.name()}"`;
  }
  return null;
}

function extractM365ExampleLines(text: string): string[] {
  const lines: string[] = [];
  for (const rawLine of text.split('\n')) {
    const line = rawLine.trim();
    if (line.includes('m365-agent-cli')) lines.push(line);
  }
  return lines;
}

/** Join shell lines ending with `\` for continuation. */
function joinShellContinuations(text: string): string {
  const lines = text.split('\n');
  const out: string[] = [];
  let buf = '';
  for (const line of lines) {
    if (buf) buf += ' ';
    const trimmedEnd = line.replace(/\s+$/u, '');
    if (trimmedEnd.endsWith('\\')) {
      buf += trimmedEnd.slice(0, -1).trimEnd();
    } else {
      out.push(buf + line);
      buf = '';
    }
  }
  if (buf) out.push(buf);
  return out.join('\n');
}

export interface CliHelpExampleVerifyResult {
  errors: string[];
}

/**
 * Parse every `--help` output (including `addHelpText`) and validate `m365-agent-cli …` snippets:
 * known command path and flags exist on the resolved Commander node (merged with ancestors).
 */
function verifyCliHelpExamples(program: Command): CliHelpExampleVerifyResult {
  const errors: string[] = [];

  forEachCommandDepthFirst(program, (cmd) => {
    let help: string;
    try {
      help = captureOutputHelp(cmd);
    } catch (e) {
      errors.push(`${cmd.name()}: failed to render help: ${e instanceof Error ? e.message : String(e)}`);
      return;
    }

    const pathHint = cmd.parent ? `${cmd.parent.name()} ${cmd.name()}` : cmd.name();
    for (const line of extractM365ExampleLines(help)) {
      const tokens = shellSplit(line);
      if (tokens.length < 2 || tokens[0] !== 'm365-agent-cli') continue;
      const err = validateTokensAgainstCommand(program, tokens);
      if (err) errors.push(`[help ${pathHint}] ${line}\n  -> ${err}`);
    }
  });

  return { errors };
}

function extractBashFences(md: string): string[] {
  const blocks: string[] = [];
  const re = /^```(?:bash|sh|shell)\s*$/gim;
  for (;;) {
    const m = re.exec(md);
    if (m === null) break;
    const start = m.index + m[0].length;
    const endFence = md.indexOf('```', start);
    if (endFence === -1) break;
    blocks.push(md.slice(start, endFence));
    re.lastIndex = endFence + 3;
  }
  return blocks;
}

/** Same flag/path checks for ```bash sections in CLI_REFERENCE.md. */
function verifyCliReferenceMarkdown(program: Command, mdPath: string): CliHelpExampleVerifyResult {
  const errors: string[] = [];
  let md: string;
  try {
    md = readFileSync(mdPath, 'utf8');
  } catch (e) {
    return { errors: [`Could not read ${mdPath}: ${e instanceof Error ? e.message : String(e)}`] };
  }

  for (const block of extractBashFences(md)) {
    const joined = joinShellContinuations(block);
    for (const rawLine of joined.split('\n')) {
      const line = rawLine.trim();
      if (!line.includes('m365-agent-cli')) continue;
      if (line.startsWith('#')) continue;
      const tokens = shellSplit(line);
      if (tokens.length < 2 || tokens[0] !== 'm365-agent-cli') continue;
      const err = validateTokensAgainstCommand(program, tokens);
      if (err) errors.push(`[${mdPath}] ${line}\n  -> ${err}`);
    }
  }

  return { errors };
}

function defaultCliReferencePath(): string {
  return join(__dirname, '..', '..', 'docs', 'CLI_REFERENCE.md');
}

export function verifyAllCliHelpAndDocExamples(program: Command): CliHelpExampleVerifyResult {
  const a = verifyCliHelpExamples(program);
  const b = verifyCliReferenceMarkdown(program, defaultCliReferencePath());
  return { errors: [...a.errors, ...b.errors] };
}

export function prepareProgramForHelpVerify(program: Command): void {
  forEachCommandDepthFirst(program, (cmd) => {
    cmd.configureOutput({
      writeOut: () => {},
      writeErr: () => {}
    });
    cmd.exitOverride();
  });
}

export function teardownProgramAfterHelpVerify(program: Command): void {
  resetCommanderHooks(program);
}
