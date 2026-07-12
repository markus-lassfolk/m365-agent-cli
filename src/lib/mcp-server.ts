/**
 * Native MCP (Model Context Protocol) stdio server: reflects the CLI's own Commander tree (via
 * `command-manifest.ts`) into MCP tools, one per leaf command, and executes a tool call by
 * spawning `bun src/cli.ts <argv>` — the exact same entry point a human or script would run —
 * rather than re-implementing each command's logic. This guarantees identical behavior (including
 * read-only mode, `--dry-run`, and the structured `--json` error envelope) between a direct CLI
 * invocation and an MCP tool call, and sidesteps Commander's per-command `Option` state being
 * shared across command module singletons (re-`parseAsync`-ing the same in-process `Command` tree
 * for unrelated tool calls would otherwise leak stale option values between calls).
 */
import { spawn } from 'node:child_process';
import { createInterface } from 'node:readline';
import { fileURLToPath } from 'node:url';
import type { CommandManifest, ManifestArgument, ManifestOption, ProgramManifest } from './command-manifest.js';

/** Top-level commands never exposed as MCP tools: self-referential, interactive, or long-running/blocking. */
const MCP_EXCLUDED_TOP_LEVEL_COMMANDS = new Set(['mcp', 'serve', 'login', 'update']);

const MCP_PROTOCOL_VERSION = '2024-11-05';

export interface McpArgSpec {
  kind: 'argument' | 'option';
  propName: string;
  required: boolean;
  variadic: boolean;
  /** Only meaningful for `kind: 'option'` — true for a value-less flag (e.g. `--force`). */
  isBoolean: boolean;
  /** Only set for `kind: 'option'` — the flag token to emit, e.g. `--mark-read`. */
  flag?: string;
}

export interface McpJsonSchema {
  type: 'object';
  properties: Record<string, unknown>;
  required: string[];
}

export interface McpToolDef {
  name: string;
  commandPath: string;
  description: string;
  inputSchema: McpJsonSchema;
  argSpecs: McpArgSpec[];
}

function sanitizePropName(raw: string): string {
  const cleaned = raw
    .trim()
    .replace(/[^a-zA-Z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '')
    .toLowerCase();
  return cleaned || 'value';
}

function uniquePropName(base: string, used: Set<string>): string {
  let name = base;
  let i = 2;
  while (used.has(name)) {
    name = `${base}_${i}`;
    i++;
  }
  used.add(name);
  return name;
}

function argumentSchema(a: ManifestArgument): unknown {
  const base = a.variadic ? { type: 'array', items: { type: 'string' } } : { type: 'string' };
  return a.description ? { ...base, description: a.description } : base;
}

function optionSchema(o: ManifestOption, isBoolean: boolean): unknown {
  const base = isBoolean
    ? { type: 'boolean' }
    : o.variadic
      ? { type: 'array', items: { type: 'string' } }
      : { type: 'string' };
  return o.description ? { ...base, description: o.description } : base;
}

/** Builds one MCP tool definition (name, JSON schema, and the argv-reconstruction spec) for a leaf command. */
export function buildToolDefForCommand(cmd: CommandManifest): McpToolDef {
  const used = new Set<string>();
  const argSpecs: McpArgSpec[] = [];
  const properties: Record<string, unknown> = {};
  const required: string[] = [];

  for (const a of cmd.arguments) {
    const propName = uniquePropName(sanitizePropName(a.name), used);
    argSpecs.push({ kind: 'argument', propName, required: a.required, variadic: a.variadic, isBoolean: false });
    properties[propName] = argumentSchema(a);
    if (a.required) required.push(propName);
  }

  for (const o of cmd.options) {
    const flag = o.long ?? o.short ?? undefined;
    if (!flag) continue;
    const isBoolean = !o.valueRequired && !o.valueOptional;
    // Repeatable options in this codebase are usually built with a custom accumulator function
    // (`(v, prev) => [...prev, v]`) and an empty-array default, not Commander's native `<x...>`
    // variadic syntax — `Option.variadic` alone would miss them, so also treat an array default
    // as "repeat this flag once per value" when reconstructing argv from an MCP tool call.
    const variadic = o.variadic || Array.isArray(o.defaultValue);
    const propName = uniquePropName(sanitizePropName(flag), used);
    argSpecs.push({ kind: 'option', propName, required: o.mandatory, variadic, isBoolean, flag });
    properties[propName] = optionSchema({ ...o, variadic }, isBoolean);
    if (o.mandatory) required.push(propName);
  }

  return {
    name: mcpToolNameForPath(cmd.path),
    commandPath: cmd.path,
    description: cmd.description || cmd.path,
    inputSchema: { type: 'object', properties, required },
    argSpecs
  };
}

export function mcpToolNameForPath(path: string): string {
  return path.trim().replace(/\s+/g, '_');
}

/** Walks the manifest and emits one tool per leaf command (a command with no subcommands), skipping the excluded set. */
export function buildMcpTools(manifest: ProgramManifest): McpToolDef[] {
  const tools: McpToolDef[] = [];
  const walk = (cmd: CommandManifest): void => {
    if (cmd.subcommands.length === 0) {
      tools.push(buildToolDefForCommand(cmd));
      return;
    }
    for (const sub of cmd.subcommands) walk(sub);
  };
  for (const cmd of manifest.commands) {
    if (MCP_EXCLUDED_TOP_LEVEL_COMMANDS.has(cmd.name)) continue;
    walk(cmd);
  }
  return tools;
}

function asArray(v: unknown): unknown[] {
  return Array.isArray(v) ? v : [v];
}

/** Reconstructs CLI argv (`[...commandPath.split(' '), ...flags, '--', ...positionals]`) from an
 *  MCP tool call's arguments object. Options come before a `--` separator and positionals after it
 *  so a positional value that happens to start with `-` (e.g. free-text search query) is never
 *  misparsed by Commander as an unknown option. When `opts.forceJson` is true and `--json` wasn't
 *  already requested via `args`, it's appended to the options segment (before the separator) here
 *  — inserting it structurally during construction, rather than searching the finished array for
 *  the separator afterward, avoids colliding with a positional/option value that happens to equal
 *  the literal string `"--"`. */
export function toolArgsToArgv(
  tool: McpToolDef,
  args: Record<string, unknown>,
  opts: { forceJson?: boolean } = {}
): string[] {
  const commandPath = tool.commandPath.split(/\s+/).filter(Boolean);
  const positionalArgv: string[] = [];
  const optionArgv: string[] = [];

  for (const spec of tool.argSpecs) {
    if (spec.kind !== 'argument') continue;
    const v = args[spec.propName];
    if (v === undefined || v === null) continue;
    if (spec.variadic) {
      for (const item of asArray(v)) positionalArgv.push(String(item));
    } else {
      positionalArgv.push(String(v));
    }
  }

  for (const spec of tool.argSpecs) {
    if (spec.kind !== 'option') continue;
    const v = args[spec.propName];
    if (v === undefined || v === null) continue;
    if (spec.isBoolean) {
      if (v === true) optionArgv.push(spec.flag as string);
    } else if (spec.variadic) {
      for (const item of asArray(v)) {
        optionArgv.push(spec.flag as string);
        optionArgv.push(String(item));
      }
    } else {
      optionArgv.push(spec.flag as string);
      optionArgv.push(String(v));
    }
  }

  if (opts.forceJson && !optionArgv.includes('--json')) {
    optionArgv.push('--json');
  }

  return positionalArgv.length > 0
    ? [...commandPath, ...optionArgv, '--', ...positionalArgv]
    : [...commandPath, ...optionArgv];
}

export interface RunCliResult {
  stdout: string;
  stderr: string;
  exitCode: number;
  timedOut: boolean;
}

const CLI_ENTRY = fileURLToPath(new URL('../cli.ts', import.meta.url));

const MCP_TOOL_TIMEOUT_MS =
  Number(process.env.M365_MCP_TOOL_TIMEOUT_MS) > 0 ? Number(process.env.M365_MCP_TOOL_TIMEOUT_MS) : 120_000;

const MCP_TOOL_KILL_GRACE_MS = 5000;

/** Sends `signal` to the whole child process tree, not just the direct child — a CLI subprocess
 *  can itself spawn (e.g. a helper process), and a lone `child.kill()` would leave it running past
 *  the tool timeout. POSIX: negative pid targets the process group created by `detached: true`
 *  below. Windows: `taskkill /t` walks the tree by pid. */
export function killChildTree(child: ReturnType<typeof spawn>, signal: NodeJS.Signals): void {
  if (process.platform === 'win32') {
    // An unlistened 'error' event on a failed spawn() throws and crashes the process — a missing
    // or blocked taskkill binary must only fail this one kill attempt, not take down the server.
    if (child.pid) spawn('taskkill', ['/pid', String(child.pid), '/t', '/f']).on('error', () => {});
    return;
  }
  try {
    if (child.pid) process.kill(-child.pid, signal);
  } catch {
    child.kill(signal);
  }
}

/** Spawns `bun src/cli.ts <argv>` (the same entry point a human runs) and captures its output. */
function runCli(argv: string[]): Promise<RunCliResult> {
  return new Promise((resolve) => {
    const child = spawn(process.execPath, [CLI_ENTRY, ...argv], {
      env: process.env,
      stdio: ['ignore', 'pipe', 'pipe'],
      detached: process.platform !== 'win32'
    });
    let stdout = '';
    let stderr = '';
    let settled = false;
    let killGraceTimer: ReturnType<typeof setTimeout> | undefined;

    const timer = setTimeout(() => {
      if (settled) return;
      settled = true;
      killChildTree(child, 'SIGTERM');
      // The child (or a descendant) may ignore SIGTERM; escalate if it's still alive shortly after.
      killGraceTimer = setTimeout(() => killChildTree(child, 'SIGKILL'), MCP_TOOL_KILL_GRACE_MS);
      resolve({ stdout, stderr, exitCode: 1, timedOut: true });
    }, MCP_TOOL_TIMEOUT_MS);

    child.stdout?.on('data', (d: Buffer) => {
      stdout += d.toString('utf8');
    });
    child.stderr?.on('data', (d: Buffer) => {
      stderr += d.toString('utf8');
    });
    child.on('close', (code) => {
      clearTimeout(killGraceTimer);
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      resolve({ stdout, stderr, exitCode: code ?? 1, timedOut: false });
    });
    child.on('error', (err) => {
      clearTimeout(killGraceTimer);
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      resolve({ stdout, stderr: stderr ? `${stderr}\n${err.message}` : err.message, exitCode: 1, timedOut: false });
    });
  });
}

export interface McpContext {
  tools: McpToolDef[];
  toolsByName: Map<string, McpToolDef>;
  serverName: string;
  serverVersion: string;
  runCli: (argv: string[]) => Promise<RunCliResult>;
}

export function buildMcpContext(
  manifest: ProgramManifest,
  overrides: Partial<Pick<McpContext, 'runCli'>> = {}
): McpContext {
  const tools = buildMcpTools(manifest);
  return {
    tools,
    toolsByName: new Map(tools.map((t) => [t.name, t])),
    serverName: manifest.name || 'm365-agent-cli',
    serverVersion: manifest.version,
    runCli: overrides.runCli ?? runCli
  };
}

function jsonRpcResult(id: unknown, result: unknown): Record<string, unknown> {
  return { jsonrpc: '2.0', id, result };
}

function jsonRpcError(id: unknown, code: number, message: string): Record<string, unknown> {
  return { jsonrpc: '2.0', id, error: { code, message } };
}

/**
 * Handles one already-parsed JSON-RPC message and returns the response object to write back, or
 * `null` when no response should be sent (a notification, or a malformed message with no `id`).
 * Pure aside from `ctx.runCli` — kept separate from the stdio loop so it's directly unit-testable.
 */
export async function handleMcpMessage(msg: unknown, ctx: McpContext): Promise<Record<string, unknown> | null> {
  if (!msg || typeof msg !== 'object') return null;
  const m = msg as { jsonrpc?: string; id?: unknown; method?: string; params?: unknown };
  const hasId = 'id' in m && m.id !== undefined && m.id !== null;
  if (typeof m.method !== 'string') return null;

  if (m.method === 'initialize') {
    const params = (m.params ?? {}) as { protocolVersion?: string };
    return jsonRpcResult(m.id, {
      protocolVersion: params.protocolVersion || MCP_PROTOCOL_VERSION,
      capabilities: { tools: {} },
      serverInfo: { name: ctx.serverName, version: ctx.serverVersion }
    });
  }

  if (m.method.startsWith('notifications/')) {
    return null;
  }

  if (m.method === 'ping') {
    return jsonRpcResult(m.id, {});
  }

  if (m.method === 'tools/list') {
    return jsonRpcResult(m.id, {
      tools: ctx.tools.map((t) => ({ name: t.name, description: t.description, inputSchema: t.inputSchema }))
    });
  }

  if (m.method === 'tools/call') {
    if (!hasId) return null;
    const params = (m.params ?? {}) as { name?: string; arguments?: Record<string, unknown> };
    const tool = params.name ? ctx.toolsByName.get(params.name) : undefined;
    if (!tool) {
      return jsonRpcError(m.id, -32602, `Unknown tool: ${params.name ?? '(missing name)'}`);
    }
    const supportsJson = tool.argSpecs.some((s) => s.kind === 'option' && s.flag === '--json');
    const argv = toolArgsToArgv(tool, params.arguments ?? {}, { forceJson: supportsJson });

    const result = await ctx.runCli(argv);
    const text = result.stdout.trim() || result.stderr.trim() || '(no output)';
    return jsonRpcResult(m.id, {
      content: [{ type: 'text', text }],
      isError: result.exitCode !== 0
    });
  }

  if (!hasId) return null;
  return jsonRpcError(m.id, -32601, `Method not found: ${m.method}`);
}

/** Runs the MCP stdio loop: one JSON-RPC message per line on stdin, one response per line on stdout. */
export async function runMcpStdioServer(manifest: ProgramManifest, overrides: Partial<McpContext> = {}): Promise<void> {
  const ctx = { ...buildMcpContext(manifest), ...overrides };
  const rl = createInterface({ input: process.stdin, terminal: false });
  for await (const line of rl) {
    const trimmed = line.trim();
    if (!trimmed) continue;
    let msg: unknown;
    try {
      msg = JSON.parse(trimmed);
    } catch {
      continue;
    }
    const response = await handleMcpMessage(msg, ctx);
    if (response) {
      process.stdout.write(`${JSON.stringify(response)}\n`);
    }
  }
}
