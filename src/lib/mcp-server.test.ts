import { describe, expect, it, test } from 'bun:test';
import { Command } from 'commander';
import { describeProgram } from './command-manifest.js';
import {
  buildMcpContext,
  buildMcpTools,
  buildToolDefForCommand,
  handleMcpMessage,
  mcpToolNameForPath,
  type RunCliResult,
  toolArgsToArgv
} from './mcp-server.js';

function buildFixtureProgram(): Command {
  const root = new Command('m365-agent-cli').description('root desc');

  const mail = new Command('mail')
    .description('mail desc')
    .argument('[folder]', 'folder name')
    .option('--reply <id>', 'reply to a message')
    .option('--force', 'force it')
    .option('--category <name>', 'category (repeatable)', (v: string, prev: string[]) => [...prev, v], [] as string[])
    .option('--json', 'output json');
  root.addCommand(mail);

  const rules = new Command('rules').description('rules desc');
  rules.addCommand(new Command('create').description('create a rule').requiredOption('--name <name>', 'rule name'));
  root.addCommand(rules);

  root.addCommand(new Command('mcp').description('mcp server'));
  root.addCommand(new Command('serve').description('webhook server'));
  root.addCommand(new Command('login').description('device code login'));
  root.addCommand(new Command('update').description('self-update: npm/bun global install'));

  return root;
}

function okResult(stdout: string): RunCliResult {
  return { stdout, stderr: '', exitCode: 0, timedOut: false };
}

describe('mcpToolNameForPath', () => {
  it('joins a space-separated path with underscores', () => {
    expect(mcpToolNameForPath('rules create')).toBe('rules_create');
    expect(mcpToolNameForPath('mail')).toBe('mail');
  });
});

describe('buildToolDefForCommand', () => {
  test('builds a JSON schema from arguments and options', () => {
    const root = buildFixtureProgram();
    const manifest = describeProgram(root);
    const mailManifest = manifest.commands.find((c) => c.name === 'mail')!;
    const tool = buildToolDefForCommand(mailManifest);

    expect(tool.name).toBe('mail');
    expect(tool.commandPath).toBe('mail');
    expect(tool.description).toBe('mail desc');

    expect(tool.inputSchema.properties.folder).toEqual({ type: 'string', description: 'folder name' });
    expect(tool.inputSchema.properties.reply).toEqual({ type: 'string', description: 'reply to a message' });
    expect(tool.inputSchema.properties.force).toEqual({ type: 'boolean', description: 'force it' });
    expect(tool.inputSchema.properties.json).toEqual({ type: 'boolean', description: 'output json' });
    expect(tool.inputSchema.required).toEqual([]);
  });

  test('marks --requiredOption() flags and positional required args as required', () => {
    const root = buildFixtureProgram();
    const manifest = describeProgram(root);
    const create = manifest.commands.find((c) => c.name === 'rules')!.subcommands.find((c) => c.name === 'create')!;
    const tool = buildToolDefForCommand(create);
    expect(tool.inputSchema.required).toEqual(['name']);
  });
});

describe('buildMcpTools', () => {
  test('emits one tool per leaf command and recurses into subcommands', () => {
    const manifest = describeProgram(buildFixtureProgram());
    const tools = buildMcpTools(manifest);
    const names = tools.map((t) => t.name).sort();
    expect(names).toEqual(['mail', 'rules_create']);
  });

  test('excludes mcp, serve, login, and update', () => {
    const manifest = describeProgram(buildFixtureProgram());
    const tools = buildMcpTools(manifest);
    expect(tools.some((t) => t.name === 'mcp')).toBe(false);
    expect(tools.some((t) => t.name === 'serve')).toBe(false);
    expect(tools.some((t) => t.name === 'login')).toBe(false);
    expect(tools.some((t) => t.name === 'update')).toBe(false);
  });
});

describe('toolArgsToArgv', () => {
  const manifest = describeProgram(buildFixtureProgram());
  const mailTool = buildMcpTools(manifest).find((t) => t.name === 'mail')!;
  const createTool = buildMcpTools(manifest).find((t) => t.name === 'rules_create')!;

  test('emits the command path, then positional args, then flags', () => {
    expect(toolArgsToArgv(mailTool, { folder: 'inbox', reply: 'msg-1' })).toEqual([
      'mail',
      'inbox',
      '--reply',
      'msg-1'
    ]);
  });

  test('omits undefined/null args entirely', () => {
    expect(toolArgsToArgv(mailTool, {})).toEqual(['mail']);
  });

  test('emits a boolean flag only when true', () => {
    expect(toolArgsToArgv(mailTool, { force: true })).toEqual(['mail', '--force']);
    expect(toolArgsToArgv(mailTool, { force: false })).toEqual(['mail']);
  });

  test('repeats a variadic option flag once per array item', () => {
    expect(toolArgsToArgv(mailTool, { category: ['a', 'b'] })).toEqual(['mail', '--category', 'a', '--category', 'b']);
  });

  test('emits required-option flags for a subcommand path', () => {
    expect(toolArgsToArgv(createTool, { name: 'Auto-archive' })).toEqual(['rules', 'create', '--name', 'Auto-archive']);
  });
});

describe('handleMcpMessage', () => {
  const manifest = describeProgram(buildFixtureProgram());

  test('initialize echoes back the client protocol version and advertises tools capability', async () => {
    const ctx = buildMcpContext(manifest, { runCli: async () => okResult('') });
    const res = await handleMcpMessage(
      { jsonrpc: '2.0', id: 1, method: 'initialize', params: { protocolVersion: '2025-03-26' } },
      ctx
    );
    expect(res?.result).toMatchObject({
      protocolVersion: '2025-03-26',
      capabilities: { tools: {} },
      serverInfo: { name: 'm365-agent-cli' }
    });
  });

  test('a notification (no id) never produces a response', async () => {
    const ctx = buildMcpContext(manifest, { runCli: async () => okResult('') });
    const res = await handleMcpMessage({ jsonrpc: '2.0', method: 'notifications/initialized' }, ctx);
    expect(res).toBeNull();
  });

  test('tools/list returns every built tool with name/description/inputSchema', async () => {
    const ctx = buildMcpContext(manifest, { runCli: async () => okResult('') });
    const res = await handleMcpMessage({ jsonrpc: '2.0', id: 2, method: 'tools/list' }, ctx);
    const tools = (res?.result as { tools: Array<{ name: string }> }).tools;
    expect(tools.some((t) => t.name === 'mail')).toBe(true);
    expect(tools.some((t) => t.name === 'rules_create')).toBe(true);
  });

  test('tools/call spawns the CLI with the reconstructed argv and returns its stdout as content', async () => {
    let capturedArgv: string[] = [];
    const ctx = buildMcpContext(manifest, {
      runCli: async (argv) => {
        capturedArgv = argv;
        return okResult('{"ok":true}');
      }
    });
    const res = await handleMcpMessage(
      { jsonrpc: '2.0', id: 3, method: 'tools/call', params: { name: 'mail', arguments: { folder: 'inbox' } } },
      ctx
    );
    expect(capturedArgv).toEqual(['mail', 'inbox', '--json']);
    expect(res?.result).toEqual({ content: [{ type: 'text', text: '{"ok":true}' }], isError: false });
  });

  test('tools/call does not double-append --json when already supported and requested', async () => {
    let capturedArgv: string[] = [];
    const ctx = buildMcpContext(manifest, {
      runCli: async (argv) => {
        capturedArgv = argv;
        return okResult('{}');
      }
    });
    await handleMcpMessage(
      { jsonrpc: '2.0', id: 4, method: 'tools/call', params: { name: 'mail', arguments: { json: true } } },
      ctx
    );
    expect(capturedArgv.filter((a) => a === '--json')).toHaveLength(1);
  });

  test('tools/call surfaces a non-zero exit code as isError with stderr text', async () => {
    const ctx = buildMcpContext(manifest, {
      runCli: async () => ({ stdout: '', stderr: 'Error: boom', exitCode: 1, timedOut: false })
    });
    const res = await handleMcpMessage(
      { jsonrpc: '2.0', id: 5, method: 'tools/call', params: { name: 'mail', arguments: {} } },
      ctx
    );
    expect(res?.result).toEqual({ content: [{ type: 'text', text: 'Error: boom' }], isError: true });
  });

  test('tools/call with an unknown tool name returns a JSON-RPC error, without invoking runCli', async () => {
    let called = false;
    const ctx = buildMcpContext(manifest, {
      runCli: async () => {
        called = true;
        return okResult('');
      }
    });
    const res = await handleMcpMessage(
      { jsonrpc: '2.0', id: 6, method: 'tools/call', params: { name: 'nope', arguments: {} } },
      ctx
    );
    expect(res?.error).toMatchObject({ code: -32602 });
    expect(called).toBe(false);
  });

  test('an unknown method with an id returns a JSON-RPC "method not found" error', async () => {
    const ctx = buildMcpContext(manifest, { runCli: async () => okResult('') });
    const res = await handleMcpMessage({ jsonrpc: '2.0', id: 7, method: 'not/a/thing' }, ctx);
    expect(res?.error).toMatchObject({ code: -32601 });
  });

  test('an unknown method with no id (malformed notification) produces no response', async () => {
    const ctx = buildMcpContext(manifest, { runCli: async () => okResult('') });
    const res = await handleMcpMessage({ jsonrpc: '2.0', method: 'not/a/thing' }, ctx);
    expect(res).toBeNull();
  });

  test('ping returns an empty result', async () => {
    const ctx = buildMcpContext(manifest, { runCli: async () => okResult('') });
    const res = await handleMcpMessage({ jsonrpc: '2.0', id: 8, method: 'ping' }, ctx);
    expect(res?.result).toEqual({});
  });

  test('a non-object message is ignored', async () => {
    const ctx = buildMcpContext(manifest, { runCli: async () => okResult('') });
    expect(await handleMcpMessage(null, ctx)).toBeNull();
    expect(await handleMcpMessage('a string', ctx)).toBeNull();
  });
});
