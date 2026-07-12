import { describe, expect, test } from 'bun:test';
import { spawn } from 'node:child_process';
import { dirname, join } from 'node:path';
import { createInterface } from 'node:readline';
import { fileURLToPath } from 'node:url';

/**
 * Real subprocess test: spawns `bun src/cli.ts mcp` (the actual entry point, not a mocked
 * in-process harness) and speaks newline-delimited JSON-RPC over its stdio, matching what a real
 * MCP client does. The unit tests in `src/lib/mcp-server.test.ts` mock `runCli` and never exercise
 * the readline loop or the real subprocess spawn — this test covers that wiring.
 */
const repoRoot = join(dirname(fileURLToPath(import.meta.url)), '../..');
const cliEntry = join(repoRoot, 'src/cli.ts');

interface RpcClient {
  request(method: string, params?: unknown, id?: number): Promise<Record<string, unknown>>;
  notify(method: string, params?: unknown): void;
  close(): void;
}

function startMcpServer(): RpcClient {
  const child = spawn(process.execPath, [cliEntry, 'mcp'], {
    cwd: repoRoot,
    stdio: ['pipe', 'pipe', 'pipe']
  });
  const rl = createInterface({ input: child.stdout, terminal: false });
  const pending = new Map<number, (msg: Record<string, unknown>) => void>();

  rl.on('line', (line) => {
    const trimmed = line.trim();
    if (!trimmed) return;
    try {
      const msg = JSON.parse(trimmed) as Record<string, unknown>;
      const id = msg.id as number | undefined;
      if (typeof id === 'number' && pending.has(id)) {
        pending.get(id)?.(msg);
        pending.delete(id);
      }
    } catch {
      // ignore non-JSON stdout noise
    }
  });

  let nextId = 1;
  return {
    request(method, params, id) {
      const useId = id ?? nextId++;
      return new Promise((resolve, reject) => {
        const timer = setTimeout(() => {
          pending.delete(useId);
          reject(new Error(`Timed out waiting for response to ${method}`));
        }, 20_000);
        pending.set(useId, (msg) => {
          clearTimeout(timer);
          resolve(msg);
        });
        child.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', id: useId, method, params })}\n`);
      });
    },
    notify(method, params) {
      child.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', method, params })}\n`);
    },
    close() {
      child.stdin.end();
      child.kill();
    }
  };
}

describe('mcp stdio server (real subprocess)', () => {
  test('initialize, tools/list, and tools/call(describe) round-trip over stdio', async () => {
    const client = startMcpServer();
    try {
      const init = await client.request('initialize', { protocolVersion: '2024-11-05', capabilities: {} });
      expect(init.result).toMatchObject({ protocolVersion: '2024-11-05', capabilities: { tools: {} } });
      client.notify('notifications/initialized');

      const list = await client.request('tools/list');
      const tools = (list.result as { tools: Array<{ name: string }> }).tools;
      expect(tools.some((t) => t.name === 'describe')).toBe(true);
      expect(tools.some((t) => t.name === 'mcp')).toBe(false);
      expect(tools.some((t) => t.name === 'serve')).toBe(false);
      expect(tools.some((t) => t.name === 'login')).toBe(false);

      const call = await client.request('tools/call', { name: 'describe', arguments: { list: true } });
      const result = call.result as { content: Array<{ type: string; text: string }>; isError: boolean };
      expect(result.isError).toBe(false);
      const parsed = JSON.parse(result.content[0].text) as Array<{ name: string }>;
      expect(parsed.some((c) => c.name === 'mail')).toBe(true);
    } finally {
      client.close();
    }
  }, 30_000);

  test('tools/call with an unknown tool name returns a JSON-RPC error', async () => {
    const client = startMcpServer();
    try {
      await client.request('initialize', { protocolVersion: '2024-11-05', capabilities: {} });
      const call = await client.request('tools/call', { name: 'not-a-real-tool', arguments: {} });
      expect(call.error).toMatchObject({ code: -32602 });
    } finally {
      client.close();
    }
  }, 30_000);
});
