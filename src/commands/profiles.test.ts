import { afterEach, beforeEach, describe, expect, test } from 'bun:test';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Command } from 'commander';
import { profilesCommand } from './profiles.js';

describe('profiles command', () => {
  let testHome: string;
  let originalConfigDir: string | undefined;
  const originalLog = console.log;
  const originalError = console.error;
  const originalExit = process.exit;
  let logs: string[];
  let errors: string[];
  let exitCode: number | undefined;

  beforeEach(async () => {
    testHome = await mkdtemp(join(tmpdir(), 'm365-profiles-cmd-'));
    originalConfigDir = process.env.M365_AGENT_CLI_CONFIG_DIR;
    process.env.M365_AGENT_CLI_CONFIG_DIR = join(testHome, '.config', 'm365-agent-cli');

    logs = [];
    errors = [];
    exitCode = undefined;
    console.log = ((s: string) => logs.push(s)) as typeof console.log;
    console.error = ((s: string) => errors.push(s)) as typeof console.error;
    process.exit = ((code?: number) => {
      exitCode = code;
      throw new Error(`process.exit(${code})`);
    }) as never;
  });

  afterEach(async () => {
    console.log = originalLog;
    console.error = originalError;
    process.exit = originalExit;
    if (originalConfigDir === undefined) {
      delete process.env.M365_AGENT_CLI_CONFIG_DIR;
    } else {
      process.env.M365_AGENT_CLI_CONFIG_DIR = originalConfigDir;
    }
    await rm(testHome, { recursive: true, force: true }).catch(() => {});
  });

  async function run(argv: string[]): Promise<void> {
    const program = new Command();
    program.exitOverride();
    program.addCommand(profilesCommand);
    await program.parseAsync(['node', 'm365-agent-cli', 'profiles', ...argv]);
  }

  test('list prints a friendly message when no profiles exist', async () => {
    await run(['list']);
    expect(logs.join('\n')).toContain('No identity profiles registered yet.');
  });

  test('set-default registers and selects a profile; list --json reflects it', async () => {
    await run(['set-default', 'doris', '--json']);
    expect(JSON.parse(logs.at(-1) as string)).toEqual({ defaultProfile: 'doris', identity: 'doris' });

    logs = [];
    await run(['list', '--json']);
    const parsed = JSON.parse(logs.at(-1) as string);
    expect(parsed.defaultProfile).toBe('doris');
    expect(parsed.profiles).toHaveLength(1);
    expect(parsed.profiles[0]).toMatchObject({
      name: 'doris',
      identity: 'doris',
      isDefault: true,
      cacheHealth: 'missing'
    });
  });

  test('show without a name uses the default profile', async () => {
    await run(['set-default', 'doris']);
    logs = [];
    await run(['show', '--json']);
    const parsed = JSON.parse(logs.at(-1) as string);
    expect(parsed.name).toBe('doris');
    expect(parsed.isDefault).toBe(true);
  });

  test('show exits 1 with a structured error when no name and no default', async () => {
    await expect(run(['show', '--json'])).rejects.toThrow();
    expect(exitCode).toBe(1);
    expect(JSON.parse(logs.at(-1) as string).error.message).toContain('No profile name given');
  });

  test('show exits 1 for an unknown profile name', async () => {
    await expect(run(['show', 'ghost', '--json'])).rejects.toThrow();
    expect(exitCode).toBe(1);
    expect(JSON.parse(logs.at(-1) as string).error.message).toContain('No such profile: ghost');
  });

  test('delete removes a profile and clears the default when it was the default', async () => {
    await run(['set-default', 'doris']);
    logs = [];
    await run(['delete', 'doris', '--json']);
    expect(JSON.parse(logs.at(-1) as string)).toEqual({ deleted: 'doris', purgedCache: false });

    logs = [];
    await run(['list', '--json']);
    const parsed = JSON.parse(logs.at(-1) as string);
    expect(parsed.defaultProfile).toBeNull();
    expect(parsed.profiles).toEqual([]);
  });

  test('delete exits 1 for an unknown profile', async () => {
    await expect(run(['delete', 'ghost', '--json'])).rejects.toThrow();
    expect(exitCode).toBe(1);
  });
});
