import { afterEach, beforeEach, describe, expect, test } from 'bun:test';
import { Buffer } from 'node:buffer';
import { mkdir, mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Command } from 'commander';
import { GRAPH_CRITICAL_DELEGATED_SCOPES } from '../lib/graph-oauth-scopes.js';
import { authCommand } from './auth.js';

const CLIENT_ID = '5f2abcea-d6ea-4460-b468-3d80d7a900eb';

const AADSTS50173_FIXTURE =
  "invalid_grant: AADSTS50173: The provided grant has expired due to it being revoked, a fresh auth token is needed. The grant was issued on '2026-01-01T00:00:00.0000000Z' and the TokensValidFrom date (before which tokens are not valid) for this user is '2026-02-01T00:00:00.0000000Z'.";

describe('auth repair', () => {
  let testHome: string;
  let originalEnv: NodeJS.ProcessEnv;
  const originalLog = console.log;
  const originalError = console.error;
  const originalExit = process.exit;
  const originalFetch = global.fetch;
  let logs: string[];
  let errors: string[];
  let exitCode: number | undefined;

  beforeEach(async () => {
    testHome = await mkdtemp(join(tmpdir(), 'm365-auth-repair-'));
    originalEnv = { ...process.env };
    process.env.M365_AGENT_CLI_CONFIG_DIR = join(testHome, '.config', 'm365-agent-cli');
    process.env.EWS_TENANT_ID = 'common';
    process.env.NODE_ENV = 'test';

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
    global.fetch = originalFetch;
    for (const key of Object.keys(process.env)) {
      if (!(key in originalEnv)) delete process.env[key];
    }
    for (const [key, value] of Object.entries(originalEnv)) {
      if (value === undefined) delete process.env[key];
      else process.env[key] = value;
    }
    await rm(testHome, { recursive: true, force: true }).catch(() => {});
  });

  /** `authCommand`/its `repair` subcommand are module-level singletons reused across every test
   *  in this file — Commander only overwrites an option when the new argv includes its flag, so a
   *  flag set `true` by one test (e.g. `--json`) otherwise leaks into the next parse. */
  function resetAuthRepairOptionLeaks(): void {
    const repair = authCommand.commands.find((c) => c.name() === 'repair');
    for (const opt of ['identity', 'startLogin', 'json', 'envFile']) {
      repair?.setOptionValue(opt, undefined);
    }
    repair?.setOptionValue('secrets', true);
  }

  async function run(argv: string[]): Promise<void> {
    resetAuthRepairOptionLeaks();
    const program = new Command();
    program.exitOverride();
    program.addCommand(authCommand);
    await program.parseAsync(['node', 'm365-agent-cli', 'auth', ...argv]);
  }

  test('--json reports repair_required with a safe recovery command when credentials are missing', async () => {
    delete process.env.EWS_CLIENT_ID;
    delete process.env.M365_REFRESH_TOKEN;
    delete process.env.GRAPH_REFRESH_TOKEN;
    delete process.env.EWS_REFRESH_TOKEN;

    await run(['repair', '--json']);
    const diag = JSON.parse(logs.at(-1) as string);
    expect(diag.status).toBe('repair_required');
    expect(diag.failureClass).toBe('missing_credentials');
    expect(diag.safeCommand).toBe('m365-agent-cli login');
    expect(diag.secretsPrinted).toBe(false);
  });

  test('exits 0 even when repair is required (status carries the signal, not the exit code)', async () => {
    delete process.env.EWS_CLIENT_ID;
    delete process.env.M365_REFRESH_TOKEN;
    await run(['repair', '--json']);
    expect(exitCode).toBeUndefined();
  });

  test('text output classifies a real-world AADSTS50173 payload as a revoked grant, not cache corruption', async () => {
    process.env.EWS_CLIENT_ID = CLIENT_ID;
    process.env.M365_REFRESH_TOKEN = 'fake-refresh-token';
    global.fetch = (async () =>
      new Response(JSON.stringify({ error: 'invalid_grant', error_description: AADSTS50173_FIXTURE }), {
        status: 400
      })) as unknown as typeof fetch;

    await run(['repair']);
    const output = logs.join('\n');
    expect(output).toContain('repair required');
    expect(output).toContain('revoked by tenant policy');
    expect(output).toContain('AADSTS50173');
    expect(output).toContain('Command: m365-agent-cli login');
    expect(output).toContain('Safety: no secrets printed');
    expect(output).not.toContain('fake-refresh-token');
  });

  test('reports healthy status for a valid cached token', async () => {
    process.env.EWS_CLIENT_ID = CLIENT_ID;
    process.env.M365_REFRESH_TOKEN = 'fake-refresh-token';
    const dir = process.env.M365_AGENT_CLI_CONFIG_DIR as string;
    await mkdir(dir, { recursive: true });
    const h = Buffer.from(JSON.stringify({ alg: 'none', typ: 'JWT' })).toString('base64url');
    const p = Buffer.from(
      JSON.stringify({
        exp: 2_000_000_000,
        appid: CLIENT_ID,
        upn: 'doris@lassfolk.net',
        scp: GRAPH_CRITICAL_DELEGATED_SCOPES.join(' ')
      })
    ).toString('base64url');
    await writeFile(
      join(dir, 'token-cache-default.json'),
      JSON.stringify({
        version: 1,
        refreshToken: 'fake-refresh-token',
        graph: { accessToken: `${h}.${p}.x`, expiresAt: Date.now() + 3_600_000 }
      }),
      'utf8'
    );

    await run(['repair', '--json']);
    const diag = JSON.parse(logs.at(-1) as string);
    expect(diag.status).toBe('healthy');
    expect(diag.signedInAs).toBe('doris@lassfolk.net');
  });
});
