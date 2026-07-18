import { afterEach, beforeEach, describe, expect, test } from 'bun:test';
import { Buffer } from 'node:buffer';
import { execFileSync } from 'node:child_process';
import { mkdir, mkdtemp, readdir, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Command } from 'commander';
import { GRAPH_CRITICAL_DELEGATED_SCOPES } from '../lib/graph-oauth-scopes.js';
import { doctorCommand } from './doctor.js';

const CLIENT_ID = '5f2abcea-d6ea-4460-b468-3d80d7a900eb';

function fixtureAccessToken(upn: string): string {
  const h = Buffer.from(JSON.stringify({ alg: 'none', typ: 'JWT' })).toString('base64url');
  const p = Buffer.from(
    JSON.stringify({ exp: 2_000_000_000, appid: CLIENT_ID, upn, scp: GRAPH_CRITICAL_DELEGATED_SCOPES.join(' ') })
  ).toString('base64url');
  return `${h}.${p}.x`;
}

describe('doctor command', () => {
  let testHome: string;
  let originalEnv: NodeJS.ProcessEnv;
  let cwdRestore: string;
  const originalLog = console.log;
  let logs: string[];

  function resetOptionLeaks(): void {
    for (const opt of ['identity', 'mailbox', 'envFile', 'json', 'redactedBundle', 'output']) {
      doctorCommand.setOptionValue(opt, undefined);
    }
    doctorCommand.setOptionValue('format', 'zip');
  }

  beforeEach(async () => {
    testHome = await mkdtemp(join(tmpdir(), 'm365-doctor-cmd-'));
    originalEnv = { ...process.env };
    process.env.M365_AGENT_CLI_CONFIG_DIR = join(testHome, '.config', 'm365-agent-cli');
    process.env.EWS_TENANT_ID = 'common';
    process.env.NODE_ENV = 'test';
    cwdRestore = process.cwd();
    process.chdir(testHome);

    logs = [];
    console.log = ((s: string) => logs.push(s)) as typeof console.log;
  });

  afterEach(async () => {
    console.log = originalLog;
    process.chdir(cwdRestore);
    for (const key of Object.keys(process.env)) {
      if (!(key in originalEnv)) delete process.env[key];
    }
    for (const [key, value] of Object.entries(originalEnv)) {
      if (value === undefined) delete process.env[key];
      else process.env[key] = value;
    }
    await rm(testHome, { recursive: true, force: true }).catch(() => {});
  });

  async function run(argv: string[]): Promise<void> {
    resetOptionLeaks();
    const program = new Command();
    program.exitOverride();
    program.addCommand(doctorCommand);
    await program.parseAsync(['node', 'm365-agent-cli', 'doctor', ...argv]);
  }

  test('--json prints the bundle with no secrets, matching buildDoctorBundle', async () => {
    delete process.env.EWS_CLIENT_ID;
    await run(['--json']);
    const bundle = JSON.parse(logs.at(-1) as string);
    expect(bundle.secretsPrinted).toBe(false);
    expect(bundle.authDiagnosis.failureClass).toBe('missing_credentials');
  });

  test('human output never prints raw refresh token material', async () => {
    process.env.EWS_CLIENT_ID = CLIENT_ID;
    process.env.M365_REFRESH_TOKEN = 'super-secret-refresh-token-value';
    const dir = process.env.M365_AGENT_CLI_CONFIG_DIR as string;
    await mkdir(dir, { recursive: true });
    await writeFile(
      join(dir, 'token-cache-default.json'),
      JSON.stringify({
        version: 1,
        refreshToken: 'super-secret-refresh-token-value',
        graph: { accessToken: fixtureAccessToken('doris@lassfolk.net'), expiresAt: Date.now() + 3_600_000 }
      }),
      'utf8'
    );

    await run([]);
    const output = logs.join('\n');
    expect(output).not.toContain('super-secret-refresh-token-value');
    expect(output).toContain('Auth status:    healthy');
    expect(output).toContain('No tokens, passwords, or message content are printed above.');
  });

  test('--redacted-bundle writes a real zip file readable by an independent unzip tool', async () => {
    delete process.env.EWS_CLIENT_ID;
    const zipPath = join(testHome, 'bundle.zip');
    await run(['--redacted-bundle', zipPath]);

    const listing = execFileSync('unzip', ['-l', zipPath], { encoding: 'utf8' });
    expect(listing).toContain('diagnostic.json');

    const extractDir = join(testHome, 'extracted');
    await mkdir(extractDir, { recursive: true });
    execFileSync('unzip', ['-o', zipPath, '-d', extractDir], { encoding: 'utf8' });
    const content = await Bun.file(join(extractDir, 'diagnostic.json')).json();
    expect(content.secretsPrinted).toBe(false);
    expect(content.authDiagnosis.failureClass).toBe('missing_credentials');
  });

  test('--redacted-bundle --format dir --output writes a diagnostic.json into that directory', async () => {
    delete process.env.EWS_CLIENT_ID;
    const outDir = join(testHome, 'diag-dir');
    await run(['--redacted-bundle', '--format', 'dir', '--output', outDir]);

    const files = await readdir(outDir);
    expect(files).toContain('diagnostic.json');
    const content = await Bun.file(join(outDir, 'diagnostic.json')).json();
    expect(content.secretsPrinted).toBe(false);
  });

  test('--redacted-bundle defaults to ./m365-diagnostic.zip when no path is given', async () => {
    delete process.env.EWS_CLIENT_ID;
    await run(['--redacted-bundle']);
    const defaultPath = join(testHome, 'm365-diagnostic.zip');
    const listing = execFileSync('unzip', ['-l', defaultPath], { encoding: 'utf8' });
    expect(listing).toContain('diagnostic.json');
  });
});
