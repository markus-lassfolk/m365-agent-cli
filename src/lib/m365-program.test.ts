import { afterEach, beforeEach, describe, expect, test } from 'bun:test';
import { Command } from 'commander';
import { Buffer } from 'node:buffer';
import { mkdir, mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { GRAPH_CRITICAL_DELEGATED_SCOPES } from './graph-oauth-scopes.js';
import { createM365Program } from './m365-program.js';

describe('createM365Program --dry-run wiring', () => {
  test('preAction hook sets M365_DRY_RUN before a command action runs', async () => {
    const program = createM365Program();
    let sawDuringAction: string | undefined;
    program.addCommand(
      new Command('__dry_run_probe__').action(() => {
        sawDuringAction = process.env.M365_DRY_RUN;
      })
    );
    program.exitOverride();

    const originalEnv = process.env.M365_DRY_RUN;
    try {
      delete process.env.M365_DRY_RUN;
      await program.parseAsync(['node', 'm365-agent-cli', '--dry-run', '__dry_run_probe__']);
      expect(sawDuringAction).toBe('1');
    } finally {
      if (originalEnv === undefined) delete process.env.M365_DRY_RUN;
      else process.env.M365_DRY_RUN = originalEnv;
    }
  });

  test('does not set M365_DRY_RUN when --dry-run is absent', async () => {
    const program = createM365Program();
    let sawDuringAction: string | undefined;
    program.addCommand(
      new Command('__dry_run_probe_off__').action(() => {
        sawDuringAction = process.env.M365_DRY_RUN;
      })
    );
    program.exitOverride();

    const originalEnv = process.env.M365_DRY_RUN;
    try {
      delete process.env.M365_DRY_RUN;
      await program.parseAsync(['node', 'm365-agent-cli', '__dry_run_probe_off__']);
      expect(sawDuringAction).toBeUndefined();
    } finally {
      if (originalEnv === undefined) delete process.env.M365_DRY_RUN;
      else process.env.M365_DRY_RUN = originalEnv;
    }
  });

  test('registers a top-level --dry-run option on the root program', () => {
    const program = createM365Program();
    expect(program.options.some((o) => o.long === '--dry-run')).toBe(true);
  });
});

describe('createM365Program --cache wiring', () => {
  test('preAction hook sets M365_CACHE_TTL from --cache before a command action runs', async () => {
    const program = createM365Program();
    let sawDuringAction: string | undefined;
    program.addCommand(
      new Command('__cache_probe__').action(() => {
        sawDuringAction = process.env.M365_CACHE_TTL;
      })
    );
    program.exitOverride();

    const originalEnv = process.env.M365_CACHE_TTL;
    try {
      delete process.env.M365_CACHE_TTL;
      await program.parseAsync(['node', 'm365-agent-cli', '--cache', '5m', '__cache_probe__']);
      expect(sawDuringAction).toBe('5m');
    } finally {
      if (originalEnv === undefined) delete process.env.M365_CACHE_TTL;
      else process.env.M365_CACHE_TTL = originalEnv;
    }
  });

  test('does not set M365_CACHE_TTL when --cache is absent', async () => {
    const program = createM365Program();
    let sawDuringAction: string | undefined;
    program.addCommand(
      new Command('__cache_probe_off__').action(() => {
        sawDuringAction = process.env.M365_CACHE_TTL;
      })
    );
    program.exitOverride();

    const originalEnv = process.env.M365_CACHE_TTL;
    try {
      delete process.env.M365_CACHE_TTL;
      await program.parseAsync(['node', 'm365-agent-cli', '__cache_probe_off__']);
      expect(sawDuringAction).toBeUndefined();
    } finally {
      if (originalEnv === undefined) delete process.env.M365_CACHE_TTL;
      else process.env.M365_CACHE_TTL = originalEnv;
    }
  });

  test('registers a top-level --cache option on the root program', () => {
    const program = createM365Program();
    expect(program.options.some((o) => o.long === '--cache')).toBe(true);
  });
});

describe('createM365Program --require-identity wiring', () => {
  const CLIENT_ID = '5f2abcea-d6ea-4460-b468-3d80d7a900eb';
  let testHome: string;
  let originalEnv: NodeJS.ProcessEnv;

  beforeEach(async () => {
    testHome = await mkdtemp(join(tmpdir(), 'm365-program-guard-'));
    originalEnv = { ...process.env };
    process.env.M365_AGENT_CLI_CONFIG_DIR = join(testHome, '.config', 'm365-agent-cli');
    process.env.EWS_TENANT_ID = 'common';
    process.env.EWS_CLIENT_ID = CLIENT_ID;
    process.env.M365_REFRESH_TOKEN = 'fake-refresh-token';
  });

  afterEach(async () => {
    for (const key of Object.keys(process.env)) {
      if (!(key in originalEnv)) delete process.env[key];
    }
    for (const [key, value] of Object.entries(originalEnv)) {
      if (value === undefined) delete process.env[key];
      else process.env[key] = value;
    }
    await rm(testHome, { recursive: true, force: true }).catch(() => {});
  });

  async function seedCache(identity: string, upn: string): Promise<void> {
    const dir = process.env.M365_AGENT_CLI_CONFIG_DIR as string;
    await mkdir(dir, { recursive: true });
    const h = Buffer.from(JSON.stringify({ alg: 'none', typ: 'JWT' })).toString('base64url');
    const p = Buffer.from(
      JSON.stringify({ exp: 2_000_000_000, appid: CLIENT_ID, upn, scp: GRAPH_CRITICAL_DELEGATED_SCOPES.join(' ') })
    ).toString('base64url');
    await writeFile(
      join(dir, `token-cache-${identity}.json`),
      JSON.stringify({
        version: 1,
        refreshToken: 'fake-refresh-token',
        graph: { accessToken: `${h}.${p}.x`, expiresAt: Date.now() + 3_600_000 }
      }),
      'utf8'
    );
  }

  test('blocks the command action when the signed-in identity does not match --require-identity', async () => {
    await seedCache('default', 'doris@lassfolk.net');
    const program = createM365Program();
    let ran = false;
    program.addCommand(
      new Command('__guard_probe__').action(() => {
        ran = true;
      })
    );
    program.exitOverride();

    const originalExit = process.exit;
    const originalError = console.error;
    let exitCode: number | undefined;
    const errors: string[] = [];
    process.exit = ((code?: number) => {
      exitCode = code;
      throw new Error(`exit ${code}`);
    }) as never;
    console.error = ((s: string) => errors.push(s)) as typeof console.error;
    try {
      await expect(
        program.parseAsync(['node', 'm365-agent-cli', '--require-identity', 'lotta@lassfolk.net', '__guard_probe__'])
      ).rejects.toThrow();
      expect(ran).toBe(false);
      expect(exitCode).toBe(1);
      expect(errors.join('\n')).toContain('Identity guard failed');
    } finally {
      process.exit = originalExit;
      console.error = originalError;
    }
  });

  test('allows the command action when the signed-in identity matches --require-identity', async () => {
    await seedCache('default', 'doris@lassfolk.net');
    const program = createM365Program();
    let ran = false;
    program.addCommand(
      new Command('__guard_probe_ok__').action(() => {
        ran = true;
      })
    );
    program.exitOverride();

    await program.parseAsync([
      'node',
      'm365-agent-cli',
      '--require-identity',
      'doris@lassfolk.net',
      '__guard_probe_ok__'
    ]);
    expect(ran).toBe(true);
  });

  test('does not run identity resolution when --require-identity/--as-delegate-of are absent', async () => {
    const program = createM365Program();
    let ran = false;
    program.addCommand(
      new Command('__guard_probe_off__').action(() => {
        ran = true;
      })
    );
    program.exitOverride();

    await program.parseAsync(['node', 'm365-agent-cli', '__guard_probe_off__']);
    expect(ran).toBe(true);
  });
});
