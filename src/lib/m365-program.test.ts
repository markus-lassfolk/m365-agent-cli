import { describe, expect, test } from 'bun:test';
import { Command } from 'commander';
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
