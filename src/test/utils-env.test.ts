import { afterEach, describe, expect, test } from 'bun:test';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { homedir, tmpdir } from 'node:os';
import { join } from 'node:path';
import { applyEnvFileOverrides, resolveEnvFilePathArgument } from '../lib/utils.js';

describe('resolveEnvFilePathArgument', () => {
  test('expands ~/ to user home', () => {
    const p = resolveEnvFilePathArgument('~/foo/bar');
    expect(p).toBe(join(homedir(), 'foo', 'bar'));
  });

  test('empty or whitespace falls back to default .env path', () => {
    const expected = join(homedir(), '.config', 'm365-agent-cli', '.env');
    expect(resolveEnvFilePathArgument('')).toBe(expected);
    expect(resolveEnvFilePathArgument('   ')).toBe(expected);
  });

  test('tilde alone resolves to homedir', () => {
    expect(resolveEnvFilePathArgument('~')).toBe(homedir());
  });
});

describe('applyEnvFileOverrides', () => {
  const keysToRestore = ['EWS_CLIENT_ID', 'M365_TEST_QUOTED', 'COMMENTED_OUT'];

  afterEach(() => {
    for (const k of keysToRestore) {
      delete process.env[k];
    }
  });

  test('no-op when file does not exist', () => {
    process.env.EWS_CLIENT_ID = 'unchanged';
    applyEnvFileOverrides(join(tmpdir(), `m365-cli-missing-env-${Date.now()}`));
    expect(process.env.EWS_CLIENT_ID).toBe('unchanged');
  });

  test('overwrites existing process.env and strips quotes', () => {
    const dir = mkdtempSync(join(tmpdir(), 'm365-cli-env-'));
    const envPath = join(dir, '.env');
    try {
      process.env.EWS_CLIENT_ID = 'old-client';
      writeFileSync(
        envPath,
        ['EWS_CLIENT_ID=new-client', 'M365_TEST_QUOTED="with spaces"', '#COMMENTED_OUT=skip', ''].join('\n'),
        'utf8'
      );
      applyEnvFileOverrides(envPath);
      expect(process.env.EWS_CLIENT_ID).toBe('new-client');
      expect(process.env.M365_TEST_QUOTED).toBe('with spaces');
      expect(process.env.COMMENTED_OUT).toBeUndefined();
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });
});
