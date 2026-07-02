import { afterEach, beforeEach, describe, expect, test } from 'bun:test';
import { homedir, tmpdir } from 'node:os';
import { join } from 'node:path';
import { getActiveEnvFilePath } from './active-env.js';

const KEYS = ['M365_AGENT_ENV_FILE'] as const;

describe('getActiveEnvFilePath', () => {
  const snapshot: Record<string, string | undefined> = {};

  beforeEach(() => {
    for (const k of KEYS) {
      snapshot[k] = process.env[k];
      delete process.env[k];
    }
  });
  afterEach(() => {
    for (const k of KEYS) {
      if (snapshot[k] === undefined) {
        delete process.env[k];
      } else {
        process.env[k] = snapshot[k];
      }
    }
  });

  test('returns the explicit caller-provided envPath when set (highest priority)', () => {
    const explicit = join(tmpdir(), 'cli-beta.env');
    expect(getActiveEnvFilePath(explicit)).toBe(explicit);
  });

  test('expands a tilde in the explicit envPath', () => {
    expect(getActiveEnvFilePath('~/m365-cli/.env.beta')).toBe(join(homedir(), 'm365-cli/.env.beta'));
  });

  test('falls back to M365_AGENT_ENV_FILE when no explicit envPath is provided', () => {
    process.env.M365_AGENT_ENV_FILE = join(tmpdir(), 'env-from-var.env');
    expect(getActiveEnvFilePath()).toBe(process.env.M365_AGENT_ENV_FILE);
  });

  test('falls back to the default global .env when neither is set', () => {
    expect(getActiveEnvFilePath()).toBe(join(homedir(), '.config', 'm365-agent-cli', '.env'));
  });

  test('explicit envPath wins over M365_AGENT_ENV_FILE', () => {
    process.env.M365_AGENT_ENV_FILE = join(tmpdir(), 'var-wins.env');
    const explicit = join(tmpdir(), 'explicit-wins.env');
    expect(getActiveEnvFilePath(explicit)).toBe(explicit);
  });

  test('whitespace-only explicit envPath falls through to M365_AGENT_ENV_FILE', () => {
    process.env.M365_AGENT_ENV_FILE = join(tmpdir(), 'fallthrough.env');
    expect(getActiveEnvFilePath('   ')).toBe(process.env.M365_AGENT_ENV_FILE);
  });

  test('whitespace-only M365_AGENT_ENV_FILE falls through to default', () => {
    process.env.M365_AGENT_ENV_FILE = '   ';
    expect(getActiveEnvFilePath()).toBe(join(homedir(), '.config', 'm365-agent-cli', '.env'));
  });
});
