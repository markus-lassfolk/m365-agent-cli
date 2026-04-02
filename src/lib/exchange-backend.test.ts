import { afterEach, beforeEach, describe, expect, test } from 'bun:test';
import { DEFAULT_EXCHANGE_BACKEND, getExchangeBackend, mayUseEws, shouldTryGraphFirst } from './exchange-backend.js';

describe('exchange-backend', () => {
  let originalEnv: NodeJS.ProcessEnv;

  beforeEach(() => {
    originalEnv = { ...process.env };
  });

  afterEach(() => {
    process.env = originalEnv;
  });

  test('defaults to graph (dev_v2 pure Graph)', () => {
    delete process.env.M365_EXCHANGE_BACKEND;
    expect(getExchangeBackend()).toBe(DEFAULT_EXCHANGE_BACKEND);
    expect(getExchangeBackend()).toBe('graph');
  });

  test('respects M365_EXCHANGE_BACKEND', () => {
    process.env.M365_EXCHANGE_BACKEND = 'ews';
    expect(getExchangeBackend()).toBe('ews');
    process.env.M365_EXCHANGE_BACKEND = 'auto';
    expect(getExchangeBackend()).toBe('auto');
    process.env.M365_EXCHANGE_BACKEND = 'graph';
    expect(getExchangeBackend()).toBe('graph');
  });

  test('invalid value falls back to default', () => {
    process.env.M365_EXCHANGE_BACKEND = 'nope';
    expect(getExchangeBackend()).toBe('graph');
  });

  test('shouldTryGraphFirst / mayUseEws', () => {
    delete process.env.M365_EXCHANGE_BACKEND;
    expect(shouldTryGraphFirst()).toBe(true);
    expect(mayUseEws()).toBe(false);

    process.env.M365_EXCHANGE_BACKEND = 'auto';
    expect(shouldTryGraphFirst()).toBe(true);
    expect(mayUseEws()).toBe(true);

    process.env.M365_EXCHANGE_BACKEND = 'ews';
    expect(shouldTryGraphFirst()).toBe(false);
    expect(mayUseEws()).toBe(true);
  });
});
