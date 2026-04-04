import { describe, expect, test } from 'bun:test';
import type { EventHint } from '@sentry/core';
import { glitchTipShouldSuppressHint } from './glitchtip-report-policy.js';

function hintWithEx(ex: unknown): EventHint {
  return { originalException: ex };
}

describe('glitchTipShouldSuppressHint', () => {
  test('does not suppress when reportAll is true', () => {
    const err = Object.assign(new Error('ECONNREFUSED'), { code: 'ECONNREFUSED' });
    expect(glitchTipShouldSuppressHint(hintWithEx(err), true)).toBe(false);
  });

  test('suppresses common network errno', () => {
    const err = Object.assign(new Error('connect failed'), { code: 'ECONNREFUSED' });
    expect(glitchTipShouldSuppressHint(hintWithEx(err), false)).toBe(true);
  });

  test('suppresses OAuth / AAD token message patterns', () => {
    expect(glitchTipShouldSuppressHint(hintWithEx(new Error('invalid_grant: token expired')), false)).toBe(true);
    expect(glitchTipShouldSuppressHint(hintWithEx(new Error('AADSTS700016: bad')), false)).toBe(true);
  });

  test('does not suppress arbitrary errors', () => {
    expect(glitchTipShouldSuppressHint(hintWithEx(new Error('Cannot read foo')), false)).toBe(false);
  });

  test('does not suppress when hint is undefined', () => {
    expect(glitchTipShouldSuppressHint(undefined, false)).toBe(false);
  });
});
