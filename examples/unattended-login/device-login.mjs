#!/usr/bin/env node
/**
 * Reference: complete an m365-agent-cli device-code sign-in in a headless browser.
 *
 * This is EXAMPLE code to copy and adapt — it is NOT shipped or executed by the CLI.
 * Read docs/UNATTENDED_LOGIN.md first for the security tradeoffs and the full flow.
 * All secrets come from environment variables at runtime; nothing is hardcoded.
 *
 * Required env:
 *   M365_EMAIL         the account UPN (e.g. agent@contoso.com)
 *   M365_PASSWORD      the account password
 *   M365_TOTP_SECRET   base32 TOTP seed for the account's authenticator app
 *   M365_USER_CODE     the user_code from the CLI's `device_code` JSON event
 * Optional env:
 *   M365_VERIFICATION_URI           default https://microsoft.com/devicelogin
 *   M365_VERIFICATION_URI_COMPLETE  if set, opened directly (code pre-filled)
 *   M365_LOGIN_TIMEOUT_MS           per-step wait budget (default 20000)
 *
 * Setup:  npm install   (installs playwright + otplib)   then   npx playwright install chromium
 * Run:    M365_EMAIL=... M365_PASSWORD=... M365_TOTP_SECRET=... M365_USER_CODE=... node device-login.mjs
 */
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { authenticator } from 'otplib';
import { chromium } from 'playwright';

function requireEnv(name) {
  const v = process.env[name]?.trim();
  if (!v) {
    console.error(`[device-login] missing required env: ${name}`);
    process.exit(2);
  }
  return v;
}

const EMAIL = requireEnv('M365_EMAIL');
const PASSWORD = requireEnv('M365_PASSWORD');
const TOTP_SECRET = requireEnv('M365_TOTP_SECRET');
const USER_CODE = process.env.M365_USER_CODE?.trim() ?? '';
const VERIFICATION_URI = process.env.M365_VERIFICATION_URI?.trim() || 'https://microsoft.com/devicelogin';
const VERIFICATION_URI_COMPLETE = process.env.M365_VERIFICATION_URI_COMPLETE?.trim() || '';
const STEP_TIMEOUT = Number(process.env.M365_LOGIN_TIMEOUT_MS || '20000');

// Log to stderr only, and never a secret value (not even its length) — page context only.
function step(msg) {
  console.error(`[device-login] ${msg}`);
}

async function logPage(page, label) {
  let body = '';
  try {
    body = (await page.evaluate(() => document.body.innerText)).slice(0, 200).replace(/\s+/g, ' ');
  } catch {
    /* page mid-navigation */
  }
  step(`${label} :: ${page.url()} :: ${body}`);
}

async function clickAny(page, selectors) {
  for (const sel of selectors) {
    const loc = page.locator(sel);
    if (await loc.count()) {
      await loc
        .first()
        .click()
        .catch(() => {});
      return true;
    }
  }
  return false;
}

async function main() {
  step('starting device-code sign-in');
  const profileDir = mkdtempSync(join(tmpdir(), 'm365-device-login-'));
  let success = false;

  const context = await chromium.launchPersistentContext(profileDir, { headless: true });
  const page = context.pages()[0] ?? (await context.newPage());

  try {
    // 1) Device-code entry page.
    await page.goto(VERIFICATION_URI_COMPLETE || VERIFICATION_URI, { waitUntil: 'domcontentloaded' });
    await logPage(page, 'loaded');

    if (!VERIFICATION_URI_COMPLETE) {
      const otc = await page.waitForSelector('input[name="otc"], input#otc, input[type="text"]', {
        timeout: STEP_TIMEOUT
      });
      await otc.fill(USER_CODE);
      await clickAny(page, ['input[type="submit"]', 'button[type="submit"]', 'button:has-text("Next")']);
    }

    // 2) Account: fresh email field OR a "Pick an account" tile (from a reused profile).
    await page
      .waitForFunction(
        () =>
          document.body.innerText.includes('Pick an account') ||
          !!document.querySelector('input[name="loginfmt"]') ||
          !!document.querySelector('input[type="password"]'),
        { timeout: STEP_TIMEOUT }
      )
      .catch(() => {});
    await logPage(page, 'account');

    if (await page.locator('input[name="loginfmt"]').count()) {
      await page.fill('input[name="loginfmt"]', EMAIL);
      await clickAny(page, ['input[type="submit"]', 'button:has-text("Next")']);
    } else if (await page.getByText(EMAIL, { exact: false }).count()) {
      await page
        .getByText(EMAIL, { exact: false })
        .first()
        .click()
        .catch(() => {});
    }

    // 3) Password.
    const pw = await page.waitForSelector('input[type="password"]', { timeout: STEP_TIMEOUT });
    await pw.fill(PASSWORD);
    await clickAny(page, ['input[type="submit"]', 'button:has-text("Sign in")']);
    await logPage(page, 'password-submitted');

    // 4) App consent — can re-render 1-2 times; click Continue until it's gone.
    for (let i = 0; i < 3; i++) {
      const continued = await clickAny(page, [
        'input[type="submit"][value="Continue"]',
        'button:has-text("Continue")'
      ]);
      if (!continued) break;
      await page.waitForTimeout(1000);
    }

    // 5) TOTP, if prompted. Retry once on the next 30-second window.
    for (let attempt = 0; attempt < 2; attempt++) {
      const otp = page.locator('input[name="otc"], input[autocomplete="one-time-code"]');
      if (!(await otp.count())) break;
      await otp.first().fill(authenticator.generate(TOTP_SECRET));
      await clickAny(page, ['input[type="submit"]', 'button:has-text("Verify")', 'button:has-text("Next")']);
      await page.waitForTimeout(2000);
      if (!(await page.locator('input[name="otc"], input[autocomplete="one-time-code"]').count())) break;
      step('totp rejected; waiting for the next code window');
      await page.waitForTimeout(31000);
    }

    // 6) Optional "Stay signed in?".
    await clickAny(page, ['input[type="submit"][value="Yes"]', 'button:has-text("Yes")']);

    // 7) Success is ONLY the final confirmation text — never a URL substring.
    await page.waitForFunction(
      () => /you have signed in|you may now close this window/i.test(document.body.innerText),
      { timeout: STEP_TIMEOUT }
    );
    success = true;
    step('signed in successfully');
  } catch (err) {
    step(`failed: ${err?.message ?? err}`);
    try {
      const shot = join(profileDir, 'failure.png');
      await page.screenshot({ path: shot, fullPage: true });
      step(`screenshot saved to ${shot}`);
    } catch {
      /* ignore */
    }
  } finally {
    await context.close();
    if (success) rmSync(profileDir, { recursive: true, force: true });
  }

  process.exit(success ? 0 : 1);
}

main();
