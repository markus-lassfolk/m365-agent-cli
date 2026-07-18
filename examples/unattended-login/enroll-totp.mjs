#!/usr/bin/env node
/**
 * Reference: enroll a software TOTP authenticator for a fresh M365 account, headlessly.
 *
 * Signs in with the account password, drives the Microsoft "Security info" wizard to add an
 * authenticator app, scrapes the base32 secret off the **Can't scan image?** screen, then activates
 * the method by entering a generated code. The captured seed is printed to STDOUT so your caller can
 * store it in a vault; everything else goes to stderr and the secret itself is never logged.
 *
 * This is EXAMPLE code to copy and adapt — it is NOT shipped or executed by the CLI.
 * Read docs/UNATTENDED_LOGIN.md first (the "Automated first-time TOTP enrollment" section) for the
 * hard constraints: it only works when the tenant permits self-service registration from this
 * host — no "require MFA to register security info" Conditional Access policy (that needs a TAP),
 * and the account has no MFA yet or you're adding another method. Selectors WILL drift; expect to
 * tune them against your tenant. Nothing is hardcoded — all secrets come from env at runtime.
 *
 * Required env:
 *   M365_EMAIL       the account UPN (e.g. agent@contoso.com)
 *   M365_PASSWORD    the account password
 * Optional env:
 *   M365_SECURITY_INFO_URL   default https://mysignins.microsoft.com/security-info
 *   M365_LOGIN_TIMEOUT_MS    per-step wait budget (default 20000)
 *
 * Setup:  npm install   (installs playwright + otplib)   then   npx playwright install chromium
 * Run:    M365_EMAIL=... M365_PASSWORD=... node enroll-totp.mjs
 *
 * Output (stdout, exactly one line on success), for the caller to capture into a secret store:
 *   {"totp_secret":"<BASE32>","account_name":"agent@contoso.com"}
 */
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { authenticator } from 'otplib';
import { chromium } from 'playwright';

function requireEnv(name) {
  const v = process.env[name]?.trim();
  if (!v) {
    console.error(`[enroll-totp] missing required env: ${name}`);
    process.exit(2);
  }
  return v;
}

const EMAIL = requireEnv('M365_EMAIL');
const PASSWORD = requireEnv('M365_PASSWORD');
const SECURITY_INFO_URL =
  process.env.M365_SECURITY_INFO_URL?.trim() || 'https://mysignins.microsoft.com/security-info';
const STEP_TIMEOUT = Number(process.env.M365_LOGIN_TIMEOUT_MS || '20000');

// Diagnostics go to stderr ONLY. The scraped seed is the one thing that goes to stdout (once).
// Never pass a secret value into step() — not the seed, not the password.
function step(msg) {
  console.error(`[enroll-totp] ${msg}`);
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

// Click the first selector that exists; returns true if something was clicked.
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

// Wait until any selector in the list appears, or resolve false after the step budget.
async function waitAny(page, selectors) {
  try {
    await page.waitForSelector(selectors.join(', '), { timeout: STEP_TIMEOUT });
    return true;
  } catch {
    return false;
  }
}

// Pull the base32 secret out of the revealed "Can't scan image?" panel. Microsoft renders it as a
// spaced, uppercase base32 blob next to a "Secret key" label, with the account name above it. We read
// the panel text and extract both. Kept selector-agnostic on purpose — the label text is far more
// stable than the DOM structure around it.
async function scrapeSeed(page) {
  const text = await page.evaluate(() => document.body.innerText);
  const seedMatch = text.match(/secret key[:\s]*([A-Z2-7][A-Z2-7 ]{14,})/i);
  if (!seedMatch) return null;
  const secret = seedMatch[1].replace(/\s+/g, '').toUpperCase();
  // Sanity-check it's plausible base32 and that otplib will accept it.
  if (!/^[A-Z2-7]{16,}$/.test(secret)) return null;
  try {
    authenticator.generate(secret);
  } catch {
    return null;
  }
  const nameMatch = text.match(/account name[:\s]*([^\s]+@[^\s]+)/i);
  return { secret, accountName: nameMatch ? nameMatch[1] : EMAIL };
}

async function main() {
  step('starting TOTP enrollment');
  const profileDir = mkdtempSync(join(tmpdir(), 'm365-enroll-totp-'));
  let success = false;
  let captured = null;

  const context = await chromium.launchPersistentContext(profileDir, { headless: true });
  const page = context.pages()[0] ?? (await context.newPage());

  try {
    // 1) Load the Security info page; this triggers a sign-in.
    await page.goto(SECURITY_INFO_URL, { waitUntil: 'domcontentloaded' });
    await logPage(page, 'loaded');

    // 2) Account + password (same page sequence as an ordinary sign-in).
    if (await page.locator('input[name="loginfmt"]').count()) {
      await page.fill('input[name="loginfmt"]', EMAIL);
      await clickAny(page, ['input[type="submit"]', 'button:has-text("Next")']);
    }
    const pw = await page.waitForSelector('input[type="password"]', { timeout: STEP_TIMEOUT });
    await pw.fill(PASSWORD);
    await clickAny(page, ['input[type="submit"]', 'button:has-text("Sign in")']);
    await logPage(page, 'password-submitted');

    // 3) "Stay signed in?" and any first-time "More information required" interstitial.
    await clickAny(page, ['input[type="submit"][value="Yes"]', 'button:has-text("Yes")']);
    await clickAny(page, ['input[type="submit"][value="Next"]', 'button:has-text("Next")']);

    // 4) Open the add-method wizard and choose "Authenticator app".
    //    On a first-time forced registration you may already be inside the wizard; the clicks below
    //    are no-ops if the element is absent.
    await clickAny(page, [
      'button:has-text("Add sign-in method")',
      'button:has-text("Add method")',
      'button:has-text("Add")'
    ]);
    // The method picker is a dropdown in the dialog; select the authenticator app option.
    await clickAny(page, ['[role="combobox"]', 'select', 'button:has-text("Choose a method")']);
    await clickAny(page, [
      'option:has-text("Authenticator app")',
      '[role="option"]:has-text("Authenticator app")',
      'text=Authenticator app'
    ]);
    await clickAny(page, ['button:has-text("Add")', 'button:has-text("Next")']);

    // 5) Skip "download the app" — take the "different authenticator app" branch, which is what
    //    exposes the manual key instead of forcing the Microsoft Authenticator push flow.
    await clickAny(page, [
      'button:has-text("I want to use a different authenticator app")',
      'a:has-text("I want to use a different authenticator app")'
    ]);
    await clickAny(page, ['button:has-text("Next")']);

    // 6) Reveal the manual key, then scrape it. Retry the reveal a few times while the panel renders.
    await waitAny(page, ['text=Can\'t scan', 'text=Set up your account', 'input[autocomplete="one-time-code"]']);
    for (let attempt = 0; attempt < 4 && !captured; attempt++) {
      await clickAny(page, [
        'button:has-text("Can\'t scan image?")',
        'a:has-text("Can\'t scan image?")',
        'button:has-text("Can\'t scan")',
        'text=Can\'t scan image?'
      ]);
      await page.waitForTimeout(1000);
      captured = await scrapeSeed(page);
    }
    if (!captured) throw new Error('could not read the secret key from the "Can\'t scan image?" panel');
    step('secret key captured');

    // 7) Advance to the verification step and activate with a generated code. Retry once on the next
    //    30-second window if the first code is rejected (clock/edge timing).
    await clickAny(page, ['button:has-text("Next")']);
    let activated = false;
    for (let attempt = 0; attempt < 2 && !activated; attempt++) {
      const otp = page.locator('input[name="otc"], input[autocomplete="one-time-code"], input[type="tel"]');
      if (!(await otp.count())) {
        // Verification field not shown yet — give the wizard a beat and re-check.
        await page.waitForTimeout(1500);
        if (!(await otp.count())) break;
      }
      await otp.first().fill(authenticator.generate(captured.secret));
      await clickAny(page, ['button:has-text("Next")', 'input[type="submit"]', 'button:has-text("Verify")']);
      await page.waitForTimeout(2500);
      activated = await page
        .waitForFunction(
          () => /was (successfully )?registered|success|you're all set|done/i.test(document.body.innerText),
          { timeout: 6000 }
        )
        .then(() => true)
        .catch(() => false);
      if (!activated && (await page.locator('input[name="otc"], input[autocomplete="one-time-code"]').count())) {
        step('verification code rejected; waiting for the next code window');
        await page.waitForTimeout(31000);
      }
    }
    if (!activated) throw new Error('captured the seed but could not confirm activation (method may be unverified)');

    success = true;
    step('authenticator registered and activated');
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

  if (success && captured) {
    // The ONE line of stdout: the seed for your caller to store in a vault. Capture it there —
    // do not echo it into a log file.
    process.stdout.write(`${JSON.stringify({ totp_secret: captured.secret, account_name: captured.accountName })}\n`);
  }
  process.exit(success ? 0 : 1);
}

main();
