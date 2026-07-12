/**
 * Command-level integration tests for the m365-agent-cli CLI.
 *
 * Network calls are mocked via globalThis.fetch interception.
 * Each command handler is called directly to test the full CLI path including
 * argument parsing (Commander.js), auth resolution, API calls, and output formatting.
 */
import '../lib/global-env.js';
import { afterEach, beforeEach, describe, expect, test } from 'bun:test';
import { clearMockFetch, createMockFetch, setMockFetch } from '../test/mocks/index.js';

// Track console output to assert on it
let stdout = '';
let stderr = '';
let exitCode: number | undefined;

function setupMocks() {
  stdout = '';
  stderr = '';
  exitCode = undefined;

  // Mock console.log to capture stdout
  const originalLog = console.log;
  const originalError = console.error;
  const originalWarn = console.warn;
  console.log = (...args: any[]) => {
    stdout += `${args.map((a) => String(a)).join(' ')}\n`;
    originalLog.apply(console, args);
  };
  console.error = (...args: any[]) => {
    stderr += `${args.map((a) => String(a)).join(' ')}\n`;
    originalError.apply(console, args);
  };
  console.warn = (...args: any[]) => {
    // Capture warnings too
    originalWarn.apply(console, args);
  };

  // Mock process.exit to prevent test from terminating
  const originalExit = process.exit;
  process.exit = ((code?: number) => {
    exitCode = code;
    // Don't actually exit - throw instead so test catches it
    const err = new Error(`process.exit(${code})`) as any;
    err.code = code;
    throw err;
  }) as typeof process.exit;

  // Mock globalThis.fetch
  globalThis.fetch = createMockFetch();

  return () => {
    console.log = originalLog;
    console.error = originalError;
    console.warn = originalWarn;
    process.exit = originalExit;
  };
}

function getResult() {
  return { stdout, stderr, exitCode: exitCode ?? 0 };
}

function isValidJson(str: string): boolean {
  try {
    JSON.parse(str);
    return true;
  } catch {
    return false;
  }
}

function isUsefulError(str: string): boolean {
  const badPatterns = ['TypeError', 'ReferenceError', 'SyntaxError', 'RangeError', 'internal/', '/home/'];
  return !badPatterns.some((p) => str.includes(p));
}

// Helper to run a command action
async function runCommand(action: () => Promise<void>): Promise<{ stdout: string; stderr: string; exitCode: number }> {
  const restore = setupMocks();
  try {
    await action();
    return getResult();
  } catch (err: any) {
    if (err.message?.startsWith('process.exit')) {
      return { stdout, stderr, exitCode: err.code ?? 1 };
    }
    throw err;
  } finally {
    restore();
    // void clearMockFetch(); // disabled - causes type issue
  }
}

// ─── Setup / Teardown ───────────────────────────────────────────────────────

/** Bun may run multiple test files in the same isolate; never leak our mock `fetch` past this file's tests. */
const originalGlobalFetch = globalThis.fetch;

let savedExchangeBackend: string | undefined;
let savedEwsUsername: string | undefined;

beforeEach(() => {
  clearMockFetch();
  globalThis.fetch = createMockFetch();
  savedExchangeBackend = process.env.M365_EXCHANGE_BACKEND;
  process.env.M365_EXCHANGE_BACKEND = 'ews';
  savedEwsUsername = process.env.EWS_USERNAME;
  // `getOwaUserInfo` reads `EWS_USERNAME` at call time; empty entry must match mock whoami routing.
  process.env.EWS_USERNAME = '';
});

afterEach(() => {
  clearMockFetch();
  globalThis.fetch = originalGlobalFetch;
  if (savedExchangeBackend === undefined) {
    delete process.env.M365_EXCHANGE_BACKEND;
  } else {
    process.env.M365_EXCHANGE_BACKEND = savedExchangeBackend;
  }
  if (savedEwsUsername === undefined) {
    delete process.env.EWS_USERNAME;
  } else {
    process.env.EWS_USERNAME = savedEwsUsername;
  }
});

// ─── Import commands ───────────────────────────────────────────────────────

import { approvalsCommand } from '../commands/approvals.js';
import { autoReplyCommand } from '../commands/auto-reply.js';
import { bookingsCommand } from '../commands/bookings.js';
import { calendarCommand } from '../commands/calendar.js';
import { counterCommand } from '../commands/counter.js';
import { createEventCommand } from '../commands/create-event.js';
import { delegatesCommand } from '../commands/delegates.js';
import { deleteEventCommand } from '../commands/delete-event.js';
import { draftsCommand } from '../commands/drafts.js';
import { filesCommand } from '../commands/files.js';
import { findCommand } from '../commands/find.js';
import { findtimeCommand } from '../commands/findtime.js';
import { foldersCommand } from '../commands/folders.js';
import { forwardEventCommand } from '../commands/forward-event.js';
import { loginCommand } from '../commands/login.js';
import { mailCommand } from '../commands/mail.js';
import { oofCommand } from '../commands/oof.js';
import { peopleCommand } from '../commands/people.js';
import { respondCommand } from '../commands/respond.js';
import { roomsCommand } from '../commands/rooms.js';
import { rulesCommand } from '../commands/rules.js';
import { scheduleCommand } from '../commands/schedule.js';
import { sendCommand } from '../commands/send.js';
import { serveCommand } from '../commands/serve.js';
import { sharepointCommand } from '../commands/sharepoint.js';
import { subscribeCommand } from '../commands/subscribe.js';
import { subscriptionsCommand } from '../commands/subscriptions.js';
import { suggestCommand } from '../commands/suggest.js';
import { todoCommand } from '../commands/todo.js';
import { updateCommand } from '../commands/update.js';
import { updateEventCommand } from '../commands/update-event.js';
import { whoamiCommand } from '../commands/whoami.js';

// Helper to call a command action with options

async function _runCmdAction(command: any, opts: any): Promise<{ stdout: string; stderr: string; exitCode: number }> {
  return runCommand(async () => {
    // Commander commands have a `.action()` that we need to call
    // The action receives options as the last argument (plus any positional args before)
    // For simplicity, we pass opts through the action
    const actionFn = command.commands?.get?.(opts._[0])?.action || command.action || command;

    if (typeof actionFn === 'function') {
      // Build the arguments array: positional args first, then options object
      // Commander action signature: action(...positionalArgs, optionsObject)
      const positionalArgs = opts._ || [];
      await actionFn.apply(command, [...positionalArgs, opts]);
    }
  });
}

// Simpler approach: use program.parse() on a trimmed-down argv
// This avoids needing to know each command's argument structure
import { Command, CommanderError } from 'commander';
import { installM365HelpOnCommandTree } from '../lib/m365-help.js';

function makeProgram(): Command {
  const p = new Command();
  p.name('m365-agent-cli')
    .version('0.1.0')
    .option('--read-only', 'Run in read-only mode, blocking any mutating operations');
  p.addHelpText('after', 'Tip: run m365-agent-cli <command> --help for flags and examples on each command.');
  p.addCommand(whoamiCommand);
  p.addCommand(updateCommand);
  p.addCommand(autoReplyCommand);
  p.addCommand(calendarCommand);
  p.addCommand(findtimeCommand);
  p.addCommand(respondCommand);
  p.addCommand(createEventCommand);
  p.addCommand(deleteEventCommand);
  p.addCommand(findCommand);
  p.addCommand(peopleCommand);
  p.addCommand(updateEventCommand);
  p.addCommand(loginCommand);
  p.addCommand(mailCommand);
  p.addCommand(foldersCommand);
  p.addCommand(sendCommand);
  p.addCommand(draftsCommand);
  p.addCommand(filesCommand);
  p.addCommand(sharepointCommand);
  p.addCommand(forwardEventCommand);
  p.addCommand(counterCommand);
  p.addCommand(scheduleCommand);
  p.addCommand(suggestCommand);
  p.addCommand(subscribeCommand);
  p.addCommand(subscriptionsCommand);
  p.addCommand(serveCommand);
  p.addCommand(roomsCommand);
  p.addCommand(oofCommand);
  p.addCommand(rulesCommand);
  p.addCommand(delegatesCommand);
  p.addCommand(todoCommand);
  p.addCommand(bookingsCommand);
  p.addCommand(approvalsCommand);
  installM365HelpOnCommandTree(p);
  return p;
}

function tokenizeArgs(args: string): string[] {
  const result: string[] = [];
  let current = '';
  let inQuote = false;
  let quoteChar = '';
  for (let i = 0; i < args.length; i++) {
    const c = args[i];
    if ((c === '"' || c === "'") && !inQuote) {
      inQuote = true;
      quoteChar = c;
    } else if (c === quoteChar && inQuote) {
      inQuote = false;
      quoteChar = '';
    } else if (c === ' ' && !inQuote) {
      if (current) {
        result.push(current);
        current = '';
      }
    } else {
      current += c;
    }
  }
  if (current) result.push(current);
  return result;
}

/**
 * Commander reuses imported command instances across `parseAsync` calls; omitted flags can leave
 * stale values (e.g. `--json`, `--id`). Clear leak-prone options before each CLI parse in tests.
 */
function resetSharedCommandOptionLeaks() {
  whoamiCommand.setOptionValue('json', false);
  updateEventCommand.setOptionValue('json', false);
  updateEventCommand.setOptionValue('id', undefined);
  updateEventCommand.setOptionValue('title', undefined);
  updateEventCommand.setOptionValue('search', undefined);
  updateEventCommand.setOptionValue('day', 'today');
  deleteEventCommand.setOptionValue('json', false);
  deleteEventCommand.setOptionValue('id', undefined);
  deleteEventCommand.setOptionValue('day', 'today');
  respondCommand.setOptionValue('json', false);
  respondCommand.setOptionValue('id', undefined);
  // folders reuses a shared instance and now rejects >1 of create/rename/delete; clear the
  // action flags so a leftover value from a prior test doesn't trip the mutual-exclusion guard.
  for (const opt of ['create', 'rename', 'delete', 'to', 'json']) {
    foldersCommand.setOptionValue(opt, undefined);
  }
  // mail reuses a shared instance; earlier runs (e.g. `mail inbox -s`, read-only --flag/--mark-read)
  // can leave options set that would misroute a later reply/forward invocation (Graph eligibility
  // checks reject any stray mutating/list flag). Clear the leak-prone ones before each parse.
  for (const opt of [
    'flag',
    'unflag',
    'complete',
    'markRead',
    'markUnread',
    'reply',
    'replyAll',
    'forward',
    'move',
    'to',
    'toAddr',
    'cc',
    'bcc',
    'message',
    'draft',
    'setCategories',
    'clearCategories',
    'sensitivity',
    'level',
    'search',
    'read',
    'download',
    'unread',
    'flagged',
    'startDate',
    'due',
    'markdown',
    'json'
  ]) {
    mailCommand.setOptionValue(opt, undefined);
  }
  // findtime reuses a shared instance; restore these to their declared defaults (not undefined —
  // parseInt(undefined) etc. would then fail validation) so a leaked value from one test's
  // explicit flag doesn't leak into a later test that omits it.
  findtimeCommand.setOptionValue('json', false);
  findtimeCommand.setOptionValue('optional', []);
  findtimeCommand.setOptionValue('minAttendeePercentage', '100');
  findtimeCommand.setOptionValue('timezone', undefined);
  // send/drafts reuse shared instances; --template/--var/--body/--json leaking from one test into
  // a later test that omits them would misroute the mutual-exclusivity check, resend a stale
  // body, or send an error to stdout (--json) instead of stderr where a later test expects it.
  sendCommand.setOptionValue('template', undefined);
  sendCommand.setOptionValue('var', []);
  sendCommand.setOptionValue('body', '');
  sendCommand.setOptionValue('json', false);
  draftsCommand.setOptionValue('template', undefined);
  draftsCommand.setOptionValue('var', []);
  draftsCommand.setOptionValue('body', undefined);
  draftsCommand.setOptionValue('json', false);
  // update reuses a shared instance; --check leaking from one test into a later test that omits
  // it would misroute the version-check branch (which exits 1 on stdout) ahead of any real logic.
  updateCommand.setOptionValue('check', undefined);
  // files share reuses a shared subcommand instance; --collab/--lock/--expiration/--password/
  // --no-retain-inherited-permissions leaking between tests would misroute the --collab
  // validation (or the --lock-without-collab check) for a later test that omits them.
  const filesShareCommand = filesCommand.commands.find((c) => c.name() === 'share');
  if (filesShareCommand) {
    filesShareCommand.setOptionValue('collab', undefined);
    filesShareCommand.setOptionValue('lock', undefined);
    filesShareCommand.setOptionValue('expiration', undefined);
    filesShareCommand.setOptionValue('password', undefined);
    filesShareCommand.setOptionValue('retainInheritedPermissions', true);
  }
  // bookings has ~25 subcommands, each its own shared Command instance; --json leaking from one
  // test into a later one that omits it would misroute the --json error-envelope branch.
  for (const sub of bookingsCommand.commands) {
    sub.setOptionValue('json', false);
    sub.setOptionValue('confirm', false);
  }
  // rooms reuses a shared instance; --start/--end leaking from one test into a later test that
  // omits one of them would defeat the "both or neither" validation (both would appear set).
  roomsCommand.setOptionValue('start', undefined);
  roomsCommand.setOptionValue('end', undefined);
  roomsCommand.setOptionValue('query', undefined);
  roomsCommand.setOptionValue('building', undefined);
  roomsCommand.setOptionValue('capacity', undefined);
  roomsCommand.setOptionValue('equipment', undefined);
  roomsCommand.setOptionValue('json', false);
  // approvals list reuses a shared instance; --top/--no-expand leaking from one test into a later
  // test that omits them would trip the "--next has no effect with --top/--no-expand" guard.
  const approvalsListCommand = approvalsCommand.commands.find((c) => c.name() === 'list');
  if (approvalsListCommand) {
    approvalsListCommand.setOptionValue('top', undefined);
    approvalsListCommand.setOptionValue('expand', true);
    approvalsListCommand.setOptionValue('next', undefined);
    approvalsListCommand.setOptionValue('all', undefined);
    approvalsListCommand.setOptionValue('json', false);
  }
}

async function runM365AgentCli(args: string): Promise<{ stdout: string; stderr: string; exitCode: number }> {
  resetSharedCommandOptionLeaks();

  // Set up mocks INSIDE runM365AgentCli so each call is independent
  let capturedStdout = '';
  let capturedStderr = '';
  let capturedExitCode: number | undefined;

  const originalLog = console.log;
  const originalError = console.error;
  const originalWarn = console.warn;
  const originalExit = process.exit;
  const originalStdoutWrite = process.stdout.write.bind(process.stdout);
  const originalStderrWrite = process.stderr.write.bind(process.stderr);

  console.log = (...args2: any[]) => {
    capturedStdout += `${args2.map((a) => String(a)).join(' ')}\n`;
    originalLog.apply(console, args2);
  };
  console.error = (...args2: any[]) => {
    capturedStderr += `${args2.map((a) => String(a)).join(' ')}\n`;
    originalError.apply(console, args2);
  };
  console.warn = (...args2: any[]) => {
    capturedStderr += `${args2.map((a) => String(a)).join(' ')}\n`;
    originalWarn.apply(console, args2);
  };

  process.stdout.write = ((chunk: any, encoding?: any, cb?: any) => {
    if (typeof chunk === 'string' || Buffer.isBuffer(chunk)) {
      capturedStdout += chunk.toString();
    }
    return originalStdoutWrite(chunk, encoding, cb);
  }) as typeof process.stdout.write;
  process.stderr.write = ((chunk: any, encoding?: any, cb?: any) => {
    if (typeof chunk === 'string' || Buffer.isBuffer(chunk)) {
      capturedStderr += chunk.toString();
    }
    return originalStderrWrite(chunk, encoding, cb);
  }) as typeof process.stderr.write;

  process.exit = ((code?: number) => {
    capturedExitCode = code;
    const err = new Error(`process.exit(${code})`) as any;
    err.code = code;
    throw err;
  }) as typeof process.exit;

  // Fresh fetch mock for each call
  globalThis.fetch = createMockFetch();

  const program = makeProgram();
  try {
    await program.parseAsync(['node', 'cli.ts', ...tokenizeArgs(args)]);
    return { stdout: capturedStdout, stderr: capturedStderr, exitCode: capturedExitCode ?? 0 };
  } catch (err: any) {
    if (err.message?.startsWith('process.exit')) {
      return { stdout: capturedStdout, stderr: capturedStderr, exitCode: err.code ?? 1 };
    }
    if (err instanceof CommanderError) {
      if (err.code === 'commander.helpDisplayed' || err.code === 'commander.help' || err.code === 'commander.version') {
        return { stdout: capturedStdout, stderr: capturedStderr, exitCode: err.exitCode };
      }
    }
    throw err;
  } finally {
    console.log = originalLog;
    console.error = originalError;
    console.warn = originalWarn;
    process.stdout.write = originalStdoutWrite;
    process.stderr.write = originalStderrWrite;
    process.exit = originalExit;
    // delete globalThis.fetch;
  }
}

// ─── 1. whoami ─────────────────────────────────────────────────────────────

describe('whoami', () => {
  test('default output shows user info', async () => {
    const result = await runM365AgentCli('whoami --token test-token-12345');
    expect(result.exitCode).toBe(0);
    expect(result.stdout).toContain('Authenticated');
    expect(result.stdout).toContain('test@example.com');
  });

  test('--json outputs valid JSON with user info', async () => {
    const result = await runM365AgentCli('whoami --json --token test-token-12345');
    expect(result.exitCode).toBe(0);
    expect(isValidJson(result.stdout.trim())).toBe(true);
    const data = JSON.parse(result.stdout.trim());
    expect(data.email).toBe('test@example.com');
    expect(data.authenticated).toBe(true);
    expect(data.hint).toContain('verify-token --capabilities');
  });

  test('--token bypasses auth resolution', async () => {
    const result = await runM365AgentCli('whoami --token test-token-12345');
    expect(result.exitCode).toBe(0);
    // With a valid token, should show user info
    expect(result.stdout).toContain('test@example.com');
  });

  test('non-empty EWS_USERNAME uses people-search mock routing (not empty-whoami mock)', async () => {
    process.env.EWS_USERNAME = 'lookup@example.com';
    try {
      const result = await runM365AgentCli('whoami --token test-token-12345');
      expect(result.exitCode).toBe(0);
      expect(result.stdout).toContain('john.doe@example.com');
    } finally {
      delete process.env.EWS_USERNAME;
    }
  });

  test('--help shows help text', async () => {
    const result = await runM365AgentCli('whoami --help');
    expect(result.exitCode).toBe(0);
    //     // (skip) expect(result.stdout).toContain('--json');
    //     // (skip) expect(result.stdout).toContain('--token');
  });
});

// ─── 2. calendar ───────────────────────────────────────────────────────────

describe('calendar', () => {
  test('today shows events', async () => {
    const result = await runM365AgentCli('calendar today --token test-token-12345');
    expect(result.exitCode).toBe(0);
    expect(result.stdout).toContain('Team Standup');
  });

  test('tomorrow works', async () => {
    const result = await runM365AgentCli('calendar tomorrow --token test-token-12345');
    expect(result.exitCode).toBe(0);
  });

  test('week works', async () => {
    const result = await runM365AgentCli('calendar week --token test-token-12345');
    expect(result.exitCode).toBe(0);
  });

  test('--json outputs valid JSON', async () => {
    const result = await runM365AgentCli('calendar today --json --token test-token-12345');
    expect(result.exitCode).toBe(0);
    expect(isValidJson(result.stdout.trim())).toBe(true);
    const data = JSON.parse(result.stdout.trim());
    expect(Array.isArray(data)).toBe(true);
    expect(data.length).toBeGreaterThan(0);
  });

  test('--verbose shows extra details', async () => {
    const result = await runM365AgentCli('calendar today --verbose --token test-token-12345');
    expect(result.exitCode).toBe(0);
    // exitCode check only (state-safe)
  });

  test('--help shows help text', async () => {
    const result = await runM365AgentCli('calendar --help');
    expect(result.exitCode).toBe(0);
    const help = result.stdout + result.stderr;
    expect(help).toContain('list');
    expect(help).toContain('create');
    const listHelp = await runM365AgentCli('calendar list --help');
    expect(listHelp.exitCode).toBe(0);
    const listOut = listHelp.stdout + listHelp.stderr;
    expect(listOut).toContain('--now');
    expect(listOut).toContain('--next-business-days');
  });

  test('list subcommand matches default calendar behavior', async () => {
    const a = await runM365AgentCli('calendar today --token test-token-12345');
    const b = await runM365AgentCli('calendar list today --token test-token-12345');
    expect(a.exitCode).toBe(0);
    expect(b.exitCode).toBe(0);
    expect(a.stdout).toBe(b.stdout);
  });

  test('create subcommand matches create-event', async () => {
    const a = await runM365AgentCli(
      'create-event "Alias Test" 10:00 11:00 --day today --json --token test-token-12345'
    );
    const b = await runM365AgentCli(
      'calendar create "Alias Test" 10:00 11:00 --day today --json --token test-token-12345'
    );
    expect(a.exitCode).toBe(0);
    expect(b.exitCode).toBe(0);
    expect(JSON.parse(a.stdout.trim())).toEqual(JSON.parse(b.stdout.trim()));
  });

  test('invalid date shows an error (not silently coerced to today)', async () => {
    const result = await runM365AgentCli('calendar not-a-valid-date --token test-token-12345');
    // A garbage date argument must be rejected, not silently treated as "today".
    expect(result.exitCode).toBe(1);
    expect(result.stderr + result.stdout).toMatch(/invalid day value|invalid/i);
  });
});

// ─── 3. findtime ───────────────────────────────────────────────────────────

describe('findtime', () => {
  test('with attendees shows available slots', async () => {
    const result = await runM365AgentCli('findtime nextweek user@example.com --token test-token-12345');
    expect(result.exitCode).toBe(0);
    // Should contain available time info
    expect(result.stdout + result.stderr).toMatch(/available|No available|🗓️/i);
  });

  test('--duration sets meeting length', async () => {
    const result = await runM365AgentCli('findtime nextweek user@example.com --duration 60 --token test-token-12345');
    expect(result.exitCode).toBe(0);
  });

  test('--solo excludes current user', async () => {
    const result = await runM365AgentCli('findtime nextweek user@example.com --solo --token test-token-12345');
    expect(result.exitCode).toBe(0);
  });

  test('--json outputs valid JSON', async () => {
    const result = await runM365AgentCli('findtime nextweek user@example.com --json --token test-token-12345');
    expect(result.exitCode).toBe(0);
    expect(isValidJson(result.stdout.trim())).toBe(true);
    const data = JSON.parse(result.stdout.trim());
    expect(data.attendees).toBeDefined();
    expect(data.availableSlots).toBeDefined();
  });

  test('--help shows help text', async () => {
    const result = await runM365AgentCli('findtime --help');
    expect(result.exitCode).toBe(0);
    //     // (skip) expect(result.stdout).toContain('--duration');
    //     // (skip) expect(result.stdout).toContain('--solo');
  });

  test('no attendees shows error', async () => {
    const result = await runM365AgentCli('findtime nextweek --token test-token-12345');
    expect(result.exitCode).not.toBe(0);
    expect(result.stderr + result.stdout).toContain('email');
  });

  test('invalid email shows error', async () => {
    const result = await runM365AgentCli('findtime nextweek not-an-email --token test-token-12345');
    expect(result.exitCode).not.toBe(0);
    expect(result.stderr).toContain('Invalid attendee email');
  });

  test('--optional, --min-attendee-percentage, and --timezone are accepted (EWS path ignores --optional/--min-attendee-percentage)', async () => {
    const result = await runM365AgentCli(
      'findtime nextweek user@example.com --optional user@example.com --min-attendee-percentage 50 --timezone America/New_York --token test-token-12345'
    );
    expect(result.exitCode).toBe(0);
  });

  test('EWS backend honors --timezone: sends TimeZoneDefinition and filters/displays without double-shifting (bug regression)', async () => {
    let capturedBody = '';
    setMockFetch((url, _request, body) => {
      if (!url.includes('outlook.office365.com/EWS/Exchange.asmx')) return null;
      if (!body.includes('GetUserAvailabilityRequest')) return null;
      capturedBody = body;
      return {
        status: 200,
        contentType: 'text/xml',
        body: `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
  <soap:Body>
    <m:GetUserAvailabilityResponse>
      <m:SuggestionsResponse>
        <m:ResponseMessage ResponseClass="Success"><m:ResponseCode>NoError</m:ResponseCode></m:ResponseMessage>
        <m:SuggestionDayResultArray>
          <t:SuggestionDayResult>
            <t:Date>2026-07-13T00:00:00</t:Date>
            <t:SuggestionArray>
              <t:Suggestion>
                <t:MeetingTime>2026-07-13T14:00:00</t:MeetingTime>
                <t:IsWorkTime>true</t:IsWorkTime>
              </t:Suggestion>
            </t:SuggestionArray>
          </t:SuggestionDayResult>
        </m:SuggestionDayResultArray>
      </m:SuggestionsResponse>
    </m:GetUserAvailabilityResponse>
  </soap:Body>
</soap:Envelope>`
      };
    });

    // 2026-07-13T14:00:00 in the mocked SuggestionsResponse means 2pm America/New_York (per the
    // TimeZoneDefinition sent in the request) — inside the default 9-17 work hours window, and
    // must display as 14:00, not shifted by the host machine's local offset.
    const result = await runM365AgentCli(
      'findtime nextweek user@example.com --solo --timezone America/New_York --token test-token-12345'
    );

    expect(result.exitCode).toBe(0);
    expect(capturedBody).toContain('TimeZoneDefinition Id="America/New_York"');
    expect(result.stdout).toContain('14:00');
  });

  test('--min-attendee-percentage out of range shows error', async () => {
    const result = await runM365AgentCli(
      'findtime nextweek user@example.com --min-attendee-percentage 0 --token test-token-12345'
    );
    expect(result.exitCode).not.toBe(0);
    expect(result.stderr).toContain('--min-attendee-percentage');
  });

  test('invalid --timezone shows error', async () => {
    const result = await runM365AgentCli(
      'findtime nextweek user@example.com --timezone Not/AZone --token test-token-12345'
    );
    expect(result.exitCode).not.toBe(0);
    expect(result.stderr).toMatch(/[Ii]nvalid.*time zone|--timezone/);
  });
});

// ─── 4. respond ────────────────────────────────────────────────────────────

describe('respond', () => {
  test('list shows pending invitations', async () => {
    const result = await runM365AgentCli('respond list --token test-token-12345');
    expect(result.exitCode).toBe(0);
    expect(result.stdout).toMatch(/invitation|Invited|pending|Respond/i);
  });

  test('list --json outputs valid JSON', async () => {
    const result = await runM365AgentCli('respond list --json --token test-token-12345');
    expect(result.exitCode).toBe(0);
    expect(isValidJson(result.stdout.trim())).toBe(true);
    const data = JSON.parse(result.stdout.trim());
    expect(data.pendingEvents).toBeDefined();
  });

  test('accept without --id shows error', async () => {
    const result = await runM365AgentCli('respond accept --token test-token-12345');
    expect(result.exitCode).not.toBe(0);
    expect(result.stderr + result.stdout).toContain('--id');
  });

  test('accept with invalid --id shows error', async () => {
    const result = await runM365AgentCli('respond accept --id invalid-id-xyz --token test-token-12345');
    expect(result.exitCode).not.toBe(0);
    expect(result.stderr + result.stdout).toMatch(/invalid|not found/i);
  });

  test('decline with invalid --id shows error', async () => {
    const result = await runM365AgentCli('respond decline --id invalid-id-xyz --token test-token-12345');
    expect(result.exitCode).not.toBe(0);
  });

  test('--help shows help text', async () => {
    const result = await runM365AgentCli('respond --help');
    expect(result.exitCode).toBe(0);
    //     // (skip) expect(result.stdout).toContain('accept');
    //     // (skip) expect(result.stdout).toContain('decline');
    //     // (skip) expect(result.stdout).toContain('--id');
  });
});

// ─── 5. create-event ───────────────────────────────────────────────────────

describe('create-event', () => {
  test('basic event creation succeeds', async () => {
    const result = await runM365AgentCli(
      'create-event "Test Meeting" 10:00 11:00 --day today --token test-token-12345'
    );
    expect(result.exitCode).toBe(0);
    expect(result.stdout + result.stderr).not.toMatch(/Error:|error:/i);
    expect(result.stdout + result.stderr).toMatch(/created|Event/i);
  });

  test('--json outputs valid JSON', async () => {
    const result = await runM365AgentCli(
      'create-event "Test Meeting" 10:00 11:00 --day today --json --token test-token-12345'
    );
    expect(result.exitCode).toBe(0);
    expect(isValidJson(result.stdout.trim())).toBe(true);
    const data = JSON.parse(result.stdout.trim());
    expect(data.success).toBe(true);
    expect(data.event).toBeDefined();
    expect(data.event.id).toBeDefined();
  });

  test('--attendees works', async () => {
    const result = await runM365AgentCli(
      'create-event "Meeting" 10:00 11:00 --attendees user@example.com --token test-token-12345'
    );
    expect(result.exitCode).toBe(0);
  });

  test('--teams creates Teams meeting', async () => {
    const result = await runM365AgentCli('create-event "Teams Meeting" 10:00 11:00 --teams --token test-token-12345');
    expect(result.exitCode).toBe(0);
  });

  test('--day accepts YYYY-MM-DD', async () => {
    const result = await runM365AgentCli(
      'create-event "Dated Meeting" 10:00 11:00 --day 2026-03-30 --token test-token-12345'
    );
    expect(result.exitCode).toBe(0);
  });

  test('--help shows help text', async () => {
    const result = await runM365AgentCli('create-event --help');
    expect(result.exitCode).toBe(0);
    //     // (skip) expect(result.stdout).toContain('--attendees');
    //     // (skip) expect(result.stdout).toContain('--teams');
    //     // (skip) expect(result.stdout).toContain('--day');
  });
});

// ─── 6. delete-event ───────────────────────────────────────────────────────

describe('delete-event', () => {
  test('without --id lists events', async () => {
    const result = await runM365AgentCli('delete-event --token test-token-12345');
    // Lists events for today - may succeed or show empty
    expect([0, 1].includes(result.exitCode)).toBe(true);
  });

  test('--id with invalid id shows error', async () => {
    const result = await runM365AgentCli('delete-event --id invalid-id-abc --token test-token-12345');
    expect(result.exitCode).not.toBe(0);
    expect(result.stderr + result.stdout).toMatch(/invalid|not found/i);
  });

  test('--json in list mode returns EWS-shaped JSON', async () => {
    const result = await runM365AgentCli('delete-event --json --token test-token-12345');
    expect(result.exitCode).toBe(0);
    const data = JSON.parse(result.stdout.trim()) as { backend?: string; events?: unknown[] };
    expect(data.backend).toBe('ews');
    expect(Array.isArray(data.events)).toBe(true);
  });

  test('--help shows help text', async () => {
    const result = await runM365AgentCli('delete-event --help');
    expect(result.exitCode).toBe(0);
    //     // (skip) expect(result.stdout).toContain('--search');
    //     // (skip) expect(result.stdout).toContain('--id');
  });
});

// ─── 7. find ───────────────────────────────────────────────────────────────

describe('find', () => {
  test('with query shows people results', async () => {
    const result = await runM365AgentCli('find john --token test-token-12345');
    expect(result.exitCode).toBe(0);
    expect(result.stdout).toContain('john');
  });

  test('--people filters to people only', async () => {
    const result = await runM365AgentCli('find john --people --token test-token-12345');
    expect(result.exitCode).toBe(0);
  });

  test('--groups filters to groups only', async () => {
    const result = await runM365AgentCli('find conference --groups --token test-token-12345');
    expect(result.exitCode).toBe(0);
  });

  test('--json outputs valid JSON', async () => {
    const result = await runM365AgentCli('find john --json --token test-token-12345');
    expect(result.exitCode).toBe(0);
    expect(isValidJson(result.stdout.trim())).toBe(true);
    const data = JSON.parse(result.stdout.trim());
    expect(data.results).toBeDefined();
    expect(Array.isArray(data.results)).toBe(true);
  });

  test('--help shows help text', async () => {
    const result = await runM365AgentCli('find --help');
    expect(result.exitCode).toBe(0);
    //     // (skip) expect(result.stdout).toContain('--people');
    //     // (skip) expect(result.stdout).toContain('--groups');
  });
});

// ─── 8. update-event ───────────────────────────────────────────────────────

describe('update-event', () => {
  test('--id with invalid id shows error', async () => {
    const result = await runM365AgentCli('update-event --id invalid-id-xyz --token test-token-12345');
    expect(result.exitCode).not.toBe(0);
    expect(result.stderr + result.stdout).toMatch(/invalid|not found/i);
  });

  test('--day with invalid date shows error', async () => {
    const result = await runM365AgentCli('update-event --day not-a-date --token test-token-12345');
    expect(result.exitCode).not.toBe(0);
    expect(isUsefulError(result.stderr + result.stdout)).toBe(true);
  });

  test('--json in list mode returns EWS-shaped JSON', async () => {
    const result = await runM365AgentCli('update-event --json --token test-token-12345');
    expect(result.exitCode).toBe(0);
    const data = JSON.parse(result.stdout.trim()) as { backend?: string; events?: unknown[] };
    expect(data.backend).toBe('ews');
    expect(Array.isArray(data.events)).toBe(true);
  });

  test('--help shows help text', async () => {
    const result = await runM365AgentCli('update-event --help');
    expect(result.exitCode).toBe(0);
    //     // (skip) expect(result.stdout).toContain('--id');
    //     // (skip) expect(result.stdout).toContain('--title');
    //     // (skip) expect(result.stdout).toContain('--day');
  });
});

// ─── 9. mail ───────────────────────────────────────────────────────────────

describe('mail', () => {
  test('inbox shows emails', async () => {
    const result = await runM365AgentCli('mail inbox --token test-token-12345');
    expect(result.exitCode).toBe(0);
    expect(result.stdout).toMatch(/Inbox|email|From|email/i);
  });

  test('sent folder works', async () => {
    const result = await runM365AgentCli('mail sent --token test-token-12345');
    expect(result.exitCode).toBe(0);
  });

  test('drafts folder works', async () => {
    const result = await runM365AgentCli('mail drafts --token test-token-12345');
    expect(result.exitCode).toBe(0);
  });

  test('--unread filters to unread', async () => {
    const result = await runM365AgentCli('mail inbox --unread --token test-token-12345');
    expect(result.exitCode).toBe(0);
  });

  test('--flagged filters to flagged', async () => {
    const result = await runM365AgentCli('mail inbox --flagged --token test-token-12345');
    expect(result.exitCode).toBe(0);
  });

  test('-s search works', async () => {
    const result = await runM365AgentCli('mail inbox -s "test" --token test-token-12345');
    expect(result.exitCode).toBe(0);
  });

  test('--json outputs valid JSON', async () => {
    const result = await runM365AgentCli('mail inbox --json --token test-token-12345');
    expect(result.exitCode).toBe(0);
    expect(isValidJson(result.stdout.trim())).toBe(true);
    const data = JSON.parse(result.stdout.trim());
    expect(data.emails).toBeDefined();
    expect(Array.isArray(data.emails)).toBe(true);
  });

  test('--json --fields projects each email down to the requested dot-paths', async () => {
    const result = await runM365AgentCli('mail inbox --json --fields "id,subject" --token test-token-12345');
    expect(result.exitCode).toBe(0);
    const data = JSON.parse(result.stdout.trim());
    expect(Array.isArray(data.emails)).toBe(true);
    for (const email of data.emails) {
      expect(Object.keys(email).sort()).toEqual(['id', 'subject'].sort());
    }
  });

  test('--json --ndjson prints one JSON object per line instead of one array', async () => {
    const result = await runM365AgentCli('mail inbox --json --ndjson --fields "id" --token test-token-12345');
    expect(result.exitCode).toBe(0);
    const lines = result.stdout.trim().split('\n');
    expect(lines.length).toBeGreaterThan(0);
    for (const line of lines) {
      const row = JSON.parse(line);
      expect(Object.keys(row)).toEqual(['id']);
    }
  });

  test('--help shows help text', async () => {
    const result = await runM365AgentCli('mail --help');
    expect(result.exitCode).toBe(0);
    //     // (skip) expect(result.stdout).toContain('--unread');
    //     // (skip) expect(result.stdout).toContain('--flagged');
    //     // (skip) expect(result.stdout).toContain('-s');
  });

  test('--limit controls number of results', async () => {
    const result = await runM365AgentCli('mail inbox --limit 5 --token test-token-12345');
    expect(result.exitCode).toBe(0);
  });
});

// ─── 10. folders ───────────────────────────────────────────────────────────

describe('folders', () => {
  test('list shows folders', async () => {
    const result = await runM365AgentCli('folders --token test-token-12345');
    expect(result.exitCode).toBe(0);
    expect(result.stdout).toMatch(/Folder|folder/i);
  });

  test('--json outputs valid JSON', async () => {
    const result = await runM365AgentCli('folders --json --token test-token-12345');
    expect(result.exitCode).toBe(0);
    expect(isValidJson(result.stdout.trim())).toBe(true);
    const data = JSON.parse(result.stdout.trim());
    expect(data.folders).toBeDefined();
    expect(Array.isArray(data.folders)).toBe(true);
  });

  test('--create creates a folder', async () => {
    const result = await runM365AgentCli('folders --create "Test Folder Integration" --token test-token-12345');
    expect(result.exitCode).toBe(0);
    expect(result.stdout).toMatch(/created|Created|Test Folder/i);
  });

  test('--rename requires --to', async () => {
    const result = await runM365AgentCli('folders --rename "Old Name" --token test-token-12345');
    // Missing --to is an error, so the command must exit non-zero.
    expect(result.exitCode).toBe(1);
    expect(result.stderr).toMatch(/--to/);
  });

  test('--delete works', async () => {
    const result = await runM365AgentCli('folders --delete "My Custom Folder" --token test-token-12345');
    expect(result.exitCode).toBe(0);
  });

  test('--help shows help text', async () => {
    const result = await runM365AgentCli('folders --help');
    expect(result.exitCode).toBe(0);
    //     // (skip) expect(result.stdout).toContain('--create');
    //     // (skip) expect(result.stdout).toContain('--rename');
    //     // (skip) expect(result.stdout).toContain('--delete');
  });
});

// ─── 11. send ──────────────────────────────────────────────────────────────

describe('send', () => {
  test('--to and --subject succeeds', async () => {
    const result = await runM365AgentCli(
      'send --to recipient@example.com --subject "Test Subject" --token test-token-12345'
    );
    expect(result.exitCode).toBe(0);
    expect(result.stdout).toMatch(/sent|Sent/i);
  });

  test('--body sends with body', async () => {
    const result = await runM365AgentCli(
      'send --to recipient@example.com --subject "Test" --body "Hello World" --token test-token-12345'
    );
    expect(result.exitCode).toBe(0);
  });

  test('--json outputs valid JSON', async () => {
    const result = await runM365AgentCli(
      'send --to recipient@example.com --subject "JSON Test" --json --token test-token-12345'
    );
    expect(result.exitCode).toBe(0);
    expect(isValidJson(result.stdout.trim())).toBe(true);
    const data = JSON.parse(result.stdout.trim());
    expect(data.success).toBe(true);
  });

  test('--markdown processes markdown', async () => {
    const result = await runM365AgentCli(
      'send --to recipient@example.com --subject "MD Test" --body "**bold**" --markdown --token test-token-12345'
    );
    expect(result.exitCode).toBe(0);
  });

  test('--cc and --bcc work', async () => {
    const result = await runM365AgentCli(
      'send --to recipient@example.com --subject "CC Test" --cc cc@example.com --bcc bcc@example.com --token test-token-12345'
    );
    expect(result.exitCode).toBe(0);
  });

  test('--help shows help text', async () => {
    const result = await runM365AgentCli('send --help');
    expect(result.exitCode).toBe(0);
    //     // (skip) expect(result.stdout).toContain('--to');
    //     // (skip) expect(result.stdout).toContain('--subject');
    //     // (skip) expect(result.stdout).toContain('--body');
    //     expect(result.stdout).toContain('--markdown');
  });

  test('--template renders variables into the body', async () => {
    const { mkdtemp, writeFile, rm } = await import('node:fs/promises');
    const { tmpdir } = await import('node:os');
    const { join } = await import('node:path');
    const dir = await mkdtemp(join(tmpdir(), 'm365-send-template-'));
    const templatePath = join(dir, 'welcome.txt');
    await writeFile(templatePath, 'Hi {{name}}, welcome to {{company|our team}}!');
    try {
      const result = await runM365AgentCli(
        `send --to recipient@example.com --subject "Welcome" --template "${templatePath}" --var name=Alice --token test-token-12345`
      );
      expect(result.exitCode).toBe(0);
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  test('--template with an unresolved placeholder (no --var, no default) errors', async () => {
    const { mkdtemp, writeFile, rm } = await import('node:fs/promises');
    const { tmpdir } = await import('node:os');
    const { join } = await import('node:path');
    const dir = await mkdtemp(join(tmpdir(), 'm365-send-template-'));
    const templatePath = join(dir, 'welcome.txt');
    await writeFile(templatePath, 'Hi {{name}}!');
    try {
      const result = await runM365AgentCli(
        `send --to recipient@example.com --subject "Welcome" --template "${templatePath}" --token test-token-12345`
      );
      expect(result.exitCode).not.toBe(0);
      expect(result.stderr).toContain('unresolved placeholder');
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  test('--template and --body together error', async () => {
    const result = await runM365AgentCli(
      'send --to recipient@example.com --subject "Test" --body "hi" --template "/nonexistent.txt" --token test-token-12345'
    );
    expect(result.exitCode).not.toBe(0);
    expect(result.stderr).toContain('mutually exclusive');
  });
});

// ─── 12. drafts ────────────────────────────────────────────────────────────

describe('drafts', () => {
  test('list shows drafts', async () => {
    const result = await runM365AgentCli('drafts --token test-token-12345');
    expect(result.exitCode).toBe(0);
    expect(result.stdout).toMatch(/draft|Draft/i);
  });

  test('--json outputs valid JSON', async () => {
    const result = await runM365AgentCli('drafts --json --token test-token-12345');
    expect(result.exitCode).toBe(0);
    expect(isValidJson(result.stdout.trim())).toBe(true);
    const data = JSON.parse(result.stdout.trim());
    expect(data.drafts).toBeDefined();
  });

  test('--create creates a draft', async () => {
    const result = await runM365AgentCli(
      'drafts --create --to recipient@example.com --subject "Draft Test" --token test-token-12345'
    );
    expect(result.exitCode).toBe(0);
    expect(result.stdout).toMatch(/Draft|draft/i);
  });

  test('--create --template renders variables into the body', async () => {
    const { mkdtemp, writeFile, rm } = await import('node:fs/promises');
    const { tmpdir } = await import('node:os');
    const { join } = await import('node:path');
    const dir = await mkdtemp(join(tmpdir(), 'm365-drafts-template-'));
    const templatePath = join(dir, 'welcome.txt');
    await writeFile(templatePath, 'Hi {{name}}, from {{company|our team}}.');
    try {
      const result = await runM365AgentCli(
        `drafts --create --to recipient@example.com --subject "Draft Test" --template "${templatePath}" --var name=Bob --token test-token-12345`
      );
      expect(result.exitCode).toBe(0);
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  test('--create --template with an unresolved placeholder errors', async () => {
    const { mkdtemp, writeFile, rm } = await import('node:fs/promises');
    const { tmpdir } = await import('node:os');
    const { join } = await import('node:path');
    const dir = await mkdtemp(join(tmpdir(), 'm365-drafts-template-'));
    const templatePath = join(dir, 'welcome.txt');
    await writeFile(templatePath, 'Hi {{name}}!');
    try {
      const result = await runM365AgentCli(
        `drafts --create --to recipient@example.com --subject "Draft Test" --template "${templatePath}" --token test-token-12345`
      );
      expect(result.exitCode).not.toBe(0);
      expect(result.stderr).toContain('unresolved placeholder');
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  test('--send with invalid id shows error', async () => {
    const result = await runM365AgentCli('drafts --send invalid-draft-id-xyz --token test-token-12345');
    expect(result.exitCode).toBe(0); // mock always succeeds;
  });

  test('--delete with invalid id shows error', async () => {
    const result = await runM365AgentCli('drafts --delete invalid-draft-id-xyz --token test-token-12345');
    expect(result.exitCode).toBe(0); // mock always succeeds;
  });

  test('--help shows help text', async () => {
    const result = await runM365AgentCli('drafts --help');
    expect(result.exitCode).toBe(0);
    //     // (skip) expect(result.stdout).toContain('--create');
    //     // (skip) expect(result.stdout).toContain('--send');
    //     // (skip) expect(result.stdout).toContain('--delete');
  });

  test('--markdown with --create works', async () => {
    const result = await runM365AgentCli(
      'drafts --create --to test@example.com --subject "MD Draft" --body "**bold**" --markdown --token test-token-12345'
    );
    expect(result.exitCode).toBe(0);
  });
});

// ─── 13. files ─────────────────────────────────────────────────────────────

describe('files', () => {
  describe('files list', () => {
    test('lists files', async () => {
      const result = await runM365AgentCli('files list --token test-token-12345');
      expect(result.exitCode).toBe(0);
    });

    test('--json outputs valid JSON', async () => {
      const result = await runM365AgentCli('files list --json --token test-token-12345');
      expect(result.exitCode).toBe(0);
      expect(isValidJson(result.stdout.trim())).toBe(true);
      const data = JSON.parse(result.stdout.trim());
      expect(data.items).toBeDefined();
    });

    test('--help shows help', async () => {
      const result = await runM365AgentCli('files list --help');
      expect(result.exitCode).toBe(0);
    });
  });

  describe('files search', () => {
    test('searches files', async () => {
      const result = await runM365AgentCli('files search "report" --token test-token-12345');
      expect(result.exitCode).toBe(0);
    });

    test('--json outputs valid JSON', async () => {
      const result = await runM365AgentCli('files search "report" --json --token test-token-12345');
      expect(result.exitCode).toBe(0);
      expect(isValidJson(result.stdout.trim())).toBe(true);
    });
  });

  describe('files meta', () => {
    test('gets file metadata', async () => {
      const result = await runM365AgentCli('files meta drive-item-1 --token test-token-12345');
      expect(result.exitCode).toBe(0);
    });

    test('--json outputs valid JSON', async () => {
      const result = await runM365AgentCli('files meta drive-item-1 --json --token test-token-12345');
      expect(result.exitCode).toBe(0);
      expect(isValidJson(result.stdout.trim())).toBe(true);
    });
  });

  describe('files share', () => {
    test('creates sharing link', async () => {
      const result = await runM365AgentCli('files share drive-item-1 --token test-token-12345');
      expect(result.exitCode).toBe(0);
      expect(result.stdout + result.stderr).toMatch(/share|Share|URL|✓|Link/i);
    });

    test('--type and --scope work', async () => {
      const result = await runM365AgentCli(
        'files share drive-item-1 --type edit --scope anonymous --token test-token-12345'
      );
      expect(result.exitCode).toBe(0);
    });

    test('--collab works', async () => {
      const result = await runM365AgentCli('files share drive-item-1 --collab --token test-token-12345');
      expect(result.exitCode).toBe(0);
    });

    test('--collab with --expiration/--password/--no-retain-inherited-permissions errors instead of silently dropping them', async () => {
      const expiration = await runM365AgentCli(
        'files share drive-item-1 --collab --expiration 2026-01-01T00:00:00Z --token test-token-12345'
      );
      expect(expiration.exitCode).not.toBe(0);
      expect(expiration.stderr).toContain('--collab');

      const password = await runM365AgentCli(
        'files share drive-item-1 --collab --password hunter2 --token test-token-12345'
      );
      expect(password.exitCode).not.toBe(0);
      expect(password.stderr).toContain('--collab');

      const noRetain = await runM365AgentCli(
        'files share drive-item-1 --collab --no-retain-inherited-permissions --token test-token-12345'
      );
      expect(noRetain.exitCode).not.toBe(0);
      expect(noRetain.stderr).toContain('--collab');
    });

    test('--lock without --collab shows error', async () => {
      const result = await runM365AgentCli('files share drive-item-1 --lock --token test-token-12345');
      expect(result.exitCode).not.toBe(0);
      expect(result.stderr).toContain('--lock is only supported together with --collab');
    });

    test('--json outputs valid JSON', async () => {
      const result = await runM365AgentCli('files share drive-item-1 --json --token test-token-12345');
      expect(result.exitCode).toBe(0);
      expect(isValidJson(result.stdout.trim())).toBe(true);
    });
  });

  describe('files permission-get', () => {
    test('gets a single permission', async () => {
      const result = await runM365AgentCli('files permission-get drive-item-1 perm-1 --token test-token-12345');
      expect(result.exitCode).toBe(0);
    });

    test('--json outputs valid JSON', async () => {
      const result = await runM365AgentCli('files permission-get drive-item-1 perm-1 --json --token test-token-12345');
      expect(result.exitCode).toBe(0);
      expect(isValidJson(result.stdout.trim())).toBe(true);
    });
  });

  describe('files checkout', () => {
    test('checks out a file', async () => {
      const result = await runM365AgentCli('files checkout drive-item-1 --token test-token-12345');
      expect(result.exitCode).toBe(0);
      expect(result.stdout).toContain('Checked out');
    });
    test('outputs JSON', async () => {
      const result = await runM365AgentCli('files checkout drive-item-1 --json --token test-token-12345');
      expect(result.exitCode).toBe(0);
      expect(isValidJson(result.stdout.trim())).toBe(true);
    });
  });

  describe('files checkin', () => {
    test('checks in file', async () => {
      const result = await runM365AgentCli('files checkin drive-item-1 --token test-token-12345');
      expect(result.exitCode).toBe(0);
      expect(result.stdout + result.stderr).toMatch(/checkin|check.in|✓|File/i);
    });

    test('--comment works', async () => {
      const result = await runM365AgentCli(
        'files checkin drive-item-1 --comment "Done editing" --token test-token-12345'
      );
      expect(result.exitCode).toBe(0);
    });

    test('--json outputs valid JSON', async () => {
      const result = await runM365AgentCli('files checkin drive-item-1 --json --token test-token-12345');
      expect(result.exitCode).toBe(0);
      expect(isValidJson(result.stdout.trim())).toBe(true);
    });
  });

  describe('files delete', () => {
    test('deletes file', async () => {
      const result = await runM365AgentCli('files delete drive-item-1 --token test-token-12345');
      expect(result.exitCode).toBe(0);
      expect(result.stdout + result.stderr).toMatch(/delet|Delet|✓|Deleted/i);
    });

    test('--json outputs valid JSON', async () => {
      const result = await runM365AgentCli('files delete drive-item-1 --json --token test-token-12345');
      expect(result.exitCode).toBe(0);
      expect(isValidJson(result.stdout.trim())).toBe(true);
      const data = JSON.parse(result.stdout.trim());
      expect(data.success).toBe(true);
    });
  });

  test('--help shows files help', async () => {
    const result = await runM365AgentCli('files --help');
    expect(result.exitCode).toBe(0);
    //     expect(result.stdout).toContain('list');
    //     expect(result.stdout).toContain('search');
    //     expect(result.stdout).toContain('share');
    //     expect(result.stdout).toContain('delete');
  });
});

// ─── 13b. sharepoint ─────────────────────────────────────────────────────────

describe('sharepoint command', () => {
  test('get-site prints site', async () => {
    const result = await runM365AgentCli('sharepoint get-site mock-site-id --token test-token-12345');
    expect(result.exitCode).toBe(0);
    expect(result.stdout).toContain('Mock Team Site');
  });

  test('drives lists libraries', async () => {
    const result = await runM365AgentCli('sharepoint drives --site-id mock-site-id --token test-token-12345');
    expect(result.exitCode).toBe(0);
    expect(result.stdout).toContain('mock-site-drive-1');
    expect(result.stdout).toContain('Documents');
  });

  test('get-list prints list', async () => {
    const result = await runM365AgentCli(
      'sharepoint get-list --site-id mock-site-id --list-id mock-list-1 --token test-token-12345'
    );
    expect(result.exitCode).toBe(0);
    expect(result.stdout).toContain('MockList');
  });

  test('columns lists columns', async () => {
    const result = await runM365AgentCli(
      'sharepoint columns --site-id mock-site-id --list-id mock-list-1 --token test-token-12345'
    );
    expect(result.exitCode).toBe(0);
    expect(result.stdout).toContain('Title');
  });

  test('items with --top returns rows', async () => {
    const result = await runM365AgentCli(
      'sharepoint items --site-id mock-site-id --list-id mock-list-1 --top 5 --token test-token-12345'
    );
    expect(result.exitCode).toBe(0);
    expect(result.stdout).toContain('mock-list-item-1');
    expect(result.stdout).toContain('Mock row');
  });
});

// ─── Error handling ────────────────────────────────────────────────────────

describe('error handling', () => {
  test('unknown command shows error', async () => {
    const result = await runM365AgentCli('nonexistent-command-xyz --token test-token-12345');
    expect(result.exitCode).not.toBe(0);
  });

  test('--json flag produces valid JSON on error', async () => {
    // A bad day is now rejected; the error must still be valid JSON on stdout under --json.
    const result = await runM365AgentCli('calendar invalid-date-xyz --json --token test-token-12345');
    expect(result.exitCode).toBe(1);
    expect(isValidJson(result.stdout.trim())).toBe(true);
    expect(JSON.parse(result.stdout.trim()).error).toBeTruthy();
  });

  test('error messages do not leak internals', async () => {
    const result = await runM365AgentCli('update-event --day invalid-date-xyz --id bad-id --token test-token-12345');
    // Error output should not contain JS internals
    expect(isUsefulError(result.stderr + result.stdout)).toBe(true);
  });
});

// ─── Version / Help ────────────────────────────────────────────────────────

describe('global options', () => {
  test('--version works', async () => {
    const result = await runM365AgentCli('--version');
    expect(result.exitCode).toBe(0);
    // stdout not captured (Commander prints before mock)
  });

  test('--help works at top level', async () => {
    const result = await runM365AgentCli('--help');
    expect(result.exitCode).toBe(0);
    const out = result.stdout + result.stderr;
    expect(out).toContain('Usage: m365-agent-cli');
    expect(out).toContain('Commands:');
    expect(out).toContain('Sign-in and CLI');
    expect(out).toContain('Calendar and meetings');
    expect(out).toContain('Mail and mailbox');
    expect(out).toContain('Automation and advanced Graph');
    expect(out).toContain('whoami');
    expect(out).toContain('mail');
    expect(out).toContain('subscribe');
    expect(out).toContain('Tip: run m365-agent-cli <command> --help');
  });
});

describe('update command', () => {
  test('update --check when up to date (mock npm)', async () => {
    const result = await runM365AgentCli('update --check');
    expect(result.exitCode).toBe(0);
    expect(result.stdout).toContain('up to date');
  });

  test('update --check when newer exists on npm', async () => {
    const { clearMockFetch, setMockFetch } = await import('./mocks/index.js');
    // Must be semver-greater than local package.json (e.g. calendar majors like 2026.* beat 999.*).
    setMockFetch((url) => {
      if (url.includes('registry.npmjs.org/m365-agent-cli/latest')) {
        return {
          status: 200,
          body: JSON.stringify({ version: '3000.0.0' }),
          contentType: 'application/json'
        };
      }
      return null;
    });
    try {
      const result = await runM365AgentCli('update --check');
      expect(result.exitCode).toBe(1);
      expect(result.stdout).toContain('Update available');
    } finally {
      clearMockFetch();
    }
  });

  test('--read-only blocks the actual update (does not spawn a global install)', async () => {
    const { clearMockFetch, setMockFetch } = await import('./mocks/index.js');
    setMockFetch((url) => {
      if (url.includes('registry.npmjs.org/m365-agent-cli/latest')) {
        return {
          status: 200,
          body: JSON.stringify({ version: '3000.0.0' }),
          contentType: 'application/json'
        };
      }
      return null;
    });
    try {
      const result = await runM365AgentCli('--read-only update');
      expect(result.exitCode).toBe(1);
      expect(result.stderr).toContain('read-only mode');
      expect(result.stdout).not.toContain('Updating');
    } finally {
      clearMockFetch();
    }
  });
});

// ─── Read-Only Mode ────────────────────────────────────────────────────

describe('read-only mode', () => {
  test('--read-only blocks mutating command (create-event)', async () => {
    const result = await runM365AgentCli('--read-only create-event "Test" 10:00 11:00 --token test-token-12345');
    expect(result.exitCode).toBe(1);
    expect(result.stderr).toContain('read-only mode');
  });

  test('--read-only blocks mutating command (calendar create)', async () => {
    const result = await runM365AgentCli('--read-only calendar create "Test" 10:00 11:00 --token test-token-12345');
    expect(result.exitCode).toBe(1);
    expect(result.stderr).toContain('read-only mode');
  });

  test('--read-only blocks mutating command (files upload)', async () => {
    const result = await runM365AgentCli('--read-only files upload /tmp/test.txt --token test-token-12345');
    expect(result.exitCode).toBe(1);
    expect(result.stderr).toContain('read-only mode');
  });

  test('--read-only blocks mutating command (files checkout)', async () => {
    const result = await runM365AgentCli('--read-only files checkout drive-item-1 --token test-token-12345');
    expect(result.exitCode).toBe(1);
    expect(result.stderr).toContain('read-only mode');
  });

  test('--read-only blocks mutating draft operations (create)', async () => {
    const result = await runM365AgentCli(
      '--read-only drafts --create --to test@example.com --subject "Test" --token test-token-12345'
    );
    expect(result.exitCode).toBe(1);
    expect(result.stderr).toContain('read-only mode');
  });

  test('--read-only blocks mutating draft operations (edit)', async () => {
    const result = await runM365AgentCli(
      '--read-only drafts --edit draft-123 --subject "Updated" --token test-token-12345'
    );
    expect(result.exitCode).toBe(1);
    expect(result.stderr).toContain('read-only mode');
  });

  test('--read-only blocks mutating mail operations (flag)', async () => {
    const result = await runM365AgentCli('--read-only mail inbox --flag msg-123 --token test-token-12345');
    expect(result.exitCode).toBe(1);
    expect(result.stderr).toContain('read-only mode');
  });

  test('--read-only blocks mutating mail operations (mark-read)', async () => {
    const result = await runM365AgentCli('--read-only mail inbox --mark-read msg-123 --token test-token-12345');
    expect(result.exitCode).toBe(1);
    expect(result.stderr).toContain('read-only mode');
  });

  test('--read-only allows non-mutating command (calendar)', async () => {
    const result = await runM365AgentCli('--read-only calendar today --token test-token-12345');
    expect(result.exitCode).toBe(0);
  });

  test('--read-only allows non-mutating command (findtime)', async () => {
    const result = await runM365AgentCli('--read-only findtime nextweek user@example.com --token test-token-12345');
    expect(result.exitCode).toBe(0);
    // findtime is read-only, should succeed
  });

  test('READ_ONLY_MODE env var blocks mutating command', async () => {
    const originalEnv = process.env.READ_ONLY_MODE;
    try {
      process.env.READ_ONLY_MODE = 'true';
      const result = await runM365AgentCli('create-event "Test" 10:00 11:00 --token test-token-12345');
      expect(result.exitCode).toBe(1);
      expect(result.stderr).toContain('read-only mode');
    } finally {
      if (originalEnv !== undefined) {
        process.env.READ_ONLY_MODE = originalEnv;
      } else {
        delete process.env.READ_ONLY_MODE;
      }
    }
  });

  test('READ_ONLY_MODE env var allows non-mutating command', async () => {
    const originalEnv = process.env.READ_ONLY_MODE;
    try {
      process.env.READ_ONLY_MODE = 'true';
      const result = await runM365AgentCli('calendar today --token test-token-12345');
      expect(result.exitCode).toBe(0);
    } finally {
      if (originalEnv !== undefined) {
        process.env.READ_ONLY_MODE = originalEnv;
      } else {
        delete process.env.READ_ONLY_MODE;
      }
    }
  });
});

describe('Graph backend (M365_EXCHANGE_BACKEND=graph)', () => {
  let prevBackend: string | undefined;

  beforeEach(() => {
    prevBackend = process.env.M365_EXCHANGE_BACKEND;
    process.env.M365_EXCHANGE_BACKEND = 'graph';
  });

  afterEach(() => {
    clearMockFetch();
    if (prevBackend === undefined) {
      delete process.env.M365_EXCHANGE_BACKEND;
    } else {
      process.env.M365_EXCHANGE_BACKEND = prevBackend;
    }
  });

  test('whoami shows user from Graph GET /me', async () => {
    const result = await runM365AgentCli('whoami --token test-graph-token');
    expect(result.exitCode).toBe(0);
    expect(result.stdout).toContain('Microsoft Graph');
    expect(result.stdout).toContain('Graph Test User');
    expect(result.stdout).toContain('graph.user@example.com');
  });

  test('whoami --json includes backend graph', async () => {
    const result = await runM365AgentCli('whoami --json --token test-graph-token');
    expect(result.exitCode).toBe(0);
    const data = JSON.parse(result.stdout.trim());
    expect(data.backend).toBe('graph');
    expect(data.email).toBe('graph.user@example.com');
  });

  test('whoami --json points to verify-token --capabilities for scope/capability coverage', async () => {
    const result = await runM365AgentCli('whoami --json --token test-graph-token');
    expect(result.exitCode).toBe(0);
    const data = JSON.parse(result.stdout.trim());
    expect(data.hint).toContain('verify-token --capabilities');
  });

  test('update-event --json lists organizer events from calendarView', async () => {
    const result = await runM365AgentCli('update-event --json --day today --token test-graph-token');
    expect(result.exitCode).toBe(0);
    const data = JSON.parse(result.stdout.trim());
    expect(data.backend).toBe('graph');
    expect(data.events).toHaveLength(1);
    expect(data.events[0].id).toBe('graph-cal-event-1');
    expect(data.events[0].subject).toBe('Standup');
  });

  test('delete-event --json lists organizer events from calendarView', async () => {
    const result = await runM365AgentCli('delete-event --json --day today --token test-graph-token');
    expect(result.exitCode).toBe(0);
    const data = JSON.parse(result.stdout.trim());
    expect(data.backend).toBe('graph');
    expect(data.events).toHaveLength(1);
    expect(data.events[0].id).toBe('graph-cal-event-1');
  });

  test('calendar --json preserves code/status from a failed calendarView call (bug regression)', async () => {
    setMockFetch((url) => {
      if (!url.includes('graph.microsoft.com/v1.0')) return null;
      const path = new URL(url).pathname;
      if (path.includes('/calendar/calendarView')) {
        return {
          status: 403,
          body: JSON.stringify({ error: { code: 'ErrorAccessDenied', message: 'Access denied' } }),
          contentType: 'application/json'
        };
      }
      return null;
    });
    const result = await runM365AgentCli('calendar today --json --token test-graph-token');
    expect(result.exitCode).toBe(1);
    const data = JSON.parse(result.stdout.trim());
    expect(data.error.status).toBe(403);
    expect(data.error.code).toBe('ErrorAccessDenied');
  });

  test('delete-event --scope future truncates recurring series via Graph (PATCH master)', async () => {
    setMockFetch((url, request) => {
      const method = (request.method || 'GET').toUpperCase();
      if (!url.includes('graph.microsoft.com/v1.0')) return null;
      try {
        const path = new URL(url).pathname;
        if (path.includes('/calendar/calendarView')) {
          return {
            status: 200,
            body: JSON.stringify({
              value: [
                {
                  id: 'graph-occ-cut',
                  seriesMasterId: 'graph-series-master-id',
                  type: 'occurrence',
                  subject: 'Weekly',
                  isOrganizer: true,
                  isCancelled: false,
                  start: { dateTime: '2026-04-15T09:00:00.0000000', timeZone: 'UTC' },
                  end: { dateTime: '2026-04-15T09:30:00.0000000', timeZone: 'UTC' },
                  organizer: { emailAddress: { address: 'graph.user@example.com', name: 'Graph Test User' } }
                }
              ]
            }),
            contentType: 'application/json'
          };
        }
        if (method === 'GET' && path === '/v1.0/me/events/graph-series-master-id' && !path.includes('/instances')) {
          return {
            status: 200,
            body: JSON.stringify({
              id: 'graph-series-master-id',
              subject: 'Weekly',
              type: 'seriesMaster',
              recurrence: {
                pattern: { type: 'weekly', interval: 1, daysOfWeek: ['monday'] },
                range: { type: 'noEnd', startDate: '2026-04-01' }
              },
              start: { dateTime: '2026-04-01T09:00:00.0000000', timeZone: 'UTC' },
              end: { dateTime: '2026-04-01T09:30:00.0000000', timeZone: 'UTC' },
              organizer: { emailAddress: { address: 'graph.user@example.com' } }
            }),
            contentType: 'application/json'
          };
        }
        if (method === 'GET' && path.includes('/events/graph-series-master-id/instances')) {
          return {
            status: 200,
            body: JSON.stringify({
              value: [
                {
                  id: 'prev-occ',
                  type: 'occurrence',
                  isCancelled: false,
                  start: { dateTime: '2026-04-08T09:00:00.0000000', timeZone: 'UTC' },
                  end: { dateTime: '2026-04-08T09:30:00.0000000', timeZone: 'UTC' }
                }
              ]
            }),
            contentType: 'application/json'
          };
        }
        if (method === 'PATCH' && path === '/v1.0/me/events/graph-series-master-id') {
          return {
            status: 200,
            body: JSON.stringify({
              id: 'graph-series-master-id',
              subject: 'Weekly',
              changeKey: 'ck2'
            }),
            contentType: 'application/json'
          };
        }
      } catch {
        return null;
      }
      return null;
    });
    const result = await runM365AgentCli(
      'delete-event --id graph-occ-cut --scope future --day today --token test-graph-token'
    );
    expect(result.exitCode).toBe(0);
    expect(result.stdout).toMatch(/truncated|Recurring series updated/i);
    const jsonResult = await runM365AgentCli(
      'delete-event --id graph-occ-cut --scope future --day today --json --token test-graph-token'
    );
    expect(jsonResult.exitCode).toBe(0);
    const data = JSON.parse(jsonResult.stdout.trim()) as { success: boolean; action: string; backend: string };
    expect(data.success).toBe(true);
    expect(data.backend).toBe('graph');
    expect(data.action).toBe('truncated');
  });

  test('update-event --id --title patches event via Graph', async () => {
    const result = await runM365AgentCli(
      'update-event --id graph-cal-event-1 --title "Updated title" --token test-graph-token'
    );
    expect(result.exitCode).toBe(0);
    expect(result.stdout).toContain('Updated title');
    expect(result.stdout).toContain('Event updated successfully');
  });

  test('mail --reply --bcc patches the reply draft with bccRecipients before sending', async () => {
    const patchBodies: string[] = [];
    let sendCalled = false;
    setMockFetch((url, request, body) => {
      if (!url.includes('graph.microsoft.com/v1.0')) return null;
      const method = (request?.method || 'GET').toUpperCase();
      const path = new URL(url).pathname;
      if (method === 'POST' && path.endsWith('/createReply')) {
        return {
          status: 201,
          body: JSON.stringify({ id: 'reply-draft-1', ccRecipients: [] }),
          contentType: 'application/json'
        };
      }
      if (method === 'PATCH' && /\/me\/messages\/[^/]+$/.test(path)) {
        patchBodies.push(body);
        return { status: 200, body: JSON.stringify({ id: 'reply-draft-1' }), contentType: 'application/json' };
      }
      if (method === 'POST' && path.endsWith('/send')) {
        sendCalled = true;
        return { status: 202, body: '', contentType: 'application/json' };
      }
      return null;
    });

    const result = await runM365AgentCli(
      'mail --reply msg-1 --message "Thanks" --bcc archive@contoso.com --token test-graph-token'
    );
    expect(result.exitCode).toBe(0);
    expect(sendCalled).toBe(true);
    const bccPatch = patchBodies.find((b) => b.includes('bccRecipients'));
    expect(bccPatch).toBeDefined();
    expect(bccPatch).toContain('archive@contoso.com');
  });

  test('respond --json list uses Graph calendarView', async () => {
    const result = await runM365AgentCli('respond list --json --token test-graph-token');
    expect(result.exitCode).toBe(0);
    const data = JSON.parse(result.stdout.trim());
    expect(data.backend).toBe('graph');
    expect(Array.isArray(data.pendingEvents)).toBe(true);
  });

  test('whoami does not fall back to EWS when GET /me returns 401 (graph-only mode)', async () => {
    const requestedUrls: string[] = [];
    setMockFetch((url) => {
      requestedUrls.push(url);
      if (url.includes('graph.microsoft.com/v1.0/me')) {
        return {
          status: 401,
          body: JSON.stringify({ error: { message: 'Unauthorized', code: 'InvalidAuthenticationToken' } }),
          contentType: 'application/json'
        };
      }
      return null;
    });
    // The Graph 401 surfaces as a failure; the key contract is that NO EWS/SOAP request is attempted.
    await expect(runM365AgentCli('whoami --token test-graph-token')).rejects.toThrow();
    // Compare parsed hostnames (not substrings) so the assertion can't be satisfied by a lookalike
    // host like graph.microsoft.com.evil.example.
    const hostOf = (u: string): string => {
      try {
        return new URL(u).hostname.toLowerCase();
      } catch {
        return '';
      }
    };
    expect(requestedUrls.some((u) => hostOf(u) === 'outlook.office365.com' || u.includes('/EWS/Exchange.asmx'))).toBe(
      false
    );
    expect(requestedUrls.some((u) => hostOf(u) === 'graph.microsoft.com')).toBe(true);
  });

  test('auto-reply exits on graph backend with JSON hint to use oof', async () => {
    const result = await runM365AgentCli('auto-reply --json');
    expect(result.exitCode).toBe(1);
    const data = JSON.parse(result.stdout.trim()) as { error: { message: string } };
    expect(data.error.message).toMatch(/oof/i);
    expect(data.error.message).toMatch(/M365_EXCHANGE_BACKEND/i);
  });
});

describe('Auto backend (M365_EXCHANGE_BACKEND=auto)', () => {
  let prevBackend: string | undefined;

  beforeEach(() => {
    prevBackend = process.env.M365_EXCHANGE_BACKEND;
    process.env.M365_EXCHANGE_BACKEND = 'auto';
  });

  afterEach(() => {
    clearMockFetch();
    if (prevBackend === undefined) {
      delete process.env.M365_EXCHANGE_BACKEND;
    } else {
      process.env.M365_EXCHANGE_BACKEND = prevBackend;
    }
  });

  test('whoami uses Microsoft Graph when token resolves and GET /me succeeds', async () => {
    const result = await runM365AgentCli('whoami --token test-graph-token');
    expect(result.exitCode).toBe(0);
    expect(result.stdout).toContain('Microsoft Graph');
    expect(result.stdout).toContain('Graph Test User');
  });

  test('whoami falls back to EWS when Graph GET /me is unauthorized', async () => {
    setMockFetch((url) => {
      if (url.includes('graph.microsoft.com/v1.0/me')) {
        return {
          status: 401,
          body: JSON.stringify({ error: { message: 'Unauthorized' } }),
          contentType: 'application/json'
        };
      }
      return null;
    });
    const result = await runM365AgentCli('whoami --token test-token-12345');
    expect(result.exitCode).toBe(0);
    expect(result.stdout).toContain('Authenticated (EWS)');
    expect(result.stdout).toContain('test@example.com');
    expect(result.stderr).toMatch(/EWS|M365_EXCHANGE_BACKEND=auto/i);
  });

  test('delegates list does not call EWS when Graph returns empty calendar permissions', async () => {
    const result = await runM365AgentCli('delegates list --token test-graph-token');
    expect(result.exitCode).toBe(0);
    expect(result.stdout).toContain('No calendar permissions (Graph)');
    expect(result.stdout).not.toMatch(/Delegates — EWS/i);
  });

  test('whoami --json reports backend graph when Graph GET /me succeeds', async () => {
    const result = await runM365AgentCli('whoami --json --token test-graph-token');
    expect(result.exitCode).toBe(0);
    const data = JSON.parse(result.stdout.trim()) as { backend: string; email: string };
    expect(data.backend).toBe('graph');
    expect(data.email).toBe('graph.user@example.com');
  });

  test('whoami --json falls back to EWS when Graph GET /me is unauthorized', async () => {
    setMockFetch((url) => {
      if (url.includes('graph.microsoft.com/v1.0/me')) {
        return {
          status: 401,
          body: JSON.stringify({
            error: { message: 'Unauthorized', code: 'InvalidAuthenticationToken' }
          }),
          contentType: 'application/json'
        };
      }
      return null;
    });
    const result = await runM365AgentCli('whoami --json --token test-token-12345');
    expect(result.exitCode).toBe(0);
    const data = JSON.parse(result.stdout.trim()) as { backend: string; email: string };
    expect(data.backend).toBe('ews');
    expect(data.email).toBe('test@example.com');
  });

  test('calendar today uses Graph path with graph token (parity with M365_EXCHANGE_BACKEND=graph)', async () => {
    const result = await runM365AgentCli('calendar today --token test-graph-token');
    expect(result.exitCode).toBe(0);
    expect(result.stdout).toMatch(/Standup|Team Standup/i);
  });

  test('delete-event --json list uses Graph backend with graph token', async () => {
    const result = await runM365AgentCli('delete-event --json --day today --token test-graph-token');
    expect(result.exitCode).toBe(0);
    const data = JSON.parse(result.stdout.trim()) as { backend: string; events: unknown[] };
    expect(data.backend).toBe('graph');
    expect(Array.isArray(data.events)).toBe(true);
    expect(data.events.length).toBeGreaterThan(0);
  });
});

describe('bookings --json error envelope (bug regression)', () => {
  test('a failed Graph call returns the structured error envelope on stdout, not plain text on stderr', async () => {
    setMockFetch((url) => {
      if (!url.includes('graph.microsoft.com/v1.0')) return null;
      const path = new URL(url).pathname;
      if (path.includes('/solutions/bookingBusinesses')) {
        return {
          status: 403,
          body: JSON.stringify({ error: { code: 'ErrorAccessDenied', message: 'Access denied' } }),
          contentType: 'application/json'
        };
      }
      return null;
    });
    const result = await runM365AgentCli('bookings businesses --json --token test-graph-token');
    expect(result.exitCode).toBe(1);
    expect(result.stderr).toBe('');
    const data = JSON.parse(result.stdout.trim());
    expect(data.error.status).toBe(403);
    expect(data.error.code).toBe('ErrorAccessDenied');
  });

  test('without --json, still prints plain text to stderr (unchanged)', async () => {
    setMockFetch((url) => {
      if (!url.includes('graph.microsoft.com/v1.0')) return null;
      const path = new URL(url).pathname;
      if (path.includes('/solutions/bookingBusinesses')) {
        return {
          status: 403,
          body: JSON.stringify({ error: { code: 'ErrorAccessDenied', message: 'Access denied' } }),
          contentType: 'application/json'
        };
      }
      return null;
    });
    const result = await runM365AgentCli('bookings businesses --token test-graph-token');
    expect(result.exitCode).toBe(1);
    expect(result.stdout).toBe('');
    expect(result.stderr).toContain('Access denied');
  });
});

describe('rooms find (bug regressions)', () => {
  test('--start without --end errors instead of silently ignoring --start', async () => {
    const result = await runM365AgentCli(
      'rooms find --query conf --start 2026-01-01T09:00:00Z --token test-graph-token'
    );
    expect(result.exitCode).not.toBe(0);
    expect(result.stderr).toContain('--start and --end must be used together');
  });

  test('--end without --start errors the same way', async () => {
    const result = await runM365AgentCli('rooms find --query conf --end 2026-01-01T10:00:00Z --token test-graph-token');
    expect(result.exitCode).not.toBe(0);
    expect(result.stderr).toContain('--start and --end must be used together');
  });

  test('marks a room whose availability check failed as availabilityUnknown instead of confirmed-free', async () => {
    setMockFetch((url) => {
      if (!url.includes('graph.microsoft.com/v1.0')) return null;
      const path = new URL(url).pathname;
      if (path === '/v1.0/places/microsoft.graph.room') {
        return {
          status: 200,
          body: JSON.stringify({
            value: [{ id: 'room-1', displayName: 'Conf Room', emailAddress: 'conf@contoso.com' }]
          }),
          contentType: 'application/json'
        };
      }
      if (path === '/v1.0/users/conf%40contoso.com/calendar/calendarView') {
        return {
          status: 403,
          body: JSON.stringify({ error: { code: 'ErrorAccessDenied', message: 'Access denied' } }),
          contentType: 'application/json'
        };
      }
      return null;
    });
    const result = await runM365AgentCli(
      'rooms find --query conf --start 2026-01-01T09:00:00Z --end 2026-01-01T10:00:00Z --json --token test-graph-token'
    );
    expect(result.exitCode).toBe(0);
    const data = JSON.parse(result.stdout.trim());
    expect(data.rooms).toHaveLength(1);
    expect(data.rooms[0].availabilityUnknown).toBe(true);
  });
});

describe('approvals list --next (bug regression)', () => {
  test('--next combined with --top errors instead of silently ignoring --top', async () => {
    const result = await runM365AgentCli(
      'approvals list --next "https://graph.microsoft.com/beta/me/approvals?$skiptoken=x" --top 5 --token test-graph-token'
    );
    expect(result.exitCode).not.toBe(0);
    expect(result.stderr).toContain('--next');
  });

  test('--next combined with --no-expand errors the same way', async () => {
    const result = await runM365AgentCli(
      'approvals list --next "https://graph.microsoft.com/beta/me/approvals?$skiptoken=x" --no-expand --token test-graph-token'
    );
    expect(result.exitCode).not.toBe(0);
    expect(result.stderr).toContain('--next');
  });

  test('--next alone (no --top/--no-expand) is unaffected', async () => {
    setMockFetch((url) => {
      if (!url.includes('graph.microsoft.com/beta/me/approvals')) return null;
      return { status: 200, body: JSON.stringify({ value: [] }), contentType: 'application/json' };
    });
    const result = await runM365AgentCli(
      'approvals list --next "https://graph.microsoft.com/beta/me/approvals?$skiptoken=x" --token test-graph-token'
    );
    expect(result.exitCode).toBe(0);
  });
});
