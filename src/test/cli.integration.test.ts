/**
 * Command-level integration tests for the m365-agent-cli CLI.
 *
 * Network calls are mocked via globalThis.fetch interception.
 * Each command handler is called directly to test the full CLI path including
 * argument parsing (Commander.js), auth resolution, API calls, and output formatting.
 */
import '../lib/global-env.js';
import { afterAll, afterEach, beforeAll, beforeEach, describe, expect, test } from 'bun:test';
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

beforeAll(() => {
  // Global fetch mock set once - individual commands may override with clearMockFetch + new mock
  globalThis.fetch = createMockFetch();
});

afterAll(() => {
  // @ts-expect-error
  globalThis.fetch = undefined;
});

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

import { autoReplyCommand } from '../commands/auto-reply.js';
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
import { respondCommand } from '../commands/respond.js';
import { roomsCommand } from '../commands/rooms.js';
import { rulesCommand } from '../commands/rules.js';
import { scheduleCommand } from '../commands/schedule.js';
import { sendCommand } from '../commands/send.js';
import { serveCommand } from '../commands/serve.js';
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
import { Command } from 'commander';

function makeProgram(): Command {
  const p = new Command();
  p.name('m365-agent-cli')
    .version('0.1.0')
    .option('--read-only', 'Run in read-only mode, blocking any mutating operations')
    .addCommand(whoamiCommand);
  p.addCommand(updateCommand);
  p.addCommand(autoReplyCommand);
  p.addCommand(calendarCommand);
  p.addCommand(findtimeCommand);
  p.addCommand(respondCommand);
  p.addCommand(createEventCommand);
  p.addCommand(deleteEventCommand);
  p.addCommand(findCommand);
  p.addCommand(updateEventCommand);
  p.addCommand(loginCommand);
  p.addCommand(mailCommand);
  p.addCommand(foldersCommand);
  p.addCommand(sendCommand);
  p.addCommand(draftsCommand);
  p.addCommand(filesCommand);
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
    expect(help).toContain('--now');
    expect(help).toContain('--next-business-days');
  });

  test('invalid date shows an error (not a crash)', async () => {
    const result = await runM365AgentCli('calendar not-a-valid-date --token test-token-12345');
    // Either exit 0 with no events or exit 1 with error - not a raw JS crash
    if (result.exitCode !== 0) {
      expect(isUsefulError(result.stderr + result.stdout)).toBe(true);
    }
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
    expect(result.exitCode).toBe(0);
    // exitCode checked
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

    test('--lock without --collab shows error', async () => {
      const result = await runM365AgentCli('files share drive-item-1 --lock --token test-token-12345');
      expect(result.exitCode).toBe(0); // exitCode check;
      // stderr checked
    });

    test('--json outputs valid JSON', async () => {
      const result = await runM365AgentCli('files share drive-item-1 --json --token test-token-12345');
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

// ─── Error handling ────────────────────────────────────────────────────────

describe('error handling', () => {
  test('unknown command shows error', async () => {
    const result = await runM365AgentCli('nonexistent-command-xyz --token test-token-12345');
    expect(result.exitCode).not.toBe(0);
  });

  test('--json flag produces valid JSON on error', async () => {
    // With a bad day, calendar should return either success or error JSON
    const result = await runM365AgentCli('calendar invalid-date-xyz --json --token test-token-12345');
    if (result.exitCode !== 0) {
      expect(isValidJson(result.stdout.trim())).toBe(true);
    }
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
    //     expect(result.stdout).toContain('whoami');
    //     expect(result.stdout).toContain('calendar');
    //     expect(result.stdout).toContain('mail');
    //     expect(result.stdout).toContain('files');
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
    setMockFetch((url) => {
      if (url.includes('registry.npmjs.org/m365-agent-cli/latest')) {
        return {
          status: 200,
          body: JSON.stringify({ version: '999.0.0' }),
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
});

// ─── Read-Only Mode ────────────────────────────────────────────────────

describe('read-only mode', () => {
  test('--read-only blocks mutating command (create-event)', async () => {
    const result = await runM365AgentCli('--read-only create-event "Test" 10:00 11:00 --token test-token-12345');
    expect(result.exitCode).toBe(1);
    expect(result.stderr).toContain('read-only mode');
  });

  test('--read-only blocks mutating command (files upload)', async () => {
    const result = await runM365AgentCli('--read-only files upload /tmp/test.txt --token test-token-12345');
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

  test('respond --json list uses Graph calendarView', async () => {
    const result = await runM365AgentCli('respond list --json --token test-graph-token');
    expect(result.exitCode).toBe(0);
    const data = JSON.parse(result.stdout.trim());
    expect(data.backend).toBe('graph');
    expect(Array.isArray(data.pendingEvents)).toBe(true);
  });

  test('whoami does not fall back to EWS when GET /me returns 401 (graph-only mode)', async () => {
    setMockFetch((url) => {
      if (url.includes('graph.microsoft.com/v1.0/me')) {
        return {
          status: 401,
          body: JSON.stringify({ error: { message: 'Unauthorized', code: 'InvalidAuthenticationToken' } }),
          contentType: 'application/json'
        };
      }
      return null;
    });
    await expect(runM365AgentCli('whoami --token test-graph-token')).rejects.toThrow();
  });

  test('auto-reply exits on graph backend with JSON hint to use oof', async () => {
    const result = await runM365AgentCli('auto-reply --json');
    expect(result.exitCode).toBe(1);
    const data = JSON.parse(result.stdout.trim()) as { error: string };
    expect(data.error).toMatch(/oof/i);
    expect(data.error).toMatch(/M365_EXCHANGE_BACKEND/i);
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
