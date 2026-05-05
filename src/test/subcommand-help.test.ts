/**
 * Grouped `--help` for parents listed in subcommand-help-groups.ts
 */
import { describe, expect, test } from 'bun:test';
import { Command, CommanderError } from 'commander';
import { calendarCommand } from '../commands/calendar.js';
import { filesCommand } from '../commands/files.js';
import { teamsCommand } from '../commands/teams.js';
import { installM365HelpOnCommandTree } from '../lib/m365-help.js';

/** `addCommand()` does not copy `configureOutput` / `exitOverride` onto attached commands (unlike `.command()`). */
function forEachCommandDepthFirst(cmd: Command, fn: (c: Command) => void): void {
  fn(cmd);
  for (const sub of cmd.commands) {
    forEachCommandDepthFirst(sub, fn);
  }
}

async function helpOutput(program: Command, argv: string[]): Promise<string> {
  const chunks: string[] = [];
  forEachCommandDepthFirst(program, (cmd) => {
    cmd.configureOutput({
      writeOut: (s) => {
        chunks.push(s);
      },
      writeErr: (s) => {
        chunks.push(s);
      }
    });
    cmd.exitOverride();
  });
  try {
    await program.parseAsync(argv);
  } catch (err) {
    if (err instanceof CommanderError && (err.code === 'commander.helpDisplayed' || err.code === 'commander.help')) {
      return chunks.join('');
    }
    throw err;
  }
  return chunks.join('');
}

describe('grouped subcommand help', () => {
  test('teams --help shows section titles', async () => {
    const program = new Command();
    program.name('m365-agent-cli');
    program.addCommand(teamsCommand);
    installM365HelpOnCommandTree(program);
    const out = await helpOutput(program, ['node', 'cli', 'teams', '--help']);
    expect(out).toContain('Teams and channels');
    expect(out).toContain('Channel tabs');
    expect(out).toContain('Apps and installations');
    expect(out).toContain('Teams List');
  });

  test('files --help shows section titles', async () => {
    const program = new Command();
    program.name('m365-agent-cli');
    program.addCommand(filesCommand);
    installM365HelpOnCommandTree(program);
    const out = await helpOutput(program, ['node', 'cli', 'files', '--help']);
    expect(out).toContain('Browse and read');
    expect(out).toContain('Upload and download');
    expect(out).toContain('List folder or drive root children');
  });

  test('calendar --help shows Calendar section', async () => {
    const program = new Command();
    program.name('m365-agent-cli');
    program.addCommand(calendarCommand);
    installM365HelpOnCommandTree(program);
    const out = await helpOutput(program, ['node', 'cli', 'calendar', '--help']);
    expect(out).toContain('Calendar');
    expect(out).toContain('List calendar events for a day or range');
  });
});
