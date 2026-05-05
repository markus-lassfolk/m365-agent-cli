import { describe, expect, test } from 'bun:test';
import { Command, Help } from 'commander';
import { buildRootCommandSections } from '../lib/root-command-groups.js';

describe('buildRootCommandSections', () => {
  test('groups registered names and collects unknown under Other commands', () => {
    const root = new Command();
    root.helpCommand(false);
    root.addCommand(new Command('whoami').description('Who'));
    root.addCommand(new Command('calendar').description('Cal'));
    root.addCommand(new Command('orphan-cmd').description('Not in registry'));
    const helper = new Help();
    const sections = buildRootCommandSections(root, helper);
    const titles = sections.map((s) => s.title);
    expect(titles.some((t) => t === 'Sign-in and CLI')).toBe(true);
    expect(titles.some((t) => t === 'Calendar and meetings')).toBe(true);
    const other = sections.find((s) => s.title === 'Other commands');
    expect(other?.commands.map((c) => c.name())).toEqual(['orphan-cmd']);
  });

  test('omits empty sections when no commands from that group are present', () => {
    const root = new Command();
    root.helpCommand(false);
    root.addCommand(new Command('only-unknown').description('x'));
    const sections = buildRootCommandSections(root, new Help());
    expect(sections).toHaveLength(1);
    expect(sections[0].title).toBe('Other commands');
    expect(sections[0].commands).toHaveLength(1);
  });
});
