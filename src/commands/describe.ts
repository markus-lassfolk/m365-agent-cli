import { Command } from 'commander';
import { describeProgram, findCommandManifestByPath } from '../lib/command-manifest.js';

export const describeCommand = new Command('describe')
  .description(
    'Machine-readable manifest of every command, subcommand, option, and argument (JSON). ' +
      'Use this to discover the CLI surface programmatically instead of parsing --help text.'
  )
  .option('--command <path>', 'Only describe this command (space-separated path, e.g. "rules create")')
  .option('--list', 'List top-level commands only (name, aliases, description) — fast overview')
  .action((options: { command?: string; list?: boolean }, cmd: Command) => {
    // `cmd.parent` is the root program instance Commander passes at runtime — this avoids a
    // circular import back to m365-program.ts (which registers this command).
    const root = cmd.parent;
    if (!root) {
      console.log(JSON.stringify({ error: 'describe: could not resolve the root command' }, null, 2));
      process.exit(1);
    }

    const manifest = describeProgram(root);

    if (options.command) {
      const found = findCommandManifestByPath(manifest, options.command);
      if (!found) {
        console.log(JSON.stringify({ error: `describe: no command found at path "${options.command}"` }, null, 2));
        process.exit(1);
      }
      console.log(JSON.stringify(found, null, 2));
      return;
    }

    if (options.list) {
      console.log(
        JSON.stringify(
          manifest.commands.map((c) => ({ name: c.name, aliases: c.aliases, description: c.description })),
          null,
          2
        )
      );
      return;
    }

    console.log(JSON.stringify(manifest, null, 2));
  });
