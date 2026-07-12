import type { Command, Option } from 'commander';
import { getPackageVersionSync } from './package-info.js';

export interface ManifestOption {
  flags: string;
  long: string | null;
  short: string | null;
  description: string;
  /** True only for options created via `.requiredOption()` (the flag itself must be supplied). */
  mandatory: boolean;
  /** True when the option takes a required value, e.g. `--foo <val>`. */
  valueRequired: boolean;
  /** True when the option takes an optional value, e.g. `--foo [val]`. */
  valueOptional: boolean;
  variadic: boolean;
  defaultValue?: unknown;
}

export interface ManifestArgument {
  name: string;
  description: string;
  required: boolean;
  variadic: boolean;
}

export interface CommandManifest {
  name: string;
  path: string;
  aliases: string[];
  description: string;
  arguments: ManifestArgument[];
  options: ManifestOption[];
  subcommands: CommandManifest[];
}

export interface ProgramManifest {
  name: string;
  version: string;
  description: string;
  globalOptions: ManifestOption[];
  commands: CommandManifest[];
}

/** Excludes the auto-added `-h, --help` / `-V, --version` options Commander injects on every command. */
const NOISE_FLAGS = new Set(['-h, --help', '-V, --version']);

function toManifestOption(opt: Option): ManifestOption {
  return {
    flags: opt.flags,
    long: opt.long ?? null,
    short: opt.short ?? null,
    description: opt.description ?? '',
    mandatory: opt.mandatory ?? false,
    valueRequired: opt.required ?? false,
    valueOptional: opt.optional ?? false,
    variadic: opt.variadic ?? false,
    ...(opt.defaultValue !== undefined ? { defaultValue: opt.defaultValue } : {})
  };
}

function toManifestOptions(cmd: Command): ManifestOption[] {
  return cmd.options.filter((o) => !NOISE_FLAGS.has(o.flags)).map(toManifestOption);
}

function toManifestArguments(cmd: Command): ManifestArgument[] {
  return cmd.registeredArguments.map((a) => ({
    name: a.name(),
    description: a.description ?? '',
    required: a.required,
    variadic: a.variadic ?? false
  }));
}

/** Recursively reflects a Commander command (and its subcommands) into a plain, JSON-safe manifest. */
export function describeCommandTree(cmd: Command, parentPath = ''): CommandManifest {
  const path = parentPath ? `${parentPath} ${cmd.name()}` : cmd.name();
  return {
    name: cmd.name(),
    path,
    aliases: cmd.aliases(),
    description: cmd.description() ?? '',
    arguments: toManifestArguments(cmd),
    options: toManifestOptions(cmd),
    subcommands: cmd.commands.map((sub) => describeCommandTree(sub, path))
  };
}

/** Builds the full manifest for a root program: top-level metadata + every command/subcommand. */
export function describeProgram(program: Command): ProgramManifest {
  return {
    name: program.name(),
    version: getPackageVersionSync(),
    description: program.description() ?? '',
    globalOptions: toManifestOptions(program),
    commands: program.commands.map((cmd) => describeCommandTree(cmd))
  };
}

/** Finds a command (or subcommand) manifest by its space-separated path, e.g. "rules create". */
export function findCommandManifestByPath(manifest: ProgramManifest, path: string): CommandManifest | undefined {
  const segments = path.trim().split(/\s+/).filter(Boolean);
  if (segments.length === 0) return undefined;
  let pool = manifest.commands;
  let found: CommandManifest | undefined;
  for (const seg of segments) {
    found = pool.find((c) => c.name === seg || c.aliases.includes(seg));
    if (!found) return undefined;
    pool = found.subcommands;
  }
  return found;
}
