import { describe, expect, test } from 'bun:test';
import { Command } from 'commander';
import { describeCommandTree, describeProgram, findCommandManifestByPath } from './command-manifest.js';

function buildFixtureProgram(): Command {
  const root = new Command('m365-agent-cli').description('root desc');
  root.option('--read-only', 'read-only mode');

  const mail = new Command('mail')
    .description('mail desc')
    .option('--reply <id>', 'reply to a message')
    .option('--json', 'output json')
    .argument('[id]', 'optional id arg');
  root.addCommand(mail);

  const rules = new Command('rules').description('rules desc');
  const rulesCreate = new Command('create')
    .description('create a rule')
    .requiredOption('--name <name>', 'rule name')
    .option('--priority <n>', 'priority', '10');
  rules.addCommand(rulesCreate);
  rules.addCommand(new Command('list').alias('ls').description('list rules'));
  root.addCommand(rules);

  return root;
}

describe('describeCommandTree', () => {
  test('reflects name, description, options, and arguments', () => {
    const root = buildFixtureProgram();
    const mail = root.commands.find((c) => c.name() === 'mail')!;
    const manifest = describeCommandTree(mail);

    expect(manifest.name).toBe('mail');
    expect(manifest.path).toBe('mail');
    expect(manifest.description).toBe('mail desc');
    expect(manifest.subcommands).toEqual([]);

    const reply = manifest.options.find((o) => o.long === '--reply');
    expect(reply).toBeDefined();
    expect(reply?.valueRequired).toBe(true);
    expect(reply?.mandatory).toBe(false);

    expect(manifest.arguments).toHaveLength(1);
    expect(manifest.arguments[0]).toEqual({
      name: 'id',
      description: 'optional id arg',
      required: false,
      variadic: false
    });
  });

  test('recurses into subcommands and builds dotted-space paths', () => {
    const root = buildFixtureProgram();
    const rules = root.commands.find((c) => c.name() === 'rules')!;
    const manifest = describeCommandTree(rules);

    expect(manifest.subcommands.map((s) => s.name)).toEqual(['create', 'list']);
    const create = manifest.subcommands.find((s) => s.name === 'create')!;
    expect(create.path).toBe('rules create');
    const nameOpt = create.options.find((o) => o.long === '--name');
    expect(nameOpt?.mandatory).toBe(true);
    expect(create.options.find((o) => o.long === '--priority')?.defaultValue).toBe('10');

    const list = manifest.subcommands.find((s) => s.name === 'list')!;
    expect(list.aliases).toEqual(['ls']);
  });

  test('excludes the auto-injected --help/--version options', () => {
    const root = buildFixtureProgram();
    root.version('1.2.3');
    const manifest = describeCommandTree(root);
    const flags = manifest.options.map((o) => o.flags);
    expect(flags).not.toContain('-h, --help');
    expect(flags).not.toContain('-V, --version');
  });

  test('reports an argument default value (bug regression)', () => {
    const cmd = new Command('mail-cmd').argument('[folder]', 'folder name', 'inbox');
    const manifest = describeCommandTree(cmd);
    expect(manifest.arguments[0]).toEqual({
      name: 'folder',
      description: 'folder name',
      required: false,
      variadic: false,
      defaultValue: 'inbox'
    });
  });

  test('a --no-* option with no positive counterpart implicitly defaults to true (bug regression)', () => {
    const cmd = new Command('approvals-list').option('--no-expand', 'Skip $expand');
    const manifest = describeCommandTree(cmd);
    const opt = manifest.options.find((o) => o.long === '--no-expand');
    expect(opt?.negate).toBe(true);
    expect(opt?.defaultValue).toBe(true);
  });

  test('a --no-* option with a registered positive counterpart does not get a forced true default', () => {
    const cmd = new Command('thing').option('--expand', 'expand it').option('--no-expand', 'do not expand');
    const manifest = describeCommandTree(cmd);
    const noExpand = manifest.options.find((o) => o.long === '--no-expand');
    expect(noExpand?.negate).toBe(true);
    expect(noExpand?.defaultValue).toBeUndefined();
  });

  test('a non-negated option has negate: false', () => {
    const cmd = new Command('thing').option('--force', 'force it');
    const manifest = describeCommandTree(cmd);
    expect(manifest.options.find((o) => o.long === '--force')?.negate).toBe(false);
  });
});

describe('describeProgram', () => {
  test('builds top-level metadata and full command list', () => {
    const root = buildFixtureProgram();
    const manifest = describeProgram(root);
    expect(manifest.name).toBe('m365-agent-cli');
    expect(manifest.description).toBe('root desc');
    expect(manifest.globalOptions.map((o) => o.long)).toEqual(['--read-only']);
    expect(manifest.commands.map((c) => c.name)).toEqual(['mail', 'rules']);
  });
});

describe('findCommandManifestByPath', () => {
  test('resolves a top-level command', () => {
    const manifest = describeProgram(buildFixtureProgram());
    expect(findCommandManifestByPath(manifest, 'mail')?.name).toBe('mail');
  });

  test('resolves a nested subcommand path', () => {
    const manifest = describeProgram(buildFixtureProgram());
    const found = findCommandManifestByPath(manifest, 'rules create');
    expect(found?.path).toBe('rules create');
  });

  test('resolves via an alias', () => {
    const manifest = describeProgram(buildFixtureProgram());
    expect(findCommandManifestByPath(manifest, 'rules ls')?.name).toBe('list');
  });

  test('returns undefined for an unknown path', () => {
    const manifest = describeProgram(buildFixtureProgram());
    expect(findCommandManifestByPath(manifest, 'nope')).toBeUndefined();
    expect(findCommandManifestByPath(manifest, 'rules nope')).toBeUndefined();
  });

  test('returns undefined for an empty path', () => {
    const manifest = describeProgram(buildFixtureProgram());
    expect(findCommandManifestByPath(manifest, '   ')).toBeUndefined();
  });
});
