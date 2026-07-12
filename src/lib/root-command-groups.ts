import type { Command, Help } from 'commander';

/** Root `m365-agent-cli --help` sections: primary command names only (not aliases). */
const ROOT_COMMAND_GROUPS: readonly { readonly title: string; readonly commands: readonly string[] }[] = [
  {
    title: 'Sign-in and CLI',
    commands: ['whoami', 'login', 'update', 'verify-token', 'describe']
  },
  {
    title: 'Calendar and meetings',
    commands: [
      'calendar',
      'graph-calendar',
      'create-event',
      'update-event',
      'delete-event',
      'respond',
      'forward-event',
      'counter',
      'findtime',
      'schedule',
      'suggest',
      'meeting',
      'rooms'
    ]
  },
  {
    title: 'Mail and mailbox',
    commands: [
      'mail',
      'outlook-graph',
      'folders',
      'send',
      'drafts',
      'contacts',
      'outlook-categories',
      'oof',
      'auto-reply',
      'mailbox-settings',
      'rules',
      'delegates'
    ]
  },
  {
    title: 'Files and content',
    commands: ['files', 'word', 'excel', 'powerpoint', 'onenote', 'sharepoint', 'pages']
  },
  {
    title: 'Teams and work',
    commands: ['teams', 'planner', 'todo', 'groups', 'bookings', 'approvals']
  },
  {
    title: 'People and organization',
    commands: ['find', 'people', 'org', 'presence', 'insights']
  },
  {
    title: 'Copilot and Viva',
    commands: ['copilot', 'viva']
  },
  {
    title: 'Automation and advanced Graph',
    commands: ['graph', 'graph-search', 'subscribe', 'subscriptions', 'serve']
  }
];

const OTHER_SECTION_TITLE = 'Other commands';

/**
 * Partition visible root subcommands into grouped sections plus any commands missing from the registry.
 */
export function buildRootCommandSections(rootCmd: Command, helper: Help): { title: string; commands: Command[] }[] {
  const visible = helper.visibleCommands(rootCmd);
  const byName = new Map<string, Command>();
  for (const c of visible) {
    byName.set(c.name(), c);
  }
  const used = new Set<string>();
  const sections: { title: string; commands: Command[] }[] = [];

  for (const g of ROOT_COMMAND_GROUPS) {
    const commands: Command[] = [];
    for (const name of g.commands) {
      const cmd = byName.get(name);
      if (cmd) {
        commands.push(cmd);
        used.add(name);
      }
    }
    if (commands.length > 0) {
      sections.push({ title: g.title, commands });
    }
  }

  const other: Command[] = [];
  for (const c of visible) {
    if (!used.has(c.name())) {
      other.push(c);
    }
  }
  if (other.length > 0) {
    sections.push({ title: OTHER_SECTION_TITLE, commands: other });
  }

  return sections;
}
