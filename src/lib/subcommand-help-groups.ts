import type { Command, Help } from 'commander';

const OTHER_SECTION_TITLE = 'Other commands';

/** Curated sections: exact subcommand names (Commander `.name()`), not argument tokens. */
const SUBCOMMAND_GROUPS_BY_PARENT: Record<
  string,
  readonly { readonly title: string; readonly commands: readonly string[] }[]
> = {
  calendar: [{ title: 'Calendar', commands: ['list', 'create'] }],
  files: [
    {
      title: 'Browse and read',
      commands: [
        'list',
        'delta',
        'search',
        'recent',
        'shared-with-me',
        'preview',
        'thumbnails',
        'meta',
        'activities',
        'analytics'
      ]
    },
    {
      title: 'Upload and download',
      commands: ['upload', 'upload-large', 'download']
    },
    {
      title: 'Item lifecycle',
      commands: ['copy', 'move', 'delete', 'permanent-delete', 'restore-deleted', 'restore', 'versions', 'convert']
    },
    {
      title: 'Sharing and permissions',
      commands: ['share', 'invite', 'permissions', 'permission-remove', 'permission-update']
    },
    {
      title: 'Checkout and list metadata',
      commands: ['checkout', 'checkin', 'list-item', 'follow', 'unfollow']
    },
    {
      title: 'Sensitivity and retention',
      commands: ['sensitivity-assign', 'sensitivity-extract', 'retention-label', 'retention-label-remove']
    }
  ],
  teams: [
    {
      title: 'Teams and channels',
      commands: [
        'list',
        'get',
        'primary-channel',
        'channel-files-folder',
        'channels',
        'all-channels',
        'incoming-channels',
        'channel-get',
        'channel-members',
        'members',
        'team-member-add',
        'channel-member-add'
      ]
    },
    {
      title: 'Channel tabs',
      commands: ['tabs', 'tab-get', 'tab-create', 'tab-update', 'tab-delete']
    },
    {
      title: 'Channel messages',
      commands: [
        'messages',
        'channel-message-get',
        'channel-message-send',
        'message-replies',
        'channel-message-reply',
        'channel-message-patch',
        'channel-message-delete',
        'channel-message-react'
      ]
    },
    {
      title: 'Chats and chat messages',
      commands: [
        'chats',
        'chat-get',
        'chat-messages',
        'chat-message-get',
        'chat-message-replies',
        'chat-message-send',
        'chat-message-reply',
        'chat-message-patch',
        'chat-message-reply-patch',
        'chat-message-delete',
        'chat-message-react',
        'chat-pinned',
        'chat-members',
        'chat-create',
        'chat-member-add'
      ]
    },
    {
      title: 'Apps and installations',
      commands: [
        'apps',
        'app-catalog',
        'app-catalog-get',
        'app-get',
        'app-add',
        'app-patch',
        'app-upgrade',
        'app-delete',
        'chat-apps',
        'chat-app-get',
        'chat-app-add',
        'chat-app-patch',
        'chat-app-upgrade',
        'chat-app-delete',
        'user-apps',
        'user-app-get',
        'user-app-add',
        'user-app-delete'
      ]
    },
    { title: 'Activity', commands: ['activity-notify'] }
  ],
  copilot: [
    {
      title: 'Search and roots',
      commands: ['retrieval', 'search', 'search-next', 'root-get', 'root-patch']
    },
    {
      title: 'Conversations and chat',
      commands: [
        'conversations-list',
        'conversation-get',
        'conversation-patch',
        'conversation-delete',
        'conversation-delete-by-thread',
        'conversation-create',
        'conversations-count',
        'messages-list',
        'message-get',
        'message-create',
        'message-patch',
        'message-delete',
        'messages-count',
        'chat',
        'chat-stream'
      ]
    },
    {
      title: 'Agents and package counts',
      commands: ['agents-list', 'agent-get', 'agents-count', 'packages-count', 'package-zip-delete']
    },
    {
      title: 'User settings',
      commands: [
        'settings-get',
        'settings-patch',
        'settings-delete',
        'settings-people-get',
        'settings-people-patch',
        'settings-people-delete',
        'settings-enhanced-personalization-get',
        'settings-enhanced-personalization-patch',
        'settings-enhanced-personalization-delete'
      ]
    },
    {
      title: 'Admin configuration',
      commands: [
        'admin-settings-get',
        'admin-settings-patch',
        'admin-settings-delete',
        'admin-limited-mode-get',
        'admin-limited-mode-patch',
        'admin-limited-mode-delete',
        'admin-nav-get',
        'admin-nav-patch',
        'admin-nav-delete',
        'admin-catalog-get',
        'admin-catalog-patch',
        'admin-catalog-delete'
      ]
    },
    {
      title: 'Communications',
      commands: [
        'communications-get',
        'communications-patch',
        'communications-delete',
        'interaction-history-nav-get',
        'interaction-history-nav-patch',
        'interaction-history-nav-delete'
      ]
    },
    {
      title: 'Meeting insights and export',
      commands: [
        'meeting-insights-list',
        'meeting-insight-get',
        'meeting-insights-count',
        'meeting-insight-create',
        'meeting-insight-patch',
        'meeting-insight-delete',
        'interactions-export',
        'interactions-export-tenant'
      ]
    },
    {
      title: 'Reports, packages, activity, AI user',
      commands: ['reports', 'packages', 'activity-feed', 'ai-user', 'notify-help']
    }
  ],
  planner: [
    {
      title: 'Current user',
      commands: ['get-me', 'update-me', 'delta']
    },
    {
      title: 'Plans and favorites',
      commands: [
        'create-plan',
        'get-plan',
        'list-plans',
        'delete-plan',
        'update-plan',
        'plan-archive',
        'plan-unarchive',
        'plan-usage-rights',
        'move-plan-to-container',
        'get-plan-details',
        'update-plan-details',
        'delete-plan-details',
        'list-favorite-plans',
        'add-favorite',
        'remove-favorite',
        'list-recent-plans',
        'list-roster-plans',
        'list-user-plans'
      ]
    },
    {
      title: 'Buckets',
      commands: ['create-bucket', 'list-buckets', 'delete-bucket', 'update-bucket']
    },
    {
      title: 'Tasks',
      commands: [
        'create-task',
        'delete-task',
        'get-task',
        'list-tasks',
        'list-my-tasks',
        'list-my-day-tasks',
        'list-user-tasks',
        'update-task',
        'get-task-details',
        'update-task-details',
        'delete-task-details',
        'get-task-board',
        'update-task-board',
        'add-checklist-item',
        'remove-checklist-item',
        'update-checklist-item',
        'add-reference',
        'remove-reference'
      ]
    },
    { title: 'Roster and task extensions', commands: ['roster', 'tasks'] }
  ],
  excel: [
    {
      title: 'Session and workbook',
      commands: ['session-create', 'session-close', 'session-refresh', 'workbook-get', 'application-calculate']
    },
    {
      title: 'Worksheets',
      commands: [
        'worksheets',
        'worksheet-get',
        'worksheet-add',
        'worksheet-delete',
        'worksheet-update',
        'worksheet-names',
        'worksheet-name-get'
      ]
    },
    {
      title: 'Range',
      commands: ['range', 'range-patch', 'range-clear', 'used-range']
    },
    {
      title: 'Tables',
      commands: [
        'tables',
        'table-get',
        'table-add',
        'table-delete',
        'table-patch',
        'table-columns',
        'table-column-get',
        'table-column-patch',
        'table-rows',
        'table-rows-add',
        'table-row-patch',
        'table-row-delete'
      ]
    },
    {
      title: 'Pivot tables and charts',
      commands: [
        'pivot-tables',
        'pivot-table-get',
        'pivot-table-create',
        'pivot-table-delete',
        'pivot-table-patch',
        'pivot-table-refresh',
        'pivot-tables-refresh-all',
        'charts',
        'chart-create',
        'chart-delete',
        'chart-patch'
      ]
    },
    {
      title: 'Comments and names',
      commands: [
        'comments-list',
        'comments-get',
        'comments-create',
        'comments-patch',
        'comments-reply',
        'names',
        'name-get'
      ]
    }
  ],
  bookings: [
    {
      title: 'Businesses',
      commands: [
        'businesses',
        'business-create',
        'business-delete',
        'business-get',
        'business-update',
        'business-publish',
        'business-unpublish'
      ]
    },
    {
      title: 'Appointments',
      commands: [
        'appointments',
        'appointment',
        'appointment-create',
        'appointment-update',
        'appointment-delete',
        'appointment-cancel',
        'calendar-view'
      ]
    },
    {
      title: 'Customers and services',
      commands: [
        'customers',
        'customer',
        'customer-create',
        'customer-update',
        'customer-delete',
        'services',
        'service-create',
        'service-get',
        'service-update',
        'service-delete'
      ]
    },
    {
      title: 'Staff',
      commands: ['staff', 'staff-create', 'staff-get', 'staff-update', 'staff-delete', 'staff-availability']
    },
    {
      title: 'Catalog and questions',
      commands: [
        'currencies',
        'currency-get',
        'custom-questions',
        'custom-question',
        'custom-question-create',
        'custom-question-update',
        'custom-question-delete'
      ]
    }
  ],
  onenote: [
    {
      title: 'Notebooks and structure',
      commands: ['notebooks', 'notebook', 'sections', 'section', 'section-group']
    },
    {
      title: 'Pages',
      commands: [
        'pages',
        'page',
        'list-pages',
        'create-page',
        'create-page-multipart',
        'patch-page-content',
        'patch-page-content-multipart',
        'delete-page',
        'copy-page',
        'page-preview',
        'content',
        'export'
      ]
    },
    {
      title: 'Resources and operations',
      commands: ['get-resource', 'resource-download', 'operation']
    }
  ],
  'outlook-graph': [
    {
      title: 'Mail folders',
      commands: [
        'list-folders',
        'get-folder',
        'create-folder',
        'update-folder',
        'delete-folder',
        'child-folders',
        'list-mail',
        'list-messages',
        'messages-delta',
        'get-message',
        'patch-message',
        'delete-message',
        'move-message',
        'copy-message',
        'send-mail',
        'send-message',
        'create-reply',
        'create-reply-all',
        'create-forward',
        'list-message-attachments',
        'get-message-attachment',
        'download-message-attachment'
      ]
    },
    {
      title: 'Contacts',
      commands: ['list-contacts', 'get-contact', 'create-contact', 'update-contact', 'delete-contact']
    }
  ],
  contacts: [
    {
      title: 'Contacts',
      commands: ['list', 'show', 'create', 'update', 'delete', 'delta', 'search', 'merge-suggestions']
    },
    {
      title: 'Folders and extensions',
      commands: ['folders', 'folder', 'attachments', 'photo', 'extension']
    }
  ]
};

function getSubcommandGroupConfig(
  parentName: string
): readonly { readonly title: string; readonly commands: readonly string[] }[] | undefined {
  return SUBCOMMAND_GROUPS_BY_PARENT[parentName];
}

/**
 * When the parent has a registry entry, partition visible subcommands into sections plus any unlisted names.
 */
export function buildNestedCommandSections(
  parentCmd: Command,
  helper: Help
): { title: string; commands: Command[] }[] | null {
  const config = getSubcommandGroupConfig(parentCmd.name());
  if (!config) return null;

  const visible = helper.visibleCommands(parentCmd);
  const byName = new Map<string, Command>();
  for (const c of visible) {
    byName.set(c.name(), c);
  }
  const used = new Set<string>();
  const sections: { title: string; commands: Command[] }[] = [];

  for (const g of config) {
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
