import { readFile, writeFile } from 'node:fs/promises';
import { basename } from 'node:path';
import { Command } from 'commander';
import { resolveAuth } from '../lib/auth.js';
import { getEmail } from '../lib/ews-client.js';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import {
  addChecklistItem,
  addLinkedResource,
  createTask,
  createTaskFileAttachment,
  createTaskLinkedResource,
  createTaskReferenceAttachment,
  createTodoList,
  deleteAttachment,
  deleteChecklistItem,
  deleteTask,
  deleteTaskLinkedResource,
  deleteTaskOpenExtension,
  deleteTodoList,
  deleteTodoListOpenExtension,
  getChecklistItem,
  getTask,
  getTaskAttachment,
  getTaskAttachmentContent,
  getTaskLinkedResource,
  getTaskOpenExtension,
  getTasks,
  getTodoList,
  getTodoListOpenExtension,
  getTodoLists,
  getTodoTasksDeltaPage,
  listAttachments,
  listTaskChecklistItems,
  listTaskLinkedResources,
  listTaskOpenExtensions,
  listTodoListOpenExtensions,
  removeLinkedResourceByWebUrl,
  setTaskOpenExtension,
  setTodoListOpenExtension,
  type TodoImportance,
  type TodoLinkedResource,
  type TodoList,
  type TodoStatus,
  type TodoTask,
  type TodoTasksQueryOptions,
  updateChecklistItem,
  updateTask,
  updateTaskLinkedResource,
  updateTaskOpenExtension,
  updateTodoList,
  updateTodoListOpenExtension,
  uploadLargeFileAttachment
} from '../lib/todo-client.js';
import { checkReadOnly } from '../lib/utils.js';

function fmtDate(iso: string | undefined): string {
  if (!iso) return '';
  try {
    return new Date(iso).toLocaleString('en-US', {
      timeZone: 'UTC',
      month: 'short',
      day: 'numeric',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
      hour12: false
    });
  } catch {
    return iso;
  }
}

function fmtDT(d: { dateTime: string; timeZone: string } | undefined): string {
  if (!d) return '';
  try {
    return new Date(d.dateTime).toLocaleString('en-US', {
      timeZone: d.timeZone || 'UTC',
      month: 'short',
      day: 'numeric',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
      hour12: false
    });
  } catch {
    return d.dateTime;
  }
}

function impEmoji(i: TodoImportance | undefined): string {
  return i === 'high' ? '\u{1F534}' : i === 'low' ? '\u{1F535}' : '\u26AA';
}
function stsEmoji(s: TodoStatus | undefined): string {
  switch (s) {
    case 'completed':
      return '\u2705';
    case 'inProgress':
      return '\u{1F504}';
    case 'waitingOnOthers':
      return '\u23F3';
    case 'deferred':
      return '\u{1F4E6}';
    case 'notStarted':
      return '\u2B1B';
    default:
      return '\u26AA';
  }
}

function emailUrl(id: string): string {
  return `https://outlook.office365.com/mail/${encodeURIComponent(id)}`;
}

function linkedTitle(lr: Pick<TodoLinkedResource, 'displayName' | 'description'>): string {
  const t = (lr.displayName || lr.description || '').trim();
  return t || '(link)';
}

async function resolveListId(
  token: string,
  nameOrId: string,
  user?: string
): Promise<{ listId: string; listDisplay: string }> {
  const listsR = await getTodoLists(token, user);
  if (!listsR.ok || !listsR.data) {
    console.error(`Error: ${listsR.error?.message}`);
    process.exit(1);
  }

  const matched = listsR.data.find(
    (l) =>
      l.id === nameOrId ||
      l.displayName.toLowerCase() === nameOrId.toLowerCase() ||
      l.wellknownListName?.toLowerCase() === nameOrId.toLowerCase()
  );

  if (matched) {
    return { listId: matched.id, listDisplay: matched.displayName };
  }

  const s = await getTodoList(token, nameOrId, user);
  if (!s.ok || !s.data) {
    console.error(`List not found: "${nameOrId}".`);
    console.error('Use "m365-agent-cli todo lists".');
    process.exit(1);
  }
  return { listId: s.data.id, listDisplay: s.data.displayName };
}

export const todoCommand = new Command('todo').description('Manage Microsoft To-Do tasks');

todoCommand
  .command('lists')
  .description('List all To-Do task lists')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user or shared mailbox (Graph delegation)')
  .action(async (opts: { json?: boolean; token?: string; identity?: string; user?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const result = await getTodoLists(auth.token!, opts.user);
    if (!result.ok || !result.data) {
      console.error(`Error: ${result.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(result.data, null, 2));
      return;
    }
    const lists: TodoList[] = result.data;
    if (lists.length === 0) {
      console.log('No task lists found.');
      return;
    }
    console.log(`\nTo-Do Lists (${lists.length}):\n`);
    for (const l of lists) {
      const tag = l.isShared ? ' [shared]' : l.isOwner === false ? ' [shared with me]' : '';
      console.log(`  ${l.displayName}${tag}`);
      console.log(`    ID: ${l.id}`);
      if (l.wellknownListName) console.log(`    Well-known: ${l.wellknownListName}`);
      console.log('');
    }
  });

todoCommand
  .command('get')
  .description('List tasks in a list, or show a single task')
  .option('-l, --list <name|id>', 'List name or ID (default: Tasks)', 'Tasks')
  .option('-t, --task <id>', 'Show detail for a specific task ID')
  .option('--status <status>', 'Filter by status: notStarted, inProgress, completed, waitingOnOthers, deferred')
  .option('--importance <importance>', 'Filter by importance: low, normal, high')
  .option('--filter <odata>', 'Raw OData $filter (not combined with --status / --importance; see Graph todoTask)')
  .option('--orderby <expr>', 'OData $orderby (e.g. lastModifiedDateTime desc)')
  .option('--select <fields>', 'OData $select (comma-separated field names)')
  .option('--top <n>', 'Page size; when set, returns a single page (no auto follow nextLink)')
  .option('--skip <n>', 'OData $skip (single-page request; combine with --top for paging)')
  .option('--expand <expr>', 'OData $expand (e.g. attachments)')
  .option('--count', 'Add $count=true (single-page response)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user or shared mailbox (Graph delegation)')
  .action(
    async (opts: {
      list?: string;
      task?: string;
      status?: string;
      importance?: string;
      filter?: string;
      orderby?: string;
      select?: string;
      top?: string;
      skip?: string;
      expand?: string;
      count?: boolean;
      json?: boolean;
      token?: string;
      identity?: string;
      user?: string;
    }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }

      const listName = opts.list || 'Tasks';
      const { listId, listDisplay } = await resolveListId(auth.token!, listName, opts.user);

      if (opts.task) {
        const r = await getTask(
          auth.token!,
          listId,
          opts.task,
          opts.user,
          opts.select ? { select: opts.select } : undefined
        );
        if (!r.ok || !r.data) {
          console.error(`Error: ${r.error?.message}`);
          process.exit(1);
        }
        const t: TodoTask = r.data;
        if (opts.json) {
          console.log(JSON.stringify(t, null, 2));
          return;
        }
        const hr = '\u2500'.repeat(60);
        console.log(`\n${hr}`);
        console.log(`Title:       ${t.title}`);
        console.log(`Status:      ${stsEmoji(t.status)} ${t.status}`);
        console.log(`Importance:  ${impEmoji(t.importance)} ${t.importance}`);
        if (t.categories?.length) console.log(`Categories:  ${t.categories.join(', ')}`);
        if (t.dueDateTime) console.log(`Due:         ${fmtDT(t.dueDateTime)} (${t.dueDateTime.timeZone})`);
        if (t.startDateTime) console.log(`Start:       ${fmtDT(t.startDateTime)} (${t.startDateTime.timeZone})`);
        if (t.isReminderOn && t.reminderDateTime) console.log(`Reminder:    ${fmtDT(t.reminderDateTime)}`);
        if (t.completedDateTime) console.log(`Completed:   ${fmtDT(t.completedDateTime)}`);
        if (t.linkedResources?.length) {
          console.log('Linked:');
          for (const lr of t.linkedResources) console.log(`  - ${linkedTitle(lr)}: ${lr.webUrl ?? ''}`);
        }
        if (t.body?.content) {
          console.log(`\n${hr}\n${t.body.content}`);
        }
        if (t.checklistItems?.length) {
          console.log('\nChecklist:');
          for (const item of t.checklistItems)
            console.log(`  ${item.isChecked ? '\u2611' : '\u2610'} ${item.displayName}`);
        }
        console.log(`\n${hr}`);
        console.log(`ID:          ${t.id}`);
        if (t.createdDateTime) console.log(`Created:     ${fmtDate(t.createdDateTime)}`);
        if (t.lastModifiedDateTime) console.log(`Modified:   ${fmtDate(t.lastModifiedDateTime)}`);
        console.log('');
        return;
      }

      if (opts.filter && (opts.status || opts.importance)) {
        console.error('Error: use either --filter or --status/--importance, not both');
        process.exit(1);
      }

      let listQuery: TodoTasksQueryOptions | string | undefined;
      const parseTopSkip = (q: TodoTasksQueryOptions) => {
        if (opts.top !== undefined) {
          const n = parseInt(opts.top, 10);
          if (Number.isNaN(n) || n < 1) {
            console.error('Error: --top must be a positive integer');
            process.exit(1);
          }
          q.top = n;
        }
        if (opts.skip !== undefined) {
          const s = parseInt(opts.skip, 10);
          if (Number.isNaN(s) || s < 0) {
            console.error('Error: --skip must be a non-negative integer');
            process.exit(1);
          }
          q.skip = s;
        }
        if (opts.expand) q.expand = opts.expand;
        if (opts.count) q.count = true;
      };

      if (opts.filter) {
        const q: TodoTasksQueryOptions = { filter: opts.filter };
        if (opts.orderby) q.orderby = opts.orderby;
        if (opts.select) q.select = opts.select;
        parseTopSkip(q);
        listQuery = q;
      } else {
        const filters: string[] = [];
        if (opts.status) {
          const validStatuses: TodoStatus[] = ['notStarted', 'inProgress', 'completed', 'waitingOnOthers', 'deferred'];
          if (!validStatuses.includes(opts.status as TodoStatus)) {
            console.error(`Error: Invalid status "${opts.status}". Valid values: ${validStatuses.join(', ')}`);
            process.exit(1);
          }
          filters.push(`status eq '${opts.status}'`);
        }
        if (opts.importance) {
          const validImportance: TodoImportance[] = ['low', 'normal', 'high'];
          if (!validImportance.includes(opts.importance as TodoImportance)) {
            console.error(
              `Error: Invalid importance "${opts.importance}". Valid values: ${validImportance.join(', ')}`
            );
            process.exit(1);
          }
          filters.push(`importance eq '${opts.importance}'`);
        }
        const filterStr = filters.join(' and ') || undefined;
        if (
          filterStr ||
          opts.orderby ||
          opts.select ||
          opts.top !== undefined ||
          opts.skip !== undefined ||
          opts.expand ||
          opts.count
        ) {
          const q: TodoTasksQueryOptions = {};
          if (filterStr) q.filter = filterStr;
          if (opts.orderby) q.orderby = opts.orderby;
          if (opts.select) q.select = opts.select;
          parseTopSkip(q);
          listQuery = q;
        }
      }

      const result = await getTasks(auth.token!, listId, listQuery, opts.user);
      if (!result.ok || !result.data) {
        console.error(`Error: ${result.error?.message}`);
        process.exit(1);
      }
      const tasks: TodoTask[] = result.data;
      if (opts.json) {
        console.log(JSON.stringify({ list: listDisplay, listId, tasks }, null, 2));
        return;
      }
      if (tasks.length === 0) {
        console.log(`\n${listDisplay}: no tasks found.\n`);
        return;
      }
      console.log(`\n${listDisplay} (${tasks.length} task${tasks.length === 1 ? '' : 's'}):\n`);
      for (const t of tasks) {
        const due = t.dueDateTime ? `\u{1F4C5} ${fmtDT(t.dueDateTime)}` : '';
        console.log(`  ${t.status === 'completed' ? '\u2705' : '  '} ${impEmoji(t.importance)} ${t.title} ${due}`);
        console.log(`      ID: ${t.id}  |  ${t.status || 'no status'}  |  ${t.importance || 'normal'}`);
        if (t.categories?.length) console.log(`      Categories: ${t.categories.join(', ')}`);
        if (t.linkedResources?.length)
          console.log(`      \u21B3 linked: ${t.linkedResources.map((l) => linkedTitle(l)).join(', ')}`);
        console.log('');
      }
    }
  );

todoCommand
  .command('create')
  .description('Create a new task')
  .requiredOption('-t, --title <text>', 'Task title')
  .option('-l, --list <name|id>', 'List name or ID (default: Tasks)', 'Tasks')
  .option('-b, --body <text>', 'Task body/notes')
  .option('-d, --due <ISO-8601>', 'Due date (e.g. 2026-04-15T17:00:00Z)')
  .option('--start <ISO-8601>', 'Start date/time')
  .option('--importance <level>', 'Importance: low, normal, high', 'normal')
  .option('--status <status>', 'Initial status: notStarted, inProgress, waitingOnOthers, deferred', 'notStarted')
  .option('--reminder <ISO-8601>', 'Reminder datetime')
  .option('--timezone <tz>', 'Default time zone for due/start/reminder (e.g. UTC, Eastern Standard Time)', 'UTC')
  .option('--due-tz <tz>', 'Time zone for due only (overrides --timezone)')
  .option('--start-tz <tz>', 'Time zone for start only')
  .option('--reminder-tz <tz>', 'Time zone for reminder only')
  .option('--link <msgId>', 'Link task to an email by message ID')
  .option('--mailbox <email>', 'Delegated or shared mailbox (with --link, for EWS message lookup)')
  .option(
    '--category <name>',
    'Category label (repeatable; To Do uses string categories)',
    (v: string, prev: string[]) => [...prev, v],
    [] as string[]
  )
  .option('--recurrence-json <path>', 'JSON file: Graph patternedRecurrence object')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--ews-identity <name>', 'EWS token cache identity for --link (default: default)')
  .option('--user <email>', 'Target user or shared mailbox for the task (Graph delegation)')
  .action(
    async (
      opts: {
        title: string;
        list?: string;
        body?: string;
        due?: string;
        start?: string;
        importance?: string;
        status?: string;
        reminder?: string;
        timezone?: string;
        dueTz?: string;
        startTz?: string;
        reminderTz?: string;
        link?: string;
        mailbox?: string;
        category?: string[];
        recurrenceJson?: string;
        json?: boolean;
        token?: string;
        identity?: string;
        ewsIdentity?: string;
        user?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }

      const listName = opts.list || 'Tasks';
      const { listId } = await resolveListId(auth.token!, listName, opts.user);

      let recurrence: Record<string, unknown> | undefined;
      if (opts.recurrenceJson) {
        const raw = await readFile(opts.recurrenceJson, 'utf-8');
        recurrence = JSON.parse(raw) as Record<string, unknown>;
      }

      let linkedResources: any[] | undefined;
      if (opts.link) {
        // Do not pass the Graph --token to EWS auth, as they require different tokens
        const ewsAuth = await resolveAuth({ identity: opts.ewsIdentity });
        if (!ewsAuth.success) {
          console.error(`EWS Auth error: ${ewsAuth.error}`);
          process.exit(1);
        }
        const er = await getEmail(ewsAuth.token!, opts.link, opts.mailbox);
        if (!er.ok || !er.data) {
          console.error(`Could not fetch email: ${er.error?.message}`);
          process.exit(1);
        }
        linkedResources = [{ webUrl: emailUrl(er.data.Id), displayName: er.data.Subject || 'Linked email' }];
      }

      const cats = (opts.category ?? []).map((c) => c.trim()).filter(Boolean);
      const result = await createTask(
        auth.token!,
        listId,
        {
          title: opts.title,
          body: opts.body,
          importance: opts.importance as TodoImportance,
          status: opts.status as TodoStatus,
          dueDateTime: opts.due,
          startDateTime: opts.start,
          reminderDateTime: opts.reminder,
          timeZone: opts.timezone,
          dueTimeZone: opts.dueTz,
          startTimeZone: opts.startTz,
          reminderTimeZone: opts.reminderTz,
          isReminderOn: !!opts.reminder,
          linkedResources,
          categories: cats.length ? cats : undefined,
          recurrence
        },
        opts.user
      );
      if (!result.ok || !result.data) {
        console.error(`Error: ${result.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(result.data, null, 2));
      else {
        console.log(`\n\u2705 Task created: "${result.data.title}"`);
        console.log(`   ID: ${result.data.id}`);
        console.log(`   List: ${listName}`);
        if (opts.link) console.log(`   \u21B3 Linked to email`);
        console.log('');
      }
    }
  );

todoCommand
  .command('update')
  .description('Update a task (title, body, due, importance, status, categories)')
  .requiredOption('-l, --list <name|id>', 'List name or ID')
  .requiredOption('-t, --task <id>', 'Task ID')
  .option('--title <text>', 'New title')
  .option('-b, --body <text>', 'New body/notes')
  .option('-d, --due <ISO-8601>', 'Due date (or omit with --clear-due)')
  .option('--clear-due', 'Remove due date')
  .option('--start <ISO-8601>', 'Start date/time (or omit with --clear-start)')
  .option('--clear-start', 'Remove start date/time')
  .option('--importance <level>', 'Importance: low, normal, high')
  .option('--status <status>', 'Status: notStarted, inProgress, completed, waitingOnOthers, deferred')
  .option('--reminder <ISO-8601>', 'Reminder datetime')
  .option('--clear-reminder', 'Turn off reminder')
  .option('--timezone <tz>', 'Default time zone when setting due/start/reminder', 'UTC')
  .option('--due-tz <tz>', 'Time zone for due only')
  .option('--start-tz <tz>', 'Time zone for start only')
  .option('--reminder-tz <tz>', 'Time zone for reminder only')
  .option(
    '--category <name>',
    'Set categories to this list (repeatable; replaces existing categories)',
    (v: string, prev: string[]) => [...prev, v],
    [] as string[]
  )
  .option('--clear-categories', 'Remove all categories')
  .option('--recurrence-json <path>', 'JSON file: patternedRecurrence (replaces recurrence)')
  .option('--clear-recurrence', 'Remove recurrence from the task')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user or shared mailbox (Graph delegation)')
  .action(
    async (
      opts: {
        list: string;
        task: string;
        title?: string;
        body?: string;
        due?: string;
        clearDue?: boolean;
        start?: string;
        clearStart?: boolean;
        importance?: string;
        status?: string;
        reminder?: string;
        clearReminder?: boolean;
        category?: string[];
        clearCategories?: boolean;
        recurrenceJson?: string;
        clearRecurrence?: boolean;
        timezone?: string;
        dueTz?: string;
        startTz?: string;
        reminderTz?: string;
        json?: boolean;
        token?: string;
        identity?: string;
        user?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const { listId } = await resolveListId(auth.token!, opts.list, opts.user);

      if (opts.clearDue && opts.due !== undefined) {
        console.error('Error: use either --due or --clear-due, not both');
        process.exit(1);
      }

      if (opts.clearRecurrence && opts.recurrenceJson) {
        console.error('Error: use either --recurrence-json or --clear-recurrence, not both');
        process.exit(1);
      }
      if (opts.clearStart && opts.start !== undefined) {
        console.error('Error: use either --start or --clear-start, not both');
        process.exit(1);
      }

      const hasField =
        opts.title !== undefined ||
        opts.body !== undefined ||
        opts.due !== undefined ||
        opts.clearDue ||
        opts.start !== undefined ||
        opts.clearStart ||
        opts.importance !== undefined ||
        opts.status !== undefined ||
        opts.reminder !== undefined ||
        opts.clearReminder ||
        (opts.category !== undefined && opts.category.length > 0) ||
        opts.clearCategories ||
        opts.recurrenceJson !== undefined ||
        opts.clearRecurrence;

      if (!hasField) {
        console.error(
          'Error: specify at least one of --title, --body, --due, --clear-due, --start, --clear-start, --importance, --status, --reminder, --clear-reminder, --category, --clear-categories, --recurrence-json, --clear-recurrence'
        );
        process.exit(1);
      }

      if (opts.clearCategories && opts.category !== undefined && opts.category.length > 0) {
        console.error('Error: use either --clear-categories or --category, not both');
        process.exit(1);
      }

      let importance: TodoImportance | undefined;
      if (opts.importance !== undefined) {
        const valid: TodoImportance[] = ['low', 'normal', 'high'];
        if (!valid.includes(opts.importance as TodoImportance)) {
          console.error(`Invalid importance: ${opts.importance}`);
          process.exit(1);
        }
        importance = opts.importance as TodoImportance;
      }
      let status: TodoStatus | undefined;
      if (opts.status !== undefined) {
        const valid: TodoStatus[] = ['notStarted', 'inProgress', 'completed', 'waitingOnOthers', 'deferred'];
        if (!valid.includes(opts.status as TodoStatus)) {
          console.error(`Invalid status: ${opts.status}`);
          process.exit(1);
        }
        status = opts.status as TodoStatus;
      }

      const updateOpts: Parameters<typeof updateTask>[3] = {};
      if (opts.title !== undefined) updateOpts.title = opts.title;
      if (opts.body !== undefined) updateOpts.body = opts.body;
      if (opts.clearDue) updateOpts.dueDateTime = null;
      else if (opts.due !== undefined) updateOpts.dueDateTime = opts.due;
      if (opts.clearStart) updateOpts.startDateTime = null;
      else if (opts.start !== undefined) updateOpts.startDateTime = opts.start;
      if (importance !== undefined) updateOpts.importance = importance;
      if (status !== undefined) updateOpts.status = status;
      if (opts.clearReminder) {
        updateOpts.isReminderOn = false;
        updateOpts.reminderDateTime = null;
      } else if (opts.reminder !== undefined) {
        updateOpts.isReminderOn = true;
        updateOpts.reminderDateTime = opts.reminder;
      }
      if (opts.clearCategories) updateOpts.clearCategories = true;
      else if (opts.category !== undefined && opts.category.length > 0) {
        updateOpts.categories = opts.category.map((c) => c.trim()).filter(Boolean);
      }
      if (opts.clearRecurrence) updateOpts.recurrence = null;
      else if (opts.recurrenceJson) {
        const raw = await readFile(opts.recurrenceJson, 'utf-8');
        updateOpts.recurrence = JSON.parse(raw) as Record<string, unknown>;
      }

      if (opts.due !== undefined || opts.start !== undefined || opts.reminder !== undefined) {
        updateOpts.timeZone = opts.timezone;
        if (opts.dueTz !== undefined) updateOpts.dueTimeZone = opts.dueTz;
        if (opts.startTz !== undefined) updateOpts.startTimeZone = opts.startTz;
        if (opts.reminderTz !== undefined) updateOpts.reminderTimeZone = opts.reminderTz;
      }

      const r = await updateTask(auth.token!, listId, opts.task, updateOpts, opts.user);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else console.log(`\n\u2705 Updated: "${r.data.title}"\n`);
    }
  );

todoCommand
  .command('complete')
  .description('Mark a task as completed')
  .requiredOption('-l, --list <name|id>', 'List name or ID')
  .requiredOption('-t, --task <id>', 'Task ID')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user or shared mailbox (Graph delegation)')
  .action(
    async (
      opts: { list: string; task: string; json?: boolean; token?: string; identity?: string; user?: string },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const { listId } = await resolveListId(auth.token!, opts.list, opts.user);
      // dateTime should not include Z/offset - keep dateTime and timeZone separate
      const nowISO = new Date().toISOString();
      const now = nowISO.replace('Z', '');
      const r = await updateTask(
        auth.token!,
        listId,
        opts.task,
        { status: 'completed', completedDateTime: now },
        opts.user
      );
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else console.log(`\n\u2705 Completed: "${r.data.title}" (${fmtDate(nowISO)})\n`);
    }
  );

todoCommand
  .command('delete')
  .description('Delete a task')
  .requiredOption('-l, --list <name|id>', 'List name or ID')
  .requiredOption('-t, --task <id>', 'Task ID')
  .option('--confirm', 'Skip confirmation prompt')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user or shared mailbox (Graph delegation)')
  .action(
    async (
      opts: { list: string; task: string; confirm?: boolean; token?: string; identity?: string; user?: string },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const { listId, listDisplay: listName } = await resolveListId(auth.token!, opts.list, opts.user);
      const taskR = await getTask(auth.token!, listId, opts.task, opts.user);
      if (!taskR.ok || !taskR.data) {
        console.error(`Task not found: ${taskR.error?.message}`);
        process.exit(1);
      }
      if (!opts.confirm) {
        console.log(`Delete "${taskR.data.title}" from "${listName}"? (ID: ${opts.task})`);
        console.log('Run with --confirm to confirm.');
        process.exit(1);
      }
      const r = await deleteTask(auth.token!, listId, opts.task, opts.user);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(`\n\u{1F5D1}  Deleted: "${taskR.data.title}"\n`);
    }
  );

todoCommand
  .command('add-checklist')
  .description('Add a checklist (subtask) item to a task')
  .requiredOption('-l, --list <name|id>', 'List name or ID')
  .requiredOption('-t, --task <id>', 'Task ID')
  .requiredOption('-n, --name <text>', 'Checklist item text')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user or shared mailbox (Graph delegation)')
  .action(
    async (
      opts: {
        list: string;
        task: string;
        name: string;
        json?: boolean;
        token?: string;
        identity?: string;
        user?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const { listId } = await resolveListId(auth.token!, opts.list, opts.user);
      const r = await addChecklistItem(auth.token!, listId, opts.task, opts.name, opts.user);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else console.log(`\n\u2705 Added: "${r.data.displayName}" (${r.data.id})\n`);
    }
  );

todoCommand
  .command('update-checklist')
  .description('Update a checklist item (rename or check/uncheck)')
  .requiredOption('-l, --list <name|id>', 'List name or ID')
  .requiredOption('-t, --task <id>', 'Task ID')
  .requiredOption('-c, --item <checklistItemId>', 'Checklist item ID')
  .option('-n, --name <text>', 'New display text')
  .option('--checked', 'Mark checked')
  .option('--unchecked', 'Mark unchecked')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user or shared mailbox (Graph delegation)')
  .action(
    async (
      opts: {
        list: string;
        task: string;
        item: string;
        name?: string;
        checked?: boolean;
        unchecked?: boolean;
        json?: boolean;
        token?: string;
        identity?: string;
        user?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      if (opts.checked && opts.unchecked) {
        console.error('Error: use either --checked or --unchecked, not both');
        process.exit(1);
      }
      if (opts.name === undefined && !opts.checked && !opts.unchecked) {
        console.error('Error: specify --name and/or --checked/--unchecked');
        process.exit(1);
      }
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const { listId } = await resolveListId(auth.token!, opts.list, opts.user);
      const patch: { displayName?: string; isChecked?: boolean } = {};
      if (opts.name !== undefined) patch.displayName = opts.name;
      if (opts.checked) patch.isChecked = true;
      if (opts.unchecked) patch.isChecked = false;
      const r = await updateChecklistItem(auth.token!, listId, opts.task, opts.item, patch, opts.user);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else console.log(`\n\u2705 Updated checklist item: "${r.data.displayName}"\n`);
    }
  );

todoCommand
  .command('delete-checklist')
  .description('Delete a checklist item from a task')
  .requiredOption('-l, --list <name|id>', 'List name or ID')
  .requiredOption('-t, --task <id>', 'Task ID')
  .requiredOption('-c, --item <checklistItemId>', 'Checklist item ID')
  .option('--confirm', 'Confirm without prompt')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user or shared mailbox (Graph delegation)')
  .action(
    async (
      opts: {
        list: string;
        task: string;
        item: string;
        confirm?: boolean;
        token?: string;
        identity?: string;
        user?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      if (!opts.confirm) {
        console.log(`Delete checklist item ${opts.item}? Run with --confirm.`);
        process.exit(1);
      }
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const { listId } = await resolveListId(auth.token!, opts.list, opts.user);
      const r = await deleteChecklistItem(auth.token!, listId, opts.task, opts.item, opts.user);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(`\n\u2705 Deleted checklist item: ${opts.item}\n`);
    }
  );

todoCommand
  .command('create-list')
  .description('Create a new To Do list')
  .requiredOption('-n, --name <displayName>', 'List display name')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user or shared mailbox (Graph delegation)')
  .action(
    async (opts: { name: string; json?: boolean; token?: string; identity?: string; user?: string }, cmd: any) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await createTodoList(auth.token!, opts.name, opts.user);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else console.log(`\n\u2705 Created list: "${r.data.displayName}" (${r.data.id})\n`);
    }
  );

todoCommand
  .command('update-list')
  .description('Rename a To Do list')
  .requiredOption('-l, --list <name|id>', 'Current list name or ID')
  .requiredOption('-n, --name <displayName>', 'New display name')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user or shared mailbox (Graph delegation)')
  .action(
    async (
      opts: { list: string; name: string; json?: boolean; token?: string; identity?: string; user?: string },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const { listId } = await resolveListId(auth.token!, opts.list, opts.user);
      const r = await updateTodoList(auth.token!, listId, opts.name, opts.user);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else console.log(`\n\u2705 Renamed list to: "${r.data.displayName}"\n`);
    }
  );

todoCommand
  .command('delete-list')
  .description('Delete a To Do list')
  .requiredOption('-l, --list <name|id>', 'List name or ID')
  .option('--confirm', 'Confirm deletion')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user or shared mailbox (Graph delegation)')
  .action(
    async (
      opts: { list: string; confirm?: boolean; json?: boolean; token?: string; identity?: string; user?: string },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const { listId, listDisplay } = await resolveListId(auth.token!, opts.list, opts.user);
      if (!opts.confirm) {
        console.log(`Delete list "${listDisplay}"? (ID: ${listId})`);
        console.log('Run with --confirm to confirm.');
        process.exit(1);
      }
      const r = await deleteTodoList(auth.token!, listId, opts.user);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify({ deleted: listId }, null, 2));
      else console.log(`\n\u2705 Deleted list: "${listDisplay}"\n`);
    }
  );

todoCommand
  .command('list-attachments')
  .description('List attachments on a task')
  .requiredOption('-l, --list <name|id>', 'List name or ID')
  .requiredOption('-t, --task <id>', 'Task ID')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user or shared mailbox (Graph delegation)')
  .action(
    async (opts: { list: string; task: string; json?: boolean; token?: string; identity?: string; user?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const { listId } = await resolveListId(auth.token!, opts.list, opts.user);
      const r = await listAttachments(auth.token!, listId, opts.task, opts.user);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else {
        for (const a of r.data) {
          console.log(`- ${a.name || a.id} (${a.id})${a.size != null ? ` ${a.size} bytes` : ''}`);
        }
      }
    }
  );

todoCommand
  .command('add-attachment')
  .description('Attach a small file to a task (base64 upload; Graph size limits apply)')
  .requiredOption('-l, --list <name|id>', 'List name or ID')
  .requiredOption('-t, --task <id>', 'Task ID')
  .requiredOption('-f, --file <path>', 'Local file path')
  .option('--name <filename>', 'Attachment name (default: file basename)')
  .option('--content-type <mime>', 'MIME type (default: application/octet-stream)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user or shared mailbox (Graph delegation)')
  .action(
    async (
      opts: {
        list: string;
        task: string;
        file: string;
        name?: string;
        contentType?: string;
        json?: boolean;
        token?: string;
        identity?: string;
        user?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const { listId } = await resolveListId(auth.token!, opts.list, opts.user);
      const buf = await readFile(opts.file);
      const b64 = buf.toString('base64');
      const attName = opts.name?.trim() || basename(opts.file);
      const ct = opts.contentType?.trim() || 'application/octet-stream';
      const r = await createTaskFileAttachment(auth.token!, listId, opts.task, attName, b64, ct, opts.user);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else console.log(`\n\u2705 Attached: ${r.data.name || r.data.id} (${r.data.id})\n`);
    }
  );

todoCommand
  .command('delete-attachment')
  .description('Remove an attachment from a task')
  .requiredOption('-l, --list <name|id>', 'List name or ID')
  .requiredOption('-t, --task <id>', 'Task ID')
  .requiredOption('-a, --attachment <attachmentId>', 'Attachment ID')
  .option('--confirm', 'Confirm without prompt')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user or shared mailbox (Graph delegation)')
  .action(
    async (
      opts: {
        list: string;
        task: string;
        attachment: string;
        confirm?: boolean;
        token?: string;
        identity?: string;
        user?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      if (!opts.confirm) {
        console.log(`Delete attachment ${opts.attachment}? Run with --confirm.`);
        process.exit(1);
      }
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const { listId } = await resolveListId(auth.token!, opts.list, opts.user);
      const r = await deleteAttachment(auth.token!, listId, opts.task, opts.attachment, opts.user);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(`\n\u2705 Deleted attachment: ${opts.attachment}\n`);
    }
  );

todoCommand
  .command('get-attachment')
  .description('Fetch metadata for one task attachment')
  .requiredOption('-l, --list <name|id>', 'List name or ID')
  .requiredOption('-t, --task <id>', 'Task ID')
  .requiredOption('-a, --attachment <attachmentId>', 'Attachment ID')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user or shared mailbox (Graph delegation)')
  .action(
    async (opts: {
      list: string;
      task: string;
      attachment: string;
      json?: boolean;
      token?: string;
      identity?: string;
      user?: string;
    }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const { listId } = await resolveListId(auth.token!, opts.list, opts.user);
      const r = await getTaskAttachment(auth.token!, listId, opts.task, opts.attachment, opts.user);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else {
        const a = r.data;
        console.log(`Name: ${a.name || '(unnamed)'}`);
        console.log(`ID: ${a.id}`);
        if (a.contentType) console.log(`Type: ${a.contentType}`);
        if (a.size != null) console.log(`Size: ${a.size}`);
      }
    }
  );

todoCommand
  .command('download-attachment')
  .description(
    'Download file attachment bytes (Graph GET .../attachments/{id}/$value); not for reference/URL attachments'
  )
  .requiredOption('-l, --list <name|id>', 'List name or ID')
  .requiredOption('-t, --task <id>', 'Task ID')
  .requiredOption('-a, --attachment <attachmentId>', 'Attachment ID')
  .requiredOption('-o, --output <path>', 'Write file to this path')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user or shared mailbox (Graph delegation)')
  .action(
    async (opts: {
      list: string;
      task: string;
      attachment: string;
      output: string;
      token?: string;
      identity?: string;
      user?: string;
    }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const { listId } = await resolveListId(auth.token!, opts.list, opts.user);
      const r = await getTaskAttachmentContent(auth.token!, listId, opts.task, opts.attachment, opts.user);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      await writeFile(opts.output, r.data);
      console.log(`Wrote ${r.data.byteLength} bytes to ${opts.output}`);
    }
  );

todoCommand
  .command('add-reference-attachment')
  .description('Add a URL reference attachment (not file bytes)')
  .requiredOption('-l, --list <name|id>', 'List name or ID')
  .requiredOption('-t, --task <id>', 'Task ID')
  .requiredOption('--url <url>', 'Target URL')
  .requiredOption('-n, --name <text>', 'Attachment display name')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user or shared mailbox (Graph delegation)')
  .action(
    async (
      opts: {
        list: string;
        task: string;
        url: string;
        name: string;
        json?: boolean;
        token?: string;
        identity?: string;
        user?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const { listId } = await resolveListId(auth.token!, opts.list, opts.user);
      const r = await createTaskReferenceAttachment(auth.token!, listId, opts.task, opts.name, opts.url, opts.user);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else console.log(`\n\u2705 Reference attachment: ${r.data.name || r.data.id} (${r.data.id})\n`);
    }
  );

todoCommand
  .command('add-linked-resource')
  .description('Merge linked resources on the task (PATCH task). Prefer todo linked-resource create for Graph POST.')
  .requiredOption('-l, --list <name|id>', 'List name or ID')
  .requiredOption('-t, --task <id>', 'Task ID')
  .requiredOption('--url <url>', 'Resource webUrl')
  .option('-d, --description <text>', 'Title (Graph displayName; legacy alias)')
  .option('--display-name <text>', 'Graph displayName (same as -d)')
  .option('--icon <url>', 'Optional icon URL')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user or shared mailbox (Graph delegation)')
  .action(
    async (
      opts: {
        list: string;
        task: string;
        url: string;
        description?: string;
        displayName?: string;
        icon?: string;
        json?: boolean;
        token?: string;
        identity?: string;
        user?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const title = opts.displayName?.trim() || opts.description?.trim();
      if (!title) {
        console.error('Error: specify --display-name or -d/--description (Graph displayName)');
        process.exit(1);
      }
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const { listId } = await resolveListId(auth.token!, opts.list, opts.user);
      const r = await addLinkedResource(
        auth.token!,
        listId,
        opts.task,
        { webUrl: opts.url, displayName: title, iconUrl: opts.icon },
        opts.user
      );
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else console.log(`\n\u2705 Linked resource added. Task: "${r.data.title}"\n`);
    }
  );

todoCommand
  .command('remove-linked-resource')
  .description('Remove a linked resource by matching webUrl')
  .requiredOption('-l, --list <name|id>', 'List name or ID')
  .requiredOption('-t, --task <id>', 'Task ID')
  .requiredOption('--url <url>', 'webUrl to remove')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user or shared mailbox (Graph delegation)')
  .action(
    async (
      opts: {
        list: string;
        task: string;
        url: string;
        json?: boolean;
        token?: string;
        identity?: string;
        user?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const { listId } = await resolveListId(auth.token!, opts.list, opts.user);
      const r = await removeLinkedResourceByWebUrl(auth.token!, listId, opts.task, opts.url, opts.user);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else console.log(`\n\u2705 Removed linked resource matching URL.\n`);
    }
  );

todoCommand
  .command('upload-attachment-large')
  .description('Upload a large file via Graph upload session (chunked PUT)')
  .requiredOption('-l, --list <name|id>', 'List name or ID')
  .requiredOption('-t, --task <id>', 'Task ID')
  .requiredOption('-f, --file <path>', 'Local file path')
  .option('-n, --name <filename>', 'Attachment name (default: file basename)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user or shared mailbox (Graph delegation)')
  .action(
    async (
      opts: {
        list: string;
        task: string;
        file: string;
        name?: string;
        json?: boolean;
        token?: string;
        identity?: string;
        user?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const { listId } = await resolveListId(auth.token!, opts.list, opts.user);
      const r = await uploadLargeFileAttachment(auth.token!, listId, opts.task, opts.file, opts.name, opts.user);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else console.log(`\n\u2705 Large attachment: ${r.data.name || r.data.id} (${r.data.id})\n`);
    }
  );

todoCommand
  .command('delta')
  .description('One page of todo task delta (use -l for first page, or --url for nextLink/deltaLink)')
  .option('-l, --list <name|id>', 'List name or ID (first page only)')
  .option('--url <fullUrl>', 'Full nextLink or deltaLink URL from a previous response')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (first page only; --url encodes scope)')
  .action(async (opts: { list?: string; url?: string; token?: string; identity?: string; user?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = opts.url
      ? await getTodoTasksDeltaPage(auth.token!, '', opts.url)
      : await (async () => {
          if (!opts.list) {
            console.error('Error: specify --list for the first delta page, or --url to follow nextLink/deltaLink');
            process.exit(1);
          }
          const { listId } = await resolveListId(auth.token!, opts.list, opts.user);
          return getTodoTasksDeltaPage(auth.token!, listId, undefined, opts.user);
        })();
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log(JSON.stringify(r.data, null, 2));
  });

todoCommand
  .command('list-checklist-items')
  .description('List checklist items via GET collection (same items as on the task object)')
  .requiredOption('-l, --list <name|id>', 'List name or ID')
  .requiredOption('-t, --task <id>', 'Task ID')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user or shared mailbox (Graph delegation)')
  .action(
    async (opts: { list: string; task: string; json?: boolean; token?: string; identity?: string; user?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const { listId } = await resolveListId(auth.token!, opts.list, opts.user);
      const r = await listTaskChecklistItems(auth.token!, listId, opts.task, opts.user);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else {
        for (const it of r.data) {
          console.log(`${it.isChecked ? '\u2611' : '\u2610'} ${it.displayName} (${it.id})`);
        }
      }
    }
  );

todoCommand
  .command('get-checklist-item')
  .description('Get one checklist item by id (Graph GET checklistItems/{id})')
  .requiredOption('-l, --list <name|id>', 'List name or ID')
  .requiredOption('-t, --task <id>', 'Task ID')
  .requiredOption('-c, --checklist-item <id>', 'Checklist item id')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user or shared mailbox (Graph delegation)')
  .action(
    async (opts: {
      list: string;
      task: string;
      checklistItem: string;
      json?: boolean;
      token?: string;
      identity?: string;
      user?: string;
    }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const { listId } = await resolveListId(auth.token!, opts.list, opts.user);
      const r = await getChecklistItem(auth.token!, listId, opts.task, opts.checklistItem, opts.user);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else {
        const it = r.data;
        console.log(`${it.isChecked ? '\u2611' : '\u2610'} ${it.displayName}`);
        console.log(`ID: ${it.id}`);
        if (it.createdDateTime) console.log(`Created: ${it.createdDateTime}`);
        if (it.checkedDateTime) console.log(`Checked: ${it.checkedDateTime}`);
      }
    }
  );

const todoLinkedResourceCommand = new Command('linked-resource').description(
  'Graph linkedResource endpoints (per-item REST)'
);

todoLinkedResourceCommand
  .command('list')
  .description('List linked resources for a task')
  .requiredOption('-l, --list <name|id>', 'List name or ID')
  .requiredOption('-t, --task <id>', 'Task ID')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user or shared mailbox (Graph delegation)')
  .action(
    async (opts: { list: string; task: string; json?: boolean; token?: string; identity?: string; user?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const { listId } = await resolveListId(auth.token!, opts.list, opts.user);
      const r = await listTaskLinkedResources(auth.token!, listId, opts.task, opts.user);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else {
        for (const lr of r.data) {
          console.log(`- ${linkedTitle(lr)}  ${lr.webUrl ?? ''}  (${lr.id})`);
        }
      }
    }
  );

todoLinkedResourceCommand
  .command('create')
  .description('POST a linkedResource (Graph native create)')
  .requiredOption('-l, --list <name|id>', 'List name or ID')
  .requiredOption('-t, --task <id>', 'Task ID')
  .option('--url <url>', 'webUrl')
  .requiredOption('-n, --name <text>', 'displayName')
  .option('--application-name <text>', 'applicationName')
  .option('--external-id <id>', 'externalId')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user or shared mailbox (Graph delegation)')
  .action(
    async (
      opts: {
        list: string;
        task: string;
        url?: string;
        name: string;
        applicationName?: string;
        externalId?: string;
        json?: boolean;
        token?: string;
        identity?: string;
        user?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const { listId } = await resolveListId(auth.token!, opts.list, opts.user);
      const r = await createTaskLinkedResource(
        auth.token!,
        listId,
        opts.task,
        {
          webUrl: opts.url,
          displayName: opts.name,
          applicationName: opts.applicationName,
          externalId: opts.externalId
        },
        opts.user
      );
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else console.log(`\n\u2705 Linked resource: ${linkedTitle(r.data)} (${r.data.id})\n`);
    }
  );

todoLinkedResourceCommand
  .command('get')
  .description('GET one linked resource by id')
  .requiredOption('-l, --list <name|id>', 'List name or ID')
  .requiredOption('-t, --task <id>', 'Task ID')
  .requiredOption('-i, --id <linkedResourceId>', 'linkedResource id')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user or shared mailbox (Graph delegation)')
  .action(
    async (opts: {
      list: string;
      task: string;
      id: string;
      json?: boolean;
      token?: string;
      identity?: string;
      user?: string;
    }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const { listId } = await resolveListId(auth.token!, opts.list, opts.user);
      const r = await getTaskLinkedResource(auth.token!, listId, opts.task, opts.id, opts.user);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else console.log(JSON.stringify(r.data, null, 2));
    }
  );

todoLinkedResourceCommand
  .command('update')
  .description('PATCH a linked resource')
  .requiredOption('-l, --list <name|id>', 'List name or ID')
  .requiredOption('-t, --task <id>', 'Task ID')
  .requiredOption('-i, --id <linkedResourceId>', 'linkedResource id')
  .option('--url <url>', 'webUrl')
  .option('-n, --name <text>', 'displayName')
  .option('--application-name <text>', 'applicationName')
  .option('--external-id <id>', 'externalId')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user or shared mailbox (Graph delegation)')
  .action(
    async (
      opts: {
        list: string;
        task: string;
        id: string;
        url?: string;
        name?: string;
        applicationName?: string;
        externalId?: string;
        json?: boolean;
        token?: string;
        identity?: string;
        user?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      if (
        opts.url === undefined &&
        opts.name === undefined &&
        opts.applicationName === undefined &&
        opts.externalId === undefined
      ) {
        console.error('Error: specify at least one of --url, --name, --application-name, --external-id');
        process.exit(1);
      }
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const { listId } = await resolveListId(auth.token!, opts.list, opts.user);
      const r = await updateTaskLinkedResource(
        auth.token!,
        listId,
        opts.task,
        opts.id,
        {
          webUrl: opts.url,
          displayName: opts.name,
          applicationName: opts.applicationName,
          externalId: opts.externalId
        },
        opts.user
      );
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else console.log(`\n\u2705 Updated linked resource ${opts.id}\n`);
    }
  );

todoLinkedResourceCommand
  .command('delete')
  .description('DELETE a linked resource by id')
  .requiredOption('-l, --list <name|id>', 'List name or ID')
  .requiredOption('-t, --task <id>', 'Task ID')
  .requiredOption('-i, --id <linkedResourceId>', 'linkedResource id')
  .option('--confirm', 'Confirm without prompt')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user or shared mailbox (Graph delegation)')
  .action(
    async (
      opts: {
        list: string;
        task: string;
        id: string;
        confirm?: boolean;
        token?: string;
        identity?: string;
        user?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      if (!opts.confirm) {
        console.log(`Delete linked resource ${opts.id}? Run with --confirm.`);
        process.exit(1);
      }
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const { listId } = await resolveListId(auth.token!, opts.list, opts.user);
      const r = await deleteTaskLinkedResource(auth.token!, listId, opts.task, opts.id, opts.user);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(`\n\u2705 Deleted linked resource ${opts.id}\n`);
    }
  );

todoCommand.addCommand(todoLinkedResourceCommand);

const todoListExtensionCommand = new Command('list-extension').description(
  'Open type extensions on a task list (Graph)'
);

todoListExtensionCommand
  .command('list')
  .description('List open extensions on a task list')
  .requiredOption('-l, --list <name|id>', 'List name or ID')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user or shared mailbox (Graph delegation)')
  .action(async (opts: { list: string; json?: boolean; token?: string; identity?: string; user?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const { listId } = await resolveListId(auth.token!, opts.list, opts.user);
    const r = await listTodoListOpenExtensions(auth.token!, listId, opts.user);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else {
      for (const ext of r.data) {
        const name = (ext.extensionName as string) || JSON.stringify(ext);
        console.log(`- ${name}`);
      }
    }
  });

todoListExtensionCommand
  .command('get')
  .description('Get one open extension on a list')
  .requiredOption('-l, --list <name|id>', 'List name or ID')
  .requiredOption('-n, --name <id>', 'extensionName')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user or shared mailbox (Graph delegation)')
  .action(
    async (opts: { list: string; name: string; json?: boolean; token?: string; identity?: string; user?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const { listId } = await resolveListId(auth.token!, opts.list, opts.user);
      const r = await getTodoListOpenExtension(auth.token!, listId, opts.name, opts.user);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else console.log(JSON.stringify(r.data, null, 2));
    }
  );

todoListExtensionCommand
  .command('set')
  .description('Create an open extension on a list')
  .requiredOption('-l, --list <name|id>', 'List name or ID')
  .requiredOption('-n, --name <id>', 'extensionName')
  .requiredOption('--json-file <path>', 'JSON object: custom properties')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user or shared mailbox (Graph delegation)')
  .action(
    async (
      opts: {
        list: string;
        name: string;
        jsonFile: string;
        json?: boolean;
        token?: string;
        identity?: string;
        user?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const { listId } = await resolveListId(auth.token!, opts.list, opts.user);
      const raw = await readFile(opts.jsonFile, 'utf-8');
      const data = JSON.parse(raw) as Record<string, unknown>;
      const r = await setTodoListOpenExtension(auth.token!, listId, opts.name, data, opts.user);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else console.log(`\n\u2705 List extension set: ${opts.name}\n`);
    }
  );

todoListExtensionCommand
  .command('update')
  .description('PATCH a list open extension')
  .requiredOption('-l, --list <name|id>', 'List name or ID')
  .requiredOption('-n, --name <id>', 'extensionName')
  .requiredOption('--json-file <path>', 'JSON patch body')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user or shared mailbox (Graph delegation)')
  .action(
    async (
      opts: { list: string; name: string; jsonFile: string; token?: string; identity?: string; user?: string },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const { listId } = await resolveListId(auth.token!, opts.list, opts.user);
      const raw = await readFile(opts.jsonFile, 'utf-8');
      const patch = JSON.parse(raw) as Record<string, unknown>;
      const r = await updateTodoListOpenExtension(auth.token!, listId, opts.name, patch, opts.user);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log('\n\u2705 List extension updated.\n');
    }
  );

todoListExtensionCommand
  .command('delete')
  .description('Delete a list open extension')
  .requiredOption('-l, --list <name|id>', 'List name or ID')
  .requiredOption('-n, --name <id>', 'extensionName')
  .option('--confirm', 'Confirm without prompt')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user or shared mailbox (Graph delegation)')
  .action(
    async (
      opts: { list: string; name: string; confirm?: boolean; token?: string; identity?: string; user?: string },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      if (!opts.confirm) {
        console.log(`Delete list extension "${opts.name}"? Run with --confirm.`);
        process.exit(1);
      }
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const { listId } = await resolveListId(auth.token!, opts.list, opts.user);
      const r = await deleteTodoListOpenExtension(auth.token!, listId, opts.name, opts.user);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(`\n\u2705 Deleted list extension: ${opts.name}\n`);
    }
  );

todoCommand.addCommand(todoListExtensionCommand);

const todoExtensionCommand = new Command('extension').description('Open type extensions on a task (Graph)');

todoExtensionCommand
  .command('list')
  .description('List open extensions on a task')
  .requiredOption('-l, --list <name|id>', 'List name or ID')
  .requiredOption('-t, --task <id>', 'Task ID')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user or shared mailbox (Graph delegation)')
  .action(
    async (opts: { list: string; task: string; json?: boolean; token?: string; identity?: string; user?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const { listId } = await resolveListId(auth.token!, opts.list, opts.user);
      const r = await listTaskOpenExtensions(auth.token!, listId, opts.task, opts.user);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else {
        for (const ext of r.data) {
          const name = (ext.extensionName as string) || JSON.stringify(ext);
          console.log(`- ${name}`);
        }
      }
    }
  );

todoExtensionCommand
  .command('get')
  .description('Get one open extension by name')
  .requiredOption('-l, --list <name|id>', 'List name or ID')
  .requiredOption('-t, --task <id>', 'Task ID')
  .requiredOption('-n, --name <id>', 'extensionName')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user or shared mailbox (Graph delegation)')
  .action(
    async (opts: {
      list: string;
      task: string;
      name: string;
      json?: boolean;
      token?: string;
      identity?: string;
      user?: string;
    }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const { listId } = await resolveListId(auth.token!, opts.list, opts.user);
      const r = await getTaskOpenExtension(auth.token!, listId, opts.task, opts.name, opts.user);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else console.log(JSON.stringify(r.data, null, 2));
    }
  );

todoExtensionCommand
  .command('set')
  .description('Create an open extension (POST); JSON file is merged with extensionName')
  .requiredOption('-l, --list <name|id>', 'List name or ID')
  .requiredOption('-t, --task <id>', 'Task ID')
  .requiredOption('-n, --name <id>', 'extensionName')
  .requiredOption('--json-file <path>', 'JSON object: custom properties (extensionName added automatically)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user or shared mailbox (Graph delegation)')
  .action(
    async (
      opts: {
        list: string;
        task: string;
        name: string;
        jsonFile: string;
        json?: boolean;
        token?: string;
        identity?: string;
        user?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const { listId } = await resolveListId(auth.token!, opts.list, opts.user);
      const raw = await readFile(opts.jsonFile, 'utf-8');
      const data = JSON.parse(raw) as Record<string, unknown>;
      const r = await setTaskOpenExtension(auth.token!, listId, opts.task, opts.name, data, opts.user);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else console.log(`\n\u2705 Extension set: ${opts.name}\n`);
    }
  );

todoExtensionCommand
  .command('update')
  .description('PATCH an open extension (partial update)')
  .requiredOption('-l, --list <name|id>', 'List name or ID')
  .requiredOption('-t, --task <id>', 'Task ID')
  .requiredOption('-n, --name <id>', 'extensionName')
  .requiredOption('--json-file <path>', 'JSON object: properties to patch')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user or shared mailbox (Graph delegation)')
  .action(
    async (
      opts: {
        list: string;
        task: string;
        name: string;
        jsonFile: string;
        token?: string;
        identity?: string;
        user?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const { listId } = await resolveListId(auth.token!, opts.list, opts.user);
      const raw = await readFile(opts.jsonFile, 'utf-8');
      const patch = JSON.parse(raw) as Record<string, unknown>;
      const r = await updateTaskOpenExtension(auth.token!, listId, opts.task, opts.name, patch, opts.user);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log('\n\u2705 Extension updated.\n');
    }
  );

todoExtensionCommand
  .command('delete')
  .description('Delete an open extension')
  .requiredOption('-l, --list <name|id>', 'List name or ID')
  .requiredOption('-t, --task <id>', 'Task ID')
  .requiredOption('-n, --name <id>', 'extensionName')
  .option('--confirm', 'Confirm without prompt')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user or shared mailbox (Graph delegation)')
  .action(
    async (
      opts: {
        list: string;
        task: string;
        name: string;
        confirm?: boolean;
        token?: string;
        identity?: string;
        user?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      if (!opts.confirm) {
        console.log(`Delete extension "${opts.name}"? Run with --confirm.`);
        process.exit(1);
      }
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const { listId } = await resolveListId(auth.token!, opts.list, opts.user);
      const r = await deleteTaskOpenExtension(auth.token!, listId, opts.task, opts.name, opts.user);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(`\n\u2705 Deleted extension: ${opts.name}\n`);
    }
  );

todoCommand.addCommand(todoExtensionCommand);
