import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import { resolveAuth } from '../lib/auth.js';
import { getEmail } from '../lib/ews-client.js';
import {
  getTodoLists,
  getTodoList,
  getTasks,
  getTask,
  createTask,
  updateTask,
  deleteTask,
  addChecklistItem,
  type TodoImportance,
  type TodoStatus,
  type TodoTask,
  type TodoList
} from '../lib/todo-client.js';

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

async function resolveListId(token: string, nameOrId: string): Promise<{ listId: string; listDisplay: string }> {
  const listsR = await getTodoLists(token);
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

  const s = await getTodoList(token, nameOrId);
  if (!s.ok || !s.data) {
    console.error(`List not found: "${nameOrId}".`);
    console.error('Use "clippy todo lists".');
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
  .action(async (opts: { json?: boolean; token?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const result = await getTodoLists(auth.token!);
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
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .action(
    async (opts: {
      list?: string;
      task?: string;
      status?: string;
      importance?: string;
      json?: boolean;
      token?: string;
    }) => {
      const auth = await resolveGraphAuth({ token: opts.token });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }

      const listName = opts.list || 'Tasks';
      const { listId, listDisplay } = await resolveListId(auth.token!, listName);

      if (opts.task) {
        const r = await getTask(auth.token!, listId, opts.task);
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
        if (t.dueDateTime) console.log(`Due:         ${fmtDT(t.dueDateTime)} (${t.dueDateTime.timeZone})`);
        if (t.isReminderOn && t.reminderDateTime) console.log(`Reminder:    ${fmtDT(t.reminderDateTime)}`);
        if (t.completedDateTime) console.log(`Completed:   ${fmtDT(t.completedDateTime)}`);
        if (t.linkedResources?.length) {
          console.log('Linked:');
          for (const lr of t.linkedResources) console.log(`  - ${lr.description}: ${lr.webUrl}`);
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

      const filters: string[] = [];
      if (opts.status) filters.push(`status eq '${opts.status}'`);
      if (opts.importance) filters.push(`importance eq '${opts.importance}'`);
      const result = await getTasks(auth.token!, listId, filters.join(' and ') || undefined);
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
        if (t.linkedResources?.length)
          console.log(`      \u21B3 linked: ${t.linkedResources.map((l) => l.description).join(', ')}`);
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
  .option('--importance <level>', 'Importance: low, normal, high', 'normal')
  .option('--status <status>', 'Initial status: notStarted, inProgress, waitingOnOthers, deferred', 'notStarted')
  .option('--reminder <ISO-8601>', 'Reminder datetime')
  .option('--link <msgId>', 'Link task to an email by message ID')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .action(
    async (opts: {
      title: string;
      list?: string;
      body?: string;
      due?: string;
      importance?: string;
      status?: string;
      reminder?: string;
      link?: string;
      json?: boolean;
      token?: string;
    }) => {
      const auth = await resolveGraphAuth({ token: opts.token });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }

      const listName = opts.list || 'Tasks';
      const { listId } = await resolveListId(auth.token!, listName);

      let linkedResources;
      if (opts.link) {
        const ewsAuth = await resolveAuth({ token: opts.token });
        if (!ewsAuth.success) {
          console.error(`EWS Auth error: ${ewsAuth.error}`);
          process.exit(1);
        }
        const er = await getEmail(ewsAuth.token!, opts.link);
        if (!er.ok || !er.data) {
          console.error(`Could not fetch email: ${er.error?.message}`);
          process.exit(1);
        }
        linkedResources = [{ webUrl: emailUrl(er.data.Id), description: er.data.Subject || 'Linked email' }];
      }

      const result = await createTask(auth.token!, listId, {
        title: opts.title,
        body: opts.body,
        importance: opts.importance as TodoImportance,
        status: opts.status as TodoStatus,
        dueDateTime: opts.due,
        reminderDateTime: opts.reminder,
        isReminderOn: !!opts.reminder,
        linkedResources
      });
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
  .command('complete')
  .description('Mark a task as completed')
  .requiredOption('-l, --list <name|id>', 'List name or ID')
  .requiredOption('-t, --task <id>', 'Task ID')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .action(async (opts: { list: string; task: string; json?: boolean; token?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const { listId } = await resolveListId(auth.token!, opts.list);
    const now = new Date().toISOString();
    const r = await updateTask(auth.token!, listId, opts.task, { status: 'completed', completedDateTime: now });
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else console.log(`\n\u2705 Completed: "${r.data.title}" (${fmtDate(now)})\n`);
  });

todoCommand
  .command('delete')
  .description('Delete a task')
  .requiredOption('-l, --list <name|id>', 'List name or ID')
  .requiredOption('-t, --task <id>', 'Task ID')
  .option('--confirm', 'Skip confirmation prompt')
  .option('--token <token>', 'Use a specific token')
  .action(async (opts: { list: string; task: string; confirm?: boolean; token?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const { listId, listDisplay: listName } = await resolveListId(auth.token!, opts.list);
    const taskR = await getTask(auth.token!, listId, opts.task);
    if (!taskR.ok || !taskR.data) {
      console.error(`Task not found: ${taskR.error?.message}`);
      process.exit(1);
    }
    if (!opts.confirm) {
      console.log(`Delete "${taskR.data.title}" from "${listName}"? (ID: ${opts.task})`);
      console.log('Run with --confirm to confirm.');
      process.exit(1);
    }
    const r = await deleteTask(auth.token!, listId, opts.task);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log(`\n\u{1F5D1}  Deleted: "${taskR.data.title}"\n`);
  });

todoCommand
  .command('add-checklist')
  .description('Add a checklist (subtask) item to a task')
  .requiredOption('-l, --list <name|id>', 'List name or ID')
  .requiredOption('-t, --task <id>', 'Task ID')
  .requiredOption('-n, --name <text>', 'Checklist item text')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .action(async (opts: { list: string; task: string; name: string; json?: boolean; token?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const { listId } = await resolveListId(auth.token!, opts.list);
    const r = await addChecklistItem(auth.token!, listId, opts.task, opts.name);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else console.log(`\n\u2705 Added: "${r.data.displayName}" (${r.data.id})\n`);
  });
