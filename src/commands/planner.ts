import { readFile } from 'node:fs/promises';
import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import {
  applyDeltaPageToState,
  readDeltaStateFile,
  resolveDeltaContinuationUrl,
  writeDeltaStateFile
} from '../lib/graph-delta-state-file.js';
import {
  addPlannerChecklistItem,
  addPlannerFavoritePlan,
  addPlannerReference,
  addPlannerRosterMember,
  archivePlannerPlan,
  buildPlannerAssignments,
  type CreatePlannerTaskExtras,
  createPlannerBucket,
  createPlannerPlan,
  createPlannerPlanForSignedInUser,
  createPlannerPlanInRoster,
  createPlannerRoster,
  createTask,
  deletePlannerBucket,
  deletePlannerPlan,
  deletePlannerPlanDetails,
  deletePlannerTask,
  deletePlannerTaskDetails,
  getAssignedToTaskBoardFormat,
  getBucketTaskBoardFormat,
  getPlanDetails,
  getPlannerBucket,
  getPlannerDeltaPage,
  getPlannerPlan,
  getPlannerPlanUsageRights,
  getPlannerRoster,
  getPlannerTaskDetails,
  getPlannerUser,
  getProgressTaskBoardFormat,
  getTask,
  listFavoritePlans,
  listGroupPlans,
  listPlanBuckets,
  listPlannerMyDayTasks,
  listPlannerPlansForUser,
  listPlannerRecentPlans,
  listPlannerRosterMembers,
  listPlannerTasksForUser,
  listPlanTasks,
  listRosterPlans,
  listUserPlans,
  listUserTasks,
  mergePlannerAssignments,
  movePlannerPlanToContainer,
  normalizeAppliedCategories,
  type PlannerCategorySlot,
  type PlannerPlanDetails,
  type PlannerTask,
  parsePlannerLabelKey,
  removePlannerChecklistItem,
  removePlannerFavoritePlan,
  removePlannerReference,
  removePlannerRosterMember,
  type UpdatePlannerPlanDetailsParams,
  type UpdatePlannerTaskDetailsParams,
  unarchivePlannerPlan,
  updateAssignedToTaskBoardFormat,
  updateBucketTaskBoardFormat,
  updatePlannerBucket,
  updatePlannerChecklistItem,
  updatePlannerPlan,
  updatePlannerPlanDetails,
  updatePlannerTaskDetails,
  updatePlannerUser,
  updateProgressTaskBoardFormat,
  updateTask
} from '../lib/planner-client.js';
import { checkReadOnly } from '../lib/utils.js';

const LABEL_SLOTS: PlannerCategorySlot[] = [
  'category1',
  'category2',
  'category3',
  'category4',
  'category5',
  'category6'
];

function formatTaskLabels(task: PlannerTask, descriptions?: PlannerPlanDetails['categoryDescriptions']): string {
  if (!task.appliedCategories) return '';
  const parts: string[] = [];
  for (const slot of LABEL_SLOTS) {
    if (task.appliedCategories[slot]) {
      const name = descriptions?.[slot]?.trim();
      parts.push(name || slot);
    }
  }
  return parts.join(', ');
}

async function printPlannerTasksHuman(token: string, tasks: PlannerTask[]): Promise<void> {
  const planDetailsCache = new Map<string, PlannerPlanDetails['categoryDescriptions']>();
  for (const t of tasks) {
    if (!planDetailsCache.has(t.planId)) {
      const d = await getPlanDetails(token, t.planId);
      planDetailsCache.set(t.planId, d.ok ? d.data?.categoryDescriptions : undefined);
    }
    const desc = planDetailsCache.get(t.planId);
    const labels = formatTaskLabels(t, desc);
    console.log(`- [${t.percentComplete === 100 ? 'x' : ' '}] ${t.title} (ID: ${t.id})`);
    console.log(`  Plan ID: ${t.planId} | Bucket ID: ${t.bucketId}${labels ? ` | Labels: ${labels}` : ''}`);
  }
}

export const plannerCommand = new Command('planner').description('Manage Microsoft Planner tasks and plans');

async function runPlannerListMyTasks(opts: { json?: boolean; token?: string; identity?: string }): Promise<void> {
  const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
  if (!auth.success) {
    console.error(`Auth error: ${auth.error}`);
    process.exit(1);
  }
  const result = await listUserTasks(auth.token!);
  if (!result.ok || !result.data) {
    console.error(`Error listing tasks: ${result.error?.message}`);
    process.exit(1);
  }
  if (opts.json) {
    console.log(JSON.stringify(result.data, null, 2));
  } else {
    await printPlannerTasksHuman(auth.token!, result.data);
  }
}

plannerCommand
  .command('list-my-tasks')
  .description('List tasks assigned to you')
  .option('--json', 'Output JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(runPlannerListMyTasks);

plannerCommand
  .command('tasks')
  .description('Alias for list-my-tasks (tasks assigned to you)')
  .option('--json', 'Output JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(runPlannerListMyTasks);

plannerCommand
  .command('list-my-day-tasks')
  .description(
    'List tasks in My Day (beta GET …/planner/myDayTasks; tasks you added to My Day or due today per Microsoft)'
  )
  .option('--json', 'Output JSON')
  .option('--user <email>', 'Target user (Graph delegation; may 403)')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (opts: { json?: boolean; user?: string; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const result = await listPlannerMyDayTasks(auth.token!, opts.user);
    if (!result.ok || !result.data) {
      console.error(`Error listing My Day tasks: ${result.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(result.data, null, 2));
    } else {
      await printPlannerTasksHuman(auth.token!, result.data);
    }
  });

plannerCommand
  .command('list-plans')
  .description('List your plans or plans for a group')
  .option('-g, --group <groupId>', 'Group ID to list plans for')
  .option('--json', 'Output JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (opts: { group?: string; json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const result = opts.group ? await listGroupPlans(auth.token!, opts.group) : await listUserPlans(auth.token!);
    if (!result.ok || !result.data) {
      console.error(`Error listing plans: ${result.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(result.data, null, 2));
    } else {
      for (const p of result.data) {
        console.log(`- ${p.title} (ID: ${p.id})`);
      }
    }
  });

plannerCommand
  .command('list-user-tasks')
  .description('List Planner tasks for a user (Graph GET /users/{id}/planner/tasks; may 403 if not permitted)')
  .requiredOption('-u, --user <userId>', 'Azure AD object id of the user')
  .option('--json', 'Output JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (opts: { user: string; json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const result = await listPlannerTasksForUser(auth.token!, opts.user);
    if (!result.ok || !result.data) {
      console.error(`Error listing tasks: ${result.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(result.data, null, 2));
    } else {
      await printPlannerTasksHuman(auth.token!, result.data);
    }
  });

plannerCommand
  .command('list-user-plans')
  .description('List Planner plans for a user (Graph GET /users/{id}/planner/plans; may 403 if not permitted)')
  .requiredOption('-u, --user <userId>', 'Azure AD object id of the user')
  .option('--json', 'Output JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (opts: { user: string; json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const result = await listPlannerPlansForUser(auth.token!, opts.user);
    if (!result.ok || !result.data) {
      console.error(`Error listing plans: ${result.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(result.data, null, 2));
    } else {
      for (const p of result.data) {
        console.log(`- ${p.title} (ID: ${p.id})`);
      }
    }
  });

plannerCommand
  .command('list-buckets')
  .description('List buckets in a plan')
  .requiredOption('-p, --plan <planId>', 'Plan ID')
  .option('--json', 'Output JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (opts: { plan: string; json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const result = await listPlanBuckets(auth.token!, opts.plan);
    if (!result.ok || !result.data) {
      console.error(`Error listing buckets: ${result.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(result.data, null, 2));
    } else {
      for (const b of result.data) {
        console.log(`- ${b.name} (ID: ${b.id})`);
      }
    }
  });

plannerCommand
  .command('list-tasks')
  .description('List tasks in a plan')
  .requiredOption('-p, --plan <planId>', 'Plan ID')
  .option('--json', 'Output JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (opts: { plan: string; json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const result = await listPlanTasks(auth.token!, opts.plan);
    if (!result.ok || !result.data) {
      console.error(`Error listing tasks: ${result.error?.message}`);
      process.exit(1);
    }
    const detailsR = await getPlanDetails(auth.token!, opts.plan);
    const descriptions = detailsR.ok ? detailsR.data?.categoryDescriptions : undefined;
    if (opts.json) {
      console.log(JSON.stringify(result.data, null, 2));
    } else {
      for (const t of result.data) {
        const labels = formatTaskLabels(t, descriptions);
        console.log(
          `- [${t.percentComplete === 100 ? 'x' : ' '}] ${t.title} (ID: ${t.id})${labels ? ` | ${labels}` : ''}`
        );
      }
    }
  });

plannerCommand
  .command('create-task')
  .description('Create a new task in a plan')
  .requiredOption('-p, --plan <planId>', 'Plan ID')
  .requiredOption('-t, --title <title>', 'Task title')
  .option('-b, --bucket <bucketId>', 'Bucket ID')
  .option('--due <ISO-8601>', 'Due date/time (PATCH after create)')
  .option('--start <ISO-8601>', 'Start date/time (PATCH after create)')
  .option(
    '--label <slot>',
    'Label slot: 1-6 or category1..category6 (repeatable; names are defined in plan details)',
    (v: string, prev: string[]) => [...prev, v],
    [] as string[]
  )
  .option(
    '--assign <userId>',
    'Assign user(s) on create (repeatable)',
    (v: string, prev: string[]) => [...prev, v],
    [] as string[]
  )
  .option('--conversation-thread <id>', 'Teams conversation thread id (PATCH after create)')
  .option('--order-hint <hint>', 'Task order hint (PATCH after create)')
  .option('--assignee-priority <hint>', 'Assignee priority order hint (PATCH after create)')
  .option('--priority <0-10>', 'Task priority: 0 highest .. 10 lowest (PATCH after create)')
  .option(
    '--preview-type <mode>',
    'Card preview: automatic | noPreview | checklist | description | reference (PATCH after create)'
  )
  .option('--json', 'Output JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      opts: {
        plan: string;
        title: string;
        bucket?: string;
        due?: string;
        start?: string;
        label?: string[];
        assign?: string[];
        conversationThread?: string;
        orderHint?: string;
        assigneePriority?: string;
        priority?: string;
        previewType?: string;
        json?: boolean;
        token?: string;
        identity?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      if (opts.priority !== undefined) {
        const p = parseInt(opts.priority, 10);
        if (Number.isNaN(p) || p < 0 || p > 10) {
          console.error('Error: --priority must be an integer from 0 to 10');
          process.exit(1);
        }
      }
      let applied: ReturnType<typeof normalizeAppliedCategories> | undefined;
      if (opts.label?.length) {
        const setTrue: PlannerCategorySlot[] = [];
        for (const raw of opts.label) {
          const slot = parsePlannerLabelKey(raw);
          if (!slot) {
            console.error(`Invalid --label "${raw}". Use 1-6 or category1..category6.`);
            process.exit(1);
          }
          setTrue.push(slot);
        }
        applied = normalizeAppliedCategories(undefined, { setTrue });
      }
      const assignments = opts.assign && opts.assign.length > 0 ? buildPlannerAssignments(opts.assign) : undefined;
      const extras: CreatePlannerTaskExtras = {};
      if (opts.due !== undefined) extras.dueDateTime = opts.due;
      if (opts.start !== undefined) extras.startDateTime = opts.start;
      if (opts.conversationThread !== undefined) extras.conversationThreadId = opts.conversationThread;
      if (opts.orderHint !== undefined) extras.orderHint = opts.orderHint;
      if (opts.assigneePriority !== undefined) extras.assigneePriority = opts.assigneePriority;
      if (opts.priority !== undefined) extras.priority = parseInt(opts.priority, 10);
      if (opts.previewType !== undefined) extras.previewType = opts.previewType;
      const extrasOut = Object.keys(extras).length > 0 ? extras : undefined;
      const result = await createTask(auth.token!, opts.plan, opts.title, opts.bucket, assignments, applied, extrasOut);
      if (!result.ok || !result.data) {
        console.error(`Error creating task: ${result.error?.message}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(result.data, null, 2));
      } else {
        console.log(`Created task: ${result.data.title} (ID: ${result.data.id})`);
      }
    }
  );

plannerCommand
  .command('update-task')
  .description('Update a task')
  .requiredOption('-i, --id <taskId>', 'Task ID')
  .option('--title <title>', 'New title')
  .option('-b, --bucket <bucketId>', 'Move to Bucket ID')
  .option('--percent <percentComplete>', 'Percent complete (0-100)')
  .option(
    '--assign <userId>',
    'Replace assignments with exactly these user IDs (repeatable)',
    (v: string, prev: string[]) => [...prev, v],
    [] as string[]
  )
  .option(
    '--add-assign <userId>',
    'Add assignee, keeping others (repeatable)',
    (v: string, prev: string[]) => [...prev, v],
    [] as string[]
  )
  .option(
    '--remove-assign <userId>',
    'Remove one assignee by user ID (repeatable)',
    (v: string, prev: string[]) => [...prev, v],
    [] as string[]
  )
  .option('--clear-assign', 'Remove all assignees')
  .option('--order-hint <hint>', 'Task order hint within bucket')
  .option('--conversation-thread <id>', 'Teams conversation thread id')
  .option('--assignee-priority <hint>', 'Assignee priority order hint')
  .option('--due <ISO-8601>', 'Due date/time')
  .option('--start <ISO-8601>', 'Start date/time')
  .option('--clear-due', 'Clear due date')
  .option('--clear-start', 'Clear start date')
  .option('--priority <0-10>', 'Task priority (0 highest .. 10 lowest)')
  .option('--clear-priority', 'Reset priority (set to null)')
  .option('--preview-type <mode>', 'Card preview: automatic | noPreview | checklist | description | reference')
  .option('--clear-preview-type', 'Clear preview type (set to null)')
  .option(
    '--label <slot>',
    'Turn on label slot (1-6 or category1..category6); repeatable',
    (v: string, prev: string[]) => [...prev, v],
    [] as string[]
  )
  .option(
    '--unlabel <slot>',
    'Turn off label slot; repeatable',
    (v: string, prev: string[]) => [...prev, v],
    [] as string[]
  )
  .option('--clear-labels', 'Clear all label slots on the task')
  .option('--json', 'Output JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      opts: {
        id: string;
        title?: string;
        bucket?: string;
        percent?: string;
        assign?: string[];
        addAssign?: string[];
        removeAssign?: string[];
        clearAssign?: boolean;
        orderHint?: string;
        conversationThread?: string;
        assigneePriority?: string;
        due?: string;
        start?: string;
        clearDue?: boolean;
        clearStart?: boolean;
        priority?: string;
        clearPriority?: boolean;
        previewType?: string;
        clearPreviewType?: boolean;
        label?: string[];
        unlabel?: string[];
        clearLabels?: boolean;
        json?: boolean;
        token?: string;
        identity?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }

      if (opts.clearPriority && opts.priority !== undefined) {
        console.error('Error: use either --priority or --clear-priority, not both');
        process.exit(1);
      }
      if (opts.clearPreviewType && opts.previewType !== undefined) {
        console.error('Error: use either --preview-type or --clear-preview-type, not both');
        process.exit(1);
      }
      if (opts.priority !== undefined) {
        const p = parseInt(opts.priority, 10);
        if (Number.isNaN(p) || p < 0 || p > 10) {
          console.error('Error: --priority must be an integer from 0 to 10');
          process.exit(1);
        }
      }

      const assignReplace = (opts.assign?.length ?? 0) > 0;
      const assignMerge = (opts.addAssign?.length ?? 0) > 0 || (opts.removeAssign?.length ?? 0) > 0;
      if (assignReplace && assignMerge) {
        console.error('Error: use either --assign (replace) or --add-assign/--remove-assign, not both');
        process.exit(1);
      }
      if (assignReplace && opts.clearAssign) {
        console.error('Error: use either --assign or --clear-assign, not both');
        process.exit(1);
      }
      if (opts.clearAssign && assignMerge) {
        console.error('Error: use either --clear-assign or --add-assign/--remove-assign, not both');
        process.exit(1);
      }
      if (opts.clearDue && opts.due !== undefined) {
        console.error('Error: use either --due or --clear-due, not both');
        process.exit(1);
      }
      if (opts.clearStart && opts.start !== undefined) {
        console.error('Error: use either --start or --clear-start, not both');
        process.exit(1);
      }

      // First, we need to get the task to retrieve its ETag.
      const taskRes = await getTask(auth.token!, opts.id);
      if (!taskRes.ok || !taskRes.data) {
        console.error(`Error fetching task: ${taskRes.error?.message}`);
        process.exit(1);
      }
      const etag = taskRes.data['@odata.etag'];
      if (!etag) {
        console.error('Task does not have an ETag');
        process.exit(1);
      }

      const updates: any = {};
      if (opts.title !== undefined) updates.title = opts.title;
      if (opts.bucket !== undefined) updates.bucketId = opts.bucket;
      if (opts.percent !== undefined) {
        const percentValue = parseInt(opts.percent, 10);
        if (Number.isNaN(percentValue) || percentValue < 0 || percentValue > 100) {
          console.error(`Invalid percent value: ${opts.percent}. Must be a number between 0 and 100.`);
          process.exit(1);
        }
        updates.percentComplete = percentValue;
      }
      if (opts.clearAssign) {
        updates.assignments = {};
      } else if (assignReplace) {
        updates.assignments = buildPlannerAssignments(opts.assign!);
      } else if (assignMerge) {
        updates.assignments = mergePlannerAssignments(
          taskRes.data.assignments as Record<string, unknown> | undefined,
          opts.addAssign ?? [],
          opts.removeAssign ?? []
        );
      }

      if (opts.orderHint !== undefined) updates.orderHint = opts.orderHint;
      if (opts.conversationThread !== undefined) updates.conversationThreadId = opts.conversationThread;
      if (opts.assigneePriority !== undefined) updates.assigneePriority = opts.assigneePriority;

      if (opts.clearDue) updates.dueDateTime = null;
      else if (opts.due !== undefined) updates.dueDateTime = opts.due;
      if (opts.clearStart) updates.startDateTime = null;
      else if (opts.start !== undefined) updates.startDateTime = opts.start;

      if (opts.clearPriority) updates.priority = null;
      else if (opts.priority !== undefined) updates.priority = parseInt(opts.priority, 10);
      if (opts.clearPreviewType) updates.previewType = null;
      else if (opts.previewType !== undefined) updates.previewType = opts.previewType;

      const labelOps = (opts.label?.length ?? 0) > 0 || (opts.unlabel?.length ?? 0) > 0 || opts.clearLabels;
      if (labelOps) {
        const setTrue: PlannerCategorySlot[] = [];
        const setFalse: PlannerCategorySlot[] = [];
        for (const raw of opts.label ?? []) {
          const slot = parsePlannerLabelKey(raw);
          if (!slot) {
            console.error(`Invalid --label "${raw}". Use 1-6 or category1..category6.`);
            process.exit(1);
          }
          setTrue.push(slot);
        }
        for (const raw of opts.unlabel ?? []) {
          const slot = parsePlannerLabelKey(raw);
          if (!slot) {
            console.error(`Invalid --unlabel "${raw}". Use 1-6 or category1..category6.`);
            process.exit(1);
          }
          setFalse.push(slot);
        }
        if (opts.clearLabels && (setTrue.length > 0 || setFalse.length > 0)) {
          console.error('Error: use --clear-labels alone, or use --label/--unlabel without --clear-labels');
          process.exit(1);
        }
        updates.appliedCategories = normalizeAppliedCategories(taskRes.data.appliedCategories, {
          clearAll: opts.clearLabels,
          setTrue: setTrue.length ? setTrue : undefined,
          setFalse: setFalse.length ? setFalse : undefined
        });
      }

      if (Object.keys(updates).length === 0) {
        console.error(
          'Error: specify at least one of --title, --bucket, --percent, --assign, --add-assign, --remove-assign, --clear-assign, --order-hint, --conversation-thread, --assignee-priority, --due, --start, --clear-due, --clear-start, --priority, --clear-priority, --preview-type, --clear-preview-type, --label, --unlabel, --clear-labels'
        );
        process.exit(1);
      }

      const result = await updateTask(auth.token!, opts.id, etag, updates);
      if (!result.ok) {
        console.error(`Error updating task: ${result.error?.message}`);
        process.exit(1);
      }

      // Since PATCH returns 204 No Content, get task again to show updated state. The update already
      // succeeded, so a failed display-refetch must NOT be reported as a failure (mirrors update-plan
      // / update-bucket / update-task-details).
      const updatedTaskRes = await getTask(auth.token!, opts.id);
      if (opts.json && updatedTaskRes.ok && updatedTaskRes.data) {
        console.log(JSON.stringify(updatedTaskRes.data, null, 2));
      } else {
        console.log(`Updated task: ${opts.id}`);
      }
    }
  );

plannerCommand
  .command('get-task')
  .description('Fetch a Planner task by ID')
  .requiredOption('-i, --id <taskId>', 'Task ID')
  .option('--with-details', 'Include task details (description, checklist, references)')
  .option('--json', 'Output JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (opts: { id: string; withDetails?: boolean; json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const taskRes = await getTask(auth.token!, opts.id);
    if (!taskRes.ok || !taskRes.data) {
      console.error(`Error: ${taskRes.error?.message}`);
      process.exit(1);
    }
    const t = taskRes.data;
    const td = opts.withDetails ? await getPlannerTaskDetails(auth.token!, opts.id) : undefined;
    if (opts.json) {
      if (opts.withDetails && td?.ok && td.data) {
        console.log(JSON.stringify({ task: t, details: td.data }, null, 2));
      } else {
        console.log(JSON.stringify(t, null, 2));
      }
    } else {
      const detailsR = await getPlanDetails(auth.token!, t.planId);
      const descriptions = detailsR.ok ? detailsR.data?.categoryDescriptions : undefined;
      const labels = formatTaskLabels(t, descriptions);
      console.log(`${t.title} (ID: ${t.id})`);
      console.log(`  Plan: ${t.planId} | Bucket: ${t.bucketId} | ${t.percentComplete}%`);
      if (t.assigneePriority) console.log(`  Assignee priority: ${t.assigneePriority}`);
      if (t.conversationThreadId) console.log(`  Conversation thread: ${t.conversationThreadId}`);
      if (t.dueDateTime) console.log(`  Due: ${t.dueDateTime}`);
      if (t.startDateTime) console.log(`  Start: ${t.startDateTime}`);
      if (t.priority !== undefined) console.log(`  Priority: ${t.priority} (0=highest..10=lowest)`);
      if (t.previewType) console.log(`  Preview type: ${t.previewType}`);
      if (labels) console.log(`  Labels: ${labels}`);
      if (opts.withDetails && td?.ok && td.data) {
        if (td.data.description) console.log(`  Description:\n${td.data.description}`);
        if (td.data.checklist && Object.keys(td.data.checklist).length) {
          console.log('  Checklist:');
          for (const [cid, item] of Object.entries(td.data.checklist)) {
            console.log(`    [${item.isChecked ? 'x' : ' '}] ${item.title} (${cid})`);
          }
        }
      }
    }
  });

plannerCommand
  .command('get-plan')
  .description('Fetch a Planner plan (for ETag before update/delete)')
  .requiredOption('-p, --plan <planId>', 'Plan ID')
  .option('--json', 'Output JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (opts: { plan: string; json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await getPlannerPlan(auth.token!, opts.plan);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else {
      console.log(`${r.data.title} (ID: ${r.data.id})`);
      if (r.data.owner) console.log(`  Owner (group): ${r.data.owner}`);
      if (r.data['@odata.etag']) console.log(`  ETag: ${r.data['@odata.etag']}`);
    }
  });

plannerCommand
  .command('delete-task')
  .description('Delete a Planner task')
  .requiredOption('-i, --id <taskId>', 'Task ID')
  .option('--confirm', 'Confirm deletion')
  .option('--json', 'Output JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (opts: { id: string; confirm?: boolean; json?: boolean; token?: string; identity?: string }, cmd: any) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const taskRes = await getTask(auth.token!, opts.id);
      if (!taskRes.ok || !taskRes.data) {
        console.error(`Error: ${taskRes.error?.message}`);
        process.exit(1);
      }
      const etag = taskRes.data['@odata.etag'];
      if (!etag) {
        console.error('Task missing ETag');
        process.exit(1);
      }
      if (!opts.confirm) {
        console.log(`Delete task "${taskRes.data.title}"? (ID: ${opts.id})`);
        console.log('Run with --confirm to confirm.');
        process.exit(1);
      }
      const del = await deletePlannerTask(auth.token!, opts.id, etag);
      if (!del.ok) {
        console.error(`Error: ${del.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify({ deleted: opts.id }, null, 2));
      else console.log(`Deleted task: ${opts.id}`);
    }
  );

plannerCommand
  .command('create-plan')
  .description(
    'Create a Planner plan: group (v1), roster (beta), or personal / user container (beta: POST /me/planner/plans; use --me)'
  )
  .option('-g, --group <groupId>', 'Microsoft 365 group that owns the plan')
  .option('-r, --roster <rosterId>', 'Beta: planner roster id (container)')
  .option('--me', 'Beta: create in the signed-in user Planner container (POST /me/planner/plans; container type user)')
  .requiredOption('-t, --title <title>', 'Plan title')
  .option('--json', 'Output JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      opts: {
        group?: string;
        roster?: string;
        me?: boolean;
        title: string;
        json?: boolean;
        token?: string;
        identity?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const hasGroup = Boolean(opts.group);
      const hasRoster = Boolean(opts.roster);
      const hasMe = opts.me === true;
      const n = (hasGroup ? 1 : 0) + (hasRoster ? 1 : 0) + (hasMe ? 1 : 0);
      if (n !== 1) {
        console.error('Error: specify exactly one of --group, --roster, or --me');
        process.exit(1);
      }
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = hasMe
        ? await createPlannerPlanForSignedInUser(auth.token!, opts.title)
        : hasRoster
          ? await createPlannerPlanInRoster(auth.token!, opts.roster!, opts.title)
          : await createPlannerPlan(auth.token!, opts.group!, opts.title);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else console.log(`Created plan: ${r.data.title} (ID: ${r.data.id})`);
    }
  );

plannerCommand
  .command('update-plan')
  .description('Rename a Planner plan')
  .requiredOption('-p, --plan <planId>', 'Plan ID')
  .requiredOption('-t, --title <title>', 'New title')
  .option('--json', 'Output JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (opts: { plan: string; title: string; json?: boolean; token?: string; identity?: string }, cmd: any) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const pr = await getPlannerPlan(auth.token!, opts.plan);
      if (!pr.ok || !pr.data) {
        console.error(`Error: ${pr.error?.message}`);
        process.exit(1);
      }
      const etag = pr.data['@odata.etag'];
      if (!etag) {
        console.error('Plan missing ETag');
        process.exit(1);
      }
      const r = await updatePlannerPlan(auth.token!, opts.plan, etag, opts.title);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      const again = await getPlannerPlan(auth.token!, opts.plan);
      if (opts.json && again.ok && again.data) console.log(JSON.stringify(again.data, null, 2));
      else console.log(`Updated plan: ${opts.plan}`);
    }
  );

plannerCommand
  .command('delete-plan')
  .description('Delete a Planner plan')
  .requiredOption('-p, --plan <planId>', 'Plan ID')
  .option('--confirm', 'Confirm deletion')
  .option('--json', 'Output JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (opts: { plan: string; confirm?: boolean; json?: boolean; token?: string; identity?: string }, cmd: any) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const pr = await getPlannerPlan(auth.token!, opts.plan);
      if (!pr.ok || !pr.data) {
        console.error(`Error: ${pr.error?.message}`);
        process.exit(1);
      }
      const etag = pr.data['@odata.etag'];
      if (!etag) {
        console.error('Plan missing ETag');
        process.exit(1);
      }
      if (!opts.confirm) {
        console.log(`Delete plan "${pr.data.title}"? (ID: ${opts.plan})`);
        console.log('Run with --confirm to confirm.');
        process.exit(1);
      }
      const r = await deletePlannerPlan(auth.token!, opts.plan, etag);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify({ deleted: opts.plan }, null, 2));
      else console.log(`Deleted plan: ${opts.plan}`);
    }
  );

plannerCommand
  .command('plan-archive')
  .description(
    'Archive a plan (Graph **beta**: POST /planner/plans/{id}/archive). Requires plan ETag (`If-Match`) and a justification string.'
  )
  .requiredOption('-p, --plan <planId>', 'Plan ID')
  .requiredOption('-j, --justification <text>', 'Reason (required by Graph)')
  .option('--json', 'Output JSON confirmation')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      opts: { plan: string; justification: string; json?: boolean; token?: string; identity?: string },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const pr = await getPlannerPlan(auth.token!, opts.plan);
      if (!pr.ok || !pr.data) {
        console.error(`Error: ${pr.error?.message}`);
        process.exit(1);
      }
      const etag = pr.data['@odata.etag'];
      if (!etag) {
        console.error('Plan missing ETag');
        process.exit(1);
      }
      const r = await archivePlannerPlan(auth.token!, opts.plan, etag, opts.justification);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify({ archived: opts.plan }, null, 2));
      else console.log(`Archived plan: ${opts.plan}`);
    }
  );

plannerCommand
  .command('plan-unarchive')
  .description(
    'Unarchive a plan (Graph **beta**: POST /planner/plans/{id}/unarchive). Requires ETag and justification.'
  )
  .requiredOption('-p, --plan <planId>', 'Plan ID')
  .requiredOption('-j, --justification <text>', 'Reason (required by Graph)')
  .option('--json', 'Output JSON confirmation')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      opts: { plan: string; justification: string; json?: boolean; token?: string; identity?: string },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const pr = await getPlannerPlan(auth.token!, opts.plan);
      if (!pr.ok || !pr.data) {
        console.error(`Error: ${pr.error?.message}`);
        process.exit(1);
      }
      const etag = pr.data['@odata.etag'];
      if (!etag) {
        console.error('Plan missing ETag');
        process.exit(1);
      }
      const r = await unarchivePlannerPlan(auth.token!, opts.plan, etag, opts.justification);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify({ unarchived: opts.plan }, null, 2));
      else console.log(`Unarchived plan: ${opts.plan}`);
    }
  );

plannerCommand
  .command('plan-usage-rights')
  .description(
    'Get usage rights for a plan (Graph **beta**: GET /planner/plans/{id}/getUsageRights()). Returns JSON from Graph.'
  )
  .requiredOption('-p, --plan <planId>', 'Plan ID')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (opts: { plan: string; json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await getPlannerPlanUsageRights(auth.token!, opts.plan);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log(JSON.stringify(r.data, null, 2));
  });

plannerCommand
  .command('move-plan-to-container')
  .description(
    'Move a plan to another container (Graph **beta**: POST …/moveToContainer). Typically from a user/roster container to a group. Body is JSON (e.g. target `container`); requires plan ETag.'
  )
  .requiredOption('-p, --plan <planId>', 'Plan ID')
  .requiredOption('--etag <etag>', 'Plan @odata.etag (If-Match), from planner get-plan --json')
  .requiredOption('--json-file <path>', 'JSON body per Graph (moveToContainer action parameters)')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (opts: { plan: string; etag: string; jsonFile: string; token?: string; identity?: string }, cmd: any) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const raw = await readFile(opts.jsonFile, 'utf-8');
      const body = JSON.parse(raw) as Record<string, unknown>;
      const r = await movePlannerPlanToContainer(auth.token!, opts.plan, opts.etag, body);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (r.data !== undefined && r.data !== null) {
        console.log(JSON.stringify(r.data, null, 2));
      } else {
        console.log('Move completed.');
      }
    }
  );

plannerCommand
  .command('create-bucket')
  .description('Create a bucket in a plan')
  .requiredOption('-p, --plan <planId>', 'Plan ID')
  .requiredOption('-n, --name <name>', 'Bucket name')
  .option('--json', 'Output JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (opts: { plan: string; name: string; json?: boolean; token?: string; identity?: string }, cmd: any) => {
    checkReadOnly(cmd);
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await createPlannerBucket(auth.token!, opts.plan, opts.name);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else console.log(`Created bucket: ${r.data.name} (ID: ${r.data.id})`);
  });

plannerCommand
  .command('update-bucket')
  .description('Rename a bucket and/or set order hint (reordering)')
  .requiredOption('-i, --id <bucketId>', 'Bucket ID')
  .option('-n, --name <name>', 'New name')
  .option('--order-hint <hint>', 'Bucket order hint string')
  .option('--json', 'Output JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      opts: { id: string; name?: string; orderHint?: string; json?: boolean; token?: string; identity?: string },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      if (opts.name === undefined && opts.orderHint === undefined) {
        console.error('Error: specify --name and/or --order-hint');
        process.exit(1);
      }
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const br = await getPlannerBucket(auth.token!, opts.id);
      if (!br.ok || !br.data) {
        console.error(`Error: ${br.error?.message}`);
        process.exit(1);
      }
      const etag = br.data['@odata.etag'];
      if (!etag) {
        console.error('Bucket missing ETag');
        process.exit(1);
      }
      const bucketUpdates: { name?: string; orderHint?: string } = {};
      if (opts.name !== undefined) bucketUpdates.name = opts.name;
      if (opts.orderHint !== undefined) bucketUpdates.orderHint = opts.orderHint;
      const r = await updatePlannerBucket(auth.token!, opts.id, etag, bucketUpdates);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      const again = await getPlannerBucket(auth.token!, opts.id);
      if (opts.json && again.ok && again.data) console.log(JSON.stringify(again.data, null, 2));
      else console.log(`Updated bucket: ${opts.id}`);
    }
  );

plannerCommand
  .command('delete-bucket')
  .description('Delete a bucket')
  .requiredOption('-i, --id <bucketId>', 'Bucket ID')
  .option('--confirm', 'Confirm deletion')
  .option('--json', 'Output JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (opts: { id: string; confirm?: boolean; json?: boolean; token?: string; identity?: string }, cmd: any) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const br = await getPlannerBucket(auth.token!, opts.id);
      if (!br.ok || !br.data) {
        console.error(`Error: ${br.error?.message}`);
        process.exit(1);
      }
      const etag = br.data['@odata.etag'];
      if (!etag) {
        console.error('Bucket missing ETag');
        process.exit(1);
      }
      if (!opts.confirm) {
        console.log(`Delete bucket "${br.data.name}"? (ID: ${opts.id})`);
        console.log('Run with --confirm to confirm.');
        process.exit(1);
      }
      const r = await deletePlannerBucket(auth.token!, opts.id, etag);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify({ deleted: opts.id }, null, 2));
      else console.log(`Deleted bucket: ${opts.id}`);
    }
  );

plannerCommand
  .command('get-task-details')
  .description('Get Planner task details (description, checklist, references)')
  .requiredOption('-i, --id <taskId>', 'Task ID')
  .option('--json', 'Output JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (opts: { id: string; json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await getPlannerTaskDetails(auth.token!, opts.id);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else {
      console.log(`Details ID: ${r.data.id}`);
      if (r.data.description) console.log(`Description:\n${r.data.description}`);
      if (r.data.checklist && Object.keys(r.data.checklist).length) {
        console.log('Checklist:');
        for (const [cid, item] of Object.entries(r.data.checklist)) {
          console.log(`  [${item.isChecked ? 'x' : ' '}] ${item.title} (${cid})`);
        }
      }
    }
  });

plannerCommand
  .command('update-task-details')
  .description('Update Planner task details (description and/or checklist/references JSON)')
  .requiredOption('-i, --id <taskId>', 'Task ID')
  .option('--description <text>', 'Task description (HTML or plain depending on client)')
  .option('--patch-json <path>', 'JSON file with PATCH body (description, checklist, references, previewType)')
  .option('--json', 'Output JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      opts: {
        id: string;
        description?: string;
        patchJson?: string;
        json?: boolean;
        token?: string;
        identity?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      if (opts.description !== undefined && opts.patchJson) {
        console.error('Error: use either --description or --patch-json, not both');
        process.exit(1);
      }
      if (opts.description === undefined && !opts.patchJson) {
        console.error('Error: specify --description and/or --patch-json');
        process.exit(1);
      }
      const dr = await getPlannerTaskDetails(auth.token!, opts.id);
      if (!dr.ok || !dr.data) {
        console.error(`Error: ${dr.error?.message}`);
        process.exit(1);
      }
      const etag = dr.data['@odata.etag'];
      const detailsId = dr.data.id;
      if (!etag) {
        console.error('Task details missing ETag');
        process.exit(1);
      }
      let body: Record<string, unknown>;
      if (opts.patchJson) {
        const raw = await readFile(opts.patchJson, 'utf-8');
        body = JSON.parse(raw) as Record<string, unknown>;
      } else {
        body = { description: opts.description };
      }
      const r = await updatePlannerTaskDetails(auth.token!, detailsId, etag, body as UpdatePlannerTaskDetailsParams);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      const again = await getPlannerTaskDetails(auth.token!, opts.id);
      if (opts.json && again.ok && again.data) console.log(JSON.stringify(again.data, null, 2));
      else console.log(`Updated task details for task: ${opts.id}`);
    }
  );

plannerCommand
  .command('get-plan-details')
  .description('Get plan details (label names, sharedWith, ETag)')
  .requiredOption('-p, --plan <planId>', 'Plan ID')
  .option('--json', 'Output JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (opts: { plan: string; json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await getPlanDetails(auth.token!, opts.plan);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else {
      console.log(`Plan details ID: ${r.data.id}`);
      if (r.data['@odata.etag']) console.log(`ETag: ${r.data['@odata.etag']}`);
      if (r.data.categoryDescriptions) {
        for (const slot of LABEL_SLOTS) {
          const n = r.data.categoryDescriptions[slot];
          if (n) console.log(`  ${slot}: ${n}`);
        }
      }
      if (r.data.sharedWith && Object.keys(r.data.sharedWith).length) {
        console.log('sharedWith:', JSON.stringify(r.data.sharedWith));
      }
    }
  });

plannerCommand
  .command('update-plan-details')
  .description('PATCH plan details (label names, sharedWith)')
  .requiredOption('-p, --plan <planId>', 'Plan ID')
  .option('--names-json <path>', 'JSON: categoryDescriptions object (category1..category6)')
  .option('--shared-with-json <path>', 'JSON: sharedWith map { userId: true|false }')
  .option('--json', 'Output JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      opts: {
        plan: string;
        namesJson?: string;
        sharedWithJson?: string;
        json?: boolean;
        token?: string;
        identity?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      if (!opts.namesJson && !opts.sharedWithJson) {
        console.error('Error: specify --names-json and/or --shared-with-json');
        process.exit(1);
      }
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const dr = await getPlanDetails(auth.token!, opts.plan);
      if (!dr.ok || !dr.data) {
        console.error(`Error: ${dr.error?.message}`);
        process.exit(1);
      }
      const etag = dr.data['@odata.etag'];
      if (!etag) {
        console.error('Plan details missing ETag');
        process.exit(1);
      }
      const body: UpdatePlannerPlanDetailsParams = {};
      if (opts.namesJson) {
        const raw = await readFile(opts.namesJson, 'utf-8');
        body.categoryDescriptions = JSON.parse(raw) as UpdatePlannerPlanDetailsParams['categoryDescriptions'];
      }
      if (opts.sharedWithJson) {
        const raw = await readFile(opts.sharedWithJson, 'utf-8');
        body.sharedWith = JSON.parse(raw) as Record<string, boolean>;
      }
      const r = await updatePlannerPlanDetails(auth.token!, opts.plan, etag, body);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      const again = await getPlanDetails(auth.token!, opts.plan);
      if (opts.json && again.ok && again.data) console.log(JSON.stringify(again.data, null, 2));
      else console.log(`Updated plan details for plan: ${opts.plan}`);
    }
  );

plannerCommand
  .command('delete-plan-details')
  .description(
    'DELETE plan details facet (`/planner/plans/{id}/details`) — destructive; labels/sharedWith data is removed'
  )
  .requiredOption('-p, --plan <planId>', 'Plan ID')
  .option('--confirm', 'Confirm deletion')
  .option('--json', 'Output JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (opts: { plan: string; confirm?: boolean; json?: boolean; token?: string; identity?: string }, cmd: any) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const dr = await getPlanDetails(auth.token!, opts.plan);
      if (!dr.ok || !dr.data) {
        console.error(`Error: ${dr.error?.message}`);
        process.exit(1);
      }
      const etag = dr.data['@odata.etag'];
      if (!etag) {
        console.error('Plan details missing ETag');
        process.exit(1);
      }
      if (!opts.confirm) {
        console.log(`Delete plan DETAILS for plan ${opts.plan}? (Plan object may remain; this removes details only.)`);
        console.log('Run with --confirm to confirm.');
        process.exit(1);
      }
      const del = await deletePlannerPlanDetails(auth.token!, opts.plan, etag);
      if (!del.ok) {
        console.error(`Error: ${del.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify({ deletedPlanDetails: opts.plan }, null, 2));
      else console.log(`Deleted plan details for plan: ${opts.plan}`);
    }
  );

plannerCommand
  .command('delete-task-details')
  .description('DELETE task details facet (`/planner/tasks/{id}/details`) — checklist/description/references removed')
  .requiredOption('-i, --id <taskId>', 'Task ID')
  .option('--confirm', 'Confirm deletion')
  .option('--json', 'Output JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (opts: { id: string; confirm?: boolean; json?: boolean; token?: string; identity?: string }, cmd: any) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const dr = await getPlannerTaskDetails(auth.token!, opts.id);
      if (!dr.ok || !dr.data) {
        console.error(`Error: ${dr.error?.message}`);
        process.exit(1);
      }
      const etag = dr.data['@odata.etag'];
      if (!etag) {
        console.error('Task details missing ETag');
        process.exit(1);
      }
      if (!opts.confirm) {
        console.log(`Delete TASK DETAILS for task ${opts.id}? (Task row may remain.)`);
        console.log('Run with --confirm to confirm.');
        process.exit(1);
      }
      const del = await deletePlannerTaskDetails(auth.token!, opts.id, etag);
      if (!del.ok) {
        console.error(`Error: ${del.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify({ deletedTaskDetails: opts.id }, null, 2));
      else console.log(`Deleted task details for task: ${opts.id}`);
    }
  );

plannerCommand
  .command('list-favorite-plans')
  .description('List favorite plans (beta Graph API; see GRAPH_BETA_URL)')
  .option('--json', 'Output JSON')
  .option('--user <email>', 'Target user (Graph delegation; may 403)')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (opts: { json?: boolean; user?: string; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await listFavoritePlans(auth.token!, opts.user);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else for (const p of r.data) console.log(`- ${p.title} (${p.id})`);
  });

plannerCommand
  .command('list-roster-plans')
  .description('List plans from rosters you belong to (beta Graph API)')
  .option('--json', 'Output JSON')
  .option('--user <email>', 'Target user (Graph delegation; may 403)')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (opts: { json?: boolean; user?: string; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await listRosterPlans(auth.token!, opts.user);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else for (const p of r.data) console.log(`- ${p.title} (${p.id})`);
  });

plannerCommand
  .command('list-recent-plans')
  .description('List recently viewed plans (beta GET …/planner/recentPlans)')
  .option('--json', 'Output JSON')
  .option('--user <email>', 'Target user (Graph delegation; may 403)')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (opts: { json?: boolean; user?: string; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await listPlannerRecentPlans(auth.token!, opts.user);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else for (const p of r.data) console.log(`- ${p.title} (${p.id})`);
  });

plannerCommand
  .command('delta')
  .description('Fetch one page of Planner delta (beta /me/planner/all/delta or --url / --state-file cursor)')
  .option('--url <url>', 'Next or delta link (overrides --state-file continuation)')
  .option('--state-file <path>', 'Read/write JSON delta cursor (kind: plannerAll)')
  .option('--json', 'Output JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (opts: { url?: string; stateFile?: string; json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const existingState = opts.stateFile ? await readDeltaStateFile(opts.stateFile) : null;
    if (existingState && existingState.kind !== 'plannerAll') {
      console.error('Error: state file is not for planner delta (kind must be plannerAll).');
      process.exit(1);
    }
    const continueUrl = resolveDeltaContinuationUrl({ explicitNext: opts.url, state: existingState });
    const r = await getPlannerDeltaPage(auth.token!, continueUrl);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.stateFile && r.data) {
      const merged = applyDeltaPageToState(existingState, 'plannerAll', r.data, {});
      await writeDeltaStateFile(opts.stateFile, merged);
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else {
      console.log(`Changes: ${(r.data.value ?? []).length} item(s)`);
      if (r.data['@odata.nextLink']) console.log(`nextLink: ${r.data['@odata.nextLink']}`);
      if (r.data['@odata.deltaLink']) console.log(`deltaLink: ${r.data['@odata.deltaLink']}`);
      if (opts.stateFile) console.log(`state-file: ${opts.stateFile} (updated)`);
    }
  });

plannerCommand
  .command('add-checklist-item')
  .description('Add a Planner checklist item (generates id)')
  .requiredOption('-i, --id <taskId>', 'Task ID')
  .requiredOption('-t, --title <text>', 'Checklist item text')
  .option('-c, --item-id <id>', 'Optional id (default: random UUID)')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (opts: { id: string; title: string; itemId?: string; token?: string; identity?: string }, cmd: any) => {
    checkReadOnly(cmd);
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await addPlannerChecklistItem(auth.token!, opts.id, opts.title, opts.itemId);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log(`OK: checklist updated for task ${opts.id}`);
  });

plannerCommand
  .command('remove-checklist-item')
  .description('Remove a Planner checklist item by id')
  .requiredOption('-i, --id <taskId>', 'Task ID')
  .requiredOption('-c, --item <checklistItemId>', 'Checklist item id')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (opts: { id: string; item: string; token?: string; identity?: string }, cmd: any) => {
    checkReadOnly(cmd);
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await removePlannerChecklistItem(auth.token!, opts.id, opts.item);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log(`OK: removed checklist item ${opts.item}`);
  });

plannerCommand
  .command('add-reference')
  .description('Add a link reference on task details')
  .requiredOption('-i, --id <taskId>', 'Task ID')
  .requiredOption('-u, --url <url>', 'Reference URL (key)')
  .requiredOption('-a, --alias <text>', 'Display alias')
  .option('--type <type>', 'Optional type string (e.g. PowerPoint)')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      opts: { id: string; url: string; alias: string; type?: string; token?: string; identity?: string },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await addPlannerReference(auth.token!, opts.id, opts.url, opts.alias, opts.type);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(`OK: reference added for task ${opts.id}`);
    }
  );

plannerCommand
  .command('remove-reference')
  .description('Remove a reference by URL key')
  .requiredOption('-i, --id <taskId>', 'Task ID')
  .requiredOption('-u, --url <url>', 'Reference URL key')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (opts: { id: string; url: string; token?: string; identity?: string }, cmd: any) => {
    checkReadOnly(cmd);
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await removePlannerReference(auth.token!, opts.id, opts.url);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log(`OK: reference removed for task ${opts.id}`);
  });

plannerCommand
  .command('update-checklist-item')
  .description('Rename or check/uncheck a Planner checklist item')
  .requiredOption('-i, --id <taskId>', 'Task ID')
  .requiredOption('-c, --item <checklistItemId>', 'Checklist item id')
  .option('-t, --title <text>', 'New title')
  .option('--checked', 'Mark checked')
  .option('--unchecked', 'Mark unchecked')
  .option('--order-hint <hint>', 'Order hint string')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      opts: {
        id: string;
        item: string;
        title?: string;
        checked?: boolean;
        unchecked?: boolean;
        orderHint?: string;
        token?: string;
        identity?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      if (opts.checked && opts.unchecked) {
        console.error('Error: use either --checked or --unchecked, not both');
        process.exit(1);
      }
      if (opts.title === undefined && !opts.checked && !opts.unchecked && opts.orderHint === undefined) {
        console.error('Error: specify --title, --checked/--unchecked, and/or --order-hint');
        process.exit(1);
      }
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const patch: { title?: string; isChecked?: boolean; orderHint?: string } = {};
      if (opts.title !== undefined) patch.title = opts.title;
      if (opts.checked) patch.isChecked = true;
      if (opts.unchecked) patch.isChecked = false;
      if (opts.orderHint !== undefined) patch.orderHint = opts.orderHint;
      const r = await updatePlannerChecklistItem(auth.token!, opts.id, opts.item, patch);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(`OK: updated checklist item ${opts.item}`);
    }
  );

plannerCommand
  .command('get-task-board')
  .description('Get task board ordering (assignedTo, bucket, or progress view)')
  .requiredOption('-i, --id <taskId>', 'Task ID')
  .requiredOption('--view <name>', 'assignedTo | bucket | progress (matches Graph task board format resources)')
  .option('--json', 'Output JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (opts: { id: string; view: string; json?: boolean; token?: string; identity?: string }) => {
    const v = opts.view.trim().toLowerCase();
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r =
      v === 'assignedto' || v === 'assigned'
        ? await getAssignedToTaskBoardFormat(auth.token!, opts.id)
        : v === 'bucket'
          ? await getBucketTaskBoardFormat(auth.token!, opts.id)
          : v === 'progress'
            ? await getProgressTaskBoardFormat(auth.token!, opts.id)
            : null;
    if (!r) {
      console.error('Error: --view must be assignedTo, bucket, or progress');
      process.exit(1);
    }
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else console.log(JSON.stringify(r.data, null, 2));
  });

plannerCommand
  .command('update-task-board')
  .description('PATCH task board ordering (use --json-file for body; etag fetched automatically)')
  .requiredOption('-i, --id <taskId>', 'Task ID')
  .requiredOption('--view <name>', 'assignedTo | bucket | progress (matches Graph task board format resources)')
  .requiredOption(
    '--json-file <path>',
    'JSON body: assignedTo = orderHintsByAssignee + unassignedOrderHint; bucket/progress = { "orderHint": "..." }'
  )
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (opts: { id: string; view: string; jsonFile: string; token?: string; identity?: string }, cmd: any) => {
    checkReadOnly(cmd);
    const v = opts.view.trim().toLowerCase();
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const raw = await readFile(opts.jsonFile, 'utf-8');
    const body = JSON.parse(raw) as Record<string, unknown>;
    if (v === 'assignedto' || v === 'assigned') {
      const gr = await getAssignedToTaskBoardFormat(auth.token!, opts.id);
      if (!gr.ok || !gr.data) {
        console.error(`Error: ${gr.error?.message}`);
        process.exit(1);
      }
      const etag = gr.data['@odata.etag'];
      if (!etag) {
        console.error('Missing ETag on assignedTo task board format');
        process.exit(1);
      }
      const patch: {
        orderHintsByAssignee?: Record<string, string> | null;
        unassignedOrderHint?: string | null;
      } = {};
      if (Object.hasOwn(body, 'orderHintsByAssignee')) {
        patch.orderHintsByAssignee = body.orderHintsByAssignee as Record<string, string> | null;
      }
      if (Object.hasOwn(body, 'unassignedOrderHint')) {
        patch.unassignedOrderHint = body.unassignedOrderHint as string | null;
      }
      if (Object.keys(patch).length === 0) {
        console.error('Error: json-file must include orderHintsByAssignee and/or unassignedOrderHint');
        process.exit(1);
      }
      const r = await updateAssignedToTaskBoardFormat(auth.token!, opts.id, etag, patch);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
    } else if (v === 'bucket') {
      const gr = await getBucketTaskBoardFormat(auth.token!, opts.id);
      if (!gr.ok || !gr.data) {
        console.error(`Error: ${gr.error?.message}`);
        process.exit(1);
      }
      const etag = gr.data['@odata.etag'];
      const orderHint = typeof body.orderHint === 'string' ? body.orderHint : null;
      if (!etag || !orderHint) {
        console.error('Error: bucket view requires ETag and json-file with { "orderHint": "..." }');
        process.exit(1);
      }
      const r = await updateBucketTaskBoardFormat(auth.token!, opts.id, etag, orderHint);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
    } else if (v === 'progress') {
      const gr = await getProgressTaskBoardFormat(auth.token!, opts.id);
      if (!gr.ok || !gr.data) {
        console.error(`Error: ${gr.error?.message}`);
        process.exit(1);
      }
      const etag = gr.data['@odata.etag'];
      const orderHint = typeof body.orderHint === 'string' ? body.orderHint : null;
      if (!etag || !orderHint) {
        console.error('Error: progress view requires ETag and json-file with { "orderHint": "..." }');
        process.exit(1);
      }
      const r = await updateProgressTaskBoardFormat(auth.token!, opts.id, etag, orderHint);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
    } else {
      console.error('Error: --view must be assignedTo, bucket, or progress');
      process.exit(1);
    }
    console.log('OK: task board updated');
  });

plannerCommand
  .command('get-me')
  .description('Get current user Planner settings (beta: favorites, recents; see GRAPH_BETA_URL)')
  .option('--json', 'Output JSON')
  .option('--user <email>', 'Target user (Graph delegation; may 403)')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (opts: { json?: boolean; user?: string; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await getPlannerUser(auth.token!, opts.user);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else console.log(JSON.stringify(r.data, null, 2));
  });

plannerCommand
  .command('update-me')
  .description(
    'PATCH plannerUser with a merge body (beta; use @odata.etag from planner get-me --json as --etag If-Match)'
  )
  .requiredOption('--etag <etag>', 'If-Match value from planner get-me')
  .requiredOption('--json-file <path>', 'JSON merge body (e.g. recentPlanReferences per Graph docs)')
  .option('--user <email>', 'Target user (Graph delegation; may 403)')
  .option('--json', 'Print updated plannerUser as JSON when Graph returns representation')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      opts: {
        etag: string;
        jsonFile: string;
        user?: string;
        json?: boolean;
        token?: string;
        identity?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      let body: Record<string, unknown>;
      try {
        body = JSON.parse(await readFile(opts.jsonFile, 'utf-8')) as Record<string, unknown>;
      } catch (e) {
        console.error(`Error reading json-file: ${e instanceof Error ? e.message : String(e)}`);
        process.exit(1);
      }
      const r = await updatePlannerUser(auth.token!, opts.user, opts.etag, body);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json && r.data) console.log(JSON.stringify(r.data, null, 2));
      else console.log('OK: planner user updated');
    }
  );

plannerCommand
  .command('add-favorite')
  .description('Add a plan to your favorites (beta PATCH /me/planner)')
  .requiredOption('-p, --plan <planId>', 'Plan ID')
  .requiredOption('-t, --title <text>', 'Plan title (shown in favorites)')
  .option('--user <email>', 'Target user (Graph delegation; may 403)')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (opts: { plan: string; title: string; user?: string; token?: string; identity?: string }, cmd: any) => {
    checkReadOnly(cmd);
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await addPlannerFavoritePlan(auth.token!, opts.plan, opts.title, opts.user);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log(`OK: favorite added for plan ${opts.plan}`);
  });

plannerCommand
  .command('remove-favorite')
  .description('Remove a plan from your favorites (beta)')
  .requiredOption('-p, --plan <planId>', 'Plan ID')
  .option('--user <email>', 'Target user (Graph delegation; may 403)')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (opts: { plan: string; user?: string; token?: string; identity?: string }, cmd: any) => {
    checkReadOnly(cmd);
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await removePlannerFavoritePlan(auth.token!, opts.plan, opts.user);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log(`OK: favorite removed for plan ${opts.plan}`);
  });

const plannerRosterCommand = new Command('roster').description(
  'Planner roster APIs (beta; rosters are alternate plan containers — see planner create-plan --roster)'
);

plannerRosterCommand
  .command('create')
  .description('Create an empty planner roster (beta POST /planner/rosters)')
  .option('--json', 'Output JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (opts: { json?: boolean; token?: string; identity?: string }, cmd: any) => {
    checkReadOnly(cmd);
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await createPlannerRoster(auth.token!);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else console.log(`Created roster (ID: ${r.data.id})`);
  });

plannerRosterCommand
  .command('get')
  .description('Get a planner roster by id (beta)')
  .requiredOption('-r, --roster <rosterId>', 'Roster ID')
  .option('--json', 'Output JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (opts: { roster: string; json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await getPlannerRoster(auth.token!, opts.roster);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else console.log(JSON.stringify(r.data, null, 2));
  });

plannerRosterCommand
  .command('list-members')
  .description('List members of a planner roster (beta)')
  .requiredOption('-r, --roster <rosterId>', 'Roster ID')
  .option('--json', 'Output JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (opts: { roster: string; json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await listPlannerRosterMembers(auth.token!, opts.roster);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else for (const m of r.data) console.log(`- user ${m.userId} (member id: ${m.id})`);
  });

plannerRosterCommand
  .command('add-member')
  .description('Add a user to a planner roster (beta)')
  .requiredOption('-r, --roster <rosterId>', 'Roster ID')
  .requiredOption('-u, --user <userId>', 'Azure AD object id of the user')
  .option('--tenant <tenantId>', 'Tenant id (optional; same-tenant only per Graph)')
  .option('--json', 'Output JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      opts: { roster: string; user: string; tenant?: string; json?: boolean; token?: string; identity?: string },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await addPlannerRosterMember(auth.token!, opts.roster, opts.user, {
        tenantId: opts.tenant
      });
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else console.log(`Added member: ${r.data.userId} (member id: ${r.data.id})`);
    }
  );

plannerRosterCommand
  .command('remove-member')
  .description(
    'Remove a member from a planner roster (beta; removing last member may delete roster/plan after retention)'
  )
  .requiredOption('-r, --roster <rosterId>', 'Roster ID')
  .requiredOption('-m, --member <memberId>', 'Roster member resource id (from list-members)')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (opts: { roster: string; member: string; token?: string; identity?: string }, cmd: any) => {
    checkReadOnly(cmd);
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await removePlannerRosterMember(auth.token!, opts.roster, opts.member);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log(`OK: removed roster member ${opts.member}`);
  });

plannerCommand.addCommand(plannerRosterCommand);
