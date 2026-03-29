import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import {
  createTask,
  getTask,
  listGroupPlans,
  listPlanBuckets,
  listPlanTasks,
  listUserPlans,
  listUserTasks,
  updateTask
} from '../lib/planner-client.js';

export const plannerCommand = new Command('planner').description('Manage Microsoft Planner tasks and plans');

plannerCommand
  .command('list-my-tasks')
  .description('List tasks assigned to you')
  .option('--json', 'Output JSON')
  .option('--token <token>', 'Use a specific token')
  .action(async (opts: { json?: boolean; token?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token });
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
      for (const t of result.data) {
        console.log(`- [${t.percentComplete === 100 ? 'x' : ' '}] ${t.title} (ID: ${t.id})`);
        console.log(`  Plan ID: ${t.planId} | Bucket ID: ${t.bucketId}`);
      }
    }
  });

plannerCommand
  .command('list-plans')
  .description('List your plans or plans for a group')
  .option('-g, --group <groupId>', 'Group ID to list plans for')
  .option('--json', 'Output JSON')
  .option('--token <token>', 'Use a specific token')
  .action(async (opts: { group?: string; json?: boolean; token?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token });
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
  .command('list-buckets')
  .description('List buckets in a plan')
  .requiredOption('-p, --plan <planId>', 'Plan ID')
  .option('--json', 'Output JSON')
  .option('--token <token>', 'Use a specific token')
  .action(async (opts: { plan: string; json?: boolean; token?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token });
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
  .action(async (opts: { plan: string; json?: boolean; token?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const result = await listPlanTasks(auth.token!, opts.plan);
    if (!result.ok || !result.data) {
      console.error(`Error listing tasks: ${result.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(result.data, null, 2));
    } else {
      for (const t of result.data) {
        console.log(`- [${t.percentComplete === 100 ? 'x' : ' '}] ${t.title} (ID: ${t.id})`);
      }
    }
  });

plannerCommand
  .command('create-task')
  .description('Create a new task in a plan')
  .requiredOption('-p, --plan <planId>', 'Plan ID')
  .requiredOption('-t, --title <title>', 'Task title')
  .option('-b, --bucket <bucketId>', 'Bucket ID')
  .option('--json', 'Output JSON')
  .option('--token <token>', 'Use a specific token')
  .action(async (opts: { plan: string; title: string; bucket?: string; json?: boolean; token?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const result = await createTask(auth.token!, opts.plan, opts.title, opts.bucket);
    if (!result.ok || !result.data) {
      console.error(`Error creating task: ${result.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(result.data, null, 2));
    } else {
      console.log(`Created task: ${result.data.title} (ID: ${result.data.id})`);
    }
  });

plannerCommand
  .command('update-task')
  .description('Update a task')
  .requiredOption('-i, --id <taskId>', 'Task ID')
  .option('--title <title>', 'New title')
  .option('-b, --bucket <bucketId>', 'Move to Bucket ID')
  .option('--percent <percentComplete>', 'Percent complete (0-100)')
  .option('--assign <userId>', 'Assign to user ID')
  .option('--json', 'Output JSON')
  .option('--token <token>', 'Use a specific token')
  .action(
    async (opts: {
      id: string;
      title?: string;
      bucket?: string;
      percent?: string;
      assign?: string;
      json?: boolean;
      token?: string;
    }) => {
      const auth = await resolveGraphAuth({ token: opts.token });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
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
      if (opts.assign) {
        // Planner API requires a specific structure for assignments.
        updates.assignments = {
          [opts.assign]: {
            '@odata.type': '#microsoft.graph.plannerAssignment',
            orderHint: ' !'
          }
        };
      }

      const result = await updateTask(auth.token!, opts.id, etag, updates);
      if (!result.ok) {
        console.error(`Error updating task: ${result.error?.message}`);
        process.exit(1);
      }

      // Since PATCH returns 204 No Content, get task again to show updated state
      const updatedTaskRes = await getTask(auth.token!, opts.id);
      if (!updatedTaskRes.ok || !updatedTaskRes.data) {
        console.error(`Error fetching updated task: ${updatedTaskRes.error?.message}`);
        process.exit(1);
      }

      if (opts.json) {
        console.log(JSON.stringify(updatedTaskRes.data, null, 2));
      } else {
        console.log(`Updated task: ${opts.id}`);
      }
    }
  );
