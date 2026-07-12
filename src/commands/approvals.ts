import { Command } from 'commander';
import {
  type Approval,
  type ApprovalStep,
  deleteApproval,
  getApproval,
  listAllMyApprovals,
  listApprovalSteps,
  listMyApprovals,
  patchApprovalStep
} from '../lib/graph-approvals-client.js';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import { checkReadOnly } from '../lib/utils.js';

export const approvalsCommand = new Command('approvals').description(
  'Microsoft Approvals (Teams Approvals + Power Automate approvals) — list, get, steps, respond, cancel (DELETE). All paths use `/me/approvals` on the **beta** Graph endpoint. Requires the **`ApprovalSolution.ReadWrite`** delegated scope (identifier `6768d3af-4562-48ff-82d2-c5e19eb21b9c`); **`ApprovalSolutionResponse.ReadWrite`** is a narrower alternative for read-and-respond only.'
);

interface BaseOpts {
  json?: boolean;
  token?: string;
  identity?: string;
}

const baseFlags = (cmd: Command) =>
  cmd
    .option('--json', 'Output raw Graph JSON')
    .option('--token <token>', 'Use a specific Graph token')
    .option('--identity <name>', 'Graph token cache identity (default: default)');

baseFlags(approvalsCommand.command('list'))
  .description(
    'List approvals visible to the signed-in user (`GET /me/approvals?$expand=steps`). Use --next or --all for paging.'
  )
  .option('--top <n>', 'Limit results (Graph $top, max 200)')
  .option('--no-expand', 'Skip $expand=steps to reduce payload size')
  .option('--next <url>', 'Follow @odata.nextLink from a previous list response (full URL)')
  .option('--all', 'Follow all pages until no @odata.nextLink (may be slow)')
  .action(async (opts: BaseOpts & { top?: string; expand?: boolean; next?: string; all?: boolean }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    if (opts.all && opts.next) {
      console.error('Error: use only one of --all or --next');
      process.exit(1);
    }
    // @odata.nextLink is a full, self-contained continuation URL — it already carries the
    // $top/$expand from the page that produced it, so combining --next with --top/--no-expand
    // would silently do nothing rather than apply a new value.
    if (opts.next && (opts.top || opts.expand === false)) {
      console.error(
        "Error: --top and --no-expand have no effect with --next (the continuation URL already carries the original page's $top/$expand). Omit --next, or drop --top/--no-expand."
      );
      process.exit(1);
    }
    const top = opts.top ? Number.parseInt(opts.top, 10) : undefined;
    if (opts.top && (!Number.isFinite(top) || (top as number) <= 0)) {
      console.error('Error: --top must be a positive integer');
      process.exit(1);
    }
    const expandSteps = opts.expand !== false;

    if (opts.all) {
      const r = await listAllMyApprovals(auth.token, { top, expandSteps });
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message ?? 'approvals list failed'}`);
        process.exit(1);
      }
      const items = r.data;
      if (opts.json) {
        console.log(JSON.stringify({ value: items }, null, 2));
        return;
      }
      if (items.length === 0) {
        console.log('No pending approvals.');
        return;
      }
      for (const a of items) renderApproval(a);
      return;
    }

    const r = await listMyApprovals(auth.token, {
      top,
      expandSteps,
      nextLink: opts.next?.trim()
    });
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message ?? 'approvals list failed'}`);
      process.exit(1);
    }
    const items = r.data.value ?? [];
    const nextLink = r.data['@odata.nextLink'];
    if (opts.json) {
      console.log(
        JSON.stringify(
          {
            value: items,
            ...(nextLink ? { '@odata.nextLink': nextLink } : {})
          },
          null,
          2
        )
      );
      return;
    }
    if (nextLink) {
      console.log(`nextLink: ${nextLink}`);
    }
    if (items.length === 0) {
      console.log('No pending approvals.');
      return;
    }
    for (const a of items) renderApproval(a);
  });

baseFlags(approvalsCommand.command('get <approvalId>'))
  .description('Get a single approval (`GET /me/approvals/{id}?$expand=steps`).')
  .option('--no-expand', 'Skip $expand=steps')
  .action(async (approvalId: string, opts: BaseOpts & { expand?: boolean }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await getApproval(auth.token, approvalId, { expandSteps: opts.expand !== false });
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message ?? 'approval get failed'}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(r.data, null, 2));
      return;
    }
    renderApproval(r.data);
  });

baseFlags(approvalsCommand.command('steps <approvalId>'))
  .description('List steps for an approval (`GET /me/approvals/{id}/steps`).')
  .action(async (approvalId: string, opts: BaseOpts) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await listApprovalSteps(auth.token, approvalId);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message ?? 'approval steps failed'}`);
      process.exit(1);
    }
    const items = r.data.value ?? [];
    if (opts.json) {
      console.log(JSON.stringify({ value: items }, null, 2));
      return;
    }
    if (items.length === 0) {
      console.log('No steps on this approval.');
      return;
    }
    for (const s of items) renderStep(s);
  });

baseFlags(approvalsCommand.command('respond <approvalId> <stepId>'))
  .description(
    'Apply approve or deny on an approval step (`PATCH /me/approvals/{id}/steps/{stepId}`). Requires `ApprovalSolution.ReadWrite` (or `ApprovalSolutionResponse.ReadWrite` for read-and-respond only).'
  )
  .requiredOption('--decision <decision>', 'approve | deny')
  .option('--justification <text>', 'Reviewer justification (optional but commonly required)')
  .action(
    async (
      approvalId: string,
      stepId: string,
      opts: BaseOpts & { decision: string; justification?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const decisionRaw = opts.decision.trim().toLowerCase();
      if (decisionRaw !== 'approve' && decisionRaw !== 'deny') {
        console.error("Error: --decision must be 'approve' or 'deny'");
        process.exit(1);
      }
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await patchApprovalStep(auth.token, approvalId, stepId, {
        reviewResult: decisionRaw === 'approve' ? 'Approve' : 'Deny',
        justification: opts.justification?.trim() || undefined
      });
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message ?? 'approval respond failed'}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(r.data, null, 2));
        return;
      }
      console.log(`✓ ${decisionRaw === 'approve' ? 'Approved' : 'Denied'} step ${stepId}`);
      renderStep(r.data);
    }
  );

baseFlags(approvalsCommand.command('cancel <approvalId>'))
  .description(
    'Delete/cancel an approval you own (`DELETE /me/approvals/{id}` with If-Match). Omit `--if-match` to fetch the latest ETag via `approvals get` first.'
  )
  .option(
    '--if-match <etag>',
    'Required `If-Match` value (`@odata.etag` from `approvals get --json`). When omitted, the CLI fetches the approval once to read the ETag.'
  )
  .action(async (approvalId: string, opts: BaseOpts & { ifMatch?: string }, cmd: Command) => {
    checkReadOnly(cmd);
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    let ifMatch = opts.ifMatch?.trim();
    if (!ifMatch) {
      const gr = await getApproval(auth.token, approvalId, { expandSteps: false });
      if (!gr.ok || !gr.data) {
        console.error(`Error: ${gr.error?.message ?? 'approvals get failed'}`);
        process.exit(1);
      }
      ifMatch = gr.data['@odata.etag']?.trim();
      if (!ifMatch) {
        console.error('Error: approval response missing @odata.etag; pass --if-match explicitly');
        process.exit(1);
      }
    }
    const r = await deleteApproval(auth.token, approvalId, ifMatch);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message ?? 'approval cancel failed'}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify({ cancelled: approvalId }, null, 2));
      return;
    }
    console.log(`Cancelled (deleted) approval ${approvalId}`);
  });

function renderApproval(a: Approval): void {
  console.log(`approval: ${a.id}`);
  if (a.steps?.length) {
    for (const s of a.steps) {
      const tag = s.assignedToMe ? '*' : ' ';
      const reviewed = s.reviewResult ? ` ${s.reviewResult}` : '';
      console.log(
        `  ${tag} step ${s.id} status=${s.status ?? '?'}${reviewed}${s.displayName ? ` "${s.displayName}"` : ''}`
      );
    }
  }
}

function renderStep(s: ApprovalStep): void {
  console.log(`step: ${s.id}`);
  console.log(`  status: ${s.status ?? '?'}`);
  if (s.reviewResult) console.log(`  reviewResult: ${s.reviewResult}`);
  if (s.justification) console.log(`  justification: ${s.justification}`);
  if (s.reviewedDateTime) console.log(`  reviewedAt: ${s.reviewedDateTime}`);
  if (s.assignedToMe) console.log('  assignedToMe: yes');
}
