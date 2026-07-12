import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import {
  listConversationThreads,
  listGroupConversations,
  listMyOutlookGroups,
  listThreadPosts,
  replyToPost
} from '../lib/graph-groups-client.js';
import { toJsonError } from '../lib/json-error.js';
import { checkReadOnly } from '../lib/utils.js';

export const groupsCommand = new Command('groups').description(
  'Outlook Groups (Microsoft 365 unified groups) — list groups, conversations, threads, posts, and reply (delegated `Group.ReadWrite.All`).'
);

interface BaseOpts {
  json?: boolean;
  token?: string;
  identity?: string;
}

/** Prints the --json structured error envelope (or the matching plain-text "Auth error: .../
 *  Error: ...") for the two failure shapes every groups subcommand hits, then exits 1. */
function failGroups(
  json: boolean | undefined,
  prefix: 'Auth error' | 'Error',
  error: unknown,
  fallbackMessage?: string
): never {
  if (json) {
    console.log(JSON.stringify({ error: toJsonError(error, fallbackMessage) }, null, 2));
  } else {
    const message =
      (typeof error === 'string' ? error : (error as { message?: string } | undefined)?.message) ?? fallbackMessage;
    console.error(`${prefix}: ${message}`);
  }
  process.exit(1);
}

/** Strip HTML tags for terminal preview (repeat until stable so nested/obfuscated tags are removed). */
function stripHtmlTagsForConsole(html: string): string {
  let cur = html;
  let prev = '';
  while (cur !== prev) {
    prev = cur;
    cur = cur.replace(/<[^>]+>/g, '');
  }
  return cur.replace(/\s+/g, ' ').trim();
}

const baseFlags = (cmd: Command) =>
  cmd
    .option('--json', 'Output as JSON')
    .option('--token <token>', 'Use a specific Graph token')
    .option('--identity <name>', 'Graph token cache identity (default: default)');

baseFlags(groupsCommand.command('list'))
  .description(
    "List Microsoft 365 / Outlook groups the user belongs to (`GET /me/memberOf/microsoft.graph.group?$filter=groupTypes/any(c:c eq 'Unified')`). Sends `ConsistencyLevel: eventual` per Graph rules."
  )
  .option('--top <n>', 'Limit results (Graph $top, max 200)')
  .action(async (opts: BaseOpts & { top?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      failGroups(opts.json, 'Auth error', auth.error);
    }
    const top = opts.top ? Number.parseInt(opts.top, 10) : undefined;
    if (opts.top && (!Number.isFinite(top) || (top as number) <= 0)) {
      console.error('Error: --top must be a positive integer');
      process.exit(1);
    }
    const r = await listMyOutlookGroups(auth.token, { top });
    if (!r.ok || !r.data) {
      failGroups(opts.json, 'Error', r.error, 'memberOf failed');
    }
    const items = r.data.value ?? [];
    if (opts.json) {
      console.log(JSON.stringify({ value: items }, null, 2));
      return;
    }
    if (items.length === 0) {
      console.log('No Microsoft 365 groups found for the signed-in user.');
      return;
    }
    for (const g of items) {
      const name = g.displayName ?? g.mailNickname ?? '(no name)';
      console.log(`${g.id}\t${name}${g.mail ? `\t${g.mail}` : ''}`);
      if (g.description) console.log(`  ${g.description}`);
    }
  });

baseFlags(groupsCommand.command('conversations <groupId>'))
  .description('List conversations in a Microsoft 365 group (`GET /groups/{id}/conversations`).')
  .option('--top <n>', 'Limit results (Graph $top, max 200)')
  .action(async (groupId: string, opts: BaseOpts & { top?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      failGroups(opts.json, 'Auth error', auth.error);
    }
    const top = opts.top ? Number.parseInt(opts.top, 10) : undefined;
    if (opts.top && (!Number.isFinite(top) || (top as number) <= 0)) {
      console.error('Error: --top must be a positive integer');
      process.exit(1);
    }
    const r = await listGroupConversations(auth.token, groupId, { top });
    if (!r.ok || !r.data) {
      failGroups(opts.json, 'Error', r.error, 'conversations failed');
    }
    const items = r.data.value ?? [];
    if (opts.json) {
      console.log(JSON.stringify({ value: items }, null, 2));
      return;
    }
    if (items.length === 0) {
      console.log('No conversations.');
      return;
    }
    for (const c of items) {
      console.log(`${c.id}\t${c.topic ?? '(no topic)'}`);
      if (c.lastDeliveredDateTime) console.log(`  last: ${c.lastDeliveredDateTime}`);
      if (c.preview) console.log(`  ${c.preview}`);
    }
  });

baseFlags(groupsCommand.command('thread <groupId> <conversationId>'))
  .description('List threads in a group conversation (`GET /groups/{id}/conversations/{id}/threads`).')
  .option('--top <n>', 'Limit results (Graph $top, max 200)')
  .action(async (groupId: string, conversationId: string, opts: BaseOpts & { top?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      failGroups(opts.json, 'Auth error', auth.error);
    }
    const top = opts.top ? Number.parseInt(opts.top, 10) : undefined;
    if (opts.top && (!Number.isFinite(top) || (top as number) <= 0)) {
      console.error('Error: --top must be a positive integer');
      process.exit(1);
    }
    const r = await listConversationThreads(auth.token, groupId, conversationId, { top });
    if (!r.ok || !r.data) {
      failGroups(opts.json, 'Error', r.error, 'threads failed');
    }
    const items = r.data.value ?? [];
    if (opts.json) {
      console.log(JSON.stringify({ value: items }, null, 2));
      return;
    }
    if (items.length === 0) {
      console.log('No threads.');
      return;
    }
    for (const t of items) {
      console.log(`${t.id}\t${t.topic ?? '(no topic)'}${t.isLocked ? '\t[locked]' : ''}`);
      if (t.lastDeliveredDateTime) console.log(`  last: ${t.lastDeliveredDateTime}`);
      if (t.preview) console.log(`  ${t.preview}`);
    }
  });

baseFlags(groupsCommand.command('posts <groupId> <conversationId> <threadId>'))
  .description('List posts in a group thread (`GET /groups/{id}/conversations/{id}/threads/{id}/posts`).')
  .option('--top <n>', 'Limit results (Graph $top, max 200)')
  .action(async (groupId: string, conversationId: string, threadId: string, opts: BaseOpts & { top?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      failGroups(opts.json, 'Auth error', auth.error);
    }
    const top = opts.top ? Number.parseInt(opts.top, 10) : undefined;
    if (opts.top && (!Number.isFinite(top) || (top as number) <= 0)) {
      console.error('Error: --top must be a positive integer');
      process.exit(1);
    }
    const r = await listThreadPosts(auth.token, groupId, conversationId, threadId, { top });
    if (!r.ok || !r.data) {
      failGroups(opts.json, 'Error', r.error, 'posts failed');
    }
    const items = r.data.value ?? [];
    if (opts.json) {
      console.log(JSON.stringify({ value: items }, null, 2));
      return;
    }
    if (items.length === 0) {
      console.log('No posts.');
      return;
    }
    for (const p of items) {
      const sender = p.from?.emailAddress?.address ?? p.sender?.emailAddress?.address ?? '';
      console.log(`${p.id}\t${p.receivedDateTime ?? p.createdDateTime ?? ''}\t${sender}`);
      const c = p.body?.content;
      if (c) {
        const stripped = stripHtmlTagsForConsole(c);
        if (stripped) console.log(`  ${stripped.slice(0, 200)}${stripped.length > 200 ? '…' : ''}`);
      }
    }
  });

baseFlags(groupsCommand.command('post-reply <groupId> <conversationId> <threadId> <postId>'))
  .description(
    'Reply to a specific post in a group thread (`POST .../posts/{id}/reply`). Use `--text` for plain text or `--html` for HTML.'
  )
  .option('--text <body>', 'Plain text reply body')
  .option('--html <body>', 'HTML reply body (overrides --text)')
  .action(
    async (
      groupId: string,
      conversationId: string,
      threadId: string,
      postId: string,
      opts: BaseOpts & { text?: string; html?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      if (!opts.text?.trim() && !opts.html?.trim()) {
        console.error('Error: provide --text or --html');
        process.exit(1);
      }
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        failGroups(opts.json, 'Auth error', auth.error);
      }
      const r = await replyToPost(
        auth.token,
        { groupId, conversationId, threadId, postId },
        opts.html?.trim()
          ? { contentType: 'html', content: opts.html }
          : { contentType: 'text', content: opts.text ?? '' }
      );
      if (!r.ok) {
        failGroups(opts.json, 'Error', r.error, 'reply failed');
      }
      if (opts.json) {
        console.log(JSON.stringify({ ok: true, postId }, null, 2));
        return;
      }
      console.log(`✓ Replied to post ${postId}`);
    }
  );
