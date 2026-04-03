import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import { getMyPresence, getUserPresence } from '../lib/graph-presence-client.js';

export const presenceCommand = new Command('presence').description(
  'User presence (Graph cloud communications); requires Presence.Read (see GRAPH_SCOPES.md)'
);

presenceCommand
  .command('me')
  .description('Get signed-in user presence (GET /me/presence)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await getMyPresence(auth.token);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log(opts.json ? JSON.stringify(r.data, null, 2) : `${r.data.availability ?? ''}\t${r.data.activity ?? ''}`);
  });

presenceCommand
  .command('user')
  .description('Get presence for a user (GET /users/{id|upn}/presence)')
  .argument('<user>', 'User id (GUID) or UPN')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (user: string, opts: { json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await getUserPresence(auth.token, user);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log(opts.json ? JSON.stringify(r.data, null, 2) : `${r.data.availability ?? ''}\t${r.data.activity ?? ''}`);
  });
