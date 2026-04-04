import { Command } from 'commander';
import { resolveAuth } from '../lib/auth.js';
import { getOwaUserInfo } from '../lib/ews-client.js';
import { getExchangeBackend } from '../lib/exchange-backend.js';
import { warnAutoGraphToEwsFallback } from '../lib/exchange-fallback-hint.js';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import { callGraph, type GraphResponse } from '../lib/graph-client.js';

interface GraphMe {
  displayName?: string;
  userPrincipalName?: string;
  mail?: string;
}

async function fetchGraphMe(token: string): Promise<GraphResponse<GraphMe>> {
  return callGraph<GraphMe>(token, '/me');
}

export const whoamiCommand = new Command('whoami')
  .description('Show authenticated user information (Graph or EWS per M365_EXCHANGE_BACKEND)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific access token (Graph bearer when using graph/auto; EWS when using ews)')
  .option('--identity <name>', 'Use a specific authentication identity (default: default)')
  .action(async (options: { json?: boolean; token?: string; identity?: string }) => {
    const backend = getExchangeBackend();

    const outputGraph = (
      displayName: string,
      email: string,
      opts: { json?: boolean; identity?: string; token?: string }
    ) => {
      if (opts.json) {
        const result: {
          displayName: string;
          email: string;
          authenticated: boolean;
          backend: string;
          identity?: string;
        } = {
          displayName,
          email,
          authenticated: true,
          backend: 'graph'
        };
        if (!opts.token) {
          result.identity = opts.identity || 'default';
        }
        console.log(JSON.stringify(result, null, 2));
      } else {
        console.log('\u2713 Authenticated (Microsoft Graph)');
        if (!opts.token) {
          console.log(`  Identity: ${opts.identity || 'default'}`);
        }
        console.log(`  Backend: graph (M365_EXCHANGE_BACKEND=${backend})`);
        console.log(`  Name: ${displayName}`);
        console.log(`  Email: ${email}`);
      }
    };

    const outputEws = (
      displayName: string,
      email: string,
      opts: { json?: boolean; identity?: string; token?: string }
    ) => {
      if (opts.json) {
        const result: {
          displayName: string;
          email: string;
          authenticated: boolean;
          backend: string;
          identity?: string;
        } = {
          displayName,
          email,
          authenticated: true,
          backend: 'ews'
        };
        if (!opts.token) {
          result.identity = opts.identity || 'default';
        }
        console.log(JSON.stringify(result, null, 2));
      } else {
        console.log('\u2713 Authenticated (EWS)');
        if (!opts.token) {
          console.log(`  Identity: ${opts.identity || 'default'}`);
        }
        console.log(`  Backend: ews (M365_EXCHANGE_BACKEND=${backend})`);
        console.log(`  Name: ${displayName}`);
        console.log(`  Email: ${email}`);
      }
    };

    const runEws = async (): Promise<void> => {
      const authResult = await resolveAuth({
        token: options.token,
        identity: options.identity
      });

      if (!authResult.success) {
        if (options.json) {
          console.log(JSON.stringify({ error: authResult.error, backend: 'ews' }, null, 2));
        } else {
          console.error(`Error: ${authResult.error}`);
          console.error('\nCheck your .env for EWS_CLIENT_ID and EWS_REFRESH_TOKEN.');
        }
        process.exit(1);
      }

      const userInfo = await getOwaUserInfo(authResult.token!);

      if (!userInfo.ok || !userInfo.data) {
        if (options.json) {
          console.log(
            JSON.stringify(
              {
                error: userInfo.error?.message || 'Failed to fetch user info',
                authenticated: true,
                backend: 'ews'
              },
              null,
              2
            )
          );
        } else {
          console.log('\u2713 Authenticated');
          console.log('  Could not fetch user details from EWS API');
        }
        process.exit(0);
      }

      const { displayName, email } = userInfo.data;
      outputEws(displayName, email, options);
    };

    const runGraph = async (): Promise<void> => {
      const authResult = await resolveGraphAuth({
        token: options.token,
        identity: options.identity
      });

      if (!authResult.success || !authResult.token) {
        if (options.json) {
          console.log(
            JSON.stringify(
              {
                error: authResult.error || 'Graph authentication failed',
                backend: 'graph'
              },
              null,
              2
            )
          );
        } else {
          console.error(`Error: ${authResult.error || 'Graph authentication failed'}`);
          console.error(
            '\nFor Graph, set EWS_CLIENT_ID and M365_REFRESH_TOKEN (Microsoft Graph scopes), or run `m365-agent-cli login`.'
          );
        }
        process.exit(1);
      }

      const userInfo = await fetchGraphMe(authResult.token);

      if (!userInfo.ok || !userInfo.data) {
        if (options.json) {
          console.log(
            JSON.stringify(
              {
                error: userInfo.error?.message || 'Failed to fetch user from Graph',
                authenticated: false,
                backend: 'graph'
              },
              null,
              2
            )
          );
        } else {
          console.error(`Error: ${userInfo.error?.message || 'Failed to fetch user from Graph'}`);
        }
        process.exit(1);
      }

      const displayName = userInfo.data.displayName || '(unknown)';
      const email = userInfo.data.mail || userInfo.data.userPrincipalName || '';
      outputGraph(displayName, email, options);
    };

    if (backend === 'graph') {
      await runGraph();
      return;
    }

    if (backend === 'ews') {
      await runEws();
      return;
    }

    // auto: Graph first, then EWS
    const graphAuth = await resolveGraphAuth({
      token: options.token,
      identity: options.identity
    });
    if (graphAuth.success && graphAuth.token) {
      try {
        const userInfo = await fetchGraphMe(graphAuth.token);
        if (userInfo.ok && userInfo.data) {
          const displayName = userInfo.data.displayName || '(unknown)';
          const email = userInfo.data.mail || userInfo.data.userPrincipalName || '';
          outputGraph(displayName, email, options);
          return;
        }
        warnAutoGraphToEwsFallback('whoami', {
          json: options.json,
          graphError: userInfo.error?.message,
          reason: 'api'
        });
      } catch (err) {
        const msg = err instanceof Error ? err.message : String(err);
        warnAutoGraphToEwsFallback('whoami', { json: options.json, graphError: msg });
      }
    } else {
      warnAutoGraphToEwsFallback('whoami', {
        json: options.json,
        graphError: graphAuth.error,
        reason: 'auth'
      });
    }

    await runEws();
  });
