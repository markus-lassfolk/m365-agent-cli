import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import {
  evaluateGraphCapabilities,
  formatCapabilityTextTable,
  type GraphTokenPayloadForCapabilities,
  graphTokenPermissionKind,
  permissionSetFromGraphPayload
} from '../lib/graph-capability-matrix.js';
import { getJwtPayloadAppId } from '../lib/jwt-utils.js';
import { applyEnvFileOverrides, getGlobalEnvFilePath, resolveEnvFilePathArgument } from '../lib/utils.js';

function kindLabel(kind: ReturnType<typeof graphTokenPermissionKind>): string {
  switch (kind) {
    case 'delegated':
      return 'Delegated (scp)';
    case 'application':
      return 'Application (roles)';
    case 'mixed':
      return 'Mixed (scp + roles)';
    default:
      return 'Unknown — no scp/roles on token';
  }
}

export const verifyTokenCommand = new Command('verify-token')
  .description('Verify Microsoft Graph token scopes and optional CLI feature coverage matrix')
  .option('--token <token>', 'Use a specific Graph token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--json', 'Output as JSON')
  .option(
    '--capabilities',
    'Show read/write checkboxes for m365-agent-cli areas (Planner, SharePoint, mail, …) from this token'
  )
  .option('--verbose', 'With --capabilities, print a detail line per feature')
  .option(
    '--env-file <path>',
    'Load EWS_CLIENT_ID and refresh token from this file (e.g. ~/.config/m365-agent-cli/.env.beta) before verifying'
  )
  .action(
    async (options: {
      token?: string;
      json?: boolean;
      identity?: string;
      envFile?: string;
      capabilities?: boolean;
      verbose?: boolean;
    }) => {
      if (options.envFile) {
        applyEnvFileOverrides(resolveEnvFilePathArgument(options.envFile));
      }
      const authResult = await resolveGraphAuth({ token: options.token, identity: options.identity });
      if (!authResult.success || !authResult.token) {
        if (options.json) {
          console.log(JSON.stringify({ error: authResult.error || 'Failed to resolve auth token' }, null, 2));
        } else {
          console.error(`Error: ${authResult.error || 'Failed to resolve auth token'}`);
        }
        process.exit(1);
      }

      const token = authResult.token;
      try {
        const parts = token.split('.');
        if (parts.length !== 3) throw new Error('Invalid JWT format');
        const payload = JSON.parse(Buffer.from(parts[1], 'base64url').toString('utf8')) as Record<string, unknown>;
        const capPayload = payload as GraphTokenPayloadForCapabilities;

        if (options.capabilities && options.json) {
          const kind = graphTokenPermissionKind(capPayload);
          const perms = permissionSetFromGraphPayload(capPayload);
          const evaluated = evaluateGraphCapabilities(perms);
          console.log(
            JSON.stringify(
              {
                tokenKind: kind,
                tokenKindLabel: kindLabel(kind),
                appId: getJwtPayloadAppId(token) ?? null,
                tenantId: typeof payload.tid === 'string' ? payload.tid : null,
                userPrincipalName: typeof payload.upn === 'string' ? payload.upn : null,
                permissions: [...perms].sort(),
                capabilities: evaluated.map((r) => ({
                  id: r.id,
                  area: r.area,
                  detail: r.detail,
                  read: r.notApplicable || r.readColumnDash ? null : r.readOk,
                  write: r.notApplicable || r.writeColumnDash || r.writeScopes.length === 0 ? null : r.writeOk,
                  notApplicable: r.notApplicable ?? false,
                  readColumnDash: Boolean(r.readColumnDash),
                  writeColumnDash: Boolean(r.writeColumnDash)
                })),
                notes: [
                  'Matrix is heuristic from Microsoft Graph permission names on this access token.',
                  'EWS commands need Exchange Online delegated permission EWS.AccessAsUser.All (not part of Graph scp).',
                  '`graph invoke` / `graph batch` depend on the path you call.',
                  'Graph Search needs entity-specific scopes for each content type; see graph-search command help.'
                ]
              },
              null,
              2
            )
          );
          return;
        }

        if (options.capabilities && !options.json) {
          const envPath = options.envFile ? resolveEnvFilePathArgument(options.envFile) : getGlobalEnvFilePath();
          console.log(`Configuration file: ${envPath}`);
          if (process.env.EWS_CLIENT_ID) {
            console.log(`EWS_CLIENT_ID (Entra app): ${process.env.EWS_CLIENT_ID}`);
          }
          console.log('');

          const kind = graphTokenPermissionKind(capPayload);
          const perms = permissionSetFromGraphPayload(capPayload);
          const evaluated = evaluateGraphCapabilities(perms);
          const tokenAppId = getJwtPayloadAppId(token);

          console.log('\u2713 Graph token — CLI feature coverage\n');
          console.log(`  Token type: ${kindLabel(kind)}`);
          console.log(`  App ID:     ${tokenAppId ?? 'N/A'}`);
          console.log(`  Tenant:     ${payload.tid || 'N/A'}`);
          console.log(`  User:       ${payload.upn || payload.email || 'N/A'}`);
          console.log(`  Name:       ${payload.name || 'N/A'}`);
          console.log(`  Distinct permissions on token: ${perms.size}`);
          console.log('');
          console.log(formatCapabilityTextTable(evaluated, { verbose: options.verbose ?? false }));
          console.log('');
          console.log('Notes:');
          console.log(
            '  • Read can be satisfied by a broader scope (e.g. Calendars.ReadWrite counts as calendar read).'
          );
          console.log(
            '  • Write shows — when this CLI area has no separate write permission (e.g. rooms are read-only).'
          );
          console.log('  • EWS is not shown: add EWS.AccessAsUser.All under Office 365 Exchange Online in Entra.');
          console.log('  • For raw `scp` / `roles`, run: verify-token (without --capabilities).');

          const envClient = process.env.EWS_CLIENT_ID?.trim();
          const tokenApp = tokenAppId?.trim();
          if (envClient && tokenApp && envClient.toLowerCase() !== tokenApp.toLowerCase()) {
            console.log(
              '\n  Warning: Access token app id does not match EWS_CLIENT_ID. Use verify-token --env-file for .env.beta, or align cache / login.'
            );
          }
          return;
        }

        if (options.json) {
          console.log(JSON.stringify(payload, null, 2));
          return;
        }

        const envPath = options.envFile ? resolveEnvFilePathArgument(options.envFile) : getGlobalEnvFilePath();
        console.log(`Configuration file: ${envPath}`);
        if (process.env.EWS_CLIENT_ID) {
          console.log(`EWS_CLIENT_ID (Entra app): ${process.env.EWS_CLIENT_ID}`);
        } else {
          console.log('EWS_CLIENT_ID: (not set — check .env or login)');
        }
        if (!options.envFile && !process.env.M365_AGENT_ENV_FILE?.trim()) {
          console.log(
            'Tip: For a second app (.env.beta), set $env:M365_AGENT_ENV_FILE to that path in PowerShell, or use: verify-token --env-file "$env:USERPROFILE\\.config\\m365-agent-cli\\.env.beta"'
          );
        }
        console.log('');

        console.log('\u2713 Token Verified\n');
        const tokenAppId = getJwtPayloadAppId(token);
        console.log(`  App ID (from access token): ${tokenAppId ?? 'N/A'}`);
        console.log(`  Tenant: ${payload.tid || 'N/A'}`);
        console.log(`  User:   ${payload.upn || payload.email || 'N/A'}`);
        console.log(`  Name:   ${payload.name || 'N/A'}`);

        if (payload.scp) {
          console.log('\n  Delegated Scopes (scp):');
          String(payload.scp)
            .split(' ')
            .forEach((scope: string) => {
              if (scope) console.log(`    - ${scope}`);
            });
        }

        if (payload.roles && Array.isArray(payload.roles)) {
          console.log('\n  Application Roles (roles):');
          (payload.roles as string[]).forEach((role: string) => {
            console.log(`    - ${role}`);
          });
        }

        if (!payload.scp && (!payload.roles || (payload.roles as unknown[]).length === 0)) {
          console.log('\n  No scopes or roles found in token.');
        }

        console.log('\n  Tip: m365-agent-cli verify-token --capabilities — see read/write coverage for CLI features.');

        const envClient = process.env.EWS_CLIENT_ID?.trim();
        const tokenApp = tokenAppId?.trim();
        if (envClient && tokenApp && envClient.toLowerCase() !== tokenApp.toLowerCase()) {
          console.log(
            '\n  Warning: Access token app id does not match EWS_CLIENT_ID. Use verify-token --env-file for .env.beta, set M365_AGENT_ENV_FILE before starting the CLI, or delete ~/.config/m365-agent-cli/token-cache-*.json and run login again.'
          );
        }
      } catch (err: unknown) {
        const message = err instanceof Error ? err.message : String(err);
        if (options.json) {
          console.log(JSON.stringify({ error: `Failed to parse token: ${message}` }, null, 2));
        } else {
          console.error(`Failed to parse token: ${message}`);
        }
        process.exit(1);
      }
    }
  );
