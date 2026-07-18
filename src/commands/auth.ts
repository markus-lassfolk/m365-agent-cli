import { Command } from 'commander';
import type { AuthFailureClass } from '../lib/auth-diagnostics.js';
import { diagnoseAuth } from '../lib/auth-diagnostics.js';
import { getDefaultProfileIdentity } from '../lib/identity-profiles.js';
import { toJsonError } from '../lib/json-error.js';
import { LoginAccountMismatchError } from '../lib/login-identity-binding.js';
import { applyEnvFileOverrides, resolveEnvFilePathArgument } from '../lib/utils.js';
import { runDeviceCodeLogin } from './login.js';

const FAILURE_DESCRIPTIONS: Record<AuthFailureClass, string> = {
  healthy: 'authenticated and ready',
  missing_credentials: 'missing EWS_CLIENT_ID or refresh token — never logged in (or env not loaded)',
  missing_cache: 'no usable cached token and refresh could not be diagnosed further',
  malformed_cache: 'local token cache file is malformed',
  refresh_grant_revoked: 'refresh grant revoked by tenant policy (e.g. password reset, admin revocation)',
  refresh_grant_expired: 'refresh grant expired (age or inactivity)',
  interaction_required: 'Azure requires interactive re-authentication (MFA/conditional access changed)',
  mfa_required: 'multi-factor authentication is required and was not satisfied',
  conditional_access_blocked: 'blocked by a Conditional Access policy',
  consent_required: 'admin or user consent is required for one or more scopes',
  tenant_client_mismatch: 'the Entra app registration (client id) or tenant does not match this token',
  unknown_error: 'authentication failed for an unclassified reason'
};

interface AuthRepairOptions {
  identity?: string;
  startLogin?: boolean;
  json?: boolean;
  envFile?: string;
  secrets?: boolean;
}

const repairCmd = new Command('repair')
  .description('Diagnose auth failures (revoked/expired delegated auth) and print the safe recovery path')
  .option('--identity <name>', 'Identity/cache slot to diagnose (default: the default profile, else "default")')
  .option('--start-login', 'If repair is required, immediately launch the device-code login flow')
  .option('--json', 'Output as JSON')
  .option('--env-file <path>', 'Load EWS_CLIENT_ID / refresh token from this file before diagnosing')
  .option(
    '--no-secrets',
    'No-op: this command never prints raw access/refresh token material regardless of this flag (kept for explicit automation intent)'
  )
  .action(async (opts: AuthRepairOptions) => {
    if (opts.envFile) {
      applyEnvFileOverrides(resolveEnvFilePathArgument(opts.envFile));
    }
    const resolvedEnvPath = opts.envFile ? resolveEnvFilePathArgument(opts.envFile) : undefined;
    const identity = opts.identity || (await getDefaultProfileIdentity()) || 'default';

    const diag = await diagnoseAuth({ identity, envPath: resolvedEnvPath });

    if (opts.json) {
      console.log(JSON.stringify(diag, null, 2));
    } else {
      console.log(`M365 auth status: ${diag.status === 'healthy' ? 'healthy' : 'repair required'}`);
      console.log(`Identity: ${diag.signedInAs ?? diag.identity}`);
      console.log(`Failure: ${FAILURE_DESCRIPTIONS[diag.failureClass]}`);
      if (diag.evidence.length > 0) {
        console.log(`Evidence: ${diag.evidence.join('; ')}`);
      }
      if (diag.status === 'repair_required') {
        console.log(`Recommended action: ${diag.recommendedAction}`);
        if (diag.safeCommand) {
          console.log(`Command: ${diag.safeCommand}`);
        }
      }
      console.log('Safety: no secrets printed');
    }

    if (diag.status === 'repair_required' && opts.startLogin) {
      if (!opts.json) {
        console.log('\nStarting interactive login...');
      }
      try {
        await runDeviceCodeLogin({ envFile: opts.envFile, identity: opts.identity });
      } catch (err) {
        const message =
          err instanceof LoginAccountMismatchError
            ? err.message
            : `Login failed: ${err instanceof Error ? err.message : String(err)}`;
        if (opts.json) {
          console.log(JSON.stringify({ error: toJsonError(message) }, null, 2));
        } else {
          console.error(`Error: ${message}`);
        }
        process.exit(1);
      }
    }

    // Exit 0 even when repair is required — `status`/`failureClass` carry that signal, per #243's
    // acceptance criteria ("exits non-zero only for tool/runtime failures").
  });

export const authCommand = new Command('auth')
  .description('Authentication diagnostics and repair')
  .addCommand(repairCmd);
