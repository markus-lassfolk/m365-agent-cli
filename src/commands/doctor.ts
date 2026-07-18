import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { Command } from 'commander';
import { buildDoctorBundle } from '../lib/doctor-bundle.js';
import { getDefaultProfileIdentity } from '../lib/identity-profiles.js';
import { createStoredZip } from '../lib/minimal-zip.js';
import { applyEnvFileOverrides, resolveEnvFilePathArgument, resolveOutputFilePath } from '../lib/utils.js';

interface DoctorOptions {
  identity?: string;
  mailbox?: string;
  envFile?: string;
  json?: boolean;
  redactedBundle?: string | boolean;
  format?: string;
  output?: string;
}

export const doctorCommand = new Command('doctor')
  .description('Non-secret diagnostic summary for auth/mailbox failures — safe to attach to an issue')
  .option('--identity <name>', 'Identity/cache slot to diagnose (default: the default profile, else "default")')
  .option('--mailbox <email>', 'Also check delegated/shared access to this mailbox')
  .option('--env-file <path>', 'Load EWS_CLIENT_ID / refresh token from this file before diagnosing')
  .option('--json', 'Print the bundle as JSON to stdout instead of a human summary')
  .option(
    '--redacted-bundle [path]',
    'Write a shareable diagnostic bundle. Path ending in .zip writes a zip archive (default ./m365-diagnostic.zip); otherwise see --format'
  )
  .option('--format <dir|zip>', 'Bundle format for --redacted-bundle (default: zip, or dir with --output)', 'zip')
  .option('--output <path>', 'Output path for --redacted-bundle (directory when --format dir)')
  .action(async (opts: DoctorOptions) => {
    if (opts.envFile) {
      applyEnvFileOverrides(resolveEnvFilePathArgument(opts.envFile));
    }
    const resolvedEnvPath = opts.envFile ? resolveEnvFilePathArgument(opts.envFile) : undefined;
    const identity = opts.identity || (await getDefaultProfileIdentity()) || 'default';

    const bundle = await buildDoctorBundle({
      identity,
      mailbox: opts.mailbox,
      envPath: resolvedEnvPath
    });

    if (opts.redactedBundle !== undefined) {
      const format = opts.format === 'dir' ? 'dir' : 'zip';
      const json = JSON.stringify(bundle, null, 2);
      // `--redacted-bundle` alone (no value) parses as boolean `true`, not a path — only use it as
      // a path when Commander actually captured a string argument.
      const explicitPath = typeof opts.redactedBundle === 'string' ? opts.redactedBundle : undefined;

      if (format === 'dir') {
        const outDir = resolveOutputFilePath(opts.output || explicitPath || './m365-diagnostic');
        await mkdir(outDir, { recursive: true });
        const target = join(outDir, 'diagnostic.json');
        await writeFile(target, json, 'utf8');
        console.log(`Wrote redacted diagnostic bundle: ${target}`);
      } else {
        const target = resolveOutputFilePath(opts.output || explicitPath || './m365-diagnostic.zip');
        await mkdir(dirname(target), { recursive: true });
        const zip = createStoredZip([{ name: 'diagnostic.json', content: Buffer.from(json, 'utf8') }]);
        await writeFile(target, zip);
        console.log(`Wrote redacted diagnostic bundle: ${target}`);
      }
      console.log('Bundle contains no tokens, passwords, or message content — safe to attach to an issue.');
      return;
    }

    if (opts.json) {
      console.log(JSON.stringify(bundle, null, 2));
      return;
    }

    console.log(`m365-agent-cli ${bundle.cli.version} — doctor\n`);
    console.log(`CLI version:    ${bundle.cli.version}`);
    console.log(`Node:           ${bundle.cli.nodeVersion}`);
    console.log(`Platform:       ${bundle.cli.platform} (${bundle.cli.arch}), ${bundle.cli.osRelease}`);
    console.log(`Config dir:     ${bundle.config.configDir}`);
    console.log(
      `Env file:       ${bundle.config.envFile.path} (${bundle.config.envFile.exists ? 'present' : 'missing'})`
    );
    console.log(`Backend:        ${bundle.exchangeBackend}`);
    console.log(`Client ID:      ${bundle.clientId ?? '(not set)'}`);
    console.log(`Default profile: ${bundle.profiles.defaultProfile ?? '(none)'}`);
    console.log('');
    console.log(`Identity:       ${bundle.identity.name}`);
    console.log(
      `Cache file:     ${bundle.identity.cacheFile.path} (${bundle.identity.cacheFile.exists ? `present, ${bundle.identity.cacheFile.sizeBytes} bytes` : 'missing'})`
    );
    console.log(`Auth status:    ${bundle.authDiagnosis.status}`);
    console.log(`Failure class:  ${bundle.authDiagnosis.failureClass}`);
    if (bundle.authDiagnosis.evidence.length > 0) {
      console.log(`Evidence:       ${bundle.authDiagnosis.evidence.join('; ')}`);
    }
    if (bundle.authDiagnosis.status !== 'healthy') {
      console.log(`Recommended:    ${bundle.authDiagnosis.recommendedAction}`);
      if (bundle.authDiagnosis.safeCommand) console.log(`Command:        ${bundle.authDiagnosis.safeCommand}`);
    }
    if (bundle.mailboxCheck) {
      console.log(
        `Mailbox check:  ${bundle.mailboxCheck.mailbox} — ${bundle.mailboxCheck.ok ? 'ok' : 'failed/unchecked'}`
      );
    }
    console.log('\nNo tokens, passwords, or message content are printed above.');
    console.log(
      'Tip: `doctor --redacted-bundle` writes a shareable file; `doctor --json` for machine-readable output.'
    );
  });
