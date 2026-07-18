import { Command } from 'commander';
import { approvalsCommand } from '../commands/approvals.js';
import { authCommand } from '../commands/auth.js';
import { autoReplyCommand } from '../commands/auto-reply.js';
import { bookingsCommand } from '../commands/bookings.js';
import { calendarCommand } from '../commands/calendar.js';
import { contactsCommand } from '../commands/contacts.js';
import { copilotCommand } from '../commands/copilot.js';
import { counterCommand } from '../commands/counter.js';
import { createEventCommand } from '../commands/create-event.js';
import { delegatesCommand } from '../commands/delegates.js';
import { deleteEventCommand } from '../commands/delete-event.js';
import { describeCommand } from '../commands/describe.js';
import { doctorCommand } from '../commands/doctor.js';
import { draftsCommand } from '../commands/drafts.js';
import { excelCommand } from '../commands/excel.js';
import { filesCommand } from '../commands/files.js';
import { findCommand } from '../commands/find.js';
import { findtimeCommand } from '../commands/findtime.js';
import { foldersCommand } from '../commands/folders.js';
import { forwardEventCommand } from '../commands/forward-event.js';
import { graphCommand } from '../commands/graph.js';
import { graphCalendarCommand } from '../commands/graph-calendar.js';
import { graphSearchCommand } from '../commands/graph-search.js';
import { groupsCommand } from '../commands/groups.js';
import { insightsCommand } from '../commands/insights.js';
import { loginCommand } from '../commands/login.js';
import { mailCommand } from '../commands/mail.js';
import { mailboxSettingsCommand } from '../commands/mailbox-settings.js';
import { mcpCommand } from '../commands/mcp.js';
import { meetingCommand } from '../commands/meeting.js';
import { onenoteCommand } from '../commands/onenote.js';
import { oofCommand } from '../commands/oof.js';
import { orgCommand } from '../commands/org.js';
import { outlookCategoriesCommand } from '../commands/outlook-categories.js';
import { outlookGraphCommand } from '../commands/outlook-graph.js';
import { peopleCommand } from '../commands/people.js';
import { plannerCommand } from '../commands/planner.js';
import { powerpointCommand } from '../commands/powerpoint.js';
import { presenceCommand } from '../commands/presence.js';
import { profilesCommand } from '../commands/profiles.js';
import { readinessCommand } from '../commands/readiness.js';
import { respondCommand } from '../commands/respond.js';
import { roomsCommand } from '../commands/rooms.js';
import { rulesCommand } from '../commands/rules.js';
import { scheduleCommand } from '../commands/schedule.js';
import { sendCommand } from '../commands/send.js';
import { serveCommand } from '../commands/serve.js';
import { sharepointCommand } from '../commands/sharepoint.js';
import { sitePagesCommand } from '../commands/site-pages.js';
import { subscribeCommand } from '../commands/subscribe.js';
import { subscriptionsCommand } from '../commands/subscriptions.js';
import { suggestCommand } from '../commands/suggest.js';
import { teamsCommand } from '../commands/teams.js';
import { todoCommand } from '../commands/todo.js';
import { updateCommand } from '../commands/update.js';
import { updateEventCommand } from '../commands/update-event.js';
import { verifyTokenCommand } from '../commands/verify-token.js';
import { vivaCommand } from '../commands/viva.js';
import { whoamiCommand } from '../commands/whoami.js';
import { wordCommand } from '../commands/word.js';
import { checkIdentityGuards } from './identity-guard.js';
import { toJsonError } from './json-error.js';
import { installM365HelpOnCommandTree } from './m365-help.js';
import { getPackageVersionSync } from './package-info.js';

/**
 * `login`, `auth` (repair), and `profiles` exist specifically to establish or fix the signed-in
 * identity that `--require-identity`/`--as-delegate-of` verify — applying the root guard to them
 * too would make e.g. `login --require-identity X` on a machine with no verifiable identity yet
 * permanently unable to succeed (the guard blocks the exact command meant to satisfy it). These
 * commands already have their own finer-grained wrong-account protection (`login-identity-binding.ts`'s
 * `LoginAccountMismatchError`, refused unless `--force-identity-switch`), so exempting them from
 * the coarse root guard doesn't remove wrong-account protection — it relies on the right layer.
 */
const IDENTITY_GUARD_EXEMPT_COMMANDS: ReadonlySet<string> = new Set(['login', 'auth', 'profiles']);

/** Walk up to the top-level command name (e.g. `repair` under `auth` → `'auth'`). */
function topLevelCommandName(cmd: Command): string {
  let current = cmd;
  while (current.parent && current.parent.parent) {
    current = current.parent;
  }
  return current.name();
}

/** Register all root subcommands and options (same tree as `src/cli.ts`). */
function registerM365Commands(program: Command): void {
  program
    .name('m365-agent-cli')
    .description('Microsoft 365 from your terminal: calendar, mail, OneDrive, Planner, Teams, Graph — one OAuth login')
    .version(getPackageVersionSync());

  program.option('--read-only', 'Run in read-only mode, blocking any mutating operations');
  program.option(
    '--dry-run',
    'Preview the exact request a mutating command would send (Graph URL/body or EWS SOAP envelope) without sending it, then exit 0. For a multi-step command, only the first mutating request is shown.'
  );
  program.option(
    '--cache <duration>',
    'Cache idempotent Microsoft Graph GET responses on disk for this long (e.g. 30s, 5m, 1h; bare number = seconds), keyed by identity token + URL. Off by default.'
  );
  program.option(
    '--require-identity <upn>',
    'Wrong-account guardrail: fail before running any command unless the signed-in identity (decoded from the access token) matches this UPN/email exactly.'
  );
  program.option(
    '--as-delegate-of <upn>',
    'Wrong-account guardrail: assert you are signed in as this UPN before operating on --mailbox (distinguishes "signed in as X" from "operating on mailbox Y via delegation"). Requires --mailbox on the command being run.'
  );

  // Sync --dry-run / --cache into process-wide env vars before the action runs: the transport
  // layer (graph-client.ts / ews-client.ts) has no access to the parsed Command instance, so this
  // is the only way to carry the flags down to where requests are actually built and sent.
  program.hook('preAction', async (_thisCommand, actionCommand) => {
    const opts = actionCommand.optsWithGlobals();
    if (opts.dryRun) {
      process.env.M365_DRY_RUN = '1';
    }
    if (opts.cache) {
      process.env.M365_CACHE_TTL = String(opts.cache);
    }
    if (
      (opts.requireIdentity || opts.asDelegateOf) &&
      !IDENTITY_GUARD_EXEMPT_COMMANDS.has(topLevelCommandName(actionCommand))
    ) {
      const guard = await checkIdentityGuards({
        identity: opts.identity,
        requireIdentity: opts.requireIdentity,
        asDelegateOf: opts.asDelegateOf,
        mailbox: opts.mailbox
      });
      if (!guard.ok) {
        if (opts.json) {
          console.log(JSON.stringify({ error: toJsonError(guard.message) }, null, 2));
        } else {
          console.error(`Error: ${guard.message}`);
        }
        process.exit(1);
      }
    }
  });

  program.addHelpText('after', 'Tip: run m365-agent-cli <command> --help for flags and examples on each command.');

  program.addCommand(whoamiCommand);
  program.addCommand(profilesCommand);
  program.addCommand(authCommand);
  program.addCommand(readinessCommand);
  program.addCommand(doctorCommand);
  program.addCommand(describeCommand);
  program.addCommand(mcpCommand);
  program.addCommand(updateCommand);
  program.addCommand(loginCommand);
  program.addCommand(sitePagesCommand);
  program.addCommand(verifyTokenCommand);
  program.addCommand(autoReplyCommand);
  program.addCommand(calendarCommand);
  program.addCommand(findtimeCommand);
  program.addCommand(respondCommand);
  program.addCommand(createEventCommand);
  program.addCommand(deleteEventCommand);
  program.addCommand(findCommand);
  program.addCommand(peopleCommand);
  program.addCommand(orgCommand);
  program.addCommand(updateEventCommand);
  program.addCommand(mailCommand);
  program.addCommand(foldersCommand);
  program.addCommand(sendCommand);
  program.addCommand(draftsCommand);
  program.addCommand(filesCommand);
  program.addCommand(excelCommand);
  program.addCommand(forwardEventCommand);
  program.addCommand(counterCommand);
  program.addCommand(scheduleCommand);
  program.addCommand(suggestCommand);
  program.addCommand(vivaCommand);
  program.addCommand(subscribeCommand);
  program.addCommand(subscriptionsCommand);
  program.addCommand(serveCommand);
  program.addCommand(roomsCommand);
  program.addCommand(oofCommand);
  program.addCommand(mailboxSettingsCommand);
  program.addCommand(rulesCommand);
  program.addCommand(delegatesCommand);
  program.addCommand(todoCommand);
  program.addCommand(contactsCommand);
  program.addCommand(copilotCommand);
  program.addCommand(meetingCommand);
  program.addCommand(onenoteCommand);

  program.addCommand(outlookCategoriesCommand);
  program.addCommand(outlookGraphCommand);
  program.addCommand(graphCalendarCommand);
  program.addCommand(graphSearchCommand);
  program.addCommand(graphCommand);
  program.addCommand(teamsCommand);
  program.addCommand(bookingsCommand);
  program.addCommand(presenceCommand);

  program.addCommand(insightsCommand);
  program.addCommand(groupsCommand);
  program.addCommand(approvalsCommand);

  program.addCommand(plannerCommand);

  program.addCommand(wordCommand);
  program.addCommand(powerpointCommand);

  program.addCommand(sharepointCommand);

  installM365HelpOnCommandTree(program);
}

export function createM365Program(): Command {
  const program = new Command();
  registerM365Commands(program);
  return program;
}
