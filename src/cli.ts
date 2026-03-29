#!/usr/bin/env bun
import { Command } from 'commander';
import { autoReplyCommand } from './commands/auto-reply.js';
import { calendarCommand } from './commands/calendar.js';
import { counterCommand } from './commands/counter.js';
import { createEventCommand } from './commands/create-event.js';
import { delegatesCommand } from './commands/delegates.js';
import { deleteEventCommand } from './commands/delete-event.js';
import { draftsCommand } from './commands/drafts.js';
import { filesCommand } from './commands/files.js';
import { findCommand } from './commands/find.js';
import { findtimeCommand } from './commands/findtime.js';
import { foldersCommand } from './commands/folders.js';
import { forwardEventCommand } from './commands/forward-event.js';
import { loginCommand } from './commands/login.js';
import { mailCommand } from './commands/mail.js';
import { oofCommand } from './commands/oof.js';
import { plannerCommand } from './commands/planner.js';
import { respondCommand } from './commands/respond.js';
import { roomsCommand } from './commands/rooms.js';
import { rulesCommand } from './commands/rules.js';
import { scheduleCommand } from './commands/schedule.js';
import { sendCommand } from './commands/send.js';
import { serveCommand } from './commands/serve.js';
import { sharepointCommand } from './commands/sharepoint.js';
import { sitePagesCommand } from './commands/site-pages.js';
import { subscribeCommand } from './commands/subscribe.js';
import { subscriptionsCommand } from './commands/subscriptions.js';
import { suggestCommand } from './commands/suggest.js';
import { todoCommand } from './commands/todo.js';
import { updateEventCommand } from './commands/update-event.js';
import { verifyTokenCommand } from './commands/verify-token.js';
import { whoamiCommand } from './commands/whoami.js';

const program = new Command();

program.name('m365-agent-cli').description('CLI for Microsoft 365/EWS').version('0.1.0');

program.addCommand(whoamiCommand);
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
program.addCommand(updateEventCommand);
program.addCommand(mailCommand);
program.addCommand(foldersCommand);
program.addCommand(sendCommand);
program.addCommand(draftsCommand);
program.addCommand(filesCommand);
program.addCommand(forwardEventCommand);
program.addCommand(counterCommand);
program.addCommand(scheduleCommand);
program.addCommand(suggestCommand);
program.addCommand(subscribeCommand);
program.addCommand(subscriptionsCommand);
program.addCommand(serveCommand);
program.addCommand(roomsCommand);
program.addCommand(oofCommand);
program.addCommand(rulesCommand);
program.addCommand(delegatesCommand);
program.addCommand(todoCommand);

program.addCommand(plannerCommand);

program.addCommand(sharepointCommand);

program.parse();
