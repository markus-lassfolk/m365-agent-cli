#!/usr/bin/env bun
import { Command } from 'commander';
import { whoamiCommand } from './commands/whoami.js';
import { autoReplyCommand } from './commands/auto-reply.js';
import { calendarCommand } from './commands/calendar.js';
import { findtimeCommand } from './commands/findtime.js';
import { respondCommand } from './commands/respond.js';
import { createEventCommand } from './commands/create-event.js';
import { deleteEventCommand } from './commands/delete-event.js';
import { findCommand } from './commands/find.js';
import { updateEventCommand } from './commands/update-event.js';
import { mailCommand } from './commands/mail.js';
import { foldersCommand } from './commands/folders.js';
import { sendCommand } from './commands/send.js';
import { draftsCommand } from './commands/drafts.js';
import { filesCommand } from './commands/files.js';
import { forwardEventCommand } from './commands/forward-event.js';
import { counterCommand } from './commands/counter.js';
import { scheduleCommand } from './commands/schedule.js';
import { suggestCommand } from './commands/suggest.js';
import { subscribeCommand } from './commands/subscribe.js';
import { subscriptionsCommand } from './commands/subscriptions.js';
import { serveCommand } from './commands/serve.js';
import { roomsCommand } from './commands/rooms.js';
import { oofCommand } from './commands/oof.js';
import { delegatesCommand } from './commands/delegates.js';
import { todoCommand } from './commands/todo.js';

const program = new Command();

program.name('clippy').description('CLI for Microsoft 365/EWS').version('0.1.0');

program.addCommand(whoamiCommand);
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
program.addCommand(delegatesCommand);
program.addCommand(todoCommand);

program.parse();
