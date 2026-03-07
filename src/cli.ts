#!/usr/bin/env bun
import { Command } from 'commander';
import { whoamiCommand } from './commands/whoami.js';
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

const program = new Command();

program
  .name('clippy')
  .description('CLI for Microsoft 365/EWS')
  .version('0.1.0');

program.addCommand(whoamiCommand);
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

program.parse();
