import { Command } from 'commander';
import { resolveAuth } from '../lib/auth.js';
import { resolveNames } from '../lib/ews-client.js';

export const findCommand = new Command('find')
  .description('Search for people or rooms')
  .argument('<query>', 'Search query (name, email, etc.)')
  .option('--rooms', 'Only show rooms')
  .option('--people', 'Only show people (exclude rooms)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .action(async (query: string, options: {
    rooms?: boolean;
    people?: boolean;
    json?: boolean;
    token?: string;
  }) => {
    const authResult = await resolveAuth({
      token: options.token,
    });

    if (!authResult.success) {
      if (options.json) {
        console.log(JSON.stringify({ error: authResult.error }, null, 2));
      } else {
        console.error(`Error: ${authResult.error}`);
        console.error('\nCheck your .env file for EWS_CLIENT_ID and EWS_REFRESH_TOKEN.');
      }
      process.exit(1);
    }

    try {
      const result = await resolveNames(authResult.token!, query);

      if (!result.ok || !result.data) {
        if (options.json) {
          console.log(JSON.stringify({ error: result.error?.message || 'Search failed' }, null, 2));
        } else {
          console.error(`Error: ${result.error?.message || 'Search failed'}`);
        }
        process.exit(1);
      }

      let results = result.data;

      // Filter by type if requested
      if (options.rooms) {
        results = results.filter(p => p.MailboxType === 'Room');
      } else if (options.people) {
        results = results.filter(p => p.MailboxType !== 'Room');
      }

      if (options.json) {
        console.log(JSON.stringify({
          results: results.map(p => ({
            name: p.DisplayName,
            email: p.EmailAddress,
            title: p.JobTitle,
            department: p.Department,
            type: p.MailboxType === 'Room' ? 'Room' : 'Person',
          })),
        }, null, 2));
        return;
      }

      if (results.length === 0) {
        console.log(`\nNo results found for "${query}"\n`);
        return;
      }

      console.log(`\nSearch results for "${query}":\n`);
      console.log('\u2500'.repeat(60));

      for (const person of results) {
        const isRoom = person.MailboxType === 'Room';
        const icon = isRoom ? '\u{1F4CD}' : '\u{1F464}';

        console.log(`\n  ${icon} ${person.DisplayName}`);
        if (person.EmailAddress) {
          console.log(`     ${person.EmailAddress}`);
        }
        if (!isRoom) {
          if (person.JobTitle) {
            console.log(`     ${person.JobTitle}`);
          }
          if (person.Department) {
            console.log(`     ${person.Department}`);
          }
        }
      }

      console.log('\n' + '\u2500'.repeat(60) + '\n');
    } catch (err) {
      const message = err instanceof Error ? err.message : 'Unknown error';
      if (options.json) {
        console.log(JSON.stringify({ error: message }, null, 2));
      } else {
        console.error(`Error: ${message}`);
      }
      process.exit(1);
    }
  });
