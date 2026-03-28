import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import { searchPeople, searchUsers, searchGroups, expandGroup } from '../lib/graph-directory.js';

export const findCommand = new Command('find')
  .description('Search for people or groups in the directory')
  .argument('<query>', 'Search query (name, email, etc.)')
  .option('--people', 'Only search people/users')
  .option('--groups', 'Only search groups')
  .option('--expand', 'Expand group members if the query matches a group')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .action(
    async (
      query: string,
      options: {
        people?: boolean;
        groups?: boolean;
        expand?: boolean;
        json?: boolean;
        token?: string;
      }
    ) => {
      const authResult = await resolveGraphAuth({
        token: options.token
      });

      if (!authResult.success) {
        if (options.json) {
          console.log(JSON.stringify({ error: authResult.error }, null, 2));
        } else {
          console.error(`Error: ${authResult.error}`);
          console.error('\nCheck your .env file or run clippy auth.');
        }
        process.exit(1);
      }

      const token = authResult.token!;
      try {
        let results: any[] = [];
        let errors: string[] = [];

        const searchAll = !options.people && !options.groups;

        if (searchAll || options.people) {
          const peopleRes = await searchPeople(token, query);
          if (peopleRes.ok && peopleRes.data) {
            results.push(...peopleRes.data.map((p: any) => ({
              id: p.id,
              type: 'Person',
              name: p.displayName,
              email: p.userPrincipalName || (p.scoredEmailAddresses && p.scoredEmailAddresses[0]?.address) || (p.emailAddresses && p.emailAddresses[0]?.address),
              title: p.jobTitle || p.title,
              department: p.department
            })));
          } else if (peopleRes.error) {
            errors.push(`People search failed: ${peopleRes.error.message || peopleRes.error.status}`);
          }

          const usersRes = await searchUsers(token, query);
          if (usersRes.ok && usersRes.data) {
            for (const u of usersRes.data) {
              if (!results.find(r => r.id === u.id)) {
                results.push({
                  id: u.id,
                  type: 'Person',
                  name: u.displayName,
                  email: u.mail || u.userPrincipalName,
                  title: u.jobTitle,
                  department: (u as any).department
                });
              }
            }
          } else if (usersRes.error) {
            errors.push(`Users search failed: ${usersRes.error.message || usersRes.error.status}`);
          }
        }

        if (searchAll || options.groups) {
          const groupsRes = await searchGroups(token, query);
          if (groupsRes.ok && groupsRes.data) {
            for (const g of groupsRes.data) {
              let groupItem: any = {
                id: g.id,
                type: 'Group',
                name: g.displayName,
                email: g.mail,
                description: g.description
              };
              
              if (options.expand) {
                const membersRes = await expandGroup(token, g.id);
                if (membersRes.ok && membersRes.data) {
                  groupItem.members = membersRes.data.map((m: any) => ({
                    id: m.id,
                    name: m.displayName,
                    email: m.mail || m.userPrincipalName
                  }));
                }
              }
              results.push(groupItem);
            }
          } else if (groupsRes.error) {
            errors.push(`Groups search failed: ${groupsRes.error.message || groupsRes.error.status}`);
          }
        }

        if (options.json) {
          const output: any = { results };
          if (errors.length > 0) output.errors = errors;
          console.log(JSON.stringify(output, null, 2));
          return;
        }

        if (errors.length > 0) {
          console.error(`Warnings:`);
          errors.forEach(e => console.error(` - ${e}`));
        }

        if (results.length === 0) {
          console.log(`\nNo results found for "${query}"\n`);
          return;
        }

        console.log(`\nSearch results for "${query}":\n`);
        console.log('\u2500'.repeat(80));

        for (const res of results) {
          const isGroup = res.type === 'Group';
          const icon = isGroup ? '\u{1F465}' : '\u{1F464}';

          console.log(`\n  ${icon} ${res.name} [${res.type}]`);
          if (res.email) console.log(`     Email: ${res.email}`);
          if (res.title) console.log(`     Title: ${res.title}`);
          if (res.department) console.log(`     Dept:  ${res.department}`);
          if (res.description) console.log(`     Desc:  ${res.description}`);

          if (isGroup && res.members) {
            console.log(`     Members (${res.members.length}):`);
            for (const member of res.members) {
              console.log(`       - ${member.name} ${member.email ? `(${member.email})` : ''}`);
            }
          }
        }

        console.log(`\n${'\u2500'.repeat(80)}\n`);
      } catch (err) {
        const message = err instanceof Error ? err.message : 'Unknown error';
        if (options.json) {
          console.log(JSON.stringify({ error: message }, null, 2));
        } else {
          console.error(`Error: ${message}`);
        }
        process.exit(1);
      }
    }
  );
