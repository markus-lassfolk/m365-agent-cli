import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import { expandGroup, searchGroups, searchPeople, searchUsers } from '../lib/graph-directory.js';

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
          console.error('\nCheck your .env file or run m365-agent-cli auth.');
        }
        process.exit(1);
      }

      const token = authResult.token!;
      try {
        const results: any[] = [];
        const errors: string[] = [];

        const searchAll = !options.people && !options.groups;

        if (searchAll || options.people) {
          const peopleRes = await searchPeople(token, query);
          if (peopleRes.ok && peopleRes.data) {
            results.push(
              ...peopleRes.data.map((p) => ({
                id: p.id,
                type: 'Person',
                name: p.displayName,
                email: p.userPrincipalName || p.scoredEmailAddresses?.[0]?.address,
                title: p.jobTitle,
                department: p.department,
                userPrincipalName: p.userPrincipalName
              }))
            );
          } else if (peopleRes.error) {
            if (peopleRes.error.status === 403) {
              // Default auth scopes may not cover People API - suppress 403 unless --people was explicitly requested
              if (options.people) {
                errors.push(`People search failed: ${peopleRes.error.message}`);
              }
            } else {
              errors.push(`People search failed: ${peopleRes.error.message}`);
            }
          }

          const usersRes = await searchUsers(token, query);
          if (usersRes.ok && usersRes.data) {
            for (const u of usersRes.data) {
              const userEmail = u.mail || u.userPrincipalName;
              const userUpn = u.userPrincipalName;
              if (
                !results.find(
                  (r) =>
                    (r.userPrincipalName && userUpn && r.userPrincipalName.toLowerCase() === userUpn.toLowerCase()) ||
                    (r.email && userEmail && r.email.toLowerCase() === userEmail.toLowerCase())
                )
              ) {
                results.push({
                  id: u.id,
                  type: 'Person',
                  name: u.displayName,
                  email: userEmail,
                  title: u.jobTitle,
                  department: u.department,
                  userPrincipalName: u.userPrincipalName
                });
              }
            }
          } else if (usersRes.error) {
            if (usersRes.error.status === 403) {
              // Default auth scopes may not cover Users API - suppress 403 unless --people was explicitly requested
              if (options.people) {
                errors.push(`Users search failed: ${usersRes.error.message}`);
              }
            } else {
              errors.push(`Users search failed: ${usersRes.error.message}`);
            }
          }
        }

        if (searchAll || options.groups) {
          const groupsRes = await searchGroups(token, query);
          if (groupsRes.ok && groupsRes.data) {
            // Only expand members of the first group to avoid N+1 API calls
            const groupsToExpand = options.expand ? groupsRes.data.slice(0, 1) : [];
            for (const g of groupsRes.data) {
              const groupItem: any = {
                id: g.id,
                type: 'Group',
                name: g.displayName,
                email: g.mail,
                description: g.description
              };

              if (groupsToExpand.some((ge) => ge.id === g.id)) {
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
            if (groupsRes.error.status === 403) {
              // Default auth scopes may not cover Groups API - suppress 403 unless --groups was explicitly requested
              if (options.groups) {
                errors.push(`Groups search failed: ${groupsRes.error.message}`);
              }
            } else {
              errors.push(`Groups search failed: ${groupsRes.error.message}`);
            }
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
          for (const e of errors) console.error(` - ${e}`);
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
