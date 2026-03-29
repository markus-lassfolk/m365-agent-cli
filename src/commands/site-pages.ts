import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import {
  getSitePage,
  listSitePages,
  publishSitePage,
  type SitePage,
  updateSitePage
} from '../lib/site-pages-client.js';
import { checkReadOnly } from '../lib/utils.js';

export const sitePagesCommand = new Command('pages').description('Manage SharePoint Site Pages');

sitePagesCommand
  .command('list <siteId>')
  .description('List site pages for a given site ID')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .action(async (siteId: string, options: { json?: boolean; token?: string }) => {
    const auth = await resolveGraphAuth({ token: options.token });
    if (!auth.success) {
      console.error(`Error: ${auth.error}`);
      process.exit(1);
    }

    const result = await listSitePages(auth.token!, siteId);
    if (!result.ok || !result.data) {
      console.error(`Error: ${result.error?.message || 'Request failed'}`);
      process.exit(1);
    }

    if (options.json) {
      console.log(JSON.stringify(result.data, null, 2));
      return;
    }

    if (!result.data || result.data.length === 0) {
      console.log('No pages found.');
      return;
    }

    for (const page of result.data) {
      const state = page.publishingState
        ? `${page.publishingState.level} (v${page.publishingState.versionId})`
        : 'Unknown';
      console.log(`- ${page.name || page.title || page.id} (${page.id}) - State: ${state}`);
    }
  });

sitePagesCommand
  .command('get <siteId> <pageId>')
  .description('Get a site page by ID')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .action(async (siteId: string, pageId: string, options: { json?: boolean; token?: string }) => {
    const auth = await resolveGraphAuth({ token: options.token });
    if (!auth.success) {
      console.error(`Error: ${auth.error}`);
      process.exit(1);
    }

    const result = await getSitePage(auth.token!, siteId, pageId);
    if (!result.ok || !result.data) {
      console.error(`Error: ${result.error?.message || 'Request failed'}`);
      process.exit(1);
    }

    if (options.json) {
      console.log(JSON.stringify(result.data, null, 2));
      return;
    }

    console.log(`ID: ${result.data.id}`);
    console.log(`Name: ${result.data.name || '-'}`);
    console.log(`Title: ${result.data.title || '-'}`);
    console.log(`Web URL: ${result.data.webUrl || '-'}`);
    if (result.data.publishingState) {
      console.log(`State: ${result.data.publishingState.level} (v${result.data.publishingState.versionId})`);
    }
  });

sitePagesCommand
  .command('update <siteId> <pageId>')
  .description('Update a site page')
  .option('--title <title>', 'New title')
  .option('--name <name>', 'New name')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .action(
    async (
      siteId: string,
      pageId: string,
      options: { title?: string; name?: string; json?: boolean; token?: string },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: options.token });
      if (!auth.success) {
        console.error(`Error: ${auth.error}`);
        process.exit(1);
      }

      const payload: Partial<SitePage> = {};
      if (options.title) payload.title = options.title;
      if (options.name) payload.name = options.name;

      if (Object.keys(payload).length === 0) {
        console.error('Error: Please provide at least one field to update (--title or --name)');
        process.exit(1);
      }

      const result = await updateSitePage(auth.token!, siteId, pageId, payload);
      if (!result.ok || !result.data) {
        console.error(`Error: ${result.error?.message || 'Request failed'}`);
        process.exit(1);
      }

      if (options.json) {
        console.log(JSON.stringify(result.data, null, 2));
        return;
      }

      console.log(`✓ Updated page ${pageId}`);
    }
  );

sitePagesCommand
  .command('publish <siteId> <pageId>')
  .description('Publish a site page')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific Graph token')
  .action(async (siteId: string, pageId: string, options: { json?: boolean; token?: string }, cmd: any) => {
    checkReadOnly(cmd);
    const auth = await resolveGraphAuth({ token: options.token });
    if (!auth.success) {
      console.error(`Error: ${auth.error}`);
      process.exit(1);
    }

    const result = await publishSitePage(auth.token!, siteId, pageId);
    if (!result.ok) {
      console.error(`Error: ${result.error?.message || 'Request failed'}`);
      process.exit(1);
    }

    if (options.json) {
      console.log(JSON.stringify({ ok: true }, null, 2));
      return;
    }

    console.log(`✓ Published page ${pageId}`);
  });
