import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import { createListItem, getListItems, getLists, updateListItem } from '../lib/sharepoint-client.js';
import { checkReadOnly } from '../lib/utils.js';

export const sharepointCommand = new Command('sharepoint').description('Manage Microsoft SharePoint Lists').alias('sp');

sharepointCommand
  .command('lists')
  .description('List all SharePoint lists in a site')
  .requiredOption('--site-id <id>', 'SharePoint Site ID')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .action(async (opts: { siteId: string; json?: boolean; token?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error || 'Unknown error'}`);
      process.exit(1);
    }
    const res = await getLists(auth.token, opts.siteId);
    if (!res.ok) {
      console.error(`Error listing lists: ${res.error?.message || 'Unknown error'}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(res.data, null, 2));
      return;
    }
    if (!res.data || res.data.length === 0) {
      console.log('No lists found in this site.');
      return;
    }
    for (const list of res.data) {
      console.log(`${list.name} (${list.id})`);
      if (list.description) console.log(`  ${list.description}`);
    }
  });

sharepointCommand
  .command('items')
  .description('Get items from a SharePoint list')
  .requiredOption('--site-id <id>', 'SharePoint Site ID')
  .requiredOption('--list-id <id>', 'SharePoint List ID')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .action(async (opts: { siteId: string; listId: string; json?: boolean; token?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error || 'Unknown error'}`);
      process.exit(1);
    }
    const res = await getListItems(auth.token, opts.siteId, opts.listId);
    if (!res.ok) {
      console.error(`Error getting list items: ${res.error?.message || 'Unknown error'}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(res.data, null, 2));
      return;
    }
    if (!res.data || res.data.length === 0) {
      console.log('No items found in this list.');
      return;
    }
    for (const item of res.data) {
      console.log(`Item ID: ${item.id}`);
      if (item.fields) {
        for (const [key, val] of Object.entries(item.fields)) {
          if (!key.startsWith('@odata')) {
            console.log(`  ${key}: ${val}`);
          }
        }
      }
      console.log('---');
    }
  });

sharepointCommand
  .command('create-item')
  .description('Create an item in a SharePoint list')
  .requiredOption('--site-id <id>', 'SharePoint Site ID')
  .requiredOption('--list-id <id>', 'SharePoint List ID')
  .requiredOption('--fields <json>', 'JSON string of fields to set (e.g. \'{"Title": "My Item"}\')')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .action(
    async (opts: { siteId: string; listId: string; fields: string; json?: boolean; token?: string }, cmd: any) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error || 'Unknown error'}`);
        process.exit(1);
      }
      let parsedFields: Record<string, any>;
      try {
        parsedFields = JSON.parse(opts.fields);
      } catch (err: any) {
        console.error(`Error parsing fields JSON: ${err.message}`);
        process.exit(1);
      }
      if (typeof parsedFields !== 'object' || parsedFields === null || Array.isArray(parsedFields)) {
        console.error(
          'Error: --fields JSON must be an object (e.g. "{"Title": "New Title"}"), not an array or primitive.'
        );
        process.exit(1);
      }
      const res = await createListItem(auth.token, opts.siteId, opts.listId, parsedFields);
      if (!res.ok) {
        console.error(`Error creating list item: ${res.error?.message || 'Unknown error'}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(res.data, null, 2));
        return;
      }
      console.log(`Successfully created item ${res.data?.id}`);
    }
  );

sharepointCommand
  .command('update-item')
  .description('Update an item in a SharePoint list')
  .requiredOption('--site-id <id>', 'SharePoint Site ID')
  .requiredOption('--list-id <id>', 'SharePoint List ID')
  .requiredOption('--item-id <id>', 'SharePoint List Item ID')
  .requiredOption('--fields <json>', 'JSON string of fields to set (e.g. \'{"Title": "New Title"}\')')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .action(
    async (
      opts: {
        siteId: string;
        listId: string;
        itemId: string;
        fields: string;
        json?: boolean;
        token?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error || 'Unknown error'}`);
        process.exit(1);
      }
      let parsedFields: Record<string, any>;
      try {
        parsedFields = JSON.parse(opts.fields);
      } catch (err: any) {
        console.error(`Error parsing fields JSON: ${err.message}`);
        process.exit(1);
      }
      if (typeof parsedFields !== 'object' || parsedFields === null || Array.isArray(parsedFields)) {
        console.error(
          'Error: --fields JSON must be an object (e.g. "{"Title": "New Title"}"), not an array or primitive.'
        );
        process.exit(1);
      }
      const res = await updateListItem(auth.token, opts.siteId, opts.listId, opts.itemId, parsedFields);
      if (!res.ok) {
        console.error(`Error updating list item: ${res.error?.message || 'Unknown error'}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(res.data, null, 2));
        return;
      }
      console.log(`Successfully updated item ${opts.itemId}`);
    }
  );
