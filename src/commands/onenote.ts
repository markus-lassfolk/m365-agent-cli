import { readFile, writeFile } from 'node:fs/promises';
import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import {
  createOneNotePageFromHtml,
  getOneNotePage,
  getOneNotePageContentHtml,
  listNotebookSections,
  listOneNoteNotebooks,
  listSectionPages
} from '../lib/onenote-graph-client.js';
import { checkReadOnly } from '../lib/utils.js';

export const onenoteCommand = new Command('onenote').description(
  'OneNote notebooks, sections, and pages via Microsoft Graph (Notes.ReadWrite.All)'
);

onenoteCommand
  .command('notebooks')
  .description('List OneNote notebooks')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
  .action(async (opts: { json?: boolean; token?: string; identity?: string; user?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await listOneNoteNotebooks(auth.token!, opts.user);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else {
      for (const n of r.data) {
        const url = n.links?.oneNoteWebUrl?.href ?? '';
        console.log(`${n.displayName ?? '(notebook)'}\t${n.id}${url ? `\t${url}` : ''}`);
      }
    }
  });

onenoteCommand
  .command('sections')
  .description('List sections in a notebook')
  .argument('<notebookId>', 'Notebook id')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
  .action(
    async (
      notebookId: string,
      opts: { json?: boolean; token?: string; identity?: string; user?: string }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await listNotebookSections(auth.token!, notebookId, opts.user);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else {
        for (const s of r.data) {
          console.log(`${s.displayName ?? '(section)'}\t${s.id}`);
        }
      }
    }
  );

onenoteCommand
  .command('pages')
  .description('List pages in a section')
  .argument('<sectionId>', 'Section id')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
  .action(
    async (sectionId: string, opts: { json?: boolean; token?: string; identity?: string; user?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await listSectionPages(auth.token!, sectionId, opts.user);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else {
        for (const p of r.data) {
          const url = p.links?.oneNoteWebUrl?.href ?? '';
          console.log(`${p.title ?? '(untitled)'}\t${p.id}${url ? `\t${url}` : ''}`);
        }
      }
    }
  );

onenoteCommand
  .command('page')
  .description('Get page metadata (JSON)')
  .argument('<pageId>', 'Page id')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
  .action(
    async (pageId: string, opts: { token?: string; identity?: string; user?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getOneNotePage(auth.token!, pageId, opts.user);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(JSON.stringify(r.data, null, 2));
    }
  );

onenoteCommand
  .command('content')
  .description('Print page HTML content to stdout')
  .argument('<pageId>', 'Page id')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
  .action(async (pageId: string, opts: { token?: string; identity?: string; user?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await getOneNotePageContentHtml(auth.token!, pageId, opts.user);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    process.stdout.write(r.data);
    if (!r.data.endsWith('\n')) console.log('');
  });

onenoteCommand
  .command('export')
  .description('Download page HTML to a file')
  .argument('<pageId>', 'Page id')
  .requiredOption('-o, --output <path>', 'Output file path')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
  .action(
    async (
      pageId: string,
      opts: { output: string; token?: string; identity?: string; user?: string }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getOneNotePageContentHtml(auth.token!, pageId, opts.user);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      await writeFile(opts.output, r.data, 'utf-8');
      console.log(`Wrote ${opts.output}`);
    }
  );

onenoteCommand
  .command('create-page')
  .description('Create a page in a section from an HTML file (POST text/html)')
  .requiredOption('--section <sectionId>', 'Section id')
  .requiredOption('--file <path>', 'HTML file to upload')
  .option('--json', 'Echo created page as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
  .action(
    async (
      opts: {
        section: string;
        file: string;
        json?: boolean;
        token?: string;
        identity?: string;
        user?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const html = await readFile(opts.file, 'utf-8');
      const r = await createOneNotePageFromHtml(auth.token!, opts.section, html, opts.user);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else {
        console.log(`Created page: ${r.data.title ?? r.data.id} (${r.data.id})`);
        const url = r.data.links?.oneNoteWebUrl?.href;
        if (url) console.log(url);
      }
    }
  );
