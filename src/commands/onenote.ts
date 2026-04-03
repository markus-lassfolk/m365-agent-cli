import { readFile, writeFile } from 'node:fs/promises';
import { Command } from 'commander';
import { requireGraphAuth } from '../lib/graph-auth.js';
import {
  copyOneNotePageToSection,
  copyOneNoteSectionToNotebook,
  copyOneNoteSectionToSectionGroup,
  createOneNoteNotebook,
  createOneNotePageFromHtml,
  createSectionGroupInNotebook,
  createSectionInNotebook,
  createSectionInSectionGroup,
  deleteOneNoteNotebook,
  deleteOneNotePage,
  deleteOneNoteSection,
  deleteOneNoteSectionGroup,
  getOneNoteNotebook,
  getOneNoteNotebookFromWebUrl,
  getOneNoteOperation,
  getOneNotePage,
  getOneNotePageContentHtml,
  getOneNotePagePreview,
  getOneNoteSection,
  getOneNoteSectionGroup,
  listAllOneNotePages,
  listNotebookSectionGroups,
  listNotebookSections,
  listOneNoteNotebooks,
  listSectionPages,
  listSectionsInSectionGroup,
  type OneNoteGraphScope,
  updateOneNoteNotebook,
  updateOneNotePageContent,
  updateOneNoteSection,
  updateOneNoteSectionGroup
} from '../lib/onenote-graph-client.js';
import { checkReadOnly } from '../lib/utils.js';

/** Resolves `--user` vs `--group` / `--site` OneNote roots (mutually exclusive). */
function parseOneNoteRoot(opts: { user?: string; group?: string; site?: string }): {
  user?: string;
  scope?: OneNoteGraphScope;
} {
  const site = opts.site?.trim();
  const group = opts.group?.trim();
  const user = opts.user?.trim();
  if (site && group) {
    console.error('Error: use only one of --site or --group');
    process.exit(1);
  }
  if ((site || group) && user) {
    console.error('Error: do not combine --user with --group or --site (pick one OneNote root)');
    process.exit(1);
  }
  if (site) return { scope: { siteId: site } };
  if (group) return { scope: { groupId: group } };
  return { user };
}

/** Options repeated on OneNote subcommands (group/site roots). */
const optGroupSite = [
  ['--group <id>', 'OneNote root: /groups/{id}/onenote'],
  ['--site <id>', 'OneNote root: /sites/{id}/onenote']
] as const;

function addOneNoteRootOptions(cmd: Command): Command {
  for (const [flags, desc] of optGroupSite) {
    cmd.option(flags, desc);
  }
  return cmd;
}

export const onenoteCommand = new Command('onenote').description(
  'OneNote via Microsoft Graph (`Notes.ReadWrite.All`): notebooks (incl. resolve by web URL), section groups, sections (copy to notebook / section group), pages, copy, operations'
);

// ─── Legacy / primary list commands (unchanged names) ─────────────────────

addOneNoteRootOptions(
  onenoteCommand
    .command('notebooks')
    .description('List OneNote notebooks')
    .option('--json', 'Output as JSON')
    .option('--token <token>', 'Use a specific token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .option('--user <email>', 'Target user (Graph delegation)')
).action(
  async (opts: { json?: boolean; token?: string; identity?: string; user?: string; group?: string; site?: string }) => {
    const token = await requireGraphAuth(opts);
    const { user, scope } = parseOneNoteRoot(opts);
    const r = await listOneNoteNotebooks(token, user, scope);
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
  }
);

addOneNoteRootOptions(
  onenoteCommand
    .command('sections')
    .description('List sections in a notebook')
    .argument('<notebookId>', 'Notebook id')
    .option('--json', 'Output as JSON')
    .option('--token <token>', 'Use a specific token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .option('--user <email>', 'Target user (Graph delegation)')
).action(
  async (
    notebookId: string,
    opts: { json?: boolean; token?: string; identity?: string; user?: string; group?: string; site?: string }
  ) => {
    const token = await requireGraphAuth(opts);
    const { user, scope } = parseOneNoteRoot(opts);
    const r = await listNotebookSections(token, notebookId, user, scope);
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

addOneNoteRootOptions(
  onenoteCommand
    .command('pages')
    .description('List pages in a section')
    .argument('<sectionId>', 'Section id')
    .option('--json', 'Output as JSON')
    .option('--token <token>', 'Use a specific token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .option('--user <email>', 'Target user (Graph delegation)')
).action(
  async (
    sectionId: string,
    opts: { json?: boolean; token?: string; identity?: string; user?: string; group?: string; site?: string }
  ) => {
    const token = await requireGraphAuth(opts);
    const { user, scope } = parseOneNoteRoot(opts);
    const r = await listSectionPages(token, sectionId, user, scope);
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

addOneNoteRootOptions(
  onenoteCommand
    .command('list-pages')
    .description('List pages across notebooks (GET …/onenote/pages; optional OData query)')
    .option('--query <odata>', 'Query string without leading ? (e.g. $top=50&$orderby=lastModifiedTime desc)')
    .option('--json', 'Output as JSON')
    .option('--token <token>', 'Use a specific token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .option('--user <email>', 'Target user (Graph delegation)')
).action(
  async (opts: {
    query?: string;
    json?: boolean;
    token?: string;
    identity?: string;
    user?: string;
    group?: string;
    site?: string;
  }) => {
    const token = await requireGraphAuth(opts);
    const { user, scope } = parseOneNoteRoot(opts);
    const r = await listAllOneNotePages(token, user, opts.query, scope);
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

addOneNoteRootOptions(
  onenoteCommand
    .command('page-preview')
    .description('Get page preview text snippet (GET …/pages/{id}/preview)')
    .argument('<pageId>', 'Page id')
    .option('--json', 'Output as JSON')
    .option('--token <token>', 'Use a specific token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .option('--user <email>', 'Target user (Graph delegation)')
).action(
  async (
    pageId: string,
    opts: { json?: boolean; token?: string; identity?: string; user?: string; group?: string; site?: string }
  ) => {
    const token = await requireGraphAuth(opts);
    const { user, scope } = parseOneNoteRoot(opts);
    const r = await getOneNotePagePreview(token, pageId, user, scope);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else console.log(r.data.previewText ?? '');
  }
);

addOneNoteRootOptions(
  onenoteCommand
    .command('page')
    .description('Get page metadata (JSON)')
    .argument('<pageId>', 'Page id')
    .option('--token <token>', 'Use a specific token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .option('--user <email>', 'Target user (Graph delegation)')
).action(
  async (pageId: string, opts: { token?: string; identity?: string; user?: string; group?: string; site?: string }) => {
    const token = await requireGraphAuth(opts);
    const { user, scope } = parseOneNoteRoot(opts);
    const r = await getOneNotePage(token, pageId, user, scope);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log(JSON.stringify(r.data, null, 2));
  }
);

addOneNoteRootOptions(
  onenoteCommand
    .command('content')
    .description('Print page HTML content to stdout')
    .argument('<pageId>', 'Page id')
    .option('--token <token>', 'Use a specific token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .option('--user <email>', 'Target user (Graph delegation)')
).action(
  async (pageId: string, opts: { token?: string; identity?: string; user?: string; group?: string; site?: string }) => {
    const token = await requireGraphAuth(opts);
    const { user, scope } = parseOneNoteRoot(opts);
    const r = await getOneNotePageContentHtml(token, pageId, user, scope);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    process.stdout.write(r.data);
    if (!r.data.endsWith('\n')) console.log('');
  }
);

addOneNoteRootOptions(
  onenoteCommand
    .command('export')
    .description('Download page HTML to a file')
    .argument('<pageId>', 'Page id')
    .requiredOption('-o, --output <path>', 'Output file path')
    .option('--token <token>', 'Use a specific token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .option('--user <email>', 'Target user (Graph delegation)')
).action(
  async (
    pageId: string,
    opts: { output: string; token?: string; identity?: string; user?: string; group?: string; site?: string }
  ) => {
    const token = await requireGraphAuth(opts);
    const { user, scope } = parseOneNoteRoot(opts);
    const r = await getOneNotePageContentHtml(token, pageId, user, scope);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    await writeFile(opts.output, r.data, 'utf-8');
    console.log(`Wrote ${opts.output}`);
  }
);

addOneNoteRootOptions(
  onenoteCommand
    .command('create-page')
    .description('Create a page in a section from an HTML file (POST text/html)')
    .requiredOption('--section <sectionId>', 'Section id')
    .requiredOption('--file <path>', 'HTML file to upload')
    .option('--json', 'Echo created page as JSON')
    .option('--token <token>', 'Use a specific token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .option('--user <email>', 'Target user (Graph delegation)')
).action(
  async (
    opts: {
      section: string;
      file: string;
      json?: boolean;
      token?: string;
      identity?: string;
      user?: string;
      group?: string;
      site?: string;
    },
    cmd: any
  ) => {
    checkReadOnly(cmd);
    const token = await requireGraphAuth(opts);
    const { user, scope } = parseOneNoteRoot(opts);
    const html = await readFile(opts.file, 'utf-8');
    const r = await createOneNotePageFromHtml(token, opts.section, html, user, scope);
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

// ─── Notebook CRUD ──────────────────────────────────────────────────────────

const notebookCmd = new Command('notebook').description('Notebook get, create, update, delete');

addOneNoteRootOptions(
  notebookCmd
    .command('from-web-url')
    .description('Resolve a notebook by OneNote web URL (Graph notebook getNotebookFromWebUrl)')
    .requiredOption('--url <webUrl>', 'Notebook web URL (https://… or onenote:…)')
    .option('--json', 'Output notebook JSON')
    .option('--token <token>', 'Use a specific token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .option('--user <email>', 'Target user')
).action(
  async (opts: {
    url: string;
    json?: boolean;
    token?: string;
    identity?: string;
    user?: string;
    group?: string;
    site?: string;
  }) => {
    const token = await requireGraphAuth(opts);
    const { user, scope } = parseOneNoteRoot(opts);
    const r = await getOneNoteNotebookFromWebUrl(token, opts.url, user, scope);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else {
      const url = r.data.links?.oneNoteWebUrl?.href ?? '';
      console.log(`${r.data.displayName ?? '(notebook)'}\t${r.data.id}${url ? `\t${url}` : ''}`);
    }
  }
);

addOneNoteRootOptions(
  notebookCmd
    .command('list')
    .description('List notebooks (same as `onenote notebooks`)')
    .option('--json', 'Output as JSON')
    .option('--token <token>', 'Use a specific token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .option('--user <email>', 'Target user')
).action(
  async (opts: { json?: boolean; token?: string; identity?: string; user?: string; group?: string; site?: string }) => {
    const token = await requireGraphAuth(opts);
    const { user, scope } = parseOneNoteRoot(opts);
    const r = await listOneNoteNotebooks(token, user, scope);
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
  }
);

addOneNoteRootOptions(
  notebookCmd
    .command('get')
    .description('Get one notebook by id')
    .argument('<notebookId>', 'Notebook id')
    .option('--json', 'Output as JSON')
    .option('--token <token>', 'Use a specific token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .option('--user <email>', 'Target user')
).action(
  async (
    notebookId: string,
    opts: { json?: boolean; token?: string; identity?: string; user?: string; group?: string; site?: string }
  ) => {
    const token = await requireGraphAuth(opts);
    const { user, scope } = parseOneNoteRoot(opts);
    const r = await getOneNoteNotebook(token, notebookId, user, scope);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else {
      const url = r.data.links?.oneNoteWebUrl?.href ?? '';
      console.log(`${r.data.displayName ?? '(notebook)'}\t${r.data.id}${url ? `\t${url}` : ''}`);
    }
  }
);

addOneNoteRootOptions(
  notebookCmd
    .command('create')
    .description('Create a notebook')
    .requiredOption('--name <displayName>', 'Display name')
    .option('--json', 'Echo created notebook as JSON')
    .option('--token <token>', 'Use a specific token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .option('--user <email>', 'Target user')
).action(
  async (
    opts: {
      name: string;
      json?: boolean;
      token?: string;
      identity?: string;
      user?: string;
      group?: string;
      site?: string;
    },
    cmd: any
  ) => {
    checkReadOnly(cmd);
    const token = await requireGraphAuth(opts);
    const { user, scope } = parseOneNoteRoot(opts);
    const r = await createOneNoteNotebook(token, opts.name, user, scope);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else console.log(`Created notebook: ${r.data.displayName ?? r.data.id} (${r.data.id})`);
  }
);

addOneNoteRootOptions(
  notebookCmd
    .command('update')
    .description('PATCH a notebook (e.g. rename)')
    .argument('<notebookId>', 'Notebook id')
    .requiredOption('--json-file <path>', 'JSON patch body')
    .option('--json', 'Echo result as JSON')
    .option('--token <token>', 'Use a specific token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .option('--user <email>', 'Target user')
).action(
  async (
    notebookId: string,
    opts: {
      jsonFile: string;
      json?: boolean;
      token?: string;
      identity?: string;
      user?: string;
      group?: string;
      site?: string;
    },
    cmd: any
  ) => {
    checkReadOnly(cmd);
    const token = await requireGraphAuth(opts);
    const { user, scope } = parseOneNoteRoot(opts);
    const patch = JSON.parse(await readFile(opts.jsonFile, 'utf-8')) as Record<string, unknown>;
    const r = await updateOneNoteNotebook(token, notebookId, patch, user, scope);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) console.log(JSON.stringify(r.data ?? null, null, 2));
    else if (r.data) console.log(`Updated notebook: ${r.data.id}`);
    else console.log('Notebook updated.');
  }
);

addOneNoteRootOptions(
  notebookCmd
    .command('delete')
    .description('Delete a notebook')
    .argument('<notebookId>', 'Notebook id')
    .option('--confirm', 'Confirm delete')
    .option('--token <token>', 'Use a specific token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .option('--user <email>', 'Target user')
).action(
  async (
    notebookId: string,
    opts: {
      confirm?: boolean;
      token?: string;
      identity?: string;
      user?: string;
      group?: string;
      site?: string;
    },
    cmd: any
  ) => {
    checkReadOnly(cmd);
    if (!opts.confirm) {
      console.error('Refusing to delete without --confirm');
      process.exit(1);
    }
    const token = await requireGraphAuth(opts);
    const { user, scope } = parseOneNoteRoot(opts);
    const r = await deleteOneNoteNotebook(token, notebookId, user, scope);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log('Notebook deleted.');
  }
);

onenoteCommand.addCommand(notebookCmd);

// ─── Section group CRUD ─────────────────────────────────────────────────────

const sectionGroupCmd = new Command('section-group').description('Section groups under a notebook');

addOneNoteRootOptions(
  sectionGroupCmd
    .command('list')
    .description('List section groups in a notebook')
    .requiredOption('--notebook <notebookId>', 'Notebook id')
    .option('--json', 'Output as JSON')
    .option('--token <token>', 'Use a specific token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .option('--user <email>', 'Target user')
).action(
  async (opts: {
    notebook: string;
    json?: boolean;
    token?: string;
    identity?: string;
    user?: string;
    group?: string;
    site?: string;
  }) => {
    const token = await requireGraphAuth(opts);
    const { user, scope } = parseOneNoteRoot(opts);
    const r = await listNotebookSectionGroups(token, opts.notebook, user, scope);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else {
      for (const g of r.data) {
        console.log(`${g.displayName ?? '(group)'}\t${g.id}`);
      }
    }
  }
);

addOneNoteRootOptions(
  sectionGroupCmd
    .command('get')
    .description('Get a section group by id')
    .argument('<sectionGroupId>', 'Section group id')
    .option('--json', 'Output as JSON')
    .option('--token <token>', 'Use a specific token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .option('--user <email>', 'Target user')
).action(
  async (
    sectionGroupId: string,
    opts: { json?: boolean; token?: string; identity?: string; user?: string; group?: string; site?: string }
  ) => {
    const token = await requireGraphAuth(opts);
    const { user, scope } = parseOneNoteRoot(opts);
    const r = await getOneNoteSectionGroup(token, sectionGroupId, user, scope);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else console.log(`${r.data.displayName ?? '(group)'}\t${r.data.id}`);
  }
);

addOneNoteRootOptions(
  sectionGroupCmd
    .command('create')
    .description('Create a section group in a notebook')
    .requiredOption('--notebook <notebookId>', 'Notebook id')
    .requiredOption('--name <displayName>', 'Display name')
    .option('--json', 'Echo as JSON')
    .option('--token <token>', 'Use a specific token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .option('--user <email>', 'Target user')
).action(
  async (
    opts: {
      notebook: string;
      name: string;
      json?: boolean;
      token?: string;
      identity?: string;
      user?: string;
      group?: string;
      site?: string;
    },
    cmd: any
  ) => {
    checkReadOnly(cmd);
    const token = await requireGraphAuth(opts);
    const { user, scope } = parseOneNoteRoot(opts);
    const r = await createSectionGroupInNotebook(token, opts.notebook, opts.name, user, scope);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else console.log(`Created section group: ${r.data.displayName ?? r.data.id} (${r.data.id})`);
  }
);

addOneNoteRootOptions(
  sectionGroupCmd
    .command('update')
    .description('PATCH a section group')
    .argument('<sectionGroupId>', 'Section group id')
    .requiredOption('--json-file <path>', 'JSON patch')
    .option('--json', 'Echo as JSON')
    .option('--token <token>', 'Use a specific token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .option('--user <email>', 'Target user')
).action(
  async (
    sectionGroupId: string,
    opts: {
      jsonFile: string;
      json?: boolean;
      token?: string;
      identity?: string;
      user?: string;
      group?: string;
      site?: string;
    },
    cmd: any
  ) => {
    checkReadOnly(cmd);
    const token = await requireGraphAuth(opts);
    const { user, scope } = parseOneNoteRoot(opts);
    const patch = JSON.parse(await readFile(opts.jsonFile, 'utf-8')) as Record<string, unknown>;
    const r = await updateOneNoteSectionGroup(token, sectionGroupId, patch, user, scope);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) console.log(JSON.stringify(r.data ?? null, null, 2));
    else if (r.data) console.log(`Updated section group: ${r.data.id}`);
    else console.log('Section group updated.');
  }
);

addOneNoteRootOptions(
  sectionGroupCmd
    .command('delete')
    .description('Delete a section group')
    .argument('<sectionGroupId>', 'Section group id')
    .option('--confirm', 'Confirm')
    .option('--token <token>', 'Use a specific token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .option('--user <email>', 'Target user')
).action(
  async (
    sectionGroupId: string,
    opts: { confirm?: boolean; token?: string; identity?: string; user?: string; group?: string; site?: string },
    cmd: any
  ) => {
    checkReadOnly(cmd);
    if (!opts.confirm) {
      console.error('Refusing to delete without --confirm');
      process.exit(1);
    }
    const token = await requireGraphAuth(opts);
    const { user, scope } = parseOneNoteRoot(opts);
    const r = await deleteOneNoteSectionGroup(token, sectionGroupId, user, scope);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log('Section group deleted.');
  }
);

onenoteCommand.addCommand(sectionGroupCmd);

// ─── Section CRUD ───────────────────────────────────────────────────────────

const sectionCmd = new Command('section').description('Sections (under notebook or section group)');

addOneNoteRootOptions(
  sectionCmd
    .command('list')
    .description('List sections — use --notebook or --section-group')
    .option('--notebook <notebookId>', 'List sections in notebook')
    .option('--section-group <sectionGroupId>', 'List sections in section group')
    .option('--json', 'Output as JSON')
    .option('--token <token>', 'Use a specific token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .option('--user <email>', 'Target user')
).action(
  async (opts: {
    notebook?: string;
    sectionGroup?: string;
    json?: boolean;
    token?: string;
    identity?: string;
    user?: string;
    group?: string;
    site?: string;
  }) => {
    const token = await requireGraphAuth(opts);
    const { user, scope } = parseOneNoteRoot(opts);
    const nb = opts.notebook?.trim();
    const sg = opts.sectionGroup?.trim();
    if ((nb && sg) || (!nb && !sg)) {
      console.error('Error: specify exactly one of --notebook or --section-group');
      process.exit(1);
    }
    const r = nb
      ? await listNotebookSections(token, nb, user, scope)
      : await listSectionsInSectionGroup(token, sg!, user, scope);
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

addOneNoteRootOptions(
  sectionCmd
    .command('get')
    .description('Get a section by id')
    .argument('<sectionId>', 'Section id')
    .option('--json', 'Output as JSON')
    .option('--token <token>', 'Use a specific token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .option('--user <email>', 'Target user')
).action(
  async (
    sectionId: string,
    opts: { json?: boolean; token?: string; identity?: string; user?: string; group?: string; site?: string }
  ) => {
    const token = await requireGraphAuth(opts);
    const { user, scope } = parseOneNoteRoot(opts);
    const r = await getOneNoteSection(token, sectionId, user, scope);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else console.log(`${r.data.displayName ?? '(section)'}\t${r.data.id}`);
  }
);

addOneNoteRootOptions(
  sectionCmd
    .command('create')
    .description('Create a section (under --notebook or --section-group)')
    .requiredOption('--name <displayName>', 'Display name')
    .option('--notebook <notebookId>', 'Create in notebook')
    .option('--section-group <sectionGroupId>', 'Create in section group')
    .option('--json', 'Echo as JSON')
    .option('--token <token>', 'Use a specific token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .option('--user <email>', 'Target user')
).action(
  async (
    opts: {
      name: string;
      notebook?: string;
      sectionGroup?: string;
      json?: boolean;
      token?: string;
      identity?: string;
      user?: string;
      group?: string;
      site?: string;
    },
    cmd: any
  ) => {
    checkReadOnly(cmd);
    const nb = opts.notebook?.trim();
    const sg = opts.sectionGroup?.trim();
    if ((nb && sg) || (!nb && !sg)) {
      console.error('Error: specify exactly one of --notebook or --section-group');
      process.exit(1);
    }
    const token = await requireGraphAuth(opts);
    const { user, scope } = parseOneNoteRoot(opts);
    const r = nb
      ? await createSectionInNotebook(token, nb, opts.name, user, scope)
      : await createSectionInSectionGroup(token, sg!, opts.name, user, scope);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else console.log(`Created section: ${r.data.displayName ?? r.data.id} (${r.data.id})`);
  }
);

addOneNoteRootOptions(
  sectionCmd
    .command('update')
    .description('PATCH a section')
    .argument('<sectionId>', 'Section id')
    .requiredOption('--json-file <path>', 'JSON patch')
    .option('--json', 'Echo as JSON')
    .option('--token <token>', 'Use a specific token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .option('--user <email>', 'Target user')
).action(
  async (
    sectionId: string,
    opts: {
      jsonFile: string;
      json?: boolean;
      token?: string;
      identity?: string;
      user?: string;
      group?: string;
      site?: string;
    },
    cmd: any
  ) => {
    checkReadOnly(cmd);
    const token = await requireGraphAuth(opts);
    const { user, scope } = parseOneNoteRoot(opts);
    const patch = JSON.parse(await readFile(opts.jsonFile, 'utf-8')) as Record<string, unknown>;
    const r = await updateOneNoteSection(token, sectionId, patch, user, scope);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) console.log(JSON.stringify(r.data ?? null, null, 2));
    else if (r.data) console.log(`Updated section: ${r.data.id}`);
    else console.log('Section updated.');
  }
);

addOneNoteRootOptions(
  sectionCmd
    .command('delete')
    .description('Delete a section')
    .argument('<sectionId>', 'Section id')
    .option('--confirm', 'Confirm')
    .option('--token <token>', 'Use a specific token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .option('--user <email>', 'Target user')
).action(
  async (
    sectionId: string,
    opts: { confirm?: boolean; token?: string; identity?: string; user?: string; group?: string; site?: string },
    cmd: any
  ) => {
    checkReadOnly(cmd);
    if (!opts.confirm) {
      console.error('Refusing to delete without --confirm');
      process.exit(1);
    }
    const token = await requireGraphAuth(opts);
    const { user, scope } = parseOneNoteRoot(opts);
    const r = await deleteOneNoteSection(token, sectionId, user, scope);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log('Section deleted.');
  }
);

addOneNoteRootOptions(
  sectionCmd
    .command('copy-to-notebook')
    .description('Copy a section into another notebook (async — poll Operation-Location with `onenote operation`)')
    .argument('<sectionId>', 'Source section id')
    .requiredOption('--notebook <notebookId>', 'Destination notebook id')
    .option('--group-id <id>', 'Request body `groupId` when the destination is a Microsoft 365 group notebook')
    .option('--rename-as <name>', 'Optional name for the copied section')
    .option('--json', 'Print operation status JSON')
    .option('--token <token>', 'Use a specific token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .option('--user <email>', 'Target user')
).action(
  async (
    sectionId: string,
    opts: {
      notebook: string;
      groupId?: string;
      renameAs?: string;
      json?: boolean;
      token?: string;
      identity?: string;
      user?: string;
      group?: string;
      site?: string;
    },
    cmd: any
  ) => {
    checkReadOnly(cmd);
    const token = await requireGraphAuth(opts);
    const { user, scope } = parseOneNoteRoot(opts);
    const r = await copyOneNoteSectionToNotebook(token, sectionId, opts.notebook, user, scope, {
      copyToNotebookGroupId: opts.groupId,
      renameAs: opts.renameAs
    });
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(r.data, null, 2));
    } else {
      console.log(`Copy accepted (HTTP ${r.data.status}).`);
      if (r.data.operationLocation) {
        console.log(`Operation-Location: ${r.data.operationLocation}`);
        console.log('Poll with: m365-agent-cli onenote operation "<url>"');
      }
    }
  }
);

addOneNoteRootOptions(
  sectionCmd
    .command('copy-to-section-group')
    .description('Copy a section into a section group (async — poll Operation-Location with `onenote operation`)')
    .argument('<sectionId>', 'Source section id')
    .requiredOption('--section-group <sectionGroupId>', 'Destination section group id')
    .option('--group-id <id>', 'Request body `groupId` when the destination is a Microsoft 365 group notebook')
    .option('--rename-as <name>', 'Optional name for the copied section')
    .option('--json', 'Print operation status JSON')
    .option('--token <token>', 'Use a specific token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .option('--user <email>', 'Target user')
).action(
  async (
    sectionId: string,
    opts: {
      sectionGroup: string;
      groupId?: string;
      renameAs?: string;
      json?: boolean;
      token?: string;
      identity?: string;
      user?: string;
      group?: string;
      site?: string;
    },
    cmd: any
  ) => {
    checkReadOnly(cmd);
    const token = await requireGraphAuth(opts);
    const { user, scope } = parseOneNoteRoot(opts);
    const r = await copyOneNoteSectionToSectionGroup(token, sectionId, opts.sectionGroup, user, scope, {
      copyToGroupId: opts.groupId,
      renameAs: opts.renameAs
    });
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(r.data, null, 2));
    } else {
      console.log(`Copy accepted (HTTP ${r.data.status}).`);
      if (r.data.operationLocation) {
        console.log(`Operation-Location: ${r.data.operationLocation}`);
        console.log('Poll with: m365-agent-cli onenote operation "<url>"');
      }
    }
  }
);

onenoteCommand.addCommand(sectionCmd);

// ─── Page delete / patch content / copy / operation poll ────────────────────

addOneNoteRootOptions(
  onenoteCommand
    .command('delete-page')
    .description('Delete a page')
    .argument('<pageId>', 'Page id')
    .option('--confirm', 'Confirm delete')
    .option('--token <token>', 'Use a specific token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .option('--user <email>', 'Target user')
).action(
  async (
    pageId: string,
    opts: { confirm?: boolean; token?: string; identity?: string; user?: string; group?: string; site?: string },
    cmd: any
  ) => {
    checkReadOnly(cmd);
    if (!opts.confirm) {
      console.error('Refusing to delete without --confirm');
      process.exit(1);
    }
    const token = await requireGraphAuth(opts);
    const { user, scope } = parseOneNoteRoot(opts);
    const r = await deleteOneNotePage(token, pageId, user, scope);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log('Page deleted.');
  }
);

addOneNoteRootOptions(
  onenoteCommand
    .command('patch-page-content')
    .description('PATCH page content (Graph JSON patch commands; see MS docs for page-update)')
    .argument('<pageId>', 'Page id')
    .requiredOption('--json-file <path>', 'JSON array of patch commands')
    .option('--token <token>', 'Use a specific token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .option('--user <email>', 'Target user')
).action(
  async (
    pageId: string,
    opts: { jsonFile: string; token?: string; identity?: string; user?: string; group?: string; site?: string },
    cmd: any
  ) => {
    checkReadOnly(cmd);
    const token = await requireGraphAuth(opts);
    const { user, scope } = parseOneNoteRoot(opts);
    const commands = JSON.parse(await readFile(opts.jsonFile, 'utf-8'));
    const r = await updateOneNotePageContent(token, pageId, commands, user, scope);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log('Page content updated.');
  }
);

addOneNoteRootOptions(
  onenoteCommand
    .command('copy-page')
    .description('Copy a page to another section (async — use Operation-Location with `onenote operation`)')
    .argument('<pageId>', 'Source page id')
    .requiredOption('--section <sectionId>', 'Destination section id')
    .option(
      '--group-id <id>',
      'Request body `groupId` when the destination section is in a group notebook (see Graph copyToSection)'
    )
    .option('--json', 'Print operation Location and status')
    .option('--token <token>', 'Use a specific token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .option('--user <email>', 'Target user')
).action(
  async (
    pageId: string,
    opts: {
      section: string;
      groupId?: string;
      json?: boolean;
      token?: string;
      identity?: string;
      user?: string;
      group?: string;
      site?: string;
    },
    cmd: any
  ) => {
    checkReadOnly(cmd);
    const token = await requireGraphAuth(opts);
    const { user, scope } = parseOneNoteRoot(opts);
    const r = await copyOneNotePageToSection(token, pageId, opts.section, user, scope, opts.groupId);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(r.data, null, 2));
    } else {
      console.log(`Copy accepted (HTTP ${r.data.status}).`);
      if (r.data.operationLocation) {
        console.log(`Operation-Location: ${r.data.operationLocation}`);
        console.log('Poll with: m365-agent-cli onenote operation "<url>"');
      }
    }
  }
);

onenoteCommand
  .command('operation')
  .description('Get async OneNote operation status (copy, etc.)')
  .argument('<operationLocationUrl>', 'Full Operation-Location URL from copy response')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (operationLocationUrl: string, opts: { json?: boolean; token?: string; identity?: string }) => {
    const token = await requireGraphAuth(opts);
    const r = await getOneNoteOperation(token, operationLocationUrl);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else {
      console.log(`status: ${r.data.status ?? '?'}`);
      if (r.data.percentComplete) console.log(`percentComplete: ${r.data.percentComplete}`);
      if (r.data.resourceLocation) console.log(`resourceLocation: ${r.data.resourceLocation}`);
      if (r.data.error?.message) console.error(`error: ${r.data.error.message}`);
    }
  });
