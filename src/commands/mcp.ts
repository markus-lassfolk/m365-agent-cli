import { Command } from 'commander';
import { describeProgram } from '../lib/command-manifest.js';
import { runMcpStdioServer } from '../lib/mcp-server.js';

export const mcpCommand = new Command('mcp')
  .description(
    'Start an MCP (Model Context Protocol) stdio server that exposes every CLI command as an MCP tool, ' +
      'for MCP-aware clients (e.g. Claude Desktop, Claude Code) to call directly instead of shelling out to this CLI.'
  )
  .addHelpText(
    'after',
    `
Each leaf command (e.g. "mail", "rules create") becomes one MCP tool, named by its command path
with spaces replaced by underscores (e.g. "rules_create"), with a JSON schema built from that
command's own arguments and options. A tool call is executed by running this same CLI as a
subprocess with the equivalent argv, so behavior (read-only mode, --dry-run, --json errors) is
identical to running the command directly. "mcp", "serve", and "login" are not exposed as tools.

Example MCP client config (stdio transport):
  { "command": "m365-agent-cli", "args": ["mcp"] }
`
  )
  .action(async (_opts, cmd) => {
    const program = cmd.parent;
    const manifest = describeProgram(program);
    await runMcpStdioServer(manifest);
  });
