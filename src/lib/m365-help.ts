import { type Command, Help } from 'commander';
import { buildRootCommandSections } from './root-command-groups.js';
import { buildNestedCommandSections } from './subcommand-help-groups.js';

/**
 * Custom Commander help: grouped subcommands for the root program and for parents
 * listed in `subcommand-help-groups.ts`. All other commands use the default layout.
 */
class M365Help extends Help {
  override formatHelp(cmd: Command, helper: Help): string {
    if (!cmd.parent) {
      return this.formatWithGroupedCommands(cmd, helper, () => buildRootCommandSections(cmd, helper));
    }
    const nested = buildNestedCommandSections(cmd, helper);
    if (nested) {
      return this.formatWithGroupedCommands(cmd, helper, () => nested);
    }
    return super.formatHelp(cmd, helper);
  }

  /**
   * Same structure as Commander 12 `Help.formatHelp` for usage/description/args/options,
   * then a grouped `Commands:` block. Keep aligned with upstream when upgrading Commander.
   */
  private formatWithGroupedCommands(
    cmd: Command,
    helper: Help,
    getSections: () => { title: string; commands: Command[] }[]
  ): string {
    const termWidth = helper.padWidth(cmd, helper);
    const helpWidth = helper.helpWidth || 80;
    const itemIndentWidth = 2;
    const itemSeparatorWidth = 2;
    const sectionTitleIndent = 2;
    const commandRowIndent = 4;

    function formatItem(term: string, description: string, wrapColumnWidthOffset: number): string {
      if (description) {
        const fullText = `${term.padEnd(termWidth + itemSeparatorWidth)}${description}`;
        return helper.wrap(fullText, helpWidth - wrapColumnWidthOffset, termWidth + itemSeparatorWidth);
      }
      return term;
    }

    function formatList(textArray: string[], indent: number): string {
      return textArray.join('\n').replace(/^/gm, ' '.repeat(indent));
    }

    let output = [`Usage: ${helper.commandUsage(cmd)}`, ''];

    const commandDescription = helper.commandDescription(cmd);
    if (commandDescription.length > 0) {
      output = output.concat([helper.wrap(commandDescription, helpWidth, 0), '']);
    }

    const argumentList = helper.visibleArguments(cmd).map((argument) => {
      return formatItem(helper.argumentTerm(argument), helper.argumentDescription(argument), itemIndentWidth);
    });
    if (argumentList.length > 0) {
      output = output.concat(['Arguments:', formatList(argumentList, itemIndentWidth), '']);
    }

    const optionList = helper.visibleOptions(cmd).map((option) => {
      return formatItem(helper.optionTerm(option), helper.optionDescription(option), itemIndentWidth);
    });
    if (optionList.length > 0) {
      output = output.concat(['Options:', formatList(optionList, itemIndentWidth), '']);
    }

    if (this.showGlobalOptions) {
      const globalOptionList = helper.visibleGlobalOptions(cmd).map((option) => {
        return formatItem(helper.optionTerm(option), helper.optionDescription(option), itemIndentWidth);
      });
      if (globalOptionList.length > 0) {
        output = output.concat(['Global Options:', formatList(globalOptionList, itemIndentWidth), '']);
      }
    }

    const sections = getSections();
    const commandBlocks: string[] = [];
    for (const section of sections) {
      const rows = section.commands.map((sub) =>
        formatItem(helper.subcommandTerm(sub), helper.subcommandDescription(sub), commandRowIndent)
      );
      commandBlocks.push(`${' '.repeat(sectionTitleIndent)}${section.title}`);
      commandBlocks.push(formatList(rows, commandRowIndent));
    }
    if (commandBlocks.length > 0) {
      output = output.concat(['Commands:', ...commandBlocks, '']);
    }

    return output.join('\n');
  }
}

/** Commander uses `command.createHelp()` on each node; assign custom help for `root` and all descendants. */
export function installM365HelpOnCommandTree(root: Command): void {
  root.createHelp = () => new M365Help();
  for (const sub of root.commands) {
    installM365HelpOnCommandTree(sub);
  }
}
