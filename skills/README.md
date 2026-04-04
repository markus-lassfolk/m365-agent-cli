# OpenClaw Agent Skills

This directory contains pre-packaged **Agent Skills** designed for [OpenClaw](https://github.com/openclaw/openclaw) or other autonomous AI agents supporting the `.skill` or `SKILL.md` format.

These skills teach the AI *how* to use the `m365-agent-cli` CLI and *how* to behave when managing your digital life.

## Included Skills

### 1. `m365-agent-cli` (The Technical Manual)
Located in `./m365-agent-cli/SKILL.md`, this is the strict technical documentation for the CLI. It teaches the AI agent the exact syntax, flags, and endpoints required to interact with Microsoft 365 (e.g., `mail`, `calendar`, **`contacts`**, **`onenote`**, **`meeting`**, `files`, `planner`, `sharepoint`). It also covers **categories/labels** (Outlook vs To Do vs Planner), **calendar ranges** (including business-day windows, **`--now`**, and **`delete-event --scope`** / recurring occurrences), and **`outlook-categories`**. The AI reads this to know how to execute actions on your behalf without hallucinating commands. The skill frontmatter **`version`** is kept in sync with the npm package via **`npm run sync-skill`** when releasing (see [docs/RELEASE.md](../docs/RELEASE.md)).

### 2. `personal-assistant` (Moved)
The **Personal Assistant** behavioral playbook and associated ecosystem recommendations have been moved to their own dedicated repository.

Please visit the **[openclaw-personal-assistant](https://github.com/markus-lassfolk/openclaw-personal-assistant)** repository for the Master Guide and installation instructions.

## Installation

To grant these superpowers to your local OpenClaw agent, simply copy the directories into your agent's workspace:

```bash
mkdir -p ~/.openclaw/workspace/skills
cp -r skills/* ~/.openclaw/workspace/skills/
```
