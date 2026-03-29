# OpenClaw Agent Skills

This directory contains pre-packaged **Agent Skills** designed for [OpenClaw](https://github.com/openclaw/openclaw) or other autonomous AI agents supporting the `.skill` or `SKILL.md` format.

These skills teach the AI *how* to use the `clippy` CLI and *how* to behave when managing your digital life.

## Included Skills

### 1. `clippy` (The Technical Manual)
Located in `skills/clippy/SKILL.md`, this is the strict technical documentation for the CLI. It teaches the AI agent the exact syntax, flags, and endpoints required to interact with Microsoft 365 (e.g., `mail`, `calendar`, `files`, `planner`, `sharepoint`). The AI reads this to know how to execute actions on your behalf without hallucinating commands.

### 2. `personal-assistant` (The Behavioral Playbook)
Located in `skills/personal-assistant/SKILL.md`, this is the behavioral framework for an Executive Assistant. Rather than just giving the AI tools, this gives it a **proactive persona**.

#### The PA Persona (High-Level)
When an agent loads this skill, it adopts the mindset of a Chief of Staff:
- **Proactive Triage:** It doesn't wait to be told to check your email. It runs background checks, flags urgent items, isolates newsletters, and prepares draft responses for you to review.
- **Calendar Defense:** It actively negotiates meeting times using `findtime` to prevent double-bookings and email ping-pong.
- **Context Retention:** It remembers facts about people, projects, and decisions in its long-term memory, surfacing that information right before your meetings.
- **Scam Defense:** It acts as a shield, labeling suspicious emails and asking for permission before moving them to Junk.

## Installation

To grant these superpowers to your local OpenClaw agent, simply copy the directories into your agent's workspace:

```bash
mkdir -p ~/.openclaw/workspace/skills
cp -r skills/* ~/.openclaw/workspace/skills/
```
