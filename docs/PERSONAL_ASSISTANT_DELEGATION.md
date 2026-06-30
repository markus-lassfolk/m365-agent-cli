# Personal-assistant and delegated workflows

How to use **`m365-agent-cli`** when you automate mail, calendar, files, Teams, and directory tasks **for yourself**, for a **shared mailbox**, or **on behalf of another user** (manager) using Graph **`/users/{id}/...`** paths. Wrong flag combinations are a common source of 403/404—use this as a map, then confirm scopes in [`GRAPH_SCOPES.md`](./GRAPH_SCOPES.md).

## 1. Three different “act as” stories

| Goal | Prefer | Notes |
| --- | --- | --- |
| **Shared mailbox (Exchange / EWS model)** | **`--mailbox <email>`** on commands that support it (calendar, mail, send, folders, …) | EWS path; not the same as Graph **`--user`**. See [`MIGRATION_TRACKING.md`](./MIGRATION_TRACKING.md) if the tenant is Graph-first. |
| **Graph “that user’s” resource** | **`--user <upn-or-id>`** on Graph-backed commands that implement it | Switches to **`/users/{id}/...`**. Examples: **`files`**, **`graph-calendar`**, **`outlook-graph`**, **`rules`**, **`oof`**, **`subscribe`**, **`teams list`**, **`org`**, etc. |
| **Calendar sharing (Graph permissions on a calendar)** | **`delegates calendar-share`** (not classic EWS delegates) | When **`M365_EXCHANGE_BACKEND=graph`**, use Graph **`calendarPermission`** via **`calendar-share`**; classic **`delegates add`** is EWS. See [`GRAPH_EWS_PARITY_MATRIX.md`](./GRAPH_EWS_PARITY_MATRIX.md) §2a. |

## 2. Manager calendar and mail (assistant scenarios)

- **Calendar view / events:** **`calendar`** / **`graph-calendar`** with **`--user`** where supported, or EWS **`--mailbox`** for shared mailboxes—pick one stack per mailbox.
- **To Do:** **`todo --user <email>`** uses Graph `/users/{id}/todo/...` and requires **Tasks.Read.Shared** / **Tasks.ReadWrite.Shared** in addition to the user's actual sharing/delegation rights.
- **Mail:** **`mail`**, **`outlook-graph`**, **`folders`** — Graph variants often support **`--user`** for delegated mailboxes your token can access (**Mail.Read\*\.Shared** scopes).
- **Subscriptions (notifications):** **`subscribe`** with **`--user`** when the resource lives under **`users/{id}`** (mail folders, events, calendar, Copilot resources—see command help).
- **Org hierarchy:** **`org manager`** (your manager or **`--user`**’s manager) and **`org direct-reports`** — routing (“who approves?”, “who reports here?”). Requires **`User.Read`** for self paths and typically **`User.Read.All`** for **`--user`** (tenant-dependent).

## 3. Teams

- **Joined teams for another person:** **`teams list --user <upn-or-id>`** → **`GET /users/{id}/joinedTeams`**. Needs **`Team.ReadBasic.All`** (or equivalent per Graph docs).
- **Channels / messages:** Once you have **`teamId`** / **`channelId`**, other **`teams`** subcommands use **`/teams/{id}/...`** and do not need **`--user`** on the path.
- **Chats:** **`teams chats`** uses **`GET /me/chats`** only. There is **no** supported Graph pattern to “list all chats for user B” the way **`/me/chats`** works for the signed-in user. Use a **manager session**, **shared channels**, or tenant-specific tools; do not expect a **`--user`** flag on **`chats`**.

## 4. Bookings and app-only tokens

- **`bookings staff-availability`** may require an **application** token (**`--token`**) per Microsoft; delegated tokens can fail by design. Not an CLI bug—see [`GRAPH_API_GAPS.md`](./GRAPH_API_GAPS.md).

## 5. Escaping the wrapped surface

- Arbitrary JSON Graph paths: **`graph invoke`**, **`graph batch`**.
- Policies that stay out of first-class commands: [`GRAPH_INVOKE_BOUNDARIES.md`](./GRAPH_INVOKE_BOUNDARIES.md) (Cloud Communications except wrapped meeting recordings/transcripts, **PSTN out of scope**, RSC, governance patterns).

## 6. Verify your token

- **`m365-agent-cli verify-token`** and **`verify-token --capabilities`** to see whether **`User.Read.All`**, **`Team.ReadBasic.All`**, shared mail/calendar scopes, etc. are present before scripting manager workflows.

## 7. CLI coverage vs Graph `/users/{id}/...` (gap hints)

These are **common 403/404** causes when mixing “act as user” expectations with what the CLI actually wraps:

| Scenario | Graph supports delegated `/users/{id}/...`? | CLI today |
| --- | --- | --- |
| Mail, calendar, contacts, files, rules, OOF, many `subscribe` resources | Often yes (with **Shared** scopes where needed) | **`--user`** on supported commands — confirm **`--help`**. |
| Joined Teams for another user | Yes (`/users/{id}/joinedTeams`) | **`teams list --user`**. |
| List **chats** for another user | No first-class list analogous to `/me/chats` | **`teams chats`** = signed-in user only; use manager session or other tooling. |
| Teams activity feed notify for another user | App-only path documented | **`teams activity-notify`** = `/me/…` or chat; app-only **`/users/{id}/teamwork/...`** → **`graph invoke`**. |
| Bookings staff availability | Typically **application** permission | **`bookings staff-availability --token`** (app token). |
| Insights (`/me/insights/...`) with delegation | Graph supports user-scoped paths | **`insights * --user`** where the command exposes the flag. |

Full gap table and pagination/delegation hardening notes: [`GRAPH_WRAPPER_GAP_AUDIT.md`](./GRAPH_WRAPPER_GAP_AUDIT.md).

---

*Related: [`skills/m365-agent-cli/SKILL.md`](../skills/m365-agent-cli/SKILL.md) (delegation bullets), [`GRAPH_SCOPES.md`](./GRAPH_SCOPES.md).*
