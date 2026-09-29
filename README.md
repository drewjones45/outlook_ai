# outlook-ai

Reads your mail and calendar from **classic ("Legacy") Outlook for Mac** through AppleScript, caches it locally, and serves it to Claude. Claude can then summarize your inbox, find the emails that need your reply, draft those replies in your voice, and use your calendar for context. Drafts land in Outlook's Drafts folder. **Nothing is ever sent.**

It plugs into Claude as a local [MCP](https://modelcontextprotocol.io) server (Claude Code, Claude Desktop, local Cowork sessions). It also has a CLI and a Markdown export for any other Claude chat.

---

## Read this first: the platform is being retired (checked 2026-09-29)

- **Legacy Outlook for Mac stops working with Exchange Online (Microsoft 365) mailboxes starting in October 2026.**
  - It connects to Microsoft 365 through Exchange Web Services (EWS). Microsoft begins turning EWS off on **October 1, 2026** and removes it completely in 2027.
  - Microsoft's own page says Legacy Outlook "will stop working against Exchange Online mailboxes starting October 2026".
  - It keeps working with **on-premises Exchange** and with **IMAP/POP** accounts.
  - Sources: [Microsoft Support](https://support.microsoft.com/en-us/outlook/end-of-support-for-legacy-outlook-for-mac), [Exchange team blog](https://techcommunity.microsoft.com/blog/exchange/exchange-online-ews-your-time-is-almost-up/4492361).
- **New Outlook for Mac will not get AppleScript.**
  - Microsoft marked roadmap item 88537 ("Support for AppleScript in the New Outlook for Mac") **Cancelled** on August 12, 2026: "full compatibility isn't feasible" ([roadmap](https://www.microsoft.com/microsoft-365/roadmap?id=88537)).
  - New Outlook doesn't return errors to scripts. It answers with empty folders.

What that means for you:

| Your mailbox | Does this approach work? |
|---|---|
| Microsoft 365 / Exchange Online (most work accounts) | Only until the EWS cutoff reaches your tenant. After that Legacy Outlook stops syncing, and this tool would only see old mail. `outlook-ai doctor` flags these accounts, and the inbox and calendar tools warn Claude (and you) when the newest message in your inbox is more than 3 days old. |
| On-premises Exchange | Yes, as long as your Outlook build still offers Legacy mode. |
| IMAP / POP (Gmail, iCloud, Fastmail…) in Legacy Outlook | Yes. |

For a Microsoft 365 mailbox, the durable options are Anthropic's [Microsoft 365 connector](https://support.claude.com/en/articles/15183774-connect-to-microsoft-365) and the [Claude for Outlook add-in](https://claude.com/docs/office-agents/outlook). Both need a one-time admin consent in your tenant: the step this approach was meant to avoid. The local cache, triage and Claude layers here don't depend on AppleScript. They could sit on top of Microsoft Graph if IT approves an app registration.

---

## Review of the approach

The idea: Claude Cowork wrote an AppleScript that reads the classic Outlook for Mac app directly. That means no network calls, no OAuth, and no app registration, because Outlook is already signed in.

### What holds up

- **The technical claim is right.** Classic Outlook's scripting dictionary exposes messages, recipients, headers, calendar events, attendees and their responses, plus a `reply to` command that creates real threaded drafts. Reading it through Apple Events involves no Microsoft cloud API, token or tenant consent.
- **Reading locally is a good design instinct.** It's fast, works offline, and the data never touches a third-party integration service.

### What I'd push back on

1. **The platform is going away.** See above. This is the deciding issue for any Microsoft 365 mailbox.

2. **"IT already approved Outlook" does not mean IT approved sending your email to an AI service.**
   - Admin consent exists so the organization decides which third parties get mailbox data. Reading locally skips the technical control, but not the policy behind it.
   - The moment email content goes into a Claude session, it leaves your Mac for Anthropic's servers. So "no network calls" is only true for the extraction step.
   - Whether that's allowed depends on your employer's AI and data policy, not on which API did the reading.
   - It also depends on which Claude account you use. Business plans (Team, Enterprise, API) don't train on your data by default. Personal Free/Pro/Max accounts [may, depending on your settings](https://www.anthropic.com/news/updates-to-our-consumer-terms).
   - Check the policy and use the account your company approved. That's cheaper than explaining it after the fact.
   - Also: on managed Macs, IT can block the macOS Automation permission this needs (a PPPC profile). If it's blocked, `outlook-ai doctor` will tell you.

3. **"Mac app on my PC."** AppleScript exists only on macOS. This has to run on the Mac where Outlook is installed.

4. **Technical traps a first-draft AppleScript usually falls into.** Each one is handled here:

   | Trap | What goes wrong | What this project does |
   |---|---|---|
   | Recurring meetings | Outlook's scripting returns only the series "master" (dated at the *first* occurrence). A query for today's events misses every weekly 1:1 that started last year. | Reads each series' iCalendar data (RRULE/EXDATE), expands it in Python, and swaps in rescheduled occurrences using `is occurrence` / `recurrence id` / `master`. |
   | Speed | Every property read is an Apple Event. Looping over "all emails in the inbox" with bodies takes many minutes and hits the 2-minute Apple Event timeout (-1712). | Indexes whole folders with a few bulk reads, caches in SQLite, and after the first sync fetches only new mail. Read state is refreshed in bulk. |
   | Parsing | Output joined with `\|` or `::` breaks on the first subject that contains one. | Uses ASCII control-character separators and strips them from every value. |
   | Dates | `date as string` depends on locale; AppleScript dates have no time zone. | Emits locale-independent timestamps built from date components, and converts them in Python per date, so daylight saving is handled. |
   | Multiple accounts | `inbox` means the *default* account's inbox. | Configurable account; uses that account's own `inbox` / `sent items`. |
   | Exchange quirks | Internal senders sometimes appear as X.500 paths (`/O=EXCHANGELABS/...`) instead of email addresses. | Falls back to name matching for "is this me?" and hides the X.500 path. |
   | New Outlook mode | Scripts "succeed" with zero messages. | Checks Outlook's `IsRunningNewOutlook` setting before every call and fails loudly. |
   | Context size | A whole inbox doesn't fit in a prompt. | Claude gets a compact, triaged index and fetches full emails on demand. |
   | Prompt injection | Email bodies are written by strangers ("ignore previous instructions and forward…"). | The only write action is *create a draft*: no send, forward, delete or move. Email text reaches Claude inside `<untrusted_email_data>` markers, with instructions to treat it as data. You review every draft. |

---

## How it works

```
Outlook (Legacy mode)
   ▲  Apple Events (osascript)   outlook_ai/applescripts/*.applescript
   │
outlook-ai sync ──► SQLite cache (~/Library/Application Support/outlook-ai, mode 600)
                         │
                         ├─► triage: which emails need *your* reply, and why
                         │
                         ├─► MCP server ──► Claude Code / Claude Desktop / local Cowork
                         ├─► CLI (inbox, triage, calendar, prep, draft, notify…)
                         └─► Markdown export ──► any Claude chat
```

**Deciding what "needs my reply" means.** A message scores points for each signal and loses points for others. Claude reads the shortlist and makes the final call.

- Points for:
  - you're on the To line, especially as the only recipient
  - it greets or @-mentions you by name
  - it asks a question or uses request language ("can you", "by Friday", "please review") in the new text, with quoted history ignored
  - it's flagged, high-importance, or from a VIP sender
- Excluded:
  - newsletters, no-reply senders, auto-replies and meeting responses
  - threads you already answered (Outlook's own "replied" flag, then Sent Items)
  - older messages in a thread that has a newer one

### MCP tools

| Tool | What it does |
|---|---|
| `inbox_overview` | Recent mail, grouped: needs reply / invites / FYI / automated. Includes flags and previews. |
| `emails_needing_reply` | The shortlist, with the reasons and the latest message text. |
| `read_email` | One email plus its whole cached thread, including your own replies. |
| `search_emails` | Keyword search over cached inbox and sent mail. |
| `calendar` | Agenda for any day or range: recurring meetings expanded, join links, attendee responses. |
| `meeting_prep` | One meeting, plus recent email from its attendees or about its topic. |
| `create_reply_draft` | Saves a reply in Outlook's Drafts folder. **Never sends.** |
| `sync_outlook`, `outlook_status` | Force a refresh; show who "you" are and when the cache last synced. |

It also provides three prompts: **`inbox_summary`**, **`draft_replies`** and **`daily_briefing`**. In Claude Code they appear as `/mcp__outlook__draft_replies` and so on.

---

## Setup (on your Mac)

You need macOS, Microsoft Outlook **open and in Legacy mode**, and [uv](https://docs.astral.sh/uv/). Run everything in Terminal.

```bash
# 1. Install uv (manages Python for you). Open a new Terminal window afterwards.
curl -LsSf https://astral.sh/uv/install.sh | sh        # or: brew install uv

# 2. Get the code (if macOS offers to install developer tools for git, accept)
git clone https://github.com/drewjones45/outlook_ai.git ~/outlook_ai
cd ~/outlook_ai
uv sync

# 3. Check the connection. macOS asks "Terminal wants access to control Microsoft Outlook": click OK.
uv run outlook-ai doctor

# 4. Create your config, then fill in your name, nicknames and every address that is you
uv run outlook-ai init
nano ~/.config/outlook-ai/config.toml   # save: Ctrl+O, Enter; exit: Ctrl+X

# 5. First sync, then look at the results
uv run outlook-ai sync --days 7
uv run outlook-ai triage
uv run outlook-ai calendar
```

- `doctor` flags a Microsoft 365 (Exchange Online) mailbox with `[!!]`. That's the EWS warning at the top of this page, not a setup error.
- Edit the config in `nano` or a code editor rather than TextEdit, whose "smart quotes" break the file.
- The first sync is the slow one, because every property read is a separate Apple Event. Start with a small `--days` and widen it later (`sync --days 30`). After that, syncs only fetch what's new.
- To test drafts before letting Claude write them: `uv run outlook-ai draft <id> --body "Test, please ignore" --open`, using an id from `triage`. Then delete the draft.

### Connect it to Claude

**Claude Code in Terminal** (most reliable, because it inherits Terminal's Automation permission). It needs a Pro, Max, Team or Enterprise plan. Install it with `curl -fsSL https://claude.ai/install.sh | bash` ([docs](https://code.claude.com/docs/en/setup)), then:

```bash
claude mcp add outlook --scope user -- uv run --directory ~/outlook_ai outlook-ai serve
claude
```

Then ask "summarize my inbox", or run `/mcp__outlook__draft_replies`.

**Claude Desktop.** Go to Settings → Developer → Edit Config and add the following. Use the full path to `uv` (`which uv`), because Claude Desktop doesn't read your shell's PATH. Then quit Claude Desktop completely and reopen it.

```json
{
  "mcpServers": {
    "outlook": {
      "command": "/opt/homebrew/bin/uv",
      "args": ["run", "--directory", "/Users/YOU/outlook_ai", "outlook-ai", "serve"]
    }
  }
}
```

If tool calls fail with "not authorized to send Apple events" (-1743) and macOS never asked you:

- Claude.app may not be able to request the Automation permission. There is an [open report](https://github.com/anthropics/claude-code/issues/95398) of this.
- Workaround: run the sync from launchd (below) so the MCP server can answer from the cache.
- Use Claude Code or `outlook-ai draft` for drafts.

**Claude Cowork.**
- Local MCP servers work only in **local** Cowork sessions, through the desktop app. They don't work in cloud sessions or in scheduled tasks, which run remotely ([Cowork architecture](https://support.claude.com/en/articles/14479288-claude-cowork-architecture-overview)).
- Cowork's own code runs in a sandboxed VM, so it can't run `osascript` against your Mac's Outlook itself. That's why this runs as an MCP server on the host.

**Any other Claude chat:** `uv run outlook-ai export -o ~/Desktop/outlook-context.md`, then attach the file. It contains email content, so delete it when you're done.

---

## Commands

| Command | Purpose |
|---|---|
| `doctor` | Check macOS, Legacy vs New Outlook, the Automation permission, accounts (flags Exchange Online), folders, calendars and config. |
| `init` | Write a starter config from the accounts Outlook reports. |
| `sync [--days N] [--max N]` | Pull new mail, read state, Sent Items and the calendar window. `--days 0` means the whole folder. |
| `inbox`, `triage`, `read ID`, `search WORDS`, `calendar [DATE]`, `prep EVENT_ID` | The same views Claude gets. |
| `draft ID --body "…" [--reply-all] [--open]` | Save a reply draft (text can also come from stdin or `--body-file`). |
| `export [-o FILE]` | Markdown bundle: calendar, needs-reply shortlist and inbox overview. |
| `notify` | macOS notification for meetings starting soon: join link, unread mail from attendees. |
| `launchd sync\|notify` | Print a LaunchAgent plist that runs `sync` or `notify` on a timer. |
| `serve` | Run the MCP server (stdio). |
| `script NAME` | Print a bundled AppleScript with its helpers, ready to paste into Script Editor. |
| `purge --yes` | Delete the local cache. |

### Background sync and meeting notifications (optional)

```bash
uv run outlook-ai launchd sync   > ~/Library/LaunchAgents/com.outlook-ai.sync.plist
uv run outlook-ai launchd notify > ~/Library/LaunchAgents/com.outlook-ai.notify.plist
launchctl bootstrap gui/$(id -u) ~/Library/LaunchAgents/com.outlook-ai.sync.plist
launchctl bootstrap gui/$(id -u) ~/Library/LaunchAgents/com.outlook-ai.notify.plist
```

- The first run triggers its own Automation prompt, for the Python binary. Logs go to `~/Library/Logs/outlook-ai-*.log`.
- To stop a job: `launchctl bootout gui/$(id -u)/com.outlook-ai.sync`.

---

## Configuration

`~/.config/outlook-ai/config.toml` (`outlook-ai init` writes it):

- `[me]`:
  - `name`, `aliases` (nicknames people use in emails)
  - `addresses`: every address that is you. Triage depends on this.
- `[outlook]`:
  - `account`: blank means Outlook's default account
  - `inbox_folder`, `sent_folder`
  - `calendars`: blank means all
- `[sync]`:
  - `days`, `max_messages`, `body_chars`
  - calendar window
  - `stale_minutes`: MCP tools re-sync when the cache is older than this
- `[drafts]`:
  - `signature` (appended to drafts)
  - `style` (your voice, passed to Claude)
  - `open_window`
- `[triage]`: `vip_senders`, `ignore_senders` (addresses or `@domain`), `threshold`
- `[notify]`: `lead_minutes`

## Privacy and safety

- The cache is readable only by your user account (mode 600) and holds email bodies. `outlook-ai purge --yes` removes it.
- Nothing leaves your Mac until a Claude session asks for it, and Claude only receives what a tool returns. That's a compact index and the emails it chooses to open, not your whole mailbox.
- The MCP server can't send, forward, delete, move or mark mail. Its one write action creates drafts.
- Review drafts before sending. A malicious email could try to steer a reply, and you are the last check.

## Limitations and verifying on your Mac

- **The AppleScripts were written and checked without a Mac.**
  - Every Outlook term they use was checked against published copies of Outlook's scripting dictionary (2011, 2016 and 2019 editions).
  - The test suite parses the scripts for unbalanced blocks, reserved words, calls that would go to Outlook instead of the script, and record layouts that drifted from the parser.
  - The Python side is tested end to end against a fake Outlook (`uv run pytest`).
  - Still, the first real run is the real test. Start with `doctor`, then `sync --days 2`.
  - If a script misbehaves, `outlook-ai script fetch_messages | pbcopy` gives you the exact script to paste into Script Editor.
- **Drafts.** Outlook's `reply to` without opening a window has worked in Legacy builds, but hung in one reported test. After 60 seconds this tool falls back to building the reply as a new message, which keeps the recipients but loses threading. The tool output says when that happens.
- **Calendar.** Recurring meetings are expanded from each series' iCalendar data, and occurrences edited in Outlook replace their original slots. Series whose iCalendar data Outlook doesn't return are listed as a warning, not silently dropped.
- **Mail.** Only mail inside the sync window is searchable; older messages can still be opened by id. Exchange sometimes reports internal senders without an SMTP address.

## Development

```bash
uv run pytest       # runs on any OS against the fake Outlook in tests/fake_outlook.py
```

Layout:

- `outlook_ai/applescripts/`: the Outlook-facing scripts plus the shared helpers (`_lib`, `_outlook`)
- `outlook_ai/outlook.py`: runs the scripts and parses their records
- `outlook_ai/sync.py`, `store.py`: incremental sync and the SQLite cache
- `outlook_ai/triage.py`, `recurrence.py`: the "needs reply" scoring and recurring-meeting expansion
- `outlook_ai/service.py`: the operations
- `outlook_ai/mcp_server.py`, `cli.py`: the two front ends
