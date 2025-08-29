# Outlook Inbox Watcher

A Windows-only utility that watches Outlook for new threads that target your team (via To/CC distribution list or a set of member addresses), detects callouts in the message text (to the team or specific members), tracks a response SLA timer (default 60 minutes), and sends a notification email if the timer expires without a team response. It can also backfill N days of history to establish baseline response patterns. All activity is written to a SQLite database for audit and reporting.

> **Requires**: Windows, Outlook desktop (MAPI), Python 3.9+, and `pywin32`.

---

## Features
- Watch **Inbox** and **Sent Items** using `pywin32` COM events
- Filter threads where **To/CC** contains your configured distribution lists or team member addresses
- Parse email body/subject for **callouts** to the team or specific members (configurable aliases)
- Maintain **SLA timers** (default 60 minutes) and **notify** if no team response arrives by the deadline
- **Backfill** N days of Inbox/Sent to bootstrap history and response patterns
- **SQLite** data store for threads, emails, mentions, timers, and notifications
- Minimal **reports**: average time-to-first-response, breach counts, and more

---

## Quick Start

1. **Install dependencies** (in an elevated shell if needed)  
   ```powershell
   py -m pip install -r requirements.txt
   ```

2. **Copy and edit the config**  
   ```powershell
   copy config.example.yaml config.yaml
   # open config.yaml and set your team, lists, and recipients
   ```

3. **Run a backfill first** (optional but recommended)  
   ```powershell
   py main.py backfill --days 14 --config config.yaml
   ```

4. **Start the watcher** (keeps running; press Ctrl+C to stop)  
   ```powershell
   py main.py watch --config config.yaml --level 2
   ```
   Levels:
   - `0` = log only (no notifications)
   - `1` = create timers and mark breaches, but **do not send** emails
   - `2` = full operation (**send notifications**)

5. **Generate a simple report**  
   ```powershell
   py main.py report --since-days 30 --config config.yaml
   ```

---

## Important Notes

- This tool uses **COM events**, which require the Python process to run continuously with a message pump. The `watch` mode sets this up for you.  
- It is designed for a **single Outlook profile** on the machine where it runs. If your team uses multiple mailboxes or shared mailboxes, you can point `outlook.profile` and `outlook.store_display_name` to the right store/folder in the config.
- To avoid feedback loops, notification emails are marked with a category (default: `OutlookWatcher`) and ignored by the watcher.
- If Outlook raises security prompts for programmatic access, see your admin about trusting the signed code or adjusting **Programmatic Access** settings in Trust Center.

---

## CLI

```
py main.py --help
py main.py watch --config config.yaml --level 2
py main.py backfill --days 30 --config config.yaml
py main.py report --since-days 90 --config config.yaml
```

---

## Data Model (SQLite)

**threads**: conversation-level state  
**emails**: one row per message processed  
**mentions**: extracted team/member callouts  
**timers**: SLA timers created on callouts (and marked responded/breached)  
**notifications**: record of notifications sent  

The database path is configurable via `database_path` in `config.yaml` (default: `outlook_watcher.db`).

---

## Safety / Scope

- This tool **does not** auto-reply to external senders. It only sends an **internal notification** using your configured template.  
- It **only** processes MailItems (`IPM.Note`). Meeting requests / receipts are ignored.  
- You can run multiple watchers (e.g., per mailbox) by deploying separate configs & DBs.

---

## License

MIT
