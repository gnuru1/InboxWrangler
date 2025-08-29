# Outlook Auto-Tag (Sent-only + SQLite)

**Purpose**  
Mine Outlook **Sent Items** (optionally across all mounted stores), keep durable features in **SQLite**, compute **contact scores** (with boosts for internal domains, Core Team, Manager Chain), build a **VIP** set, and **auto-categorize upcoming calendar events**.

**Requirements**
- Windows + Outlook desktop (same profile as the user)
- Python 3.8+
- `pip install -r requirements.txt`

**Quick start**
```bash
pip install -r requirements.txt
# Edit config (internal domains, team lists):
#   outlook_autotag/config.py
python -m outlook_autotag.main                   # dry-run (no changes)
python -m outlook_autotag.main --apply          # apply categories
```

**Config highlights**
- Multiple **internal domains** (e.g. `acme.com`, `acme.co`)
- Explicit **CORE_TEAM** and **MANAGER_CHAIN**
- Optional **VIP_ALLOWLIST**
- Category names and colors configurable (e.g. rename "External" to "Client")

**What it writes**
- SQLite DB (default `outlook_auto_tag.db`)
- `contact_scores_sent_sqlite.csv` (for inspection)

**Safety**
- Default is **dry-run**: shows what would change and logs to DB audit; use `--apply` to write Outlook categories.
