from __future__ import annotations

import logging
from .db import DB

LOGGER = logging.getLogger(__name__)

def render_report(db: DB, since_days: int = 30) -> str:
    stats = db.report_summary(since_days=since_days)
    lines = []
    lines.append("=== Outlook Watcher Report (since {} days) ===".format(since_days))
    lines.append("Threads:         {}".format(stats.get('threads')))
    lines.append("Emails:          {}".format(stats.get('total_emails')))
    lines.append("Timers:          {}".format(stats.get('timers')))
    lines.append("  Responded:     {}".format(stats.get('responded')))
    lines.append("  Breached:      {}".format(stats.get('breaches')))
    lines.append("  Open:          {}".format(stats.get('open')))
    if stats.get('avg_minutes_to_respond') is not None:
        lines.append("Avg TTR (mins):  {}".format(stats['avg_minutes_to_respond']))
    return "\n".join(lines)
