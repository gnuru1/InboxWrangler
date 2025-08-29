from __future__ import annotations

import argparse
import logging
import sys

from outlook_watcher.config import load_config
from outlook_watcher.db import DB
from outlook_watcher.outlook_events import OutlookWatcher
from outlook_watcher.backfill import backfill
from outlook_watcher.report import render_report

def setup_logging(level: str):
    lvl = getattr(logging, level.upper(), logging.INFO)
    logging.basicConfig(
        level=lvl,
        format="%(asctime)s %(levelname)s %(name)s - %(message)s"
    )

def cli(argv=None):
    parser = argparse.ArgumentParser(description="Outlook Inbox Watcher")
    parser.add_argument("--config", default="config.yaml", help="Path to config YAML")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_watch = sub.add_parser("watch", help="Run live watcher")
    p_watch.add_argument("--level", type=int, default=2, choices=[0,1,2], help="0=log only, 1=no email, 2=send notifications")

    p_back = sub.add_parser("backfill", help="Process historical Inbox/Sent items into DB")
    p_back.add_argument("--days", type=int, default=14, help="Days of history to process")

    p_rep = sub.add_parser("report", help="Print a simple response summary")
    p_rep.add_argument("--since-days", type=int, default=30, help="Days window for summary")

    args = parser.parse_args(argv)

    cfg = load_config(args.config)
    setup_logging(cfg.logging.level)
    db = DB(cfg.database_path)

    if args.cmd == "watch":
        watcher = OutlookWatcher(cfg, db)
        watcher.run(level=args.level)
    elif args.cmd == "backfill":
        backfill(cfg, db, days=args.days)
    elif args.cmd == "report":
        print(render_report(db, since_days=args.since_days))

if __name__ == "__main__":
    cli()
