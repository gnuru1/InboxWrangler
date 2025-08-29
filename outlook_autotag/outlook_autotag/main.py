import argparse

from .config import CONFIG
from . import db
from . import outlook_utils as o
from . import sent_scanner
from . import calendar_learning
from . import scoring
from . import categorizer

def main(argv=None):
    parser = argparse.ArgumentParser(description="Sent-only Outlook auto-categorizer with SQLite persistence")
    parser.add_argument("--db", default=CONFIG["DB_PATH"], help="SQLite path (default: outlook_auto_tag.db)")
    parser.add_argument("--days-back-sent", type=int, default=CONFIG["DAYS_BACK_SENT"])
    parser.add_argument("--days-back-cal", type=int, default=CONFIG["DAYS_BACK_CAL"])
    parser.add_argument("--days-fwd-cal", type=int, default=CONFIG["DAYS_FWD_CAL"])
    parser.add_argument("--vip-top-n", type=int, default=CONFIG["VIP_TOP_N"])
    parser.add_argument("--no-scan-all-stores", action="store_true", help="Scan only default Sent Items")
    parser.add_argument("--apply", action="store_true", help="Apply categories (disable dry run)")
    args = parser.parse_args(argv)

    # Apply CLI overrides to runtime config (do not mutate file defaults on disk)
    cfg = dict(CONFIG)
    cfg["DB_PATH"] = args.db
    cfg["DAYS_BACK_SENT"] = args.days_back_sent
    cfg["DAYS_BACK_CAL"]  = args.days_back_cal
    cfg["DAYS_FWD_CAL"]   = args.days_fwd_cal
    cfg["VIP_TOP_N"]      = args.vip_top_n
    cfg["SCAN_ALL_SENT_STORES"] = not args.no_scan_all_stores
    cfg["DRY_RUN"] = not args.apply

    internal_domains = {d.lower() for d in cfg["INTERNAL_DOMAINS"]}
    ignore_parts = {s.lower() for s in cfg["IGNORE_LOCALPART_SUBSTR"]}
    core_team = set(map(str.lower, cfg["CORE_TEAM"]))
    manager_chain = set(map(str.lower, cfg["MANAGER_CHAIN"]))
    vip_allowlist = set(map(str.lower, cfg["VIP_ALLOWLIST"]))

    print(f"Profile: {o.ME}")
    print(f"Internal domains: {sorted(internal_domains)}")
    print(f"Core Team: {len(core_team)} | Manager Chain: {len(manager_chain)} | VIP allowlist: {len(vip_allowlist)}")
    print(f"DB: {cfg['DB_PATH']} | Dry-run: {cfg['DRY_RUN']} | Scan all stores: {cfg['SCAN_ALL_SENT_STORES']}")

    conn = db.connect(cfg["DB_PATH"])
    db.init_schema(conn)
    run_id = db.start_run(conn, cfg["CODE_VERSION"], cfg["DRY_RUN"])

    try:
        # Seed config contacts
        db.seed_contacts_from_config(conn, core_team, manager_chain, vip_allowlist)

        # 1) Mine Sent Items (incremental via checkpoints)
        print("Mining Sent Items…")
        msg_count, row_count = sent_scanner.mine_sent(
            conn, cfg["DAYS_BACK_SENT"], cfg["SCAN_ALL_SENT_STORES"], internal_domains, ignore_parts
        )
        print(f"  Inserted {msg_count} message rows, {row_count} recipient feature rows.")

        # 2) Learn from Calendar (organizer/accepted)
        print("Learning from Calendar…")
        meet_rows = calendar_learning.learn_from_calendar(conn, cfg["DAYS_BACK_CAL"], cfg["DAYS_FWD_CAL"])
        print(f"  Learned {meet_rows} meeting relationships.")

        # 3) Rebuild metrics & compute scores
        print("Rebuilding metrics & computing scores…")
        db.rebuild_contact_metrics(conn)
        scoring.compute_and_update_scores(conn, internal_domains, core_team, manager_chain)
        db.update_contacts_last_seen(conn)

        # 4) Build VIP set & tag upcoming calendar
        vip_set = scoring.build_vip_set(conn, vip_allowlist, core_team, manager_chain, cfg["VIP_TOP_N"], internal_domains)
        print(f"Tagging calendar (next {cfg['DAYS_FWD_CAL']} days)…")
        changed = categorizer.tag_calendar(conn, run_id, internal_domains, cfg, vip_set)
        if not cfg["DRY_RUN"]:
            print(f"Applied categories to {changed} appointments.")

        # 5) Export CSV for inspection
        db.export_scores_csv(conn, "contact_scores_sent_sqlite.csv")
        print("Wrote contact_scores_sent_sqlite.csv")

    finally:
        db.finish_run(conn, run_id)
        conn.close()

if __name__ == "__main__":
    raise SystemExit(main())
