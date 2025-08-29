from . import outlook_utils as o

def ensure_categories(config):
    cats = o.session.Categories
    existing = {c.Name for c in cats}
    defs = []
    for name, color_name in config["CATEGORY_DEFS"]:
        if name == "External":
            name = config["EXTERNAL_CATEGORY_NAME"]
        defs.append((name, o.cat_color(color_name)))

    for name, color in defs:
        if name not in existing:
            cats.Add(name, color)

def classify_appointment(appt, scores_map, vip_set, internal_domains, config):
    cats = []
    req, opt = o.attendees_from_appt(appt)
    participants = set(p.lower() for p in (req + opt))
    size = len(participants)
    if size == 0:
        return cats

    core_set = set(e.lower() for e in config["CORE_TEAM"])
    mgr_set  = set(e.lower() for e in config["MANAGER_CHAIN"])

    has_mgr   = any(p in mgr_set for p in participants)
    has_core  = any(p in core_set for p in participants)
    has_vip   = any(p in vip_set for p in participants)
    has_ext   = any(not o.is_internal(p, internal_domains) for p in participants)

    if has_vip:  cats.append("VIP")
    if has_mgr:  cats.append("Manager Chain")
    if has_core: cats.append("Core Team")
    if has_ext and config["EXTERNAL_CATEGORY_NAME"] not in cats:
        cats.append(config["EXTERNAL_CATEGORY_NAME"])

    # Focus: 1:1 with VIP/Manager/Core or strong score
    if size == 1:
        only = next(iter(participants))
        if (only in vip_set or only in mgr_set or only in core_set or scores_map.get(only, 0.0) >= 0.80):
            cats.append("Focus")

    # Low-signal: large or nobody with history
    if size >= 8 or all(scores_map.get(p, 0.0) < 0.20 for p in participants):
        cats.append("Low-signal")

    return cats

def set_categories(appt, new_cats: list):
    current = (appt.Categories or "").strip()
    curr_set = set([c.strip() for c in current.split(";") if c.strip()]) if current else set()
    merged = curr_set | set(new_cats)
    merged_str = "; ".join(sorted(merged))
    if merged == curr_set:
        return False, current, merged_str
    appt.Categories = merged_str
    appt.Save()
    return True, current, merged_str

def tag_calendar(conn, run_id: int, internal_domains, config, vip_set):
    ensure_categories(config)

    # Build score map
    cur = conn.cursor()
    cur.execute("SELECT email, score FROM contact_metrics")
    scores_map = {e.lower(): (s or 0.0) for (e, s) in cur.fetchall()}

    # Time window
    start = __import__("datetime").datetime.now()
    end   = __import__("datetime").datetime.now() + __import__("datetime").timedelta(days=config["DAYS_FWD_CAL"])
    items = o.ns.GetDefaultFolder(o.constants.olFolderCalendar).Items
    items.IncludeRecurrences = True
    items.Sort("[Start]")
    restriction = f"[Start] >= '{o.dt_us_str(start)}' AND [Start] <= '{o.dt_us_str(end)}'"

    tagged = 0
    dry = config["DRY_RUN"]
    for appt in items.Restrict(restriction):
        cats = classify_appointment(appt, scores_map, vip_set, internal_domains, config)
        if not cats:
            continue
        ak = o.appt_key(appt)
        conn.execute("""                INSERT OR IGNORE INTO appts(appt_entry_id, global_id, start_utc, end_utc, subject)
            VALUES (?, ?, ?, ?, ?)
        """, (ak, getattr(appt, "GlobalAppointmentID", None),
              o.to_iso_utc(appt.Start), o.to_iso_utc(appt.End),
              (getattr(appt, "Subject", "") or "")[:500]))

        if dry:
            old_cats = (appt.Categories or "").strip()
            new_cats = "; ".join(sorted(set([c.strip() for c in old_cats.split(';') if c.strip()]) | set(cats)))
            conn.execute("""                    INSERT OR IGNORE INTO appt_tags(appt_entry_id, run_id, old_categories, new_categories, tagged_at_utc)
                VALUES (?, ?, ?, ?, ?)
            """, (ak, run_id, old_cats, new_cats, o.now_utc_iso()))
            print(f"[DRY] {appt.Start}  {appt.Subject} -> {cats}")
        else:
            changed, old_cats, new_cats = set_categories(appt, cats)
            if changed:
                tagged += 1
                conn.execute("""                        INSERT OR IGNORE INTO appt_tags(appt_entry_id, run_id, old_categories, new_categories, tagged_at_utc)
                    VALUES (?, ?, ?, ?, ?)
                """, (ak, run_id, old_cats, new_cats, o.now_utc_iso()))
    conn.commit()
    return tagged
