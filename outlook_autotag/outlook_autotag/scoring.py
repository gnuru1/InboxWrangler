import datetime as dt

from . import outlook_utils as o

def compute_and_update_scores(conn, internal_domains, core_team, manager_chain):
    cur = conn.cursor()
    cur.execute("""            SELECT cm.email, cm.sent_msgs, cm.sent_replies, cm.sent_initiated, cm.direct_msgs, cm.broad_msgs,
               cm.importance_high, cm.attachments_sent, cm.meetings_organized_with, cm.meetings_accepted_from,
               cm.last_sent_utc,
               COALESCE(c.is_core,0) AS is_core, COALESCE(c.is_manager,0) AS is_manager, COALESCE(c.is_vip_seed,0) AS is_vip_seed
        FROM contact_metrics cm
        LEFT JOIN contacts c ON c.email = cm.email
    """)
    rows = cur.fetchall()

    max_sent = max([r[1] for r in rows] or [1])
    updates = []
    for (email, sent_msgs, sent_replies, sent_initiated, direct_msgs, broad_msgs,
         importance_high, attachments_sent, meetings_organized_with, meetings_accepted_from,
         last_sent_utc, is_core, is_manager, is_vip_seed) in rows:

        if not email or sent_msgs < 2:
            score = 0.0
        else:
            vol_norm        = (sent_msgs / max_sent) if max_sent else 0.0
            direct_ratio    = direct_msgs / max(sent_msgs, 1)
            broad_ratio     = broad_msgs  / max(sent_msgs, 1)
            initiated_ratio = sent_initiated / max(sent_msgs, 1)
            importance_sig  = min(1.0, importance_high / 5.0)
            attach_sig      = min(1.0, attachments_sent / 4.0)
            meeting_sig     = min(1.0, (meetings_organized_with + meetings_accepted_from) / 5.0)

            internal_boost  = 0.02 if o.is_internal(email, internal_domains) else 0.0
            core_boost      = 0.08 if is_core else 0.0
            mgr_boost       = 0.12 if is_manager else 0.0

            score = (
                0.40 * direct_ratio +
                0.18 * initiated_ratio -
                0.12 * broad_ratio +
                0.10 * vol_norm +
                0.10 * meeting_sig +
                0.04 * importance_sig +
                0.03 * attach_sig +
                internal_boost + core_boost + mgr_boost
            )

            if last_sent_utc:
                try:
                    days = (dt.datetime.utcnow() - dt.datetime.fromisoformat(last_sent_utc.replace("Z",""))).days
                    if days <= 7:   score += 0.06
                    elif days <= 30: score += 0.03
                except Exception:
                    pass

            score = round(max(0.0, min(1.5, score)), 4)

        updates.append((score, email))

    conn.executemany("UPDATE contact_metrics SET score=? WHERE email=?", updates)
    conn.commit()

def build_vip_set(conn, vip_allowlist, core_team, manager_chain, vip_top_n, internal_domains):
    vip = set(e.lower() for e in (vip_allowlist | core_team | manager_chain))
    cur = conn.cursor()
    cur.execute("SELECT email, score FROM contact_metrics WHERE score IS NOT NULL ORDER BY score DESC")
    rows = cur.fetchall()
    budget = len(vip) + vip_top_n
    for (email, score) in rows:
        if len(vip) >= budget:
            break
        vip.add(email.lower())
    return vip
