import datetime as dt

from . import db
from . import outlook_utils as o

def mine_sent(conn, days_back: int, scan_all: bool, internal_domains, ignore_parts):
    start_window = dt.datetime.now() - dt.timedelta(days=days_back)
    inserted_msgs = 0
    inserted_rows = 0

    for folder, store_name in o.iter_sent_folders(scan_all):
        checkpoint_iso = db.get_checkpoint(conn, store_name)
        if checkpoint_iso:
            cp_dt = o.from_iso_local(checkpoint_iso)
            start_dt = max(start_window, cp_dt - dt.timedelta(minutes=1))
        else:
            start_dt = start_window

        items = o.restrict_by_senton(folder.Items, start_dt)
        cur = conn.cursor()
        batch = 0
        latest_seen_local = checkpoint_iso and o.from_iso_local(checkpoint_iso) or start_dt

        for item in items:
            mclass = getattr(item, "MessageClass", "")

            # Meeting requests sent (count organizer signals via calendar learning; still store message provenance)
            if mclass.startswith("IPM.Schedule.Meeting.Request"):
                sk = o.get_search_key(item)
                if sk:
                    cur.execute("""                            INSERT OR IGNORE INTO sent_items(search_key, internet_message_id, store_name, sent_on_local, sent_on_utc, subject)
                        VALUES (?, ?, ?, ?, ?, ?)
                    """, (sk, o.get_internet_message_id(item), store_name,
                          o.to_iso_local(item.SentOn), o.to_iso_utc(item.SentOn),
                          (getattr(item, "Subject", "") or "")[:500]))
                    if cur.rowcount > 0:
                        inserted_msgs += 1
                latest_seen_local = max(latest_seen_local, item.SentOn)
                batch += 1
                if batch % 500 == 0:
                    conn.commit()
                continue

            # Regular Sent mail
            if not mclass.startswith("IPM.Note"):
                continue

            _, _, rcpts = o.recipients_from_mail(item)
            if not rcpts:
                continue

            # Filter machine/system addresses
            rcpts = {r for r in rcpts if not o.is_ignored(r, ignore_parts)}
            if not rcpts:
                continue

            sk = o.get_search_key(item)
            if not sk:
                continue  # skip if no durable key

            sent_local_iso = o.to_iso_local(item.SentOn)
            sent_utc_iso   = o.to_iso_utc(item.SentOn)
            imid = o.get_internet_message_id(item)
            is_reply = 1 if o.get_in_reply_to(item) else 0
            is_direct = 1 if (len(rcpts) == 1) else 0
            is_broad  = 1 if (len(rcpts) >= 4) else 0
            has_attach = 0
            try:
                has_attach = 1 if item.Attachments.Count > 0 else 0
            except Exception:
                pass
            imp_high = 0
            try:
                from win32com.client import constants
                imp_high = 1 if item.Importance == constants.olImportanceHigh else 0
            except Exception:
                pass

            # Message row
            cur.execute("""                    INSERT OR IGNORE INTO sent_items(search_key, internet_message_id, store_name, sent_on_local, sent_on_utc, subject)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (sk, imid, store_name, sent_local_iso, sent_utc_iso, (getattr(item, "Subject", "") or "")[:500]))
            if cur.rowcount > 0:
                inserted_msgs += 1

            # Recipient features
            for r in rcpts:
                cur.execute("""                        INSERT OR IGNORE INTO sent_recipients
                    (search_key, recipient, sent_on_local, sent_on_utc, is_reply, is_direct, is_broad, importance_high, attachments_count)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (sk, r.lower(), sent_local_iso, sent_utc_iso, is_reply, is_direct, is_broad, imp_high, has_attach))
                if cur.rowcount > 0:
                    inserted_rows += 1

            latest_seen_local = max(latest_seen_local, item.SentOn)
            batch += 1
            if batch % 500 == 0:
                conn.commit()

        # Update checkpoint for this store
        db.set_checkpoint(conn, store_name, o.to_iso_local(latest_seen_local))
        conn.commit()

    return inserted_msgs, inserted_rows
