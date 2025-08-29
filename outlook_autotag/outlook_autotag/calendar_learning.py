import datetime as dt

from . import outlook_utils as o

def learn_from_calendar(conn, days_back: int, days_fwd: int):
    start = dt.datetime.now() - dt.timedelta(days=days_back)
    end   = dt.datetime.now() + dt.timedelta(days=days_fwd)
    items = o.ns.GetDefaultFolder(o.constants.olFolderCalendar).Items
    items.IncludeRecurrences = True
    items.Sort("[Start]")
    restriction = f"[Start] >= '{o.dt_us_str(start)}' AND [End] <= '{o.dt_us_str(end)}'"

    cur = conn.cursor()
    count_rows = 0

    for appt in items.Restrict(restriction):
        try:
            ak = o.appt_key(appt)
            cur.execute("""                    INSERT OR IGNORE INTO appts(appt_entry_id, global_id, start_utc, end_utc, subject)
                VALUES (?, ?, ?, ?, ?)
            """, (ak, getattr(appt, "GlobalAppointmentID", None),
                  o.to_iso_utc(appt.Start), o.to_iso_utc(appt.End),
                  (getattr(appt, "Subject", "") or "")[:500]))

            # If you're the organizer, credit attendees
            try:
                if appt.ResponseStatus == o.constants.olResponseOrganized:
                    req, opt = o.attendees_from_appt(appt)
                    for p in set(req + opt):
                        cur.execute("""                                INSERT OR IGNORE INTO meetings_with_contacts(appt_entry_id, email, role)
                            VALUES (?, ?, 'organized_with')
                        """, (ak, p.lower()))
                        count_rows += cur.rowcount
                    continue
            except Exception:
                pass

            # If you accepted someone else's meeting, credit the organizer
            try:
                if appt.ResponseStatus == o.constants.olResponseAccepted:
                    org_smtp = o.organizer_smtp_from_name((appt.Organizer or "").strip())
                    if org_smtp and org_smtp.lower() != o.ME:
                        cur.execute("""                                INSERT OR IGNORE INTO meetings_with_contacts(appt_entry_id, email, role)
                            VALUES (?, ?, 'accepted_from')
                        """, (ak, org_smtp.lower()))
                        count_rows += cur.rowcount
            except Exception:
                pass
        except Exception:
            continue

    conn.commit()
    return count_rows
