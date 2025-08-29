import sqlite3
import datetime as dt

SCHEMA_SQL = """
PRAGMA journal_mode = WAL;
PRAGMA synchronous = NORMAL;
PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS runs (
  run_id INTEGER PRIMARY KEY,
  started_at_utc TEXT NOT NULL,
  finished_at_utc TEXT,
  code_version TEXT,
  dry_run INTEGER NOT NULL
);

CREATE TABLE IF NOT EXISTS contacts (
  email TEXT PRIMARY KEY,
  domain TEXT,
  is_core INTEGER NOT NULL DEFAULT 0,
  is_manager INTEGER NOT NULL DEFAULT 0,
  is_vip_seed INTEGER NOT NULL DEFAULT 0,
  last_seen_utc TEXT
);

CREATE TABLE IF NOT EXISTS sent_items (
  search_key BLOB PRIMARY KEY,
  internet_message_id TEXT,
  store_name TEXT,
  sent_on_local TEXT NOT NULL,
  sent_on_utc TEXT NOT NULL,
  subject TEXT
);
CREATE INDEX IF NOT EXISTS idx_sent_senton_local ON sent_items(sent_on_local);
CREATE INDEX IF NOT EXISTS idx_sent_imid        ON sent_items(internet_message_id);

CREATE TABLE IF NOT EXISTS sent_recipients (
  search_key BLOB,
  recipient TEXT,
  sent_on_local TEXT NOT NULL,
  sent_on_utc TEXT NOT NULL,
  is_reply INTEGER NOT NULL,
  is_direct INTEGER NOT NULL,
  is_broad  INTEGER NOT NULL,
  importance_high INTEGER NOT NULL,
  attachments_count INTEGER NOT NULL,
  PRIMARY KEY (search_key, recipient),
  FOREIGN KEY (search_key) REFERENCES sent_items(search_key) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS idx_sr_recipient ON sent_recipients(recipient);
CREATE INDEX IF NOT EXISTS idx_sr_senton    ON sent_recipients(sent_on_local);

CREATE TABLE IF NOT EXISTS appts (
  appt_entry_id TEXT PRIMARY KEY,
  global_id TEXT,
  start_utc TEXT,
  end_utc TEXT,
  subject TEXT
);

CREATE TABLE IF NOT EXISTS meetings_with_contacts (
  appt_entry_id TEXT,
  email TEXT,
  role TEXT,  -- 'organized_with' | 'accepted_from'
  PRIMARY KEY (appt_entry_id, email, role),
  FOREIGN KEY (appt_entry_id) REFERENCES appts(appt_entry_id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS idx_mwc_email ON meetings_with_contacts(email);

CREATE TABLE IF NOT EXISTS contact_metrics (
  email TEXT PRIMARY KEY,
  sent_msgs INTEGER,
  sent_replies INTEGER,
  sent_initiated INTEGER,
  direct_msgs INTEGER,
  broad_msgs INTEGER,
  importance_high INTEGER,
  attachments_sent INTEGER,
  meetings_organized_with INTEGER,
  meetings_accepted_from INTEGER,
  last_sent_utc TEXT,
  score REAL,
  FOREIGN KEY (email) REFERENCES contacts(email) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS appt_tags (
  appt_entry_id TEXT,
  run_id INTEGER,
  old_categories TEXT,
  new_categories TEXT,
  tagged_at_utc TEXT,
  PRIMARY KEY (appt_entry_id, run_id),
  FOREIGN KEY (appt_entry_id) REFERENCES appts(appt_entry_id) ON DELETE CASCADE,
  FOREIGN KEY (run_id) REFERENCES runs(run_id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS checkpoints (
  store_name TEXT PRIMARY KEY,
  last_sent_on_local TEXT
);
"""

def connect(db_path: str) -> sqlite3.Connection:
    conn = sqlite3.connect(db_path)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn

def init_schema(conn: sqlite3.Connection):
    conn.executescript(SCHEMA_SQL)
    conn.commit()

def start_run(conn: sqlite3.Connection, code_version: str, dry_run: bool) -> int:
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO runs(started_at_utc, code_version, dry_run) VALUES (?, ?, ?)",
        (dt.datetime.utcnow().replace(microsecond=0).isoformat() + "Z", code_version, 1 if dry_run else 0),
    )
    conn.commit()
    return cur.lastrowid

def finish_run(conn: sqlite3.Connection, run_id: int):
    conn.execute(
        "UPDATE runs SET finished_at_utc=? WHERE run_id=?",
        (dt.datetime.utcnow().replace(microsecond=0).isoformat() + "Z", run_id),
    )
    conn.commit()

def seed_contacts_from_config(conn: sqlite3.Connection, core_team, manager_chain, vip_allowlist):
    all_seed = set(map(str.lower, core_team | manager_chain | vip_allowlist))
    if not all_seed:
        return
    cur = conn.cursor()
    for e in all_seed:
        domain = e.split("@")[-1] if "@" in e else ""
        cur.execute("""                INSERT INTO contacts(email, domain, is_core, is_manager, is_vip_seed)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(email) DO UPDATE SET
                domain=excluded.domain,
                is_core=MAX(contacts.is_core, excluded.is_core),
                is_manager=MAX(contacts.is_manager, excluded.is_manager),
                is_vip_seed=MAX(contacts.is_vip_seed, excluded.is_vip_seed)
        """, (e, domain, 1 if e in core_team else 0, 1 if e in manager_chain else 0, 1 if e in vip_allowlist else 0))
    conn.commit()

def get_checkpoint(conn: sqlite3.Connection, store_name: str):
    cur = conn.cursor()
    cur.execute("SELECT last_sent_on_local FROM checkpoints WHERE store_name=?", (store_name,))
    row = cur.fetchone()
    return row[0] if row else None

def set_checkpoint(conn: sqlite3.Connection, store_name: str, last_sent_on_local: str):
    conn.execute("""            INSERT INTO checkpoints(store_name, last_sent_on_local)
        VALUES (?, ?)
        ON CONFLICT(store_name) DO UPDATE SET last_sent_on_local=excluded.last_sent_on_local
    """, (store_name, last_sent_on_local))
    conn.commit()

def rebuild_contact_metrics(conn: sqlite3.Connection):
    conn.execute("DELETE FROM contact_metrics")
    conn.execute("""            WITH sentagg AS (
          SELECT recipient AS email,
                 COUNT(*) AS sent_msgs,
                 SUM(is_reply) AS sent_replies,
                 SUM(1 - is_reply) AS sent_initiated,
                 SUM(is_direct) AS direct_msgs,
                 SUM(is_broad) AS broad_msgs,
                 SUM(importance_high) AS importance_high,
                 SUM(attachments_count) AS attachments_sent,
                 MAX(sent_on_utc) AS last_sent_utc
          FROM sent_recipients
          GROUP BY recipient
        ),
        meetagg AS (
          SELECT email,
                 SUM(CASE WHEN role='organized_with' THEN 1 ELSE 0 END) AS meetings_organized_with,
                 SUM(CASE WHEN role='accepted_from' THEN 1 ELSE 0 END) AS meetings_accepted_from
          FROM meetings_with_contacts
          GROUP BY email
        ),
        emails AS (
          SELECT email FROM sent_recipients GROUP BY email
          UNION
          SELECT email FROM meetings_with_contacts GROUP BY email
        )
        INSERT INTO contact_metrics(
            email, sent_msgs, sent_replies, sent_initiated, direct_msgs, broad_msgs,
            importance_high, attachments_sent, meetings_organized_with, meetings_accepted_from,
            last_sent_utc, score
        )
        SELECT e.email,
               COALESCE(sa.sent_msgs,0), COALESCE(sa.sent_replies,0), COALESCE(sa.sent_initiated,0),
               COALESCE(sa.direct_msgs,0), COALESCE(sa.broad_msgs,0),
               COALESCE(sa.importance_high,0), COALESCE(sa.attachments_sent,0),
               COALESCE(ma.meetings_organized_with,0), COALESCE(ma.meetings_accepted_from,0),
               sa.last_sent_utc, NULL
        FROM emails e
        LEFT JOIN sentagg sa ON sa.email=e.email
        LEFT JOIN meetagg ma ON ma.email=e.email
    """)
    conn.commit()

def update_contacts_last_seen(conn: sqlite3.Connection):
    cur = conn.cursor()
    cur.execute("SELECT email, last_sent_utc FROM contact_metrics WHERE last_sent_utc IS NOT NULL")
    for email, last_seen in cur.fetchall():
        domain = email.split("@")[-1] if "@" in email else ""
        conn.execute("""                INSERT INTO contacts(email, domain, last_seen_utc)
            VALUES (?, ?, ?)
            ON CONFLICT(email) DO UPDATE SET
              domain=excluded.domain,
              last_seen_utc=excluded.last_seen_utc
        """, (email, domain, last_seen))
    conn.commit()

def export_scores_csv(conn: sqlite3.Connection, csv_path: str):
    cur = conn.cursor()
    cur.execute("""            SELECT cm.email, cm.score, cm.sent_msgs, cm.sent_replies, cm.sent_initiated,
               cm.direct_msgs, cm.broad_msgs, cm.importance_high, cm.attachments_sent,
               cm.meetings_organized_with, cm.meetings_accepted_from, cm.last_sent_utc
        FROM contact_metrics cm
        ORDER BY cm.score DESC
    """)
    rows = cur.fetchall()
    import csv
    with open(csv_path, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["email","score","sent_msgs","sent_replies","sent_initiated","direct","broad",
                    "importance_high","attachments_sent","organized_with","accepted_from","last_sent_utc"])
        for r in rows:
            w.writerow(r)
