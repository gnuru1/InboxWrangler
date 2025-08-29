from __future__ import annotations

import sqlite3
from typing import Optional, Iterable, Dict, Any, List, Tuple
from dataclasses import dataclass
from contextlib import contextmanager
import logging
import time

LOGGER = logging.getLogger(__name__)

SCHEMA = '''
PRAGMA journal_mode=WAL;
CREATE TABLE IF NOT EXISTS threads (
  conversation_id TEXT PRIMARY KEY,
  subject TEXT,
  first_message_time TEXT,
  last_message_time TEXT,
  last_inbound_time TEXT,
  last_outbound_time TEXT,
  last_team_response_at TEXT,
  open_callouts INTEGER DEFAULT 0
);
CREATE TABLE IF NOT EXISTS emails (
  entry_id TEXT PRIMARY KEY,
  conversation_id TEXT,
  direction TEXT CHECK(direction IN ('inbound','outbound')),
  subject TEXT,
  sender_email TEXT,
  sender_name TEXT,
  to_emails TEXT,
  cc_emails TEXT,
  time TEXT,
  body_preview TEXT,
  categories TEXT
);
CREATE TABLE IF NOT EXISTS mentions (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  entry_id TEXT,
  mention TEXT,
  mention_type TEXT CHECK(mention_type IN ('team','member')),
  member_email TEXT
);
CREATE TABLE IF NOT EXISTS timers (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  conversation_id TEXT,
  trigger_entry_id TEXT,
  created_at TEXT,
  due_at TEXT,
  responded_at TEXT,
  breached_at TEXT,
  status TEXT CHECK(status IN ('OPEN','RESPONDED','BREACHED')) DEFAULT 'OPEN',
  notified INTEGER DEFAULT 0
);
CREATE TABLE IF NOT EXISTS notifications (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  timer_id INTEGER,
  sent_entry_id TEXT,
  sent_at TEXT
);
CREATE INDEX IF NOT EXISTS idx_timers_due ON timers(status, due_at);
CREATE INDEX IF NOT EXISTS idx_emails_conv ON emails(conversation_id);
CREATE INDEX IF NOT EXISTS idx_mentions_entry ON mentions(entry_id);
'''

class DB:
    def __init__(self, path: str):
        self.path = path
        self.conn = sqlite3.connect(self.path, check_same_thread=False)
        self.conn.execute('PRAGMA foreign_keys=ON')
        self.conn.executescript(SCHEMA)
        self.conn.commit()

    @contextmanager
    def tx(self):
        try:
            yield self.conn
            self.conn.commit()
        except Exception:
            self.conn.rollback()
            raise

    def upsert_thread(self, conversation_id: str, subject: str, message_time: str, direction: str):
        with self.tx() as c:
            c.execute('SELECT conversation_id, first_message_time, last_message_time FROM threads WHERE conversation_id=?', (conversation_id,))
            row = c.fetchone()
            if row is None:
                c.execute('INSERT INTO threads(conversation_id, subject, first_message_time, last_message_time) VALUES (?,?,?,?)',
                          (conversation_id, subject, message_time, message_time))
            else:
                # Update last_message_time and subject (subject may evolve with prefixes)
                c.execute('UPDATE threads SET subject=?, last_message_time=? WHERE conversation_id=?',
                          (subject, message_time, conversation_id))

            if direction == 'inbound':
                c.execute('UPDATE threads SET last_inbound_time=? WHERE conversation_id=?', (message_time, conversation_id))
            else:
                c.execute('UPDATE threads SET last_outbound_time=? WHERE conversation_id=?', (message_time, conversation_id))

    def insert_email(self, **kw):
        # kw keys: entry_id, conversation_id, direction, subject, sender_email, sender_name, to_emails, cc_emails, time, body_preview, categories
        with self.tx() as c:
            try:
                c.execute('INSERT OR IGNORE INTO emails(entry_id, conversation_id, direction, subject, sender_email, sender_name, to_emails, cc_emails, time, body_preview, categories) VALUES (?,?,?,?,?,?,?,?,?,?,?)',
                          (kw['entry_id'], kw['conversation_id'], kw['direction'], kw['subject'], kw['sender_email'], kw['sender_name'], kw['to_emails'], kw['cc_emails'], kw['time'], kw['body_preview'], kw['categories']))
            except Exception:
                LOGGER.exception("Failed to insert email", extra={'entry_id': kw.get('entry_id')})
                raise

    def insert_mentions(self, entry_id: str, mentions: List[Dict[str, str]]):
        if not mentions:
            return
        with self.tx() as c:
            c.executemany('INSERT INTO mentions(entry_id, mention, mention_type, member_email) VALUES (?,?,?,?)',
                          [(entry_id, m['text'], m['type'], m.get('member_email')) for m in mentions])

    def create_timer(self, conversation_id: str, trigger_entry_id: str, created_at: str, due_at: str) -> int:
        with self.tx() as c:
            c.execute('INSERT INTO timers(conversation_id, trigger_entry_id, created_at, due_at, status) VALUES (?,?,?,?,?)',
                      (conversation_id, trigger_entry_id, created_at, due_at, 'OPEN'))
            timer_id = c.execute('SELECT last_insert_rowid()').fetchone()[0]
            c.execute('UPDATE threads SET open_callouts = COALESCE(open_callouts,0) + 1 WHERE conversation_id=?', (conversation_id,))
        return timer_id

    def respond_timers_for_conversation(self, conversation_id: str, responded_at: str):
        with self.tx() as c:
            c.execute("""                UPDATE timers
                   SET responded_at = COALESCE(responded_at, ?),
                       status = CASE WHEN COALESCE(responded_at, ?) <= due_at THEN 'RESPONDED' ELSE status END
                 WHERE conversation_id = ?
                   AND status = 'OPEN'
                   AND due_at >= ?
            """, (responded_at, responded_at, conversation_id, responded_at))
            c.execute('UPDATE threads SET last_team_response_at=? WHERE conversation_id=?', (responded_at, conversation_id))
            # Also drop open_callouts count to 0 (we could be smarter per-callout; this is sufficient for SLA monitor)
            c.execute('UPDATE threads SET open_callouts = 0 WHERE conversation_id=?', (conversation_id,))

    def mark_breached(self, timer_id: int, breached_at: str):
        with self.tx() as c:
            c.execute('UPDATE timers SET breached_at=?, status="BREACHED" WHERE id=? AND status="OPEN"', (breached_at, timer_id))

    def timers_due_unnotified(self, now_str: str) -> List[Dict[str, Any]]:
        with self.tx() as c:
            rows = c.execute('''
                SELECT id, conversation_id, trigger_entry_id, created_at, due_at, responded_at, breached_at, status
                  FROM timers
                 WHERE status = 'OPEN'
                   AND due_at <= ?
            ''', (now_str,)).fetchall()
        return [dict(id=r[0], conversation_id=r[1], trigger_entry_id=r[2], created_at=r[3], due_at=r[4], responded_at=r[5], breached_at=r[6], status=r[7]) for r in rows]

    def mark_notified(self, timer_id: int, sent_entry_id: Optional[str], sent_at: str):
        with self.tx() as c:
            c.execute('INSERT INTO notifications(timer_id, sent_entry_id, sent_at) VALUES (?,?,?)', (timer_id, sent_entry_id, sent_at))
            c.execute('UPDATE timers SET notified=1 WHERE id=?', (timer_id,))

    def get_email(self, entry_id: str) -> Optional[Dict[str, Any]]:
        with self.tx() as c:
            r = c.execute('SELECT entry_id, conversation_id, direction, subject, sender_email, sender_name, to_emails, cc_emails, time, body_preview FROM emails WHERE entry_id=?', (entry_id,)).fetchone()
        if not r:
            return None
        keys = ['entry_id','conversation_id','direction','subject','sender_email','sender_name','to_emails','cc_emails','time','body_preview']
        return dict(zip(keys, r))

    def get_thread_subject(self, conversation_id: str) -> Optional[str]:
        with self.tx() as c:
            r = c.execute('SELECT subject FROM threads WHERE conversation_id=?', (conversation_id,)).fetchone()
        return r[0] if r else None

    # Simple reporting helpers
    def report_summary(self, since_days: int) -> Dict[str, Any]:
        with self.tx() as c:
            stats = {}
            stats['total_emails'] = c.execute('SELECT COUNT(*) FROM emails').fetchone()[0]
            stats['threads'] = c.execute('SELECT COUNT(*) FROM threads').fetchone()[0]
            stats['timers'] = c.execute('SELECT COUNT(*) FROM timers').fetchone()[0]
            stats['breaches'] = c.execute('SELECT COUNT(*) FROM timers WHERE status="BREACHED"').fetchone()[0]
            stats['responded'] = c.execute('SELECT COUNT(*) FROM timers WHERE status="RESPONDED"').fetchone()[0]
            stats['open'] = c.execute('SELECT COUNT(*) FROM timers WHERE status="OPEN"').fetchone()[0]
            # average TTR (time to respond) for responded timers
            row = c.execute('''
                SELECT AVG(julianday(responded_at) - julianday(created_at)) * 24 * 60
                  FROM timers
                 WHERE status = 'RESPONDED' AND responded_at IS NOT NULL
            ''').fetchone()
            stats['avg_minutes_to_respond'] = round(row[0], 1) if row and row[0] is not None else None
        return stats
