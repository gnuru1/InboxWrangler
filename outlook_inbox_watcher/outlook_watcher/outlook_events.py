from __future__ import annotations

import datetime as dt
import logging
import time
from typing import Callable, Optional, List

import pythoncom
import win32com.client  # type: ignore

from .config import AppConfig
from .db import DB
from . import outlook_utils as ou

LOGGER = logging.getLogger(__name__)

MAILITEM_CLASS = 43  # OlObjectClass.olMail

def is_mail_item(item) -> bool:
    try:
        return int(getattr(item, "Class", 0)) == MAILITEM_CLASS and str(getattr(item, "MessageClass", "")).startswith("IPM.Note")
    except Exception:
        return False

def _iso(t) -> str:
    if hasattr(t, "isoformat"):
        return t.isoformat(timespec='seconds')
    try:
        return str(t)
    except Exception:
        return ""

def _now_iso() -> str:
    return dt.datetime.now().isoformat(timespec='seconds')

def _to_list_str(lst: Optional[List[str]]) -> str:
    lst = lst or []
    return ", ".join(lst)

class ItemsEventSink:
    def __init__(self, process: Callable):
        self.process = process
    def OnItemAdd(self, item):
        try:
            self.process(item)
        except Exception:
            LOGGER.exception("Error in OnItemAdd handler")

class OutlookWatcher:
    def __init__(self, cfg: AppConfig, db: DB):
        self.cfg = cfg
        self.db = db
        self.outlook = None
        self.ns = None
        self.inbox_items = None
        self.sent_items = None

    def connect(self):
        # Initialize COM and connect to profile
        pythoncom.CoInitialize()
        self.outlook = win32com.client.gencache.EnsureDispatch("Outlook.Application")
        self.ns = self.outlook.GetNamespace("MAPI")
        profile = self.cfg.outlook.profile
        if profile:
            self.ns.Logon(ProfileName=profile, ShowDialog=False, NewSession=False)
        # Choose store
        inbox = self._get_folder(self.cfg.outlook.inbox_folder)
        sent = self._get_folder(self.cfg.outlook.sent_folder)
        if inbox is None or sent is None:
            raise RuntimeError("Could not find Inbox or Sent Items. Check outlook.store_display_name and folder names in config.")
        # Hook events
        self.inbox_items = win32com.client.WithEvents(inbox.Items, ItemsEventSink)
        self.sent_items = win32com.client.WithEvents(sent.Items, ItemsEventSink)
        self.inbox_items.process = self._handle_inbound
        self.sent_items.process = self._handle_outbound
        LOGGER.info("Connected. Watching Inbox='%s', Sent='%s'", self.cfg.outlook.inbox_folder, self.cfg.outlook.sent_folder)

    def _get_folder(self, folder_name: str):
        # Default store or specific store by display name
        stores = self.ns.Stores
        target_store = None
        if self.cfg.outlook.store_display_name:
            for i in range(1, stores.Count+1):
                s = stores.Item(i)
                if s.DisplayName == self.cfg.outlook.store_display_name:
                    target_store = s
                    break
        else:
            target_store = stores.DefaultStore
        if not target_store:
            return None
        root = target_store.GetDefaultFolder(6)  # OlDefaultFolders.olFolderInbox = 6
        # If they requested Inbox or a subfolder by exact name under root owner (Inbox/Sent Items live at root)
        if folder_name == "Inbox":
            return root
        # Sent Items default folder id is 5
        if folder_name == "Sent Items":
            return target_store.GetDefaultFolder(5)
        # Otherwise, try to find top-level folder by name
        folders = target_store.GetRootFolder().Folders
        try:
            return folders.Item(folder_name)
        except Exception:
            # Try under Inbox
            try:
                return root.Folders.Item(folder_name)
            except Exception:
                return None

    def _handle_inbound(self, item):
        if not is_mail_item(item):
            return
        # Category ignore
        categories = ou.safe_getattr(item, "Categories", "") or ""
        if self.cfg.outlook.ignore_category and self.cfg.outlook.ignore_category in categories:
            return
        # Basic fields
        conv_id = ou.safe_getattr(item, "ConversationID", None) or "NO-CONV"
        subject = ou.safe_getattr(item, "Subject", "") or ""
        received = ou.safe_getattr(item, "ReceivedTime", None)
        received_str = _iso(received) or _now_iso()
        sender_email = ou.sender_smtp(item) or ""
        sender_name = ou.safe_getattr(item, "SenderName", "") or ""
        to_list, cc_list = ou.recipients_smtp_list(item)
        body = ou.safe_getattr(item, "Body", "") or ""
        body_prev = ou.preview(body, 700)
        # Filter to/cc
        if not ou.matches_target(to_list, cc_list, self.cfg.targets.distribution_lists, self.cfg.targets.member_addresses):
            return
        # Persist
        self.db.upsert_thread(conv_id, subject, received_str, 'inbound')
        self.db.insert_email(entry_id=ou.safe_getattr(item, "EntryID", ""), conversation_id=conv_id, direction='inbound', subject=subject,
                             sender_email=sender_email, sender_name=sender_name, to_emails=_to_list_str(to_list), cc_emails=_to_list_str(cc_list),
                             time=received_str, body_preview=body_prev, categories=categories)

        # Mentions -> create timers
        patterns = ou.compile_callout_patterns(self.cfg.team.name, self.cfg.team.mention_aliases, [m.__dict__ for m in self.cfg.team.members])
        mentions = ou.extract_mentions(subject + " " + body, patterns)
        self.db.insert_mentions(ou.safe_getattr(item, "EntryID", ""), mentions)
        if mentions:
            # SLA timer
            created_at = received_str
            due_at = (dt.datetime.fromisoformat(received_str) + dt.timedelta(minutes=self.cfg.response_sla_minutes)).isoformat(timespec='seconds')
            timer_id = self.db.create_timer(conv_id, ou.safe_getattr(item, "EntryID", ""), created_at, due_at)
            LOGGER.info("Created SLA timer %s for conv=%s due=%s", timer_id, conv_id, due_at)

    def _handle_outbound(self, item):
        if not is_mail_item(item):
            return
        categories = ou.safe_getattr(item, "Categories", "") or ""
        if self.cfg.outlook.ignore_category and self.cfg.outlook.ignore_category in categories:
            return
        conv_id = ou.safe_getattr(item, "ConversationID", None) or "NO-CONV"
        subject = ou.safe_getattr(item, "Subject", "") or ""
        sent = ou.safe_getattr(item, "SentOn", None) or ou.safe_getattr(item, "CreationTime", None)
        sent_str = _iso(sent) or _now_iso()
        sender_email = ou.sender_smtp(item) or ""
        sender_name = ou.safe_getattr(item, "SenderName", "") or ""
        to_list, cc_list = ou.recipients_smtp_list(item)
        body = ou.safe_getattr(item, "Body", "") or ""
        body_prev = ou.preview(body, 700)

        # Only treat as a team response if the sender is a configured team member
        team_emails = {e.lower() for e in (self.cfg.targets.member_addresses or [])}
        if sender_email and sender_email.lower() in team_emails:
            self.db.upsert_thread(conv_id, subject, sent_str, 'outbound')
            self.db.insert_email(entry_id=ou.safe_getattr(item, "EntryID", ""), conversation_id=conv_id, direction='outbound', subject=subject,
                                 sender_email=sender_email, sender_name=sender_name, to_emails=", ".join(to_list), cc_emails=", ".join(cc_list),
                                 time=sent_str, body_preview=body_prev, categories=categories)
            # Mark timers as responded (for open timers with due_at >= sent time)
            self.db.respond_timers_for_conversation(conv_id, sent_str)

    def run(self, level: int = 2):
        """Run the watcher loop.

        level:
          0 = log only, compute timers but do not send notify
          1 = compute timers, mark breaches, NO email
          2 = full: send notifications
        """
        if self.outlook is None:
            self.connect()
        check_interval = int(self.cfg.check_interval_seconds or 30)
        LOGGER.info("Starting message pump (interval=%ss, level=%s)", check_interval, level)
        try:
            while True:
                self._check_timers(level=level)
                pythoncom.PumpWaitingMessages()
                time.sleep(check_interval)
        except KeyboardInterrupt:
            LOGGER.info("Watcher stopped.")
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def _check_timers(self, level: int):
        now = _now_iso()
        due = self.db.timers_due_unnotified(now)
        if not due:
            return
        for t in due:
            # Update BREACHED status
            self.db.mark_breached(t['id'], now)
            if level < 2 or not self.cfg.notify.enabled:
                LOGGER.warning("SLA BREACH (no notify): conv=%s due=%s trigger=%s", t['conversation_id'], t['due_at'], t['trigger_entry_id'])
                continue
            # Compose and send notification
            sent_entry = self._send_notify(t)
            self.db.mark_notified(t['id'], sent_entry, now)
            LOGGER.warning("SLA BREACH notified: conv=%s due=%s trigger=%s", t['conversation_id'], t['due_at'], t['trigger_entry_id'])

    def _send_notify(self, timer_row) -> Optional[str]:
        conv_id = timer_row['conversation_id']
        trig = self.db.get_email(timer_row['trigger_entry_id'])
        subject = self.db.get_thread_subject(conv_id) or "(no subject)"
        # Build template fields
        fmt = {
            'team_name': self.cfg.team.name,
            'subject': subject,
            'conversation_id': conv_id,
            'received_time': trig['time'] if trig else '',
            'due_at': timer_row['due_at'],
            'sla_minutes': self.cfg.response_sla_minutes,
            'from_name': trig['sender_name'] if trig else '',
            'from_email': trig['sender_email'] if trig else '',
            'body_preview': trig['body_preview'] if trig else '',
        }
        mail = self.outlook.CreateItem(0)
        mail.To = "; ".join(self.cfg.notify.to or [])
        if self.cfg.notify.cc:
            mail.CC = "; ".join(self.cfg.notify.cc or [])
        mail.Subject = (self.cfg.notify.subject_template or "").format(**fmt)
        mail.Body = (self.cfg.notify.body_template or "").format(**fmt)
        # Tag with ignore category
        if self.cfg.outlook.ignore_category:
            try:
                mail.Categories = (mail.Categories + ";" if mail.Categories else "") + self.cfg.outlook.ignore_category
            except Exception:
                pass
        try:
            mail.Send()
            # EntryID may not be immediately available after Send. We return None.
            return None
        except Exception:
            LOGGER.exception("Failed to send notification email.")
            return None
