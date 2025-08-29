from __future__ import annotations

import datetime as dt
import logging
from typing import List, Dict, Any

import pythoncom
import win32com.client  # type: ignore

from .config import AppConfig
from .db import DB
from . import outlook_utils as ou

LOGGER = logging.getLogger(__name__)

def _iso(t) -> str:
    try:
        return t.isoformat(timespec='seconds')
    except Exception:
        return str(t)

def _date_for_restrict(d: dt.datetime) -> str:
    # Outlook Restrict expects US style: mm/dd/yyyy hh:mm AMPM
    return d.strftime("%m/%d/%Y %I:%M %p")

def is_mail_item(item) -> bool:
    try:
        return int(getattr(item, "Class", 0)) == 43 and str(getattr(item, "MessageClass", "")).startswith("IPM.Note")
    except Exception:
        return False

def _iter_items(items):
    # Items in Outlook are 1-based indexable; iteration via Items can be slow; restrict range first.
    count = items.Count
    for i in range(1, count+1):
        yield items.Item(i)

def backfill(cfg: AppConfig, db: DB, days: int = 14):
    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.gencache.EnsureDispatch("Outlook.Application")
        ns = outlook.GetNamespace("MAPI")
        if cfg.outlook.profile:
            ns.Logon(ProfileName=cfg.outlook.profile, ShowDialog=False, NewSession=False)
        # Folders
        stores = ns.Stores
        store = stores.DefaultStore
        if cfg.outlook.store_display_name:
            for i in range(1, stores.Count+1):
                s = stores.Item(i)
                if s.DisplayName == cfg.outlook.store_display_name:
                    store = s
                    break
        inbox = store.GetDefaultFolder(6)
        sent = store.GetDefaultFolder(5)

        since = dt.datetime.now() - dt.timedelta(days=days)

        # Backfill Inbox
        inbox_items = inbox.Items
        inbox_items = inbox_items.Restrict(f"[ReceivedTime] >= '{_date_for_restrict(since)}'")
        inbox_items.Sort("[ReceivedTime]")
        LOGGER.info("Backfilling Inbox from %s (%s items may match)...", since.date(), inbox_items.Count)
        patterns = ou.compile_callout_patterns(cfg.team.name, cfg.team.mention_aliases, [m.__dict__ for m in cfg.team.members])
        target_dls = cfg.targets.distribution_lists
        target_members = cfg.targets.member_addresses
        for it in _iter_items(inbox_items):
            if not is_mail_item(it):
                continue
            categories = ou.safe_getattr(it, "Categories", "") or ""
            if cfg.outlook.ignore_category and cfg.outlook.ignore_category in categories:
                continue
            conv = ou.safe_getattr(it, "ConversationID", None) or "NO-CONV"
            subject = ou.safe_getattr(it, "Subject", "") or ""
            received = ou.safe_getattr(it, "ReceivedTime", None)
            received_str = _iso(received)
            sender_email = ou.sender_smtp(it) or ""
            sender_name = ou.safe_getattr(it, "SenderName", "") or ""
            to_list, cc_list = ou.recipients_smtp_list(it)
            body_prev = ou.preview(ou.safe_getattr(it, "Body", "") or "", 700)

            if not ou.matches_target(to_list, cc_list, target_dls, target_members):
                continue

            db.upsert_thread(conv, subject, received_str, 'inbound')
            db.insert_email(entry_id=ou.safe_getattr(it, "EntryID", ""), conversation_id=conv, direction='inbound', subject=subject,
                            sender_email=sender_email, sender_name=sender_name, to_emails=", ".join(to_list), cc_emails=", ".join(cc_list),
                            time=received_str, body_preview=body_prev, categories=categories)

            mentions = ou.extract_mentions(subject + " " + ou.safe_getattr(it, "Body", "") or "", patterns)
            db.insert_mentions(ou.safe_getattr(it, "EntryID", ""), mentions)
            if mentions:
                due_at = (received + dt.timedelta(minutes=cfg.response_sla_minutes)).isoformat(timespec='seconds')
                db.create_timer(conv, ou.safe_getattr(it, "EntryID", ""), received_str, due_at)

        # Build outbound map by conversation
        sent_items = sent.Items
        sent_items = sent_items.Restrict(f"[SentOn] >= '{_date_for_restrict(since)}'")
        sent_items.Sort("[SentOn]")
        LOGGER.info("Backfilling Sent Items from %s (%s items may match)...", since.date(), sent_items.Count)
        team_emails = {e.lower() for e in (cfg.targets.member_addresses or [])}
        for it in _iter_items(sent_items):
            if not is_mail_item(it):
                continue
            categories = ou.safe_getattr(it, "Categories", "") or ""
            if cfg.outlook.ignore_category and cfg.outlook.ignore_category in categories:
                continue
            sender_email = (ou.sender_smtp(it) or "").lower()
            if sender_email not in team_emails:
                continue
            conv = ou.safe_getattr(it, "ConversationID", None) or "NO-CONV"
            subject = ou.safe_getattr(it, "Subject", "") or ""
            sent_on = ou.safe_getattr(it, "SentOn", None) or ou.safe_getattr(it, "CreationTime", None)
            sent_str = _iso(sent_on)
            to_list, cc_list = ou.recipients_smtp_list(it)
            body_prev = ou.preview(ou.safe_getattr(it, "Body", "") or "", 700)

            db.upsert_thread(conv, subject, sent_str, 'outbound')
            db.insert_email(entry_id=ou.safe_getattr(it, "EntryID", ""), conversation_id=conv, direction='outbound', subject=subject,
                            sender_email=sender_email, sender_name=ou.safe_getattr(it, "SenderName", "") or "", to_emails=", ".join(to_list), cc_emails=", ".join(cc_list),
                            time=sent_str, body_preview=body_prev, categories=categories)
            db.respond_timers_for_conversation(conv, sent_str)

        LOGGER.info("Backfill complete.")
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
