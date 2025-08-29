from __future__ import annotations

import re
from typing import List, Dict, Any, Optional, Tuple
import logging

LOGGER = logging.getLogger(__name__)

SMTP_PROP = 'http://schemas.microsoft.com/mapi/proptag/0x39FE001E'

def safe_getattr(obj, name: str, default=None):
    try:
        return getattr(obj, name)
    except Exception:
        return default

def get_smtp_from_address_entry(address_entry) -> Optional[str]:
    # Try common paths to get SMTP regardless of EX/SMTP
    try:
        pa = address_entry.PropertyAccessor
        if pa:
            smtp = pa.GetProperty(SMTP_PROP)
            if smtp:
                return smtp
    except Exception:
        pass
    # Fallbacks
    try:
        t = safe_getattr(address_entry, "Type", None)
        if t == "EX":
            exu = safe_getattr(address_entry, "GetExchangeUser", None)
            if exu:
                exu = exu()
                smtp = safe_getattr(exu, "PrimarySmtpAddress", None)
                if smtp:
                    return smtp
        # SMTP type or others
        addr = safe_getattr(address_entry, "Address", None)
        if addr and '@' in addr:
            return addr
    except Exception:
        return None
    return None

def sender_smtp(mail_item) -> Optional[str]:
    try:
        pa = mail_item.PropertyAccessor
        if pa:
            smtp = pa.GetProperty(SMTP_PROP)
            if smtp:
                return smtp
    except Exception:
        pass
    try:
        se = safe_getattr(mail_item, "Sender", None)
        if se:
            ae = safe_getattr(se, "AddressEntry", None)
            if ae:
                return get_smtp_from_address_entry(ae)
    except Exception:
        pass
    addr = safe_getattr(mail_item, "SenderEmailAddress", None)
    if addr and '@' in addr:
        return addr
    return None

def recipients_smtp_list(mail_item) -> Tuple[List[str], List[str]]:
    to_list, cc_list = [], []
    try:
        for r in mail_item.Recipients:
            # MailItem.Recipients returns recipients across To/CC/BCC; we need split by Type:
            # 1 = To, 2 = CC, 3 = BCC
            rtype = getattr(r, "Type", 1)
            smtp = None
            try:
                ae = r.AddressEntry
                smtp = get_smtp_from_address_entry(ae) if ae else None
            except Exception:
                pass
            if not smtp:
                # Last resort
                smtp = getattr(r, "Address", None)
            if not smtp:
                smtp = getattr(r, "Name", None)
            if not smtp:
                continue
            if rtype == 1:
                to_list.append(smtp)
            elif rtype == 2:
                cc_list.append(smtp)
            # We ignore BCC for matching
    except Exception:
        LOGGER.exception("Failed to enumerate recipients for a message.")
    return to_list, cc_list

def normalize_text(text: str) -> str:
    if not text:
        return ""
    # Strip common quote markers, collapse whitespace
    text = re.sub(r'(?m)^>.*$', '', text)  # quoted lines
    text = re.sub(r'(?is)From:.*?Sent:.*?Subject:.*?(\r?\n){2}', '', text)  # rough "Original Message" block
    text = re.sub(r'\s+', ' ', text).strip()
    return text

def preview(text: str, max_len: int = 700) -> str:
    t = normalize_text(text)
    return t[:max_len]

def compile_callout_patterns(team_name: str, team_aliases: List[str], members: List[Dict[str, Any]]) -> Dict[str, Any]:
    aliases = [re.escape(a.strip()) for a in (team_aliases or []) if a.strip()]
    if team_name:
        aliases.append(re.escape(team_name))
    team_pat = None
    if aliases:
        team_pat = re.compile(r'\b(?:' + '|'.join(aliases) + r')\b', re.IGNORECASE)

    member_pats = []
    for m in members or []:
        toks = []
        if m.get('name'):
            toks.append(re.escape(m['name']))
        for a in m.get('aliases', []) or []:
            toks.append(re.escape(a))
        if toks:
            member_pats.append( (m, re.compile(r'\b(?:' + '|'.join(toks) + r')\b', re.IGNORECASE)) )
    return {'team': team_pat, 'members': member_pats}

def extract_mentions(text: str, patterns: Dict[str, Any]) -> list:
    out = []
    if not text:
        return out
    if patterns.get('team'):
        if patterns['team'].search(text):
            out.append({'type': 'team', 'text': patterns['team'].pattern})
    for m, pat in patterns.get('members', []):
        if pat.search(text):
            # pick first email for member if present
            mem_email = (m.get('emails') or [None])[0]
            out.append({'type': 'member', 'text': m.get('name') or 'member', 'member_email': mem_email})
    return out

def matches_target(to_list: List[str], cc_list: List[str], distribution_lists: List[str], member_addresses: List[str]) -> bool:
    target_set = {s.lower() for s in (distribution_lists or []) + (member_addresses or [])}
    if not target_set:
        return False
    all_rcpts = [*(to_list or []), *(cc_list or [])]
    for r in all_rcpts:
        if r and r.lower() in target_set:
            return True
    # Also allow simple name contains if SMTP not resolved
    for r in all_rcpts:
        rl = r.lower()
        for t in target_set:
            if '@' not in t and t in rl:
                return True
    return False
