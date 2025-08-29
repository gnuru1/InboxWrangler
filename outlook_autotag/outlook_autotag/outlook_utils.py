import re
import time
import datetime as dt

import win32com.client as win32
from win32com.client import constants

# --- Outlook bootstrap (single COM session for the process) ---
outlook = win32.Dispatch("Outlook.Application")
session = outlook.Session
ns = outlook.GetNamespace("MAPI")

def user_primary_smtp() -> str:
    try:
        return session.Accounts[0].SmtpAddress.lower()
    except Exception:
        try:
            return session.CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress.lower()
        except Exception:
            return ""

ME = user_primary_smtp()

# --- Helpers: time / format ---
def now_utc_iso() -> str:
    return dt.datetime.utcnow().replace(microsecond=0).isoformat() + "Z"

def to_iso_utc(local_dt: dt.datetime) -> str:
    if local_dt.tzinfo is None:
        offset = -time.timezone
        if time.daylight and time.localtime().tm_isdst:
            offset = -time.altzone
        utc = local_dt - dt.timedelta(seconds=offset)
        return utc.replace(microsecond=0).isoformat() + "Z"
    return local_dt.astimezone(dt.timezone.utc).replace(microsecond=0).isoformat()

def to_iso_local(local_dt: dt.datetime) -> str:
    return local_dt.replace(microsecond=0).isoformat()

def from_iso_local(s: str) -> dt.datetime:
    return dt.datetime.fromisoformat(s)

def dt_us_str(d: dt.datetime) -> str:
    return d.strftime("%m/%d/%Y %I:%M %p")

# --- Helpers: addresses/domains ---
def domain_of(email: str) -> str:
    return email.split("@")[-1].lower() if "@" in email else ""

def is_internal(email: str, internal_domains) -> bool:
    return domain_of(email) in internal_domains

def is_ignored(email: str, ignores) -> bool:
    local = email.split("@")[0].lower()
    return any(sub in local for sub in ignores)

def resolve_smtp(addr_entry) -> str:
    """Return SMTP for EX/SMTP AddressEntry (handles Exchange users & DLs)."""
    try:
        if addr_entry.Type == "EX":
            try:
                exu = addr_entry.GetExchangeUser()
                if exu and exu.PrimarySmtpAddress:
                    return exu.PrimarySmtpAddress.lower()
            except Exception:
                pass
            try:
                edl = addr_entry.GetExchangeDistributionList()
                if edl and edl.PrimarySmtpAddress:
                    return edl.PrimarySmtpAddress.lower()
            except Exception:
                pass
            return (addr_entry.Address or "").lower()
        else:
            return (addr_entry.Address or "").lower()
    except Exception:
        return (getattr(addr_entry, "Address", "") or "").lower()

# --- MAPI properties ---
PROP_HEADERS   = "http://schemas.microsoft.com/mapi/proptag/0x007D001E"  # PR_TRANSPORT_MESSAGE_HEADERS
PROP_IMID      = "http://schemas.microsoft.com/mapi/proptag/0x1035001E"  # PR_INTERNET_MESSAGE_ID
PROP_SEARCHKEY = "http://schemas.microsoft.com/mapi/proptag/0x300B0102"  # PR_SEARCH_KEY (binary)

def get_headers(item) -> str:
    try:
        return item.PropertyAccessor.GetProperty(PROP_HEADERS) or ""
    except Exception:
        return ""

def parse_header_value(headers: str, name: str) -> str:
    import re
    m = re.search(rf"^{re.escape(name)}:\s*(.+)$", headers, re.IGNORECASE | re.MULTILINE)
    return m.group(1).strip() if m else ""

def get_in_reply_to(item) -> str:
    return parse_header_value(get_headers(item), "In-Reply-To").lower()

def get_internet_message_id(item) -> str:
    try:
        v = item.PropertyAccessor.GetProperty(PROP_IMID) or ""
        return v.strip().lower()
    except Exception:
        return ""

def get_search_key(item) -> bytes:
    try:
        sk = item.PropertyAccessor.GetProperty(PROP_SEARCHKEY)
        return bytes(sk) if sk is not None else None
    except Exception:
        return None

# --- Outlook object helpers ---
def iter_sent_folders(scan_all=True):
    if scan_all:
        for store in session.Stores:
            try:
                f = store.GetDefaultFolder(constants.olFolderSentMail)
                yield f, store.DisplayName
            except Exception:
                continue
    else:
        yield ns.GetDefaultFolder(constants.olFolderSentMail), "Default Store"

def recipients_from_mail(item):
    tos, ccs = set(), set()
    try:
        for r in item.Recipients:
            smtp = resolve_smtp(r.AddressEntry)
            if not smtp or smtp == ME:
                continue
            if r.Type == 1:
                tos.add(smtp)
            elif r.Type == 2:
                ccs.add(smtp)
            # r.Type==3: BCC/resources; ignore
    except Exception:
        pass
    return tos, ccs, (tos | ccs)

def restrict_by_senton(items, start_local_dt: dt.datetime):
    items.Sort("[SentOn]")
    return items.Restrict(f"[SentOn] >= '{dt_us_str(start_local_dt)}'")

def attendees_from_appt(appt):
    req, opt = [], []
    try:
        for r in appt.Recipients:
            smtp = resolve_smtp(r.AddressEntry)
            if not smtp or smtp == ME or r.Type == 3:
                continue  # skip self + rooms/resources
            if r.Type == 1: req.append(smtp)
            elif r.Type == 2: opt.append(smtp)
    except Exception:
        pass
    return req, opt

def organizer_smtp_from_name(name: str) -> str:
    if not name:
        return ""
    try:
        rec = session.CreateRecipient(name)
        if rec.Resolve():
            return resolve_smtp(rec.AddressEntry)
    except Exception:
        pass
    return ""

def appt_key(appt) -> str:
    gaid = getattr(appt, "GlobalAppointmentID", None)
    return gaid or appt.EntryID

def cat_color(color_name: str) -> int:
    # Map string to Outlook constant (fallback to DarkGray)
    c = getattr(constants, color_name, None)
    if c is None:
        c = getattr(constants, "olCategoryColorDarkGray")
    return c
