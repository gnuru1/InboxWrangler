# Centralized configuration. Edit values here.

CONFIG = {
    # --- Org specifics (EDIT ME) ---
    "INTERNAL_DOMAINS": {"acme.com", "acme.co"},  # add two+ internal domains (lowercase)
    "CORE_TEAM": {
        # "alice@acme.com", "bob@acme.com"
    },
    "MANAGER_CHAIN": {
        # "manager@acme.com", "director@acme.com", "vp@acme.com"
    },
    "VIP_ALLOWLIST": set(),  # always VIP (execs / critical partners)

    # --- Behavior & windows ---
    "SCAN_ALL_SENT_STORES": True,  # scan Sent Items in all mounted stores
    "DAYS_BACK_SENT": 180,         # lookback for Sent Items
    "DAYS_BACK_CAL": 180,          # calendar learning window (past)
    "DAYS_FWD_CAL": 60,            # calendar tagging window (future)
    "VIP_TOP_N": 15,               # add top-N scorers to VIP (in addition to config)
    "DRY_RUN": True,               # False to actually set categories

    # --- Data & audit ---
    "DB_PATH": "outlook_auto_tag.db",
    "CODE_VERSION": "v2025-08-29.1",

    # --- Categories ---
    "EXTERNAL_CATEGORY_NAME": "External",  # change to "Client" if desired
    # (name, Outlook color constant name — leave as strings)
    "CATEGORY_DEFS": [
        ("VIP",                   "olCategoryColorRed"),
        ("Manager Chain",         "olCategoryColorMaroon"),   # robust fallback across Outlook versions
        ("Core Team",             "olCategoryColorDarkGreen"),
        ("External",              "olCategoryColorDarkBlue"), # runtime-renamed to EXTERNAL_CATEGORY_NAME
        ("Focus",                 "olCategoryColorPurple"),
        ("Low-signal",            "olCategoryColorDarkGray"),
    ],

    # --- Filtering ---
    "IGNORE_LOCALPART_SUBSTR": {"no-reply", "noreply", "donotreply", "mailer-daemon"},
}
