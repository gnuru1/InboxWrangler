from __future__ import annotations

import os
import yaml
from dataclasses import dataclass, field
from typing import List, Optional, Dict, Any

@dataclass
class Member:
    name: str
    emails: List[str]
    aliases: List[str] = field(default_factory=list)

@dataclass
class Team:
    name: str
    mention_aliases: List[str] = field(default_factory=list)
    members: List[Member] = field(default_factory=list)

@dataclass
class Notify:
    enabled: bool = True
    to: List[str] = field(default_factory=list)
    cc: List[str] = field(default_factory=list)
    subject_template: str = "[SLA Breach] {team_name} unanswered in '{subject}'"
    body_template: str = "SLA breach."

@dataclass
class OutlookConfig:
    profile: Optional[str] = None
    store_display_name: Optional[str] = None
    inbox_folder: str = "Inbox"
    sent_folder: str = "Sent Items"
    ignore_category: str = "OutlookWatcher"

@dataclass
class Targets:
    distribution_lists: List[str] = field(default_factory=list)
    member_addresses: List[str] = field(default_factory=list)

@dataclass
class LoggingConfig:
    level: str = "INFO"

@dataclass
class AppConfig:
    database_path: str = "outlook_watcher.db"
    response_sla_minutes: int = 60
    check_interval_seconds: int = 30
    notify: Notify = field(default_factory=Notify)
    outlook: OutlookConfig = field(default_factory=OutlookConfig)
    targets: Targets = field(default_factory=Targets)
    team: Team = field(default_factory=Team)
    logging: LoggingConfig = field(default_factory=LoggingConfig)

def load_config(path: str) -> AppConfig:
    if not os.path.exists(path):
        raise FileNotFoundError(f"Config not found: {path}")
    with open(path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}
    # Deep merge onto defaults via dataclass construction
    def dict_get(d: Dict[str, Any], key: str, default):
        v = d.get(key, None)
        return v if v is not None else default

    notify = dict_get(data, "notify", {}) or {}
    outlook = dict_get(data, "outlook", {}) or {}
    targets = dict_get(data, "targets", {}) or {}
    team = dict_get(data, "team", {}) or {}
    logging_cfg = dict_get(data, "logging", {}) or {}

    cfg = AppConfig(
        database_path=data.get("database_path", "outlook_watcher.db"),
        response_sla_minutes=int(data.get("response_sla_minutes", 60)),
        check_interval_seconds=int(data.get("check_interval_seconds", 30)),
        notify=Notify(
            enabled=bool(notify.get("enabled", True)),
            to=list(notify.get("to", []) or []),
            cc=list(notify.get("cc", []) or []),
            subject_template=notify.get("subject_template", "[SLA Breach] {team_name} unanswered in '{subject}'"),
            body_template=notify.get("body_template", "SLA breach."),
        ),
        outlook=OutlookConfig(
            profile=outlook.get("profile"),
            store_display_name=outlook.get("store_display_name"),
            inbox_folder=outlook.get("inbox_folder", "Inbox"),
            sent_folder=outlook.get("sent_folder", "Sent Items"),
            ignore_category=outlook.get("ignore_category", "OutlookWatcher"),
        ),
        targets=Targets(
            distribution_lists=list(targets.get("distribution_lists", []) or []),
            member_addresses=list(targets.get("member_addresses", []) or []),
        ),
        team=Team(
            name=team.get("name", "Team"),
            mention_aliases=list(team.get("mention_aliases", []) or []),
            members=[Member(name=m.get("name", ""), emails=list(m.get("emails", []) or []), aliases=list(m.get("aliases", []) or []))
                     for m in team.get("members", []) or []],
        ),
        logging=LoggingConfig(level=logging_cfg.get("level", "INFO"))
    )
    return cfg
