"""User configuration (TOML) and filesystem locations."""

from __future__ import annotations

import os
import sys
import tomllib
from dataclasses import dataclass, field
from pathlib import Path


def config_path() -> Path:
    if env := os.environ.get("OUTLOOK_AI_CONFIG"):
        return Path(env).expanduser()
    return Path.home() / ".config" / "outlook-ai" / "config.toml"


def data_dir() -> Path:
    if env := os.environ.get("OUTLOOK_AI_DATA_DIR"):
        return Path(env).expanduser()
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Application Support" / "outlook-ai"
    return Path.home() / ".local" / "share" / "outlook-ai"


def db_path() -> Path:
    return data_dir() / "cache.sqlite3"


@dataclass
class MeConfig:
    name: str = ""
    # Other names people use for you in email bodies ("Drew", "AJ").
    aliases: list[str] = field(default_factory=list)
    # Every address that is "you" (primary, aliases, old addresses).
    addresses: list[str] = field(default_factory=list)


@dataclass
class OutlookConfig:
    # Outlook account name as shown in Outlook > Settings > Accounts.
    # Empty means "use Outlook's default inbox / sent items".
    account: str = ""
    inbox_folder: str = "Inbox"
    sent_folder: str = "Sent Items"
    # Calendar names to include. Empty means every calendar.
    calendars: list[str] = field(default_factory=list)


@dataclass
class SyncConfig:
    days: int = 30
    max_messages: int = 1500
    body_chars: int = 20000
    batch_size: int = 50
    calendar_days_back: int = 1
    calendar_days_ahead: int = 14
    # MCP tools re-sync automatically when the cache is older than this.
    stale_minutes: int = 10
    # Seconds before an individual osascript call is abandoned.
    osascript_timeout: int = 600


@dataclass
class DraftsConfig:
    # Appended to every draft body (plain text; newlines preserved).
    signature: str = ""
    # Free-text guidance Claude follows when writing drafts in your voice.
    style: str = "Concise, warm and direct. Match the formality of the sender. No filler."
    # Open each new draft in an Outlook window so you can review it immediately.
    open_window: bool = False


@dataclass
class TriageConfig:
    # Addresses or @domains whose mail always ranks high.
    vip_senders: list[str] = field(default_factory=list)
    # Addresses or @domains to treat as automated/bulk.
    ignore_senders: list[str] = field(default_factory=list)
    # Minimum score for "needs your response" candidates.
    threshold: float = 3.0


@dataclass
class NotifyConfig:
    # Minutes before a meeting starts to post a macOS notification.
    lead_minutes: int = 10


@dataclass
class Config:
    me: MeConfig = field(default_factory=MeConfig)
    outlook: OutlookConfig = field(default_factory=OutlookConfig)
    sync: SyncConfig = field(default_factory=SyncConfig)
    drafts: DraftsConfig = field(default_factory=DraftsConfig)
    triage: TriageConfig = field(default_factory=TriageConfig)
    notify: NotifyConfig = field(default_factory=NotifyConfig)
    source: Path | None = None

    @property
    def my_addresses(self) -> set[str]:
        return {a.strip().lower() for a in self.me.addresses if a.strip()}

    @property
    def my_names(self) -> list[str]:
        names = []
        if self.me.name.strip():
            names.append(self.me.name.strip())
            names.append(self.me.name.strip().split()[0])
        names.extend(a.strip() for a in self.me.aliases if a.strip())
        seen: set[str] = set()
        out = []
        for n in names:
            if n.lower() not in seen:
                seen.add(n.lower())
                out.append(n)
        return out


_SECTIONS = {
    "me": MeConfig,
    "outlook": OutlookConfig,
    "sync": SyncConfig,
    "drafts": DraftsConfig,
    "triage": TriageConfig,
    "notify": NotifyConfig,
}


def _build(cls, raw: dict, section: str):
    known = cls.__dataclass_fields__
    unknown = sorted(set(raw) - set(known))
    if unknown:
        raise ValueError(f"Unknown setting(s) in [{section}]: {', '.join(unknown)}")
    defaults = cls()
    values = {}
    for key, value in raw.items():
        expected = type(getattr(defaults, key))
        if expected is float and isinstance(value, int):
            value = float(value)
        if not isinstance(value, expected):
            raise ValueError(
                f"[{section}] {key} should be {expected.__name__}, got {type(value).__name__}"
            )
        values[key] = value
    return cls(**values)


def load_config(path: Path | None = None) -> Config:
    path = path or config_path()
    if not path.exists():
        return Config()
    with path.open("rb") as fh:
        raw = tomllib.load(fh)
    unknown = sorted(set(raw) - set(_SECTIONS))
    if unknown:
        raise ValueError(f"Unknown section(s) in {path}: {', '.join(unknown)}")
    cfg = Config(**{name: _build(cls, raw.get(name, {}), name) for name, cls in _SECTIONS.items()})
    cfg.source = path
    return cfg


def _toml_str(value: str) -> str:
    escaped = value.replace("\\", "\\\\").replace('"', '\\"').replace("\n", "\\n")
    return f'"{escaped}"'


def _toml_list(values: list[str]) -> str:
    return "[" + ", ".join(_toml_str(v) for v in values) + "]"


def render_starter_config(name: str, addresses: list[str], account: str) -> str:
    """Starter config.toml, pre-filled with whatever `doctor` could detect."""
    first = name.split()[0] if name.strip() else ""
    return f"""# outlook-ai configuration. Edit freely; `outlook-ai doctor` validates it.

[me]
name = {_toml_str(name)}
# Other names people use for you in email bodies.
aliases = {_toml_list([first] if first else [])}
# Every address that is you. Used to tell "sent to me" from "CC'd" and to skip your own mail.
addresses = {_toml_list(addresses)}

[outlook]
# Account name from Outlook > Settings > Accounts. Leave empty to use Outlook's default inbox.
account = {_toml_str(account)}
inbox_folder = "Inbox"
sent_folder = "Sent Items"
# Calendar names to include; empty = all calendars.
calendars = []

[sync]
days = 30            # fetch full details for inbox + sent mail from the last N days
max_messages = 1500  # cap on detail fetches per sync (first sync is the slow one)
body_chars = 20000   # longest body stored per message
calendar_days_back = 1
calendar_days_ahead = 14
stale_minutes = 10   # MCP tools re-sync when the cache is older than this

[drafts]
signature = ""
style = "Concise, warm and direct. Match the formality of the sender. No filler."
open_window = false  # true = open each new draft in an Outlook window

[triage]
vip_senders = []     # e.g. ["boss@company.com", "@bigclient.com"]
ignore_senders = []  # e.g. ["@notifications.example.com"]
threshold = 3.0

[notify]
lead_minutes = 10
"""
