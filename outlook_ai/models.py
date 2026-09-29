"""Plain data types shared by the bridge, cache, triage and renderers."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime


@dataclass(frozen=True)
class Person:
    name: str = ""
    address: str = ""

    @property
    def email(self) -> str:
        return self.address.strip().lower()

    @property
    def is_x500(self) -> bool:
        # Exchange sometimes reports internal people by legacy X.500 DN
        # ("/O=EXCHANGELABS/OU=...") instead of an SMTP address.
        return self.email.startswith("/o=")

    def display(self) -> str:
        if self.is_x500:
            return self.name or "(internal sender)"
        if self.name and self.address and self.name.lower() != self.address.lower():
            return f"{self.name} <{self.address}>"
        return self.name or self.address or "(unknown)"


@dataclass
class Message:
    id: int
    folder: str  # "inbox" or "sent"
    subject: str = ""
    sender: Person = field(default_factory=Person)
    to: list[Person] = field(default_factory=list)
    cc: list[Person] = field(default_factory=list)
    received: datetime | None = None
    sent: datetime | None = None
    is_read: bool = True
    flag: str = ""  # "", "flagged" or "completed"
    priority: str = ""  # "", "high", "normal" or "low"
    categories: list[str] = field(default_factory=list)
    attachments: list[str] = field(default_factory=list)
    # Selected RFC 5322 headers (lower-cased names) used for threading and bulk detection.
    headers: dict[str, str] = field(default_factory=dict)
    body: str = ""
    body_truncated: bool = False
    replied: bool | None = None  # Outlook's own "replied to" marker when available

    @property
    def when(self) -> datetime | None:
        return self.received or self.sent

    @property
    def message_id(self) -> str:
        return self.headers.get("message-id", "").strip()

    @property
    def in_reply_to(self) -> str:
        return self.headers.get("in-reply-to", "").strip()


@dataclass(frozen=True)
class Attendee:
    name: str = ""
    address: str = ""
    kind: str = ""  # "required", "optional" or "resource"
    status: str = ""  # "accepted", "declined", "tentative", "none", ""

    def display(self) -> str:
        return Person(self.name, self.address).display()


@dataclass
class Event:
    id: int
    subject: str = ""
    start: datetime | None = None
    end: datetime | None = None
    all_day: bool = False
    location: str = ""
    organizer: str = ""
    calendar: str = ""
    is_recurring: bool = False
    free_busy: str = ""
    attendees: list[Attendee] = field(default_factory=list)
    body: str = ""
    # iCalendar text for recurring series, used to expand occurrences.
    ical: str = ""
    # An edited single occurrence of a series ("exception"): which series, and
    # the start time the occurrence originally had.
    is_occurrence: bool = False
    recurrence_id: datetime | None = None
    master_id: int | None = None

    def occurrence_key(self) -> str:
        return f"{self.id}@{self.start.isoformat() if self.start else ''}"
