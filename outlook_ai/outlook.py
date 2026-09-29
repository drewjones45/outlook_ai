"""Typed Python API over the AppleScripts in ./applescripts.

Each method runs one script and parses its records. The field order of every
record type is defined here and must match the corresponding script.
"""

from __future__ import annotations

import html
from dataclasses import dataclass, field
from datetime import datetime

from .config import Config
from .models import Event, Message
from .osa import OutlookError, Runner, osascript_runner
from .recurrence import occurs_between, series_stub
from .models import Attendee
from .protocol import (
    decode_enum,
    format_local_datetime,
    html_to_text,
    looks_like_html,
    normalize_newlines,
    parse_attendees,
    parse_bool,
    parse_header_block,
    parse_local_datetime,
    parse_people,
    parse_person,
    parse_records,
    parse_strings,
)

# MSG record fields (after the tag), see fetch_messages.applescript
MSG_FIELDS = (
    "id", "subject", "sender", "received", "sent", "is_read", "to", "cc", "flag",
    "priority", "categories", "attachments", "headers", "replied", "truncated", "body",
)
# EVT record fields (after the tag), see list_events.applescript
EVT_FIELDS = (
    "id", "subject", "start", "end", "all_day", "location", "organizer", "calendar",
    "is_recurring", "free_busy", "attendees", "body", "ical", "is_occurrence",
    "recurrence_id", "master_id",
)
# An edited occurrence's (series id, original start): its original slot must
# not be produced again when the series is expanded.
MovedSlot = tuple[int, datetime]


def _fields(record: list[str], names: tuple[str, ...]) -> dict[str, str]:
    values = record[1:] + [""] * max(0, len(names) - (len(record) - 1))
    return dict(zip(names, values))


def _int(value: str) -> int | None:
    try:
        return int(value.strip())
    except (ValueError, AttributeError):
        return None


@dataclass
class Diagnosis:
    version: str = ""
    accounts: list[dict[str, str]] = field(default_factory=list)
    folders: list[dict[str, str]] = field(default_factory=list)
    calendars: list[dict[str, str]] = field(default_factory=list)
    inbox: str = ""
    sent: str = ""
    notes: list[str] = field(default_factory=list)


class Outlook:
    def __init__(self, cfg: Config, runner: Runner | None = None):
        self.cfg = cfg
        self.run = runner or osascript_runner

    def _call(self, script: str, *args: str) -> list[list[str]]:
        output = self.run(script, list(args), self.cfg.sync.osascript_timeout)
        return parse_records(normalize_newlines(output))

    def _folder_args(self, role: str) -> list[str]:
        o = self.cfg.outlook
        name = o.inbox_folder if role == "inbox" else o.sent_folder
        return [role, o.account, name]

    # -- read -------------------------------------------------------------------

    def diagnose(self) -> Diagnosis:
        d = Diagnosis()
        for rec in self._call("diagnose", self.cfg.outlook.account):
            tag = rec[0]
            if tag == "APP":
                d.version = rec[1] if len(rec) > 1 else ""
            elif tag == "ACCT":
                kind, name, email, default, full_name, ms_online = (rec + [""] * 7)[1:7]
                d.accounts.append(
                    {"kind": kind, "name": name, "email": email, "default": default,
                     "full_name": full_name, "microsoft_online": ms_online}
                )
            elif tag == "FOLDER":
                account, name, unread, total = (rec + [""] * 5)[1:5]
                d.folders.append({"account": account, "name": name, "unread": unread, "total": total})
            elif tag == "CAL":
                name, account = (rec + [""] * 3)[1:3]
                d.calendars.append({"name": name, "account": account})
            elif tag == "INBOX":
                d.inbox = " / ".join(x for x in rec[1:] if x)
            elif tag == "SENT":
                d.sent = " / ".join(x for x in rec[1:] if x)
            elif tag == "NOTE":
                d.notes.append(rec[1] if len(rec) > 1 else "")
        return d

    def index_folder(self, role: str) -> list[tuple[int, datetime | None, bool]]:
        """Every message in the folder as (id, received-or-sent time, is_read)."""
        rows = []
        for rec in self._call("index_folder", *self._folder_args(role)):
            if rec[0] != "IDX" or len(rec) < 4:
                continue
            mid = _int(rec[1])
            if mid is None:
                continue
            rows.append((mid, parse_local_datetime(rec[2]), parse_bool(rec[3])))
        return rows

    def fetch_messages(self, ids: list[int], role: str, body_chars: int) -> tuple[list[Message], list[str]]:
        if not ids:
            return [], []
        messages: list[Message] = []
        errors: list[str] = []
        records = self._call("fetch_messages", ",".join(str(i) for i in ids), str(body_chars))
        for rec in records:
            if rec[0] == "ERR":
                errors.append(f"message {rec[1] if len(rec) > 1 else '?'}: {rec[2] if len(rec) > 2 else ''}")
                continue
            if rec[0] != "MSG":
                continue
            f = _fields(rec, MSG_FIELDS)
            mid = _int(f["id"])
            if mid is None:
                continue
            replied = f["replied"].strip()
            body = f["body"]
            if looks_like_html(body):
                body = html_to_text(body)
            messages.append(
                Message(
                    id=mid,
                    folder=role,
                    subject=f["subject"].strip(),
                    sender=parse_person(f["sender"]),
                    to=parse_people(f["to"]),
                    cc=parse_people(f["cc"]),
                    received=parse_local_datetime(f["received"]),
                    sent=parse_local_datetime(f["sent"]),
                    is_read=parse_bool(f["is_read"]),
                    flag=decode_enum("flag", f["flag"]),
                    priority=decode_enum("priority", f["priority"]),
                    categories=parse_strings(f["categories"]),
                    attachments=parse_strings(f["attachments"]),
                    headers=parse_header_block(f["headers"]),
                    replied=None if replied == "" else parse_bool(replied),
                    body_truncated=parse_bool(f["truncated"]),
                    body=body.strip(),
                )
            )
        return messages, errors

    def list_events(self, start: datetime, end: datetime) -> tuple[list[Event], list[MovedSlot]]:
        """Events overlapping [start, end), the recurring series that occur in
        it (with iCalendar text for expansion), and the original slots of
        edited occurrences.

        Two passes keep this cheap: every series costs two Apple Events in the
        first pass, and only series that occur in the window get the full
        (attendees, notes, ...) fetch in the second.
        """
        events: list[Event] = []
        moved: list[MovedSlot] = []
        series: dict[int, tuple[str, str]] = {}
        calendars = "\n".join(self.cfg.outlook.calendars)
        records = self._call("list_events", format_local_datetime(start), format_local_datetime(end), calendars)
        for rec in records:
            if rec[0] == "XCP":
                master, when = _int(rec[1] if len(rec) > 1 else ""), parse_local_datetime(rec[2] if len(rec) > 2 else "")
                if master is not None and when is not None:
                    moved.append((master, when))
            elif rec[0] == "SER":
                sid = _int(rec[1] if len(rec) > 1 else "")
                if sid is not None and len(rec) > 3 and rec[3].strip():
                    series[sid] = (rec[2], rec[3])
            elif rec[0] == "EVT":
                ev = self._event(rec)
                if ev is not None:
                    events.append(ev)
        for ev in events:
            if ev.is_occurrence and ev.master_id is not None and ev.recurrence_id is not None:
                moved.append((ev.master_id, ev.recurrence_id))

        active = {sid for sid, (_, ical) in series.items() if occurs_between(ical, start, end)}
        have = {ev.id: ev for ev in events if ev.id in active}
        missing = sorted(active - set(have))
        if missing:
            for ev in self._event_details(missing):
                have[ev.id] = ev
        for sid in sorted(active):
            calendar, ical = series[sid]
            ev = have.get(sid) or series_stub(sid, calendar, ical)
            ev.is_recurring, ev.is_occurrence, ev.ical = True, False, ical
            if sid not in {e.id for e in events}:
                events.append(ev)
        return events, moved

    def _event_details(self, ids: list[int]) -> list[Event]:
        out = []
        batch = max(1, self.cfg.sync.batch_size)
        for i in range(0, len(ids), batch):
            for rec in self._call("event_details", ",".join(str(x) for x in ids[i : i + batch])):
                if rec[0] == "EVT":
                    ev = self._event(rec)
                    if ev is not None:
                        out.append(ev)
        return out

    @staticmethod
    def _event(rec: list[str]) -> Event | None:
        f = _fields(rec, EVT_FIELDS)
        eid = _int(f["id"])
        if eid is None:
            return None
        return Event(
            id=eid,
            subject=f["subject"].strip(),
            start=parse_local_datetime(f["start"]),
            end=parse_local_datetime(f["end"]),
            all_day=parse_bool(f["all_day"]),
            location=f["location"].strip(),
            organizer=f["organizer"].strip(),
            calendar=f["calendar"].strip(),
            is_recurring=parse_bool(f["is_recurring"]),
            free_busy=decode_enum("free_busy", f["free_busy"]),
            attendees=[
                Attendee(a.name, a.address, decode_enum("attendee_type", a.kind), decode_enum("status", a.status))
                for a in parse_attendees(f["attendees"])
            ],
            body=f["body"].strip(),
            ical=f["ical"],
            is_occurrence=parse_bool(f["is_occurrence"]),
            recurrence_id=parse_local_datetime(f["recurrence_id"]),
            master_id=_int(f["master_id"]),
        )

    # -- write (drafts only) ------------------------------------------------------

    def create_reply_draft(
        self, message_id: int, text: str, *, reply_all: bool = False, open_window: bool = False
    ) -> dict[str, str]:
        """Create a reply in Outlook's Drafts folder. Never sends anything."""
        body = text.rstrip()
        if self.cfg.drafts.signature.strip():
            body += "\n\n" + self.cfg.drafts.signature.strip()
        records = self._call(
            "create_reply_draft",
            str(message_id),
            body_to_html(body),
            body + "\n\n",
            "1" if reply_all else "0",
            "1" if open_window else "0",
            "\n".join(sorted(self.cfg.my_addresses)),
        )
        for rec in records:
            if rec[0] == "DRAFT":
                rec = rec + [""] * 6
                return {
                    "draft_id": rec[1],
                    "subject": rec[2],
                    "to": ", ".join(p.display() for p in parse_people(rec[3])),
                    "method": rec[4],
                    "folder": rec[5],
                }
        raise OutlookError("Outlook didn't confirm the draft was created.")

    def notify(self, title: str, subtitle: str, message: str) -> None:
        self._call("notify", title[:120], subtitle[:120], message[:240])


def body_to_html(text: str) -> str:
    """Plain draft text -> simple HTML paragraphs for Outlook's HTML editor."""
    paragraphs = [p for p in text.replace("\r\n", "\n").split("\n\n")]
    parts = []
    for p in paragraphs:
        escaped = "<br>".join(html.escape(line) for line in p.split("\n"))
        parts.append(f"<p>{escaped}</p>")
    return "<div>" + "".join(parts) + "</div><br>"
