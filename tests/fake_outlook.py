"""A stand-in for osascript + Outlook that speaks the same wire format.

It lets the whole pipeline (bridge parsing, sync, cache, triage, service,
MCP tools) run on any OS. It does not prove the AppleScripts themselves work;
that needs a Mac with classic Outlook (see README "Verifying on your Mac").
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime

from outlook_ai.models import Attendee, Event, Message, Person
from outlook_ai.protocol import FS, GS, RS, US, format_local_datetime


def _people(people: list[Person]) -> str:
    return GS.join(FS.join([p.name, p.address]) for p in people)


def _attendees(attendees: list[Attendee]) -> str:
    return GS.join(FS.join([a.name, a.address, KIND[a.kind], STATUS[a.status]]) for a in attendees)


def _iso(dt: datetime | None) -> str:
    return format_local_datetime(dt) if dt else ""


def _headers(headers: dict[str, str]) -> str:
    # Outlook hands back the raw header block, CRLF line endings included.
    lines = ["Received: from mx.example.com by mail.example.com;", "\tTue, 29 Sep 2026 10:00:00 -0700"]
    lines += [f"{k.title()}: {v}" for k, v in headers.items()]
    return "\r\n".join(lines) + "\r\n"


# How AppleScript renders Outlook enum values when coerced to text: the raw
# form when terminology isn't available (the fake uses both forms).
FLAG = {"": "\u00abconstant ****FlNF\u00bb", "flagged": "not completed", "completed": "completed"}
PRIORITY = {"": "", "high": "priority high", "normal": "\u00abconstant ****PrNr\u00bb", "low": "priority low"}
STATUS = {"accepted": "\u00abconstant ****ASac\u00bb", "tentative": "tentatively accepted", "declined": "declined",
          "none": "not responded", "": ""}
KIND = {"required": "required attendee type", "optional": "\u00abconstant ****ATop\u00bb", "resource": "resource attendee type", "": ""}


@dataclass
class FakeOutlook:
    inbox: list[Message] = field(default_factory=list)
    sent: list[Message] = field(default_factory=list)
    events: list[Event] = field(default_factory=list)
    calls: list[tuple[str, list[str]]] = field(default_factory=list)
    drafts: list[dict] = field(default_factory=list)
    fail: dict[str, Exception] = field(default_factory=dict)

    def __call__(self, name: str, args, timeout: int) -> str:
        args = list(args)
        self.calls.append((name, args))
        if name in self.fail:
            raise self.fail[name]
        handler = getattr(self, f"_{name}")
        return RS.join(handler(args)) + "\n"

    def _diagnose(self, args):
        return [
            US.join(["APP", "16.89"]),
            US.join(["ACCT", "exchange", "Work", "me@corp.example", "1", "Pat Example", "1"]),
            US.join(["FOLDER", "Work", "Inbox", "2", str(len(self.inbox))]),
            US.join(["CAL", "Calendar", "Work"]),
            US.join(["INBOX", "Work", "Inbox"]),
        ]

    def _index_folder(self, args):
        role = args[0]
        items = self.inbox if role == "inbox" else self.sent
        out = []
        for m in items:
            t = m.received if role == "inbox" else m.sent
            out.append(US.join(["IDX", str(m.id), _iso(t), "1" if m.is_read else "0"]))
        return out

    def _fetch_messages(self, args):
        wanted = [int(x) for x in args[0].split(",") if x]
        limit = int(args[1])
        by_id = {m.id: m for m in self.inbox + self.sent}
        out = []
        for mid in wanted:
            m = by_id.get(mid)
            if m is None:
                out.append(US.join(["ERR", str(mid), "Can't get message id %d. (-1728)" % mid]))
                continue
            body = m.body[:limit] if limit > 0 else m.body
            out.append(
                US.join(
                    [
                        "MSG", str(m.id), m.subject, _people([m.sender]), _iso(m.received), _iso(m.sent),
                        "1" if m.is_read else "0", _people(m.to), _people(m.cc), FLAG[m.flag], PRIORITY[m.priority],
                        GS.join(m.categories), GS.join(m.attachments), _headers(m.headers),
                        "" if m.replied is None else ("1" if m.replied else "0"),
                        "1" if len(m.body) > len(body) else "0", body.replace("\n", "\r"),
                    ]
                )
            )
        return out

    def _list_events(self, args):
        start = datetime.fromisoformat(args[0]).astimezone()
        end = datetime.fromisoformat(args[1]).astimezone()
        wanted = [c for c in (args[2].split("\n") if len(args) > 2 else []) if c]
        out = []
        for e in self.events:
            if wanted and e.calendar not in wanted:
                continue
            overlaps = e.start < end and e.end > start
            if overlaps:
                out.append(self._evt(e))
            if e.is_recurring and not e.is_occurrence:
                out.append(US.join(["SER", str(e.id), e.calendar, e.ical]))
            if e.is_occurrence and e.recurrence_id and start <= e.recurrence_id < end:
                out.append(US.join(["XCP", str(e.master_id or ""), _iso(e.recurrence_id)]))
        return out

    def _event_details(self, args):
        wanted = {int(x) for x in args[0].split(",") if x}
        return [self._evt(e) for e in self.events if e.id in wanted]

    @staticmethod
    def _evt(e: Event) -> str:
        return US.join(
            [
                "EVT", str(e.id), e.subject, _iso(e.start), _iso(e.end), "1" if e.all_day else "0",
                e.location, e.organizer, e.calendar, "1" if e.is_recurring else "0",
                {"busy": "\u00abconstant ****eSBu\u00bb", "free": "free"}.get(e.free_busy, e.free_busy),
                _attendees(e.attendees), e.body, "",
                "1" if e.is_occurrence else "0", _iso(e.recurrence_id), str(e.master_id or ""),
            ]
        )

    def _create_reply_draft(self, args):
        mid, html_body, plain_body, reply_all, open_window, my_addresses = args
        original = next(m for m in self.inbox + self.sent if m.id == int(mid))
        draft_id = str(90000 + len(self.drafts))
        self.drafts.append(
            {"id": draft_id, "reply_to": int(mid), "html": html_body, "plain": plain_body,
             "reply_all": reply_all == "1", "open": open_window == "1", "me": my_addresses.split("\n")}
        )
        return [US.join(["DRAFT", draft_id, "RE: " + original.subject, _people([original.sender]), "reply", "Drafts"])]

    def _notify(self, args):
        return []
