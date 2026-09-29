from __future__ import annotations

import os
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent))

os.environ["TZ"] = "America/Los_Angeles"
time.tzset()

from fake_outlook import FakeOutlook  # noqa: E402

from outlook_ai.config import Config, MeConfig  # noqa: E402
from outlook_ai.models import Attendee, Event, Message, Person  # noqa: E402
from outlook_ai.outlook import Outlook  # noqa: E402
from outlook_ai.service import Service  # noqa: E402
from outlook_ai.store import Store  # noqa: E402

ME = Person("Pat Example", "me@corp.example")
JANE = Person("Jane Doe", "jane@corp.example")
ALICE = Person("Alice Smith", "alice@partner.example")
BOB = Person("Bob Lee", "bob@corp.example")
CAROL = Person("Carol King", "carol@corp.example")


@pytest.fixture(autouse=True)
def isolated_dirs(tmp_path, monkeypatch):
    monkeypatch.setenv("OUTLOOK_AI_DATA_DIR", str(tmp_path / "data"))
    monkeypatch.setenv("OUTLOOK_AI_CONFIG", str(tmp_path / "config.toml"))


@pytest.fixture
def cfg() -> Config:
    c = Config(me=MeConfig(name="Pat Example", aliases=["Pat"], addresses=["me@corp.example"]))
    c.drafts.signature = "Pat"
    return c


def now() -> datetime:
    return datetime.now().astimezone().replace(microsecond=0)


def weekly_ical(summary: str, first: datetime, minutes: int = 30, freq: str = "DAILY", exdate: datetime | None = None) -> str:
    fmt = "%Y%m%dT%H%M%S"
    end = first + timedelta(minutes=minutes)
    lines = [
        "BEGIN:VCALENDAR",
        "PRODID:-//Microsoft Corporation//Outlook for Mac MIMEDIR//EN",
        "VERSION:2.0",
        "BEGIN:VTIMEZONE",
        "TZID:Pacific Standard Time",
        "BEGIN:STANDARD",
        "DTSTART:16011104T020000",
        "RRULE:FREQ=YEARLY;BYDAY=1SU;BYMONTH=11",
        "TZOFFSETFROM:-0700",
        "TZOFFSETTO:-0800",
        "END:STANDARD",
        "BEGIN:DAYLIGHT",
        "DTSTART:16010311T020000",
        "RRULE:FREQ=YEARLY;BYDAY=2SU;BYMONTH=3",
        "TZOFFSETFROM:-0800",
        "TZOFFSETTO:-0700",
        "END:DAYLIGHT",
        "END:VTIMEZONE",
        "BEGIN:VEVENT",
        f"UID:{summary.replace(' ', '-')}@test",
        f"SUMMARY:{summary}",
        f"DTSTART;TZID=Pacific Standard Time:{first.strftime(fmt)}",
        f"DTEND;TZID=Pacific Standard Time:{end.strftime(fmt)}",
        f"RRULE:FREQ={freq}",
    ]
    if exdate is not None:
        lines.append(f"EXDATE;TZID=Pacific Standard Time:{exdate.strftime(fmt)}")
    lines += ["END:VEVENT", "END:VCALENDAR"]
    return "\r\n".join(lines)


def sample_mailbox() -> FakeOutlook:
    n = now()
    inbox = [
        Message(101, "inbox", "Q3 budget review", JANE, [ME], [], received=n - timedelta(hours=2), is_read=False,
                body="Hi Pat,\n\nCan you review the attached budget by Friday? Let me know if the numbers look right.\n\nThanks,\nJane",
                attachments=["Q3 budget.xlsx"], headers={"message-id": "<101@corp>"}),
        Message(102, "inbox", "This week in widgets", Person("Widget News", "news@vendor.example"), [ME], [],
                received=n - timedelta(hours=5), body="Top stories... Click here to read more?",
                headers={"list-unsubscribe": "<mailto:unsub@vendor.example>"}, flag="flagged", priority="low"),
        Message(103, "inbox", "Offsite logistics", BOB, [Person("Team", "team@corp.example")], [ME],
                received=n - timedelta(days=1), body="FYI the offsite moved to Thursday. Same place."),
        Message(104, "inbox", "Contract question", ALICE, [ME], [], received=n - timedelta(days=3),
                body="Pat - quick question: are we OK with net 60 payment terms?", headers={"message-id": "<104@partner>"}),
        Message(105, "inbox", "Hiring plan for next quarter", CAROL, [ME], [], received=n - timedelta(days=5),
                body="Can you send me the hiring plan?"),
        Message(106, "inbox", "RE: Hiring plan for next quarter", CAROL, [ME], [], received=n - timedelta(days=1),
                is_read=False,
                body="Following up on this - can you send it today?\n\nFrom: Carol King\nSent: Monday\nTo: Pat\nSubject: Hiring plan\n\nCan you send me the hiring plan?"),
        Message(107, "inbox", "Your password expires soon", Person("IT", "no-reply@system.example"), [ME], [],
                received=n - timedelta(hours=8), body="Please reset your password?"),
        Message(108, "inbox", "Accepted: Weekly sync", BOB, [ME], [], received=n - timedelta(hours=9), body=""),
        Message(109, "inbox", "Launch checklist", Person("Pat Example", "/O=EXCHANGELABS/OU=EXCHANGE/CN=RECIPIENTS/CN=DAN"),
                [ME], [], received=n - timedelta(hours=3), body="Notes to self"),
        Message(110, "inbox", "Vendor onboarding", Person("Dan Wu", "/O=EXCHANGELABS/OU=EXCHANGE/CN=RECIPIENTS/CN=DANWU"),
                [Person("Pat Example", "/O=EXCHANGELABS/OU=EXCHANGE/CN=RECIPIENTS/CN=PAT")], [],
                received=n - timedelta(hours=4), body="Could you approve the vendor form in the portal?"),
    ]
    sent = [
        Message(201, "sent", "RE: Contract question", ME, [ALICE], [], sent=n - timedelta(days=2),
                body="Net 60 works for us.", headers={"in-reply-to": "<104@partner>"}),
    ]
    today = n.replace(hour=0, minute=0, second=0)
    events = [
        Event(301, "Budget sync", today + timedelta(hours=14), today + timedelta(hours=15), location="Room 5",
              organizer="Jane Doe", calendar="Calendar",
              attendees=[Attendee("Jane Doe", "jane@corp.example", "required", "accepted"),
                         Attendee("Pat Example", "me@corp.example", "required", "none")],
              body="Join: https://teams.microsoft.com/l/meetup-join/19%3ameeting_abc%40thread.v2/0?context=x"),
        Event(302, "Daily standup", today - timedelta(days=60) + timedelta(hours=9),
              today - timedelta(days=60) + timedelta(hours=9, minutes=15), calendar="Calendar", is_recurring=True,
              ical=weekly_ical("Daily standup", today - timedelta(days=60) + timedelta(hours=9), minutes=15)),
        Event(303, "Company holiday", today + timedelta(days=1), today + timedelta(days=2), all_day=True,
              calendar="Calendar"),
        # A series that ended long ago: must not cost a detail fetch.
        Event(305, "Old weekly sync", today - timedelta(days=900), today - timedelta(days=900) + timedelta(hours=1),
              calendar="Calendar", is_recurring=True,
              ical=weekly_ical("Old weekly sync", today - timedelta(days=900), minutes=60,
                               freq="WEEKLY;UNTIL=20250101T000000Z")),
        # Tomorrow's standup was moved from 09:00 to 11:00.
        Event(304, "Daily standup", today + timedelta(days=1, hours=11), today + timedelta(days=1, hours=11, minutes=15),
              calendar="Calendar", is_recurring=True, is_occurrence=True,
              recurrence_id=today + timedelta(days=1, hours=9), master_id=302),
    ]
    return FakeOutlook(inbox=inbox, sent=sent, events=events)


@pytest.fixture
def fake() -> FakeOutlook:
    return sample_mailbox()


@pytest.fixture
def store(tmp_path) -> Store:
    s = Store(tmp_path / "data" / "cache.sqlite3")
    yield s
    s.close()


@pytest.fixture
def service(cfg, store, fake) -> Service:
    return Service(cfg, store, Outlook(cfg, runner=fake))
