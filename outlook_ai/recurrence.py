"""Expand recurring calendar series into the occurrences inside a window.

Outlook's AppleScript hands back a recurring meeting as its series "master":
one event whose start time is the *first* occurrence, possibly years ago. A
plain "events between Monday and Friday" query therefore misses every weekly
1:1 that started before Monday. The calendar script also returns each
master's iCalendar text; the RRULE/EXDATE/RECURRENCE-ID data in it is expanded
here with the `recurring-ical-events` library.
"""

from __future__ import annotations

import dataclasses
from datetime import date, datetime, time, timedelta

import icalendar
import recurring_ical_events

from .models import Event


def _aware(value: date | datetime) -> datetime:
    if isinstance(value, datetime):
        return value if value.tzinfo is not None else value.astimezone()
    return datetime.combine(value, time.min).astimezone()


def _occurrences(ical_text: str, start: datetime, end: datetime) -> list[tuple[datetime, datetime, bool]]:
    text = ical_text.strip()
    if "BEGIN:VCALENDAR" not in text.upper():
        text = f"BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//outlook-ai//EN\r\n{text}\r\nEND:VCALENDAR"
    cal = icalendar.Calendar.from_ical(text)
    out = []
    for comp in recurring_ical_events.of(cal).between(start, end):
        dtstart = comp.get("DTSTART")
        if dtstart is None:
            continue
        raw_start = dtstart.dt
        all_day = not isinstance(raw_start, datetime)
        s = _aware(raw_start)
        if comp.get("DTEND") is not None:
            e = _aware(comp["DTEND"].dt)
        elif comp.get("DURATION") is not None:
            e = s + comp["DURATION"].dt
        else:
            e = s + (timedelta(days=1) if all_day else timedelta(0))
        out.append((s, e, all_day))
    return out


def occurs_between(ical_text: str, start: datetime, end: datetime) -> bool:
    """Does this series have an occurrence in [start, end)? Unparseable data
    counts as yes, so the problem surfaces in expand_events instead of vanishing."""
    try:
        return bool(_occurrences(ical_text, start, end))
    except Exception:
        return True


def series_stub(event_id: int, calendar: str, ical_text: str) -> Event:
    """Minimal event for a series whose full details couldn't be fetched."""
    subject = location = ""
    try:
        cal = icalendar.Calendar.from_ical(ical_text)
        for comp in cal.walk("VEVENT"):
            subject = str(comp.get("SUMMARY", "") or "")
            location = str(comp.get("LOCATION", "") or "")
            break
    except Exception:
        pass
    return Event(id=event_id, subject=subject, location=location, calendar=calendar, is_recurring=True, ical=ical_text)


def _same_slot(a: Event, b: Event) -> bool:
    if (a.subject or "").strip().lower() != (b.subject or "").strip().lower():
        return False
    if a.start is None or b.start is None:
        return False
    return a.start.astimezone().date() == b.start.astimezone().date()


def _minute(dt: datetime) -> int:
    return int(dt.timestamp() // 60)


def expand_events(
    raw: list[Event],
    start: datetime,
    end: datetime,
    moved: list[tuple[int, datetime]] | tuple = (),
) -> tuple[list[Event], list[str]]:
    """Single events in the window + expanded occurrences of recurring series.

    `moved` lists (series id, original start) of occurrences that were edited
    (rescheduled or cancelled); their original slots are not generated again.
    Returns (events sorted by start, problems). A series whose iCalendar data
    can't be parsed is reported as a problem rather than silently dropped.
    """
    moved_slots = {(mid, _minute(when)) for mid, when in moved}
    series_ids = {ev.id for ev in raw if ev.is_recurring and not ev.is_occurrence and ev.ical.strip()}
    singles: list[Event] = []
    expanded: list[Event] = []
    problems: list[str] = []
    missing_ical = 0
    for ev in raw:
        overlaps = ev.start is not None and ev.end is not None and ev.start < end and ev.end > start
        if ev.is_recurring and not ev.is_occurrence and ev.ical.strip():
            try:
                slots = _occurrences(ev.ical, start, end)
            except Exception as exc:  # malformed or unsupported iCalendar data
                problems.append(f"couldn't expand recurring event '{ev.subject}': {exc}")
                if overlaps:
                    singles.append(dataclasses.replace(ev, ical=""))
                continue
            for s, e, all_day in slots:
                if (ev.id, _minute(s)) in moved_slots:
                    continue
                expanded.append(
                    dataclasses.replace(ev, start=s, end=e, all_day=ev.all_day or all_day, ical="")
                )
        elif ev.id in series_ids and not ev.is_occurrence:
            continue  # the same series master, also returned by the window query
        elif overlaps:
            singles.append(dataclasses.replace(ev, ical=""))
        elif ev.is_recurring and not ev.is_occurrence:
            missing_ical += 1
    if missing_ical:
        problems.append(
            f"{missing_ical} recurring series came back without iCalendar data, so their "
            "occurrences in this window can't be listed"
        )

    # An edited occurrence Outlook couldn't link to its series: fall back to
    # "same title, same day" to avoid listing the meeting twice.
    unlinked = [x for x in singles if x.is_occurrence and x.master_id is None]
    kept = [x for x in expanded if not any(_same_slot(x, u) for u in unlinked)]
    events = singles + kept
    seen: set[tuple[str, int]] = set()
    unique = []
    for ev in sorted(events, key=lambda e: e.start):
        # Keyed on title + start rather than id: an edited occurrence has its
        # own id but may also be present inside the series' iCalendar data.
        key = ((ev.subject or "").strip().lower(), _minute(ev.start))
        if key not in seen:
            seen.add(key)
            unique.append(ev)
    return unique, problems
