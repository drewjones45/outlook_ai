from datetime import datetime, timedelta

from conftest import weekly_ical

from outlook_ai.models import Event
from outlook_ai.recurrence import expand_events


def local(*args) -> datetime:
    return datetime(*args).astimezone()


def test_weekly_series_expands_into_window_with_exdate():
    first = datetime(2026, 1, 5, 10, 0)  # a Monday
    master = Event(
        1, "1:1 with Sam", local(2026, 1, 5, 10), local(2026, 1, 5, 10, 30), is_recurring=True,
        ical=weekly_ical("1:1 with Sam", first, freq="WEEKLY;BYDAY=MO", exdate=datetime(2026, 9, 7, 10, 0)),
    )
    events, problems = expand_events([master], local(2026, 9, 1), local(2026, 10, 1))
    assert problems == []
    starts = [e.start.strftime("%m-%d %H:%M %z") for e in events]
    # Sep 7 is excluded by EXDATE; times stay 10:00 Pacific across the series.
    assert starts == ["09-14 10:00 -0700", "09-21 10:00 -0700", "09-28 10:00 -0700"]
    assert all(e.end - e.start == timedelta(minutes=30) for e in events)
    assert all(e.is_recurring and e.id == 1 for e in events)


def test_series_across_dst_change_keeps_wall_clock_time():
    master = Event(
        2, "Standup", local(2026, 10, 1, 9), local(2026, 10, 1, 9, 15), is_recurring=True,
        ical=weekly_ical("Standup", datetime(2026, 10, 1, 9, 0), minutes=15),
    )
    events, _ = expand_events([master], local(2026, 10, 31), local(2026, 11, 3))
    assert [e.start.strftime("%d %H:%M %z") for e in events] == [
        "31 09:00 -0700", "01 09:00 -0800", "02 09:00 -0800",
    ]


def _planning_master():
    return Event(
        3, "Planning", local(2026, 1, 5, 10), local(2026, 1, 5, 11), is_recurring=True,
        ical=weekly_ical("Planning", datetime(2026, 1, 5, 10, 0), minutes=60, freq="WEEKLY;BYDAY=MO"),
    )


def test_rescheduled_occurrence_replaces_series_slot():
    moved = Event(4, "Planning", local(2026, 9, 14, 15), local(2026, 9, 14, 16), is_recurring=True,
                  is_occurrence=True, recurrence_id=local(2026, 9, 14, 10), master_id=3)
    events, _ = expand_events([_planning_master(), moved], local(2026, 9, 14), local(2026, 9, 15),
                              [(3, local(2026, 9, 14, 10))])
    assert [(e.id, e.start.hour) for e in events] == [(4, 15)]


def test_occurrence_moved_out_of_window_leaves_its_slot_empty():
    # Only the XCP record (series id + original start) says the 14th moved away.
    events, _ = expand_events([_planning_master()], local(2026, 9, 14), local(2026, 9, 22),
                              [(3, local(2026, 9, 14, 10))])
    assert [e.start.day for e in events] == [21]


def test_unlinked_edited_occurrence_falls_back_to_same_day_match():
    moved = Event(4, "Planning", local(2026, 9, 14, 15), local(2026, 9, 14, 16), is_occurrence=True)
    events, _ = expand_events([_planning_master(), moved], local(2026, 9, 14), local(2026, 9, 15))
    assert [(e.id, e.start.hour) for e in events] == [(4, 15)]


def test_series_master_in_window_is_not_listed_twice():
    master = _planning_master()
    first = Event(3, "Planning", local(2026, 1, 5, 10), local(2026, 1, 5, 11), is_recurring=True)
    events, _ = expand_events([first, master], local(2026, 1, 5), local(2026, 1, 6))
    assert len(events) == 1 and events[0].start.hour == 10


def test_single_events_are_filtered_to_window_and_bad_ical_is_reported():
    inside = Event(5, "Lunch", local(2026, 9, 14, 12), local(2026, 9, 14, 13))
    outside = Event(6, "Old", local(2026, 8, 1, 12), local(2026, 8, 1, 13))
    broken = Event(7, "Broken", local(2026, 9, 14, 8), local(2026, 9, 14, 9), is_recurring=True, ical="NOT ICAL")
    no_ical = Event(8, "Mystery series", local(2025, 1, 1, 8), local(2025, 1, 1, 9), is_recurring=True)
    events, problems = expand_events([inside, outside, broken, no_ical], local(2026, 9, 14), local(2026, 9, 15))
    assert [e.id for e in events] == [7, 5]
    assert any("Broken" in p for p in problems)
    assert any("without iCalendar data" in p for p in problems)
