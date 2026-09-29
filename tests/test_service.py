import re
from datetime import timedelta

import pytest
from conftest import now

from outlook_ai.models import Message, Person
from outlook_ai.osa import OutlookError


def test_first_sync_fetches_window_then_only_new_mail(service, fake):
    report = service.syncer.sync_all()
    assert report.inbox_total == len(fake.inbox)
    assert report.inbox_fetched == len(fake.inbox)
    assert report.sent_fetched == 1
    assert report.events >= 2

    fake.inbox.append(Message(111, "inbox", "New thing", Person("Z", "z@corp.example"),
                              [Person("Pat Example", "me@corp.example")], received=now(), body="Hello?"))
    fake.calls.clear()
    report = service.syncer.sync_all()
    assert report.inbox_fetched == 1
    fetches = [args for name, args in fake.calls if name == "fetch_messages"]
    assert fetches[0][0] == "111"


def test_sync_tracks_read_state_and_removed_messages(service, fake, store):
    service.syncer.sync_all()
    fake.inbox[0].is_read = True  # read in Outlook
    removed = fake.inbox.pop()  # deleted or moved in Outlook
    service.syncer.sync_all()
    assert store.message(101).is_read
    assert all(m.id != removed.id for m in store.messages("inbox"))


def test_folder_changed_during_index_is_retried(service, fake):
    calls = {"n": 0}
    original = fake._index_folder

    def flaky(args):
        calls["n"] += 1
        if calls["n"] == 1:
            raise OutlookError("Folder changed while indexing", 1002)
        return original(args)

    fake._index_folder = flaky
    report = service.syncer.sync_all()
    assert report.inbox_total == len(fake.inbox)


def test_inbox_overview_groups_and_ids(service):
    text = service.inbox_overview(days=7)
    assert "NEEDS YOUR REPLY" in text and "AUTOMATED" in text
    assert "[101]" in text and "Q3 budget review" in text
    assert text.count("<untrusted_email_data>") == 1


def test_emails_needing_reply(service):
    text = service.emails_needing_reply(days=14)
    ids = re.findall(r"### \d+\. \[(\d+)\]", text)
    assert set(ids) == {"101", "106", "110"}
    assert "Hiring plan" in text and "Following up" in text


def test_read_email_includes_thread_and_my_reply(service):
    service.syncer.sync_all()
    text = service.read_email(104)
    assert "net 60" in text
    assert "YOU: Net 60 works for us." in text


def test_read_email_fetches_uncached_message_live(service, fake, store):
    service.syncer.sync_all()
    old = Message(150, "inbox", "Ancient", Person("Old", "old@x.com"), [], received=now() - timedelta(days=400),
                  body="From long ago")
    fake.inbox.append(old)
    assert "From long ago" in service.read_email(150)
    assert store.message(150) is not None


def test_search(service):
    text = service.search_emails("budget")
    assert "[101]" in text
    assert "No cached messages" in service.search_emails("nonexistent-term-xyz")


def test_calendar_expands_recurring_and_shows_join_link(service):
    text = service.calendar("today")
    assert "Budget sync" in text
    assert "Daily standup" in text  # series started 60 days ago
    assert "join: https://teams.microsoft.com/l/meetup-join/" in text
    assert "your response: no response" in text
    assert "Company holiday" in service.calendar("tomorrow")


def test_only_series_in_window_get_a_detail_fetch(service, fake):
    service.calendar("today")
    detail_calls = [args[0] for name, args in fake.calls if name == "event_details"]
    assert detail_calls == ["302"]  # the standup; not the series that ended in 2025


def test_series_details_failure_falls_back_to_ical_summary(service, fake):
    fake._event_details = lambda args: []
    text = service.calendar("today")
    assert "Daily standup" in text


def test_moved_occurrence_shows_at_new_time_only(service):
    text = service.calendar("tomorrow")
    assert "11:00–11:15 Daily standup" in text
    assert "09:00–09:15 Daily standup" not in text


def test_flag_and_priority_survive_the_round_trip(service, store):
    service.syncer.sync_all()
    msg = store.message(102)
    assert msg.flag == "flagged" and msg.priority == "low"
    assert msg.headers.get("list-unsubscribe")
    assert store.message(101).priority == ""


def test_calendar_outside_cached_window_queries_outlook(service, fake):
    service.calendar("today")
    fake.calls.clear()
    far = (now() + timedelta(days=60)).strftime("%Y-%m-%d")
    service.calendar(far)
    assert any(name == "list_events" for name, _ in fake.calls)


def test_meeting_prep_links_attendee_email(service):
    text = service.meeting_prep(301)
    assert "Budget sync" in text
    assert "[101]" in text  # Jane's email about the budget


def test_create_reply_draft_sends_html_and_plain_with_signature(service, fake):
    service.syncer.sync_all()
    out = service.create_reply_draft(101, "Looks good <b>to me</b>.\n\nOne question on row 4.", reply_all=True)
    assert "not sent" in out
    draft = fake.drafts[-1]
    assert draft["reply_to"] == 101 and draft["reply_all"]
    assert "&lt;b&gt;to me&lt;/b&gt;" in draft["html"]  # escaped, not injected as markup
    assert "<p>One question on row 4.</p>" in draft["html"]
    assert draft["plain"].rstrip().endswith("Pat")  # signature from config
    assert draft["me"] == ["me@corp.example"]  # so reply-all fallback can leave you off
    assert "Drafts" in out


def test_empty_draft_is_rejected(service, fake):
    assert "empty" in service.create_reply_draft(101, "   ")
    assert not fake.drafts


def test_outlook_unreachable_falls_back_to_cache_with_note(service, fake, cfg):
    service.syncer.sync_all()
    cfg.sync.stale_minutes = 0
    fake.fail["index_folder"] = OutlookError("Not authorized to send Apple events", -1743)
    text = service.inbox_overview()
    assert "couldn't refresh from Outlook" in text
    assert "Privacy & Security" in text
    assert "[101]" in text


def test_first_sync_inside_a_tool_call_is_capped(service, fake):
    text = service.inbox_overview()
    assert "first sync" in text
    idx = [args for name, args in fake.calls if name == "fetch_messages"]
    assert idx  # something was fetched right away


def test_status(service):
    service.syncer.sync_all()
    text = service.status()
    assert "Pat Example" in text and "me@corp.example" in text


@pytest.mark.parametrize("value", ["today", "tomorrow", "yesterday", "friday", "2026-09-29"])
def test_parse_day_accepts(value):
    from outlook_ai.service import parse_day

    assert parse_day(value).hour == 0


def test_parse_day_rejects_garbage():
    from outlook_ai.service import parse_day

    with pytest.raises(ValueError):
        parse_day("next blue moon")


def test_stale_mailbox_warning(service, fake):
    for m in fake.inbox:
        m.received = m.received - timedelta(days=10)
    text = service.inbox_overview(days=30)
    assert "may have stopped syncing" in text


def test_fresh_mailbox_has_no_warning(service):
    assert "may have stopped syncing" not in service.inbox_overview(days=7)


def test_read_email_reports_outlook_down_instead_of_missing(service, fake):
    fake.fail["fetch_messages"] = OutlookError("Microsoft Outlook isn't running.", -600)
    with pytest.raises(OutlookError):
        service.read_email(424242)


def test_live_fetched_message_is_not_listed_as_inbox(service, fake, store):
    service.syncer.sync_all()
    fake.inbox.append(Message(160, "inbox", "Archived thing", Person("Old", "old@x.com"), [],
                              received=now() - timedelta(days=400), body="archived"))
    service.read_email(160)
    assert store.message(160).folder == "other"
    assert all(m.id != 160 for m in store.messages("inbox"))


def test_first_sync_with_unbounded_window_still_caps_to_a_week(service, fake, store, cfg):
    cfg.sync.days = 0  # "entire folder" must not turn the in-tool first sync into a full backfill
    fake.inbox.append(Message(170, "inbox", "Month-old", Person("Old", "old@x.com"), [],
                              received=now() - timedelta(days=30), body="old"))
    assert "first sync" in service.inbox_overview()
    assert store.message(170) is None
    assert store.message(101) is not None
