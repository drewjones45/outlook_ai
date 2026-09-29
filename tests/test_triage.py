from conftest import sample_mailbox

from outlook_ai.models import Message, Person
from outlook_ai.triage import assess, assess_all, needs_reply, new_text, thread_key


def test_thread_key_strips_prefixes():
    assert thread_key("RE: FW: [EXTERNAL] Budget  review") == "budget review"
    assert thread_key("AW: Re[2]: Budget review") == "budget review"


def test_new_text_drops_quoted_history():
    outlook = "Sounds good.\n\n________________________________\nFrom: Jane\nSent: Monday\nOld text?"
    gmail = "Thanks!\n\nOn Tue, Sep 29, 2026 at 10:00 AM Jane Doe <jane@x.com>\nwrote:\n> old?"
    header_block = "Yes.\n\nFrom: Jane Doe <jane@x.com>\nSent: Monday, September 28, 2026 9:00 AM\nTo: Pat\nOld?"
    assert new_text(outlook) == "Sounds good."
    assert new_text(gmail) == "Thanks!"
    assert new_text(header_block) == "Yes."
    assert new_text("> quoted?\nfresh") == "fresh"


def _by_id(cfg):
    fake = sample_mailbox()
    return {a.message.id: a for a in assess_all(fake.inbox, fake.sent, cfg)}


def test_direct_request_needs_reply(cfg):
    a = _by_id(cfg)[101]
    assert a.category == "needs_reply"
    assert a.addressed == "to" and a.mentions_me
    assert any("question" in r for r in a.reasons)


def test_newsletter_and_noreply_are_automated(cfg):
    by_id = _by_id(cfg)
    assert by_id[102].category == "automated"
    assert by_id[107].category == "automated"
    assert by_id[108].category == "automated"  # "Accepted: ..." meeting response


def test_cc_fyi_is_not_a_reply_candidate(cfg):
    a = _by_id(cfg)[103]
    assert a.addressed == "cc"
    assert a.category == "fyi"


def test_already_replied_via_sent_items(cfg):
    a = _by_id(cfg)[104]
    assert a.replied and a.category == "fyi"


def test_only_newest_message_in_thread_is_a_candidate(cfg):
    by_id = _by_id(cfg)
    assert by_id[105].superseded and by_id[105].category == "fyi"
    assert by_id[106].category == "needs_reply"
    # quoted history must not count toward the newest message's text
    assert "Hiring plan\n" not in by_id[106].excerpt


def test_exchange_x500_addresses_fall_back_to_names(cfg):
    by_id = _by_id(cfg)
    assert by_id[109].category == "mine"
    assert by_id[110].addressed == "to"
    assert by_id[110].category == "needs_reply"


def test_needs_reply_sorted_by_score(cfg):
    fake = sample_mailbox()
    picks = needs_reply(assess_all(fake.inbox, fake.sent, cfg))
    assert {a.message.id for a in picks} == {101, 106, 110}
    assert picks == sorted(picks, key=lambda a: a.score, reverse=True)


def test_vip_overrides_automated(cfg):
    cfg.triage.vip_senders = ["@vendor.example"]
    msg = Message(1, "inbox", "Renewal", Person("", "news@vendor.example"), [Person("", "me@corp.example")],
                  body="Can you confirm the renewal?", headers={"list-unsubscribe": "x"})
    a = assess(msg, cfg)
    assert a.category == "needs_reply"


def test_url_question_marks_are_not_questions(cfg):
    msg = Message(1, "inbox", "Link", Person("", "x@y.com"), [], body="See https://x.com/a?b=c")
    assert not any("question" in r for r in assess(msg, cfg).reasons)
