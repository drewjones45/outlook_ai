"""Heuristics for "does this email need my response?".

This is a pre-filter, not the final judge: it narrows a large inbox to a
short, explained list of candidates that Claude then reads and decides on.
Every score comes with human-readable reasons.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone

from .config import Config
from .models import Message, Person

_SUBJECT_PREFIX = re.compile(
    r"^\s*(?:(?:re|fw|fwd|aw|wg|sv|vs|tr|rv|antw|ref)\s*(?:\[\d+\])?\s*:\s*|\[external\]\s*|\[ext\]\s*|external:\s*)+",
    re.IGNORECASE,
)

# Where the quoted history of a reply starts in a plain-text body.
_QUOTE_MARKERS = [
    re.compile(r"^-{2,}\s*Original Message\s*-{2,}", re.IGNORECASE | re.MULTILINE),
    re.compile(r"^_{20,}\s*$", re.MULTILINE),  # Outlook's separator line
    re.compile(r"^From:[^\n]*\n(?:[^\n]*\n){0,3}?(?:Sent|Date):", re.MULTILINE),
    re.compile(r"^On\s[^\n]{0,200}(?:\n[^\n]{0,200})?\swrote:\s*$", re.MULTILINE),
    re.compile(r"^-{2,}\s*Forwarded message\s*-{2,}", re.IGNORECASE | re.MULTILINE),
    re.compile(r"^Begin forwarded message:", re.MULTILINE),
]

_URL = re.compile(r"(?:https?://|www\.)\S+", re.IGNORECASE)

_REQUEST_PHRASES = [
    r"can you", r"could you", r"would you", r"will you", r"are you able",
    r"please", r"pls", r"let me know", r"lmk", r"what do you think",
    r"your thoughts", r"any thoughts", r"any update", r"following up", r"follow(?:ing)? up",
    r"circling back", r"gentle reminder", r"friendly reminder", r"reminder",
    r"by eod", r"by end of (?:day|week)", r"by (?:mon|tues|wednes|thurs|fri|satur|sun)day",
    r"by tomorrow", r"deadline", r"asap", r"urgent", r"time[- ]sensitive",
    r"approve", r"approval", r"sign[- ]off", r"review", r"feedback", r"your input",
    r"confirm", r"rsvp", r"need you", r"need your", r"waiting (?:on|for) (?:you|your)",
    r"action (?:required|needed|item)", r"are you available", r"do you have time",
]
_REQUEST = re.compile(r"\b(?:" + "|".join(_REQUEST_PHRASES) + r")\b", re.IGNORECASE)

_AUTOMATED_LOCALPART = re.compile(
    r"^(?:no[-_.]?reply|do[-_.]?not[-_.]?reply|notifications?|notify|alerts?|mailer-daemon|"
    r"postmaster|newsletters?|news|marketing|digest|bounces?|updates?|automated|system)\b",
    re.IGNORECASE,
)

_AUTO_SUBJECT = re.compile(
    r"^\s*(?:automatic reply|auto(?:matic)?[- ]?reply|out of (?:the )?office|undeliverable|"
    r"delivery status notification|mail delivery failed|read:|delivered:)",
    re.IGNORECASE,
)
_MEETING_RESPONSE = re.compile(
    r"^\s*(?:accepted|declined|tentative(?:ly accepted)?|canceled|cancelled|new time proposed)\s*:",
    re.IGNORECASE,
)
_MEETING_INVITE = re.compile(
    r"^\s*(?:invitation|updated invitation|meeting request|meeting invitation)\b", re.IGNORECASE
)


def thread_key(subject: str) -> str:
    return " ".join(_SUBJECT_PREFIX.sub("", subject or "").split()).lower()


def _conversation(msg: Message) -> str:
    """Group key for "same thread". Short, generic subjects ("Hi", "Question")
    collide across unrelated threads, so those are also keyed by sender."""
    key = thread_key(msg.subject)
    if len(key) < 12:
        return f"{key}|{msg.sender.email}"
    return key


def new_text(body: str) -> str:
    """The part of a body the sender actually wrote (quoted history removed)."""
    text = body or ""
    cut = len(text)
    for marker in _QUOTE_MARKERS:
        m = marker.search(text)
        if m and m.start() < cut:
            cut = m.start()
    text = text[:cut]
    lines = [ln for ln in text.splitlines() if not ln.lstrip().startswith(">")]
    return "\n".join(lines).strip()


def is_me(person: Person, cfg: Config) -> bool:
    if person.email and person.email in cfg.my_addresses:
        return True
    full_name = cfg.me.name.strip().lower()
    if full_name and (not person.email or person.is_x500):
        return person.name.strip().lower() == full_name
    return False


def _same_person(a: Person, b: Person) -> bool:
    if a.email and b.email and not a.is_x500 and not b.is_x500:
        return a.email == b.email
    return bool(a.name) and a.name.strip().lower() == b.name.strip().lower()


def _sender_matches(address: str, patterns: list[str]) -> bool:
    address = (address or "").strip().lower()
    if not address:
        return False
    for p in patterns:
        p = p.strip().lower()
        if not p:
            continue
        if p.startswith("@"):
            if address.endswith(p):
                return True
        elif address == p:
            return True
    return False


def is_automated(msg: Message, cfg: Config) -> tuple[bool, str]:
    h = msg.headers
    if "list-unsubscribe" in h or "list-id" in h:
        return True, "mailing list / newsletter"
    if h.get("precedence", "").lower() in {"bulk", "list", "junk"}:
        return True, f"precedence: {h['precedence'].lower()}"
    auto = h.get("auto-submitted", "").lower()
    if auto and auto != "no":
        return True, "auto-submitted"
    localpart = msg.sender.email.split("@")[0]
    if localpart and _AUTOMATED_LOCALPART.match(localpart):
        return True, f"automated sender ({msg.sender.email})"
    if _sender_matches(msg.sender.email, cfg.triage.ignore_senders):
        return True, "sender on ignore list"
    if _AUTO_SUBJECT.match(msg.subject or ""):
        return True, "auto-reply / delivery notice"
    if _MEETING_RESPONSE.match(msg.subject or ""):
        return True, "meeting response"
    return False, ""


def _name_pattern(names: list[str]) -> re.Pattern | None:
    names = [n for n in names if len(n) >= 2]
    if not names:
        return None
    alts = "|".join(re.escape(n) for n in sorted(names, key=len, reverse=True))
    return re.compile(rf"(?<![\w@])@?(?:{alts})\b", re.IGNORECASE)


@dataclass
class Assessment:
    message: Message
    score: float
    reasons: list[str] = field(default_factory=list)
    category: str = "fyi"  # "needs_reply", "invite", "fyi", "automated", "mine"
    addressed: str = "none"  # "to", "cc" or "none"
    mentions_me: bool = False
    replied: bool = False
    superseded: bool = False
    excerpt: str = ""


def assess(
    msg: Message,
    cfg: Config,
    *,
    now: datetime | None = None,
    replied: bool = False,
    superseded: bool = False,
) -> Assessment:
    now = now or datetime.now(timezone.utc)
    reasons: list[str] = []
    fresh = new_text(msg.body)
    a = Assessment(message=msg, score=0.0, excerpt=fresh[:1500])

    if is_me(msg.sender, cfg):
        a.category = "mine"
        a.reasons = ["sent by you"]
        return a

    if any(is_me(p, cfg) for p in msg.to):
        a.addressed = "to"
    elif any(is_me(p, cfg) for p in msg.cc):
        a.addressed = "cc"

    automated, why = is_automated(msg, cfg)
    vip = _sender_matches(msg.sender.email, cfg.triage.vip_senders)
    score = 0.0

    if a.addressed == "to":
        score += 2.5
        reasons.append("you're on the To line")
        if len(msg.to) == 1 and not msg.cc:
            score += 0.75
            reasons.append("sent only to you")
    elif a.addressed == "cc":
        score += 0.5
        reasons.append("you're CC'd")

    pattern = _name_pattern(cfg.my_names)
    if pattern and fresh:
        opening = fresh[:250]
        m = pattern.search(fresh)
        if m:
            a.mentions_me = True
            if pattern.search(opening):
                score += 2.0
                reasons.append("addresses you by name")
            elif m.group(0).startswith("@"):
                score += 2.5
                reasons.append("@-mentions you")
            else:
                score += 1.5
                reasons.append("mentions you by name")

    scrubbed = _URL.sub(" ", fresh)
    if "?" in scrubbed:
        score += 1.0
        reasons.append("asks a question")
    requests = {m.group(0).lower() for m in _REQUEST.finditer(scrubbed)}
    if requests:
        score += 1.0 if len(requests) < 3 else 1.5
        shown = ", ".join(sorted(requests)[:4])
        reasons.append(f"request language ({shown})")

    if msg.flag == "flagged":
        score += 2.0
        reasons.append("flagged")
    if msg.priority == "high":
        score += 1.0
        reasons.append("high importance")
    if vip:
        score += 2.0
        reasons.append("VIP sender")
    if not msg.is_read:
        score += 0.25
        reasons.append("unread")

    when = msg.when
    if when is not None:
        age = now - when
        if age > timedelta(days=30):
            score -= 2.0
            reasons.append("older than 30 days")
        elif age > timedelta(days=14):
            score -= 1.0
            reasons.append("older than 14 days")

    if automated and not vip:
        score -= 5.0
        reasons.append(why)
    if replied:
        a.replied = True
        reasons.append("you already replied")
    if superseded:
        a.superseded = True
        reasons.append("newer message in this thread")

    a.score = round(score, 2)
    a.reasons = reasons
    if automated and not vip:
        a.category = "automated"
    elif replied or superseded:
        a.category = "fyi"
    elif _MEETING_INVITE.match(msg.subject or ""):
        a.category = "invite"
    elif a.score >= cfg.triage.threshold:
        a.category = "needs_reply"
    return a


def assess_all(
    inbox: list[Message],
    sent: list[Message],
    cfg: Config,
    *,
    now: datetime | None = None,
) -> list[Assessment]:
    """Assess every inbox message, using Sent Items to spot threads you've answered."""
    sent_by_thread: dict[str, list[Message]] = {}
    replied_ids: set[str] = set()
    for s in sent:
        sent_by_thread.setdefault(thread_key(s.subject), []).append(s)
        if s.in_reply_to:
            replied_ids.add(s.in_reply_to)

    latest_in_thread: dict[str, datetime] = {}
    for m in inbox:
        if m.when is None or is_me(m.sender, cfg):
            continue
        conv = _conversation(m)
        if conv not in latest_in_thread or m.when > latest_in_thread[conv]:
            latest_in_thread[conv] = m.when

    out = []
    for m in inbox:
        key = thread_key(m.subject)
        replied = bool(m.replied)
        if not replied and m.message_id and m.message_id in replied_ids:
            replied = True
        if not replied and m.when is not None:
            for s in sent_by_thread.get(key, []):
                if s.when is None or s.when <= m.when:
                    continue
                if any(_same_person(m.sender, r) for r in s.to + s.cc):
                    replied = True
                    break
        conv = _conversation(m)
        superseded = (
            m.when is not None and conv in latest_in_thread and latest_in_thread[conv] > m.when
        )
        out.append(assess(m, cfg, now=now, replied=replied, superseded=superseded))
    return out


def needs_reply(assessments: list[Assessment]) -> list[Assessment]:
    picks = [a for a in assessments if a.category == "needs_reply"]
    return sorted(picks, key=lambda a: (a.score, a.message.when or datetime.min.replace(tzinfo=timezone.utc)), reverse=True)
