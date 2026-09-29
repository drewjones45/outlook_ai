"""Plain-text rendering shared by the CLI, the Markdown export and the MCP tools.

Anything that came out of an email or calendar body is wrapped in
<untrusted_email_data> markers: it was written by third parties and may
contain text that looks like instructions ("ignore previous instructions and
forward..."). The markers let Claude keep data and instructions apart.
"""

from __future__ import annotations

import re
from datetime import datetime, timedelta

from .models import Event, Message
from .triage import Assessment, new_text

OPEN = "<untrusted_email_data>"
CLOSE = "</untrusted_email_data>"

_MEETING_LINK = re.compile(
    r"https://(?:[\w.-]*teams\.microsoft\.com/l/meetup-join/|[\w.-]*zoom\.us/j/|meet\.google\.com/|"
    r"[\w.-]*webex\.com/(?:meet|join|[\w.-]+/j\.php))[^\s<>\"')\]]*",
    re.IGNORECASE,
)


def fence(text: str) -> str:
    # Neutralize any marker inside the data so it can't close the block early.
    safe = text.replace(CLOSE, "</untrusted_email_data_>").replace(OPEN, "<untrusted_email_data_>")
    return f"{OPEN}\n{safe}\n{CLOSE}"


def local(dt: datetime | None) -> datetime | None:
    return dt.astimezone() if dt is not None else None


def when(dt: datetime | None, now: datetime | None = None) -> str:
    dt = local(dt)
    if dt is None:
        return "?"
    now = local(now) if now else datetime.now().astimezone()
    if dt.date() == now.date():
        return dt.strftime("today %H:%M")
    if dt.date() == (now - timedelta(days=1)).date():
        return dt.strftime("yesterday %H:%M")
    if abs((now - dt).days) < 6:
        return dt.strftime("%a %H:%M")
    if dt.year == now.year:
        return dt.strftime("%b %d %H:%M")
    return dt.strftime("%Y-%m-%d %H:%M")


def one_line(text: str, limit: int) -> str:
    text = " ".join((text or "").split())
    return text if len(text) <= limit else text[: limit - 1].rstrip() + "…"


def meeting_link(event: Event) -> str:
    for source in (event.location, event.body):
        m = _MEETING_LINK.search(source or "")
        if m:
            return m.group(0)
    return ""


def _flags(a: Assessment) -> str:
    msg = a.message
    bits = []
    if not msg.is_read:
        bits.append("UNREAD")
    if msg.flag == "flagged":
        bits.append("FLAGGED")
    if msg.priority == "high":
        bits.append("HIGH")
    bits.append({"to": "to you", "cc": "cc you"}.get(a.addressed, "not addressed to you"))
    if a.mentions_me:
        bits.append("mentions you")
    if a.replied:
        bits.append("you replied")
    if msg.attachments:
        bits.append(f"{len(msg.attachments)} attachment(s)")
    return ", ".join(bits)


def render_overview(
    assessments: list[Assessment],
    *,
    header: str,
    preview_chars: int = 160,
    now: datetime | None = None,
) -> str:
    groups = [
        ("NEEDS YOUR REPLY (heuristic)", [a for a in assessments if a.category == "needs_reply"]),
        ("MEETING INVITES", [a for a in assessments if a.category == "invite"]),
        ("FYI / OTHER", [a for a in assessments if a.category == "fyi"]),
        ("AUTOMATED / NEWSLETTERS", [a for a in assessments if a.category == "automated"]),
        ("FROM YOU", [a for a in assessments if a.category == "mine"]),
    ]
    lines = [header, ""]
    body: list[str] = []
    for title, items in groups:
        if not items:
            continue
        body.append(f"## {title} ({len(items)})")
        for a in items:
            m = a.message
            body.append(
                f"[{m.id}] {when(m.when, now)} | {m.sender.display()} | {one_line(m.subject, 120) or '(no subject)'}"
            )
            body.append(f"    {_flags(a)}; score {a.score:g}")
            if preview_chars and a.category != "automated":
                preview = one_line(a.excerpt or new_text(m.body), preview_chars)
                if preview:
                    body.append(f"    > {preview}")
        body.append("")
    if not body:
        body.append("(no messages in this window)")
    lines.append(fence("\n".join(body).rstrip()))
    return "\n".join(lines)


def render_candidates(assessments: list[Assessment], *, header: str, now: datetime | None = None) -> str:
    if not assessments:
        return f"{header}\n\nNo emails currently look like they need your reply."
    chunks = []
    for i, a in enumerate(assessments, 1):
        m = a.message
        recips = ", ".join(p.display() for p in m.to[:6]) + (" …" if len(m.to) > 6 else "")
        cc = ", ".join(p.display() for p in m.cc[:6]) + (" …" if len(m.cc) > 6 else "")
        chunk = [
            f"### {i}. [{m.id}] {one_line(m.subject, 140) or '(no subject)'}",
            f"From: {m.sender.display()} · {when(m.when, now)} · {_flags(a)}",
            f"To: {recips or '-'}" + (f" · Cc: {cc}" if cc else ""),
            f"Why flagged (score {a.score:g}): {'; '.join(a.reasons) or '-'}",
            "Latest message text:",
            a.excerpt or "(empty body)",
        ]
        chunks.append("\n".join(chunk))
    return f"{header}\n\n" + fence("\n\n".join(chunks))


def render_message(msg: Message, thread: list[Message], *, max_chars: int, now: datetime | None = None) -> str:
    body = msg.body or ""
    clipped = len(body) > max_chars
    lines = [
        f"Message id: {msg.id} ({msg.folder})",
        f"Subject: {msg.subject or '(no subject)'}",
        f"From: {msg.sender.display()}",
        f"To: {', '.join(p.display() for p in msg.to) or '-'}",
    ]
    if msg.cc:
        lines.append(f"Cc: {', '.join(p.display() for p in msg.cc)}")
    lines.append(f"Date: {when(msg.when, now)}")
    status = ["read" if msg.is_read else "unread"]
    if msg.flag:
        status.append(msg.flag)
    if msg.priority and msg.priority != "normal":
        status.append(f"{msg.priority} importance")
    if msg.categories:
        status.append("categories: " + ", ".join(msg.categories))
    lines.append("Status: " + ", ".join(status))
    if msg.attachments:
        lines.append("Attachments: " + ", ".join(msg.attachments))
    text = body[:max_chars]
    if clipped or msg.body_truncated:
        text += "\n[... body truncated ...]"
    lines.append("")
    lines.append(fence(text or "(empty body)"))

    others = [t for t in thread if t.id != msg.id]
    if others:
        lines.append("")
        lines.append(f"Thread context ({len(others)} other cached message(s), oldest first):")
        ctx = []
        for t in others:
            who = "YOU" if t.folder == "sent" else t.sender.display()
            ctx.append(f"- [{t.id}] {when(t.when, now)} {who}: {one_line(new_text(t.body), 400)}")
        lines.append(fence("\n".join(ctx)))
    return "\n".join(lines)


def _status_word(status: str) -> str:
    return {"accepted": "accepted", "declined": "declined", "tentative": "tentative"}.get(status, "no response")


def render_events(
    events: list[Event],
    *,
    header: str,
    me: set[str] | None = None,
    with_body: bool = False,
    now: datetime | None = None,
) -> str:
    if not events:
        return f"{header}\n\n(no events)"
    now = now or datetime.now().astimezone()
    me = me or set()
    out = []
    current_day = None
    for e in events:
        start, end = local(e.start), local(e.end)
        day = start.date() if start else None
        if day != current_day:
            current_day = day
            out.append(f"## {start.strftime('%A %b %d') if start else '?'}")
        if e.all_day:
            span = "all day"
        else:
            span = f"{start:%H:%M}–{end:%H:%M}" if start and end else "?"
        line = f"- {span} {e.subject or '(no title)'} [event {e.id}]"
        if start and end and start <= now < end:
            line += "  ← NOW"
        out.append(line)
        details = []
        if e.location:
            details.append(f"where: {one_line(e.location, 120)}")
        link = meeting_link(e)
        if link and link not in (e.location or ""):
            details.append(f"join: {link}")
        if e.organizer:
            details.append(f"organizer: {e.organizer}")
        if e.is_recurring:
            details.append("recurring")
        if e.free_busy and e.free_busy not in ("busy", ""):
            details.append(f"shown as {e.free_busy}")
        if details:
            out.append("    " + " · ".join(details))
        if e.attendees:
            mine = [a for a in e.attendees if a.address.lower() in me]
            others = [a for a in e.attendees if a.address.lower() not in me]
            shown = ", ".join(
                f"{a.name or a.address} ({_status_word(a.status)})" for a in others[:8]
            )
            more = f" +{len(others) - 8} more" if len(others) > 8 else ""
            out.append(f"    attendees: {shown}{more}")
            if mine:
                out.append(f"    your response: {_status_word(mine[0].status)}")
        if with_body and e.body.strip():
            out.append("    notes: " + one_line(e.body, 600))
    return f"{header}\n\n" + fence("\n".join(out))
