"""Operations shared by the MCP server and the CLI.

Every method returns ready-to-read text; the MCP tools and CLI commands are
thin wrappers around these.
"""

from __future__ import annotations

import re
from datetime import datetime, timedelta

from .config import Config, db_path, load_config
from .models import Event, Message
from .osa import OutlookError
from .outlook import Outlook
from .recurrence import expand_events
from .render import (
    fence,
    one_line,
    render_candidates,
    render_events,
    render_message,
    render_overview,
    when,
)
from .store import Store
from .sync import Syncer
from .triage import assess_all, needs_reply, thread_key

_WEEKDAYS = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]


def parse_day(value: str, now: datetime | None = None) -> datetime:
    """'today', 'tomorrow', 'yesterday', a weekday name or YYYY-MM-DD -> local midnight."""
    now = (now or datetime.now()).astimezone()
    today = now.replace(hour=0, minute=0, second=0, microsecond=0)
    v = (value or "today").strip().lower()
    if v in ("", "today", "now"):
        return today
    if v == "tomorrow":
        return today + timedelta(days=1)
    if v == "yesterday":
        return today - timedelta(days=1)
    if v in _WEEKDAYS:
        delta = (_WEEKDAYS.index(v) - today.weekday()) % 7
        return today + timedelta(days=delta)
    try:
        return datetime.strptime(v, "%Y-%m-%d").astimezone()
    except ValueError as exc:
        raise ValueError(f"Unrecognized date '{value}'. Use today, tomorrow, a weekday or YYYY-MM-DD.") from exc


def _clamp(value: int, low: int, high: int) -> int:
    return max(low, min(high, int(value)))


class Service:
    def __init__(self, cfg: Config | None = None, store: Store | None = None, outlook: Outlook | None = None):
        self.cfg = cfg or load_config()
        self.store = store or Store(db_path())
        self.outlook = outlook or Outlook(self.cfg)
        self.syncer = Syncer(self.cfg, self.store, self.outlook)

    # -- helpers ------------------------------------------------------------------

    def _now(self) -> datetime:
        return datetime.now().astimezone()

    def _freshness(self, what: str) -> str:
        last = self.store.last_sync(what)
        if last is None:
            return "never synced"
        minutes = int((self._now().timestamp() - last) / 60)
        return "synced just now" if minutes < 1 else f"synced {minutes} min ago"

    def _assessments(self, days: int):
        since = self._now() - timedelta(days=days)
        inbox = self.store.messages("inbox", since)
        # Replies can come later than the message, so look at all cached sent mail from the window on.
        sent = self.store.messages("sent", since)
        return assess_all(inbox, sent, self.cfg, now=self._now())

    def _with_note(self, note: str, text: str) -> str:
        notes = [n for n in (note, self._stale_mailbox_note()) if n]
        return "\n".join(notes) + f"\n\n{text}" if notes else text

    def _stale_mailbox_note(self) -> str:
        """Warn when Outlook itself seems to have stopped receiving mail.

        Legacy Outlook for Mac talks to Exchange Online over EWS, which Microsoft
        began retiring on 2026-10-01; after that the app (and so this cache)
        silently stops getting new mail. Old data must not pass as current.
        """
        newest = self.store.newest_indexed("inbox")
        if newest is None:
            return ""
        age = self._now() - newest
        if age < timedelta(hours=72):
            return ""
        return (
            f"Warning: the newest message in Outlook's inbox arrived {age.days} day(s) ago. If you normally "
            "get mail daily, Outlook may have stopped syncing (Legacy Outlook for Mac can't connect to "
            "Exchange Online after Microsoft's October 2026 EWS retirement). Treat this data as possibly stale."
        )

    # -- mail -----------------------------------------------------------------------

    def inbox_overview(
        self, days: int = 7, unread_only: bool = False, include_automated: bool = True, limit: int = 100
    ) -> str:
        days, limit = _clamp(days, 1, 365), _clamp(limit, 1, 500)
        note = self.syncer.ensure_fresh(mail=True)
        items = self._assessments(days)
        if unread_only:
            items = [a for a in items if not a.message.is_read]
        if not include_automated:
            items = [a for a in items if a.category != "automated"]
        order = {"needs_reply": 0, "invite": 1, "fyi": 2, "automated": 3, "mine": 4}
        items.sort(key=lambda a: (order.get(a.category, 9), -(a.message.when.timestamp() if a.message.when else 0)))
        shown = items[:limit]
        total, unread = self.store.folder_counts("inbox")
        header = (
            f"Inbox: {total} messages in folder ({unread} unread) · {len(items)} in the last {days} day(s)"
            f"{' (unread only)' if unread_only else ''}, showing {len(shown)} · {self._freshness('mail')}.\n"
            "Ids in [brackets] work with read_email and create_reply_draft. Categories are heuristic."
        )
        return self._with_note(note, render_overview(shown, header=header, now=self._now()))

    def emails_needing_reply(self, days: int = 14, limit: int = 15) -> str:
        days, limit = _clamp(days, 1, 365), _clamp(limit, 1, 100)
        note = self.syncer.ensure_fresh(mail=True)
        picks = needs_reply(self._assessments(days))[:limit]
        header = (
            f"{len(picks)} email(s) from the last {days} day(s) likely need your reply "
            f"(heuristic pre-filter; read each with read_email and use judgment) · {self._freshness('mail')}."
        )
        return self._with_note(note, render_candidates(picks, header=header, now=self._now()))

    def _get_message(self, message_id: int) -> Message | None:
        msg = self.store.message(message_id)
        if msg is not None:
            return msg
        # Not cached (outside the sync window, or in another folder): fetch it
        # straight from Outlook. Filed as "other" so it never shows up in inbox
        # listings; a later sync relabels it if it is in the inbox after all.
        found, _ = self.outlook.fetch_messages([message_id], "other", self.cfg.sync.body_chars)
        if found:
            self.store.upsert_messages(found)
            return found[0]
        return None

    def read_email(self, message_id: int, max_chars: int = 12000) -> str:
        msg = self._get_message(int(message_id))
        if msg is None:
            return f"No message with id {message_id} in the cache or in Outlook."
        thread = self.store.by_thread_key(thread_key(msg.subject), thread_key)
        return render_message(msg, thread, max_chars=_clamp(max_chars, 500, 100000), now=self._now())

    def search_emails(self, query: str, days: int = 90, limit: int = 20) -> str:
        days, limit = _clamp(days, 1, 3650), _clamp(limit, 1, 100)
        note = self.syncer.ensure_fresh(mail=True)
        since = self._now() - timedelta(days=days)
        hits = self.store.search(query, since, limit)
        if not hits:
            return self._with_note(
                note,
                f"No cached messages from the last {days} day(s) match '{query}'. "
                "Only synced mail is searchable; widen the window with `outlook-ai sync --days N`.",
            )
        lines = [
            f"[{m.id}] {when(m.when, self._now())} | {'YOU -> ' + ', '.join(p.display() for p in m.to[:3]) if m.folder == 'sent' else m.sender.display()}"
            f" | {one_line(m.subject, 120)}\n    > {one_line(m.body, 200)}"
            for m in hits
        ]
        header = f"{len(hits)} match(es) for '{query}' in the last {days} day(s):"
        return self._with_note(note, f"{header}\n\n" + fence("\n".join(lines)))

    def create_reply_draft(self, message_id: int, body: str, reply_all: bool = False) -> str:
        if not body.strip():
            return "Draft body is empty; nothing created."
        # The script itself checks the id against Outlook, so an uncached message is fine.
        msg = self.store.message(int(message_id)) or Message(int(message_id), "other")
        result = self.outlook.create_reply_draft(
            int(message_id), body, reply_all=reply_all, open_window=self.cfg.drafts.open_window
        )
        folder = result.get("folder") or "Drafts"
        where = f"saved in Outlook's '{folder}' folder"
        if self.cfg.drafts.open_window:
            where += " and opened in a window"
        text = (
            f"Draft reply {where} (not sent). Draft id {result.get('draft_id') or '?'} · "
            f"subject: {result.get('subject') or msg.subject} · to: {result.get('to') or msg.sender.display()}"
        )
        if result.get("method", "reply") != "reply":
            text += f"\nNote: created as a {result['method']}; it won't be threaded with the original."
        return text

    # -- calendar -------------------------------------------------------------------

    def _events(self, start: datetime, end: datetime) -> tuple[list[Event], str]:
        note = self.syncer.ensure_fresh(calendar=True)
        window = self.store.calendar_window()
        if window and window[0] <= start and end <= window[1]:
            return self.store.events(start, end), note
        try:
            raw, moved = self.outlook.list_events(start, end)
            events, problems = expand_events(raw, start, end, moved)
        except OutlookError as exc:
            return self.store.events(start, end), f"Note: range is outside the cached window and Outlook is unreachable ({exc})."
        wanted = {c.lower() for c in self.cfg.outlook.calendars}
        if wanted:
            events = [e for e in events if e.calendar.lower() in wanted]
        extra = "; ".join(problems)
        return events, " ".join(x for x in (note, extra) if x)

    def calendar(self, date: str = "today", days: int = 1, include_notes: bool = False) -> str:
        start = parse_day(date, self._now())
        days = _clamp(days, 1, 62)
        end = start + timedelta(days=days)
        events, note = self._events(start, end)
        span = start.strftime("%a %b %d") + (f" – {(end - timedelta(days=1)).strftime('%a %b %d')}" if days > 1 else "")
        header = f"Calendar {span}: {len(events)} event(s) · {self._freshness('calendar')}."
        text = render_events(events, header=header, me=self.cfg.my_addresses, with_body=include_notes, now=self._now())
        return self._with_note(note, text)

    def meeting_prep(self, event_id: int, date: str = "") -> str:
        now = self._now()
        mail_note = self.syncer.ensure_fresh(mail=True)
        if date:
            start = parse_day(date, now)
            candidates, note = self._events(start, start + timedelta(days=1))
        else:
            candidates, note = self._events(now - timedelta(hours=12), now + timedelta(days=self.cfg.sync.calendar_days_ahead))
        note = " ".join(x for x in (mail_note, note) if x)
        matches = [e for e in candidates if e.id == int(event_id)]
        if not matches:
            return self._with_note(note, f"No event {event_id} found{' on ' + date if date else ' in the upcoming window'}.")
        upcoming = [e for e in matches if e.end and e.end >= now]
        event = (upcoming or matches)[0]
        parts = [render_events([event], header="Meeting:", me=self.cfg.my_addresses, with_body=True, now=now)]

        others = {a.address.lower() for a in event.attendees if a.address} - self.cfg.my_addresses
        since = now - timedelta(days=21)
        related: dict[int, Message] = {m.id: m for m in self.store.from_addresses(others, since, 15)}
        words = sorted({w for w in re.findall(r"[\w'-]{4,}", event.subject or "")}, key=len, reverse=True)[:2]
        if words:
            for m in self.store.search(" ".join(words), since, 5):
                related.setdefault(m.id, m)
        if related:
            rows = sorted(related.values(), key=lambda m: m.when or now, reverse=True)
            lines = [
                f"[{m.id}] {when(m.when, now)} | {m.sender.display() if m.folder == 'inbox' else 'YOU'} | "
                f"{one_line(m.subject, 100)}\n    > {one_line(m.body, 240)}"
                for m in rows
            ]
            parts.append(
                f"Recent email (last 21 days) from attendees or about '{event.subject}':\n" + fence("\n".join(lines))
            )
        else:
            parts.append("No recent cached email from these attendees or about this topic.")
        return self._with_note(note, "\n\n".join(parts))

    # -- maintenance ------------------------------------------------------------------

    def sync(self, days: int = 0) -> str:
        report = self.syncer.sync_all(days=days or None)
        return "Synced from Outlook: " + report.summary()

    def status(self) -> str:
        stats = self.store.stats()
        total, unread = self.store.folder_counts("inbox")
        me = self.cfg.me
        lines = [
            f"User: {me.name or '(name not set)'} · addresses: {', '.join(me.addresses) or '(none set)'}",
            f"Inbox folder: {total} messages, {unread} unread · cached details: {stats['inbox_cached']} inbox, {stats['sent_cached']} sent",
            f"Calendar: {stats['events_cached']} events cached",
            f"Mail {self._freshness('mail')} · calendar {self._freshness('calendar')}",
            f"Config: {self.cfg.source or '(defaults, no config file)'} · cache: {self.store.path}",
        ]
        return "\n".join(lines)
