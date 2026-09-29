"""Incremental sync from Outlook into the local cache."""

from __future__ import annotations

import fcntl
import os
import time
from contextlib import contextmanager
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from typing import Callable, Iterator

from .config import Config, data_dir
from .osa import OutlookError
from .outlook import Outlook
from .recurrence import expand_events
from .store import Store

Progress = Callable[[str], None]


@dataclass
class SyncReport:
    inbox_total: int = 0
    inbox_unread: int = 0
    inbox_fetched: int = 0
    sent_fetched: int = 0
    events: int = 0
    errors: list[str] = field(default_factory=list)
    seconds: float = 0.0

    def summary(self) -> str:
        parts = [
            f"inbox: {self.inbox_total} messages ({self.inbox_unread} unread), {self.inbox_fetched} new fetched",
            f"sent: {self.sent_fetched} new fetched",
            f"calendar: {self.events} events cached",
            f"took {self.seconds:.1f}s",
        ]
        if self.errors:
            parts.append(f"{len(self.errors)} problem(s): " + "; ".join(self.errors[:3]))
        return " · ".join(parts)


class SyncBusy(RuntimeError):
    pass


@contextmanager
def sync_lock(block: bool) -> Iterator[None]:
    """One sync at a time across the CLI, MCP server and launchd jobs."""
    path = data_dir() / "sync.lock"
    path.parent.mkdir(parents=True, exist_ok=True)
    fd = os.open(path, os.O_CREAT | os.O_RDWR, 0o600)
    try:
        try:
            fcntl.flock(fd, fcntl.LOCK_EX | (0 if block else fcntl.LOCK_NB))
        except BlockingIOError as exc:
            raise SyncBusy("another sync is already running") from exc
        yield
    finally:
        os.close(fd)


class Syncer:
    def __init__(self, cfg: Config, store: Store, outlook: Outlook, progress: Progress | None = None):
        self.cfg = cfg
        self.store = store
        self.outlook = outlook
        self.progress = progress or (lambda _msg: None)

    # -- mail -----------------------------------------------------------------

    def _index(self, role: str) -> list[tuple[int, datetime | None, bool]]:
        try:
            return self.outlook.index_folder(role)
        except OutlookError as exc:
            if exc.code == 1002:  # folder changed mid-index; one retry is enough
                return self.outlook.index_folder(role)
            raise

    def _sync_folder(self, role: str, days: int, limit: int, body_chars: int, report: SyncReport) -> int:
        self.progress(f"Indexing {role}...")
        rows = self._index(role)
        self.store.replace_folder_index(role, rows)
        since = datetime.now().astimezone() - timedelta(days=days) if days > 0 else None
        todo = self.store.ids_to_fetch(role, since, limit)
        fetched = 0
        batch = max(1, self.cfg.sync.batch_size)
        for i in range(0, len(todo), batch):
            chunk = todo[i : i + batch]
            self.progress(f"Fetching {role} {i + 1}-{i + len(chunk)} of {len(todo)}...")
            messages, errors = self.outlook.fetch_messages(chunk, role, body_chars)
            self.store.upsert_messages(messages)
            fetched += len(messages)
            report.errors.extend(errors)
        return fetched

    def sync_mail(self, report: SyncReport, days: int | None = None, max_messages: int | None = None) -> None:
        s = self.cfg.sync
        days = s.days if days is None else days
        limit = s.max_messages if max_messages is None else max_messages
        report.inbox_fetched = self._sync_folder("inbox", days, limit, s.body_chars, report)
        report.inbox_total, report.inbox_unread = self.store.folder_counts("inbox")
        try:
            # Sent Items covers the same window, so "already replied" works for every inbox message.
            report.sent_fetched = self._sync_folder("sent", days, limit, 4000, report)
        except OutlookError as exc:
            # Sent Items only improves "already replied" detection; don't fail the sync.
            report.errors.append(f"sent items: {exc}")
        self.store.mark_synced("mail")

    # -- calendar -------------------------------------------------------------

    def sync_calendar(self, report: SyncReport) -> None:
        s = self.cfg.sync
        today = datetime.now().astimezone().replace(hour=0, minute=0, second=0, microsecond=0)
        start = today - timedelta(days=s.calendar_days_back)
        end = today + timedelta(days=s.calendar_days_ahead + 1)
        self.progress("Reading calendar...")
        raw, moved = self.outlook.list_events(start, end)
        events, problems = expand_events(raw, start, end, moved)
        wanted = {c.lower() for c in self.cfg.outlook.calendars}
        if wanted:
            events = [e for e in events if e.calendar.lower() in wanted]
        self.store.replace_events(events, (start, end))
        self.store.mark_synced("calendar")
        report.events = len(events)
        report.errors.extend(problems)

    # -- entry points -----------------------------------------------------------

    def sync_all(
        self,
        *,
        days: int | None = None,
        max_messages: int | None = None,
        mail: bool = True,
        calendar: bool = True,
        block: bool = True,
    ) -> SyncReport:
        report = SyncReport()
        started = time.monotonic()
        with sync_lock(block):
            if mail:
                self.sync_mail(report, days=days, max_messages=max_messages)
            if calendar:
                try:
                    self.sync_calendar(report)
                except OutlookError as exc:
                    if not mail:
                        raise
                    report.errors.append(f"calendar: {exc}")
        report.seconds = time.monotonic() - started
        return report

    def is_stale(self, what: str) -> bool:
        last = self.store.last_sync(what)
        return last is None or time.time() - last > self.cfg.sync.stale_minutes * 60

    def ensure_fresh(self, *, mail: bool = False, calendar: bool = False) -> str:
        """Re-sync stale parts of the cache. Returns a note for the reader when
        the data may be out of date (Outlook unreachable, sync in progress)."""
        want_mail = mail and self.is_stale("mail")
        want_cal = calendar and self.is_stale("calendar")
        if not (want_mail or want_cal):
            return ""
        first_time = want_mail and self.store.last_sync("mail") is None
        try:
            if first_time:
                # A full first sync can take minutes, longer than a tool call should.
                # Grab the last week now; `outlook-ai sync` backfills the rest.
                days = self.cfg.sync.days
                self.sync_all(
                    days=7 if days <= 0 else min(7, days), max_messages=200,
                    mail=True, calendar=want_cal, block=False,
                )
                return (
                    "Note: first sync, so only the last 7 days were loaded. "
                    "Run `outlook-ai sync` in Terminal to backfill the full window."
                )
            self.sync_all(mail=want_mail, calendar=want_cal, block=False)
            return ""
        except SyncBusy:
            return "Note: a sync is already running; showing cached data."
        except OutlookError as exc:
            last = self.store.last_sync("mail" if want_mail else "calendar")
            age = f"last synced {int((time.time() - last) / 60)} min ago" if last else "never synced"
            return f"Note: couldn't refresh from Outlook ({exc}). Showing cached data ({age})."
