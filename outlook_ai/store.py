"""Local SQLite cache of mail and calendar data.

AppleScript is slow (every property read is an Apple Event round trip), so the
cache is what makes repeated questions from Claude fast: Outlook is only asked
for what changed since the last sync. The file holds email content, so it is
created readable by your user account only.
"""

from __future__ import annotations

import json
import os
import sqlite3
import threading
import time
from datetime import datetime, timezone
from pathlib import Path

from .models import Attendee, Event, Message, Person

_SCHEMA = """
CREATE TABLE IF NOT EXISTS messages (
    id INTEGER PRIMARY KEY,
    folder TEXT NOT NULL,
    subject TEXT NOT NULL DEFAULT '',
    sender_name TEXT NOT NULL DEFAULT '',
    sender_address TEXT NOT NULL DEFAULT '',
    to_json TEXT NOT NULL DEFAULT '[]',
    cc_json TEXT NOT NULL DEFAULT '[]',
    received_ts REAL,
    sent_ts REAL,
    is_read INTEGER NOT NULL DEFAULT 1,
    flag TEXT NOT NULL DEFAULT '',
    priority TEXT NOT NULL DEFAULT '',
    categories_json TEXT NOT NULL DEFAULT '[]',
    attachments_json TEXT NOT NULL DEFAULT '[]',
    headers_json TEXT NOT NULL DEFAULT '{}',
    body TEXT NOT NULL DEFAULT '',
    body_truncated INTEGER NOT NULL DEFAULT 0,
    replied INTEGER,
    present INTEGER NOT NULL DEFAULT 1,
    fetched_at REAL NOT NULL
);
CREATE INDEX IF NOT EXISTS messages_by_folder ON messages(folder, present, received_ts);

CREATE TABLE IF NOT EXISTS folder_index (
    folder TEXT NOT NULL,
    id INTEGER NOT NULL,
    ts REAL,
    is_read INTEGER NOT NULL DEFAULT 1,
    PRIMARY KEY (folder, id)
);

CREATE TABLE IF NOT EXISTS events (
    key TEXT PRIMARY KEY,
    id INTEGER NOT NULL,
    subject TEXT NOT NULL DEFAULT '',
    start_ts REAL,
    end_ts REAL,
    all_day INTEGER NOT NULL DEFAULT 0,
    location TEXT NOT NULL DEFAULT '',
    organizer TEXT NOT NULL DEFAULT '',
    calendar TEXT NOT NULL DEFAULT '',
    is_recurring INTEGER NOT NULL DEFAULT 0,
    free_busy TEXT NOT NULL DEFAULT '',
    attendees_json TEXT NOT NULL DEFAULT '[]',
    body TEXT NOT NULL DEFAULT ''
);
CREATE INDEX IF NOT EXISTS events_by_start ON events(start_ts);

CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS notified (key TEXT PRIMARY KEY, ts REAL NOT NULL);
"""


def _ts(dt: datetime | None) -> float | None:
    return dt.timestamp() if dt is not None else None


def _dt(ts: float | None) -> datetime | None:
    if ts is None:
        return None
    return datetime.fromtimestamp(ts, tz=timezone.utc).astimezone()


def _people_json(people: list[Person]) -> str:
    return json.dumps([{"name": p.name, "address": p.address} for p in people])


def _people(raw: str) -> list[Person]:
    return [Person(name=p.get("name", ""), address=p.get("address", "")) for p in json.loads(raw or "[]")]


class Store:
    def __init__(self, path: Path):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        try:
            os.chmod(self.path.parent, 0o700)
        except OSError:
            pass
        new = not self.path.exists()
        self._lock = threading.RLock()
        self._db = sqlite3.connect(self.path, timeout=30, check_same_thread=False)
        self._db.row_factory = sqlite3.Row
        if new:
            os.chmod(self.path, 0o600)
        self._db.execute("PRAGMA journal_mode=WAL")
        self._db.executescript(_SCHEMA)
        self._db.commit()

    def close(self) -> None:
        with self._lock:
            self._db.close()

    # -- meta ---------------------------------------------------------------

    def get_meta(self, key: str, default: str = "") -> str:
        with self._lock:
            row = self._db.execute("SELECT value FROM meta WHERE key = ?", (key,)).fetchone()
        return row["value"] if row else default

    def set_meta(self, key: str, value: str) -> None:
        with self._lock, self._db:
            self._db.execute(
                "INSERT INTO meta(key, value) VALUES(?, ?) "
                "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
                (key, value),
            )

    def last_sync(self, what: str) -> float | None:
        raw = self.get_meta(f"last_sync_{what}")
        return float(raw) if raw else None

    def mark_synced(self, what: str, when: float | None = None) -> None:
        self.set_meta(f"last_sync_{what}", str(when if when is not None else time.time()))

    # -- folder index ---------------------------------------------------------

    def replace_folder_index(self, folder: str, rows: list[tuple[int, datetime | None, bool]]) -> None:
        with self._lock, self._db:
            self._db.execute("DELETE FROM folder_index WHERE folder = ?", (folder,))
            self._db.executemany(
                "INSERT OR REPLACE INTO folder_index(folder, id, ts, is_read) VALUES(?, ?, ?, ?)",
                [(folder, mid, _ts(when), int(read)) for mid, when, read in rows],
            )
            present = {mid for mid, _, _ in rows}
            ids = [r["id"] for r in self._db.execute("SELECT id FROM messages WHERE folder = ?", (folder,))]
            gone = [(mid,) for mid in ids if mid not in present]
            back = [(mid,) for mid in ids if mid in present]
            self._db.executemany("UPDATE messages SET present = 0 WHERE id = ?", gone)
            self._db.executemany("UPDATE messages SET present = 1 WHERE id = ?", back)
            self._db.executemany(
                "UPDATE messages SET is_read = ? WHERE id = ?",
                [(int(read), mid) for mid, _, read in rows],
            )

    def folder_counts(self, folder: str) -> tuple[int, int]:
        with self._lock:
            row = self._db.execute(
                "SELECT COUNT(*) AS total, SUM(CASE WHEN is_read = 0 THEN 1 ELSE 0 END) AS unread "
                "FROM folder_index WHERE folder = ?",
                (folder,),
            ).fetchone()
        return int(row["total"] or 0), int(row["unread"] or 0)

    def newest_indexed(self, folder: str) -> datetime | None:
        with self._lock:
            row = self._db.execute("SELECT MAX(ts) AS ts FROM folder_index WHERE folder = ?", (folder,)).fetchone()
        return _dt(row["ts"]) if row and row["ts"] is not None else None

    def ids_to_fetch(self, folder: str, since: datetime | None, limit: int) -> list[int]:
        """Indexed ids in the window that aren't cached yet, newest first."""
        sql = (
            "SELECT f.id FROM folder_index f LEFT JOIN messages m ON m.id = f.id "
            "WHERE f.folder = ? AND (m.id IS NULL OR m.folder != f.folder)"
        )
        params: list = [folder]
        if since is not None:
            sql += " AND f.ts >= ?"
            params.append(since.timestamp())
        sql += " ORDER BY f.ts DESC LIMIT ?"
        params.append(limit)
        with self._lock:
            return [r["id"] for r in self._db.execute(sql, params)]

    # -- messages -------------------------------------------------------------

    def upsert_messages(self, messages: list[Message]) -> None:
        now = time.time()
        rows = [
            (
                m.id, m.folder, m.subject, m.sender.name, m.sender.address,
                _people_json(m.to), _people_json(m.cc), _ts(m.received), _ts(m.sent),
                int(m.is_read), m.flag, m.priority, json.dumps(m.categories),
                json.dumps(m.attachments), json.dumps(m.headers), m.body, int(m.body_truncated),
                None if m.replied is None else int(m.replied), now,
            )
            for m in messages
        ]
        with self._lock, self._db:
            self._db.executemany(
                """INSERT INTO messages(id, folder, subject, sender_name, sender_address, to_json,
                    cc_json, received_ts, sent_ts, is_read, flag, priority, categories_json,
                    attachments_json, headers_json, body, body_truncated, replied, present, fetched_at)
                VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, ?)
                ON CONFLICT(id) DO UPDATE SET
                    folder = excluded.folder, subject = excluded.subject,
                    sender_name = excluded.sender_name, sender_address = excluded.sender_address,
                    to_json = excluded.to_json, cc_json = excluded.cc_json,
                    received_ts = excluded.received_ts, sent_ts = excluded.sent_ts,
                    is_read = excluded.is_read, flag = excluded.flag, priority = excluded.priority,
                    categories_json = excluded.categories_json,
                    attachments_json = excluded.attachments_json,
                    headers_json = excluded.headers_json, body = excluded.body,
                    body_truncated = excluded.body_truncated, replied = excluded.replied,
                    present = 1, fetched_at = excluded.fetched_at""",
                rows,
            )

    @staticmethod
    def _message(row: sqlite3.Row) -> Message:
        return Message(
            id=row["id"],
            folder=row["folder"],
            subject=row["subject"],
            sender=Person(row["sender_name"], row["sender_address"]),
            to=_people(row["to_json"]),
            cc=_people(row["cc_json"]),
            received=_dt(row["received_ts"]),
            sent=_dt(row["sent_ts"]),
            is_read=bool(row["is_read"]),
            flag=row["flag"],
            priority=row["priority"],
            categories=json.loads(row["categories_json"]),
            attachments=json.loads(row["attachments_json"]),
            headers=json.loads(row["headers_json"]),
            body=row["body"],
            body_truncated=bool(row["body_truncated"]),
            replied=None if row["replied"] is None else bool(row["replied"]),
        )

    def messages(
        self,
        folder: str,
        since: datetime | None = None,
        *,
        unread_only: bool = False,
        limit: int | None = None,
    ) -> list[Message]:
        sql = "SELECT * FROM messages WHERE folder = ? AND present = 1"
        params: list = [folder]
        if since is not None:
            sql += " AND COALESCE(received_ts, sent_ts) >= ?"
            params.append(since.timestamp())
        if unread_only:
            sql += " AND is_read = 0"
        sql += " ORDER BY COALESCE(received_ts, sent_ts) DESC"
        if limit:
            sql += " LIMIT ?"
            params.append(limit)
        with self._lock:
            return [self._message(r) for r in self._db.execute(sql, params)]

    def message(self, message_id: int) -> Message | None:
        with self._lock:
            row = self._db.execute("SELECT * FROM messages WHERE id = ?", (message_id,)).fetchone()
        return self._message(row) if row else None

    def by_thread_key(self, key: str, thread_key_fn) -> list[Message]:
        """All cached messages (inbox and sent) whose subject normalizes to `key`."""
        if not key:
            return []
        # Cheap SQL prefilter on a distinctive fragment, exact match in Python.
        fragment = max(key.split(), key=len)
        with self._lock:
            rows = self._db.execute(
                "SELECT * FROM messages WHERE subject LIKE ? ESCAPE '\\' "
                "ORDER BY COALESCE(received_ts, sent_ts)",
                (f"%{_like_escape(fragment)}%",),
            ).fetchall()
        return [m for m in map(self._message, rows) if thread_key_fn(m.subject) == key]

    def search(self, query: str, since: datetime | None, limit: int) -> list[Message]:
        terms = [t for t in query.split() if t]
        if not terms:
            return []
        sql = "SELECT * FROM messages WHERE present = 1"
        params: list = []
        for term in terms:
            like = f"%{_like_escape(term)}%"
            sql += (
                " AND (subject LIKE ? ESCAPE '\\' OR sender_name LIKE ? ESCAPE '\\'"
                " OR sender_address LIKE ? ESCAPE '\\' OR to_json LIKE ? ESCAPE '\\'"
                " OR body LIKE ? ESCAPE '\\')"
            )
            params.extend([like] * 5)
        if since is not None:
            sql += " AND COALESCE(received_ts, sent_ts) >= ?"
            params.append(since.timestamp())
        sql += " ORDER BY COALESCE(received_ts, sent_ts) DESC LIMIT ?"
        params.append(limit)
        with self._lock:
            return [self._message(r) for r in self._db.execute(sql, params)]

    def from_addresses(self, addresses: set[str], since: datetime | None, limit: int) -> list[Message]:
        if not addresses:
            return []
        marks = ",".join("?" * len(addresses))
        sql = f"SELECT * FROM messages WHERE present = 1 AND folder = 'inbox' AND lower(sender_address) IN ({marks})"
        params: list = sorted(addresses)
        if since is not None:
            sql += " AND received_ts >= ?"
            params.append(since.timestamp())
        sql += " ORDER BY received_ts DESC LIMIT ?"
        params.append(limit)
        with self._lock:
            return [self._message(r) for r in self._db.execute(sql, params)]

    # -- calendar -------------------------------------------------------------

    def replace_events(self, events: list[Event], window: tuple[datetime, datetime]) -> None:
        with self._lock, self._db:
            self._db.execute("DELETE FROM events")
            self._db.executemany(
                """INSERT OR REPLACE INTO events(key, id, subject, start_ts, end_ts, all_day, location,
                    organizer, calendar, is_recurring, free_busy, attendees_json, body)
                VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                [
                    (
                        e.occurrence_key(), e.id, e.subject, _ts(e.start), _ts(e.end), int(e.all_day),
                        e.location, e.organizer, e.calendar, int(e.is_recurring), e.free_busy,
                        json.dumps([a.__dict__ for a in e.attendees]), e.body,
                    )
                    for e in events
                ],
            )
        self.set_meta("calendar_window", json.dumps([window[0].timestamp(), window[1].timestamp()]))

    def calendar_window(self) -> tuple[datetime, datetime] | None:
        raw = self.get_meta("calendar_window")
        if not raw:
            return None
        start, end = json.loads(raw)
        return _dt(start), _dt(end)

    def events(self, start: datetime, end: datetime) -> list[Event]:
        with self._lock:
            rows = self._db.execute(
                "SELECT * FROM events WHERE start_ts < ? AND end_ts > ? ORDER BY start_ts",
                (end.timestamp(), start.timestamp()),
            ).fetchall()
        return [
            Event(
                id=r["id"],
                subject=r["subject"],
                start=_dt(r["start_ts"]),
                end=_dt(r["end_ts"]),
                all_day=bool(r["all_day"]),
                location=r["location"],
                organizer=r["organizer"],
                calendar=r["calendar"],
                is_recurring=bool(r["is_recurring"]),
                free_busy=r["free_busy"],
                attendees=[Attendee(**a) for a in json.loads(r["attendees_json"])],
                body=r["body"],
            )
            for r in rows
        ]

    # -- notifications --------------------------------------------------------

    def claim_notification(self, key: str) -> bool:
        """True the first time `key` is claimed; False if already notified."""
        with self._lock, self._db:
            self._db.execute("DELETE FROM notified WHERE ts < ?", (time.time() - 7 * 86400,))
            cur = self._db.execute(
                "INSERT OR IGNORE INTO notified(key, ts) VALUES(?, ?)", (key, time.time())
            )
            return cur.rowcount == 1

    # -- housekeeping ---------------------------------------------------------

    def stats(self) -> dict[str, int]:
        with self._lock:
            out = {}
            for folder in ("inbox", "sent"):
                out[f"{folder}_cached"] = self._db.execute(
                    "SELECT COUNT(*) FROM messages WHERE folder = ? AND present = 1", (folder,)
                ).fetchone()[0]
            out["events_cached"] = self._db.execute("SELECT COUNT(*) FROM events").fetchone()[0]
        return out


def _like_escape(text: str) -> str:
    return text.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_")
