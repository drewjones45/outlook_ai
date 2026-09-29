"""Wire format between the AppleScripts and Python.

AppleScript has no JSON encoder, and ad-hoc delimiters such as "|" or "::"
break the first time an email subject contains them. The scripts therefore
use the ASCII information-separator control characters, which never occur in
normal text, and strip them from every value before emitting it:

    record  := TAG US field US field ...      records are separated by RS
    list    := item GS item GS ...            for list-valued fields
    item    := part FS part ...               e.g. name FS address

Dates are emitted as local wall-clock time "YYYY-MM-DDTHH:MM:SS" (the Mac's
time zone, which is what Outlook hands AppleScript) and converted here.
"""

from __future__ import annotations

import html
import re
from datetime import datetime

from .models import Attendee, Person

FS = "\x1c"  # file separator: parts of a list item
GS = "\x1d"  # group separator: items of a list field
RS = "\x1e"  # record separator
US = "\x1f"  # unit separator: fields of a record


def normalize_newlines(text: str) -> str:
    # Outlook's plain-text bodies often use classic-Mac CR line endings.
    return text.replace("\r\n", "\n").replace("\r", "\n")


def parse_records(output: str) -> list[list[str]]:
    """Split osascript stdout into records; each record is a list of fields."""
    records = []
    for chunk in output.split(RS):
        chunk = chunk.strip("\n")
        if not chunk:
            continue
        records.append(chunk.split(US))
    return records


def parse_list(value: str) -> list[list[str]]:
    if not value:
        return []
    return [item.split(FS) for item in value.split(GS) if item]


def parse_people(value: str) -> list[Person]:
    people = []
    for parts in parse_list(value):
        name = parts[0].strip() if parts else ""
        address = parts[1].strip() if len(parts) > 1 else ""
        if name or address:
            people.append(Person(name=name, address=address))
    return people


def parse_person(value: str) -> Person:
    people = parse_people(value)
    return people[0] if people else Person()


def parse_attendees(value: str) -> list[Attendee]:
    out = []
    for parts in parse_list(value):
        parts = parts + [""] * (4 - len(parts))
        name, address, kind, status = (p.strip() for p in parts[:4])
        if name or address:
            out.append(Attendee(name=name, address=address, kind=kind, status=status))
    return out


def parse_strings(value: str) -> list[str]:
    return [parts[0].strip() for parts in parse_list(value) if parts and parts[0].strip()]


def parse_bool(value: str) -> bool:
    return value.strip().lower() in {"1", "true", "yes"}


def parse_local_datetime(value: str) -> datetime | None:
    """Local wall-clock "YYYY-MM-DDTHH:MM:SS" -> timezone-aware datetime."""
    value = value.strip()
    if not value:
        return None
    try:
        naive = datetime.fromisoformat(value)
    except ValueError:
        return None
    if naive.tzinfo is not None:
        return naive
    # astimezone() on a naive datetime interprets it in the system time zone,
    # picking the right UTC offset for that date (DST included).
    return naive.astimezone()


def format_local_datetime(dt: datetime) -> str:
    """Aware or naive datetime -> the local wall-clock form the scripts parse."""
    if dt.tzinfo is not None:
        dt = dt.astimezone()
    return dt.strftime("%Y-%m-%dT%H:%M:%S")


# The only headers worth keeping: threading and bulk-mail detection.
WANTED_HEADERS = (
    "message-id", "in-reply-to", "references", "list-unsubscribe", "list-id",
    "precedence", "auto-submitted",
)


def parse_header_block(raw: str) -> dict[str, str]:
    """Raw RFC 5322 header block -> {lower-name: unfolded value} for WANTED_HEADERS."""
    headers: dict[str, str] = {}
    current = None
    for line in normalize_newlines(raw or "").split("\n"):
        if not line.strip():
            if headers or current:
                break  # blank line ends the header block
            continue
        if line[0] in " \t":
            if current:
                headers[current] += " " + line.strip()
            continue
        name, sep, value = line.partition(":")
        name = name.strip().lower()
        current = None
        if sep and name in WANTED_HEADERS and name not in headers:
            headers[name] = value.strip()
            current = name
    return {k: " ".join(v.split()) for k, v in headers.items()}


# Outlook enum values reach Python as text: either the dictionary term
# ("not completed") or, when AppleScript lacks the terminology at coercion
# time, the raw form ("«constant ****FlNC»"). Codes are from Outlook's sdef.
_ENUMS = {
    "flag": [("FlNF", "not flagged", ""), ("FlNC", "not completed", "flagged"), ("FlCo", "completed", "completed")],
    "priority": [("PrHi", "priority high", "high"), ("PrLo", "priority low", "low"), ("PrNr", "priority normal", "normal")],
    "attendee_type": [("ATrq", "required attendee type", "required"), ("ATop", "optional attendee type", "optional"),
                      ("ATrs", "resource attendee type", "resource")],
    "status": [("ASte", "tentatively accepted", "tentative"), ("ASac", "accepted", "accepted"),
               ("ASde", "declined", "declined"), ("ASnr", "not responded", "none")],
    "free_busy": [("eSFr", "free", "free"), ("eSBu", "busy", "busy"), ("eSTe", "tentative", "tentative"),
                  ("eSOO", "out of office", "out of office")],
}


def decode_enum(kind: str, raw: str) -> str:
    value = (raw or "").strip()
    if not value:
        return ""
    for code, _term, out in _ENUMS[kind]:
        if code in value:
            return out
    lowered = value.lower()
    for _code, term, out in _ENUMS[kind]:
        if lowered == term:
            return out
    return ""


_TAG = re.compile(r"<[^>]+>")
_DROP = re.compile(r"<(script|style|head)\b.*?</\1\s*>", re.IGNORECASE | re.DOTALL)
_BREAK = re.compile(r"<\s*(br|/p|/div|/tr|/li|/h\d)\b[^>]*>", re.IGNORECASE)


def html_to_text(value: str) -> str:
    """Crude HTML -> text, for the rare message whose plain-text body is empty."""
    text = _DROP.sub(" ", value)
    text = _BREAK.sub("\n", text)
    text = html.unescape(_TAG.sub("", text))
    lines = [" ".join(line.split()) for line in text.split("\n")]
    return re.sub(r"\n{3,}", "\n\n", "\n".join(lines)).strip()


def looks_like_html(value: str) -> bool:
    head = value[:2000].lower()
    return "<html" in head or "<body" in head or ("<div" in head and "</div>" in value.lower())
