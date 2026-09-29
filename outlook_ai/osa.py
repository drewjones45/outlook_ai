"""Run the bundled AppleScripts with osascript and translate their errors."""

from __future__ import annotations

import os
import re
import subprocess
import sys
import tempfile
import time
from importlib import resources
from typing import Callable, Sequence

from .protocol import normalize_newlines

# (script name, args, timeout seconds) -> stdout text
Runner = Callable[[str, Sequence[str], int], str]

_HINTS = {
    -1743: (
        "macOS blocked this program from controlling Microsoft Outlook. Open System Settings > "
        "Privacy & Security > Automation, find the app that runs outlook-ai (Terminal, iTerm, "
        "Claude, Python...) and switch on 'Microsoft Outlook'. If no switch is listed, run "
        "`tccutil reset AppleEvents` and try again so macOS asks. On a company-managed Mac, IT "
        "may have disabled this permission."
    ),
    -1712: (
        "Outlook did not answer in time (Apple Event timed out). Close any open Outlook dialog, "
        "then retry with a smaller window, e.g. `outlook-ai sync --days 7`."
    ),
    -600: "Microsoft Outlook isn't running (or just crashed). Open Outlook and try again.",
    -609: "Lost the connection to Outlook (it quit mid-request). Reopen Outlook and try again.",
    -1728: "Outlook couldn't find that item. It may have been moved or deleted; run a sync.",
    -1708: (
        "Outlook didn't understand the request. New Outlook for Mac supports very little "
        "AppleScript: switch to Legacy Outlook (Outlook menu > Legacy Outlook) and try again."
    ),
    -10000: (
        "Outlook refused the request (-10000). If you're on New Outlook, switch to Legacy "
        "Outlook (Outlook menu > Legacy Outlook)."
    ),
    -2741: (
        "The AppleScript failed to compile against Outlook's dictionary. Make sure Microsoft "
        "Outlook is installed in /Applications and run `outlook-ai doctor`."
    ),
    -2753: (
        "The AppleScript failed to compile against Outlook's dictionary. Make sure Microsoft "
        "Outlook is installed in /Applications and run `outlook-ai doctor`."
    ),
    1001: "Check [outlook] account in your config against `outlook-ai doctor` output.",
    1002: "A folder changed while it was being indexed. Run the sync again.",
    1003: (
        "Switch Outlook to Legacy Outlook (Outlook menu > Legacy Outlook, or Help > Revert to Legacy "
        "Outlook) if your build still offers it. Microsoft cancelled AppleScript support for New "
        "Outlook in August 2026, so there is no fix on this side."
    ),
}

_CODE = re.compile(r"\((-?\d+)\)\s*$")


class OutlookError(RuntimeError):
    def __init__(self, message: str, code: int | None = None):
        self.code = code
        self.hint = _HINTS.get(code) if code is not None else None
        text = message if not self.hint else f"{message}\n{self.hint}"
        super().__init__(text)


def parse_osascript_error(stderr: str) -> OutlookError:
    text = stderr.strip()
    code = None
    last = text.splitlines()[-1] if text else ""
    m = _CODE.search(last)
    if m:
        code = int(m.group(1))
    # "…/get_messages.applescript:120:140: execution error: Microsoft Outlook got an error: … (-1728)"
    detail = re.sub(r"^.*?(?:execution|syntax) error:\s*", "", last) or text or "osascript failed"
    return OutlookError(detail, code)


# Scripts that never talk to Outlook skip the Outlook helpers, so they compile
# even where Outlook's dictionary can't be loaded.
_NO_OUTLOOK = {"notify"}


def script_source(name: str) -> str:
    """A bundled script with the shared helper libraries appended."""
    pkg = resources.files("outlook_ai") / "applescripts"
    parts = [(pkg / f"{name}.applescript").read_text(encoding="utf-8").rstrip()]
    parts.append((pkg / "_lib.applescript").read_text(encoding="utf-8").rstrip())
    if name not in _NO_OUTLOOK:
        parts.append((pkg / "_outlook.applescript").read_text(encoding="utf-8").rstrip())
    return "\n\n".join(parts) + "\n"


_new_outlook_cache: tuple[float, bool] | None = None


def new_outlook_enabled() -> bool:
    """True when Outlook's "New Outlook" switch is on.

    New Outlook answers AppleScript with empty folders instead of errors, so a
    sync would silently "succeed" with no mail. Checked via Outlook's own
    preference; cached for a minute.
    """
    global _new_outlook_cache
    now = time.monotonic()
    if _new_outlook_cache and now - _new_outlook_cache[0] < 60:
        return _new_outlook_cache[1]
    try:
        out = subprocess.run(
            ["/usr/bin/defaults", "read", "com.microsoft.Outlook", "IsRunningNewOutlook"],
            capture_output=True, text=True, timeout=10,
        )
        enabled = out.returncode == 0 and out.stdout.strip() == "1"
    except (OSError, subprocess.TimeoutExpired):
        enabled = False
    _new_outlook_cache = (now, enabled)
    return enabled


def osascript_runner(name: str, args: Sequence[str], timeout: int) -> str:
    if sys.platform != "darwin":
        raise OutlookError(
            "outlook-ai talks to Outlook through AppleScript, which only exists on macOS. "
            "Run it on the Mac where Outlook is installed."
        )
    if name not in _NO_OUTLOOK and new_outlook_enabled():
        raise OutlookError("Outlook is running in New Outlook mode, which doesn't support AppleScript.", 1003)
    fd, path = tempfile.mkstemp(prefix=f"outlook-ai-{name}-", suffix=".applescript")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as fh:
            fh.write(script_source(name))
        try:
            # launchd jobs start without a locale; make sure names with accents
            # or emoji come back as UTF-8.
            env = dict(os.environ)
            if "UTF-8" not in (env.get("LC_ALL") or env.get("LC_CTYPE") or env.get("LANG") or "").upper():
                env["LANG"] = "en_US.UTF-8"
            proc = subprocess.run(
                ["/usr/bin/osascript", path, *args],
                capture_output=True,
                timeout=timeout,
                env=env,
            )
        except subprocess.TimeoutExpired as exc:
            raise OutlookError(
                f"Outlook script '{name}' ran longer than {timeout}s and was stopped.", -1712
            ) from exc
    finally:
        try:
            os.unlink(path)
        except OSError:
            pass
    if proc.returncode != 0:
        raise parse_osascript_error(proc.stderr.decode("utf-8", "replace"))
    return normalize_newlines(proc.stdout.decode("utf-8", "replace"))
