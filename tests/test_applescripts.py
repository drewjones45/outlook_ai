"""Static checks for the bundled AppleScripts.

AppleScript can only be compiled on a Mac, so these tests catch the mistakes
that would otherwise surface there: unbalanced blocks, handler calls that
Outlook would receive instead of the script, reserved words used as variable
names, and records whose field count drifted from the Python parser.
"""

import re
from importlib import resources

import pytest

from outlook_ai.cli import SCRIPTS
from outlook_ai.osa import script_source
from outlook_ai.outlook import EVT_FIELDS, MSG_FIELDS

PKG = resources.files("outlook_ai") / "applescripts"
ALL_FILES = sorted(p.name for p in PKG.iterdir() if p.name.endswith(".applescript"))

RESERVED = set(
    """about above after against and apart around as aside at back before beginning behind below beneath
    beside between but by considering contain contains continue copy div does eighth else end equal equals
    error every exit false fifth first for fourth from front get given global if ignoring in instead into is
    it its last local me middle mod my ninth not of on onto or out over prop property put ref reference repeat
    return returning script second set seventh since sixth some tell tenth that the then third through thru
    timeout times to transaction true try until where while whose with without result text list record date
    string number integer real item character word paragraph count length offset class""".split()
)
# Single-word terms from Standard Additions (StandardAdditions.sdef). Scripting-addition
# terms are in scope everywhere, so a variable named e.g. "kind" compiles as a property
# of whatever the script is talking to and fails at run time.
RESERVED |= set(
    """alias appletalk as ask beep before buttons caution critical delay desktop displaying down editors eof
    extensions folder fonts for from has help host in informational invisibles ip kind locked message
    modulation name no note of offset password path pitch plugins port preferences printmonitor properties
    read replacing round rounding say scheme short showing size startup stationery stop subtitle summarize
    to trash until up url using visible voices volume warning write yes""".split()
)


def _code_lines(source: str) -> list[str]:
    lines = []
    for line in source.splitlines():
        line = re.sub(r'"(?:[^"\\]|\\.)*"', '""', line)  # blank out string literals
        line = line.split("--", 1)[0].rstrip()
        lines.append(line)
    return lines


def _handlers(source: str) -> set[str]:
    return set(re.findall(r"^\s*on\s+(\w+)\(", source, re.MULTILINE))


@pytest.mark.parametrize("name", ALL_FILES)
def test_scripts_are_ascii(name):
    # osascript reads plain-text scripts most reliably when they are ASCII.
    (PKG / name).read_text(encoding="ascii")


@pytest.mark.parametrize("name", SCRIPTS)
def test_blocks_balance(name):
    stack: list[str] = []
    for n, line in enumerate(_code_lines(script_source(name)), 1):
        s = line.strip()
        if not s:
            continue
        m = re.match(r"end(?:\s+(\w+))?$", s)
        if m:
            kind = m.group(1) or (stack[-1] if stack else "")
            assert stack, f"{name}:{n}: '{s}' with nothing open"
            top = stack.pop()
            assert kind == top, f"{name}:{n}: '{s}' closes '{top}'"
            continue
        if re.match(r"tell\s+application\s+\"\"\s*$", s):
            stack.append("tell")
        elif re.match(r"(?:else\s+)?if\b.*\bthen$", s):
            if not s.startswith("else"):
                stack.append("if")
        elif re.match(r"repeat\b", s):
            stack.append("repeat")
        elif s == "try":
            stack.append("try")
        elif re.match(r"with\s+timeout\b", s):
            stack.append("timeout")
        elif re.match(r"script\s+\w+$", s):
            stack.append("script")
        elif m := re.match(r"on\s+(\w+)\s*(?:\(|\w|$)", s):
            if m.group(1) != "error":  # "on error" belongs to try
                stack.append(m.group(1))
    assert not stack, f"{name}: unclosed {stack}"


@pytest.mark.parametrize("name", SCRIPTS)
def test_handler_calls_inside_tell_blocks_use_my(name):
    source = script_source(name)
    handlers = _handlers(source)
    depth = 0
    for n, line in enumerate(_code_lines(source), 1):
        s = line.strip()
        if re.match(r"tell\s+application\s+\"\"\s*$", s):
            depth += 1
        elif s == "end tell":
            depth -= 1
        elif depth:
            for call in re.finditer(r"(\bmy\s+)?\b(\w+)\(", s):
                if call.group(2) in handlers:
                    assert call.group(1), f"{name}:{n}: '{call.group(2)}(' inside a tell block needs 'my'"


@pytest.mark.parametrize("name", SCRIPTS)
def test_no_reserved_words_as_variables(name):
    source = "\n".join(_code_lines(script_source(name)))
    names = set(re.findall(r"\bset\s+(?:end of\s+)?(?:\w+'s\s+)?(\w+)\s+to\b", source))
    names |= set(re.findall(r"\brepeat with\s+(\w+)\s+(?:in|from)\b", source))
    names |= set(re.findall(r"\bproperty\s+(\w+)\s*:", source))
    for params in re.findall(r"^\s*on\s+\w+\(([^)]*)\)", source, re.MULTILINE):
        names |= {p.strip() for p in params.split(",") if p.strip()}
    bad = sorted(v for v in names if v.lower() in RESERVED)
    assert not bad, f"{name}: reserved words used as variables: {bad}"


@pytest.mark.parametrize("name", SCRIPTS)
def test_called_handlers_exist(name):
    source = script_source(name)
    defined = _handlers(source) | {"run"}
    called = set(re.findall(r"\bmy\s+(\w+)\(", source))
    assert called <= defined, f"{name}: undefined handlers {sorted(called - defined)}"


def _record_arity(source: str, tag: str) -> int:
    start = source.index(f'makeRecord({{"{tag}"')
    i = source.index("{", start)
    depth, fields = 0, 1
    for ch in source[i:]:
        if ch in "({":
            depth += 1
        elif ch in ")}":
            depth -= 1
            if depth == 0:
                return fields
        elif ch == "," and depth == 1:
            fields += 1
    raise AssertionError("unterminated record")


def test_record_layouts_match_python_parsers():
    assert _record_arity(script_source("fetch_messages"), "MSG") == 1 + len(MSG_FIELDS)
    assert _record_arity(script_source("list_events"), "EVT") == 1 + len(EVT_FIELDS)
