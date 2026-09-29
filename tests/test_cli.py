import plistlib

from outlook_ai.cli import build_parser, main


def test_parser_has_all_commands():
    parser = build_parser()
    choices = parser._subparsers._group_actions[0].choices
    assert {"doctor", "init", "sync", "inbox", "triage", "read", "search", "calendar", "prep", "draft",
            "export", "notify", "launchd", "serve", "script", "purge"} <= set(choices)


def test_script_command_prints_script_with_helpers(capsys):
    assert main(["script", "index_folder"]) == 0
    out = capsys.readouterr().out
    assert "on run argv" in out and "on isoDate(d)" in out


def test_launchd_plist_is_valid(capsys):
    assert main(["launchd", "notify"]) == 0
    plist = plistlib.loads(capsys.readouterr().out.encode())
    assert plist["Label"] == "com.outlook-ai.notify"
    assert plist["ProgramArguments"][-1] == "notify"
    assert plist["StartInterval"] == 300


def test_non_mac_errors_are_friendly(capsys, monkeypatch):
    monkeypatch.setattr("sys.platform", "linux")
    assert main(["sync"]) == 2
    assert "only exists on macOS" in capsys.readouterr().err
