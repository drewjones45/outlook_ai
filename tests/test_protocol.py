from datetime import timedelta

from outlook_ai.protocol import (
    FS,
    GS,
    RS,
    US,
    decode_enum,
    format_local_datetime,
    html_to_text,
    looks_like_html,
    parse_attendees,
    parse_header_block,
    parse_local_datetime,
    parse_people,
    parse_records,
)


def test_records_survive_delimiter_lookalikes():
    subject = 'Re: a | b :: c, "quoted"; tab\there'
    out = RS.join([US.join(["MSG", "1", subject]), US.join(["MSG", "2", "line1\nline2"])]) + "\n"
    records = parse_records(out)
    assert records == [["MSG", "1", subject], ["MSG", "2", "line1\nline2"]]


def test_empty_output_has_no_records():
    assert parse_records("") == []
    assert parse_records("\n") == []


def test_people_and_attendees():
    raw = GS.join([FS.join(["Jane Doe", "jane@x.com"]), FS.join(["", "bob@x.com"]), FS.join(["", ""])])
    people = parse_people(raw)
    assert [p.display() for p in people] == ["Jane Doe <jane@x.com>", "bob@x.com"]
    att = parse_attendees(FS.join(["Jane", "jane@x.com", "required", "accepted"]))
    assert att[0].status == "accepted" and att[0].kind == "required"


def test_local_datetime_round_trip_across_dst():
    # TZ is America/Los_Angeles in tests: July is -07:00, December is -08:00.
    summer = parse_local_datetime("2026-07-01T09:30:00")
    winter = parse_local_datetime("2026-12-01T09:30:00")
    assert summer.utcoffset() == timedelta(hours=-7)
    assert winter.utcoffset() == timedelta(hours=-8)
    assert format_local_datetime(summer) == "2026-07-01T09:30:00"
    assert parse_local_datetime("") is None
    assert parse_local_datetime("garbage") is None


def test_header_block_keeps_wanted_headers_and_unfolds():
    raw = (
        "Received: from a by b;\r\n\tTue, 29 Sep 2026\r\n"
        "Message-ID: <abc@x>\r\n"
        "References: <one@x>\r\n <two@x>\r\n"
        "LIST-UNSUBSCRIBE: <mailto:u@x>\r\n"
        "Subject: ignored\r\n"
    )
    assert parse_header_block(raw) == {
        "message-id": "<abc@x>",
        "references": "<one@x> <two@x>",
        "list-unsubscribe": "<mailto:u@x>",
    }
    assert parse_header_block("") == {}


def test_enum_values_decode_from_terms_and_raw_codes():
    assert decode_enum("flag", "not completed") == "flagged"
    assert decode_enum("flag", "\u00abconstant ****FlNC\u00bb") == "flagged"
    assert decode_enum("flag", "completed") == "completed"
    assert decode_enum("flag", "not flagged") == ""
    assert decode_enum("priority", "\u00abconstant ****PrHi\u00bb") == "high"
    assert decode_enum("status", "tentatively accepted") == "tentative"
    assert decode_enum("status", "accepted") == "accepted"
    assert decode_enum("free_busy", "out of office") == "out of office"
    assert decode_enum("status", "something else") == ""


def test_html_fallback_body():
    html_body = "<html><head><style>p{}</style></head><body><p>Hi&nbsp;Pat,</p><div>Line two<br>three</div></body></html>"
    assert looks_like_html(html_body)
    assert html_to_text(html_body) == "Hi Pat,\nLine two\nthree"
    assert not looks_like_html("plain text with a < sign")
