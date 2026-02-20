"""Tests for the TXT parser."""

import os
import tempfile

from xlbridge.parser import parse_txt


def _write_tmp(content: str) -> str:
    fd, path = tempfile.mkstemp(suffix=".txt")
    with os.fdopen(fd, "w", encoding="utf-8") as f:
        f.write(content)
    return path


def test_parse_basic():
    path = _write_tmp(
        "# comment\n"
        "\n"
        "[Sheet1]!A1|Hello\n"
        "[Sheet1]!B2|World\n"
    )
    entries = parse_txt(path)
    os.unlink(path)

    assert len(entries) == 2
    assert entries[0].sheet == "Sheet1"
    assert entries[0].coord == "A1"
    assert entries[0].value == "Hello"
    assert entries[1].coord == "B2"
    assert entries[1].value == "World"


def test_parse_multiline_escaped():
    path = _write_tmp("[Sheet1]!A1|Line1\\nLine2\n")
    entries = parse_txt(path)
    os.unlink(path)

    assert len(entries) == 1
    assert entries[0].value == "Line1\nLine2"


def test_parse_skips_invalid_lines():
    path = _write_tmp(
        "# header\n"
        "bad line\n"
        "[Sheet1]!A1|OK\n"
    )
    entries = parse_txt(path)
    os.unlink(path)

    assert len(entries) == 1
    assert entries[0].value == "OK"


def test_parse_pipe_in_value():
    path = _write_tmp("[Sheet1]!A1|val|ue|with|pipes\n")
    entries = parse_txt(path)
    os.unlink(path)

    assert len(entries) == 1
    assert entries[0].value == "val|ue|with|pipes"


def test_parse_japanese():
    path = _write_tmp("[Sheet1]!A1|プロジェクト名\n")
    entries = parse_txt(path)
    os.unlink(path)

    assert entries[0].value == "プロジェクト名"
